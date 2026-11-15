/**
 * BrandDeck deterministic deck generation: user input + style -> CardDeck.
 *
 * This is the "make me a deck out of this text, in this style" entry point. It is
 * deliberately LLM-free and pure: the same input always produces byte-identical
 * output (no `Date`, no randomness, no I/O), so the panel can offer a "regenerate"
 * button that never surprises the user. The user then fine-tunes geometry in the
 * panel (or in the Markdown outline) — this module only has to produce a valid,
 * non-overlapping starting point in the house style.
 *
 * TWO STAGES
 * 1. {@link parseSourceToOutline} — text -> outline
 *    `{title, subtitle?, slides: [{title, bullets, notes?, kind?}]}`.
 *    Tolerant: never throws, on empty/odd input it returns
 *    `{title: '未命名演示', slides: []}`.
 * 2. {@link composeDeck} — outline + theme + style -> `{deck, warnings}`.
 *    The deck is schema-valid (`CardDeckSchema`) and already inside `innerBox(theme)`,
 *    so `repairDeck` must report zero *geometry* warnings for it.
 *
 * LAYOUT MODEL (1280x720, see schema.js)
 * - Everything lives inside `innerBox(theme)` = {x:64, y:64, w:1152, h:592}.
 * - The top-right logo reserve (150x48 + 16 gap) is kept free of TEXT by
 *   shortening titles to `FRAME.titleRight` (1050px). Body copy starts below
 *   y=128 so the full safe width is available for it. Decorative shapes may pass
 *   under the reserve — a logo sits on top of a bar, never on top of a sentence.
 * - Text boxes are measured with the same `estimateLines` the validator uses, so
 *   a box whose height equals `ceil(estimate.heightPx)` can never be reported as
 *   overflowing. Heights are grown, never cropped (except for the last-resort
 *   emergency shrink documented in `placeList`).
 * - Density caps bullets per slide and spills the rest onto `<title>（续）`
 *   continuation slides. `maxSlides` is a hard cap (cover + closing included).
 * - `cards` rows are sized to their content (`textTop + textH + padBottom`), never
 *   to a fixed band: a one-line bullet gets a short card, not a mostly-empty box.
 *   Both cards of a row share the taller text height so the grid stays even, and a
 *   slide whose rows would leave the safe box spills onto a continuation slide.
 *
 * COLOURS
 * - `theme` supplies the tokens, `style.colors` overrides them (hex or
 *   `$colors.x`), and the RESOLVED hex is written into every element. The deck is
 *   therefore self-contained: it renders the same even if the theme is re-parsed.
 * - `style.fonts` is applied by the CALLER on the theme (`applyStyleToTheme` in
 *   styles.js) and the compiler resolves a face per text role, so font overrides
 *   reach the exported .pptx without touching the DSL (whose text style is strict
 *   and carries no `fontFace`). `deck.metadata.fonts` records what is in force.
 *
 * Judgement calls are documented next to the code that makes them and summarised
 * in the module's report; the ones that matter for callers:
 * - An outline slide with no bullets is drawn as a section slide (`kind:'section'`).
 * - Bullets longer than `MAX_BULLET_CHARS` are split on 。；;！？!? and newlines
 *   before layout, and a warning is emitted per split bullet.
 * - Continuation slides are all titled `<title>（续）`; duplicates are intentional
 *   because `slideId` is the authoritative key and the panel shows slide order.
 */
import {
  CANVAS,
  LOGO_RESERVE,
  SAFE_AREA,
  estimateLines,
  innerBox,
  resolveColor,
} from './schema.js'
import { makeDefaultTheme } from './defaults.js'

/* ------------------------------------------------------------------ *
 * Public constants
 * ------------------------------------------------------------------ */

/**
 * Neutral default style. Every field of the style object is optional; these are
 * the values `composeDeck` falls back to.
 * @type {Readonly<{
 *   density: 'airy' | 'standard' | 'dense',
 *   tone: 'formal' | 'bold' | 'minimal',
 *   bulletLayout: 'list' | 'cards',
 *   maxSlides: number,
 *   accentBar: boolean,
 *   coverVariant: 'left' | 'centered',
 *   closingSlide: boolean,
 *   colors: Record<string, string>,
 *   fonts: Record<string, string>,
 * }>}
 */
export const DEFAULT_STYLE = Object.freeze({
  density: 'standard',
  tone: 'formal',
  bulletLayout: 'list',
  maxSlides: 12,
  accentBar: true,
  coverVariant: 'left',
  closingSlide: true,
  colors: Object.freeze({}),
  fonts: Object.freeze({}),
})

/** Bullets longer than this are split before layout (design rule #6). */
export const MAX_BULLET_CHARS = 90

/* ------------------------------------------------------------------ *
 * Internal tables
 * ------------------------------------------------------------------ */

/**
 * Bullets per slide per density, plus the vertical gap between them. `list` and
 * `cards` have different caps because a card needs much more room than a line.
 * `gapMax` bounds how far leftover space may be spread into the gap: `airy` is
 * supposed to breathe, `dense` should stay packed.
 */
const DENSITY = Object.freeze({
  airy: Object.freeze({ list: 3, cards: 2, gap: 24, gapMax: 44 }),
  standard: Object.freeze({ list: 4, cards: 4, gap: 20, gapMax: 32 }),
  dense: Object.freeze({ list: 6, cards: 4, gap: 16, gapMax: 24 }),
})

/**
 * Tone table. A tone only changes *colour usage*, never geometry — so the layout
 * is identical across tones and tests can rely on it.
 * - formal: one primary accent bar, neutral surface cards.
 * - bold: primary accent bar, primary card ticks, primary section titles.
 * - minimal: no accent bar at all, textPrimary only (no tinted cards).
 */
const TONES = Object.freeze({
  formal: Object.freeze({
    coverBar: true,
    bar: true,
    sectionRule: 'primary',
    sectionTitle: 'textPrimary',
    cardFill: 'surface',
    cardTick: false,
    marker: true,
    sub: 'textSecondary',
  }),
  bold: Object.freeze({
    coverBar: true,
    bar: true,
    sectionRule: 'primary',
    sectionTitle: 'primary',
    cardFill: 'surface',
    cardTick: true,
    marker: true,
    sub: 'textSecondary',
  }),
  minimal: Object.freeze({
    coverBar: false,
    bar: false,
    sectionRule: 'textPrimary',
    sectionTitle: 'textPrimary',
    cardFill: null,
    cardTick: false,
    marker: false,
    sub: 'textPrimary',
  }),
})

/** Font sizes per role. Ranges follow the layout rules in the module header. */
const TYPE = Object.freeze({
  coverTitle: Object.freeze([64, 60, 56]),
  coverSubtitle: 22,
  slideTitle: Object.freeze([40, 38, 36]),
  sectionTitle: 52,
  body: Object.freeze([22, 20]),
  bodyEmergency: Object.freeze([18, 16]),
  card: Object.freeze([22, 20, 18]),
  closingTitle: 64,
  closingCaption: 20,
  closingCaptionText: '结束',
  closingTitleText: '谢谢',
})

/** Horizontal limits: titles stop before the logo reserve, body copy does not. */
const FRAME = Object.freeze({
  /** Right edge of the full safe area (1216). */
  right: CANVAS.width - SAFE_AREA.right,
  /** Right edge for text that may sit beside the logo (1050). */
  titleRight: CANVAS.width - SAFE_AREA.right - LOGO_RESERVE.width - LOGO_RESERVE.gap,
})

/** Hard-coded last-resort hex values, used only when the theme is unusable. */
const FALLBACK_COLORS = Object.freeze({
  background: '#FFFFFF',
  surface: '#F4F6F9',
  textPrimary: '#1F2937',
  textSecondary: '#6B7280',
  primary: '#2F5BD2',
  accent: '#F59E0B',
})
const PALETTE_KEYS = Object.freeze(Object.keys(FALLBACK_COLORS))

const CLOSING_LIKE = /谢谢|感谢|致谢|总结|回顾|答疑|问答|q\s*&\s*a|thanks?\b|end\b/i
const CJK = /[\u2e80-\u9fff\u3000-\u303f\uf900-\ufaff\uff00-\uffef]/
const DEFAULT_DECK_TITLE = '未命名演示'

/* ------------------------------------------------------------------ *
 * Stage 1: text -> outline
 * ------------------------------------------------------------------ */

/**
 * Parse pasted text / a Markdown body into a deck outline.
 *
 * Rules (see the task contract):
 * - The first `# ` heading is the deck title; a second `# ` becomes a slide.
 * - Every `## ` heading (or a setext `===` / `---` underline) starts a slide.
 * - `- `, `* `, `+ ` and `1. ` lines are bullets; leading indentation is ignored.
 * - Other non-empty lines are paragraphs. Consecutive lines are hard-wrapped
 *   fragments of ONE paragraph, joined CJK-aware, and become one bullet.
 * - `备注：…` / `Notes: …` becomes `slide.notes` (repeated lines are appended).
 * - A heading with no bullets becomes `{kind:'section'}`.
 * - A `> quote` line is ignored when it is the machine header BrandDeck emits
 *   (`> deck: … · template: …`), otherwise it is used as the deck subtitle.
 * - With no slide heading at all, paragraphs are chunked into slides of at most
 *   `paragraphBullets` (default 5) bullets, titled from the first sentence.
 * - Never throws. Empty/odd input -> `{title:'未命名演示', slides: []}`.
 *
 * @param {string} text
 * @param {{subtitle?: string, fallbackTitle?: string, paragraphBullets?: number, maxTitleChars?: number}} [options]
 * @returns {{title: string, subtitle?: string, slides: Array<{title: string, bullets: string[], notes?: string, kind?: 'section'}>}}
 */
export function parseSourceToOutline(text, options = {}) {
  const opts = isObject(options) ? options : {}
  const fallbackTitle =
    typeof opts.fallbackTitle === 'string' && opts.fallbackTitle.trim() !== ''
      ? opts.fallbackTitle.trim()
      : DEFAULT_DECK_TITLE
  const perSlide =
    Number.isFinite(opts.paragraphBullets) && opts.paragraphBullets >= 1
      ? Math.floor(opts.paragraphBullets)
      : 5
  const maxTitleChars =
    Number.isFinite(opts.maxTitleChars) && opts.maxTitleChars >= 4
      ? Math.floor(opts.maxTitleChars)
      : 24

  const source = typeof text === 'string' ? text.replace(/\r\n?/g, '\n') : ''
  const lines = source.split('\n')

  /** @type {string | null} */
  let deckTitle = null
  /** @type {string | null} */
  let subtitle = typeof opts.subtitle === 'string' && opts.subtitle.trim() !== '' ? opts.subtitle.trim() : null
  /** @type {Array<{title: string, bullets: string[], notes?: string, kind?: 'section'}>} */
  const slides = []
  /** @type {{title: string, bullets: string[], notes?: string} | null} */
  let current = null
  /** @type {string | null} */
  let paragraph = null
  /** @type {string[]} */
  const preamble = []

  /** Append a finished paragraph: to the current slide, or to the preamble. */
  const flushParagraph = () => {
    if (paragraph === null) return
    const value = paragraph.trim()
    paragraph = null
    if (value === '') return
    if (current !== null) current.bullets.push(value)
    else preamble.push(value)
  }

  /** @param {string} title */
  const startSlide = (title) => {
    flushParagraph()
    const slide = { title: title === '' ? '未命名页面' : title, bullets: /** @type {string[]} */ ([]) }
    slides.push(slide)
    current = slide
  }

  /** @param {string} note */
  const appendNotes = (note) => {
    if (current === null) return
    const value = note.trim()
    if (value === '') return
    current.notes = current.notes === undefined || current.notes === '' ? value : `${current.notes}\n${value}`
  }

  for (let index = 0; index < lines.length; index += 1) {
    const trimmed = lines[index].trim()

    if (trimmed === '') {
      flushParagraph()
      continue
    }

    const note = matchNotesLine(trimmed)
    if (note !== null) {
      flushParagraph()
      appendNotes(note)
      continue
    }

    if (trimmed.startsWith('>')) {
      flushParagraph()
      const quote = trimmed.replace(/^>\s?/, '').trim()
      // `> deck: …` is BrandDeck's own machine header (see outline.js); a user
      // describing a deck should not get it back as a cover subtitle.
      if (quote !== '' && subtitle === null && !/^(deck|template|模板)\s*[:：]/i.test(quote)) {
        subtitle = cleanInline(quote)
      }
      continue
    }

    const nextLine = index + 1 < lines.length ? lines[index + 1].trim() : ''
    const isPlain = !isBulletLine(trimmed) && !isAtxHeading(trimmed)
    if (isPlain && isSetextUnderline(nextLine, '=')){
      flushParagraph()
      const headingText = cleanInline(stripTrailingId(trimmed))
      if (deckTitle === null && slides.length === 0 && headingText !== '') {
        deckTitle = headingText
      } else if (headingText !== '') {
        startSlide(headingText)
      }
      index += 1
      continue
    }
    if (isPlain && isSetextUnderline(nextLine, '-')) {
      flushParagraph()
      startSlide(cleanInline(stripTrailingId(trimmed)))
      index += 1
      continue
    }

    const atx = /^(#{1,6})\s*(.*)$/.exec(trimmed)
    if (atx !== null) {
      flushParagraph()
      const level = atx[1].length
      const headingText = cleanInline(stripTrailingId(atx[2]))
      if (level === 1 && deckTitle === null && slides.length === 0 && headingText !== '') {
        deckTitle = headingText
        continue
      }
      if (level <= 2) {
        startSlide(headingText)
        continue
      }
      // `###` and deeper is content, not structure: keep the text.
      if (headingText !== '') {
        if (current !== null) current.bullets.push(headingText)
        else preamble.push(headingText)
      }
      continue
    }

    // A horizontal rule is a separator, not content: `---` / `***` / `___`.
    // (A `---` directly under text was already consumed as a setext heading above.)
    if (/^(?:-{3,}|\*{3,}|_{3,})$/.test(trimmed)) {
      flushParagraph()
      continue
    }

    const bullet = matchBulletLine(trimmed)
    if (bullet !== null) {
      flushParagraph()
      const value = cleanInline(bullet)
      if (value !== '') {
        if (current !== null) current.bullets.push(value)
        else preamble.push(value)
      }
      continue
    }

    // A plain line: hard-wrapped text joins the paragraph being built.
    paragraph = paragraph === null ? trimmed : joinWrapped(paragraph, trimmed)
  }
  flushParagraph()

  // No `##` heading anywhere: chunk the collected paragraphs into slides.
  if (slides.length === 0 && preamble.length > 0) {
    const derived = deriveTitle(preamble[0], maxTitleChars)
    const chunks = chunk(preamble, perSlide)
    chunks.forEach((bullets, chunkIndex) => {
      slides.push({
        title: chunkIndex === 0 ? derived : `${derived}（${chunkIndex + 1}）`,
        bullets,
      })
    })
  } else if (preamble.length > 0) {
    if (deckTitle !== null && deckTitle !== '') {
      // A `#` title already names the deck, so text before the first `##` is framing
      // copy, not a page. Turning it into a lead slide invents a heading out of the
      // first bullet (models like to write a one-line thesis there) and produces a
      // page whose only content is that thesis. Use it as the cover subtitle instead.
      // This function is pure and returns no warnings channel, so the mapping is
      // simply the sensible one: thesis line → subtitle, remaining framing dropped.
      if (subtitle === null) subtitle = preamble[0]
    } else {
      // No deck title at all: this is the "plain prose" case, where the first
      // paragraph legitimately becomes the title and the rest chunk into slides.
      const derived = deriveTitle(preamble[0], maxTitleChars)
      const lead = chunk(preamble, perSlide).map((bullets, chunkIndex) => ({
        title: chunkIndex === 0 ? derived : `${derived}（${chunkIndex + 1}）`,
        bullets,
      }))
      slides.unshift(...lead)
    }
  }

  const title =
    deckTitle === null || deckTitle === ''
      ? preamble.length > 0
        ? deriveTitle(preamble[0], maxTitleChars)
        : fallbackTitle
      : deckTitle

  /** @type {Array<{title: string, bullets: string[], notes?: string, kind?: 'section'}>} */
  const normalized = slides.map((slide) => {
    /** @type {{title: string, bullets: string[], notes?: string, kind?: 'section'}} */
    const out = { title: slide.title === '' ? '未命名页面' : slide.title, bullets: slide.bullets }
    if (slide.notes !== undefined && slide.notes !== '') out.notes = slide.notes
    if (slide.bullets.length === 0) out.kind = 'section'
    return out
  })

  /** @type {{title: string, subtitle?: string, slides: typeof normalized}} */
  const result = { title }
  if (subtitle !== null && subtitle !== '') result.subtitle = subtitle
  result.slides = normalized
  return result
}

/* ------------------------------------------------------------------ *
 * Stage 2: outline + theme + style -> deck
 * ------------------------------------------------------------------ */

/**
 * Compose a schema-valid CardDeck from an outline.
 *
 * @param {{
 *   outline: ReturnType<typeof parseSourceToOutline> | string,
 *   theme?: import('./schema.js').ThemeTokens,
 *   style?: Partial<typeof DEFAULT_STYLE>,
 *   deckId?: string,
 *   maxSlides?: number,
 * }} input
 * @returns {{deck: import('./schema.js').CardDeck, warnings: string[]}}
 */
export function composeDeck(input = {}) {
  const request = isObject(input) ? input : {}
  /** @type {string[]} */
  const warnings = []

  const theme = normalizeTheme(request.theme)
  const style = normalizeStyle(request.style, request.maxSlides)
  const box = innerBox(theme)
  const palette = resolvePalette(theme, style.colors)
  const outline = normalizeOutline(request.outline)
  const toneSpec = TONES[style.tone]

  const title = clampChars(oneLine(outline.title), 200) || DEFAULT_DECK_TITLE
  if (oneLine(outline.title).length > 200) {
    warnings.push('封面 · 标题超过 200 字，已截断')
  }

  /** @type {LocalContext} */
  const ctx = { theme, box, palette, style, toneSpec, warnings }

  const content = buildContentSlides(outline, style, ctx, warnings)
  const wantsClosing = style.closingSlide && style.maxSlides >= 2
  const allowed = Math.max(0, style.maxSlides - 1 - (wantsClosing ? 1 : 0))
  if (content.length > allowed) {
    warnings.push(`已按上限截断 ${content.length - allowed} 页`)
    content.length = allowed
  }

  /** @type {Array<import('./schema.js').CardSlide>} */
  const slides = []
  let counter = 0
  const nextSlideId = () => `s${(counter += 1)}`

  slides.push(buildCoverSlide(nextSlideId(), ctx, title, outline.subtitle))

  for (const item of content) {
    const slideId = nextSlideId()
    if (item.kind === 'section') {
      slides.push(buildSectionSlide(slideId, ctx, item))
    } else {
      slides.push(buildBulletSlide(slideId, ctx, item))
    }
  }

  const lastTitle = content.length > 0 ? content[content.length - 1].title : ''
  if (wantsClosing && !CLOSING_LIKE.test(lastTitle)) {
    slides.push(buildClosingSlide(nextSlideId(), ctx))
  }

  /** @type {import('./schema.js').CardDeck} */
  const deck = {
    deckId: resolveDeckId(request.deckId, title),
    title,
    templateId: theme.templateId,
    slides,
  }

  // Provenance for the caller. Fonts are deliberately NOT copied here: a style's
  // font overrides take effect through the *theme* (`applyStyleToTheme` in
  // styles.js) and the compiler resolves a face per text role, so the exported
  // .pptx already carries them. `TextElementSchema.style` is strict and has no
  // `fontFace`, and `CardDeckSchema` is non-strict, so a copy here would be
  // stripped on save and would only mislead a reader of the stored JSON.
  deck.metadata = {
    generator: 'branddeck/generate',
    style: {
      density: style.density,
      tone: style.tone,
      bulletLayout: style.bulletLayout,
      maxSlides: style.maxSlides,
      accentBar: style.accentBar,
      coverVariant: style.coverVariant,
      closingSlide: style.closingSlide,
    },
    /** Fonts actually in force, already applied to the theme by the caller. */
    fonts: { title: theme.fonts.title, body: theme.fonts.body },
  }

  return { deck, warnings: [...new Set(warnings)] }
}

/* ------------------------------------------------------------------ *
 * Outline normalization + continuation slides
 * ------------------------------------------------------------------ */

/**
 * Coerce arbitrary outline input into the canonical shape. A raw string is parsed
 * with {@link parseSourceToOutline} so callers can pass "just the pasted text".
 *
 * @param {any} outline
 * @param {ReturnType<typeof normalizeStyle>} style
 */
function normalizeOutline(outline) {
  const source = typeof outline === 'string' ? parseSourceToOutline(outline) : outline
  const src = isObject(source) ? source : {}
  /** @type {Array<{title: string, bullets: string[], notes?: string, kind: 'section' | 'bullets'}>} */
  const slides = []
  for (const raw of Array.isArray(src.slides) ? src.slides : []) {
    if (!isObject(raw)) continue
    const slideTitle = typeof raw.title === 'string' ? oneLine(raw.title) : ''
    const bullets = (Array.isArray(raw.bullets) ? raw.bullets : [])
      .filter((bullet) => typeof bullet === 'string')
      .map((bullet) => oneLine(bullet))
      .filter((bullet) => bullet !== '')
    const isSection = raw.kind === 'section' || bullets.length === 0
    /** @type {{title: string, bullets: string[], notes?: string, kind: 'section' | 'bullets'}} */
    const item = {
      title: slideTitle === '' ? '未命名页面' : slideTitle,
      bullets: isSection ? [] : bullets,
      kind: isSection ? 'section' : 'bullets',
    }
    if (typeof raw.notes === 'string' && raw.notes.trim() !== '') item.notes = raw.notes.trim()
    slides.push(item)
  }
  return {
    title: typeof src.title === 'string' ? oneLine(src.title) : '',
    subtitle: typeof src.subtitle === 'string' ? oneLine(src.subtitle) : '',
    slides,
  }
}

/**
 * Expand outline slides into the flat list of content slides the layout pass
 * draws. Handles long-bullet splitting and both kinds of spilling:
 * - count: the density cap (rule 4) spills a title's bullets onto `<title>（续）`;
 * - height: `cards` rows are sized to their content, so a slide whose rows would
 *   leave the safe box also spills (see {@link planCardChunks}).
 *
 * @param {{slides: Array<{title: string, bullets: string[], notes?: string, kind: string}>}} outline
 * @param {ReturnType<typeof normalizeStyle>} style
 * @param {LocalContext} ctx
 * @param {string[]} warnings
 * @returns {Array<{kind: 'section' | 'bullets', title: string, bullets: string[], notes?: string}>}
 */
function buildContentSlides(outline, style, ctx, warnings) {
  /** @type {Array<{kind: 'section' | 'bullets', title: string, bullets: string[], notes?: string}>} */
  const out = []
  const cap = DENSITY[style.density][style.bulletLayout]

  outline.slides.forEach((slide, index) => {
    const label = `第 ${index + 1} 页「${clampChars(slide.title, 16)}」`
    if (slide.kind === 'section' || slide.bullets.length === 0) {
      out.push({ kind: 'section', title: slide.title, bullets: [] })
      if (slide.notes !== undefined) out[out.length - 1].notes = slide.notes
      return
    }

    /** @type {string[]} */
    const pieces = []
    for (const bullet of slide.bullets) {
      const parts = splitLongBullet(bullet, MAX_BULLET_CHARS)
      if (parts.length > 1) {
        warnings.push(`${label} · 要点超过 ${MAX_BULLET_CHARS} 字，已拆分为 ${parts.length} 条`)
      }
      pieces.push(...parts)
    }

    const chunks =
      style.bulletLayout === 'cards'
        ? planCardChunks(pieces, cap, ctx, slide.title)
        : chunk(pieces, cap)
    const countChunks = Math.ceil(pieces.length / Math.max(1, cap))
    if (chunks.length > countChunks) {
      warnings.push(`${label} · 卡片按内容高度排布，已溢出到 ${chunks.length - countChunks} 页「（续）」`)
    }

    chunks.forEach((bullets, chunkIndex) => {
      /** @type {{kind: 'bullets', title: string, bullets: string[], notes?: string}} */
      const item = {
        kind: 'bullets',
        title: chunkIndex === 0 ? slide.title : `${slide.title}（续）`,
        bullets,
      }
      if (chunkIndex === 0 && slide.notes !== undefined) item.notes = slide.notes
      out.push(item)
    })
  })

  return out
}

/* ------------------------------------------------------------------ *
 * Slide builders
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} LocalContext
 * @property {import('./schema.js').ThemeTokens} theme
 * @property {ReturnType<typeof innerBox>} box
 * @property {Record<string, string>} palette resolved hex per token
 * @property {ReturnType<typeof normalizeStyle>} style
 * @property {typeof TONES['formal']} toneSpec
 * @property {string[]} warnings
 */

/**
 * Cover slide: left or centered title, optional subtitle, accent bar outside the
 * logo reserve. Nothing overlaps the 150x48 logo corner.
 *
 * @param {string} slideId
 * @param {LocalContext} ctx
 * @param {string} deckTitle
 * @param {string} subtitle
 * @returns {import('./schema.js').CardSlide}
 */
function buildCoverSlide(slideId, ctx, deckTitle, subtitle) {
  const { box, palette, style, toneSpec, warnings } = ctx
  const builder = makeSlideBuilder(slideId)
  const centered = style.coverVariant === 'centered'
  const withBar = style.accentBar && toneSpec.coverBar
  const textX = centered ? box.x : box.x + (withBar ? 34 : 0)
  const width = centered ? box.w : FRAME.titleRight - textX

  const size = chooseSize(TYPE.coverTitle, deckTitle, width, 3)
  const fitted = clampToLines(deckTitle, width, size, 3)
  if (fitted.truncated) warnings.push('封面 · 标题过长，已截断到 3 行以内')
  const titleH = textHeight(fitted.text, width, size)

  const subText = subtitle === undefined ? '' : oneLine(subtitle)
  const subSize = TYPE.coverSubtitle
  const fittedSub = subText === '' ? null : clampToLines(subText, width, subSize, 2)
  const subH = fittedSub === null ? 0 : textHeight(fittedSub.text, width, subSize)

  const gapAfterTitle = subH > 0 ? 22 : 0
  const gapBeforeRule = withBar ? 34 : 0
  const ruleH = withBar ? 2 : 0
  const total = titleH + gapAfterTitle + subH + gapBeforeRule + ruleH
  // Start at least 72px below the safe top: that keeps the centered variant's
  // full-width title (x..1216) below the logo reserve (y 64..112).
  const y = Math.max(box.y + 72, Math.round((CANVAS.height - total) / 2) - 24)

  if (withBar && !centered) {
    builder.shape('bar', 'rect', { x: box.x, y: y - 8, w: 8, h: titleH + gapAfterTitle + subH + 16 }, {
      fill: palette.primary,
    })
  }
  if (withBar && centered) {
    builder.shape('bar', 'rect', {
      x: Math.round(CANVAS.width / 2) - 32,
      y: y - 30,
      w: 64,
      h: 6,
    }, { fill: palette.primary })
  }

  builder.text('title', fitted.text, { x: textX, y, w: width, h: titleH }, textStyle(size, palette.textPrimary, {
    bold: true,
    align: centered ? 'center' : 'left',
  }))

  let cursor = y + titleH
  if (fittedSub !== null) {
    cursor += gapAfterTitle
    builder.text('subtitle', fittedSub.text, { x: textX, y: cursor, w: width, h: subH }, textStyle(
      subSize,
      palette[toneSpec.sub],
      { align: centered ? 'center' : 'left' },
    ))
    cursor += subH
  }

  if (withBar) {
    const ruleY = cursor + gapBeforeRule
    if (centered) {
      builder.shape('rule', 'line', {
        x: Math.round(CANVAS.width / 2) - 140,
        y: ruleY,
        w: 280,
        h: 2,
      }, { stroke: palette.surface, strokeWidth: 2 })
    } else {
      builder.shape('rule', 'line', { x: textX, y: ruleY, w: width, h: 2 }, {
        stroke: palette.surface,
        strokeWidth: 2,
      })
    }
  }

  return slideShell(slideId, builder, palette, fitted.text)
}

/**
 * Section slide: one big centered title plus a short accent rule, no bullets.
 *
 * @param {string} slideId
 * @param {LocalContext} ctx
 * @param {{title: string, notes?: string}} item
 * @returns {import('./schema.js').CardSlide}
 */
function buildSectionSlide(slideId, ctx, item) {
  const { box, palette, style, toneSpec, warnings } = ctx
  const builder = makeSlideBuilder(slideId)
  const width = box.w
  const size = TYPE.sectionTitle
  const fitted = clampToLines(item.title, width, size, 2)
  if (fitted.truncated) warnings.push(`「${clampChars(item.title, 16)}」· 章节标题过长，已截断`)
  const titleH = textHeight(fitted.text, width, size)
  const ruleH = style.tone === 'bold' ? 8 : 6
  const ruleW = style.tone === 'bold' ? 160 : 120
  const total = titleH + 32 + ruleH
  const y = Math.max(box.y + 80, Math.round((CANVAS.height - total) / 2) - 12)

  builder.text('title', fitted.text, { x: box.x, y, w: width, h: titleH }, textStyle(
    size,
    palette[toneSpec.sectionTitle],
    { bold: true, align: 'center', valign: 'middle' },
  ))
  builder.shape('rule', 'rect', {
    x: Math.round((CANVAS.width - ruleW) / 2),
    y: y + titleH + 32,
    w: ruleW,
    h: ruleH,
  }, { fill: palette[toneSpec.sectionRule] })

  const slide = slideShell(slideId, builder, palette, fitted.text, item.notes)
  return slide
}

/**
 * Bullet slide: title + optional accent bar, then the bullets as a vertical list
 * or as 2-column cards. Both layouts place every bullet once, inside `innerBox`.
 *
 * @param {string} slideId
 * @param {LocalContext} ctx
 * @param {{title: string, bullets: string[], notes?: string}} item
 * @returns {import('./schema.js').CardSlide}
 */
function buildBulletSlide(slideId, ctx, item) {
  const { box, palette, style, toneSpec, warnings } = ctx
  const builder = makeSlideBuilder(slideId)
  const width = FRAME.titleRight - box.x
  const size = chooseSize(TYPE.slideTitle, item.title, width, 2)
  const fitted = clampToLines(item.title, width, size, 2)
  const label = `「${clampChars(fitted.text, 16)}」`
  if (fitted.truncated) warnings.push(`${label} · 页标题过长，已截断到 2 行以内`)
  const titleH = textHeight(fitted.text, width, size)

  builder.text('title', fitted.text, { x: box.x, y: box.y, w: width, h: titleH }, textStyle(
    size,
    palette.textPrimary,
    { bold: true },
  ))

  let cursor = box.y + titleH
  if (style.accentBar && toneSpec.bar) {
    const barH = style.tone === 'bold' ? 8 : 6
    const barW = style.tone === 'bold' ? 120 : 96
    builder.shape('bar', 'rect', { x: box.x, y: cursor + 14, w: barW, h: barH }, { fill: palette.primary })
    cursor += 14 + barH
  }
  const contentTop = cursor + 34

  if (style.bulletLayout === 'cards') {
    placeCards(builder, ctx, item.bullets, contentTop, label)
  } else {
    placeList(builder, ctx, item.bullets, contentTop, label)
  }

  const slide = slideShell(slideId, builder, palette, fitted.text, item.notes)
  return slide
}

/**
 * Closing slide: centered 谢谢 + optional 结束 caption and accent rule.
 *
 * @param {string} slideId
 * @param {LocalContext} ctx
 * @returns {import('./schema.js').CardSlide}
 */
function buildClosingSlide(slideId, ctx) {
  const { box, palette, style, toneSpec } = ctx
  const builder = makeSlideBuilder(slideId)
  const titleH = textHeight(TYPE.closingTitleText, box.w, TYPE.closingTitle)
  const captionH = textHeight(TYPE.closingCaptionText, box.w, TYPE.closingCaption)
  const withBar = style.accentBar && toneSpec.bar
  const ruleH = withBar ? 6 : 0
  const total = titleH + 28 + ruleH + (ruleH > 0 ? 24 : 0) + captionH
  const y = Math.round((CANVAS.height - total) / 2)

  builder.text('title', TYPE.closingTitleText, { x: box.x, y, w: box.w, h: titleH }, textStyle(
    TYPE.closingTitle,
    palette.textPrimary,
    { bold: true, align: 'center' },
  ))
  let cursor = y + titleH + 28
  if (withBar) {
    builder.shape('rule', 'rect', {
      x: Math.round(CANVAS.width / 2) - 32,
      y: cursor,
      w: 64,
      h: ruleH,
    }, { fill: palette.primary })
    cursor += ruleH + 24
  }
  builder.text('caption', TYPE.closingCaptionText, { x: box.x, y: cursor, w: box.w, h: captionH }, textStyle(
    TYPE.closingCaption,
    palette.textSecondary,
    { align: 'center' },
  ))

  return slideShell(slideId, builder, palette, TYPE.closingTitleText)
}

/**
 * Common slide envelope: `slideId` (stable by order), the title the panel/outline
 * projection shows, the page background and the element list.
 *
 * @param {string} slideId
 * @param {ReturnType<typeof makeSlideBuilder>} builder
 * @param {Record<string, string>} palette
 * @param {string} title
 * @param {string} [notes]
 * @returns {import('./schema.js').CardSlide}
 */
function slideShell(slideId, builder, palette, title, notes) {
  /** @type {import('./schema.js').CardSlide} */
  const slide = {
    slideId,
    title,
    background: palette.background,
    elements: builder.elements,
  }
  if (notes !== undefined && notes !== '') slide.speakerNotes = notes
  return slide
}

/* ------------------------------------------------------------------ *
 * Bullet layouts
 * ------------------------------------------------------------------ */

/**
 * Vertical list: one text box per bullet, a small square marker per bullet when
 * the tone allows it. Font size is 22 when everything fits at the density gap,
 * else 20; the 18/16 step exists only to keep pathological input inside the safe
 * area (documented judgement call: containment beats the 20px minimum).
 *
 * @param {ReturnType<typeof makeSlideBuilder>} builder
 * @param {LocalContext} ctx
 * @param {string[]} bullets
 * @param {number} top
 * @param {string} label
 */
function placeList(builder, ctx, bullets, top, label) {
  const { box, palette, style, toneSpec, warnings } = ctx
  const marker = toneSpec.marker
  const textX = marker ? box.x + 30 : box.x
  const width = FRAME.right - textX
  const available = box.y + box.h - top
  const baseGap = DENSITY[style.density].gap

  let fontSize = TYPE.body[0]
  let gap = baseGap
  let heights = bullets.map((bullet) => textHeight(bullet, width, fontSize))
  let total = totalHeight(heights, gap)

  if (total > available) {
    fontSize = TYPE.body[1]
    gap = Math.min(baseGap, 16)
    heights = bullets.map((bullet) => textHeight(bullet, width, fontSize))
    total = totalHeight(heights, gap)
  }
  if (total > available) {
    for (const size of TYPE.bodyEmergency) {
      fontSize = size
      gap = Math.min(baseGap, 10)
      heights = bullets.map((bullet) => textHeight(bullet, width, size))
      total = totalHeight(heights, gap)
      if (total <= available) break
    }
    warnings.push(`${label} · 要点过多，字号已缩小到 ${fontSize}px`)
  }
  if (total > available) {
    // Unreachable for bullets that went through `splitLongBullet`; kept so a
    // hand-built outline can never push text outside the safe area.
    const scale = available / total
    heights = heights.map((height) => Math.max(8, Math.floor(height * scale)))
    gap = Math.max(4, Math.floor(gap * scale))
    warnings.push(`${label} · 要点过多，已压缩行高以适应页面`)
  } else if (bullets.length > 1) {
    // Leftover vertical space is spread into the gap (bounded by `gapMax`) so an
    // `airy` 3-bullet slide fills the page instead of hugging the title.
    const spec = DENSITY[style.density]
    const slack = available - total
    gap = Math.min(spec.gapMax, gap + Math.floor(slack / (bullets.length - 1)))
  }

  let y = top
  bullets.forEach((bullet, index) => {
    const height = heights[index]
    if (marker) {
      builder.shape('marker', 'roundRect', {
        x: box.x,
        y: Math.round(y + (Math.ceil(fontSize * 1.35) / 2) - 4),
        w: 8,
        h: 8,
      }, { fill: palette.primary, radius: 4 })
    }
    builder.text('body', bullet, { x: textX, y, w: width, h: height }, textStyle(
      fontSize,
      palette.textPrimary,
    ))
    const lines = estimateLines(bullet, width, fontSize).lines
    if (lines > 4) {
      warnings.push(`${label} · 要点过长（约 ${lines} 行），建议精简为 ${MAX_BULLET_CHARS} 字以内`)
    }
    y += height + gap
  })
}

/**
 * Card geometry. A card is sized to its content: `textTop + textH + padBottom`,
 * where `textTop` leaves room for the `bold` tone's accent tick. Cards in the same
 * visual row share one height (the tallest text of the row) and one text-box
 * height, so the grid stays even and no card carries a large empty band.
 */
const CARD = Object.freeze({
  /** Horizontal gap between the two columns. */
  gapX: 24,
  /** Vertical gap between card rows (design rule: >= 16px). */
  gapY: 24,
  /** Upper bound when leftover height is spread into the row gap. */
  gapYMax: 48,
  /** Left/right padding inside a card. */
  padX: 28,
  /** Bottom padding inside a card (breathing room below the text). */
  padBottom: 32,
  /** Space between a card's top and its text, with the accent tick. */
  tickTop: 24,
  tickH: 6,
  tickGap: 22,
  /** Accent tick width. */
  tickW: 48,
  /** Minimum text height so an empty-ish card still renders as a card. */
  minTextH: 30,
})

/** Text-column width inside one card. */
function cardTextWidth(box) {
  return Math.round((box.w - CARD.gapX) / 2) - CARD.padX * 2
}

/**
 * @param {boolean} tick
 */
function cardTextTop(tick) {
  return tick ? CARD.tickTop + CARD.tickH + CARD.tickGap : CARD.padX
}

/**
 * Measure content-sized card rows for one slide: pair the bullets into the two
 * columns, take the tallest text per row, and derive the row's card height.
 *
 * @param {string[]} bullets
 * @param {number} textW
 * @param {number} fontSize
 * @param {boolean} tick
 * @returns {{rows: Array<{cardH: number, textH: number, items: Array<{bullet: string, textH: number}>}>, total: number}}
 */
function measureCardRows(bullets, textW, fontSize, tick) {
  /** @type {Array<{cardH: number, textH: number, items: Array<{bullet: string, textH: number}>}>} */
  const rows = []
  const textTop = cardTextTop(tick)
  for (let index = 0; index < bullets.length; index += 2) {
    const items = bullets.slice(index, index + 2).map((bullet) => ({
      bullet,
      textH: textHeight(bullet, textW, fontSize),
    }))
    const textH = Math.max(CARD.minTextH, ...items.map((item) => item.textH))
    rows.push({ cardH: textTop + textH + CARD.padBottom, textH, items })
  }
  const total = rows.reduce((sum, row) => sum + row.cardH, 0) + CARD.gapY * Math.max(0, rows.length - 1)
  return { rows, total }
}

/**
 * Top edge of the bullet area on a bullet slide, i.e. below the title and the
 * optional accent bar. Shared by the card planner and the actual layout so the
 * height budget used for spilling is the real one.
 *
 * @param {LocalContext} ctx
 * @param {string} title
 */
function bulletContentTop(ctx, title) {
  const { box, style, toneSpec } = ctx
  const width = FRAME.titleRight - box.x
  const size = chooseSize(TYPE.slideTitle, title, width, 2)
  const fitted = clampToLines(title, width, size, 2)
  let cursor = box.y + textHeight(fitted.text, width, size)
  if (style.accentBar && toneSpec.bar) cursor += 14 + (style.tone === 'bold' ? 8 : 6)
  return cursor + 34
}

/**
 * Split a title's bullets into card slides that each fit the safe box, honouring
 * the density cap as a count cap and the measured row heights as a height cap.
 *
 * @param {string[]} bullets
 * @param {number} cap
 * @param {LocalContext} ctx
 * @param {string} title
 * @returns {string[][]}
 */
function planCardChunks(bullets, cap, ctx, title) {
  /** @type {string[][]} */
  const chunks = []
  /** @type {string[]} */
  let current = []
  const textW = cardTextWidth(ctx.box)
  const tick = ctx.toneSpec.cardTick
  const bottom = ctx.box.y + ctx.box.h

  /** @param {number} count number of already-built chunks */
  const topFor = (count) => bulletContentTop(ctx, count === 0 ? title : `${title}（续）`)

  for (const bullet of bullets) {
    if (current.length > 0 && current.length >= cap) {
      chunks.push(current)
      current = []
    }
    if (current.length > 0) {
      const top = topFor(chunks.length)
      const total = measureCardRows([...current, bullet], textW, TYPE.card[0], tick).total
      if (top + total > bottom) {
        chunks.push(current)
        current = []
      }
    }
    current.push(bullet)
  }
  if (current.length > 0) chunks.push(current)
  return chunks.length === 0 ? [[]] : chunks
}

/**
 * 2-column roundRect cards, each row sized to its own content. `airy` allows 2 per
 * slide, `standard`/`dense` 4; a card slide therefore never holds more than 4
 * bullets (height-driven spill is planned earlier by {@link planCardChunks}).
 *
 * A lone card (odd bullet count) keeps the column width instead of stretching to
 * the full 1128px: a full-width card next to three half-width ones reads as an
 * error, not as emphasis.
 *
 * @param {ReturnType<typeof makeSlideBuilder>} builder
 * @param {LocalContext} ctx
 * @param {string[]} bullets
 * @param {number} top
 * @param {string} label
 */
function placeCards(builder, ctx, bullets, top, label) {
  const { box, palette, toneSpec, warnings } = ctx
  const cardW = Math.round((box.w - CARD.gapX) / 2)
  const textW = cardTextWidth(box)
  const tick = toneSpec.cardTick
  const textTop = cardTextTop(tick)
  const bottom = box.y + box.h

  // Largest body size whose measured rows fit the safe box; the planner sized the
  // chunk at TYPE.card[0], so this only ever steps down for hand-built outlines.
  let plan = null
  let fontSize = TYPE.card[TYPE.card.length - 1]
  for (const size of TYPE.card) {
    const candidate = measureCardRows(bullets, textW, size, tick)
    if (top + candidate.total <= bottom) {
      plan = candidate
      fontSize = size
      break
    }
  }
  if (plan === null) {
    fontSize = TYPE.card[TYPE.card.length - 1]
    plan = measureCardRows(bullets, textW, fontSize, tick)
    // Last resort (unreachable for planned chunks): keep the grid on the page and
    // report that the copy had to be squeezed.
    const scale = (bottom - top) / plan.total
    if (scale < 1) {
      plan = {
        rows: plan.rows.map((row) => {
          const textH = Math.max(8, Math.floor(row.textH * scale))
          return { cardH: textTop + textH + CARD.padBottom, textH, items: row.items }
        }),
        total: 0,
      }
      warnings.push(`${label} · 卡片内容过高，已压缩卡片高度以适应页面`)
    }
  }

  let rowTop = top
  // Content-sized rows leave slack on a sparsely filled page; spread it into the
  // row gap (bounded, never past the safe box) so the grid does not bunch at the
  // top. A single row keeps the base gap and stays title-anchored.
  const rowsHeight = plan.rows.reduce((sum, row) => sum + row.cardH, 0)
  const slack = Math.max(0, bottom - top - rowsHeight)
  const rowGap =
    plan.rows.length > 1
      ? Math.min(CARD.gapYMax, CARD.gapY + Math.floor(slack / (plan.rows.length - 1)))
      : CARD.gapY

  plan.rows.forEach((row) => {
    row.items.forEach((item, column) => {
      const x = box.x + column * (cardW + CARD.gapX)
      if (toneSpec.cardFill !== null) {
        builder.shape('card', 'roundRect', { x, y: rowTop, w: cardW, h: row.cardH }, {
          fill: palette[toneSpec.cardFill],
          radius: 18,
        })
      }
      if (tick) {
        builder.shape('cardtick', 'rect', {
          x: x + CARD.padX,
          y: rowTop + CARD.tickTop,
          w: CARD.tickW,
          h: CARD.tickH,
        }, { fill: palette.primary })
      }
      builder.text('body', item.bullet, {
        x: x + CARD.padX,
        y: rowTop + textTop,
        w: textW,
        h: row.textH,
      }, textStyle(fontSize, palette.textPrimary))
    })
    rowTop += row.cardH + rowGap
  })
}

/* ------------------------------------------------------------------ *
 * Style / theme normalization
 * ------------------------------------------------------------------ */

/**
 * @param {any} rawStyle
 * @param {number} [requestMaxSlides]
 */
function normalizeStyle(rawStyle, requestMaxSlides) {
  const src = isObject(rawStyle) ? rawStyle : {}
  let maxSlides = Number.isFinite(src.maxSlides) ? Math.floor(src.maxSlides) : DEFAULT_STYLE.maxSlides
  if (Number.isFinite(requestMaxSlides)) maxSlides = Math.floor(requestMaxSlides)
  maxSlides = clamp(maxSlides, 1, 60)

  /** @type {Record<string, string>} */
  const colors = {}
  if (isObject(src.colors)) {
    for (const key of PALETTE_KEYS) {
      const value = src.colors[key]
      if (typeof value === 'string' && value.trim() !== '') colors[key] = value.trim()
    }
  }
  /** @type {{title?: string, body?: string}} */
  const fonts = {}
  if (isObject(src.fonts)) {
    if (typeof src.fonts.title === 'string' && src.fonts.title.trim() !== '') fonts.title = src.fonts.title.trim()
    if (typeof src.fonts.body === 'string' && src.fonts.body.trim() !== '') fonts.body = src.fonts.body.trim()
  }

  return {
    density: oneOf(src.density, Object.keys(DENSITY), DEFAULT_STYLE.density),
    tone: oneOf(src.tone, Object.keys(TONES), DEFAULT_STYLE.tone),
    bulletLayout: oneOf(src.bulletLayout, ['list', 'cards'], DEFAULT_STYLE.bulletLayout),
    maxSlides,
    accentBar: src.accentBar !== false,
    coverVariant: oneOf(src.coverVariant, ['left', 'centered'], DEFAULT_STYLE.coverVariant),
    closingSlide: src.closingSlide !== false,
    colors,
    fonts,
  }
}

/**
 * Fill in a usable theme: missing/invalid tokens fall back to the built-in theme
 * so `innerBox` and `resolveColor` always have something to work with.
 *
 * @param {any} rawTheme
 * @returns {import('./schema.js').ThemeTokens}
 */
function normalizeTheme(rawTheme) {
  const base = makeDefaultTheme()
  if (!isObject(rawTheme)) return base

  const colors = isObject(rawTheme.colors) ? rawTheme.colors : {}
  /** @type {Record<string, string>} */
  const mergedColors = {}
  for (const key of PALETTE_KEYS) {
    const source = colors[key]
    mergedColors[key] = typeof source === 'string' && source.trim() !== '' ? source : base.colors[key]
  }
  mergedColors.chartPalette =
    Array.isArray(colors.chartPalette) && colors.chartPalette.every((value) => typeof value === 'string') && colors.chartPalette.length > 0
      ? colors.chartPalette.slice(0, 8)
      : base.colors.chartPalette

  const fonts = isObject(rawTheme.fonts) ? rawTheme.fonts : {}
  const brandAssets = isObject(rawTheme.brandAssets) ? rawTheme.brandAssets : {}
  /** @type {{logoPath?: string, logoWidthPx?: number, logoHeightPx?: number, icons?: Record<string, string>}} */
  const mergedAssets = { icons: isObject(brandAssets.icons) ? { ...brandAssets.icons } : {} }
  if (typeof brandAssets.logoPath === 'string' && brandAssets.logoPath !== '') mergedAssets.logoPath = brandAssets.logoPath
  if (Number.isInteger(brandAssets.logoWidthPx)) mergedAssets.logoWidthPx = brandAssets.logoWidthPx
  if (Number.isInteger(brandAssets.logoHeightPx)) mergedAssets.logoHeightPx = brandAssets.logoHeightPx

  return {
    templateId:
      typeof rawTheme.templateId === 'string' && rawTheme.templateId !== ''
        ? clampChars(rawTheme.templateId, 64)
        : base.templateId,
    templateName:
      typeof rawTheme.templateName === 'string' && rawTheme.templateName !== ''
        ? rawTheme.templateName
        : base.templateName,
    colors: /** @type {any} */ (mergedColors),
    fonts: {
      title: typeof fonts.title === 'string' && fonts.title !== '' ? fonts.title : base.fonts.title,
      body: typeof fonts.body === 'string' && fonts.body !== '' ? fonts.body : base.fonts.body,
      latin: typeof fonts.latin === 'string' && fonts.latin !== '' ? fonts.latin : base.fonts.latin,
    },
    brandAssets: /** @type {any} */ (mergedAssets),
  }
}

/**
 * Resolve the six palette slots to literal `#RRGGBB` values: theme token, then
 * style override (hex or `$colors.x`), then the built-in fallback.
 *
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {Record<string, string>} overrides
 * @returns {Record<string, string>}
 */
function resolvePalette(theme, overrides) {
  /** @type {Record<string, string>} */
  const palette = {}
  for (const key of PALETTE_KEYS) {
    const token = typeof theme.colors[key] === 'string' && theme.colors[key] !== ''
      ? theme.colors[key]
      : FALLBACK_COLORS[key]
    palette[key] = resolveColor(overrides[key], theme, resolveColor(token, theme, FALLBACK_COLORS[key]))
  }
  return palette
}

/**
 * deckId: an explicit safe id wins; otherwise a deterministic path-safe slug of
 * the title (CJK-only titles get a stable digest, never a random id).
 *
 * @param {any} requested
 * @param {string} title
 */
function resolveDeckId(requested, title) {
  if (typeof requested === 'string' && requested.trim() !== '') {
    const trimmed = requested.trim()
    if (/^[a-zA-Z0-9._-]{1,64}$/.test(trimmed)) return trimmed
    const slug = deckIdFromTitle(trimmed)
    if (slug !== '') return slug
  }
  return deckIdFromTitle(title)
}

/** @param {string} title */
function deckIdFromTitle(title) {
  const base = String(title ?? '')
    .normalize('NFKD')
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/[-._]{2,}/g, '-')
    .replace(/^[-._]+|[-._]+$/g, '')
    .slice(0, 48)
  if (/[a-z]/.test(base) && base.length > 0) return base
  return `deck-${digest(title)}`
}

/**
 * Deterministic 32-bit FNV-1a digest in base36. Used only for ids.
 * @param {string} text
 */
function digest(text) {
  let hash = 0x811c9dc5
  for (const char of String(text ?? '')) {
    hash ^= char.codePointAt(0) ?? 0
    hash = Math.imul(hash, 0x01000193) >>> 0
  }
  return hash.toString(36).padStart(7, '0').slice(0, 7)
}

/* ------------------------------------------------------------------ *
 * Geometry + text helpers
 * ------------------------------------------------------------------ */

/**
 * Deterministic `<slideId>-<role>-<n>` ids, unique within a slide.
 * @param {string} slideId
 */
function makeSlideBuilder(slideId) {
  /** @type {Map<string, number>} */
  const counts = new Map()
  /** @type {any[]} */
  const elements = []
  const nextId = (role) => {
    const n = (counts.get(role) ?? 0) + 1
    counts.set(role, n)
    const id = `${slideId}-${role}-${n}`
    return id.length <= 64 ? id : `${slideId.slice(0, 40)}-${role}-${n}`
  }
  return {
    elements,
    /**
     * @param {string} role
     * @param {string} content
     * @param {{x: number, y: number, w: number, h: number}} geo
     * @param {any} style
     */
    text(role, content, geo, style) {
      elements.push({
        type: /** @type {const} */ ('text'),
        id: nextId(role),
        x: px(geo.x),
        y: px(geo.y),
        w: px(geo.w),
        h: px(geo.h),
        role,
        content,
        style,
      })
    },
    /**
     * @param {string} role id segment, e.g. `bar` / `card`
     * @param {'rect' | 'roundRect' | 'ellipse' | 'line' | 'rightArrow'} shapeType
     * @param {{x: number, y: number, w: number, h: number}} geo
     * @param {any} style
     */
    shape(role, shapeType, geo, style) {
      elements.push({
        type: /** @type {const} */ ('shape'),
        id: nextId(role),
        shapeType,
        x: px(geo.x),
        y: px(geo.y),
        w: px(geo.w),
        h: px(geo.h),
        style,
      })
    },
  }
}

/**
 * Height needed by a text box. `estimateLines` reports `lines * fontSize * 1.35`;
 * rounding up keeps `estimateTextOverflow` (which allows 1px of slack) quiet.
 *
 * @param {string} content
 * @param {number} width
 * @param {number} fontSize
 */
function textHeight(content, width, fontSize) {
  return Math.max(8, Math.ceil(estimateLines(content, width, fontSize).heightPx))
}

/**
 * Largest candidate size at which the text stays within `maxLines`.
 * @param {readonly number[]} sizes
 * @param {string} text
 * @param {number} width
 * @param {number} maxLines
 */
function chooseSize(sizes, text, width, maxLines) {
  for (const size of sizes) {
    if (estimateLines(text, width, size).lines <= maxLines) return size
  }
  return sizes[sizes.length - 1]
}

/**
 * Shorten `text` with an ellipsis until it fits `maxLines`, so an over-long title
 * from user input can never overflow its box (or push later elements down).
 *
 * @param {string} text
 * @param {number} width
 * @param {number} fontSize
 * @param {number} maxLines
 * @returns {{text: string, truncated: boolean}}
 */
function clampToLines(text, width, fontSize, maxLines) {
  if (estimateLines(text, width, fontSize).lines <= maxLines) return { text, truncated: false }
  const chars = Array.from(text)
  for (let keep = chars.length - 1; keep > 1; keep -= 1) {
    const candidate = `${chars.slice(0, keep).join('')}…`
    if (estimateLines(candidate, width, fontSize).lines <= maxLines) return { text: candidate, truncated: true }
  }
  return { text: chars.slice(0, Math.max(1, Math.min(2, chars.length))).join(''), truncated: true }
}

/**
 * Text style object: only the keys the strict `TextElementSchema` allows.
 * @param {number} fontSize
 * @param {string} color resolved `#RRGGBB`
 * @param {{bold?: boolean, align?: 'left' | 'center' | 'right', valign?: 'top' | 'middle' | 'bottom'}} [extra]
 */
function textStyle(fontSize, color, extra = {}) {
  /** @type {{fontSize: number, color: string, bold?: boolean, align?: 'left' | 'center' | 'right', valign?: 'top' | 'middle' | 'bottom'}} */
  const style = { fontSize, color }
  if (extra.bold !== undefined) style.bold = extra.bold
  if (extra.align !== undefined) style.align = extra.align
  if (extra.valign !== undefined) style.valign = extra.valign
  return style
}

/** @param {number[]} heights @param {number} gap */
function totalHeight(heights, gap) {
  if (heights.length === 0) return 0
  return heights.reduce((sum, height) => sum + height, 0) + gap * (heights.length - 1)
}

/* ------------------------------------------------------------------ *
 * Text utilities
 * ------------------------------------------------------------------ */

/**
 * Split one over-long bullet into pieces of at most `limit` characters, on
 * 。；;！？!? and newlines, merging short fragments back together.
 *
 * @param {string} text
 * @param {number} limit
 * @returns {string[]}
 */
function splitLongBullet(text, limit) {
  const value = oneLine(text)
  if (Array.from(value).length <= limit) return value === '' ? [] : [value]

  const fragments = String(text)
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .flatMap((line) => line.split(/(?<=[。；;！？!?])/))
    .map((fragment) => fragment.trim())
    .filter((fragment) => fragment !== '')

  /** @type {string[]} */
  const parts = []
  /** @param {string} piece */
  const push = (piece) => {
    const last = parts.length > 0 ? parts[parts.length - 1] : null
    if (last !== null && Array.from(last).length + Array.from(piece).length <= limit) {
      parts[parts.length - 1] = `${last}${piece}`
      return
    }
    parts.push(piece)
  }

  for (const fragment of fragments) {
    const chars = Array.from(fragment)
    if (chars.length <= limit) {
      push(fragment)
      continue
    }
    for (let start = 0; start < chars.length; start += limit) {
      parts.push(chars.slice(start, start + limit).join(''))
    }
  }
  return parts.length === 0 ? [value] : parts
}

/** @param {unknown} value @param {number} size */
function chunk(value, size) {
  const items = Array.isArray(value) ? value : []
  const width = Math.max(1, Math.floor(size))
  /** @type {any[][]} */
  const out = []
  for (let index = 0; index < items.length; index += width) out.push(items.slice(index, index + width))
  return out
}

/** `^- `, `^* `, `^+ ` and `^1. ` all start a bullet. */
function matchBulletLine(line) {
  const match = /^\s*(?:[-*+]|\d{1,2}[.)])\s+(.*)$/.exec(line)
  return match === null ? null : match[1]
}

function isBulletLine(line) {
  return matchBulletLine(line) !== null
}

function isAtxHeading(line) {
  return /^#{1,6}\s*\S/.test(line) || /^#{1,6}$/.test(line)
}

/** @param {string} line @param {'=' | '-'} marker */
function isSetextUnderline(line, marker) {
  return new RegExp(`^\\${marker}{2,}$`).test(line)
}

/** `备注：…` / `注:` / `Notes: …` @param {string} line */
function matchNotesLine(line) {
  const zh = /^(?:备注|说明|注)\s*[:：]\s?(.*)$/.exec(line)
  if (zh !== null) return zh[1]
  const en = /^notes?\s*[:：]\s?(.*)$/i.exec(line)
  if (en !== null) return en[1]
  return null
}

/**
 * Drop a trailing BrandDeck `{#id}` marker from a heading (our own outline format
 * uses it; a user pasting an outline should not see it in the deck).
 * @param {string} text
 */
function stripTrailingId(text) {
  return text.replace(/\s*\{#[^}]*\}\s*$/, '').trim()
}

/**
 * Strip inline Markdown emphasis/code/link syntax: the DSL renders plain text.
 * @param {string} text
 */
function cleanInline(text) {
  return String(text ?? '')
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/`{1,3}([^`]*)`{1,3}/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/__([^_]+)__/g, '$1')
    .replace(/(^|[\s(（])\*([^*\s][^*]*)\*(?=$|[\s)）,.，。;；])/g, '$1$2')
    .replace(/~~([^~]+)~~/g, '$1')
    .trim()
}

/**
 * Join two hard-wrapped fragments without inserting a space between CJK glyphs.
 * @param {string} left
 * @param {string} right
 */
function joinWrapped(left, right) {
  if (left === '') return right
  if (right === '') return left
  const tail = left.slice(-1)
  const head = right.slice(0, 1)
  if (CJK.test(tail) || CJK.test(head)) return `${left}${right}`
  return `${left} ${right}`
}

/**
 * Title from the first sentence of a paragraph: <= `maxChars` code points, with
 * an ellipsis when shortened.
 * @param {string} text
 * @param {number} maxChars
 */
function deriveTitle(text, maxChars) {
  const flat = oneLine(text)
  const sentence = /^[^。！？!?；;\n]+/.exec(flat)
  let title = (sentence === null ? flat : sentence[0]).trim().replace(/[。．.,，;；:：\s]+$/, '')
  if (title === '') title = flat
  const chars = Array.from(title)
  if (chars.length > maxChars) return `${chars.slice(0, maxChars).join('')}…`
  return title
}

/** @param {unknown} value */
function oneLine(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim()
}

/** @param {string} value @param {number} max */
function clampChars(value, max) {
  const chars = Array.from(String(value ?? ''))
  if (chars.length <= max) return String(value ?? '')
  return `${chars.slice(0, Math.max(1, max - 1)).join('')}…`
}

/** @param {unknown} value */
function isObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

/** @param {number} value */
function px(value) {
  return Math.round(value)
}

/** @param {number} value @param {number} min @param {number} max */
function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value))
}

/** @param {unknown} value @param {string[]} allowed @param {string} fallback */
function oneOf(value, allowed, fallback) {
  return typeof value === 'string' && allowed.includes(value) ? value : fallback
}