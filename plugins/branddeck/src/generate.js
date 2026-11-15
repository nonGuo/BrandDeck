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
 * LAYOUT INTENTS (版式意图)
 * A slide carries `kind`, either declared by the author (`版式：对比` / `layout: stats`)
 * or inferred from its bullets by {@link inferKind}. Every kind has its own builder,
 * so 对比 / 时间线 / 指标 stop looking like three identical bullet lists:
 * - `section`  — a heading and nothing else.
 * - `bullets`  — the classic list / 2-column cards (density decides the cap).
 * - `compare`  — two columns (旧/现状 vs 新/目标), a heading per column, a hairline
 *                divider between them, short body copy per column.
 * - `timeline` — a horizontal band of at most 5 numbered dots joined by segments,
 *                one short caption under each. Extra bullets fold into the last node
 *                as `+N` and their text is preserved in the speaker notes.
 * - `stats`    — 2–4 stat cards in a row: the extracted number big on top, the rest
 *                of the words as the label underneath.
 * The three new kinds are geometry-bound, not list-bound, so they deliberately
 * ignore `bulletLayout` and the density *caps* (each builder documents why); tone
 * still changes colour usage only, so the geometry is identical across tones. The
 * chosen kind is written back onto `CardSlide.kind` (a first-class DSL field since
 * the outline learned `版式：`), which is what lets the text projection round-trip it.
 *
 * TWO-TRACK TEXT (短要点 + 演讲备注)
 * The page shows 3–5 short bullets; detail belongs in the notes. `composeDeck` never
 * truncates a slide's `notes` — they land verbatim in `CardSlide.speakerNotes` — and
 * for a slide WITHOUT notes an over-long bullet is split at its first clause: the
 * clause stays visible (shortened to <= `MAX_BULLET_CHARS`), everything beyond it
 * moves into `speakerNotes` with the `要点过长，细节已移入演讲备注` warning, so no word
 * is ever dropped. A slide that already has notes keeps the older behaviour (split
 * into several bullets) because its author is already using the notes track.
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
 *   before layout, and a warning is emitted per split bullet — unless the slide has
 *   no notes, in which case the overflow becomes the notes (see TWO-TRACK TEXT).
 * - Inference (bullet shape → kind) only ever applies when the author did not say:
 *   an explicit `版式：` line is never overridden, and a slide whose rules do not
 *   fire keeps `kind` absent (= plain bullets) exactly as before.
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

/** Warning emitted when an over-long bullet's detail is moved into the notes. */
export const DETAIL_TO_NOTES_WARNING = '要点过长，细节已移入演讲备注'

/**
 * Layout kinds an OUTLINE may declare, i.e. the subset of `SlideKindSchema` that the
 * text projection can express with a `版式：…` line (`cover` / `cards` are decided by
 * the builder, not by the author). Every value here is a valid `CardSlide.kind`.
 * @type {ReadonlyArray<'section' | 'bullets' | 'compare' | 'timeline' | 'stats'>}
 */
export const CONTENT_KINDS = Object.freeze(['section', 'bullets', 'compare', 'timeline', 'stats'])

/**
 * @typedef {'section' | 'bullets' | 'compare' | 'timeline' | 'stats'} SlideKind
 */

/**
 * `版式：…` / `layout: …` aliases, Chinese and English.
 *
 * outline.js keeps its own label table for the deck ⇄ Markdown round-trip (it also
 * knows `卡片`, which only exists as a *presentation* of `bullets`). This table is
 * the parser's: it accepts a superset, and unknown values are silently ignored.
 * @type {ReadonlyMap<string, 'section' | 'bullets' | 'compare' | 'timeline' | 'stats'>}
 */
const LAYOUT_ALIASES = new Map([
  ['对比', 'compare'],
  ['比较', 'compare'],
  ['两栏', 'compare'],
  ['compare', 'compare'],
  ['contrast', 'compare'],
  ['时间线', 'timeline'],
  ['时间轴', 'timeline'],
  ['流程', 'timeline'],
  ['步骤', 'timeline'],
  ['timeline', 'timeline'],
  ['process', 'timeline'],
  ['steps', 'timeline'],
  ['step', 'timeline'],
  ['指标', 'stats'],
  ['数据', 'stats'],
  ['stat', 'stats'],
  ['stats', 'stats'],
  ['kpi', 'stats'],
  ['列表', 'bullets'],
  ['清单', 'bullets'],
  ['bullets', 'bullets'],
  ['bullet', 'bullets'],
  ['list', 'bullets'],
  ['章节', 'section'],
  ['过渡', 'section'],
  ['section', 'section'],
])

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
 * Bullets per page for the geometry-bound layouts. These are *width* limits, not
 * style choices: `stats` needs 2–4 cards to stay readable, `compare` two columns of
 * at most two lines each. Density therefore does not change them (a `dense` stat row
 * of six cards is unreadable at any tone); the remainder spills onto a `（续）` page.
 */
const LAYOUT_BULLET_CAP = Object.freeze({ compare: 4, stats: 4 })

/** Timeline nodes per slide; extra bullets fold into the last node as `+N`. */
const TIMELINE_MAX_NODES = 5

/** Human labels for warnings ("卡片版式单页最多 …"). */
const KIND_LABEL = Object.freeze({ compare: '对比', timeline: '时间线', stats: '指标' })

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

/**
 * Body sizes a *sparse* list may step up to. A four-bullet page at 22px leaves the
 * bottom half of a 720px canvas empty; the same copy at 28px reads as a composed page.
 * Ordered large → small and only ever tried upward from `TYPE.body[0]`, so a dense page
 * is untouched.
 */
const SPARSE_BODY_SIZES = Object.freeze([32, 30, 28, 26, 24])

/** How far a sparse list may be pushed down while centring (keeps it near its title). */
const BLOCK_CENTER_MAX = 48

/**
 * A trailing list page with fewer bullets than this reads as an empty slide. The planner
 * stretches the density cap to avoid one (see `planListChunks`), never past `LIST_HARD_MAX`
 * — beyond that the page has stopped being a list you can read at a glance.
 */
const LIST_MIN_PAGE = 3
const LIST_HARD_MAX = 8

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
 * - `版式：对比` / `版式: 时间线` / `layout: stats` sets `slide.kind` explicitly (see
 *   {@link parseLayoutLine}); the line is metadata and never becomes a bullet.
 * - Without such a line, `kind` is inferred from the bullets by {@link inferKind}.
 * - A heading with no bullets becomes `{kind:'section'}`.
 * - A `> quote` line is ignored when it is the machine header BrandDeck emits
 *   (`> deck: … · template: …`), otherwise it is used as the deck subtitle.
 * - With no slide heading at all, paragraphs are chunked into slides of at most
 *   `paragraphBullets` (default 5) bullets, titled from the first sentence.
 * - Never throws. Empty/odd input -> `{title:'未命名演示', slides: []}`.
 *
 * @param {string} text
 * @param {{subtitle?: string, fallbackTitle?: string, paragraphBullets?: number, maxTitleChars?: number}} [options]
 * @returns {{title: string, subtitle?: string, slides: Array<{title: string, bullets: string[], notes?: string, kind?: SlideKind}>}}
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
  /** @type {Array<{title: string, bullets: string[], notes?: string, kind?: SlideKind}>} */
  const slides = []
  /** @type {{title: string, bullets: string[], notes?: string, kind?: SlideKind} | null} */
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
    const named = cleanSlideTitle(title)
    /** @type {{title: string, bullets: string[], notes?: string, kind?: SlideKind}} */
    const slide = { title: named === '' ? '未命名页面' : named, bullets: /** @type {string[]} */ ([]) }
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

    // `版式：对比` / `layout: stats` — layout intent, never copy. Recognised anywhere
    // in a slide section (also as `- 版式：…`); an unknown value is dropped silently
    // (this function is pure and has no warnings channel) and, since it is still a
    // `版式：` line, it is removed rather than echoed back as a bullet.
    const layout = parseLayoutLine(trimmed)
    if (layout !== null) {
      flushParagraph()
      if (current !== null && layout.kind !== null) current.kind = layout.kind
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

  /**
   * Normalise + type. Author-declared kinds win; otherwise {@link inferKind} may tag
   * a slide, and a slide it leaves alone gets NO `kind` property at all (the old
   * contract: absent means "plain bullets", and outline-test pins that for section-
   * less bullet slides). `section` is structural: it is only ever set when the slide
   * really has no bullets, so an explicit `版式：章节` on a slide that *does* carry
   * copy cannot silently drop that copy.
   *
   * @type {Array<{title: string, bullets: string[], notes?: string, kind?: SlideKind}>}
   */
  const normalized = slides.map((slide) => {
    /** @type {{title: string, bullets: string[], notes?: string, kind?: SlideKind}} */
    const out = { title: slide.title === '' ? '未命名页面' : slide.title, bullets: slide.bullets }
    if (slide.notes !== undefined && slide.notes !== '') out.notes = slide.notes
    if (slide.bullets.length === 0) {
      out.kind = 'section'
      return out
    }
    if (slide.kind !== undefined && slide.kind !== 'section') {
      out.kind = slide.kind
      return out
    }
    const inferred = inferKind(slide.bullets)
    if (inferred !== null) out.kind = inferred
    return out
  })

  /** @type {{title: string, subtitle?: string, slides: typeof normalized}} */
  const result = { title }
  if (subtitle !== null && subtitle !== '') result.subtitle = subtitle
  result.slides = normalized
  return result
}

/* ------------------------------------------------------------------ *
 * Outline: layout intent (explicit line + inference)
 * ------------------------------------------------------------------ */

/**
 * Read one `版式：<name>` / `排版:` / `layout:` line.
 *
 * Also matches the same line written as a bullet (`- 版式：对比`), because the
 * content prompt asks the model for bulleted copy and a model will occasionally
 * decorate the layout line the same way. Both colon widths are accepted; the value
 * is trimmed, lower-cased and stripped of a trailing sentence mark, so
 * `layout: Compare。` still maps.
 *
 * @param {string} line
 * @returns {{kind: SlideKind | null} | null} `null` when the line is not a layout
 *   line at all; `{kind: null}` when it is one with an unknown value
 */
function parseLayoutLine(line) {
  const match = /^(?:[-*+]\s+)?(?:版式|排版|layout)\s*[:：]\s*(.*)$/i.exec(line.trim())
  if (match === null) return null
  const value = oneLine(match[1])
    .toLowerCase()
    .replace(/[。．.；;，,、\s]+$/, '')
  const kind = LAYOUT_ALIASES.get(value)
  return { kind: kind === undefined ? null : kind }
}

/** First-character contrast cues (spec list). `i` covers `VS`/`Before`/`After`. */
const CONTRAST_CUE = /^(?:对比|相比|旧|新|之前|以前|原先|现在|目前|当前|vs|before|after|old|new)/i
/** A contrast word anywhere in the bullet, for the `标签：` form. */
const CONTRAST_WORD = /对比|相比|旧|新|之前|以前|原先|现在|目前|当前|vs\b|before|after|old|new/i
const LABEL_COLON = /[：:]/

/**
 * A "statistic" bullet: a digit run glued to a unit, and short enough to sit on a
 * card (`STAT_MAX_CHARS` code points — a superset of the 汉字 count, so the rule is
 * conservative in the direction that matters: it never fires on prose).
 *
 * The unit list deliberately includes *operational* units (天/小时/周/人/次/页/个/%),
 * not just money and percentages: a real deck says "出片时间从 5 天降到 4 小时" far
 * more often than "ROI 138%", and a unit-less number would leave the card with no
 * headline at all.
 */
const STAT_UNIT = '天|小时|分钟|周|月|年|人|次|页|个|条|项|家|台|套|万|亿|元|美元|块|个百分点|倍|%|％|k|K|M|B'
/**
 * Chinese numerals used as quantities: 三分之一 / 三倍 / 两成 / 四分之三.
 * A "statistics" page in a Chinese deck says 成本降至三分之一 far more often than 33%,
 * so the extractor has to understand it or the card silently loses its headline.
 */
const STAT_CN_NUMERAL = '(?:[一二三四五六七八九十百千万亿两半]+分之[一二三四五六七八九十百千]+|[一二三四五六七八九十百千万亿两]+(?:倍|成|成多|个百分点))'
const STAT_NUMBER = new RegExp(
  `([$¥￥]\\s*\\d+(?:[.,]\\d+)?\\s*(?:${STAT_UNIT})?|\\d+(?:[.,]\\d+)?\\s*(?:${STAT_UNIT})|${STAT_CN_NUMERAL})`,
)
/**
 * A change expressed as two numbers: `从 5 天降到 4 小时` / `450ms → 18ms` / `5→2`.
 * Rendered whole, because "5 天" alone loses the point of the metric.
 */
const STAT_RANGE = /(?:从\s*)?(\d+(?:[.,]\d+)?\s*(?:天|小时|分钟|周|月|年|人|次|页|个|条|项|家|台|套|ms|s|%|％|倍|元|万)?)\s*(?:降(?:低|到|至)?|提升(?:到|至)?|缩短(?:到|至)?|增长(?:到|至)?|上升到|→|->|~|至|到)\s*(\d+(?:[.,]\d+)?\s*(?:天|小时|分钟|周|月|年|人|次|页|个|条|项|家|台|套|ms|s|%|％|倍|元|万)?)/
const STAT_MAX_CHARS = 24

/** Ordered cues that make a bullet read as a step: 第一步 / 阶段 2 / Q1 / 1. / →. */
const TIMELINE_CUE = /^(?:第\s*[一二三四五六七八九十\d]+\s*[步阶段章节]|阶段\s*[一二三四五六七八九十\d]*|Q[1-4]\b|q[1-4]\b|\d{1,2}\s*[.、)]|→)/

/** @param {string} bullet */
function isContrastBullet(bullet) {
  const flat = oneLine(bullet)
  if (CONTRAST_CUE.test(flat)) return true
  // `数据看板：交付从三天缩短到半小时` — a labelled pair that names a contrast.
  return LABEL_COLON.test(flat) && CONTRAST_WORD.test(flat)
}

/** @param {string} bullet */
function isStatBullet(bullet) {
  const flat = oneLine(bullet)
  return Array.from(flat).length <= STAT_MAX_CHARS && STAT_NUMBER.test(flat)
}

/** @param {string} bullet */
function isTimelineBullet(bullet) {
  const flat = oneLine(bullet)
  return TIMELINE_CUE.test(flat) || flat.includes('→')
}

/**
 * Infer a layout kind from a bullet list — deterministic, order-preserving, and only
 * ever consulted when the author did not declare one.
 *
 * Rules (first hit wins, so the result is stable for any input):
 * 1. `compare`  — >= 2 bullets read as a contrast: they start with a cue
 *    (对比/相比/旧/新/之前/现在/vs/before/after) or carry a `标签：` colon together
 *    with a contrast word.
 * 2. `stats`    — >= 2 bullets are dominated by a number (a digit run with a unit
 *    from % 倍 万 亿 元 $) and are <= `STAT_MAX_CHARS` code points long.
 * 3. `timeline` — >= 3 bullets start with an ordered cue (第一步/阶段/Q1/第1/1./→)
 *    or contain `→`.
 * 4. Otherwise  — `null`, i.e. leave `kind` absent (= plain bullets). A slide with
 *    no bullets is a `section` and never reaches this function.
 *
 * Order matters: a `第一步：成本下降 138%` deck scores on both stats and timeline, and
 * numbers are the stronger signal (a stat row is what the user wants to see).
 *
 * @param {string[]} bullets
 * @returns {SlideKind | null}
 */
function inferKind(bullets) {
  const list = bullets.map((bullet) => oneLine(bullet)).filter((bullet) => bullet !== '')
  if (list.length === 0) return null
  if (list.filter(isContrastBullet).length >= 2) return 'compare'
  if (list.filter(isStatBullet).length >= 2) return 'stats'
  if (list.filter(isTimelineBullet).length >= 3) return 'timeline'
  return null
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

  // `kind` is a first-class DSL field (schema.js `SlideKindSchema`): the generator
  // states the layout it actually drew, that is what the compiler/panel read and what
  // the text outline round-trips as a `版式：…` line. `cover` / `section` are the
  // structural kinds; the outline does not offer them for editing, so they stay
  // implicit there and explicit here.
  const coverSlide = buildCoverSlide(nextSlideId(), ctx, title, outline.subtitle)
  coverSlide.kind = 'cover'
  slides.push(coverSlide)

  for (const item of content) {
    const slideId = nextSlideId()
    /** @type {import('./schema.js').CardSlide} */
    let slide
    if (item.kind === 'section') {
      slide = buildSectionSlide(slideId, ctx, item)
    } else if (item.kind === 'compare') {
      slide = buildCompareSlide(slideId, ctx, item)
    } else if (item.kind === 'timeline') {
      slide = buildTimelineSlide(slideId, ctx, item)
    } else if (item.kind === 'stats') {
      slide = buildStatsSlide(slideId, ctx, item)
    } else if (item.kind === 'breather') {
      slide = buildBreatherSlide(slideId, ctx, item)
    } else {
      slide = buildBulletSlide(slideId, ctx, item)
    }
    // `item.kind` is what buildContentSlides decided to draw (it already resolved
    // `cards` vs `bullets`), so the stored deck never claims a layout it does not
    // show.
    slide.kind = item.kind
    slides.push(slide)
  }

  const lastTitle = content.length > 0 ? content[content.length - 1].title : ''
  if (wantsClosing && !CLOSING_LIKE.test(lastTitle)) {
    const closingSlide = buildClosingSlide(nextSlideId(), ctx)
    closingSlide.kind = 'cover'
    slides.push(closingSlide)
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
  /** @type {Array<{title: string, bullets: string[], notes?: string, kind: SlideKind}>} */
  const slides = []
  for (const raw of Array.isArray(src.slides) ? src.slides : []) {
    if (!isObject(raw)) continue
    const slideTitle = typeof raw.title === 'string' ? oneLine(raw.title) : ''
    const bullets = (Array.isArray(raw.bullets) ? raw.bullets : [])
      .filter((bullet) => typeof bullet === 'string')
      .map((bullet) => oneLine(bullet))
      .filter((bullet) => bullet !== '')
    // Unknown kinds are ignored (`CONTENT_KINDS` is the outline-reachable subset of
    // `SlideKindSchema`), and `section` is structural: only a slide with no bullets is
    // one. An explicit `版式：章节` on a slide that carries copy falls back to bullets
    // rather than dropping the copy on the floor.
    const declared = CONTENT_KINDS.includes(raw.kind) ? raw.kind : undefined
    const isSection = bullets.length === 0
    const kind = isSection ? 'section' : declared === undefined || declared === 'section' ? 'bullets' : declared
    /** @type {{title: string, bullets: string[], notes?: string, kind: SlideKind}} */
    const item = {
      title: slideTitle === '' ? '未命名页面' : slideTitle,
      bullets: isSection ? [] : bullets,
      kind,
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
 * draws. Handles the two-track text pass and both kinds of spilling:
 * - count: the density cap (rule 4) spills a title's bullets onto `<title>（续）`;
 * - height: `cards` rows are sized to their content, so a slide whose rows would
 *   leave the safe box also spills (see {@link planCardChunks}).
 *
 * `compare` / `stats` are width-bound layouts, so they always render at most
 * `LAYOUT_BULLET_CAP` bullets per page and spill the rest the same way. `timeline`
 * never spills: extra bullets fold into the last node (`+N`) and their text is kept
 * in the notes (see {@link foldTimeline}).
 *
 * @param {{slides: Array<{title: string, bullets: string[], notes?: string, kind: SlideKind}>}} outline
 * @param {ReturnType<typeof normalizeStyle>} style
 * @param {LocalContext} ctx
 * @param {string[]} warnings
 * @returns {Array<{kind: SlideKind, title: string, bullets: string[], notes?: string}>}
 */
function buildContentSlides(outline, style, ctx, warnings) {
  /** @type {Array<{kind: SlideKind, title: string, bullets: string[], notes?: string}>} */
  const out = []
  const cap = DENSITY[style.density][style.bulletLayout]

  outline.slides.forEach((slide, index) => {
    const label = `第 ${index + 1} 页「${clampChars(slide.title, 16)}」`
    if (slide.kind === 'section' || slide.bullets.length === 0) {
      out.push({ kind: 'section', title: slide.title, bullets: [] })
      if (slide.notes !== undefined) out[out.length - 1].notes = slide.notes
      return
    }

    // Two-track text: long bullets either keep today's split, or — when there is no
    // notes track yet — hand their overflow detail to the speaker notes.
    const track = planTwoTrack(slide.bullets, slide.notes, label, warnings)
    let kind = slide.kind

    if (kind === 'timeline') {
      const folded = foldTimeline(track, label, warnings)
      /** @type {{kind: SlideKind, title: string, bullets: string[], notes?: string}} */
      const item = { kind, title: slide.title, bullets: folded.bullets }
      if (folded.notes !== undefined) item.notes = folded.notes
      out.push(item)
      return
    }

    // `compare` / `stats` are bound by their column/card width, not by the density
    // cap: `LAYOUT_BULLET_CAP` is what keeps a row readable. More bullets than that and
    // the page is no longer a row/two-column page at all — so it degrades to a list
    // instead of spilling the remainder onto a nearly empty `（续）` slide.
    if ((kind === 'compare' || kind === 'stats') && track.bullets.length > LAYOUT_BULLET_CAP[kind]) {
      warnings.push(
        `${label} · ${KIND_LABEL[kind]}版式单页最多 ${LAYOUT_BULLET_CAP[kind]} 条，已改为列表呈现`,
      )
      kind = 'bullets'
    }

    if (kind === 'compare' || kind === 'stats') {
      const chunks = chunk(track.bullets, LAYOUT_BULLET_CAP[kind])
      chunks.forEach((bullets, chunkIndex) => {
        /** @type {{kind: SlideKind, title: string, bullets: string[], notes?: string}} */
        const item = {
          kind,
          title: chunkIndex === 0 ? slide.title : `${slide.title}（续）`,
          bullets,
        }
        if (chunkIndex === 0 && track.notes !== undefined) item.notes = track.notes
        out.push(item)
      })
      return
    }

    const chunks =
      style.bulletLayout === 'cards'
        ? planCardChunks(track.bullets, cap, ctx, slide.title)
        : planListChunks(track.bullets, cap, ctx, slide.title)
    const countChunks = Math.ceil(track.bullets.length / Math.max(1, cap))
    if (chunks.length > countChunks) {
      warnings.push(`${label} · 要点按高度排布，已溢出到 ${chunks.length - countChunks} 页「（续）」`)
    }

    chunks.forEach((bullets, chunkIndex) => {
      /** @type {{kind: SlideKind, title: string, bullets: string[], notes?: string}} */
      const item = {
        // The slide's `kind` is what the compiler/panel and the text projection read:
        // `cards` is a distinct rendering, so it gets its own value.
        kind: style.bulletLayout === 'cards' ? 'cards' : 'bullets',
        title: chunkIndex === 0 ? slide.title : `${slide.title}（续）`,
        bullets,
      }
      if (chunkIndex === 0 && track.notes !== undefined) item.notes = track.notes
      out.push(item)
    })
  })

  // Rhythm: break up monotony without inventing meaning (see the two helpers below).
  const varied = varyAdjacentLayouts(out, style)
  const paced = insertBreathers(varied, style)
  return paced
}

/* ------------------------------------------------------------------ *
 * Rhythm: adjacent-layout variety and deliberate breathing room
 * ------------------------------------------------------------------ */

/** Kinds that present the same thing — an unordered set of points. */
const LISTY = new Set(['bullets', 'cards'])

/**
 * Keep two neighbouring slides from being visually identical, using only swaps that
 * do not change what a slide *means*.
 *
 * The honest constraint here: we must not turn a list into a timeline just to look
 * different — that invents a sequence the author never wrote. So the allowed moves are
 * those that re-render the same content:
 *   · two bullets as `bullets` right after a list → `compare` (two columns, neutral);
 *   · a `cards` slide sandwiched between lists → plain `bullets` (same semantics).
 * Anything else is left alone; a repeated layout is much better than a lie.
 *
 * @param {Array<{kind: SlideKind, title: string, bullets: string[], notes?: string}>} slides
 * @param {ReturnType<typeof normalizeStyle>} style
 * @param {string[]} warnings
 */
function varyAdjacentLayouts(slides, style) {
  if (style.varyLayout === false) return slides
  const out = slides.slice()
  let run = 0
  for (let index = 0; index < out.length; index += 1) {
    const item = out[index]
    const previous = index === 0 ? null : out[index - 1]
    const isListy = LISTY.has(item.kind)
    run = isListy && previous !== null && LISTY.has(previous.kind) ? run + 1 : 0
    // Only react once the page would be the third list-like slide in a row.
    if (run < 2) continue

    if (item.kind === 'cards') {
      out[index] = { ...item, kind: 'bullets' }
      run = 0
    }
  }
  return out
}

/** A statement that is already a number, so a breather page never invents data. */
const BREATHER_NUMBER = /(\d+(?:[.,]\d+)?\s*(?:%|％|倍|天|小时|分钟|周|月|年|人|次|页|个|条|项|万|亿|元|ms|s))/

/**
 * Insert a "breathing" page every `breatherEvery` content slides.
 *
 * A breather is one oversized statement and one short caption — the density contrast
 * a human designer uses to let a dense deck breathe. The statement is *taken from the
 * deck*: the strongest existing number if the slides around it carry one, otherwise
 * the section's own phrasing. We never make a number up; if nothing numeric exists the
 * page is skipped rather than faked.
 *
 * Off by default (`style.breather`), because forced white space is just another
 * mannerism when it does not match the material.
 *
 * @param {Array<{kind: SlideKind, title: string, bullets: string[], notes?: string}>} slides
 * @param {ReturnType<typeof normalizeStyle>} style
 * @param {LocalContext} _ctx
 */
function insertBreathers(slides, style) {
  if (style?.breather !== true) return slides
  const every =
    Number.isFinite(style.breatherEvery) && style.breatherEvery >= 3 ? Math.floor(style.breatherEvery) : 5
  const out = []
  let dense = 0
  for (let index = 0; index < slides.length; index += 1) {
    const item = slides[index]
    out.push(item)
    const isContent =
      LISTY.has(item.kind) || item.kind === 'compare' || item.kind === 'stats' || item.kind === 'timeline'
    if (!isContent) continue
    dense += 1
    if (dense % every !== 0) continue
    // Never place a breather as the last page: a lone number followed by the closing
    // slide reads like a mistake rather than a pause.
    if (index >= slides.length - 1) continue

    // Borrow the most concrete number from the slides around this point; if the
    // material has none, skip the page instead of inventing data.
    let numeric
    for (let offset = -2; offset <= 2 && numeric === undefined; offset += 1) {
      const neighbour = slides[index + offset]
      if (neighbour === undefined) continue
      for (const bullet of neighbour.bullets) {
        const match = BREATHER_NUMBER.exec(bullet)
        if (match !== null) {
          numeric = { bullet, match }
          break
        }
      }
    }
    if (numeric === undefined) continue
    // Only lift the number out of the sentence when the sentence is short enough to be
    // read as a label. Otherwise the split mangles the meaning: "出片时间从 5 天降到
    // 4 小时" would become "5 天 / 出片时间从 降到 4 小时", which also reads as if the
    // number were the *gain* when it was the starting point. For a longer sentence the
    // whole line is the statement and the slide's own title is the caption.
    const flat = oneLine(numeric.bullet)
    const short = Array.from(flat).length <= 14
    // A breather is a *statement*, not a sentence: if the shortest numeric bullet
    // around this point is still long, skip the page entirely. A two-line wrap that
    // breaks after "小时" is worse than no pause at all ("宁缺毋滥").
    if (!short) continue
    const headline = short ? oneLine(numeric.match[0]).replace(/\s+/g, '') : flat
    const caption = short
      ? oneLine(flat.replace(numeric.match[0], ''))
          .replace(/^[，,。；;：:、\s]+/, '')
          .replace(/[，,。；;：:、\s]+$/, '')
      : item.title
    out.push({
      kind: 'breather',
      title: '',
      bullets: [headline, caption === '' ? item.title : caption],
    })
  }
  return out
}


/* ------------------------------------------------------------------ *
 * Two-track text: 短要点 on the page, detail in the speaker notes
 * ------------------------------------------------------------------ */

/** Shortest acceptable visible clause when slicing an over-long bullet. */
const MIN_CLAUSE_CHARS = 8

/**
 * How an over-long bullet is cut at its first clause: the head stays on the page,
 * the tail becomes speaker-note detail. Clause separators are 。；;，, (and newlines,
 * already flattened by `oneLine`).
 */
const CLAUSE_SEPARATORS = '。；;，,'

/**
 * Split one over-long bullet into `{head, tail}`: `head` is the first clause (a short
 * 动宾短语 that reads as a page bullet), `tail` is everything the page no longer has
 * to show. Both halves are trimmed, the head loses its trailing separator, and no
 * character is dropped — `head + tail` is the original, minus whitespace.
 *
 * The first clause is preferred when it is at least `MIN_CLAUSE_CHARS` long (so a
 * bullet starting "首先，" does not become a two-word page bullet); when the whole
 * text has no separator at all the head is clamped to `limit` and the remainder
 * moves into the tail, which is the "finally shorten the visible bullet" step.
 *
 * @param {string} text
 * @param {number} limit `MAX_BULLET_CHARS`
 * @returns {{head: string, tail: string}}
 */
function splitFirstClause(text, limit) {
  const chars = Array.from(oneLine(text))
  let cut = -1
  let firstCut = -1
  for (let index = 0; index < chars.length; index += 1) {
    if (!CLAUSE_SEPARATORS.includes(chars[index])) continue
    if (firstCut === -1) firstCut = index + 1
    if (index + 1 >= MIN_CLAUSE_CHARS && index + 1 <= limit) {
      cut = index + 1
      break
    }
  }
  if (cut === -1) cut = firstCut === -1 ? chars.length : firstCut
  const rawHead = chars.slice(0, cut).join('')
  const tail = chars.slice(cut).join('')
  const stripped = rawHead.replace(/[。；;，,、\s]+$/, '')
  // An empty head would be a slide bullet with no content: fall back to the raw cut
  // (or the whole text when the bullet was nothing but separators).
  let head = stripped === '' ? rawHead.trim() : stripped
  if (head === '') head = oneLine(text)
  if (Array.from(head).length <= limit) return { head, tail: tail.trim() }
  const headChars = Array.from(head)
  return {
    head: headChars.slice(0, limit).join('').trim(),
    tail: `${headChars.slice(limit).join('')}${tail}`.trim(),
  }
}

/**
 * Two-track prep for one slide: shorten what is on the page and keep the detail in
 * `notes`.
 *
 * - Every bullet at or below `MAX_BULLET_CHARS` is untouched.
 * - A slide WITH notes keeps the historical behaviour: over-long bullets are split
 *   into several visible bullets (`要点超过 N 字，已拆分为 M 条`) and the notes are
 *   passed through verbatim — they are never re-cut or truncated.
 * - A slide WITHOUT notes gets the new behaviour: the first clause stays visible and
 *   the rest becomes the notes (`要点过长，细节已移入演讲备注`). That is what turns a
 *   paragraph-shaped page into 3–5 short 动宾短语 plus a speaker script.
 *
 * @param {string[]} bullets
 * @param {string | undefined} notes author-supplied notes for this slide
 * @param {string} label warning prefix
 * @param {string[]} warnings
 * @returns {{bullets: string[], notes?: string}}
 */
function planTwoTrack(bullets, notes, label, warnings) {
  const existing = typeof notes === 'string' && notes.trim() !== '' ? notes.trim() : ''
  /** @type {string[]} */
  const visible = []
  /** @type {string[]} */
  const detail = []
  let shortened = 0

  for (const bullet of bullets) {
    const flat = oneLine(bullet)
    if (Array.from(flat).length <= MAX_BULLET_CHARS) {
      if (flat !== '') visible.push(flat)
      continue
    }
    if (existing !== '') {
      const parts = splitLongBullet(bullet, MAX_BULLET_CHARS)
      if (parts.length > 1) {
        warnings.push(`${label} · 要点超过 ${MAX_BULLET_CHARS} 字，已拆分为 ${parts.length} 条`)
      }
      visible.push(...parts)
      continue
    }
    const { head, tail } = splitFirstClause(flat, MAX_BULLET_CHARS)
    visible.push(head)
    if (tail !== '') {
      detail.push(tail)
      shortened += 1
    }
  }

  /** @type {{bullets: string[], notes?: string}} */
  const out = { bullets: visible }
  if (existing !== '') out.notes = existing
  else if (detail.length > 0) out.notes = detail.join('\n')
  if (shortened > 0) warnings.push(`${label} · ${DETAIL_TO_NOTES_WARNING}（${shortened} 条）`)
  return out
}

/**
 * Timeline slides never spill: the first `TIMELINE_MAX_NODES` bullets become nodes and
 * the rest fold into the last one as `+N`. Their full text is appended to the notes so
 * folding never loses copy (the page only loses the words, not the information).
 *
 * @param {{bullets: string[], notes?: string}} track
 * @param {string} label
 * @param {string[]} warnings
 * @returns {{bullets: string[], notes?: string}}
 */
function foldTimeline(track, label, warnings) {
  const hidden = track.bullets.slice(TIMELINE_MAX_NODES)
  if (hidden.length === 0) return track
  warnings.push(
    `${label} · 时间线仅展示 ${TIMELINE_MAX_NODES} 个节点，其余 ${hidden.length} 条已折叠为 +${hidden.length}（原文进入演讲备注）`,
  )
  const folded = hidden.map((bullet) => `折叠节点：${bullet}`).join('\n')
  const notes = track.notes === undefined || track.notes === '' ? folded : `${track.notes}\n${folded}`
  return { bullets: track.bullets, notes }
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
 * Shared slide header: fitted title at the top of the safe area plus the optional
 * accent bar. Every content layout starts with exactly these two elements and then
 * places its body below `top`, so the title/logo rules are stated once.
 *
 * @param {ReturnType<typeof makeSlideBuilder>} builder
 * @param {LocalContext} ctx
 * @param {string} title
 * @returns {{top: number, label: string, text: string}} `label` is the `「…」` warning
 *   prefix; `text` is the fitted title the slide envelope should carry
 */
function buildHeader(builder, ctx, title) {
  const { box, palette, style, toneSpec, warnings } = ctx
  const width = FRAME.titleRight - box.x
  const size = chooseSize(TYPE.slideTitle, title, width, 2)
  const fitted = clampToLines(title, width, size, 2)
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
  return { top: cursor + 34, label, text: fitted.text }
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
  const { palette, style } = ctx
  const builder = makeSlideBuilder(slideId)
  const header = buildHeader(builder, ctx, item.title)
  const contentTop = header.top

  if (style.bulletLayout === 'cards') {
    placeCards(builder, ctx, item.bullets, contentTop, header.label)
  } else {
    placeList(builder, ctx, item.bullets, contentTop, header.label)
  }

  const slide = slideShell(slideId, builder, palette, header.text, item.notes)
  return slide
}

/* ------------------------------------------------------------------ *
 * Layout intents: compare / timeline / stats
 * ------------------------------------------------------------------ */

/**
 * Compare geometry. The two columns take the full safe width minus one gap; the gap
 * is where the divider lives, so the hairline never touches a sentence.
 */
const COMPARE = Object.freeze({
  gapX: 56,
  dividerW: 2,
  headingSize: 26,
  headingGap: 22,
  bodySizes: Object.freeze([22, 20, 18, 16]),
  /** Copy that needs more lines than this is clamped (and kept in the notes). */
  maxLines: 4,
})

/** "Current side" cues: the left column. */
const LEFT_CUE = /^(?:旧|原来|原先|之前|以前|当前|目前|现状|before|old|current|as-is)/i
/** "New side" cues: the right column. */
const RIGHT_CUE = /^(?:新|现在|之后|未来|目标|after|new|target|to-be)/i
/** `标签：正文` — the label becomes the column heading. */
const LABEL_PREFIX = /^([^：:]{1,12})\s*[：:]\s*(.+)$/

/**
 * Split bullets into the two sides of a contrast.
 *
 * Clean split: every bullet opens with a side cue (旧/之前/之前 vs 新/现在/目标); the
 * columns keep their relative order. Otherwise the list is halved (`ceil(n/2)` on the
 * left), which is the deterministic fallback the design note asks for. A single
 * bullet cannot be split, so it stays in the left column and the slide degrades to
 * one full-width column (the builder skips the empty column and the divider).
 *
 * @param {string[]} bullets
 * @returns {{left: string[], right: string[]}}
 */
function splitCompareColumns(bullets) {
  const list = bullets.map((bullet) => oneLine(bullet)).filter((bullet) => bullet !== '')
  if (list.length === 0) return { left: [], right: [] }

  /** @type {string[]} */
  const left = []
  /** @type {string[]} */
  const right = []
  let clean = true
  for (const bullet of list) {
    if (LEFT_CUE.test(bullet)) left.push(bullet)
    else if (RIGHT_CUE.test(bullet)) right.push(bullet)
    else {
      clean = false
      break
    }
  }
  if (clean && left.length > 0 && right.length > 0) return { left, right }

  const half = Math.ceil(list.length / 2)
  return { left: list.slice(0, half), right: list.slice(half) }
}

/**
 * Turn one column's bullets into `{heading, label, bullets}`: the heading comes from
 * the first bullet's `X：` prefix (and that prefix is then stripped from every bullet
 * of the column so the page does not repeat it), else from the caller's fallback.
 *
 * @param {string[]} bullets
 * @param {string} fallback
 * @returns {{heading: string, label: string | null, bullets: string[]}}
 */
function prepareCompareColumn(bullets, fallback) {
  const first = bullets.length > 0 ? oneLine(bullets[0]) : ''
  const labelled = LABEL_PREFIX.exec(first)
  if (labelled === null) {
    // `fallback === ''` means the caller decided this page should not carry column
    // headings at all (the author wrote no labels, so any heading would be invented).
    return {
      heading: fallback,
      label: null,
      bullets: bullets.map((bullet) => oneLine(bullet)),
    }
  }
  const label = oneLine(labelled[1])
  return {
    heading: clampChars(label, 14),
    label,
    bullets: bullets.map((bullet) => stripLabel(oneLine(bullet), label)),
  }
}

/**
 * Remove a `标签：` prefix from a bullet, keeping the bullet intact when what is left
 * would be empty (a bare `现状：` is still copy, not nothing).
 *
 * @param {string} bullet
 * @param {string} label
 */
function stripLabel(bullet, label) {
  const stripped = bullet.replace(new RegExp(`^${escapeRegExp(label)}\\s*[：:]\\s*`), '').trim()
  return stripped === '' ? bullet : stripped
}

/**
 * Compare slide: 旧/现状 on the left, 新/目标 on the right, a heading per column and a
 * hairline divider between them.
 *
 * Density decides the body font size (it has to fit a column, not a full-width line);
 * `bulletLayout` does not apply — a compare page on a `cards` deck would be the same
 * page, so honouring the flag here would only make the deck inconsistent.
 *
 * @param {string} slideId
 * @param {LocalContext} ctx
 * @param {{title: string, bullets: string[], notes?: string}} item
 * @returns {import('./schema.js').CardSlide}
 */
function buildCompareSlide(slideId, ctx, item) {
  const { box, palette, style, toneSpec, warnings } = ctx
  const builder = makeSlideBuilder(slideId)
  const header = buildHeader(builder, ctx, item.title)
  const { left, right } = splitCompareColumns(item.bullets)

  /** @type {Array<{heading: string, label: string | null, bullets: string[]}>} */
  const columns = []
  // No author-written label means no column heading: the renderer must not invent
  // 现状 / 目标. An applicability list ("适合季度汇报 / 不适合一次性创意") under those
  // headings reads as a before/after the author never wrote.
  if (left.length > 0) columns.push(prepareCompareColumn(left, ''))
  if (right.length > 0) columns.push(prepareCompareColumn(right, ''))
  if (columns.length === 0) {
    return buildSectionSlide(slideId, ctx, { title: header.text, notes: item.notes })
  }
  // Both halves of a halved list can carry the SAME `标签：` prefix (e.g. the author
  // wrote `现状：…` on every bullet). Two identical column headings would read as a
  // bug, so drop the headings — still stripping the shared label from the copy, so the
  // page does not say it twice.
  if (
    columns.length === 2 &&
    columns[0].label !== null &&
    columns[0].label === columns[1].label
  ) {
    const shared = columns[0].label
    const strip = (bullet) => stripLabel(oneLine(bullet), shared)
    columns[0] = { heading: '', label: null, bullets: left.map(strip) }
    columns[1] = { heading: '', label: null, bullets: right.map(strip) }
  }

  const twoColumns = columns.length > 1
  const colW = twoColumns ? Math.round((box.w - COMPARE.gapX) / 2) : box.w
  // Headings only exist when the author labelled the columns (`X：` prefix). Without a
  // label the page is two balanced columns and no invented 现状/目标 row.
  const anyHeading = columns.some((column) => oneLine(column.heading) !== '')
  const headingH = anyHeading
    ? Math.max(...columns.map((column) => textHeight(column.heading, colW, COMPARE.headingSize)))
    : 0
  const bodyTop = header.top + (anyHeading ? headingH + COMPARE.headingGap : 0)
  const available = box.y + box.h - bodyTop
  const baseGap = DENSITY[style.density].gap

  /** @type {Array<{heights: number[]}>} */
  let plan = []
  let fontSize = COMPARE.bodySizes[COMPARE.bodySizes.length - 1]
  for (const candidate of COMPARE.bodySizes) {
    const measured = columns.map((column) => {
      const heights = column.bullets.map((bullet) => textHeight(bullet, colW, candidate))
      return { heights }
    })
    const worst = Math.max(...measured.map((one) => totalHeight(one.heights, baseGap)))
    plan = measured
    fontSize = candidate
    if (worst <= available) break
  }

  // Unreachable for bullets that went through `planTwoTrack` (<= 90 chars fit a column
  // at 16px), but a hand-built outline must not be able to push text out of the box:
  // clamp each bullet to `maxLines`, keep the original in the notes and, if the page
  // is still too tall, compress the heights.
  const worstFinal = Math.max(...plan.map((one) => totalHeight(one.heights, baseGap)))
  if (worstFinal > available) {
    warnings.push(`${header.label} · 对比要点过长，已按 ${COMPARE.maxLines} 行截断（原文进入演讲备注）`)
    /** @type {string[]} */
    const dropped = []
    columns.forEach((column, columnIndex) => {
      column.bullets = column.bullets.map((bullet) => {
        const fitted = clampToLines(bullet, colW, fontSize, COMPARE.maxLines)
        if (fitted.truncated) dropped.push(oneLine(bullet))
        return fitted.text
      })
      plan[columnIndex] = {
        heights: column.bullets.map((bullet) => textHeight(bullet, colW, fontSize)),
      }
    })
    if (dropped.length > 0) item.notes = mergeNotes(item.notes, dropped)
  }

  const tallest = Math.max(...plan.map((one) => totalHeight(one.heights, baseGap)))
  const columnsCount = Math.max(...columns.map((column) => column.bullets.length))
  let gap =
    columnsCount > 1
      ? Math.min(DENSITY[style.density].gapMax, baseGap + Math.max(0, Math.floor((available - tallest) / (columnsCount - 1))))
      : baseGap

  // Final containment guard (same philosophy as `placeList`): if the copy still does
  // not fit after clamping, shrink the gap and then the boxes rather than letting a
  // hand-built outline push text past `innerBox`. Unreachable for generated copy,
  // because `planTwoTrack` caps a bullet at 90 chars and the layout caps the count.
  const usable = Math.max(0, available)
  const used = Math.max(...plan.map((one) => totalHeight(one.heights, gap)))
  if (used > usable) {
    gap = Math.max(4, Math.floor(gap * (usable / used)))
    const usedAgain = Math.max(...plan.map((one) => totalHeight(one.heights, gap)))
    if (usedAgain > usable) {
      const scale = usable / usedAgain
      plan = plan.map((one) => ({
        heights: one.heights.map((height) => Math.max(8, Math.floor(height * scale))),
      }))
      warnings.push(`${header.label} · 对比内容过高，已压缩行高以适应页面`)
    }
  }

  let dividerBottom = header.top
  columns.forEach((column, columnIndex) => {
    const x = twoColumns ? box.x + columnIndex * (colW + COMPARE.gapX) : box.x
    if (anyHeading && oneLine(column.heading) !== '') {
      builder.text('subtitle', column.heading, { x, y: header.top, w: colW, h: headingH }, textStyle(
        COMPARE.headingSize,
        palette[toneSpec.sectionTitle],
        { bold: true },
      ), 'heading')
    }
    let y = bodyTop
    column.bullets.forEach((bullet, bulletIndex) => {
      const height = plan[columnIndex].heights[bulletIndex]
      builder.text('body', bullet, { x, y, w: colW, h: height }, textStyle(fontSize, palette.textPrimary))
      y += height + gap
    })
    dividerBottom = Math.max(dividerBottom, y - gap)
  })

  if (twoColumns) {
    // Subtle hairline in the middle of the gap. Decorative (surface, not primary), and
    // it spans the copy rather than the full box so it reads as a divider, not a frame.
    builder.shape('divider', 'rect', {
      x: box.x + colW + Math.round((COMPARE.gapX - COMPARE.dividerW) / 2),
      y: header.top,
      w: COMPARE.dividerW,
      h: Math.max(24, dividerBottom - header.top),
    }, { fill: palette.surface })
  }

  return slideShell(slideId, builder, palette, header.text, item.notes)
}

/** Timeline geometry: one dot per node, its index inside, a caption underneath. */
const TIMELINE = Object.freeze({
  dotSize: 44,
  indexSize: 18,
  indexW: 32,
  captionSize: 16,
  captionGap: 14,
  captionMaxLines: 2,
  // A timeline label is an abbreviation, and the character cap is what keeps a
  // 5-node page inside the deck-wide 150-character budget: five 3-line captions
  // count as ~200 characters even though each looks short on screen, which trips
  // repairDeck's "shorten copy" warning — advice the user cannot act on, because
  // the text was already ellipsized.
  captionMaxChars: 26,
  captionMaxW: 280,
  segH: 3,
  /** Gap between a dot and the connector segment (so they never touch). */
  segGap: 10,
})

/** Ordered prefix the node caption does not need (the index already says the order). */
const TIMELINE_PREFIX = /^(?:第\s*[一二三四五六七八九十\d]+\s*[步阶段章节]\s*[:：、.．]?\s*|阶段\s*[一二三四五六七八九十\d]*\s*[:：、.．]?\s*|Q[1-4]\b\s*[:：、.．]?\s*|\d{1,2}\s*[.、)]\s*|→\s*)/i

/**
 * Stats geometry: one row of 2–4 cards, number on top, label underneath. Cards share
 * a height (the tallest content) so the row reads as a grid; the row is centred in
 * the space under the title, because a band of cards hugging the title looks like a
 * mistake on a sparse page.
 */
const STATS = Object.freeze({
  gapX: 24,
  padX: 24,
  padTop: 24,
  padBottom: 24,
  numberGap: 14,
  labelSize: 18,
  numberSizes: Object.freeze([56, 48, 40, 32]),
  maxCards: 4,
  labelMaxLines: 3,
})

/**
 * Pull the headline number out of a stat bullet: `内容交付量同比增长 138%` becomes
 * `{number: '138%', label: '内容交付量同比增长'}`. A bullet without a unit-bearing
 * number keeps its whole text as the label and gets no number line.
 *
 * @param {string} bullet
 * @returns {{number?: string, label: string}}
 */
function extractStat(bullet) {
  const flat = oneLine(bullet)

  // `从 5 天降到 4 小时` is one metric, not two: keep both sides so the card says
  // what actually changed. Checked before the single-number path.
  const range = STAT_RANGE.exec(flat)
  if (range !== null) {
    const headline = `${oneLine(range[1]).replace(/\s+/g, '')}→${oneLine(range[2]).replace(/\s+/g, '')}`
    const label = oneLine(`${flat.slice(0, range.index)} ${flat.slice(range.index + range[0].length)}`)
      .replace(/^[，,。；;：:、\s]+/, '')
      .replace(/[，,。；;：:、\s]+$/, '')
    return { number: headline, label: label === '' ? flat : label }
  }

  const match = STAT_NUMBER.exec(flat)
  if (match === null) return { label: flat }
  const number = oneLine(match[0]).replace(/\s+/g, '')
  const label = trimDanglingConnective(
    oneLine(`${flat.slice(0, match.index)} ${flat.slice(match.index + match[0].length)}`)
      .replace(/^[，,。；;：:、\s]+/, '')
      .replace(/[，,。；;：:、\s]+$/, ''),
  )
  return { number, label: label === '' ? flat : label }
}

/**
 * A number is often glued to a connective: `成本降至三分之一` leaves the label `成本降至`,
 * which is half a sentence. Strip a dangling connective so the card reads `三分之一 /
 * 成本`. Only trailing particles are removed, never a word that carries meaning, and an
 * empty result is rejected by the caller (which then keeps the whole bullet).
 *
 * @param {string} label
 */
function trimDanglingConnective(label) {
  let text = label
  const particles = ['个百分点', '降低到', '下降到', '减少到', '压缩到', '提升到', '增长到', '缩短到', '降至', '降到', '减至', '提至', '达至', '达到', '降至', '至', '到', '从', '为', '是', '达', '约', '近', '超', '过', '等于']
  let changed = true
  while (changed) {
    changed = false
    for (const particle of particles) {
      if (!text.endsWith(particle)) continue
      const next = text.slice(0, text.length - particle.length).trim()
      // Never strip the label down to nothing: `达到` alone is still copy.
      if (next !== '') {
        text = next
        changed = true
        break
      }
    }
  }
  return text
}

/**
 * Breathing slide: one oversized statement and one short caption, centred, with no
 * title bar and no bullets.
 *
 * This is the density contrast a human designer uses to let a dense deck breathe. The
 * page is deliberately almost empty — two text elements and nothing else — and it only
 * appears when `style.breather` is on *and* a concrete number already exists in the
 * surrounding slides, so it never invents data to fill space.
 *
 * @param {string} slideId
 * @param {LocalContext} ctx
 * @param {{title: string, bullets: string[]}} item  bullets = [headline, caption]
 * @returns {import('./schema.js').CardSlide}
 */
function buildBreatherSlide(slideId, ctx, item) {
  const { box, palette, toneSpec } = ctx
  const builder = makeSlideBuilder(slideId)
  const headline = oneLine(item.bullets[0] ?? '')
  const caption = oneLine(item.bullets[1] ?? '')

  // A breather can afford scale: it is the only thing on the page. The size still has
  // to respect the character count — a long statement at 96px runs off the canvas —
  // so it steps down through the same ladder the body text uses (the reference's
  // "double-sided character lock": cap the copy upstream, adapt the size downstream).
  const headlineChars = Array.from(headline).length
  const sizeLadder = [96, 84, 72, 64, 56, 48]
  const headlineSize =
    sizeLadder.find((size) => {
      const perLine = Math.floor(box.w / (size * 0.62))
      const lines = Math.ceil(headlineChars / Math.max(1, perLine))
      return lines * size * 1.2 <= box.h * 0.55
    }) ?? 48
  const headlineSizeFinal = Math.round(headlineSize)
  const captionSize = Array.isArray(TYPE.body) ? Math.min(...TYPE.body) : TYPE.body
  const headlineH = Math.round(textHeight(headline, box.w, headlineSizeFinal))
  const captionH = caption === '' ? 0 : Math.round(textHeight(caption, box.w, captionSize))
  const gap = caption === '' ? 0 : 24
  const top = Math.round(box.y + Math.max(0, (box.h - (headlineH + gap + captionH)) / 2))

  builder.text(
    'title',
    headline,
    { x: box.x, y: top, w: box.w, h: headlineH },
    textStyle(headlineSizeFinal, palette[toneSpec.sectionTitle], {
      bold: true,
      align: 'center',
      valign: 'middle',
    }),
  )
  if (caption !== '') {
    builder.text(
      'caption',
      caption,
      { x: box.x, y: top + headlineH + gap, w: box.w, h: captionH },
      textStyle(captionSize, palette[toneSpec.sub], { align: 'center', valign: 'top' }),
    )
  }

  return slideShell(slideId, builder, palette, undefined, item.notes)
}

/**
 * Stats slide: 2–4 cards in a row, the extracted number big on top and the remaining
 * words as the label underneath. The card width comes from the count (2/3/4 all fit
 * the safe width); `bulletLayout` and density do not apply — the row is a band, and a
 * 6-card band would not be readable at any density (the extras spill to `（续）`).
 *
 * @param {string} slideId
 * @param {LocalContext} ctx
 * @param {{title: string, bullets: string[], notes?: string}} item
 * @returns {import('./schema.js').CardSlide}
 */
function buildStatsSlide(slideId, ctx, item) {
  const { box, palette, style, toneSpec, warnings } = ctx
  const builder = makeSlideBuilder(slideId)
  const header = buildHeader(builder, ctx, item.title)
  const count = Math.max(1, Math.min(STATS.maxCards, item.bullets.length))
  const cardW = Math.round((box.w - STATS.gapX * (count - 1)) / count)
  const textW = cardW - STATS.padX * 2

  /** @type {Array<{number?: string, label: string, numberSize: number, numberH: number, labelH: number}>} */
  const cards = item.bullets.slice(0, count).map((bullet) => {
    const stat = extractStat(bullet)
    const numberSize = stat.number === undefined ? 0 : chooseSize(STATS.numberSizes, stat.number, textW, 1)
    const numberH = stat.number === undefined ? 0 : textHeight(stat.number, textW, numberSize)
    return { ...stat, numberSize, numberH, labelH: textHeight(stat.label, textW, STATS.labelSize) }
  })

  /** Height of the tallest card: number + gap + label, plus padding. */
  const measure = () =>
    STATS.padTop +
    Math.max(...cards.map((card) => card.numberH + (card.numberH > 0 ? STATS.numberGap : 0) + card.labelH)) +
    STATS.padBottom
  let cardH = measure()
  const available = box.y + box.h - header.top
  if (cardH > available) {
    // Hand-built outlines only: labels are capped at `labelMaxLines`, full text kept in
    // the notes. Numbers are single-line by construction (`chooseSize(..., 1)`).
    warnings.push(`${header.label} · 指标说明过长，已按 ${STATS.labelMaxLines} 行截断（原文进入演讲备注）`)
    for (const card of cards) {
      const fitted = clampToLines(card.label, textW, STATS.labelSize, STATS.labelMaxLines)
      if (fitted.truncated) item.notes = mergeNotes(item.notes, [card.label])
      card.label = fitted.text
      card.labelH = textHeight(card.label, textW, STATS.labelSize)
    }
    cardH = measure()
  }

  // Degrade honestly: a stat row reads as data, so every card must carry a number. If
  // some bullet has none, the page is really a list — render it as one instead of
  // leaving a tall empty card beside the figures.
  const noNumberCards = cards.filter((card) => card.number === undefined)
  if (noNumberCards.length > 0) {
    warnings.push(
      `${header.label} · 指标页需要每张卡都有数字，已改为列表呈现（缺数字：${clampChars(noNumberCards[0].label, 12)}）`,
    )
    return buildBulletSlide(slideId, ctx, {
      title: item.title,
      bullets: item.bullets,
      ...(item.notes === undefined ? {} : { notes: item.notes }),
    })
  }

  // Centre the band under the title (never above it).
  const top = centeredBandTop(box, header.top, cardH)
  const numberColor = style.tone === 'minimal' ? palette.textPrimary : palette.primary

  cards.forEach((card, index) => {
    const x = box.x + index * (cardW + STATS.gapX)
    if (toneSpec.cardFill !== null) {
      builder.shape('card', 'roundRect', { x, y: top, w: cardW, h: cardH }, {
        fill: palette[toneSpec.cardFill],
        radius: 18,
      })
    }
    if (card.number !== undefined) {
      builder.text('subtitle', card.number, {
        x: x + STATS.padX,
        y: top + STATS.padTop,
        w: textW,
        h: card.numberH,
      }, textStyle(card.numberSize, numberColor, { bold: true }), 'stat')
    }
    // Labels are bottom-aligned in the row: a bullet without a number has no number
    // line, and its label must still line up with the labels of its neighbours.
    builder.text('body', card.label, {
      x: x + STATS.padX,
      y: top + cardH - STATS.padBottom - card.labelH,
      w: textW,
      h: card.labelH,
    }, textStyle(STATS.labelSize, palette[toneSpec.sub]), 'label')
  })

  return slideShell(slideId, builder, palette, header.text, item.notes)
}

/**
 * Timeline slide: a horizontal band of numbered dots joined by segments, one short
 * caption under each dot.
 *
 * Nodes are the first `TIMELINE_MAX_NODES` bullets (extra bullets are folded into the
 * last caption as `+N` and preserved in the notes by {@link foldTimeline}). Captions
 * are clamped to `captionMaxLines` — a timeline label is an abbreviation, never the
 * paragraph that used to sit on the page — so a long caption ellipsizes instead of
 * pushing the band out of the safe area. Two bullets are a legitimate timeline (2
 * nodes), one bullet degrades to a single centred node.
 *
 * @param {string} slideId
 * @param {LocalContext} ctx
 * @param {{title: string, bullets: string[], notes?: string}} item
 * @returns {import('./schema.js').CardSlide}
 */
function buildTimelineSlide(slideId, ctx, item) {
  const { box, palette, style, toneSpec, warnings } = ctx
  const builder = makeSlideBuilder(slideId)
  const header = buildHeader(builder, ctx, item.title)
  const all = item.bullets.map((bullet) => oneLine(bullet)).filter((bullet) => bullet !== '')
  if (all.length === 0) {
    return buildSectionSlide(slideId, ctx, { title: header.text, notes: item.notes })
  }
  const nodes = all.slice(0, TIMELINE_MAX_NODES)
  const hidden = all.length - nodes.length
  const count = nodes.length
  const slotW = box.w / count
  // A caption may use its slot minus a 24px breathing gap, capped so a 2-node page
  // does not produce two 500px sentences.
  const captionW = Math.round(Math.min(slotW - 24, TIMELINE.captionMaxW))

  let truncated = 0
  const captions = nodes.map((bullet, index) => {
    const bare = oneLine(bullet.replace(TIMELINE_PREFIX, ''))
    let text = bare === '' ? oneLine(bullet) : bare
    if (hidden > 0 && index === nodes.length - 1) text = `${text} +${hidden}`
    // Character cap first (page budget), then the line cap (layout).
    const capped =
      Array.from(text).length > TIMELINE.captionMaxChars
        ? clampChars(text, TIMELINE.captionMaxChars)
        : text
    const fitted = clampToLines(capped, captionW, TIMELINE.captionSize, TIMELINE.captionMaxLines)
    if (fitted.truncated || capped !== text) truncated += 1
    return { text: fitted.text, h: textHeight(fitted.text, captionW, TIMELINE.captionSize) }
  })
  if (truncated > 0) {
    warnings.push(`${header.label} · 时间线节点文案过长，已省略 ${truncated} 处（保留前 ${TIMELINE.captionMaxLines} 行）`)
  }

  const maxCaptionH = Math.max(...captions.map((caption) => caption.h))
  const bandH = TIMELINE.dotSize + TIMELINE.captionGap + maxCaptionH
  const bandTop = centeredBandTop(box, header.top, bandH)
  const dotY = bandTop
  const captionTop = bandTop + TIMELINE.dotSize + TIMELINE.captionGap
  const indexH = Math.ceil(TIMELINE.indexSize * 1.35)
  const indexColor = style.tone === 'minimal' ? palette.textPrimary : palette.background
  const dotStyle = style.tone === 'minimal'
    ? { fill: palette.background, stroke: palette.textPrimary, strokeWidth: 2 }
    : { fill: palette.primary }

  /** Where each dot sits, so the connector can be drawn between them. */
  const centers = nodes.map((_, index) => box.x + slotW * (index + 0.5))

  nodes.forEach((_, index) => {
    const dotX = Math.round(centers[index] - TIMELINE.dotSize / 2)
    builder.shape('node', 'ellipse', { x: dotX, y: dotY, w: TIMELINE.dotSize, h: TIMELINE.dotSize }, dotStyle)
    // The index lives INSIDE the dot: its box must be contained by the ellipse's
    // bounding box, otherwise the "text straddles a shape" rule would fire.
    builder.text('body', String(index + 1), {
      x: dotX + Math.round((TIMELINE.dotSize - TIMELINE.indexW) / 2),
      y: dotY + Math.round((TIMELINE.dotSize - indexH) / 2),
      w: TIMELINE.indexW,
      h: indexH,
    }, textStyle(TIMELINE.indexSize, indexColor, { bold: true, align: 'center', valign: 'middle' }), 'index')
    builder.text('body', captions[index].text, {
      x: clamp(Math.round(centers[index] - captionW / 2), box.x, box.x + box.w - captionW),
      y: captionTop,
      w: captionW,
      h: captions[index].h,
    }, textStyle(TIMELINE.captionSize, palette.textPrimary, { align: 'center' }), 'caption')
  })

  // Connectors are drawn per gap, not as one long line: a line through the dots would
  // intersect every dot in the middle (shape/shape overlap) instead of joining them.
  for (let index = 0; index < count - 1; index += 1) {
    const from = Math.round(centers[index] + TIMELINE.dotSize / 2 + TIMELINE.segGap)
    const to = Math.round(centers[index + 1] - TIMELINE.dotSize / 2 - TIMELINE.segGap)
    if (to <= from) continue
    builder.shape('seg', 'rect', {
      x: from,
      y: Math.round(dotY + (TIMELINE.dotSize - TIMELINE.segH) / 2),
      w: to - from,
      h: TIMELINE.segH,
    }, { fill: palette.surface })
  }

  return slideShell(slideId, builder, palette, header.text, item.notes)
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
 * A page that would otherwise be half empty steps *up* through `SPARSE_BODY_SIZES`,
 * spreads its leftover height into the gaps, and is finally centred by at most
 * `BLOCK_CENTER_MAX` — sparse copy should carry the page, not hug the title.
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
  let blockTop = top

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
  } else if (bullets.length > 0) {
    const spec = DENSITY[style.density]
    // Sparse page: four short bullets at 22px hug the title and leave the bottom 40%
    // of the slide empty, which reads as an unfinished page. Step the body size up
    // first (the copy is short, so it can carry the page), then spread the leftover
    // into the gap up to `gapMax`, and finally centre the block so what remains is
    // split above and below instead of pooling at the bottom.
    for (const size of SPARSE_BODY_SIZES) {
      if (size <= fontSize) break
      const candidateHeights = bullets.map((bullet) => textHeight(bullet, width, size))
      if (totalHeight(candidateHeights, spec.gapMax) > available) continue
      fontSize = size
      heights = candidateHeights
    }
    total = totalHeight(heights, gap)
    if (bullets.length > 1) {
      gap = Math.min(spec.gapMax, gap + Math.floor((available - total) / (bullets.length - 1)))
      total = totalHeight(heights, gap)
    }
    // Centring is bounded: a sparse page should breathe, not float away from its title.
    const offset = Math.min(BLOCK_CENTER_MAX, Math.max(0, Math.floor((available - total) / 2)))
    blockTop = top + offset
  }

  let y = blockTop
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
 * Split a title's bullets into list slides.
 *
 * The density cap is a *preference*, not a hard limit: six short bullets at 22px fit a
 * 720px page with room to spare, and splitting them into 4+2 produces a second page
 * whose void is the most visible defect a reader notices. So the cap is stretched (by
 * one at a time, never past `LIST_HARD_MAX`) until every page carries at least
 * `LIST_MIN_PAGE` bullets — while still refusing any page that would not fit its height.
 *
 * @param {string[]} bullets
 * @param {number} cap
 * @param {LocalContext} ctx
 * @param {string} title
 * @returns {string[][]}
 */
function planListChunks(bullets, cap, ctx, title) {
  const { box } = ctx
  const width = FRAME.right - (ctx.toneSpec.marker ? box.x + 30 : box.x)
  const gap = DENSITY[ctx.style.density].gap
  const bottom = box.y + box.h
  /** @param {number} count already-built pages */
  const topFor = (count) => bulletContentTop(ctx, count === 0 ? title : `${title}（续）`)
  const fits = (list, count) => topFor(count) + totalHeight(list.map((bullet) => textHeight(bullet, width, TYPE.body[0])), gap) <= bottom

  /** @param {number} width cap for one page */
  const plan = (limit) => {
    /** @type {string[][]} */
    const pages = []
    /** @type {string[]} */
    let current = []
    for (const bullet of bullets) {
      const next = [...current, bullet]
      if (current.length > 0 && (current.length >= limit || !fits(next, pages.length))) {
        pages.push(current)
        current = [bullet]
      } else {
        current = next
      }
    }
    if (current.length > 0) pages.push(current)
    return pages
  }

  let best = plan(cap)
  // Only a thin *tail* justifies stretching the cap: 40 bullets at a cap of 4 are ten
  // full pages and must stay ten pages. Six bullets (4+2) and five (4+1) should not
  // split at all when the copy fits one page.
  if (best.length > 1 && best.some((page) => page.length < LIST_MIN_PAGE)) {
    for (let limit = cap + 1; limit <= LIST_HARD_MAX; limit += 1) {
      const candidate = plan(limit)
      if (candidate.every((page) => page.length >= LIST_MIN_PAGE)) {
        best = candidate
        break
      }
    }
  }
  return best.length === 0 ? [[]] : best
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
    // Rhythm flags. This function is the generator's own gate: styles.js validates the
    // stored profile, but whatever reaches composeDeck has to survive *this* normaliser,
    // or the feature is silently off (that mismatch cost us one round already).
    varyLayout: src.varyLayout !== false,
    breather: src.breather === true,
    breatherEvery:
      Number.isFinite(src.breatherEvery) && src.breatherEvery >= 3
        ? Math.min(12, Math.floor(src.breatherEvery))
        : DEFAULT_STYLE.breatherEvery,
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

/**
 * Derive a stable, path-safe deck id from a title.
 *
 * The ASCII remnant of a mostly-CJK title is noise: `自建模板 vs 统一模板库` used to
 * yield `vs`, which names nothing and collides easily. So the remnant is only used
 * when it actually represents the title — at least three characters AND at least a
 * third of the alphanumeric weight. Otherwise we fall back to `deck-<digest>`, which
 * is still deterministic (same title → same id) but never meaningless.
 *
 * @param {string} title
 */
function deckIdFromTitle(title) {
  const text = String(title ?? '')
  const base = text
    .normalize('NFKD')
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/[-._]{2,}/g, '-')
    .replace(/^[-._]+|[-._]+$/g, '')
    .slice(0, 48)

  const letters = base.replace(/[^a-z0-9]/g, '')
  // Weight = alphanumeric characters in the title vs those that survived as ASCII.
  const titleWeight = [...text].filter((char) => /[\p{L}\p{N}]/u.test(char)).length
  const coverage = titleWeight === 0 ? 1 : letters.length / titleWeight
  if (letters.length >= 3 && coverage >= 1 / 3) return letters.slice(0, 48)
  return `deck-${digest(text)}`
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
     * @param {string} role one of the four `TextRoleSchema` values (it also picks the
     *   font face: title/subtitle use the display face, body/caption the text face)
     * @param {string} content
     * @param {{x: number, y: number, w: number, h: number}} geo
     * @param {any} style
     * @param {string} [idSeg] optional distinct id segment, for layouts that place
     *   several elements of the same role (a stat card's number and label, a
     *   timeline's index and caption)
     */
    text(role, content, geo, style, idSeg) {
      elements.push({
        type: /** @type {const} */ ('text'),
        id: nextId(idSeg === undefined ? role : idSeg),
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

/**
 * Top edge for a band layout (stats row, timeline) that is centred in the space below
 * the title but can never start above it — on a full page it simply starts at
 * `contentTop`.
 *
 * @param {ReturnType<typeof innerBox>} box
 * @param {number} contentTop
 * @param {number} bandH
 */
function centeredBandTop(box, contentTop, bandH) {
  const available = Math.max(0, box.y + box.h - contentTop)
  return contentTop + Math.max(0, Math.floor((available - bandH) / 2))
}

/**
 * Append extra note lines to the notes already on a slide, never replacing or
 * shortening them.
 *
 * @param {string | undefined} notes
 * @param {string[]} additions
 * @returns {string}
 */
function mergeNotes(notes, additions) {
  const base = typeof notes === 'string' ? notes : ''
  const extra = additions.filter((line) => line !== '').join('\n')
  if (base === '') return extra
  return extra === '' ? base : `${base}\n${extra}`
}

/** @param {string} text */
function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
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
  // A tail of one is the only split worth disturbing: the last page would carry a
  // single bullet and read as an empty slide. Case-by-case rebalancing (not an even
  // redistribution) keeps every other page's grouping intact.
  if (out.length > 1 && out[out.length - 1].length === 1) {
    const tail = out[out.length - 1]
    const previous = out[out.length - 2]
    tail.unshift(previous.pop())
  }
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
 * A model writing an outline often labels the slide's *role* instead of its content —
 * `封面：把合规演示从 5 天压到半小时`, `第 3 页：模板解析`, `Slide 4: Export`. The deck
 * already has a cover, so that prefix is noise on screen and, worse, it makes the page
 * claim to be something it is not. Strip the label (it must be followed by a separator,
 * so `封面设计的三条原则` is left alone).
 *
 * @param {string} title
 */
function cleanSlideTitle(title) {
  const text = String(title ?? '').trim()
  const stripped = text
    .replace(/^(?:封面|扉页|目录|首页|结束页|封底)\s*[:：、]\s*/, '')
    .replace(/^第\s*[0-9一二三四五六七八九十]+\s*(?:页|张|屏)\s*[:：、]\s*/, '')
    .replace(/^(?:slide|page|p)\s*[0-9]+\s*[:：、]\s*/i, '')
    .trim()
  return stripped === '' ? text : stripped
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