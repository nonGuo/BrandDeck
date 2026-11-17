/**
 * BrandDeck shared data contract.
 *
 * This module is the single source of truth for the deck DSL that the model
 * authors, the Web panel edits and the PPTX compiler consumes. Everything here
 * is pure data + pure functions (no I/O) so it can be unit-tested in isolation.
 *
 * Coordinate system: a 1280x720 virtual canvas (16:9). The compiler converts
 * px -> inches with `px / 96` so a 1280x720 deck maps exactly onto
 * PowerPoint's `LAYOUT_WIDE` (13.333in x 7.5in).
 */
import { z } from 'zod'

import { mermaidToSvg } from './mermaid.js'

export const CANVAS = Object.freeze({ width: 1280, height: 720 })

/** Margin kept free on every side; the top-right corner additionally reserves the logo. */
export const SAFE_AREA = Object.freeze({ top: 64, right: 64, bottom: 64, left: 64 })

/** Smallest margin a template-derived content box may claim (sanity floor, not a design rule). */
export const FRAME_MIN_MARGIN = 24

/** Logo reservation inside the top-right corner of the safe area. */
export const LOGO_RESERVE = Object.freeze({ width: 150, height: 48, gap: 16 })

/**
 * Default page text budget, in characters. It is the *default* rather than the law:
 * a deployment can move it (`slideCharBudget` in the plugin config, or per request),
 * and `repairDeck` takes the effective value as an option. Kept exported so callers and
 * the docs have one source of truth for the number.
 */
export const SLIDE_TEXT_BUDGET = 150

/** Clamp a caller-supplied page budget into a range that still means something. */
export function resolveSlideCharBudget(value) {
  const number = Number(value)
  if (!Number.isFinite(number)) return SLIDE_TEXT_BUDGET
  return Math.min(600, Math.max(40, Math.floor(number)))
}

export const elementBaseShape = {
  id: z.string().min(1).max(64),
  x: z.number().finite(),
  y: z.number().finite(),
  w: z.number().finite(),
  h: z.number().finite(),
}

const colorValue = z.string().min(1).max(64)

export const TextRoleSchema = z.enum(['title', 'subtitle', 'body', 'caption'])
export const ShapeTypeSchema = z.enum(['rect', 'roundRect', 'ellipse', 'line', 'rightArrow'])
export const ChartTypeSchema = z.enum(['bar', 'line', 'pie'])

export const TextElementSchema = z.object({
  type: z.literal('text'),
  ...elementBaseShape,
  role: TextRoleSchema,
  content: z.string(),
  style: z
    .object({
      fontSize: z.number().finite().positive(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      color: colorValue,
      align: z.enum(['left', 'center', 'right']).optional(),
      valign: z.enum(['top', 'middle', 'bottom']).optional(),
      lineSpacing: z.number().finite().positive().optional(),
    })
    .strict(),
})

export const ShapeElementSchema = z.object({
  type: z.literal('shape'),
  ...elementBaseShape,
  shapeType: ShapeTypeSchema,
  style: z
    .object({
      fill: colorValue.optional(),
      stroke: colorValue.optional(),
      strokeWidth: z.number().finite().nonnegative().optional(),
      radius: z.number().finite().nonnegative().optional(),
      opacity: z.number().min(0).max(1).optional(),
    })
    .strict(),
})

export const SvgElementSchema = z.object({
  type: z.literal('svg'),
  ...elementBaseShape,
  description: z.string().optional(),
  svgCode: z.string().min(1),
})

/**
 * A diagram authored as text (Mermaid flowchart) rather than as SVG.
 *
 * The model is good at `A --> B` and bad at "put the box at x=312". `repairDeck` compiles
 * this into an ordinary `svg` element with the theme's colours baked in, so the panel,
 * the rasteriser and the PPTX compiler never need to know a diagram existed.
 */
export const DiagramElementSchema = z.object({
  type: z.literal('diagram'),
  ...elementBaseShape,
  description: z.string().optional(),
  /** Mermaid source; the supported subset is documented in src/mermaid.js. */
  mermaid: z.string().min(1),
})

export const ChartElementSchema = z.object({
  type: z.literal('chart'),
  ...elementBaseShape,
  chartType: ChartTypeSchema,
  data: z.object({
    labels: z.array(z.string()).min(1).max(24),
    series: z
      .array(
        z.object({
          name: z.string().min(1),
          values: z.array(z.number().finite()).min(1).max(24),
        }),
      )
      .min(1)
      .max(6),
  }),
  style: z
    .object({
      legend: z.boolean().optional(),
      showValues: z.boolean().optional(),
    })
    .strict()
    .optional(),
})

export const ImageElementSchema = z.object({
  type: z.literal('image'),
  ...elementBaseShape,
  assetKey: z.string().min(1),
  alt: z.string().optional(),
})

export const SlideElementSchema = z.discriminatedUnion('type', [
  TextElementSchema,
  ShapeElementSchema,
  SvgElementSchema,
  DiagramElementSchema,
  ChartElementSchema,
  ImageElementSchema,
])

/**
 * Layout intent of a slide.
 *
 * The generator states it, the text outline round-trips it as a `版式：…` line, and
 * the compiler/panel may use it. It has to be a first-class DSL field: as an unknown
 * key it was silently stripped on save (Zod drops what it does not declare), so every
 * content slide ended up rendered as the same bullet list.
 *
 * Absent means "infer from the elements", which keeps decks written before this
 * field existed valid.
 */
export const SlideKindSchema = z.enum([
  'cover',
  'section',
  'toc',
  'bullets',
  'cards',
  'compare',
  'timeline',
  'stats',
  /** A page built around a figure (Mermaid-compiled architecture/flow diagram). */
  'diagram',
  /** 2–4 tinted panels, one per point (the reference decks' agenda/three-point page). */
  'columns',
  /** One or two oversized numbers with a label and context. */
  'figures',
  /** Labelled rows with hairlines, the first row marked in the accent colour. */
  'rankList',
  /** A native chart driven by the page's own values (bar / line / pie). */
  'chart',
  /** Generator-only pacing page: one oversized statement, no title, no bullets. */
  'breather',
])

export const CardSlideSchema = z.object({
  slideId: z.string().min(1).max(64),
  title: z.string().optional(),
  /** Layout intent; absent means "infer from the elements" (legacy decks). */
  kind: SlideKindSchema.optional(),
  background: colorValue.optional(),
  speakerNotes: z.string().optional(),
  elements: z.array(SlideElementSchema).min(1).max(40),
})

export const CardDeckSchema = z.object({
  deckId: z.string().min(1).max(64),
  title: z.string().min(1).max(200),
  templateId: z.string().min(1).max(64),
  slides: z.array(CardSlideSchema).min(1).max(60),
})

/** Where the template places one piece of background art, in inches. */
const frameInches = z.object({
  x: z.number().finite(),
  y: z.number().finite(),
  w: z.number().finite(),
  h: z.number().finite(),
})

/**
 * Background art reference: the extracted file plus the frame the template draws it in.
 * A bare string is accepted for forward/backward compatibility and treated as full-bleed.
 */
const backdropRef = z.union([
  z.string().min(1),
  z.object({ path: z.string().min(1), frame: frameInches.optional() }),
])

export const ThemeTokensSchema = z.object({
  templateId: z.string().min(1),
  templateName: z.string().min(1),
  colors: z.object({
    background: colorValue,
    surface: colorValue,
    textPrimary: colorValue,
    textSecondary: colorValue,
    primary: colorValue,
    accent: colorValue,
    chartPalette: z.array(colorValue).min(1).max(8),
  }),
  fonts: z.object({
    title: z.string().min(1),
    body: z.string().min(1),
    latin: z.string().min(1),
  }),
  brandAssets: z
    .object({
      logoPath: z.string().optional(),
      logoWidthPx: z.number().int().positive().optional(),
      logoHeightPx: z.number().int().positive().optional(),
      /**
       * Where the template's master actually places the logo, in inches. Not decorative
       * detail: a corporate master puts its lockup bottom-right (10.07in, 6.54in in the
       * Huawei template), and drawing it in a top-right corner instead is the single most
       * obvious "this was not made with our template" tell.
       */
      logoFrame: z
        .object({
          x: z.number().finite(),
          y: z.number().finite(),
          w: z.number().finite(),
          h: z.number().finite(),
        })
        .optional(),
      /**
       * The content box the template itself uses, in **canvas px** (1280×720). Derived from
       * a content layout's `BODY` placeholder, so a deck generated from a template keeps that
       * template's margins instead of our built-in 64px guess. Absent ⇒ 64px safe area.
       */
      frame: z
        .object({
          x: z.number().finite(),
          y: z.number().finite(),
          w: z.number().finite(),
          h: z.number().finite(),
        })
        .optional(),
      /**
       * Full-bleed background pictures extracted from the template's own layouts.
       *
       * `cover` is the layout art the template uses behind a title, `section` the variant
       * used for chapter breaks, `closing` the end page. Paths are absolute, like `logoPath`,
       * and the files are written next to `logo.png` at parse time.
       */
      backdrops: z
        .object({
          cover: backdropRef.optional(),
          section: backdropRef.optional(),
          closing: backdropRef.optional(),
        })
        .optional(),
      /**
       * Fixed footer strings the template repeats on every page (a confidentiality marker, a
       * security-level label). Inherited rather than invented: a template without them gets
       * none, because a fabricated "机密" line is worse than an empty footer.
       */
      footlines: z.array(z.string()).max(4).optional(),
      icons: z.record(z.string(), z.string()).default({}),
    })
    .default({ icons: {} }),
})

/** @typedef {import('zod').infer<typeof CardDeckSchema>} CardDeck */
/** @typedef {import('zod').infer<typeof CardSlideSchema>} CardSlide */
/** @typedef {import('zod').infer<typeof SlideElementSchema>} SlideElement */
/** @typedef {import('zod').infer<typeof ThemeTokensSchema>} ThemeTokens */

const HEX = /^#[0-9a-fA-F]{3,8}$/
const TOKEN_REF = /^\$?colors\.([a-zA-Z0-9_-]+)$/

/**
 * Resolve a color that may be a HEX literal or a theme token reference
 * (`$colors.primary`, `colors.primary`, `$accent`).
 * @param {string | undefined} value
 * @param {ThemeTokens} theme
 * @param {string} fallback
 * @returns {string} `#RRGGBB` (alpha preserved when the source carried it)
 */
export function resolveColor(value, theme, fallback = '#000000') {
  if (typeof value !== 'string' || value.length === 0) return fallback
  const raw = value.trim()
  if (HEX.test(raw)) return normalizeHex(raw)
  let match = TOKEN_REF.exec(raw)
  if (match === null) {
    const short = /^\$?([a-zA-Z0-9_-]+)$/.exec(raw)
    if (short !== null && short[1] in theme.colors) match = [, short[1]]
  }
  if (match !== null) {
    const key = match[1]
    const palette = theme.colors
    if (key in palette && typeof palette[key] === 'string') return normalizeHex(palette[key])
    const paletteMatch = /^chart([1-9])$/.exec(key)
    if (paletteMatch !== null) {
      const hex = Array.isArray(palette.chartPalette)
        ? palette.chartPalette[Number(paletteMatch[1]) - 1]
        : undefined
      if (typeof hex === 'string') return normalizeHex(hex)
    }
  }
  return fallback
}

/**
 * Normalize `#abc` / `abc` / `AABBCC` into `#AABBCC` (or `#AARRGGBB` when alpha is given).
 * @param {string} value
 * @returns {string}
 */
export function normalizeHex(value) {
  let hex = String(value).trim()
  if (!hex.startsWith('#')) hex = `#${hex}`
  const body = hex.slice(1)
  if (body.length === 3) {
    return `#${body[0]}${body[0]}${body[1]}${body[1]}${body[2]}${body[2]}`.toUpperCase()
  }
  return hex.toUpperCase()
}

/**
 * Split `#RRGGBB` / `#AARRGGBB` into a bare 6-digit hex plus an opacity, the two
 * values PowerPoint APIs want separately.
 * @param {string} value
 * @returns {{ hex: string, alpha: number }}
 */
export function splitAlpha(value) {
  const normalized = normalizeHex(value)
  const body = normalized.slice(1)
  if (body.length === 8) {
    const alpha = parseInt(body.slice(0, 2), 16) / 255
    return { hex: `#${body.slice(2)}`, alpha }
  }
  return { hex: normalized, alpha: 1 }
}

/**
 * The drawing area every element must stay inside: the template's own content box when it
 * has one, else the built-in 64px safe area.
 *
 * Why prefer the template's box: a corporate deck has already decided where content lives
 * (the Huawei content page puts its body at 0.81in/1.64in, width 11.74in — 78px/157px in
 * canvas px, not 64px/64px). Drawing inside our guess instead makes every page sit slightly
 * off the template's own grid, which reads as "pasted in".
 *
 * The frame is clamped to the safe area's bounds so a hostile or half-parsed template can
 * never push content off the canvas.
 *
 * @param {ThemeTokens} theme
 * @returns {{x: number, y: number, w: number, h: number, logo: {width: number, height: number, gap: number} | null}}
 */
export function innerBox(theme) {
  const reserveLogo = Boolean(theme?.brandAssets?.logoPath)
  const fallback = {
    x: SAFE_AREA.left,
    y: SAFE_AREA.top,
    w: CANVAS.width - SAFE_AREA.left - SAFE_AREA.right,
    h: CANVAS.height - SAFE_AREA.top - SAFE_AREA.bottom,
  }
  const frame = theme?.brandAssets?.frame
  const usable =
    frame !== undefined &&
    frame !== null &&
    [frame.x, frame.y, frame.w, frame.h].every((value) => Number.isFinite(value)) &&
    frame.w >= 480 &&
    frame.h >= 240
  if (usable === false) return { ...fallback, logo: reserveLogo ? LOGO_RESERVE : null }

  // Clamp to a *sanity* floor, not to our own 64px guess: a corporate template may legally
  // place its title at 0.50in (48px) — the Huawei content page does — and clamping that up
  // to 64px would silently re-break the alignment this frame exists to provide. 24px keeps
  // content off the physical edge; the logo reserve is enforced separately by repairDeck.
  const floor = FRAME_MIN_MARGIN
  const x = Math.max(floor, Math.round(/** @type {any} */ (frame).x))
  const y = Math.max(floor, Math.round(/** @type {any} */ (frame).y))
  const right = Math.min(CANVAS.width - floor, x + Math.round(/** @type {any} */ (frame).w))
  const bottom = Math.min(CANVAS.height - floor, y + Math.round(/** @type {any} */ (frame).h))
  if (right - x < 480 || bottom - y < 240) return { ...fallback, logo: reserveLogo ? LOGO_RESERVE : null }
  return { x, y, w: right - x, h: bottom - y, logo: reserveLogo ? LOGO_RESERVE : null }
}

const clamp = (value, min, max) => Math.min(max, Math.max(min, value))

function roundRect(elements) {
  for (const element of elements) {
    element.x = Math.round(element.x)
    element.y = Math.round(element.y)
    element.w = Math.round(element.w)
    element.h = Math.round(element.h)
  }
}

/**
 * Repair a deck in place-ish (returns a new deck) so that downstream renderers and
 * the compiler never see out-of-bounds geometry. Each repair is reported as a
 * human-readable warning; the UI surfaces them and the agent can fix the deck.
 *
 * @param {CardDeck} deck
 * @param {ThemeTokens} theme
 * @param {{slideCharBudget?: number}} [options] page text budget (default 150 chars)
 * @returns {{ deck: CardDeck, warnings: string[] }}
 */
export function repairDeck(deck, theme, options = {}) {
  const budget = resolveSlideCharBudget(options.slideCharBudget)
  /** @type {string[]} */
  const warnings = []
  const next = structuredClone(deck)
  const box = innerBox(theme)

  next.slides.forEach((slide, slideIndex) => {
    const label = `slide ${slideIndex + 1}`
    if (typeof slide.background === 'string') {
      slide.background = resolveColor(slide.background, theme, theme.colors.background)
    }
    rotateDuplicateIds(slide, slideIndex, warnings)
    compileDiagrams(slide, theme, label, warnings)
    for (const element of slide.elements) {
      const before = { x: element.x, y: element.y, w: element.w, h: element.h }
      // Text needs a usable box; decorative shapes may legitimately be hairlines
      // (a 2px divider), so they only get a 1px floor.
      const minSide = element.type === 'text' ? 8 : 1
      if (!Number.isFinite(element.w) || element.w <= 0) element.w = 200
      if (!Number.isFinite(element.h) || element.h <= 0) element.h = 80
      element.w = Math.max(minSide, Math.min(element.w, CANVAS.width))
      element.h = Math.max(minSide, Math.min(element.h, CANVAS.height))
      if (!Number.isFinite(element.x)) element.x = box.x
      if (!Number.isFinite(element.y)) element.y = box.y
      element.x = clamp(element.x, 0, Math.max(0, CANVAS.width - element.w))
      element.y = clamp(element.y, 0, Math.max(0, CANVAS.height - element.h))
      if (
        element.x !== before.x ||
        element.y !== before.y ||
        element.w !== before.w ||
        element.h !== before.h
      ) {
        warnings.push(
          `${label} · element "${element.id}" geometry clamped ` +
            `(${fmtBox(before)} → ${fmtBox(element)})`,
        )
      }
      // Keep the logo corner readable: text that lands under it gets pushed down.
      if (box.logo !== null && element.type === 'text') {
        const logoX = CANVAS.width - SAFE_AREA.right - box.logo.width
        const overlapsLogo =
          element.x + element.w > logoX &&
          element.y < SAFE_AREA.top + box.logo.height &&
          element.role !== 'caption'
        if (overlapsLogo && element.y < SAFE_AREA.top + box.logo.height + box.logo.gap) {
          element.y = SAFE_AREA.top + box.logo.height + box.logo.gap
          warnings.push(`${label} · element "${element.id}" moved out of the logo corner`)
        }
      }
      repairElement(element, theme, label, warnings)
    }
    roundRect(slide.elements)
    const overflow = estimateTextOverflow(slide, theme)
    if (overflow.totalCharacters > budget) {
      warnings.push(
        `${label} · ${overflow.totalCharacters} characters of text ` +
          `(budget ${budget}); shorten copy to avoid overflow`,
      )
    }
    for (const box of overflow.overflowing) {
      warnings.push(`${label} · text "${box.id}" overflows its box (${box.estimatedLines} lines)`)
    }
  })

  return { deck: next, warnings }
}

function fmtBox(box) {
  return `${box.x},${box.y} ${box.w}x${box.h}`
}

/**
 * Compile `diagram` elements (Mermaid source) into plain `svg` elements.
 *
 * This happens at repair time so that "authoring" and "rendering" stay separate: the model
 * may write a flowchart, but everything downstream — the panel canvas, the 3x rasteriser,
 * the PPTX compiler — only ever sees an SVG string. A diagram the renderer cannot draw is
 * dropped with a warning that names the reason, because a deck with a missing figure is
 * recoverable while a deck that fails to save is not.
 *
 * @param {any} slide
 * @param {ThemeTokens} theme
 * @param {string} label
 * @param {string[]} warnings
 */
function compileDiagrams(slide, theme, label, warnings) {
  const elements = Array.isArray(slide.elements) ? slide.elements : []
  for (let index = 0; index < elements.length; index += 1) {
    const element = elements[index]
    if (element === null || typeof element !== 'object' || element.type !== 'diagram') continue
    try {
      const rendered = mermaidToSvg(element.mermaid, {
        width: Math.max(120, Math.round(element.w)),
        height: Math.max(80, Math.round(element.h)),
        colors: {
          textPrimary: theme.colors.textPrimary,
          secondary: theme.colors.textSecondary,
          primary: theme.colors.primary,
          background: theme.colors.background,
          surface: theme.colors.surface,
          textSecondary: theme.colors.textSecondary,
        },
        fontSize: diagramFontSize(element),
      })
      elements[index] = {
        type: 'svg',
        id: element.id,
        x: element.x,
        y: element.y,
        w: element.w,
        h: element.h,
        ...(typeof element.description === 'string' ? { description: element.description } : {}),
        svgCode: rendered.svg,
      }
      warnings.push(
        `${label} · 架构图已由 Mermaid 编译为 SVG（${rendered.nodes} 个节点 / ${rendered.edges} 条连线 / ${rendered.groups} 个分组）`,
      )
    } catch (error) {
      warnings.push(
        `${label} · 架构图无法绘制，已移除：${error instanceof Error ? error.message : String(error)}`,
      )
      elements.splice(index, 1)
      index -= 1
    }
  }
}

/** Diagram label size: shrink a little when the figure is small. */
function diagramFontSize(element) {
  const height = Number.isFinite(element?.h) ? element.h : 400
  return Math.max(10, Math.min(16, Math.round(height / 26)))
}

function rotateDuplicateIds(slide, slideIndex, warnings) {
  const seen = new Set()
  for (const element of slide.elements) {
    if (typeof element.id !== 'string' || element.id.trim() === '') {
      element.id = `s${slideIndex + 1}-${slide.elements.indexOf(element)}`
    }
    if (seen.has(element.id)) {
      const unique = `${element.id}-${seen.size}`
      warnings.push(`slide ${slideIndex + 1} · duplicate element id "${element.id}" renamed to "${unique}"`)
      element.id = unique
    }
    seen.add(element.id)
  }
}

function repairElement(element, theme, label, warnings) {
  if (element.type === 'text') {
    if (element.style === undefined) element.style = { fontSize: 20, color: '$colors.textPrimary' }
    element.style.fontSize = clamp(element.style.fontSize ?? 20, 8, 120)
    element.style.color = resolveColor(element.style.color, theme, theme.colors.textPrimary)
    element.role = element.role ?? 'body'
  } else if (element.type === 'shape') {
    if (element.style === undefined) element.style = {}
    if (element.style.fill !== undefined) {
      element.style.fill = resolveColor(element.style.fill, theme, '')
    }
    if (element.style.stroke !== undefined) {
      element.style.stroke = resolveColor(element.style.stroke, theme, '')
    }
    if (element.style.radius !== undefined && element.shapeType !== 'roundRect') {
      element.style.radius = undefined
    }
  } else if (element.type === 'svg') {
    const check = validateSvg(element.svgCode)
    if (check.problems.length > 0) {
      for (const problem of check.problems) warnings.push(`${label} · svg "${element.id}": ${problem}`)
      element.svgCode = check.sanitized
    }
  } else if (element.type === 'chart') {
    const series = element.data?.series ?? []
    const labels = element.data?.labels ?? []
    for (const one of series) {
      if (one.values.length !== labels.length) {
        const target = Math.min(one.values.length, labels.length)
        warnings.push(
          `${label} · chart "${element.id}" series "${one.name}" had ${one.values.length} values ` +
            `for ${labels.length} labels; truncated to ${target}`,
        )
        one.values = one.values.slice(0, target)
      }
    }
    if (series.length > 0 && labels.length === 0) {
      element.data.labels = series[0].values.map((_, index) => String(index + 1))
    }
  } else if (element.type === 'image') {
    if (typeof element.assetKey !== 'string' || element.assetKey.length === 0) {
      element.assetKey = 'logo'
    }
  }
}

const ALLOWED_SVG_TAGS = new Set([
  'svg',
  'g',
  'defs',
  'title',
  'desc',
  'rect',
  'circle',
  'ellipse',
  'line',
  'path',
  'polygon',
  'polyline',
  'text',
  'tspan',
  'linearGradient',
  'radialGradient',
  'stop',
  'clipPath',
  'marker',
])

/**
 * Sanitize model-authored SVG. `@resvg/resvg-js` is a pure vector renderer: it
 * cannot draw HTML, so `foreignObject`/`script`/event handlers are rejected and
 * stripped rather than exported (the design doc calls this out explicitly).
 *
 * @param {string} svgCode
 * @returns {{ sanitized: string, problems: string[] }}
 */
export function validateSvg(svgCode) {
  /** @type {string[]} */
  const problems = []
  let out = String(svgCode ?? '').trim()
  if (!out.startsWith('<svg')) {
    const start = out.indexOf('<svg')
    if (start === -1) {
      return { sanitized: '', problems: ['missing <svg> root element'] }
    }
    out = out.slice(start)
  }
  for (const forbidden of ['foreignObject', 'script', 'iframe', 'image', 'use', 'style']) {
    const re = new RegExp(`<\\s*${forbidden}\\b[\\s\\S]*?<\\s*/\\s*${forbidden}\\s*>|<\\s*${forbidden}\\b[^>]*/>`, 'i')
    if (re.test(out)) {
      problems.push(`<${forbidden}> is not supported by the SVG rasterizer and was removed`)
      out = out.replace(re, '')
    }
  }
  out = out.replace(/\son[a-zA-Z]+\s*=\s*"[^"]*"/g, '').replace(/\son[a-zA-Z]+\s*=\s*'[^']*'/g, '')
  const tagRe = /<\s*\/?\s*([a-zA-Z][a-zA-Z0-9:]*)/g
  let tag
  while ((tag = tagRe.exec(out)) !== null) {
    const name = tag[1]
    if (name.includes(':')) continue
    if (!ALLOWED_SVG_TAGS.has(name)) problems.push(`unsupported element <${name}>`)
  }
  if (!/\swidth\s*=/.test(out) || !/\sheight\s*=/.test(out)) {
    const viewBox = /\sviewBox\s*=\s*"([^"]+)"/.exec(out)
    if (viewBox !== null) {
      const parts = viewBox[1].split(/[\s,]+/).map(Number)
      if (parts.length === 4 && parts.every(Number.isFinite)) {
        out = out.replace('<svg', `<svg width="${parts[2]}" height="${parts[3]}"`)
      }
    }
  }
  if (!/\sxmlns\s*=/.test(out)) {
    out = out.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"')
  }
  return { sanitized: out, problems: [...new Set(problems)] }
}

/**
 * Rough line-count estimate used both by the panel (live warning) and by the
 * agent-facing validator. CJK glyphs are full-width, Latin glyphs roughly half.
 * @param {string} text
 * @param {number} boxWidthPx
 * @param {number} fontSizePx
 * @returns {{ lines: number, charsPerLine: number, heightPx: number }}
 */
export function estimateLines(text, boxWidthPx, fontSizePx) {
  const safeFont = Math.max(8, fontSizePx)
  let width = 0
  let explicitLines = 1
  for (const char of String(text ?? '')) {
    if (char === '\n') {
      explicitLines += 1
      width = 0
      continue
    }
    width += isWide(char) ? safeFont : safeFont * 0.55
  }
  const charsPerLine = Math.max(1, Math.floor(boxWidthPx / (safeFont * 0.55)))
  const linesFromWidth = Math.ceil(width / Math.max(1, boxWidthPx))
  const lines = Math.max(explicitLines, linesFromWidth)
  return { lines, charsPerLine, heightPx: lines * safeFont * 1.35 }
}

function isWide(char) {
  const code = char.codePointAt(0) ?? 0
  return (
    (code >= 0x1100 && code <= 0x115f) ||
    (code >= 0x2e80 && code <= 0xa4cf) ||
    (code >= 0xac00 && code <= 0xd7a3) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xfe30 && code <= 0xfe6f) ||
    (code >= 0xff00 && code <= 0xff60) ||
    (code >= 0xffe0 && code <= 0xffe6)
  )
}

/**
 * Text-overflow report for one slide: total characters (budget check) and the
 * individual boxes whose estimated rendered height exceeds their declared `h`.
 * @param {CardSlide} slide
 * @param {ThemeTokens} theme
 */
export function estimateTextOverflow(slide, theme) {
  let totalCharacters = 0
  const overflowing = []
  for (const element of slide.elements) {
    if (element.type !== 'text') continue
    totalCharacters += String(element.content ?? '').replace(/\s+/g, '').length
    const fontSize = element.style?.fontSize ?? 20
    const estimate = estimateLines(element.content, element.w, fontSize)
    if (estimate.heightPx > element.h + 1) {
      overflowing.push({
        id: element.id,
        estimatedLines: estimate.lines,
        neededHeight: Math.ceil(estimate.heightPx),
        declaredHeight: element.h,
      })
    }
  }
  return { totalCharacters, overflowing }
}

/**
 * Parse + repair in one step. Returns structured issues so the calling tool can
 * report exactly what it changed.
 * @param {unknown} input
 * @param {ThemeTokens} theme
 * @param {{slideCharBudget?: number}} [options]
 */
export function parseAndRepairDeck(input, theme, options = {}) {
  const parsed = CardDeckSchema.safeParse(input)
  if (!parsed.success) {
    return {
      ok: /** @type {const} */ (false),
      issues: parsed.error.issues.map((issue) => ({
        path: issue.path.join('.'),
        message: issue.message,
      })),
    }
  }
  const { deck, warnings } = repairDeck(parsed.data, theme, options)
  return { ok: /** @type {const} */ (true), deck, warnings }
}

/** Fallback page token values used when a template cannot be parsed. */
export const FALLBACK_THEME_PROXY = '/theme-fallback'
