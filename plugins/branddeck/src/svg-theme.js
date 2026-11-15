/**
 * SVG theme normalisation: make model-authored diagrams belong to the deck.
 *
 * WHY: an LLM writes `<svg>` with its own palette — Mermaid's default purple/green,
 * a random "modern gradient", black text on a dark brand block. Dropped into a
 * Huawei-red deck it reads as a foreign object, which is exactly the style conflict
 * we must avoid. The reference approach is to *re-skin* the SVG with the deck's own
 * design tokens before it is rasterised, rather than to trust the model's colours.
 *
 * PRINCIPLE (learned from the reference, not copied): we never invent a palette. Every
 * colour is mapped onto the theme's own tokens, and the mapping preserves the SVG's
 * *semantic intent* — "this shape was the accented one" stays accented, "this was the
 * muted one" stays muted. Fonts likewise come from the template, never from the model.
 *
 * This module is pure string work on purpose: no XML parser dependency, no I/O, and
 * the output must stay a valid SVG for `@resvg/resvg-js`.
 */

const HEX = /^#(?:[0-9a-f]{3}|[0-9a-f]{4}|[0-9a-f]{6}|[0-9a-f]{8})$/i
const RGB = /^rgba?\(\s*([\d.]+)\s*[, ]\s*([\d.]+)\s*[, ]\s*([\d.]+)\s*(?:[,/]\s*([\d.]+)\s*)?\)$/i

/**
 * @typedef {object} SvgThemeInput
 * @property {import('./schema.js').ThemeTokens} theme
 */

/**
 * Relative luminance of a hex colour (0 = black, 1 = white). Accepts #rgb/#rrggbb.
 * @param {string} hex
 * @returns {number}
 */
export function luminance(hex) {
  const parsed = parseColor(hex)
  if (parsed === null) return 0.5
  const [r, g, b] = parsed
  const channel = (value) => {
    const srgb = value / 255
    return srgb <= 0.03928 ? srgb / 12.92 : ((srgb + 0.055) / 1.055) ** 2.4
  }
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)
}

/**
 * @param {string} value
 * @returns {[number, number, number] | null}
 */
export function parseColor(value) {
  const text = String(value ?? '').trim()
  if (HEX.test(text)) {
    const body = text.slice(1)
    if (body.length === 3 || body.length === 4) {
      return [0, 1, 2].map((index) => parseInt(body[index] + body[index], 16))
    }
    return [0, 1, 2].map((index) => parseInt(body.slice(index * 2, index * 2 + 2), 16))
  }
  const rgb = RGB.exec(text)
  if (rgb !== null) {
    return [rgb[1], rgb[2], rgb[3]].map((part) => Math.max(0, Math.min(255, Math.round(Number(part)))))
  }
  return null
}

/** @param {number} value */
const clamp255 = (value) => Math.max(0, Math.min(255, Math.round(value)))

/**
 * Classify a colour into the role it should play in the theme's own vocabulary.
 *
 * We deliberately keep this coarse: the goal is a coherent slide, not colour fidelity.
 * Roles: `primary` (brand accent), `accent` (secondary accent), `text` (readable ink),
 * `surface` (light container), `background` (page), `muted` (hairlines/secondary ink).
 *
 * @param {string} value
 * @param {import('./schema.js').ThemeTokens} theme
 * @returns {'primary' | 'accent' | 'text' | 'surface' | 'background' | 'muted'}
 */
export function classifyColor(value, theme) {
  const parsed = parseColor(value)
  const colors = theme?.colors ?? {}
  // Saturated, mid-to-dark colours are almost always the model's "highlight".
  if (parsed !== null) {
    const [r, g, b] = parsed
    const max = Math.max(r, g, b)
    const min = Math.min(r, g, b)
    const chroma = (max - min) / 255
    const lum = luminance(value)
    if (chroma >= 0.25 && lum <= 0.72) {
      // A saturated highlight: keep the deck's own accent, and give the second
      // saturated colour to `accent` so multi-series diagrams stay distinguishable.
      return 'primary'
    }
    if (lum >= 0.9) return 'background'
    if (lum >= 0.72) return 'surface'
    if (lum <= 0.35) return 'text'
    return 'muted'
  }
  return 'muted'
}

/**
 * Rewrite every hard-coded colour in an SVG to the theme's tokens, and force the
 * text font family to the template's.
 *
 * Preserves structure: only attribute *values* change, plus a `font-family` on text
 * elements. `none`, `transparent`, `inherit`, `currentColor` and `url(#…)` are left
 * alone — they are semantic, not palette.
 *
 * @param {string} svg
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {{distinctAccents?: boolean}} [options]
 * @returns {{ svg: string, remapped: number, palette: Record<string, string> }}
 */
export function normalizeSvgTheme(svg, theme, options = {}) {
  const colors = theme?.colors ?? {}
  const palette = {
    primary: colors.primary ?? '#C7000A',
    accent: colors.accent ?? colors.primary ?? '#E9002F',
    text: colors.textPrimary ?? '#1D1D1A',
    surface: colors.surface ?? '#EBEBEB',
    background: colors.background ?? '#FFFFFF',
    muted: colors.textSecondary ?? '#666666',
  }
  const chartPalette = Array.isArray(colors.chartPalette) && colors.chartPalette.length > 0
    ? colors.chartPalette
    : [palette.primary]

  /**
   * Every saturated source colour becomes the SAME brand `primary`.
   *
   * This is the deliberate difference from a naive palette-cycling approach: in a
   * diagram, three boxes with three different (model-chosen) colours almost always
   * mean three nodes of ONE kind — a flow, a pipeline, a list. Painting them red,
   * magenta and gold invents a distinction the author never made, which is exactly the
   * "copy the reference too literally" failure mode. A coherent slide beats a
   * colourful one.
   *
   * Figures that genuinely need separation (chart-like multi-series shapes) opt in with
   * `options.distinctAccents === true`, and only then walk the template's chart palette.
   */
  const assigned = new Map()
  let highlightCursor = 0
  const pick = (role, source) => {
    if (role !== 'primary') return palette[role]
    if (options.distinctAccents !== true) return palette.primary
    const key = String(source).toLowerCase()
    const existing = assigned.get(key)
    if (existing !== undefined) return existing
    const slot = chartPalette[highlightCursor % chartPalette.length]
    const chosen = typeof slot === 'string' ? slot : palette.primary
    highlightCursor += 1
    assigned.set(key, chosen)
    return chosen
  }

  let remapped = 0
  let text = String(svg ?? '')

  // `fill="#abc"`, `stroke='red'`, `style="fill:#abc"`, `stop-color="#abc"` …
  text = text.replace(
    /\b(fill|stroke|stop-color|flood-color|color)\s*=\s*("|')([^"']*)\2/gi,
    (whole, attribute, quote, value) => {
      const trimmed = String(value).trim()
      if (trimmed === '' || /^(none|transparent|inherit|currentcolor|url\()/i.test(trimmed)) return whole
      if (!HEX.test(trimmed) && !RGB.test(trimmed) && !/^[a-z]+$/i.test(trimmed)) return whole
      if (/^[a-z]+$/i.test(trimmed) && !/^(black|white|red|green|blue|gray|grey|orange|purple|yellow)$/i.test(trimmed)) {
        return whole
      }
      const role = classifyColor(trimmed, theme)
      const replacement = pick(role, trimmed)
      remapped += 1
      return `${attribute}=${quote}${replacement}${quote}`
    },
  )

  // Inline `style="fill:#abc;stroke:#def"` declarations.
  text = text.replace(/style\s*=\s*("|')([^"']*)\1/gi, (whole, quote, css) => {
    const next = String(css).replace(
      /\b(fill|stroke|stop-color)\s*:\s*([^;"']+)/gi,
      (decl, property, value) => {
        const trimmed = String(value).trim()
        if (trimmed === '' || /^(none|transparent|inherit|currentcolor|url\()/i.test(trimmed)) return decl
        const role = classifyColor(trimmed, theme)
        remapped += 1
        return `${property}:${pick(role, trimmed)}`
      },
    )
    return next === css ? whole : `style=${quote}${next}${quote}`
  })

  // Text must use the template's typeface, or the diagram reads as a foreign element.
  // The font is applied by `applySvgFont()`; the palette report just names it.
  const fontFamily = theme?.fonts?.body ?? theme?.fonts?.title ?? 'Arial'

  return { svg: text, remapped, palette: { ...palette, fontFamily } }
}

/**
 * Inject the template's font family into every `<text>` / `<tspan>` that carries text.
 * Separated from the colour pass so the compiler can call it after `prepareSvg()` has
 * settled the viewBox and size.
 *
 * @param {string} svg
 * @param {import('./schema.js').ThemeTokens} theme
 * @returns {string}
 */
export function applySvgFont(svg, theme) {
  const family = theme?.fonts?.body ?? theme?.fonts?.title ?? 'Arial'
  const escaped = String(family).replace(/"/g, '&quot;')
  let text = String(svg ?? '')
  // Drop any model-chosen family first, so the template wins.
  text = text.replace(/\sfont-family\s*=\s*("|')[^"']*\1/gi, '')
  text = text.replace(/<text\b([^>]*)>/gi, (whole, attributes) => {
    if (/\sfont-family\s*=/.test(attributes)) return whole
    return `<text${attributes} font-family="${escaped}">`
  })
  return text
}

/**
 * One call the compiler uses: colour remap + font injection.
 * @param {string} svg
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {{distinctAccents?: boolean}} [options]
 */
export function themeSvg(svg, theme, options = {}) {
  const normalized = normalizeSvgTheme(svg, theme, options)
  return {
    svg: applySvgFont(normalized.svg, theme),
    remapped: normalized.remapped,
    palette: normalized.palette,
  }
}
