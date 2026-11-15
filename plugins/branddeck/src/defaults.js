/**
 * BrandDeck fallback theme tokens.
 *
 * Used when no corporate `.pptx` template has been parsed yet (design doc Phase 1:
 * "内置一套默认 ThemeTokens"), and by the template parser as the last link of every
 * fallback chain when a template part is missing or unusable.
 *
 * The object is deep-frozen: consumers must never mutate shared defaults. Use
 * `makeDefaultTheme()` to get a mutable, schema-valid copy.
 */

const HEX = /^#[0-9A-F]{6}$/

/**
 * A deliberately neutral, print-safe brand deck: white page, cool grey surface,
 * one corporate blue and one warm accent. Every colour is a 6-digit HEX so it can be
 * handed to pptxgenjs (which wants `RRGGBB` without `#`) and to CSS verbatim.
 * @type {import('./schema.js').ThemeTokens}
 */
export const DEFAULT_THEME = Object.freeze({
  templateId: 'branddeck-default',
  templateName: 'BrandDeck Default',
  colors: Object.freeze({
    background: '#FFFFFF',
    surface: '#F4F6F9',
    textPrimary: '#1F2937',
    textSecondary: '#6B7280',
    primary: '#2F5BD2',
    accent: '#F59E0B',
    chartPalette: Object.freeze(['#2F5BD2', '#F59E0B', '#10B981', '#EF4444', '#8B5CF6', '#0EA5E9']),
  }),
  fonts: Object.freeze({
    title: 'Microsoft YaHei',
    body: 'Microsoft YaHei',
    latin: 'Arial',
  }),
  brandAssets: Object.freeze({ icons: Object.freeze({}) }),
})

/** Template id used for the built-in theme (path safe: `^[a-zA-Z0-9._-]{1,64}$`). */
export const DEFAULT_TEMPLATE_ID = DEFAULT_THEME.templateId

/** Default CJK/Latin font stack, exported separately for the parser fallback chains. */
export const DEFAULT_FONTS = DEFAULT_THEME.fonts

/**
 * Fresh, mutable copy of {@link DEFAULT_THEME} with an optional identity override.
 * @param {{templateId?: string, templateName?: string}} [identity]
 * @returns {import('./schema.js').ThemeTokens} a value that satisfies `ThemeTokensSchema`
 */
export function makeDefaultTheme(identity = {}) {
  const theme = structuredClone(DEFAULT_THEME)
  if (typeof identity.templateId === 'string' && identity.templateId.length > 0) {
    theme.templateId = identity.templateId
  }
  if (typeof identity.templateName === 'string' && identity.templateName.length > 0) {
    theme.templateName = identity.templateName
  }
  return theme
}

/**
 * True when every colour of a theme is a plain 6-digit HEX. The compiler uses it to
 * refuse to emit colours pptxgenjs would silently drop.
 * @param {unknown} theme
 * @returns {boolean}
 */
export function hasOnlyHexColors(theme) {
  const colors = theme?.colors
  if (colors === undefined || colors === null) return false
  const flat = [
    colors.background,
    colors.surface,
    colors.textPrimary,
    colors.textSecondary,
    colors.primary,
    colors.accent,
    ...(Array.isArray(colors.chartPalette) ? colors.chartPalette : []),
  ]
  return flat.every((value) => typeof value === 'string' && HEX.test(value))
}
