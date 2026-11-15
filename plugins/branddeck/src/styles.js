/**
 * Style profiles — the reusable 「风格配置文件」.
 *
 * A style is the answer to "make it look like this" without editing any deck:
 * density/tone/layout choices plus optional colour and font overrides on top of a
 * template's tokens. It is stored as JSON under `<store>/styles/<styleId>.json`,
 * so a style authored once (by hand in the panel, or from a conversation through
 * `branddeck_style_save`) is reusable for every later deck.
 *
 * The module is pure data + small validators; layout lives in `generate.js`.
 */
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { z } from 'zod'
import { storeRoot, storeError, assertSafeId } from './store.js'

export const DENSITIES = ['airy', 'standard', 'dense']
export const TONES = ['formal', 'bold', 'minimal']
export const BULLET_LAYOUTS = ['list', 'cards']
export const COVER_VARIANTS = ['left', 'centered']

const hexOrToken = z.string().min(1).max(64)

export const StyleSchema = z
  .object({
    styleId: z.string().min(1).max(64),
    name: z.string().min(1).max(80),
    /** Free-form intent, e.g. the sentence the user typed when they asked for it. */
    description: z.string().max(2000).optional(),
    /** Template this style is tuned for; generation may override it. */
    templateId: z.string().min(1).max(64).optional(),
    density: z.enum(DENSITIES).default('standard'),
    tone: z.enum(TONES).default('formal'),
    bulletLayout: z.enum(BULLET_LAYOUTS).default('list'),
    coverVariant: z.enum(COVER_VARIANTS).default('left'),
    accentBar: z.boolean().default(true),
    closingSlide: z.boolean().default(true),
    /** Break up a run of identical layouts with a neutral re-render (default on). */
    varyLayout: z.boolean().default(true),
    /** Insert a sparse "density contrast" page every `breatherEvery` content slides. */
    breather: z.boolean().default(false),
    breatherEvery: z.number().int().min(3).max(12).default(5),
    maxSlides: z.number().int().min(1).max(60).default(12),
    colors: z
      .object({
        background: hexOrToken.optional(),
        surface: hexOrToken.optional(),
        textPrimary: hexOrToken.optional(),
        textSecondary: hexOrToken.optional(),
        primary: hexOrToken.optional(),
        accent: hexOrToken.optional(),
      })
      .default({}),
    fonts: z
      .object({ title: z.string().min(1).max(80).optional(), body: z.string().min(1).max(80).optional() })
      .default({}),
    /** Where the style came from: hand-authored in the panel, or distilled from chat. */
    source: z.enum(['manual', 'chat', 'template']).default('manual'),
    createdAt: z.string().optional(),
    updatedAt: z.string().optional(),
  })
  .strict()

/** @typedef {import('zod').infer<typeof StyleSchema>} StyleProfile */

/** Neutral style used when the caller names none. */
export const DEFAULT_STYLE = Object.freeze({
  styleId: 'default',
  name: '标准风格',
  description: '中性商务风：一条品牌色强调条、列表式正文、标准信息密度。',
  density: 'standard',
  tone: 'formal',
  bulletLayout: 'list',
  coverVariant: 'left',
  accentBar: true,
  closingSlide: true,
  varyLayout: true,
  breather: false,
  breatherEvery: 5,
  maxSlides: 12,
  colors: {},
  fonts: {},
  source: 'template',
})

/**
 * Derive a style from a template's tokens. This is the "one click from a
 * template" path: it keeps the template's palette and only records intent.
 * @param {{templateId: string, templateName?: string}} theme
 * @param {string} [styleId]
 * @returns {StyleProfile}
 */
export function styleFromTemplate(theme, styleId) {
  const templateId = String(theme?.templateId ?? 'default')
  return StyleSchema.parse({
    ...DEFAULT_STYLE,
    styleId: styleId ?? `${templateId}-standard`,
    name: `${theme?.templateName ?? templateId} · 标准`,
    description: `沿用模板「${theme?.templateName ?? templateId}」的色板与字体，标准密度。`,
    templateId,
    source: 'template',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  })
}

/** @param {{root?: string, home?: string}} [config] */
function stylesDir(config = {}) {
  return join(storeRoot(config), 'styles')
}

/**
 * Validate + normalise a style payload, filling ids and timestamps.
 * @param {unknown} input
 * @param {{styleId?: string, existing?: StyleProfile | null}} [hints]
 * @returns {StyleProfile}
 */
export function normalizeStyle(input, hints = {}) {
  const raw = input !== null && typeof input === 'object' ? { ...input } : {}
  const existing = hints.existing ?? null
  // Clamp numeric ranges before validation: a style authored by hand (or by a model)
  // should not fail to save because one number drifted out of range.
  if (Number.isFinite(raw.breatherEvery)) {
    raw.breatherEvery = Math.min(12, Math.max(3, Math.floor(Number(raw.breatherEvery))))
  }
  const candidate = {
    ...DEFAULT_STYLE,
    ...(existing ?? {}),
    ...raw,
    styleId: String(hints.styleId ?? raw.styleId ?? existing?.styleId ?? slugStyleId(raw.name ?? 'style')),
  }
  const now = new Date().toISOString()
  candidate.createdAt = existing?.createdAt ?? raw.createdAt ?? now
  candidate.updatedAt = now
  const parsed = StyleSchema.safeParse(candidate)
  if (!parsed.success) {
    const error = storeError(
      'invalid_style',
      `style rejected: ${parsed.error.issues.map((issue) => `${issue.path.join('.')}: ${issue.message}`).join('; ')}`,
    )
    error.details = parsed.error.issues.map((issue) => ({
      path: issue.path.join('.'),
      message: issue.message,
    }))
    throw error
  }
  return parsed.data
}

/** Display name → path-safe id (CJK falls back to a digest, like deck ids). */
export function slugStyleId(name) {
  const ascii = String(name ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
  if (ascii.length > 0) return ascii.slice(0, 48)
  let digest = 0
  const text = String(name ?? 'style')
  for (let index = 0; index < text.length; index += 1) digest = (digest * 31 + text.charCodeAt(index)) >>> 0
  return `style-${digest.toString(36)}`
}

/**
 * Persist a style.
 * @param {unknown} input
 * @param {{root?: string, home?: string}} [config]
 * @param {{styleId?: string}} [hints]
 * @returns {StyleProfile}
 */
export function writeStyle(input, config = {}, hints = {}) {
  const dir = stylesDir(config)
  mkdirSync(dir, { recursive: true })
  const styleId = String(hints.styleId ?? input?.styleId ?? '')
  if (styleId !== '') assertSafeId(styleId)
  const existing = styleId === '' ? null : readStyleOrNull(styleId, config)
  const style = normalizeStyle(input, { styleId: styleId === '' ? undefined : styleId, existing })
  assertSafeId(style.styleId)
  const file = join(dir, `${style.styleId}.json`)
  const temp = `${file}.${process.pid}.${Date.now().toString(36)}.tmp`
  const body = Buffer.from(`${JSON.stringify(style, null, 2)}\n`, 'utf8')
  // Atomic-ish: write a sibling then rename, so a reader never sees half a file.
  writeFileSync(temp, body)
  try {
    renameSync(temp, file)
  } catch {
    writeFileSync(file, body)
    try {
      unlinkSync(temp)
    } catch {
      /* best effort */
    }
  }
  return style
}

/**
 * @param {string} styleId
 * @param {{root?: string, home?: string}} [config]
 * @returns {StyleProfile}
 */
export function readStyle(styleId, config = {}) {
  const style = readStyleOrNull(styleId, config)
  if (style === null) throw storeError('style_not_found', `style "${styleId}" not found`)
  return style
}

/** @param {string} styleId */
export function readStyleOrNull(styleId, config = {}) {
  try {
    assertSafeId(styleId)
  } catch {
    return null
  }
  const file = join(stylesDir(config), `${styleId}.json`)
  if (!existsSync(file)) return null
  try {
    const parsed = StyleSchema.safeParse(JSON.parse(readFileSync(file, 'utf8')))
    return parsed.success ? parsed.data : null
  } catch {
    return null
  }
}

/**
 * List styles, newest first, always starting with the built-in default.
 * @param {{root?: string, home?: string}} [config]
 */
export function listStyles(config = {}) {
  const dir = stylesDir(config)
  /** @type {StyleProfile[]} */
  const styles = [{ ...DEFAULT_STYLE }]
  if (existsSync(dir)) {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      if (!entry.isFile() || !entry.name.endsWith('.json')) continue
      const style = readStyleOrNull(entry.name.slice(0, -'.json'.length), config)
      if (style !== null && style.styleId !== DEFAULT_STYLE.styleId) styles.push(style)
    }
  }
  const rest = styles
    .slice(1)
    .sort((a, b) => String(b.updatedAt ?? '').localeCompare(String(a.updatedAt ?? '')))
  return [styles[0], ...rest]
}

/**
 * @param {string} styleId
 * @param {{root?: string, home?: string}} [config]
 * @returns {boolean}
 */
export function deleteStyle(styleId, config = {}) {
  if (styleId === DEFAULT_STYLE.styleId) {
    throw storeError('protected_style', 'the built-in default style cannot be deleted')
  }
  assertSafeId(styleId)
  const file = join(stylesDir(config), `${styleId}.json`)
  if (!existsSync(file)) return false
  unlinkSync(file)
  return true
}

/**
 * Apply a style's overrides on top of a template's tokens.
 * Colour overrides accept hex or token references; fonts override the two roles.
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {StyleProfile | null | undefined} style
 * @returns {import('./schema.js').ThemeTokens}
 */
export function applyStyleToTheme(theme, style) {
  const base = JSON.parse(JSON.stringify(theme))
  if (style === null || style === undefined) return base
  const colors = style.colors ?? {}
  for (const [key, value] of Object.entries(colors)) {
    if (typeof value === 'string' && value.length > 0) base.colors[key] = value
  }
  const fonts = style.fonts ?? {}
  if (typeof fonts.title === 'string' && fonts.title.length > 0) base.fonts.title = fonts.title
  if (typeof fonts.body === 'string' && fonts.body.length > 0) base.fonts.body = fonts.body
  return base
}

/** Directory holding style files (exported for diagnostics/tests). */
export function styleDir(config = {}) {
  return resolve(stylesDir(config))
}
