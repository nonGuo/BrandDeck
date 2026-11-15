/**
 * `.pptx` → ThemeTokens + logo + media inventory.
 *
 * This is the orchestrator the Host tool `branddeck_template_parse` calls. It is kept
 * free of *persistence* side effects so it can be unit-tested without a store: it only
 * reads the package and computes the absolute `logo.png` path that
 * `store.writeTemplate()` will write to (same `templateId`, same store root), so
 * `theme.brandAssets.logoPath` is already correct when the caller persists the result.
 *
 * Contract: `parsePptxTemplate({bytes|path, templateId?, templateName?, storeConfig?})`
 * → `{ theme, logoBytes, meta, media, raw }`.
 */
import { basename, extname, join } from 'node:path'
import { createHash } from 'node:crypto'
import { loadPackage } from './pptx/read.js'
import { discoverStructure, extractThemeTokens } from './pptx/theme.js'
import { findBrandLogo, listMedia } from './pptx/media.js'
import { ThemeTokensSchema } from './schema.js'
import { storeRoot } from './store.js'

const SAFE_ID = /^[a-zA-Z0-9._-]{1,64}$/

/**
 * Turn a display name into a path-safe template id (`^[a-zA-Z0-9._-]{1,64}$`).
 *
 * Names that lose information (non-ASCII, e.g. `PPT模板-HW`) get a short deterministic
 * hash suffix so that two different Chinese names never collide on the same file name,
 * while the same name always maps to the same id (no random ids in the store).
 *
 * @param {string} name
 * @returns {string} e.g. `ppt-hw-3f9c21`, `hw` or `template-1a2b3c`
 */
export function namespaceTemplateId(name) {
  const raw = String(name ?? '').trim()
  const ascii = raw.normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
  const slug = ascii
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/[._-]{2,}/g, '-')
    .replace(/^[-._]+|[-._]+$/g, '')
    .slice(0, 40)
  // Non-ASCII names (e.g. `PPT模板-HW`) lose characters in the slug, so fold in a
  // deterministic digest to keep different names apart while staying reproducible.
  const lossy = slug.length === 0 || /[^\x00-\x7f]/.test(raw)
  const digest = createHash('sha1').update(raw, 'utf8').digest('hex').slice(0, 6)
  const id = lossy ? (slug.length === 0 ? `template-${digest}` : `${slug}-${digest}`) : slug
  const finalId = id.slice(0, 64)
  return SAFE_ID.test(finalId) && !finalId.includes('..') ? finalId : `template-${digest}`
}

/**
 * Derive a display name from a file path (`C:\x\PPT模板-HW.pptx` → `PPT模板-HW`).
 * @param {string | undefined} path
 * @returns {string | undefined}
 */
function nameFromPath(path) {
  if (typeof path !== 'string' || path.length === 0) return undefined
  const file = basename(path)
  const extension = extname(file)
  const stem = extension.length > 0 ? file.slice(0, -extension.length) : file
  return stem.length > 0 ? stem : undefined
}

/**
 * @typedef {object} ParsedTemplate
 * @property {import('./schema.js').ThemeTokens} theme
 * @property {Buffer | null} logoBytes
 * @property {object} meta
 * @property {Array<object>} media  JSON-safe media inventory
 * @property {object} raw           parser diagnostics (clrScheme/fontScheme/slideSize/notes)
 */

/**
 * Parse a corporate `.pptx` template into BrandDeck design tokens.
 *
 * @param {object} input
 * @param {Buffer | Uint8Array} [input.bytes]
 * @param {string} [input.path]
 * @param {string} [input.templateId]   defaults to `namespaceTemplateId(name)`
 * @param {string} [input.templateName] defaults to the file name stem
 * @param {{root?: string, home?: string}} [input.storeConfig] store config used to
 *        compute `theme.brandAssets.logoPath`
 * @returns {Promise<ParsedTemplate>}
 */
export async function parsePptxTemplate(input = {}) {
  const sourcePath = typeof input.path === 'string' ? input.path : undefined
  const templateName =
    trimmed(input.templateName) ?? nameFromPath(sourcePath) ?? 'BrandDeck Template'
  const templateId = trimmed(input.templateId) ?? namespaceTemplateId(templateName)
  if (!SAFE_ID.test(templateId) || templateId.includes('..')) {
    const error = new Error(
      `unsafe templateId "${templateId}": only [a-zA-Z0-9._-] are allowed, max 64 characters`,
    )
    error.code = 'invalid_id'
    throw error
  }

  const pkg = await loadPackage(
    input.bytes !== undefined ? { bytes: input.bytes } : { path: /** @type {string} */ (sourcePath) },
  )
  const structure = discoverStructure(pkg)
  const { tokens, raw } = extractThemeTokens(pkg, { templateId, templateName })

  /* --- logo ------------------------------------------------------------ */
  const logo = findBrandLogo(pkg)
  const storeConfig = input.storeConfig ?? {}
  const logoPath = logo === null ? null : join(storeRoot(storeConfig), 'templates', templateId, 'logo.png')

  tokens.brandAssets = {
    icons: tokens.brandAssets?.icons ?? {},
  }
  if (logo !== null && logoPath !== null) {
    tokens.brandAssets.logoPath = logoPath
    tokens.brandAssets.logoWidthPx = logo.widthPx
    tokens.brandAssets.logoHeightPx = logo.heightPx
    if (logo.frameInches !== null) tokens.brandAssets.logoFrame = { ...logo.frameInches }
  }

  const checked = ThemeTokensSchema.safeParse(tokens)
  if (!checked.success) {
    const error = new Error(
      `extracted tokens failed ThemeTokensSchema: ${checked.error.issues
        .map((issue) => `${issue.path.join('.')}: ${issue.message}`)
        .join('; ')}`,
    )
    error.code = 'invalid_theme'
    throw error
  }

  /* --- media inventory (JSON safe: no Buffers) ------------------------- */
  const media = listMedia(pkg).map((entry) => ({
    path: entry.path,
    bytes: entry.bytes,
    contentType: entry.contentType,
    format: entry.format,
    widthPx: entry.widthPx,
    heightPx: entry.heightPx,
    referencedBy: entry.referencedBy,
  }))

  const meta = {
    templateId,
    templateName,
    sourcePath: sourcePath ?? null,
    slideCount: raw.slides.length,
    masters: raw.masters.length,
    layouts: raw.layouts.length,
    themes: pkg.partNames().filter((part) => /^ppt\/theme\/theme\d+\.xml$/.test(part)).length,
    slideSize: raw.slideSize,
    logo:
      logo === null
        ? null
        : {
            partPath: logo.partPath,
            bytes: logo.bytes,
            widthPx: logo.widthPx,
            heightPx: logo.heightPx,
            aspect: logo.aspect,
            frameInches: logo.frameInches,
            masterParts: logo.masterParts,
            logoPath,
          },
    media,
    notes: raw.notes,
  }

  return {
    theme: restoreExtraAssets(checked.data, logo, logoPath),
    logoBytes: logo === null ? null : logo.buffer,
    meta,
    media,
    raw: { ...raw, themePart: structure.themePart },
  }
}

/**
 * `ThemeTokensSchema` strips keys it does not know, but the compiler honours
 * `brandAssets.logoFrame` (the real logo sits bottom-right on the master, not in a
 * top-right corner). Re-attach the asset extras the schema drops.
 *
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {import('./pptx/media.js').BrandLogo | null} logo
 * @param {string | null} logoPath
 * @returns {import('./schema.js').ThemeTokens}
 */
function restoreExtraAssets(theme, logo, logoPath) {
  theme.brandAssets = { icons: theme.brandAssets?.icons ?? {} }
  if (logo === null || logoPath === null) return theme
  theme.brandAssets.logoPath = logoPath
  theme.brandAssets.logoWidthPx = logo.widthPx
  theme.brandAssets.logoHeightPx = logo.heightPx
  if (logo.frameInches !== null) theme.brandAssets.logoFrame = { ...logo.frameInches }
  return theme
}

/**
 * @param {unknown} value
 * @returns {string | undefined}
 */
function trimmed(value) {
  if (typeof value !== 'string') return undefined
  const text = value.trim()
  return text.length > 0 ? text : undefined
}
