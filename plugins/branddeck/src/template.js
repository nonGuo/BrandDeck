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
import { readTemplateArtwork } from './pptx/artwork.js'
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

  /* --- the template's visual system -----------------------------------
   * Backdrop art, the content box it draws in, and its fixed footer strings. These are the
   * parts of a corporate template that survive the trip to pptxgenjs; without them the deck
   * is drawn on a white canvas with our own margins and does not look like the template.
   */
  const artwork = readTemplateArtwork(pkg, structure)
  const backdropDir = join(storeRoot(storeConfig), 'templates', templateId, 'backdrops')
  /** @type {Array<{name: string, bytes: Buffer, mediaPart: string}>} */
  const backdropFiles = []
  if (artwork.backdrops.length > 0) {
    // One file per *distinct* backdrop part, named by layout so the mapping stays readable
    // in the store (`cover-1.png`, `cover-2.jpg`, …).
    const seen = new Set()
    artwork.backdrops.forEach((backdrop, index) => {
      if (seen.has(backdrop.mediaPart)) return
      seen.add(backdrop.mediaPart)
      const bytes = pkg.buffer(backdrop.mediaPart)
      if (bytes === null || bytes === undefined || bytes.length === 0) return
      const extension = (backdrop.mediaPart.split('.').pop() ?? 'png').toLowerCase()
      const file = `backdrop-${index + 1}.${extension}`
      backdropFiles.push({ name: file, bytes: Buffer.from(bytes), mediaPart: backdrop.mediaPart })
      void backdropDir
    })
  }
  if (backdropFiles.length > 0) {
    /** @type {Map<string, {path: string, frame: any}>} */
    const byPart = new Map()
    for (const file of backdropFiles) {
      const source = artwork.backdrops.find((entry) => entry.mediaPart === file.mediaPart)
      byPart.set(file.mediaPart, {
        path: join(backdropDir, file.name),
        ...(source === undefined ? {} : { frame: { ...source.frameInches } }),
      })
    }
    const refs = artwork.backdrops
      .map((backdrop) => byPart.get(backdrop.mediaPart))
      .filter((ref) => ref !== undefined)
    if (refs.length > 0) {
      // Every theme layout gets art; the first doubles as section/closing unless the
      // template ships something more specific. Nothing is invented: a template without
      // artwork simply has no `backdrops` key and the deck stays on a clean background.
      tokens.brandAssets.backdrops = {
        cover: refs[0],
        section: refs[refs.length > 1 ? 1 : 0],
        closing: refs[0],
      }
    }
  }
  if (artwork.frame !== null) tokens.brandAssets.frame = { ...artwork.frame }
  if (artwork.footlines.length > 0) tokens.brandAssets.footlines = [...artwork.footlines]

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

  // Format the template once, here: the layouts inventory + a human-readable spec
  // are stored next to the tokens so a later run never re-derives them.
  //
  // The spec builder is loaded lazily and optionally: the spec is a convenience
  // artefact, and a missing/broken builder must never take the whole plugin down
  // (a static import would fail module evaluation for every caller).
  let doc = { markdown: null, layouts: null }
  try {
    const builder = await import('./template-doc.js')
    doc = builder.buildTemplateDoc({
      pkg,
      structure,
      theme: restoreExtraAssets(checked.data, logo, logoPath),
      meta,
      media,
    })
  } catch (error) {
    // A spec problem must never fail template parsing: the tokens are the contract.
    meta.specError = error instanceof Error ? error.message : String(error)
  }

  return {
    theme: restoreExtraAssets(checked.data, logo, logoPath),
    logoBytes: logo === null ? null : logo.buffer,
    /** Backdrop files the store must write next to `logo.png`; paths are already in the tokens. */
    backdropFiles,
    /** What the template offered visually, for the spec/docs and for debugging. */
    artwork,
    meta,
    media,
    layouts: doc.layouts ?? null,
    markdown: doc.markdown ?? null,
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
  // Everything the template gave us beyond the four colour/font groups. This used to
  // rebuild `brandAssets` from `{icons}` only, which silently dropped the frame, the
  // backdrop art and the footer markers the parser had just extracted.
  const carried = theme.brandAssets ?? {}
  theme.brandAssets = {
    icons: carried.icons ?? {},
    ...(carried.frame === undefined ? {} : { frame: { ...carried.frame } }),
    ...(carried.backdrops === undefined ? {} : { backdrops: { ...carried.backdrops } }),
    ...(carried.footlines === undefined ? {} : { footlines: [...carried.footlines] }),
  }
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
