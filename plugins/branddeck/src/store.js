/**
 * Template / deck persistence under `$DSH_HOME/branddeck` (ARCHITECTURE §3).
 *
 * Layout:
 *   <root>/templates/<templateId>/theme.json    ThemeTokens (with brandAssets.logoPath)
 *   <root>/templates/<templateId>/logo.png      logo extracted from the master (optional)
 *   <root>/templates/<templateId>/source.pptx   original template copy (optional)
 *   <root>/templates/<templateId>/meta.json     { templateId, templateName, sourcePath,
 *                                                 slideCount, updatedAt, media[] }
 *   <root>/decks/<deckId>.json                  CardDeck
 *
 * Every write is atomic-ish (write a sibling temp file, then `rename`), so a crash or a
 * concurrent reader never observes a half-written JSON file. `templateId`/`deckId` are
 * path-safe by construction: `assertSafeId` rejects anything outside
 * `^[a-zA-Z0-9._-]{1,64}$` and any `..` segment.
 *
 * `config` is always the optional **last** argument and accepts `{ root }` (the
 * `branddeck` root itself) or `{ home }` (a DSH home; the store then lives in
 * `<home>/branddeck`). Tests pass `{ root: '.tmp/store-xxx' }`.
 */
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { randomBytes } from 'node:crypto'
import { createHash } from 'node:crypto'
import { homedir } from 'node:os'
import { join, dirname, resolve } from 'node:path'
import { CardDeckSchema, ThemeTokensSchema } from './schema.js'
import { DEFAULT_THEME, makeDefaultTheme } from './defaults.js'
import { deckToOutline, applyOutline } from './outline.js'

const SAFE_ID = /^[a-zA-Z0-9._-]{1,64}$/
const STORE_DIRNAME = 'branddeck'

let writeCounter = 0

/* ------------------------------------------------------------------ *
 * Paths
 * ------------------------------------------------------------------ */

/**
 * Resolve the BrandDeck store root.
 * `config.root` wins (used verbatim), then `config.home`, then `$DSH_HOME`, then
 * `~/.dsh` — the latter three get `/branddeck` appended.
 * @param {{root?: string, home?: string}} [config]
 * @returns {string} absolute path
 */
export function storeRoot(config = {}) {
  if (typeof config.root === 'string' && config.root.trim().length > 0) {
    return resolve(config.root.trim())
  }
  const home = typeof config.home === 'string' && config.home.trim().length > 0
    ? config.home.trim()
    : process.env.DSH_HOME ?? join(homedir(), '.dsh')
  return resolve(home, STORE_DIRNAME)
}

/**
 * Create `templates/` and `decks/` if they do not exist yet.
 * @param {{root?: string, home?: string}} [config]
 * @returns {{root: string, templatesDir: string, decksDir: string}}
 */
export function ensureStore(config = {}) {
  const root = storeRoot(config)
  const templatesDir = join(root, 'templates')
  const decksDir = join(root, 'decks')
  mkdirSync(templatesDir, { recursive: true })
  mkdirSync(decksDir, { recursive: true })
  return { root, templatesDir, decksDir }
}

/**
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 */
function templateDir(id, config = {}) {
  assertSafeId(id)
  return join(storeRoot(config), 'templates', id)
}

/**
 * Directory holding one deck's artefacts (`card.json` / `outline.md` / `meta.json`).
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 */
function deckDir(id, config = {}) {
  assertSafeId(id)
  return join(storeRoot(config), 'decks', id)
}

/**
 * Canonical card JSON inside the deck directory.
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 */
function deckFile(id, config = {}) {
  assertSafeId(id)
  return join(deckDir(id, config), 'card.json')
}

/**
 * Pre-0.2 single-file deck location, still read so existing stores keep working.
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 */
function legacyDeckFile(id, config = {}) {
  assertSafeId(id)
  return join(storeRoot(config), 'decks', `${id}.json`)
}

/**
 * Reject ids that could escape the store directory.
 * @param {unknown} id
 * @returns {string} the validated id
 * @throws {Error & {code:'invalid_id'}}
 */
export function assertSafeId(id) {
  if (typeof id !== 'string' || id.length === 0) {
    throw storeError('invalid_id', 'id must be a non-empty string')
  }
  if (id === '.' || id === '..' || id.includes('..')) {
    throw storeError('invalid_id', `unsafe id "${id}": parent-directory segments are not allowed`)
  }
  if (!SAFE_ID.test(id)) {
    throw storeError(
      'invalid_id',
      `unsafe id "${id}": only [a-zA-Z0-9._-] are allowed, max 64 characters`,
    )
  }
  return id
}

/**
 * Error factory for the store layer (the Host tool layer maps `error.code`).
 * @param {string} code
 * @param {string} message
 * @param {unknown} [cause]
 * @returns {Error & {code: string, details?: unknown, cause?: unknown}}
 */
export function storeError(code, message, cause) {
  const error = new Error(message)
  error.code = code
  if (cause !== undefined) error.cause = cause
  return error
}

/* ------------------------------------------------------------------ *
 * Atomic file helpers
 * ------------------------------------------------------------------ */

/**
 * Write JSON atomically (temp file + rename), UTF-8, pretty-printed.
 * @param {string} file
 * @param {unknown} value
 */
function writeJsonAtomic(file, value) {
  const text = `${JSON.stringify(value, null, 2)}\n`
  writeFileAtomic(file, Buffer.from(text, 'utf8'))
}

/**
 * @param {string} file
 * @param {Buffer | Uint8Array} bytes
 */
function writeFileAtomic(file, bytes) {
  mkdirSync(dirname(file), { recursive: true })
  writeCounter += 1
  const temp = `${file}.${process.pid}.${writeCounter}.${randomBytes(4).toString('hex')}.tmp`
  writeFileSync(temp, bytes)
  try {
    renameSync(temp, file)
  } catch (cause) {
    try {
      unlinkSync(temp)
    } catch {
      /* best effort */
    }
    throw storeError('write_failed', `cannot persist ${file}`, cause)
  }
}

/**
 * @param {string} file
 * @returns {any | null}
 */
function readJsonIfPresent(file) {
  if (!existsSync(file)) return null
  try {
    return JSON.parse(readFileSync(file, 'utf8'))
  } catch (cause) {
    throw storeError('corrupt_store', `cannot parse store file ${file}`, cause)
  }
}

/** @param {string} file */
function mtimeIso(file) {
  try {
    return statSync(file).mtime.toISOString()
  } catch {
    return null
  }
}

/**
 * Slugify a display name into a safe id (`^[a-zA-Z0-9._-]{1,64}$`).
 * Keeps CJK out of ids on purpose: ids become file names. When the name carries no
 * usable ASCII letters (e.g. a pure-Chinese deck title) a deterministic digest is used
 * instead of a random one, so the same title always maps to the same id.
 * @param {string} name
 * @param {string} [prefix] fallback prefix, e.g. `deck`
 * @returns {string}
 */
export function slugifyId(name, prefix = 'item') {
  const raw = String(name ?? '').trim()
  const base = raw
    .normalize('NFKD')
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/[-._]{2,}/g, '-')
    .replace(/^[-._]+|[-._]+$/g, '')
    .slice(0, 48)
  const digest = createHash('sha1').update(raw, 'utf8').digest('hex').slice(0, 6)
  if (base.length === 0) return `${prefix}-${digest}`
  if (!/[a-z]/.test(base)) return `${prefix}-${base}`.slice(0, 64)
  const id = /^[a-z]/.test(base) ? base : `${prefix}-${base}`
  return id.slice(0, 64)
}

/* ------------------------------------------------------------------ *
 * Templates
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} TemplateSummary
 * @property {string} templateId
 * @property {string} templateName
 * @property {boolean} hasLogo
 * @property {Record<string, string | string[]>} colors
 * @property {Record<string, string>} fonts
 * @property {string | null} updatedAt
 * @property {number} slideCount
 * @property {number} mediaCount
 */

/**
 * Persist a parsed template.
 *
 * @param {object} input
 * @param {string} input.templateId
 * @param {string} input.templateName
 * @param {import('./schema.js').ThemeTokens} input.tokens  theme tokens (schema-validated)
 * @param {Buffer | Uint8Array} [input.logoBytes]           logo to write as `logo.png`
 * @param {object} [input.meta]                             extra meta fields
 * @param {Array<object>} [input.media]                     media inventory for meta.json
 * @param {Buffer | Uint8Array} [input.sourceBytes]         original .pptx copy
 * @param {{root?: string, home?: string}} [config]
 * @returns {TemplateSummary}
 */
export function writeTemplate(input, config = {}) {
  const { templateId, templateName, tokens } = input ?? {}
  assertSafeId(templateId)
  if (typeof templateName !== 'string' || templateName.trim().length === 0) {
    throw storeError('invalid_template', 'templateName must be a non-empty string')
  }
  const parsed = ThemeTokensSchema.safeParse(tokens)
  if (!parsed.success) {
    throw storeError(
      'invalid_theme',
      `theme tokens rejected by ThemeTokensSchema: ${parsed.error.issues
        .map((issue) => `${issue.path.join('.')}: ${issue.message}`)
        .join('; ')}`,
    )
  }

  ensureStore(config)
  const dir = templateDir(templateId, config)
  mkdirSync(dir, { recursive: true })

  // `ThemeTokensSchema` strips unknown keys, so anything beyond the contract
  // (`brandAssets.logoFrame`, which the compiler honours) has to be carried over from
  // the *input* tokens explicitly.
  const inputAssets = tokens?.brandAssets ?? {}
  /** @type {{x: number, y: number, w: number, h: number} | undefined} */
  const logoFrame = inputAssets.logoFrame
  /** @type {import('./schema.js').ThemeTokens} */
  const theme = {
    templateId: parsed.data.templateId,
    templateName: parsed.data.templateName,
    colors: parsed.data.colors,
    fonts: parsed.data.fonts,
    brandAssets: { icons: parsed.data.brandAssets?.icons ?? {} },
  }
  const logoPath = join(dir, 'logo.png')
  const hasLogo = input.logoBytes !== undefined && input.logoBytes !== null && input.logoBytes.length > 0
  if (hasLogo) {
    // brandAssets.logoPath must point at the file we are about to write (absolute).
    theme.brandAssets.logoPath = logoPath
    if (inputAssets.logoWidthPx !== undefined) theme.brandAssets.logoWidthPx = inputAssets.logoWidthPx
    if (inputAssets.logoHeightPx !== undefined) theme.brandAssets.logoHeightPx = inputAssets.logoHeightPx
    if (logoFrame !== undefined) theme.brandAssets.logoFrame = logoFrame
    writeFileAtomic(logoPath, /** @type {Buffer | Uint8Array} */ (input.logoBytes))
  }

  if (input.sourceBytes !== undefined && input.sourceBytes !== null && input.sourceBytes.length > 0) {
    writeFileAtomic(join(dir, 'source.pptx'), /** @type {Buffer | Uint8Array} */ (input.sourceBytes))
  }

  writeJsonAtomic(join(dir, 'theme.json'), theme)

  // Reusable style sheet: the layouts inventory and the human-readable spec are
  // written next to the tokens so the next run (or a reviewer) never has to
  // re-derive them from the .pptx.
  const layoutsInventory = input.layoutsInventory ?? null
  if (layoutsInventory !== null && layoutsInventory !== undefined) {
    writeJsonAtomic(join(dir, 'layouts.json'), layoutsInventory)
  }
  const hasSpec = typeof input.markdown === 'string' && input.markdown.trim().length > 0
  if (hasSpec) {
    writeFileAtomic(join(dir, 'template.md'), Buffer.from(/** @type {string} */ (input.markdown), 'utf8'))
  }

  const media = Array.isArray(input.media)
    ? input.media.map((entry) => ({
        path: entry?.path ?? '',
        bytes: entry?.bytes ?? 0,
        contentType: entry?.contentType ?? 'application/octet-stream',
        widthPx: entry?.widthPx ?? null,
        heightPx: entry?.heightPx ?? null,
        referencedBy: Array.isArray(entry?.referencedBy) ? entry.referencedBy : [],
      }))
    : []

  const meta = {
    templateId,
    templateName: templateName.trim(),
    sourcePath: input.meta?.sourcePath ?? null,
    sourceCopy: existsSync(join(dir, 'source.pptx')) ? 'source.pptx' : null,
    slideCount: Number.isFinite(input.meta?.slideCount) ? input.meta.slideCount : 0,
    masters: input.meta?.masters ?? null,
    layouts: input.meta?.layouts ?? null,
    themes: input.meta?.themes ?? null,
    slideSize: input.meta?.slideSize ?? null,
    raw: input.meta?.raw ?? null,
    hasLogo,
    hasLayouts: layoutsInventory !== null && layoutsInventory !== undefined,
    hasSpec,
    updatedAt: new Date().toISOString(),
    media,
  }
  if (input.meta?.notes !== undefined) meta.notes = input.meta.notes
  writeJsonAtomic(join(dir, 'meta.json'), meta)

  return summarizeTemplate({ theme, meta, hasLogo })
}

/**
 * Read a persisted template.
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 * @returns {{theme: import('./schema.js').ThemeTokens, meta: any, logoPath: string | null, hasLogo: boolean, dir: string}}
 */
export function readTemplate(id, config = {}) {
  const dir = templateDir(id, config)
  const themeFile = join(dir, 'theme.json')
  if (!existsSync(themeFile)) {
    throw storeError('template_not_found', `template "${id}" not found in ${dir}`)
  }
  /** @type {import('./schema.js').ThemeTokens} */
  const theme = readJsonIfPresent(themeFile)
  const meta = readJsonIfPresent(join(dir, 'meta.json')) ?? {
    templateId: id,
    templateName: theme?.templateName ?? id,
    updatedAt: mtimeIso(themeFile),
    media: [],
  }
  const logoPath = join(dir, 'logo.png')
  const hasLogo = existsSync(logoPath)
  // Always hand back the absolute path that exists *now* (the store root can move).
  const assets = { icons: {}, ...(theme.brandAssets ?? {}) }
  if (hasLogo) assets.logoPath = logoPath
  else delete assets.logoPath
  theme.brandAssets = assets
  const layoutsFile = join(dir, 'layouts.json')
  const specFile = join(dir, 'template.md')
  return {
    theme,
    meta,
    logoPath: hasLogo ? logoPath : null,
    hasLogo,
    dir,
    layouts: existsSync(layoutsFile) ? readJsonIfPresent(layoutsFile) : null,
    markdown: existsSync(specFile) ? readFileSync(specFile, 'utf8') : null,
    sourcePath: existsSync(join(dir, 'source.pptx')) ? join(dir, 'source.pptx') : null,
  }
}

/**
 * List persisted templates, newest first.
 * @param {{root?: string, home?: string}} [config]
 * @returns {TemplateSummary[]}
 */
export function listTemplates(config = {}) {
  const templatesDir = join(storeRoot(config), 'templates')
  if (!existsSync(templatesDir)) return []
  /** @type {TemplateSummary[]} */
  const summaries = []
  for (const entry of readdirSync(templatesDir, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue
    if (!SAFE_ID.test(entry.name) || entry.name.includes('..')) continue
    try {
      const result = readTemplate(entry.name, config)
      summaries.push(summarizeTemplate(result))
    } catch {
      // A half-written or foreign directory must not break the listing.
    }
  }
  return summaries.sort((a, b) => String(b.updatedAt ?? '').localeCompare(String(a.updatedAt ?? '')))
}

/**
 * Compact template description for the panel/tool layer (HTTP `TemplateSummary`).
 * Accepts the result of `readTemplate()` or `{theme, meta, hasLogo}`.
 * @param {{theme?: import('./schema.js').ThemeTokens, meta?: any, hasLogo?: boolean}} input
 * @returns {TemplateSummary}
 */
export function summarizeTemplate(input) {
  const theme = input?.theme ?? makeDefaultTheme()
  const meta = input?.meta ?? {}
  const hasLogo = input?.hasLogo ?? Boolean(theme?.brandAssets?.logoPath)
  return {
    templateId: theme.templateId ?? meta.templateId ?? 'unknown',
    templateName: theme.templateName ?? meta.templateName ?? 'unknown',
    hasLogo,
    colors: {
      background: theme.colors?.background ?? DEFAULT_THEME.colors.background,
      surface: theme.colors?.surface ?? DEFAULT_THEME.colors.surface,
      textPrimary: theme.colors?.textPrimary ?? DEFAULT_THEME.colors.textPrimary,
      textSecondary: theme.colors?.textSecondary ?? DEFAULT_THEME.colors.textSecondary,
      primary: theme.colors?.primary ?? DEFAULT_THEME.colors.primary,
      accent: theme.colors?.accent ?? DEFAULT_THEME.colors.accent,
      chartPalette: [...(theme.colors?.chartPalette ?? DEFAULT_THEME.colors.chartPalette)],
    },
    fonts: {
      title: theme.fonts?.title ?? DEFAULT_THEME.fonts.title,
      body: theme.fonts?.body ?? DEFAULT_THEME.fonts.body,
      latin: theme.fonts?.latin ?? DEFAULT_THEME.fonts.latin,
    },
    updatedAt: meta.updatedAt ?? null,
    slideCount: Number.isFinite(meta.slideCount) ? meta.slideCount : 0,
    mediaCount: Array.isArray(meta.media) ? meta.media.length : 0,
    ...(typeof meta.hasLayouts === 'boolean' ? { hasLayouts: meta.hasLayouts } : {}),
    ...(typeof meta.hasSpec === 'boolean' ? { hasSpec: meta.hasSpec } : {}),
    ...(typeof meta.sourceCopy === 'string' ? { sourceCopy: meta.sourceCopy } : {}),
  }
}

/* ------------------------------------------------------------------ *
 * Decks
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} DeckSummary
 * @property {string} deckId
 * @property {string} title
 * @property {string} templateId
 * @property {number} slideCount
 * @property {string | null} updatedAt
 */

/**
 * Persist a deck (schema-validated) as a **deck bundle**: the card JSON plus the
 * editable text twin, a meta record and a revision snapshot.
 *
 * Why a directory instead of one JSON: generation is rarely one-shot — people
 * adjust copy and layout over several rounds. Keeping `card.json` (geometry
 * truth) next to `outline.md` (content truth) and `revisions/` (history) makes
 * every round inspectable, diffable and reversible.
 *
 * @param {import('./schema.js').CardDeck} deck
 * @param {{root?: string, home?: string}} [config]
 * @param {{outline?: string, theme?: import('./schema.js').ThemeTokens, keepRevisions?: number, revisionReason?: string}} [options]
 * @returns {DeckSummary}
 */
export function writeDeck(deck, config = {}, options = {}) {
  const parsed = CardDeckSchema.safeParse(deck)
  if (!parsed.success) {
    const error = storeError(
      'invalid_deck',
      `deck rejected by CardDeckSchema: ${parsed.error.issues
        .map((issue) => `${issue.path.join('.')}: ${issue.message}`)
        .join('; ')}`,
    )
    error.details = parsed.error.issues.map((issue) => ({
      path: issue.path.join('.'),
      message: issue.message,
    }))
    throw error
  }
  ensureStore(config)
  const id = parsed.data.deckId
  const dir = deckDir(id, config)
  mkdirSync(dir, { recursive: true })

  const cardFile = deckFile(id, config)
  writeJsonAtomic(cardFile, parsed.data)

  const outline =
    typeof options.outline === 'string'
      ? options.outline
      : deckToOutline(parsed.data, options.theme ?? makeDefaultTheme())
  writeFileAtomic(join(dir, 'outline.md'), Buffer.from(outline, 'utf8'))

  const metaPath = join(dir, 'meta.json')
  const previous = readJsonIfPresent(metaPath) ?? {}
  const revision = Number(previous.revision ?? 0) + 1
  const updatedAt = mtimeIso(cardFile)
  const baseMeta = {
    deckId: id,
    title: parsed.data.title,
    templateId: parsed.data.templateId,
    slideCount: parsed.data.slides.length,
    revision,
    updatedAt,
    previousUpdatedAt: previous.updatedAt ?? null,
    revisions: Array.isArray(previous.revisions) ? previous.revisions : [],
  }

  // Snapshot before bumping the counter so `revisions/0001` is the first saved state.
  const snapshot = saveRevision(dir, previous, standardize(parsed.data), outline, updatedAt, options.revisionReason)
  baseMeta.revisions = snapshot.entries
  baseMeta.revision = snapshot.revision
  writeJsonAtomic(metaPath, baseMeta)

  // Legacy single-file location is removed so `listDecks` cannot show the deck twice.
  const legacy = legacyDeckFile(id, config)
  if (existsSync(legacy)) {
    try {
      unlinkSync(legacy)
    } catch {
      /* best effort */
    }
  }

  return summarizeDeck(parsed.data, { updatedAt, revision: baseMeta.revision, dir })
}

/**
 * Snapshot the current state into `revisions/NNNN-{card.json,outline.md}` and prune
 * old generations. The counter comes from the meta record so ids never repeat.
 * @param {string} dir
 * @param {{revision?: number, revisions?: any[]}} previous
 * @param {import('./schema.js').CardDeck} deck
 * @param {string} outline
 * @param {string | null} updatedAt
 * @param {string | undefined} reason
 * @param {number} [keep]
 */
function saveRevision(dir, previous, deck, outline, updatedAt, reason, keep = 20) {
  const revisionsDir = join(dir, 'revisions')
  mkdirSync(revisionsDir, { recursive: true })
  const next = Number(previous.revision ?? 0) + 1
  const stamp = String(next).padStart(4, '0')
  writeJsonAtomic(join(revisionsDir, `${stamp}-card.json`), deck)
  writeFileAtomic(join(revisionsDir, `${stamp}-outline.md`), Buffer.from(outline, 'utf8'))

  /** @type {{revision: number, updatedAt: string | null, reason: string | null, card: string, outline: string}[]} */
  const entries = [...(Array.isArray(previous.revisions) ? previous.revisions : [])]
  entries.push({
    revision: next,
    updatedAt,
    reason: typeof reason === 'string' && reason.length > 0 ? reason : null,
    card: `revisions/${stamp}-card.json`,
    outline: `revisions/${stamp}-outline.md`,
  })

  // Keep the newest `keep` generations; drop the oldest files to bound growth.
  while (entries.length > keep) {
    const dropped = entries.shift()
    for (const rel of [dropped?.card, dropped?.outline]) {
      if (typeof rel !== 'string') continue
      const file = join(dir, rel)
      if (existsSync(file)) {
        try {
          unlinkSync(file)
        } catch {
          /* best effort */
        }
      }
    }
  }
  return { revision: next, entries }
}

/**
 * Deep-clone a deck into plain JSON (revision snapshots must not alias the caller).
 * @param {import('./schema.js').CardDeck} deck
 */
function standardize(deck) {
  return JSON.parse(JSON.stringify(deck))
}

/**
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 * @returns {import('./schema.js').CardDeck}
 */
export function readDeck(id, config = {}) {
  const dir = deckDir(id, config)
  const cardFile = deckFile(id, config)
  const legacy = legacyDeckFile(id, config)
  const source = existsSync(cardFile) ? cardFile : existsSync(legacy) ? legacy : null
  if (source === null) throw storeError('deck_not_found', `deck "${id}" not found (${dir})`)
  const deck = readJsonIfPresent(source)
  const parsed = CardDeckSchema.safeParse(deck)
  if (!parsed.success) {
    const error = storeError('invalid_deck', `stored deck "${id}" no longer matches CardDeckSchema`)
    error.details = parsed.error.issues.map((issue) => ({
      path: issue.path.join('.'),
      message: issue.message,
    }))
    throw error
  }
  return parsed.data
}

/**
 * The editable text twin of a deck. Regenerated on demand for decks written by an
 * older version that only had `card.json`.
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 * @param {import('./schema.js').ThemeTokens} [theme]
 * @returns {{markdown: string, revision: number, updatedAt: string | null}}
 */
export function readDeckOutline(id, config = {}, theme) {
  const dir = deckDir(id, config)
  const outlineFile = join(dir, 'outline.md')
  const meta = readJsonIfPresent(join(dir, 'meta.json')) ?? {}
  if (existsSync(outlineFile)) {
    return {
      markdown: readFileSync(outlineFile, 'utf8'),
      revision: Number(meta.revision ?? 0),
      updatedAt: meta.updatedAt ?? null,
    }
  }
  const deck = readDeck(id, config)
  return {
    markdown: deckToOutline(deck, theme ?? makeDefaultTheme()),
    revision: Number(meta.revision ?? 0),
    updatedAt: meta.updatedAt ?? null,
  }
}

/**
 * Apply edited outline text back onto the stored deck, then persist both twins.
 * Geometry is preserved for every element the markdown still mentions.
 * @param {string} id
 * @param {string} markdown
 * @param {{root?: string, home?: string}} [config]
 * @param {{theme?: import('./schema.js').ThemeTokens, prune?: boolean}} [options]
 * @returns {{deck: import('./schema.js').CardDeck, warnings: string[], changes: any, revision: number}}
 */
export function writeDeckOutline(id, markdown, config = {}, options = {}) {
  const deck = readDeck(id, config)
  const theme = options.theme ?? makeDefaultTheme()
  const applied = applyOutline(markdown, deck, theme, { prune: options.prune === true })
  const meta = writeDeck(applied.deck, config, {
    outline: deckToOutline(applied.deck, theme),
    theme,
    revisionReason: 'outline edit',
  })
  return {
    deck: readDeck(id, config),
    warnings: applied.warnings,
    changes: applied.changes,
    revision: Number(meta.revision ?? 0),
  }
}

/**
 * Copy a deck bundle (card/outline/meta) into a user-chosen directory, e.g. the
 * workspace, so it can be reviewed, diffed and committed with git.
 * @param {string} id
 * @param {string} targetDir
 * @param {{root?: string, home?: string}} [config]
 * @param {{theme?: import('./schema.js').ThemeTokens}} [options]
 * @returns {{dir: string, files: string[]}}
 */
export function materializeDeck(id, targetDir, config = {}, options = {}) {
  const deck = readDeck(id, config)
  const theme = options.theme ?? makeDefaultTheme()
  const dir = resolve(targetDir)
  mkdirSync(dir, { recursive: true })
  const outline = readDeckOutline(id, config, theme)
  const files = []
  writeJsonAtomic(join(dir, 'card.json'), deck)
  files.push(join(dir, 'card.json'))
  writeFileAtomic(join(dir, 'outline.md'), Buffer.from(outline.markdown, 'utf8'))
  files.push(join(dir, 'outline.md'))
  const meta = readJsonIfPresent(join(deckDir(id, config), 'meta.json')) ?? {}
  writeJsonAtomic(join(dir, 'meta.json'), {
    deckId: deck.deckId,
    title: deck.title,
    templateId: deck.templateId,
    slideCount: deck.slides.length,
    revision: Number(meta.revision ?? 0),
    updatedAt: meta.updatedAt ?? null,
    materializedAt: new Date().toISOString(),
    note: 'card.json 是几何真相，outline.md 是文字真相；用 branddeck_outline_save 回写。',
  })
  files.push(join(dir, 'meta.json'))
  return { dir, files }
}

/**
 * List decks, newest first (by file mtime).
 * @param {{root?: string, home?: string}} [config]
 * @returns {DeckSummary[]}
 */
export function listDecks(config = {}) {
  const decksDir = join(storeRoot(config), 'decks')
  if (!existsSync(decksDir)) return []
  /** @type {DeckSummary[]} */
  const summaries = []
  /** @type {Set<string>} */
  const seen = new Set()

  for (const entry of readdirSync(decksDir, { withFileTypes: true })) {
    // New layout: one directory per deck. Legacy layout: one JSON file per deck.
    const isDir = entry.isDirectory()
    const id = isDir ? entry.name : entry.name.endsWith('.json') ? entry.name.slice(0, -'.json'.length) : ''
    if (id === '' || !SAFE_ID.test(id) || id.includes('..')) continue
    if (seen.has(id)) continue
    seen.add(id)
    try {
      const deck = readDeck(id, config)
      const stampFile = isDir ? join(decksDir, id, 'card.json') : join(decksDir, entry.name)
      const meta = isDir ? readJsonIfPresent(join(decksDir, id, 'meta.json')) : null
      summaries.push(
        summarizeDeck(deck, {
          updatedAt: mtimeIso(stampFile),
          revision: Number(meta?.revision ?? 0),
          dir: join(decksDir, id),
        }),
      )
    } catch {
      // Skip unparsable decks instead of failing the whole listing.
    }
  }
  return summaries.sort((a, b) => String(b.updatedAt ?? '').localeCompare(String(a.updatedAt ?? '')))
}

/**
 * Delete a deck. Returns `true` when a file was removed.
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 * @returns {boolean}
 */
export function deleteDeck(id, config = {}) {
  const file = deckFile(id, config)
  if (!existsSync(file)) return false
  try {
    unlinkSync(file)
    return true
  } catch (cause) {
    throw storeError('delete_failed', `cannot delete deck "${id}"`, cause)
  }
}

/**
 * Compact deck description (HTTP `DeckSummary`).
 * @param {Partial<import('./schema.js').CardDeck>} deck
 * @param {{updatedAt?: string | null}} [extra]
 * @returns {DeckSummary}
 */
export function summarizeDeck(deck, extra = {}) {
  const slides = Array.isArray(deck?.slides) ? deck.slides : []
  return {
    deckId: String(deck?.deckId ?? 'unknown'),
    title: String(deck?.title ?? '未命名演示'),
    templateId: String(deck?.templateId ?? DEFAULT_THEME.templateId),
    slideCount: slides.length,
    updatedAt: extra.updatedAt ?? null,
    ...(Number.isFinite(extra.revision) ? { revision: Number(extra.revision) } : {}),
    ...(typeof extra.dir === 'string' ? { dir: extra.dir } : {}),
  }
}

/* ------------------------------------------------------------------ *
 * Starter deck
 * ------------------------------------------------------------------ */

/** Safe-area geometry for the generated cover (1280×720 canvas, 64 px margins). */
const COVER = Object.freeze({
  left: 96,
  contentWidth: 1024,
  accentBar: { x: 96, y: 196, w: 88, h: 8 },
  title: { x: 96, y: 216, w: 1024, h: 112, fontSize: 44 },
  subtitle: { x: 96, y: 340, w: 1024, h: 64, fontSize: 22 },
  divider: { x: 96, y: 444, w: 1024, h: 2 },
  caption: { x: 96, y: 468, w: 1024, h: 36, fontSize: 14 },
})

/**
 * Build a schema-valid starter deck: a cover slide with a title, a subtitle, an accent
 * bar and a hairline divider, all inside the 64 px safe area and clear of the top-right
 * logo reservation.
 *
 * Colours are token references (`$colors.primary`, …) so the deck follows the theme it
 * was created for and stays editable in the panel.
 *
 * Pure by default: the deck is returned, not written. Pass `{ persist: true }` (or call
 * `writeDeck`) to store it.
 *
 * @param {object} input
 * @param {string} input.title
 * @param {string} [input.templateId]
 * @param {string} [input.deckId]
 * @param {import('./schema.js').ThemeTokens} [input.theme]
 * @param {string} [input.subtitle]
 * @param {boolean} [input.persist]
 * @param {{root?: string, home?: string}} [config] store config (also honoured from `input.config`)
 * @returns {import('./schema.js').CardDeck}
 */
export function createDeck(input, config = {}) {
  const theme = input?.theme ?? DEFAULT_THEME
  const title = typeof input?.title === 'string' && input.title.trim().length > 0
    ? input.title.trim().slice(0, 200)
    : '未命名演示'
  const templateId = typeof input?.templateId === 'string' && input.templateId.trim().length > 0
    ? input.templateId.trim()
    : theme.templateId ?? DEFAULT_THEME.templateId
  assertSafeId(templateId)
  const deckId = input?.deckId === undefined || input.deckId === null || input.deckId === ''
    ? slugifyId(title, 'deck')
    : assertSafeId(String(input.deckId))

  const subtitle = typeof input?.subtitle === 'string' && input.subtitle.trim().length > 0
    ? input.subtitle.trim()
    : `BrandDeck · ${theme.templateName ?? DEFAULT_THEME.templateName}`

  /** @type {import('./schema.js').CardDeck} */
  const deck = {
    deckId,
    title,
    templateId,
    slides: [
      {
        slideId: 'slide-1',
        title,
        background: '$colors.background',
        speakerNotes: '封面页：在面板中双击文字即可修改；导出时 Logo 由模板母版自动放置。',
        elements: [
          {
            type: 'shape',
            id: 'kicker-bar',
            shapeType: 'rect',
            x: COVER.accentBar.x,
            y: COVER.accentBar.y,
            w: COVER.accentBar.w,
            h: COVER.accentBar.h,
            style: { fill: '$colors.primary' },
          },
          {
            type: 'text',
            id: 'title',
            role: 'title',
            content: title,
            x: COVER.title.x,
            y: COVER.title.y,
            w: COVER.title.w,
            h: COVER.title.h,
            style: {
              fontSize: COVER.title.fontSize,
              bold: true,
              color: '$colors.textPrimary',
              align: 'left',
              valign: 'bottom',
            },
          },
          {
            type: 'text',
            id: 'subtitle',
            role: 'subtitle',
            content: subtitle,
            x: COVER.subtitle.x,
            y: COVER.subtitle.y,
            w: COVER.subtitle.w,
            h: COVER.subtitle.h,
            style: {
              fontSize: COVER.subtitle.fontSize,
              color: '$colors.textSecondary',
              align: 'left',
              valign: 'top',
            },
          },
          {
            type: 'shape',
            id: 'divider',
            shapeType: 'rect',
            x: COVER.divider.x,
            y: COVER.divider.y,
            w: COVER.divider.w,
            h: COVER.divider.h,
            style: { fill: '$colors.surface' },
          },
          {
            type: 'text',
            id: 'caption',
            role: 'caption',
            content: '1280×720 画布 · 安全区 64px · 单页正文 ≤150 字',
            x: COVER.caption.x,
            y: COVER.caption.y,
            w: COVER.caption.w,
            h: COVER.caption.h,
            style: {
              fontSize: COVER.caption.fontSize,
              color: '$colors.textSecondary',
              align: 'left',
              valign: 'middle',
            },
          },
        ],
      },
    ],
  }

  const parsed = CardDeckSchema.safeParse(deck)
  if (!parsed.success) {
    throw storeError(
      'invalid_deck',
      `starter deck failed validation: ${parsed.error.issues
        .map((issue) => `${issue.path.join('.')}: ${issue.message}`)
        .join('; ')}`,
    )
  }

  const storeConfig = input?.config ?? config
  if (input?.persist === true) writeDeck(parsed.data, storeConfig, { theme })
  return parsed.data
}

/**
 * Remove a template directory (used by `branddeck_template_delete`, if it is ever added).
 * @param {string} id
 * @param {{root?: string, home?: string}} [config]
 * @returns {boolean}
 */
export function deleteTemplate(id, config = {}) {
  const dir = templateDir(id, config)
  if (!existsSync(dir)) return false
  rmSync(dir, { recursive: true, force: true })
  return true
}
