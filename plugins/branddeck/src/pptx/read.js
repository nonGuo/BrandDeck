/**
 * OOXML package reader: jszip unpacking + fast-xml-parser wrappers + relationship
 * resolution.
 *
 * Design notes
 * ------------
 * - The whole package is inflated once, eagerly, into memory (`loadPackage`). A
 *   corporate 16:9 template is a couple of megabytes, so this is cheap, and it makes
 *   the returned handle **fully synchronous** — every consumer (theme/media/template)
 *   can walk the package without `await`, which keeps the parsing logic readable and
 *   keeps the tool layer free of promise plumbing.
 * - XML is parsed with `parseTagValue:false` / `parseAttributeValue:false` so values
 *   stay verbatim strings. This matters: `lastClr="000000"` must not become the number
 *   0, and `<a:srgbClr val="1D1D1A"/>` must survive as `'1D1D1A'` (padding zeros are
 *   meaningful for the diagnostics in ARCHITECTURE §2).
 * - `removeNSPrefix:true` strips `a:`/`p:`/`r:` prefixes from element *and* attribute
 *   names, so `r:embed="rId6"` is read as `['@_embed']`. Callers must use the
 *   prefix-free form.
 */
import { posix } from 'node:path'
import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'

/** Shared fast-xml-parser options — the only XML parse configuration we use. */
export const XML_OPTIONS = Object.freeze({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: true,
  allowBooleanAttributes: true,
})

const parser = new XMLParser({ ...XML_OPTIONS })

/**
 * Parse an XML string into a plain object tree (namespace prefixes removed).
 * Attribute access uses the `@_` prefix, e.g. `node['@_val']`.
 * @param {string} text
 * @returns {any}
 */
export function parseXml(text) {
  if (typeof text !== 'string' || text.length === 0) return {}
  return parser.parse(text)
}

/**
 * fast-xml-parser collapses a single child element into an object instead of a
 * one-element array. Every consumer that iterates siblings needs this helper.
 *
 * `asArray(undefined)` → `[]`, `asArray(node)` → `[node]`, arrays pass through.
 * @template T
 * @param {T | T[] | undefined | null} value
 * @returns {T[]}
 */
export function asArray(value) {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? value : [value]
}

/**
 * Error factory for the package layer. The Host tool layer maps `error.code` straight
 * into its `{ ok:false, error:{ code, message } }` envelope (ARCHITECTURE §4).
 * @param {string} code
 * @param {string} message
 * @param {unknown} [cause]
 * @returns {Error & { code: string, cause?: unknown }}
 */
export function pptxError(code, message, cause) {
  const error = new Error(message)
  error.code = code
  if (cause !== undefined) error.cause = cause
  return error
}

const ZIP_MAGICS = [
  [0x50, 0x4b, 0x03, 0x04], // regular zip
  [0x50, 0x4b, 0x05, 0x06], // empty zip
  [0x50, 0x4b, 0x07, 0x08], // spanned zip
]

/**
 * @param {Uint8Array} bytes
 * @returns {boolean} true when the buffer starts with a PKZIP local/central header
 */
function looksLikeZip(bytes) {
  if (bytes === undefined || bytes === null || bytes.length < 4) return false
  return ZIP_MAGICS.some((magic) => magic.every((byte, index) => bytes[index] === byte))
}

/**
 * Normalize a package part path: strip a leading slash, collapse `.`/`..` segments.
 * @param {string} part
 * @returns {string}
 */
export function normalizePart(part) {
  let value = String(part ?? '').trim()
  if (value.startsWith('/')) value = value.slice(1)
  if (value.length === 0) return ''
  const normalized = posix.normalize(value)
  return normalized === '.' ? '' : normalized
}

/**
 * Resolve a relationship target against the part that owns the relationship.
 * Handles absolute (`/ppt/...`), relative (`../media/x.png`) and percent-encoded
 * targets. Returns `null` for external targets (`TargetMode="External"`).
 *
 * @param {string} sourcePart e.g. `ppt/slideMasters/slideMaster1.xml`
 * @param {string} target raw `Target` attribute value
 * @param {string} [mode] raw `TargetMode` attribute value
 * @returns {string | null} normalized zip path, or null when external/unresolvable
 */
export function resolveRelTarget(sourcePart, target, mode) {
  if (typeof target !== 'string' || target.length === 0) return null
  if (String(mode ?? '').toLowerCase() === 'external') return null
  if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return null // http:, https:, file: …
  let decoded = target
  if (/%[0-9a-fA-F]{2}/.test(decoded)) {
    try {
      decoded = decodeURIComponent(decoded)
    } catch {
      /* keep the raw target: a stray % is not fatal */
    }
  }
  const baseDir = posix.dirname(normalizePart(sourcePart))
  const joined = decoded.startsWith('/') ? decoded.slice(1) : posix.join(baseDir, decoded)
  const normalized = normalizePart(joined)
  return normalized.length === 0 ? null : normalized
}

/**
 * Compute the `.rels` part path that holds the relationships of `sourcePart`:
 * `ppt/slides/slide1.xml` → `ppt/slides/_rels/slide1.xml.rels`.
 * @param {string} sourcePart
 * @returns {string}
 */
export function relsPathFor(sourcePart) {
  const clean = normalizePart(sourcePart)
  const dir = posix.dirname(clean)
  const file = posix.basename(clean)
  return dir === '.' || dir === '' ? `_rels/${file}.rels` : `${dir}/_rels/${file}.rels`
}

/**
 * @typedef {object} Relationship
 * @property {string} id            `Id` attribute (e.g. `rId6`)
 * @property {string} type          relationship type URI
 * @property {string} target        raw `Target` attribute
 * @property {string} mode          `Internal` | `External`
 * @property {string | null} resolvedPart normalized zip path, `null` when external
 */

/**
 * @typedef {object} PackageHandle
 * @property {() => string[]} partNames       all file parts (no directory entries)
 * @property {(part: string) => boolean} has
 * @property {(part: string) => Buffer} buffer
 * @property {(part: string) => string} text
 * @property {(part: string) => any} xml
 * @property {(sourcePart: string) => Relationship[]} relationships
 * @property {number} byteLength             size of the original package
 */

/**
 * Open a `.pptx` (or any OPC zip) package.
 *
 * @param {{bytes?: Buffer|Uint8Array} | {path?: string} | Buffer | Uint8Array | string} input
 *   `{path}`/string → read from disk; `{bytes}`/Buffer → use as-is.
 * @returns {Promise<PackageHandle>}
 * @throws {Error & {code:'not_a_pptx'|'missing_input'|'part_not_found'}}
 */
export async function loadPackage(input) {
  const { bytes, sourcePath } = await readInput(input)
  if (!looksLikeZip(bytes)) {
    throw pptxError(
      'not_a_pptx',
      `not a PowerPoint (zip) package: ${sourcePath ?? 'buffer'} does not start with a PKZIP header`,
    )
  }

  /** @type {JSZip} */
  let zip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch (cause) {
    throw pptxError('not_a_pptx', `not a PowerPoint (zip) package: unable to unzip ${sourcePath ?? 'buffer'}`, cause)
  }

  /** @type {Map<string, Uint8Array>} */
  const rawParts = new Map()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir === true) continue
    rawParts.set(normalizePart(name), await entry.async('uint8array'))
  }

  if (!rawParts.has('ppt/presentation.xml')) {
    throw pptxError(
      'not_a_pptx',
      `not a PowerPoint presentation: ${sourcePath ?? 'package'} has no ppt/presentation.xml ` +
        `(${rawParts.size} parts found)`,
    )
  }

  const names = [...rawParts.keys()].sort()
  const textCache = new Map()
  const xmlCache = new Map()
  const relCache = new Map()
  const bufferCache = new Map()

  /** @param {string} part */
  const requirePart = (part) => {
    const key = normalizePart(part)
    const found = rawParts.get(key)
    if (found === undefined) throw pptxError('part_not_found', `package part not found: ${part}`)
    return /** @type {[string, Uint8Array]} */ ([key, found])
  }

  /** @type {PackageHandle} */
  const handle = {
    partNames: () => [...names],
    has: (part) => rawParts.has(normalizePart(part)),
    buffer: (part) => {
      const [key, raw] = requirePart(part)
      let cached = bufferCache.get(key)
      if (cached === undefined) {
        cached = Buffer.from(raw)
        bufferCache.set(key, cached)
      }
      return cached
    },
    text: (part) => {
      const [key, raw] = requirePart(part)
      let cached = textCache.get(key)
      if (cached === undefined) {
        cached = Buffer.from(raw).toString('utf8')
        if (cached.charCodeAt(0) === 0xfeff) cached = cached.slice(1) // BOM
        textCache.set(key, cached)
      }
      return cached
    },
    xml: (part) => {
      const [key] = requirePart(part)
      let cached = xmlCache.get(key)
      if (cached === undefined) {
        cached = parseXml(handle.text(key))
        xmlCache.set(key, cached)
      }
      return cached
    },
    relationships: (sourcePart) => {
      const source = normalizePart(sourcePart)
      const cached = relCache.get(source)
      if (cached !== undefined) return cached
      const relsPart = relsPathFor(source)
      /** @type {Relationship[]} */
      const resolved = []
      if (rawParts.has(relsPart)) {
        const doc = handle.xml(relsPart)
        const nodes = asArray(doc?.Relationships?.Relationship)
        for (const node of nodes) {
          const id = String(node?.['@_Id'] ?? '')
          const target = String(node?.['@_Target'] ?? '')
          const modeRaw = node?.['@_TargetMode']
          const mode = modeRaw === undefined ? 'Internal' : String(modeRaw)
          const resolvedPart = resolveRelTarget(source, target, mode)
          resolved.push({
            id,
            type: String(node?.['@_Type'] ?? ''),
            target,
            mode,
            resolvedPart,
          })
        }
      }
      relCache.set(source, resolved)
      return resolved
    },
    byteLength: bytes.length,
  }

  return handle
}

/**
 * Resolve the `loadPackage` input into bytes (+ a label for error messages).
 * @param {any} input
 * @returns {Promise<{bytes: Uint8Array, sourcePath?: string}>}
 */
async function readInput(input) {
  if (Buffer.isBuffer(input) || input instanceof Uint8Array) {
    return { bytes: input }
  }
  if (typeof input === 'string') {
    return { bytes: await readFileBytes(input), sourcePath: input }
  }
  if (input !== null && typeof input === 'object') {
    if (input.bytes !== undefined) return { bytes: toBytes(input.bytes) }
    if (typeof input.path === 'string') {
      return { bytes: await readFileBytes(input.path), sourcePath: input.path }
    }
  }
  throw pptxError('missing_input', 'template input missing: pass { path } or { bytes }')
}

/** @param {any} value */
function toBytes(value) {
  if (Buffer.isBuffer(value)) return value
  if (value instanceof Uint8Array) return value
  if (value instanceof ArrayBuffer) return new Uint8Array(value)
  throw pptxError('missing_input', `unsupported byte input type: ${typeof value}`)
}

/** @param {string} path */
async function readFileBytes(path) {
  const { readFile } = await import('node:fs/promises')
  try {
    return await readFile(path)
  } catch (cause) {
    throw pptxError('template_not_found', `cannot read template file: ${path}`, cause)
  }
}
