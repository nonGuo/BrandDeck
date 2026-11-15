/**
 * Media inventory + brand-logo discovery + image size sniffing.
 *
 * Traps this module handles (see `research/template-report.md` §7):
 * - "media part referenced by exactly one master" is **not** a usable logo heuristic:
 *   `PPT模板-HW.pptx` references `image1.png` from slideMaster1 *and* slideMaster4, and
 *   `image6.png` (the same artwork at 242 px) from slideMaster3.
 *   The rule we use instead: candidates are media parts whose referencing relationships
 *   are **all** master relationships (any number of masters, zero layout/slide refs),
 *   ranked by how many masters place them, then by whether a master's shape tree holds
 *   them in a `<p:pic>` (a `userDrawn` picture with no `<p:ph>`).
 * - The real logo sits **bottom-right** on the master (x 10.07in, y 6.54in), not in a
 *   top-right corner. We report the master's frame verbatim in inches and let the
 *   compiler honour it.
 * - Pixel sizes are read from the file headers (PNG IHDR / JPEG SOFn / GIF / BMP),
 *   never from the part path or a DPI extension (`<a14:useLocalDpi>`).
 */
import { posix } from 'node:path'
import { asArray } from './read.js'
import { discoverStructure, EMU_PER_INCH } from './theme.js'

/** Extensions we know how to label even when the magic bytes are unsupported. */
const EXTENSION_TYPES = Object.freeze({
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  svg: 'image/svg+xml',
  ico: 'image/x-icon',
  webp: 'image/webp',
})

/** Media part prefix inside an OPC package. */
const MEDIA_PREFIX = 'ppt/media/'

/* ------------------------------------------------------------------ *
 * Image size sniffing
 * ------------------------------------------------------------------ */

/**
 * Read pixel dimensions out of an image header. Never throws: unknown formats return
 * `widthPx: null, heightPx: null`.
 *
 * @param {Buffer | Uint8Array | ArrayBuffer} bytes
 * @param {string} [partPath] used only to label the format when the magic is unknown
 * @returns {{format: string, contentType: string, widthPx: number | null, heightPx: number | null}}
 */
export function imageSize(bytes, partPath = '') {
  const extension = posix.extname(String(partPath)).replace('.', '').toLowerCase()
  /** @type {{format: string, contentType: string, widthPx: number | null, heightPx: number | null}} */
  const fallback = {
    format: extension.length > 0 ? extension : 'unknown',
    contentType: EXTENSION_TYPES[extension] ?? 'application/octet-stream',
    widthPx: null,
    heightPx: null,
  }

  let buffer
  try {
    buffer = toBuffer(bytes)
  } catch {
    return fallback
  }
  if (buffer.length < 10) return fallback

  try {
    const png = readPngSize(buffer)
    if (png !== null) return { format: 'png', contentType: 'image/png', ...png }

    const jpeg = readJpegSize(buffer)
    if (jpeg !== null) return { format: 'jpeg', contentType: 'image/jpeg', ...jpeg }

    const gif = readGifSize(buffer)
    if (gif !== null) return { format: 'gif', contentType: 'image/gif', ...gif }

    const bmp = readBmpSize(buffer)
    if (bmp !== null) return { format: 'bmp', contentType: 'image/bmp', ...bmp }

    const vector = readVectorFormat(buffer)
    if (vector !== null) {
      return { format: vector, contentType: EXTENSION_TYPES[vector] ?? 'application/octet-stream', widthPx: null, heightPx: null }
    }
  } catch {
    // A malformed header must never break template parsing.
    return fallback
  }
  return fallback
}

/** @param {Buffer | Uint8Array | ArrayBuffer} bytes */
function toBuffer(bytes) {
  if (Buffer.isBuffer(bytes)) return bytes
  if (bytes instanceof Uint8Array) return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (bytes instanceof ArrayBuffer) return Buffer.from(bytes)
  throw new TypeError('not bytes')
}

const PNG_MAGIC = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]

/** @param {Buffer} buffer */
function readPngSize(buffer) {
  if (buffer.length < 24) return null
  if (!PNG_MAGIC.every((byte, index) => buffer[index] === byte)) return null
  if (buffer.toString('latin1', 12, 16) !== 'IHDR') return null
  return { widthPx: buffer.readUInt32BE(16), heightPx: buffer.readUInt32BE(20) }
}

/** SOF markers that carry the frame size (0xC4/0xC8/0xCC are not SOF). */
const SOF_MARKERS = new Set([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf])

/** @param {Buffer} buffer */
function readJpegSize(buffer) {
  if (buffer.length < 4 || buffer[0] !== 0xff || buffer[1] !== 0xd8) return null
  let offset = 2
  while (offset + 3 < buffer.length) {
    if (buffer[offset] !== 0xff) {
      offset += 1
      continue
    }
    const marker = buffer[offset + 1]
    if (marker === 0xff) {
      offset += 1
      continue
    }
    // Standalone markers (no payload length).
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      offset += 2
      continue
    }
    if (marker === 0xd9 || marker === 0xda) return null // end of image / start of scan
    if (offset + 8 >= buffer.length) return null
    if (SOF_MARKERS.has(marker)) {
      return { widthPx: buffer.readUInt16BE(offset + 7), heightPx: buffer.readUInt16BE(offset + 5) }
    }
    const length = buffer.readUInt16BE(offset + 2)
    if (length <= 0) return null
    offset += 2 + length
  }
  return null
}

/** @param {Buffer} buffer */
function readGifSize(buffer) {
  if (buffer.length < 10) return null
  const header = buffer.toString('latin1', 0, 6)
  if (header !== 'GIF87a' && header !== 'GIF89a') return null
  return { widthPx: buffer.readUInt16LE(6), heightPx: buffer.readUInt16LE(8) }
}

/** @param {Buffer} buffer */
function readBmpSize(buffer) {
  if (buffer.length < 26) return null
  if (buffer[0] !== 0x42 || buffer[1] !== 0x4d) return null // 'BM'
  const widthPx = buffer.readInt32LE(18)
  const heightPx = Math.abs(buffer.readInt32LE(22))
  if (widthPx <= 0 || heightPx <= 0) return null
  return { widthPx, heightPx }
}

/** EMF/WMF are vector containers: we knowingly report no pixel size. */
function readVectorFormat(buffer) {
  if (buffer.length >= 44 && buffer.readUInt32LE(0) === 1 && buffer.toString('latin1', 40, 44) === ' EMF') {
    return 'emf'
  }
  if (buffer.length >= 4 && buffer[0] === 0xd7 && buffer[1] === 0xcd) return 'wmf'
  if (buffer.length >= 4 && buffer.readUInt32LE(0) === 0x9ac6cdd7) return 'wmf'
  if (buffer.length >= 4) {
    const head = buffer.toString('latin1', 0, 4)
    if (head === 'II*\u0000' || head === 'MM\u0000*') return 'tiff'
    if (head === 'RIFF' && buffer.toString('latin1', 8, 12) === 'WEBP') return 'webp'
  }
  return null
}

/* ------------------------------------------------------------------ *
 * Media inventory
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} MediaRef
 * @property {string} relsPart     e.g. `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @property {string} sourcePart   e.g. `ppt/slideMasters/slideMaster1.xml`
 * @property {string} sourceKind   `master` | `layout` | `slide` | `notesMaster` | …
 * @property {string} relId
 */

/**
 * @typedef {object} MediaPart
 * @property {string} path
 * @property {number} bytes            byte length (matches expected-tokens.json)
 * @property {Buffer} buffer           the raw bytes (not JSON serializable)
 * @property {string} contentType
 * @property {string} format
 * @property {number | null} widthPx
 * @property {number | null} heightPx
 * @property {MediaRef[]} refs
 * @property {string[]} referencedBy   `.rels` part paths (JSON friendly, as in ARCHITECTURE §3)
 */

/**
 * List every media part with its pixel size and the relationships that reference it.
 * @param {import('./read.js').PackageHandle} pkg
 * @returns {MediaPart[]}
 */
export function listMedia(pkg) {
  /** @type {Map<string, MediaRef[]>} */
  const references = new Map()
  for (const part of pkg.partNames()) {
    if (!part.includes('_rels/') || !part.endsWith('.rels')) continue
    const sourcePart = sourcePartOfRels(part)
    if (sourcePart === null) continue
    const sourceKind = classifySource(sourcePart)
    for (const rel of pkg.relationships(sourcePart)) {
      if (rel.resolvedPart === null) continue
      if (!rel.resolvedPart.startsWith(MEDIA_PREFIX)) continue
      const entries = references.get(rel.resolvedPart) ?? []
      entries.push({ relsPart: part, sourcePart, sourceKind, relId: rel.id })
      references.set(rel.resolvedPart, entries)
    }
  }

  /** @type {MediaPart[]} */
  const media = []
  for (const part of pkg.partNames()) {
    if (!part.startsWith(MEDIA_PREFIX)) continue
    const buffer = pkg.buffer(part)
    const size = imageSize(buffer, part)
    const refs = references.get(part) ?? []
    media.push({
      path: part,
      bytes: buffer.length,
      buffer,
      contentType: size.contentType,
      format: size.format,
      widthPx: size.widthPx,
      heightPx: size.heightPx,
      refs,
      referencedBy: [...new Set(refs.map((ref) => ref.relsPart))],
    })
  }
  return media.sort((a, b) => a.path.localeCompare(b.path, 'en'))
}

/**
 * `ppt/slides/_rels/slide1.xml.rels` → `ppt/slides/slide1.xml`.
 * @param {string} relsPart
 * @returns {string | null}
 */
function sourcePartOfRels(relsPart) {
  const match = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(relsPart)
  if (match === null) return null
  return `${match[1]}/${match[2]}`
}

/** @param {string} sourcePart */
function classifySource(sourcePart) {
  if (sourcePart.startsWith('ppt/slideMasters/')) return 'master'
  if (sourcePart.startsWith('ppt/slideLayouts/')) return 'layout'
  if (sourcePart.startsWith('ppt/slides/')) return 'slide'
  if (sourcePart.startsWith('ppt/notesMasters/')) return 'notesMaster'
  if (sourcePart.startsWith('ppt/notesSlides/')) return 'notesSlide'
  if (sourcePart.startsWith('ppt/handoutMasters/')) return 'handoutMaster'
  return 'other'
}

/* ------------------------------------------------------------------ *
 * Master shape-tree walking (for `<p:pic>` placement + frame geometry)
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} PicturePlacement
 * @property {string} masterPart
 * @property {string} relId
 * @property {string} mediaPart
 * @property {{x: number, y: number, w: number, h: number} | null} frameInches
 * @property {string | null} name
 * @property {boolean} hasPlaceholder
 */

/**
 * Find every picture placed by a master's shape tree, mapped to its media part.
 * Walks nested `<p:grpSp>` groups.
 * @param {import('./read.js').PackageHandle} pkg
 * @param {string} masterPart
 * @returns {PicturePlacement[]}
 */
function masterPictures(pkg, masterPart) {
  const doc = pkg.xml(masterPart)
  const spTree = doc?.sldMaster?.cSld?.spTree
  if (spTree === undefined || spTree === null) return []
  const relById = new Map(pkg.relationships(masterPart).map((rel) => [rel.id, rel]))
  /** @type {PicturePlacement[]} */
  const pictures = []

  /** @param {any} container */
  const walk = (container) => {
    for (const pic of asArray(container?.pic)) {
      const embed = pic?.blipFill?.blip?.['@_embed'] ?? pic?.blipFill?.blip?.['@_link']
      const rel = embed === undefined ? undefined : relById.get(String(embed))
      const mediaPart = rel?.resolvedPart ?? null
      if (mediaPart === null || mediaPart === undefined) continue
      pictures.push({
        masterPart,
        relId: String(embed),
        mediaPart,
        frameInches: readFrameInches(pic?.spPr?.xfrm),
        name: pic?.nvPicPr?.cNvPr?.['@_name'] === undefined ? null : String(pic.nvPicPr.cNvPr['@_name']),
        hasPlaceholder: pic?.nvPicPr?.nvPr?.ph !== undefined,
      })
    }
    for (const group of asArray(container?.grpSp)) walk(group)
  }

  walk(spTree)
  return pictures
}

/**
 * `<a:xfrm><a:off x y/><a:ext cx cy/></a:xfrm>` → inches (4 dp).
 * @param {any} xfrm
 * @returns {{x: number, y: number, w: number, h: number} | null}
 */
function readFrameInches(xfrm) {
  const off = xfrm?.off
  const ext = xfrm?.ext
  if (off === undefined || ext === undefined) return null
  const x = Number(off['@_x'])
  const y = Number(off['@_y'])
  const w = Number(ext['@_cx'])
  const h = Number(ext['@_cy'])
  if (![x, y, w, h].every(Number.isFinite)) return null
  return toInches({ x, y, w, h })
}

/**
 * Convert an EMU rectangle to inches.
 * @param {{x: number, y: number, w: number, h: number}} emu
 */
function toInches(emu) {
  const convert = (value) => Math.round((value / EMU_PER_INCH) * 10000) / 10000
  return { x: convert(emu.x), y: convert(emu.y), w: convert(emu.w), h: convert(emu.h) }
}

/* ------------------------------------------------------------------ *
 * Brand logo discovery
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} BrandLogo
 * @property {string} partPath
 * @property {number} bytes
 * @property {Buffer} buffer
 * @property {number} widthPx
 * @property {number} heightPx
 * @property {number} aspect
 * @property {{x: number, y: number, w: number, h: number} | null} frameInches
 * @property {string | null} masterPart   master that places the picture (null when only related)
 * @property {string[]} masterParts       every master relationship pointing at the part
 * @property {string} contentType
 */

/** Plausibility window for a logo lockup. */
const LOGO_LIMITS = Object.freeze({
  minAspect: 0.8,
  maxAspect: 8,
  minWidthPx: 40,
  maxWidthPx: 1200,
  minHeightPx: 16,
  maxHeightPx: 400,
})

/**
 * Locate the corporate logo.
 *
 * Rules (documented because the naive "one master" rule fails on the real template):
 * 1. Candidate = media part whose referencing relationships are all master rels
 *    (no layout/slide/notes refs).
 * 2. Keep parts with plausible lockup dimensions
 *    (aspect 0.8…8, width 40…1200 px, height 16…400 px).
 * 3. Rank: number of masters placing it (desc) → placed by a master `<p:pic>` (desc) →
 *    aspect closeness to a 3.5:1 lockup → pixel width (desc) → byte size (asc, a
 *    low-resolution duplicate must lose to the canonical part).
 * 4. Frame geometry = the first placing master's `<a:xfrm>` in inches.
 * 5. If no candidate is master-only, fall back to the smallest master-referenced image
 *    so callers still get *something* brand-ish; `null` when nothing is plausible.
 *
 * @param {import('./read.js').PackageHandle} pkg
 * @returns {BrandLogo | null}
 */
export function findBrandLogo(pkg) {
  const media = listMedia(pkg)
  if (media.length === 0) return null
  const structure = discoverStructure(pkg)

  /** Every `<p:pic>` placement across masters, keyed by media part. */
  /** @type {Map<string, PicturePlacement[]>} */
  const placements = new Map()
  for (const master of structure.masters) {
    for (const picture of masterPictures(pkg, master.part)) {
      const entries = placements.get(picture.mediaPart) ?? []
      entries.push(picture)
      placements.set(picture.mediaPart, entries)
    }
  }

  /** @type {(entry: MediaPart) => boolean} */
  const hasPlausibleSize = (entry) => {
    if (entry.widthPx === null || entry.heightPx === null) return false
    if (entry.widthPx < LOGO_LIMITS.minWidthPx || entry.widthPx > LOGO_LIMITS.maxWidthPx) return false
    if (entry.heightPx < LOGO_LIMITS.minHeightPx || entry.heightPx > LOGO_LIMITS.maxHeightPx) return false
    const aspect = entry.widthPx / entry.heightPx
    return aspect >= LOGO_LIMITS.minAspect && aspect <= LOGO_LIMITS.maxAspect
  }

  /** @type {(entry: MediaPart) => boolean} */
  const isMasterOnly = (entry) =>
    entry.refs.length > 0 && entry.refs.every((ref) => ref.sourceKind === 'master')

  /** @type {(entry: MediaPart) => boolean} */
  const isMasterReferenced = (entry) => entry.refs.some((ref) => ref.sourceKind === 'master')

  const masterPartsOf = (entry) => [...new Set(entry.refs.map((ref) => ref.sourcePart))].sort()

  /**
   * @param {MediaPart} entry
   * @param {PicturePlacement[]} entryPlacements
   * @returns {BrandLogo}
   */
  const toLogo = (entry, entryPlacements) => {
    const ordered = [...entryPlacements].sort(
      (a, b) =>
        structure.masters.findIndex((master) => master.part === a.masterPart) -
        structure.masters.findIndex((master) => master.part === b.masterPart),
    )
    const placed = ordered[0] ?? null
    return {
      partPath: entry.path,
      bytes: entry.bytes,
      buffer: entry.buffer,
      widthPx: /** @type {number} */ (entry.widthPx),
      heightPx: /** @type {number} */ (entry.heightPx),
      aspect: Number((/** @type {number} */ (entry.widthPx) / /** @type {number} */ (entry.heightPx)).toFixed(4)),
      frameInches: placed?.frameInches ?? null,
      masterPart: placed?.masterPart ?? null,
      masterParts: masterPartsOf(entry),
      contentType: entry.contentType,
    }
  }

  const masterOnly = media.filter((entry) => isMasterOnly(entry) && hasPlausibleSize(entry))
  if (masterOnly.length > 0) {
    const ranked = [...masterOnly].sort((a, b) => {
      const mastersA = masterPartsOf(a).length
      const mastersB = masterPartsOf(b).length
      if (mastersA !== mastersB) return mastersB - mastersA
      const placedA = placements.get(a.path)?.length ?? 0
      const placedB = placements.get(b.path)?.length ?? 0
      if (placedA !== placedB) return placedB - placedA
      const aspectA = /** @type {number} */ (a.widthPx) / /** @type {number} */ (a.heightPx)
      const aspectB = /** @type {number} */ (b.widthPx) / /** @type {number} */ (b.heightPx)
      const aspectDelta = Math.abs(Math.log(aspectA / 3.5)) - Math.abs(Math.log(aspectB / 3.5))
      if (Math.abs(aspectDelta) > 1e-6) return aspectDelta
      if (a.widthPx !== b.widthPx) return /** @type {number} */ (b.widthPx) - /** @type {number} */ (a.widthPx)
      if (a.bytes !== b.bytes) return a.bytes - b.bytes
      return a.path.localeCompare(b.path, 'en')
    })
    return toLogo(ranked[0], placements.get(ranked[0].path) ?? [])
  }

  // Fallback: a small image that at least one master references (it may also be used by
  // a layout, which is common for a header/footer strip). Prefer the smallest pixel area.
  const fallback = media.filter((entry) => isMasterReferenced(entry) && entry.widthPx !== null)
  if (fallback.length === 0) return null
  const smallest = [...fallback].sort((a, b) => {
    const areaA = /** @type {number} */ (a.widthPx) * /** @type {number} */ (a.heightPx)
    const areaB = /** @type {number} */ (b.widthPx) * /** @type {number} */ (b.heightPx)
    if (areaA !== areaB) return areaA - areaB
    return a.path.localeCompare(b.path, 'en')
  })[0]
  return toLogo(smallest, placements.get(smallest.path) ?? [])
}
