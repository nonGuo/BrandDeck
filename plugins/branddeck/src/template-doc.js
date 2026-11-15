/**
 * Template documentation: the **layouts inventory** (`layouts.json`) and the
 * human/model-readable spec (`template.md`).
 *
 * Why this module exists
 * ----------------------
 * `theme.json` answers "which colours and faces does the brand use?" but not "which
 * page shapes does the deck offer?". A model authoring a deck needs the latter, and a
 * reviewer needs a file they can read without opening PowerPoint. Both are derived
 * **once** when a template is parsed, persisted next to `theme.json`, and reused by
 * later runs (`branddeck_deck_create { templateId }`) without re-parsing the `.pptx`.
 *
 * Contract
 * --------
 * - `buildLayoutsInventory(pkg, structure)` → `{ slideSize, masters, counts }`
 * - `buildTemplateMarkdown({ theme, meta, media, layouts })` → markdown string
 * - `buildTemplateDoc({ pkg, structure, theme, meta, media })` → `{ markdown, layouts }`
 *
 * Geometry decisions (deliberate, see report)
 * -------------------------------------------
 * 1. **Inches, 4 decimals.** EMU → inches is `value / 914400` (OOXML fixed point),
 *    rounded to 4 decimals so the JSON is stable and readable. `EMU_PER_INCH` is
 *    imported from `theme.js` rather than re-declared.
 * 2. **Unusable `<a:xfrm>` is skipped.** A shape needs `<a:off x y/>` **and**
 *    `<a:ext cx cy/>` with four finite numbers and a positive width/height;
 *    placeholders that only inherit their frame from the master are therefore not
 *    reported (they carry no geometry of their own).
 * 3. **Out-of-canvas geometry is CLAMPED, never dropped.** `PPT模板-HW.pptx` has a
 *    decorative group on `slideMaster3` at x ≈ 13.36 in, i.e. just past the
 *    13.3385 in page width. Every rectangle this module emits is fitted inside
 *    `slideSize`: width/height are capped to the page, then the origin is pulled back
 *    so `x + w <= widthIn` and `y + h <= heightIn`. Clamping (rather than skipping)
 *    keeps a real element visible on the page edge instead of silently losing it, and
 *    it guarantees every number we hand downstream is inside the slide. The group
 *    itself is never emitted — only its descendant placeholders/pictures, each with
 *    its own `<a:xfrm>` — so decorative master groups cannot leak off-canvas shapes
 *    into the inventory.
 * 4. **Order is document order per element kind.** fast-xml-parser keeps sibling
 *    elements in arrays but cannot interleave `sp`/`pic`/`graphicFrame` order, so we
 *    walk kinds in a fixed order (`sp`, `pic`, `graphicFrame`, then nested `grpSp`).
 *    That is deterministic, which is what the persisted inventory needs.
 *
 * The markdown is Chinese-labelled by design (the template is a Chinese corporate
 * deck), deterministic, and never prints `undefined`/`NaN`/`[object Object]`.
 */
import { asArray } from './pptx/read.js'
import { EMU_PER_INCH, discoverStructure, hexColor } from './pptx/theme.js'
import { listMedia } from './pptx/media.js'

/** Decimal places kept for inch geometry (EMU resolution is far finer than this). */
const INCH_DECIMALS = 4
const INCH_FACTOR = 10 ** INCH_DECIMALS

/** Canonical 16:9 page, used only when `presentation.xml` has no `<p:sldSz>`. */
const FALLBACK_SLIDE = Object.freeze({ cx: 12192000, cy: 6858000 })

/** Maximum placeholder bullets printed per layout before the `…` marker. */
const PLACEHOLDER_LINE_LIMIT = 8

/** Marker used instead of an absent value (never `undefined`). */
const DASH = '—'

/* ------------------------------------------------------------------ *
 * Small value helpers
 * ------------------------------------------------------------------ */

/**
 * @param {unknown} value
 * @returns {string | undefined} trimmed non-empty string, else undefined
 */
function trimmed(value) {
  if (typeof value !== 'string') return undefined
  const text = value.trim()
  return text.length > 0 ? text : undefined
}

/**
 * @param {unknown} value
 * @returns {number | undefined}
 */
function finiteNumber(value) {
  if (value === undefined || value === null || value === '') return undefined
  const number = Number(value)
  return Number.isFinite(number) ? number : undefined
}

/**
 * Round to 4 decimals without producing `-0`.
 * @param {number} value
 * @returns {number}
 */
function round4(value) {
  const rounded = Math.round(value * INCH_FACTOR) / INCH_FACTOR
  return Object.is(rounded, -0) ? 0 : rounded
}

/**
 * Format a number for markdown: 4 decimals max, trailing zeros dropped, never `NaN`.
 * @param {unknown} value
 * @returns {string}
 */
function fmt(value) {
  const number = finiteNumber(value)
  if (number === undefined) return '0'
  return String(round4(number))
}

/** `ppt/slideMasters/slideMaster1.xml` → `slideMaster1.xml` */
function baseName(part) {
  const text = String(part ?? '')
  const index = text.lastIndexOf('/')
  return index === -1 ? text : text.slice(index + 1)
}

/* ------------------------------------------------------------------ *
 * Package readers (pure: no disk access)
 * ------------------------------------------------------------------ */

/**
 * Read `<p:sldSz>` from `ppt/presentation.xml` and convert to inches.
 *
 * The real template is `cx=12196763` (13.3385 in), **not** the canonical 12192000 —
 * the page size is always read, never assumed.
 *
 * @param {import('./pptx/read.js').PackageHandle} pkg
 * @returns {{cx: number, cy: number, widthIn: number, heightIn: number}}
 */
function readSlideSize(pkg) {
  const doc = pkg.has('ppt/presentation.xml') ? pkg.xml('ppt/presentation.xml') : null
  const sldSz = doc?.presentation?.sldSz ?? doc?.sldSz ?? null
  const cx = finiteNumber(sldSz?.['@_cx']) ?? FALLBACK_SLIDE.cx
  const cy = finiteNumber(sldSz?.['@_cy']) ?? FALLBACK_SLIDE.cy
  return {
    cx,
    cy,
    widthIn: round4(cx / EMU_PER_INCH),
    heightIn: round4(cy / EMU_PER_INCH),
  }
}

/**
 * The solid background a part paints itself (`<p:cSld><p:bg><p:bgPr><a:solidFill>`).
 * Only `<a:srgbClr>` is resolved: a `<a:schemeClr>` needs the theme's colour map,
 * which this module deliberately does not re-derive (the token layer owns that).
 *
 * @param {import('./pptx/read.js').PackageHandle} pkg
 * @param {string} part
 * @returns {string | undefined} `#RRGGBB`
 */
function readPaintedBackground(pkg, part) {
  if (!pkg.has(part)) return undefined
  const doc = pkg.xml(part)
  const root = doc?.sldMaster ?? doc?.sldLayout ?? doc?.sld
  const solidFill = root?.cSld?.bg?.bgPr?.solidFill ?? root?.bg?.bgPr?.solidFill
  return hexColor(solidFill?.srgbClr?.['@_val'])
}

/**
 * `<a:xfrm><a:off x y/><a:ext cx cy/></a:xfrm>` → inches (unclamped).
 *
 * Returns `null` for a missing/incomplete frame or a non-positive extent: such a
 * shape has no usable geometry and is skipped.
 *
 * @param {any} xfrm
 * @returns {{x: number, y: number, w: number, h: number} | null}
 */
function readRectInches(xfrm) {
  const off = xfrm?.off
  const ext = xfrm?.ext
  if (off === undefined || off === null || ext === undefined || ext === null) return null
  const x = Number(off['@_x'])
  const y = Number(off['@_y'])
  const w = Number(ext['@_cx'])
  const h = Number(ext['@_cy'])
  if (![x, y, w, h].every(Number.isFinite)) return null
  if (w <= 0 || h <= 0) return null
  return { x: x / EMU_PER_INCH, y: y / EMU_PER_INCH, w: w / EMU_PER_INCH, h: h / EMU_PER_INCH }
}

/**
 * Fit a rectangle inside the slide (the documented CLAMP rule).
 *
 * Width/height are capped to the page first, then the origin is pulled back so the
 * whole box stays on-canvas. The final re-fit after rounding guarantees the invariant
 * even when decimal rounding nudges an edge.
 *
 * @param {{x: number, y: number, w: number, h: number}} rect inches
 * @param {{widthIn: number, heightIn: number}} slide
 * @returns {{x: number, y: number, w: number, h: number}} inches, 4 decimals
 */
function clampToSlide(rect, slide) {
  const width = Math.min(Math.max(rect.w, 0), slide.widthIn)
  const height = Math.min(Math.max(rect.h, 0), slide.heightIn)
  const x = Math.min(Math.max(rect.x, 0), Math.max(0, slide.widthIn - width))
  const y = Math.min(Math.max(rect.y, 0), Math.max(0, slide.heightIn - height))
  const clamped = { x: round4(x), y: round4(y), w: round4(width), h: round4(height) }
  if (clamped.x + clamped.w > slide.widthIn) {
    clamped.x = Math.max(0, round4(slide.widthIn - clamped.w))
  }
  if (clamped.y + clamped.h > slide.heightIn) {
    clamped.y = Math.max(0, round4(slide.heightIn - clamped.h))
  }
  return clamped
}

/**
 * Master display name: an explicit `@name` if present, else the name of the master's
 * own theme (`1_Title Slide`, `Chart page`, …), else the part's file name.
 *
 * Office does not store master names, so the per-master theme name is the closest
 * human label the file actually carries; the part path is still printed alongside it
 * in the markdown so nothing is ambiguous.
 *
 * @param {import('./pptx/read.js').PackageHandle} pkg
 * @param {string} part
 * @returns {string}
 */
function readMasterName(pkg, part) {
  const doc = pkg.has(part) ? pkg.xml(part) : null
  const declared =
    trimmed(doc?.sldMaster?.['@_name']) ?? trimmed(doc?.sldMaster?.cSld?.['@_name'])
  if (declared !== undefined) return declared
  const themeRel = pkg.relationships(part).find((rel) => rel.type.endsWith('/theme'))
  if (themeRel?.resolvedPart !== null && themeRel?.resolvedPart !== undefined && pkg.has(themeRel.resolvedPart)) {
    const themeName = trimmed(pkg.xml(themeRel.resolvedPart)?.theme?.['@_name'])
    if (themeName !== undefined) return themeName
  }
  return baseName(part)
}

/**
 * Layout display name: `<p:sldLayout><p:cSld name="…">` (`攀登`, `End page`, …),
 * falling back to `@name`, the standard `@type`, then the file name.
 * @param {import('./pptx/read.js').PackageHandle} pkg
 * @param {string} part
 * @returns {string}
 */
function readLayoutName(pkg, part) {
  const doc = pkg.has(part) ? pkg.xml(part) : null
  const root = doc?.sldLayout
  const declared =
    trimmed(root?.cSld?.['@_name']) ?? trimmed(root?.['@_name']) ?? trimmed(root?.['@_type'])
  return declared ?? baseName(part)
}

/**
 * Collect the placeholders of a part (layout or master).
 *
 * Placeholders are `<p:sp>`, `<p:pic>` and `<p:graphicFrame>` nodes carrying a
 * `<p:ph>`; nested `<p:grpSp>` groups are walked. A placeholder whose `<a:xfrm>` is
 * missing/unusable is skipped (it inherits its frame from the master, so there is
 * nothing to report), and every emitted frame is clamped to the slide.
 *
 * @param {import('./pptx/read.js').PackageHandle} pkg
 * @param {string} part
 * @param {{widthIn: number, heightIn: number}} slide
 * @returns {Array<{type: string, idx?: number, name?: string, x: number, y: number, w: number, h: number}>}
 */
function readPlaceholders(pkg, part, slide) {
  if (!pkg.has(part)) return []
  const doc = pkg.xml(part)
  const root = doc?.sldLayout ?? doc?.sldMaster ?? null
  /** @type {Array<{type: string, idx?: number, name?: string, x: number, y: number, w: number, h: number}>} */
  const placeholders = []

  /**
   * @param {any} node shape node
   * @param {any} ph `<p:ph>` node
   * @param {any} xfrm the node's own `<a:xfrm>`
   */
  const add = (node, ph, xfrm) => {
    if (ph === undefined || ph === null) return
    const rect = readRectInches(xfrm)
    if (rect === null) return
    const frame = clampToSlide(rect, slide)
    // `<p:ph>` without @type means "obj" in OOXML (a generic content placeholder).
    const type = trimmed(ph['@_type']) ?? 'obj'
    const rawIdx = finiteNumber(ph['@_idx'])
    const idx = rawIdx === undefined ? undefined : Math.trunc(rawIdx)
    const name =
      trimmed(node?.nvSpPr?.cNvPr?.['@_name']) ??
      trimmed(node?.nvPicPr?.cNvPr?.['@_name']) ??
      trimmed(node?.nvGraphicFramePr?.cNvPr?.['@_name'])
    placeholders.push({
      type,
      ...(idx === undefined ? {} : { idx }),
      ...(name === undefined ? {} : { name }),
      x: frame.x,
      y: frame.y,
      w: frame.w,
      h: frame.h,
    })
  }

  /** @param {any} container */
  const visit = (container) => {
    for (const sp of asArray(container?.sp)) add(sp, sp?.nvSpPr?.nvPr?.ph, sp?.spPr?.xfrm)
    for (const pic of asArray(container?.pic)) add(pic, pic?.nvPicPr?.nvPr?.ph, pic?.spPr?.xfrm)
    for (const frame of asArray(container?.graphicFrame)) {
      add(frame, frame?.nvGraphicFramePr?.nvPr?.ph, frame?.xfrm)
    }
    for (const group of asArray(container?.grpSp)) visit(group)
  }

  visit(root?.cSld?.spTree)
  return placeholders
}

/**
 * Frame of the first picture a master places (depth-first: direct `<p:pic>`s, then
 * nested groups), clamped to the slide. This is where the master draws the brand
 * lockup — layout-level images are page artwork, not the logo placement.
 *
 * @param {import('./pptx/read.js').PackageHandle} pkg
 * @param {string} part
 * @param {{widthIn: number, heightIn: number}} slide
 * @returns {{x: number, y: number, w: number, h: number} | undefined}
 */
function readMasterLogoFrame(pkg, part, slide) {
  if (!pkg.has(part)) return undefined
  const doc = pkg.xml(part)
  const tree = doc?.sldMaster?.cSld?.spTree
  /** @type {{x: number, y: number, w: number, h: number} | null} */
  let found = null
  /** @param {any} container */
  const visit = (container) => {
    if (found !== null) return
    for (const pic of asArray(container?.pic)) {
      if (found !== null) return
      const rect = readRectInches(pic?.spPr?.xfrm)
      if (rect === null) continue
      found = clampToSlide(rect, slide)
    }
    for (const group of asArray(container?.grpSp)) visit(group)
  }
  visit(tree)
  return found ?? undefined
}

/* ------------------------------------------------------------------ *
 * Layouts inventory
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} RectInches
 * @property {number} x
 * @property {number} y
 * @property {number} w
 * @property {number} h
 */

/**
 * @typedef {object} LayoutPlaceholder
 * @property {string} type        `<p:ph @type>`, `obj` when absent
 * @property {number} [idx]       `<p:ph @idx>`
 * @property {string} [name]      shape name (`cNvPr @name`)
 * @property {number} x
 * @property {number} y
 * @property {number} w
 * @property {number} h
 */

/**
 * @typedef {object} Layout
 * @property {string} part
 * @property {string} name
 * @property {number} placeholderCount  number of placeholders with usable geometry
 * @property {LayoutPlaceholder[]} placeholders
 * @property {string} [background]      `#RRGGBB` when the layout paints its own page
 */

/**
 * @typedef {object} Master
 * @property {string} part
 * @property {string} name
 * @property {string} [background]      `#RRGGBB` painted by `<p:bg>`
 * @property {RectInches} [logoFrame]   frame of the master's first picture, inches
 * @property {Layout[]} layouts
 */

/**
 * @typedef {object} LayoutsInventory
 * @property {{cx: number, cy: number, widthIn: number, heightIn: number}} slideSize
 * @property {Master[]} masters
 * @property {{masters: number, layouts: number}} counts
 */

/**
 * Build the reusable layouts inventory (`layouts.json`).
 *
 * @param {import('./pptx/read.js').PackageHandle} pkg
 * @param {import('./pptx/theme.js').PackageStructure | undefined} structure
 *   result of `discoverStructure(pkg)`; re-discovered when omitted
 * @returns {LayoutsInventory}
 */
export function buildLayoutsInventory(pkg, structure) {
  if (pkg === undefined || pkg === null || typeof pkg.xml !== 'function') {
    throw new TypeError('buildLayoutsInventory(pkg, structure) requires an opened package handle')
  }
  const resolved = structure ?? discoverStructure(pkg)
  const slideSize = readSlideSize(pkg)
  const slide = { widthIn: slideSize.widthIn, heightIn: slideSize.heightIn }

  /** @type {Master[]} */
  const masters = []
  for (const entry of asArray(resolved?.masters)) {
    const part = trimmed(entry?.part)
    if (part === undefined) continue

    /** @type {Layout[]} */
    const layouts = []
    for (const layoutPart of asArray(entry?.layouts)) {
      const path = trimmed(layoutPart)
      if (path === undefined || !pkg.has(path)) continue
      const placeholders = readPlaceholders(pkg, path, slide)
      /** @type {Layout} */
      const layout = {
        part: path,
        name: readLayoutName(pkg, path),
        placeholderCount: placeholders.length,
        placeholders,
      }
      const background = readPaintedBackground(pkg, path)
      if (background !== undefined) layout.background = background
      layouts.push(layout)
    }

    /** @type {Master} */
    const master = { part, name: readMasterName(pkg, part), layouts }
    const background = readPaintedBackground(pkg, part)
    if (background !== undefined) master.background = background
    const logoFrame = readMasterLogoFrame(pkg, part, slide)
    if (logoFrame !== undefined) master.logoFrame = logoFrame
    masters.push(master)
  }

  const layoutCount = masters.reduce((total, master) => total + master.layouts.length, 0)
  return { slideSize, masters, counts: { masters: masters.length, layouts: layoutCount } }
}

/* ------------------------------------------------------------------ *
 * Media / logo classification (markdown input is JSON-safe)
 * ------------------------------------------------------------------ */

/**
 * Relationship source kind of a `.rels` part path.
 * @param {unknown} relsPart
 * @returns {string}
 */
function sourceKindOfRels(relsPart) {
  const path = String(relsPart ?? '')
  if (path.startsWith('ppt/slideMasters/')) return 'master'
  if (path.startsWith('ppt/slideLayouts/')) return 'layout'
  if (path.startsWith('ppt/slides/')) return 'slide'
  if (path.startsWith('ppt/notesMasters/')) return 'notesMaster'
  if (path.startsWith('ppt/notesSlides/')) return 'notesSlide'
  if (path.startsWith('ppt/handoutMasters/')) return 'handoutMaster'
  return 'other'
}

/**
 * Which kinds of part reference a media entry. Accepts both shapes we may be handed:
 * raw `listMedia()` entries (`refs[].sourceKind`) and the JSON-safe inventory from
 * `parsePptxTemplate` (`referencedBy` rels paths).
 * @param {any} entry
 * @returns {Set<string>}
 */
function mediaSourceKinds(entry) {
  const kinds = new Set()
  for (const ref of Array.isArray(entry?.refs) ? entry.refs : []) {
    const kind = trimmed(ref?.sourceKind) ?? (ref?.relsPart === undefined ? undefined : sourceKindOfRels(ref.relsPart))
    if (kind !== undefined) kinds.add(kind)
  }
  for (const relsPart of Array.isArray(entry?.referencedBy) ? entry.referencedBy : []) {
    if (typeof relsPart === 'string' && relsPart.length > 0) kinds.add(sourceKindOfRels(relsPart))
  }
  return kinds
}

/**
 * The source part paths that reference a media entry (`…/_rels/x.xml.rels` → `…/x.xml`).
 * @param {any} entry
 * @returns {string[]}
 */
function mediaSourceParts(entry) {
  /** @type {Set<string>} */
  const parts = new Set()
  for (const ref of Array.isArray(entry?.refs) ? entry.refs : []) {
    const source = trimmed(ref?.sourcePart)
    if (source !== undefined) parts.add(source)
  }
  for (const relsPart of Array.isArray(entry?.referencedBy) ? entry.referencedBy : []) {
    const match = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(String(relsPart ?? ''))
    if (match !== null) parts.add(`${match[1]}/${match[2]}`)
  }
  return [...parts].sort((a, b) => a.localeCompare(b, 'en'))
}

/**
 * Purpose label for the media table:
 * master-only → 徽标; referenced by a layout → 背景/配图; otherwise → 图片.
 * @param {any} entry
 * @returns {string}
 */
function mediaPurpose(entry) {
  const kinds = mediaSourceKinds(entry)
  if (kinds.size === 0) return '图片'
  if (kinds.size === 1 && kinds.has('master')) return '徽标'
  if (kinds.has('layout')) return '背景/配图'
  return '图片'
}

/**
 * `{x,y,w,h}` → `{x,y,w,h}` when all four are finite, else `null`.
 * @param {any} value
 * @returns {RectInches | null}
 */
function frameOf(value) {
  if (value === null || typeof value !== 'object') return null
  const x = finiteNumber(value.x)
  const y = finiteNumber(value.y)
  const w = finiteNumber(value.w)
  const h = finiteNumber(value.h)
  if (x === undefined || y === undefined || w === undefined || h === undefined) return null
  return { x, y, w, h }
}

/**
 * @typedef {object} LogoFacts
 * @property {string | undefined} partPath
 * @property {number | undefined} widthPx
 * @property {number | undefined} heightPx
 * @property {number | undefined} bytes
 * @property {RectInches | null} frame
 * @property {string[]} masters
 */

/**
 * Resolve the logo facts the markdown needs.
 *
 * Priority: `meta.logo` (produced by `parsePptxTemplate`) → the media inventory
 * (a part referenced **only** by masters, ranked by master count, then pixel width)
 * → nothing. The frame comes from `meta.logo.frameInches`, else from the logo-placing
 * master in the layouts inventory.
 *
 * @param {{meta: any, media: any[], layouts: LayoutsInventory | null}} input
 * @returns {LogoFacts}
 */
function resolveLogo({ meta, media, layouts }) {
  const declared = meta?.logo !== null && typeof meta?.logo === 'object' ? meta.logo : null
  const metaMedia = Array.isArray(meta?.media) ? meta.media : []
  const pool = media.length > 0 ? media : metaMedia

  /** @type {any} */
  let chosen = null
  const declaredPath = trimmed(declared?.partPath)
  if (declaredPath !== undefined) {
    chosen = pool.find((entry) => trimmed(entry?.path) === declaredPath) ?? null
  }
  if (chosen === null) {
    const candidates = pool.filter((entry) => {
      const kinds = mediaSourceKinds(entry)
      if (kinds.size === 0) return false
      if (![...kinds].every((kind) => kind === 'master')) return false
      return finiteNumber(entry?.widthPx) !== undefined && finiteNumber(entry?.heightPx) !== undefined
    })
    const ranked = [...candidates].sort((a, b) => {
      const mastersA = mediaSourceParts(a).length
      const mastersB = mediaSourceParts(b).length
      if (mastersA !== mastersB) return mastersB - mastersA
      const widthA = finiteNumber(a?.widthPx) ?? 0
      const widthB = finiteNumber(b?.widthPx) ?? 0
      if (widthA !== widthB) return widthB - widthA
      return String(a?.path).localeCompare(String(b?.path), 'en')
    })
    chosen = ranked[0] ?? null
  }

  const partPath = declaredPath ?? trimmed(chosen?.path)
  const mastersFromPool = chosen === null ? [] : mediaSourceParts(chosen).filter((part) => part.startsWith('ppt/slideMasters/'))
  const declaredMasters = Array.isArray(declared?.masterParts)
    ? declared.masterParts.map((part) => trimmed(part)).filter((part) => part !== undefined)
    : []
  const masters = declaredMasters.length > 0 ? declaredMasters : mastersFromPool

  let frame = frameOf(declared?.frameInches)
  if (frame === null) {
    for (const part of masters) {
      const master = asArray(layouts?.masters).find((entry) => entry?.part === part)
      const candidate = frameOf(master?.logoFrame)
      if (candidate !== null) {
        frame = candidate
        break
      }
    }
  }

  return {
    partPath,
    widthPx: finiteNumber(declared?.widthPx) ?? finiteNumber(chosen?.widthPx),
    heightPx: finiteNumber(declared?.heightPx) ?? finiteNumber(chosen?.heightPx),
    bytes: finiteNumber(declared?.bytes) ?? finiteNumber(chosen?.bytes),
    frame,
    masters,
  }
}

/* ------------------------------------------------------------------ *
 * Markdown rendering
 * ------------------------------------------------------------------ */

/**
 * Sanitize one table cell: no newlines, no raw pipes, no surrounding blanks.
 * @param {unknown} value
 * @returns {string}
 */
function cell(value) {
  const text = value === undefined || value === null ? DASH : String(value)
  return text.replace(/\r?\n/g, ' ').replace(/\|/g, '｜').trim()
}

/**
 * Render a markdown table (header + separator + rows), one string per line.
 * @param {string[]} headers
 * @param {Array<Array<unknown>>} rows
 * @returns {string[]}
 */
function table(headers, rows) {
  const lines = [
    `| ${headers.map(cell).join(' | ')} |`,
    `| ${headers.map(() => '---').join(' | ')} |`,
  ]
  for (const row of rows) {
    lines.push(`| ${row.map(cell).join(' | ')} |`)
  }
  return lines
}

/**
 * Colour text for the palette table. Missing values become `—` (never `undefined`).
 * @param {unknown} value
 * @returns {string}
 */
function colorCell(value) {
  const text = trimmed(value)
  return text === undefined ? DASH : text.toUpperCase()
}

/**
 * Human aspect label derived from the real page size (`16:9`, `4:3`, `1.50:1`).
 * @param {number} widthIn
 * @param {number} heightIn
 * @returns {string}
 */
function aspectLabel(widthIn, heightIn) {
  if (!(widthIn > 0) || !(heightIn > 0)) return DASH
  const ratio = widthIn / heightIn
  if (Math.abs(ratio - 16 / 9) <= 0.01) return '16:9'
  if (Math.abs(ratio - 4 / 3) <= 0.01) return '4:3'
  return `${ratio.toFixed(2)}:1`
}

/**
 * Aspect label plus the PowerPoint name of that canvas (`16:9（宽屏）`).
 * @param {string} label
 * @returns {string}
 */
function aspectNote(label) {
  if (label === '16:9') return '16:9（宽屏）'
  if (label === '4:3') return '4:3（标准）'
  return label
}

/**
 * Render `template.md`: a deterministic, Chinese-labelled spec of the template.
 *
 * Sections: title + summary, 色板, 字体, 页面, 版式清单, 品牌素材, 复用方式.
 * Every input is optional; the function never throws for missing media, an empty
 * theme or a template without layouts, and never prints `undefined`/`NaN`.
 *
 * @param {object} input
 * @param {import('./schema.js').ThemeTokens | null} [input.theme]
 * @param {any} [input.meta]  `parsePptxTemplate().meta`
 * @param {any[]} [input.media] JSON-safe media inventory
 * @param {LayoutsInventory | null} [input.layouts]  `buildLayoutsInventory()` result
 * @returns {string}
 */
export function buildTemplateMarkdown({ theme, meta, media, layouts } = {}) {
  const themeSafe = theme !== null && typeof theme === 'object' ? theme : {}
  const metaSafe = meta !== null && typeof meta === 'object' ? meta : {}
  const inventory = layouts !== null && typeof layouts === 'object' ? layouts : null
  const mediaList = (Array.isArray(media) ? media : []).filter(
    (entry) => entry !== null && typeof entry === 'object',
  )
  const colors = themeSafe.colors !== null && typeof themeSafe.colors === 'object' ? themeSafe.colors : {}
  const fonts = themeSafe.fonts !== null && typeof themeSafe.fonts === 'object' ? themeSafe.fonts : {}

  const templateName =
    trimmed(metaSafe.templateName) ?? trimmed(themeSafe.templateName) ?? 'BrandDeck 模板'
  const templateId = trimmed(metaSafe.templateId) ?? trimmed(themeSafe.templateId)

  const slide = inventory?.slideSize ?? metaSafe.slideSize ?? {}
  const widthIn = finiteNumber(slide?.widthIn) ?? round4((finiteNumber(slide?.cx) ?? FALLBACK_SLIDE.cx) / EMU_PER_INCH)
  const heightIn =
    finiteNumber(slide?.heightIn) ?? round4((finiteNumber(slide?.cy) ?? FALLBACK_SLIDE.cy) / EMU_PER_INCH)
  const cx = finiteNumber(slide?.cx)
  const cy = finiteNumber(slide?.cy)

  // The inventory is authoritative; a caller that only has `meta` still gets counts.
  const masterCount =
    finiteNumber(inventory?.counts?.masters) ??
    finiteNumber(metaSafe.masters) ??
    asArray(inventory?.masters).length
  const layoutCount =
    finiteNumber(inventory?.counts?.layouts) ??
    finiteNumber(metaSafe.layouts) ??
    asArray(inventory?.masters).reduce((total, master) => total + asArray(master?.layouts).length, 0)
  const themeCount = finiteNumber(metaSafe.themes)

  const masters = asArray(inventory?.masters)
  const logo = resolveLogo({ meta: metaSafe, media: mediaList, layouts: inventory })

  /** @type {string[]} */
  const lines = []

  /* --- 1. title + summary ---------------------------------------------- */
  lines.push(`# ${templateName}`, '')
  lines.push(
    `> 页面 ${fmt(widthIn)} × ${fmt(heightIn)} 英寸（宽高比 ${aspectLabel(widthIn, heightIn)}）· ` +
      `母版 ${masterCount} 个 · 版式 ${layoutCount} 个 · 主题 ${themeCount ?? DASH} 套。`,
    '',
  )

  /* --- 2. 色板 ---------------------------------------------------------- */
  /** @type {Array<[string, unknown, string]>} */
  const paletteRows = [
    ['背景', colors.background, '页面底色（母版背景与版式继承）'],
    ['容器', colors.surface, '卡片 / 内容块填充'],
    ['主文字', colors.textPrimary, '标题与正文主色'],
    ['次文字', colors.textSecondary, '注释、页脚与辅助信息'],
    ['主色', colors.primary, '品牌主色：强调、关键数据'],
    ['强调色', colors.accent, '次级强调：高亮、CTA'],
  ]
  const chartPalette = Array.isArray(colors.chartPalette) ? colors.chartPalette : []
  chartPalette.forEach((value, index) => {
    paletteRows.push([`图表${index + 1}`, value, `图表第 ${index + 1} 个系列颜色`])
  })
  lines.push('## 色板', '')
  lines.push(...table(['语义', 'HEX', '用途'], paletteRows.map(([label, hex, usage]) => [label, colorCell(hex), usage])), '')

  /* --- 3. 字体 ---------------------------------------------------------- */
  lines.push('## 字体', '')
  lines.push(
    ...table(
      ['角色', '字体名'],
      [
        ['标题', trimmed(fonts.title) ?? DASH],
        ['正文', trimmed(fonts.body) ?? DASH],
        ['西文', trimmed(fonts.latin) ?? DASH],
      ],
    ),
    '',
  )

  /* --- 4. 页面 ---------------------------------------------------------- */
  lines.push('## 页面', '')
  lines.push(
    `- 尺寸：${fmt(widthIn)} × ${fmt(heightIn)} 英寸` +
      (cx !== undefined && cy !== undefined ? `（EMU ${Math.round(cx)} × ${Math.round(cy)}）` : ''),
  )
  lines.push(`- 宽高比：${aspectNote(aspectLabel(widthIn, heightIn))}`)
  lines.push(
    '- 16:9 说明：编译器按 1280 × 720 px 画布映射（px/96 = 13.333 × 7.5 英寸），' +
      '与 PowerPoint 的 LAYOUT_WIDE 对齐；模板实际页宽与标准值的小差异由画布缩放吸收。',
  )
  lines.push('')

  /* --- 5. 版式清单 ------------------------------------------------------ */
  lines.push('## 版式清单', '')
  if (masters.length === 0) {
    lines.push(`- ${DASH} 模板中没有可用的母版 / 版式。`, '')
  } else {
    for (const master of masters) {
      const part = trimmed(master?.part) ?? DASH
      const name = trimmed(master?.name) ?? baseName(part)
      const masterBackground = trimmed(master?.background)
      lines.push(`### ${name}（${part}）`, '')

      const layouts = asArray(master?.layouts)
      const rows = layouts.map((layout) => {
        const layoutPart = trimmed(layout?.part) ?? DASH
        const layoutName = trimmed(layout?.name) ?? baseName(layoutPart)
        const ownBackground = trimmed(layout?.background)
        const backgroundCell =
          ownBackground !== undefined
            ? `版式自绘 ${ownBackground.toUpperCase()}`
            : masterBackground !== undefined
              ? `继承母版 ${masterBackground.toUpperCase()}`
              : DASH
        const placeholderCount =
          finiteNumber(layout?.placeholderCount) ?? asArray(layout?.placeholders).length
        return [layoutName, String(placeholderCount), backgroundCell]
      })
      if (rows.length === 0) {
        const backgroundCell = masterBackground === undefined ? DASH : `继承母版 ${masterBackground.toUpperCase()}`
        rows.push([`${DASH} 无版式`, '0', backgroundCell])
      }
      lines.push(...table(['名称', '占位符数', '背景色'], rows), '')

      for (const layout of layouts) {
        const placeholders = asArray(layout?.placeholders)
        if (placeholders.length === 0) continue
        const layoutPart = trimmed(layout?.part) ?? DASH
        const layoutName = trimmed(layout?.name) ?? baseName(layoutPart)
        lines.push(`#### ${layoutName}（${layoutPart}）`, '')
        for (const placeholder of placeholders.slice(0, PLACEHOLDER_LINE_LIMIT)) {
          const type = trimmed(placeholder?.type) ?? 'obj'
          const idx =
            placeholder?.idx === undefined || placeholder?.idx === null
              ? ''
              : ` (idx ${Number.isFinite(Number(placeholder.idx)) ? Math.trunc(Number(placeholder.idx)) : placeholder.idx})`
          lines.push(
            `- ${type}${idx}: ${fmt(placeholder?.x)},${fmt(placeholder?.y)},` +
              `${fmt(placeholder?.w)},${fmt(placeholder?.h)}`,
          )
        }
        if (placeholders.length > PLACEHOLDER_LINE_LIMIT) lines.push('…')
        lines.push('')
      }
    }
  }

  /* --- 6. 品牌素材 ------------------------------------------------------ */
  lines.push('## 品牌素材', '')
  if (logo.partPath !== undefined) {
    const size =
      logo.widthPx !== undefined && logo.heightPx !== undefined
        ? `${Math.round(logo.widthPx)}×${Math.round(logo.heightPx)} px`
        : '像素尺寸未知'
    const bytesText = logo.bytes === undefined ? '' : `，${Math.round(logo.bytes)} 字节`
    lines.push(`- 徽标文件：${logo.partPath}（${size}${bytesText}）`)
  } else {
    lines.push('- 徽标文件：未从母版解析到徽标图片。')
  }
  if (logo.frame !== null) {
    lines.push(
      `- 徽标框（母版坐标，英寸）：x ${fmt(logo.frame.x)} · y ${fmt(logo.frame.y)} · ` +
        `w ${fmt(logo.frame.w)} · h ${fmt(logo.frame.h)}`,
    )
  } else {
    lines.push('- 徽标框：未定位（母版未放置徽标图片）。')
  }
  lines.push(`- 引用母版：${logo.masters.length === 0 ? '无' : logo.masters.join('、')}`, '')

  const mediaRows = [...mediaList]
    .sort((a, b) => String(a?.path ?? '').localeCompare(String(b?.path ?? ''), 'en'))
    .map((entry) => {
      const widthPx = finiteNumber(entry?.widthPx)
      const heightPx = finiteNumber(entry?.heightPx)
      const size =
        widthPx !== undefined && heightPx !== undefined
          ? `${Math.round(widthPx)}×${Math.round(heightPx)} px`
          : DASH
      const bytes = finiteNumber(entry?.bytes)
      return [trimmed(entry?.path) ?? DASH, size, bytes === undefined ? '0' : String(Math.round(bytes)), mediaPurpose(entry)]
    })
  if (mediaRows.length === 0) {
    mediaRows.push([DASH, DASH, '0', '模板中未包含图片素材'])
  }
  lines.push(...table(['文件', '尺寸', '字节', '用途'], mediaRows), '')

  /* --- 7. 复用方式 ------------------------------------------------------ */
  lines.push('## 复用方式', '')
  lines.push(
    '- 模板已物化为 theme.json（令牌）、layouts.json（版式）与 source.pptx（原件副本），' +
      '下次使用无需再解析 .pptx。',
  )
  lines.push(
    '- 本文件 template.md（当前说明）与 theme.json、layouts.json 一并保存在模板目录，' +
      '任何评审者都可以直接阅读。',
  )
  lines.push(
    `- 下一轮直接调用 branddeck_deck_create { templateId }` +
      (templateId === undefined ? '（templateId 见模板目录名）' : `（本模板 templateId = "${templateId}"）`) +
      ' 即可复用，无需重新解析。',
  )

  // Exactly one trailing newline, no trailing blank line, no trailing spaces: the file
  // is diffed and reviewed, so its byte shape has to be stable.
  return `${lines.join('\n').replace(/\n+$/, '')}\n`
}

/* ------------------------------------------------------------------ *
 * Orchestration
 * ------------------------------------------------------------------ */

/**
 * Convert a raw `listMedia()` inventory into JSON-safe entries (no `Buffer`).
 * @param {import('./pptx/read.js').PackageHandle} pkg
 * @returns {any[]}
 */
function toMediaInventory(pkg) {
  return listMedia(pkg).map((entry) => ({
    path: entry.path,
    bytes: entry.bytes,
    contentType: entry.contentType,
    format: entry.format,
    widthPx: entry.widthPx,
    heightPx: entry.heightPx,
    referencedBy: entry.referencedBy,
    refs: entry.refs.map((ref) => ({
      relsPart: ref.relsPart,
      sourcePart: ref.sourcePart,
      sourceKind: ref.sourceKind,
      relId: ref.relId,
    })),
  }))
}

/** @param {import('./pptx/read.js').PackageHandle} pkg */
function countThemeParts(pkg) {
  return pkg.partNames().filter((part) => /^ppt\/theme\/theme\d+\.xml$/.test(part)).length
}

/**
 * Build both persisted artefacts in one call.
 *
 * @param {object} input
 * @param {import('./pptx/read.js').PackageHandle} input.pkg
 * @param {import('./pptx/theme.js').PackageStructure} [input.structure]
 * @param {import('./schema.js').ThemeTokens} [input.theme]
 * @param {any} [input.meta]
 * @param {any[]} [input.media]
 * @returns {{markdown: string, layouts: LayoutsInventory}}
 */
export function buildTemplateDoc({ pkg, structure, theme, meta, media } = {}) {
  if (pkg === undefined || pkg === null || typeof pkg.xml !== 'function') {
    throw new TypeError('buildTemplateDoc({ pkg }) requires an opened package handle')
  }
  const resolved = structure ?? discoverStructure(pkg)
  const layouts = buildLayoutsInventory(pkg, resolved)
  const mediaList = Array.isArray(media) ? media : toMediaInventory(pkg)

  // Back-fill the counts the spec line needs, so a caller may pass a minimal `meta`
  // (only templateName/templateId) and still get an accurate document.
  const enrichedMeta = meta !== null && typeof meta === 'object' ? { ...meta } : {}
  if (trimmed(enrichedMeta.templateId) === undefined) {
    const fromTheme = trimmed(theme?.templateId)
    if (fromTheme !== undefined) enrichedMeta.templateId = fromTheme
  }
  if (trimmed(enrichedMeta.templateName) === undefined) {
    const fromTheme = trimmed(theme?.templateName)
    if (fromTheme !== undefined) enrichedMeta.templateName = fromTheme
  }
  enrichedMeta.masters = layouts.counts.masters
  enrichedMeta.layouts = layouts.counts.layouts
  if (finiteNumber(enrichedMeta.themes) === undefined) enrichedMeta.themes = countThemeParts(pkg)
  enrichedMeta.slideSize = { ...layouts.slideSize }

  const markdown = buildTemplateMarkdown({ theme, meta: enrichedMeta, media: mediaList, layouts })
  return { markdown, layouts }
}
