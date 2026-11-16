/**
 * What a corporate template actually offers *visually*: background artwork, the content
 * box it uses, and the fixed footer strings it repeats on every page.
 *
 * WHY THIS EXISTS: parsing a template used to yield colours, fonts and a logo — and then
 * generation drew on a plain white canvas with our own 64px margins. A template whose
 * layouts carry full-bleed artwork, a `Security Level:` marker and a 11.7in content column
 * produced a deck that looked nothing like it. The three artefacts below are the parts of
 * the template's visual system that survive the trip to pptxgenjs:
 *
 *   · `backdrops` — the layout pictures that cover the whole slide (cover art). They become
 *     the *layout-level* `<p:bg><a:blipFill>` of the generated deck, i.e. the same mechanism
 *     the template itself uses.
 *   · `frame` — the geometry of the template's content placeholder, in canvas px. Drawing
 *     inside it is what makes generated pages line up with the template's own grid.
 *   · `footlines` — short static strings on the master ("Security Level:", "Huawei
 *     Confidential"). Inherited when present, absent otherwise: inventing a confidentiality
 *     marker on a template that has none would be worse than an empty footer.
 *
 * Everything here is read-only and tolerant: a malformed template yields empty results
 * rather than throwing, because a template that parses "less well" must still parse.
 */
import { asArray } from './read.js'
import { EMU_PER_INCH } from './theme.js'

/** A picture covering at least this share of the slide counts as background art. */
const BACKDROP_MIN_FILL = 0.6

/** Canvas the DSL draws on; template inches are mapped onto it. */
const CANVAS = Object.freeze({ width: 1280, height: 720 })

/** Placeholder prompt text that must never be mistaken for a footer. */
const PROMPT_TEXT = /^(?:单击此处|点击此处|click to|add (?:a )?(?:title|text)|type (?:a )?(?:title|text))/i

/** A footer line is short and static; anything longer is body copy. */
const FOOTLINE_MAX_CHARS = 48

/** Compliance markers we *want* to inherit ("Security Level:", "机密", "Confidential"). */
const COMPLIANCE_TEXT = /(?:保密|机密|内部|受限|confidential|security|internal|proprietary|restricted)\s*[:：]?/i

/** Design annotations on a master are not footers ("PANTONE186CRGB 200/16/46"). */
const ANNOTATION_TEXT = /pantone|rgb|cmyk|#?[0-9a-f]{6}\b|字号|font\s*size|\bred\b/i

/**
 * @typedef {object} Backdrop
 * @property {string} layoutPart  e.g. `ppt/slideLayouts/slideLayout1.xml`
 * @property {string} layoutName  the layout's display name when the package has one
 * @property {string} mediaPart   e.g. `ppt/media/image2.png`
 * @property {{x: number, y: number, w: number, h: number}} frameInches  where the template places it
 * @property {number} fillRatio   picture area / slide area (1.0 = full bleed)
 * @property {boolean} fullBleed  covers the slide within the tolerance below
 */

/**
 * @typedef {object} TemplateArtwork
 * @property {Backdrop[]} backdrops
 * @property {{x: number, y: number, w: number, h: number} | null} frame  canvas px
 * @property {{x: number, y: number, w: number, h: number} | null} coverFrame  title zone, canvas px
 * @property {string[]} footlines
 * @property {{widthEmu: number, heightEmu: number} | null} slideSize
 */

/**
 * Read the template's visual system.
 *
 * @param {import('./read.js').PackageHandle} pkg
 * @param {import('./theme.js').PackageStructure} structure
 * @returns {TemplateArtwork}
 */
export function readTemplateArtwork(pkg, structure) {
  const slideSize = readSlideSize(pkg)
  const scale = canvasScale(slideSize)
  /** @type {Backdrop[]} */
  const backdrops = []
  /** @type {Array<{part: string, name: string, body: {x: number, y: number, w: number, h: number} | null, title: {x: number, y: number, w: number, h: number} | null, area: number}>} */
  const candidates = []
  /** @type {string[]} */
  const footlines = []

  for (const layoutPart of structure.layouts) {
    const doc = pkg.xml(layoutPart)
    const layout = doc?.sldLayout ?? doc
    const spTree = layout?.cSld?.spTree
    if (spTree === undefined || spTree === null) continue
    const relById = new Map(pkg.relationships(layoutPart).map((rel) => [rel.id, rel]))
    const name = layoutNameOf(pkg, layoutPart) ?? basenameOf(layoutPart)

    // --- pictures that cover the slide ------------------------------------
    /** @type {{x: number, y: number, w: number, h: number} | null} */
    let body = null
    /** @type {{x: number, y: number, w: number, h: number} | null} */
    let title = null
    /** @type {number} */
    let bodyArea = 0

    for (const shape of walkShapes(spTree)) {
      if (shape.kind === 'pic') {
        const embed = shape.node?.blipFill?.blip?.['@_embed']
        const rel = embed === undefined ? undefined : relById.get(String(embed))
        const mediaPart = rel?.resolvedPart ?? null
        const frame = readFrameEmu(shape.node?.spPr?.xfrm)
        if (mediaPart === null) continue
        if (frame === null) continue
        const fill = fillRatio(frame, slideSize)
        if (fill < BACKDROP_MIN_FILL) continue
        backdrops.push({
          layoutPart,
          layoutName: name,
          mediaPart,
          // The template's art is 13.34×6.13in on a 13.34×7.5in slide: full width, 82% of
          // the height, with the footer band left white on purpose. Stretching it to
          // full-bleed would squash the picture and bury the footer, so the frame travels
          // with the backdrop and compile places it exactly.
          frameInches: {
            x: round4(frame.x / EMU_PER_INCH),
            y: round4(frame.y / EMU_PER_INCH),
            w: round4(frame.w / EMU_PER_INCH),
            h: round4(frame.h / EMU_PER_INCH),
          },
          fillRatio: Number(fill.toFixed(3)),
          fullBleed: fill >= 0.92,
        })
        continue
      }
      if (shape.kind !== 'placeholder') continue
      const type = String(shape.node?.nvSpPr?.nvPr?.ph?.['@_type'] ?? 'body').toLowerCase()
      const frame = readFrameEmu(shape.node?.spPr?.xfrm)
      if (frame === null) continue
      const rect = toCanvas(frame, scale)
      const area = rect.w * rect.h
      // A "title" is whatever the template uses for the heading slot — `subTitle` in
      // PowerPoint is often the eyebrow line above a heading, and a real corporate content
      // page may well name its heading `subTitle` (the Huawei content page does). Treat any
      // heading-ish type as a title candidate and keep the topmost one.
      if (type === 'ctrtitle' || type === 'title' || type === 'subtitle') {
        if (title === null || rect.y < title.y) title = rect
        continue
      }
      if (area > bodyArea) {
        bodyArea = area
        body = rect
      }
    }

    if (body !== null) {
      // The page box is the union of the title and body zones: `innerBox` is the area the
      // *whole page* is laid out in (our title sits at its top edge), not just the body.
      const top = title === null ? body.y : Math.min(title.y, body.y)
      const bottom = Math.max(body.y + body.h, title === null ? 0 : title.y + title.h)
      const left = title === null ? body.x : Math.min(title.x, body.x)
      const right = Math.max(body.x + body.w, title === null ? 0 : title.x + title.w)
      candidates.push({
        part: layoutPart,
        name,
        body: { x: left, y: top, w: right - left, h: bottom - top },
        title,
        area: bodyArea,
      })
    }
  }

  // Footers come from the MASTER, never from layouts: layout text is page furniture
  // ("目录", "Thank you."), and treating it as a footer would stamp "目录" on every slide.
  // Master placeholders do count — "Security Level:" is a placeholder on the Huawei master.
  for (const master of structure.masters) {
    const doc = pkg.xml(master.part)
    const spTree = doc?.sldMaster?.cSld?.spTree
    if (spTree === null || spTree === undefined) continue
    for (const shape of walkShapes(spTree)) {
      if (shape.kind === 'pic') continue
      if (shape.kind === 'placeholder') {
        const type = String(shape.node?.nvSpPr?.nvPr?.ph?.['@_type'] ?? '').toLowerCase()
        // A title placeholder on a master is the deck title slot, not a footer.
        if (type === 'title' || type === 'ctrtitle' || type === 'subtitle') continue
      }
      pushFootline(footlines, textOf(shape.node))
    }
  }

  // The content frame is the *biggest* body placeholder: a cover layout's body is a
  // single-line subtitle, a content layout's is the page. Sorting by area picks the page
  // without needing to understand layout names in any language.
  const content = [...candidates].sort((a, b) => b.area - a.area)[0] ?? null
  const coverSource = [...candidates].sort((a, b) => (b.title === null ? 0 : 1) - (a.title === null ? 0 : 1))[0] ?? null

  // Prefer compliance markers; fall back to nothing at all. A template that only carries
  // design annotations gets no footer rather than a stamp of its own colour spec.
  const compliance = footlines.filter((line) => COMPLIANCE_TEXT.test(line) && !ANNOTATION_TEXT.test(line))
  return {
    backdrops,
    frame: content?.body ?? null,
    coverFrame: coverSource?.title ?? null,
    footlines: (compliance.length > 0 ? compliance : []).slice(0, 2),
    slideSize,
  }
}

/* ------------------------------------------------------------------ *
 * Shape walking
 * ------------------------------------------------------------------ */

/**
 * Flatten a `<p:spTree>` into pictures and placeholders, following nested groups.
 * @param {any} spTree
 * @param {number} [depth]
 * @returns {Array<{kind: 'pic' | 'placeholder' | 'text', node: any}>}
 */
function walkShapes(spTree, depth = 0) {
  if (spTree === null || spTree === undefined || depth > 6) return []
  /** @type {Array<{kind: 'pic' | 'placeholder' | 'text', node: any}>} */
  const out = []
  for (const pic of asArray(spTree.pic)) out.push({ kind: 'pic', node: pic })
  for (const sp of asArray(spTree.sp)) {
    const isPlaceholder = sp?.nvSpPr?.nvPr?.ph !== undefined
    out.push({ kind: isPlaceholder ? 'placeholder' : 'text', node: sp })
  }
  for (const group of asArray(spTree.grpSp)) out.push(...walkShapes(group, depth + 1))
  return out
}

/** @param {any} sp */
function textOf(sp) {
  const body = sp?.txBody
  if (body === null || body === undefined) return ''
  const paragraphs = asArray(body.p)
  const parts = []
  for (const paragraph of paragraphs) {
    for (const run of asArray(paragraph?.r)) {
      const text = run?.t?.['#text'] ?? run?.t
      if (typeof text === 'string') parts.push(text)
    }
  }
  return parts.join('').replace(/\s+/g, ' ').trim()
}

/** @param {string[]} list @param {string} text */
function pushFootline(list, text) {
  const value = String(text ?? '').trim()
  if (value === '') return
  if (Array.from(value).length > FOOTLINE_MAX_CHARS) return
  if (PROMPT_TEXT.test(value)) return
  if (list.includes(value)) return
  list.push(value)
}

/* ------------------------------------------------------------------ *
 * Geometry
 * ------------------------------------------------------------------ */

/**
 * `ppt/presentation.xml` `<p:sldSz cx cy>` in EMU.
 * @param {import('./read.js').PackageHandle} pkg
 */
function readSlideSize(pkg) {
  const doc = pkg.xml('ppt/presentation.xml')
  const size = doc?.presentation?.sldSz ?? doc?.sldSz
  const widthEmu = Number(size?.['@_cx'])
  const heightEmu = Number(size?.['@_cy'])
  if (!Number.isFinite(widthEmu) || !Number.isFinite(heightEmu)) return null
  if (widthEmu <= 0 || heightEmu <= 0) return null
  return { widthEmu, heightEmu }
}

/**
 * Template inches → canvas px. The template is usually 13.34in wide against our 13.333in,
 * so x and y are scaled separately instead of assuming one ratio.
 * @param {{widthEmu: number, heightEmu: number} | null} slideSize
 */
function canvasScale(slideSize) {
  const widthIn = slideSize === null ? 13.3333 : slideSize.widthEmu / EMU_PER_INCH
  const heightIn = slideSize === null ? 7.5 : slideSize.heightEmu / EMU_PER_INCH
  return { x: CANVAS.width / widthIn, y: CANVAS.height / heightIn }
}

/**
 * `<a:xfrm><a:off/><a:ext/></a:xfrm>` → EMU rect.
 * @param {any} xfrm
 * @returns {{x: number, y: number, w: number, h: number} | null}
 */
function readFrameEmu(xfrm) {
  const off = xfrm?.off
  const ext = xfrm?.ext
  if (off === undefined || ext === undefined) return null
  const x = Number(off['@_x'])
  const y = Number(off['@_y'])
  const w = Number(ext['@_cx'])
  const h = Number(ext['@_cy'])
  if (![x, y, w, h].every(Number.isFinite)) return null
  return { x, y, w, h }
}

/**
 * EMU rect → canvas px rect.
 * @param {{x: number, y: number, w: number, h: number}} emu
 * @param {{x: number, y: number}} scale
 */
function toCanvas(emu, scale) {
  return {
    x: Math.round((emu.x / EMU_PER_INCH) * scale.x),
    y: Math.round((emu.y / EMU_PER_INCH) * scale.y),
    w: Math.round((emu.w / EMU_PER_INCH) * scale.x),
    h: Math.round((emu.h / EMU_PER_INCH) * scale.y),
  }
}

/**
 * Share of the slide a picture covers.
 * @param {{x: number, y: number, w: number, h: number}} frame  EMU
 * @param {{widthEmu: number, heightEmu: number} | null} slideSize
 */
function fillRatio(frame, slideSize) {
  const width = slideSize?.widthEmu ?? 12192000
  const height = slideSize?.heightEmu ?? 6858000
  const slideArea = width * height
  if (slideArea <= 0) return 0
  return (frame.w * frame.h) / slideArea
}

/**
 * The layout's `name` attribute (`sldLayout.name`), if present.
 * @param {import('./read.js').PackageHandle} pkg
 * @param {string} layoutPart
 * @returns {string | null}
 */
function layoutNameOf(pkg, layoutPart) {
  const doc = pkg.xml(layoutPart)
  const layout = doc?.sldLayout ?? doc
  const name = layout?.['@_name'] ?? layout?.cSld?.['@_name']
  return typeof name === 'string' && name.trim() !== '' ? name.trim() : null
}

/** @param {number} value */
function round4(value) {
  return Math.round(value * 10000) / 10000
}

/** @param {string} part */
function basenameOf(part) {
  const file = part.split('/').pop() ?? part
  return file.replace(/\.xml$/i, '')
}
