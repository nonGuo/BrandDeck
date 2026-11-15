/**
 * CardDeck → native `.pptx` compiler (design doc §三.4, ARCHITECTURE §2).
 *
 * Mapping rules:
 * - `pres.layout = 'LAYOUT_WIDE'` (13.333 × 7.5 in) and the 1280×720 canvas maps exactly
 *   with `inch = px / 96`; font sizes are `pt = px * 0.75`.
 * - One slide master carries the deck background, the logo and the footer/page number.
 *   The logo honours `theme.brandAssets.logoFrame` when present — on the real HW
 *   template the logo lives **bottom-right**, so a hard-coded top-right slot would be
 *   wrong. Without a frame we fall back to the top-right `LOGO_RESERVE` slot and fit the
 *   logo inside it by its real aspect ratio.
 * - text → `addText` (`shrinkText:true` so cross-platform font differences never spill),
 *   shape → `addShape`, chart → `addChart` (native editable data), svg → 3× PNG
 *   (`addImage`), image → `addImage`.
 * - Every failure that can be survived is a warning (`assets` missing, un-rasterisable
 *   SVG, off-budget text): the export never throws away the whole deck because one
 *   element was bad.
 * - The produced buffer is re-opened with jszip before it is returned
 *   (`{ verified: true }`), so a corrupt OPC package fails here and not in PowerPoint.
 */
import { existsSync, mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'
import { randomBytes } from 'node:crypto'
import { dirname, join, resolve } from 'node:path'
import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'
import {
  CANVAS,
  LOGO_RESERVE,
  SAFE_AREA,
  estimateTextOverflow,
  normalizeHex,
  resolveColor,
  splitAlpha,
} from './schema.js'
import { DEFAULT_THEME } from './defaults.js'
import { imageSize } from './pptx/media.js'
import { rasterizeSvg } from './svg.js'

/** Master name registered on every compiled presentation. */
export const MASTER_NAME = 'BRANDDECK'

const PX_PER_INCH = 96
const PT_PER_PX = 0.75

/**
 * px → inches (4 decimal places, rounded for stable output).
 * @param {number} value
 * @returns {number}
 */
export function pxToIn(value) {
  return round(Math.max(0, Number(value) || 0) / PX_PER_INCH, 4)
}

/**
 * px → points (2 decimals; the DSL expresses font sizes in canvas px).
 * @param {number} value
 * @returns {number}
 */
export function pxToPt(value) {
  return round(Math.max(0, Number(value) || 0) * PT_PER_PX, 2)
}

/** Shape names accepted by the DSL → pptxgenjs `ShapeType` keys. */
const SHAPE_TYPES = Object.freeze({
  rect: 'rect',
  roundRect: 'roundRect',
  ellipse: 'ellipse',
  line: 'line',
  rightArrow: 'rightArrow',
})

/** Chart names accepted by the DSL → pptxgenjs `ChartType` keys. */
const CHART_TYPES = Object.freeze({ bar: 'bar', line: 'line', pie: 'pie' })

/**
 * @typedef {object} CompileOptions
 * @property {number} [svgScale]          rasterization scale, default 3 (288 dpi)
 * @property {string[]} [fontFiles]       explicit font files for resvg
 * @property {string[]} [fontDirs]        explicit font directories for resvg
 * @property {boolean} [loadSystemFonts]  default true
 * @property {boolean} [includeLogo]      default true
 * @property {boolean} [includeFooter]    default true
 * @property {string} [footerText]        default = deck title
 * @property {boolean} [slideNumbers]     default true
 * @property {boolean} [verify]           default true (re-open the buffer with jszip)
 * @property {string} [masterName]
 * @property {string} [cjkFont]           East-Asian typeface for `<a:ea>`; defaults to the
 *                                        theme's CJK face (see `resolveCjkFont`)
 */

/**
 * @typedef {object} CompileResult
 * @property {Buffer} buffer
 * @property {number} slideCount
 * @property {string[]} warnings
 * @property {boolean} verified
 * @property {string} cjkFont   the East-Asian face written into `<a:ea>`
 */

/**
 * Compile a CardDeck into a `.pptx` Buffer.
 *
 * @param {object} input
 * @param {import('./schema.js').CardDeck} input.deck
 * @param {import('./schema.js').ThemeTokens} [input.theme]
 * @param {Buffer | Uint8Array} [input.logoBytes]
 * @param {Record<string, Buffer | Uint8Array | string>} [input.assets] asset key → bytes or absolute path
 * @param {CompileOptions} [input.options]
 * @returns {Promise<CompileResult>}
 */
export async function compileDeck({ deck, theme, logoBytes, assets, options } = /** @type {any} */ ({})) {
  const opts = /** @type {CompileOptions} */ (options ?? {})
  /** @type {string[]} */
  const warnings = []

  if (deck === undefined || deck === null || !Array.isArray(deck.slides) || deck.slides.length === 0) {
    throw compileError('invalid_deck', 'compileDeck requires a deck with at least one slide')
  }
  const activeTheme = normalizeTheme(theme)
  const deckTitle = typeof deck.title === 'string' && deck.title.trim().length > 0
    ? deck.title.trim()
    : 'BrandDeck'
  const resolvedLogo = resolveLogo(logoBytes, activeTheme, opts.includeLogo !== false, warnings)
  const cjkFont = resolveCjkFont(activeTheme, opts)

  const pres = new PptxGenJS()
  pres.layout = 'LAYOUT_WIDE'
  pres.title = deckTitle
  pres.author = 'BrandDeck'
  pres.company = 'BrandDeck'

  defineMaster(pres, activeTheme, resolvedLogo, deckTitle, opts)

  let svgImages = 0
  deck.slides.forEach((slideDef, slideIndex) => {
    const label = `slide ${slideIndex + 1}`
    const slide = pres.addSlide({ masterName: opts.masterName ?? MASTER_NAME })
    const background = slideDef.background === undefined
      ? activeTheme.colors.background
      : resolveColor(slideDef.background, activeTheme, activeTheme.colors.background)
    slide.background = { color: bareHex(background) }

    reportOverflow(slideDef, activeTheme, label, warnings)

    for (const element of slideDef.elements ?? []) {
      try {
        switch (element.type) {
          case 'text':
            addTextElement(slide, element, activeTheme)
            break
          case 'shape':
            addShapeElement(slide, element, activeTheme, pres, label, warnings)
            break
          case 'svg': {
            if (addSvgElement(slide, element, activeTheme, opts, label, warnings)) svgImages += 1
            break
          }
          case 'chart':
            addChartElement(slide, element, activeTheme, pres, label, warnings)
            break
          case 'image':
            addImageElement(slide, element, activeTheme, resolvedLogo, assets ?? {}, label, warnings)
            break
          default:
            warnings.push(`${label} · unknown element type "${element.type}" skipped`)
        }
      } catch (cause) {
        warnings.push(`${label} · element "${element.id}" skipped: ${messageOf(cause)}`)
      }
    }

    const notes = typeof slideDef.speakerNotes === 'string' ? slideDef.speakerNotes.trim() : ''
    if (notes.length > 0) slide.addNotes(notes)
  })

  const rawBuffer = await pres.write({ outputType: 'nodebuffer' })
  // pptxgenjs has no per-face font option: it writes our `fontFace` into <a:latin>,
  // <a:ea> *and* <a:cs>. A Latin face in <a:ea> makes renderers fall back for CJK text
  // and (LibreOffice) draw the glyphs twice. Fix the faces and the slide-number field
  // artefacts in one post-processing pass.
  const buffer = await finalizePackage(Buffer.isBuffer(rawBuffer) ? rawBuffer : Buffer.from(rawBuffer), {
    cjkFont,
    warnings,
  })
  const slideCount = deck.slides.length
  const verified = opts.verify === false ? true : await verifyPackage(buffer, slideCount, warnings)
  if (svgImages > 0) {
    // Purely informational; the panel shows it as an export note.
    warnings.push(`${svgImages} SVG diagram(s) were rasterized to PNG at ${opts.svgScale ?? 3}x`)
  }

  return {
    buffer,
    slideCount,
    warnings,
    verified,
    cjkFont,
  }
}

/**
 * Compile and write to disk atomically (temp file + rename).
 *
 * Without an explicit `outputPath` the file lands in
 * `<cwd>/branddeck/<deckId>.pptx` (ARCHITECTURE §5 default); the tool layer normally
 * passes the workspace path explicitly.
 *
 * @param {Parameters<typeof compileDeck>[0]} input
 * @param {string} [outputPath]
 * @returns {Promise<CompileResult & {path: string, bytes: number}>}
 */
export async function compileDeckToFile(input, outputPath) {
  const result = await compileDeck(input)
  const deckId = String(input?.deck?.deckId ?? 'deck')
  const target = outputPath === undefined || outputPath === null || String(outputPath).length === 0
    ? join(process.cwd(), 'branddeck', `${deckId}.pptx`)
    : resolve(String(outputPath))
  writeFileAtomic(target, result.buffer)
  return {
    ...result,
    buffer: result.buffer,
    path: target,
    bytes: result.buffer.length,
  }
}

/* ------------------------------------------------------------------ *
 * Master / logo / footer
 * ------------------------------------------------------------------ */

/**
 * @param {PptxGenJS} pres
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {{buffer: Buffer, aspect: number} | null} logo
 * @param {string} deckTitle
 * @param {CompileOptions} opts
 */
function defineMaster(pres, theme, logo, deckTitle, opts) {
  /** @type {any[]} */
  const objects = []

  if (logo !== null) {
    const frame = logoFrame(theme, logo.aspect)
    objects.push({
      image: {
        data: `image/png;base64,${logo.buffer.toString('base64')}`,
        x: frame.x,
        y: frame.y,
        w: frame.w,
        h: frame.h,
        altText: 'brand logo',
      },
    })
  }

  if (opts.includeFooter !== false) {
    const footerText = typeof opts.footerText === 'string' && opts.footerText.trim().length > 0
      ? opts.footerText.trim()
      : deckTitle
    objects.push({
      text: {
        text: footerText,
        options: {
          x: pxToIn(SAFE_AREA.left),
          y: pxToIn(CANVAS.height - SAFE_AREA.bottom + 20),
          w: pxToIn(720),
          h: pxToIn(28),
          fontFace: theme.fonts.body,
          fontSize: 9,
          color: bareHex(theme.colors.textSecondary),
          align: 'left',
          valign: 'middle',
          margin: 0,
          isTextBox: true,
        },
      },
    })
  }

  pres.defineSlideMaster({
    title: opts.masterName ?? MASTER_NAME,
    background: { color: bareHex(theme.colors.background) },
    objects,
    ...(opts.slideNumbers === false
      ? {}
      : {
          slideNumber: {
            x: pxToIn(CANVAS.width - SAFE_AREA.right - 48),
            y: pxToIn(CANVAS.height - SAFE_AREA.bottom + 20),
            w: pxToIn(48),
            h: pxToIn(28),
            color: bareHex(theme.colors.textSecondary),
            fontFace: theme.fonts.body,
            fontSize: 9,
            align: 'right',
            valign: 'middle',
          },
        }),
  })
}

/**
 * Where to draw the logo, in inches.
 *
 * 1. `theme.brandAssets.logoFrame` — the frame measured on the template's primary
 *    master. In `PPT模板-HW.pptx` that is bottom-right (`x 10.07, y 6.54, w 2.47,
 *    h 0.53`), so this is the only way to keep the corporate placement.
 * 2. Otherwise the top-right `LOGO_RESERVE` slot (150×48 px inside the safe area),
 *    with the logo fitted by its real aspect ratio (never stretched).
 *
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {number} aspect width/height of the logo image
 * @returns {{x: number, y: number, w: number, h: number}}
 */
function logoFrame(theme, aspect) {
  const declared = theme.brandAssets?.logoFrame
  const slideWidthIn = CANVAS.width / PX_PER_INCH
  const slideHeightIn = CANVAS.height / PX_PER_INCH
  if (
    declared !== undefined &&
    declared !== null &&
    [declared.x, declared.y, declared.w, declared.h].every(Number.isFinite) &&
    declared.w > 0 &&
    declared.h > 0
  ) {
    // Clamp into the slide (the template's own slide is 13.3385in wide vs our 13.3333in).
    const w = Math.min(declared.w, slideWidthIn)
    const h = Math.min(declared.h, slideHeightIn)
    const x = Math.min(Math.max(0, declared.x), Math.max(0, slideWidthIn - w))
    const y = Math.min(Math.max(0, declared.y), Math.max(0, slideHeightIn - h))
    return { x: round(x, 4), y: round(y, 4), w: round(w, 4), h: round(h, 4) }
  }

  const box = {
    x: pxToIn(CANVAS.width - SAFE_AREA.right - LOGO_RESERVE.width),
    y: pxToIn(SAFE_AREA.top),
    w: pxToIn(LOGO_RESERVE.width),
    h: pxToIn(LOGO_RESERVE.height),
  }
  const safeAspect = Number.isFinite(aspect) && aspect > 0 ? aspect : 4
  const boxAspect = box.w / box.h
  let w = box.w
  let h = box.w / safeAspect
  if (h > box.h) {
    h = box.h
    w = box.h * safeAspect
  }
  return {
    x: round(box.x + (box.w - w), 4), // right-aligned in the reserve slot
    y: round(box.y + (box.h - h) / 2, 4), // vertically centred
    w: round(w, 4),
    h: round(h, 4),
  }
}

/**
 * Resolve the logo bytes: explicit argument → `theme.brandAssets.logoPath` on disk.
 * @param {Buffer | Uint8Array | undefined} logoBytes
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {boolean} include
 * @param {string[]} warnings
 * @returns {{buffer: Buffer, aspect: number} | null}
 */
function resolveLogo(logoBytes, theme, include, warnings) {
  if (!include) return null
  let buffer = null
  if (logoBytes !== undefined && logoBytes !== null && logoBytes.length > 0) {
    buffer = Buffer.isBuffer(logoBytes) ? logoBytes : Buffer.from(logoBytes)
  } else if (typeof theme.brandAssets?.logoPath === 'string' && existsSync(theme.brandAssets.logoPath)) {
    try {
      buffer = readFileSync(theme.brandAssets.logoPath)
    } catch {
      warnings.push(`logo file could not be read: ${theme.brandAssets.logoPath}`)
      return null
    }
  } else {
    return null
  }
  const size = imageSize(buffer, 'logo.png')
  const aspect = size.widthPx !== null && size.heightPx !== null && size.heightPx > 0
    ? size.widthPx / size.heightPx
    : (theme.brandAssets?.logoWidthPx && theme.brandAssets?.logoHeightPx
        ? theme.brandAssets.logoWidthPx / theme.brandAssets.logoHeightPx
        : 4)
  return { buffer, aspect }
}

/* ------------------------------------------------------------------ *
 * Element mappers
 * ------------------------------------------------------------------ */

/**
 * @param {any} slide
 * @param {import('./schema.js').TextElement} element
 * @param {import('./schema.js').ThemeTokens} theme
 */
function addTextElement(slide, element, theme) {
  const style = element.style ?? {}
  const color = splitAlpha(resolveColor(style.color, theme, theme.colors.textPrimary))
  const isDisplayRole = element.role === 'title' || element.role === 'subtitle'
  const fontFace = isDisplayRole ? theme.fonts.title : theme.fonts.body
  /** @type {Record<string, any>} */
  const textOptions = {
    x: pxToIn(element.x),
    y: pxToIn(element.y),
    w: pxToIn(element.w),
    h: pxToIn(element.h),
    fontSize: pxToPt(style.fontSize ?? 18),
    color: bareHex(color.hex),
    fontFace,
    bold: style.bold === true,
    italic: style.italic === true,
    align: style.align ?? 'left',
    valign: style.valign ?? 'top',
    margin: 0,
    breakLine: false,
    isTextBox: true,
    // Design doc §三.4: "开启 shrinkText: true 防止跨设备字体微差导致溢出".
    shrinkText: true,
    wrap: true,
  }
  if (style.lineSpacing !== undefined && Number.isFinite(style.lineSpacing)) {
    textOptions.lineSpacing = pxToPt(style.lineSpacing)
  }
  if (color.alpha < 1) textOptions.transparency = Math.round((1 - color.alpha) * 100)
  slide.addText(String(element.content ?? ''), textOptions)
}

/**
 * @param {any} slide
 * @param {import('./schema.js').ShapeElement} element
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {PptxGenJS} pres
 * @param {string} label
 * @param {string[]} warnings
 */
function addShapeElement(slide, element, theme, pres, label, warnings) {
  const style = element.style ?? {}
  const shapeName = SHAPE_TYPES[element.shapeType]
  if (shapeName === undefined) {
    warnings.push(`${label} · shape "${element.id}" has unknown shapeType "${element.shapeType}"`)
    return
  }
  /** @type {Record<string, any>} */
  const shapeOptions = {
    x: pxToIn(element.x),
    y: pxToIn(element.y),
    w: pxToIn(element.w),
    h: pxToIn(element.h),
  }

  const fill = style.fill === undefined ? null : splitAlpha(resolveColor(style.fill, theme, ''))
  const stroke = style.stroke === undefined ? null : splitAlpha(resolveColor(style.stroke, theme, ''))
  const opacity = Number.isFinite(style.opacity) ? Math.min(1, Math.max(0, style.opacity)) : 1

  if (fill !== null) {
    shapeOptions.fill = { color: bareHex(fill.hex), transparency: Math.round((1 - fill.alpha * opacity) * 100) }
  } else if (element.shapeType !== 'line') {
    shapeOptions.fill = { color: bareHex(theme.colors.surface), transparency: 100 } // no fill
  }
  if (stroke !== null) {
    shapeOptions.line = {
      color: bareHex(stroke.hex),
      width: Number.isFinite(style.strokeWidth) ? Math.max(0.25, style.strokeWidth) : 1,
      transparency: Math.round((1 - stroke.alpha * opacity) * 100),
    }
  } else if (element.shapeType === 'line') {
    // A line with no stroke is meaningless: default to the brand primary.
    shapeOptions.line = { color: bareHex(theme.colors.primary), width: 1 }
  }
  if (element.shapeType === 'roundRect') {
    shapeOptions.rectRadius = pxToIn(Number.isFinite(style.radius) ? style.radius : 12)
  }
  slide.addShape(pres.ShapeType[shapeName], shapeOptions)
}

/**
 * @param {any} slide
 * @param {import('./schema.js').SvgElement} element
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {CompileOptions} opts
 * @param {string} label
 * @param {string[]} warnings
 * @returns {boolean} true when an image part was added
 */
function addSvgElement(slide, element, theme, opts, label, warnings) {
  try {
    const png = rasterizeSvg(element.svgCode, {
      widthPx: element.w,
      heightPx: element.h,
      scale: opts.svgScale ?? 3,
      fontFiles: opts.fontFiles,
      fontDirs: opts.fontDirs,
      loadSystemFonts: opts.loadSystemFonts,
      color: theme.colors.primary,
    })
    slide.addImage({
      data: `image/png;base64,${png.toString('base64')}`,
      x: pxToIn(element.x),
      y: pxToIn(element.y),
      w: pxToIn(element.w),
      h: pxToIn(element.h),
      altText: element.description ?? `SVG diagram ${element.id}`,
    })
    return true
  } catch (cause) {
    warnings.push(`${label} · svg "${element.id}" could not be rasterized: ${messageOf(cause)}`)
    return false
  }
}

/**
 * @param {any} slide
 * @param {import('./schema.js').ChartElement} element
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {PptxGenJS} pres
 * @param {string} label
 * @param {string[]} warnings
 */
function addChartElement(slide, element, theme, pres, label, warnings) {
  const chartName = CHART_TYPES[element.chartType]
  if (chartName === undefined) {
    warnings.push(`${label} · chart "${element.id}" has unknown chartType "${element.chartType}"`)
    return
  }
  const labels = Array.isArray(element.data?.labels) ? element.data.labels.map(String) : []
  const series = Array.isArray(element.data?.series) ? element.data.series : []
  if (labels.length === 0 || series.length === 0) {
    warnings.push(`${label} · chart "${element.id}" has no data and was skipped`)
    return
  }
  const data = series.map((one) => ({
    name: String(one.name ?? 'series'),
    labels,
    values: one.values.map((value) => (Number.isFinite(value) ? value : 0)),
  }))
  const palette = theme.colors.chartPalette.map((hex) => bareHex(hex))
  /** @type {Record<string, any>} */
  const chartOptions = {
    x: pxToIn(element.x),
    y: pxToIn(element.y),
    w: pxToIn(element.w),
    h: pxToIn(element.h),
    chartColors: palette,
    showLegend: element.style?.legend === true,
    showValue: element.style?.showValues === true,
    showTitle: false,
    fontFace: theme.fonts.body,
    catAxisLabelFontFace: theme.fonts.body,
    valAxisLabelFontFace: theme.fonts.body,
    legendFontFace: theme.fonts.body,
    dataLabelFontFace: theme.fonts.body,
    catAxisLabelColor: bareHex(theme.colors.textSecondary),
    valAxisLabelColor: bareHex(theme.colors.textSecondary),
    dataLabelColor: bareHex(theme.colors.textPrimary),
  }
  try {
    slide.addChart(pres.ChartType[chartName], data, chartOptions)
  } catch (cause) {
    warnings.push(`${label} · chart "${element.id}" could not be added: ${messageOf(cause)}`)
  }
}

/**
 * @param {any} slide
 * @param {import('./schema.js').ImageElement} element
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {{buffer: Buffer} | null} logo
 * @param {Record<string, Buffer | Uint8Array | string>} assets
 * @param {string} label
 * @param {string[]} warnings
 * @returns {boolean}
 */
function addImageElement(slide, element, theme, logo, assets, label, warnings) {
  const asset = resolveAsset(element.assetKey, logo, assets, theme)
  if (asset === null) {
    warnings.push(`${label} · image "${element.id}" skipped: asset "${element.assetKey}" not found`)
    return false
  }
  /** @type {Record<string, any>} */
  const imageOptions = {
    x: pxToIn(element.x),
    y: pxToIn(element.y),
    w: pxToIn(element.w),
    h: pxToIn(element.h),
    altText: element.alt ?? element.assetKey,
  }
  if (asset.kind === 'data') {
    slide.addImage({ ...imageOptions, data: `image/png;base64,${asset.bytes.toString('base64')}` })
  } else {
    slide.addImage({ ...imageOptions, path: asset.path })
  }
  return true
}

/**
 * @param {string} assetKey
 * @param {{buffer: Buffer} | null} logo
 * @param {Record<string, Buffer | Uint8Array | string>} assets
 * @param {import('./schema.js').ThemeTokens} theme
 * @returns {{kind:'data', bytes: Buffer} | {kind:'path', path: string} | null}
 */
function resolveAsset(assetKey, logo, assets, theme) {
  const key = String(assetKey ?? '')
  if (key === 'logo' || key === '$logo') {
    if (logo !== null) return { kind: 'data', bytes: logo.buffer }
    const logoPath = theme.brandAssets?.logoPath
    if (typeof logoPath === 'string' && existsSync(logoPath)) return { kind: 'path', path: logoPath }
    return null
  }
  const candidate = assets?.[key]
  if (candidate === undefined || candidate === null) return null
  if (Buffer.isBuffer(candidate)) return { kind: 'data', bytes: candidate }
  if (candidate instanceof Uint8Array) return { kind: 'data', bytes: Buffer.from(candidate) }
  if (typeof candidate === 'string' && candidate.length > 0) {
    if (candidate.startsWith('data:')) {
      const base64 = candidate.slice(candidate.indexOf(',') + 1)
      return { kind: 'data', bytes: Buffer.from(base64, 'base64') }
    }
    const path = resolve(candidate)
    if (existsSync(path)) return { kind: 'path', path }
  }
  return null
}

/* ------------------------------------------------------------------ *
 * Post-processing: East-Asian font faces + slide-number field artefacts
 * ------------------------------------------------------------------ */

/** Fonts that can actually draw CJK glyphs (used to pick the `<a:ea>` face). */
const CJK_FACE_PATTERN =
  /[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]|yahei|pingfang|simhei|simsun|simkai|fangsong|dengxian|heiti|songti|kaiti|noto sans (sc|cjk|tc)|source han|思源|黑体|宋体|雅黑|楷体|ＭＳ|meiryo|malgun|ms gothic|hiragino/i

/** Standard placeholder text for an auto-numbered field (same as the HW template). */
const SLIDE_NUMBER_PLACEHOLDER = '‹#›'

/**
 * PptxGenJS writes a single `fontFace` into `<a:latin>`, `<a:ea>` and `<a:cs>`. A Latin
 * face in `<a:ea>` forces renderers to fall back for CJK text — LibreOffice then draws
 * the fallback glyphs *on top of* the Latin pass, which reads as doubled/overstruck text.
 *
 * @param {string} face
 * @returns {boolean} true when the face can render CJK
 */
export function isCjkCapableFace(face) {
  return typeof face === 'string' && face.trim().length > 0 && CJK_FACE_PATTERN.test(face)
}

/**
 * Pick the East-Asian typeface: explicit option → theme title face (CJK-capable) →
 * theme body face (CJK-capable) → `Microsoft YaHei`.
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {CompileOptions} [options]
 * @returns {string}
 */
export function resolveCjkFont(theme, options = {}) {
  const explicit = typeof options.cjkFont === 'string' ? options.cjkFont.trim() : ''
  if (explicit.length > 0) return explicit
  const candidates = [theme?.fonts?.title, theme?.fonts?.body, 'Microsoft YaHei']
  for (const candidate of candidates) {
    if (isCjkCapableFace(candidate)) return String(candidate).trim()
  }
  return 'Microsoft YaHei'
}

/**
 * Rewrite `<a:ea typeface="…"/>` so East-Asian runs use a CJK-capable face while
 * `<a:latin>` / `<a:cs>` keep the Latin face. Only literal typefaces are touched —
 * scheme references such as `+mj-ea` / `+mn-ea` stay as they are.
 *
 * @param {string} xml
 * @param {string} cjkFont
 * @returns {string}
 */
export function rewriteEastAsianFaces(xml, cjkFont) {
  return xml.replace(/<a:ea\b([^>]*?)(\/?)>/g, (full, attributes) => {
    const match = /\btypeface="([^"]*)"/.exec(attributes)
    const face = match === null ? '' : match[1]
    if (face.length === 0 || face.startsWith('+') || face === cjkFont) return full
    return full.replace(/\btypeface="[^"]*"/, `typeface="${escapeXmlAttribute(cjkFont)}"`)
  })
}

/**
 * Remove the junk text pptxgenjs leaves inside `<a:fld type="slidenum">`:
 * the master gets `<a:t>null</a:t>` (stringified "no slide number yet") and layouts get
 * the placeholder index (e.g. `<a:t>1002</a:t>`). Real slide numbers are preserved;
 * anything else is normalized to the standard `‹#›` placeholder.
 *
 * @param {string} xml
 * @param {'slide'|'template'} partKind
 * @returns {string}
 */
export function rewriteSlideNumberFields(xml, partKind) {
  return xml.replace(
    /(<a:fld\b[^>]*\btype="slidenum"[^>]*>)([\s\S]*?)(<\/a:fld>)/g,
    (full, open, inner, close) => {
      const next = inner.replace(/<a:t>([^<]*)<\/a:t>/, (textNode, text) => {
        const value = String(text).trim()
        if (partKind === 'slide' && /^\d+$/.test(value)) return textNode
        return `<a:t>${SLIDE_NUMBER_PLACEHOLDER}</a:t>`
      })
      return open + next + close
    },
  )
}

/**
 * Turn the page-number `sldNum` *placeholder* on a slide into a plain text box that
 * still carries the `<a:fld type="slidenum">` field.
 *
 * Why: PowerPoint only outlines placeholders in edit mode, but LibreOffice renders the
 * empty-placeholder prompt border in slide exports — the page number came out inside a
 * dashed rectangle. The HW template itself places its page number as a `userDrawn` text
 * box with a field (no `<p:ph>`), which is exactly the shape we produce here. The master
 * keeps its real placeholder so PowerPoint's Slide Number command still works.
 *
 * @param {string} xml
 * @returns {string}
 */
export function convertSlideNumberPlaceholders(xml) {
  return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (block) => {
    if (!/type="sldNum"/.test(block)) return block
    let next = block
      .replace(/<p:ph\b[^>]*type="sldNum"[^>]*\/?>/g, '')
      .replace(/<p:cNvSpPr>\s*<a:spLocks\b[^>]*\/>\s*<\/p:cNvSpPr>/g, '<p:cNvSpPr txBox="1"/>')
      .replace(/name="Slide Number Placeholder[^"]*"/g, 'name="Page Number"')
      .replace(/<p:nvPr>\s*<\/p:nvPr>/g, '<p:nvPr></p:nvPr>')
    const hasLine = /<a:ln\b[^>]*\/>|<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/.test(next)
    if (!hasLine) {
      // No border, no fill: the field text should sit directly on the page.
      next = next.replace(/<\/p:spPr>/, '<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>')
    }
    return next
  })
}

/**
 * Apply the XML fix-ups to the assembled package and re-zip it.
 *
 * @param {Buffer} buffer
 * @param {{cjkFont: string, warnings: string[]}} input
 * @returns {Promise<Buffer>}
 */
async function finalizePackage(buffer, input) {
  const { cjkFont, warnings } = input
  let zip
  try {
    zip = await JSZip.loadAsync(buffer)
  } catch (cause) {
    throw compileError('compile_verify_failed', 'pptxgenjs produced an unreadable package', cause)
  }

  const partPattern = /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/
  let rewritten = 0
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir === true || !partPattern.test(name)) continue
    const original = await zip.file(name).async('string')
    const isSlide = name.startsWith('ppt/slides/')
    let next = rewriteEastAsianFaces(original, cjkFont)
    next = rewriteSlideNumberFields(next, isSlide ? 'slide' : 'template')
    if (isSlide) next = convertSlideNumberPlaceholders(next)
    if (next !== original) {
      zip.file(name, next)
      rewritten += 1
    }
  }

  if (rewritten === 0) return buffer
  warnings.push(`font/field fix-up applied to ${rewritten} slide part(s) (East-Asian face → ${cjkFont})`)
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

/**
 * @param {string} value
 * @returns {string}
 */
function escapeXmlAttribute(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

/* ------------------------------------------------------------------ *
 * Verification / helpers
 * ------------------------------------------------------------------ */

/**
 * Re-open the compiled buffer and check it is a usable presentation.
 * @param {Buffer} buffer
 * @param {number} expectedSlides
 * @param {string[]} warnings
 * @returns {Promise<boolean>}
 */
async function verifyPackage(buffer, expectedSlides, warnings) {
  /** @type {JSZip} */
  let zip
  try {
    zip = await JSZip.loadAsync(buffer)
  } catch (cause) {
    throw compileError('compile_verify_failed', 'compiled output is not a readable zip package', cause)
  }
  if (zip.file('ppt/presentation.xml') === null) {
    throw compileError('compile_verify_failed', 'compiled output has no ppt/presentation.xml')
  }
  const slideParts = Object.keys(zip.files).filter(
    (name) => /^ppt\/slides\/slide\d+\.xml$/.test(name) && zip.files[name].dir !== true,
  )
  if (slideParts.length !== expectedSlides) {
    warnings.push(
      `compiled package contains ${slideParts.length} slide part(s) but the deck has ${expectedSlides} slide(s)`,
    )
    return false
  }
  const masterParts = Object.keys(zip.files).filter((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name))
  if (masterParts.length === 0) {
    warnings.push('compiled package has no slide master part')
    return false
  }
  return true
}

/**
 * Text overflow / volume warnings for one slide (design doc §三.2: ≤150 chars).
 * @param {import('./schema.js').CardSlide} slideDef
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {string} label
 * @param {string[]} warnings
 */
function reportOverflow(slideDef, theme, label, warnings) {
  let report
  try {
    report = estimateTextOverflow({ elements: slideDef.elements ?? [] }, theme)
  } catch {
    return
  }
  for (const box of report.overflowing ?? []) {
    warnings.push(
      `${label} · text "${box.id}" likely overflows (${box.estimatedLines} lines ≈ ` +
        `${box.neededHeight}px in a ${box.declaredHeight}px box); shrinkText will fit it in PowerPoint`,
    )
  }
}

/**
 * Fill in any missing part of the theme so the compiler never reads `undefined`.
 * @param {import('./schema.js').ThemeTokens | undefined} theme
 * @returns {import('./schema.js').ThemeTokens}
 */
function normalizeTheme(theme) {
  const base = structuredClone(DEFAULT_THEME)
  if (theme === undefined || theme === null) return base
  return {
    templateId: theme.templateId ?? base.templateId,
    templateName: theme.templateName ?? base.templateName,
    colors: {
      background: normalizeColor(theme.colors?.background, base.colors.background),
      surface: normalizeColor(theme.colors?.surface, base.colors.surface),
      textPrimary: normalizeColor(theme.colors?.textPrimary, base.colors.textPrimary),
      textSecondary: normalizeColor(theme.colors?.textSecondary, base.colors.textSecondary),
      primary: normalizeColor(theme.colors?.primary, base.colors.primary),
      accent: normalizeColor(theme.colors?.accent, base.colors.accent),
      chartPalette:
        Array.isArray(theme.colors?.chartPalette) && theme.colors.chartPalette.length > 0
          ? theme.colors.chartPalette.map((hex) => normalizeColor(hex, base.colors.primary))
          : base.colors.chartPalette,
    },
    fonts: {
      title: theme.fonts?.title ?? base.fonts.title,
      body: theme.fonts?.body ?? base.fonts.body,
      latin: theme.fonts?.latin ?? base.fonts.latin,
    },
    brandAssets: {
      icons: theme.brandAssets?.icons ?? {},
      ...(theme.brandAssets?.logoPath === undefined ? {} : { logoPath: theme.brandAssets.logoPath }),
      ...(theme.brandAssets?.logoWidthPx === undefined ? {} : { logoWidthPx: theme.brandAssets.logoWidthPx }),
      ...(theme.brandAssets?.logoHeightPx === undefined ? {} : { logoHeightPx: theme.brandAssets.logoHeightPx }),
      ...(theme.brandAssets?.logoFrame === undefined ? {} : { logoFrame: theme.brandAssets.logoFrame }),
    },
  }
}

/**
 * @param {unknown} value
 * @param {string} fallback
 * @returns {string}
 */
function normalizeColor(value, fallback) {
  if (typeof value !== 'string') return fallback
  try {
    return normalizeHex(value)
  } catch {
    return fallback
  }
}

/**
 * @param {string} code
 * @param {string} message
 * @param {unknown} [cause]
 */
function compileError(code, message, cause) {
  const error = new Error(message)
  error.code = code
  if (cause !== undefined) error.cause = cause
  return error
}

/** @param {unknown} cause */
function messageOf(cause) {
  return cause instanceof Error ? cause.message : String(cause)
}

/**
 * @param {number} value
 * @param {number} digits
 * @returns {number}
 */
function round(value, digits) {
  const factor = 10 ** digits
  return Math.round(value * factor) / factor
}

/**
 * @param {string} bare `RRGGBB` or `#RRGGBB`
 * @returns {string} `RRGGBB`
 */
function bareHex(bare) {
  return String(bare).replace('#', '').toUpperCase().slice(-6)
}

/**
 * Atomic write (temp + rename) so a failed export never truncates a good file.
 * @param {string} target
 * @param {Buffer} bytes
 */
function writeFileAtomic(target, bytes) {
  mkdirSync(dirname(target), { recursive: true })
  const temp = `${target}.${process.pid}.${randomBytes(4).toString('hex')}.tmp`
  writeFileSync(temp, bytes)
  try {
    renameSync(temp, target)
  } catch (cause) {
    try {
      unlinkSync(temp)
    } catch {
      /* best effort */
    }
    throw compileError('write_failed', `cannot write ${target}`, cause)
  }
}
