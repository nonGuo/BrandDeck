/**
 * SVG → PNG rasterization with `@resvg/resvg-js`.
 *
 * The compiler never embeds model-authored SVG into the `.pptx`: Office versions
 * disagree about complex SVG, so we rasterise at 3× (design doc §三.4) and embed a PNG.
 *
 * Resolution: the canvas is 1280×720 CSS px, i.e. 96 dpi. Rendering at `scale: 3`
 * yields 288 dpi — within 4 % of the design doc's 300 DPI target, and the same number
 * the design doc calls out ("3 倍分辨率（300 DPI）").
 *
 * CJK pitfall (design doc pitfall #1): resvg does **not** read fontconfig on a bare
 * container, so callers can pass explicit `fontFiles`/`fontDirs`. On a Windows desktop
 * `loadSystemFonts: true` (the default here) picks up `C:\Windows\Fonts`.
 */
import { Resvg } from '@resvg/resvg-js'

/** Hard cap on model-authored SVG source (design doc: inline diagrams only). */
export const MAX_SVG_SOURCE_BYTES = 200 * 1024

/** Allowed rasterization scale window (3 = 288 dpi). */
const MIN_SCALE = 0.1
const MAX_SCALE = 8

/**
 * Typed error for the SVG pipeline.
 * @param {'svg_too_large'|'svg_invalid_box'|'svg_raster_failed'|'svg_missing_root'} code
 * @param {string} message
 * @param {unknown} [cause]
 * @returns {Error & {code: string, cause?: unknown}}
 */
function svgError(code, message, cause) {
  const error = new Error(message)
  error.code = code
  if (cause !== undefined) error.cause = cause
  return error
}

/**
 * Normalize model-authored SVG so resvg renders exactly the element box:
 * - add `xmlns` when missing,
 * - force `width`/`height` to the element box in px,
 * - keep an existing `viewBox`, synthesize `0 0 w h` when absent,
 * - optionally substitute CSS `currentColor` (resvg has no cascade to inherit from, so
 *   a bare `currentColor` would otherwise render black).
 *
 * @param {string} svgCode
 * @param {number} widthPx
 * @param {number} heightPx
 * @param {string} [color] explicit colour for `currentColor` (e.g. `#C7000A`)
 * @returns {string}
 * @throws {Error & {code:'svg_missing_root'|'svg_invalid_box'}}
 */
export function prepareSvg(svgCode, widthPx, heightPx, color) {
  const rawWidth = Number(widthPx)
  const rawHeight = Number(heightPx)
  if (
    !Number.isFinite(rawWidth) ||
    !Number.isFinite(rawHeight) ||
    rawWidth <= 0 ||
    rawHeight <= 0
  ) {
    throw svgError('svg_invalid_box', `SVG box must be positive, got ${widthPx}x${heightPx}`)
  }
  const width = roundPx(rawWidth)
  const height = roundPx(rawHeight)

  let out = String(svgCode ?? '').trim()
  const start = out.indexOf('<svg')
  if (start === -1) throw svgError('svg_missing_root', 'SVG source has no <svg> root element')
  out = out.slice(start)
  out = out.replace(/<\?xml[^>]*\?>/gi, '')
  if (color !== undefined && color !== null && color !== '') {
    out = out.replace(/currentColor/g, color)
  }

  const rootTag = /^<svg\b[^>]*>/i.exec(out)
  if (rootTag === null) throw svgError('svg_missing_root', 'SVG root element is not terminated')
  let tag = rootTag[0]

  if (!/\sxmlns\s*=/.test(tag)) {
    tag = tag.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"')
  }
  const viewBox = /\sviewBox\s*=\s*"([^"]*)"/.exec(tag)
  tag = tag.replace(/\swidth\s*=\s*"[^"]*"/i, '').replace(/\sheight\s*=\s*"[^"]*"/i, '')
  let attributes = tag.slice('<svg'.length, -1).trimEnd()
  attributes += ` width="${width}" height="${height}"`
  if (viewBox === null) attributes += ` viewBox="0 0 ${width} ${height}"`
  const rebuilt = `<svg${attributes.startsWith(' ') ? '' : ' '}${attributes}>`

  return rebuilt + out.slice(rootTag[0].length)
}

/**
 * @typedef {object} RasterizeOptions
 * @property {number} [widthPx]  element box width (px)
 * @property {number} [heightPx] element box height (px)
 * @property {number} [scale]    output scale, default 3 (288 dpi)
 * @property {string[]} [fontFiles] explicit `.ttf`/`.otf` files (container-safe)
 * @property {string[]} [fontDirs]  explicit font directories
 * @property {boolean} [loadSystemFonts] default `true` (Windows/LibreOffice hosts)
 * @property {string} [color]    colour substituted for `currentColor`
 * @property {string} [background] default transparent
 */

/**
 * Rasterize SVG source to a transparent PNG Buffer.
 *
 * @param {string} svgCode
 * @param {RasterizeOptions} [options]
 * @returns {Buffer} PNG bytes
 * @throws {Error & {code:'svg_too_large'|'svg_invalid_box'|'svg_missing_root'|'svg_raster_failed'}}
 */
export function rasterizeSvg(svgCode, options = {}) {
  const source = String(svgCode ?? '')
  const byteLength = Buffer.byteLength(source, 'utf8')
  if (byteLength > MAX_SVG_SOURCE_BYTES) {
    throw svgError(
      'svg_too_large',
      `SVG source is ${byteLength} bytes, above the ${MAX_SVG_SOURCE_BYTES} byte limit`,
    )
  }
  const { widthPx, heightPx } = resolveBox(source, options)
  const scale = clampScale(options.scale ?? 3)
  const prepared = prepareSvg(source, widthPx, heightPx, options.color)

  const font = {
    fontFiles: Array.isArray(options.fontFiles) ? [...options.fontFiles] : [],
    fontDirs: Array.isArray(options.fontDirs) ? [...options.fontDirs] : [],
    loadSystemFonts: options.loadSystemFonts !== false,
  }

  try {
    const renderer = new Resvg(prepared, {
      background: options.background ?? 'rgba(0,0,0,0)',
      fitTo: { mode: 'zoom', value: scale },
      font,
    })
    const png = renderer.render().asPng()
    if (png === undefined || png === null || png.length === 0) {
      throw svgError('svg_raster_failed', 'resvg returned an empty PNG')
    }
    return Buffer.from(png)
  } catch (cause) {
    if (/** @type {any} */ (cause)?.code === 'svg_raster_failed') throw cause
    const message = cause instanceof Error ? cause.message : String(cause)
    throw svgError('svg_raster_failed', `SVG rasterization failed: ${message}`, cause)
  }
}

/**
 * Determine the element box, falling back to the SVG's own width/height/viewBox when the
 * caller did not supply one.
 * @param {string} source
 * @param {RasterizeOptions} options
 * @returns {{widthPx: number, heightPx: number}}
 */
function resolveBox(source, options) {
  let widthPx = Number(options.widthPx)
  let heightPx = Number(options.heightPx)
  if (widthPx > 0 && heightPx > 0) return { widthPx, heightPx }

  const rootTag = /<svg\b[^>]*>/i.exec(source)
  const tag = rootTag === null ? '' : rootTag[0]
  const attr = (name) => {
    const match = new RegExp(`\\s${name}\\s*=\\s*"([^"]*)"`, 'i').exec(tag)
    if (match === null) return undefined
    const value = parseFloat(match[1])
    return Number.isFinite(value) && value > 0 ? value : undefined
  }
  const viewBox = /\sviewBox\s*=\s*"([^"]*)"/i.exec(tag)
  const box = viewBox === null ? null : viewBox[1].split(/[\s,]+/).map(Number)

  if (!(widthPx > 0)) {
    widthPx = attr('width') ?? (box !== null && Number.isFinite(box[2]) ? box[2] : undefined)
  }
  if (!(heightPx > 0)) {
    heightPx = attr('height') ?? (box !== null && Number.isFinite(box[3]) ? box[3] : undefined)
  }
  if (!(widthPx > 0) || !(heightPx > 0)) {
    throw svgError('svg_invalid_box', 'SVG box is empty: pass widthPx/heightPx or set width/height/viewBox')
  }
  return { widthPx, heightPx }
}

/** @param {number} value */
function clampScale(value) {
  const scale = Number(value)
  if (!Number.isFinite(scale) || scale <= 0) return 3
  return Math.min(MAX_SCALE, Math.max(MIN_SCALE, scale))
}

/** Render at whole pixels to keep resvg output deterministic. */
function roundPx(value) {
  const number = Number(value)
  return Number.isFinite(number) ? Math.max(1, Math.round(number)) : 0
}
