/**
 * svg.js: prepareSvg normalization, 3× CJK rasterization, typed failures and the
 * source-size / zero-box guards.
 *
 * Run: node test/engine-svg.mjs
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  OUT_DIR,
  assert,
  assertCode,
  assertEqual,
  assertMatch,
  run,
  test,
} from './engine-helpers.mjs'
import { MAX_SVG_SOURCE_BYTES, prepareSvg, rasterizeSvg } from '../src/svg.js'

mkdirSync(OUT_DIR, { recursive: true })

const CJK_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="240" viewBox="0 0 640 240">
  <rect width="640" height="240" fill="#FFFFFF"/>
  <rect x="24" y="24" width="8" height="80" fill="#C7000A"/>
  <text x="52" y="80" font-family="Microsoft YaHei, sans-serif" font-size="40" font-weight="700" fill="#1D1D1A">品牌智能演示平台</text>
  <text x="52" y="140" font-family="Microsoft YaHei, sans-serif" font-size="24" fill="#666666">增长趋势 · 漏斗模型 · 对比矩阵</text>
  <text x="52" y="196" font-family="Arial" font-size="26" fill="#C7000A">Revenue +38.6% YoY</text>
</svg>`

/**
 * Read the pixel size out of a PNG buffer (IHDR).
 * @param {Buffer} png
 * @returns {{width: number, height: number}}
 */
function pngSize(png) {
  assert(Buffer.isBuffer(png), 'PNG is a Buffer')
  assertMatch(png.subarray(0, 8).toString('hex'), /^89504e470d0a1a0a$/, 'PNG magic bytes')
  assertEqual(png.toString('latin1', 12, 16), 'IHDR', 'first chunk is IHDR')
  return { width: png.readUInt32BE(16), height: png.readUInt32BE(20) }
}

await test('prepareSvg: forces the box, adds xmlns, keeps or synthesizes viewBox', () => {
  const withoutViewBox = prepareSvg('<svg><rect width="10" height="10"/></svg>', 320, 180)
  assertMatch(withoutViewBox, /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg"/, 'xmlns added first')
  assertMatch(withoutViewBox, /width="320"/, 'width forced to the element box')
  assertMatch(withoutViewBox, /height="180"/, 'height forced to the element box')
  assertMatch(withoutViewBox, /viewBox="0 0 320 180"/, 'viewBox synthesized when absent')
  const rootTag = /^<svg\b[^>]*>/.exec(withoutViewBox)[0]
  assertEqual((rootTag.match(/\swidth=/g) ?? []).length, 1, 'exactly one width on the root element')
  assert(!rootTag.includes('width="10"'), 'the root width attribute is replaced, not duplicated')

  const withViewBox = prepareSvg('<svg xmlns="http://www.w3.org/2000/svg" width="1000" height="500" viewBox="0 0 1000 500"><g/></svg>', 500, 250)
  assertMatch(withViewBox, /viewBox="0 0 1000 500"/, 'existing viewBox is preserved')
  assertMatch(withViewBox, /width="500"/, 'width replaced')
  assertMatch(withViewBox, /height="250"/, 'height replaced')

  const colored = prepareSvg('<svg><rect fill="currentColor"/></svg>', 10, 10, '#C7000A')
  assert(!colored.includes('currentColor'), 'currentColor substituted')
  assertMatch(colored, /fill="#C7000A"/, 'explicit colour injected')

  const untouched = prepareSvg('<svg><rect fill="currentColor"/></svg>', 10, 10)
  assert(untouched.includes('currentColor'), 'currentColor kept when no colour is supplied')

  const prefixed = prepareSvg('<?xml version="1.0"?>\n<!-- comment --><svg width="4" height="4"/>', 4, 4)
  assert(!prefixed.includes('<?xml'), 'XML declaration removed')
  assertMatch(prefixed, /^<svg/, 'root starts at <svg')
  return 'normalization ok'
})

await test('prepareSvg: typed errors for a missing root and an empty box', async () => {
  await assertCode(() => prepareSvg('<div>not svg</div>', 100, 100), 'svg_missing_root', 'no <svg> root')
  await assertCode(() => prepareSvg('<svg><g/></svg>', 0, 100), 'svg_invalid_box', 'zero width')
  await assertCode(() => prepareSvg('<svg><g/></svg>', 100, -5), 'svg_invalid_box', 'negative height')
  return 'typed errors'
})

await test('rasterizeSvg: CJK at 3× — PNG magic, 3× pixel size, real glyphs', () => {
  const png = rasterizeSvg(CJK_SVG, { widthPx: 640, heightPx: 240, scale: 3 })
  const size = pngSize(png)
  assertEqual(size, { width: 1920, height: 720 }, '3× of a 640×240 box')
  assert(png.length > 5000, `PNG is not suspiciously small (${png.length} bytes)`)
  writeFileSync(join(OUT_DIR, 'svg-cjk-3x.png'), png)

  const single = rasterizeSvg(CJK_SVG, { widthPx: 640, heightPx: 240, scale: 1 })
  const singleSize = pngSize(single)
  assertEqual(singleSize, { width: 640, height: 240 }, '1× of the same box')
  assert(png.length > single.length, 'the 3× PNG is bigger than the 1× PNG (more glyph detail)')
  return { bytes3x: png.length, bytes1x: single.length, px: `${size.width}x${size.height}` }
})

await test('rasterizeSvg: CJK renders with explicit font options (container-safe path)', () => {
  const png = rasterizeSvg(CJK_SVG, {
    widthPx: 320,
    heightPx: 120,
    scale: 3,
    loadSystemFonts: true,
    fontFiles: [],
    fontDirs: [],
  })
  const size = pngSize(png)
  assertEqual(size, { width: 960, height: 360 }, 'font options do not change the geometry')
  writeFileSync(join(OUT_DIR, 'svg-cjk-fontopts.png'), png)
  return `${png.length} bytes`
})

await test('rasterizeSvg: guards and typed resvg failures', async () => {
  // >200 KiB of source is refused before resvg ever runs.
  const big = `<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><desc>${'x'.repeat(MAX_SVG_SOURCE_BYTES)}</desc></svg>`
  await assertCode(
    () => rasterizeSvg(big, { widthPx: 10, heightPx: 10 }),
    'svg_too_large',
    'oversized source is refused',
  )
  // Zero/negative box is refused by prepareSvg.
  await assertCode(
    () => rasterizeSvg('<svg><rect/></svg>', { widthPx: 0, heightPx: 0 }),
    'svg_invalid_box',
    'empty box is refused',
  )
  // No <svg> root at all.
  await assertCode(() => rasterizeSvg('totally not svg', { widthPx: 10, heightPx: 10 }), 'svg_missing_root', 'garbage source')
  // A structurally broken document must surface as svg_raster_failed with the resvg message.
  const broken = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="10" height="10"></svg>'
  const error = await assertCode(
    () => rasterizeSvg(broken, { widthPx: 100, heightPx: 100 }),
    'svg_raster_failed',
    'resvg parse failure is typed',
  )
  assertMatch(error.message, /SVG rasterization failed/, 'error message carries the resvg failure')
  return 'guards ok'
})

await test('rasterizeSvg: box falls back to the SVG width/height when options omit it', () => {
  const png = rasterizeSvg(CJK_SVG, { scale: 2 })
  assertEqual(pngSize(png), { width: 1280, height: 480 }, '2× of the declared 640×240')
  return 'fallback box'
})

run('engine-svg')
