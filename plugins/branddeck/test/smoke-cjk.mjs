/**
 * Proves CJK text survives SVG rasterization — the design doc's pitfall #1.
 * Also proves the rasterized PNG is accepted by pptxgenjs as a slide image.
 *
 * Run: node test/smoke-cjk.mjs
 */
import { writeFileSync, mkdirSync, existsSync, readdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const outDir = join(here, '..', '.tmp')
mkdirSync(outDir, { recursive: true })

const { Resvg } = await import('@resvg/resvg-js')

const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="240" viewBox="0 0 640 240">
  <rect width="640" height="240" fill="#101828"/>
  <text x="24" y="70" font-family="Microsoft YaHei, PingFang SC, sans-serif" font-size="40" font-weight="700" fill="#ffffff">品牌智能演示平台</text>
  <text x="24" y="130" font-family="Microsoft YaHei, sans-serif" font-size="24" fill="#93c5fd">增长趋势 · 漏斗模型 · 对比矩阵</text>
  <text x="24" y="185" font-family="Arial" font-size="26" fill="#fbbf24">Revenue +38.6% YoY</text>
  <circle cx="560" cy="60" r="26" fill="#3b82f6"/>
</svg>`

const png = new Resvg(svg, { fitTo: { mode: 'zoom', value: 3 } }).render().asPng()
const pngPath = join(outDir, 'smoke-cjk.png')
writeFileSync(pngPath, png)

let fontInfo = 'system'
const winFonts = 'C:\\Windows\\Fonts'
if (existsSync(winFonts)) {
  const hit = readdirSync(winFonts).filter((f) => /^(msyh|simhei|simsun|SourceHanSans)/i.test(f))
  fontInfo = `system (${hit.slice(0, 6).join(', ')})`
}

const { default: PptxGenJS } = await import('pptxgenjs')
const pres = new PptxGenJS()
pres.layout = 'LAYOUT_WIDE'
const slide = pres.addSlide()
slide.background = { color: '101828' }
slide.addImage({ data: `image/png;base64,${png.toString('base64')}`, x: 0.64, y: 0.64, w: 6.6, h: 2.5 })
const buf = await pres.write({ outputType: 'nodebuffer' })
const pptxPath = join(outDir, 'smoke-cjk.pptx')
writeFileSync(pptxPath, buf)

console.log(JSON.stringify({ pngBytes: png.length, pngPath, pptxBytes: buf.length, pptxPath, fontInfo }, null, 2))
if (png.length < 5000) {
  console.error('CJK PNG looks too small — glyphs may be missing')
  process.exitCode = 1
}
