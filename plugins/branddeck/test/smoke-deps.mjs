/**
 * Dependency smoke test: proves the pure-JS + native stack actually loads and
 * performs the two operations the compiler depends on (SVG rasterization and
 * PPTX serialization) in this exact Node runtime.
 *
 * Run: node test/smoke-deps.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const outDir = join(here, '..', '.tmp')
mkdirSync(outDir, { recursive: true })

const results = []

async function check(name, fn) {
  const started = Date.now()
  try {
    const value = await fn()
    results.push({ name, ok: true, ms: Date.now() - started, value })
    console.log(`PASS ${name} (${Date.now() - started}ms) ${value === undefined ? '' : JSON.stringify(value)}`)
  } catch (error) {
    results.push({ name, ok: false, ms: Date.now() - started, error: String(error?.stack ?? error) })
    console.log(`FAIL ${name}: ${String(error?.stack ?? error)}`)
  }
}

await check('jszip', async () => {
  const m = await import('jszip')
  const zip = new m.default()
  zip.file('hello.xml', '<a:hello/>')
  const buf = await zip.generateAsync({ type: 'nodebuffer' })
  const reopened = await m.default.loadAsync(buf)
  const text = await reopened.file('hello.xml').async('string')
  if (text !== '<a:hello/>') throw new Error('roundtrip mismatch')
  return { bytes: buf.length }
})

await check('fast-xml-parser', async () => {
  const { XMLParser } = await import('fast-xml-parser')
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    removeNSPrefix: true,
  })
  const doc = parser.parse(
    '<a:theme xmlns:a="x"><a:clrScheme name="H"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1></a:clrScheme></a:theme>',
  )
  const dk1 = doc.theme.clrScheme.dk1.sysClr
  if (dk1['@_lastClr'] !== '000000') throw new Error(`unexpected parse: ${JSON.stringify(doc)}`)
  return { root: Object.keys(doc) }
})

await check('zod', async () => {
  const { z } = await import('zod')
  const s = z.object({ a: z.number() })
  if (!s.safeParse({ a: 1 }).success) throw new Error('zod rejected valid input')
  return { version: z.string().constructor.name }
})

await check('pptxgenjs write', async () => {
  const { default: PptxGenJS } = await import('pptxgenjs')
  const pres = new PptxGenJS()
  pres.layout = 'LAYOUT_WIDE'
  const slide = pres.addSlide()
  slide.background = { color: 'FFFFFF' }
  slide.addText('BrandDeck 品牌演示', {
    x: 0.64, y: 0.64, w: 8, h: 1,
    fontSize: 32, color: '1A1A1A', bold: true, fontFace: 'Microsoft YaHei',
    shrinkText: true,
  })
  slide.addShape(pres.ShapeType.roundRect, { x: 0.64, y: 2, w: 3, h: 1.2, fill: { color: 'EEF2F7' }, line: { color: '2F5BD2', width: 1 } })
  slide.addChart(pres.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], { x: 5, y: 2, w: 3, h: 2 })
  const buf = await pres.write({ outputType: 'nodebuffer' })
  const path = join(outDir, 'smoke.pptx')
  writeFileSync(path, buf)
  return { bytes: buf.length, path }
})

await check('@resvg/resvg-js rasterize', async () => {
  const { Resvg } = await import('@resvg/resvg-js')
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200" viewBox="0 0 400 200">
    <rect x="0" y="0" width="400" height="200" fill="#0f2f6b"/>
    <circle cx="80" cy="100" r="50" fill="#3d7bfd"/>
    <path d="M150 140 L230 60 L310 140 Z" fill="#ffd166"/>
    <text x="200" y="185" font-size="22" fill="#ffffff" text-anchor="middle">BrandDeck SVG</text>
  </svg>`
  const resvg = new Resvg(svg, { background: 'rgba(0,0,0,0)', fitTo: { mode: 'zoom', value: 3 } })
  const png = resvg.render().asPng()
  const path = join(outDir, 'smoke-svg.png')
  writeFileSync(path, png)
  const resvgDefault = new Resvg(svg, { fitTo: { mode: 'width', value: 800 } })
  const png2 = resvgDefault.render().asPng()
  if (png.length < 1000) throw new Error('rasterized PNG suspiciously small')
  return { bytes: png.length, scaledBytes: png2.length, path }
})

const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} checks passed`)
if (failed.length > 0) process.exitCode = 1
