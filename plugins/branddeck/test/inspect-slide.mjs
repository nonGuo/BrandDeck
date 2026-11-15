// Inspect one slide's text bodies in an exported pptx, to explain render artifacts.
import { readFileSync } from 'node:fs'
import JSZip from 'jszip'

const pptx = process.argv[2]
const slidePart = process.argv[3] ?? 'ppt/slides/slide1.xml'
const zip = await JSZip.loadAsync(readFileSync(pptx))
const xml = await zip.file(slidePart).async('string')

// Crude but effective: list every <a:t> run with its containing shape id.
const shapes = xml.split('<p:sp>').slice(1)
console.log(`shapes: ${shapes.length}`)
for (const shape of shapes) {
  const name = /<p:cNvPr[^>]*name="([^"]*)"/.exec(shape)?.[1] ?? '?'
  const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(shape)
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(shape)
  const bodyPr = /<a:bodyPr[^>]*>/.exec(shape)?.[0] ?? ''
  const runs = [...shape.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1])
  if (runs.length === 0 && !off) continue
  const geom = off && ext
    ? `x=${(Number(off[1]) / 12192000 * 13.333 * 96).toFixed(0)} y=${(Number(off[2]) / 6858000 * 7.5 * 96).toFixed(0)} w=${(Number(ext[1]) / 12192000 * 13.333 * 96).toFixed(0)} h=${(Number(ext[2]) / 6858000 * 7.5 * 96).toFixed(0)}`
    : 'no-xfrm'
  console.log(`- ${name} | ${geom} | bodyPr=${bodyPr.replace(/\s+/g, ' ').slice(0, 90)}`)
  for (const run of runs) console.log(`    run: ${JSON.stringify(run)}`)
}
