// Diagnostic: dump autofit + run properties of slide 1 and the master object list.
import { readFileSync } from 'node:fs'
import JSZip from 'jszip'

const pptx = process.argv[2] ?? '.tmp/e2e/hw-demo.pptx'
const zip = await JSZip.loadAsync(readFileSync(pptx))

const slide = await zip.file('ppt/slides/slide1.xml').async('string')
console.log('--- autofit on slide1')
for (const m of slide.matchAll(/<a:(normAutofit|spAutoFit|noAutofit)[^>]*\/>/g)) console.log('  ', m[0])
console.log('--- rPr on slide1')
for (const m of slide.matchAll(/<a:rPr[^>]*>/g)) console.log('  ', m[0])
console.log('--- pPr chart of first shape')
for (const m of slide.matchAll(/<a:pPr[^>]*>/g)) console.log('  ', m[0])

const master = await zip.file('ppt/slideMasters/slideMaster1.xml').async('string')
console.log('--- master shape names')
for (const m of master.matchAll(/<p:cNvPr id="\d+" name="([^"]*)"/g)) console.log('  ', m[1])
console.log('--- master text runs')
for (const m of master.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)) console.log('  ', JSON.stringify(m[1]))
console.log('--- master pictures')
for (const m of master.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)) console.log('  pic present, len', m[0].length)

console.log('--- slide rels')
const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string')
console.log(rels.replace(/></g, '>\n<'))
