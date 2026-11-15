// Which pptxgenjs option actually writes font faces into the OOXML?
import { writeFileSync } from 'node:fs'
import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'

const variants = {
  fontFace: { fontFace: 'Microsoft YaHei' },
  typeface: { typeface: 'Microsoft YaHei' },
  both: { fontFace: 'Microsoft YaHei', typeface: 'Microsoft YaHei' },
}

for (const [name, options] of Object.entries(variants)) {
  const pres = new PptxGenJS()
  const slide = pres.addSlide()
  slide.addText('品牌标题 Brand', { x: 1, y: 1, w: 6, h: 1, fontSize: 24, ...options })
  const buffer = await pres.write({ outputType: 'nodebuffer' })
  const zip = await JSZip.loadAsync(buffer)
  const xml = await zip.file('ppt/slides/slide1.xml').async('string')
  const fonts = [...xml.matchAll(/<a:(?:rFonts|latin|ea)[^>]*\/>/g)].map((m) => m[0])
  const theme = await zip.file('ppt/theme/theme1.xml').async('string')
  const themeFonts = [...theme.matchAll(/<a:(?:latin|ea)[^>]*typeface="([^"]*)"/g)].map((m) => m[1])
  console.log(`--- ${name}`)
  console.log('  slide font tags:', fonts.length === 0 ? '(none)' : fonts.join(' '))
  console.log('  rPr sample:', (/<a:rPr[^>]*>/.exec(xml) ?? [''])[0])
  console.log('  theme major/minor faces:', themeFonts.slice(0, 2).join(' / '))
  writeFileSync(`.tmp/font-${name}.pptx`, buffer)
}
