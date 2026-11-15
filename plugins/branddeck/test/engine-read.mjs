/**
 * read.js: OPC unpacking, XML shape, relationship resolution, error codes.
 *
 * Run: node test/engine-read.mjs
 */
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  TEMPLATE_PATH,
  TMP_DIR,
  assert,
  assertCode,
  assertEqual,
  assertMatch,
  run,
  test,
} from './engine-helpers.mjs'
import { XML_OPTIONS, asArray, loadPackage, normalizePart, parseXml, relsPathFor } from '../src/pptx/read.js'
import { mkdirSync } from 'node:fs'

mkdirSync(TMP_DIR, { recursive: true })

/**
 * Build a minimal but valid-enough OPC zip in memory (used only for relationship edge
 * cases the real template does not contain, e.g. `TargetMode="External"`).
 * @param {{rels: string, extra?: Record<string, string>}} input
 * @returns {Promise<Buffer>}
 */
async function syntheticPackage(input) {
  const zip = new JSZip()
  zip.file(
    'ppt/presentation.xml',
    '<?xml version="1.0"?><p:presentation xmlns:p="p" xmlns:r="r"><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
  )
  zip.file('ppt/_rels/presentation.xml.rels', input.rels)
  for (const [part, text] of Object.entries(input.extra ?? {})) zip.file(part, text)
  return zip.generateAsync({ type: 'nodebuffer' })
}

await test('loadPackage(path): opens the real HW template and lists its parts', async () => {
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const parts = pkg.partNames()
  assert(parts.length > 50, `expected >50 parts, got ${parts.length}`)
  assert(pkg.has('ppt/presentation.xml'), 'has ppt/presentation.xml')
  assert(pkg.has('/ppt/theme/theme1.xml'), 'has() tolerates a leading slash')
  assert(!pkg.has('ppt/nope.xml'), 'missing part reports false')
  assertEqual(relsPathFor('ppt/slides/slide1.xml'), 'ppt/slides/_rels/slide1.xml.rels', 'rels path helper')
  return `${parts.length} parts, ${pkg.byteLength} bytes`
})

await test('loadPackage: text()/buffer()/xml() read the real parts, xml() is cached', async () => {
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const presentationText = pkg.text('ppt/presentation.xml')
  assertMatch(presentationText, /<p:sldSz cx="12196763" cy="6858000"\/>/, 'slide size survives verbatim')
  const presentationXml = pkg.xml('ppt/presentation.xml')
  assert(pkg.xml('ppt/presentation.xml') === presentationXml, 'xml() returns a cached tree')
  assertEqual(presentationXml.presentation.sldSz['@_cx'], '12196763', 'attributes stay strings')
  const themeBuffer = pkg.buffer('ppt/theme/theme1.xml')
  assert(Buffer.isBuffer(themeBuffer) && themeBuffer.length > 1000, 'buffer() returns bytes')
  await assertCode(() => pkg.text('ppt/nope.xml'), 'part_not_found', 'missing part throws part_not_found')
  return { bytes: pkg.byteLength, cx: presentationXml.presentation.sldSz['@_cx'] }
})

await test('relationships: presentation → theme1, slide1 → ../slideLayouts resolved', async () => {
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const presRels = pkg.relationships('ppt/presentation.xml')
  const themeRel = presRels.find((rel) => rel.type.endsWith('/theme'))
  assert(themeRel !== undefined, 'presentation rels contain a theme relationship')
  assertEqual(themeRel.resolvedPart, 'ppt/theme/theme1.xml', 'theme relationship resolves to theme1.xml')
  assertEqual(themeRel.mode, 'Internal', 'theme relationship is internal')
  const masters = presRels.filter((rel) => rel.type.endsWith('/slideMaster'))
  assertEqual(masters.length, 4, 'four slide masters are related from presentation.xml')

  const slideRels = pkg.relationships('ppt/slides/slide1.xml')
  const layout = slideRels.find((rel) => rel.type.endsWith('/slideLayout'))
  assertEqual(layout.resolvedPart, 'ppt/slideLayouts/slideLayout1.xml', '../ target normalized')
  assertEqual(pkg.relationships('ppt/notesMasters/notesMaster1.xml').length > 0, true, 'notes master has rels')
  return { presRels: presRels.length, slideRels: slideRels.map((rel) => rel.resolvedPart) }
})

await test('relationships: external targets and odd targets are handled', async () => {
  const rels = [
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://x/hyperlink" Target="https://example.com/a" TargetMode="External"/>',
    '<Relationship Id="rId2" Type="http://x/image" Target="/ppt/media/image%201.png"/>',
    '<Relationship Id="rId3" Type="http://x/theme" Target="theme/theme1.xml"/>',
    '<Relationship Id="rId4" Type="http://x/theme" Target=""/>',
    '<Relationship Id="rId5" Type="http://x/image" Target="../ppt/media/image2.png"/>',
    '</Relationships>',
  ].join('')
  const bytes = await syntheticPackage({ rels })
  const pkg = await loadPackage({ bytes })
  const resolved = pkg.relationships('ppt/presentation.xml')
  const external = resolved.find((rel) => rel.id === 'rId1')
  assertEqual(external.mode, 'External', 'TargetMode is preserved')
  assertEqual(external.resolvedPart, null, 'external target is not resolved')
  assertEqual(resolved.find((rel) => rel.id === 'rId2').resolvedPart, 'ppt/media/image 1.png', 'absolute + percent-decoded target')
  assertEqual(resolved.find((rel) => rel.id === 'rId3').resolvedPart, 'ppt/theme/theme1.xml', 'relative target resolved against the source dir')
  assertEqual(resolved.find((rel) => rel.id === 'rId4').resolvedPart, null, 'empty target ignored')
  assertEqual(
    resolved.find((rel) => rel.id === 'rId5').resolvedPart,
    'ppt/media/image2.png',
    'parent traversal resolved (ppt/../ppt/media)',
  )
  assertEqual(normalizePart('/a/./b/../c.xml'), 'a/c.xml', 'normalizePart collapses . and ..')
  assertEqual(pkg.relationships('ppt/slides/slide9.xml').length, 0, 'part without rels → empty array')
  return `${resolved.length} relationships`
})

await test('loadPackage: rejects non-zip, non-pptx and missing input with typed errors', async () => {
  await assertCode(() => loadPackage(Buffer.from('definitely not a zip')), 'not_a_pptx', 'garbage buffer')
  const zip = new JSZip()
  zip.file('hello.txt', 'not a presentation')
  const notPptx = await zip.generateAsync({ type: 'nodebuffer' })
  await assertCode(() => loadPackage({ bytes: notPptx }), 'not_a_pptx', 'zip without ppt/presentation.xml')
  await assertCode(() => loadPackage({ path: join(TMP_DIR, 'does-not-exist.pptx') }), 'template_not_found', 'missing file')
  await assertCode(() => loadPackage({}), 'missing_input', 'no path/bytes')
  // An empty directory masquerading as a file, and a real zip truncated mid-way.
  const real = readFileSync(TEMPLATE_PATH)
  await assertCode(() => loadPackage(real.subarray(0, 4096)), 'not_a_pptx', 'truncated zip')
  return 'not_a_pptx / template_not_found / missing_input all reached'
})

await test('parseXml/asArray: values stay verbatim strings (sysClr lastClr trap)', () => {
  assertEqual(XML_OPTIONS.parseTagValue, false, 'parseTagValue disabled')
  assertEqual(XML_OPTIONS.parseAttributeValue, false, 'parseAttributeValue disabled')
  assertEqual(XML_OPTIONS.removeNSPrefix, true, 'removeNSPrefix enabled')
  const doc = parseXml(
    '<a:theme xmlns:a="x"><a:clrScheme name="2210"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
      '<a:lt1><a:srgbClr val="1D1D1A"/></a:lt1></a:clrScheme>' +
      '<a:one/><a:many>1</a:many><a:many>2</a:many></a:theme>',
  )
  assertEqual(doc.theme.clrScheme.dk1.sysClr['@_lastClr'], '000000', 'lastClr keeps leading zeros as a string')
  assertEqual(doc.theme.clrScheme.dk1.sysClr['@_val'], 'windowText', 'system colour hint survives')
  assertEqual(doc.theme.clrScheme.lt1.srgbClr['@_val'], '1D1D1A', 'srgbClr val survives verbatim')
  assertEqual(asArray(doc.theme.one).length, 1, 'single child becomes a one-element array')
  assertEqual(asArray(doc.theme.many).length, 2, 'repeated child passes through')
  assertEqual(asArray(undefined), [], 'undefined → []')
  assertEqual(parseXml('').constructor, Object, 'empty input → empty object')
  return 'verbatim values preserved'
})

run('engine-read')
