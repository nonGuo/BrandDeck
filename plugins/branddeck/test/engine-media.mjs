/**
 * media.js: pixel-size sniffing, media inventory with referencing relationships, and
 * brand-logo discovery (including the "referenced by two masters" trap).
 *
 * Run: node test/engine-media.mjs
 */
import { existsSync, readFileSync } from 'node:fs'
import JSZip from 'jszip'
import {
  EXPECTED_TOKENS_PATH,
  TEMPLATE_PATH,
  assert,
  assertEqual,
  diffTable,
  run,
  test,
} from './engine-helpers.mjs'
import { loadPackage } from '../src/pptx/read.js'
import { findBrandLogo, imageSize, listMedia } from '../src/pptx/media.js'

/**
 * Minimal PNG header (signature + IHDR) — enough for `imageSize`, which must read the
 * dimensions out of the file header and not from the part name.
 * @param {number} width
 * @param {number} height
 * @returns {Buffer}
 */
function pngHeader(width, height) {
  const buffer = Buffer.alloc(33)
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]).copy(buffer, 0)
  buffer.writeUInt32BE(13, 8)
  buffer.write('IHDR', 12, 'latin1')
  buffer.writeUInt32BE(width, 16)
  buffer.writeUInt32BE(height, 20)
  buffer[24] = 8 // bit depth
  buffer[25] = 6 // RGBA
  return buffer
}

/**
 * Minimal JPEG with a SOF0 frame (dimensions inside the marker, not the file name).
 * @param {number} width
 * @param {number} height
 * @returns {Buffer}
 */
function jpegHeader(width, height) {
  const buffer = Buffer.alloc(30)
  buffer[0] = 0xff
  buffer[1] = 0xd8 // SOI
  buffer[2] = 0xff
  buffer[3] = 0xc0 // SOF0
  buffer.writeUInt16BE(17, 4) // segment length
  buffer[6] = 8 // precision
  buffer.writeUInt16BE(height, 7)
  buffer.writeUInt16BE(width, 9)
  buffer[11] = 3 // components
  return buffer
}

/**
 * Build a synthetic package with a fully controlled master/layout/slide → media
 * relationship graph.
 * @param {object} input
 * @param {Record<string, Buffer>} input.media            part name → bytes
 * @param {Record<string, string>} input.masterRefs        master part → media part
 * @param {Record<string, string[]>} [input.layoutRefs]    layout part → media parts
 * @param {Record<string, string[]>} [input.slideRefs]     slide part → media parts
 * @param {boolean} [input.picOnMaster]                    place the picture in the master shape tree
 * @returns {Promise<Buffer>}
 */
async function syntheticMediaPackage(input) {
  const zip = new JSZip()
  const masterParts = Object.keys(input.masterRefs)
  const layoutParts = Object.keys(input.layoutRefs ?? {})
  const slideParts = Object.keys(input.slideRefs ?? {})

  zip.file(
    'ppt/presentation.xml',
    '<?xml version="1.0"?><p:presentation xmlns:p="p" xmlns:r="r">' +
      '<p:sldMasterIdLst>' +
      masterParts.map((_, index) => `<p:sldMasterId id="${2147483648 + index}" r:id="rId${index + 1}"/>`).join('') +
      '</p:sldMasterIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
  )

  const presentationRels = [
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ...masterParts.map(
      (part, index) =>
        `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="${part.replace('ppt/', '')}"/>`,
    ),
    `<Relationship Id="rId${masterParts.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`,
    '</Relationships>',
  ].join('')
  zip.file('ppt/_rels/presentation.xml.rels', presentationRels)

  zip.file(
    'ppt/theme/theme1.xml',
    '<?xml version="1.0"?><a:theme xmlns:a="a" name="T"><a:themeElements><a:clrScheme name="T">' +
      '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="333333"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="C7000A"/></a:accent1><a:accent2><a:srgbClr val="E9002F"/></a:accent2>' +
      '</a:clrScheme></a:themeElements></a:theme>',
  )

  masterParts.forEach((masterPart) => {
    const mediaPart = input.masterRefs[masterPart]
    const pic =
      input.picOnMaster === true
        ? '<p:pic><p:nvPicPr><p:cNvPr id="10" name="logo"/><p:cNvPicPr/><p:nvPr userDrawn="1"/></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rId1"/></p:blipFill>' +
          '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="457200"/></a:xfrm>' +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        : ''
    zip.file(
      masterPart,
      `<?xml version="1.0"?><p:sldMaster xmlns:a="a" xmlns:p="p" xmlns:r="r"><p:cSld><p:spTree>${pic}</p:spTree></p:cSld></p:sldMaster>`,
    )
    zip.file(
      `ppt/slideMasters/_rels/${masterPart.split('/').pop()}.rels`,
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaPart.split('/').pop()}"/>` +
        '</Relationships>',
    )
  })

  for (const [layoutPart, mediaParts] of Object.entries(input.layoutRefs ?? {})) {
    zip.file(
      layoutPart,
      '<?xml version="1.0"?><p:sldLayout xmlns:a="a" xmlns:p="p" xmlns:r="r"><p:cSld><p:spTree/></p:cSld></p:sldLayout>',
    )
    zip.file(
      `ppt/slideLayouts/_rels/${layoutPart.split('/').pop()}.rels`,
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        mediaParts
          .map(
            (mediaPart, index) =>
              `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaPart.split('/').pop()}"/>`,
          )
          .join('') +
        '</Relationships>',
    )
  }

  for (const [slidePart, mediaParts] of Object.entries(input.slideRefs ?? {})) {
    zip.file(
      slidePart,
      '<?xml version="1.0"?><p:sld xmlns:a="a" xmlns:p="p" xmlns:r="r"><p:cSld><p:spTree/></p:cSld></p:sld>',
    )
    zip.file(
      `ppt/slides/_rels/${slidePart.split('/').pop()}.rels`,
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        mediaParts
          .map(
            (mediaPart, index) =>
              `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaPart.split('/').pop()}"/>`,
          )
          .join('') +
        '</Relationships>',
    )
  }

  for (const [part, bytes] of Object.entries(input.media)) zip.file(part, bytes)
  return zip.generateAsync({ type: 'nodebuffer' })
}

await test('imageSize: reads PNG IHDR / JPEG SOFn / GIF / BMP from the header', () => {
  assertEqual(imageSize(pngHeader(529, 116), 'ppt/media/image1.png'), {
    format: 'png',
    contentType: 'image/png',
    widthPx: 529,
    heightPx: 116,
  }, 'PNG dimensions come from IHDR')

  const jpeg = jpegHeader(1949, 917)
  assertEqual(imageSize(jpeg, 'ppt/media/image3.jpeg'), {
    format: 'jpeg',
    contentType: 'image/jpeg',
    widthPx: 1949,
    heightPx: 917,
  }, 'JPEG dimensions come from SOF0')

  const gif = Buffer.from('4749463839610a0020000000', 'hex') // GIF89a, 10x32
  const gifSize = imageSize(gif, 'x.gif')
  assertEqual(gifSize.format, 'gif', 'GIF magic detected')
  assertEqual([gifSize.widthPx, gifSize.heightPx], [10, 32], 'GIF dimensions are little-endian')

  const bmp = Buffer.alloc(26)
  bmp.write('BM', 0, 'latin1')
  bmp.writeInt32LE(64, 18)
  bmp.writeInt32LE(-48, 22)
  const bmpSize = imageSize(bmp, 'x.bmp')
  assertEqual([bmpSize.format, bmpSize.widthPx, bmpSize.heightPx], ['bmp', 64, 48], 'BMP dimensions (abs height)')
  return 'png/jpeg/gif/bmp'
})

await test('imageSize: EMF/WMF report no pixel size and never throw', () => {
  const emf = Buffer.alloc(64)
  emf.writeUInt32LE(1, 0)
  emf.write(' EMF', 40, 'latin1')
  assertEqual(imageSize(emf, 'ppt/media/image7.emf'), {
    format: 'emf',
    contentType: 'image/x-emf',
    widthPx: null,
    heightPx: null,
  }, 'EMF has no pixel size')

  const wmf = Buffer.from([0xd7, 0xcd, 0xc6, 0x9a, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
  assertEqual(imageSize(wmf, 'ppt/media/image8.wmf').format, 'wmf', 'WMF magic detected')

  for (const garbage of [Buffer.from('hello world, not an image'), Buffer.alloc(0), Buffer.alloc(12), null, undefined]) {
    const size = imageSize(garbage, 'ppt/media/x.bin')
    assertEqual(size.widthPx, null, `garbage input ${String(garbage?.length)} yields null width`)
  }
  assertEqual(imageSize(Buffer.from('hello world, not an image'), 'ppt/media/x.png').contentType, 'image/png', 'extension fallback for the content type')
  return 'emf/wmf/garbage safe'
})

await test('listMedia: six parts, pixel sizes and the relationships that reference them', async () => {
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const media = listMedia(pkg)
  assertEqual(media.length, 6, 'six media parts')
  assertEqual(media.map((entry) => entry.path), [
    'ppt/media/image1.png',
    'ppt/media/image2.png',
    'ppt/media/image3.jpeg',
    'ppt/media/image4.jpg',
    'ppt/media/image5.jpeg',
    'ppt/media/image6.png',
  ], 'media parts in package order')

  const logoPart = media.find((entry) => entry.path === 'ppt/media/image1.png')
  assertEqual(logoPart.bytes, 22308, 'image1 byte length')
  assertEqual([logoPart.widthPx, logoPart.heightPx], [529, 116], 'image1 pixel size')
  assertEqual(logoPart.contentType, 'image/png', 'image1 content type')
  assertEqual(logoPart.referencedBy, [
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    'ppt/slideMasters/_rels/slideMaster4.xml.rels',
  ], 'image1 is referenced by TWO masters')
  assert(Buffer.isBuffer(logoPart.buffer) && logoPart.buffer.length === 22308, 'bytes are available')

  const photo = media.find((entry) => entry.path === 'ppt/media/image2.png')
  assertEqual([photo.widthPx, photo.heightPx], [1949, 917], 'image2 pixel size')
  assertEqual(photo.referencedBy, ['ppt/slideLayouts/_rels/slideLayout1.xml.rels'], 'image2 comes from a layout')
  for (const entry of media) {
    assert(entry.refs.every((ref) => ref.sourceKind === 'master' || ref.sourceKind === 'layout'), `${entry.path} is referenced only from masters/layouts`)
  }
  return media.map((entry) => `${entry.path} ${entry.widthPx}x${entry.heightPx}`)
})

await test('findBrandLogo: image1.png bottom-right frame, cross-checked with expected-tokens', async () => {
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const logo = findBrandLogo(pkg)
  assert(logo !== null, 'a logo was found')
  assertEqual(logo.partPath, 'ppt/media/image1.png', 'canonical logo part')
  assertEqual(logo.bytes, 22308, 'logo byte length')
  assertEqual([logo.widthPx, logo.heightPx], [529, 116], 'logo pixel size')
  assertEqual(logo.masterPart, 'ppt/slideMasters/slideMaster1.xml', 'frame read from the primary master')
  assertEqual(logo.masterParts, ['ppt/slideMasters/slideMaster1.xml', 'ppt/slideMasters/slideMaster4.xml'], 'logo is placed by two masters')
  assertEqual(logo.frameInches, { x: 10.0686, y: 6.5356, w: 2.4698, h: 0.5277 }, 'master1 frame in inches (bottom-right)')
  assert(Math.abs(logo.aspect - 4.5603) < 0.001, `aspect ≈ 4.5603, got ${logo.aspect}`)
  assert(Buffer.isBuffer(logo.buffer), 'logo bytes returned')

  if (existsSync(EXPECTED_TOKENS_PATH)) {
    const expected = JSON.parse(readFileSync(EXPECTED_TOKENS_PATH, 'utf8')).logo
    const mismatches = diffTable([
      { field: 'logo.partPath', expected: expected.partPath, actual: logo.partPath },
      { field: 'logo.bytes', expected: expected.bytes, actual: logo.bytes },
      { field: 'logo.widthPx', expected: expected.widthPx, actual: logo.widthPx },
      { field: 'logo.heightPx', expected: expected.heightPx, actual: logo.heightPx },
      { field: 'logo.frameInches', expected: expected.frameInches, actual: logo.frameInches },
    ])
    assertEqual(mismatches.length, 0, `logo fields match expected-tokens.json: ${mismatches.map((row) => row.field).join(', ')}`)
  } else {
    console.log('  NOTE research/expected-tokens.json absent — engine dump is authoritative')
  }
  return `${logo.partPath} ${logo.widthPx}x${logo.heightPx} @ (${logo.frameInches.x},${logo.frameInches.y})`
})

await test('findBrandLogo: master-only rule + pic frame, and a layout-only package yields no logo', async () => {
  // image1 is referenced from a master AND a layout → not a candidate for the
  // "master-only" rule; image2 is master-only → the logo.
  const bytes = await syntheticMediaPackage({
    media: {
      'ppt/media/image1.png': pngHeader(400, 100),
      'ppt/media/image2.png': pngHeader(320, 80),
    },
    masterRefs: { 'ppt/slideMasters/slideMaster1.xml': 'ppt/media/image1.png' },
    layoutRefs: { 'ppt/slideLayouts/slideLayout1.xml': ['ppt/media/image1.png'] },
    picOnMaster: true,
  })
  // Add image2 as the second master's placed picture so a frame exists for it.
  const zip = await JSZip.loadAsync(bytes)
  zip.file(
    'ppt/presentation.xml',
    '<?xml version="1.0"?><p:presentation xmlns:p="p" xmlns:r="r">' +
      '<p:sldMasterIdLst><p:sldMasterId id="1" r:id="rId1"/><p:sldMasterId id="2" r:id="rId2"/></p:sldMasterIdLst>' +
      '<p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/slideMasters/slideMaster2.xml',
    '<?xml version="1.0"?><p:sldMaster xmlns:a="a" xmlns:p="p" xmlns:r="r"><p:cSld><p:spTree>' +
      '<p:pic><p:nvPicPr><p:cNvPr id="20" name="logo2"/><p:cNvPicPr/><p:nvPr userDrawn="1"/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId1"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="457200" y="2743200"/><a:ext cx="1828800" cy="457200"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
      '</p:spTree></p:cSld></p:sldMaster>',
  )
  zip.file(
    'ppt/slideMasters/_rels/slideMaster2.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>' +
      '</Relationships>',
  )
  const pkg = await loadPackage({ bytes: await zip.generateAsync({ type: 'nodebuffer' }) })
  const logo = findBrandLogo(pkg)
  assert(logo !== null, 'a logo was found')
  assertEqual(logo.partPath, 'ppt/media/image2.png', 'only the master-only part is a candidate')
  assertEqual(logo.frameInches, { x: 0.5, y: 3, w: 2, h: 0.5 }, 'frame geometry recovered from the master <a:xfrm>')

  // A package where every image sits on a layout: no master reference at all → null.
  const layoutOnly = await syntheticMediaPackage({
    media: { 'ppt/media/image1.png': pngHeader(120, 40) },
    masterRefs: {},
    layoutRefs: { 'ppt/slideLayouts/slideLayout1.xml': ['ppt/media/image1.png'] },
  })
  const layoutPkg = await loadPackage({ bytes: layoutOnly })
  assertEqual(findBrandLogo(layoutPkg), null, 'no master-referenced media → null')
  return logo.frameInches
})

run('engine-media')
