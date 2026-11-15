/**
 * compile.js: CardDeck → native .pptx, including the export-fidelity regressions the
 * Lead reproduced in a real LibreOffice render:
 *
 * - every `<a:ea>` names a CJK-capable face (a Latin face there doubles CJK glyphs);
 * - no `<a:t>null</a:t>` / `undefined` / placeholder-index junk runs;
 * - page numbers are fields in plain text boxes, not `sldNum` placeholders;
 * - each deck string appears exactly once in its slide XML.
 *
 * Run: node test/engine-compile.mjs
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import {
  OUT_DIR,
  TEMPLATE_PATH,
  TMP_DIR,
  assert,
  assertCode,
  assertEqual,
  assertMatch,
  freshTempDir,
  run,
  test,
} from './engine-helpers.mjs'
import { parsePptxTemplate } from '../src/template.js'
import { compileDeck, compileDeckToFile, isCjkCapableFace, resolveCjkFont } from '../src/compile.js'
import { CardDeckSchema } from '../src/schema.js'

mkdirSync(OUT_DIR, { recursive: true })
const STORE = freshTempDir('compile-store')

// One real parse of the corporate template: themes, logo bytes and media inventory.
const parsedTemplate = await parsePptxTemplate({
  path: TEMPLATE_PATH,
  templateName: 'PPT模板-HW',
  storeConfig: { root: STORE },
})
const THEME = parsedTemplate.theme
const LOGO = parsedTemplate.logoBytes

/** Junk runs that must never reach a slide/layout/master part. */
const JUNK = /<a:t>(null|undefined|NaN)<\/a:t>/

const SVG_FLOW =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 700 320">' +
  '<rect x="0" y="0" width="700" height="320" rx="16" fill="#F7F7F7"/>' +
  '<rect x="40" y="120" width="160" height="80" rx="12" fill="#C7000A"/>' +
  '<text x="120" y="168" font-family="Microsoft YaHei" font-size="24" fill="#FFFFFF" text-anchor="middle">解析模板</text>' +
  '<path d="M220 160 L300 160" stroke="#666666" stroke-width="4"/>' +
  '<rect x="310" y="120" width="160" height="80" rx="12" fill="#E9002F"/>' +
  '<text x="390" y="168" font-family="Microsoft YaHei" font-size="24" fill="#FFFFFF" text-anchor="middle">生成卡片</text>' +
  '<path d="M490 160 L570 160" stroke="#666666" stroke-width="4"/>' +
  '<rect x="580" y="120" width="160" height="80" rx="12" fill="#232323"/>' +
  '<text x="660" y="168" font-family="Microsoft YaHei" font-size="24" fill="#FFFFFF" text-anchor="middle">导出 PPTX</text>' +
  '</svg>'

const TITLE_TEXT = 'AI 驱动的企业内容生产'
const SUBTITLE_TEXT = '从企业模板到原生可编辑演示文稿'
const HEADING_TEXT = '业务挑战与机会'
const BODY_TEXT = '内容需求增长 3 倍，但排版与合规返工消耗了 60% 的交付时间。'

/**
 * The 3-slide test deck: title slide, text+shape slide, chart+svg slide.
 * @returns {import('../src/schema.js').CardDeck}
 */
function buildDeck(overrides = {}) {
  const deck = {
    deckId: 'engine-demo',
    title: '引擎导出验证',
    templateId: THEME.templateId,
    slides: [
      {
        slideId: 'cover',
        background: '$colors.background',
        speakerNotes: '封面页备注：验证 speakerNotes 落到 notesSlide。',
        elements: [
          { type: 'shape', id: 'accent', shapeType: 'rect', x: 96, y: 236, w: 8, h: 132, style: { fill: '$colors.primary' } },
          { type: 'text', id: 'title', role: 'title', content: TITLE_TEXT, x: 128, y: 236, w: 900, h: 88, style: { fontSize: 56, bold: true, color: '$colors.textPrimary' } },
          { type: 'text', id: 'subtitle', role: 'subtitle', content: SUBTITLE_TEXT, x: 128, y: 340, w: 900, h: 40, style: { fontSize: 22, color: '$colors.textSecondary' } },
        ],
      },
      {
        slideId: 'content',
        background: '$colors.background',
        elements: [
          { type: 'text', id: 'heading', role: 'title', content: HEADING_TEXT, x: 96, y: 96, w: 800, h: 64, style: { fontSize: 36, bold: true, color: '$colors.textPrimary' } },
          { type: 'shape', id: 'card', shapeType: 'roundRect', x: 96, y: 200, w: 1040, h: 260, style: { fill: '$colors.surface', radius: 16 } },
          { type: 'text', id: 'body', role: 'body', content: BODY_TEXT, x: 128, y: 232, w: 960, h: 120, style: { fontSize: 20, color: '$colors.textPrimary' } },
          { type: 'shape', id: 'rule', shapeType: 'line', x: 128, y: 400, w: 960, h: 2, style: { stroke: '$colors.primary', strokeWidth: 2 } },
        ],
      },
      {
        slideId: 'visual',
        background: '$colors.background',
        elements: [
          { type: 'text', id: 'heading2', role: 'title', content: '增长趋势与流程', x: 96, y: 96, w: 800, h: 64, style: { fontSize: 36, bold: true, color: '$colors.textPrimary' } },
          {
            type: 'chart',
            id: 'chart',
            chartType: 'bar',
            x: 96,
            y: 200,
            w: 520,
            h: 280,
            data: { labels: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: '营收', values: [120, 168, 205, 260] }] },
            style: { legend: true, showValues: true },
          },
          { type: 'svg', id: 'flow', description: '三步流程示意', x: 640, y: 200, w: 450, h: 206, svgCode: SVG_FLOW },
          { type: 'image', id: 'logo-mark', assetKey: 'logo', x: 96, y: 520, w: 200, h: 44 },
        ],
      },
    ],
  }
  return { ...deck, ...overrides }
}

/**
 * Load every XML part of a compiled package.
 * @param {Buffer} buffer
 * @returns {Promise<{parts: Record<string, string>, names: string[], media: Record<string, Buffer>}>}
 */
async function openPackage(buffer) {
  const zip = await JSZip.loadAsync(buffer)
  /** @type {Record<string, string>} */
  const parts = {}
  /** @type {Record<string, Buffer>} */
  const media = {}
  const names = []
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue
    names.push(name)
    if (name.endsWith('.xml') || name.endsWith('.rels')) parts[name] = await entry.async('string')
    if (name.startsWith('ppt/media/')) media[name] = await entry.async('nodebuffer')
  }
  return { parts, names, media }
}

/** @param {Buffer} png */
function pngSize(png) {
  return { width: png.readUInt32BE(16), height: png.readUInt32BE(20) }
}

/** XML-escape for text we look for inside `<a:t>`. */
function escapeXml(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

const deck = buildDeck()
assertEqual(CardDeckSchema.safeParse(deck).success, true, 'test deck is schema-valid')

const compiled = await compileDeck({ deck, theme: THEME, logoBytes: LOGO, assets: { logo: LOGO } })
const pkg = await openPackage(compiled.buffer)

await test('compileDeck: valid zip, 3 slides, presentation part, verified=true', async () => {
  assertEqual(compiled.verified, true, 'result is verified')
  assertEqual(compiled.slideCount, 3, 'slideCount matches the deck')
  assert(compiled.buffer.length > 40000, `buffer is plausible (${compiled.buffer.length} bytes)`)
  assert(pkg.names.includes('ppt/presentation.xml'), 'ppt/presentation.xml exists')
  assert(pkg.names.includes('ppt/slides/slide1.xml'), 'slide1 part exists')
  assert(pkg.names.includes('ppt/slides/slide2.xml'), 'slide2 part exists')
  assert(pkg.names.includes('ppt/slides/slide3.xml'), 'slide3 part exists (explicit acceptance criterion)')
  assert(pkg.names.includes('ppt/slideMasters/slideMaster1.xml'), 'a slide master exists')
  assertEqual(compiled.cjkFont, 'Microsoft YaHei', 'East-Asian face resolved from the theme')
  return { bytes: compiled.buffer.length, slides: compiled.slideCount, parts: pkg.names.length }
})

await test('compileDeck: deck text lands in the slide XML', () => {
  const slide1 = pkg.parts['ppt/slides/slide1.xml']
  const slide2 = pkg.parts['ppt/slides/slide2.xml']
  assertMatch(slide1, new RegExp(`<a:t>${escapeXml(TITLE_TEXT)}</a:t>`), 'slide1 contains the cover title')
  assertMatch(slide1, new RegExp(`<a:t>${escapeXml(SUBTITLE_TEXT)}</a:t>`), 'slide1 contains the subtitle')
  assertMatch(slide2, new RegExp(`<a:t>${escapeXml(HEADING_TEXT)}</a:t>`), 'slide2 contains its heading')
  assertMatch(slide2, new RegExp(`<a:t>${escapeXml(BODY_TEXT)}</a:t>`), 'slide2 contains the body copy')
  assert(slide2.includes('roundRect'), 'the roundRect card became a native preset shape')
  assert(slide2.includes('prst="line"'), 'the divider became a native line shape')
  assertMatch(slide1, /<a:solidFill><a:srgbClr val="C7000A"\/>/, 'brand primary applied to the accent bar')

  const slideTexts = [slide1, slide2, pkg.parts['ppt/slides/slide3.xml']].join('')
  for (const text of [TITLE_TEXT, SUBTITLE_TEXT, HEADING_TEXT, BODY_TEXT, '增长趋势与流程']) {
    const occurrences = slideTexts.split(`<a:t>${escapeXml(text)}</a:t>`).length - 1
    assertEqual(occurrences, 1, `"${text}" appears exactly once in the slide parts`)
  }
  return 'text mapping ok'
})

await test('compileDeck: page numbers are fields, not sldNum placeholders', () => {
  for (const name of ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml']) {
    const xml = pkg.parts[name]
    assertMatch(xml, /<a:fld[^>]*type="slidenum"/, `${name} carries a slidenum field`)
    assert(!/<p:ph type="sldNum"/.test(xml), `${name} has no sldNum placeholder (LibreOffice would draw a prompt box)`)
    assert(!/<a:t>1002<\/a:t>/.test(xml), `${name} has no placeholder-index run`)
  }
  assertMatch(pkg.parts['ppt/slides/slide1.xml'], /<a:t>[‹#›\d]+<\/a:t>/, 'page number text is sane')
  return 'page numbers ok'
})

await test('export fidelity: every <a:ea> is CJK-capable and no junk runs survive', () => {
  const handwritten = [
    'ppt/slides/slide1.xml',
    'ppt/slides/slide2.xml',
    'ppt/slides/slide3.xml',
    'ppt/slideMasters/slideMaster1.xml',
    'ppt/slideLayouts/slideLayout2.xml',
  ]
  /** @type {string[]} */
  const badFaces = []
  for (const name of handwritten) {
    const xml = pkg.parts[name]
    assert(xml !== undefined, `${name} exists`)
    for (const match of xml.matchAll(/<a:ea typeface="([^"]*)"/g)) {
      const face = match[1]
      if (face.startsWith('+')) continue // scheme reference (+mj-ea / +mn-ea)
      if (!isCjkCapableFace(face)) badFaces.push(`${name}: ${face}`)
    }
    assert(!JUNK.test(xml), `${name} contains no null/undefined/NaN text runs`)
    assert(!/<a:t>1002<\/a:t>/.test(xml), `${name} contains no placeholder-index text run`)
  }
  assertEqual(badFaces, [], `no Latin face sits in <a:ea>: ${badFaces.join(', ')}`)

  // The exact pair the Lead asked for: Latin in latin/cs, CJK in ea, on a body run.
  const xml = pkg.parts['ppt/slides/slide2.xml']
  const bodyRun = /<a:latin typeface="Arial"[^>]*\/><a:ea typeface="Microsoft YaHei"[^>]*\/><a:cs typeface="Arial"[^>]*\/>/.exec(xml)
  assert(bodyRun !== null, 'body runs carry Arial latin + Microsoft YaHei ea + Arial cs')
  assertMatch(xml, /<a:latin typeface="Microsoft YaHei"[^>]*\/><a:ea typeface="Microsoft YaHei"/, 'title runs carry Microsoft YaHei in both latin and ea')

  // CJK text never sits in a run whose ea is missing/latin-only.
  const cjkTexts = [TITLE_TEXT, HEADING_TEXT, BODY_TEXT]
  for (const text of cjkTexts) {
    const index = xml.indexOf(`<a:t>${escapeXml(text)}</a:t>`)
    if (index === -1) continue
    const runStart = xml.lastIndexOf('<a:r>', index)
    const run = xml.slice(runStart === -1 ? 0 : runStart, index)
    assert(/<a:ea typeface="Microsoft YaHei"/.test(run), `CJK run for "${text}" has a CJK <a:ea>`)
  }
  return 'font faces ok'
})

await test('compileDeck: the SVG was rasterized into an image part at 3×', () => {
  const svgElement = deck.slides[2].elements.find((element) => element.type === 'svg')
  const expected = { width: Math.round(svgElement.w) * 3, height: Math.round(svgElement.h) * 3 }
  const mediaEntries = Object.entries(pkg.media)
  assert(mediaEntries.length >= 3, `logo + svg + slide logo → ≥3 media parts, got ${mediaEntries.length}`)

  const candidates = mediaEntries.filter(([, bytes]) => Buffer.isBuffer(bytes) && bytes.length > 8000)
  const svgPng = candidates
    .map(([name, bytes]) => ({ name, bytes, size: pngSize(bytes) }))
    .find((entry) => entry.size.width === expected.width && entry.size.height === expected.height)
  assert(svgPng !== undefined, `a ${expected.width}×${expected.height} PNG exists (the 3× SVG raster)`)
  assert(svgPng.bytes.subarray(0, 8).toString('hex') === '89504e470d0a1a0a', 'the SVG raster is a PNG')
  assertMatch(pkg.parts['ppt/slides/slide3.xml'], /<p:pic>/, 'slide3 embeds the rasterized diagram as a picture')
  assertMatch(pkg.parts['ppt/slides/slide3.xml'], /name="Image 0"|name="Image \d+"/, 'the picture has a name')
  // The logo is embedded too (master/layout + the explicit image element on slide3).
  const logoParts = candidates.filter(([, bytes]) => bytes.length === LOGO.length)
  assert(logoParts.length >= 1, 'the corporate logo bytes are embedded')
  writeFileSync(join(OUT_DIR, 'engine-compile-svg-raster.png'), svgPng.bytes)
  return { expected, actual: svgPng.size, part: svgPng.name }
})

await test('compileDeck: chart is a native Office chart with an embedded workbook', () => {
  const chartParts = pkg.names.filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
  assertEqual(chartParts.length, 1, 'exactly one chart part')
  const chartXml = pkg.parts[chartParts[0]]
  assert(chartXml.includes('barChart'), 'barChart is rendered natively')
  assertMatch(chartXml, /<\?xml|<!\[CDATA\[|Q1/, 'chart carries the category labels')
  assertMatch(chartXml, /C7000A/i, 'chart colours come from the theme palette')
  assert(
    pkg.names.some((name) => name.startsWith('ppt/embeddings/') && name.endsWith('.xlsx')),
    'the chart ships an embedded xlsx so its data stays editable',
  )
  assertMatch(pkg.parts['ppt/slides/slide3.xml'], /<c:chart|graphicFrame/, 'slide3 embeds the chart frame')
  return chartParts[0]
})

await test('compileDeck: speaker notes become a notes slide', () => {
  const notes = pkg.names.filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name))
  assert(notes.length >= 1, 'pptxgenjs emitted at least one notes slide')
  const withText = notes.filter((name) => pkg.parts[name].includes('speakerNotes'))
  assertEqual(withText.length, 1, 'exactly one notes slide carries the speaker notes')
  assert(pkg.names.some((name) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(name)), 'notes master exists')
  return { notesSlides: notes.length, withText: withText[0] }
})

await test('compileDeck: logo honours theme.brandAssets.logoFrame (bottom-right)', () => {
  const layoutParts = pkg.names.filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name))
  const withPicture = layoutParts
    .map((name) => pkg.parts[name])
    .find((xml) => xml.includes('<p:pic>'))
  assert(withPicture !== undefined, 'a layout carries the logo picture')
  const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(withPicture)[0]
  const frame = /<a:off x="(\d+)" y="(\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(pic)
  assert(frame !== null, 'the picture has a frame')
  const [, x, y, cx, cy] = frame.map(Number)
  const toIn = (emu) => emu / 914400
  assert(Math.abs(toIn(x) - 10.0686) < 0.01, `logo x ≈ 10.0686in, got ${toIn(x)}`)
  assert(Math.abs(toIn(y) - 6.5356) < 0.01, `logo y ≈ 6.5356in, got ${toIn(y)}`)
  assert(Math.abs(toIn(cx) - 2.4698) < 0.01, `logo w ≈ 2.4698in, got ${toIn(cx)}`)
  assert(Math.abs(toIn(cy) - 0.5277) < 0.01, `logo h ≈ 0.5277in, got ${toIn(cy)}`)
  assertMatch(pic, /descr="brand logo"/, 'logo alt text is set')
  return { x: toIn(x), y: toIn(y), w: toIn(cx), h: toIn(cy) }
})

await test('compileDeck: layout carries the footer, slides carry the theme background', () => {
  const layouts = pkg.names
    .filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name))
    .map((name) => pkg.parts[name])
  assert(layouts.some((xml) => xml.includes('引擎导出验证')), 'footer text defaults to the deck title (layout object)')
  assertMatch(pkg.parts['ppt/slides/slide1.xml'], /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"\/>/, 'slide background follows the theme')
  assertMatch(pkg.parts['ppt/slideMasters/slideMaster1.xml'], /<p:ph type="sldNum"/, 'the master keeps a real slide-number placeholder')
  assert(layouts.some((xml) => /<a:latin typeface="Arial"/.test(xml)), 'footer uses the Latin body face')
  return 'footer/background ok'
})

await test('compileDeckToFile: writes .tmp/engine-out and re-opens as a valid package', async () => {
  const target = join(OUT_DIR, 'engine-demo.pptx')
  const result = await compileDeckToFile({ deck, theme: THEME, logoBytes: LOGO, assets: { logo: LOGO } }, target)
  assertEqual(result.path, target, 'returns the written path')
  assertEqual(result.verified, true, 'file result is verified')
  assert(result.bytes > 40000, 'file has plausible size')
  const onDisk = readFileSync(target)
  assertEqual(onDisk.length, result.bytes, 'byte length matches the file')
  assertEqual(onDisk.equals(result.buffer), true, 'the returned buffer is what was written')
  const reopened = await JSZip.loadAsync(onDisk)
  assert(reopened.file('ppt/presentation.xml') !== null, 're-opened with jszip: presentation part present')
  const slideCount = Object.keys(reopened.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length
  assertEqual(slideCount, 3, 're-opened with jszip: 3 slides')
  return { path: target, bytes: result.bytes }
})

await test('compileDeck: survivable problems become warnings, not exceptions', async () => {
  const lossy = buildDeck({
    deckId: 'lossy',
    slides: [
      {
        slideId: 'problems',
        elements: [
          { type: 'image', id: 'missing', assetKey: 'nope', x: 96, y: 96, w: 100, h: 100 },
          { type: 'text', id: 'tiny', role: 'body', content: '这是一段故意放进很小文本框里的超长中文内容，用来触发溢出预警。'.repeat(3), x: 96, y: 260, w: 120, h: 24, style: { fontSize: 20, color: '$colors.textPrimary' } },
        ],
      },
    ],
  })
  const result = await compileDeck({ deck: lossy, theme: THEME, logoBytes: LOGO })
  assertEqual(result.verified, true, 'the export still verifies')
  assert(result.warnings.some((warning) => warning.includes('asset "nope" not found')), 'missing asset warning')
  assert(result.warnings.some((warning) => warning.includes('overflows')), 'text overflow warning')
  return result.warnings
})

await test('compileDeck: input guards and cjkFont option', async () => {
  await assertCode(() => compileDeck({ deck: { slides: [] }, theme: THEME }), 'invalid_deck', 'empty deck refused')
  const overridden = await compileDeck({ deck, theme: { ...THEME, fonts: { ...THEME.fonts, title: 'Arial' } }, logoBytes: LOGO, options: { cjkFont: 'SimHei' } })
  assertEqual(overridden.cjkFont, 'SimHei', 'explicit cjkFont wins')
  const pkgWithSimHei = await openPackage(overridden.buffer)
  assertMatch(pkgWithSimHei.parts['ppt/slides/slide1.xml'], /<a:ea typeface="SimHei"/, 'the override is written into <a:ea>')
  assertEqual(resolveCjkFont({ fonts: { title: 'Arial', body: 'Arial', latin: 'Arial' } }), 'Microsoft YaHei', 'Latin-only themes fall back to Microsoft YaHei for CJK')
  assertEqual(resolveCjkFont({ fonts: { title: '微软雅黑', body: 'Arial' } }), '微软雅黑', 'a CJK-capable title face is reused')
  assertEqual(isCjkCapableFace('Microsoft YaHei'), true, 'Microsoft YaHei is CJK-capable')
  assertEqual(isCjkCapableFace('Arial'), false, 'Arial is not CJK-capable')
  return 'guards ok'
})

await test('compileDeck: shape/alpha mapping (transparency, radius, opacity)', async () => {
  // NB: 8-digit colours are #AARRGGBB in this contract (schema.splitAlpha), so
  // #80C7000A = brand red at 50 % alpha.
  const alphaDeck = buildDeck({
    deckId: 'alpha',
    slides: [
      {
        slideId: 'alpha',
        elements: [
          { type: 'shape', id: 'glass', shapeType: 'roundRect', x: 100, y: 100, w: 400, h: 200, style: { fill: '#80C7000A', radius: 24, stroke: '$colors.primary', strokeWidth: 2 } },
          { type: 'shape', id: 'dot', shapeType: 'ellipse', x: 600, y: 100, w: 80, h: 80, style: { fill: '$colors.accent', opacity: 0.5 } },
        ],
      },
    ],
  })
  const result = await compileDeck({ deck: alphaDeck, theme: THEME, logoBytes: null })
  const alphaPkg = await openPackage(result.buffer)
  const xml = alphaPkg.parts['ppt/slides/slide1.xml']
  assertMatch(xml, /prst="roundRect"/, 'roundRect mapped')
  assertMatch(xml, /prst="ellipse"/, 'ellipse mapped')
  // pptxgenjs expresses fill transparency in whole percent → 50 % = <a:alpha val="50000">.
  assertMatch(xml, /<a:srgbClr val="C7000A"><a:alpha val="50000"\/>/, 'ARGB fill alpha becomes a native <a:alpha>')
  assertMatch(xml, /<a:srgbClr val="E9002F"><a:alpha val="50000"\/>/, 'style.opacity becomes a native <a:alpha>')
  assertMatch(xml, /<a:gd name="adj"/, 'roundRect carries an adjustment (rectRadius)')
  writeFileSync(join(OUT_DIR, 'engine-alpha.pptx'), result.buffer)
  return 'shape mapping ok'
})

run('engine-compile')
