/**
 * End-to-end: real corporate template → tokens → store → starter deck → CardDeck →
 * compiled .pptx → re-opened package.
 *
 * Every step uses the real `PPT模板-HW.pptx`; nothing is mocked.
 *
 * Run: node test/engine-e2e.mjs
 */
import { existsSync, mkdirSync, readFileSync, statSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import {
  EXPECTED_TOKENS_PATH,
  OUT_DIR,
  TEMPLATE_PATH,
  assert,
  assertEqual,
  assertMatch,
  diffTable,
  freshTempDir,
  run,
  test,
} from './engine-helpers.mjs'
import { parsePptxTemplate, namespaceTemplateId } from '../src/template.js'
import {
  createDeck,
  ensureStore,
  listDecks,
  listTemplates,
  readDeck,
  readTemplate,
  writeDeck,
  writeTemplate,
} from '../src/store.js'
import { compileDeckToFile } from '../src/compile.js'
import { CardDeckSchema } from '../src/schema.js'

mkdirSync(OUT_DIR, { recursive: true })
const STORE = freshTempDir('e2e-store')
const storeConfig = { root: STORE }

/** @type {any} */
let state = {}

const FLOW_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 720 300">' +
  '<rect width="720" height="300" rx="18" fill="#F7F7F7"/>' +
  '<rect x="40" y="110" width="180" height="80" rx="12" fill="#C7000A"/>' +
  '<text x="130" y="158" font-family="Microsoft YaHei" font-size="26" fill="#FFFFFF" text-anchor="middle">解析模板</text>' +
  '<path d="M235 150 L305 150" stroke="#666666" stroke-width="4"/>' +
  '<rect x="320" y="110" width="180" height="80" rx="12" fill="#E9002F"/>' +
  '<text x="410" y="158" font-family="Microsoft YaHei" font-size="26" fill="#FFFFFF" text-anchor="middle">生成卡片</text>' +
  '<path d="M515 150 L585 150" stroke="#666666" stroke-width="4"/>' +
  '<rect x="600" y="110" width="90" height="80" rx="12" fill="#232323"/>' +
  '<text x="645" y="158" font-family="Microsoft YaHei" font-size="22" fill="#FFFFFF" text-anchor="middle">导出</text>' +
  '</svg>'

await test('parse → writeTemplate → readTemplate (real HW template)', async () => {
  ensureStore(storeConfig)
  const parsed = await parsePptxTemplate({
    path: TEMPLATE_PATH,
    templateName: 'PPT模板-HW',
    templateId: 'ppt-hw',
    storeConfig,
  })
  assertEqual(parsed.theme.templateId, 'ppt-hw', 'template id honoured')
  assertEqual(parsed.theme.brandAssets.logoPath, join(STORE, 'templates', 'ppt-hw', 'logo.png'), 'logoPath is the absolute store path')
  assertEqual(parsed.logoBytes.length, 22308, 'logo bytes are the master logo')

  const summary = writeTemplate(
    {
      templateId: parsed.theme.templateId,
      templateName: parsed.theme.templateName,
      tokens: parsed.theme,
      logoBytes: parsed.logoBytes,
      meta: parsed.meta,
      media: parsed.media,
      sourceBytes: readFileSync(TEMPLATE_PATH),
    },
    storeConfig,
  )
  assertEqual(summary.hasLogo, true, 'template summary reports the logo')
  assertEqual(summary.slideCount, 5, 'meta slideCount from the real deck')
  assert(statSync(join(STORE, 'templates', 'ppt-hw', 'logo.png')).size === 22308, 'logo.png persisted with the right size')
  assert(existsSync(join(STORE, 'templates', 'ppt-hw', 'source.pptx')), 'original template copied for re-parsing')

  const read = readTemplate('ppt-hw', storeConfig)
  assertEqual(read.hasLogo, true, 'readTemplate has the logo')
  assertEqual(read.theme.brandAssets.logoFrame, { x: 10.0686, y: 6.5356, w: 2.4698, h: 0.5277 }, 'logo frame persisted')
  assertEqual(listTemplates(storeConfig).map((entry) => entry.templateId), ['ppt-hw'], 'template listed')

  state = { parsed, read }
  return { templateId: parsed.theme.templateId, media: parsed.media.length }
})

await test('tokens match research/expected-tokens.json (colors, fonts, logo, size)', () => {
  if (!existsSync(EXPECTED_TOKENS_PATH)) {
    console.log(`  NOTE ${EXPECTED_TOKENS_PATH} is absent — engine dump is authoritative`)
    return 'expected-tokens.json absent'
  }
  const expected = JSON.parse(readFileSync(EXPECTED_TOKENS_PATH, 'utf8'))
  const theme = state.read.theme
  const rows = [
    { field: 'colors.background', expected: expected.colors.background, actual: theme.colors.background },
    { field: 'colors.surface', expected: expected.colors.surface, actual: theme.colors.surface },
    { field: 'colors.textPrimary', expected: expected.colors.textPrimary, actual: theme.colors.textPrimary },
    { field: 'colors.textSecondary', expected: expected.colors.textSecondary, actual: theme.colors.textSecondary },
    { field: 'colors.primary', expected: expected.colors.primary, actual: theme.colors.primary },
    { field: 'colors.accent', expected: expected.colors.accent, actual: theme.colors.accent },
    { field: 'colors.chartPalette', expected: expected.colors.chartPalette, actual: theme.colors.chartPalette },
    { field: 'fonts.title', expected: expected.fonts.title, actual: theme.fonts.title },
    { field: 'fonts.body', expected: expected.fonts.body, actual: theme.fonts.body },
    { field: 'fonts.latin', expected: expected.fonts.latin, actual: theme.fonts.latin },
    { field: 'logo.widthPx', expected: expected.logo.widthPx, actual: state.parsed.meta.logo.widthPx },
    { field: 'logo.heightPx', expected: expected.logo.heightPx, actual: state.parsed.meta.logo.heightPx },
    { field: 'logo.frameInches', expected: expected.logo.frameInches, actual: state.parsed.meta.logo.frameInches },
    { field: 'slideSize.cx', expected: expected.slideSize.cx, actual: state.parsed.raw.slideSize.cx },
    { field: 'slideSize.widthIn', expected: expected.slideSize.widthIn, actual: state.parsed.raw.slideSize.widthIn },
    { field: 'masters', expected: expected.masters, actual: state.parsed.meta.masters },
    { field: 'layouts', expected: expected.layouts, actual: state.parsed.meta.layouts },
  ]
  const mismatches = diffTable(rows)
  assertEqual(mismatches.length, 0, `tokens match expected-tokens.json (${mismatches.map((row) => row.field).join(', ')})`)
  return `${rows.length} fields match`
})

await test('namespaceTemplateId: display names become path-safe ids', () => {
  const ascii = namespaceTemplateId('HW Corporate 2026')
  assertMatch(ascii, /^[a-zA-Z0-9._-]{1,64}$/, 'ascii name is safe')
  assertEqual(ascii, 'hw-corporate-2026', 'ascii name slugged')
  const cjk = namespaceTemplateId('PPT模板-HW')
  assertMatch(cjk, /^[a-zA-Z0-9._-]{1,64}$/, 'CJK name is safe')
  assertEqual(cjk, namespaceTemplateId('PPT模板-HW'), 'mapping is deterministic')
  assert(cjk !== namespaceTemplateId('PPT模板-HW2'), 'different CJK names differ')
  assertEqual(namespaceTemplateId(''), namespaceTemplateId(''), 'empty name is stable')
  return { ascii, cjk }
})

await test('createDeck → writeDeck → readDeck with real tokens', () => {
  const theme = state.read.theme
  const deck = createDeck({ title: '品牌智能演示', templateId: 'ppt-hw', theme })
  assertEqual(CardDeckSchema.safeParse(deck).success, true, 'starter deck is schema-valid')
  writeDeck(deck, storeConfig)
  const reread = readDeck(deck.deckId, storeConfig)
  assertEqual(reread.deckId, deck.deckId, 'deck id round-trips')
  assertEqual(listDecks(storeConfig)[0].deckId, deck.deckId, 'deck listed')
  state.deck = reread
  return reread.deckId
})

await test('compileDeckToFile → re-open the .pptx with jszip', async () => {
  const theme = state.read.theme
  const logoBytes = readFileSync(state.read.logoPath)
  const deck = structuredClone(state.deck)

  // Grow the starter cover into a 3-slide deck that exercises every element type.
  deck.slides.push({
    slideId: 'content',
    background: '$colors.background',
    elements: [
      { type: 'text', id: 'h', role: 'title', content: '业务挑战与机会', x: 96, y: 96, w: 1000, h: 64, style: { fontSize: 36, bold: true, color: '$colors.textPrimary' } },
      { type: 'shape', id: 'card', shapeType: 'roundRect', x: 96, y: 190, w: 480, h: 300, style: { fill: '$colors.surface', radius: 16 } },
      { type: 'text', id: 'lead', role: 'body', content: '内容需求增长 3 倍，排版返工消耗 60% 交付时间。', x: 128, y: 222, w: 420, h: 240, style: { fontSize: 20, color: '$colors.textPrimary' } },
      { type: 'chart', id: 'chart', chartType: 'bar', x: 620, y: 190, w: 560, h: 300, data: { labels: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: '营收', values: [120, 168, 205, 260] }] }, style: { legend: true, showValues: true } },
    ],
    speakerNotes: '内容页备注',
  })
  deck.slides.push({
    slideId: 'flow',
    background: '$colors.background',
    elements: [
      { type: 'text', id: 'h2', role: 'title', content: '品牌流程示意', x: 96, y: 96, w: 1000, h: 64, style: { fontSize: 36, bold: true, color: '$colors.textPrimary' } },
      { type: 'svg', id: 'flow', description: '三步流程', x: 96, y: 200, w: 720, h: 300, svgCode: FLOW_SVG },
      { type: 'image', id: 'logo', assetKey: 'logo', x: 900, y: 480, w: 280, h: 60 },
    ],
  })
  writeDeck(deck, storeConfig)

  const outputPath = join(OUT_DIR, 'engine-e2e-hw.pptx')
  const result = await compileDeckToFile({ deck, theme, logoBytes, assets: { logo: logoBytes } }, outputPath)
  assertEqual(result.path, outputPath, 'export path')
  assertEqual(result.verified, true, 'compiled package verified')
  assertEqual(result.slideCount, 3, 'three slides')

  const onDisk = readFileSync(outputPath)
  assert(onDisk.length > 40000, `export is plausible (${onDisk.length} bytes)`)
  const zip = await JSZip.loadAsync(onDisk)
  assert(zip.file('ppt/presentation.xml') !== null, 'presentation part present')
  const slideParts = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  assertEqual(slideParts.length, 3, 'three slide parts')
  const slide1 = await zip.file('ppt/slides/slide1.xml').async('string')
  assertMatch(slide1, /<a:t>品牌智能演示<\/a:t>/, 'cover title survived the round trip')
  const slide2 = await zip.file('ppt/slides/slide2.xml').async('string')
  assertMatch(slide2, /<a:t>业务挑战与机会<\/a:t>/, 'content title written')
  assertMatch(slide2, /barChart|c:chart/, 'chart is a native chart object')
  const media = Object.keys(zip.files).filter((name) => name.startsWith('ppt/media/'))
  assert(media.length >= 2, `logo + SVG raster embedded (${media.length} media parts)`)
  assert(
    Object.keys(zip.files).some((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)),
    'chart part exists',
  )

  // Fidelity guards on the final artefact.
  const allXml = [slide1, slide2, await zip.file('ppt/slides/slide3.xml').async('string')].join('')
  assert(!/<a:t>(null|undefined|NaN)<\/a:t>/.test(allXml), 'no junk text runs')
  assert(!/<a:ea typeface="Arial"/.test(allXml), 'no Latin face in <a:ea> (CJK would double in renders)')
  assertMatch(allXml, /<a:ea typeface="Microsoft YaHei"/, 'CJK face written into <a:ea>')

  const titleOccurrences = allXml.split('<a:t>品牌智能演示</a:t>').length - 1
  assertEqual(titleOccurrences, 1, 'the cover title appears exactly once')
  state.export = { path: outputPath, bytes: onDisk.length }
  return { path: outputPath, bytes: onDisk.length, media: media.length }
})

await test('the exported file is a real 16:9 presentation (13.333 × 7.5 in)', async () => {
  const zip = await JSZip.loadAsync(readFileSync(state.export.path))
  const presentation = await zip.file('ppt/presentation.xml').async('string')
  const sldSz = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(presentation)
  assert(sldSz !== null, 'presentation.xml declares a slide size')
  const widthIn = Number(sldSz[1]) / 914400
  const heightIn = Number(sldSz[2]) / 914400
  assert(Math.abs(widthIn - 13.3333) < 0.01, `width ≈ 13.333in, got ${widthIn}`)
  assertEqual(heightIn, 7.5, 'height is 7.5in')
  const app = await zip.file('docProps/app.xml')?.async('string')
  assert(app !== undefined && app.includes('<Slides>3</Slides>'), 'docProps/app.xml reports 3 slides')
  return `${widthIn}x${heightIn} in`
})

run('engine-e2e')
