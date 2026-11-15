/**
 * End-to-end Host check: real template -> tokens -> stored deck -> compiled .pptx.
 *
 * This runs entirely in Node (no DSH runtime, no browser): it is the regression
 * gate for "does BrandDeck actually produce a correct PowerPoint file".
 *
 * Run: node test/e2e-hw.mjs
 */
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const pkgRoot = resolve(here, '..')
const repoRoot = resolve(pkgRoot, '..', '..')
const templatePath = join(repoRoot, 'PPT模板-HW.pptx')
const outDir = join(pkgRoot, '.tmp', 'e2e')
const storeRoot = join(pkgRoot, '.tmp', 'e2e-store')

const { parsePptxTemplate } = await import('../src/template.js')
const store = await import('../src/store.js')
const { compileDeckToFile } = await import('../src/compile.js')
const { parseAndRepairDeck } = await import('../src/schema.js')
const { HW_DEMO_DECK } = await import('./fixtures/hw-demo-deck.mjs')

const results = []
function record(name, ok, detail) {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${detail === undefined ? '' : ` — ${detail}`}`)
}
function assert(condition, message) {
  if (!condition) throw new Error(message)
}

await rm(outDir, { recursive: true, force: true })
await rm(storeRoot, { recursive: true, force: true })
await mkdir(outDir, { recursive: true })
const storeConfig = { root: storeRoot }

/* 1. Parse the real corporate template ------------------------------------ */
let theme
let parsed
try {
  const bytes = await readFile(templatePath)
  parsed = await parsePptxTemplate({
    bytes,
    path: templatePath,
    templateId: 'ppt-hw',
    templateName: 'PPT模板-HW',
    storeConfig,
  })
  theme = parsed.theme
  await store.writeTemplate(
    {
      templateId: 'ppt-hw',
      templateName: 'PPT模板-HW',
      tokens: theme,
      logoBytes: parsed.logoBytes,
      meta: parsed.meta,
      media: parsed.media,
      sourcePath: templatePath,
    },
    storeConfig,
  )
  record(
    'parse PPT模板-HW.pptx',
    true,
    `colors=${Object.keys(theme.colors).length} fonts=${theme.fonts.title}/${theme.fonts.body} logo=${parsed.logoBytes ? `${parsed.logoBytes.length}B` : 'none'}`,
  )
} catch (error) {
  record('parse PPT模板-HW.pptx', false, error.message)
  throw error
}

/* 2. Cross-check against the independently produced reference ------------- */
try {
  const expectedPath = join(repoRoot, 'research', 'expected-tokens.json')
  if (!existsSync(expectedPath)) {
    record('cross-check tokens vs research/expected-tokens.json', true, 'reference file absent — skipped')
  } else {
    const expected = JSON.parse(await readFile(expectedPath, 'utf8'))
    const mismatches = []
    for (const [key, hex] of Object.entries(expected.colors)) {
      if (Array.isArray(hex)) continue
      if (String(theme.colors[key]).toUpperCase() !== String(hex).toUpperCase()) {
        mismatches.push(`${key}: expected ${hex} got ${theme.colors[key]}`)
      }
    }
    const palette = expected.colors.chartPalette ?? []
    palette.forEach((hex, index) => {
      const actual = theme.colors.chartPalette[index]
      if (String(actual).toUpperCase() !== String(hex).toUpperCase()) {
        mismatches.push(`chartPalette[${index}]: expected ${hex} got ${actual}`)
      }
    })
    for (const key of ['title', 'latin']) {
      const pair = [expected.fonts[key], theme.fonts[key]].map((v) => String(v).toLowerCase())
      const compatible = pair[0] === pair[1] || pair[0].includes(pair[1]) || pair[1].includes(pair[0])
      if (!compatible) mismatches.push(`fonts.${key}: expected ${expected.fonts[key]} got ${theme.fonts[key]}`)
    }
    assert(mismatches.length === 0, mismatches.join(' | '))
    record('cross-check tokens vs research/expected-tokens.json', true, 'colors + palette + fonts match')
  }
} catch (error) {
  record('cross-check tokens vs research/expected-tokens.json', false, error.message)
}

/* 3. Store the deck through the shared validator -------------------------- */
let deck
try {
  const repaired = parseAndRepairDeck({ ...HW_DEMO_DECK, templateId: theme.templateId }, theme)
  assert(repaired.ok, JSON.stringify(repaired.issues))
  deck = repaired.deck
  await store.writeDeck(deck, storeConfig)
  const readBack = await store.readDeck(deck.deckId, storeConfig)
  assert(readBack !== null, 'deck did not read back')
  assert(readBack.slides.length === deck.slides.length, 'slide count changed on round-trip')
  const decks = await store.listDecks(storeConfig)
  record(
    'store deck round-trip',
    true,
    `${decks.length} deck(s), ${deck.slides.length} slides, ${repaired.warnings.length} warning(s)`,
  )
  if (repaired.warnings.length > 0) {
    for (const warning of repaired.warnings.slice(0, 8)) console.log(`      warn: ${warning}`)
  }
} catch (error) {
  record('store deck round-trip', false, error.message)
  throw error
}

/* 4. Compile to a native .pptx ------------------------------------------- */
const pptxPath = join(outDir, 'hw-demo.pptx')
let compileResult
try {
  compileResult = await compileDeckToFile(
    {
      deck,
      theme,
      logoBytes: parsed.logoBytes ?? null,
      assets: {},
      options: { svgRasterScale: 3, loadSystemFonts: true },
    },
    pptxPath,
  )
  assert(existsSync(pptxPath), 'compiled file missing')
  record(
    'compile deck to .pptx',
    true,
    `${compileResult.slideCount} slides, ${(await readFile(pptxPath)).length} bytes, ${(compileResult.warnings ?? []).length} warning(s)`,
  )
  for (const warning of (compileResult.warnings ?? []).slice(0, 8)) console.log(`      warn: ${warning}`)
} catch (error) {
  record('compile deck to .pptx', false, error.message)
  throw error
}

/* 5. Structural verification of the produced ZIP -------------------------- */
try {
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(await readFile(pptxPath))
  const names = Object.keys(zip.files)
  const slides = names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
  assert(slides.length === deck.slides.length, `slide parts ${slides.length} != ${deck.slides.length}`)
  assert(names.includes('ppt/presentation.xml'), 'missing presentation.xml')
  assert(
    names.includes('ppt/slideMasters/slideMaster1.xml'),
    'missing slide master (defineSlideMaster should emit one)',
  )
  const media = names.filter((n) => n.startsWith('ppt/media/'))
  assert(media.length >= 2, `expected logo + rasterized svg images, found ${media.length}`)

  const coverXml = await zip.file('ppt/slides/slide1.xml').async('string')
  assert(coverXml.includes('AI 驱动的企业内容生产'), 'cover title text missing from slide1.xml')
  const roadmapXml = await zip.file('ppt/slides/slide5.xml').async('string')
  const roadmapXmlLower = roadmapXml.toLowerCase()
  const hasChartOrImage = roadmapXmlLower.includes('graphicframe') || roadmapXmlLower.includes('pic')
  assert(hasChartOrImage, 'roadmap slide has neither a picture nor a graphic frame')

  const chartSlide = await zip.file('ppt/slides/slide4.xml').async('string')
  assert(/graphicFrame/i.test(chartSlide), 'chart slide has no graphic frame')
  assert(names.some((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)), 'no native chart part was written')

  const presXml = await zip.file('ppt/presentation.xml').async('string')
  const size = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presXml)
  assert(size !== null, 'presentation.xml has no sldSz')
  const widthIn = Number(size[1]) / 914400
  const heightIn = Number(size[2]) / 914400
  assert(Math.abs(widthIn - 13.333) < 0.05, `slide width ${widthIn}in is not 16:9 wide`)
  assert(Math.abs(heightIn - 7.5) < 0.05, `slide height ${heightIn}in is not 7.5in`)

  record(
    'verify .pptx structure',
    true,
    `${slides.length} slides, ${media.length} media parts, chart part present, ${widthIn.toFixed(3)}x${heightIn.toFixed(2)}in`,
  )
} catch (error) {
  record('verify .pptx structure', false, error.message)
}

/* 6. Notes + summary ------------------------------------------------------ */
const summaryPath = join(outDir, 'e2e-summary.json')
await writeFile(
  summaryPath,
  JSON.stringify(
    {
      template: {
        colors: theme.colors,
        fonts: theme.fonts,
        logo: parsed.logoBytes ? `${parsed.logoBytes.length} bytes` : null,
        meta: parsed.meta,
      },
      deck: { deckId: deck.deckId, slides: deck.slides.length },
      pptx: pptxPath,
      compile: { slideCount: compileResult?.slideCount, warnings: compileResult?.warnings ?? [] },
      results,
    },
    null,
    2,
  ),
  'utf8',
)

const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} end-to-end checks passed`)
console.log(`pptx: ${pptxPath}`)
if (failed.length > 0) process.exitCode = 1
