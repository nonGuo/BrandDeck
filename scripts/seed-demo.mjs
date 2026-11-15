/**
 * Seed the live BrandDeck store the running Harness uses, and export a demo deck.
 *
 * This drives the SAME modules the installed plugin runs (parser + store +
 * compiler) against the real profile data directory, so whatever it writes is
 * exactly what the Web panel will list.
 *
 * Run: node scripts/seed-demo.mjs
 */
import { mkdir } from 'node:fs/promises'
import { readFile } from 'node:fs/promises'
import { homedir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const repoRoot = resolve(here, '..')
const pkgRoot = join(repoRoot, 'plugins', 'branddeck')

const { parsePptxTemplate } = await import(pathToFileURL(join(pkgRoot, 'src', 'template.js')).href)
const store = await import(pathToFileURL(join(pkgRoot, 'src', 'store.js')).href)
const { compileDeckToFile } = await import(pathToFileURL(join(pkgRoot, 'src', 'compile.js')).href)
const { parseAndRepairDeck } = await import(pathToFileURL(join(pkgRoot, 'src', 'schema.js')).href)
const { HW_DEMO_DECK } = await import(pathToFileURL(join(pkgRoot, 'test', 'fixtures', 'hw-demo-deck.mjs')).href)

// Same root the plugin resolves at runtime: $DSH_HOME/branddeck (default ~/.dsh).
const dshHome = process.env.DSH_HOME ?? join(homedir(), '.dsh')
const storeConfig = { root: join(dshHome, 'branddeck') }
console.log(`store root: ${store.storeRoot(storeConfig)}`)

const templatePath = join(repoRoot, 'PPT模板-HW.pptx')
const bytes = await readFile(templatePath)
const parsed = await parsePptxTemplate({
  bytes,
  path: templatePath,
  templateId: 'ppt-hw',
  templateName: 'PPT模板-HW',
  storeConfig,
})
await store.writeTemplate(
  {
    templateId: 'ppt-hw',
    templateName: 'PPT模板-HW',
    tokens: parsed.theme,
    logoBytes: parsed.logoBytes,
    meta: parsed.meta,
    media: parsed.media,
    sourcePath: templatePath,
  },
  storeConfig,
)
console.log(
  `template: ${parsed.theme.templateId} · primary ${parsed.theme.colors.primary} · ` +
    `title font ${parsed.theme.fonts.title} · logo ${parsed.logoBytes ? `${parsed.logoBytes.length}B` : 'none'}`,
)

const repaired = parseAndRepairDeck({ ...HW_DEMO_DECK, templateId: 'ppt-hw' }, parsed.theme)
if (!repaired.ok) {
  console.error('demo deck failed validation', repaired.issues)
  process.exit(1)
}
await store.writeDeck(repaired.deck, storeConfig)
console.log(`deck: ${repaired.deck.deckId} · ${repaired.deck.slides.length} slides · ${repaired.warnings.length} warnings`)

const demoDir = join(repoRoot, 'demo')
await mkdir(demoDir, { recursive: true })
const outPath = join(demoDir, 'HW品牌演示-BrandDeck样例.pptx')
const result = await compileDeckToFile(
  {
    deck: repaired.deck,
    theme: parsed.theme,
    logoBytes: parsed.logoBytes,
    assets: {},
    options: { svgRasterScale: 3, loadSystemFonts: true },
  },
  outPath,
)
console.log(`exported: ${result.path} (${result.bytes} bytes, ${result.slideCount} slides)`)
for (const warning of result.warnings ?? []) console.log(`  warn: ${warning}`)

const decks = await store.listDecks(storeConfig)
const templates = await store.listTemplates(storeConfig)
console.log(`store now has ${decks.length} deck(s), ${templates.length} template(s)`)
