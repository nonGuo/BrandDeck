/**
 * Host API contract test: exercises the exact Connection RPC endpoint table the
 * Web panel calls, without a browser or a DSH runtime.
 *
 * `dispatchEndpoint` is the handler body registered on `/branddeck`, so calling
 * it directly asserts the panel contract (envelopes, payloads, error codes).
 *
 * Run: node test/api-test.mjs
 */
import { mkdir, readFile, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const pkgRoot = resolve(here, '..')
const repoRoot = resolve(pkgRoot, '..', '..')
const templatePath = join(repoRoot, 'PPT模板-HW.pptx')
const storeRoot = join(pkgRoot, '.tmp', 'api-store')
const outDir = join(pkgRoot, '.tmp', 'api-out')

await rm(storeRoot, { recursive: true, force: true })
await rm(outDir, { recursive: true, force: true })
await mkdir(outDir, { recursive: true })

const { __internals } = await import('../index.js')
const { HW_DEMO_DECK } = await import('./fixtures/hw-demo-deck.mjs')

const config = { svgRasterScale: 3, loadSystemFonts: true }
const storeConfig = { root: storeRoot }
const deps = { config, storeConfig }
const call = (endpoint, payload) => __internals.dispatchEndpoint(endpoint, payload, deps)

const results = []
function check(name, fn) {
  try {
    const detail = fn()
    results.push({ name, ok: true })
    console.log(`PASS ${name}${detail === undefined ? '' : ` — ${detail}`}`)
  } catch (error) {
    results.push({ name, ok: false, error })
    console.log(`FAIL ${name}: ${error?.message ?? error}`)
  }
  }
const assert = (condition, message) => {
  if (!condition) throw new Error(message)
}
const okOf = (envelope, label) => {
  assert(envelope?.ok === true, `${label}: ${JSON.stringify(envelope?.error ?? envelope)}`)
  return envelope.value
}

/* Empty store -------------------------------------------------------------- */
const empty = await call('state', {})
check('state on an empty store is an empty state, not an error', () => {
  const value = okOf(empty, 'state')
  assert(value.deck === null, 'expected deck null')
  assert(Array.isArray(value.decks) && value.decks.length === 0, 'expected no decks')
  assert(value.theme && typeof value.theme.colors === 'object', 'expected a fallback theme')
  assert(value.capability?.canvas?.width === 1280, 'expected canvas capability')
})

/* Template parse ----------------------------------------------------------- */
const parsed = await call('templates.parse', { path: templatePath, templateId: 'ppt-hw', templateName: 'PPT模板-HW' })
check('templates.parse stores the corporate template', () => {
  const value = okOf(parsed, 'templates.parse')
  assert(value.template.templateId === 'ppt-hw', 'templateId mismatch')
  assert(value.theme.colors.primary.toUpperCase() === '#C7000A', `primary: ${value.theme.colors.primary}`)
  assert(value.theme.fonts.title.length > 0, 'missing title font')
  return `${value.theme.colors.primary} / ${value.theme.fonts.title}`
})

const logo = await call('templates.logo', { templateId: 'ppt-hw' })
check('templates.logo returns a PNG data URL', () => {
  const value = okOf(logo, 'templates.logo')
  assert(typeof value.dataUrl === 'string' && value.dataUrl.startsWith('data:image/png;base64,'), 'not a png data url')
  assert(value.widthPx > 0 && value.heightPx > 0, 'missing dimensions')
  return `${value.widthPx}x${value.heightPx}, ${value.dataUrl.length} chars`
})

const missingLogo = await call('templates.logo', { templateId: 'nope' })
check('templates.logo fails cleanly for an unknown template', () => {
  assert(missingLogo.ok === false, 'expected failure envelope')
  assert(missingLogo.error.code === 'logo_not_found' || missingLogo.error.code === 'template_not_found', `code: ${missingLogo.error.code}`)
  return missingLogo.error.code
})

/* Deck creation + save ----------------------------------------------------- */
const created = await call('decks.create', { title: '面板新建卡组', templateId: 'ppt-hw' })
const createdDeck = okOf(created, 'decks.create').deck
check('decks.create returns a stored, schema-valid starter deck', () => {
  assert(createdDeck.deckId.length > 0, 'missing deckId')
  assert(createdDeck.slides.length >= 1, 'no slides')
  assert(createdDeck.templateId === 'ppt-hw', 'template not applied')
  return `${createdDeck.deckId}, ${createdDeck.slides.length} slide(s)`
})

const listed = await call('decks.list', {})
check('decks.list sees the created deck', () => {
  const value = okOf(listed, 'decks.list')
  assert(value.decks.some((d) => d.deckId === createdDeck.deckId), 'created deck not listed')
  return `${value.decks.length} deck(s)`
})

const saved = await call('decks.save', { deck: { ...HW_DEMO_DECK, templateId: 'ppt-hw' } })
check('decks.save repairs and stores the demo deck', () => {
  const value = okOf(saved, 'decks.save')
  assert(value.deck.slides.length === HW_DEMO_DECK.slides.length, 'slide count changed')
  assert(Array.isArray(value.warnings), 'warnings must be an array')
  return `${value.deck.slides.length} slides, ${value.warnings.length} warning(s)`
})

const fetched = await call('decks.get', { deckId: 'hw-demo' })
check('decks.get reads the saved deck back with its theme', () => {
  const value = okOf(fetched, 'decks.get')
  assert(value.deck.deckId === 'hw-demo', 'wrong deck')
  assert(value.theme.colors.primary.toUpperCase() === '#C7000A', 'theme not resolved')
  return `${value.deck.slides.length} slides`
})

const badSave = await call('decks.save', { deck: { deckId: 'x', title: '', templateId: 't', slides: [] } })
check('decks.save rejects a schema-invalid deck with structured issues', () => {
  assert(badSave.ok === false, 'expected failure')
  assert(badSave.error.code === 'invalid_deck', `code: ${badSave.error.code}`)
  assert(Array.isArray(badSave.error.details) && badSave.error.details.length > 0, 'missing details')
  return `${badSave.error.details.length} issue(s)`
})

const notFound = await call('decks.get', { deckId: 'does-not-exist' })
check('decks.get reports deck_not_found (panel maps this to the empty state)', () => {
  assert(notFound.ok === false, 'expected failure')
  assert(notFound.error.code === 'deck_not_found', `code: ${notFound.error.code}`)
})

/* Export through the panel path -------------------------------------------- */
const exported = await call('export', { deckId: 'hw-demo' })
check('export compiles the deck into a .pptx on disk', () => {
  const value = okOf(exported, 'export')
  assert(existsSync(value.path), `missing file ${value.path}`)
  assert(value.slideCount === HW_DEMO_DECK.slides.length, `slideCount ${value.slideCount}`)
  assert(value.path.includes('exports'), `export should land in the store exports dir: ${value.path}`)
  return `${value.path} (${value.slideCount} slides)`
})

const JSZip = (await import('jszip')).default
const zip = await JSZip.loadAsync(await readFile(exported.value.path))
check('exported package contains slides, a master and media', () => {
  const names = Object.keys(zip.files)
  const slides = names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
  const media = names.filter((n) => n.startsWith('ppt/media/'))
  assert(slides.length === HW_DEMO_DECK.slides.length, `slides ${slides.length}`)
  assert(names.includes('ppt/presentation.xml'), 'no presentation.xml')
  assert(names.some((n) => n.startsWith('ppt/slideMasters/')), 'no master')
  assert(media.length >= 2, `media ${media.length}`)
  return `${slides.length} slides, ${media.length} media, ${(names.length)} parts`
})

let cjkReport = 'n/a'
for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
  const xml = await zip.file(name).async('string')
  const eaFaces = [...xml.matchAll(/<a:ea typeface="([^"]*)"/g)].map((m) => m[1])
  const latinInEa = eaFaces.filter((face) => /arial|calibri|helvetica/i.test(face))
  if (latinInEa.length > 0) {
    cjkReport = `${name} still has latin <a:ea>: ${latinInEa[0]}`
    break
  }
  cjkReport = `${eaFaces.length} ea face(s): ${[...new Set(eaFaces)].join(', ')}`
}
check('no Latin typeface is written into <a:ea>', () => {
  assert(!cjkReport.includes('still has'), cjkReport)
  return cjkReport
})

const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} API contract checks passed`)
if (failed.length > 0) process.exitCode = 1
