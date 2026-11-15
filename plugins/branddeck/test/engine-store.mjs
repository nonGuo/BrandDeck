/**
 * store.js: paths, atomic JSON, template + deck CRUD, id safety and the starter deck.
 *
 * Everything runs inside a throwaway root under `.tmp/` — the real `$DSH_HOME` is never
 * touched (and the test asserts that).
 *
 * Run: node test/engine-store.mjs
 */
import { existsSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import {
  CANVAS,
  CardDeckSchema,
  LOGO_RESERVE,
  SAFE_AREA,
  ThemeTokensSchema,
} from '../src/schema.js'
import {
  assert,
  assertCode,
  assertEqual,
  assertMatch,
  freshTempDir,
  run,
  test,
} from './engine-helpers.mjs'
import {
  assertSafeId,
  createDeck,
  deleteDeck,
  ensureStore,
  listDecks,
  listTemplates,
  readDeck,
  readTemplate,
  slugifyId,
  storeRoot,
  summarizeDeck,
  summarizeTemplate,
  writeDeck,
  writeTemplate,
} from '../src/store.js'
import { DEFAULT_THEME, makeDefaultTheme } from '../src/defaults.js'

const ROOT = freshTempDir('store')

/** A theme with the out-of-contract `logoFrame` asset the compiler honours. */
function templateTokens(templateId = 'hw-test') {
  const theme = makeDefaultTheme({ templateId, templateName: 'HW 测试模板' })
  theme.colors = {
    background: '#FFFFFF',
    surface: '#EBEBEB',
    textPrimary: '#1D1D1A',
    textSecondary: '#666666',
    primary: '#C7000A',
    accent: '#E9002F',
    chartPalette: ['#C7000A', '#E9002F', '#F4A100', '#FFFF00', '#232323', '#666666'],
  }
  theme.fonts = { title: 'Microsoft YaHei', body: 'Arial', latin: 'Arial' }
  theme.brandAssets = {
    icons: {},
    logoPath: join(ROOT, 'templates', templateId, 'logo.png'),
    logoWidthPx: 529,
    logoHeightPx: 116,
    logoFrame: { x: 10.0686, y: 6.5356, w: 2.4698, h: 0.5277 },
  }
  return theme
}

/** A 1×1 transparent PNG (small, valid, and byte-identical after the round trip). */
const LOGO_BYTES = Buffer.from(
  '89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63fcffff3f030005fe02fea735e10f0000000049454e44ae426082',
  'hex',
)

await test('storeRoot: honours DSH_HOME, config.root and config.home', () => {
  const previous = process.env.DSH_HOME
  try {
    process.env.DSH_HOME = join(ROOT, 'fake-dsh-home')
    assertMatch(storeRoot(), /fake-dsh-home[\\/]branddeck$/, 'DSH_HOME gets /branddeck appended')
    assertMatch(storeRoot({ root: join(ROOT, 'custom') }), /custom$/, 'config.root is used verbatim')
    assertMatch(storeRoot({ home: join(ROOT, 'other-home'), root: undefined }), /other-home[\\/]branddeck$/, 'config.home gets /branddeck')
  } finally {
    if (previous === undefined) delete process.env.DSH_HOME
    else process.env.DSH_HOME = previous
  }
  return storeRoot({ root: ROOT })
})

await test('ensureStore: creates templates/ and decks/', () => {
  const paths = ensureStore({ root: ROOT })
  assertEqual(paths.root, resolve(ROOT), 'root resolved')
  assert(existsSync(paths.templatesDir), 'templates/ exists')
  assert(existsSync(paths.decksDir), 'decks/ exists')
  assert(statSync(paths.templatesDir).isDirectory(), 'templates/ is a directory')
  return Object.keys(paths)
})

await test('assertSafeId: rejects traversal and odd ids, accepts safe ones', async () => {
  for (const bad of ['../x', '..', '.', 'a/b', 'a\\b', '', 'x'.repeat(65), 'has space', 'nul\u0000', 'a*b']) {
    await assertCode(() => assertSafeId(bad), 'invalid_id', `id ${JSON.stringify(bad)} is rejected`)
  }
  for (const good of ['hw', 'hw-2', 'A.b_c-1', 'x'.repeat(64), 'deck-46e38c']) {
    assertEqual(assertSafeId(good), good, `id ${JSON.stringify(good)} is accepted`)
  }
  return 'id safety ok'
})

await test('writeTemplate/readTemplate/listTemplates: full round trip with logo + meta', () => {
  const tokens = templateTokens('hw-test')
  const summary = writeTemplate(
    {
      templateId: 'hw-test',
      templateName: 'HW 测试模板',
      tokens,
      logoBytes: LOGO_BYTES,
      meta: {
        sourcePath: 'C:/templates/PPT模板-HW.pptx',
        slideCount: 5,
        masters: 4,
        layouts: 7,
        slideSize: { cx: 12196763, cy: 6858000, widthIn: 13.3385, heightIn: 7.5 },
      },
      media: [{ path: 'ppt/media/image1.png', bytes: 22308, contentType: 'image/png', widthPx: 529, heightPx: 116, referencedBy: ['ppt/slideMasters/_rels/slideMaster1.xml.rels'] }],
      sourceBytes: Buffer.from('fake original pptx bytes'),
    },
    { root: ROOT },
  )
  assertEqual(summary.templateId, 'hw-test', 'summary carries the id')
  assertEqual(summary.hasLogo, true, 'summary reports the logo')
  assertEqual(summary.colors.primary, '#C7000A', 'summary carries the palette')
  assertMatch(summary.updatedAt, /^\d{4}-\d{2}-\d{2}T/, 'summary has an ISO timestamp')

  const dir = join(ROOT, 'templates', 'hw-test')
  for (const file of ['theme.json', 'logo.png', 'meta.json', 'source.pptx']) {
    assert(existsSync(join(dir, file)), `${file} was persisted`)
  }
  const themeOnDisk = JSON.parse(readFileSync(join(dir, 'theme.json'), 'utf8'))
  assertEqual(themeOnDisk.brandAssets.logoPath, join(dir, 'logo.png'), 'logoPath points at the persisted logo')
  assertEqual(themeOnDisk.brandAssets.logoFrame, { x: 10.0686, y: 6.5356, w: 2.4698, h: 0.5277 }, 'logoFrame survives the schema strip')
  assertMatch(readFileSync(join(dir, 'theme.json'), 'utf8'), /\n  "templateId"/, 'JSON is pretty-printed')
  assertEqual(readFileSync(join(dir, 'logo.png')).equals(LOGO_BYTES), true, 'logo bytes are byte-identical')
  assertEqual(JSON.parse(readFileSync(join(dir, 'meta.json'), 'utf8')).slideCount, 5, 'meta.json keeps slideCount')

  const read = readTemplate('hw-test', { root: ROOT })
  assertEqual(read.hasLogo, true, 'readTemplate reports the logo')
  assertEqual(read.logoPath, join(dir, 'logo.png'), 'readTemplate returns the absolute logo path')
  assertEqual(read.theme.brandAssets.logoPath, join(dir, 'logo.png'), 'theme.brandAssets.logoPath is absolute')
  assert(read.logoPath !== null && existsSync(read.logoPath), 'the logo file really exists')
  assertEqual(ThemeTokensSchema.safeParse(read.theme).success, true, 'read theme still satisfies the schema')

  const listed = listTemplates({ root: ROOT })
  assertEqual(listed.length, 1, 'one template listed')
  assertEqual(listed[0].templateId, 'hw-test', 'listing shows the template')
  assertEqual(listed[0].mediaCount, 1, 'listing counts media')
  return listed[0]
})

await test('writeTemplate: rejects invalid ids and invalid themes', async () => {
  await assertCode(
    () => writeTemplate({ templateId: '../evil', templateName: 'x', tokens: templateTokens('ok') }, { root: ROOT }),
    'invalid_id',
    'traversal id refused',
  )
  await assertCode(
    () => writeTemplate({ templateId: 'bad-theme', templateName: 'x', tokens: { templateId: 'bad-theme' } }, { root: ROOT }),
    'invalid_theme',
    'incomplete tokens refused',
  )
  await assertCode(
    () => writeTemplate({ templateId: 'no-name', templateName: '   ', tokens: templateTokens('no-name') }, { root: ROOT }),
    'invalid_template',
    'blank template name refused',
  )
  await assertCode(() => readTemplate('does-not-exist', { root: ROOT }), 'template_not_found', 'missing template')
  return 'validation ok'
})

await test('createDeck: schema-valid starter deck inside the safe area', () => {
  const theme = templateTokens('hw-test')
  const deck = createDeck({ title: 'Brand Deck 2026', templateId: 'hw-test', theme })
  assertEqual(CardDeckSchema.safeParse(deck).success, true, 'starter deck satisfies CardDeckSchema')
  assertEqual(deck.deckId, 'brand-deck-2026', 'deckId slugged from the title')
  assertEqual(deck.templateId, 'hw-test', 'templateId carried over')
  assertEqual(deck.slides.length, 1, 'one cover slide')
  const slide = deck.slides[0]
  assertEqual(slide.elements.length, 5, 'cover has 5 elements')
  assertEqual(slide.background, '$colors.background', 'background is a token reference')

  const logoX = CANVAS.width - SAFE_AREA.right - LOGO_RESERVE.width
  for (const element of slide.elements) {
    assert(element.x >= SAFE_AREA.left, `${element.id} respects the left margin`)
    assert(element.y >= SAFE_AREA.top, `${element.id} respects the top margin`)
    assert(element.x + element.w <= CANVAS.width - SAFE_AREA.right, `${element.id} respects the right margin`)
    assert(element.y + element.h <= CANVAS.height - SAFE_AREA.bottom, `${element.id} respects the bottom margin`)
    if (element.type === 'text' && element.role !== 'caption') {
      const clearOfLogo =
        element.x + element.w <= logoX || element.y >= SAFE_AREA.top + LOGO_RESERVE.height + LOGO_RESERVE.gap
      assert(clearOfLogo, `${element.id} stays clear of the top-right logo reserve`)
    }
  }
  const colors = slide.elements.filter((element) => element.style?.fill ?? element.style?.color)
  assert(
    colors.every((element) => String(element.style.fill ?? element.style.color).startsWith('$colors.')),
    'starter styling uses $colors token references',
  )
  const title = slide.elements.find((element) => element.id === 'title')
  assertEqual(title.content, 'Brand Deck 2026', 'title text set')
  assertEqual(title.role, 'title', 'title role set')
  assert(typeof slide.speakerNotes === 'string' && slide.speakerNotes.length > 0, 'speaker notes filled')

  // A pure-CJK title must still produce a path-safe, deterministic id.
  const cjk = createDeck({ title: '品牌智能演示', templateId: 'hw-test', theme })
  assertMatch(cjk.deckId, /^deck-[0-9a-f]{6}$/, 'CJK-only title falls back to a deterministic digest id')
  assertEqual(createDeck({ title: '品牌智能演示', templateId: 'hw-test', theme }).deckId, cjk.deckId, 'the digest id is stable')
  return { deckId: deck.deckId, cjkDeckId: cjk.deckId }
})

await test('writeDeck/readDeck/listDecks/deleteDeck: CRUD round trip, newest first', async () => {
  const theme = templateTokens('hw-test')
  const first = createDeck({ title: 'First Deck', templateId: 'hw-test', theme })
  const second = createDeck({ title: 'Second Deck', templateId: 'hw-test', theme, deckId: 'second-deck' })
  const firstSummary = writeDeck(first, { root: ROOT })
  assertEqual(firstSummary.slideCount, 1, 'summary counts slides')
  await new Promise((done) => setTimeout(done, 20)) // ensure a distinct mtime
  writeDeck(second, { root: ROOT })

  assert(existsSync(join(ROOT, 'decks', 'first-deck.json')), 'deck file written')
  assertEqual(readDeck('first-deck', { root: ROOT }).title, 'First Deck', 'readDeck returns the deck')
  const listed = listDecks({ root: ROOT })
  assertEqual(listed.map((entry) => entry.deckId), ['second-deck', 'first-deck'], 'newest first')
  assertEqual(summarizeDeck(first, { updatedAt: 'x' }).slideCount, 1, 'summarizeDeck')

  assertEqual(deleteDeck('first-deck', { root: ROOT }), true, 'deleteDeck removes an existing deck')
  assertEqual(deleteDeck('first-deck', { root: ROOT }), false, 'deleteDeck returns false when absent')
  await assertCode(() => readDeck('first-deck', { root: ROOT }), 'deck_not_found', 'deleted deck is gone')
  await assertCode(() => readDeck('../escape', { root: ROOT }), 'invalid_id', 'traversal id refused by readDeck')
  await assertCode(() => writeDeck({ deckId: 'broken' }, { root: ROOT }), 'invalid_deck', 'invalid deck refused')
  return listed.map((entry) => entry.deckId)
})

await test('createDeck({persist:true}) writes the deck and nothing escapes the store root', () => {
  const deck = createDeck({ title: 'Persisted Cover', templateId: 'hw-test', theme: templateTokens('hw-test'), persist: true, config: { root: ROOT } })
  assertEqual(readDeck(deck.deckId, { root: ROOT }).title, 'Persisted Cover', 'persist:true wrote the deck')

  // Everything the store produced must live under ROOT.
  const resolveRoot = resolve(ROOT)
  /** @param {string} dir */
  const walk = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const absolute = resolve(join(dir, entry.name))
      assert(absolute.startsWith(resolveRoot), `${absolute} is inside the store root`)
      if (entry.isDirectory()) walk(absolute)
    }
  }
  walk(ROOT)
  const files = readdirSync(join(ROOT, 'decks'))
  assert(files.every((name) => name.endsWith('.json')), 'decks/ holds only .json files')
  return `${files.length} deck file(s)`
})

await test('summarizeTemplate/summarizeDeck: stable shapes for the HTTP API', () => {
  const theme = templateTokens('hw-test')
  const templateSummary = summarizeTemplate({ theme, meta: { updatedAt: '2026-01-01T00:00:00.000Z', slideCount: 5, media: [1, 2] }, hasLogo: true })
  assertEqual(Object.keys(templateSummary).sort(), [
    'colors',
    'fonts',
    'hasLogo',
    'mediaCount',
    'slideCount',
    'templateId',
    'templateName',
    'updatedAt',
  ], 'TemplateSummary keys')
  const deckSummary = summarizeDeck({ deckId: 'd', title: 't', templateId: 'x', slides: [{}, {}] })
  assertEqual(Object.keys(deckSummary).sort(), ['deckId', 'slideCount', 'templateId', 'title', 'updatedAt'], 'DeckSummary keys')
  assertEqual(deckSummary.slideCount, 2, 'slide count')
  assertEqual(summarizeDeck({}).deckId, 'unknown', 'summarizeDeck tolerates junk')

  // Defaults must stay schema-valid and frozen.
  assertEqual(ThemeTokensSchema.safeParse(DEFAULT_THEME).success, true, 'DEFAULT_THEME is schema-valid')
  assert(Object.isFrozen(DEFAULT_THEME), 'DEFAULT_THEME is frozen')
  const copy = makeDefaultTheme({ templateId: 'custom' })
  copy.colors.primary = '#000000'
  assertEqual(DEFAULT_THEME.colors.primary, '#2F5BD2', 'makeDefaultTheme returns an independent copy')
  return templateSummary.templateId
})

await test('slugifyId: ascii slugs stay readable, junk never escapes the id grammar', () => {
  assertEqual(slugifyId('Brand Deck 2026', 'deck'), 'brand-deck-2026', 'ascii title slugged')
  assertMatch(slugifyId('!!!', 'deck'), /^deck-[0-9a-f]{6}$/, 'punctuation-only falls back to a digest')
  assertMatch(slugifyId('2026', 'deck'), /^deck-2026$/, 'digit-leading gets a prefix')
  assertMatch(slugifyId('品牌演示', 'deck'), /^deck-[0-9a-f]{6}$/, 'CJK falls back to a digest')
  for (const input of ['a b', '../x', 'A/B', 'x'.repeat(120)]) {
    assertMatch(slugifyId(input, 'deck'), /^[a-zA-Z0-9._-]{1,64}$/, `${JSON.stringify(input)} slug is safe`)
  }
  return 'slug ok'
})

run('engine-store')
