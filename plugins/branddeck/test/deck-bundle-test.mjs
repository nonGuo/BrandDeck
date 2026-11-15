/**
 * Deck-bundle test: the intermediate artefacts users keep tuning.
 *
 * A deck is stored as a directory so generation is never one-shot:
 *   card.json      geometry truth (the「卡片视图」)
 *   outline.md     content truth  (the「文字版本」)
 *   meta.json      revision counter + timestamp
 *   revisions/NNNN-{card.json,outline.md}   history, pruned to a bound
 *
 * Also covers the template side (theme.json + layouts.json + template.md +
 * source.pptx) and the migration from the old single-file layout.
 *
 * Run: node test/deck-bundle-test.mjs
 */
import { existsSync, readFileSync, readdirSync, rmSync, writeFileSync, mkdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import * as store from '../src/store.js'
import { HW_DEMO_DECK } from './fixtures/hw-demo-deck.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const root = join(here, '..', '.tmp', 'bundle-store')
rmSync(root, { recursive: true, force: true })
mkdirSync(root, { recursive: true })
const config = { root }

const HW_THEME = {
  templateId: 'ppt-hw',
  templateName: 'PPT模板-HW',
  colors: {
    background: '#FFFFFF',
    surface: '#EBEBEB',
    textPrimary: '#1D1D1A',
    textSecondary: '#666666',
    primary: '#C7000A',
    accent: '#E9002F',
    chartPalette: ['#C7000A', '#E9002F', '#F4A100', '#FFFF00', '#232323', '#666666'],
  },
  fonts: { title: 'Microsoft YaHei', body: 'Arial', latin: 'Arial' },
  brandAssets: { logoPath: 'D:/tmp/logo.png', logoWidthPx: 529, logoHeightPx: 116, icons: {} },
}

const results = []
function check(name, fn) {
  try {
    const detail = fn()
    results.push({ name, ok: true })
    console.log(`PASS ${name}${detail === undefined ? '' : ` — ${detail}`}`)
  } catch (error) {
    results.push({ name, ok: false })
    console.log(`FAIL ${name}: ${error?.message ?? error}`)
  }
}
const assert = (condition, message) => {
  if (!condition) throw new Error(message)
}
const clone = (value) => structuredClone(value)

const deckDir = join(root, 'decks', 'hw-demo')

/* 1. Bundle layout -------------------------------------------------------- */
check('writeDeck materialises card + outline + meta + a revision snapshot', () => {
  const summary = store.writeDeck(clone(HW_DEMO_DECK), config, { theme: HW_THEME })
  assert(summary.revision === 1, `first revision: ${summary.revision}`)
  for (const name of ['card.json', 'outline.md', 'meta.json']) {
    assert(existsSync(join(deckDir, name)), `missing ${name}`)
  }
  const revisions = readdirSync(join(deckDir, 'revisions')).sort()
  assert(
    revisions.join(',') === '0001-card.json,0001-outline.md',
    `revision files: ${revisions.join(',')}`,
  )
  // The legacy single-file location must not linger, or the deck lists twice.
  assert(!existsSync(join(root, 'decks', 'hw-demo.json')), 'legacy deck file still present')
  return `rev${summary.revision}, ${revisions.length} snapshot file(s)`
})

/* 2. The text twin is real markdown with stable ids ---------------------- */
check('outline.md is readable markdown carrying element ids', () => {
  const outline = store.readDeckOutline('hw-demo', config, HW_THEME)
  assert(outline.markdown.startsWith('# '), 'outline should start with the deck title')
  assert(outline.markdown.includes('{#cover-title}'), 'element ids missing')
  assert(outline.markdown.includes('AI 驱动的企业内容生产'), 'cover copy missing')
  // Trailing whitespace means spaces/tabs at the end of a line — not the file's
  // final newline, which is required.
  const dirtyLine = outline.markdown.split('\n').find((line) => /[ \t]+$/.test(line))
  assert(dirtyLine === undefined, `trailing whitespace on a line: ${JSON.stringify(dirtyLine)}`)
  assert(outline.markdown.endsWith('\n') && !outline.markdown.endsWith('\n\n'), 'outline must end with exactly one newline')
  return `${outline.markdown.length} chars, revision ${outline.revision}`
})

/* 3. Editing the text keeps the geometry --------------------------------- */
check('saving an edited outline updates copy and preserves geometry/style', () => {
  const before = store.readDeck('hw-demo', config)
  const beforeElement = before.slides[0].elements.find((element) => element.id === 'cover-title')
  const outline = store.readDeckOutline('hw-demo', config, HW_THEME)
  const edited = outline.markdown.replace(
    '### 标题 {#cover-title}\nAI 驱动的企业内容生产',
    '### 标题 {#cover-title}\n改稿后的主标题',
  )
  assert(edited !== outline.markdown, 'edit did not change the markdown')

  const applied = store.writeDeckOutline('hw-demo', edited, config, { theme: HW_THEME })
  const after = applied.deck.slides[0].elements.find((element) => element.id === 'cover-title')
  assert(after.content === '改稿后的主标题', `content: ${after.content}`)
  for (const field of ['x', 'y', 'w', 'h']) {
    assert(after[field] === beforeElement[field], `${field} drifted: ${after[field]}`)
  }
  assert(
    JSON.stringify(after.style) === JSON.stringify(beforeElement.style),
    'style drifted',
  )
  assert(applied.changes.updated === 1, `updated: ${applied.changes.updated}`)
  assert(applied.changes.untouched > 0, 'untouched count missing')
  return `updated=${applied.changes.updated} untouched=${applied.changes.untouched}`
})

/* 4. History grows and stays bounded ------------------------------------- */
check('each save appends a revision snapshot and the counter only moves forward', () => {
  const first = store.writeDeck(clone(HW_DEMO_DECK), config, { theme: HW_THEME })
  const second = store.writeDeck(clone(HW_DEMO_DECK), config, { theme: HW_THEME })
  assert(second.revision === first.revision + 1, `${first.revision} -> ${second.revision}`)
  const files = readdirSync(join(deckDir, 'revisions')).sort()
  const newest = files.filter((name) => name.endsWith('-card.json')).pop()
  assert(newest !== undefined && newest.startsWith(String(second.revision).padStart(4, '0')), `newest snapshot: ${newest}`)
  const meta = JSON.parse(readFileSync(join(deckDir, 'meta.json'), 'utf8'))
  assert(Array.isArray(meta.revisions) && meta.revisions.length === second.revision, 'meta.revisions out of sync')
  assert(meta.revisions[meta.revisions.length - 1].card === `revisions/${newest}`, 'meta entry points at the wrong file')
  return `${second.revision} revisions, ${files.length} files`
})

/* 5. Listing sees one entry with its revision ---------------------------- */
check('listDecks reports one entry per deck, with its revision', () => {
  const decks = store.listDecks(config)
  assert(decks.length === 1, `expected 1 deck, got ${decks.length}`)
  assert(decks[0].deckId === 'hw-demo', decks[0].deckId)
  assert(Number.isFinite(decks[0].revision) && decks[0].revision > 0, `revision: ${decks[0].revision}`)
  return `rev${decks[0].revision}`
})

/* 6. Migration from the old single-file store ---------------------------- */
check('a legacy <deckId>.json still reads, and the next save migrates it', () => {
  const legacyId = 'legacy-deck'
  const legacyFile = join(root, 'decks', `${legacyId}.json`)
  writeFileSync(legacyFile, JSON.stringify({ ...clone(HW_DEMO_DECK), deckId: legacyId }, null, 2), 'utf8')
  const read = store.readDeck(legacyId, config)
  assert(read.deckId === legacyId, 'legacy deck unreadable')

  store.writeDeck(read, config, { theme: HW_THEME })
  assert(!existsSync(legacyFile), 'legacy file not removed after migration')
  assert(existsSync(join(root, 'decks', legacyId, 'card.json')), 'migrated card.json missing')
  assert(existsSync(join(root, 'decks', legacyId, 'outline.md')), 'migrated outline missing')
  return 'read + migrated'
})

/* 7. Materialising into a visible directory ------------------------------ */
check('materializeDeck writes card.json + outline.md + meta.json for review', () => {
  const target = join(here, '..', '.tmp', 'bundle-materialized')
  rmSync(target, { recursive: true, force: true })
  const result = store.materializeDeck('hw-demo', target, config, { theme: HW_THEME })
  for (const name of ['card.json', 'outline.md', 'meta.json']) {
    assert(existsSync(join(result.dir, name)), `missing ${name}`)
  }
  const meta = JSON.parse(readFileSync(join(result.dir, 'meta.json'), 'utf8'))
  assert(meta.deckId === 'hw-demo', 'materialized meta has the wrong deck')
  assert(typeof meta.note === 'string' && meta.note.length > 0, 'meta note missing')
  return result.files.length + ' files'
})

/* 8. Template side: token + layout + spec + source copy ------------------ */
check('a template is stored as tokens + layouts + spec + source copy', () => {
  const templateDir = join(root, 'templates', 'spec-test')
  mkdirSync(templateDir, { recursive: true })
  const summary = store.writeTemplate(
    {
      templateId: 'spec-test',
      templateName: 'Spec Test',
      tokens: { ...clone(HW_THEME), templateId: 'spec-test', templateName: 'Spec Test' },
      logoBytes: Buffer.from([0x89, 0x50, 0x4e, 0x47]),
      meta: { slideCount: 3, masters: 1, layouts: 2, slideSize: { cx: 12192000, cy: 6858000, widthIn: 13.333, heightIn: 7.5 } },
      media: [],
      sourceBytes: Buffer.from('pptx-bytes'),
      layoutsInventory: { slideSize: { cx: 12192000, cy: 6858000, widthIn: 13.333, heightIn: 7.5 }, masters: [], counts: { masters: 0, layouts: 0 } },
      markdown: '# Spec Test\n\n## 色板\n',
    },
    config,
  )
  assert(summary.hasLayouts === true, 'hasLayouts flag missing')
  assert(summary.hasSpec === true, 'hasSpec flag missing')
  for (const name of ['theme.json', 'layouts.json', 'template.md', 'source.pptx', 'meta.json', 'logo.png']) {
    assert(existsSync(join(templateDir, name)), `missing ${name}`)
  }
  const read = store.readTemplate('spec-test', config)
  assert(read.layouts !== null, 'layouts.json not read back')
  assert(read.markdown !== null && read.markdown.startsWith('# Spec Test'), 'spec not read back')
  assert(read.sourcePath !== null, 'source copy not exposed')
  return '6 files, layouts + spec + source reusable'
})

const failed = results.filter((entry) => !entry.ok)
console.log(`\n${results.length - failed.length}/${results.length} deck-bundle checks passed`)
if (failed.length > 0) process.exitCode = 1
