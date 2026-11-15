/**
 * Test for src/outline.js — the Markdown "文字版本" twin of a CardDeck.
 *
 * Run: C:\Users\Administrator\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe test/outline-test.mjs
 */
import { HW_DEMO_DECK } from './fixtures/hw-demo-deck.mjs'
import { deckToOutline, applyOutline } from '../src/outline.js'
import { CANVAS, SAFE_AREA, CardDeckSchema, innerBox } from '../src/schema.js'

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
    results.push({ name, ok: false, error })
    console.log(`FAIL ${name}: ${error?.message ?? error}`)
  }
}
const assert = (condition, message) => {
  if (!condition) throw new Error(message)
}
const clone = (value) => structuredClone(value)
const stable = (value) => JSON.stringify(value)
const deepEqual = (a, b) => stable(a) === stable(b)

const deck = clone(HW_DEMO_DECK)
const theme = clone(HW_THEME)
const outline = deckToOutline(deck, theme)
const textElements = (aDeck) =>
  aDeck.slides.flatMap((slide) => slide.elements.filter((element) => element.type === 'text'))
const textCount = textElements(deck).length
const slideById = (aDeck, slideId) => aDeck.slides.find((slide) => slide.slideId === slideId)
const elementById = (aDeck, slideId, elementId) =>
  slideById(aDeck, slideId)?.elements.find((element) => element.id === elementId)
const overlaps = (a, b) =>
  a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y

/** Remove a whole `### … {#id}` block (heading + content) from the outline. */
function removeBlock(markdown, elementId) {
  const lines = markdown.split('\n')
  const out = []
  let skipping = false
  for (const line of lines) {
    const startsBlock = line.startsWith('### ') || line.startsWith('## ') || line.startsWith('<!--')
    if (line.startsWith('### ') && line.includes(`{#${elementId}}`)) {
      skipping = true
      continue
    }
    if (skipping && (startsBlock || /^备注[:：]/.test(line))) skipping = false
    if (!skipping) out.push(line)
  }
  return out.join('\n')
}

/** Remove a whole `## … {#slideId}` section from the outline. */
function removeSlideSection(markdown, slideId) {
  const lines = markdown.split('\n')
  const out = []
  let skipping = false
  for (const line of lines) {
    if (line.startsWith('## ') && line.includes(`{#${slideId}}`)) {
      skipping = true
      continue
    }
    if (skipping && line.startsWith('## ')) skipping = false
    if (!skipping) out.push(line)
  }
  return out.join('\n')
}

/** Insert Markdown text at the end of one slide section. */
function appendToSlide(markdown, slideId, block) {
  const lines = markdown.split('\n')
  let start = -1
  for (let i = 0; i < lines.length; i += 1) {
    if (lines[i].startsWith('## ') && lines[i].includes(`{#${slideId}}`)) {
      start = i
      break
    }
  }
  assert(start !== -1, `slide ${slideId} not found in outline`)
  let end = lines.length
  for (let i = start + 1; i < lines.length; i += 1) {
    if (lines[i].startsWith('## ')) {
      end = i
      break
    }
  }
  // Keep the trailing blank line that separates this section from the next one.
  let insertAt = end
  while (insertAt > start + 1 && lines[insertAt - 1].trim() === '') insertAt -= 1
  const inserted = ['', ...block.split('\n')]
  lines.splice(insertAt, 0, ...inserted)
  return lines.join('\n')
}

/** Replace the single `备注：…` line inside one slide section. */
function replaceNotes(markdown, slideId, replacement) {
  const lines = markdown.split('\n')
  let start = -1
  for (let i = 0; i < lines.length; i += 1) {
    if (lines[i].startsWith('## ') && lines[i].includes(`{#${slideId}}`)) {
      start = i
      break
    }
  }
  assert(start !== -1, `slide ${slideId} not found in outline`)
  for (let i = start + 1; i < lines.length && !lines[i].startsWith('## '); i += 1) {
    if (/^备注[:：]/.test(lines[i])) {
      lines[i] = replacement
      return lines.join('\n')
    }
  }
  throw new Error(`no 备注 line in slide ${slideId}`)
}

check('deck -> outline -> deck round-trips every text element and id', () => {
  const applied = applyOutline(outline, deck, theme)
  assert(applied.warnings.length === 0, `unexpected warnings: ${applied.warnings.join(' | ')}`)
  assert(deepEqual(applied.deck, deck), 'round-trip changed the deck')
  assert(
    applied.changes.updated === 0 &&
      applied.changes.added === 0 &&
      applied.changes.removed === 0 &&
      applied.changes.untouched === textCount,
    `changes: ${JSON.stringify(applied.changes)}`,
  )
  const before = textElements(deck)
  const after = textElements(applied.deck)
  assert(before.length === after.length, `text element count ${before.length} -> ${after.length}`)
  for (let i = 0; i < before.length; i += 1) {
    assert(before[i].id === after[i].id, `id ${before[i].id} -> ${after[i].id}`)
    assert(before[i].content === after[i].content, `content drift on ${before[i].id}`)
  }
  return `${after.length} text elements, ${before.length} ids`
})

check('editing one element changes only that element', () => {
  const oldText = '从企业模板到原生可编辑演示文稿，一条链路打通设计与内容'
  const newText = '新的副标题：AI 原生可编辑，品牌一致'
  assert(outline.includes(oldText), 'fixture copy missing from the outline')
  const edited = outline.replace(
    `### 副标题 {#cover-subtitle}\n${oldText}`,
    `### 副标题 {#cover-subtitle}\n${newText}`,
  )
  assert(edited !== outline, 'markdown edit did not apply')

  const applied = applyOutline(edited, deck, theme)
  const target = elementById(applied.deck, 'cover', 'cover-subtitle')
  assert(target.content === newText, `content: ${target.content}`)
  assert(
    deepEqual(target, { ...elementById(deck, 'cover', 'cover-subtitle'), content: newText }),
    'geometry/style of the edited element changed',
  )
  for (const neighbour of ['cover-kicker', 'cover-title', 'cover-meta']) {
    assert(
      deepEqual(elementById(applied.deck, 'cover', neighbour), elementById(deck, 'cover', neighbour)),
      `neighbour ${neighbour} changed`,
    )
  }
  for (const slide of deck.slides.slice(1)) {
    assert(
      deepEqual(slideById(applied.deck, slide.slideId), slide),
      `unrelated slide ${slide.slideId} changed`,
    )
  }
  let differing = 0
  for (const before of textElements(deck)) {
    const after = elementById(applied.deck, findSlideId(deck, before.id), before.id)
    if (before.content !== after.content) differing += 1
  }
  assert(differing === 1, `${differing} text elements changed`)
  assert(
    applied.changes.updated === 1 &&
      applied.changes.added === 0 &&
      applied.changes.removed === 0 &&
      applied.changes.untouched === textCount - 1,
    `changes: ${JSON.stringify(applied.changes)}`,
  )
  assert(applied.warnings.length === 0, `warnings: ${applied.warnings.join(' | ')}`)
  return `1 updated, ${textCount - 1} untouched`
})

// Helper used above (declared after use is fine for function declarations).
function findSlideId(aDeck, elementId) {
  for (const slide of aDeck.slides) {
    if (slide.elements.some((element) => element.id === elementId)) return slide.slideId
  }
  throw new Error(`element ${elementId} not found`)
}

check('adding a 正文 block creates one safe, non-overlapping element', () => {
  const withNew = appendToSlide(outline, 'challenge', '### 正文 {#new-1}\n新增的正文段落')
  const applied = applyOutline(withNew, deck, theme)
  assert(applied.changes.added === 1, `added: ${applied.changes.added}`)
  assert(applied.warnings.some((w) => w.includes('added element challenge-t1 (auto-layout)')), `warnings: ${applied.warnings.join(' | ')}`)

  const slide = slideById(applied.deck, 'challenge')
  const added = slide.elements.filter(
    (element) => element.type === 'text' && element.content === '新增的正文段落',
  )
  assert(added.length === 1, `${added.length} matching elements`)
  const element = added[0]
  assert(element.id === 'challenge-t1', `id: ${element.id}`)
  assert(element.role === 'body', `role: ${element.role}`)
  assert(element.style.fontSize === 20, `fontSize: ${element.style.fontSize}`)

  const box = innerBox(theme)
  assert(element.x >= box.x && element.y >= box.y, `outside top-left: ${element.x},${element.y}`)
  assert(element.x + element.w <= box.x + box.w, 'exceeds safe width')
  assert(element.y + element.h <= box.y + box.h, 'exceeds safe height')
  assert(element.x + element.w <= CANVAS.width - SAFE_AREA.right, 'exceeds canvas width')

  for (const other of slide.elements) {
    if (other === element) continue
    if (other.type !== 'text') continue
    assert(!overlaps(element, other), `overlaps text box ${other.id}`)
  }
  // Every pre-existing element keeps its content and geometry.
  for (const before of slideById(deck, 'challenge').elements) {
    const after = elementById(applied.deck, 'challenge', before.id)
    assert(deepEqual(after, before), `element ${before.id} changed`)
  }
  for (const other of deck.slides) {
    if (other.slideId === 'challenge') continue
    assert(deepEqual(slideById(applied.deck, other.slideId), other), `slide ${other.slideId} changed`)
  }
  return `${element.id} at ${element.x},${element.y} ${element.w}x${element.h}`
})

check('repeated additions get distinct ids and do not overlap each other', () => {
  const withNew = appendToSlide(
    outline,
    'end',
    '### 正文 {#first}\n第一段新增正文\n\n### 注释 {#second}\n第二段新增注释',
  )
  const applied = applyOutline(withNew, deck, theme)
  assert(applied.changes.added === 2, `added: ${applied.changes.added}`)
  const slide = slideById(applied.deck, 'end')
  const first = elementById(applied.deck, 'end', 'end-t1')
  const second = elementById(applied.deck, 'end', 'end-t2')
  assert(first !== undefined && second !== undefined, 'deterministic ids end-t1/end-t2 missing')
  assert(first.role === 'body', `first role: ${first.role}`)
  assert(second.role === 'caption', `second role: ${second.role}`)
  assert(second.style.fontSize === 14, `caption fontSize: ${second.style.fontSize}`)
  assert(!overlaps(first, second), 'added elements overlap each other')
  for (const element of [first, second]) {
    for (const other of slide.elements) {
      if (other === element || other.type !== 'text') continue
      assert(!overlaps(element, other), `addition ${element.id} overlaps ${other.id}`)
    }
  }
  return `${first.id}@${first.y} ${second.id}@${second.y}`
})

check('removing a section keeps the element by default and prunes on demand', () => {
  const withoutFootnote = removeBlock(outline, 'challenge-footnote')
  assert(!withoutFootnote.includes('challenge-footnote'), 'section not removed from markdown')
  assert(!withoutFootnote.includes('数据来源'), 'block content leaked into the markdown')

  const kept = applyOutline(withoutFootnote, deck, theme)
  const keptElement = elementById(kept.deck, 'challenge', 'challenge-footnote')
  assert(keptElement !== undefined, 'element removed without prune')
  assert(deepEqual(keptElement, elementById(deck, 'challenge', 'challenge-footnote')), 'kept element changed')
  assert(kept.changes.removed === 0, `removed: ${kept.changes.removed}`)
  assert(kept.warnings.length === 0, `warnings: ${kept.warnings.join(' | ')}`)

  const pruned = applyOutline(withoutFootnote, deck, theme, { prune: true })
  assert(elementById(pruned.deck, 'challenge', 'challenge-footnote') === undefined, 'element not pruned')
  assert(pruned.changes.removed === 1, `removed: ${pruned.changes.removed}`)
  assert(pruned.warnings.some((w) => w.includes('slide 3 · pruned element "challenge-footnote"')), `warnings: ${pruned.warnings.join(' | ')}`)
  assert(
    slideById(pruned.deck, 'challenge').elements.length ===
      slideById(deck, 'challenge').elements.length - 1,
    'unexpected element count after prune',
  )

  // A slide section that disappears entirely is also "absent from the markdown".
  const withoutGrowth = removeSlideSection(outline, 'growth')
  assert(!withoutGrowth.includes('{#growth}'), 'slide section not removed from markdown')
  const growthKept = applyOutline(withoutGrowth, deck, theme)
  assert(deepEqual(slideById(growthKept.deck, 'growth'), slideById(deck, 'growth')), 'unmentioned slide changed')
  const growthPruned = applyOutline(withoutGrowth, deck, theme, { prune: true })
  const growthText = slideById(growthPruned.deck, 'growth').elements.filter(
    (element) => element.type === 'text',
  ).length
  assert(growthText === 0, `${growthText} text elements left on the pruned slide`)
  assert(growthPruned.changes.removed === 7, `removed: ${growthPruned.changes.removed}`)
  assert(growthPruned.warnings.every((w) => w.includes('slide 4 · pruned element')), `warnings: ${growthPruned.warnings.join(' | ')}`)
  const prunedDeck = CardDeckSchema.safeParse(growthPruned.deck)
  assert(prunedDeck.success, JSON.stringify(prunedDeck.error?.issues?.slice(0, 3)))
  return 'element kept by default, pruned on demand; whole slide section too'
})

check('备注：/ 备注: / Notes: map to speakerNotes', () => {
  const zhFull = replaceNotes(outline, 'cover', '备注：改过的备注文案')
  const appliedZh = applyOutline(zhFull, deck, theme)
  assert(slideById(appliedZh.deck, 'cover').speakerNotes === '改过的备注文案', 'full-width 备注：')

  const zhHalf = replaceNotes(outline, 'cover', '备注: 半角冒号备注')
  const appliedHalf = applyOutline(zhHalf, deck, theme)
  assert(slideById(appliedHalf.deck, 'cover').speakerNotes === '半角冒号备注', 'half-width 备注:')

  const en = replaceNotes(outline, 'end', 'Notes: English speaker notes')
  const appliedEn = applyOutline(en, deck, theme)
  assert(slideById(appliedEn.deck, 'end').speakerNotes === 'English speaker notes', 'Notes:')
  assert(
    slideById(appliedEn.deck, 'cover').speakerNotes === slideById(deck, 'cover').speakerNotes,
    'untouched slide notes drifted',
  )
  return '3 note syntaxes'
})

check('output is byte-stable and has no trailing whitespace', () => {
  const first = deckToOutline(deck, theme)
  const second = deckToOutline(deck, theme)
  assert(first === second, 'two calls differ')
  assert(first.endsWith('\n') && !first.endsWith('\n\n'), 'expected exactly one trailing newline')
  for (const [index, line] of first.split('\n').entries()) {
    assert(!/[ \t]+$/.test(line), `trailing whitespace on line ${index + 1}: ${JSON.stringify(line)}`)
  }
  const applied = applyOutline(first, deck, theme)
  assert(deckToOutline(applied.deck, theme) === first, 'projection is not a fixed point after apply')
  assert(!first.includes('\r'), 'expected LF line endings')
  return `${first.split('\n').length} lines, ${Buffer.byteLength(first, 'utf8')} bytes`
})

check('English role labels round-trip too', () => {
  const en = deckToOutline(deck, theme, { labels: 'en' })
  assert(en.includes('### Title {#cover-title}'), 'missing Title label')
  assert(en.includes('### Caption {#cover-kicker}'), 'missing Caption label')
  assert(en.includes('### Subtitle {#cover-subtitle}'), 'missing Subtitle label')
  assert(en.includes('### Body {#cover-meta}') === false, 'caption emitted as Body')
  const applied = applyOutline(en, deck, theme)
  assert(deepEqual(applied.deck, deck), 'english-label round-trip changed the deck')
  assert(applied.warnings.length === 0, `warnings: ${applied.warnings.join(' | ')}`)
  return 'title/subtitle/body/caption'
})

check('applied decks remain schema-valid after add + prune + notes', () => {
  let edited = appendToSlide(outline, 'challenge', '### 正文 {#new-1}\n新增的正文段落')
  edited = appendToSlide(edited, 'growth', '### 标题 {#new-title}\n新增的标题')
  edited = removeBlock(edited, 'agenda-num-3')
  edited = replaceNotes(edited, 'end', '备注：新增的收尾页备注')
  const applied = applyOutline(edited, deck, theme, { prune: true })
  const parsed = CardDeckSchema.safeParse(applied.deck)
  assert(parsed.success, JSON.stringify(parsed.error?.issues?.slice(0, 3), null, 2))
  assert(applied.changes.added === 2, `added: ${applied.changes.added}`)
  assert(applied.changes.removed === 1, `removed: ${applied.changes.removed}`)
  assert(
    slideById(applied.deck, 'end').speakerNotes === '新增的收尾页备注',
    'notes not applied',
  )
  return `${applied.changes.added} added, ${applied.changes.removed} pruned, schema ok`
})

check('garbage input warns and leaves the deck unchanged, never throws', () => {
  const inputs = [
    '',
    '这是一个没有标题也没有任何卡片的随机段落。\n第二行同样只是散文。\n',
    '# 标题\n### 标题 {#cover-title}\n孤立内容\n## 没有 id 的标题\n正文\n### 正文 {#new-1}\n没有父级\n',
  ]
  const notes = []
  for (const input of inputs) {
    const before = clone(deck)
    let applied
    try {
      applied = applyOutline(input, deck, theme, { prune: true })
    } catch (error) {
      throw new Error(`threw on ${JSON.stringify(input)}: ${error?.message ?? error}`)
    }
    assert(deepEqual(applied.deck, before), `deck changed for ${JSON.stringify(input)}`)
    assert(applied.warnings.length > 0, `expected warnings for ${JSON.stringify(input)}`)
    const { updated, added, removed } = applied.changes
    assert(updated === 0 && added === 0 && removed === 0, `changes: ${JSON.stringify(applied.changes)}`)
    notes.push(applied.warnings.length)
  }
  return `3 inputs, ${notes.join('/')} warnings`
})

const failures = results.filter((result) => !result.ok)
console.log(
  `\n${results.length - failures.length}/${results.length} checks passed` +
    (failures.length === 0 ? '' : `, ${failures.length} failed`),
)
if (failures.length > 0) process.exitCode = 1
