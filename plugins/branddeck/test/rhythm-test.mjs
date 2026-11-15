/**
 * Rhythm test — the two "pacing" behaviours and their off-switches.
 *
 * The rule under test is a judgement call as much as a layout rule:
 *   · adjacent-layout variety may only swap between renderings of the SAME content
 *     (never invent a sequence or a comparison the author did not write);
 *   · a breathing page appears only when it is explicitly enabled AND the material
 *     already contains a short statement, and never as the last page;
 *   · neither feature may add warnings (clean copy in ⇒ zero warnings out), which is
 *     the contract the rest of the generator's checks rely on.
 *
 * Run: node test/rhythm-test.mjs
 */
import { parseSourceToOutline, composeDeck } from '../src/generate.js'
import { CardDeckSchema } from '../src/schema.js'
import { normalizeStyle, DEFAULT_STYLE } from '../src/styles.js'

const THEME = {
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
  brandAssets: { icons: {} },
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

const md = (slides) =>
  ['# 节奏测试', '', ...slides.flatMap(([title, bullets]) => [`## ${title}`, ...bullets.map((b) => `- ${b}`), ''])].join('\n')

const LISTY_FOUR = [
  ['第一页', ['甲', '乙', '丙']],
  ['第二页', ['丁', '戊', '己']],
  // Two bullets: the only slide a *neutral* two-column re-render can apply to.
  ['第三页', ['庚', '辛']],
  ['第四页', ['壬', '癸']],
]

const kinds = (deck) => deck.slides.map((slide) => slide.kind)

/* 1. Anti-repetition ------------------------------------------------------- */
check('adjacent variety re-renders a repeated card run as plain lists', () => {
  // The ONLY safe swap: cards -> bullets. Same content, different shape. A compare
  // swap was removed because a comparison layout splits "现状：庚" into a heading plus
  // an item, i.e. it rewrites the user's words — see the comment in generate.js.
  const bulletStyle = { bulletLayout: 'cards' }
  const on = composeDeck({
    outline: parseSourceToOutline(md(LISTY_FOUR)),
    theme: THEME,
    style: { ...bulletStyle },
    deckId: 'r1',
  })
  const off = composeDeck({
    outline: parseSourceToOutline(md(LISTY_FOUR)),
    theme: THEME,
    style: { ...bulletStyle, varyLayout: false },
    deckId: 'r2',
  })
  const onKinds = kinds(on.deck)
  const offKinds = kinds(off.deck)
  assert(offKinds.filter((kind) => kind === 'cards').length > onKinds.filter((kind) => kind === 'cards').length, `variety did not reduce cards: on=${onKinds.join(' → ')} off=${offKinds.join(' → ')}`)
  assert(onKinds.includes('bullets'), `expected a plain list somewhere: ${onKinds.join(' → ')}`)
  // The copy must be identical between the two runs.
  const strip = (deck) =>
    deck.slides.map((slide) =>
      slide.elements.filter((element) => element.type === 'text').map((element) => element.content).join('|'),
    )
  assert(JSON.stringify(strip(on.deck)) === JSON.stringify(strip(off.deck)), 'variety changed the copy')
  return `on=${onKinds.join(' → ')} | off=${offKinds.join(' → ')}`
})

check('variety never fabricates a sequence or a comparison', () => {
  // Four 3-bullet slides: nothing here is a pair, so no compare may be invented.
  const plain = [
    ['一', ['甲', '乙', '丙']],
    ['二', ['丁', '戊', '己']],
    ['三', ['庚', '辛', '壬']],
    ['四', ['癸', '子', '丑']],
  ]
  const { deck } = composeDeck({ outline: parseSourceToOutline(md(plain)), theme: THEME, style: {}, deckId: 'r3' })
  assert(!kinds(deck).includes('compare'), `invented a comparison: ${kinds(deck).join(' → ')}`)
  assert(!kinds(deck).includes('timeline'), 'invented a sequence')
  assert(!kinds(deck).includes('stats'), 'invented statistics')
  return kinds(deck).join(' → ')
})

check('variety adds no warnings (clean copy ⇒ zero warnings)', () => {
  const { warnings } = composeDeck({ outline: parseSourceToOutline(md(LISTY_FOUR)), theme: THEME, style: {}, deckId: 'r4' })
  assert(warnings.length === 0, `warnings: ${warnings.join(' | ')}`)
})

/* 2. Breathing pages ------------------------------------------------------- */
const SHORT = [
  ['一', ['成本下降 70%', '其它甲', '其它乙']],
  ['二', ['内容甲', '内容乙', '内容丙']],
  ['三', ['内容丁', '内容戊', '内容己']],
  ['四', ['内容庚', '内容辛', '内容壬']],
]

check('breather is off by default and on only when asked', () => {
  const base = composeDeck({ outline: parseSourceToOutline(md(SHORT)), theme: THEME, style: {}, deckId: 'b1' })
  assert(!kinds(base.deck).includes('breather'), 'breather appeared without being enabled')
  assert(DEFAULT_STYLE.breather === false, 'DEFAULT_STYLE.breather should be false')

  const on = composeDeck({
    outline: parseSourceToOutline(md(SHORT)),
    theme: THEME,
    style: { breather: true, breatherEvery: 3 },
    deckId: 'b2',
  })
  assert(kinds(on.deck).includes('breather'), `no breather: ${kinds(on.deck).join(' → ')}`)
  assert(CardDeckSchema.safeParse(on.deck).success, 'deck with a breather fails CardDeckSchema')
  return kinds(on.deck).join(' → ')
})

check('a breather carries a short, already-present statement (never invented)', () => {
  const { deck } = composeDeck({
    outline: parseSourceToOutline(md(SHORT)),
    theme: THEME,
    style: { breather: true, breatherEvery: 3 },
    deckId: 'b3',
  })
  const breather = deck.slides.find((slide) => slide.kind === 'breather')
  assert(breather !== undefined, 'no breather slide')
  const texts = breather.elements.filter((element) => element.type === 'text').map((element) => element.content)
  assert(texts[0] === '70%', `headline should be the existing number, got ${JSON.stringify(texts[0])}`)
  // The number must exist in the source material verbatim.
  assert(md(SHORT).includes(texts[0]), 'headline was not present in the source')
  return texts.join(' / ')
})

check('long sentences are skipped rather than split into nonsense', () => {
  const longOnly = [
    ['一', ['出片时间从 5 天降到 4 小时', '其它甲', '其它乙']],
    ['二', ['内容甲', '内容乙', '内容丙']],
    ['三', ['内容丁', '内容戊', '内容己']],
    ['四', ['内容庚', '内容辛', '内容壬']],
  ]
  const { deck } = composeDeck({
    outline: parseSourceToOutline(md(longOnly)),
    theme: THEME,
    style: { breather: true, breatherEvery: 3 },
    deckId: 'b4',
  })
  assert(!kinds(deck).includes('breather'), 'a long sentence produced a breather page')
  return 'skipped as expected'
})

check('a breather is never the last page before the closing slide', () => {
  // Exactly three content slides with `breatherEvery: 3` would land on the last one.
  const three = [
    ['一', ['成本下降 70%', '甲', '乙']],
    ['二', ['丙', '丁', '戊']],
    ['三', ['己', '庚', '辛']],
  ]
  const { deck } = composeDeck({
    outline: parseSourceToOutline(md(three)),
    theme: THEME,
    style: { breather: true, breatherEvery: 3 },
    deckId: 'b5',
  })
  const list = kinds(deck)
  const breatherIndex = list.indexOf('breather')
  assert(breatherIndex === -1 || breatherIndex < list.length - 1, `breather placed last: ${list.join(' → ')}`)
  return list.join(' → ')
})

/* 3. The style contract carries the flags ---------------------------------- */
check('the style contract preserves the rhythm flags', () => {
  const normalized = normalizeStyle({ name: 'x', varyLayout: false, breather: true, breatherEvery: 4 })
  assert(normalized.varyLayout === false, `varyLayout dropped: ${normalized.varyLayout}`)
  assert(normalized.breather === true, `breather dropped: ${normalized.breather}`)
  assert(normalized.breatherEvery === 4, `breatherEvery dropped: ${normalized.breatherEvery}`)
  // Out-of-range values clamp instead of throwing.
  assert(normalizeStyle({ name: 'x', breatherEvery: 99 }).breatherEvery <= 12, 'breatherEvery not clamped')
  return 'varyLayout / breather / breatherEvery survive normalization'
})

const failed = results.filter((entry) => !entry.ok)
console.log(`\n${results.length - failed.length}/${results.length} rhythm checks passed`)
if (failed.length > 0) process.exitCode = 1
