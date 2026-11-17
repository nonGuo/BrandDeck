/**
 * Tests for the fallback layout library — the reference-deck vocabulary.
 *
 * WHY: `example/bussiness.pptx` and `example/technology.pptx` are reference *decks*, not
 * template layout libraries. What they teach is composition: a kicker line, a rule under
 * the header, tinted panels, oversized figures, ranked rows. Those are the shapes this
 * plugin now falls back to when the input template only defines a skin.
 *
 * The tests below pin the two things that are easy to get wrong and invisible until a
 * customer sees the deck: which layout a page is *given*, and the geometry inside it.
 *
 * Run: node test/layout-library-test.mjs
 */
import { CANVAS, innerBox } from '../src/schema.js'
import { composeDeck } from '../src/generate.js'
import { makeDefaultTheme } from '../src/defaults.js'

const THEME = makeDefaultTheme()

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

function assert(condition, message) {
  if (!condition) throw new Error(message)
}

const stable = (value) => JSON.stringify(value)
const kinds = (deck) => deck.slides.map((slide) => slide.kind)
const texts = (slide) => slide.elements.filter((element) => element.type === 'text').map((element) => element.content)

/** Build an outline from `[chapter, [bullets]]` pairs. */
function outline(title, pages) {
  const lines = [`# ${title}`, '']
  for (const [chapter, bullets] of pages) {
    lines.push('## 页面', `章节：${chapter}`)
    for (const bullet of bullets) lines.push(`- ${bullet}`)
    lines.push('')
  }
  return lines.join('\n')
}

/* ------------------------------------------------------------------ *
 * Page grammar
 * ------------------------------------------------------------------ */

check('1a · every content page carries a kicker naming its chapter', () => {
  const { deck } = composeDeck({
    outline: outline('章节检验', [
      ['甲', ['口径：中汽协口径统计', '时间：2025 全年', '边界：不含散件']],
      ['乙', ['内容甲：这是第一条说明', '内容乙：这是第二条说明', '内容丙：这是第三条说明']],
    ]),
    theme: THEME,
    style: { maxSlides: 10 },
  })
  const content = deck.slides.filter((slide) => !['cover', 'section', 'toc'].includes(slide.kind))
  assert(content.length > 0, 'no content page produced')
  for (const slide of content) {
    const kicker = slide.elements.find((element) => element.id.endsWith('-kicker-1'))
    assert(kicker !== undefined, `page "${slide.title}" has no kicker`)
    assert(/^(0?\d)\s·\s/.test(kicker.content), `kicker "${kicker.content}" is not "NN · chapter"`)
  }
  return `${content.length} content pages with kickers`
})

check('1b · the header closes with a rule, and the accent is a segment of it', () => {
  const { deck } = composeDeck({
    outline: outline('页眉检验', [['甲', ['口径：内容说明一', '时间：内容说明二', '边界：内容说明三']]]),
    theme: THEME,
    style: { maxSlides: 8 },
  })
  // Target a *content* page: section dividers and the compare layout also draw `-rule-1`.
  const slide = deck.slides.find(
    (entry) =>
      !['cover', 'section', 'toc'].includes(entry.kind) &&
      entry.elements.some((element) => element.id.endsWith('-rule-1')),
  )
  assert(slide !== undefined, 'no page drew a header rule')
  const rule = slide.elements.find((element) => element.id.endsWith('-rule-1'))
  const accent = slide.elements.find((element) => element.id.endsWith('-accent-1'))
  assert(accent !== undefined, 'no accent tick on the rule')
  // Adjacent, not overlapping: the overlap gate in repairDeck rejects touching shapes.
  assert(accent.x + accent.w <= rule.x + 0.5, `accent (${accent.x}+${accent.w}) overlaps rule at ${rule.x}`)
  assert(Math.abs(accent.y - rule.y) <= 1, `accent y ${accent.y} is not on the rule y ${rule.y}`)
  return `accent 0..${accent.w}px, rule ${rule.w}px`
})

/* ------------------------------------------------------------------ *
 * Layout selection
 * ------------------------------------------------------------------ */

check('2a · labelled points become panels, not a bullet list', () => {
  const { deck } = composeDeck({
    outline: outline('分栏检验', [
      ['甲', ['口径：中汽协口径，另有乘联会口径并列', '时间：2025 年出口数据', '结构：按车型与动力拆分', '边界：不含二手车出口']],
    ]),
    theme: THEME,
    style: { maxSlides: 8 },
  })
  const page = deck.slides.find((slide) => slide.kind === 'columns')
  assert(page !== undefined, `expected a columns page, got ${stable(kinds(deck))}`)
  const panels = page.elements.filter((element) => element.id.endsWith('-panel-1'))
  assert(panels.length === 1, `expected one panel (not the count), got ${panels.length}`)
  const all = page.elements.filter((element) => element.type === 'shape' && /-panel-\d+$/.test(element.id))
  assert(all.length === 4, `expected 4 panels, got ${all.length}`)
  // Panels share geometry: same height, same top, equal widths.
  const tops = new Set(all.map((panel) => panel.y))
  const heights = new Set(all.map((panel) => panel.h))
  assert(tops.size === 1 && heights.size === 1, `panels disagree: tops ${stable([...tops])} heights ${stable([...heights])}`)
  return `4 panels of ${all[0].w}×${all[0].h}px`
})

check('2b · a page that is nothing but labelled rows becomes a ranked list', () => {
  const { deck } = composeDeck({
    outline: outline('排行检验', [
      ['甲', ['墨西哥 — 62.52 万辆', '俄罗斯 — 58.27 %', '阿联酋 — 37 个', '巴西 — 30.1 天', '英国 — 28.4 倍']],
    ]),
    theme: THEME,
    style: { maxSlides: 8 },
  })
  const page = deck.slides.find((slide) => slide.kind === 'rankList')
  assert(page !== undefined, `expected a rankList page, got ${stable(kinds(deck))}`)
  const highlight = page.elements.find((element) => element.id.endsWith('-highlight-1'))
  assert(highlight !== undefined, 'the first row is not marked')
  const values = page.elements.filter((element) => element.id.includes('-value-')).map((element) => element.content)
  assert(values.some((value) => value.startsWith('62.52')), `values missing: ${stable(values)}`)
  return `${values.length} ranked rows, first row marked`
})

check('2c · two bare figures become a figures page, and the numbers stay verbatim', () => {
  const { deck } = composeDeck({
    outline: outline('数字检验', [['甲', ['2025 年新能源汽车出口 261.5 万辆', '同期汽车总出口 343 万辆']]]),
    theme: THEME,
    style: { maxSlides: 8 },
  })
  const page = deck.slides.find((slide) => slide.kind === 'figures')
  assert(page !== undefined, `expected a figures page, got ${stable(kinds(deck))}`)
  const figures = page.elements.filter((element) => element.id.endsWith('-figure-1') || element.id.endsWith('-figure-2'))
  assert(figures.length === 2, `expected 2 figures, got ${figures.length}`)
  assert(figures[0].content === '261.5 万辆', `figure 1 was "${figures[0].content}"`)
  assert(figures[1].content === '343 万辆', `figure 2 was "${figures[1].content}"`)
  // The figure must be visually dominant, or it is just a bold line.
  assert(figures[0].style.fontSize >= 60, `figure size ${figures[0].style.fontSize} is not dominant`)
  return `2 figures at ${figures[0].style.fontSize}px`
})

check('2d · three pages of the same shape do not all draw the same way', () => {
  const labelled = (chapter) => [
    chapter,
    ['口径：中汽协口径统计说明', '时间：2025 年全年数据', '结构：按车型与动力拆分', '边界：不含散件出口'],
  ]
  const { deck } = composeDeck({
    outline: outline('去重检验', [labelled('甲'), labelled('乙'), labelled('丙'), labelled('丁')]),
    theme: THEME,
    style: { maxSlides: 12 },
  })
  const content = kinds(deck).filter((kind) => !['cover', 'section', 'toc'].includes(kind))
  let run = 1
  let worst = 1
  for (let index = 1; index < content.length; index += 1) {
    run = content[index] === content[index - 1] ? run + 1 : 1
    worst = Math.max(worst, run)
  }
  assert(worst <= 2, `three pages in a row share a layout: ${stable(content)}`)
  return content.join(' → ')
})

/* ------------------------------------------------------------------ *
 * Geometry
 * ------------------------------------------------------------------ */

check('3a · panels stay inside the frame and never overlap each other', () => {
  const { deck } = composeDeck({
    outline: outline('几何检验', [['甲', ['口径：说明一足够长以形成两行文字内容', '时间：说明二足够长以形成两行文字内容', '边界：说明三足够长以形成两行文字内容']]]),
    theme: THEME,
    style: { maxSlides: 8 },
  })
  const box = innerBox(THEME)
  const page = deck.slides.find((slide) => slide.kind === 'columns')
  assert(page !== undefined, 'no columns page')
  const panels = page.elements.filter((element) => element.type === 'shape' && /-panel-\d+$/.test(element.id))
  for (const panel of panels) {
    assert(panel.x >= box.x - 0.5, `panel starts left of the frame: ${panel.x} < ${box.x}`)
    assert(panel.x + panel.w <= box.x + box.w + 0.5, `panel runs past the frame: ${panel.x + panel.w}`)
    assert(panel.y >= box.y - 0.5, `panel starts above the frame: ${panel.y}`)
    assert(panel.y + panel.h <= box.y + box.h + 0.5, `panel runs below the frame: ${panel.y + panel.h}`)
  }
  for (let index = 1; index < panels.length; index += 1) {
    const left = panels[index - 1]
    assert(left.x + left.w <= panels[index].x + 0.5, `panels ${index - 1}/${index} overlap`)
  }
  return `${panels.length} panels inside ${box.w}×${box.h}`
})

check('3b · the layout survives a deck with no chapters (no kicker, no divider)', () => {
  const { deck } = composeDeck({
    outline: '# 无章节\n\n## 一页\n- 口径：说明一足够长\n- 时间：说明二足够长\n- 边界：说明三足够长\n',
    theme: THEME,
    style: { maxSlides: 8 },
  })
  assert(!kinds(deck).includes('section'), 'invented a chapter divider')
  assert(!kinds(deck).includes('toc'), 'invented a contents page')
  const content = deck.slides.find((slide) => !['cover', 'toc', 'section'].includes(slide.kind))
  const kicker = content.elements.find((element) => element.id.endsWith('-kicker-1'))
  assert(kicker === undefined, 'drew a kicker for a deck with no chapters')
  return kinds(deck).join(' → ')
})

check('3c · a page never exceeds the canvas', () => {
  const long = Array.from({ length: 6 }, (_, index) => `要点 ${index + 1}：这是一条比较长的说明文字，用来撑满一行以上`)
  const { deck } = composeDeck({
    outline: outline('溢出检验', [['甲', long]]),
    theme: THEME,
    style: { maxSlides: 8, density: 'dense' },
  })
  for (const slide of deck.slides) {
    for (const element of slide.elements) {
      assert(element.x >= 0 && element.y >= 0, `element ${element.id} at negative origin`)
      assert(element.x + element.w <= CANVAS.width + 0.5, `element ${element.id} past the right edge`)
      assert(element.y + element.h <= CANVAS.height + 0.5, `element ${element.id} past the bottom edge`)
    }
  }
  return `${deck.slides.length} pages inside the canvas`
})

/* ------------------------------------------------------------------ *
 * Summary
 * ------------------------------------------------------------------ */

const failed = results.filter((result) => !result.ok)
console.log(`\n${results.length} checks, ${results.length - failed.length} passed, ${failed.length} failed`)
if (failed.length > 0) {
  console.log(`  FAILED: ${failed.map((result) => result.name).join('; ')}`)
  process.exitCode = 1
} else {
  console.log('ALL PASS')
}
