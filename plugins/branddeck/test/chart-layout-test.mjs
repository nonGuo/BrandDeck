/**
 * Tests for the two mechanisms borrowed from the reference pipeline:
 *
 *   1. the **type scale** — every role as a ratio of the body size (README §6.1)
 *   2. **chart routing** — values that parse into one unit become a native chart, chosen by
 *      the information relationship rather than by how the page looks (README §6.2)
 *
 * Both exist because the deck read as "thin and monotonous": the sizes were ad-hoc numbers,
 * and a page of figures came out as a bullet list because nothing routed it to a chart.
 *
 * Run: node test/chart-layout-test.mjs
 */
import { TYPE_RATIO, composeDeck, typeScale } from '../src/generate.js'
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
const chartOf = (slide) => slide.elements.find((element) => element.type === 'chart')

/** Build a one-page deck from bullets (plus optional page metadata lines). */
function page(bullets, extra = []) {
  const lines = ['# 检验', '', '## 一页', '章节：甲', ...extra]
  for (const bullet of bullets) lines.push(`- ${bullet}`)
  return composeDeck({ outline: lines.join('\n'), theme: THEME, style: { maxSlides: 6 } })
}

/* ------------------------------------------------------------------ *
 * Type scale
 * ------------------------------------------------------------------ */

check('1a · every role sits on the ratio ladder, snapped to even px', () => {
  const scale = typeScale(22)
  const bands = {
    hero: [2.5, 5],
    chapter: [2, 2.5],
    title: [1.5, 2],
    subtitle: [1.2, 1.5],
    lead: [1.1, 1.4],
    bilingual: [0.75, 0.9],
    annotation: [0.7, 0.85],
    footnote: [0.5, 0.65],
  }
  for (const [role, [low, high]] of Object.entries(bands)) {
    const ratio = scale[role] / scale.body
    assert(ratio >= low - 0.06 && ratio <= high + 0.06, `${role} ratio ${ratio.toFixed(2)} outside ${low}–${high}`)
    assert(scale[role] % 2 === 0, `${role} = ${scale[role]}px is not an even number`)
  }
  assert(scale.body === 22, `body anchor moved: ${scale.body}`)
  // The ladder must be strictly ordered, or "role" means nothing.
  const order = ['footnote', 'annotation', 'bilingual', 'body', 'lead', 'subtitle', 'title', 'chapter', 'hero']
  const sizes = order.map((role) => scale[role])
  for (let index = 1; index < sizes.length; index += 1) {
    assert(sizes[index] > sizes[index - 1], `ladder broken at ${order[index]}: ${stable(sizes)}`)
  }
  return order.map((role) => `${role}=${scale[role]}`).join(' ')
})

check('1b · the ladder rescales with the anchor instead of drifting', () => {
  const small = typeScale(18)
  const large = typeScale(26)
  assert(small.body === 18 && large.body === 26, `anchors not honoured: ${small.body}/${large.body}`)
  // A smaller body must not produce a larger title (the bug absolute constants cause).
  assert(large.hero > small.hero, `hero did not scale: ${small.hero} → ${large.hero}`)
  assert(small.hero / small.body > 2.4 && large.hero / large.body < 2.9, 'ratio drifted across anchors')
  return `18→hero ${small.hero}, 26→hero ${large.hero}`
})

check('1c · TYPE_RATIO is the single source of truth', () => {
  const scale = typeScale(22)
  for (const [role, ratio] of Object.entries(TYPE_RATIO)) {
    if (role === 'body') continue
    const expected = Math.round((22 * ratio) / 2) * 2
    assert(scale[role] === expected, `${role}: ${scale[role]} ≠ ratio ${ratio} → ${expected}`)
  }
  return `${Object.keys(TYPE_RATIO).length} roles derived from one table`
})

/* ------------------------------------------------------------------ *
 * Chart routing
 * ------------------------------------------------------------------ */

check('2a · category values route to a bar chart, values unchanged', () => {
  const { deck } = page(['墨西哥 — 62.52 万辆', '俄罗斯 — 58.27 万辆', '阿联酋 — 37.2 万辆', '巴西 — 30.1 万辆'])
  const slide = deck.slides.find((entry) => entry.kind === 'chart')
  assert(slide !== undefined, `expected a chart page, got ${stable(kinds(deck))}`)
  const chart = chartOf(slide)
  assert(chart.chartType === 'bar', `chartType was ${chart.chartType}`)
  assert(stable(chart.data.labels) === stable(['墨西哥', '俄罗斯', '阿联酋', '巴西']), `labels ${stable(chart.data.labels)}`)
  assert(stable(chart.data.series[0].values) === stable([62.52, 58.27, 37.2, 30.1]), `values ${stable(chart.data.series[0].values)}`)
  return `bar, ${chart.data.labels.length} rows`
})

check('2b · a time axis routes to a line chart', () => {
  const { deck } = page(['2021 — 12.4 万辆', '2022 — 67.9 万辆', '2023 — 120.3 万辆', '2024 — 128.4 万辆', '2025 — 201.5 万辆'])
  const slide = deck.slides.find((entry) => entry.kind === 'chart')
  assert(slide !== undefined, `expected a chart page, got ${stable(kinds(deck))}`)
  assert(chartOf(slide).chartType === 'line', `chartType was ${chartOf(slide).chartType}`)
  return `line, 5 points`
})

check('2c · parts that add up to 100 route to a pie chart', () => {
  const { deck } = page(['纯电 — 62 %', '插混 — 24 %', '燃油 — 14 %'])
  const slide = deck.slides.find((entry) => entry.kind === 'chart')
  assert(slide !== undefined, `expected a chart page, got ${stable(kinds(deck))}`)
  assert(chartOf(slide).chartType === 'pie', `chartType was ${chartOf(slide).chartType}`)
  // 62+24+14 must still be 100 after the round trip — the data is not renormalised.
  const values = chartOf(slide).data.series[0].values
  assert(values.reduce((a, b) => a + b, 0) === 100, `values changed: ${stable(values)}`)
  return 'pie, sums to 100'
})

check('2d · mixed units are refused rather than charted from partial data', () => {
  const { deck } = page(['甲 — 12 %', '乙 — 3 万辆', '丙 — 5 个', '丁 — 7 天'])
  assert(!kinds(deck).includes('chart'), `invented a chart from mixed units: ${stable(kinds(deck))}`)
  return 'refused, page stayed a list'
})

check('2e · an unparseable row is refused (no chart from half the data)', () => {
  const { deck } = page(['墨西哥 — 62.52 万辆', '俄罗斯 — 58.27 万辆', '其余国家合计保持增长'])
  assert(!kinds(deck).includes('chart'), `charted an incomplete set: ${stable(kinds(deck))}`)
  return 'refused'
})

check('2f · 结论 becomes the reading beside the chart; 来源 becomes the source line', () => {
  const { deck } = page(
    ['墨西哥 — 62.52 万辆', '俄罗斯 — 58.27 万辆', '阿联酋 — 37.2 万辆'],
    ['来源：中国汽车工业协会 / caam.org.cn', '结论：墨西哥首次超过俄罗斯'],
  )
  const slide = deck.slides.find((entry) => entry.kind === 'chart')
  assert(slide !== undefined, 'no chart page')
  const texts = slide.elements.filter((element) => element.type === 'text').map((element) => element.content)
  assert(texts.includes('墨西哥首次超过俄罗斯'), `conclusion missing: ${stable(texts)}`)
  assert(texts.includes('中国汽车工业协会 / caam.org.cn'), `source missing: ${stable(texts)}`)
  // The source must be set at the footnote role, not at body size.
  const source = slide.elements.find((element) => element.id.endsWith('-source-1'))
  assert(source.story === undefined && source.style.fontSize < 18, `source size ${source.style.fontSize} is not a footnote`)
  // The chart yields width to the reading when there is one.
  const chart = chartOf(slide)
  assert(chart.w < 1152 * 0.8, `chart kept full width (${chart.w}) despite a conclusion`)
  return `chart ${chart.w}px wide + reading + ${source.style.fontSize}px source`
})

check('2g · without 来源 no citation is invented', () => {
  const { deck } = page(['墨西哥 — 62.52 万辆', '俄罗斯 — 58.27 万辆', '阿联酋 — 37.2 万辆'])
  const slide = deck.slides.find((entry) => entry.kind === 'chart')
  const source = slide.elements.find((element) => element.id.endsWith('-source-1'))
  assert(source === undefined, 'drew a source line that was never authored')
  return 'no invented citation'
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
