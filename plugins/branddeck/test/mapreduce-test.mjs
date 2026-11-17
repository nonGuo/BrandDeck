/**
 * Tests for src/mapreduce.js — the pure half of the long-material pipeline.
 *
 * The model calls live in assist.js and are not exercised here; what is tested is the
 * part that can silently corrupt a deck: splitting (must not lose or reorder text),
 * fact merging (must not drop a number), and bounded concurrency (must not open one
 * call per block).
 *
 * Run:
 *   C:\Users\Administrator\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe test/mapreduce-test.mjs
 */
import {
  MATERIAL_DEFAULTS,
  buildMapPrompt,
  buildReducePrompt,
  mergeFacts,
  needsMapReduce,
  parseFacts,
  progressLabel,
  resolveMaterialOptions,
  runWithConcurrency,
  sliceAtSentence,
  splitIntoBlocks,
} from '../src/mapreduce.js'

/* ------------------------------------------------------------------ *
 * Harness
 * ------------------------------------------------------------------ */

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

async function checkAsync(name, fn) {
  try {
    const detail = await fn()
    results.push({ name, ok: true })
    console.log(`PASS ${name}${detail === undefined ? '' : ` — ${detail}`}`)
  } catch (error) {
    results.push({ name, ok: false })
    console.log(`FAIL ${name}: ${error?.message ?? error}`)
  }
}

const stable = (value) => JSON.stringify(value)

/**
 * Compare text ignoring what a block layout is allowed to change: whitespace runs, and
 * the `#` markers of headings (which travel inline with the block text).
 */
function flatten(text) {
  return String(text ?? '')
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/\s+/g, '')
}

/* ------------------------------------------------------------------ *
 * Fixtures
 * ------------------------------------------------------------------ */

const SECTIONS = [
  ['一、为什么做这件事', '企业做一份合规的演示文稿平均要花五天。其中大部分时间不是花在内容上，而是花在找模板、对齐品牌色、调字体、修版式溢出上。'],
  ['二、整体架构', '整条链路分成四层：模板特征提取、内容生成、可视化微调、原生编译。层与层之间用数据契约连接，每层只做一件事。'],
  ['三、模板解析的坑', '主题色在 clrScheme 里是反的，直接读 accent1 会把白色当作品牌主色。母版 Logo 的坐标在 spTree 里，不在占位符上。'],
  ['四、卡片 DSL', '画布固定 1280×720，安全边距 64 像素，右上角预留 Logo 区。模型只产出内容，坐标由确定性代码计算。'],
  ['五、版式路由', '有对比关系用对比版式，有阶段用时间线，以数字为主用指标卡。作者没有声明时由确定性规则推断，推断不出来就退回列表。'],
  ['六、SVG 示意图', '模型生成的 SVG 先用模板令牌换肤，再按三倍分辨率光栅化嵌入。不直接嵌 SVG，因为不同版本和平台的渲染差异太大。'],
  ['七、长材料', '超过阈值的长材料走 Map-Reduce：按章节切块、逐块提炼事实、合并去重、再一次编排成大纲。分块必须零丢失。'],
  ['八、两种中间产物', '卡片版给面板微调用，文字版给 review 和 git diff 用。两份由同一个 deck 派生，改文字不动位置。'],
  ['九、验证方式', '不靠"应该能用"：模板解析对真实模板取证，生成器有二十项断言，导出后用 LibreOffice 渲染回图片逐页看。'],
  ['十、已知边界', '复杂动画、母版级版式、SmartArt 不在范围内。图表只有柱状、折线、饼图三种，够用但不万能。'],
]

/** A document with 10 sections, ~1,100 characters — realistic Markdown material. */
function buildDocument() {
  return SECTIONS.map(([heading, body]) => `## ${heading}\n\n${body}`).join('\n\n')
}

/* ------------------------------------------------------------------ *
 * Splitting
 * ------------------------------------------------------------------ */

check('1a · splitIntoBlocks loses no character at any block size', () => {
  const text = buildDocument()
  const report = []
  for (const blockChars of [500, 1000, 2000, 4000, 20000]) {
    const blocks = splitIntoBlocks(text, { blockChars, maxBlocks: 60 })
    const joined = flatten(blocks.map((block) => block.text).join('\n\n'))
    const original = flatten(text)
    assert(joined === original, `blockChars=${blockChars}: lost/duplicated text (${original.length} vs ${joined.length})`)
    const oversize = blocks.filter((block) => block.chars > blockChars)
    assert(oversize.length === 0, `blockChars=${blockChars}: ${oversize.length} block(s) over the limit (${oversize.map((block) => block.chars).join(',')})`)
    report.push(`${blockChars}:${blocks.length}`)
  }
  return `blocks per blockChars ${report.join(' ')}`
})

check('1b · block count follows the size target, not the section count', () => {
  const text = buildDocument()
  const blocks = splitIntoBlocks(text, { blockChars: 4000, maxBlocks: 12 })
  // 10 short sections must NOT become 10 calls: they pack into a couple of blocks.
  assert(blocks.length <= 2, `expected packing into <=2 blocks, got ${blocks.length}`)
  const chars = blocks.map((block) => block.chars)
  assert(chars.every((count) => count > 0), `empty block in ${stable(chars)}`)
  return `${blocks.length} blocks (${chars.join(', ')} chars)`
})

check('1c · headings travel inline so the map step sees structure', () => {
  const blocks = splitIntoBlocks(buildDocument(), { blockChars: 1000, maxBlocks: 60 })
  const shouldHaveHeadingMarkers = blocks.slice(0, -1)
  for (const block of shouldHaveHeadingMarkers) {
    assert(/^##\s/m.test(block.text), `block ${block.index} has no inline heading: ${block.text.slice(0, 40)}`)
  }
  // The recorded heading is the section the block starts in, for progress/errors.
  assert(blocks[0].heading === SECTIONS[0][0], `first heading was "${blocks[0].heading}"`)
  return `${blocks.length} blocks, headings inline`
})

check('1d · an oversized section is cut at a sentence boundary', () => {
  const body = '这是一句用来测试分块边界的话，长度大约二十几个字。'.repeat(60)
  const blocks = splitIntoBlocks(`## 长章节\n\n${body}`, { blockChars: 500, maxBlocks: 60 })
  assert(blocks.length > 1, `expected several blocks, got ${blocks.length}`)
  for (const block of blocks.slice(0, -1)) {
    assert(block.text.endsWith('。'), `block ${block.index} does not end at a sentence ender: ...${block.text.slice(-12)}`)
  }
  const joined = flatten(blocks.map((block) => block.text).join('\n\n'))
  assert(joined === flatten(`## 长章节\n\n${body}`), 'slicing lost text')
  return `${blocks.length} sentence-aligned blocks`
})

check('1e · maxBlocks merges the tail instead of dropping it', () => {
  // Four copies of the document ≈ 2,800 chars → with 500-char blocks that is well past
  // the 3-block budget, so the merge path is actually exercised.
  const text = [buildDocument(), buildDocument(), buildDocument(), buildDocument()].join('\n\n')
  const blocks = splitIntoBlocks(text, { blockChars: 500, maxBlocks: 3 })
  assert(blocks.length === 3, `expected 3 blocks, got ${blocks.length}`)
  const joined = flatten(blocks.map((block) => block.text).join('\n\n'))
  assert(joined === flatten(text), `tail merge lost text (${flatten(text).length} vs ${joined.length})`)
  const indexes = blocks.map((block) => block.index)
  assert(stable(indexes) === stable([0, 1, 2]), `indexes were ${stable(indexes)}`)
  return `3 blocks, ${blocks.map((block) => block.chars).join('/')} chars, lossless`
})

check('1f · empty and whitespace-only input produce no blocks', () => {
  for (const input of ['', '   ', '\n\n\n']) {
    const blocks = splitIntoBlocks(input, {})
    assert(blocks.length === 0, `"${input}" produced ${blocks.length} blocks`)
  }
  return 'empty inputs → 0 blocks'
})

check('1g · sliceAtSentence respects the limit and prefers an ender', () => {
  const text = '第一句话在这里。第二句话在这里。第三句话在这里。'
  const short = '只有一个短句。'
  assert(sliceAtSentence(short, 100) === short, 'short text must come back unchanged')
  const slice = sliceAtSentence(text, 12)
  assert(Array.from(slice).length <= 12, `slice of ${Array.from(slice).length} chars exceeds 12`)
  assert(slice.endsWith('。'), `slice did not end at an ender: "${slice}"`)
  // Pathological: no ender at all inside the window → a hard cut, never an empty string.
  const hard = sliceAtSentence('字'.repeat(50), 10)
  assert(Array.from(hard).length === 10, `hard cut returned ${Array.from(hard).length} chars`)
  return `"${slice}" (${Array.from(slice).length} chars)`
})

/* ------------------------------------------------------------------ *
 * Options and threshold
 * ------------------------------------------------------------------ */

check('2a · needsMapReduce switches exactly at the threshold', () => {
  const threshold = 100
  assert(needsMapReduce('字'.repeat(100), { threshold }) === false, 'equal length must stay single-pass')
  assert(needsMapReduce('字'.repeat(101), { threshold }) === true, 'one over the threshold must map-reduce')
  assert(needsMapReduce('字'.repeat(7000), {}) === true, 'default threshold should trip at 7000 chars')
  assert(needsMapReduce('', {}) === false, 'empty material must not map-reduce')
  return 'threshold is exclusive, defaults apply'
})

check('2b · resolveMaterialOptions clamps hostile config', () => {
  const defaults = resolveMaterialOptions(undefined)
  assert(defaults.threshold === MATERIAL_DEFAULTS.threshold, 'threshold default lost')
  const hostile = resolveMaterialOptions({ materialThreshold: -5, materialBlockChars: 1e9, materialMaxBlocks: 0, materialConcurrency: 99, materialFactsPerBlock: 'abc' })
  assert(hostile.threshold === 500, `threshold clamped to ${hostile.threshold}`)
  assert(hostile.blockChars === 40000, `blockChars clamped to ${hostile.blockChars}`)
  assert(hostile.maxBlocks === 1, `maxBlocks clamped to ${hostile.maxBlocks}`)
  assert(hostile.concurrency === 8, `concurrency clamped to ${hostile.concurrency}`)
  assert(hostile.factsPerBlock === MATERIAL_DEFAULTS.factsPerBlock, 'non-numeric factsPerBlock must fall back')
  return 'clamps to [500,200000] / [500,40000] / [1,60] / [1,8]'
})

/* ------------------------------------------------------------------ *
 * Facts
 * ------------------------------------------------------------------ */

check('3a · parseFacts reads bullets and numbered lists, ignores prose', () => {
  const answer = [
    '下面是提炼结果：',
    '- 接口耗时从 450ms 降到 18ms',
    '* 三级缓存命中率 92%',
    '1. 2026 Q3 上线，覆盖 12 个事业部',
    '- **加粗的要点也要能读出来**',
    '- （本段无可用事实）',
    '',
    '以上。',
  ].join('\n')
  const facts = parseFacts(answer)
  assert(facts.length === 4, `parsed ${facts.length} facts: ${stable(facts)}`)
  assert(facts[0] === '接口耗时从 450ms 降到 18ms', `first fact "${facts[0]}"`)
  assert(facts[3] === '加粗的要点也要能读出来', `bold fact "${facts[3]}"`)
  assert(!facts.some((fact) => fact.includes('本段无')), 'placeholder fact leaked')
  return `${facts.length} facts, placeholder dropped`
})

check('3b · mergeFacts de-duplicates exact repeats and keeps order', () => {
  const groups = [
    ['接口耗时从 450ms 降到 18ms', '三级缓存命中率 92%'],
    ['接口耗时从450ms降到18ms。', '覆盖 12 个事业部'],
  ]
  const { facts, dropped } = mergeFacts(groups)
  assert(facts.length === 3, `kept ${facts.length}: ${stable(facts)}`)
  assert(dropped === 1, `dropped ${dropped}`)
  assert(facts[0] === '接口耗时从 450ms 降到 18ms', 'order changed')
  assert(facts[1] === '三级缓存命中率 92%', 'order changed')
  return `${facts.length} kept, ${dropped} dropped`
})

check('3c · mergeFacts keeps a number that only the longer fact carries', () => {
  // Conservative on purpose: "命中率 92%" and "缓存命中率" are different facts, and
  // dropping the short one would lose the number that makes the page credible.
  const groups = [['缓存命中率 92%', '缓存命中率'], ['接口耗时 18ms', '接口耗时']]
  const { facts } = mergeFacts(groups)
  assert(facts.includes('缓存命中率 92%'), 'lost the fact carrying the number')
  assert(facts.includes('接口耗时 18ms'), 'lost the fact carrying the number')
  return `${facts.length} facts kept (${stable(facts)})`
})

check('3d · mergeFacts caps at max and reports an accurate count', () => {
  const groups = [Array.from({ length: 30 }, (_, i) => `第 ${i + 1} 条要点带有编号内容`)]
  const { facts, dropped } = mergeFacts(groups, { max: 5 })
  assert(facts.length === 5, `kept ${facts.length}`)
  // 30 facts in, 5 kept → exactly 25 lost, however they were lost.
  assert(dropped === 25, `dropped ${dropped}, expected 25`)
  return `${facts.length} kept, ${dropped} dropped`
})

check('3e · mergeFacts cannot loop forever on duplicates only', () => {
  const groups = [['同一句话重复出现三次', '同一句话重复出现三次'], ['同一句话重复出现三次']]
  const { facts, dropped } = mergeFacts(groups)
  assert(facts.length === 1, `kept ${facts.length}`)
  assert(dropped === 2, `dropped ${dropped}`)
  return `1 kept, ${dropped} dropped`
})

check('3f · a duplicate hit on the same pass as the cap is not counted twice', () => {
  // Regression: `dropped` used to add the cap overflow on top of the duplicates already
  // counted, so it reported 6 losses for 10 facts where 5 were kept.
  const facts = [
    '第一条要点内容甲', '第一条要点内容甲', '第三条要点内容丙', '第四条要点内容丁',
    '第五条要点内容戊', '第六条要点内容己', '第七条要点内容庚', '第八条要点内容辛',
  ]
  const { facts: kept, dropped } = mergeFacts([facts], { max: 5 })
  assert(kept.length === 5, `kept ${kept.length}`)
  assert(dropped === 3, `dropped ${dropped}, expected 3 (8 in, 5 kept)`)
  return `${kept.length} kept, ${dropped} dropped`
})

/* ------------------------------------------------------------------ *
 * Prompts
 * ------------------------------------------------------------------ */

check('4a · the map prompt carries the block, its position and the fact budget', () => {
  const [block] = splitIntoBlocks(buildDocument(), { blockChars: 4000, maxBlocks: 12 })
  const { system, prompt } = buildMapPrompt({ block, total: 3, factsPerBlock: 6 })
  assert(prompt.includes('1/3'), `prompt lacks position: ${prompt.slice(0, 60)}`)
  assert(prompt.includes(block.text.slice(0, 20)), 'prompt lacks the block text')
  // The floor is `max(3, ceil(factsPerBlock / 2))`, so 6 becomes "3~6 行".
  assert(system.includes('3~6 行'), `fact budget not stated: ${system}`)
  assert(system.includes('数字'), 'the number-preservation rule is missing')
  return `${system.length + prompt.length} chars`
})

check('4b · the reduce prompt restates the outline convention and the page band', () => {
  const { system, prompt } = buildReducePrompt({ facts: ['要点一 138%', '要点二 74%'], maxSlides: 8 })
  assert(system.includes('# <卡组标题>'), 'title convention missing')
  assert(system.includes('## <页面标题>'), 'page convention missing')
  // The cap is a *band*: "no more than N" made the model answer with the minimum, which is
  // exactly how a 29k-character report became four pages.
  assert(system.includes('8 页'), `page band missing: ${system.slice(-200)}`)
  assert(system.includes('上限不是目标'), 'the cap is still phrased as a ceiling')
  assert(system.includes('6~8 页'), `target band missing: ${system.slice(-200)}`)
  assert(prompt.includes('1. 要点一 138%'), `facts not numbered: ${prompt}`)
  return `${system.length + prompt.length} chars`
})

check('4c · progressLabel is human readable', () => {
  assert(progressLabel(3, 7) === '已提炼 3/7 段', progressLabel(3, 7))
  return progressLabel(3, 7)
})

/* ------------------------------------------------------------------ *
 * Concurrency
 * ------------------------------------------------------------------ */

await checkAsync('5a · runWithConcurrency bounds the in-flight count and keeps order', async () => {
  const items = Array.from({ length: 12 }, (_, i) => i)
  let inFlight = 0
  let peak = 0
  const results = await runWithConcurrency(items, 3, async (item) => {
    inFlight += 1
    peak = Math.max(peak, inFlight)
    await new Promise((resolve) => setTimeout(resolve, 5 + (item % 3)))
    inFlight -= 1
    return item * 2
  })
  assert(peak <= 3, `peak concurrency was ${peak}`)
  assert(peak === 3, `concurrency never reached the limit (peak ${peak})`)
  assert(stable(results) === stable(items.map((item) => item * 2)), 'results out of order')
  return `peak ${peak}/3 over ${items.length} items`
})

await checkAsync('5b · runWithConcurrency handles an empty list and a limit above the list', async () => {
  const empty = await runWithConcurrency([], 3, async () => 1)
  assert(Array.isArray(empty) && empty.length === 0, `empty list returned ${stable(empty)}`)
  const two = await runWithConcurrency([1, 2], 8, async (item) => item + 1)
  assert(stable(two) === stable([2, 3]), `short list returned ${stable(two)}`)
  return 'empty and oversized limit both safe'
})

await checkAsync('5c · a failing task rejects instead of resolving a hole', async () => {
  let rejected = false
  try {
    await runWithConcurrency([1, 2, 3], 2, async (item) => {
      if (item === 2) throw new Error('boom')
      return item
    })
  } catch (error) {
    rejected = String(error?.message).includes('boom')
  }
  assert(rejected, 'a throwing task must propagate its error')
  return 'error propagates'
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
