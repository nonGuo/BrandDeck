/**
 * Tests for the densify pass — "outline is too thin, fill in the detail".
 *
 * The measurement half is pure (`measureOutline` / `thinPages`) and is pinned here.
 * The model call itself lives in `densifyOutline` and is exercised live; what is testable
 * deterministically is the *contract*: metadata is not copy, a diagram page is not "thin",
 * and the pass must never report a page as dense because of its `备注：`.
 *
 * Run: node test/densify-test.mjs
 */
import { DENSIFY, buildEnrichPrompt, measureOutline, thinPages } from '../src/assist.js'

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

/** A thin deck: the shape the real generator produced before this pass existed. */
const THIN = [
  '# 卡组标题',
  '',
  '## 定位',
  '章节：结论',
  '- 五天',
  '- 半小时',
  '备注：企业合规演示平均耗时五天，目标压缩到半小时，这一段很长的备注不应该算作正文。',
  '',
  '## 架构',
  '章节：实现',
  '版式：架构图',
  '图：mermaid',
  'flowchart LR',
  '  A[解析] --> B[生成]',
  '',
  '## 证据',
  '章节：实现',
  '- 企业合规演示平均耗时五天，含排版、配色与三轮评审返工',
  '- 单页正文控制在八十到一百四十字，读者只看这一页也要能读懂',
].join('\n')

/* ------------------------------------------------------------------ *
 * Measurement
 * ------------------------------------------------------------------ */

check('1a · visible copy is title + bullets, and metadata is excluded', () => {
  const { pages } = measureOutline(THIN)
  assert(pages.length === 3, `pages ${pages.length}`)
  const [positioning] = pages
  // "定位"(2) + "五天"(2) + "半小时"(3) = 7 — the 备注 must not be counted.
  assert(positioning.chars === 7, `chars ${positioning.chars} (${stable(positioning)})`)
  assert(positioning.bullets === 2, `bullets ${positioning.bullets}`)
  return `${pages.map((page) => page.chars).join('/')} chars`
})

check('1b · a diagram page counts its bullets but not the Mermaid source', () => {
  const { pages } = measureOutline(THIN)
  const diagram = pages[1]
  // Title "架构"(2), no bullets; the flowchart lines are source, not copy.
  assert(diagram.chars === 2, `diagram page chars ${diagram.chars}`)
  assert(diagram.bullets === 0, `diagram page bullets ${diagram.bullets}`)
  return `diagram page = ${diagram.chars} chars`
})

check('1c · a fenced Mermaid block is skipped entirely', () => {
  const fenced = ['## 图页', '版式：架构图', '- 结论一句', '图：mermaid', '```mermaid', 'flowchart LR', '  A[甲] --> B[乙]', '```'].join('\n')
  const { pages } = measureOutline(fenced)
  assert(pages.length === 1, `pages ${pages.length}`)
  // "图页"(2) + "结论一句"(4) = 6
  assert(pages[0].chars === 6, `chars ${pages[0].chars}`)
  return `fenced block ignored (${pages[0].chars} chars)`
})

check('1d · an empty or metadata-only outline measures to zero', () => {
  for (const input of ['', '# 只有标题', '## 只有页标题\n版式：列表\n章节：甲']) {
    const measured = measureOutline(input)
    assert(measured.pages.every((page) => page.bullets === 0 || page.chars >= 0), 'bad measurement')
  }
  const onlyMeta = measureOutline('## 页\n版式：列表\n- \n')
  assert(onlyMeta.total >= 2, `total ${onlyMeta.total}`)
  return 'odd inputs measured without throwing'
})

/* ------------------------------------------------------------------ *
 * Selection
 * ------------------------------------------------------------------ */

check('2a · thin pages are the short ones with bullets, in page order', () => {
  // Both pages with copy are under the 70-char floor here (7 and 56 chars); the diagram
  // page is skipped (see 2b). Order must follow the outline.
  const thin = thinPages(THIN, { floor: 70 })
  assert(thin.length === 2, `thin ${thin.length}: ${stable(thin.map((page) => page.title))}`)
  assert(stable(thin.map((page) => page.title)) === stable(['定位', '证据']), `thin pages ${stable(thin.map((page) => page.title))}`)
  assert(thin[0].index === 0 && thin[1].index === 2, `indexes ${stable(thin.map((page) => page.index))}`)
  return `2 thin pages of 3 (${thin.map((page) => page.chars).join(', ')} chars < 70)`
})

check('2b · a diagram page is never "thin" — the figure is the content', () => {
  const thin = thinPages(THIN, { floor: 200 })
  assert(!thin.some((page) => page.title === '架构'), 'the diagram page was selected for densifying')
  return `selected ${stable(thin.map((page) => page.title))} with a high floor`
})

check('2c · the floor defaults to DENSIFY.floor and 0 selects nothing', () => {
  assert(Number.isFinite(DENSIFY.floor) && DENSIFY.floor > 0, 'DENSIFY.floor is not a number')
  const none = thinPages(THIN, { floor: 0 })
  assert(none.length === 0, `floor 0 selected ${none.length} pages`)
  const all = thinPages(THIN, { floor: 1000 })
  assert(all.every((page) => page.bullets > 0), 'a page without bullets slipped in')
  return `default ${DENSIFY.floor}, floor 0 → 0 pages`
})

/* ------------------------------------------------------------------ *
 * Prompt
 * ------------------------------------------------------------------ */

check('3a · the densify prompt freezes structure and names the target length', () => {
  const pages = thinPages(THIN, { floor: 70 })
  const { system, prompt } = buildEnrichPrompt({ markdown: THIN, pages, charBudget: 150 })
  // Structure freeze: the model must not rename pages or drop chapters.
  assert(system.includes('完全不变'), 'structure freeze is not stated')
  assert(system.includes('页数、页序、每页标题都不许变'), 'page count/order rule missing')
  // Target length: 55%..100% of the budget.
  const low = Math.round(150 * 0.55)
  const high = Math.round(150 * 0.85)
  assert(system.includes(`${low}~${high} 字`), `target length missing: ${system.slice(0, 160)}`)
  assert(system.includes('不许编造'), 'the no-invention rule is missing')
  // The prompt must carry the full outline so the model can keep the parts it keeps.
  assert(prompt.includes('## 定位') && prompt.includes('## 证据'), 'outline not included')
  assert(prompt.includes('「定位」'), 'the thin page is not named')
  return `${system.length + prompt.length} chars, target 82~150`
})

check('3b · the budget is clamped into a sane range', () => {
  const pages = thinPages(THIN, { floor: 70 })
  // 10 chars is below the floor: it clamps to 40, and the target band is 55%~85% of it.
  const low = buildEnrichPrompt({ markdown: THIN, pages, charBudget: 10 })
  const lowBand = `${Math.round(40 * 0.55)}~${Math.round(40 * 0.85)} 字`
  assert(low.system.includes(lowBand), `low budget not clamped to 40 (expected "${lowBand}"): ${low.system.slice(0, 160)}`)
  // 9999 is above the ceiling: it clamps to DENSIFY.ceiling.
  const high = buildEnrichPrompt({ markdown: THIN, pages, charBudget: 9999 })
  const highBand = `${Math.round(DENSIFY.ceiling * 0.85)} 字`
  assert(high.system.includes(highBand), `high budget not clamped to ${DENSIFY.ceiling}: ${high.system.slice(0, 160)}`)
  return `clamped to [40, ${DENSIFY.ceiling}]`
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
