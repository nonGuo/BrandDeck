/**
 * Long-material pipeline: Map-Reduce hierarchical extraction.
 *
 * WHY: a single prompt only sees a bounded window, and even inside it a 20,000-character
 * report degrades badly — the model latches onto the first few sections and pads the
 * rest. The reference approach is Map-Reduce: extract facts per block (map), then
 * cluster and de-duplicate them into one outline (reduce). This module owns the *pure*
 * half of that (splitting, prompt building, fact parsing/merging, concurrency); the
 * model calls live in `assist.js`.
 *
 * Two rules that matter for quality, both learned from the reference:
 *   · Never split a block mid-sentence, and never drop text: every character of the
 *     source must land in exactly one block, so a fact cannot vanish between blocks.
 *   · The map phase extracts *facts with their numbers, names and dates intact*; the
 *     reduce phase is the only place allowed to compress wording.
 */

/** A block of source text plus where it came from, for progress messages and errors. */
/**
 * @typedef {object} MaterialBlock
 * @property {number} index 0-based block number
 * @property {string} text   the block's text, trimmed
 * @property {number} chars  code-point count of `text`
 * @property {string} heading the nearest preceding Markdown heading, when there was one
 */

/** Defaults, overridable through the plugin config. */
export const MATERIAL_DEFAULTS = Object.freeze({
  /** Above this many characters, one prompt is no longer trustworthy → map-reduce. */
  threshold: 6000,
  /** Target size of one map block. */
  blockChars: 4000,
  /** Hard cap on map calls, so a huge document cannot become an unbounded spend. */
  maxBlocks: 12,
  /** How many map calls may be in flight at once. */
  concurrency: 3,
  /** Facts requested per block. Keep modest: facts are input to a fixed output budget. */
  factsPerBlock: 6,
})

/** @param {unknown} value */
const isObject = (value) => value !== null && typeof value === 'object' && !Array.isArray(value)

/** @param {string} value */
const charCount = (value) => Array.from(String(value ?? '')).length

/**
 * Resolve the tuning knobs from a config object, clamped to sane ranges.
 * @param {any} [config]
 */
export function resolveMaterialOptions(config) {
  const raw = isObject(config) ? config : {}
  const pick = (value, fallback, min, max) => {
    const number = Number(value)
    if (!Number.isFinite(number)) return fallback
    return Math.min(max, Math.max(min, Math.floor(number)))
  }
  return {
    threshold: pick(raw.materialThreshold, MATERIAL_DEFAULTS.threshold, 500, 200000),
    blockChars: pick(raw.materialBlockChars, MATERIAL_DEFAULTS.blockChars, 500, 40000),
    maxBlocks: pick(raw.materialMaxBlocks, MATERIAL_DEFAULTS.maxBlocks, 1, 60),
    concurrency: pick(raw.materialConcurrency, MATERIAL_DEFAULTS.concurrency, 1, 8),
    factsPerBlock: pick(raw.materialFactsPerBlock, MATERIAL_DEFAULTS.factsPerBlock, 3, 20),
  }
}

/**
 * Decide whether material needs the map-reduce path.
 * @param {string} text
 * @param {{threshold?: number}} [options]
 */
export function needsMapReduce(text, options = {}) {
  const limit = Number.isFinite(options.threshold) ? Number(options.threshold) : MATERIAL_DEFAULTS.threshold
  return charCount(text) > limit
}

/**
 * Split material into blocks without losing a character.
 *
 * Strategy, in order of preference:
 *   1. Markdown headings (`#`..`######`) are hard boundaries — a block never spans two
 *      sections, because the reduce phase wants coherent units.
 *   2. Within a section, paragraphs (`\n\n`) are the unit; a run of short paragraphs is
 *      packed until it reaches `blockChars`.
 *   3. A single paragraph longer than `blockChars` is split at sentence enders
 *      (`。！？；.!?;` and newlines), and only as a last resort mid-sentence.
 *
 * @param {string} text
 * @param {{blockChars?: number, maxBlocks?: number}} [options]
 * @returns {MaterialBlock[]}
 */
export function splitIntoBlocks(text, options = {}) {
  const source = String(text ?? '').replace(/\r\n?/g, '\n')
  const limit = Number.isFinite(options.blockChars) ? Math.max(200, options.blockChars) : MATERIAL_DEFAULTS.blockChars
  const maxBlocks = Number.isFinite(options.maxBlocks) ? Math.max(1, options.maxBlocks) : MATERIAL_DEFAULTS.maxBlocks
  if (source.trim() === '') return []

  // 1. Group the source into sections, keeping the heading as an inline marker: a block
  //    may span several short sections, and the model needs to see where they split.
  /** @type {{heading: string, line: string, text: string}[]} */
  const sections = []
  let heading = ''
  let headingLine = ''
  let buffer = []
  const flushSection = () => {
    const body = buffer.join('\n').trim()
    if (body !== '') sections.push({ heading, line: headingLine, text: body })
    buffer = []
  }
  for (const line of source.split('\n')) {
    if (/^#{1,6}\s+/.test(line)) {
      flushSection()
      headingLine = line.trim()
      heading = line.replace(/^#{1,6}\s+/, '').trim()
      continue
    }
    buffer.push(line)
  }
  flushSection()
  if (sections.length === 0) return []

  // 2. Pack sections into blocks no larger than `limit`, so a document of many short
  //    sections becomes a few full blocks rather than a dozen tiny calls.
  /** @type {MaterialBlock[]} */
  const blocks = []
  /** @param {string} body @param {string} blockHeading */
  const pushBlock = (body, blockHeading) => {
    const trimmed = body.trim()
    if (trimmed === '') return
    blocks.push({ index: blocks.length, text: trimmed, chars: charCount(trimmed), heading: blockHeading })
  }

  let pending = ''
  let pendingHeading = ''
  for (const section of sections) {
    // The heading travels with the text (as a Markdown line) so structure survives.
    const piece = section.line === '' ? section.text : `${section.line}\n${section.text}`
    const candidate = pending === '' ? piece : `${pending}\n\n${piece}`
    if (charCount(candidate) <= limit) {
      if (pending === '') pendingHeading = section.heading
      pending = candidate
      continue
    }
    if (pending !== '') pushBlock(pending, pendingHeading)
    // The new section alone may still be oversized: slice it at sentence boundaries, and
    // keep the heading on its first slice.
    let rest = piece
    let firstSliceHeading = section.heading
    while (charCount(rest) > limit) {
      const slice = sliceAtSentence(rest, limit)
      pushBlock(slice, firstSliceHeading)
      firstSliceHeading = section.heading
      rest = rest.slice(slice.length).replace(/^\n+/, '')
    }
    pending = rest
    pendingHeading = section.heading
  }
  if (pending.trim() !== '') pushBlock(pending, pendingHeading)

  if (blocks.length === 0) return []

  // 3. Respect the budget: merge the tail when a document is enormous, so we never drop
  //    material (the last block simply carries more text).
  if (blocks.length > maxBlocks) {
    const kept = blocks.slice(0, maxBlocks - 1)
    const tail = blocks.slice(maxBlocks - 1)
    kept.push({
      index: kept.length,
      text: tail.map((block) => block.text).join('\n\n'),
      chars: tail.reduce((sum, block) => sum + block.chars, 0),
      heading: tail[0].heading,
    })
    return kept.map((block, index) => ({ ...block, index }))
  }
  return blocks
}

/**
 * Cut `text` at the last sentence ender at or before `limit`, falling back to a hard cut.
 * @param {string} text
 * @param {number} limit
 * @returns {string}
 */
export function sliceAtSentence(text, limit) {
  const chars = Array.from(text)
  if (chars.length <= limit) return text
  const window = chars.slice(0, limit).join('')
  const cut = Math.max(
    window.lastIndexOf('。'),
    window.lastIndexOf('！'),
    window.lastIndexOf('？'),
    window.lastIndexOf('；'),
    window.lastIndexOf('\n'),
    window.lastIndexOf('. '),
  )
  // Keep the ender itself, and avoid returning almost nothing on pathological input.
  if (cut >= Math.floor(limit * 0.3)) return window.slice(0, cut + 1)
  return window
}

/**
 * Prompt for one map step: a block of source → short, number-preserving facts.
 * @param {{block: MaterialBlock, total: number, factsPerBlock: number, deckTitle?: string}} input
 */
export function buildMapPrompt(input) {
  const { block, total, factsPerBlock } = input
  return {
    system: [
      '你在为一份演示文稿做资料提炼（Map 阶段）。',
      '输入是长材料中的一个片段，你要把它压缩成若干条"可上页的事实"。',
      '要求：',
      `- 输出 ${Math.max(3, Math.ceil(factsPerBlock / 2))}~${factsPerBlock} 行，每行以 "- " 开头，不要编号、不要小标题`,
      '- **必须保留原始的人名、工具名、具体数字、百分比、时间节点**，这些是可信度的来源，不要概括掉',
      '- 每条不超过 30 个汉字，写成陈述句，不要评价、不要空话（禁止"赋能/抓手/闭环/矩阵/重塑/全链路/深耕"这类词）',
      '- 只提炼本片段里真实出现的内容，不要补充常识、不要推测',
      '- 如果本片段没有任何可用信息，输出一行 "- （本段无可用事实）"',
      '只输出这些行，不要任何解释。',
    ].join('\n'),
    prompt: [
      block.heading === '' ? `片段 ${block.index + 1}/${total}` : `片段 ${block.index + 1}/${total}（所属章节：${block.heading}）`,
      '',
      block.text,
    ].join('\n'),
  }
}

/**
 * Prompt for the reduce step: merged facts → a Markdown outline in BrandDeck's format.
 * @param {{facts: string[], maxSlides: number, deckTitle?: string, topic?: string}} input
 */
export function buildReducePrompt(input) {
  const cap = Number.isFinite(input.maxSlides) ? input.maxSlides : 10
  return {
    system: [
      '你在为一份演示文稿做结构编排（Reduce 阶段）。',
      '输入是从长材料各片段提炼出的要点（可能重复、可能零散）。你的任务：去重合并、按逻辑聚类成章节、排成页面。',
      '结构（必须遵守，这是"总-分-总"）：',
      '- 第一行 `# <卡组标题>`（含具体对象，不要"XX方案"这类空标题）',
      '- **每个 `##` 页块里写一行 `章节：<章节名>`**：同名章节会被编成一章，生成器据此插入目录页与章节隔页',
      '  · 章节名沿用材料自己的大标题（"一、总体规模"、"二、区域市场"这种），不要另起一套说法；材料有几章就写几章，一章都不要并掉',
      '  · 每章 1~3 页：内容多的章多给页，不要为了压缩页数把两章挤成一章',
      '  · 第一章是**总**（结论与全景），中间各章是**分**，最后一章是**总**（回收结论与下一步）；材料没有总结章时自己补一章"结论与展望"',
      '页面：',
      '- 每个 `## <页面标题>` 下写 3~6 条 `- <要点>`，**每条 12~28 个汉字**，写成「主张 + 关键数据/限定条件」的短句',
      '- **若材料里有可比较的同类数字**（同一指标的几个类别 / 几个年份 / 几个部分），把这一页写成数据行而不是句子：',
      '  每条 `- <标签> — <数值><单位>`，至少 3 行、单位一致（例：`- 墨西哥 — 62.52 万辆`）。',
      '  生成器会自动选图（类别→柱状、年份→折线、加总为 100%→饼图）；单位不一致或不足 3 行时不会成图。',
      '- **单页「标题 + 要点」合计 80~125 字**（预算 150 字，标题也算在内）：不要把页面写成关键词清单——读者只看这一页也要能读懂',
      '- 页面标题写成**带结论或数字的完整判断**（如"三级缓存把接口耗时从 450ms 降到 18ms"），不要"性能优化方案"',
      '- 必须保留输入里的具体数字、人名、工具名、时间节点；宁可少一条要点，也不要丢掉数字',
      '- 更长的解释、案例、来龙去脉写进 `备注：<一段话>`（页面正文与备注是两条轨，各司其职）',
      '版式（该页块内一行 `版式：<类型>`）：',
      '- `对比`：A 与 B / 旧方案与新方案（栏标题用 `标签：` 前缀，没标签就不画栏标题）',
      '- `时间线`：阶段、步骤、里程碑',
      '- `指标`：整页以数字为主（每张卡都要有数字）',
      '- `架构图`：讲分层、模块关系、数据流、处理流程 → 再写一行 `图：mermaid` 与一段 `flowchart TD|LR`',
      '  · 节点 `A[文字]`/`B(文字)`/`C{判断}`，连线 `A --> B`、`A -->|说明| B`；分层用 `subgraph 名 ... end`',
      '  · 节点文字 ≤12 汉字、节点 4~9 个；不要 classDef/style/click；该页要点只写 2~3 条结论',
      '- 不要写 Markdown 表格、代码块或解释性文字',
      `- **总页数目标 ${Math.max(3, Math.min(cap, Math.round(cap * 0.75)))}~${cap} 页**（含封面、目录与章节隔页）：${cap} 是上限不是目标，`,
      `  页数偏少会让长材料显得单薄；材料越厚、章节越多，越应该接近上限`,
    ].join('\n'),
    prompt: [
      input.topic === undefined || input.topic === '' ? '' : `主题：${input.topic}`,
      '以下是各片段的要点：',
      '',
      ...input.facts.map((fact, index) => `${index + 1}. ${fact}`),
    ]
      .filter((line) => line !== '')
      .join('\n'),
  }
}

/**
 * Read the facts out of a map answer: lines starting with `-` (or `•`), cleaned.
 * Tolerates a numbered list and stray prose, which models produce anyway.
 * @param {string} text
 * @returns {string[]}
 */
export function parseFacts(text) {
  const lines = String(text ?? '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
  /** @type {string[]} */
  const facts = []
  for (const line of lines) {
    const match = /^\s*(?:[-*•]|\d{1,2}[.)、])\s*(.+)$/.exec(line)
    if (match === null) continue
    const fact = match[1]
      .replace(/\*\*/g, '')
      .replace(/^[（(]本段无可用事实[）)]$/, '')
      .trim()
    if (fact === '' || fact.startsWith('（本段无')) continue
    facts.push(fact)
  }
  return facts
}

/**
 * Merge the per-block facts for the reduce step.
 *
 * De-duplication is deliberately conservative (normalised text equality + "one fact is a
 * prefix of another") because two facts that merely look similar usually carry different
 * numbers, and those numbers are exactly what we must not lose.
 *
 * @param {string[][]} groups facts per block, in block order
 * @param {{max?: number}} [options]
 * @returns {{facts: string[], dropped: number}}
 */
export function mergeFacts(groups, options = {}) {
  const max = Number.isFinite(options.max) ? Math.max(1, options.max) : 80
  const normalise = (fact) =>
    fact
      .toLowerCase()
      .replace(/[\s，,。.；;：:、\-—…"'“”‘’()（）]/g, '')
  /** @type {string[]} */
  const kept = []
  const seen = new Set()
  const total = groups.reduce((sum, group) => sum + group.length, 0)
  let dropped = 0

  for (const group of groups) {
    for (const fact of group) {
      const key = normalise(fact)
      if (key === '') {
        dropped += 1
        continue
      }
      if (seen.has(key)) {
        dropped += 1
        continue
      }
      // A shorter fact that a longer one already contains adds no information… unless
      // the longer one is just this one plus a tail, which is the same thing.
      const covered = kept.some((existing) => {
        const existingKey = normalise(existing)
        return existingKey.length > key.length && existingKey.includes(key) && key.length >= 8
      })
      if (covered) {
        dropped += 1
        continue
      }
      seen.add(key)
      kept.push(fact)
      if (kept.length >= max) {
        // Everything not kept is dropped, whether it was a duplicate or never reached.
        // `total - kept.length` is the only count that stays right when the cap and a
        // duplicate are hit on the same pass.
        return { facts: kept, dropped: total - kept.length }
      }
    }
  }
  return { facts: kept, dropped }
}

/**
 * Run `task` over `items` with at most `limit` in flight, preserving result order.
 *
 * Bounded on purpose: a 12-block document must not open 12 simultaneous model calls
 * (rate limits, and a partial failure would be harder to reason about).
 *
 * @template T, R
 * @param {T[]} items
 * @param {number} limit
 * @param {(item: T, index: number) => Promise<R>} task
 * @returns {Promise<R[]>}
 */
export async function runWithConcurrency(items, limit, task) {
  const size = Math.max(1, Math.min(limit, items.length))
  /** @type {R[]} */
  const results = new Array(items.length)
  let cursor = 0
  const workers = Array.from({ length: size }, async () => {
    for (;;) {
      const index = cursor
      cursor += 1
      if (index >= items.length) return
      results[index] = await task(items[index], index)
    }
  })
  await Promise.all(workers)
  return results
}

/**
 * Rough progress text for the transcript ("已提炼 3/7 段").
 * @param {number} done @param {number} total
 */
export function progressLabel(done, total) {
  return `已提炼 ${done}/${total} 段`
}
