/**
 * In-panel conversation: turn a couple of sentences into a style or a deck.
 *
 * WHY THIS LIVES ON THE HOST: the Web Client half of a bundle plugin is a *sibling*
 * of the shipped client plugins, so it cannot reach their `connection` service (that
 * service is provided on the Connection plugin's own scope). The Host half, however,
 * sits on the same context as the `llm` adapter registry, so it can stream a model
 * call directly. The panel therefore posts a sentence to its own Host route and gets
 * a finished artefact back — no Agent turn, no session, no chat round-trip.
 *
 * Everything here is defensive: a missing LLM, an unparsable answer, or a cancelled
 * request must degrade to a clear error code the panel can show, never break the
 * plugin (that lesson cost us one full web-boot failure already).
 */

/** Provider used when the profile has no explicit default model selection. */
import {
  MATERIAL_DEFAULTS,
  buildMapPrompt,
  buildReducePrompt,
  mergeFacts,
  needsMapReduce,
  parseFacts,
  progressLabel,
  runWithConcurrency,
  splitIntoBlocks,
} from './mapreduce.js'

const FALLBACK_PROVIDER = 'model-tier'

/**
 * Ceiling for the one retry a truncated-and-empty answer gets. A reasoning model's
 * thinking is billed against `maxTokens`, so the retry needs real headroom; the cap
 * keeps a runaway retry from being expensive.
 */
const RETRY_MAX_TOKENS = 16000

/**
 * Stream one completion and return its text.
 *
 * @param {any} ctx plugin context (needs `llm`; `agentDefaultModel` is optional)
 * @param {{system: string, prompt: string, maxTokens?: number, temperature?: number, signal?: AbortSignal}} request
 * @returns {Promise<string>}
 */
export async function callModel(ctx, request) {
  const llm = ctx === null || ctx === undefined ? undefined : ctx.get?.('llm')
  if (llm === undefined || llm === null || typeof llm.stream !== 'function') {
    throw codeError('llm_unavailable', 'this Harness profile exposes no llm service to plugins')
  }

  let selection = {}
  try {
    const owner = ctx.get?.('agentDefaultModel')
    const current = owner?.currentSelection?.()
    if (current !== null && typeof current === 'object') selection = current
  } catch {
    /* selection is advisory; the provider default below still applies */
  }
  const provider = typeof selection.provider === 'string' && selection.provider !== '' ? selection.provider : FALLBACK_PROVIDER
  const model = typeof selection.model === 'string' ? selection.model : ''
  if (model === '') {
    throw codeError('llm_unavailable', 'no default model is selected in this profile')
  }

  /** @type {any} */
  const options = {
    provider,
    model,
    system: request.system,
    messages: [{ role: 'user', content: [{ type: 'text', text: request.prompt }] }],
    temperature: Number.isFinite(request.temperature) ? request.temperature : 0.3,
    // These are utility calls (extract facts / write an outline), not deep reasoning. On a
    // reasoning model the thinking budget is charged against maxTokens, and a long Chinese
    // answer then never gets emitted: we saw finish=max-tokens with zero text. Keep the
    // effort low so the budget goes to output.
    reasoningEffort: request.reasoningEffort ?? 'low',
    maxTokens: Number.isFinite(request.maxTokens) ? request.maxTokens : 1600,
    // Some provider routes require a session identity to route the request
    // (`x-opencode-session`). Prefer the session that initiated this call when there
    // is one, otherwise mint a stable-per-panel-call id so the header is always set.
    sessionId: resolveSessionId(ctx, request.sessionId),
    ...(request.signal === undefined ? {} : { signal: request.signal }),
  }

  /**
   * One streamed attempt: collect text plus how it finished.
   * @param {number} maxTokens
   * @returns {Promise<{text: string, failure: any, finishKind: string | null}>}
   */
  const runOnce = async (maxTokens) => {
    /** @type {{text: string, failure: any, finishKind: string | null}} */
    const state = { text: '', failure: null, finishKind: null }
    try {
      for await (const chunk of llm.stream({ ...options, maxTokens })) {
        if (chunk === null || typeof chunk !== 'object') continue
        if (chunk.type === 'text-delta' && typeof chunk.text === 'string') state.text += chunk.text
        else if (chunk.type === 'finish') {
          state.finishKind = chunk.reason?.kind ?? null
          if (chunk.reason?.kind === 'error' || chunk.reason?.kind === 'aborted') {
            state.failure = chunk.reason.failure ?? { code: 'llm_failed', message: 'model call failed' }
          }
        }
      }
    } catch (error) {
      throw codeError('llm_failed', error instanceof Error ? error.message : String(error))
    }
    return state
  }

  const first = await runOnce(options.maxTokens)
  let { text, failure, finishKind } = first

  // A reasoning model charges its thinking against `maxTokens`, so a long Chinese answer
  // can spend the whole budget before emitting one character (observed: finish=max-tokens
  // with empty text). Failing a five-call pipeline on that coin flip is unacceptable, so
  // exactly one retry with room to answer is allowed.
  if (failure === null && finishKind === 'max-tokens' && text.trim() === '') {
    const retryBudget = Math.min(RETRY_MAX_TOKENS, Math.max(options.maxTokens * 2, options.maxTokens + 2000))
    ;({ text, failure, finishKind } = await runOnce(retryBudget))
  }

  if (failure !== null) {
    throw codeError(String(failure.code ?? 'llm_failed'), String(failure.message ?? 'model call failed'))
  }
  // finish=max-tokens means the model was cut off, not that it failed. A long Chinese
  // outline routinely hits the cap while still holding usable lines, so return the
  // prefix and let the caller parse whatever is complete.
  if (finishKind === 'max-tokens' && text.trim() !== '') return text
  if (text.trim() === '') {
    throw codeError('llm_empty', `the model returned no text${finishKind === null ? '' : ` (finish: ${finishKind})`}`)
  }
  return text
}

/**
 * Pick the session identity for a model call.
 *
 * Order: an explicit id, the agent that initiated this work, then a generated id.
 * The generated form is deliberately readable — a provider dashboard showing
 * `branddeck-panel-…` is easier to explain than a bare uuid.
 *
 * @param {any} ctx
 * @param {string | undefined} explicit
 * @returns {string}
 */
export function resolveSessionId(ctx, explicit) {
  if (typeof explicit === 'string' && explicit !== '') return explicit
  try {
    const initiator = ctx?.get?.('agents')?.currentInitiator?.()
    const id = initiator?.session?.id
    if (typeof id === 'string' && id !== '') return id
  } catch {
    /* no initiator inside an HTTP request context, which is the normal case here */
  }
  return `branddeck-panel-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}

/**
 * Pull the first JSON object out of a model answer (tolerates ```json fences and
 * a sentence before/after the object).
 * @param {string} text
 * @returns {unknown | null}
 */
export function extractJsonObject(text) {
  if (typeof text !== 'string') return null
  const fenced = /```(?:json)?\s*([\s\S]*?)```/i.exec(text)
  const candidates = [fenced === null ? null : fenced[1], text].filter((value) => typeof value === 'string')
  for (const candidate of candidates) {
    const start = candidate.indexOf('{')
    const end = candidate.lastIndexOf('}')
    if (start === -1 || end <= start) continue
    try {
      const parsed = JSON.parse(candidate.slice(start, end + 1))
      if (parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed)) return parsed
    } catch {
      /* try the next candidate */
    }
  }
  return null
}

/** Strip a Markdown fence / leading prose so an outline answer parses cleanly. */
export function extractMarkdown(text) {
  if (typeof text !== 'string') return ''
  const fenced = /```(?:markdown|md)?\s*([\s\S]*?)```/i.exec(text)
  const body = fenced === null ? text : fenced[1]
  const firstHeading = body.search(/^#\s+/m)
  return (firstHeading === -1 ? body : body.slice(firstHeading)).trim()
}

/**
 * Prompt: a sentence or two → a style profile.
 * The schema is spelled out so the model answers with exactly the keys we validate.
 * @param {{description: string, theme: any, styleIds: string[]}} input
 */
export function buildStylePrompt(input) {
  const fonts = input.theme?.fonts ?? {}
  const colors = input.theme?.colors ?? {}
  return {
    system: [
      '你是企业演示文稿的品牌设计助理。',
      '用户会用一两句话描述想要的演示风格。你把这句话翻译成 BrandDeck 的「风格配置文件」。',
      '只输出一个 JSON 对象，不要任何解释、不要 Markdown 代码块以外的东西。',
      '字段与取值约束：',
      '- name: 中文短名（<=12 字），能一眼看出风格',
      '- description: 用一句话复述并具体化用户意图',
      '- density: "airy" | "standard" | "dense"（每页要点数量）',
      '- tone: "formal" | "bold" | "minimal"（强调色用量）',
      '- bulletLayout: "list" | "cards"',
      '- coverVariant: "left" | "centered"',
      '- accentBar: boolean（是否画品牌强调条）',
      '- closingSlide: boolean（是否加致谢页）',
      '- maxSlides: 6..30 的整数',
      '- colors: 只允许 background/surface/textPrimary/textSecondary/primary/accent，值为 "#RRGGBB"，仅在用户明确要求改色时给出，可省略',
      '- fonts: 只允许 title/body，值为字体名，仅在用户明确要求换字体时给出，可省略',
      '不要编造用户没提的字段值；不确定就用默认。',
    ].join('\n'),
    prompt: [
      `用户描述：${input.description}`,
      '',
      '当前模板的品牌令牌（可作为默认值参考，不要原样复述）：',
      `- 主色 primary=${colors.primary ?? '#C7000A'}，强调 accent=${colors.accent ?? '#E9002F'}`,
      `- 背景 ${colors.background ?? '#FFFFFF'}，容器 ${colors.surface ?? '#EBEBEB'}`,
      `- 字体：标题 ${fonts.title ?? 'Microsoft YaHei'}，正文 ${fonts.body ?? 'Arial'}`,
      '',
      '请输出风格 JSON。',
    ].join('\n'),
  }
}

/**
 * Prompt: a sentence or two → a Markdown outline in BrandDeck's own convention.
 *
 * Two rules carry most of the quality here, both learned from reviewing generated
 * decks: (1) the page carries short phrases while the detail belongs in the notes
 * ("落入备注" beats truncating), and (2) the model states the *layout intent* of a
 * slide so对比/时间线/指标 render differently instead of as three identical lists.
 *
 * @param {{description: string, maxSlides?: number}} input
 */
/**
 * The outline conventions every prompt builder shares.
 *
 * Kept in one place because the single-prompt path and the Reduce step must produce the
 * *same* document: a deck whose structure depends on which path built it is a bug the
 * reader experiences as inconsistency.
 *
 * What is asked for is 总-分-总. The opening chapter states the whole, the middle chapters
 * carry the parts, the last chapter puts them back together. `章节：` is a hard
 * requirement rather than a hint: it is what lets the composer draw the 目录 page and the
 * dividers, and it is how "the logic is unclear" gets fixed structurally instead of by
 * rewording titles.
 *
 * @param {number} cap page cap
 */
function outlineConventions(cap) {
  return [
    '结构（必须遵守，这是"总-分-总"）：',
    '- 第一行 `# <卡组标题>`；`#` 与第一个 `##` 之间**不要写要点**（那是封面前言，会被丢弃）',
    '- **每个 `##` 页块里写一行 `章节：<章节名>`**：同名章节会被编成一章，生成器据此插入目录页与章节隔页',
    '- 3~5 章为宜，每章 2~4 页',
    '- 第一章是**总**：先给结论与全景（"要解决什么、结论是什么"）',
    '- 中间各章是**分**：一章一个主题；页标题写成带结论的完整判断句，不要"XX 方案"这类空标题',
    '- 最后一章是**总**：回收结论、给出下一步（章名可用"总结"）',
    '页面与文字：',
    '- 每页 3~6 条 `- <要点>`，**每条 12~28 个汉字**，写成「主张 + 关键数据/限定条件」的短句，不要关键词清单',
    '- **若材料里有可比较的同类数字**（同一指标的几个类别 / 几个年份 / 几个部分），把这一页写成数据行而不是句子：',
    '  每条 `- <标签> — <数值><单位>`，至少 3 行、单位一致（例：`- 墨西哥 — 62.52 万辆`）。',
    '  生成器会自动选图（类别→柱状、年份→折线、加总为 100%→饼图）；单位不一致或不足 3 行时不会成图。',
    '- **单页「标题 + 要点」合计 80~125 字**（预算 150 字，标题也算在内）：读者只看这一页也要能读懂，不要把内容全推给备注',
    '- **细节不要丢**：更长的解释、案例、来龙去脉写进 `备注：<一段话>`（页面正文与备注是两条轨，各司其职）',
    '- 不要写 Markdown 表格、不要写代码块、不要写解释性前后文',
    `- 总页数不超过 ${cap} 页（含封面、目录与结尾）`,
    '版式（在该页块内写一行 `版式：<类型>`）：',
    '- `对比`：A 与 B / 旧方案与新方案 / 优缺点 → 两栏；栏标题用 `标签：` 前缀给出，没标签就不画栏标题',
    '- `时间线`：阶段、步骤、里程碑、Q1-Q4 → 编号节点',
    '- `指标`：整页以数字、百分比、金额为主 → 指标卡（每张卡都要有数字）',
    '- `架构图`：讲系统分层、模块关系、数据流、处理流程 → 见下条',
    '- 其余用 `列表`（可省略不写）',
    '架构图 / 流程图（`版式：架构图`）：',
    '- 该页除了 2~3 条要点，再写一行 `图：mermaid`，随后另起一行写 `flowchart LR`（幻灯片是 16:9 横版，优先 LR；只有严格的分层/流水线才用 TD）',
    '- 只用 flowchart：节点 `A[文字]` / `B(文字)` / `C{判断}`，连线 `A --> B`、`A -->|说明| B`',
    '- 分层用 `subgraph 分组名` … `end`（例如 接入层 / 服务层 / 存储层）',
    '- 节点文字不超过 12 个汉字，节点 4~9 个；不要用 classDef、style、click 等装饰语法',
    '- 图是这一页的主体，要点只补一句结论，不要复述图里的节点',
  ].join('\n')
}

/** @param {{description: string, maxSlides?: number}} input */
export function buildContentPrompt(input) {
  const cap = Number.isFinite(input.maxSlides) ? input.maxSlides : 10
  return {
    system: [
      '你是企业演示文稿的内容架构师。用户会用一两句话描述想做的演示，你产出可直接排版的 Markdown 大纲。',
      outlineConventions(cap),
      '内容要具体、可执行，避免空话；如果用户只给了主题，就按"现状 → 问题 → 方案 → 价值 → 行动"组织成章节。',
    ].join('\n'),
    prompt: ['用户描述：', input.description, '', '请输出 Markdown 大纲。'].join('\n'),
  }
}

/** @param {string} code @param {string} message */
export function codeError(code, message) {
  const error = new Error(message)
  error.code = code
  return error
}

/**
 * Long material -> outline, via Map-Reduce.
 *
 * One prompt cannot carry a 20,000-character report: the model anchors on the opening
 * sections and pads the rest. So above the threshold we extract facts per block (map,
 * bounded concurrency), merge them (de-duplicate, keep every number), then ask once for
 * the outline (reduce). Below the threshold the single-prompt path is kept because it is
 * cheaper and better on short input.
 *
 * @param {any} ctx
 * @param {{text: string, maxSlides?: number, topic?: string, options?: any,
 *          onProgress?: (label: string) => void, signal?: AbortSignal}} input
 * @returns {Promise<{markdown: string, blocks: number, calls: number, dropped: number}>}
 */
export async function assistOutlineFromMaterial(ctx, input) {
  const material = typeof input.text === 'string' ? input.text : ''
  if (material.trim() === '') throw codeError('invalid_body', 'material is empty')
  const options = input.options ?? {}
  const stats = { blocks: 0, calls: 0, dropped: 0 }

  // Short material: exactly the previous behaviour, one call.
  if (needsMapReduce(material, options) === false) {
    const prompt = buildContentPrompt({ description: material, maxSlides: input.maxSlides })
    const answer = await callModel(ctx, { ...prompt, maxTokens: 6000, signal: input.signal })
    stats.calls = 1
    stats.blocks = 1
    const markdown = extractMarkdown(answer)
    return { ...(await finishWithDensify(ctx, markdown, options, input.signal, stats)), ...stats }
  }

  const blocks = splitIntoBlocks(material, options)
  if (blocks.length === 0) throw codeError('empty_source', 'material produced no blocks')
  stats.blocks = blocks.length

  let done = 0
  const factsPerBlock =
    Number.isFinite(options.factsPerBlock) ? options.factsPerBlock : MATERIAL_DEFAULTS.factsPerBlock
  const groups = await runWithConcurrency(
    blocks,
    Number.isFinite(options.concurrency) ? options.concurrency : MATERIAL_DEFAULTS.concurrency,
    async (block) => {
      const prompt = buildMapPrompt({ block, total: blocks.length, factsPerBlock, deckTitle: input.topic })
      const answer = await callModel(ctx, {
        ...prompt,
        maxTokens: 2200,
        temperature: 0.2,
        signal: input.signal,
      })
      stats.calls += 1
      done += 1
      if (typeof input.onProgress === 'function') input.onProgress(progressLabel(done, blocks.length))
      return parseFacts(answer)
    },
  )

  const merged = mergeFacts(groups, { max: Math.max(24, blocks.length * factsPerBlock) })
  stats.dropped = merged.dropped
  if (merged.facts.length === 0) {
    throw codeError('empty_source', 'the map phase produced no facts — the material may be unreadable')
  }

  const reducePrompt = buildReducePrompt({
    facts: merged.facts,
    maxSlides: input.maxSlides,
    topic: input.topic,
  })
  const answer = await callModel(ctx, { ...reducePrompt, maxTokens: 8000, signal: input.signal })
  stats.calls += 1
  const markdown = extractMarkdown(answer)
  return { ...(await finishWithDensify(ctx, markdown, options, input.signal, stats)), ...stats }
}

/**
 * One densify pass, then the outline the caller asked for.
 *
 * It lives here rather than in `index.js` because every entry point (command, tool, panel)
 * must get the same treatment and because the call has to be counted in `pipeline.calls`.
 * Failure is swallowed on purpose: a fuller outline is a quality improvement, never a
 * reason to lose a deck that was already generated.
 *
 * @param {any} ctx
 * @param {string} markdown
 * @param {any} options material options (also carries the densify flags)
 * @param {AbortSignal | undefined} signal
 * @param {{calls: number}} stats
 */
async function finishWithDensify(ctx, markdown, options, signal, stats) {
  if (options.densify === false) return { markdown, densified: false }
  try {
    const result = await densifyOutline(ctx, {
      markdown,
      charBudget: options.charBudget,
      floor: options.pageCharFloor,
      signal,
    })
    if (result.densified) stats.calls += 1
    return { markdown: result.markdown, densified: result.densified }
  } catch {
    return { markdown, densified: false }
  }
}


/* ------------------------------------------------------------------ *
 * Outline densification ("填充细节")
 * ------------------------------------------------------------------ */

/** Metadata lines carry structure, not copy — they never count toward a page's text. */
const OUTLINE_META = /^(?:版式|排版|layout|章节|章|chapter|备注|notes?|图|示意图|diagram|mermaid)\s*[:：]/i
/** The one metadata line that opens a multi-line diagram block. */
const DIAGRAM_META = /^(?:图|示意图|diagram|mermaid)\s*[:：]/i

/**
 * Measure the visible copy of an outline, page by page.
 *
 * "Visible" means what a reader sees on the slide: the page title plus its bullets.
 * Metadata (`版式：` / `章节：` / `备注：` / `图：`) and Mermaid source lines are excluded —
 * counting them would make a page look dense when it is not, which is exactly the mistake
 * that produced 20-character pages while the budget said 150.
 *
 * Pure and dependency-free so it can be unit-tested and reused by the panel.
 *
 * @param {string} markdown
 * @returns {{pages: Array<{title: string, chars: number, bullets: number}>, total: number}}
 */
export function measureOutline(markdown) {
  const pages = []
  /** @type {{title: string, chars: number, bullets: number} | null} */
  let current = null
  let inFence = false
  // A `图：mermaid` block may be fenced or bare. A bare one has no closing marker, so it
  // is tracked explicitly: otherwise every `flowchart` / `A --> B` line counts as copy and
  // a thin diagram page looks dense (which would hide it from the densify pass).
  let inDiagram = false
  for (const line of String(markdown ?? '').replace(/\r\n?/g, '\n').split('\n')) {
    const trimmed = line.trim()
    if (trimmed.startsWith('```')) {
      inFence = !inFence
      inDiagram = false
      continue
    }
    if (inFence) continue
    const heading = /^##\s+(.*)$/.exec(trimmed)
    if (heading !== null) {
      inDiagram = false
      const title = heading[1].replace(/\{#.*?\}/g, '').trim()
      current = { title, chars: charCount(title), bullets: 0 }
      pages.push(current)
      continue
    }
    if (trimmed === '') {
      inDiagram = false
      continue
    }
    if (trimmed.startsWith('#')) {
      inDiagram = false
      continue
    }
    if (OUTLINE_META.test(trimmed)) {
      // `图：` opens a diagram; the other metadata lines are single lines.
      inDiagram = DIAGRAM_META.test(trimmed)
      continue
    }
    if (inDiagram) continue
    if (current === null) continue
    const bullet = /^(?:[-*+]|\d{1,2}[.)、])\s*(.+)$/.exec(trimmed)
    if (bullet !== null) {
      current.chars += charCount(bullet[1])
      current.bullets += 1
      continue
    }
    // A plain paragraph inside a page is still visible copy.
    current.chars += charCount(trimmed)
  }
  const total = pages.reduce((sum, page) => sum + page.chars, 0)
  return { pages, total }
}

/** @param {string} text */
function charCount(text) {
  return Array.from(String(text ?? '')).length
}

/**
 * Pages that are below the floor, i.e. the ones the densify pass should rewrite.
 *
 * A page with a diagram is deliberately sparse (the figure is the content), so it is not
 * counted as thin — expanding it would fight the layout.
 *
 * @param {string} markdown
 * @param {{floor?: number}} [options]
 */
export function thinPages(markdown, options = {}) {
  const floor = Number.isFinite(options.floor) ? Math.max(1, Math.floor(options.floor)) : DENSIFY.floor
  const { pages } = measureOutline(markdown)
  return pages
    .map((page, index) => ({ ...page, index }))
    .filter((page) => page.bullets > 0 && page.chars < floor)
}

/** Densify knobs. `floor` is a *floor*, not a target: the target is the page budget. */
export const DENSIFY = Object.freeze({
  /** A page below this many visible characters is considered thin. */
  floor: 70,
  /** Never ask for more than this per page, whatever the budget says. */
  ceiling: 150,
})

/**
 * Prompt for the densify pass: thin pages → the same pages with fuller copy.
 *
 * The pass is deliberately conservative about structure: titles, chapters, layout intents
 * and diagrams must come back unchanged, otherwise the deck the user approved is not the
 * deck they get. Only the bullet text may grow.
 *
 * @param {{markdown: string, pages: Array<{index: number, title: string, chars: number}>, charBudget?: number}} input
 */
export function buildEnrichPrompt(input) {
  const budget = Number.isFinite(input.charBudget)
    ? Math.min(DENSIFY.ceiling, Math.max(40, Math.floor(input.charBudget)))
    : DENSIFY.ceiling
  const list = input.pages.map((page) => `- 第 ${page.index + 1} 页「${page.title}」（当前 ${page.chars} 字）`).join('\n')
  return {
    system: [
      '你在给一份演示文稿的大纲"补细节"。上一轮大纲的页面太薄了：读者只看页面读不懂，必须翻备注。',
      '你的任务：把指定页面的要点写充实，同时**严格保持结构不变**。',
      '必须遵守：',
      '- 保持 `# 卡组标题`、每页 `## 页面标题`、`章节：`、`版式：`、`图：` 与其后的 mermaid 代码**完全不变**',
      '- 只重写这些页面的 `- ` 要点；页数、页序、每页标题都不许变',
      `- 每页正文（标题+要点）写到 **${Math.round(budget * 0.55)}~${Math.round(budget * 0.85)} 字**，每条要点 12~28 个汉字，3~6 条`,
      '- 要点写成「主张 + 关键数据/限定条件」的短句，不要关键词堆砌，不要空话（禁止"赋能/抓手/闭环/矩阵/重塑/全链路"）',
      '- 只能用原文材料里出现过的事实、数字、人名、工具名；不许编造，没料就少写一条',
      '- 保留每页已有的 `备注：` 行（可补充，不要删除已有关键信息）',
      '- 只输出完整大纲，不要解释、不要代码块围栏',
    ].join('\n'),
    prompt: [
      '需要补细节的页面：',
      list,
      '',
      '当前完整大纲：',
      input.markdown,
      '',
      '请输出补细节后的完整大纲。',
    ].join('\n'),
  }
}

/**
 * Run the densify pass once and return the (possibly) fuller outline.
 *
 * Failing soft is the whole point: this runs *after* a successful outline, so a bad or
 * truncated answer must never destroy the deck. The result is accepted only when it keeps
 * every page title, keeps the page count, and is actually denser; otherwise the original
 * markdown comes back untouched and the caller reports `densified: false`.
 *
 * @param {any} ctx
 * @param {{markdown: string, charBudget?: number, floor?: number, signal?: AbortSignal}} input
 * @returns {Promise<{markdown: string, densified: boolean, pages: number, before: number, after: number}>}
 */
export async function densifyOutline(ctx, input) {
  const original = String(input.markdown ?? '')
  const thin = thinPages(original, { floor: input.floor })
  const before = measureOutline(original)
  if (thin.length === 0) {
    return { markdown: original, densified: false, pages: 0, before: before.total, after: before.total }
  }
  const prompt = buildEnrichPrompt({
    markdown: original,
    pages: thin,
    charBudget: input.charBudget,
  })
  const answer = await callModel(ctx, {
    ...prompt,
    // The densify answer is the whole outline again, rewritten denser: it needs more room
    // than the outline it replaces, not the same amount.
    maxTokens: 8000,
    temperature: 0.3,
    signal: input.signal,
  })
  const candidate = extractMarkdown(answer)
  const after = measureOutline(candidate)
  const sameShape =
    after.pages.length === before.pages.length &&
    after.pages.every((page, index) => page.title === before.pages[index].title)
  if (sameShape === false || after.total <= before.total) {
    return { markdown: original, densified: false, pages: thin.length, before: before.total, after: before.total }
  }
  return { markdown: candidate, densified: true, pages: thin.length, before: before.total, after: after.total }
}
