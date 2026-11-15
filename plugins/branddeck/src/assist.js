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
const FALLBACK_PROVIDER = 'model-tier'

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
    maxTokens: Number.isFinite(request.maxTokens) ? request.maxTokens : 1600,
    // Some provider routes require a session identity to route the request
    // (`x-opencode-session`). Prefer the session that initiated this call when there
    // is one, otherwise mint a stable-per-panel-call id so the header is always set.
    sessionId: resolveSessionId(ctx, request.sessionId),
    ...(request.signal === undefined ? {} : { signal: request.signal }),
  }

  let text = ''
  let failure = null
  let finishKind = null
  try {
    for await (const chunk of llm.stream(options)) {
      if (chunk === null || typeof chunk !== 'object') continue
      if (chunk.type === 'text-delta' && typeof chunk.text === 'string') text += chunk.text
      else if (chunk.type === 'finish') {
        finishKind = chunk.reason?.kind ?? null
        if (chunk.reason?.kind === 'error' || chunk.reason?.kind === 'aborted') {
          failure = chunk.reason.failure ?? { code: 'llm_failed', message: 'model call failed' }
        }
      }
    }
  } catch (error) {
    throw codeError('llm_failed', error instanceof Error ? error.message : String(error))
  }

  if (failure !== null) {
    throw codeError(String(failure.code ?? 'llm_failed'), String(failure.message ?? 'model call failed'))
  }
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
 * @param {{description: string, maxSlides?: number}} input
 */
export function buildContentPrompt(input) {
  const cap = Number.isFinite(input.maxSlides) ? input.maxSlides : 10
  return {
    system: [
      '你是企业演示文稿的内容架构师。用户会用一两句话描述想做的演示，你产出可直接排版的 Markdown 大纲。',
      '严格使用这套约定：',
      '- 第一行 `# <卡组标题>`',
      '- 紧接着每个 `## <页面标题>` 下写要点；`#` 与第一个 `##` 之间**不要写任何要点**（那是封面前言，会被丢弃）',
      '- 每页 3~5 条 `- <要点>`，每条不超过 40 个汉字，一条一个观点',
      '- 需要演讲备注时写 `备注：<一句话>`',
      '- 不要写 Markdown 表格、不要写代码块、不要写解释性前后文',
      `- 总页数不超过 ${cap} 页（含封面与结尾）`,
      '内容要具体、可执行，避免空话；如果用户只给了主题，就按"现状 → 问题 → 方案 → 价值 → 行动"组织。',
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
