/**
 * BrandDeck — Host half.
 *
 * Responsibilities (see ARCHITECTURE.md):
 *  1. Agent tools (`branddeck_*`): parse a corporate .pptx template into theme
 *     tokens, store card decks, and compile decks into native editable .pptx.
 *  2. An authenticated Connection RPC channel (`/branddeck`) consumed by the Web panel.
 *  3. A short system-prompt section teaching the deck DSL.
 *
 * Everything numeric lives in the shared contract (`src/schema.js`) and the
 * engine modules under `src/`; this file only wires capabilities together.
 */
import { mkdir, readFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { join, resolve as resolvePath } from 'node:path'

import { DEFAULT_THEME } from './src/defaults.js'
import {
  CANVAS,
  SAFE_AREA,
  parseAndRepairDeck,
  repairDeck,
  resolveColor,
  estimateTextOverflow,
} from './src/schema.js'
import * as store from './src/store.js'
import * as styles from './src/styles.js'
import { parsePptxTemplate } from './src/template.js'
import { compileDeckToFile } from './src/compile.js'

/** Cordis services this plugin cannot work without. */
export const inject = ['tools']

/** Tool names, exported so tests and docs cannot drift. */
export const TOOL_NAMES = Object.freeze({
  templateParse: 'branddeck_template_parse',
  templateList: 'branddeck_template_list',
  templateSpec: 'branddeck_template_spec',
  themeGet: 'branddeck_theme_get',
  deckCreate: 'branddeck_deck_create',
  deckSave: 'branddeck_deck_save',
  deckGet: 'branddeck_deck_get',
  deckList: 'branddeck_deck_list',
  deckExport: 'branddeck_deck_export',
  outlineGet: 'branddeck_outline_get',
  outlineSave: 'branddeck_outline_save',
  deckMaterialize: 'branddeck_deck_materialize',
  styleList: 'branddeck_style_list',
  styleSave: 'branddeck_style_save',
  generate: 'branddeck_generate',
})

const PROMPT_ORDER = 2650

/**
 * @param {import('@deepseek-ai/cordis').Context} ctx
 * @param {Record<string, unknown>} [rawConfig]
 */
export function apply(ctx, rawConfig = {}) {
  const config = normalizeConfig(rawConfig)
  const storeConfig = { root: config.dataDir }

  ctx.logger?.info?.(`branddeck: store root ${store.storeRoot(storeConfig)}`)

  registerTools(ctx, config, storeConfig)
  registerPanelApi(ctx, config, storeConfig)
  registerCommands(ctx, config, storeConfig)
  if (config.promptSection) registerPromptSection(ctx, config)
}

/* ------------------------------------------------------------------ *
 * Session command: /branddeck
 * ------------------------------------------------------------------ */

/** `/branddeck` sub-command help, kept next to the parser so they cannot drift. */
const COMMAND_USAGE = [
  '用法：',
  '  /branddeck deck <一句话描述>     生成卡组（用当前模板与默认风格）',
  '  /branddeck deck --style <id>     指定风格生成',
  '  /branddeck deck --file <路径>    用一份长材料（.md/.txt）生成，可附一句主题',
  '  /branddeck style <一句话描述>    生成并保存一份风格配置',
  '  /branddeck list                  列出卡组',
  '  /branddeck templates             列出已解析模板',
  '  /branddeck                       查看本帮助与当前状态',
].join('\n')

/**
 * Parse the arguments of `/branddeck deck …`.
 *
 * Two flags, order-independent, because they are written by hand:
 *   --style <id>   pick a style profile
 *   --file <path>  read the material from a file (the long-document entry)
 * Everything left over is the description; when a file is given, the description is an
 * optional topic hint rather than the content itself.
 *
 * @param {string} tail text after `deck`
 */
export function parseDeckArguments(tail) {
  let rest = String(tail ?? '').trim()
  /** @type {{kind: 'deck', description: string, styleId?: string, materialPath?: string}} */
  const result = { kind: 'deck', description: '' }

  const styleFlag = /(?:^|\s)--style\s+(\S+)\s*/i.exec(rest)
  if (styleFlag !== null) {
    result.styleId = styleFlag[1]
    rest = rest.replace(styleFlag[0], ' ').trim()
  }
  const fileFlag = /(?:^|\s)--file\s+("([^"]+)"|\S+)\s*/i.exec(rest)
  if (fileFlag !== null) {
    result.materialPath = (fileFlag[2] ?? fileFlag[1]).trim()
    rest = rest.replace(fileFlag[0], ' ').trim()
  }
  result.description = rest
  return result
}

/**
 * Parse `/branddeck …` into `{ kind, description, styleId }`.
 *
 * The bare form (`/branddeck 给周会做一页`) means "generate a deck", because that is
 * what someone types 90% of the time.
 *
 * @param {string} rawInput text after the command name
 */
export function parseBrandDeckCommand(rawInput) {
  const text = String(rawInput ?? '').trim()
  if (text === '') return { kind: 'help' }
  const [first, ...rest] = text.split(/\s+/)
  const head = first.toLowerCase()
  const tail = text.slice(first.length).trim()

  if (head === 'help' || head === '?' || head === '--help' || head === '-h') return { kind: 'help' }
  if (head === 'list' || head === 'ls' || head === 'decks') return { kind: 'list' }
  if (head === 'templates' || head === 'template' || head === 'tpl') return { kind: 'templates' }
  if (head === 'style' || head === 'styles' || head === 'theme') {
    return tail === '' ? { kind: 'style-help' } : { kind: 'style', description: tail }
  }
  if (head === 'deck' || head === 'gen' || head === 'generate' || head === 'make') {
    return parseDeckArguments(tail)
  }
  if (head === 'file' || head === 'from') {
    // `/branddeck file <path> [说明]` — the long-material entry.
    const parts = tail.split(/\s+/).filter((part) => part !== '')
    if (parts.length === 0) return { kind: 'deck-help-file' }
    return { kind: 'deck', materialPath: parts[0], description: parts.slice(1).join(' ') }
  }
  if (head.startsWith('--style') || head.startsWith('--file')) {
    return parseDeckArguments(text)
  }
  if (head.startsWith('--style')) {
    const hint = /^--style\s+(\S+)\s*/i.exec(text)
    return hint === null
      ? { kind: 'deck', description: rest.join(' ') }
      : { kind: 'deck', styleId: hint[1], description: text.slice(hint[0].length).trim() }
  }
  // Bare text: treat the whole line as the deck description.
  return { kind: 'deck', description: text }
}

/**
 * Register the session-native conversation entry.
 *
 * This is deliberately a *command*, not a panel widget: it runs in the current
 * session, answers in the transcript, and can use the session identity for the
 * model call. `definitionId` is package-scoped so a later version can evolve the
 * behaviour without colliding with other commands named `branddeck`.
 */
function registerCommands(ctx, config, storeConfig) {
  ctx.inject(['commands'], (scope) => {
    scope.effect(
      () =>
        scope.commands.register({
          definitionId: '@local/dsh-branddeck',
          name: 'branddeck',
          description: '从一句话生成品牌演示卡组或风格配置（BrandDeck）',
          input: { hint: '[deck <描述>|style <描述>|list|templates]' },
          handler: (invocation) =>
            executeBrandDeckCommand(invocation, { ctx, config, storeConfig }).catch((error) => ({
              kind: 'error',
              text: `BrandDeck 执行失败：${error instanceof Error ? error.message : String(error)}`,
            })),
        }),
      'branddeck: /branddeck command',
    )
  })
}

/**
 * Execute one `/branddeck …` invocation and render the transcript answer.
 * @param {{agent: any, rawInput: string}} invocation
 * @param {{ctx: any, config: any, storeConfig: {root?: string}}} deps
 */
async function executeBrandDeckCommand(invocation, deps) {
  const parsed = parseBrandDeckCommand(invocation?.rawInput)
  const { config, storeConfig } = deps
  const sessionId = (() => {
    try {
      const id = invocation?.agent?.session?.id
      return typeof id === 'string' ? id : undefined
    } catch {
      return undefined
    }
  })()

  const status = async () => {
    const templates = await store.listTemplates(storeConfig)
    const decks = await store.listDecks(storeConfig)
    const styleList = styles.listStyles(storeConfig)
    return [
      `模板 ${templates.length} 个${templates[0] === undefined ? '' : `（最近：${templates[0].templateId}）`}`,
      `卡组 ${decks.length} 个${decks[0] === undefined ? '' : `（最近：${decks[0].deckId}，${decks[0].slideCount} 页）`}`,
      `风格 ${styleList.length} 个`,
    ].join(' · ')
  }

  if (parsed.kind === 'help') {
    return {
      kind: 'success',
      text: [`BrandDeck —— 一句话生成品牌演示。`, '', COMMAND_USAGE, '', `当前：${await status()}`].join('\n'),
    }
  }

  if (parsed.kind === 'style-help') {
    return {
      kind: 'error',
      text: `风格需要在描述后给出内容，例如：\n  /branddeck style 华为红、信息密度高、正文卡片式\n\n${COMMAND_USAGE}`,
    }
  }

  if (parsed.kind === 'list') {
    const decks = await store.listDecks(storeConfig)
    if (decks.length === 0) return { kind: 'success', text: '还没有卡组。试试：/branddeck deck <一句话描述>' }
    return {
      kind: 'success',
      text: ['卡组：', ...decks.map((entry) => `  · ${entry.deckId} — ${entry.title}（${entry.slideCount} 页，r${entry.revision ?? 0}）`)].join('\n'),
    }
  }

  if (parsed.kind === 'templates') {
    const templates = await store.listTemplates(storeConfig)
    if (templates.length === 0) {
      return { kind: 'success', text: '还没有模板。先运行：branddeck_template_parse { path: "<.pptx 路径>" }' }
    }
    return {
      kind: 'success',
      text: ['模板：', ...templates.map((entry) => `  · ${entry.templateId} — ${entry.templateName}（${entry.colors?.primary ?? ''}，版式 ${entry.hasLayouts ? '已格式化' : '未格式化'}）`)].join('\n'),
    }
  }

  if (parsed.kind === 'style') {
    const templateId = config.defaultTemplateId ?? (await firstTemplateId(storeConfig)) ?? null
    if (templateId === null) {
      return { kind: 'error', text: '还没有模板，无法对齐品牌。先运行 branddeck_template_parse。' }
    }
    const theme = await themeFor({ templateId }, config, storeConfig)
    let draft
    try {
      draft = await assistStyle({ description: parsed.description, theme }, { ctx: deps.ctx, storeConfig })
    } catch (error) {
      return { kind: 'error', text: `调用模型失败：${error instanceof Error ? error.message : String(error)}` }
    }
    if (draft.ok !== true) {
      return { kind: 'error', text: `没听懂这个风格：${draft.error.message}` }
    }
    const saved = styles.writeStyle({ ...draft.style, templateId }, storeConfig)
    return {
      kind: 'success',
      text: [
        `已保存风格「${saved.name}」(${saved.styleId})`,
        `  密度 ${saved.density} · 强调 ${saved.tone} · 正文 ${saved.bulletLayout} · 上限 ${saved.maxSlides} 页`,
        saved.colors?.primary === undefined ? '' : `  主色 ${saved.colors.primary}`,
        `  用它可以：/branddeck deck --style ${saved.styleId} <一句话描述>`,
      ]
        .filter((line) => line !== '')
        .join('\n'),
    }
  }

  // kind === 'deck'
  const description = typeof parsed.description === 'string' ? parsed.description.trim() : ''
  const materialPath = typeof parsed.materialPath === 'string' ? parsed.materialPath.trim() : ''
  if (description === '' && materialPath === '') {
    return { kind: 'error', text: `请给一句话描述，或用 --file 指定材料文件。\n\n${COMMAND_USAGE}` }
  }

  // `--file`: the long-document path. Read it here so the pipeline can report a clear
  // "cannot read" instead of failing inside the model call.
  let material = ''
  if (materialPath !== '') {
    try {
      material = await readFile(resolvePath(materialPath), 'utf8')
    } catch {
      return { kind: 'error', text: `读不到材料文件：${materialPath}` }
    }
    if (material.trim() === '') return { kind: 'error', text: `材料文件是空的：${materialPath}` }
  }
  const templateId = config.defaultTemplateId ?? (await firstTemplateId(storeConfig)) ?? null
  if (templateId === null) {
    return { kind: 'error', text: '还没有模板，无法对齐品牌。先运行 branddeck_template_parse。' }
  }
  let generated
  try {
    generated = await assistContent(
      {
        material: material !== '' ? material : description,
        description,
        title: material !== '' ? description : undefined,
        templateId,
        styleId: parsed.styleId,
        maxSlides: config.commandMaxSlides,
      },
      { ctx: deps.ctx, config, storeConfig, sessionId },
    )
  } catch (error) {
    return { kind: 'error', text: `调用模型失败：${error instanceof Error ? error.message : String(error)}` }
  }
  if (generated.ok !== true) {
    // Distinguish "the model could not be called" from "the model answered, but the
    // answer did not become a deck": the two need different next actions (check the
    // profile's llm wiring vs. rephrase the material).
    const code = typeof generated.error?.code === 'string' ? generated.error.code : ''
    const reason = generated.error?.message ?? '未知原因'
    return {
      kind: 'error',
      text: code.startsWith('llm') ? `调用模型失败：${reason}` : `没有生成出卡组：${reason}`,
    }
  }
  const warnings = Array.isArray(generated.warnings) ? generated.warnings.slice(0, 3) : []
  return {
    kind: 'success',
    text: [
      `已生成卡组「${generated.deck.title}」(${generated.deckId})，共 ${generated.slideCount} 页。`,
      generated.styleId === undefined ? '' : `  风格 ${generated.styleId} · 模板 ${generated.templateId}`,
      warnings.length > 0 ? `  提示：${warnings.join('；')}` : '',
      generated.pipeline === undefined
        ? ''
        : generated.pipeline.mode === 'map-reduce'
          ? `  材料 ${generated.pipeline.blocks} 段 · ${generated.pipeline.calls} 次模型调用（Map-Reduce 分层提炼）`
          : '  单次生成（材料不长，无需分层）',
      '',
      '下一步：',
      `  在侧栏 BrandDeck 面板查看/微调，或 /branddeck list 查看全部`,
      `  导出：branddeck_deck_export { deckId: "${generated.deckId}" }`,
    ]
      .filter((line) => line !== '')
      .join('\n'),
  }
}

/**
 * Validate + default the row config. The plugin deliberately ships no `Config`
 * export: keeping config plain keeps this package dependency-free and means the
 * user's `cordis.patch.yml` layer survives upgrades untouched.
 * @param {Record<string, unknown>} raw
 */
function normalizeConfig(raw) {
  const dataDir =
    typeof raw.dataDir === 'string' && raw.dataDir.trim() !== ''
      ? resolvePath(raw.dataDir)
      : undefined
  const fontFiles = Array.isArray(raw.fontFiles)
    ? raw.fontFiles.filter((entry) => typeof entry === 'string')
    : []
  const fontDirs = Array.isArray(raw.fontDirs)
    ? raw.fontDirs.filter((entry) => typeof entry === 'string')
    : []
  return {
    dataDir,
    fontFiles,
    fontDirs,
    loadSystemFonts: raw.loadSystemFonts !== false,
    svgRasterScale: clampNumber(raw.svgRasterScale, 1, 6, 3),
    promptSection: raw.promptSection !== false,
    logoSide: raw.logoSide === 'top-right' || raw.logoSide === 'none' ? raw.logoSide : 'auto',
    defaultTemplateId: typeof raw.defaultTemplateId === 'string' ? raw.defaultTemplateId : undefined,
    commandMaxSlides: clampNumber(raw.commandMaxSlides, 3, 30, 10),
    // Page copy budgets. Both used to be module constants (150 chars per page, 90 per
    // bullet); they shape the *generated* deck, so a deployment that writes denser or
    // sparser decks must be able to move them. Both are also accepted per request.
    slideCharBudget: clampNumber(raw.slideCharBudget, 40, 600, 150),
    bulletCharBudget: clampNumber(raw.bulletCharBudget, 20, 200, 90),
    // Long-material pipeline (see src/mapreduce.js). Above materialThreshold chars the
    // outline is built Map-Reduce instead of from one oversized prompt.
    materialThreshold: clampNumber(raw.materialThreshold, 500, 200000, 6000),
    materialBlockChars: clampNumber(raw.materialBlockChars, 500, 40000, 4000),
    materialMaxBlocks: clampNumber(raw.materialMaxBlocks, 1, 60, 12),
    materialConcurrency: clampNumber(raw.materialConcurrency, 1, 8, 3),
    materialFactsPerBlock: clampNumber(raw.materialFactsPerBlock, 3, 20, 8),
  }
}

function clampNumber(value, min, max, fallback) {
  const number = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(number)) return fallback
  return Math.min(max, Math.max(min, number))
}

/**
 * The copy budgets every repair/compose call needs, config-first with a per-request
 * override.
 *
 * Kept in one function because a repair that used a different budget from the compose is
 * how a deck ends up "saved clean" and then warned about on export: the panel, the tools
 * and the command must all measure the same page.
 *
 * @param {any} config resolved plugin config
 * @param {any} [body] request body (may override either budget)
 * @returns {{slideCharBudget: number, bulletCharBudget: number}}
 */
function budgetOptions(config, body) {
  const source = body !== null && typeof body === 'object' ? body : {}
  const raw = config !== null && typeof config === 'object' ? config : {}
  return {
    slideCharBudget: clampNumber(
      source.slideCharBudget ?? raw.slideCharBudget,
      40,
      600,
      150,
    ),
    bulletCharBudget: clampNumber(
      source.bulletCharBudget ?? raw.bulletCharBudget,
      20,
      200,
      90,
    ),
  }
}

/* ------------------------------------------------------------------ *
 * Helpers
 * ------------------------------------------------------------------ */

/** Envelope used by both the tools and the HTTP API. */
const ok = (value) => ({ ok: true, value })
const fail = (code, message, details) => ({
  ok: false,
  error: { code, message, ...(details === undefined ? {} : { details }) },
})

/** Turn a thrown error into a stable failure envelope. */
function failureOf(error) {
  const code =
    typeof error?.code === 'string'
      ? error.code
      : typeof error?.name === 'string' && error.name !== 'Error'
        ? error.name
        : 'branddeck_error'
  const message = error instanceof Error ? error.message : String(error)
  return fail(code, message, Array.isArray(error?.details) ? error.details : undefined)
}

/** Resolve the theme to use for a deck: its own template, else default. */
async function themeFor(deck, config, storeConfig) {
  const templateId = deck?.templateId
  if (typeof templateId === 'string' && templateId !== '') {
    try {
      const template = await store.readTemplate(templateId, storeConfig)
      if (template?.theme) return template.theme
    } catch {
      /* fall through to the default theme */
    }
  }
  if (config.defaultTemplateId !== undefined) {
    try {
      const template = await store.readTemplate(config.defaultTemplateId, storeConfig)
      if (template?.theme) return template.theme
    } catch {
      /* ignore */
    }
  }
  return DEFAULT_THEME
}

/** Read a deck, or null when it does not exist / no longer parses. */
function readDeckOrNull(deckId, storeConfig) {
  try {
    return store.readDeck(deckId, storeConfig)
  } catch {
    return null
  }
}

/** Deck to show when the caller did not name one: most recently written. */
async function activeDeck(storeConfig) {
  const decks = await store.listDecks(storeConfig)
  if (decks.length === 0) return null
  return readDeckOrNull(decks[0].deckId, storeConfig)
}

/** Directory a tool call should write deliverables into (the session cwd). */
function workspaceOf(exec) {
  const agent = exec?.agent
  const cwd = agent?.session?.header?.cwd
  if (typeof cwd === 'string' && cwd !== '') return cwd
  return process.cwd()
}

/* ------------------------------------------------------------------ *
 * Agent tools
 * ------------------------------------------------------------------ */

/**
 * Register one tool against the Host registry.
 *
 * `ctx.tools.register` consumes a plain ToolDefinition (`name`, `description`,
 * JSON-Schema `parameters`, `output { schema, render }`, `execute`). The body
 * returns `{ value }`; thrown errors become a failure envelope in the tool
 * result instead of aborting the agent turn.
 *
 * @param {(args: any, exec: any) => Promise<{ value: unknown, text?: string }>} body
 */
function registerTool(ctx, name, description, parameters, body) {
  ctx.tools.register({
    name,
    description,
    parameters,
    output: {
      schema: { type: 'object', additionalProperties: true },
      render(args, value) {
        const text =
          typeof value?.text === 'string' ? value.text : JSON.stringify(value ?? null, null, 2)
        return [{ type: 'text', text }]
      },
    },
    async execute(args, exec) {
      try {
        const result = await body(args ?? {}, exec)
        return {
          ...result.value,
          text: result.text ?? JSON.stringify(result.value, null, 2),
        }
      } catch (error) {
        const envelope = failureOf(error)
        return {
          ...envelope,
          text: `BrandDeck error (${envelope.error.code}): ${envelope.error.message}`,
        }
      }
    },
  })
}

function registerTools(ctx, config, storeConfig) {
  const register = (name, description, parameters, body) =>
    registerTool(ctx, name, description, parameters, body)

  register(
    TOOL_NAMES.templateParse,
    'Parse a corporate .pptx template into BrandDeck design tokens (colors, CJK/Latin fonts, master logo) and register it for deck authoring. Run this once per template before creating a deck from it.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        path: {
          type: 'string',
          description: 'Absolute path (or workspace-relative path) of the .pptx template to parse.',
        },
        templateId: {
          type: 'string',
          description: 'Stable id to store the template under. Defaults to a slug of the file name.',
        },
        templateName: {
          type: 'string',
          description: 'Human-readable template name shown in the panel.',
        },
      },
      required: ['path'],
    },
    async (args, exec) => {
      const cwd = workspaceOf(exec)
      const templatePath = resolvePath(cwd, args.path)
      const bytes = await readFile(templatePath)
      const templateName =
        typeof args.templateName === 'string' && args.templateName.trim() !== ''
          ? args.templateName.trim()
          : templatePath.split(/[\\/]/).pop().replace(/\.pptx$/i, '')
      const templateId =
        typeof args.templateId === 'string' && args.templateId.trim() !== ''
          ? args.templateId.trim()
          : slugify(templateName)
      const parsed = await parsePptxTemplate({
        bytes,
        path: templatePath,
        templateId,
        templateName,
        storeConfig,
      })
      await store.writeTemplate(
        {
          templateId,
          templateName,
          tokens: parsed.theme,
          logoBytes: parsed.logoBytes,
          meta: parsed.meta,
          media: parsed.media,
          sourcePath: templatePath,
          sourceBytes: bytes,
          layoutsInventory: parsed.layouts ?? null,
          markdown: parsed.markdown ?? null,
        },
        storeConfig,
      )
      const theme = parsed.theme
      return {
        value: {
          ok: true,
          templateId,
          templateName,
          colors: theme.colors,
          fonts: theme.fonts,
          hasLogo: Boolean(parsed.logoBytes),
          logoPath: theme.brandAssets?.logoPath ?? null,
          mediaCount: parsed.media?.length ?? 0,
          diagnostics: parsed.meta?.diagnostics ?? [],
        },
      }
    },
  )

  register(
    TOOL_NAMES.templateList,
    'List the .pptx templates already parsed into BrandDeck tokens, newest first.',
    { type: 'object', additionalProperties: false, properties: {} },
    async () => ({
      value: { ok: true, templates: await store.listTemplates(storeConfig) },
    }),
  )

  register(
    TOOL_NAMES.themeGet,
    'Read the design tokens of a parsed template (or of the default template) so a deck can reference $colors.* and the brand fonts.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        templateId: { type: 'string', description: 'Template id from branddeck_template_list.' },
      },
    },
    async (args) => {
      const templates = await store.listTemplates(storeConfig)
      const templateId = args.templateId ?? config.defaultTemplateId ?? templates[0]?.templateId
      if (templateId === undefined) {
        return { value: { ok: true, theme: DEFAULT_THEME, source: 'default' } }
      }
      try {
        const template = await store.readTemplate(templateId, storeConfig)
        return { value: { ok: true, theme: template.theme, source: 'template' } }
      } catch {
        return { value: { ok: true, theme: DEFAULT_THEME, source: 'default' } }
      }
    },
  )

  register(
    TOOL_NAMES.templateSpec,
    'Read the formatted, reusable spec of a parsed template: a Markdown brand sheet (palette, fonts, page size, layout list, logo, media) plus the layouts inventory (masters/layouts/placeholder geometry). Re-parses the stored source copy automatically when an older template lacks it.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        templateId: {
          type: 'string',
          description: 'Template id from branddeck_template_list; omit for the default/most recent template.',
        },
      },
    },
    async (args) => {
      const templates = await store.listTemplates(storeConfig)
      const templateId = args.templateId ?? config.defaultTemplateId ?? templates[0]?.templateId
      if (templateId === undefined) return { value: fail('template_not_found', 'No template has been parsed yet') }
      const resolved = await ensureTemplateDoc(templateId, storeConfig)
      if (resolved === null) return { value: fail('template_not_found', `template "${templateId}" not found`) }
      return {
        value: {
          ok: true,
          templateId,
          templateName: resolved.theme.templateName,
          hasLayouts: resolved.layouts !== null,
          markdown: resolved.markdown,
          layouts: resolved.layouts,
        },
      }
    },
  )

  register(
    TOOL_NAMES.deckCreate,
    `Create a new BrandDeck card deck (1280x720 virtual canvas) with a starter cover slide. Fill it with elements via ${TOOL_NAMES.deckSave}, or open the BrandDeck panel in the sidebar to edit visually.`,
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        title: { type: 'string', description: 'Deck title (also used for the cover headline).' },
        templateId: {
          type: 'string',
          description: 'Template id whose tokens this deck follows; omit for the default theme.',
        },
        deckId: { type: 'string', description: 'Optional explicit deck id.' },
      },
      required: ['title'],
    },
    async (args) => {
      const templateId =
        args.templateId ?? config.defaultTemplateId ?? (await firstTemplateId(storeConfig))
      const theme =
        templateId === undefined
          ? DEFAULT_THEME
          : await themeFor({ templateId }, config, storeConfig)
      const deck = await store.createDeck(
        {
          title: args.title,
          templateId: theme.templateId,
          deckId: args.deckId,
          theme,
        },
        storeConfig,
      )
      return { value: { ok: true, deck, theme } }
    },
  )

  register(
    TOOL_NAMES.deckSave,
    `Validate and save a complete CardDeck. Out-of-bounds geometry is clamped into the ${CANVAS.width}x${CANVAS.height} canvas, SVG is sanitized for the vector rasterizer, and overflow is reported in "warnings" — fix them before exporting.`,
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        deck: {
          type: 'object',
          additionalProperties: true,
          description:
            'Complete CardDeck object: { deckId, title, templateId, slides: [{ slideId, background?, speakerNotes?, elements: [...] }] }.',
        },
      },
      required: ['deck'],
    },
    async (args) => {
      const candidate = args.deck
      if (candidate === null || typeof candidate !== 'object') {
        return { value: fail('invalid_deck', 'deck must be an object') }
      }
      const theme = await themeFor(candidate, config, storeConfig)
      const repaired = parseAndRepairDeck(candidate, theme, budgetOptions(config))
      if (!repaired.ok) {
        return {
          value: fail('invalid_deck', 'The deck does not match the CardDeck schema', repaired.issues),
        }
      }
      const normalized = {
        ...repaired.deck,
        deckId: repaired.deck.deckId || slugify(repaired.deck.title ?? 'deck'),
        templateId: repaired.deck.templateId || theme.templateId,
      }
      await store.writeDeck(normalized, storeConfig)
      return {
        value: {
          ok: true,
          deckId: normalized.deckId,
          slideCount: normalized.slides.length,
          warnings: repaired.warnings,
          theme,
        },
      }
    },
  )

  register(
    TOOL_NAMES.deckGet,
    'Read one stored deck back, including the repair warnings a save would produce.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        deckId: { type: 'string', description: 'Deck id; omit for the most recently saved deck.' },
      },
    },
    async (args) => {
      const deck =
        args.deckId === undefined
          ? await activeDeck(storeConfig)
          : readDeckOrNull(args.deckId, storeConfig)
      if (deck === null) return { value: fail('deck_not_found', 'No deck has been saved yet') }
      const theme = await themeFor(deck, config, storeConfig)
      const repaired = repairDeck(deck, theme, budgetOptions(config))
      return { value: { ok: true, deck: repaired.deck, theme, warnings: repaired.warnings } }
    },
  )

  register(
    TOOL_NAMES.deckList,
    'List stored decks (newest first) with slide counts, so you can pick one to edit or export.',
    { type: 'object', additionalProperties: false, properties: {} },
    async () => ({ value: { ok: true, decks: await store.listDecks(storeConfig) } }),
  )

  register(
    TOOL_NAMES.deckExport,
    'Compile a deck into a native, fully editable .pptx (13.333x7.5in wide layout): text becomes real text boxes, shapes become native shapes, charts become Office charts, and SVG figures are rasterized at 3x. Returns the written file path.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        deckId: { type: 'string', description: 'Deck id; omit for the most recently saved deck.' },
        outputPath: {
          type: 'string',
          description:
            'Where to write the .pptx. Relative paths resolve against the session working directory. Defaults to <cwd>/branddeck/<deckId>.pptx.',
        },
      },
    },
    async (args, exec) => {
      const deck =
        args.deckId === undefined
          ? await activeDeck(storeConfig)
          : readDeckOrNull(args.deckId, storeConfig)
      if (deck === null) return { value: fail('deck_not_found', 'No deck has been saved yet') }
      const theme = await themeFor(deck, config, storeConfig)
      const repaired = parseAndRepairDeck(deck, theme, budgetOptions(config))
      if (!repaired.ok) {
        return {
          value: fail('invalid_deck', 'The stored deck no longer matches the schema', repaired.issues),
        }
      }
      const cwd = workspaceOf(exec)
      const outputPath =
        typeof args.outputPath === 'string' && args.outputPath.trim() !== ''
          ? resolvePath(cwd, args.outputPath)
          : resolvePath(cwd, 'branddeck', `${deck.deckId}.pptx`)
      const result = await compileDeckToFile(
        {
          deck: repaired.deck,
          theme,
          logoBytes: await readLogoBytes(theme),
          assets: await resolveIconAssets(theme),
          options: compileOptions(config),
        },
        outputPath,
      )
      return {
        value: {
          ok: true,
          path: outputPath,
          slideCount: result.slideCount,
          warnings: [...repaired.warnings, ...(result.warnings ?? [])],
        },
      }
    },
  )

  register(
    TOOL_NAMES.outlineGet,
    'Read the editable TEXT version of a deck (Markdown outline: slide headings, text elements as `### 角色 {#id}`, speaker notes). This is the cheap surface to revise copy on; geometry lives in the card JSON and is preserved when you save it back.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        deckId: { type: 'string', description: 'Deck id; omit for the most recently saved deck.' },
      },
    },
    async (args) => {
      const deck =
        args.deckId === undefined
          ? await activeDeck(storeConfig)
          : readDeckOrNull(args.deckId, storeConfig)
      if (deck === null) return { value: fail('deck_not_found', 'No deck has been saved yet') }
      const theme = await themeFor(deck, config, storeConfig)
      const outline = store.readDeckOutline(deck.deckId, storeConfig, theme)
      return {
        value: {
          ok: true,
          deckId: deck.deckId,
          revision: outline.revision,
          updatedAt: outline.updatedAt,
          markdown: outline.markdown,
        },
      }
    },
  )

  register(
    TOOL_NAMES.outlineSave,
    'Apply an edited TEXT version of a deck back onto the card version. Text elements matched by {#id} keep their exact position and style; new `### 角色 {#new-id}` blocks are auto-laid out in the safe area; graphics/charts/SVG are untouched. Returns per-element changes and warnings.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        deckId: { type: 'string', description: 'Deck id to update.' },
        markdown: { type: 'string', description: 'The complete outline Markdown (as returned by branddeck_outline_get, edited).' },
        prune: {
          type: 'boolean',
          description: 'Remove text elements missing from the Markdown (default false: they are kept).',
        },
      },
      required: ['deckId', 'markdown'],
    },
    async (args) => {
      const existing = readDeckOrNull(args.deckId, storeConfig)
      if (existing === null) return { value: fail('deck_not_found', `deck "${args.deckId}" not found`) }
      const theme = await themeFor(existing, config, storeConfig)
      const applied = store.writeDeckOutline(args.deckId, args.markdown, storeConfig, {
        theme,
        prune: args.prune === true,
      })
      return {
        value: {
          ok: true,
          deckId: applied.deck.deckId,
          revision: applied.revision,
          changes: applied.changes,
          warnings: applied.warnings,
        },
      }
    },
  )

  register(
    TOOL_NAMES.deckMaterialize,
    'Write the deck bundle (card.json + outline.md + meta.json) into a visible directory — e.g. the workspace — so the user can review, diff and commit both the card and text versions with git.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        deckId: { type: 'string', description: 'Deck id; omit for the most recently saved deck.' },
        dir: {
          type: 'string',
          description: 'Target directory. Relative paths resolve against the session working directory. Defaults to <cwd>/branddeck/<deckId>.',
        },
      },
    },
    async (args, exec) => {
      const deck =
        args.deckId === undefined
          ? await activeDeck(storeConfig)
          : readDeckOrNull(args.deckId, storeConfig)
      if (deck === null) return { value: fail('deck_not_found', 'No deck has been saved yet') }
      const theme = await themeFor(deck, config, storeConfig)
      const cwd = workspaceOf(exec)
      const dir =
        typeof args.dir === 'string' && args.dir.trim() !== ''
          ? resolvePath(cwd, args.dir)
          : resolvePath(cwd, 'branddeck', deck.deckId)
      const result = store.materializeDeck(deck.deckId, dir, storeConfig, { theme })
      return { value: { ok: true, dir: result.dir, files: result.files, deckId: deck.deckId } }
    },
  )

  register(
    TOOL_NAMES.styleList,
    'List the reusable style profiles (「风格配置文件」): density, tone, bullet layout, colour/font overrides on a template. Styles are what the panel and branddeck_generate offer as "make it look like this".',
    { type: 'object', additionalProperties: false, properties: {} },
    async () => ({ value: { ok: true, styles: styles.listStyles(storeConfig) } }),
  )

  register(
    TOOL_NAMES.styleSave,
    'Create or update a style profile from a user description — this is how a conversation becomes a reusable 「风格配置文件」. Describe the intent in `description` (e.g.「华为红、信息密度高、卡片式正文」) and set density/tone/bulletLayout/colours; the result is stored and instantly selectable in the panel.',
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        styleId: { type: 'string', description: 'Existing style id to update; omit to create one (derived from name).' },
        name: { type: 'string', description: 'Human-readable style name.' },
        description: { type: 'string', description: 'The user intent this style encodes.' },
        templateId: { type: 'string', description: 'Template the style is tuned for.' },
        density: { type: 'string', enum: ['airy', 'standard', 'dense'], description: 'Information density (bullets per slide).' },
        tone: { type: 'string', enum: ['formal', 'bold', 'minimal'], description: 'Accent usage.' },
        bulletLayout: { type: 'string', enum: ['list', 'cards'], description: 'How bullet slides are drawn.' },
        coverVariant: { type: 'string', enum: ['left', 'centered'] },
        accentBar: { type: 'boolean', description: 'Draw the brand accent bar.' },
        closingSlide: { type: 'boolean', description: 'Append a closing slide.' },
        maxSlides: { type: 'number', description: 'Hard page cap including cover and closing.' },
        colors: {
          type: 'object',
          additionalProperties: true,
          description: 'Token overrides on the template palette: background/surface/textPrimary/textSecondary/primary/accent.',
        },
        fonts: {
          type: 'object',
          additionalProperties: true,
          description: 'Font overrides: title/body.',
        },
        source: { type: 'string', enum: ['manual', 'chat', 'template'], description: 'Where this style came from.' },
      },
      required: ['name'],
    },
    async (args) => {
      const payload = { ...args }
      delete payload.styleId
      const saved = styles.writeStyle(payload, storeConfig, { styleId: args.styleId })
      return { value: { ok: true, style: saved, styles: styles.listStyles(storeConfig) } }
    },
  )

  register(
    TOOL_NAMES.generate,
    `Generate a deck from user input + a chosen style (no LLM): parses pasted text or a Markdown/plain-text file into an outline, lays it out on the 1280x720 canvas, then stores it so the panel can refine it. This is the "describe it and get slides" entry point.`,
    {
      type: 'object',
      additionalProperties: false,
      properties: {
        text: { type: 'string', description: 'Pasted content or a description of the deck to build.' },
        path: { type: 'string', description: 'A .md/.txt file to read instead of `text` (workspace-relative or absolute).' },
        templateId: { type: 'string', description: 'Template for brand tokens; omit for the default/most recent template.' },
        styleId: { type: 'string', description: 'Style profile id; omit for the neutral default style.' },
        deckId: { type: 'string', description: 'Deck id to create/overwrite.' },
        title: { type: 'string', description: 'Override the deck title (defaults to the first heading).' },
        maxSlides: { type: 'number', description: 'Page cap override for this generation.' },
        outputPath: {
          type: 'string',
          description: 'When given, also compile the .pptx here (relative paths resolve against the session cwd).',
        },
      },
    },
    async (args, exec) => {
      const generated = await generateDeckFrom({
        text: args.text,
        path: args.path,
        templateId: args.templateId,
        styleId: args.styleId,
        deckId: args.deckId,
        title: args.title,
        maxSlides: args.maxSlides,
        config,
        storeConfig,
        cwd: workspaceOf(exec),
      })
      if (generated.ok !== true) return { value: generated }
      if (typeof args.outputPath === 'string' && args.outputPath.trim() !== '') {
        const outputPath = resolvePath(workspaceOf(exec), args.outputPath)
        const compiled = await compileDeckToFile(
          {
            deck: generated.deck,
            theme: generated.theme,
            logoBytes: await readLogoBytes(generated.theme),
            assets: await resolveIconAssets(generated.theme),
            options: compileOptions(config),
          },
          outputPath,
        )
        return {
          value: { ...generated, path: outputPath, slideCount: compiled.slideCount, warnings: [...generated.warnings, ...(compiled.warnings ?? [])] },
        }
      }
      return { value: generated }
    },
  )
}

async function firstTemplateId(storeConfig) {
  const templates = await store.listTemplates(storeConfig)
  return templates[0]?.templateId
}

/**
 * Load the conversation helpers lazily, for the same reason as the generator: a
 * missing module must degrade one endpoint, not the plugin's activation.
 */
async function loadAssist() {
  try {
    return await import('./src/assist.js')
  } catch (error) {
    const failure = fail(
      'assist_unavailable',
      `conversation helpers unavailable: ${error instanceof Error ? error.message : String(error)}`,
    )
    const wrapper = new Error(failure.error.message)
    wrapper.code = failure.error.code
    throw wrapper
  }
}

/**
 * "Say a sentence, get a style": ask the model to translate the description into a
 * style profile, validate it against the same schema the tool uses, and return a
 * *draft* (the panel shows it for review before saving).
 *
 * @param {{description: string, theme: any}} input
 * @param {{ctx: any, storeConfig: {root?: string}}} deps
 */
async function assistStyle(input, deps) {
  const assist = await loadAssist()
  const prompt = assist.buildStylePrompt({
    description: input.description,
    theme: input.theme,
    styleIds: styles.listStyles(deps.storeConfig).map((entry) => entry.styleId),
  })
  const answer = await assist.callModel(deps.ctx, { ...prompt, maxTokens: 900 })
  const parsed = assist.extractJsonObject(answer)
  if (parsed === null) {
    return {
      ok: false,
      error: { code: 'style_unparsable', message: '模型没有返回可解析的风格 JSON' },
      raw: String(answer).slice(0, 1200),
    }
  }
  // Validate through the same normaliser the manual path uses, so a conversational
  // style cannot smuggle in fields the schema rejects.
  let style
  try {
    style = styles.normalizeStyle(
      { ...parsed, styleId: undefined, templateId: input.theme?.templateId, source: 'chat' },
      {},
    )
  } catch (error) {
    return {
      ok: false,
      error: { code: 'style_invalid', message: error instanceof Error ? error.message : String(error) },
      raw: String(answer).slice(0, 1200),
    }
  }
  style.description = typeof parsed.description === 'string' && parsed.description.trim() !== ''
    ? parsed.description.trim()
    : input.description
  return { ok: true, style, raw: String(answer).slice(0, 1200) }
}

/**
 * "Say a sentence, get slides": the model writes a Markdown outline in our own
 * convention, and the deterministic generator lays it out. Layout stays code-owned,
 * so a bad answer can never produce an unrenderable deck.
 *
 * @param {{description: string, templateId: string | null, styleId?: string, maxSlides?: number}} input
 * @param {{ctx: any, config: any, storeConfig: {root?: string}}} deps
 */
async function assistContent(input, deps) {
  const assist = await loadAssist()
  const templateId = input.templateId ?? (await firstTemplateId(deps.storeConfig))
  const baseTheme = templateId === undefined ? DEFAULT_THEME : await themeFor({ templateId }, deps.config, deps.storeConfig)
  const style = input.styleId === undefined || input.styleId === '' ? null : styles.readStyleOrNull(input.styleId, deps.storeConfig)
  const effective = style ?? styles.DEFAULT_STYLE
  const theme = styles.applyStyleToTheme(baseTheme, effective)
  theme.templateId = baseTheme.templateId
  theme.templateName = baseTheme.templateName

  // One pipeline, two gears: a sentence goes through a single prompt, a long document
  // through Map-Reduce (per-block facts, then one structuring pass). The threshold lives
  // in the plugin config so a deployment can tune it without a code change.
  const material = typeof input.material === 'string' && input.material.trim() !== '' ? input.material : input.description
  let markdown
  let pipeline = { blocks: 1, calls: 1, dropped: 0, mode: 'single' }
  try {
    const outlined = await assist.assistOutlineFromMaterial(deps.ctx, {
      text: material,
      maxSlides: input.maxSlides ?? effective.maxSlides,
      topic: input.title,
      options: {
        threshold: input.materialThreshold ?? deps.config.materialThreshold,
        blockChars: deps.config.materialBlockChars,
        maxBlocks: deps.config.materialMaxBlocks,
        concurrency: deps.config.materialConcurrency,
        factsPerBlock: deps.config.materialFactsPerBlock,
      },
      onProgress: deps.onProgress,
      signal: deps.signal,
    })
    markdown = outlined.markdown
    pipeline = {
      blocks: outlined.blocks,
      calls: outlined.calls,
      dropped: outlined.dropped,
      mode: outlined.blocks > 1 ? 'map-reduce' : 'single',
    }
  } catch (error) {
    const code = typeof error?.code === 'string' ? error.code : 'llm_failed'
    return { ok: false, error: { code, message: error instanceof Error ? error.message : String(error) } }
  }
  if (typeof markdown !== 'string' || markdown.trim() === '') {
    return { ok: false, error: { code: 'outline_unparsable', message: '模型没有返回可用的大纲' } }
  }

  const generated = await generateDeckFrom({
    text: markdown,
    templateId: templateId ?? undefined,
    styleId: effective.styleId,
    maxSlides: input.maxSlides ?? effective.maxSlides,
    config: deps.config,
    storeConfig: deps.storeConfig,
    slideCharBudget: input.slideCharBudget,
    bulletCharBudget: input.bulletCharBudget,
  })
  if (generated.ok !== true) return generated
  return { ...generated, markdown, pipeline }
}

/**
 * One generation pipeline shared by the tool and the panel endpoint.
 *
 * `src/generate.js` is imported lazily: it is the only module whose absence used
 * to be able to break plugin activation, and generation is a feature — not a
 * reason for the whole plugin to fail to load.
 *
 * @param {{text?: string, path?: string, templateId?: string, styleId?: string, deckId?: string,
 *          title?: string, maxSlides?: number, config: any, storeConfig: {root?: string}, cwd?: string}} input
 */
async function generateDeckFrom(input) {
  const { config, storeConfig } = input
  const cwd = input.cwd ?? process.cwd()

  let source = typeof input.text === 'string' ? input.text : ''
  if (source.trim() === '' && typeof input.path === 'string' && input.path.trim() !== '') {
    const file = resolvePath(cwd, input.path)
    try {
      source = await readFile(file, 'utf8')
    } catch {
      return fail('source_not_found', `cannot read source file ${file}`)
    }
  }
  if (source.trim() === '') {
    return fail('invalid_body', 'Expected `text` or a readable `path` as the content source')
  }

  const templateId = input.templateId ?? config.defaultTemplateId ?? (await firstTemplateId(storeConfig))
  const baseTheme = templateId === undefined ? DEFAULT_THEME : await themeFor({ templateId }, config, storeConfig)
  const style = input.styleId === undefined || input.styleId === '' ? null : styles.readStyleOrNull(input.styleId, storeConfig)
  if (input.styleId !== undefined && input.styleId !== '' && style === null && input.styleId !== 'default') {
    return fail('style_not_found', `style "${input.styleId}" not found`)
  }
  const effective = style ?? styles.DEFAULT_STYLE
  const theme = styles.applyStyleToTheme(baseTheme, effective)
  // The template id must follow the theme (a style may be applied over another template).
  theme.templateId = baseTheme.templateId
  theme.templateName = baseTheme.templateName

  let generator
  try {
    generator = await import('./src/generate.js')
  } catch (error) {
    return fail('generator_unavailable', `deck generator unavailable: ${error instanceof Error ? error.message : String(error)}`)
  }

  const outline = generator.parseSourceToOutline(source)
  if (typeof input.title === 'string' && input.title.trim() !== '') outline.title = input.title.trim()
  if (outline.slides.length === 0) {
    return fail('empty_source', 'The source produced no slides — give it headings or at least a few paragraphs')
  }

  const budgets = budgetOptions(config, input)
  const composed = generator.composeDeck({
    outline,
    theme,
    style: effective,
    deckId: input.deckId,
    maxSlides: Number.isFinite(input.maxSlides) ? input.maxSlides : effective.maxSlides,
    slideCharBudget: budgets.slideCharBudget,
    bulletCharBudget: budgets.bulletCharBudget,
  })

  const repaired = parseAndRepairDeck(composed.deck, theme, budgets)
  if (!repaired.ok) {
    return fail('invalid_deck', 'Generated deck failed validation', repaired.issues)
  }
  await store.writeDeck(repaired.deck, storeConfig, { theme })
  return {
    ok: true,
    deckId: repaired.deck.deckId,
    deck: repaired.deck,
    theme,
    styleId: effective.styleId,
    templateId: theme.templateId,
    slideCount: repaired.deck.slides.length,
    outline: { title: outline.title, slides: outline.slides.length },
    warnings: [...(composed.warnings ?? []), ...repaired.warnings],
  }
}

/**
 * Read a template's formatted spec, back-filling `layouts.json` / `template.md`
 * for templates parsed by an older version.
 *
 * This is the "next time you can use it directly" guarantee: the spec is derived
 * once, stored next to `theme.json`, and only re-derived when it is missing and a
 * `source.pptx` copy exists.
 *
 * @param {string} templateId
 * @param {{root?: string}} storeConfig
 * @returns {Promise<{theme: any, layouts: any, markdown: string | null} | null>}
 */
async function ensureTemplateDoc(templateId, storeConfig) {
  let template
  try {
    template = await store.readTemplate(templateId, storeConfig)
  } catch {
    return null
  }
  if (template.markdown !== null && template.layouts !== null) {
    return { theme: template.theme, layouts: template.layouts, markdown: template.markdown }
  }
  const source = template.sourcePath ?? template.meta?.sourcePath ?? null
  if (source === null || !existsSync(source)) {
    return { theme: template.theme, layouts: template.layouts, markdown: template.markdown }
  }
  try {
    const parsed = await parsePptxTemplate({
      path: source,
      templateId,
      templateName: template.theme.templateName,
      storeConfig,
    })
    await store.writeTemplate(
      {
        templateId,
        templateName: parsed.theme.templateName,
        tokens: parsed.theme,
        logoBytes: parsed.logoBytes,
        meta: parsed.meta,
        media: parsed.media,
        sourceBytes: await readFile(source),
        sourcePath: source,
        layoutsInventory: parsed.layouts ?? null,
        markdown: parsed.markdown ?? null,
      },
      storeConfig,
    )
    const refreshed = await store.readTemplate(templateId, storeConfig)
    return { theme: refreshed.theme, layouts: refreshed.layouts, markdown: refreshed.markdown }
  } catch {
    return { theme: template.theme, layouts: template.layouts, markdown: template.markdown }
  }
}

/** Logo bytes for a theme, or null when the template had no master logo. */
async function readLogoBytes(theme) {
  const logoPath = theme?.brandAssets?.logoPath
  if (typeof logoPath !== 'string' || logoPath === '') return null
  try {
    return await readFile(logoPath)
  } catch {
    return null
  }
}

/** Resolve `brandAssets.icons` (semantic name → path) into bytes for the compiler. */
async function resolveIconAssets(theme) {
  const icons = theme?.brandAssets?.icons ?? {}
  /** @type {Record<string, Buffer>} */
  const assets = {}
  for (const [key, value] of Object.entries(icons)) {
    if (typeof value !== 'string' || value === '') continue
    try {
      assets[key] = await readFile(value)
    } catch {
      /* missing icons degrade to a compiler warning, not a hard failure */
    }
  }
  return assets
}

function slugify(input) {
  const ascii = String(input ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, '-')
    .replace(/^-+|-+$/g, '')
  const normalized = ascii.replace(/[^a-z0-9-]+/g, '')
  return (normalized !== '' ? normalized : `deck-${Date.now().toString(36)}`).slice(0, 48)
}

/* ------------------------------------------------------------------ *
 * System prompt section
 * ------------------------------------------------------------------ */

function registerPromptSection(ctx, config) {
  ctx.inject(['systemPrompt'], (scope) => {
    scope.effect(
      () =>
        scope.systemPrompt.section({
          name: 'branddeck',
          order: PROMPT_ORDER,
          text: () => PROMPT_TEXT,
        }),
      'branddeck: prompt section',
    )
  })
}

const PROMPT_TEXT = `## BrandDeck（品牌演示卡组）

需要制作 PPT/演示文稿时，用 BrandDeck 工具把内容做成"卡片 DSL"再导出为原生可编辑的 .pptx。
不要手写 pptxgenjs 脚本，也不要直接改 .pptx 文件。

工作流：
1. 模板：${TOOL_NAMES.templateParse} { path } —— 从企业 .pptx 模板提取品牌色板、中英文字体与母版 Logo，
   并一次性格式化为 theme.json（令牌）+ layouts.json（版式）+ template.md（规格书）+ source.pptx（原件副本）。
   （解析过就跳过；用 ${TOOL_NAMES.templateList} 查看，用 ${TOOL_NAMES.templateSpec} 读规格书。）
2. 对齐品牌：${TOOL_NAMES.themeGet} { templateId } —— 拿到 colors/fonts，写元素时用 "$colors.primary" 这类令牌，不要硬编码颜色。
3. 建卡组：${TOOL_NAMES.deckCreate} { title, templateId }。
4. 写内容：${TOOL_NAMES.deckSave} { deck } —— 一次提交整份 deck；返回的 warnings 必须处理（溢出/越界/图表长度不匹配）。
5. 改文案优先用文字版，比重发整份 JSON 便宜且不动版式：
   ${TOOL_NAMES.outlineGet} { deckId } 读 Markdown 大纲 → 改文字 → ${TOOL_NAMES.outlineSave} { deckId, markdown } 回写。
   元素用 {#id} 锚定，位置与样式保持不变；新增一条「### 正文 {#新id}」会自动排版进安全区。
6. 需要让用户 review / 进 git 时：${TOOL_NAMES.deckMaterialize} { deckId, dir } —— 落出 card.json + outline.md + meta.json。
7. 导出：${TOOL_NAMES.deckExport} { deckId, outputPath }。

卡组 DSL（虚拟画布 ${CANVAS.width}×${CANVAS.height}，px）：
- 安全区：上下左右各 ${SAFE_AREA.top}px；右上角预留 Logo 区，避免把标题放进去。
- 元素：{ type:'text', id, role:'title'|'subtitle'|'body'|'caption', content, x,y,w,h, style:{fontSize,bold?,color,align?} }
  { type:'shape', shapeType:'rect'|'roundRect'|'ellipse'|'line'|'rightArrow', x,y,w,h, style:{fill?,stroke?,strokeWidth?,radius?} }
  { type:'svg', description, x,y,w,h, svgCode }
  { type:'chart', chartType:'bar'|'line'|'pie', x,y,w,h, data:{labels[],series:[{name,values[]}]} }
  { type:'image', assetKey, x,y,w,h }
- 单页正文总量 ≤150 字（超出会警告），一页一个核心观点。
- svg 只能使用纯矢量标签（rect/circle/ellipse/line/path/polygon/polyline/text/g/defs），
  禁止 foreignObject/script/image；viewBox 用元素自身 w/h，颜色显式写成品牌 hex。
- 流程图/漏斗/对比矩阵优先用 svg 或 shape 组合，不要塞进一大段正文。
- 用户可在侧栏 BrandDeck 面板里可视化微调并导出，所以 deck 要保存到插件存储（deckSave），不要只贴在回复里。`

/* ------------------------------------------------------------------ *
 * Panel API (Web Client ↔ Host)
 * ------------------------------------------------------------------ */

/** Same-origin HTTP prefix the Web panel talks to. */
export const API_PREFIX = '/branddeck-api'

/**
 * Register the panel transport.
 *
 * The Web Client half is a plain-JS bundle plugin and loads as a *sibling* of
 * the shipped client plugins, so it cannot see services those plugins provide
 * on their own contexts (`connection` included — it is installed with
 * `ctx.provide`, visible only to that plugin's children). Injecting it would
 * leave our client entry pending ("loading") and fail the whole web boot.
 *
 * The transport is therefore a same-origin HTTP route on the Host web server
 * (the same carrier the client-module registry uses for `/plugins`), called
 * from the browser with `fetch(API_PREFIX + '/' + endpoint, { method: 'POST' })`.
 * The page's own `dsh-auth` cookie is irrelevant here and no credential is
 * invented: the route only serves JSON to a same-origin caller.
 */
function registerPanelApi(ctx, config, storeConfig) {
  ctx.inject(['webServer'], (scope) => {
    scope.effect(
      () =>
        scope.webServer.register({
          kind: 'prefix',
          path: API_PREFIX,
          handler: (req, res) => {
            handleApiRequest(req, res, { config, storeConfig, ctx }).catch((error) => {
              const envelope = failureOf(error)
              if (!res.headersSent) writeJson(res, 500, envelope)
              else res.destroy()
            })
          },
        }),
      'branddeck: panel http route',
    )
  })
}

/**
 * One HTTP call: `POST <API_PREFIX>/<endpoint>` with the JSON payload as body.
 * Answers the same `{ ok, value } | { ok, error }` envelope the endpoint table
 * returns, so the panel has one shape to handle.
 * @param {import('node:http').IncomingMessage} req
 * @param {import('node:http').ServerResponse} res
 */
async function handleApiRequest(req, res, deps) {
  const site = String(req.headers['sec-fetch-site'] ?? '')
  if (site === 'cross-site') {
    writeJson(res, 403, fail('forbidden', 'cross-site requests are not allowed'))
    return
  }
  const method = (req.method ?? 'GET').toUpperCase()
  if (method !== 'POST' && method !== 'GET') {
    writeJson(res, 405, fail('method_not_allowed', 'Use POST'))
    return
  }
  const url = new URL(req.url ?? '/', 'http://branddeck.invalid')
  const endpoint = decodeURIComponent(url.pathname.slice(API_PREFIX.length).replace(/^\/+/, ''))
  if (endpoint === '') {
    writeJson(res, 400, fail('invalid_request', `Expected ${API_PREFIX}/<endpoint>`))
    return
  }
  const payload = method === 'GET' ? {} : ((await readJsonBody(req)) ?? {})
  const envelope = await dispatchEndpoint(endpoint, payload, deps).catch((error) => failureOf(error))
  writeJson(res, envelope.ok === true ? 200 : statusForError(envelope.error?.code), envelope)
}

/** Map a stable error code onto an HTTP status (the body carries the real code). */
function statusForError(code) {
  switch (code) {
    case 'not_found':
    case 'deck_not_found':
    case 'template_not_found':
    case 'logo_not_found':
      return 404
    case 'invalid_body':
    case 'invalid_deck':
    case 'invalid_request':
      return 422
    default:
      return 500
  }
}

/** Read and parse a JSON request body with a small size bound. */
async function readJsonBody(req, limitBytes = 8 * 1024 * 1024) {
  const chunks = []
  let total = 0
  for await (const chunk of req) {
    total += chunk.length
    if (total > limitBytes) {
      const error = new Error('Request body is too large')
      error.code = 'body_too_large'
      throw error
    }
    chunks.push(chunk)
  }
  if (total === 0) return undefined
  try {
    return JSON.parse(Buffer.concat(chunks).toString('utf8'))
  } catch {
    const error = new Error('Request body is not valid JSON')
    error.code = 'invalid_json'
    throw error
  }
}

function writeJson(res, status, payload) {
  const body = Buffer.from(JSON.stringify(payload), 'utf8')
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'content-length': String(body.byteLength),
    'cache-control': 'no-store',
  })
  res.end(body)
}

/**
 * One endpoint table for the panel. Keeping it a plain switch means the payload
 * shapes stay next to the store/compiler calls that consume them.
 * @param {string} endpoint
 * @param {any} payload
 * @param {{ config: ReturnType<typeof normalizeConfig>, storeConfig: { root?: string } }} deps
 */
async function dispatchEndpoint(endpoint, payload, deps) {
  const { config, storeConfig } = deps
  const ctx = deps.ctx ?? null
  const body = payload ?? {}

  switch (endpoint) {
    case 'state': {
      const decks = await store.listDecks(storeConfig)
      const templates = await store.listTemplates(storeConfig)
      const deck = await activeDeck(storeConfig)
      const theme = deck === null ? DEFAULT_THEME : await themeFor(deck, config, storeConfig)
      const repaired = deck === null ? null : repairDeck(deck, theme, budgetOptions(config))
      return ok({
        decks,
        templates,
        deck: repaired?.deck ?? null,
        theme,
        warnings: repaired?.warnings ?? [],
        activeDeckId: deck?.deckId ?? null,
        activeTemplateId: theme.templateId,
        capability: {
          canvas: CANVAS,
          safeArea: SAFE_AREA,
          dataDir: store.storeRoot(storeConfig),
        },
      })
    }

    case 'decks.list': {
      return ok({ decks: await store.listDecks(storeConfig) })
    }

    case 'decks.get': {
      const decks = await store.listDecks(storeConfig)
      const deckId = typeof body.deckId === 'string' && body.deckId !== '' ? body.deckId : null
      const deck = deckId === null ? await activeDeck(storeConfig) : readDeckOrNull(deckId, storeConfig)
      if (deck === null) return fail('deck_not_found', 'No deck has been saved yet')
      const theme = await themeFor(deck, config, storeConfig)
      const repaired = repairDeck(deck, theme, budgetOptions(config))
      const outline = store.readDeckOutline(deck.deckId, storeConfig, theme)
      return ok({
        deck: repaired.deck,
        theme,
        warnings: repaired.warnings,
        decks,
        outline: outline.markdown,
        revision: outline.revision,
        updatedAt: outline.updatedAt,
      })
    }

    case 'outline.get': {
      const deckId = typeof body.deckId === 'string' && body.deckId !== '' ? body.deckId : null
      const deck = deckId === null ? await activeDeck(storeConfig) : readDeckOrNull(deckId, storeConfig)
      if (deck === null) return fail('deck_not_found', 'No deck has been saved yet')
      const theme = await themeFor(deck, config, storeConfig)
      const outline = store.readDeckOutline(deck.deckId, storeConfig, theme)
      return ok({
        deckId: deck.deckId,
        markdown: outline.markdown,
        revision: outline.revision,
        updatedAt: outline.updatedAt,
      })
    }

    case 'outline.save': {
      const deckId = typeof body.deckId === 'string' && body.deckId !== '' ? body.deckId : null
      if (deckId === null) return fail('invalid_body', 'Expected { deckId, markdown }')
      if (typeof body.markdown !== 'string') return fail('invalid_body', 'Expected { deckId, markdown }')
      const existing = readDeckOrNull(deckId, storeConfig)
      if (existing === null) return fail('deck_not_found', `deck "${deckId}" not found`)
      const theme = await themeFor(existing, config, storeConfig)
      const applied = store.writeDeckOutline(deckId, body.markdown, storeConfig, {
        theme,
        prune: body.prune === true,
      })
      const outline = store.readDeckOutline(deckId, storeConfig, theme)
      return ok({
        deck: applied.deck,
        theme,
        markdown: outline.markdown,
        revision: applied.revision,
        changes: applied.changes,
        warnings: applied.warnings,
      })
    }

    case 'decks.materialize': {
      const deckId = typeof body.deckId === 'string' && body.deckId !== '' ? body.deckId : null
      const deck = deckId === null ? await activeDeck(storeConfig) : readDeckOrNull(deckId, storeConfig)
      if (deck === null) return fail('deck_not_found', 'No deck has been saved yet')
      const theme = await themeFor(deck, config, storeConfig)
      const target =
        typeof body.dir === 'string' && body.dir.trim() !== ''
          ? resolvePath(body.dir)
          : join(store.storeRoot(storeConfig), 'decks', deck.deckId)
      const result = store.materializeDeck(deck.deckId, target, storeConfig, { theme })
      return ok({ dir: result.dir, files: result.files, deckId: deck.deckId })
    }

    case 'templates.spec': {      const templates = await store.listTemplates(storeConfig)
      const requested = typeof body.templateId === 'string' && body.templateId !== '' ? body.templateId : null
      const templateId = requested ?? config.defaultTemplateId ?? templates[0]?.templateId ?? null
      if (templateId === null) return fail('template_not_found', 'No template has been parsed yet')
      const resolved = await ensureTemplateDoc(templateId, storeConfig)
      if (resolved === null) return fail('template_not_found', `template "${templateId}" not found`)
      return ok({
        templateId,
        templateName: resolved.theme.templateName,
        markdown: resolved.markdown,
        layouts: resolved.layouts,
      })
    }

    case 'styles.list': {
      return ok({ styles: styles.listStyles(storeConfig), defaultStyleId: styles.DEFAULT_STYLE.styleId })
    }

    case 'styles.save': {
      const payload = body.style !== null && typeof body.style === 'object' ? { ...body.style } : { ...body }
      delete payload.styleId
      const styleId = typeof body.styleId === 'string' && body.styleId !== '' ? body.styleId : undefined
      const saved = styles.writeStyle(payload, storeConfig, { styleId })
      return ok({ style: saved, styles: styles.listStyles(storeConfig) })
    }

    case 'styles.delete': {
      const styleId = typeof body.styleId === 'string' ? body.styleId : ''
      if (styleId === '') return fail('invalid_body', 'Expected { styleId }')
      const removed = styles.deleteStyle(styleId, storeConfig)
      if (!removed) return fail('style_not_found', `style "${styleId}" not found`)
      return ok({ removed: styleId, styles: styles.listStyles(storeConfig) })
    }

    case 'generate': {
      const generated = await generateDeckFrom({
        text: typeof body.text === 'string' ? body.text : undefined,
        path: typeof body.path === 'string' ? body.path : undefined,
        templateId: typeof body.templateId === 'string' ? body.templateId : undefined,
        styleId: typeof body.styleId === 'string' ? body.styleId : undefined,
        deckId: typeof body.deckId === 'string' ? body.deckId : undefined,
        title: typeof body.title === 'string' ? body.title : undefined,
        maxSlides: typeof body.maxSlides === 'number' ? body.maxSlides : undefined,
        config,
        storeConfig,
        cwd: typeof body.cwd === 'string' ? body.cwd : process.cwd(),
      })
      return generated.ok === true ? ok(generated) : generated
    }

    case 'command.run': {
      // Programmatic equivalent of typing the line in the composer: same parser,
      // same handler, same transcript result. A caller may name a session; without
      // one the most recently registered live agent is used.
      const line = typeof body.line === 'string' ? body.line : ''
      if (line.trim() === '') return fail('invalid_body', 'Expected { line }')
      const agents = ctx?.get?.('agents') ?? null
      let live = []
      try {
        live = typeof agents?.list === 'function' ? agents.list() : []
      } catch {
        live = []
      }
      const wanted = typeof body.sessionId === 'string' ? body.sessionId : null
      const agent = (() => {
        if (wanted !== null) {
          try {
            return agents?.get?.(wanted) ?? null
          } catch {
            return null
          }
        }
        return live.length === 0 ? null : live[live.length - 1]
      })()
      if (agent === null || agent === undefined) {
        return fail('no_agent', 'No live session is available to run the command in')
      }
      const rawInput = line.replace(/^\s*\/?branddeck\s*/i, '')
      const answer = await executeBrandDeckCommand(
        { agent, rawInput, attachments: [] },
        { ctx, config, storeConfig },
      ).catch((error) => ({
        kind: 'error',
        text: error instanceof Error ? error.message : String(error),
      }))
      return ok({ line, rawInput, kind: answer.kind, text: answer.text })
    }

    case 'command.check': {
      // Deterministic proof that /branddeck is registered for live agents:
      // resolve the command registry and ask it for this very command.
      const commands = ctx?.get?.('commands') ?? null
      const agents = ctx?.get?.('agents') ?? null
      const live = (() => {
        try {
          return typeof agents?.list === 'function' ? agents.list() : []
        } catch {
          return []
        }
      })()
      const rows = live.slice(0, 5).map((agent) => {
        const sessionId = (() => {
          try {
            return agent?.session?.id ?? null
          } catch {
            return null
          }
        })()
        let found = null
        try {
          const definition = typeof commands?.find === 'function' ? commands.find(agent, 'branddeck') : undefined
          found = definition === undefined ? null : { name: definition.name, description: definition.description }
        } catch (error) {
          found = { error: error instanceof Error ? error.message : String(error) }
        }
        return { sessionId, found }
      })
      return ok({
        hasCommands: commands !== null && typeof commands?.find === 'function',
        hasAgents: agents !== null,
        liveAgents: live.length,
        agents: rows,
        usage: COMMAND_USAGE,
      })
    }

    case 'assist.providers': {
      // Diagnostics: which model routes this profile offers and which one the panel
      // would use. Makes "why did the sentence-to-style call fail" answerable without
      // reading Harness logs.
      const llm = ctx?.get?.('llm') ?? null
      const selection = (() => {
        try {
          return ctx?.get?.('agentDefaultModel')?.currentSelection?.() ?? null
        } catch {
          return null
        }
      })()
      let providers = []
      try {
        providers = typeof llm?.listProviders === 'function' ? llm.listProviders() : []
      } catch {
        providers = []
      }
      return ok({
        hasLlm: llm !== null && typeof llm?.stream === 'function',
        selection,
        providers,
        note: 'The panel authors styles/decks with the default selection above; a provider that requires extra routing headers may reject plugin-initiated calls.',
      })
    }

    case 'style.assist': {
      const description = typeof body.description === 'string' ? body.description.trim() : ''
      if (description === '') return fail('invalid_body', 'Expected { description }')
      const templates = await store.listTemplates(storeConfig)
      const requested = typeof body.templateId === 'string' && body.templateId !== '' ? body.templateId : null
      const templateId = requested ?? config.defaultTemplateId ?? templates[0]?.templateId ?? null
      const theme = templateId === null ? DEFAULT_THEME : await themeFor({ templateId }, config, storeConfig)
      try {
        return ok(await assistStyle({ description, theme }, { ctx, storeConfig }))
      } catch (error) {
        return failureOf(error)
      }
    }

    case 'content.assist': {
      // Two shapes, one endpoint: a sentence ('description' / 'text') or a long document
      // ('text' or a 'path' to read). The pipeline decides single-pass vs Map-Reduce.
      let material = typeof body.text === 'string' && body.text.trim() !== '' ? body.text : ''
      if (material === '' && typeof body.description === 'string') material = body.description.trim()
      if (material === '' && typeof body.path === 'string' && body.path.trim() !== '') {
        try {
          material = await readFile(resolvePath(body.path), 'utf8')
        } catch {
          return fail('source_not_found', `cannot read material file ${body.path}`)
        }
      }
      if (material.trim() === '') return fail('invalid_body', 'Expected { description }, { text } or { path }')
      const templates = await store.listTemplates(storeConfig)
      const requested = typeof body.templateId === 'string' && body.templateId !== '' ? body.templateId : null
      const templateId = requested ?? config.defaultTemplateId ?? templates[0]?.templateId ?? null
      try {
        const result = await assistContent(
          {
            material,
            description: material,
            title: typeof body.title === 'string' ? body.title : undefined,
            templateId,
            styleId: typeof body.styleId === 'string' ? body.styleId : undefined,
            maxSlides: Number.isFinite(body.maxSlides) ? Number(body.maxSlides) : undefined,
            // Per-request overrides. The same knobs live in the plugin config, but a
            // caller that knows its material ("this report is 900 chars, keep it
            // single-pass") must not have to edit the profile to say so.
            slideCharBudget: Number.isFinite(body.slideCharBudget) ? Number(body.slideCharBudget) : undefined,
            bulletCharBudget: Number.isFinite(body.bulletCharBudget) ? Number(body.bulletCharBudget) : undefined,
            materialThreshold: Number.isFinite(body.materialThreshold) ? Number(body.materialThreshold) : undefined,
          },
          { ctx, config, storeConfig },
        )
        return result.ok === true ? ok(result) : result
      } catch (error) {
        return failureOf(error)
      }
    }

    case 'decks.save': {
      const candidate = body.deck
      if (candidate === null || typeof candidate !== 'object') {
        return fail('invalid_body', 'Expected { deck }')
      }
      const theme = await themeFor(candidate, config, storeConfig)
      const repaired = parseAndRepairDeck(candidate, theme, budgetOptions(config))
      if (!repaired.ok) {
        return fail('invalid_deck', 'The deck does not match the CardDeck schema', repaired.issues)
      }
      const normalized = {
        ...repaired.deck,
        deckId: repaired.deck.deckId || slugify(repaired.deck.title ?? 'deck'),
        templateId: repaired.deck.templateId || theme.templateId,
      }
      const saved = await store.writeDeck(normalized, storeConfig, { theme })
      return ok({
        deck: normalized,
        theme,
        warnings: repaired.warnings,
        revision: Number(saved?.revision ?? 0),
        updatedAt: saved?.updatedAt ?? null,
      })
    }

    case 'decks.create': {
      const title =
        typeof body.title === 'string' && body.title.trim() !== '' ? body.title.trim() : '未命名卡组'
      const templateId = typeof body.templateId === 'string' ? body.templateId : undefined
      const theme = await themeFor({ templateId }, config, storeConfig)
      const deck = await store.createDeck(
        { title, templateId: theme.templateId, deckId: body.deckId, theme, persist: true },
        storeConfig,
      )
      return ok({ deck, theme, warnings: [] })
    }

    case 'export': {
      const deckId = typeof body.deckId === 'string' && body.deckId !== '' ? body.deckId : null
      const deck = deckId === null ? await activeDeck(storeConfig) : readDeckOrNull(deckId, storeConfig)
      if (deck === null) return fail('deck_not_found', 'No deck has been saved yet')
      const theme = await themeFor(deck, config, storeConfig)
      const repaired = parseAndRepairDeck(deck, theme, budgetOptions(config))
      if (!repaired.ok) {
        return fail('invalid_deck', 'The stored deck no longer matches the schema', repaired.issues)
      }
      // The panel has no session working directory, so exports land in the store's
      // export folder; the agent-facing tool still defaults to the session cwd.
      const exportDir = join(store.storeRoot(storeConfig), 'exports')
      await mkdir(exportDir, { recursive: true })
      const outputPath =
        typeof body.outputPath === 'string' && body.outputPath.trim() !== ''
          ? resolvePath(body.outputPath)
          : join(exportDir, `${repaired.deck.deckId}.pptx`)
      const result = await compileDeckToFile(
        {
          deck: repaired.deck,
          theme,
          logoBytes: await readLogoBytes(theme),
          assets: await resolveIconAssets(theme),
          options: compileOptions(config),
        },
        outputPath,
      )
      return ok({
        path: outputPath,
        slideCount: result.slideCount,
        warnings: [...repaired.warnings, ...(result.warnings ?? [])],
      })
    }

    case 'templates.parse': {
      const path = typeof body.path === 'string' ? body.path : ''
      if (path === '') return fail('invalid_body', 'Expected { path }')
      const absolute = resolvePath(path)
      const bytes = await readFile(absolute)
      const templateName =
        typeof body.templateName === 'string' && body.templateName.trim() !== ''
          ? body.templateName.trim()
          : absolute.split(/[\\/]/).pop().replace(/\.pptx$/i, '')
      const templateId =
        typeof body.templateId === 'string' && body.templateId.trim() !== ''
          ? body.templateId.trim()
          : slugify(templateName)
      const parsed = await parsePptxTemplate({
        bytes,
        path: absolute,
        templateId,
        templateName,
        storeConfig,
      })
      await store.writeTemplate(
        {
          templateId,
          templateName,
          tokens: parsed.theme,
          logoBytes: parsed.logoBytes,
          meta: parsed.meta,
          media: parsed.media,
          sourcePath: absolute,
          sourceBytes: bytes,
          layoutsInventory: parsed.layouts ?? null,
          markdown: parsed.markdown ?? null,
        },
        storeConfig,
      )
      return ok({
        template: {
          templateId,
          templateName,
          hasLogo: Boolean(parsed.logoBytes),
          colors: parsed.theme.colors,
          fonts: parsed.theme.fonts,
        },
        theme: parsed.theme,
      })
    }

    case 'templates.logo': {
      const templateId = typeof body.templateId === 'string' ? body.templateId : ''
      if (templateId === '') return fail('invalid_body', 'Expected { templateId }')
      let template
      try {
        template = await store.readTemplate(templateId, storeConfig)
      } catch {
        return fail('logo_not_found', `template "${templateId}" is not registered`)
      }
      const logoPath = template?.theme?.brandAssets?.logoPath
      if (typeof logoPath !== 'string' || logoPath === '') {
        return fail('logo_not_found', 'This template has no master logo')
      }
      try {
        const bytes = await readFile(logoPath)
        return ok({
          templateId,
          mime: 'image/png',
          dataUrl: `data:image/png;base64,${bytes.toString('base64')}`,
          widthPx: template.theme.brandAssets?.logoWidthPx ?? null,
          heightPx: template.theme.brandAssets?.logoHeightPx ?? null,
        })
      } catch {
        return fail('logo_not_found', 'The stored logo file is missing')
      }
    }

    default:
      return fail('not_found', `Unknown BrandDeck endpoint "${endpoint}"`)
  }
}

/** Compiler options shared by the tool and the panel export paths. */
function compileOptions(config) {
  return {
    fontFiles: config.fontFiles,
    fontDirs: config.fontDirs,
    loadSystemFonts: config.loadSystemFonts,
    svgRasterScale: config.svgRasterScale,
    logoSide: config.logoSide,
  }
}

export const __internals = { normalizeConfig, slugify, resolveColor, estimateTextOverflow, dispatchEndpoint, handleApiRequest, executeBrandDeckCommand }
