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
import { parsePptxTemplate } from './src/template.js'
import { compileDeckToFile } from './src/compile.js'

/** Cordis services this plugin cannot work without. */
export const inject = ['tools']

/** Tool names, exported so tests and docs cannot drift. */
export const TOOL_NAMES = Object.freeze({
  templateParse: 'branddeck_template_parse',
  templateList: 'branddeck_template_list',
  themeGet: 'branddeck_theme_get',
  deckCreate: 'branddeck_deck_create',
  deckSave: 'branddeck_deck_save',
  deckGet: 'branddeck_deck_get',
  deckList: 'branddeck_deck_list',
  deckExport: 'branddeck_deck_export',
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
  if (config.promptSection) registerPromptSection(ctx, config)
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
  }
}

function clampNumber(value, min, max, fallback) {
  const number = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(number)) return fallback
  return Math.min(max, Math.max(min, number))
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
      const repaired = parseAndRepairDeck(candidate, theme)
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
      const repaired = repairDeck(deck, theme)
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
      const repaired = parseAndRepairDeck(deck, theme)
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
}

async function firstTemplateId(storeConfig) {
  const templates = await store.listTemplates(storeConfig)
  return templates[0]?.templateId
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
1. 模板：${TOOL_NAMES.templateParse} { path } —— 从企业 .pptx 模板提取品牌色板、中英文字体与母版 Logo。（已解析过就跳过，用 ${TOOL_NAMES.templateList} 查看）
2. 对齐品牌：${TOOL_NAMES.themeGet} { templateId } —— 拿到 colors/fonts，写元素时用 "$colors.primary" 这类令牌，不要硬编码颜色。
3. 建卡组：${TOOL_NAMES.deckCreate} { title, templateId }。
4. 写内容：${TOOL_NAMES.deckSave} { deck } —— 一次提交整份 deck；返回的 warnings 必须处理（溢出/越界/图表长度不匹配）。
5. 导出：${TOOL_NAMES.deckExport} { deckId, outputPath }。

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
            handleApiRequest(req, res, { config, storeConfig }).catch((error) => {
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
  const body = payload ?? {}

  switch (endpoint) {
    case 'state': {
      const decks = await store.listDecks(storeConfig)
      const templates = await store.listTemplates(storeConfig)
      const deck = await activeDeck(storeConfig)
      const theme = deck === null ? DEFAULT_THEME : await themeFor(deck, config, storeConfig)
      const repaired = deck === null ? null : repairDeck(deck, theme)
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
      const repaired = repairDeck(deck, theme)
      return ok({ deck: repaired.deck, theme, warnings: repaired.warnings, decks })
    }

    case 'decks.save': {
      const candidate = body.deck
      if (candidate === null || typeof candidate !== 'object') {
        return fail('invalid_body', 'Expected { deck }')
      }
      const theme = await themeFor(candidate, config, storeConfig)
      const repaired = parseAndRepairDeck(candidate, theme)
      if (!repaired.ok) {
        return fail('invalid_deck', 'The deck does not match the CardDeck schema', repaired.issues)
      }
      const normalized = {
        ...repaired.deck,
        deckId: repaired.deck.deckId || slugify(repaired.deck.title ?? 'deck'),
        templateId: repaired.deck.templateId || theme.templateId,
      }
      await store.writeDeck(normalized, storeConfig)
      return ok({ deck: normalized, theme, warnings: repaired.warnings })
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
      const repaired = parseAndRepairDeck(deck, theme)
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

export const __internals = { normalizeConfig, slugify, resolveColor, estimateTextOverflow, dispatchEndpoint, handleApiRequest }
