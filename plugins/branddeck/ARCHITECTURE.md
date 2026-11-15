# BrandDeck — DSH 插件实现契约（Architecture & Build Contract）

> 本文是并行开发的唯一契约。任何模块的公开签名、HTTP 协议、文件归属以本文为准。
> 设计依据：`设计文档.md`。产品名：**BrandDeck 品牌演示**。

## 1. 形态

一个 **DSH bundle 插件包**（Cordis Host 半 + Web Client 半），通过 `plugin_manager install_bundle`
安装进 profile `desktop`，跨会话持久生效。

```
plugins/branddeck/
  package.json          # bundle 清单：dsh.bundle.patch + dsh.client + exports(locale/icon/client)
  cordis.patch.yml      # 插入 host 插件行
  index.js              # ★ Host 半：apply(ctx, config)，注册工具 + HTTP API + 系统提示段
  client.js             # ★ Web Client 半：window.__ModuleLoader__.load({...})，面板 + 侧栏图标
  locale/en.json        # meta.title / meta.description（插件管理器显示）
  locale/zh.json
  icon.svg              # 原创图标（品牌卡组），≤256KiB
  src/schema.js         # ★ 已完成：CardDeck DSL、Zod 校验、坐标修复、SVG 消毒、溢出估算
  src/defaults.js       # 默认 ThemeTokens（无模板时的兜底）
  src/store.js          # 模板/卡组持久化（$DSH_HOME/branddeck）
  src/template.js       # .pptx → ThemeTokens + Logo + 媒体清单
  src/pptx/read.js      # jszip 解包 + fast-xml-parser 封装 + rels 解析
  src/pptx/theme.js     # clrScheme / fontScheme / sldSz 提取
  src/pptx/media.js     # 母版 Logo 定位 + 图片尺寸嗅探
  src/svg.js            # @resvg/resvg-js 光栅化（含 CJK 字体处理）
  src/compile.js        # CardDeck → pptxgenjs → .pptx Buffer/文件
  test/*.mjs            # 可执行的 Node 测试（无需 DSH 运行时）
```

依赖（已 `pnpm install` 通过，smoke test 全绿）：`jszip`、`fast-xml-parser`、`zod`、
`pptxgenjs`、`@resvg/resvg-js`（本地原生预编译包，Windows x64 可用）。

**归属（写作用域，互不重叠）**

| 角色 | 文件 |
| --- | --- |
| Lead | `index.js`, `cordis.patch.yml`, `package.json`, `locale/*`, `icon.svg`, `src/schema.js`, `README.md` |
| host-engine | `src/defaults.js`, `src/store.js`, `src/template.js`, `src/pptx/*.js`, `src/svg.js`, `src/compile.js`, `test/engine-*.mjs` |
| client-ui | `client.js`, `test/client-*.mjs`（只做语法/静态检查） |

## 2. 数据契约

唯一来源 `src/schema.js`（已实现，勿改签名）：

- `CANVAS = { width: 1280, height: 720 }`，`SAFE_AREA = {top:64,right:64,bottom:64,left:64}`
- `CardDeckSchema` / `CardSlideSchema` / `SlideElementSchema` / `ThemeTokensSchema`（Zod）
- `parseAndRepairDeck(input, theme) -> {ok:false,issues[]} | {ok:true, deck, warnings[]}`
- `repairDeck(deck, theme) -> {deck, warnings[]}` —— 坐标钳制、Logo 角落避让、SVG 消毒、图表截断
- `resolveColor(value, theme, fallback)` —— `$colors.primary` / `colors.primary` / `#HEX` → `#RRGGBB`
- `validateSvg(code) -> {sanitized, problems[]}`
- `estimateLines(text, w, fontSize)`、`estimateTextOverflow(slide, theme)`
- `innerBox(theme)` —— 安全区（含 Logo 预留）

**坐标换算**：`inch = px / 96`；字号 `pt = px * 0.75`。1280×720 → 13.333×7.5in（`LAYOUT_WIDE`）。

## 3. 存储布局（`src/store.js`）

根目录 `BRANDDECK_HOME` = `process.env.DSH_HOME ?? ~/.dsh` + `/branddeck`（可被插件 Config 覆盖）。

```
<root>/templates/<templateId>/theme.json     # ThemeTokens（含 brandAssets.logoPath）
<root>/templates/<templateId>/logo.png       # 母版提取的 Logo（可缺）
<root>/templates/<templateId>/source.pptx    # 原始模板副本（可选，便于重解析）
<root>/templates/<templateId>/meta.json      # {templateId, templateName, sourcePath, slideCount, updatedAt, media[]}
<root>/decks/<deckId>.json                   # CardDeck
```

`templateId` / `deckId` 必须路径安全（`^[a-zA-Z0-9._-]{1,64}$`），否则拒绝并回错误。

## 4. Host 半：模型侧工具（`index.js`）

工具名统一前缀 `branddeck_`。**所有工具都返回 `{ ok, ... }` JSON 文本**，失败带 `error.code`+`error.message`。

| 工具 | 入参 | 作用 |
| --- | --- | --- |
| `branddeck_template_parse` | `{ path, templateId?, templateName? }` | 解析磁盘上的 `.pptx` 模板 → 提取 ThemeTokens + Logo，落盘，返回色板/字体摘要 |
| `branddeck_template_list` | `{}` | 已解析模板列表（含色板摘要） |
| `branddeck_theme_get` | `{ templateUrl?, templateId? }` | 读取 tokens（供模型写 Deck 时对齐品牌色） |
| `branddeck_deck_create` | `{ title, templateId?, deckId? }` | 建空 Deck（含封面骨架，元素由模型填） |
| `branddeck_deck_save` | `{ deck }` | Zod 校验 + 修复 + 落盘，返回 `warnings` |
| `branddeck_deck_get` | `{ deckId }` | 读回 Deck |
| `branddeck_deck_list` | `{}` | Deck 列表 |
| `branddeck_deck_export` | `{ deckId, outputPath? }` | 编译原生 `.pptx`（默认 `<cwd>/<deckId>.pptx`），返回路径/页数 |
| `branddeck_deck_preview` | `{ deckId, slideIndex?, outputDir? }` | 光栅化单页为 PNG（调试用，走 resvg 渲染 SVG 元素 + 简易排版） |

系统提示段 `branddeck`：说明“先解析模板 → 读 tokens → 用 1280×720 网格写 CardDeck DSL →
`branddeck_deck_save` → 打开面板微调 → `branddeck_deck_export`”，并给出 DSL 关键规则
（安全区 64px、右上 Logo 预留 150×48、单页正文 ≤150 字、SVG 仅允许纯矢量标签）。
该段仅在配置 `enablePromptSection` 为 true 时注册。

**LLM 流水线的落地方式**：DSH 里“纯文本大模型生成”由当前会话的 Agent 承担——
插件提供工具（结构化输出 + 强校验 + 修复）与提示段（信息密度/版式规范），
不在 Host 内部二次调用模型。这样卡片 DSL 的生成、按页重排、精简文案都能被用户在对话里直接驱动。

## 5. Host 半：面板通道（同源 HTTP 路由）

面板与 Host 之间走**同源 HTTP 路由**：Host 侧
`ctx.webServer.register({ kind: 'prefix', path: '/branddeck-api', handler })`，
浏览器侧 `fetch('/branddeck-api/<endpoint>', { method: 'POST', body: JSON.stringify(payload) })`。

> **为什么不走 `ctx.connection.rpc`**：Client 半是普通 bundle 插件，与内置的客户端插件是
> **兄弟关系**；`connection` 由 `dsh-client-connection` 用 `ctx.provide` 装在自己的 ctx 上，
> 兄弟插件看不到。把它写进 `inject` 会让本插件永久 pending，表现为启动失败
> `@local/dsh-branddeck: loading`（`web boot: 1 entry did not activate`）。
> 因此客户端 `inject` 只声明 `['slots']`，数据走 web server 路由——与 `/plugins`
> 客户端模块用的是同一个载体。

统一信封

```ts
{ ok: true, value } | { ok: false, error: { code, message, details? } }
```

端点表（`POST /branddeck-api/<endpoint>`，请求体即 `payload`）：

| endpoint | payload | value |
| --- | --- | --- |
| `state` | `{}` | `{ decks, templates, deck, theme, warnings, activeDeckId, activeTemplateId, capability }` |
| `decks.list` | `{}` | `{ decks }` |
| `decks.get` | `{ deckId? }` | `{ deck, theme, warnings, decks }` |
| `decks.save` | `{ deck }` | `{ deck, theme, warnings }`（校验+修复后回写） |
| `decks.create` | `{ title, templateId?, deckId? }` | `{ deck, theme, warnings }` |
| `export` | `{ deckId?, outputPath? }` | `{ path, slideCount, warnings }` |
| `templates.parse` | `{ path, templateId?, templateName? }` | `{ template, theme }` |
| `templates.logo` | `{ templateId }` | `{ dataUrl, mime, widthPx, heightPx }` |

HTTP 状态码只作提示，**`error.code` 才是业务真相**：
`deck_not_found` / `template_not_found` / `logo_not_found` / `not_found` → 404，
`invalid_deck` / `invalid_body` / `invalid_request` → 422，其余 → 500。
跨站请求（`sec-fetch-site: cross-site`）→ 403，非 GET/POST → 405。

`DeckSummary = { deckId, title, templateId, slideCount, updatedAt }`
`TemplateSummary = { templateId, templateName, hasLogo, colors, fonts, updatedAt }`
`DeckView = { deck: CardDeck, theme, warnings: string[] }`

- 未指定 `deckId` 时取最近修改的 Deck（`activeDeckId`）。
- 面板导出落到 `<store>/exports/<deckId>.pptx`（面板无会话工作目录）；
  Agent 工具导出默认落到 `<session cwd>/branddeck/<deckId>.pptx`，可用 `outputPath` 覆盖。

## 6. Web Client 半（`client.js`）

单文件、纯 JS、无 JSX（用 `React.createElement` 或 `h` 助手）、无外部 import，
`window.__ModuleLoader__.load({ id: '@local/dsh-branddeck', factory(require) {...} })`。

- 注册 `main`（keyed，`key: 'branddeck'`）+ `sidebar.panellist`（`id: 'branddeck'`，`order` 略大于 10）
  ；`ctx.slots.inject(name, () => ctx.slots.register(opts, Component))`。
- `main` 面板组件**不接收 props**（根作用域），所有数据自取。
- 数据通道：`fetch('/branddeck-api/<endpoint>', { method: 'POST', body: JSON.stringify(payload) })`（见 §5）；
  客户端 `inject` **只**声明 `['slots']`；接口失败时渲染可读错误态，且 `apply` 内任何异常都不得冒泡
  （否则整个 web boot 失败：`web boot: 1 entry did not activate`）。
- 面板结构：顶栏（Deck 选择 + 模板 + 保存状态 + 导出 + 刷新）／左栏缩略图列表／
  中间 16:9 画布（`transform: scale()` 自适应宽度，基准 1280×720）／右侧属性条（选中元素的数值坐标）。
- 交互（设计文档模块三）：双击改字（contentEditable，失焦/回车提交）、拖拽移动、
  角点缩放、方向键微调（Shift=10px）、Delete 删除、Ctrl/Cmd+D 复制、缩略图上下移、自动保存（600ms 防抖 `decks.save`）。
- 视觉预警：文本溢出（用 `estimateTextOverflow` 同款逻辑在前端估算）→ 红色描边 + 角标。
- 样式：**仅**使用 `--dsw-alias-*` 主题令牌做外壳；卡片内部颜色来自 `ThemeTokens`（属“作品”）。
- 无数据/接口失败时要给出可读空状态与错误态，不得抛错导致 slot entry 崩溃。
- 语言：文案取 `ctx.locale`（`zh-CN` / `en`）二选一，未知回落 `en`。

## 7. 验证

1. `node test/engine-*.mjs`（host-engine 全绿）。
2. `node --check client.js` + `node --check index.js`。
3. `plugin_manager install_bundle` → `application: applied`。
4. 真机闭环：用 `PPT模板-HW.pptx` 解析 → 生成示例 Deck → 导出 `.pptx` →
   LibreOffice 渲染 PNG 人工核对（CJK 不乱码、Logo 在位、无溢出）→ 面板中编辑保存 → 再次导出。
