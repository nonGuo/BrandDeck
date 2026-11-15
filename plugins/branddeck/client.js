/**
 * BrandDeck — Web Client half (DSH client plugin).
 *
 * A single self-contained module file: the Host side (`index.js`) registers the
 * tools and the HTTP API, this file renders the interactive 16:9 deck editor in
 * the Web GUI. It is plain ES2020 JavaScript — no JSX, no TypeScript, no
 * bundler, no external network assets — and it loads React from the host module
 * registry through the single `require` call at the top of the factory.
 *
 * Transport contract: the Host half mounts `POST /branddeck-api/<endpoint>` on
 * the Harness web server, so every call is a same-origin `fetch` that resolves
 * the envelope `{ ok: true, value }` or
 * `{ ok: false, error: { code, message, details? } }`. Calls reject only on
 * transport failure; business failures arrive as `ok: false`.
 *
 * The plugin deliberately injects nothing but `slots`: services provided by the
 * shipped client plugins (the Connection service in particular) live on their own
 * contexts and are invisible to a sibling bundle plugin, and declaring one as a
 * hard dependency would leave this entry pending and break the whole web boot.
 *
 * The geometry helpers at the top are faithful re-implementations of the pure
 * helpers in `src/schema.js` (`resolveColor`, `estimateLines`,
 * `estimateTextOverflow`, `innerBox`) because the browser half cannot load that
 * ES module. Keep them in sync with the Host contract: same constants, same
 * behaviour, same warning thresholds.
 */
window.__ModuleLoader__.load({
  id: '@local/dsh-branddeck',
  factory(require) {
    const React = require('react')
    const h = React.createElement

    /** Locale namespace of this plugin (slot registrations + dictionaries). */
    const NS = 'branddeck'
    /** Key/id of both slot registrations: the panel page and its sidebar icon. */
    const PANEL_ID = 'branddeck'
    /** Same-origin route prefix the Host half mounts for this panel. */
    const API_BASE = '/branddeck-api'
    /** Ordinary panel calls are abandoned after this long (UI safety bound). */
    const REQUEST_TIMEOUT_MS = 8000
    /** Export compiles a real .pptx on the Host, so it gets a longer bound. */
    const EXPORT_TIMEOUT_MS = 120000
    /** Autosave debounce (design doc: immediate feedback, quiet network). */
    const AUTOSAVE_DEBOUNCE_MS = 600

    /* ------------------------------------------------------------------ *
     * Geometry constants — mirror of src/schema.js.
     * ------------------------------------------------------------------ */

    /** Base coordinate space of the deck DSL: a 16:9 virtual canvas. */
    const CANVAS = Object.freeze({ width: 1280, height: 720 })
    /** Margin the model keeps free on every side. */
    const SAFE_AREA = Object.freeze({ top: 64, right: 64, bottom: 64, left: 64 })
    /** Logo reservation inside the top-right corner of the safe area. */
    const LOGO_RESERVE = Object.freeze({ width: 150, height: 48, gap: 16 })

    /** Smallest element box the panel lets the user drag out (schema floor). */
    const MIN_ELEMENT_SIZE = 8
    /** Schema limits, enforced in the UI so local edits always validate. */
    const MAX_ELEMENTS_PER_SLIDE = 40
    const MAX_SLIDES = 60
    /** Keyboard nudge distances in base pixels. */
    const NUDGE_STEP = 1
    const NUDGE_STEP_FAST = 10
    /** Offset applied to a duplicated element so it does not hide the original. */
    const DUPLICATE_OFFSET = 16
    /** Corner handle edge length in screen pixels (shrunk by the canvas scale). */
    const HANDLE_SIZE = 10
    /** Thumbnail rail geometry: the mini render is the real canvas at this scale. */
    const THUMB_WIDTH = 144
    const THUMB_SCALE = THUMB_WIDTH / CANVAS.width
    /** Smallest canvas scale we allow; keeps the deck readable in a narrow panel. */
    const MIN_CANVAS_SCALE = 0.1

    /** Neutral palette used when the Host reports no theme (or none is loadable). */
    const FALLBACK_THEME = Object.freeze({
      templateId: 'fallback',
      templateName: 'BrandDeck',
      colors: Object.freeze({
        background: '#FFFFFF',
        surface: '#F4F6F9',
        textPrimary: '#1B1F24',
        textSecondary: '#5B6470',
        primary: '#247BBF',
        accent: '#F2A03D',
        chartPalette: Object.freeze(['#247BBF', '#F2A03D', '#4CAF93', '#9B7BD4', '#E06C75', '#56A8D8']),
      }),
      fonts: Object.freeze({
        title: 'Microsoft YaHei',
        body: 'Microsoft YaHei',
        latin: 'Arial',
      }),
      brandAssets: Object.freeze({ icons: Object.freeze({}) }),
    })

    /* ------------------------------------------------------------------ *
     * Dictionaries. `const ZH` / `const EN` are also read by the static
     * smoke test, which asserts every Chinese literal of this file lives in ZH
     * and that both dictionaries carry exactly the same key set.
     * ------------------------------------------------------------------ */

    const ZH = {
      'panel': '品牌演示',
      'appTitle': 'BrandDeck 品牌演示',
      'action.refresh': '刷新',
      'view.switch': '切换视图',
      'view.cards': '卡片视图',
      'view.text': '文字视图',
      'view.gen': '生成',
      'gen.stepTemplate': '模板',
      'gen.stepStyle': '风格',
      'gen.stepContent': '内容',
      'gen.noTemplate': '还没有解析过模板',
      'gen.importPlaceholder': '导入 .pptx：粘贴绝对路径',
      'gen.import': '导入模板',
      'gen.importing': '解析中…',
      'gen.importFailed': '导入失败',
      'gen.imported': '模板已导入',
      'gen.needPath': '请先填写 .pptx 的路径',
      'gen.noSpec': '选择或导入模板后，这里会显示色板、字体与版式数量。',
      'gen.fonts': '字体',
      'gen.layouts': '版式',
      'gen.fromChat': '来自对话',
      'gen.derive': '从模板派生风格',
      'gen.derived': '派生风格',
      'gen.derivedDesc': '从模板令牌派生：沿用色板与字体，标准信息密度。',
      'gen.saveStyle': '保存风格',
      'gen.saving': '保存中…',
      'gen.styleFailed': '风格保存失败',
      'gen.styleSaved': '风格已保存',
      'gen.deleteStyle': '删除风格',
      'gen.density': '信息密度',
      'gen.density.airy': '疏（每页 3 条）',
      'gen.density.standard': '标准（每页 4 条）',
      'gen.density.dense': '密（每页 6 条）',
      'gen.tone': '强调用量',
      'gen.tone.formal': '克制',
      'gen.tone.bold': '强烈',
      'gen.tone.minimal': '极简',
      'gen.bullets': '正文排布',
      'gen.bullets.list': '列表',
      'gen.bullets.cards': '卡片',
      'gen.maxSlides': '页数上限',
      'gen.maxSlidesShort': '页数上限',
      'gen.describe': '风格说明',
      'gen.describePlaceholder': '例如：华为红、信息密度高、卡片式正文',
      'gen.askAgentTitle': '用对话生成风格配置（复制给 Agent）',
      'gen.askAgent': '请把下面的描述整理成一个 BrandDeck 风格配置，调用：',
      'gen.copy': '复制指令',
      'gen.agentExample': 'branddeck_style_save { name: "风格名", description: "刚才那段描述", density: "dense", tone: "bold", bulletLayout: "cards", templateId: "{templateId}" }',
      'gen.textPlaceholder': '粘贴会议纪要、文档或一段描述；支持 Markdown：# 标题 为卡组名，## 小节 为每页，-  为要点，备注： 为演讲者备注。',
      'gen.filePlaceholder': '或填一个 .md/.txt 路径',
      'gen.titlePlaceholder': '覆盖卡组标题（可选）',
      'gen.generate': '生成卡组',
      'gen.generating': '生成中…',
      'gen.failed': '生成失败',
      'gen.done': '已生成',
      'gen.slides': '页',
      'gen.chatStyleTitle': '说两句，出风格',
      'gen.chatStylePlaceholder': '例：华为红、信息密度高、正文用卡片、别超过 10 页',
      'gen.chatStyleAction': '生成风格草稿',
      'gen.chatStyleBusy': '正在理解…',
      'gen.styleDraft': '已生成风格草稿「{name}」，检查后点「保存风格」即可复用。',
      'gen.chatContentTitle': '说两句，出内容',
      'gen.chatContentPlaceholder': '例：给部门周会做一页汇报，讲本周做完模板解析、下周启动导出引擎',
      'gen.chatContentAction': '生成卡组',
      'gen.chatContentBusy': '正在成稿…',
      'gen.needSentence': '先用一句话描述你想要的',
      'gen.assistFailed': '生成失败',
      'text.hint': '直接改文字即可；保存后按元素 id 回写，卡片里的位置与样式保持不变。',
      'text.ariaLabel': '卡组文字版（Markdown）',
      'text.loading': '读取文字版…',
      'text.saving': '应用中…',
      'text.dirty': '未应用',
      'text.synced': '与卡片一致',
      'text.failed': '应用失败',
      'text.revert': '还原',
      'text.apply': '保存并应用到卡片',
      'text.applied': '已应用 · 修订 {revision} · 更新 {updated} 新增 {added} 删除 {removed}',
      'action.export': '导出 PPTX',
      'action.retry': '重试',
      'action.duplicateSlide': '+ 复制页',
      'action.duplicateElement': '复制元素',
      'action.deleteElement': '删除元素',
      'action.moveUp': '上移',
      'action.moveDown': '下移',
      'action.deleteSlide': '删除本页',
      'action.dismiss': '忽略',
      'action.close': '关闭',
      'label.deck': '卡组',
      'label.template': '模板',
      'label.slideCount': '{count} 页',
      'label.slides': '页面',
      'label.slideIndex': '第 {index} / {total} 页',
      'label.position': '位置与尺寸',
      'label.typography': '文本样式',
      'label.fontSize': '字号',
      'label.align': '对齐',
      'label.color': '颜色',
      'label.x': 'X',
      'label.y': 'Y',
      'label.w': '宽',
      'label.h': '高',
      'label.noSelection': '未选择元素',
      'label.selectHint': '在画布上点击元素即可编辑',
      'label.noDeck': '没有卡组',
      'align.left': '左',
      'align.center': '中',
      'align.right': '右',
      'empty.title': '还没有卡组，让 Agent 用 branddeck_deck_create 生成',
      'empty.hint': '在对话里让 Agent 生成 1280×720 卡组后，这里会自动出现。',
      'state.loading': '加载中…',
      'save.idle': '就绪',
      'save.saving': '保存中…',
      'save.saved': '已保存',
      'save.dirty': '未保存',
      'save.error': '保存失败',
      'error.load': '加载失败',
      'error.request': '请求失败（{code}）',
      'error.timeout': '请求超时，请稍后重试',
      'error.transport': '与 Host 的连接失败，请稍后重试',
      'error.transportUnavailable': 'Host 连接服务不可用，请重新加载插件',
      'error.render': '面板渲染出错',
      'overflow.badge': '溢出 +{lines} 行',
      'warnings.title': 'Host 校验提示',
      'warnings.more': '另有 {count} 条',
      'export.title': '导出',
      'export.running': '正在导出 PPTX…',
      'export.done': '已导出：{path}（{count} 页）',
      'export.failed': '导出失败：{message}',
      'confirm.deleteSlide': '确认删除这一页？',
      'thumb.label': '第 {index} 页',
      'image.missing': '缺少图片资源：{key}',
      'hint.keys': '双击改字 · 拖拽移动 · 方向键微调 · Ctrl/Cmd+S 保存',
      'hint.safeArea': '当前元素超出 {margin}px 安全区',
      'hint.lastElement': '每页至少保留一个元素',
    }

    const EN = {
      'panel': 'BrandDeck',
      'appTitle': 'BrandDeck',
      'action.refresh': 'Refresh',
      'view.switch': 'Switch view',
      'view.cards': 'Cards',
      'view.text': 'Text',
      'view.gen': 'Generate',
      'gen.stepTemplate': 'Template',
      'gen.stepStyle': 'Style',
      'gen.stepContent': 'Content',
      'gen.noTemplate': 'No template parsed yet',
      'gen.importPlaceholder': 'Import a .pptx: paste its absolute path',
      'gen.import': 'Import',
      'gen.importing': 'Parsing…',
      'gen.importFailed': 'Import failed',
      'gen.imported': 'Template imported',
      'gen.needPath': 'Enter the .pptx path first',
      'gen.noSpec': 'Pick or bring in a template to see its palette, fonts and layout count.',
      'gen.fonts': 'Fonts',
      'gen.layouts': 'Layouts',
      'gen.fromChat': 'from chat',
      'gen.derive': 'Derive from template',
      'gen.derived': 'derived style',
      'gen.derivedDesc': 'Derived from the template tokens: same palette and fonts, standard density.',
      'gen.saveStyle': 'Save style',
      'gen.saving': 'Saving…',
      'gen.styleFailed': 'Style save failed',
      'gen.styleSaved': 'Style saved',
      'gen.deleteStyle': 'Delete style',
      'gen.density': 'Density',
      'gen.density.airy': 'Airy (3 bullets)',
      'gen.density.standard': 'Standard (4 bullets)',
      'gen.density.dense': 'Dense (6 bullets)',
      'gen.tone': 'Accent usage',
      'gen.tone.formal': 'Formal',
      'gen.tone.bold': 'Bold',
      'gen.tone.minimal': 'Minimal',
      'gen.bullets': 'Bullet layout',
      'gen.bullets.list': 'List',
      'gen.bullets.cards': 'Cards',
      'gen.maxSlides': 'Page cap',
      'gen.maxSlidesShort': 'Max pages',
      'gen.describe': 'Style note',
      'gen.describePlaceholder': 'e.g. brand red, dense, card-style body text',
      'gen.askAgentTitle': 'Author a style by conversation (copy for the Agent)',
      'gen.askAgent': 'Turn the description below into a BrandDeck style profile and call:',
      'gen.copy': 'Copy instruction',
      'gen.agentExample': 'branddeck_style_save { name: "style name", description: "the description above", density: "dense", tone: "bold", bulletLayout: "cards", templateId: "{templateId}" }',
      'gen.textPlaceholder': 'Paste notes, a document or a description. Markdown works: # title = deck, ## section = slide, -  = bullet, a notes prefix = speaker notes.',
      'gen.filePlaceholder': 'or a .md/.txt path',
      'gen.titlePlaceholder': 'Override deck title (optional)',
      'gen.generate': 'Generate deck',
      'gen.generating': 'Generating…',
      'gen.failed': 'Generation failed',
      'gen.done': 'Generated',
      'gen.slides': 'slides',
      'gen.chatStyleTitle': 'Say a sentence, get a style',
      'gen.chatStylePlaceholder': 'e.g. brand red, dense, card-style body, max 10 pages',
      'gen.chatStyleAction': 'Draft style',
      'gen.chatStyleBusy': 'Understanding…',
      'gen.styleDraft': 'Draft「{name}」is ready — review it, then press Save style.',
      'gen.chatContentTitle': 'Say a sentence, get slides',
      'gen.chatContentPlaceholder': 'e.g. a weekly update for my team: template parsing shipped, export engine starts next week',
      'gen.chatContentAction': 'Generate deck',
      'gen.chatContentBusy': 'Writing…',
      'gen.needSentence': 'Describe it in one sentence first',
      'gen.assistFailed': 'Generation failed',
      'text.hint': 'Edit the copy directly; saving maps it back by element id and keeps every card position and style.',
      'text.ariaLabel': 'Deck text outline (Markdown)',
      'text.loading': 'Loading text…',
      'text.saving': 'Applying…',
      'text.dirty': 'Not applied',
      'text.synced': 'In sync with cards',
      'text.failed': 'Apply failed',
      'text.revert': 'Revert',
      'text.apply': 'Save and apply to cards',
      'text.applied': 'Applied · revision {revision} · updated {updated} added {added} removed {removed}',
      'action.export': 'Export PPTX',
      'action.retry': 'Retry',
      'action.duplicateSlide': '+ Duplicate slide',
      'action.duplicateElement': 'Duplicate',
      'action.deleteElement': 'Delete',
      'action.moveUp': 'Move up',
      'action.moveDown': 'Move down',
      'action.deleteSlide': 'Delete slide',
      'action.dismiss': 'Dismiss',
      'action.close': 'Close',
      'label.deck': 'Deck',
      'label.template': 'Template',
      'label.slideCount': '{count} slides',
      'label.slides': 'Slides',
      'label.slideIndex': 'Slide {index} / {total}',
      'label.position': 'Position & size',
      'label.typography': 'Text style',
      'label.fontSize': 'Font size',
      'label.align': 'Align',
      'label.color': 'Color',
      'label.x': 'X',
      'label.y': 'Y',
      'label.w': 'W',
      'label.h': 'H',
      'label.noSelection': 'No element selected',
      'label.selectHint': 'Click an element on the canvas to edit it',
      'label.noDeck': 'No deck',
      'align.left': 'L',
      'align.center': 'C',
      'align.right': 'R',
      'empty.title': 'No deck yet - ask the Agent to create one with branddeck_deck_create',
      'empty.hint': 'A 1280x720 deck written by the Agent in the conversation shows up here.',
      'state.loading': 'Loading…',
      'save.idle': 'Ready',
      'save.saving': 'Saving…',
      'save.saved': 'Saved',
      'save.dirty': 'Unsaved',
      'save.error': 'Save failed',
      'error.load': 'Failed to load',
      'error.request': 'Request failed ({code})',
      'error.timeout': 'The request timed out. Try again.',
      'error.transport': 'The Host connection failed - try again',
      'error.transportUnavailable': 'The Host connection service is unavailable - reload the plugin',
      'error.render': 'The panel hit a render error',
      'overflow.badge': '+{lines} lines',
      'warnings.title': 'Host warnings',
      'warnings.more': '{count} more',
      'export.title': 'Export',
      'export.running': 'Exporting PPTX…',
      'export.done': 'Exported: {path} ({count} slides)',
      'export.failed': 'Export failed: {message}',
      'confirm.deleteSlide': 'Delete this slide?',
      'thumb.label': 'Slide {index}',
      'image.missing': 'Missing image asset: {key}',
      'hint.keys': 'Double-click to edit - drag to move - arrows to nudge - Ctrl/Cmd+S to save',
      'hint.safeArea': 'This element sits outside the {margin}px safe area',
      'hint.lastElement': 'A slide keeps at least one element',
    }

    const DICT = { zh: ZH, en: EN }

    /* ------------------------------------------------------------------ *
     * Runtime seat. `apply()` fills it once; every component and the slot
     * label read it. Keeping it module-scoped means components need no props
     * and the slot label can be a stable closure.
     * ------------------------------------------------------------------ */

    let runtime = null

    /** @returns {'zh'|'en'} the active language id, English when unknown. */
    function activeLocaleId() {
      const ctx = runtime === null ? null : runtime.ctx
      try {
        const locale = ctx === null || ctx === undefined ? null : ctx.locale
        if (locale === null || locale === undefined) return 'en'
        const snapshot =
          typeof locale.getLocale === 'function'
            ? locale.getLocale()
            : typeof locale.getSnapshot === 'function'
              ? locale.getSnapshot()
              : null
        const raw = snapshot === null || snapshot === undefined ? null : snapshot.active ?? snapshot.id ?? snapshot.locale
        if (typeof raw === 'string' && /^zh/i.test(raw)) return 'zh'
        if (typeof raw === 'string' && /^en/i.test(raw)) return 'en'
      } catch (error) {
        /* A broken locale service must never break the panel. */
      }
      return 'en'
    }

    /**
     * Replace `{name}` placeholders with values from `params`.
     * @param {string} template
     * @param {Record<string, unknown> | undefined} params
     * @returns {string}
     */
    function interpolate(template, params) {
      if (typeof template !== 'string') return ''
      if (params === null || params === undefined) return template
      return template.replace(/\{(\w+)\}/g, (whole, key) =>
        Object.prototype.hasOwnProperty.call(params, key) ? String(params[key]) : whole,
      )
    }

    /**
     * Translate a key for the active language. The registered namespace
     * translation is preferred (so language packs can extend us); the built-in
     * dictionary is the fallback when the locale service is unavailable.
     * @param {string} key
     * @param {Record<string, unknown>} [params]
     * @returns {string}
     */
    function t(key, params) {
      if (runtime !== null && runtime.bound !== null) {
        try {
          const translated = runtime.bound(key, params)
          if (typeof translated === 'string' && translated.length > 0 && translated !== key) return translated
        } catch (error) {
          /* fall through to the built-in dictionary */
        }
      }
      const dict = DICT[activeLocaleId()] ?? EN
      const template = Object.prototype.hasOwnProperty.call(dict, key) ? dict[key] : EN[key]
      const resolved = typeof template === 'string' ? template : key
      return interpolate(resolved, params)
    }

    /* ------------------------------------------------------------------ *
     * Pure geometry / colour helpers (behaviour-identical to src/schema.js).
     * Every one of them tolerates a missing theme: a render must never throw.
     * ------------------------------------------------------------------ */

    const clamp = (value, min, max) => Math.min(max, Math.max(min, value))

    /** @param {unknown} theme @returns {typeof FALLBACK_THEME} */
    function normalizeTheme(theme) {
      if (theme === null || typeof theme !== 'object' || theme.colors === null || typeof theme.colors !== 'object') {
        return FALLBACK_THEME
      }
      const colors = theme.colors
      const fonts = theme.fonts !== null && typeof theme.fonts === 'object' ? theme.fonts : {}
      const brandAssets = theme.brandAssets !== null && typeof theme.brandAssets === 'object' ? theme.brandAssets : {}
      const palette =
        Array.isArray(colors.chartPalette) && colors.chartPalette.length > 0
          ? colors.chartPalette.filter((value) => typeof value === 'string')
          : FALLBACK_THEME.colors.chartPalette
      return {
        templateId: typeof theme.templateId === 'string' ? theme.templateId : FALLBACK_THEME.templateId,
        templateName: typeof theme.templateName === 'string' ? theme.templateName : FALLBACK_THEME.templateName,
        colors: {
          background: pickColor(colors.background, FALLBACK_THEME.colors.background),
          surface: pickColor(colors.surface, FALLBACK_THEME.colors.surface),
          textPrimary: pickColor(colors.textPrimary, FALLBACK_THEME.colors.textPrimary),
          textSecondary: pickColor(colors.textSecondary, FALLBACK_THEME.colors.textSecondary),
          primary: pickColor(colors.primary, FALLBACK_THEME.colors.primary),
          accent: pickColor(colors.accent, FALLBACK_THEME.colors.accent),
          chartPalette: palette.length > 0 ? palette : FALLBACK_THEME.colors.chartPalette,
        },
        fonts: {
          title: pickFont(fonts.title, FALLBACK_THEME.fonts.title),
          body: pickFont(fonts.body, FALLBACK_THEME.fonts.body),
          latin: pickFont(fonts.latin, FALLBACK_THEME.fonts.latin),
        },
        brandAssets: {
          logoPath: typeof brandAssets.logoPath === 'string' ? brandAssets.logoPath : undefined,
          icons: brandAssets.icons !== null && typeof brandAssets.icons === 'object' ? brandAssets.icons : {},
        },
      }
    }

    function pickColor(value, fallback) {
      return typeof value === 'string' && value.length > 0 ? value : fallback
    }

    function pickFont(value, fallback) {
      return typeof value === 'string' && value.length > 0 ? value : fallback
    }

    /**
     * True when a theme already carries the fields the helpers read, i.e. it
     * came from `normalizeTheme`. Keeps the colour hot path allocation-free.
     * @param {unknown} theme
     * @returns {boolean}
     */
    function isTheme(theme) {
      if (theme === null || typeof theme !== 'object') return false
      const colors = theme.colors
      if (colors === null || typeof colors !== 'object') return false
      if (typeof colors.primary !== 'string' || typeof colors.textPrimary !== 'string') return false
      return Array.isArray(colors.chartPalette) && theme.fonts !== null && typeof theme.fonts === 'object'
    }

    const HEX = /^#[0-9a-fA-F]{3,8}$/
    const TOKEN_REF = /^\$?colors\.([a-zA-Z]+)$/

    /**
     * Normalize `#abc` / `abc` / `AABBCC` into `#AABBCC` (alpha preserved).
     * @param {string} value
     * @returns {string}
     */
    function normalizeHex(value) {
      let hex = String(value).trim()
      if (!hex.startsWith('#')) hex = `#${hex}`
      const body = hex.slice(1)
      if (body.length === 3) {
        return `#${body[0]}${body[0]}${body[1]}${body[1]}${body[2]}${body[2]}`.toUpperCase()
      }
      return hex.toUpperCase()
    }

    /**
     * Resolve a colour that may be a HEX literal or a theme token reference
     * (`$colors.primary`, `colors.primary`, `$accent`, `$colors.chart1`).
     * @param {string | undefined} value
     * @param {unknown} theme
     * @param {string} fallback
     * @returns {string}
     */
    function resolveColor(value, theme, fallback = '#000000') {
      /* Callers pass the normalized render theme on hot paths (every element of
       * every frame), so only an unusable theme pays for normalization. */
      const safeTheme = isTheme(theme) ? theme : normalizeTheme(theme)
      if (typeof value !== 'string' || value.length === 0) return fallback
      const raw = value.trim()
      if (HEX.test(raw)) return normalizeHex(raw)
      let match = TOKEN_REF.exec(raw)
      if (match === null) {
        const short = /^\$?([a-zA-Z]+)$/.exec(raw)
        if (short !== null && short[1] in safeTheme.colors) match = [, short[1]]
      }
      if (match !== null) {
        const key = match[1]
        const palette = safeTheme.colors
        if (key in palette && typeof palette[key] === 'string') return normalizeHex(palette[key])
        const paletteMatch = /^chart(\d)$/.exec(key)
        if (paletteMatch !== null) {
          const hex = palette.chartPalette[Number(paletteMatch[1]) - 1]
          if (typeof hex === 'string') return normalizeHex(hex)
        }
      }
      return fallback
    }

    /**
     * Full-width (CJK) glyph detection, matching the Host estimator.
     * @param {string} char
     * @returns {boolean}
     */
    function isWide(char) {
      const code = char.codePointAt(0) ?? 0
      return (
        (code >= 0x1100 && code <= 0x115f) ||
        (code >= 0x2e80 && code <= 0xa4cf) ||
        (code >= 0xac00 && code <= 0xd7a3) ||
        (code >= 0xf900 && code <= 0xfaff) ||
        (code >= 0xfe30 && code <= 0xfe6f) ||
        (code >= 0xff00 && code <= 0xff60) ||
        (code >= 0xffe0 && code <= 0xffe6)
      )
    }

    /**
     * Rough line-count / height estimate, identical to `estimateLines`.
     * @param {string} text
     * @param {number} boxWidthPx
     * @param {number} fontSizePx
     * @returns {{ lines: number, charsPerLine: number, heightPx: number }}
     */
    function estimateLines(text, boxWidthPx, fontSizePx) {
      const safeFont = Math.max(8, Number(fontSizePx) || 20)
      let width = 0
      let explicitLines = 1
      for (const char of String(text ?? '')) {
        if (char === '\n') {
          explicitLines += 1
          width = 0
          continue
        }
        width += isWide(char) ? safeFont : safeFont * 0.55
      }
      const charsPerLine = Math.max(1, Math.floor(boxWidthPx / (safeFont * 0.55)))
      const linesFromWidth = Math.ceil(width / Math.max(1, boxWidthPx))
      const lines = Math.max(explicitLines, linesFromWidth)
      return { lines, charsPerLine, heightPx: lines * safeFont * 1.35 }
    }

    /**
     * Live overflow report for one slide (`estimateTextOverflow` equivalent).
     * The theme parameter is accepted for call-site parity with the Host helper;
     * the per-box estimate itself only needs the declared box and font size.
     * @param {any} slide
     * @param {unknown} theme
     * @returns {{ totalCharacters: number, overflowing: Array<{ id: string, estimatedLines: number, neededHeight: number, declaredHeight: number }> }}
     */
    function estimateTextOverflow(slide, theme) {
      void theme
      const overflowing = []
      let totalCharacters = 0
      const elements = slideElements(slide)
      for (const element of elements) {
        if (element.type !== 'text') continue
        totalCharacters += String(element.content ?? '').replace(/\s+/g, '').length
        const fontSize = element.style?.fontSize ?? 20
        const estimate = estimateLines(element.content, element.w, fontSize)
        if (estimate.heightPx > element.h + 1) {
          overflowing.push({
            id: element.id,
            estimatedLines: estimate.lines,
            neededHeight: Math.ceil(estimate.heightPx),
            declaredHeight: element.h,
          })
        }
      }
      return { totalCharacters, overflowing }
    }

    /**
     * Safe drawing area plus the logo reservation (`innerBox` equivalent).
     * @param {unknown} theme
     * @returns {{ x: number, y: number, w: number, h: number, logo: { width: number, height: number, gap: number } | null }}
     */
    function innerBox(theme) {
      const safeTheme = normalizeTheme(theme)
      const reserveLogo = Boolean(safeTheme.brandAssets.logoPath)
      return {
        x: SAFE_AREA.left,
        y: SAFE_AREA.top,
        w: CANVAS.width - SAFE_AREA.left - SAFE_AREA.right,
        h: CANVAS.height - SAFE_AREA.top - SAFE_AREA.bottom,
        logo: reserveLogo ? LOGO_RESERVE : null,
      }
    }

    /**
     * Round and clamp a box into the canvas, exactly like the Host repair pass.
     * @param {{ x: number, y: number, w: number, h: number }} box
     * @returns {{ x: number, y: number, w: number, h: number }}
     */
    function clampBox(box) {
      const w = clamp(Math.round(Number(box.w) || 0), MIN_ELEMENT_SIZE, CANVAS.width)
      const height = clamp(Math.round(Number(box.h) || 0), MIN_ELEMENT_SIZE, CANVAS.height)
      const x = clamp(Math.round(Number(box.x) || 0), 0, Math.max(0, CANVAS.width - w))
      const y = clamp(Math.round(Number(box.y) || 0), 0, Math.max(0, CANVAS.height - height))
      return { x, y, w, h: height }
    }

    /** @param {any} deck @returns {any[]} */
    function deckSlides(deck) {
      return deck !== null && typeof deck === 'object' && Array.isArray(deck.slides) ? deck.slides : []
    }

    /** @param {any} slide @returns {any[]} */
    function slideElements(slide) {
      return slide !== null && typeof slide === 'object' && Array.isArray(slide.elements) ? slide.elements : []
    }

    /** @param {unknown} list @returns {string[]} */
    function normalizeWarnings(list) {
      return Array.isArray(list) ? list.filter((item) => typeof item === 'string' && item.length > 0) : []
    }

    /**
     * Accept both documented shapes of a deck payload: the envelope `value` of
     * `GET /decks`/`PUT /decks` (`{ deck, theme, warnings }`) and the
     * `GET /state` variant (`{ deck: DeckView, theme, warnings }`).
     * @param {any} value
     * @returns {{ deck: any | null, theme: any | null, warnings: string[] }}
     */
    function unpackDeckView(value) {
      if (value === null || typeof value !== 'object') return { deck: null, theme: null, warnings: [] }
      const nested = value.deck
      if (nested !== null && typeof nested === 'object' && Array.isArray(nested.slides)) {
        return {
          deck: nested,
          theme: value.theme !== null && typeof value.theme === 'object' ? value.theme : null,
          warnings: normalizeWarnings(value.warnings),
        }
      }
      if (Array.isArray(value.slides)) {
        return { deck: value, theme: null, warnings: normalizeWarnings(value.warnings) }
      }
      return { deck: null, theme: null, warnings: [] }
    }

    /**
     * Normalize `GET /state` into the reducer's view shape.
     * @param {any} value
     */
    function normalizeStatePayload(value) {
      const payload = value !== null && typeof value === 'object' ? value : {}
      const view = unpackDeckView(payload.deck)
      const decks = Array.isArray(payload.decks)
        ? payload.decks
            .filter((item) => item !== null && typeof item === 'object' && typeof item.deckId === 'string')
            .map((item) => ({
              deckId: item.deckId,
              title: typeof item.title === 'string' && item.title.length > 0 ? item.title : item.deckId,
              templateId: typeof item.templateId === 'string' ? item.templateId : '',
              slideCount: Number.isFinite(item.slideCount) ? item.slideCount : 0,
              updatedAt: typeof item.updatedAt === 'string' ? item.updatedAt : '',
            }))
        : []
      const stateWarnings = normalizeWarnings(payload.warnings)
      return {
        decks,
        deck: view.deck,
        /* `/state` carries the active theme next to `deck`; `/decks` nests it. */
        theme: view.theme ?? (payload.theme !== null && typeof payload.theme === 'object' ? payload.theme : null),
        warnings: view.warnings.length > 0 ? view.warnings : stateWarnings,
        activeDeckId: view.deck?.deckId ?? (typeof payload.activeDeckId === 'string' ? payload.activeDeckId : null),
      }
    }

    /**
     * Empty view used when the Host reports that no deck exists yet
     * (`deck_not_found`): the panel then renders its empty state.
     * @returns {{ decks: any[], deck: null, theme: null, warnings: string[], activeDeckId: null }}
     */
    function emptyDeckView() {
      return { decks: [], deck: null, theme: null, warnings: [], activeDeckId: null }
    }

    /* ------------------------------------------------------------------ *
     * Connection RPC client.
     * ------------------------------------------------------------------ */

    /**
     * UI safety bound for one call. `AbortSignal.timeout` is optional so an old
     * engine simply runs without a local deadline (the transport still governs).
     * @param {number} timeoutMs
     * @returns {AbortSignal | undefined}
     */
    function timeoutSignal(timeoutMs) {
      if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) return undefined
      if (typeof AbortSignal === 'undefined' || typeof AbortSignal.timeout !== 'function') return undefined
      try {
        return AbortSignal.timeout(timeoutMs)
      } catch (error) {
        return undefined
      }
    }

    /**
     * One call on the plugin's same-origin Host route. Never throws: transport
     * failures and business failures both resolve an envelope.
     *
     * The Host half mounts `POST /branddeck-api/<endpoint>` on the Harness web
     * server (the same carrier the client-module registry uses for `/plugins`).
     * A bundle client half cannot reach the shipped client plugins' `connection`
     * service — it is provided on their own contexts — so a same-origin fetch is
     * the supported carrier here.
     * @param {string} endpoint endpoint such as `state` or `decks.save`
     * @param {Record<string, unknown>} [payload]
     * @param {number} [timeoutMs]
     * @returns {Promise<{ ok: true, value: any } | { ok: false, error: { code: string, message: string } }>}
     */
    async function callRpc(endpoint, payload, timeoutMs) {
      const url = `${API_BASE}/${endpoint}`
      let response
      try {
        response = await fetch(url, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify(payload === undefined ? {} : payload),
          credentials: 'same-origin',
          signal: timeoutSignal(Number.isFinite(timeoutMs) ? timeoutMs : REQUEST_TIMEOUT_MS),
        })
      } catch (error) {
        const name = error !== null && typeof error === 'object' ? String(error.name ?? '') : ''
        if (name === 'TimeoutError' || name === 'AbortError') return { ok: false, error: { code: 'timeout', message: '' } }
        const message = error !== null && typeof error === 'object' ? String(error.message ?? '') : ''
        return { ok: false, error: { code: 'transport', message } }
      }
      let envelope
      try {
        envelope = await response.json()
      } catch (error) {
        return { ok: false, error: { code: 'transport', message: `HTTP ${response.status}` } }
      }
      if (envelope !== null && envelope !== undefined && envelope.ok === true) {
        return { ok: true, value: envelope.value }
      }
      const failure =
        envelope !== null && envelope !== undefined && envelope.error !== null && typeof envelope.error === 'object'
          ? envelope.error
          : {}
      return {
        ok: false,
        error: {
          code: typeof failure.code === 'string' && failure.code.length > 0 ? failure.code : 'unknown',
          message: typeof failure.message === 'string' ? failure.message : `HTTP ${response.status}`,
        },
      }
    }

    /** Business code the Host uses for "there is no such deck yet". */
    const CODE_DECK_NOT_FOUND = 'deck_not_found'

    /**
     * Human-readable, localized error text; Host messages win when present.
     * @param {{ code: string, message: string } | null | undefined} error
     * @returns {string}
     */
    function readableError(error) {
      const code = error !== null && error !== undefined && typeof error.code === 'string' ? error.code : ''
      const message =
        error !== null && error !== undefined && typeof error.message === 'string' ? error.message.trim() : ''
      if (message.length > 0) return code.length > 0 ? `${message} (${code})` : message
      if (code === 'timeout') return t('error.timeout')
      if (code === 'transport') return t('error.transport')
      if (code === 'transport_unavailable') return t('error.transportUnavailable')
      return t('error.request', { code: code.length > 0 ? code : 'unknown' })
    }

    /* ------------------------------------------------------------------ *
     * Element / slide mutation helpers used by the reducer.
     * ------------------------------------------------------------------ */

    let idSequence = 0

    /**
     * Build an id that is unique among `taken` and within the schema's 64 chars.
     * @param {string} base
     * @param {Set<string>} taken
     * @returns {string}
     */
    function uniqueId(base, taken) {
      const root = String(base).slice(0, 56)
      let candidate = root
      while (taken.has(candidate)) {
        idSequence += 1
        candidate = `${root}-${idSequence}`
      }
      return candidate.slice(0, 64)
    }

    /**
     * Resize a box from one corner; the opposite corner stays anchored.
     * @param {{ x: number, y: number, w: number, h: number }} origin
     * @param {'nw'|'ne'|'sw'|'se'} corner
     * @param {number} dx base-pixel delta
     * @param {number} dy base-pixel delta
     */
    function resizeBox(origin, corner, dx, dy) {
      let x = origin.x
      let y = origin.y
      let w = origin.w
      let height = origin.h
      if (corner.indexOf('e') >= 0) w = origin.w + dx
      if (corner.indexOf('s') >= 0) height = origin.h + dy
      if (corner.indexOf('w') >= 0) {
        w = Math.max(MIN_ELEMENT_SIZE, origin.w - dx)
        x = origin.x + (origin.w - w)
      }
      if (corner.indexOf('n') >= 0) {
        height = Math.max(MIN_ELEMENT_SIZE, origin.h - dy)
        y = origin.y + (origin.h - height)
      }
      return clampBox({ x, y, w, h: height })
    }

    /* ------------------------------------------------------------------ *
     * Reducer. One object holds the whole panel: data, selection, edit and
     * save/export lifecycle. Every deck mutation bumps `rev` (the autosave
     * effect restarts) and marks the deck dirty.
     * ------------------------------------------------------------------ */

    /** @returns {any} */
    function createInitialState() {
      return {
        status: 'loading',
        error: null,
        decks: [],
        deck: null,
        theme: null,
        /* Rasterized template logo (`templates.logo` dataUrl), cached per template. */
        logo: null,
        activeDeckId: null,
        warnings: [],
        warningsDismissed: false,
        slideIndex: 0,
        selectionId: null,
        editingId: null,
        dirty: false,
        rev: 0,
        saveState: 'idle',
        saveError: null,
        exporting: false,
        exportResult: null,
      }
    }

    /**
     * Replace the element list of the current slide and mark the deck dirty.
     * @param {any} state
     * @param {number} slideIndex
     * @param {any[]} elements
     */
    function mutateElements(state, slideIndex, elements) {
      const slides = deckSlides(state.deck).map((slide, index) =>
        index === slideIndex ? { ...slide, elements } : slide,
      )
      return {
        ...state,
        deck: { ...state.deck, slides },
        dirty: true,
        rev: state.rev + 1,
        saveState: 'dirty',
        saveError: null,
      }
    }

    /**
     * Replace the whole slide list (reorder / duplicate / delete) and mark dirty.
     * @param {any} state
     * @param {any[]} slides
     * @param {{ slideIndex?: number, selectionId?: string | null }} [next]
     */
    function mutateSlides(state, slides, next) {
      const patch = next !== null && typeof next === 'object' ? next : {}
      return {
        ...state,
        deck: { ...state.deck, slides },
        slideIndex: typeof patch.slideIndex === 'number' ? patch.slideIndex : state.slideIndex,
        selectionId: patch.selectionId === undefined ? state.selectionId : patch.selectionId,
        editingId: null,
        dirty: true,
        rev: state.rev + 1,
        saveState: 'dirty',
        saveError: null,
      }
    }

    /**
     * Apply `updater` to the selected-slide element with this id.
     * @param {any} state
     * @param {string} id
     * @param {(element: any) => any} updater
     */
    function updateElement(state, id, updater) {
      if (state.deck === null) return state
      const slides = deckSlides(state.deck)
      const slideIndex = clamp(state.slideIndex, 0, Math.max(0, slides.length - 1))
      const slide = slides[slideIndex]
      if (slide === null || slide === undefined) return state
      const elements = slideElements(slide)
      if (!elements.some((element) => element.id === id)) return state
      return mutateElements(
        state,
        slideIndex,
        elements.map((element) => (element.id === id ? updater(element) : element)),
      )
    }

    /** @returns {number} index of the current slide, clamped to the slide list. */
    function currentSlideIndex(state) {
      const slides = deckSlides(state.deck)
      return clamp(state.slideIndex, 0, Math.max(0, slides.length - 1))
    }

    /**
     * @param {any} state
     * @param {any} action
     * @returns {any}
     */
    function reducer(state, action) {
      switch (action.type) {
        case 'load/start':
          return { ...state, status: 'loading', error: null }

        case 'view/loaded':
          return {
            ...state,
            status: 'ready',
            error: null,
            decks: Array.isArray(action.view.decks) ? action.view.decks : state.decks,
            deck: action.view.deck,
            theme: action.view.theme,
            activeDeckId: action.view.deck?.deckId ?? action.view.activeDeckId ?? state.activeDeckId,
            warnings: normalizeWarnings(action.view.warnings),
            warningsDismissed: false,
            slideIndex: 0,
            selectionId: null,
            editingId: null,
            dirty: false,
            rev: 0,
            saveState: 'idle',
            saveError: null,
          }

        case 'load/error':
          return { ...state, status: 'error', error: action.error }

        case 'deck/pending':
          return { ...state, activeDeckId: action.deckId }

        case 'slide/select':
          return {
            ...state,
            slideIndex: clamp(action.index, 0, Math.max(0, deckSlides(state.deck).length - 1)),
            selectionId: null,
            editingId: null,
          }

        case 'element/select':
          return { ...state, selectionId: action.id, editingId: null }

        case 'edit/start':
          return { ...state, selectionId: action.id, editingId: action.id }

        case 'edit/cancel':
          return { ...state, editingId: null }

        case 'edit/commit':
          return {
            ...updateElement(state, action.id, (element) => ({ ...element, content: action.content })),
            editingId: null,
          }

        case 'element/move':
          return updateElement(state, action.id, (element) => ({
            ...element,
            ...clampBox({ x: action.x, y: action.y, w: element.w, h: element.h }),
          }))

        case 'element/nudge':
          return updateElement(state, action.id, (element) => ({
            ...element,
            ...clampBox({ x: element.x + action.dx, y: element.y + action.dy, w: element.w, h: element.h }),
          }))

        case 'element/resize':
          return updateElement(state, action.id, (element) => ({
            ...element,
            ...resizeBox(
              { x: element.x, y: element.y, w: element.w, h: element.h },
              action.corner,
              action.dx,
              action.dy,
            ),
          }))

        case 'element/geometry': {
          /* Inspector numeric fields: patch one or more box values, then clamp. */
          const patch = action.patch !== null && typeof action.patch === 'object' ? action.patch : {}
          return updateElement(state, action.id, (element) => ({
            ...element,
            ...clampBox({
              x: Number.isFinite(patch.x) ? patch.x : element.x,
              y: Number.isFinite(patch.y) ? patch.y : element.y,
              w: Number.isFinite(patch.w) ? patch.w : element.w,
              h: Number.isFinite(patch.h) ? patch.h : element.h,
            }),
          }))
        }

        case 'element/style':
          return updateElement(state, action.id, (element) => ({
            ...element,
            style: styleWithPatch(element, action.patch),
          }))

        case 'element/delete': {
          if (state.deck === null) return state
          const slideIndex = currentSlideIndex(state)
          const elements = slideElements(deckSlides(state.deck)[slideIndex])
          if (!elements.some((element) => element.id === action.id)) return state
          /* The DSL requires at least one element per slide (schema min 1), so the
           * last remaining element stays: deleting it would fail Host validation. */
          if (elements.length <= 1) return { ...state, selectionId: null, editingId: null }
          return {
            ...mutateElements(
              state,
              slideIndex,
              elements.filter((element) => element.id !== action.id),
            ),
            selectionId: null,
            editingId: null,
          }
        }

        case 'element/duplicate': {
          if (state.deck === null) return state
          const slideIndex = currentSlideIndex(state)
          const elements = slideElements(deckSlides(state.deck)[slideIndex])
          if (elements.length >= MAX_ELEMENTS_PER_SLIDE) return state
          const source = elements.find((element) => element.id === action.id)
          if (source === undefined) return state
          const id = uniqueId(`${source.id}-copy`, new Set(elements.map((element) => element.id)))
          const clone = {
            ...source,
            id,
            ...clampBox({
              x: source.x + DUPLICATE_OFFSET,
              y: source.y + DUPLICATE_OFFSET,
              w: source.w,
              h: source.h,
            }),
          }
          return {
            ...mutateElements(state, slideIndex, [...elements, clone]),
            selectionId: id,
            editingId: null,
          }
        }

        case 'slide/duplicate': {
          const slides = deckSlides(state.deck)
          if (slides.length >= MAX_SLIDES) return state
          const slideIndex = currentSlideIndex(state)
          const source = slides[slideIndex]
          if (source === undefined) return state
          const taken = new Set(slides.map((slide) => slide.slideId))
          const slideId = uniqueId(`${source.slideId}-copy`, taken)
          const elementIds = new Set()
          const elements = slideElements(source).map((element) => {
            const id = uniqueId(`${element.id}-copy`, elementIds)
            elementIds.add(id)
            return { ...element, id }
          })
          const copy = { ...source, slideId, elements }
          const nextSlides = [...slides.slice(0, slideIndex + 1), copy, ...slides.slice(slideIndex + 1)]
          return mutateSlides(state, nextSlides, { slideIndex: slideIndex + 1, selectionId: null })
        }

        case 'slide/delete': {
          const slides = deckSlides(state.deck)
          if (slides.length <= 1) return state
          /* The rail deletes the slide whose trash button was pressed. */
          const target = Number.isFinite(action.index) ? action.index : currentSlideIndex(state)
          const slideIndex = clamp(target, 0, slides.length - 1)
          const nextSlides = slides.filter((slide, index) => index !== slideIndex)
          return mutateSlides(state, nextSlides, {
            slideIndex: clamp(slideIndex, 0, nextSlides.length - 1),
            selectionId: null,
          })
        }

        case 'slide/reorder': {
          const slides = deckSlides(state.deck)
          const from = action.from
          const to = action.to
          if (from < 0 || from >= slides.length || to < 0 || to >= slides.length || from === to) return state
          const nextSlides = slides.slice()
          const moved = nextSlides.splice(from, 1)[0]
          nextSlides.splice(to, 0, moved)
          return mutateSlides(state, nextSlides, { slideIndex: to, selectionId: null })
        }

        case 'save/start':
          return { ...state, saveState: 'saving', saveError: null }

        case 'save/success': {
          const view = unpackDeckView(action.view)
          const sameRevision = action.rev === state.rev
          return {
            ...state,
            /* Only trust the repaired deck when nothing changed while saving;
             * otherwise keep the newer local copy and let autosave run again. */
            deck: sameRevision && view.deck !== null ? view.deck : state.deck,
            theme: view.theme !== null ? view.theme : state.theme,
            warnings: view.warnings,
            dirty: sameRevision ? false : state.dirty,
            saveState: sameRevision ? 'saved' : 'dirty',
            saveError: null,
          }
        }

        case 'save/error':
          return { ...state, saveState: 'error', saveError: action.error }

        case 'warnings/dismiss':
          return { ...state, warningsDismissed: true }

        case 'logo/loaded':
          return { ...state, logo: { templateId: action.templateId, dataUrl: action.dataUrl } }

        case 'export/start':
          return { ...state, exporting: true, exportResult: null }

        case 'export/success':
          return {
            ...state,
            exporting: false,
            exportResult: { ok: true, path: action.path, slideCount: action.slideCount },
          }

        case 'export/error':
          return { ...state, exporting: false, exportResult: { ok: false, message: action.message } }

        case 'export/dismiss':
          return { ...state, exportResult: null }

        default:
          return state
      }
    }

    /**
     * Merge an inspector style patch into an element's style object.
     * @param {any} element
     * @param {{ fontSize?: number, color?: string, align?: 'left'|'center'|'right' }} patch
     */
    function styleWithPatch(element, patch) {
      const style = element.style !== null && typeof element.style === 'object' ? { ...element.style } : {}
      if (Number.isFinite(patch.fontSize)) style.fontSize = clamp(Math.round(patch.fontSize), 8, 120)
      if (typeof patch.color === 'string' && patch.color.length > 0) style.color = patch.color
      if (patch.align === 'left' || patch.align === 'center' || patch.align === 'right') style.align = patch.align
      return style
    }

    /* ------------------------------------------------------------------ *
     * Stylesheet. Every selector is prefixed `branddeck-`; chrome colours come
     * only from the host `--dsw-alias-*` tokens (with neutral fallbacks), the
     * 16:9 card itself is painted with the deck's own colours.
     * ------------------------------------------------------------------ */

    const CSS_TEXT = `
/* branddeck:stylesheet:start */
.branddeck-root{--branddeck-font-mono:ui-monospace,SFMono-Regular,Consolas,"Liberation Mono",Menlo,monospace;display:flex;flex-direction:column;height:100%;min-height:0;overflow:hidden;font-size:13px;line-height:1.5;color:var(--dsw-alias-label-primary,#1B1F24);background:var(--dsw-alias-bg-base,#FFFFFF);font-family:inherit}
.branddeck-toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:8px 12px;border-bottom:.5px solid var(--dsw-alias-border-l2,#E3E6EA);background:var(--dsw-alias-bg-layer-1,#F7F8FA)}
.branddeck-toolbar-title{font-size:14px;font-weight:600;margin-right:2px;white-space:nowrap}
.branddeck-toolbar-spacer{flex:1 1 auto}
.branddeck-viewtoggle{display:inline-flex;border:1px solid var(--dsw-alias-border-l2);border-radius:var(--branddeck-radius-sm,8px);overflow:hidden}
.branddeck-viewtoggle-btn{appearance:none;border:0;background:transparent;color:var(--dsw-alias-label-secondary);font:inherit;font-size:12px;line-height:1;padding:6px 10px;cursor:pointer}
.branddeck-viewtoggle-btn:hover{background:var(--dsw-alias-interactive-bg-hover)}
.branddeck-viewtoggle-on{background:var(--dsw-alias-interactive-bg-active);color:var(--dsw-alias-label-primary);font-weight:600}
.branddeck-textview{display:flex;flex-direction:column;gap:8px;flex:1 1 auto;min-height:0;padding:12px}
.branddeck-textbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.branddeck-textbar-hint{font-size:12px;color:var(--dsw-alias-label-tertiary)}
.branddeck-textarea{flex:1 1 auto;min-height:240px;width:100%;resize:none;box-sizing:border-box;padding:12px 14px;border-radius:var(--branddeck-radius-md,12px);border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);font-family:var(--branddeck-font-mono);font-size:12.5px;line-height:1.65;tab-size:2;white-space:pre-wrap}
.branddeck-textarea:focus{outline:2px solid var(--dsw-alias-button-primary-fill);outline-offset:-1px}
.branddeck-banner-ok{border-color:var(--dsw-alias-state-success-primary)}
.branddeck-gen{display:flex;flex-direction:column;gap:12px;flex:1 1 auto;min-height:0;overflow:auto;padding:14px}
.branddeck-gen-block{display:flex;flex-direction:column;gap:10px;padding:12px 14px;border:1px solid var(--dsw-alias-border-l2);border-radius:var(--branddeck-radius-md,12px);background:var(--dsw-alias-bg-layer-1)}
.branddeck-gen-head{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:600;color:var(--dsw-alias-label-primary)}
.branddeck-gen-step{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:var(--dsw-alias-button-primary-fill);color:var(--dsw-alias-label-primary-inverted);font-size:11px;font-weight:700}
.branddeck-gen-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.branddeck-gen-row .branddeck-input,.branddeck-gen-row .branddeck-select{flex:1 1 160px;min-width:0}
.branddeck-input{appearance:none;box-sizing:border-box;padding:6px 9px;border-radius:var(--branddeck-radius-sm,8px);border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font:inherit;font-size:12.5px;line-height:1.2}
.branddeck-input:focus{outline:2px solid var(--dsw-alias-button-primary-fill);outline-offset:-1px}
.branddeck-input-num{flex:0 0 96px;width:96px;text-align:right}
.branddeck-gen-hint{font-size:12px;color:var(--dsw-alias-label-tertiary)}
.branddeck-gen-spec{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.branddeck-swatches{display:inline-flex;gap:4px}
.branddeck-swatch{width:18px;height:18px;border-radius:4px;border:1px solid var(--dsw-alias-border-l2)}
.branddeck-gen-meta{font-size:12px;color:var(--dsw-alias-label-secondary)}
.branddeck-gen-form{display:flex;gap:10px;flex-wrap:wrap}
.branddeck-field-wide{flex:1 1 100%}
.branddeck-gen-details{font-size:12px;color:var(--dsw-alias-label-secondary)}
.branddeck-gen-details summary{cursor:pointer;padding:2px 0}
.branddeck-gen-code{margin:6px 0;padding:10px;border-radius:var(--branddeck-radius-sm,8px);border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-primary);font-family:var(--branddeck-font-mono);font-size:11.5px;line-height:1.6;white-space:pre-wrap;word-break:break-word}
.branddeck-gen-textarea{min-height:160px;flex:0 0 auto}
.branddeck-chat{display:flex;flex-direction:column;gap:6px;padding:10px 12px;border:1px dashed var(--dsw-alias-interactive-bg-active);border-radius:var(--branddeck-radius-sm,8px);background:var(--dsw-alias-bg-layer-2)}
.branddeck-chat-title{font-size:12px;font-weight:600;color:var(--dsw-alias-label-primary)}
.branddeck-chat-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.branddeck-chat-row .branddeck-input{flex:1 1 240px;min-width:0}
.branddeck-chat-note{font-size:12px;color:var(--dsw-alias-state-success-primary)}
.branddeck-field{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--dsw-alias-label-secondary,#5B6470);white-space:nowrap}
.branddeck-select{height:26px;max-width:220px;padding:0 6px;border-radius:6px;font-family:inherit;font-size:12px;color:var(--dsw-alias-label-primary,#1B1F24);background:var(--dsw-alias-bg-layer-2,#FFFFFF);border:.5px solid var(--dsw-alias-border-l3,#CFD4DA);outline:none}
.branddeck-select:focus-visible{border-color:var(--dsw-alias-brand-primary,#247BBF)}
.branddeck-chip{max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:2px 8px;border-radius:999px;font-size:11px;color:var(--dsw-alias-label-secondary,#5B6470);background:var(--dsw-alias-bg-module-platform,#EEF2F7);border:.5px solid var(--dsw-alias-border-l3,#CFD4DA)}
.branddeck-save{font-size:12px;font-variant-numeric:tabular-nums;white-space:nowrap}
.branddeck-save-idle{color:var(--dsw-alias-label-tertiary,#8A929C)}
.branddeck-save-saving{color:var(--dsw-alias-label-secondary,#5B6470)}
.branddeck-save-saved{color:var(--dsw-alias-state-success-primary,#1A7F4B)}
.branddeck-save-dirty{color:var(--dsw-alias-state-warn-label,#8A5B00)}
.branddeck-save-error{color:var(--dsw-alias-state-error-primary,#C0281F)}
.branddeck-btn{height:28px;padding:0 10px;border-radius:6px;font-family:inherit;font-size:12px;cursor:pointer;color:var(--dsw-alias-label-primary,#1B1F24);background:var(--dsw-alias-button-tool-bar-fill,transparent);border:.5px solid var(--dsw-alias-border-l3,#CFD4DA)}
.branddeck-btn:hover:not(:disabled){background:var(--dsw-alias-interactive-bg-hover,#EEF1F5)}
.branddeck-btn:disabled{opacity:.45;cursor:default}
.branddeck-btn-primary{color:var(--dsw-alias-label-primary-inverted,#FFFFFF);background:var(--dsw-alias-button-primary-fill,#247BBF);border-color:transparent}
.branddeck-btn-primary:hover:not(:disabled){background:var(--dsw-alias-button-primary-hover,#1C639C)}
.branddeck-banner{display:flex;align-items:flex-start;gap:8px;padding:8px 12px;font-size:12px;border-bottom:.5px solid transparent}
.branddeck-banner-error{color:var(--dsw-alias-state-error-primary,#C0281F);background:var(--dsw-alias-state-error-secondary,#FDECEC);border-bottom-color:var(--dsw-alias-state-error-primary,#E5484D)}
.branddeck-banner-warn{color:var(--dsw-alias-state-warn-label,#8A5B00);background:var(--dsw-alias-state-warn-tertiary,#FFF6E5);border-bottom-color:var(--dsw-alias-state-warn-secondary,#F2C56B)}
.branddeck-banner-success{color:var(--dsw-alias-state-success-primary,#1A7F4B);background:var(--dsw-alias-state-success-secondary,#E8F7EE);border-bottom-color:var(--dsw-alias-state-success-primary,#1A7F4B)}
.branddeck-banner-text{flex:1 1 auto;min-width:0;white-space:pre-wrap;overflow-wrap:break-word}
.branddeck-banner-title{font-weight:600}
.branddeck-banner-list{margin:0;padding-left:16px;display:flex;flex-direction:column;gap:2px;overflow-wrap:break-word}
.branddeck-body{display:flex;flex:1 1 auto;min-height:0}
.branddeck-rail{flex:0 0 172px;display:flex;flex-direction:column;gap:9px;padding:10px;overflow-y:auto;border-right:.5px solid var(--dsw-alias-border-l2,#E3E6EA);background:var(--dsw-alias-bg-layer-1,#F7F8FA)}
.branddeck-rail-title{font-size:11px;font-weight:600;letter-spacing:.4px;text-transform:uppercase;color:var(--dsw-alias-label-tertiary,#8A929C)}
.branddeck-rail-item{display:flex;flex-direction:column;gap:4px}
.branddeck-rail-footer{margin-top:auto;display:flex;flex-direction:column;gap:6px;padding-top:6px}
.branddeck-thumb{position:relative;width:144px;height:81px;padding:0;border-radius:5px;overflow:hidden;cursor:pointer;background:#FFFFFF;border:1px solid var(--dsw-alias-border-l3,#CFD4DA);outline:none}
.branddeck-thumb-selected{border-color:var(--dsw-alias-brand-primary,#247BBF);box-shadow:0 0 0 1px var(--dsw-alias-brand-primary,#247BBF)}
.branddeck-thumb-inner{position:absolute;left:0;top:0;width:1280px;height:720px;transform-origin:0 0;pointer-events:none;overflow:hidden}
.branddeck-thumb-index{position:absolute;right:2px;bottom:2px;padding:0 4px;border-radius:3px;font-size:10px;line-height:14px;color:#FFFFFF;background:rgba(0,0,0,.55)}
.branddeck-rail-actions{display:flex;gap:4px}
.branddeck-btn-icon{height:22px;min-width:22px;padding:0 5px;border-radius:5px;font-family:inherit;font-size:11px;line-height:20px;cursor:pointer;color:var(--dsw-alias-label-secondary,#5B6470);background:var(--dsw-alias-button-tool-bar-fill-invisible,transparent);border:.5px solid var(--dsw-alias-border-l3,#CFD4DA)}
.branddeck-btn-icon:hover:not(:disabled){background:var(--dsw-alias-interactive-bg-hover,#EEF1F5)}
.branddeck-btn-icon:disabled{opacity:.4;cursor:default}
.branddeck-stage{flex:1 1 auto;min-width:0;overflow:auto;padding:16px;background:var(--dsw-alias-bg-layer-2,#F1F3F6);display:flex;align-items:flex-start;justify-content:center}
.branddeck-stage-inner{width:100%;min-width:0;display:flex;align-items:flex-start;justify-content:center}
.branddeck-frame{position:relative;flex:0 0 auto;overflow:hidden;background:#FFFFFF;box-shadow:0 2px 14px rgba(0,0,0,.16)}
.branddeck-canvas{position:absolute;left:0;top:0;transform-origin:0 0;overflow:hidden;outline:none;user-select:none;touch-action:none}
.branddeck-el{position:absolute;box-sizing:border-box}
.branddeck-el-selected{outline:2px solid var(--dsw-alias-brand-primary,#247BBF);outline-offset:0}
.branddeck-el-overflow{outline:2px solid var(--dsw-alias-state-error-primary,#E5484D)}
.branddeck-handle{position:absolute;box-sizing:border-box;z-index:3;background:var(--dsw-alias-brand-primary,#247BBF);border:1px solid #FFFFFF;border-radius:2px}
.branddeck-overflow-badge{position:absolute;right:0;top:0;z-index:3;border-radius:3px;padding:1px 4px;font-family:inherit;white-space:nowrap;pointer-events:none;color:#FFFFFF;background:var(--dsw-alias-state-error-primary,#E5484D)}
.branddeck-el-text{box-sizing:border-box;overflow:hidden}
.branddeck-el-svg{width:100%;height:100%;overflow:hidden}
.branddeck-el-svg svg{display:block;width:100%;height:100%}
.branddeck-el-chart{width:100%;height:100%;overflow:hidden}
.branddeck-placeholder{display:flex;align-items:center;justify-content:center;box-sizing:border-box;width:100%;height:100%;padding:4px;overflow:hidden;text-align:center;font-size:12px;border:1px dashed currentColor;opacity:.6}
.branddeck-editor{box-sizing:border-box;width:100%;height:100%;padding:0;overflow:hidden;white-space:pre-wrap;overflow-wrap:break-word;cursor:text;user-select:text;background:rgba(255,255,255,.96);outline:2px solid var(--dsw-alias-brand-primary,#247BBF);outline-offset:1px}
.branddeck-inspector{flex:0 0 236px;display:flex;flex-direction:column;gap:12px;padding:12px;overflow-y:auto;border-left:.5px solid var(--dsw-alias-border-l2,#E3E6EA);background:var(--dsw-alias-bg-layer-1,#F7F8FA)}
.branddeck-section{display:flex;flex-direction:column;gap:6px}
.branddeck-section-title{font-size:11px;font-weight:600;letter-spacing:.4px;text-transform:uppercase;color:var(--dsw-alias-label-tertiary,#8A929C)}
.branddeck-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
.branddeck-field-row{display:flex;align-items:center;gap:6px}
.branddeck-field-label{flex:0 0 22px;font-size:11px;color:var(--dsw-alias-label-tertiary,#8A929C)}
.branddeck-num{box-sizing:border-box;width:100%;height:26px;padding:0 6px;border-radius:6px;font-family:inherit;font-size:12px;font-variant-numeric:tabular-nums;color:var(--dsw-alias-label-primary,#1B1F24);background:var(--dsw-alias-bg-layer-2,#FFFFFF);border:.5px solid var(--dsw-alias-border-l3,#CFD4DA);outline:none}
.branddeck-num:focus-visible{border-color:var(--dsw-alias-brand-primary,#247BBF)}
.branddeck-swatches{display:flex;flex-wrap:wrap;gap:6px}
.branddeck-swatch{width:20px;height:20px;padding:0;border-radius:4px;cursor:pointer;border:.5px solid var(--dsw-alias-border-l3,#CFD4DA)}
.branddeck-swatch-active{box-shadow:0 0 0 2px var(--dsw-alias-brand-primary,#247BBF)}
.branddeck-actions{display:flex;flex-wrap:wrap;gap:6px}
.branddeck-hint{font-size:11px;line-height:16px;color:var(--dsw-alias-label-caption,#8A929C);overflow-wrap:break-word}
.branddeck-empty{flex:1 1 auto;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;padding:40px 24px;text-align:center;color:var(--dsw-alias-label-secondary,#5B6470)}
.branddeck-empty-title{max-width:520px;font-size:14px;font-weight:600;color:var(--dsw-alias-label-primary,#1B1F24)}
.branddeck-empty-hint{max-width:520px;font-size:12px;color:var(--dsw-alias-label-tertiary,#8A929C)}
.branddeck-loading{display:flex;align-items:center;gap:8px;padding:24px;color:var(--dsw-alias-label-secondary,#5B6470)}
.branddeck-modal-mask{position:fixed;inset:0;z-index:60;display:flex;align-items:center;justify-content:center;background:var(--dsw-alias-bg-mask-1,rgba(15,23,32,.35))}
.branddeck-modal{display:flex;flex-direction:column;gap:12px;min-width:300px;max-width:520px;padding:18px 20px;border-radius:10px;color:var(--dsw-alias-label-primary,#1B1F24);background:var(--dsw-alias-bg-overlay,#FFFFFF);border:.5px solid var(--dsw-alias-border-l3,#CFD4DA);box-shadow:0 12px 32px rgba(0,0,0,.24)}
.branddeck-modal-title{font-size:14px;font-weight:600}
.branddeck-spinner{flex:0 0 auto;width:16px;height:16px;border-radius:50%;border:2px solid var(--dsw-alias-border-l3,#CFD4DA);border-top-color:var(--dsw-alias-brand-primary,#247BBF);animation:branddeck-spin .9s linear infinite}
@keyframes branddeck-spin{to{transform:rotate(360deg)}}
/* branddeck:stylesheet:end */
`

    /**
     * Append the panel stylesheet once for the lifetime of the panel body.
     */
    function useStylesheet() {
      React.useEffect(() => {
        if (typeof document === 'undefined' || document.head === null || document.head === undefined) {
          return undefined
        }
        const style = document.createElement('style')
        style.setAttribute('data-branddeck', 'panel-stylesheet')
        style.textContent = CSS_TEXT
        document.head.appendChild(style)
        return () => {
          if (style.parentNode !== null) style.parentNode.removeChild(style)
        }
      }, [])
    }

    /* ------------------------------------------------------------------ *
     * Render helpers (pure functions returning element trees).
     * ------------------------------------------------------------------ */

    /**
     * Font stack for a text role, taken from the deck theme.
     * @param {any} element
     * @param {typeof FALLBACK_THEME} theme
     */
    function fontStack(element, theme) {
      const isHeading = element.role === 'title' || element.role === 'subtitle'
      const primary = isHeading ? theme.fonts.title : theme.fonts.body
      return `"${primary}", "${theme.fonts.latin}", sans-serif`
    }

    /**
     * Pick a foreground that stays legible on the white inline-edit backdrop.
     * @param {string} color resolved `#RRGGBB`
     */
    function readableOnWhite(color) {
      const hex = normalizeHex(color || '#111111').slice(1)
      if (hex.length < 6) return '#111827'
      const r = parseInt(hex.slice(0, 2), 16)
      const g = parseInt(hex.slice(2, 4), 16)
      const b = parseInt(hex.slice(4, 6), 16)
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
      return luminance > 0.72 ? '#111827' : color
    }

    /**
     * Strip anything executable from model-authored svg before it is displayed.
     * The Host already sanitizes; this is defence in depth for the DOM.
     * @param {string} code
     * @returns {string} empty string when there is nothing safe to draw
     */
    function sanitizeSvgForDisplay(code) {
      let out = typeof code === 'string' ? code.trim() : ''
      if (out.length === 0) return ''
      if (!out.startsWith('<svg')) {
        const start = out.indexOf('<svg')
        if (start === -1) return ''
        out = out.slice(start)
      }
      out = out
        .replace(/<\s*script[\s\S]*?<\s*\/\s*script\s*>/gi, '')
        .replace(/\son[a-zA-Z]+\s*=\s*("[^"]*"|'[^']*')/g, '')
        .replace(/javascript:/gi, '')
      return out.startsWith('<svg') ? out : ''
    }

    /** @param {number} cx @param {number} cy @param {number} radius */
    function arcPath(cx, cy, radius, startAngle, endAngle) {
      const start = { x: cx + radius * Math.cos(startAngle), y: cy + radius * Math.sin(startAngle) }
      const end = { x: cx + radius * Math.cos(endAngle), y: cy + radius * Math.sin(endAngle) }
      const largeArc = endAngle - startAngle > Math.PI ? 1 : 0
      const format = (value) => value.toFixed(2)
      return [
        `M ${format(cx)} ${format(cy)}`,
        `L ${format(start.x)} ${format(start.y)}`,
        `A ${format(radius)} ${format(radius)} 0 ${largeArc} 1 ${format(end.x)} ${format(end.y)}`,
        'Z',
      ].join(' ')
    }

    /**
     * Chart preview for the `chart` element: bar, line and pie, drawn from the
     * deck data with the theme's chart palette so the canvas stays brand-accurate.
     * @param {any} element
     * @param {typeof FALLBACK_THEME} theme
     */
    function renderChartContent(element, theme) {
      const width = Math.max(1, element.w)
      const height = Math.max(1, element.h)
      const palette = theme.colors.chartPalette
      const colorAt = (index) => palette[index % palette.length]
      const labels = element.data !== null && element.data !== undefined && Array.isArray(element.data.labels)
        ? element.data.labels
        : []
      const series = element.data !== null && element.data !== undefined && Array.isArray(element.data.series)
        ? element.data.series
        : []
      const viewBox = `0 0 ${width} ${height}`
      const pad = Math.min(width, height) * 0.08

      if (element.chartType === 'pie') {
        const first = series[0]
        const values = first !== undefined && Array.isArray(first.values) ? first.values.map((value) => Math.max(0, Number(value) || 0)) : []
        const total = values.reduce((sum, value) => sum + value, 0)
        if (total <= 0) return h('div', { className: 'branddeck-placeholder' }, '')
        const radius = Math.max(1, Math.min(width, height) / 2 - pad)
        const cx = width / 2
        const cy = height / 2
        if (values.length <= 1) {
          return h(
            'svg',
            { viewBox, width: '100%', height: '100%', preserveAspectRatio: 'none' },
            h('circle', { key: 'pie-single', cx, cy, r: radius, fill: colorAt(0) }),
          )
        }
        let angle = -Math.PI / 2
        const slices = values.map((value, index) => {
          const sweep = (value / total) * Math.PI * 2
          const path = arcPath(cx, cy, radius, angle, angle + sweep)
          angle += sweep
          return h('path', { key: `pie-slice-${index}`, d: path, fill: colorAt(index) })
        })
        return h('svg', { viewBox, width: '100%', height: '100%', preserveAspectRatio: 'none' }, slices)
      }

      const allValues = series.reduce((values, one) => {
        const list = one !== null && one !== undefined && Array.isArray(one.values) ? one.values : []
        return values.concat(list.filter((value) => Number.isFinite(value)))
      }, [])
      const max = allValues.length > 0 ? Math.max(1, ...allValues) : 1
      const plotWidth = Math.max(1, width - pad * 2)
      const plotHeight = Math.max(1, height - pad * 2)
      const groupWidth = plotWidth / Math.max(1, labels.length)
      const baseline = h('line', {
        key: 'chart-baseline',
        x1: pad,
        y1: height - pad,
        x2: width - pad,
        y2: height - pad,
        stroke: theme.colors.textSecondary,
        strokeWidth: 1,
        opacity: 0.35,
      })

      if (element.chartType === 'line') {
        const seriesCount = Math.max(1, series.length)
        const lines = series.map((one, seriesIndex) => {
          const values = Array.isArray(one.values) ? one.values : []
          const points = values.map((value, index) => {
            const x = pad + groupWidth * (index + 0.5)
            const y = height - pad - (Math.max(0, Number(value) || 0) / max) * plotHeight
            return `${x.toFixed(2)},${y.toFixed(2)}`
          })
          return h('polyline', {
            key: `chart-line-${seriesIndex}`,
            points: points.join(' '),
            fill: 'none',
            stroke: colorAt(seriesIndex),
            strokeWidth: Math.max(1.5, Math.min(4, plotHeight / 60)),
            strokeLinejoin: 'round',
            strokeLinecap: 'round',
            opacity: seriesCount > 0 ? 1 : 0.8,
          })
        })
        return h('svg', { viewBox, width: '100%', height: '100%', preserveAspectRatio: 'none' }, [baseline, ...lines])
      }

      const seriesCount = Math.max(1, series.length)
      const barSlot = groupWidth / (seriesCount + 0.6)
      const bars = []
      series.forEach((one, seriesIndex) => {
        const values = Array.isArray(one.values) ? one.values : []
        values.forEach((value, index) => {
          const barHeight = (Math.max(0, Number(value) || 0) / max) * plotHeight
          bars.push(
            h('rect', {
              key: `chart-bar-${seriesIndex}-${index}`,
              x: pad + groupWidth * index + barSlot * (seriesIndex + 0.3),
              y: height - pad - barHeight,
              width: Math.max(1, barSlot * 0.8),
              height: Math.max(0.5, barHeight),
              fill: colorAt(seriesIndex),
              rx: Math.min(3, barSlot * 0.2),
            }),
          )
        })
      })
      return h('svg', { viewBox, width: '100%', height: '100%', preserveAspectRatio: 'none' }, [baseline, ...bars])
    }

    /**
     * Draw one element's own content (no selection chrome).
     * @param {any} element
     * @param {typeof FALLBACK_THEME} theme
     */
    function renderElementContent(element, theme) {
      if (element === null || element === undefined) return null

      if (element.type === 'text') {
        const style = element.style !== null && typeof element.style === 'object' ? element.style : {}
        const align = style.align ?? 'left'
        const valign = style.valign ?? 'top'
        return h(
          'div',
          {
            key: 'text',
            className: 'branddeck-el-text',
            style: {
              width: '100%',
              height: '100%',
              display: 'flex',
              flexDirection: 'column',
              justifyContent: valign === 'middle' ? 'center' : valign === 'bottom' ? 'flex-end' : 'flex-start',
              alignItems: align === 'center' ? 'center' : align === 'right' ? 'flex-end' : 'flex-start',
              textAlign: align,
              fontSize: style.fontSize ?? 20,
              fontWeight: style.bold === true ? 700 : 400,
              fontStyle: style.italic === true ? 'italic' : 'normal',
              lineHeight: style.lineSpacing ?? 1.35,
              color: resolveColor(style.color, theme, theme.colors.textPrimary),
              fontFamily: fontStack(element, theme),
              whiteSpace: 'pre-wrap',
              overflowWrap: 'break-word',
            },
          },
          String(element.content ?? ''),
        )
      }

      if (element.type === 'shape') {
        const style = element.style !== null && typeof element.style === 'object' ? element.style : {}
        const fill = style.fill === undefined ? 'transparent' : resolveColor(style.fill, theme, 'transparent')
        const stroke = style.stroke === undefined ? 'transparent' : resolveColor(style.stroke, theme, 'transparent')
        const strokeWidth = Number.isFinite(style.strokeWidth) ? style.strokeWidth : 0
        if (element.shapeType === 'line') {
          const lineColor = stroke !== 'transparent' ? stroke : fill !== 'transparent' ? fill : theme.colors.primary
          return h(
            'svg',
            {
              key: 'shape',
              viewBox: `0 0 ${Math.max(1, element.w)} ${Math.max(1, element.h)}`,
              width: '100%',
              height: '100%',
              preserveAspectRatio: 'none',
            },
            h('line', {
              x1: 0,
              y1: Math.max(1, element.h) / 2,
              x2: Math.max(1, element.w),
              y2: Math.max(1, element.h) / 2,
              stroke: lineColor,
              strokeWidth: strokeWidth > 0 ? strokeWidth : 2,
            }),
          )
        }
        if (element.shapeType === 'rightArrow') {
          return h('div', {
            key: 'shape',
            className: 'branddeck-el-shape',
            style: {
              width: '100%',
              height: '100%',
              background: fill !== 'transparent' ? fill : theme.colors.primary,
              opacity: style.opacity ?? 1,
              clipPath: 'polygon(0 28%, 58% 28%, 58% 0, 100% 50%, 58% 100%, 58% 72%, 0 72%)',
            },
          })
        }
        const isEllipse = element.shapeType === 'ellipse'
        const radius = isEllipse
          ? '50%'
          : element.shapeType === 'roundRect'
            ? Math.round(style.radius ?? Math.min(element.w, element.h) * 0.12)
            : 0
        return h('div', {
          key: 'shape',
          className: 'branddeck-el-shape',
          style: {
            width: '100%',
            height: '100%',
            boxSizing: 'border-box',
            background: fill,
            border: strokeWidth > 0 ? `${strokeWidth}px solid ${stroke}` : stroke === 'transparent' ? 'none' : `1px solid ${stroke}`,
            borderRadius: typeof radius === 'number' ? `${radius}px` : radius,
            opacity: style.opacity ?? 1,
          },
        })
      }

      if (element.type === 'svg') {
        const code = sanitizeSvgForDisplay(element.svgCode)
        if (code === '') {
          return h('div', { key: 'svg', className: 'branddeck-placeholder' }, element.description ?? 'SVG')
        }
        return h('div', {
          key: 'svg',
          className: 'branddeck-el-svg',
          dangerouslySetInnerHTML: { __html: code },
        })
      }

      if (element.type === 'chart') {
        return h('div', { key: 'chart', className: 'branddeck-el-chart' }, renderChartContent(element, theme))
      }

      if (element.type === 'image') {
        /* Icons are theme asset keys; the template logo is injected by the panel
         * as a `data:` URL under the `logo` key once `templates.logo` answered. */
        const icons = theme.brandAssets.icons
        const key = typeof element.assetKey === 'string' ? element.assetKey : ''
        const candidate = typeof icons[key] === 'string' ? icons[key] : null
        const usable = candidate !== null && /^(https?:|data:|\/)/.test(candidate)
        if (!usable) {
          return h('div', { key: 'image', className: 'branddeck-placeholder' }, t('image.missing', { key }))
        }
        return h('img', {
          key: 'image',
          src: candidate,
          alt: element.alt ?? key,
          draggable: false,
          style: { width: '100%', height: '100%', objectFit: 'contain', display: 'block' },
        })
      }

      return h('div', { key: 'unknown', className: 'branddeck-placeholder' }, String(element.type ?? ''))
    }

    /** Corner handles, in base-pixel space scaled for constant screen size. */
    const HANDLE_CORNERS = Object.freeze([
      { id: 'nw', x: 0, y: 0, cursor: 'nwse-resize' },
      { id: 'ne', x: 1, y: 0, cursor: 'nesw-resize' },
      { id: 'sw', x: 0, y: 1, cursor: 'nesw-resize' },
      { id: 'se', x: 1, y: 1, cursor: 'nwse-resize' },
    ])

    /**
     * Inline text editor: an uncontrolled `contentEditable` seeded with the
     * current copy. Enter commits, Shift+Enter inserts a newline, Escape cancels,
     * blur commits (the design doc's 双击直接改字).
     */
    function TextEditor(props) {
      const nodeRef = React.useRef(null)
      const settledRef = React.useRef(false)
      const element = props.element
      const theme = props.theme
      const style = element.style !== null && typeof element.style === 'object' ? element.style : {}

      React.useEffect(() => {
        const node = nodeRef.current
        if (node === null || node === undefined) return
        node.textContent = String(element.content ?? '')
        if (typeof node.focus === 'function') node.focus()
        if (typeof window !== 'undefined' && window.getSelection !== undefined) {
          try {
            const range = document.createRange()
            range.selectNodeContents(node)
            const selection = window.getSelection()
            if (selection !== null && selection !== undefined) {
              selection.removeAllRanges()
              selection.addRange(range)
            }
          } catch (error) {
            /* selection is a nicety; editing still works without it */
          }
        }
      }, [])

      const settle = (commit) => {
        if (settledRef.current) return
        settledRef.current = true
        if (commit) {
          const node = nodeRef.current
          const text = node === null || node === undefined ? String(element.content ?? '') : String(node.innerText ?? node.textContent ?? '')
          props.onCommit(text.replace(/\r\n?/g, '\n'))
        } else {
          props.onCancel()
        }
      }

      return h('div', {
        ref: nodeRef,
        className: 'branddeck-editor',
        contentEditable: true,
        suppressContentEditableWarning: true,
        spellCheck: false,
        style: {
          fontSize: style.fontSize ?? 20,
          fontWeight: style.bold === true ? 700 : 400,
          fontStyle: style.italic === true ? 'italic' : 'normal',
          lineHeight: style.lineSpacing ?? 1.35,
          textAlign: style.align ?? 'left',
          fontFamily: fontStack(element, theme),
          color: readableOnWhite(resolveColor(style.color, theme, theme.colors.textPrimary)),
        },
        onPointerDown: (event) => event.stopPropagation(),
        onDoubleClick: (event) => event.stopPropagation(),
        onBlur: () => settle(true),
        onKeyDown: (event) => {
          if (event.key === 'Enter' && event.shiftKey !== true) {
            event.preventDefault()
            settle(true)
            return
          }
          if (event.key === 'Escape') {
            event.preventDefault()
            event.stopPropagation()
            settle(false)
          }
        },
      })
    }

    /**
     * One positioned element on the canvas (or a read-only thumbnail copy).
     */
    function ElementView(props) {
      const element = props.element
      const theme = props.theme
      const selected = props.selected === true
      const overflows = props.overflows === true
      const scale = Number.isFinite(props.scale) && props.scale > 0 ? props.scale : 1
      const className = ['branddeck-el', selected ? 'branddeck-el-selected' : '', overflows ? 'branddeck-el-overflow' : '']
      const boxStyle = { left: element.x, top: element.y, width: element.w, height: element.h }

      if (props.editing === true) {
        return h(
          'div',
          { className: className.filter(Boolean).join(' '), style: boxStyle, 'data-branddeck-element': element.id },
          h(TextEditor, {
            element,
            theme,
            onCommit: (content) => props.onEditCommit(element.id, content),
            onCancel: () => props.onEditCancel(),
          }),
        )
      }

      const children = [renderElementContent(element, theme)]

      if (selected) {
        const size = HANDLE_SIZE / scale
        for (const corner of HANDLE_CORNERS) {
          children.push(
            h('div', {
              key: `handle-${corner.id}`,
              className: 'branddeck-handle',
              style: {
                width: size,
                height: size,
                cursor: corner.cursor,
                left: corner.x === 0 ? -size / 2 : undefined,
                right: corner.x === 1 ? -size / 2 : undefined,
                top: corner.y === 0 ? -size / 2 : undefined,
                bottom: corner.y === 1 ? -size / 2 : undefined,
              },
              onPointerDown: (event) => props.onResizeStart(event, corner.id),
              onPointerMove: props.onPointerMove,
              onPointerUp: props.onPointerUp,
              onPointerCancel: props.onPointerUp,
              onLostPointerCapture: props.onPointerUp,
            }),
          )
        }
      }

      if (overflows) {
        children.push(
          h(
            'div',
            {
              key: 'overflow-badge',
              className: 'branddeck-overflow-badge',
              style: { fontSize: Math.max(9, 11 / scale), padding: `${Math.max(1, 2 / scale)}px ${Math.max(2, 4 / scale)}px` },
            },
            t('overflow.badge', { lines: props.neededLines ?? 0 }),
          ),
        )
      }

      return h(
        'div',
        {
          className: className.filter(Boolean).join(' '),
          style: boxStyle,
          'data-branddeck-element': element.id,
          onPointerDown: (event) => props.onDragStart(event, element),
          onPointerMove: props.onPointerMove,
          onPointerUp: props.onPointerUp,
          onPointerCancel: props.onPointerUp,
          onLostPointerCapture: props.onPointerUp,
          onDoubleClick: (event) => props.onDoubleClick(event, element),
        },
        children,
      )
    }

    /**
     * Numeric inspector field: local text while typing, commit on Enter/blur.
     */
    function NumberField(props) {
      const [text, setText] = React.useState(String(Math.round(props.value)))
      React.useEffect(() => {
        setText(String(Math.round(props.value)))
      }, [props.value])

      const commit = () => {
        const parsed = Number(text)
        if (Number.isFinite(parsed)) props.onCommit(parsed)
        else setText(String(Math.round(props.value)))
      }

      return h(
        'label',
        { className: 'branddeck-field-row' },
        h('span', { className: 'branddeck-field-label' }, props.label),
        h('input', {
          className: 'branddeck-num',
          type: 'number',
          step: 1,
          min: props.min,
          max: props.max,
          value: text,
          'aria-label': props.ariaLabel ?? props.label,
          onChange: (event) => setText(event.target.value),
          onBlur: commit,
          onKeyDown: (event) => {
            if (event.key === 'Enter') {
              event.preventDefault()
              commit()
              return
            }
            if (event.key === 'Escape') {
              event.preventDefault()
              setText(String(Math.round(props.value)))
            }
          },
        }),
      )
    }

    /** Text-colour swatches: theme token references plus their resolved colour. */
    const TEXT_COLOR_TOKENS = Object.freeze([
      '$colors.textPrimary',
      '$colors.textSecondary',
      '$colors.primary',
      '$colors.accent',
      '$colors.chart1',
      '$colors.chart2',
    ])

    /**
     * Right-hand inspector for the selected element: geometry, typography,
     * alignment, colour and the element actions.
     */
    function Inspector(props) {
      const element = props.element
      const theme = props.theme
      if (element === null || element === undefined) {
        return h(
          'div',
          { className: 'branddeck-inspector' },
          h('div', { className: 'branddeck-section-title' }, t('label.noSelection')),
          h('div', { className: 'branddeck-hint' }, t('label.selectHint')),
        )
      }

      const isText = element.type === 'text'
      const style = element.style !== null && typeof element.style === 'object' ? element.style : {}
      const align = style.align ?? 'left'
      const currentColor = typeof style.color === 'string' ? style.color : ''
      const currentResolved = currentColor.length > 0 ? resolveColor(currentColor, theme, '') : ''
      /* The design doc's 64px safe area (plus the logo corner) is advisory: the
       * deck may use the full canvas, but bleeding off the margin is a print risk
       * worth flagging while the element is selected. */
      const safe = innerBox(theme)
      const outsideSafe =
        element.x < safe.x ||
        element.y < safe.y ||
        element.x + element.w > safe.x + safe.w ||
        element.y + element.h > safe.y + safe.h
      const geometry = [
        { key: 'x', label: t('label.x'), value: element.x, min: 0, max: CANVAS.width },
        { key: 'y', label: t('label.y'), value: element.y, min: 0, max: CANVAS.height },
        { key: 'w', label: t('label.w'), value: element.w, min: MIN_ELEMENT_SIZE, max: CANVAS.width },
        { key: 'h', label: t('label.h'), value: element.h, min: MIN_ELEMENT_SIZE, max: CANVAS.height },
      ]

      const sections = [
        h(
          'div',
          { className: 'branddeck-section', key: 'geometry' },
          h('div', { className: 'branddeck-section-title' }, t('label.position')),
          h(
            'div',
            { className: 'branddeck-grid' },
            geometry.map((field) =>
              h(NumberField, {
                key: field.key,
                label: field.label,
                ariaLabel: `${t('label.position')} ${field.label}`,
                value: field.value,
                min: field.min,
                max: field.max,
                onCommit: (value) => props.onGeometry(field.key, value),
              }),
            ),
          ),
          outsideSafe
            ? h('div', { className: 'branddeck-hint' }, t('hint.safeArea', { margin: SAFE_AREA.left }))
            : null,
        ),
      ]

      if (isText) {
        sections.push(
          h(
            'div',
            { className: 'branddeck-section', key: 'typography' },
            h('div', { className: 'branddeck-section-title' }, t('label.typography')),
            h(NumberField, {
              label: t('label.fontSize'),
              value: style.fontSize ?? 20,
              min: 8,
              max: 120,
              onCommit: (value) => props.onStyle({ fontSize: value }),
            }),
            h(
              'div',
              { className: 'branddeck-actions' },
              ['left', 'center', 'right'].map((value) =>
                h(
                  'button',
                  {
                    key: `align-${value}`,
                    type: 'button',
                    className: `branddeck-btn${align === value ? ' branddeck-btn-primary' : ''}`,
                    title: `${t('label.align')} ${t(`align.${value}`)}`,
                    onClick: () => props.onStyle({ align: value }),
                  },
                  t(`align.${value}`),
                ),
              ),
            ),
            h(
              'div',
              { className: 'branddeck-swatches' },
              TEXT_COLOR_TOKENS.map((token) =>
                h('button', {
                  key: `swatch-${token}`,
                  type: 'button',
                  className: `branddeck-swatch${
                    currentColor === token || (currentResolved !== '' && currentResolved === resolveColor(token, theme, ''))
                      ? ' branddeck-swatch-active'
                      : ''
                  }`,
                  title: `${t('label.color')} ${token}`,
                  'aria-label': `${t('label.color')} ${token}`,
                  style: { background: resolveColor(token, theme, '#000000') },
                  onClick: () => props.onStyle({ color: token }),
                }),
              ),
            ),
          ),
        )
      }

      sections.push(
        h(
          'div',
          { className: 'branddeck-section', key: 'actions' },
          h('div', { className: 'branddeck-section-title' }, t('label.slideIndex', { index: props.slideNumber, total: props.slideTotal })),
          h(
            'div',
            { className: 'branddeck-actions' },
            h(
              'button',
              {
                type: 'button',
                className: 'branddeck-btn',
                onClick: () => props.onDuplicateElement(element.id),
              },
              t('action.duplicateElement'),
            ),
            h(
              'button',
              {
                type: 'button',
                className: 'branddeck-btn',
                /* The DSL keeps at least one element per slide, so explain the
                 * disabled state instead of silently ignoring the click. */
                disabled: props.canDelete === false,
                title: props.canDelete === false ? t('hint.lastElement') : undefined,
                onClick: () => props.onDeleteElement(element.id),
              },
              t('action.deleteElement'),
            ),
          ),
          h('div', { className: 'branddeck-hint' }, t('hint.keys')),
        ),
      )

      return h('div', { className: 'branddeck-inspector' }, sections)
    }

    /**
     * Left thumbnail rail: the same 16:9 canvas rendered at thumbnail scale.
     */
    function SlideRail(props) {
      const slides = props.slides
      const theme = props.theme
      const selectedIndex = props.selectedIndex

      const items = slides.map((slide, index) =>
        h(
          'div',
          { className: 'branddeck-rail-item', key: slide.slideId ?? `slide-${index}` },
          h(
            'button',
            {
              type: 'button',
              className: `branddeck-thumb${index === selectedIndex ? ' branddeck-thumb-selected' : ''}`,
              title: t('thumb.label', { index: index + 1 }),
              'aria-label': t('thumb.label', { index: index + 1 }),
              onClick: () => props.onSelect(index),
            },
            h(
              'div',
              {
                className: 'branddeck-thumb-inner',
                style: {
                  transform: `scale(${THUMB_SCALE})`,
                  background: resolveColor(slide.background, theme, theme.colors.background),
                },
              },
              slideElements(slide).map((element) =>
                h(
                  'div',
                  {
                    key: element.id,
                    className: 'branddeck-el',
                    style: { left: element.x, top: element.y, width: element.w, height: element.h },
                  },
                  renderElementContent(element, theme),
                ),
              ),
            ),
            h('span', { className: 'branddeck-thumb-index' }, String(index + 1)),
          ),
          h(
            'div',
            { className: 'branddeck-rail-actions' },
            h(
              'button',
              {
                type: 'button',
                className: 'branddeck-btn-icon',
                title: t('action.moveUp'),
                'aria-label': t('action.moveUp'),
                disabled: index === 0,
                onClick: () => props.onMove(index, index - 1),
              },
              '↑',
            ),
            h(
              'button',
              {
                type: 'button',
                className: 'branddeck-btn-icon',
                title: t('action.moveDown'),
                'aria-label': t('action.moveDown'),
                disabled: index === slides.length - 1,
                onClick: () => props.onMove(index, index + 1),
              },
              '↓',
            ),
            h(
              'button',
              {
                type: 'button',
                className: 'branddeck-btn-icon',
                title: t('action.deleteSlide'),
                'aria-label': t('action.deleteSlide'),
                disabled: slides.length <= 1,
                onClick: () => props.onDelete(index),
              },
              '🗑',
            ),
          ),
        ),
      )

      return h(
        'div',
        { className: 'branddeck-rail' },
        h('div', { className: 'branddeck-rail-title' }, t('label.slides')),
        items,
        h(
          'div',
          { className: 'branddeck-rail-footer' },
          h('button', { type: 'button', className: 'branddeck-btn', onClick: props.onDuplicate }, t('action.duplicateSlide')),
        ),
      )
    }

    /**
     * Error boundary around the panel body: a render bug shows a readable,
     * retryable message instead of blanking the whole slot.
     */
    class PanelBoundary extends React.Component {
      constructor(props) {
        super(props)
        this.state = { error: null }
        this.reset = this.reset.bind(this)
      }

      static getDerivedStateFromError(error) {
        return { error }
      }

      componentDidCatch(error) {
        /* Keep the detail for the message below; never re-throw. */
        this.setState({ error })
      }

      reset() {
        this.setState({ error: null })
      }

      render() {
        const error = this.state.error
        if (error === null || error === undefined) return this.props.children
        const detail = typeof error === 'object' && error !== null ? String(error.message ?? error) : String(error)
        return h(
          'div',
          { className: 'branddeck-root' },
          h(
            'div',
            { className: 'branddeck-banner branddeck-banner-error' },
            h('span', { className: 'branddeck-banner-text' }, `${t('error.render')}: ${detail}`),
            h('button', { type: 'button', className: 'branddeck-btn', onClick: this.reset }, t('action.retry')),
          ),
        )
      }
    }

    /** Sidebar icon: a brand card with copy lines, drawn in `currentColor`. */
    function SidebarIcon(props) {
      const size = props !== null && props !== undefined && Number.isFinite(props.size) ? props.size : 20
      return h(
        'svg',
        {
          width: size,
          height: size,
          viewBox: '0 0 24 24',
          fill: 'none',
          focusable: 'false',
          'aria-hidden': 'true',
          style: { display: 'block' },
        },
        h('rect', {
          key: 'frame',
          x: 2.25,
          y: 4.25,
          width: 19.5,
          height: 13,
          rx: 2.4,
          stroke: 'currentColor',
          strokeWidth: 1.6,
        }),
        h('path', {
          key: 'title-line',
          d: 'M6.2 8.6h6.4',
          stroke: 'currentColor',
          strokeWidth: 1.8,
          strokeLinecap: 'round',
        }),
        h('path', {
          key: 'body-line-1',
          d: 'M6.2 12.1h11.6',
          stroke: 'currentColor',
          strokeWidth: 1.4,
          strokeLinecap: 'round',
          opacity: 0.75,
        }),
        h('path', {
          key: 'body-line-2',
          d: 'M6.2 14.8h8.2',
          stroke: 'currentColor',
          strokeWidth: 1.4,
          strokeLinecap: 'round',
          opacity: 0.55,
        }),
        h('path', {
          key: 'stand',
          d: 'M8.6 20.2h6.8',
          stroke: 'currentColor',
          strokeWidth: 1.6,
          strokeLinecap: 'round',
          opacity: 0.75,
        }),
      )
    }

    /* ------------------------------------------------------------------ *
     * Panel body — the interactive editor.
     * ------------------------------------------------------------------ */

    /** Scale factor that makes the 1280x720 canvas fit the stage width. */
    function computeScale(availableWidth) {
      const width = Number.isFinite(availableWidth) && availableWidth > 0 ? availableWidth : CANVAS.width
      return Math.max(MIN_CANVAS_SCALE, Math.floor((width / CANVAS.width) * 100) / 100)
    }

    /** True when a key event originates from a field that owns its own keys. */
    function isTypingTarget(node) {
      if (node === null || node === undefined || typeof node !== 'object') return false
      const tag = typeof node.tagName === 'string' ? node.tagName.toUpperCase() : ''
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return true
      return node.isContentEditable === true
    }

    /**
     * The TEXT version of the deck (「文字版」).
     *
     * Looks like a document, edits like one, but round-trips through the Host:
     * `outline.get` renders the Markdown twin of the deck, `outline.save` applies
     * the edited text back onto the card JSON. Elements keep their geometry
     * because the Host matches them by `{#id}`; new `### 角色 {#id}` blocks are
     * auto-laid out in the safe area.
     */
    function TextView(props) {
      const { deckId, t, disabled, onApplied } = props
      const [markdown, setMarkdown] = React.useState('')
      const [origin, setOrigin] = React.useState('')
      const [status, setStatus] = React.useState('loading')
      const [error, setError] = React.useState(null)
      const [feedback, setFeedback] = React.useState(null)

      const load = React.useCallback(() => {
        if (typeof deckId !== 'string' || deckId.length === 0) {
          setStatus('empty')
          return
        }
        let alive = true
        setStatus('loading')
        setError(null)
        callRpc('outline.get', { deckId }).then((result) => {
          if (!alive) return
          if (result.ok) {
            const text = typeof result.value.markdown === 'string' ? result.value.markdown : ''
            setMarkdown(text)
            setOrigin(text)
            setStatus('idle')
            return
          }
          setError(readableError(result.error))
          setStatus('error')
        })
        return () => {
          alive = false
        }
      }, [deckId])

      React.useEffect(() => load(), [load])

      const dirty = markdown !== origin

      async function apply() {
        if (typeof deckId !== 'string' || deckId.length === 0) return
        setStatus('saving')
        setError(null)
        const result = await callRpc('outline.save', { deckId, markdown })
        if (!result.ok) {
          setError(readableError(result.error))
          setStatus('error')
          return
        }
        const value = result.value ?? {}
        const text = typeof value.markdown === 'string' ? value.markdown : markdown
        setMarkdown(text)
        setOrigin(text)
        setStatus('idle')
        const changes = value.changes ?? {}
        setFeedback({
          revision: value.revision ?? null,
          updated: Number(changes.updated ?? 0),
          added: Number(changes.added ?? 0),
          removed: Number(changes.removed ?? 0),
          warnings: Array.isArray(value.warnings) ? value.warnings : [],
        })
        if (typeof onApplied === 'function') onApplied()
      }

      if (status === 'empty') {
        return h('div', { className: 'branddeck-textview' }, h('div', { className: 'branddeck-empty' }, t('empty.title')))
      }

      const statusText =
        status === 'loading'
          ? t('text.loading')
          : status === 'saving'
            ? t('text.saving')
            : dirty
              ? t('text.dirty')
              : status === 'error'
                ? t('text.failed')
                : t('text.synced')

      return h(
        'div',
        { className: 'branddeck-textview' },
        h(
          'div',
          { className: 'branddeck-textbar' },
          h('span', { className: 'branddeck-textbar-hint' }, t('text.hint')),
          h('span', { className: 'branddeck-toolbar-spacer' }),
          h('span', { className: `branddeck-save branddeck-save-${dirty ? 'dirty' : 'saved'}` }, statusText),
          h(
            'button',
            {
              type: 'button',
              className: 'branddeck-btn',
              disabled: status === 'loading' || status === 'saving' || !dirty,
              onClick: () => {
                setMarkdown(origin)
                setFeedback(null)
              },
            },
            t('text.revert'),
          ),
          h(
            'button',
            {
              type: 'button',
              className: 'branddeck-btn branddeck-btn-primary',
              disabled: disabled || status === 'loading' || status === 'saving' || !dirty,
              onClick: () => void apply(),
            },
            t('text.apply'),
          ),
        ),
        error === null
          ? null
          : h('div', { className: 'branddeck-banner branddeck-banner-error' }, h('span', { className: 'branddeck-banner-text' }, error)),
        feedback === null
          ? null
          : h(
              'div',
              { className: 'branddeck-banner branddeck-banner-ok' },
              h(
                'span',
                { className: 'branddeck-banner-text' },
                t('text.applied', {
                  revision: String(feedback.revision ?? '—'),
                  updated: feedback.updated,
                  added: feedback.added,
                  removed: feedback.removed,
                }) + (feedback.warnings.length > 0 ? ` · ${feedback.warnings.slice(0, 3).join(' / ')}` : ''),
              ),
            ),
        h('textarea', {
          className: 'branddeck-textarea',
          spellCheck: false,
          value: markdown,
          'aria-label': t('text.ariaLabel'),
          onChange: (event) => {
            setMarkdown(event.target.value)
            setFeedback(null)
          },
        }),
      )
    }

    /**
     * The GENERATE entry point (「生成」视图) — the panel's front door.
     *
     * Three blocks, in the order a person actually thinks:
     *   1. 模板 — pick a parsed template or bring in a new .pptx by path (parsed once,
     *      then reusable forever); shows the palette/font/layout summary.
     *   2. 风格 — pick a reusable style profile («风格配置文件»), derive one from the
     *      template, tweak density/tone/layout, save it, and copy a ready-made
     *      instruction for the Agent to distil a style from a conversation.
     *   3. 内容 — paste text or point at a .md/.txt file, then 生成.
     */
    /* Label helpers use literal t() keys so every dictionary entry stays discoverable. */
    const densityLabel = (value, t) =>
      value === 'airy' ? t('gen.density.airy') : value === 'dense' ? t('gen.density.dense') : t('gen.density.standard')
    const toneLabel = (value, t) =>
      value === 'bold' ? t('gen.tone.bold') : value === 'minimal' ? t('gen.tone.minimal') : t('gen.tone.formal')
    const bulletLabel = (value, t) => (value === 'cards' ? t('gen.bullets.cards') : t('gen.bullets.list'))

    function GenerateView(props) {
      const { t, onGenerated, disabled } = props
      const [templates, setTemplates] = React.useState([])
      const [styles, setStyles] = React.useState([])
      const [templateId, setTemplateId] = React.useState('')
      const [styleId, setStyleId] = React.useState('default')
      const [spec, setSpec] = React.useState(null)
      const [importPath, setImportPath] = React.useState('')
      const [busy, setBusy] = React.useState(null)
      const [error, setError] = React.useState(null)
      const [result, setResult] = React.useState(null)

      const [text, setText] = React.useState('')
      const [sourcePath, setSourcePath] = React.useState('')
      const [title, setTitle] = React.useState('')
      const [maxSlides, setMaxSlides] = React.useState('')

      /* In-panel conversation: a sentence in, an artefact out. */
      const [styleChat, setStyleChat] = React.useState('')
      const [contentChat, setContentChat] = React.useState('')
      const [assistNote, setAssistNote] = React.useState(null)

      /* style form: a working copy of the selected style */
      const [draft, setDraft] = React.useState(null)

      const reload = React.useCallback(() => {
        let alive = true
        callRpc('styles.list', {}).then((response) => {
          if (!alive || !response.ok) return
          const list = Array.isArray(response.value.styles) ? response.value.styles : []
          setStyles(list)
          setStyleId((current) => (list.some((entry) => entry.styleId === current) ? current : list[0]?.styleId ?? 'default'))
        })
        callRpc('state', {}).then((response) => {
          if (!alive || !response.ok) return
          const list = Array.isArray(response.value.templates) ? response.value.templates : []
          setTemplates(list)
          setTemplateId((current) => (current !== '' && list.some((entry) => entry.templateId === current) ? current : list[0]?.templateId ?? ''))
        })
        return () => {
          alive = false
        }
      }, [])

      React.useEffect(() => reload(), [reload])

      /* keep the working copy in sync with the selection */
      React.useEffect(() => {
        const selected = styles.find((entry) => entry.styleId === styleId) ?? null
        setDraft(selected === null ? null : JSON.parse(JSON.stringify(selected)))
      }, [styleId, styles])

      /* template spec (palette/fonts/layouts) for the picker */
      React.useEffect(() => {
        if (templateId === '') {
          setSpec(null)
          return undefined
        }
        let alive = true
        callRpc('templates.spec', { templateId }).then((response) => {
          if (!alive) return
          setSpec(response.ok ? response.value : null)
        })
        return () => {
          alive = false
        }
      }, [templateId])

      async function importTemplate() {
        const path = importPath.trim()
        if (path === '') {
          setError(t('gen.needPath'))
          return
        }
        setBusy('import')
        setError(null)
        const response = await callRpc('templates.parse', { path }, 120000)
        setBusy(null)
        if (!response.ok) {
          setError(`${t('gen.importFailed')}: ${response.error.message || response.error.code}`)
          return
        }
        const parsed = response.value?.template ?? {}
        setImportPath('')
        setResult({ kind: 'import', templateId: parsed.templateId, name: parsed.templateName })
        if (typeof parsed.templateId === 'string') setTemplateId(parsed.templateId)
        reload()
      }

      /** 「说两句，出风格」: the Host asks the model and returns a draft to review. */
      async function askStyle() {
        const description = styleChat.trim()
        if (description === '') {
          setError(t('gen.needSentence'))
          return
        }
        setBusy('askStyle')
        setError(null)
        setAssistNote(null)
        const response = await callRpc('style.assist', { description, templateId }, 120000)
        setBusy(null)
        if (!response.ok) {
          setError(`${t('gen.assistFailed')}: ${response.error.message || response.error.code}`)
          return
        }
        const draftStyle = response.value.style
        setDraft(draftStyle)
        setStyleChat('')
        setAssistNote(t('gen.styleDraft', { name: draftStyle.name ?? '' }))
      }

      /** 「说两句，出内容」: model writes the outline, code lays it out, deck is stored. */
      async function askContent() {
        const description = contentChat.trim()
        if (description === '') {
          setError(t('gen.needSentence'))
          return
        }
        setBusy('askContent')
        setError(null)
        setResult(null)
        const response = await callRpc(
          'content.assist',
          { description, templateId, styleId, maxSlides: maxSlides.trim() !== '' ? Number(maxSlides) : undefined },
          180000,
        )
        setBusy(null)
        if (!response.ok) {
          setError(`${t('gen.assistFailed')}: ${response.error.message || response.error.code}`)
          return
        }
        setContentChat('')
        setResult({
          kind: 'deck',
          deckId: response.value.deckId,
          slides: response.value.slideCount,
          warnings: Array.isArray(response.value.warnings) ? response.value.warnings : [],
        })
        if (typeof onGenerated === 'function') onGenerated(response.value.deckId)
      }

      async function saveStyle() {
        if (draft === null) return
        setBusy('style')
        setError(null)
        const response = await callRpc('styles.save', { style: draft, styleId: draft.styleId })
        setBusy(null)
        if (!response.ok) {
          setError(`${t('gen.styleFailed')}: ${response.error.message || response.error.code}`)
          return
        }
        setResult({ kind: 'style', name: response.value.style.name })
        reload()
      }

      /** Derive a new style from the selected template's tokens. */
      async function deriveStyle() {
        if (templateId === '') return
        const base = styles.find((entry) => entry.styleId === styleId) ?? {}
        const nextId = `${templateId}-${Date.now().toString(36).slice(-4)}`
        const payload = {
          ...base,
          styleId: nextId,
          name: `${templateId} · ${t('gen.derived')}`,
          description: t('gen.derivedDesc'),
          templateId,
          source: 'template',
        }
        setBusy('derive')
        const response = await callRpc('styles.save', { style: payload, styleId: nextId })
        setBusy(null)
        if (!response.ok) {
          setError(response.error.message || response.error.code)
          return
        }
        setResult({ kind: 'style', name: response.value.style.name })
        reload()
        setStyleId(nextId)
      }

      async function removeStyle() {
        if (styleId === 'default') return
        setBusy('delete')
        const response = await callRpc('styles.delete', { styleId })
        setBusy(null)
        if (!response.ok) {
          setError(response.error.message || response.error.code)
          return
        }
        setStyleId('default')
        reload()
      }

      async function generate() {
        setBusy('generate')
        setError(null)
        setResult(null)
        const payload = { templateId, styleId }
        if (text.trim() !== '') payload.text = text
        if (sourcePath.trim() !== '') payload.path = sourcePath.trim()
        if (title.trim() !== '') payload.title = title.trim()
        if (maxSlides.trim() !== '' && Number.isFinite(Number(maxSlides))) payload.maxSlides = Number(maxSlides)
        const response = await callRpc('generate', payload, 120000)
        setBusy(null)
        if (!response.ok) {
          setError(`${t('gen.failed')}: ${response.error.message || response.error.code}`)
          return
        }
        setResult({
          kind: 'deck',
          deckId: response.value.deckId,
          slides: response.value.slideCount,
          warnings: Array.isArray(response.value.warnings) ? response.value.warnings : [],
        })
        if (typeof onGenerated === 'function') onGenerated(response.value.deckId)
      }

      /** Ready-made instruction to paste into the chat so the Agent writes a style. */
      const styleInstruction =
        `${t('gen.askAgent')}\n` + t('gen.agentExample', { templateId: templateId || 'templateId' })

      const canGenerate = !disabled && (text.trim() !== '' || sourcePath.trim() !== '') && busy === null

      return h(
        'div',
        { className: 'branddeck-gen' },
        error === null
          ? null
          : h(
              'div',
              { className: 'branddeck-banner branddeck-banner-error' },
              h('span', { className: 'branddeck-banner-text' }, error),
            ),

        /* ---- 1. 模板 ---- */
        h(
          'section',
          { className: 'branddeck-gen-block' },
          h('div', { className: 'branddeck-gen-head' }, h('span', { className: 'branddeck-gen-step' }, '1'), t('gen.stepTemplate')),
          h(
            'div',
            { className: 'branddeck-gen-row' },
            h(
              'select',
              {
                className: 'branddeck-select',
                value: templateId,
                'aria-label': t('gen.stepTemplate'),
                onChange: (event) => setTemplateId(event.target.value),
              },
              templates.length === 0
                ? h('option', { value: '' }, t('gen.noTemplate'))
                : templates.map((entry) => h('option', { key: entry.templateId, value: entry.templateId }, entry.templateName)),
            ),
            h('input', {
              className: 'branddeck-input',
              type: 'text',
              value: importPath,
              placeholder: t('gen.importPlaceholder'),
              'aria-label': t('gen.importPlaceholder'),
              onChange: (event) => setImportPath(event.target.value),
            }),
            h(
              'button',
              { type: 'button', className: 'branddeck-btn', disabled: busy !== null, onClick: () => void importTemplate() },
              busy === 'import' ? t('gen.importing') : t('gen.import'),
            ),
          ),
          spec === null
            ? h('div', { className: 'branddeck-gen-hint' }, t('gen.noSpec'))
            : h(
                'div',
                { className: 'branddeck-gen-spec' },
                h(
                  'div',
                  { className: 'branddeck-swatches' },
                  ['primary', 'accent', 'textPrimary', 'surface', 'background'].map((key) =>
                    h('span', {
                      key,
                      className: 'branddeck-swatch',
                      title: `${key} ${spec.theme?.colors?.[key] ?? ''}`,
                      style: { background: spec.theme?.colors?.[key] ?? 'transparent' },
                    }),
                  ),
                ),
                h(
                  'span',
                  { className: 'branddeck-gen-meta' },
                  `${t('gen.fonts')}: ${spec.theme?.fonts?.title ?? '—'} / ${spec.theme?.fonts?.body ?? '—'}`,
                ),
                h('span', { className: 'branddeck-gen-meta' }, `${t('gen.layouts')}: ${spec.layouts?.counts?.layouts ?? 0}`),
              ),
        ),

        /* ---- 2. 风格 ---- */
        h(
          'section',
          { className: 'branddeck-gen-block' },
          h('div', { className: 'branddeck-gen-head' }, h('span', { className: 'branddeck-gen-step' }, '2'), t('gen.stepStyle')),
          h(
            'div',
            { className: 'branddeck-gen-row' },
            h(
              'select',
              {
                className: 'branddeck-select',
                value: styleId,
                'aria-label': t('gen.stepStyle'),
                onChange: (event) => setStyleId(event.target.value),
              },
              styles.map((entry) =>
                h(
                  'option',
                  { key: entry.styleId, value: entry.styleId },
                  `${entry.name}${entry.source === 'chat' ? ` · ${t('gen.fromChat')}` : ''}`,
                ),
              ),
            ),
            h(
              'button',
              { type: 'button', className: 'branddeck-btn', disabled: busy !== null || templateId === '', onClick: () => void deriveStyle() },
              t('gen.derive'),
            ),
            h(
              'button',
              { type: 'button', className: 'branddeck-btn', disabled: busy !== null || draft === null, onClick: () => void saveStyle() },
              busy === 'style' ? t('gen.saving') : t('gen.saveStyle'),
            ),
            h(
              'button',
              { type: 'button', className: 'branddeck-btn', disabled: busy !== null || styleId === 'default', onClick: () => void removeStyle() },
              t('gen.deleteStyle'),
            ),
          ),
          h(
            'div',
            { className: 'branddeck-chat' },
            h('div', { className: 'branddeck-chat-title' }, t('gen.chatStyleTitle')),
            h(
              'div',
              { className: 'branddeck-chat-row' },
              h('input', {
                className: 'branddeck-input',
                type: 'text',
                value: styleChat,
                placeholder: t('gen.chatStylePlaceholder'),
                'aria-label': t('gen.chatStyleTitle'),
                onChange: (event) => setStyleChat(event.target.value),
                onKeyDown: (event) => {
                  if (event.key === 'Enter') void askStyle()
                },
              }),
              h(
                'button',
                {
                  type: 'button',
                  className: 'branddeck-btn branddeck-btn-primary',
                  disabled: busy !== null || styleChat.trim() === '',
                  onClick: () => void askStyle(),
                },
                busy === 'askStyle' ? t('gen.chatStyleBusy') : t('gen.chatStyleAction'),
              ),
            ),
            assistNote === null ? null : h('div', { className: 'branddeck-chat-note' }, assistNote),
          ),
          draft === null
            ? null
            : h(
                'div',
                { className: 'branddeck-gen-form' },
                h(
                  'label',
                  { className: 'branddeck-field' },
                  h('span', null, t('gen.density')),
                  h(
                    'select',
                    { className: 'branddeck-select', value: draft.density, onChange: (event) => setDraft({ ...draft, density: event.target.value }) },
                    ['airy', 'standard', 'dense'].map((value) => h('option', { key: value, value }, densityLabel(value, t))),
                  ),
                ),
                h(
                  'label',
                  { className: 'branddeck-field' },
                  h('span', null, t('gen.tone')),
                  h(
                    'select',
                    { className: 'branddeck-select', value: draft.tone, onChange: (event) => setDraft({ ...draft, tone: event.target.value }) },
                    ['formal', 'bold', 'minimal'].map((value) => h('option', { key: value, value }, toneLabel(value, t))),
                  ),
                ),
                h(
                  'label',
                  { className: 'branddeck-field' },
                  h('span', null, t('gen.bullets')),
                  h(
                    'select',
                    { className: 'branddeck-select', value: draft.bulletLayout, onChange: (event) => setDraft({ ...draft, bulletLayout: event.target.value }) },
                    ['list', 'cards'].map((value) => h('option', { key: value, value }, bulletLabel(value, t))),
                  ),
                ),
                h(
                  'label',
                  { className: 'branddeck-field' },
                  h('span', null, t('gen.maxSlides')),
                  h('input', {
                    className: 'branddeck-input branddeck-input-num',
                    type: 'number',
                    min: 1,
                    max: 60,
                    value: draft.maxSlides,
                    onChange: (event) =>
                      setDraft({ ...draft, maxSlides: Math.max(1, Math.min(60, Number(event.target.value) || 1)) }),
                  }),
                ),
                h(
                  'label',
                  { className: 'branddeck-field branddeck-field-wide' },
                  h('span', null, t('gen.describe')),
                  h('input', {
                    className: 'branddeck-input',
                    type: 'text',
                    value: draft.description ?? '',
                    placeholder: t('gen.describePlaceholder'),
                    onChange: (event) => setDraft({ ...draft, description: event.target.value }),
                  }),
                ),
              ),
          h(
            'details',
            { className: 'branddeck-gen-details' },
            h('summary', null, t('gen.askAgentTitle')),
            h('pre', { className: 'branddeck-gen-code' }, styleInstruction),
            h(
              'button',
              {
                type: 'button',
                className: 'branddeck-btn',
                onClick: () => {
                  if (typeof navigator !== 'undefined' && navigator.clipboard !== undefined) {
                    navigator.clipboard.writeText(styleInstruction).catch(() => {})
                  }
                },
              },
              t('gen.copy'),
            ),
          ),
        ),

        /* ---- 3. 内容 ---- */
        h(
          'section',
          { className: 'branddeck-gen-block' },
          h('div', { className: 'branddeck-gen-head' }, h('span', { className: 'branddeck-gen-step' }, '3'), t('gen.stepContent')),
          h(
            'div',
            { className: 'branddeck-chat' },
            h('div', { className: 'branddeck-chat-title' }, t('gen.chatContentTitle')),
            h(
              'div',
              { className: 'branddeck-chat-row' },
              h('input', {
                className: 'branddeck-input',
                type: 'text',
                value: contentChat,
                placeholder: t('gen.chatContentPlaceholder'),
                'aria-label': t('gen.chatContentTitle'),
                onChange: (event) => setContentChat(event.target.value),
                onKeyDown: (event) => {
                  if (event.key === 'Enter') void askContent()
                },
              }),
              h(
                'button',
                {
                  type: 'button',
                  className: 'branddeck-btn branddeck-btn-primary',
                  disabled: busy !== null || contentChat.trim() === '',
                  onClick: () => void askContent(),
                },
                busy === 'askContent' ? t('gen.chatContentBusy') : t('gen.chatContentAction'),
              ),
            ),
          ),
          h('textarea', {
            className: 'branddeck-textarea branddeck-gen-textarea',
            spellCheck: false,
            value: text,
            placeholder: t('gen.textPlaceholder'),
            'aria-label': t('gen.textPlaceholder'),
            onChange: (event) => setText(event.target.value),
          }),
          h(
            'div',
            { className: 'branddeck-gen-row' },
            h('input', {
              className: 'branddeck-input',
              type: 'text',
              value: sourcePath,
              placeholder: t('gen.filePlaceholder'),
              'aria-label': t('gen.filePlaceholder'),
              onChange: (event) => setSourcePath(event.target.value),
            }),
            h('input', {
              className: 'branddeck-input',
              type: 'text',
              value: title,
              placeholder: t('gen.titlePlaceholder'),
              'aria-label': t('gen.titlePlaceholder'),
              onChange: (event) => setTitle(event.target.value),
            }),
            h('input', {
              className: 'branddeck-input branddeck-input-num',
              type: 'number',
              min: 1,
              max: 60,
              value: maxSlides,
              placeholder: t('gen.maxSlidesShort'),
              'aria-label': t('gen.maxSlides'),
              onChange: (event) => setMaxSlides(event.target.value),
            }),
            h(
              'button',
              { type: 'button', className: 'branddeck-btn branddeck-btn-primary', disabled: !canGenerate, onClick: () => void generate() },
              busy === 'generate' ? t('gen.generating') : t('gen.generate'),
            ),
          ),
        ),

        result === null
          ? null
          : h(
              'div',
              { className: 'branddeck-banner branddeck-banner-ok' },
              h(
                'span',
                { className: 'branddeck-banner-text' },
                result.kind === 'deck'
                  ? `${t('gen.done')}: ${result.deckId} · ${result.slides} ${t('gen.slides')}${
                      result.warnings.length > 0 ? ` · ${result.warnings.slice(0, 3).join(' / ')}` : ''
                    }`
                  : result.kind === 'import'
                    ? `${t('gen.imported')}: ${result.name}`
                    : `${t('gen.styleSaved')}: ${result.name}`,
              ),
            ),
      )
    }

    function DeckPanelBody() {
      const [state, dispatch] = React.useReducer(reducer, undefined, createInitialState)
      const [scale, setScale] = React.useState(computeScale(CANVAS.width))
      const [reloadToken, setReloadToken] = React.useState(0)
      /* 「卡片视图」/「文字视图」 — two editable faces of the same deck. */
      const [view, setView] = React.useState('cards')

      const rootRef = React.useRef(null)
      const stageRef = React.useRef(null)
      const canvasRef = React.useRef(null)
      const gestureRef = React.useRef(null)
      /** Latest committed state, for event handlers bound once. */
      const stateRef = React.useRef(state)
      stateRef.current = state
      const scaleRef = React.useRef(scale)
      scaleRef.current = scale
      /** Latest "save now" implementation, used by the Ctrl/Cmd+S binding. */
      const saveNowRef = React.useRef(() => {})

      const theme = React.useMemo(() => normalizeTheme(state.theme), [state.theme])

      /* Theme plus the rasterized template logo, so `<img>` elements resolve a
       * real `data:` source. Rebuilt only when one of the two inputs changes. */
      const renderTheme = React.useMemo(() => {
        if (state.logo === null) return theme
        return {
          ...theme,
          brandAssets: {
            ...theme.brandAssets,
            icons: { ...theme.brandAssets.icons, logo: state.logo.dataUrl },
          },
        }
      }, [theme, state.logo])

      useStylesheet()

      /* ---- data loading ---------------------------------------------- */

      React.useEffect(() => {
        let alive = true
        dispatch({ type: 'load/start' })
        callRpc('state', {}).then((result) => {
          if (!alive) return
          if (result.ok) {
            dispatch({ type: 'view/loaded', view: normalizeStatePayload(result.value) })
            return
          }
          /* No deck on the Host yet is an empty state, not a failure. */
          if (result.error.code === CODE_DECK_NOT_FOUND) {
            dispatch({ type: 'view/loaded', view: emptyDeckView() })
            return
          }
          dispatch({ type: 'load/error', error: readableError(result.error) })
        })
        return () => {
          alive = false
        }
      }, [reloadToken])

      /** Re-read one deck (deck picker). */
      async function switchDeck(deckId) {
        dispatch({ type: 'deck/pending', deckId })
        dispatch({ type: 'load/start' })
        const result = await callRpc('decks.get', { deckId })
        if (result.ok) {
          dispatch({ type: 'view/loaded', view: normalizeStatePayload(result.value) })
          return
        }
        if (result.error.code === CODE_DECK_NOT_FOUND) {
          dispatch({ type: 'view/loaded', view: emptyDeckView() })
          return
        }
        dispatch({ type: 'load/error', error: readableError(result.error) })
      }

      /* ---- template logo --------------------------------------------- */

      /* The theme only names a logo file; the rasterized pixels come from
       * `templates.logo`, cached here per template id. */
      React.useEffect(() => {
        const templateId = theme.templateId
        if (theme.brandAssets.logoPath === undefined) return undefined
        const cached = state.logo
        if (cached !== null && cached.templateId === templateId) return undefined
        let alive = true
        callRpc('templates.logo', { templateId }).then((result) => {
          if (!alive || !result.ok) return
          const value = result.value
          const dataUrl = value !== null && typeof value === 'object' ? value.dataUrl : undefined
          if (typeof dataUrl === 'string' && dataUrl.startsWith('data:')) {
            dispatch({ type: 'logo/loaded', templateId, dataUrl })
          }
        })
        return () => {
          alive = false
        }
      }, [theme.templateId, theme.brandAssets.logoPath, state.logo])

      /* ---- autosave --------------------------------------------------- */

      /** Persist one deck snapshot through the Host validator. */
      async function saveDeck(deckSnapshot, revision) {
        dispatch({ type: 'save/start' })
        const result = await callRpc('decks.save', { deck: deckSnapshot })
        if (result.ok) dispatch({ type: 'save/success', rev: revision, view: result.value })
        else dispatch({ type: 'save/error', error: readableError(result.error) })
      }

      saveNowRef.current = () => {
        const current = stateRef.current
        if (current.deck !== null) void saveDeck(current.deck, current.rev)
      }

      React.useEffect(() => {
        if (state.dirty !== true || state.deck === null) return undefined
        /* Wait for the answer of an in-flight save: a superseded answer keeps the
         * deck dirty and this effect re-runs when `saveState` settles. */
        if (state.saveState === 'saving') return undefined
        const snapshot = state.deck
        const revision = state.rev
        const timer = setTimeout(() => {
          void saveDeck(snapshot, revision)
        }, AUTOSAVE_DEBOUNCE_MS)
        return () => clearTimeout(timer)
        // The deck identity changes on every mutation, restarting the debounce.
      }, [state.dirty, state.rev, state.deck, state.saveState])

      /* ---- responsive 16:9 viewport ---------------------------------- */

      React.useEffect(() => {
        const node = stageRef.current
        if (node === null || node === undefined) return undefined
        const apply = () => setScale(computeScale(node.clientWidth))
        apply()
        if (typeof ResizeObserver === 'undefined') {
          if (typeof window === 'undefined') return undefined
          window.addEventListener('resize', apply)
          return () => window.removeEventListener('resize', apply)
        }
        const observer = new ResizeObserver(apply)
        observer.observe(node)
        return () => observer.disconnect()
      }, [])

      /* ---- keyboard -------------------------------------------------- */

      React.useEffect(() => {
        if (typeof window === 'undefined') return undefined
        const onKeyDown = (event) => {
          const root = rootRef.current
          if (root === null || root === undefined) return
          const active = typeof document === 'undefined' ? null : document.activeElement
          if (active === null || active === undefined || root.contains(active) !== true) return
          const current = stateRef.current
          if (current.editingId !== null) return

          const meta = event.ctrlKey === true || event.metaKey === true
          /* Ctrl/Cmd+S is a global save shortcut, so it works from inspector
           * fields too; every other binding defers to the focused field. */
          if (meta === true && (event.key === 's' || event.key === 'S')) {
            event.preventDefault()
            saveNowRef.current()
            return
          }
          if (isTypingTarget(active)) return

          if (meta === true && (event.key === 'd' || event.key === 'D')) {
            if (current.selectionId === null) return
            event.preventDefault()
            dispatch({ type: 'element/duplicate', id: current.selectionId })
            return
          }
          if (event.key === 'Escape') {
            if (current.selectionId === null) return
            event.preventDefault()
            dispatch({ type: 'element/select', id: null })
            return
          }
          if (event.key === 'Delete' || event.key === 'Backspace') {
            if (current.selectionId === null) return
            event.preventDefault()
            dispatch({ type: 'element/delete', id: current.selectionId })
            return
          }
          const deltas = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }
          const delta = deltas[event.key]
          if (delta === undefined || current.selectionId === null) return
          event.preventDefault()
          const step = event.shiftKey === true ? NUDGE_STEP_FAST : NUDGE_STEP
          dispatch({ type: 'element/nudge', id: current.selectionId, dx: delta[0] * step, dy: delta[1] * step })
        }
        window.addEventListener('keydown', onKeyDown)
        return () => window.removeEventListener('keydown', onKeyDown)
      }, [])

      /* ---- gestures -------------------------------------------------- */

      /**
       * Start a drag or a resize gesture. Pointer capture keeps the gesture alive
       * even when the pointer leaves the element.
       */
      function startGesture(event, box, id, mode, corner) {
        if (event.button !== undefined && event.button !== 0) return
        if (stateRef.current.editingId !== null) return
        /* Only stop the canvas deselection. Calling preventDefault() here would
         * suppress the compatibility mouse events, which breaks the double-click
         * that starts inline text editing. The canvas already disables text
         * selection, so the drag has nothing else to cancel. */
        event.stopPropagation()
        dispatch({ type: 'element/select', id })
        if (canvasRef.current !== null && typeof canvasRef.current.focus === 'function') {
          canvasRef.current.focus({ preventScroll: true })
        }
        const target = event.currentTarget
        if (target !== null && target !== undefined && typeof target.setPointerCapture === 'function') {
          try {
            target.setPointerCapture(event.pointerId)
          } catch (error) {
            /* capture is best effort */
          }
        }
        gestureRef.current = {
          mode,
          corner: corner ?? null,
          id,
          pointerId: event.pointerId,
          startX: event.clientX,
          startY: event.clientY,
          origin: { x: box.x, y: box.y, w: box.w, h: box.h },
        }
      }

      function onDragStart(event, element) {
        startGesture(event, element, element.id, 'move', null)
      }

      function onResizeStart(event, corner) {
        const current = stateRef.current
        const slide = deckSlides(current.deck)[currentSlideIndex(current)]
        const element = slideElements(slide).find((candidate) => candidate.id === current.selectionId)
        if (element === undefined) return
        startGesture(event, element, element.id, 'resize', corner)
      }

      function onPointerMove(event) {
        const gesture = gestureRef.current
        if (gesture === null || gesture.pointerId !== event.pointerId) return
        const currentScale = scaleRef.current > 0 ? scaleRef.current : 1
        const dx = (event.clientX - gesture.startX) / currentScale
        const dy = (event.clientY - gesture.startY) / currentScale
        if (gesture.mode === 'move') {
          dispatch({ type: 'element/move', id: gesture.id, x: gesture.origin.x + dx, y: gesture.origin.y + dy })
          return
        }
        dispatch({ type: 'element/resize', id: gesture.id, corner: gesture.corner, dx, dy })
      }

      function onPointerUp(event) {
        const gesture = gestureRef.current
        if (gesture === null) return
        gestureRef.current = null
        const target = event.currentTarget
        if (target !== null && target !== undefined && typeof target.releasePointerCapture === 'function') {
          try {
            target.releasePointerCapture(gesture.pointerId)
          } catch (error) {
            /* already released */
          }
        }
      }

      function onElementDoubleClick(event, element) {
        event.stopPropagation()
        if (element.type !== 'text') return
        event.preventDefault()
        dispatch({ type: 'element/select', id: element.id })
        dispatch({ type: 'edit/start', id: element.id })
      }

      /* ---- slide + element actions ----------------------------------- */

      function deleteSlide(index) {
        const slides = deckSlides(stateRef.current.deck)
        if (slides.length <= 1) return
        const confirmed = typeof window !== 'undefined' && typeof window.confirm === 'function' ? window.confirm(t('confirm.deleteSlide')) : true
        if (confirmed !== true) return
        dispatch({ type: 'slide/delete', index })
      }

      /* ---- export ---------------------------------------------------- */

      async function runExport() {
        const current = stateRef.current
        const deckId = current.deck?.deckId ?? current.activeDeckId
        if (typeof deckId !== 'string' || deckId.length === 0) return
        dispatch({ type: 'export/start' })
        /* Compiling the real .pptx takes seconds, so it gets its own bound. */
        const result = await callRpc('export', { deckId }, EXPORT_TIMEOUT_MS)
        if (result.ok) {
          const value = result.value !== null && typeof result.value === 'object' ? result.value : {}
          dispatch({
            type: 'export/success',
            path: typeof value.path === 'string' ? value.path : '',
            slideCount: Number.isFinite(value.slideCount) ? value.slideCount : deckSlides(current.deck).length,
          })
          return
        }
        dispatch({ type: 'export/error', message: readableError(result.error) })
      }

      /* ---- derived view data ----------------------------------------- */

      const slides = deckSlides(state.deck)
      const slideIndex = clamp(state.slideIndex, 0, Math.max(0, slides.length - 1))
      const slide = slides[slideIndex] ?? null
      const elements = slideElements(slide)
      const selectedElement = elements.find((element) => element.id === state.selectionId) ?? null
      const overflow = estimateTextOverflow(slide, renderTheme)
      const overflowIds = new Set(overflow.overflowing.map((item) => item.id))
      const overflowLines = new Map(overflow.overflowing.map((item) => [item.id, item.estimatedLines]))
      const canvasBackground = resolveColor(slide === null ? undefined : slide.background, renderTheme, renderTheme.colors.background)
      const saveLabel = t(`save.${state.saveState}`)
      const deckOptions = state.decks.slice()
      if (state.deck !== null && !deckOptions.some((summary) => summary.deckId === state.deck.deckId)) {
        deckOptions.push({
          deckId: state.deck.deckId,
          title: state.deck.title ?? state.deck.deckId,
          templateId: state.deck.templateId ?? '',
          slideCount: slides.length,
          updatedAt: '',
        })
      }

      /* ---- toolbar --------------------------------------------------- */

      const toolbar = h(
        'div',
        { className: 'branddeck-toolbar' },
        h('span', { className: 'branddeck-toolbar-title' }, t('appTitle')),
        h(
          'label',
          { className: 'branddeck-field' },
          h('span', null, t('label.deck')),
          h(
            'select',
            {
              className: 'branddeck-select',
              value: state.activeDeckId ?? '',
              disabled: deckOptions.length === 0,
              'aria-label': t('label.deck'),
              onChange: (event) => {
                const deckId = event.target.value
                if (deckId.length > 0 && deckId !== state.activeDeckId) void switchDeck(deckId)
              },
            },
            deckOptions.length === 0
              ? h('option', { value: '' }, t('label.noDeck'))
              : deckOptions.map((summary) =>
                  h(
                    'option',
                    { key: summary.deckId, value: summary.deckId },
                    `${summary.title} · ${t('label.slideCount', { count: summary.slideCount })}`,
                  ),
                ),
          ),
        ),
        state.deck === null
          ? null
          : h('span', { className: 'branddeck-chip', title: theme.templateName }, `${t('label.template')}: ${theme.templateName}`),
        h('span', { className: `branddeck-save branddeck-save-${state.saveState}` }, saveLabel),
        h('span', { className: 'branddeck-toolbar-spacer' }),
        h(
          'div',
          { className: 'branddeck-viewtoggle', role: 'group', 'aria-label': t('view.switch') },
          h(
            'button',
            {
              type: 'button',
              className: `branddeck-viewtoggle-btn${view === 'cards' ? ' branddeck-viewtoggle-on' : ''}`,
              'aria-pressed': view === 'cards',
              onClick: () => setView('cards'),
            },
            t('view.cards'),
          ),
          h(
            'button',
            {
              type: 'button',
              className: `branddeck-viewtoggle-btn${view === 'text' ? ' branddeck-viewtoggle-on' : ''}`,
              'aria-pressed': view === 'text',
              onClick: () => setView('text'),
            },
            t('view.text'),
          ),
          h(
            'button',
            {
              type: 'button',
              className: `branddeck-viewtoggle-btn${view === 'gen' ? ' branddeck-viewtoggle-on' : ''}`,
              'aria-pressed': view === 'gen',
              onClick: () => setView('gen'),
            },
            t('view.gen'),
          ),
        ),
        h(
          'button',
          { type: 'button', className: 'branddeck-btn', onClick: () => setReloadToken((token) => token + 1) },
          t('action.refresh'),
        ),
        h(
          'button',
          {
            type: 'button',
            className: 'branddeck-btn branddeck-btn-primary',
            disabled: state.deck === null || state.exporting === true,
            onClick: () => void runExport(),
          },
          t('action.export'),
        ),
      )

      /* ---- banners --------------------------------------------------- */

      const banners = []
      if (state.error !== null) {
        banners.push(
          h(
            'div',
            { className: 'branddeck-banner branddeck-banner-error', key: 'load-error' },
            h('span', { className: 'branddeck-banner-text' }, `${t('error.load')}: ${state.error}`),
            h(
              'button',
              { type: 'button', className: 'branddeck-btn', onClick: () => setReloadToken((token) => token + 1) },
              t('action.retry'),
            ),
          ),
        )
      }
      if (state.saveError !== null) {
        banners.push(
          h(
            'div',
            { className: 'branddeck-banner branddeck-banner-error', key: 'save-error' },
            h('span', { className: 'branddeck-banner-text' }, `${t('save.error')}: ${state.saveError}`),
            h(
              'button',
              { type: 'button', className: 'branddeck-btn', onClick: () => saveNowRef.current() },
              t('action.retry'),
            ),
          ),
        )
      }
      if (state.warnings.length > 0 && state.warningsDismissed !== true) {
        const shown = state.warnings.slice(0, 6)
        const hiddenCount = state.warnings.length - shown.length
        banners.push(
          h(
            'div',
            { className: 'branddeck-banner branddeck-banner-warn', key: 'warnings' },
            h(
              'div',
              { className: 'branddeck-banner-text' },
              h('div', { className: 'branddeck-banner-title' }, t('warnings.title')),
              h(
                'ul',
                { className: 'branddeck-banner-list' },
                shown.map((warning, index) => h('li', { key: `warning-${index}` }, warning)),
              ),
              hiddenCount > 0 ? h('div', null, t('warnings.more', { count: hiddenCount })) : null,
            ),
            h('button', { type: 'button', className: 'branddeck-btn', onClick: () => dispatch({ type: 'warnings/dismiss' }) }, t('action.dismiss')),
          ),
        )
      }
      if (state.exportResult !== null) {
        const result = state.exportResult
        banners.push(
          h(
            'div',
            {
              className: `branddeck-banner ${result.ok === true ? 'branddeck-banner-success' : 'branddeck-banner-error'}`,
              key: 'export-result',
            },
            h(
              'span',
              { className: 'branddeck-banner-text' },
              result.ok === true
                ? t('export.done', { path: result.path, count: result.slideCount })
                : t('export.failed', { message: result.message }),
            ),
            h('button', { type: 'button', className: 'branddeck-btn', onClick: () => dispatch({ type: 'export/dismiss' }) }, t('action.close')),
          ),
        )
      }

      /* ---- stage ----------------------------------------------------- */

      let stageContent = null
      if (state.status === 'loading') {
        stageContent = h(
          'div',
          { className: 'branddeck-loading' },
          h('span', { className: 'branddeck-spinner' }),
          h('span', null, t('state.loading')),
        )
      } else if (slide === null) {
        stageContent = h(
          'div',
          { className: 'branddeck-empty' },
          h('div', { className: 'branddeck-empty-title' }, t('empty.title')),
          h('div', { className: 'branddeck-empty-hint' }, t('empty.hint')),
          h(
            'button',
            { type: 'button', className: 'branddeck-btn', onClick: () => setReloadToken((token) => token + 1) },
            t('action.refresh'),
          ),
        )
      } else {
        stageContent = h(
          'div',
          { className: 'branddeck-frame', style: { width: CANVAS.width * scale, height: CANVAS.height * scale } },
          h(
            'div',
            {
              ref: canvasRef,
              className: 'branddeck-canvas',
              tabIndex: 0,
              style: {
                width: CANVAS.width,
                height: CANVAS.height,
                transform: `scale(${scale})`,
                background: canvasBackground,
              },
              onPointerDown: (event) => {
                /* Pressing empty canvas clears the selection, like a design tool. */
                if (event.target === event.currentTarget) {
                  dispatch({ type: 'element/select', id: null })
                }
              },
            },
            elements.map((element) =>
              h(ElementView, {
                key: element.id,
                element,
                theme: renderTheme,
                scale,
                selected: element.id === state.selectionId,
                overflows: overflowIds.has(element.id),
                neededLines: overflowLines.get(element.id) ?? 0,
                editing: element.id === state.editingId,
                onDragStart,
                onResizeStart,
                onPointerMove,
                onPointerUp,
                onDoubleClick: onElementDoubleClick,
                onEditCommit: (id, content) => dispatch({ type: 'edit/commit', id, content }),
                onEditCancel: () => dispatch({ type: 'edit/cancel' }),
              }),
            ),
          ),
        )
      }

      const body =
        view === 'gen'
          ? h(GenerateView, {
              t,
              disabled: false,
              onGenerated: (deckId) => {
                /* Fresh deck: reload the panel data, then show it as cards. */
                setReloadToken((token) => token + 1)
                setView('cards')
              },
            })
          : view === 'text'
          ? h(TextView, {
              deckId: state.deck === null ? '' : state.deck.deckId,
              t,
              disabled: state.deck === null,
              onApplied: () => setReloadToken((token) => token + 1),
            })
          : h(
        'div',
        { className: 'branddeck-body' },
        h(SlideRail, {
          slides,
          theme: renderTheme,
          selectedIndex: slideIndex,
          onSelect: (index) => dispatch({ type: 'slide/select', index }),
          onMove: (from, to) => dispatch({ type: 'slide/reorder', from, to }),
          onDelete: (index) => deleteSlide(index),
          onDuplicate: () => dispatch({ type: 'slide/duplicate' }),
        }),
        h('div', { className: 'branddeck-stage' }, h('div', { className: 'branddeck-stage-inner', ref: stageRef }, stageContent)),
        h(Inspector, {
          element: selectedElement,
          theme: renderTheme,
          slideNumber: slideIndex + 1,
          slideTotal: Math.max(1, slides.length),
          canDelete: elements.length > 1,
          onGeometry: (field, value) => {
            if (selectedElement === null) return
            dispatch({ type: 'element/geometry', id: selectedElement.id, patch: { [field]: value } })
          },
          onStyle: (patch) => {
            if (selectedElement === null) return
            dispatch({ type: 'element/style', id: selectedElement.id, patch })
          },
          onDuplicateElement: (id) => dispatch({ type: 'element/duplicate', id }),
          onDeleteElement: (id) => dispatch({ type: 'element/delete', id }),
        }),
      )
      const modal =
        state.exporting === true
          ? h(
              'div',
              { className: 'branddeck-modal-mask' },
              h(
                'div',
                { className: 'branddeck-modal' },
                h('div', { className: 'branddeck-modal-title' }, t('export.title')),
                h(
                  'div',
                  { className: 'branddeck-loading' },
                  h('span', { className: 'branddeck-spinner' }),
                  h('span', null, t('export.running')),
                ),
              ),
            )
          : null

      return h('div', { className: 'branddeck-root', ref: rootRef }, toolbar, banners, body, modal)
    }

    /** Slot entry point: boundary first, so a render bug cannot blank the slot. */
    function DeckPanel() {
      return h(PanelBoundary, null, h(DeckPanelBody))
    }

    /* ------------------------------------------------------------------ *
     * Plugin registration.
     * ------------------------------------------------------------------ */

    return {
      /* `connection` is deliberately NOT declared here: it is installed on the
       * Connection plugin's own scope, so a sibling bundle plugin can never see
       * it, and injecting it would leave this entry pending ("loading") and fail
       * the whole web boot. Data goes over the Host's same-origin route instead. */
      inject: ['slots'],
      apply(ctx) {
        try {
          register(ctx)
        } catch (error) {
          /* Never fail web boot for a panel problem: report it in the console
           * (and therefore in the slot's own error boundary) instead. */
          console.error('[branddeck] apply failed:', error)
        }
      },
    }

    /** Register locale copy and both slots; throws only into apply()'s guard. */
    function register(ctx) {
        // `locale` is read defensively rather than injected: a missing or
        // strict-context service must not prevent the slots from registering.
        let locale
        try {
          locale = ctx === null || ctx === undefined ? undefined : (ctx.locale ?? (typeof ctx.get === 'function' ? ctx.get('locale') : undefined))
        } catch (error) {
          locale = undefined
        }
        const bound = locale !== undefined && locale !== null && typeof locale.bind === 'function' ? locale.bind(NS) : null
        runtime = { ctx, bound }

        /* Register the namespace dictionary as an owned effect so the panel copy
         * participates in the host locale service (and language switches). */
        ctx.effect(() => {
          try {
            const dispose =
              locale !== undefined && locale !== null && typeof locale.register === 'function'
                ? locale.register(NS, { zh: ZH, en: EN })
                : undefined
            return () => {
              if (typeof dispose === 'function') {
                try {
                  dispose()
                } catch (error) {
                  /* unregister is best effort */
                }
              }
            }
          } catch (error) {
            /* An already-registered namespace must not block the panel. */
            return () => {}
          }
        }, 'branddeck: locale dictionaries')

        /* Main page (keyed slot, root scope: the component receives no props). */
        ctx.slots.inject('main', () =>
          ctx.slots.register(
            {
              name: 'main',
              key: PANEL_ID,
              locale: NS,
            },
            DeckPanel,
          ),
        )

        /* Sidebar icon that opens the page (list slot, receives `{ size }`). */
        ctx.slots.inject('sidebar.panellist', () =>
          ctx.slots.register(
            {
              name: 'sidebar.panellist',
              id: PANEL_ID,
              order: 20,
              locale: NS,
              label: () => t('panel'),
            },
            SidebarIcon,
          ),
        )
    }
  },
})