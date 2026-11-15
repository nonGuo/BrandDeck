# BrandDeck 品牌演示（DSH 插件）

> 把企业 `.pptx` 模板解构成**设计令牌**，让大模型在 **1280×720 虚拟画布**上用卡片 DSL 自由排版，
> 在 DSH 侧栏面板里**所见即所得地微调**，最后编译成**原生可编辑**的 PowerPoint 文件。

本插件是对 `设计文档.md` 的完整落地：模板解析、LLM 生成流水线、前端卡片交互、PPTX 编译导出四大模块
全部实现，前后端同源 TypeScript 化设计改为**无构建步骤的 ESM + 单文件浏览器模块**，
因此不需要 `electron-rebuild`、不需要前端打包器，`pnpm install` 之后即可安装。

---

## 1. 能力一览

| 模块 | 实现 | 说明 |
| --- | --- | --- |
| 模板特征提取 | `src/pptx/*` + `src/template.js` | `jszip` 解包 + `fast-xml-parser` 解析 `theme1.xml`：`a:clrScheme`（兼容 `srgbClr` 与 `sysClr/@lastClr`）、`a:fontScheme`（latin / ea / cs）、母版 Logo 精准定位（含真实摆放坐标）、`sldSz` 真实尺寸 |
| 生成流水线 | 8 个 `branddeck_*` 工具 + 系统提示段 | 由**当前会话的 Agent** 担任“纯文本大模型”：插件负责结构化输出契约、Zod 强校验、坐标修复与溢出预算 |
| 卡片交互视图 | `client.js`（Web 面板） | 16:9 等比视口、双击改字、拖拽/缩放、方向键微调、缩略图排序、实时溢出预警、自动保存 |
| PPTX 编译导出 | `src/compile.js` | `pptxgenjs` 从零构建 `LAYOUT_WIDE`(13.333×7.5in)：文本→原生文本框、图形→原生形状、图表→原生 Office 图表、SVG→`@resvg/resvg-js` 3× 光栅化后嵌入 |
| 面板通道 | `/branddeck-api` 同源 HTTP 路由 | 客户端 `inject` 只声明 `['slots']`：`connection` 服务装在内置插件的私有作用域里，兄弟插件取不到，声明为依赖会让客户端条目永久 pending 并导致启动失败 |

## 2. 安装

```bash
# 1) 安装插件自身依赖（只需一次；含 @resvg/resvg-js 的 win32-x64 预编译原生包）
#    依赖必须装在本包目录内：bundle 以 link: 方式装入 profile，
#    Node 会从插件的真实路径向上找 node_modules。
cd plugins/branddeck
pnpm install

# 2) 在 DSH 里安装 bundle（绝对路径）
plugin_manager install_bundle { "target": "D:\\Projects\\ai\\BrandDeck\\plugins\\branddeck" }
```

> 若 `pnpm` 不在 PATH（DSH 内置运行时），可显式调用：
> `node "<dsh>/dependencies/pnpm/bin/pnpm.mjs" install`。
> 安装完成后**刷新一次页面**：已安装的 bundle 行会即时生效，但浏览器需要重新加载
> 客户端模块（`/plugins/…/client.js`）才能在侧栏出现 BrandDeck 图标。

## 3. 面板（Web Client）

- **顶栏**：Deck 选择、模板徽标、保存状态（保存中/已保存/未保存/保存失败）、刷新、导出 `.pptx`。
- **左栏**：每页缩略图（按相同 DSL 等比缩小渲染）、上移/下移/删除/复制页、`+ 复制页`。
- **画布**：固定 1280×720 基准，`transform: scale()` 自适应面板宽度；元素按基准像素绝对定位。
- **交互**：单击选中（四角控制点）／双击文本原位编辑（`Enter` 提交、`Shift+Enter` 换行、`Esc` 取消）／
  拖拽移动／角点缩放／方向键 1px（`Shift` 10px）／`Delete` 删除／`Ctrl+Cmd+D` 复制／`Ctrl+Cmd+S` 立即保存。
- **预警**：文本按 CJK 宽度估算行高，超出文本框时红框 + 角标；Host 返回的修复告警在顶部条中可关闭。
- **样式**：外壳仅使用 `--dsw-alias-*` 主题令牌（明/暗主题自动适配），卡片内部使用模板品牌色。

## 4. 工具（模型侧）

| 工具 | 用途 |
| --- | --- |
| `branddeck_template_parse` | `{ path, templateId? }` 解析磁盘上的 `.pptx` 模板 → 品牌色/字体/Logo，落盘并返回摘要 |
| `branddeck_template_list` | 已解析模板列表 |
| `branddeck_theme_get` | `{ templateId? }` 读取令牌（写 Deck 时用 `$colors.*` 引用） |
| `branddeck_deck_create` | `{ title, templateId? }` 新建带封面骨架的 Deck |
| `branddeck_deck_save` | `{ deck }` 校验 + 修复 + 落盘，返回 `warnings` |
| `branddeck_deck_get` / `branddeck_deck_list` | 读回 / 列举 |
| `branddeck_deck_export` | `{ deckId?, outputPath? }` 编译原生 `.pptx`（默认 `<cwd>/branddeck/<deckId>.pptx`） |

系统提示段 `branddeck` 教会模型 DSL 规则：安全区 64px、右上 Logo 预留、单页正文 ≤150 字、
SVG 仅允许纯矢量标签（禁 `foreignObject`/`script`）、流程图优先用 `svg`/`shape` 而非长文本。

## 5. 配置（`cordis.patch.yml` 的 row config）

```yaml
- insert:
    - id: branddeck
      name: '@local/dsh-branddeck'
      config:
        dataDir: 'D:/dsh-data/branddeck'   # 默认 $DSH_HOME/branddeck
        promptSection: true                # 是否注入系统提示段
        svgRasterScale: 3                  # SVG 光栅化倍数（3 ≈ 288 DPI）
        fontFiles: []                      # 例如容器内显式挂载 SourceHanSansCN-Regular.ttf
        fontDirs: []
        loadSystemFonts: true              # false 时仅用 fontFiles/fontDirs（容器推荐）
        logoSide: auto                     # auto | top-right | none
        defaultTemplateId: ppt-hw
```

> 插件刻意**不导出 `Config`**：配置保持纯对象，避免依赖 `schemastery` 造成跨副本实例不兼容，
> 用户写在 patch 层里的配置也不会因插件升级被覆盖。

## 6. 数据布局

```
<dataDir>/                       # 默认 $DSH_HOME/branddeck
  templates/<templateId>/theme.json     # ThemeTokens（含真实 Logo 路径与母版摆框）
  templates/<templateId>/logo.png       # 从母版提取的 Logo
  templates/<templateId>/meta.json      # 模板元信息 + 媒体清单 + 诊断
  decks/<deckId>.json                   # CardDeck（DSL）
  exports/<deckId>.pptx                 # 面板导出的产物
```

`templateId` / `deckId` 必须匹配 `^[a-zA-Z0-9._-]{1,64}$`，否则拒绝（路径穿越防护）。

## 7. DSL 速览

```jsonc
{
  "deckId": "hw-demo", "title": "…", "templateId": "ppt-hw",
  "slides": [{
    "slideId": "cover", "background": "$colors.background", "speakerNotes": "…",
    "elements": [
      { "type": "text",  "id": "t1", "role": "title", "content": "标题",
        "x": 112, "y": 258, "w": 860, "h": 96,
        "style": { "fontSize": 60, "bold": true, "color": "$colors.textPrimary" } },
      { "type": "shape", "id": "s1", "shapeType": "roundRect",
        "x": 80, "y": 220, "w": 340, "h": 280,
        "style": { "fill": "$colors.surface", "radius": 20 } },
      { "type": "chart", "id": "c1", "chartType": "bar",
        "x": 80, "y": 176, "w": 700, "h": 400,
        "data": { "labels": ["Q1"], "series": [{ "name": "交付量", "values": [120] }] } },
      { "type": "svg", "id": "v1", "x": 80, "y": 190, "w": 1100, "h": 250, "svgCode": "<svg …>" },
      { "type": "image", "id": "i1", "assetKey": "logo", "x": 1000, "y": 600, "w": 200, "h": 56 }
    ]
  }]
}
```

`repairDeck()` 会做：坐标钳制进画布、Logo 角落避让、文本最小盒、SVG 消毒（剔 `foreignObject`/
`script`/事件属性、补 `xmlns` 与尺寸）、图表序列长度对齐、重复 id 重命名，并把每一处修改写进 `warnings`。

## 8. 验证

```bash
node test/smoke-deps.mjs      # 依赖可用性（jszip / fast-xml-parser / zod / pptxgenjs / resvg）
node test/smoke-cjk.mjs       # CJK 在 resvg 里不乱码（设计文档坑点 1）
node test/schema-test.mjs     # 共享契约（13 项）
node test/engine-*.mjs        # 解析 / 令牌 / 存储 / 光栅化 / 编译
node test/client-smoke.mjs    # 浏览器半的静态审计 + 沙箱求值
node test/e2e-hw.mjs          # 真实模板 → 令牌 → Deck → 原生 pptx 全链路
```

`e2e-hw.mjs` 会把 `PPT模板-HW.pptx` 解析结果与 `research/expected-tokens.json`（独立分析产物）逐字段比对，
并用 LibreOffice Kit 渲染导出的 `.pptx` 做人工视觉核对。

## 9. 已知边界

- 模板解析只读取**主题/母版/版式级**的品牌信息，不复刻模板里的具体版式（这是“自由排版”路线的取舍）。
- SVG 走位图嵌入（设计文档要求），因此 SVG 元素在 PowerPoint 内不可再编辑；文本/图形/图表均为原生可编辑对象。
- 面板根作用域拿不到会话 id，Deck 存储因此是**工作区无关的全局存储**（`dataDir`），
  多会话共享同一批 Deck；需要按项目隔离时配置不同的 `dataDir`。
- 导出使用本机字体；容器部署请通过 `fontFiles` 显式挂载中文字体并设置 `loadSystemFonts: false`。
