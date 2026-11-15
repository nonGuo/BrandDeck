# BrandDeck 品牌演示

把企业 `.pptx` 模板解构成**设计令牌**，让大模型在 **1280×720 虚拟画布**上用卡片 DSL 自由排版，
在 DSH 侧栏面板里**所见即所得地微调**，最后编译成**原生可编辑**的 PowerPoint 文件。

本仓库是 [`设计文档.md`](设计文档.md) 的完整落地，产物是一个可安装进 DeepSeek Harness（DSH）的
bundle 插件 `plugins/branddeck`，并以真实企业模板 [`PPT模板-HW.pptx`](PPT模板-HW.pptx) 作为验证输入。

```
[企业 .pptx 模板] ──jszip + fast-xml-parser──▶ 设计令牌(色板/字体/Logo) + 真实摆位
        │
[会话 Agent] ──CardDeck DSL + Zod 强校验 + 坐标修复──▶ 1280×720 卡片
        │                                       ├──▶ DSH 侧栏面板：拖拽/改字/溢出预警
        └───────────────────────────────────────┴──▶ pptxgenjs + resvg ──▶ 原生 .pptx
```

## 当前版本状态

> **v0.1.0 · commit `e3ed1d3` · 2026-11-15 · 验证通过、可用**

下表是该版本在**本机**跑出的真实结果（命令可复现，见[第 7 节](#7-验证记录)）。

| 维度 | 结果 |
| --- | --- |
| 自动化测试 | **11 套套件 / 171 项断言全部通过**（另有 1 个 CJK 冒烟脚本按退出码判定） |
| 真实模板解析 | `PPT模板-HW.pptx` → 与独立分析产物 28 个字段**逐项一致** |
| 活体插件（Host 半） | 面板 API `state` / `outline.get` / `templates.spec` / `decks.save` 均 **200 `ok:true`** |
| 活体插件（Client 半） | 客户端模块下发 **HTTP 200 / 117 KB**，启动无崩溃日志 |
| 导出产物 | 6 页 `.pptx`，13.333×7.5in，原生图表 + 3 个媒体部件 |
| 视觉核对 | LibreOffice 渲染 6 页，CJK 单影、Logo 在位、无溢出（样例件：[`demo/HW品牌演示-BrandDeck样例.pptx`](demo/HW品牌演示-BrandDeck样例.pptx)） |

**未覆盖**：浏览器内的交互手感（拖拽/缩放的指针反馈、暗色主题观感）没有截图级验证 ——
我无法操作你的浏览器窗口；面板的数据通路与渲染崩溃风险已由 live API + 无崩溃日志 + 静态沙箱渲染测试覆盖。

## 1. 四大模块

| 设计文档模块 | 实现 | 关键点 |
| --- | --- | --- |
| 一、模板特征提取 | [`src/pptx/`](plugins/branddeck/src/pptx) + [`src/template.js`](plugins/branddeck/src/template.js) | `a:clrScheme` 兼容 `srgbClr` 与 `sysClr/@lastClr`；`a:fontScheme` 取 latin/ea/cs；母版 Logo 定位含**真实摆放坐标**；`sldSz` 真实尺寸 |
| 二、生成流水线 | [`index.js`](plugins/branddeck/index.js) 的 8 个 `branddeck_*` 工具 + 系统提示段 | 由**当前会话的 Agent** 担任"纯文本大模型"，插件负责契约、Zod 校验、坐标钳制、150 字预算 |
| 三、卡片交互视图 | [`client.js`](plugins/branddeck/client.js) | 16:9 等比视口、双击改字、拖拽/缩放/方向键、缩略图排序、实时溢出预警、600ms 防抖自动保存 |
| 四、PPTX 编译导出 | [`src/compile.js`](plugins/branddeck/src/compile.js) | `LAYOUT_WIDE`；文本/形状/图表→**原生对象**，SVG→`@resvg/resvg-js` 3× 光栅化嵌入 |

## 2. 目录结构

```
.
├── 设计文档.md                  # 需求与架构原稿（输入）
├── PPT模板-HW.pptx              # 真实企业模板（e2e 测试输入）
├── plugins/branddeck/           # ★ 插件本体（可独立安装）
│   ├── package.json             # bundle 清单：dsh.bundle.patch + dsh.client
│   ├── cordis.patch.yml         # 插入 Host 插件行
│   ├── index.js                 # Host 半：工具 / HTTP 路由 / 提示段
│   ├── client.js                # Web Client 半：16:9 编辑器面板（单文件、无构建）
│   ├── src/                     # 纯 Node 引擎：解析 / 契约 / 存储 / 光栅化 / 编译
│   ├── test/                    # 8 套自动化测试
│   ├── ARCHITECTURE.md          # 实现契约（模块签名、HTTP 协议、文件归属）
│   └── README.md                # 插件使用手册
├── research/                    # 模板事实取证 + DSH API 备忘（测试引用其中 expected-tokens.json）
├── scripts/                     # probe.mjs（minified 源码检索）、seed-demo.mjs（灌示例数据）
└── demo/                        # 已编译的样例产物
```

## 3. 快速开始

```bash
# 1) 装插件依赖（含 @resvg/resvg-js 的 win32-x64 预编译原生包）
cd plugins/branddeck && pnpm install

# 2) 安装进 DSH profile（绝对路径）
#    plugin_manager install_bundle { "target": "<repo>/plugins/branddeck" }

# 3) 重启 DeepSeek Harness（Host 半的 JS 需要新的模块生成），再刷新页面
```

装好后：左侧栏面板图标 → **BrandDeck**；或直接用工具：

```
branddeck_template_parse { path: "PPT模板-HW.pptx" }
branddeck_deck_create    { title: "季度汇报", templateId: "ppt-hw" }
branddeck_deck_save      { deck: {...} }        # 返回 warnings，按提示收敛
branddeck_deck_export    { deckId: "季报" }     # → 原生 .pptx
```

配置见 [`plugins/branddeck/README.md`](plugins/branddeck/README.md#5-配置cordis-patchyml-的-row-config)
（`dataDir`、`promptSection`、`svgRasterScale`、`fontFiles`、`logoSide` 等）。

## 4. 中间产物：同一份卡组，两种可编辑版本

生成很少一次到位，所以每份卡组都以**目录**形式保存两套真相 + 历史，随时可回去改：

```
$DSH_HOME/branddeck/decks/<deckId>/
├── card.json          # 「卡片视图」真相：几何 + 样式 + 内容（面板读写的就是它）
├── outline.md         # 「文字版本」真相：只读文案，改完按元素 id 回写
├── meta.json          # 修订号、时间戳、历史索引
└── revisions/
    ├── 0001-card.json     # 每次保存留档（默认保留最近 20 版）
    └── 0001-outline.md
```

- **文字 ↔ 卡片 无痛互转**：`outline.md` 里 `### 标题 {#cover-title}` 这种 `{#id}` 就是锚点 ——
  改文字保存后，Host 按 id 把内容写回同一元素，**位置与样式一字不动**；新增的 `### 正文 {#new-1}`
  会自动排版进安全区，不压到已有文字上。
- **面板内两种视图**：右上角 `卡片视图 / 文字视图` 切换。文字视图就是一个 Markdown 文本框 +
  「保存并应用到卡片」，应用后立即回到卡片视图看到效果。
- **可交付可 diff**：`branddeck_deck_materialize` 把 `card.json + outline.md + meta.json`
  写进工作区任意目录，于是**文案与版式都能进 git、能 review、能回滚**。
- **工具面**：`branddeck_outline_get` / `branddeck_outline_save` 让 Agent 也走同一条路，
  改文案不必重发整份 JSON。

## 5. 模板：解析一次，格式化后长期复用

模板不是每次都重新解析，而是**一次性物化**为四份可复用产物：

```
$DSH_HOME/branddeck/templates/<templateId>/
├── theme.json       # 设计令牌：色板（语义化）、字体、Logo 引用
├── layouts.json     # 版式清单：母版/版式/占位符几何（英寸），可直接读
├── template.md      # 人类可读品牌规格书：色板表、字体、页面、版式清单、素材、复用方式
├── source.pptx      # 原件副本（源文件被移动/删除也能重解析）
└── logo.png
```

- `branddeck_template_spec` 直接返回 `template.md` + `layouts.json`；
  老模板若缺这两样，会自动从 `source.pptx` 副本补齐。
- 下一轮只需 `branddeck_deck_create { templateId }`，**不再需要原始 .pptx**。
- 规格书是给人看的：`template.md` 里色板带语义和用途、版式带每个占位符的 `x,y,w,h`。
  以 `PPT模板-HW.pptx` 为例，它记录了 4 母版 / 7 版式（攀登、智能、探索、灯塔、1_Contents page、
  1_Chinese text page、End page）与 11 个占位符的几何。

## 6. 一次真实运行的事实

| 项 | 值 |
| --- | --- |
| 主色 / 强调色 | `#C7000A` / `#E9002F`（华为红） |
| 背景 / 容器 | `#FFFFFF` / `#EBEBEB` |
| 文字主 / 次 | `#1D1D1A` / `#666666` |
| 图表色板 | `#C7000A #E9002F #F4A100 #FFFF00 #232323 #666666` |
| 字体 | 标题 `Microsoft YaHei`，正文 `Arial` |
| Logo | `ppt/media/image1.png`，529×116 px，**右下角**（x 10.07 / y 6.54 in） |
| 页面尺寸 | cx 12196763 × cy 6858000 EMU = 13.3385 × 7.5 in（16:9） |

这份模板有几个会把朴素解析器带偏的坑，已逐一处理并在代码注释里写明理由：
`dk1 == lt1`（背景反相，白底来自显式 `<p:bg>`）、非标准 `sldSz`、Logo 被两个母版引用且在下侧、
`<a:ea>` 为空串需回退母版 `txStyles`。

## 7. 验证记录

环境：Windows 11 专业版 10.0.26100 · Node v24.21.0 · DSH desktop 0.2.1-rc.3 ·
jszip 3.10.2 / fast-xml-parser 5.12.1 / zod 4.6.7 / pptxgenjs 4.0.1 / @resvg/resvg-js 2.6.2

```bash
cd plugins/branddeck
node test/smoke-deps.mjs          #  5/5   依赖可用（含原生 resvg）
node test/smoke-cjk.mjs           #  退出码 0：CJK 在 resvg 中不乱码（设计文档坑点 1）
node test/schema-test.mjs         # 13/13  共享 DSL 契约
node test/engine-all.mjs          # 8 文件 / 54 项：read 6 · theme 8 · media 5 · store 10 · svg 6 · compile 13 · e2e 6 · helpers
node test/outline-test.mjs        # 10/10  文字↔卡片：整份往返、局部改写、自动排版、prune、备注、字节稳定
node test/template-doc-test.mjs   #  6/6   版式清单与模板规格书（真实模板、确定性、健壮性）
node test/client-smoke.mjs        # 45 项   浏览器半静态审计 + 沙箱求值/浅渲染
node test/api-test.mjs            # 13/13  端点表契约
node test/http-api-test.mjs       # 12/12  真实 node:http 上的传输契约
node test/deck-bundle-test.mjs    #  8/8   中间产物：card/outline/meta/修订、旧库迁移、materialize、模板四件套
node test/e2e-hw.mjs              #  5/5   真实模板 → 令牌 → Deck → .pptx 全链路
```

合计 **11 套 / 171 项断言全部通过**（另有 1 个 CJK 冒烟脚本按退出码判定）。

**活体验证**（运行中的 DSH，端口 6625）：

| 操作 | 实际结果 |
| --- | --- |
| `POST /branddeck-api/state` | 200 `ok:true`，`decks=[hw-demo(6)]`，`templates=[ppt-hw]`，`primary=#C7000A`，`canvas=1280x720` |
| `POST /branddeck-api/decks.save` | 200 `ok:true`，6 页，`warnings=0` |
| `POST /branddeck-api/outline.get` | 200 `ok:true`，返回 1,839 字符的 Markdown 文字版 |
| `POST /branddeck-api/templates.spec` | 200 `ok:true`，`template.md` 2,958 字符 + 4 个母版的版式清单 |
| `POST /branddeck-api/export` | 200 `ok:true`，落盘 `%DSH_HOME%/branddeck/exports/hw-demo.pptx`（121,611 B） |
| 客户端模块 | `/plugins/??@local/dsh-branddeck/client.js&rev=…` → 200 |
| 启动崩溃日志 | 仅存修复前那条（22:45），修复后**无新增** |

**视觉核对**：导出件经 LibreOffice Kit 渲染 6 页（`missingFonts: []`）后逐页检查 ——
封面（红条 + YaHei 标题 + 右下 Logo + 页码）、三卡片页、原生柱状图页（品牌红双系列 + CJK 图例）、
SVG 四段流程页（3× 光栅化清晰）、致谢页，均无叠字、无越界、无缺字。

## 8. 开发中发现并修复的两个真实缺陷

1. **CJK 双影**：pptxgenjs 把同一个 `fontFace` 同时写进 `<a:latin>` 与 `<a:ea>`；
   `Arial` 没有中文字形，渲染器回退时把中文画两遍（截图可见重影）。
   修法：导出后处理，让 `<a:ea>` 用模板的中文字面（正文 → `<a:latin Arial/><a:ea Microsoft YaHei/>`）。
2. **启动崩溃 `web boot: 1 entry did not activate: @local/dsh-branddeck: loading`**：
   客户端把 `connection` 写进 `inject`，而该服务由内置插件用 `ctx.provide` 装在自己的 ctx 上，
   **兄弟插件不可见**，于是条目永久 pending、整个 web boot 失败。
   修法：`inject` 只声明 `['slots']`，数据改走同源 HTTP 路由 `POST /branddeck-api/<endpoint>`，
   并给 `apply` 加兜底（面板出问题不许炸启动）。

## 9. 已知边界

- 模板解析只取**主题/母版/版式级**品牌信息，不复刻模板的具体版式（"自由排版"路线的取舍）。
- SVG 按设计文档要求以位图嵌入，故 PowerPoint 内不可再编辑；文本/形状/图表均为原生对象。
- 面板根作用域拿不到会话 id，Deck 存储是**工作区无关的全局存储**（`$DSH_HOME/branddeck`）。
- 导出用本机字体；容器部署请配 `fontFiles` + `loadSystemFonts: false`。

## 10. 版本管理约定

- `.gitignore` 排除：`.tmp/` 测试产物、`node_modules/`、`.dsh-src/`（从 `app.asar` 解包的 413 MB 源码副本）、
  编译默认输出目录 `/branddeck/`、编辑器与系统垃圾。
- `.gitattributes`：`* text=auto eol=lf`，`*.pptx`/图片标为 binary，避免 Windows 上 CRLF 改写产生噪音 diff。
- 提交信息用 `<type>: <摘要>`（feat/fix/docs/test/chore），正文写清"改了什么 + 为什么"。
