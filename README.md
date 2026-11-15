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

下表是该版本在**本机**跑出的真实结果（命令可复现，见[第 5 节](#5-验证记录)）。

| 维度 | 结果 |
| --- | --- |
| 自动化测试 | **8 套套件 / 147 项断言全部通过**（另有 1 个 CJK 冒烟脚本按退出码判定） |
| 真实模板解析 | `PPT模板-HW.pptx` → 与独立分析产物 28 个字段**逐项一致** |
| 活体插件（Host 半） | 面板 API `POST /branddeck-api/state` → **200 `ok:true`**，读到 6 页卡组 |
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

## 4. 一次真实运行的事实

用 `PPT模板-HW.pptx` 解析出的品牌令牌（与 `research/expected-tokens.json` 一致）：

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

## 5. 验证记录

环境：Windows 11 专业版 10.0.26100 · Node v24.21.0 · DSH desktop 0.2.1-rc.3 ·
jszip 3.10.2 / fast-xml-parser 5.12.1 / zod 4.6.7 / pptxgenjs 4.0.1 / @resvg/resvg-js 2.6.2

```bash
cd plugins/branddeck
node test/smoke-deps.mjs      #  5/5  依赖可用（含原生 resvg）
node test/smoke-cjk.mjs       #  退出码 0：CJK 在 resvg 中不乱码（设计文档坑点 1）
node test/schema-test.mjs     # 13/13 共享 DSL 契约
node test/engine-all.mjs      # 8 个文件 / 54 项：read 6 · theme 8 · media 5 · store 10 · svg 6 · compile 13 · e2e 6 · helpers
node test/client-smoke.mjs    # 45 项 浏览器半静态审计 + 沙箱求值/浅渲染
node test/api-test.mjs        # 13/13 端点表契约
node test/http-api-test.mjs   # 12/12 真实 node:http 上的传输契约
node test/e2e-hw.mjs          #  5/5  真实模板 → 令牌 → Deck → .pptx 全链路
```

**活体验证**（运行中的 DSH，端口 6625）：

| 操作 | 实际结果 |
| --- | --- |
| `POST /branddeck-api/state` | 200 `ok:true`，`decks=[hw-demo(6)]`，`templates=[ppt-hw]`，`primary=#C7000A`，`canvas=1280x720` |
| `POST /branddeck-api/decks.save` | 200 `ok:true`，6 页，`warnings=0` |
| `POST /branddeck-api/export` | 200 `ok:true`，落盘 `%DSH_HOME%/branddeck/exports/hw-demo.pptx`（121,611 B） |
| 客户端模块 | `/plugins/??@local/dsh-branddeck/client.js&rev=…` → 200，117,040 B |
| 启动崩溃日志 | 仅存修复前那条（22:45），修复后**无新增** |

**视觉核对**：导出件经 LibreOffice Kit 渲染 6 页（`missingFonts: []`）后逐页检查 ——
封面（红条 + YaHei 标题 + 右下 Logo + 页码）、三卡片页、原生柱状图页（品牌红双系列 + CJK 图例）、
SVG 四段流程页（3× 光栅化清晰）、致谢页，均无叠字、无越界、无缺字。

## 6. 开发中发现并修复的两个真实缺陷

1. **CJK 双影**：pptxgenjs 把同一个 `fontFace` 同时写进 `<a:latin>` 与 `<a:ea>`；
   `Arial` 没有中文字形，渲染器回退时把中文画两遍（截图可见重影）。
   修法：导出后处理，让 `<a:ea>` 用模板的中文字面（正文 → `<a:latin Arial/><a:ea Microsoft YaHei/>`）。
2. **启动崩溃 `web boot: 1 entry did not activate: @local/dsh-branddeck: loading`**：
   客户端把 `connection` 写进 `inject`，而该服务由内置插件用 `ctx.provide` 装在自己的 ctx 上，
   **兄弟插件不可见**，于是条目永久 pending、整个 web boot 失败。
   修法：`inject` 只声明 `['slots']`，数据改走同源 HTTP 路由 `POST /branddeck-api/<endpoint>`，
   并给 `apply` 加兜底（面板出问题不许炸启动）。

## 7. 已知边界

- 模板解析只取**主题/母版/版式级**品牌信息，不复刻模板的具体版式（"自由排版"路线的取舍）。
- SVG 按设计文档要求以位图嵌入，故 PowerPoint 内不可再编辑；文本/形状/图表均为原生对象。
- 面板根作用域拿不到会话 id，Deck 存储是**工作区无关的全局存储**（`$DSH_HOME/branddeck`）。
- 导出用本机字体；容器部署请配 `fontFiles` + `loadSystemFonts: false`。

## 8. 版本管理约定

- `.gitignore` 排除：`.tmp/` 测试产物、`node_modules/`、`.dsh-src/`（从 `app.asar` 解包的 413 MB 源码副本）、
  编译默认输出目录 `/branddeck/`、编辑器与系统垃圾。
- `.gitattributes`：`* text=auto eol=lf`，`*.pptx`/图片标为 binary，避免 Windows 上 CRLF 改写产生噪音 diff。
- 提交信息用 `<type>: <摘要>`（feat/fix/docs/test/chore），正文写清"改了什么 + 为什么"。
