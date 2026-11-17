# ppt-master 的排版方案拆解，以及我们要补的三件事

> 调研对象：[`example/ppt-master`](../example/ppt-master)（技能/工作流型 PPT 生成包）。
> 对照对象：本项目当前的生成管线（`plugins/branddeck/src/generate.js` 等）。
> 所有判断都附文件行号，便于逐条核对或反驳。本文第 6 节列出据此落地的改动。

## 0. 一句话结论

ppt-master 不是"提示词写得好"，而是**把"画什么"和"怎么画"拆成两层产物，各带一份可校验的契约**：

- **规划层**：`design_spec.md`（散文规格）+ `spec_lock.md`（schema 校验的硬锁）——决定每页的语义内容、阅读模式、节奏标签、版式原型绑定；
- **执行层**：把那层规格画成自包含 SVG，并无权删信息。

我们当前是"大纲 → 直接出几何"，中间**缺这一层 spec**。页面内容少、图表缺席、字号没体系，都源于此。

## 1. 图表/表格为什么不会缺席

不是靠"希望模型想起来画图"，而是四层机制串联。

### 1.1 按信息关系分流，不按源文件类型

三个互斥分支，先判关系再画：

| 信息模型 | 分支 |
| --- | --- |
| 值/日期/时长决定几何（柱长、点位置、弧度、气泡半径） | [`executor-chart.md`](../example/ppt-master/skills/ppt-master/references/executor-chart.md#L9-L13) |
| 行表头 × 列表头共同定位每个事实 | [`executor-table.md`](../example/ppt-master/skills/ppt-master/references/executor-table.md#L11-L19) |
| 定性顺序、分组、包含、因果、命名区域 | `executor-structure.md` |

两条硬规则值得直接抄：

- **物理网格不算表格语义**："a PowerPoint table object or rectangular drawing grid does not establish Table semantics" —— 反过来，源文件里是表格、但日期/时长驱动任务条位置，就得走 Gantt（chart 分支）。
- **判据是关系，不是载体**：`plan-core.md` 要求先读页面的 `Relationships`（order/link/parent/membership/contrast/overlap），再决定几何要不要承担它。

这堵住了"看到表格就排表格、看到列表就排列表"的偷懒路径——也正是我们现在的毛病。

### 1.2 词表 + 原型目录（最实体的一层）

```
templates/charts/   33 个图表 SVG 原型（column/line/donut/waterfall/funnel/heatmap/gantt/…）
templates/tables/    6 个表格原型（record/metric/comparison/feature/rating/hierarchical）
templates/layouts/  86 个版式原型
```

- 规划侧按 `chart/<key>` 引用，[`chart-vocabulary.md`](../example/ppt-master/skills/ppt-master/templates/charts/chart-vocabulary.md#L1-L12) 按**"编码了什么关系"**分类（变化趋势 / 类别比较与排序 / 目标与进度 / 构成 / 分布…），而不是按图形外观；
- 词表明确"**零选择也是合法的**"——它不做配额，只保证"想画图时有可选项"；
- 执行侧从原型改造，不是从零发明。

### 1.3 硬规则：数据拥有标记

> derive every quantitative mark from the authoritative values and one explicit scale — **Do not eyeball positions, preserve sample values from a catalog preview, or alter data to improve composition**
> — [`executor-chart.md`](../example/ppt-master/skills/ppt-master/references/executor-chart.md#L19)

配套的编码纪律：单色深浅方案（不用彩虹）、直接标注优先于图例、只在关键数据点用强调色、柱状必须从零起、非零基线必须显式声明、双轴只在同域同单位时合法。

### 1.4 生成后单独校准（闭环）

[`verify-charts.md`](../example/ppt-master/skills/ppt-master/workflows/stages/verify-charts.md#L11-L14) 是**独立阶段**：所有页画完、导出前，按 `charts_index.json` 的 key 逐类校准 plot area 与标记；`svg_position_calculator.py` 只在此时跑，初稿阶段禁止使用。复合图表也不放过：能换算标 `decomposable-calc`，不能换算标 `manual-verify`，**不许静默跳过**。

一句话：**图表是被"规划+词表+校准"三件套保证的，不是被提示词祝福的。**

## 2. 每页信息量为什么够

不是"每页至少 N 字"，是三份约束叠出来的。

### 2.1 阅读模式决定页语法

[`plan-core.md`](../example/ppt-master/skills/ppt-master/references/plan-core.md#L44-L50)：

| 模式 | 页面语法 | 粒度 |
| --- | --- | --- |
| `text`（留档/审阅） | 完整判断句、短段落、图表、必要细节；列表只用于真正并列或有序的项 | 更少但更满的页，偏 `dense` |
| `balanced`（默认，商务） | 一个主判断 + 简明解释 + 结构化证据 | 中等粒度，节奏混合 |
| `presentation`（投屏） | 一页一个判断 + 大图或英雄数字；禁止段落切碎 | 更多但更疏的页，偏 `anchor`/`breathing` |

同一份源材料，`presentation` 与 `text` 两份 deck 必须"在页语法、页数、文字量、视觉负担、密度、节奏上都不同，而不只是字号不同"。

### 2.2 §IX 先写"完整措辞"，执行层无权删

> §IX owns each page's semantic content — **complete preferred wording** and block texture at `complete` depth
> Executor may: paraphrase, condense repetition, regroup… **never: add a claim, move content across pages, or drop information to fit a layout**
> — [`executor-base.md`](../example/ppt-master/skills/ppt-master/references/executor-base.md#L189-L197)

**页面内容少，是因为规划产物里只有四句话，不是排版把它删了。** 它把内容量锁在规划层；我们的提示词直接把"最终长度"当产出，所以每一环都在做减法。

还有一条反向规则专治我们的毛病：

> A list that is merely easier to lay out, or a template slot that expects a list, is **not a reason**: widen, reflow, or **drop the card** before converting prose to fill it
> — [`executor-base.md`](../example/ppt-master/skills/ppt-master/references/executor-base.md#L231)

### 2.3 节奏标签 + 版式结构参考

每页必须打一个 `page_rhythm` 标签（[`executor-base.md`](../example/ppt-master/skills/ppt-master/references/executor-base.md#L235-L243)）：

| 标签 | 排版纪律 |
| --- | --- |
| `anchor` | 结构页（封面/章节/目录/结尾） |
| `dense` | 信息密集：卡片网格、多栏、KPI 面板、表格、图表都允许（基线） |
| `breathing` | 低密度冲击页：裸文字、分隔、留白、满版图 |

并要求"数据密集页之后跟一个呼吸页"。

版式给的是**内容关系 → 结构 + 起步几何**对照表，并明确：

> **Repeating symmetric card grids without a page job is the failure mode** these structures exist to avoid
> — [`executor-base.md`](../example/ppt-master/skills/ppt-master/references/executor-base.md#L76-L86)

对应的结构清单：一页一主张 / 对等比较（对称切分或四象限）/ 主证据+结论（3:7 或 2:8 非对称）/ 并行序列（三栏、流程线、雪佛龙）/ 中心辐射 / 宽图+说明（视觉≥55%）。

## 3. 双语与来源：排版角色 + 事实溯源

### 3.1 双语是一条字号角色，不是生成能力

[`plan-core.md`](../example/ppt-master/skills/ppt-master/references/plan-core.md#L172-L184) 给的是**相对正文的比例表**：

| 复现角色 | 相对正文 |
| --- | ---: |
| 封面标题 / 单一主视觉 | 2.5–5× |
| 章节标题 | 2–2.5× |
| 页标题 / KPI 英雄数字 | 1.5–2× |
| 副标题 | 1.2–1.5× |
| Lead / 小标题 | 1.1–1.4× |
| 正文 | 1× |
| **双语页的第二语言行** | **0.75–0.9×** |
| 注释 | 0.7–0.85× |
| **脚注 / 页码** | **0.5–0.65×** |

配套规则：正文锚点取自规范刻度，**声明每个复现角色**（`lead` 至少正文大小、`footnote`、图表标注、英雄数字），落到干净的偶数 px（正文 24 → 标题 42、副标题 32、lead 30、注释 18、脚注 16）。

"看起来像设计过"不是花样多，而是**只有几种字号，且比例固定**。

### 3.2 来源脚注是硬契约

[`executor-base.md`](../example/ppt-master/skills/ppt-master/references/executor-base.md#L268)：

- 每个外部定量/事实必须有 `Fact IDs: F001…`，值从 `sources/*.facts.json` 取，**不许改**；
- 空间允许时渲染紧凑来源脚注（名称 + 短 URL/域名）；
- 虚构口径打 `Data class: scenario`，页面上必须印可见的「**情景数据**」标签；
- 来源没说过的框架/分组/标签，页面上标「**整理** / our reading」，不能冒充来源的结构。

## 4. 我们与它的差距（逐项）

| | ppt-master | 我们（改造前） |
| --- | --- | --- |
| 内容深度由谁决定 | 规划产物 §IX（完整措辞，锁定） | 大纲提示词，无锁定，每一环都在减 |
| 版式由谁决定 | 规划给关系，执行从 86 个原型里选 | 生成器规则推断（7 种），靠 `版式：` 提示 |
| 图表 | 33 词表 + 原型 + 校准阶段 | DSL 有 chart 元素，**无词表、无原型、无校准** |
| 字号/间距 | 比例阶梯 + 5 个间距锚点，声明式 | 少量绝对 px 常量，散落各处 |
| 来源/口径 | fact_id + 每页脚注 + 情景/整理 标签 | 无 |
| 校验 | 检查器（重叠 1px / 溢出 5% / 容量）+ 图表校准 | Zod + `repairDeck` 告警 |

**根因**：它多一层 spec（散文规格 + 硬锁），我们是从大纲直接到几何。

## 5. 能搬什么、不能搬什么

**能搬**

1. 信息模型路由（值驱动几何 / 行×列网格 / 拓扑关系三判据）；
2. **字号阶梯改比例制**（相对正文 0.5×–5×），替掉绝对 px；
3. `footnote` / `bilingual-line` / `hero-number` / `lead` 作为**声明式角色**；
4. 事实溯源：`facts` + 每页脚注 + `情景数据` / `整理` 标签；
5. `page_rhythm` 三档 + "密集页后跟呼吸页"；
6. **图表词表**：按信息关系选图，而不是等模型想起来。

**不能照搬**

- **LLM 手写绝对坐标**：它靠 422KB 检查器 + 内嵌字体度量表兜底；我们是确定性出几何，退回手写坐标会丢掉可复现性；
- **图表 SVG 原型**：我们的图表是 pptxgenjs **原生 Office 图表**（可编辑、可换数），等价物是"图表规格 + 数据契约"，不是 SVG 模板；
- **两轮人工确认 + 逐页授权**：它的质量成本是延迟，我们是单次生成；
- **它也没有硬密度校验**：密度是软约束（reading mode + rhythm），别指望它拒绝过满的页。

## 6. 据此落地的三件事（本次实现）

按"性价比 × 与现有架构的兼容性"排序。原则：**只搬能被确定性代码兑现的部分**，不搬需要人工确认或 LLM 手写几何的部分。

### 6.1 字号阶梯改比例制（对应 §3.1）

- 新增 `typeScale(style)`：以 `bodySize`（默认 22px）为锚点，按上表比例算出全部角色字号，**对齐到偶数 px**；
- 角色集合：`hero`（2.5–5×，取 2.6×）、`chapter`（2.2×）、`title`（1.9×）、`subtitle`（1.4×）、`lead`（1.3×）、`body`（1×）、`bilingual`（0.8×）、`annotation`（0.75×）、`footnote`（0.6×）；
- 现有 `TYPE.*` 常量改为从 `typeScale` 取值，保留向后兼容（原名字仍在）；
- 验收：任意 body 锚点下，所有角色字号落在比例带内且为偶数；`test/typescale-test.mjs`。

### 6.2 图表词表 + 按信息关系选图（对应 §1.1 / §1.2）

我们不走 SVG 原型，改为**数据契约 → 原生图表**：

- `extractChartData(bullets)`：识别 `标签 — 数值` 行（≥3 行、数值可解析、单位一致）→ `{labels, values, unit}`；
- 选择规则（按信息关系，不按外观）：
  - 类别比较 / 排序 → `bar`（水平条，首项强调）；
  - 时间序列（标签像年份/季度/月份且 ≥4 点）→ `line`；
  - 构成（含"占比/份额"且和≈100%）→ `pie`；
- 新增 `chart` 版式：图占主导（≥60% 宽），右侧给"结论 + 来源"；
- **数据不改写**：数值原样进图，单位统一才成图，解析不出就退回 `rankList`（不猜）；
- 验收：`test/chart-layout-test.mjs` 钉住"选图规则 + 数值不被改写 + 不可解析时不强行出图"。

### 6.3 每页脚注 + 双语行（对应 §3.2）

- 大纲新增 `来源：<名称 + 域名>` 元数据 → 渲染为字号 `footnote` 的页脚（小、次要色、不抢正文）；
- `来源` 缺省时不画（不编造）；
- 摘要页/数据页若带 `情景` 标记，则在图/卡旁打可见「情景数据」标签；
- 双语：`副标题：<英文>` 或大纲 `英文：` 行 → 字号 `bilingual`（0.8× body），紧跟主标题；
- 验收：`test/footnote-test.mjs`。

### 6.4 不做（明确记录）

- LLM 手写绝对坐标、SVG 图表原型、双轮人工确认——理由见 §5；
- 硬密度校验：继续用"超预算告警 + densify"而不是拒绝出页。

## 7. 与已有改造的关系

本次三件事与前面几轮是叠加关系，不是替代：

| 已做 | 本次叠加 |
| --- | --- |
| 模板视觉系统（背景图/内容框/脚注/Logo） | 字号阶梯统一比例，页面更"成套" |
| 版式库（columns/figures/rankList + kicker + 细线） | 新增 chart 版式，补上"数据页"这一格 |
| densify（薄页补细节） | 页脚/来源/bilingual 角色把它补出来的内容标好出处 |

## 8. 参考

- 本次对照的参考 deck：`example/bussiness.pptx`（15 页）、`example/technology.pptx`（12 页）——均为成品 deck（单一 Blank 版式），学到的是构图语法，不是母版结构；
- ppt-master 的关键文件：`references/executor-base.md`（页表达核心）、`references/executor-chart.md`、`references/executor-table.md`、`references/plan-core.md`（字号阶梯/阅读模式）、`workflows/stages/verify-charts.md`、`templates/charts/chart-vocabulary.md`。
