/**
 * Tests for src/generate.js — deterministic text+style -> CardDeck generation.
 *
 * Run:
 *   C:\Users\Administrator\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe test/generate-test.mjs
 *
 * Coverage:
 * 1. parseSourceToOutline on realistic Chinese Markdown (headings/bullets/notes/paragraph)
 * 2. plain prose (no headings) + empty/odd input, never throws
 * 3. composeDeck is schema-valid and repair-clean for 3 densities x 2 layouts x 3 tones
 * 4. no overlapping elements, everything inside innerBox, logo reserve clear of text,
 *    and (regression pin) cards are sized to their content, not to a fixed band
 * 5. maxSlides truncation + 40-bullet continuation slides ending in （续）
 * 6. determinism (identical JSON) and no mutation of inputs
 * 7. style.colors.primary actually changes the rendered accent colour
 * 8. layout intents + two-track text:
 *    8a 版式： lines route to all 5 kinds (zh/en, as a bullet, mid-section)
 *    8b inference picks compare/timeline/stats/bullets and never overrides an explicit kind
 *    8c compare/timeline/stats render with 0 repair warnings and 0 overlaps (3 densities x 3 tones)
 *    8d stats number extraction (2/3/4 cards) + timeline 2 nodes and 6 → +N collapse
 *    8e long bullets move their detail into speakerNotes, authored notes pass through verbatim
 */
import {
  DEFAULT_STYLE,
  DETAIL_TO_NOTES_WARNING,
  MAX_BULLET_CHARS,
  composeDeck,
  parseSourceToOutline,
} from '../src/generate.js'
import {
  CANVAS,
  CardDeckSchema,
  LOGO_RESERVE,
  SAFE_AREA,
  estimateLines,
  innerBox,
  repairDeck,
} from '../src/schema.js'

/* ------------------------------------------------------------------ *
 * Harness
 * ------------------------------------------------------------------ */

const results = []
function check(name, fn) {
  try {
    const detail = fn()
    results.push({ name, ok: true })
    console.log(`PASS ${name}${detail === undefined ? '' : ` — ${detail}`}`)
  } catch (error) {
    results.push({ name, ok: false })
    console.log(`FAIL ${name}: ${error?.message ?? error}`)
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message)
}
const stable = (value) => JSON.stringify(value)
const clone = (value) => structuredClone(value)
const jsonRoundTrip = (value) => JSON.parse(JSON.stringify(value))

/* ------------------------------------------------------------------ *
 * Fixtures
 * ------------------------------------------------------------------ */

/** Real corporate theme incl. a logo, so the logo reserve is actually active. */
const HW_THEME = {
  templateId: 'ppt-hw',
  templateName: 'PPT模板-HW',
  colors: {
    background: '#FFFFFF',
    surface: '#EBEBEB',
    textPrimary: '#1D1D1A',
    textSecondary: '#666666',
    primary: '#C7000A',
    accent: '#E9002F',
    chartPalette: ['#C7000A', '#E9002F', '#F4A100', '#FFFF00', '#232323', '#666666'],
  },
  fonts: { title: 'Microsoft YaHei', body: 'Arial', latin: 'Arial' },
  brandAssets: { logoPath: 'D:/tmp/logo.png', logoWidthPx: 529, logoHeightPx: 116, icons: {} },
}

/** Same brand, no logo: `innerBox(theme).logo` is then null. */
const NO_LOGO_THEME = { ...clone(HW_THEME), templateId: 'ppt-hw-nologo', brandAssets: { icons: {} } }

const MARKDOWN_SAMPLE = `# 品牌内容中台 2026 年度规划

> 用一套模板，把内容生产交给流程

## 01 业务挑战

- 内容需求增长 3 倍，但排版与合规返工消耗了 60% 的交付时间。
- 各团队自建模板，品牌色与字体漂移，评审前需要人工逐页修版。
- 每次交付都从零开始，资产难以沉淀，同一个卖点被重复写五遍。

一份来自内容运营团队的复盘显示，返工主要集中在版式溢出与字号错位两类问题上，
而这两类问题都可以在生成阶段被机械地消灭。

备注：先讲现象，再讲根因，最后给出路径。

## 02 增长数据

1. 内容交付量同比增长 138%
2. 人工返工工时下降 74%
3. 单页产出效率提升 3.2 倍

## 03 落地路径

## 04 结语

- 模板只解析一次，卡组可以复用到下一次汇报。
`

/** No headings at all: 7 paragraphs (one hard-wrapped over 3 lines). */
const PROSE_SAMPLE = `品牌内容中台把模板解析、卡组编排、内容生成与原生导出串成一条链路。

第二个段落说明这条链路如何降低返工率。

第三段讲成本：一次解析，长期复用，边际成本接近零。

第四段讨论评审流程，稿件直接是可编辑的 pptx，不需要二次排版。

第五段的例子来自市场部的一次季度汇报，
这次汇报原本需要三个人三天，
最终只用了半小时完成初稿。

第六段给出结论。

第七段留一个问题给听众。`

/**
 * Short outline used for the 18 style combinations: 6 short bullets (so the
 * per-slide 150-character budget of `repairDeck` is never exceeded, which lets the
 * test demand zero warnings rather than just zero *geometry* warnings).
 */
const SHORT_OUTLINE = {
  title: '组合测试',
  subtitle: '覆盖密度 · 版式 · 语气',
  slides: [
    {
      title: '要点页',
      bullets: ['模板只解析一次', '卡组可以复用', '品牌色统一', '字体不再漂移', '评审即定稿', '导出可编辑'],
      notes: '六个短要点，用于覆盖 dense 的 6 条上限。',
    },
    { title: '章节页' },
    { title: '第二页', bullets: ['一页一个观点', '留白优先', '层级清晰'] },
  ],
}

/**
 * Stress outline: 90-character bullets, long titles, a section page. Used to prove
 * the layout still stays inside the safe box when copy is at the length limit.
 */
const STRESS_OUTLINE = {
  title: '压力测试：极长的封面标题也要留在安全区内不被裁切也不越界',
  slides: [
    {
      title: '这是一个接近两行上限的页面标题，用来压缩内容区域的可支配高度',
      bullets: [
        '内容需求增长三倍但排版与合规返工消耗了六成交付时间跨部门协作成本被反复支付每一页都要重排。',
        '各团队自建模板导致品牌色与字体漂移评审前需要人工逐页修版同一套视觉规范被重复解释五遍以上。',
        '每次交付都从零开始资产难以沉淀同一个卖点被重复写五遍而每次写作都要重新确认口径与合规边界。',
        '一次性沉淀品牌资产减少重复沟通成本把模板解析结果变成可检索可复用的卡组并在导出时保持原生可编辑。',
        '评审阶段直接给出可编辑的 pptx避免二次排版让业务方在自己的电脑上就能改字改图改配色而不破坏版式。',
        '所有改动都以卡片为单位迭代版式由生成器保证合规因此设计师只需要维护模板而不需要逐页修版。',
      ],
    },
    { title: '一个只有标题的章节页，用来验证章节版式' },
  ],
}

/**
 * Mixed-copy outline for the card regression pin: one slide pairs 1-line bullets
 * with a 3-line bullet, so content-sized rows must differ from each other.
 */
const CARDS_OUTLINE = {
  title: '卡片尺寸测试',
  slides: [
    {
      title: '混合长度',
      bullets: [
        '模板只解析一次',
        '卡组可以复用到下一次汇报',
        '内容中台把模板解析、卡组编排与原生导出串成一条流水线，评审阶段直接给出可编辑的 pptx',
        '品牌色与字体不再漂移',
      ],
    },
    { title: '全部短文案', bullets: ['第一点', '第二点', '第三点', '第四点'] },
  ],
}

const GEOMETRY_WARNING = /clamped|logo corner|duplicate element id/

/* ------------------------------------------------------------------ *
 * Geometry helpers
 * ------------------------------------------------------------------ */

const intersects = (a, b) => a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y

const contains = (outer, inner) =>
  inner.x >= outer.x &&
  inner.y >= outer.y &&
  inner.x + inner.w <= outer.x + outer.w &&
  inner.y + inner.h <= outer.y + outer.h

const textElements = (slide) => slide.elements.filter((element) => element.type === 'text')
const shapeElements = (slide) => slide.elements.filter((element) => element.type === 'shape')

/** Top-right corner the logo occupies: no text may ever be drawn there. */
const LOGO_RECT = {
  x: CANVAS.width - SAFE_AREA.right - LOGO_RESERVE.width,
  y: SAFE_AREA.top,
  w: LOGO_RESERVE.width,
  h: LOGO_RESERVE.height,
}

/**
 * Collect every geometry violation on one slide.
 * @param {any} slide
 * @param {any} theme
 * @param {string[]} errors
 */
function collectGeometryErrors(slide, theme, errors) {
  const label = slide.slideId
  const box = innerBox(theme)
  const texts = textElements(slide)
  const shapes = shapeElements(slide)

  for (const element of slide.elements) {
    if (!contains(box, element)) {
      errors.push(`${label} · "${element.id}" outside innerBox (${element.x},${element.y} ${element.w}x${element.h})`)
    }
  }

  for (let i = 0; i < texts.length; i += 1) {
    if (intersects(texts[i], LOGO_RECT)) {
      errors.push(`${label} · text "${texts[i].id}" enters the logo reserve`)
    }
    for (let j = i + 1; j < texts.length; j += 1) {
      if (intersects(texts[i], texts[j])) {
        errors.push(`${label} · texts "${texts[i].id}" and "${texts[j].id}" overlap`)
      }
    }
  }

  for (let i = 0; i < shapes.length; i += 1) {
    for (let j = i + 1; j < shapes.length; j += 1) {
      if (!intersects(shapes[i], shapes[j])) continue
      // Nesting is the house pattern (card + accent tick), crossing is a bug.
      if (contains(shapes[i], shapes[j]) || contains(shapes[j], shapes[i])) continue
      errors.push(`${label} · shapes "${shapes[i].id}" and "${shapes[j].id}" overlap`)
    }
  }

  for (const text of texts) {
    for (const shape of shapes) {
      if (!intersects(text, shape)) continue
      // Text may sit ON a card; it may never hang over its edge.
      if (contains(shape, text)) continue
      errors.push(`${label} · text "${text.id}" straddles shape "${shape.id}"`)
    }
  }

  const ids = new Set()
  for (const element of slide.elements) {
    if (ids.has(element.id)) errors.push(`${label} · duplicate element id "${element.id}"`)
    ids.add(element.id)
  }
}

/**
 * Regression pin for the "cards are mostly empty" defect: a card is sized to its
 * content, so `cardH - textH` must stay within [2*padding, 2*padding + 40].
 * 2*28 = 56 is the padding floor (28 top + 28 bottom, the `formal` geometry); the
 * `bold` tone adds a 52px accent-tick header, staying inside +40 of the floor.
 *
 * @param {any} slide
 * @param {string[]} errors
 * @returns {number[]} card heights, for the "content-sized, not uniform" check
 */
function collectCardErrors(slide, errors) {
  const cards = shapeElements(slide).filter((element) => element.shapeType === 'roundRect')
  const bodies = textElements(slide).filter((element) => element.role === 'body')
  const heights = []
  for (const card of cards) {
    heights.push(card.h)
    const inside = bodies.filter((body) => contains(card, body))
    if (inside.length !== 1) {
      errors.push(`${slide.slideId} · card "${card.id}" contains ${inside.length} text elements (expected 1)`)
      continue
    }
    const text = inside[0]
    const delta = card.h - text.h
    if (delta < 56) {
      errors.push(`${slide.slideId} · card "${card.id}" too tight: h=${card.h}, text h=${text.h}`)
    }
    if (delta > 56 + 40) {
      errors.push(`${slide.slideId} · card "${card.id}" has ${delta - 56}px of empty space below its text`)
    }
    const needed = estimateLines(text.content, text.w, text.style.fontSize).heightPx
    if (needed > text.h + 1) {
      errors.push(`${slide.slideId} · text "${text.id}" needs ${Math.ceil(needed)}px but got ${text.h}px`)
    }
  }
  return heights
}

/* ------------------------------------------------------------------ *
 * 1. Realistic Chinese Markdown
 * ------------------------------------------------------------------ */

check('1 · parseSourceToOutline: Chinese Markdown (headings, bullets, notes, paragraph)', () => {
  const outline = parseSourceToOutline(MARKDOWN_SAMPLE)
  console.log('    parsed outline:')
  console.log(`      title    : ${outline.title}`)
  console.log(`      subtitle : ${outline.subtitle}`)
  outline.slides.forEach((slide, index) => {
    const kind = slide.kind === 'section' ? ' [section]' : ''
    console.log(`      slide ${index + 1}: ${slide.title}${kind} — ${slide.bullets.length} bullets`)
    slide.bullets.forEach((bullet) => console.log(`        · ${bullet}`))
    if (slide.notes !== undefined) console.log(`        notes: ${slide.notes}`)
  })

  assert(outline.title === '品牌内容中台 2026 年度规划', `title was "${outline.title}"`)
  assert(outline.subtitle === '用一套模板，把内容生产交给流程', `subtitle was "${outline.subtitle}"`)
  assert(outline.slides.length === 4, `expected 4 slides, got ${outline.slides.length}`)
  assert(
    stable(outline.slides.map((slide) => slide.title)) ===
      stable(['01 业务挑战', '02 增长数据', '03 落地路径', '04 结语']),
    `titles were ${stable(outline.slides.map((slide) => slide.title))}`,
  )

  const challenge = outline.slides[0]
  assert(challenge.bullets.length === 4, `expected 3 bullets + 1 paragraph, got ${challenge.bullets.length}`)
  assert(challenge.notes === '先讲现象，再讲根因，最后给出路径。', `notes were "${challenge.notes}"`)
  // The hard-wrapped paragraph must become ONE bullet, joined CJK-aware.
  assert(
    challenge.bullets[3].includes('两类问题上，而这两类问题都可以'),
    `hard-wrapped paragraph not joined: "${challenge.bullets[3]}"`,
  )
  assert(!challenge.bullets[3].includes('， 而'), 'CJK hard-wrap join inserted a space')
  assert(
    outline.slides[1].bullets[0] === '内容交付量同比增长 138%',
    `numbered bullet not stripped: "${outline.slides[1].bullets[0]}"`,
  )
  assert(outline.slides[2].kind === 'section', 'heading without bullets must be kind:"section"')
  assert(outline.slides[2].bullets.length === 0, 'section slide must have no bullets')
  assert(outline.slides[3].kind === undefined, 'a bullet slide must not be marked section')
  return `${outline.slides.length} slides, ${outline.slides.reduce((n, s) => n + s.bullets.length, 0)} bullets, 1 notes block`
})

/* ------------------------------------------------------------------ *
 * 2. Plain prose + empty/odd input
 * ------------------------------------------------------------------ */

check('2a · parseSourceToOutline: plain prose without headings', () => {
  const outline = parseSourceToOutline(PROSE_SAMPLE)
  console.log(`    derived title: ${outline.title}`)
  console.log(`    slides: ${outline.slides.map((slide) => `${slide.title}(${slide.bullets.length})`).join(' | ')}`)

  assert(outline.slides.length === 2, `7 paragraphs should chunk into 2 slides of <=5, got ${outline.slides.length}`)
  assert(outline.slides[0].bullets.length === 5, `first chunk should hold 5 bullets, got ${outline.slides[0].bullets.length}`)
  assert(outline.slides[1].bullets.length === 2, `second chunk should hold 2 bullets, got ${outline.slides[1].bullets.length}`)
  // The 3-line paragraph is one bullet.
  assert(
    outline.slides[0].bullets[4].includes('原本需要三个人三天，最终只用了半小时完成初稿'),
    `hard-wrapped 3-line paragraph not joined: "${outline.slides[0].bullets[4]}"`,
  )
  assert(outline.title.endsWith('…'), `long derived title should end with an ellipsis: "${outline.title}"`)
  assert(Array.from(outline.title).length <= 25, `derived title too long: ${Array.from(outline.title).length}`)
  assert(outline.title.startsWith('品牌内容中台把模板解析'), `derived title was "${outline.title}"`)
  assert(outline.slides[1].title !== outline.slides[0].title, 'chunk titles must stay distinguishable')
  return `${outline.slides.length} slides, title "${outline.title}"`
})

check('2b · parseSourceToOutline: empty / odd input never throws', () => {
  const inputs = ['', '   \n\t\n  ', null, undefined, 42, [], {}, '###', '- \n- \n', '# 只有一个标题\n']
  for (const input of inputs) {
    const outline = parseSourceToOutline(input)
    assert(typeof outline.title === 'string' && outline.title.length > 0, `empty title for ${stable(input)}`)
    assert(Array.isArray(outline.slides), `missing slides for ${stable(input)}`)
  }
  const empty = parseSourceToOutline('')
  assert(empty.title === '未命名演示', `empty input title was "${empty.title}"`)
  assert(empty.slides.length === 0, `empty input produced ${empty.slides.length} slides`)
  assert(stable(parseSourceToOutline('   \n\n ')) === stable({ title: '未命名演示', slides: [] }), 'whitespace input not safe-default')
  const titleOnly = parseSourceToOutline('# 只有一个标题\n')
  assert(titleOnly.title === '只有一个标题', `title-only input lost the title: "${titleOnly.title}"`)
  assert(titleOnly.slides.length === 0, 'title-only input should have no slides')

  // Setext headings: `===` is level 1 (deck title), `---` is level 2 (slide).
  const setext = parseSourceToOutline('主标题\n=======\n\n## 二级\n- 要点一\n\n子标题\n-------\n- 要点二')
  assert(setext.title === '主标题', `setext title was "${setext.title}"`)
  assert(
    stable(setext.slides.map((slide) => slide.title)) === stable(['二级', '子标题']),
    `setext slides were ${stable(setext.slides.map((slide) => slide.title))}`,
  )

  // A horizontal rule is a separator, not a bullet and not a slide.
  const rules = parseSourceToOutline('## A\n- x\n\n---\n\n## B\n- y')
  assert(rules.slides.length === 2, `horizontal rule created ${rules.slides.length} slides`)
  assert(stable(rules.slides[0].bullets) === stable(['x']), `horizontal rule leaked into bullets: ${stable(rules.slides[0].bullets)}`)
  // English notes marker + inline Markdown stripping.
  const en = parseSourceToOutline('## 页\n- **粗体** 与 [链接](http://x) 和 `code`\nNotes: hello')
  assert(en.slides[0].notes === 'hello', `English notes were "${en.slides[0].notes}"`)
  assert(en.slides[0].bullets[0] === '粗体 与 链接 和 code', `inline markdown kept: "${en.slides[0].bullets[0]}"`)
  return `${inputs.length} odd inputs handled + setext/hr/inline-markdown cases`
})

/* ------------------------------------------------------------------ *
 * 3. composeDeck: schema-valid + repair-clean over 18 style combinations
 * ------------------------------------------------------------------ */

const DENSITIES = ['airy', 'standard', 'dense']
const LAYOUTS = ['list', 'cards']
const TONES = ['formal', 'bold', 'minimal']

check('2c · framing bullets before the first ## become the cover subtitle, not a phantom slide', () => {
  // Real model output shape (seen live): a thesis line written under the deck title
  // before any section heading. It must NOT become a page whose heading is that
  // sentence — that produced a one-bullet "（续）" slide in a rendered deck.
  const md = [
    '# BrandDeck 上手三步走',
    '',
    '- 三步完成一份符合品牌规范的演示物料',
    '- 全程不需要设计基础',
    '',
    '## 第一步：导入模板',
    '',
    '- 打开 BrandDeck',
    '- 上传现有 PPT',
  ].join('\n')
  const outline = parseSourceToOutline(md)
  assert(outline.title === 'BrandDeck 上手三步走', `title: ${outline.title}`)
  assert(outline.subtitle === '三步完成一份符合品牌规范的演示物料', `subtitle: ${outline.subtitle}`)
  assert(outline.slides.length === 1, `phantom lead slide: ${outline.slides.length} slides`)
  assert(outline.slides[0].title === '第一步：导入模板', `first slide: ${outline.slides[0].title}`)
  assert(outline.slides[0].bullets.length === 2, `first slide bullets: ${outline.slides[0].bullets.length}`)

  // Without a `#` title the old prose behaviour must survive: first paragraph is
  // the title and the rest chunk into slides.
  const prose = parseSourceToOutline('第一段是标题\n\n第二段\n\n第三段')
  assert(prose.title.startsWith('第一段'), `prose title: ${prose.title}`)
  assert(prose.slides.length >= 1, 'prose input produced no slides')
  return `subtitle="${outline.subtitle}", ${outline.slides.length} slide, prose kept`
})

check('8f · deckId is stable and meaningful (no ASCII dregs from a CJK title)', () => {
  // Regression: "自建模板 vs 统一模板库…" used to derive the id "vs", which names
  // nothing and collides easily. The slug must represent the title, or we fall back
  // to a deterministic digest.
  const cases = [
    ['自建模板 vs 统一模板库：差别与三步落地', /^deck-[a-z0-9]{5,}$/],
    ['BrandDeck 上手三步走', /^branddeck/],
    ['2026 Q3 Review', /^2026q3review$/],
    ['季度汇报', /^deck-[a-z0-9]{5,}$/],
  ]
  const seen = new Set()
  for (const [title, pattern] of cases) {
    const outline = parseSourceToOutline('#' + ' ' + title + '\n\n## 一页\n- 要点')
    const { deck } = composeDeck({ outline, theme: HW_THEME, style: {} })
    assert(pattern.test(deck.deckId), `"${title}" -> ${deck.deckId} (expected ${pattern})`)
    assert(/^[a-zA-Z0-9._-]{1,64}$/.test(deck.deckId), `unsafe id: ${deck.deckId}`)
    assert(!seen.has(deck.deckId), `duplicate id: ${deck.deckId}`)
    seen.add(deck.deckId)
  }
  // The same title always yields the same id (determinism).
  const a = composeDeck({ outline: parseSourceToOutline('# 季度汇报\n\n## 一页\n- 要点'), theme: HW_THEME, style: {} })
  const b = composeDeck({ outline: parseSourceToOutline('# 季度汇报\n\n## 一页\n- 要点'), theme: HW_THEME, style: {} })
  assert(a.deck.deckId === b.deck.deckId, 'deckId is not deterministic')
  return `${cases.length} titles, deterministic`
})

check('3a · composeDeck: 18 combinations are schema-valid and repair-clean', () => {
  let combinations = 0
  for (const density of DENSITIES) {
    for (const bulletLayout of LAYOUTS) {
      for (const tone of TONES) {
        const { deck, warnings } = composeDeck({ outline: SHORT_OUTLINE, theme: HW_THEME, style: { density, bulletLayout, tone } })
        const tag = `${density}/${bulletLayout}/${tone}`

        const parsed = CardDeckSchema.safeParse(deck)
        assert(parsed.success, `${tag}: schema rejected — ${JSON.stringify(parsed.error?.issues ?? [])}`)

        const repaired = repairDeck(deck, HW_THEME)
        const geometry = repaired.warnings.filter((warning) => GEOMETRY_WARNING.test(warning))
        assert(geometry.length === 0, `${tag}: repair warnings — ${geometry.join(' | ')}`)
        assert(repaired.warnings.length === 0, `${tag}: unexpected warnings — ${repaired.warnings.join(' | ')}`)
        assert(warnings.length === 0, `${tag}: compose warnings — ${warnings.join(' | ')}`)

        combinations += 1
      }
    }
  }
  assert(combinations === 18, `expected 18 combinations, ran ${combinations}`)
  return `${combinations}/18 combinations clean (0 schema errors, 0 repair warnings, 0 compose warnings)`
})

check('3b · composeDeck: stress copy (90-char bullets, long titles) stays repairable', () => {
  let combinations = 0
  let contentWarnings = 0
  for (const density of DENSITIES) {
    for (const bulletLayout of LAYOUTS) {
      for (const tone of TONES) {
        const { deck, warnings } = composeDeck({ outline: STRESS_OUTLINE, theme: NO_LOGO_THEME, style: { density, bulletLayout, tone } })
        const tag = `${density}/${bulletLayout}/${tone}`
        const parsed = CardDeckSchema.safeParse(deck)
        assert(parsed.success, `${tag}: schema rejected`)
        const repaired = repairDeck(deck, NO_LOGO_THEME)
        const geometry = repaired.warnings.filter((warning) => GEOMETRY_WARNING.test(warning))
        assert(geometry.length === 0, `${tag}: geometry warnings — ${geometry.join(' | ')}`)
        contentWarnings += warnings.filter((warning) => !GEOMETRY_WARNING.test(warning)).length
        combinations += 1
      }
    }
  }
  return `${combinations}/18 combinations, 0 geometry warnings (${contentWarnings} copy-length warnings reported)`
})

/* ------------------------------------------------------------------ *
 * 4. Geometry: no overlap, inside innerBox, logo reserve, content-sized cards
 * ------------------------------------------------------------------ */

check('4a · geometry: no overlap, inside innerBox, logo reserve clear (18 combinations)', () => {
  let slidesChecked = 0
  let elementsChecked = 0
  for (const density of DENSITIES) {
    for (const bulletLayout of LAYOUTS) {
      for (const tone of TONES) {
        const tag = `${density}/${bulletLayout}/${tone}`
        const { deck } = composeDeck({ outline: SHORT_OUTLINE, theme: HW_THEME, style: { density, bulletLayout, tone } })
        const errors = []
        deck.slides.forEach((slide, index) => {
          const expectedId = `s${index + 1}`
          if (slide.slideId !== expectedId) errors.push(`slide ${index + 1} has slideId "${slide.slideId}" (expected "${expectedId}")`)
          if (slide.elements.length === 0) errors.push(`${slide.slideId} has no elements`)
          collectGeometryErrors(slide, HW_THEME, errors)
          slidesChecked += 1
          elementsChecked += slide.elements.length
        })
        assert(errors.length === 0, `${tag}: ${errors.slice(0, 5).join(' | ')}`)
      }
    }
  }
  return `${slidesChecked} slides / ${elementsChecked} elements checked, 0 violations`
})

check('4b · geometry: same checks with a logo-less theme + stress copy', () => {
  let slidesChecked = 0
  for (const density of DENSITIES) {
    for (const bulletLayout of LAYOUTS) {
      for (const tone of TONES) {
        const tag = `${density}/${bulletLayout}/${tone}`
        const { deck } = composeDeck({ outline: STRESS_OUTLINE, theme: NO_LOGO_THEME, style: { density, bulletLayout, tone } })
        const errors = []
        deck.slides.forEach((slide) => {
          collectGeometryErrors(slide, NO_LOGO_THEME, errors)
          slidesChecked += 1
        })
        assert(errors.length === 0, `${tag}: ${errors.slice(0, 5).join(' | ')}`)
      }
    }
  }
  return `${slidesChecked} slides checked, 0 violations`
})

check('4c · cards: every card is sized to its content (not a fixed band)', () => {
  let cardsChecked = 0
  let mixedSlides = 0
  const distinct = []
  for (const density of DENSITIES) {
    for (const tone of ['formal', 'bold']) {
      const tag = `${density}/cards/${tone}`
      const { deck } = composeDeck({ outline: CARDS_OUTLINE, theme: HW_THEME, style: { density, bulletLayout: 'cards', tone } })
      const errors = []
      /** @type {number[]} */
      const heights = []
      for (const slide of deck.slides) {
        const slideHeights = collectCardErrors(slide, errors)
        heights.push(...slideHeights)
        if (new Set(slideHeights).size > 1) mixedSlides += 1
      }
      assert(errors.length === 0, `${tag}: ${errors.slice(0, 4).join(' | ')}`)
      cardsChecked += heights.length
      distinct.push(`${tag} → [${[...new Set(heights)].sort((a, b) => a - b).join(', ')}]`)
    }
  }
  console.log('    distinct card heights (content-sized, so rows differ):')
  for (const line of distinct) console.log(`      ${line}`)
  // 1-line copy must not produce the same card height as multi-line copy.
  assert(
    mixedSlides > 0,
    'card heights were uniform on every slide despite mixed copy length — cards look like fixed bands again',
  )
  return `${cardsChecked} cards within [text+56, text+96]px, ${mixedSlides} slides mix 1-line and multi-line rows`
})

/* ------------------------------------------------------------------ *
 * 5. maxSlides truncation + continuation slides
 * ------------------------------------------------------------------ */

check('5a · maxSlides truncates and warns', () => {
  // Truncation is about too many *pages*, so the fixture declares many pages: one page
  // with 40 bullets is a single page now (its overflow goes to the notes, see 5b).
  const slides = Array.from({ length: 12 }, (_, index) => ({
    title: `第 ${index + 1} 页`,
    bullets: ['第一条要点内容', '第二条要点内容', '第三条要点内容'],
  }))
  const outline = { title: '上限测试', slides }
  const limited = composeDeck({ outline, theme: HW_THEME, style: { maxSlides: 6 } })

  assert(limited.deck.slides.length <= 6, `maxSlides 6 produced ${limited.deck.slides.length} slides`)
  const truncation = limited.warnings.find((warning) => /^已按上限截断 \d+ 页$/.test(warning))
  assert(truncation !== undefined, `no truncation warning in ${stable(limited.warnings)}`)
  assert(CardDeckSchema.safeParse(limited.deck).success, 'truncated deck is not schema-valid')
  const geometry = repairDeck(limited.deck, HW_THEME).warnings.filter((warning) => GEOMETRY_WARNING.test(warning))
  assert(geometry.length === 0, `truncated deck has geometry warnings: ${geometry.join(' | ')}`)

  const cappedOnce = composeDeck({ outline, theme: HW_THEME, style: { maxSlides: 1 } })
  assert(cappedOnce.deck.slides.length === 1, `maxSlides 1 produced ${cappedOnce.deck.slides.length} slides`)
  assert(CardDeckSchema.safeParse(cappedOnce.deck).success, 'cover-only deck is not schema-valid')
  return `maxSlides 6 -> ${limited.deck.slides.length} slides, warning "${truncation}"; maxSlides 1 -> 1 slide`
})

check('5b · a page never grows a continuation page: overflow moves to the notes', () => {
  // The deck reads as complete pages, so a page can never be titled `<title>（续）`.
  // What does not fit stays in the speaker notes — nothing is lost.
  const bullets = Array.from({ length: 40 }, (_, index) => `第 ${index + 1} 条要点：内容中台把模板解析与原生导出串成一条流水线。`)
  const outline = { title: '续页测试', slides: [{ title: '长清单', bullets }] }
  const { deck, warnings } = composeDeck({ outline, theme: HW_THEME, style: { maxSlides: 60, density: 'standard', bulletLayout: 'list' } })

  const titles = deck.slides.map((slide) => slide.title)
  console.log(`    slide titles: ${titles.join(' | ')}`)
  const continuation = titles.filter((title) => typeof title === 'string' && /（续）$/.test(title))
  assert(continuation.length === 0, `continuation pages came back: ${stable(titles)}`)

  const content = deck.slides.find((slide) => slide.title === '长清单')
  assert(content !== undefined, 'the page disappeared instead of clamping')
  const shown = textElements(content).filter((element) => element.id.includes('-body-')).length
  assert(shown > 0 && shown < bullets.length, `expected a clamped page, showed ${shown}/${bullets.length}`)

  // The dropped copy must still be reachable: it is in the notes, not deleted.
  const notes = content.speakerNotes ?? ''
  const droppedBullet = bullets[bullets.length - 1]
  assert(notes.includes(droppedBullet), 'overflowing bullet was not preserved in the notes')
  assert(notes.includes('单页容量以外的要点'), `notes did not explain the clamp: "${notes.slice(0, 40)}"`)

  assert(CardDeckSchema.safeParse(deck).success, 'deck is not schema-valid')
  const geometry = repairDeck(deck, HW_THEME).warnings.filter((warning) => GEOMETRY_WARNING.test(warning))
  assert(geometry.length === 0, `clamped deck has geometry warnings: ${geometry.join(' | ')}`)
  assert(warnings.some((warning) => warning.includes('已并入演讲备注')), `no clamp warning in ${stable(warnings)}`)
  return `${deck.slides.length} slides, ${shown}/${bullets.length} bullets on the page, rest in notes`
})

/* ------------------------------------------------------------------ *
 * 6. Determinism + input immutability
 * ------------------------------------------------------------------ */

check('6a · determinism: composing twice yields identical JSON', () => {
  const outline = parseSourceToOutline(MARKDOWN_SAMPLE)
  const style = { density: 'dense', tone: 'bold', bulletLayout: 'cards', coverVariant: 'centered' }
  const first = composeDeck({ outline, theme: HW_THEME, style })
  const second = composeDeck({ outline, theme: HW_THEME, style })
  assert(stable(first.deck) === stable(second.deck), 'deck JSON differs between two identical calls')
  assert(stable(first.warnings) === stable(second.warnings), 'warnings differ between two identical calls')
  // And the deck survives a JSON round-trip unchanged: it is plain data.
  assert(stable(jsonRoundTrip(first.deck)) === stable(first.deck), 'deck is not plain JSON')
  return `${stable(first.deck).length} chars, byte-identical across runs`
})

check('6b · no mutation: inputs are untouched', () => {
  const outline = parseSourceToOutline(MARKDOWN_SAMPLE)
  const theme = clone(HW_THEME)
  const style = { density: 'airy', tone: 'minimal', bulletLayout: 'cards', colors: { primary: '#00A366' }, fonts: { title: 'HarmonyOS Sans' } }
  const before = { outline: stable(outline), theme: stable(theme), style: stable(style) }
  composeDeck({ outline, theme, style })
  assert(stable(outline) === before.outline, 'outline was mutated')
  assert(stable(theme) === before.theme, 'theme was mutated')
  assert(stable(style) === before.style, 'style was mutated')
  // DEFAULT_STYLE must stay pristine and frozen.
  assert(Object.isFrozen(DEFAULT_STYLE), 'DEFAULT_STYLE is not frozen')
  assert(DEFAULT_STYLE.density === 'standard' && DEFAULT_STYLE.maxSlides === 12, 'DEFAULT_STYLE was altered')
  return 'outline/theme/style unchanged'
})

/* ------------------------------------------------------------------ *
 * 7. Style colour overrides
 * ------------------------------------------------------------------ */

check('7 · style.colors.primary changes the rendered accent colour', () => {
  const base = composeDeck({ outline: SHORT_OUTLINE, theme: HW_THEME, style: { tone: 'formal' } })
  const overridden = composeDeck({ outline: SHORT_OUTLINE, theme: HW_THEME, style: { tone: 'formal', colors: { primary: '#00A366' } } })
  const tokenForm = composeDeck({ outline: SHORT_OUTLINE, theme: HW_THEME, style: { tone: 'formal', colors: { primary: '$colors.accent' } } })

  const baseJson = stable(base.deck)
  const overriddenJson = stable(overridden.deck)
  assert(baseJson.includes('#C7000A'), 'base deck does not carry the theme primary')
  assert(!baseJson.includes('#00A366'), 'override colour leaked into the base deck')
  assert(overriddenJson.includes('#00A366'), 'override colour is missing from the deck')
  assert(!overriddenJson.includes('#C7000A'), 'theme primary still rendered after the override')

  const accentFill = overridden.deck.slides[1].elements.find((element) => element.type === 'shape' && element.id.endsWith('-bar-1'))
  assert(accentFill !== undefined, 'no accent bar on the first bullet slide')
  assert(accentFill.style.fill === '#00A366', `accent bar fill was "${accentFill.style.fill}"`)

  // A `$colors.*` token in the override resolves against the theme.
  assert(stable(tokenForm.deck).includes('#E9002F'), '$colors.accent token did not resolve to the theme accent')
  return `#C7000A -> #00A366 on ${overridden.deck.slides[1].elements.filter((e) => e.type === 'shape').length} shapes, token form resolved to #E9002F`
})

/* ------------------------------------------------------------------ *
 * 8. Layout intents (版式) + two-track text
 * ------------------------------------------------------------------ */

/**
 * One Markdown slide carrying an explicit layout line, optionally with the line
 * written as a bullet and/or placed after the first bullet (the contract says
 * "anywhere inside a slide section").
 */
function layoutMarkdown(spelling, { asBullet = false, afterFirst = false, bullets = ['要点一', '要点二'] } = {}) {
  const line = `${asBullet ? '- ' : ''}版式：${spelling}`
  const body = afterFirst && bullets.length > 1
    ? [`- ${bullets[0]}`, line, ...bullets.slice(1).map((bullet) => `- ${bullet}`)]
    : [line, ...bullets.map((bullet) => `- ${bullet}`)]
  return `# 版式路由\n\n## 页面\n${body.join('\n')}\n`
}

check('8a · 版式： lines route to all 5 kinds (zh + en) and never become bullets', () => {
  const cases = [
    ['对比', 'compare'],
    ['compare', 'compare'],
    ['时间线', 'timeline'],
    ['timeline', 'timeline'],
    ['流程', 'timeline'],
    ['步骤', 'timeline'],
    ['process', 'timeline'],
    ['steps', 'timeline'],
    ['指标', 'stats'],
    ['数据', 'stats'],
    ['stat', 'stats'],
    ['stats', 'stats'],
    ['kpi', 'stats'],
    ['列表', 'bullets'],
    ['清单', 'bullets'],
    ['bullets', 'bullets'],
    ['list', 'bullets'],
  ]
  const seen = new Map()
  cases.forEach(([spelling, expected], index) => {
    // Cover the three shapes of the line across the table: plain, written as a
    // bullet, and written after the first bullet.
    const options = {
      asBullet: index % 3 === 1,
      afterFirst: index % 3 === 2,
    }
    const outline = parseSourceToOutline(layoutMarkdown(spelling, options))
    assert(outline.slides.length === 1, `${spelling}: ${outline.slides.length} slides`)
    const slide = outline.slides[0]
    assert(slide.kind === expected, `版式：${spelling} -> kind ${JSON.stringify(slide.kind)}, expected "${expected}"`)
    assert(
      stable(slide.bullets) === stable(['要点一', '要点二']),
      `版式：${spelling} leaked into bullets: ${stable(slide.bullets)}`,
    )
    seen.set(expected, (seen.get(expected) ?? 0) + 1)
  })

  // `section` is structural: it needs a slide with no bullets, otherwise the copy
  // would be dropped (and losing copy is not an acceptable "route").
  for (const spelling of ['章节', '过渡', 'section']) {
    const outline = parseSourceToOutline(`# 版式路由\n\n## 章节页\n版式：${spelling}\n`)
    assert(outline.slides[0].kind === 'section', `版式：${spelling} -> ${outline.slides[0].kind}`)
    assert(outline.slides[0].bullets.length === 0, 'section slide gained bullets')
  }
  // ... and an explicit `章节` on a slide WITH copy must not swallow the copy.
  const guarded = parseSourceToOutline(layoutMarkdown('章节'))
  assert(guarded.slides[0].bullets.length === 2, 'explicit section dropped the bullets')
  assert(guarded.slides[0].kind !== 'section', 'explicit section with copy stayed a section')

  // Half-width colon, English spelling, a trailing full stop and an unknown value.
  const halfWidth = parseSourceToOutline('## 页\n版式: 时间线\n- 第一步：a\n- 第二步：b\n- 第三步：c')
  assert(halfWidth.slides[0].kind === 'timeline', `版式: 时间线 -> ${halfWidth.slides[0].kind}`)
  const english = parseSourceToOutline('## 页\nlayout: compare\n- 旧：a\n- 新：b')
  assert(english.slides[0].kind === 'compare', `layout: compare -> ${english.slides[0].kind}`)
  const punctuation = parseSourceToOutline('## 页\n版式：指标。\n- 增长 138%\n- 下降 74%')
  assert(punctuation.slides[0].kind === 'stats', `版式：指标。 -> ${punctuation.slides[0].kind}`)
  const unknown = parseSourceToOutline('## 页\n版式：瀑布流\n- 要点一\n- 要点二')
  assert(unknown.slides[0].kind === undefined, `unknown layout value produced kind ${unknown.slides[0].kind}`)
  assert(
    stable(unknown.slides[0].bullets) === stable(['要点一', '要点二']),
    `unknown layout line leaked into bullets: ${stable(unknown.slides[0].bullets)}`,
  )

  const summary = [...seen.entries()].map(([kind, count]) => `${kind}:${count}`).join(' ')
  return `${cases.length} spellings → ${summary}; section via 3 spellings, unknown ignored`
})

check('8b · inference picks compare / timeline / stats / bullets (4 fixtures)', () => {
  const fixtures = [
    {
      label: 'compare (cue prefixes)',
      bullets: ['旧流程：从零开始排版', '新流程：模板一次解析'],
      kind: 'compare',
    },
    {
      label: 'compare (标签 pairs)',
      bullets: ['自建模板相比统一模板：交付慢三倍', '统一模板相比自建模板：交付快三倍'],
      kind: 'compare',
    },
    {
      label: 'timeline (ordered cues)',
      bullets: ['第一步：导入模板', '第二步：解析品牌', '第三步：生成卡组'],
      kind: 'timeline',
    },
    {
      label: 'stats (numeric + short)',
      bullets: ['内容交付量同比增长 138%', '人工返工工时下降 74%'],
      kind: 'stats',
    },
    {
      label: 'bullets (nothing matches)',
      bullets: ['模板只解析一次', '卡组可以复用到下一次汇报'],
      kind: undefined,
    },
  ]
  for (const fixture of fixtures) {
    const markdown = `# 推理\n\n## 页面\n${fixture.bullets.map((bullet) => `- ${bullet}`).join('\n')}\n`
    const slide = parseSourceToOutline(markdown).slides[0]
    console.log(`    ${fixture.label.padEnd(26)} → kind=${String(slide.kind)}  ${fixture.bullets.join(' / ')}`)
    assert(slide.kind === fixture.kind, `${fixture.label}: kind ${JSON.stringify(slide.kind)}, expected ${JSON.stringify(fixture.kind)}`)
  }

  // Timeline needs >= 3 cues: two steps stay a bullet list (conservative).
  const twoSteps = parseSourceToOutline('## 页\n- 第一步：导入模板\n- 第二步：生成卡组')
  assert(twoSteps.slides[0].kind === undefined, `2 ordered bullets became ${twoSteps.slides[0].kind}`)

  // A single contrast or a single number is not enough.
  const oneCue = parseSourceToOutline('## 页\n- 旧流程：从零开始排版\n- 模板只解析一次')
  assert(oneCue.slides[0].kind === undefined, `1 contrast bullet became ${oneCue.slides[0].kind}`)
  const oneNumber = parseSourceToOutline('## 页\n- 内容交付量同比增长 138%\n- 卡组可以复用')
  assert(oneNumber.slides[0].kind === undefined, `1 numeric bullet became ${oneNumber.slides[0].kind}`)

  // Long numeric prose is not a stat row (the <= 24-char rule).
  const longNumber = parseSourceToOutline(
    '## 页\n- 内容交付量同比增长 138%，主要来自模板复用与评审流程的合并，返工工时下降 74%\n- 人工返工工时下降 74%',
  )
  assert(longNumber.slides[0].kind === undefined, `long numeric prose became ${longNumber.slides[0].kind}`)

  // An explicit kind is NEVER overridden by inference.
  const explicit = parseSourceToOutline('## 页\n版式：列表\n- 第一步：a\n- 第二步：b\n- 第三步：c')
  assert(explicit.slides[0].kind === 'bullets', `explicit 列表 was overridden to ${explicit.slides[0].kind}`)

  // Stable: the same input always yields the same kinds.
  const source = '# 稳定\n\n## A\n- 旧：a\n- 新：b\n\n## B\n- 增长 138%\n- 下降 74%\n\n## C\n- 第一步：a\n- 第二步：b\n- 第三步：c'
  assert(stable(parseSourceToOutline(source)) === stable(parseSourceToOutline(source)), 'inference is not stable')
  return `${fixtures.length} fixtures, explicit wins, 3 negative cases, stable`
})

check('8c · the three new layouts render cleanly (3 densities x 3 tones)', () => {
  const LAYOUT_FIXTURES = {
    compare: ['旧：逐页手工排版', '旧：字体各处漂移', '新：模板一次解析', '新：导出即可编辑'],
    timeline: ['第一步：导入模板', '第二步：解析品牌', '第三步：生成卡组'],
    stats: ['交付量增长 138%', '返工下降 74%', '效率提升 3.2 倍'],
  }
  const printed = []
  let combinations = 0
  let slidesChecked = 0
  let elementsChecked = 0
  for (const [kind, bullets] of Object.entries(LAYOUT_FIXTURES)) {
    for (const density of DENSITIES) {
      for (const tone of TONES) {
        const tag = `${kind}/${density}/${tone}`
        const outline = { title: '版式测试', slides: [{ title: '内容页', kind, bullets }] }
        const { deck, warnings } = composeDeck({ outline, theme: HW_THEME, style: { density, tone } })

        const parsed = CardDeckSchema.safeParse(deck)
        assert(parsed.success, `${tag}: schema rejected — ${JSON.stringify(parsed.error?.issues ?? [])}`)

        const repaired = repairDeck(deck, HW_THEME)
        assert(repaired.warnings.length === 0, `${tag}: repair warnings — ${repaired.warnings.join(' | ')}`)
        assert(warnings.length === 0, `${tag}: compose warnings — ${warnings.join(' | ')}`)

        const contentSlide = deck.slides[1]
        assert(contentSlide.kind === kind, `${tag}: deck slide kind is ${contentSlide.kind}`)
        const errors = []
        collectGeometryErrors(contentSlide, HW_THEME, errors)
        assert(errors.length === 0, `${tag}: ${errors.join(' | ')}`)

        // Prove the routing is real: each kind leaves its own structural fingerprint
        // instead of a bullet list.
        const shapes = shapeElements(contentSlide)
        if (kind === 'compare') {
          assert(shapes.some((element) => element.id.includes('-divider-')), `${tag}: no divider`)
          assert(
            contentSlide.elements.filter((element) => element.type === 'text' && element.id.includes('-heading-')).length === 2,
            `${tag}: expected 2 column headings`,
          )
        } else if (kind === 'timeline') {
          assert(shapes.filter((element) => element.shapeType === 'ellipse').length === bullets.length, `${tag}: node count`)
          assert(shapes.filter((element) => element.id.includes('-seg-')).length === bullets.length - 1, `${tag}: segment count`)
        } else {
          // One number + one label per stat bullet. The card surface is a *tone*
          // decision (`minimal` draws none), so it is only asserted where the tone
          // actually fills cards — geometry must not depend on the tone.
          const statTexts = textElements(contentSlide).filter((element) => element.id.includes('-stat-'))
          const labelTexts = textElements(contentSlide).filter((element) => element.id.includes('-label-'))
          const cards = shapes.filter((element) => element.shapeType === 'roundRect')
          assert(labelTexts.length === bullets.length, `${tag}: ${labelTexts.length} labels for ${bullets.length} bullets`)
          assert(statTexts.length === bullets.length, `${tag}: ${statTexts.length} numbers for ${bullets.length} bullets`)
          assert(
            cards.length === (tone === 'minimal' ? 0 : bullets.length),
            `${tag}: ${cards.length} cards (tone ${tone})`,
          )
        }

        slidesChecked += 1
        elementsChecked += contentSlide.elements.length
        if (tone === 'formal' && density === 'standard') {
          printed.push(`${kind}: ${contentSlide.elements.length} elements (${textElements(contentSlide).length} text, ${shapes.length} shape)`)
        }
        combinations += 1
      }
    }
  }
  assert(combinations === 27, `expected 27 combinations, ran ${combinations}`)
  console.log('    rendered (standard/formal):')
  for (const line of printed) console.log(`      ${line}`)
  return `${combinations}/27 combinations clean (0 schema errors, 0 repair warnings, 0 overlaps), ${slidesChecked} slides / ${elementsChecked} elements`
})

check('8d · stats extracts numbers (2/3/4 cards) and timeline handles 2 and 6 bullets', () => {
  // --- stats: number extraction -------------------------------------------
  const statsOutline = {
    title: '指标',
    slides: [{ title: '关键指标', kind: 'stats', bullets: ['内容交付量同比增长 138%', '人工返工工时下降 74%'] }],
  }
  const statsDeck = composeDeck({ outline: statsOutline, theme: HW_THEME, style: {} })
  const statsSlide = statsDeck.deck.slides[1]
  const numbers = textElements(statsSlide).filter((element) => element.id.includes('-stat-')).map((element) => element.content)
  const labels = textElements(statsSlide).filter((element) => element.id.includes('-label-')).map((element) => element.content)
  console.log(`    stats numbers: ${stable(numbers)}`)
  console.log(`    stats labels : ${stable(labels)}`)
  assert(stable(numbers) === stable(['138%', '74%']), `numbers were ${stable(numbers)}`)
  assert(stable(labels) === stable(['内容交付量同比增长', '人工返工工时下降']), `labels were ${stable(labels)}`)
  assert(textElements(statsSlide).find((element) => element.id.includes('-stat-')).style.fontSize >= 40, 'stat number is not big')

  // --- stats: 2 / 3 / 4 cards all fit the safe width ----------------------
  const box = innerBox(HW_THEME)
  const cardWidths = []
  const cardBullets = [
    ['内容交付量同比增长 138%', '人工返工工时下降 74%'],
    ['内容交付量同比增长 138%', '人工返工工时下降 74%', '效率提升 3.2 倍'],
    ['内容交付量同比增长 138%', '人工返工工时下降 74%', '效率提升 3.2 倍', '成本降至三分之一'],
  ]
  cardBullets.forEach((bullets, index) => {
    const { deck, warnings } = composeDeck({
      outline: { title: '指标', slides: [{ title: '指标页', kind: 'stats', bullets }] },
      theme: HW_THEME,
      style: {},
    })
    const slide = deck.slides[1]
    const cards = shapeElements(slide).filter((element) => element.shapeType === 'roundRect')
    assert(cards.length === bullets.length, `${bullets.length} bullets → ${cards.length} cards`)
    assert(warnings.length === 0, `${bullets.length} cards: warnings ${stable(warnings)}`)
    assert(repairDeck(deck, HW_THEME).warnings.length === 0, `${bullets.length} cards: repair warnings`)
    assert(cards[0].x === box.x, `${bullets.length} cards: first card x=${cards[0].x}`)
    const last = cards[cards.length - 1]
    assert(last.x + last.w === box.x + box.w, `${bullets.length} cards: last card ends at ${last.x + last.w}`)
    // The row must be a row: every card on one line, no overlaps.
    for (const card of cards) assert(card.y === cards[0].y, `${bullets.length} cards: card y ${card.y}`)
    // A bullet without a number still gets a label and no number line.
    const numberCount = textElements(slide).filter((element) => element.id.includes('-stat-')).length
    cardWidths.push(`${bullets.length}:${cards[0].w}px/${numberCount} numbers`)
    if (index === 2) {
      // `成本降至三分之一` carries a Chinese-numeral quantity, so it does have a
      // headline: 4 bullets → 4 numbers. Arabic-digit-only extraction used to miss it
      // and leave a tall empty card beside the figures.
      assert(numberCount === 4, `4-card slide should have 4 numbers (三分之一 is a quantity), got ${numberCount}`)
      const fourth = textElements(slide).find((element) => element.id === 's2-stat-4')
      assert(fourth !== undefined && fourth.content === '三分之一', `4th number was ${fourth?.content}`)
      // The label must not keep the dangling connective: `成本降至` is half a sentence.
      const fourthLabel = textElements(slide).find((element) => element.id === 's2-label-4')
      assert(fourthLabel !== undefined && fourthLabel.content === '成本', `4th label was ${fourthLabel?.content}`)
    }
  })
  console.log(`    stats card widths: ${cardWidths.join(', ')}`)

  // The case that *does* degrade: a card with no quantity at all. A stat row reads as
  // data, so one missing number turns the page into a list rather than a tall empty card.
  const degrade = composeDeck({
    outline: {
      title: '指标',
      slides: [{ title: '指标页', kind: 'stats', bullets: ['内容交付量同比增长 138%', '人工返工工时下降 74%', '交付节奏更稳'] }],
    },
    theme: HW_THEME,
    style: {},
  })
  const degradeSlide = degrade.deck.slides[1]
  const degradeStats = textElements(degradeSlide).filter((element) => element.id.includes('-stat-') || element.id.includes('-label-'))
  assert(degradeStats.length === 0, `number-less stats page should degrade to a list, got ${degradeStats.length} stat elements`)
  assert(
    degrade.warnings.some((warning) => warning.includes('改为列表')),
    `no degrade warning in ${stable(degrade.warnings)}`,
  )
  const degradeCopy = textElements(degradeSlide).map((element) => element.content)
  assert(degradeCopy.includes('交付节奏更稳'), `degraded page lost copy: ${stable(degradeCopy)}`)

  // --- timeline: 2 nodes and 6 → +1 collapse ------------------------------
  const two = composeDeck({
    outline: { title: '时间线', slides: [{ title: '两步', kind: 'timeline', bullets: ['第一步：导入模板', '第二步：生成卡组'] }] },
    theme: HW_THEME,
    style: {},
  })
  const twoNodes = shapeElements(two.deck.slides[1]).filter((element) => element.shapeType === 'ellipse')
  assert(twoNodes.length === 2, `2 bullets → ${twoNodes.length} nodes`)
  assert(two.warnings.length === 0, `2-node timeline warnings: ${stable(two.warnings)}`)

  const sixBullets = ['第一步：导入模板', '第二步：解析品牌', '第三步：生成卡组', '第四步：导出', '第五步：评审', '第六步：归档']
  const six = composeDeck({
    outline: { title: '时间线', slides: [{ title: '六步', kind: 'timeline', bullets: sixBullets }] },
    theme: HW_THEME,
    style: {},
  })
  const sixSlide = six.deck.slides[1]
  const sixNodes = shapeElements(sixSlide).filter((element) => element.shapeType === 'ellipse')
  const captions = textElements(sixSlide).filter((element) => element.id.includes('-caption-')).map((element) => element.content)
  console.log(`    6 bullets → ${sixNodes.length} nodes, captions ${stable(captions)}`)
  console.log(`    fold warning: ${six.warnings.find((warning) => warning.includes('折叠'))}`)
  assert(sixNodes.length === 5, `6 bullets → ${sixNodes.length} nodes (expected 5)`)
  assert(captions[captions.length - 1].endsWith('+1'), `last caption was "${captions[captions.length - 1]}"`)
  assert(
    six.warnings.some((warning) => warning.includes('折叠为 +1')),
    `no fold warning in ${stable(six.warnings)}`,
  )
  // Folding must not lose the text: it goes to the notes.
  assert(
    sixSlide.speakerNotes !== undefined && sixSlide.speakerNotes.includes('第六步：归档'),
    `folded bullet not kept in notes: ${JSON.stringify(sixSlide.speakerNotes)}`,
  )
  assert(repairDeck(six.deck, HW_THEME).warnings.length === 0, '6-node timeline deck needs repair')
  return `numbers ${stable(numbers)}, cards ${cardWidths.join(' / ')}, timeline 2→2 nodes, 6→5 nodes +1 folded to notes`
})

check('8e · two-track: a long bullet moves its detail into speakerNotes', () => {
  // 98 characters: over MAX_BULLET_CHARS, with an early clause boundary after 32.
  const longBullet =
    '内容中台把模板解析、卡组编排、内容生成与原生导出串成一条完整链路，评审阶段直接给出可编辑的 pptx，业务方可以在自己的电脑上改字改图改配色而不破坏版式，因此一次解析可以长期复用，边际成本接近零。'
  assert(Array.from(longBullet).length > MAX_BULLET_CHARS, 'fixture is not long enough')
  const outline = { title: '双轨', slides: [{ title: '长要点', bullets: [longBullet, '短要点'] }] }
  const { deck, warnings } = composeDeck({ outline, theme: HW_THEME, style: {} })
  const slide = deck.slides[1]
  const visible = textElements(slide).filter((element) => element.role === 'body').map((element) => element.content)
  console.log(`    visible bullets: ${stable(visible)}`)
  console.log(`    speakerNotes   : ${JSON.stringify(slide.speakerNotes)}`)

  assert(
    warnings.some((warning) => warning.includes(DETAIL_TO_NOTES_WARNING)),
    `no "${DETAIL_TO_NOTES_WARNING}" warning in ${stable(warnings)}`,
  )
  assert(typeof slide.speakerNotes === 'string' && slide.speakerNotes !== '', 'notes were not generated')
  assert(visible.length === 2, `expected 2 visible bullets, got ${visible.length}`)
  for (const bullet of visible) {
    assert(Array.from(bullet).length <= MAX_BULLET_CHARS, `visible bullet is ${Array.from(bullet).length} chars`)
  }
  assert(visible[0] !== longBullet, 'the long bullet was not shortened')
  // No text is lost: the note holds everything the page dropped (modulo the clause
  // separator, which a page bullet does not end with).
  const flat = longBullet.replace(/\s+/g, ' ').trim()
  assert(flat.startsWith(visible[0]), `visible bullet "${visible[0]}" is not a prefix of the original`)
  const remainder = flat.slice(visible[0].length).replace(/^[，,。；;：:\s]+/, '')
  assert(remainder.length > 0, 'nothing was moved out of the bullet')
  assert(slide.speakerNotes.includes(remainder), `notes do not contain the removed text "${remainder}"`)
  assert(repairDeck(deck, HW_THEME).warnings.length === 0, 'two-track deck needs repair')

  // A slide that already HAS notes keeps the old split behaviour and its notes are
  // passed through verbatim (never truncated, never rewritten).
  const authored = '作者备注：先讲现象，再讲根因；这一段不能被生成器改动或截断。'
  const withNotes = composeDeck({
    outline: { title: '双轨', slides: [{ title: '长要点', bullets: [longBullet, '短要点'], notes: authored }] },
    theme: HW_THEME,
    style: {},
  })
  const noted = withNotes.deck.slides[1]
  const notedVisible = textElements(noted).filter((element) => element.role === 'body').map((element) => element.content)
  assert(noted.speakerNotes === authored, `authored notes changed: ${JSON.stringify(noted.speakerNotes)}`)
  assert(
    withNotes.warnings.some((warning) => warning.includes('已拆分')),
    `expected the historical split warning, got ${stable(withNotes.warnings)}`,
  )
  assert(notedVisible.length > 2, `expected the long bullet to split, got ${notedVisible.length} bullets`)
  assert(
    notedVisible.every((bullet) => bullet.length > 0),
    'split produced an empty bullet',
  )

  // Short copy never grows notes and never warns.
  const short = composeDeck({
    outline: { title: '短', slides: [{ title: '短要点', bullets: ['模板只解析一次', '卡组可以复用'] }] },
    theme: HW_THEME,
    style: {},
  })
  assert(short.deck.slides[1].speakerNotes === undefined, 'short slide gained speaker notes')
  assert(short.warnings.length === 0, `short slide warned: ${stable(short.warnings)}`)

  // And an explicit notes value always survives, even without long bullets.
  const notesOnly = composeDeck({
    outline: { title: '备注', slides: [{ title: '备注页', bullets: ['要点'], notes: '只有备注' }] },
    theme: HW_THEME,
    style: {},
  })
  assert(notesOnly.deck.slides[1].speakerNotes === '只有备注', 'short slide lost its notes')
  return `long ${Array.from(longBullet).length} → visible ${Array.from(visible[0]).length} chars, ${slide.speakerNotes.length}-char note, authored notes intact`
})

/* ------------------------------------------------------------------ *
 * Spill handling and sparse-page composition
 * ------------------------------------------------------------------ */

check('8g · a page never spills a single trailing bullet onto a 续 page', () => {
  // Two rules hold now: short copy that fits one page is not split at all, and copy that
  // does not fit is clamped into the notes rather than continued on a second page.
  const short = ['第一条要点', '第二条要点', '第三条要点', '第四条要点', '第五条要点']
  const one = composeDeck({
    outline: { title: '溢出', slides: [{ title: '五条', bullets: short }] },
    theme: HW_THEME,
    style: { bulletLayout: 'list', density: 'standard' },
  })
  const onePages = one.deck.slides.filter((slide) => slide.kind === 'bullets')
  const oneCounts = onePages.map((slide) => textElements(slide).filter((element) => element.id.includes('-body-')).length)
  assert(oneCounts.length === 1 && oneCounts[0] === 5, `5 short bullets should stay on one page, got ${stable(oneCounts)}`)
  assert(one.warnings.length === 0, `unexpected warnings: ${stable(one.warnings)}`)

  // Copy long enough that it cannot fit: still one page, no `（续）`, overflow to notes.
  const long = Array.from({ length: 10 }, (_, index) =>
    `第 ${index + 1} 条要点：内容中台把模板解析、卡组编排、内容生成与原生导出串成一条完整链路，每一层都用数据契约连接，评审阶段直接给出可编辑的 pptx`)
  const many = composeDeck({
    outline: { title: '溢出', slides: [{ title: '十条长要点', bullets: long }] },
    theme: HW_THEME,
    style: { bulletLayout: 'list', density: 'dense' },
  })
  const pages = many.deck.slides.filter((slide) => slide.kind === 'bullets')
  assert(pages.length === 1, `expected one page, got ${pages.length}`)
  assert(!many.deck.slides.some((slide) => /（续）$/.test(slide.title ?? '')), 'a 续 title survived')
  const counts = textElements(pages[0]).filter((element) => element.id.includes('-body-')).length
  assert(counts > 0 && counts < long.length, `page was not clamped (${counts}/${long.length})`)
  const notes = pages[0].speakerNotes ?? ''
  assert(notes.includes(long[long.length - 1]), 'clamped copy was not preserved in the notes')
  return `1 page for 5 short bullets; long page kept ${counts}/${long.length}, rest in notes`
})

check('8h · an over-full stats/compare page degrades to a list, not a 续 page', () => {
  // A `stats` row is a row of at most 4 cards. Five bullets is not a stat row, and
  // spilling the fifth onto its own page produces a nearly empty slide.
  const statsBullets = [
    '交付量增长 138%', '返工下降 74%', '效率提升 3.2 倍',
    '成本降至三分之一', '覆盖 12 个事业部', '上线 5 个模板',
  ]
  const { deck, warnings } = composeDeck({
    outline: { title: '指标', slides: [{ title: '指标页', kind: 'stats', bullets: statsBullets }] },
    theme: HW_THEME,
    style: { density: 'dense' },
  })
  const titles = deck.slides.map((slide) => slide.title)
  assert(!titles.some((title) => title.includes('（续）')), `a 续 page survived: ${stable(titles)}`)
  assert(
    warnings.some((warning) => warning.includes('已改为列表呈现')),
    `no degrade warning in ${stable(warnings)}`,
  )
  const kind = deck.slides[1].kind
  assert(kind === 'bullets', `degraded page should be bullets, got ${kind}`)
  return `${titles.length} pages, no 续 page`
})

check('8i · a sparse page steps the body size up instead of hugging the title', () => {
  // Four short bullets at 22px fill barely half of a 720px canvas. The copy is short
  // enough to carry the page, so it is set larger and centred within a bounded shift.
  const bullets = ['画布固定 1280×720', '安全边距 64 像素', 'Logo 位 150×48', '坐标基于同一基准']
  const { deck } = composeDeck({
    outline: { title: '稀疏', slides: [{ title: '四条短要点', bullets }] },
    theme: HW_THEME,
    style: { density: 'standard', bulletLayout: 'list' },
  })
  const slide = deck.slides[1]
  const body = textElements(slide).filter((element) => element.id.includes('-body-'))
  const size = body[0].style.fontSize
  assert(size > 22, `sparse page kept the dense size (${size}px)`)
  const bottom = Math.max(...body.map((element) => element.y + element.h))
  // Centred by at most BLOCK_CENTER_MAX, so it must stay in the upper two thirds…
  assert(bottom < 640, `body ran past the safe box (bottom ${bottom})`)
  // …and no longer sit right under the title bar.
  assert(body[0].y > 172, `body still hugs the title (y=${body[0].y})`)
  return `${size}px, first bullet y=${body[0].y}, bottom=${bottom}`
})

check('8j · a role label in the slide title is stripped, a real title is not', () => {
  // Models label the slide's role: `封面：…` / `第 3 页：…` / `Slide 4: …`. The deck has
  // its own cover, so the label is noise on screen — but `封面设计的三条原则` is a title.
  const cases = [
    ['封面：把合规演示从 5 天压到半小时', '把合规演示从 5 天压到半小时'],
    ['第 3 页：模板解析', '模板解析'],
    ['第三张、整体架构', '整体架构'],
    ['Slide 4: Export pipeline', 'Export pipeline'],
    ['封面设计的三条原则', '封面设计的三条原则'],
    ['目录页与结束页的关系', '目录页与结束页的关系'],
  ]
  for (const [input, expected] of cases) {
    const outline = parseSourceToOutline(`# 卡组\n\n## ${input}\n- 要点`)
    assert(outline.slides[0].title === expected, `"${input}" → "${outline.slides[0].title}" (expected "${expected}")`)
  }
  return `${cases.length} titles normalised`
})

check('8k · chapters produce a 目录 page with true page numbers and a divider each', () => {
  // 总-分-总: the outline names chapters with `章节：`; the composer groups them, numbers
  // the pages it actually produced, and puts a divider before each chapter.
  const source = [
    '# 卡组名',
    '',
    '## 总览一页',
    '章节：总览',
    '- 要点甲',
    '',
    '## 分节一页',
    '章节：实现',
    '- 要点乙',
    '',
    '## 分节二页',
    '章节：实现',
    '- 要点丙',
    '',
    '## 总结一页',
    '章节：总结',
    '- 要点丁',
  ].join('\n')
  const { deck, warnings } = composeDeck({ outline: source, theme: HW_THEME, style: { maxSlides: 20 } })
  const kinds = deck.slides.map((slide) => slide.kind)
  const titles = deck.slides.map((slide) => slide.title ?? '')
  console.log(`    kinds: ${kinds.join(',')} | titles: ${titles.join(' | ')}`)

  assert(kinds[0] === 'cover', `first slide is ${kinds[0]}`)
  assert(kinds[1] === 'toc', `second slide should be the 目录, got ${kinds[1]}`)

  // The TOC lists the three chapters in order…
  const toc = deck.slides[1]
  const tocText = textElements(toc).map((element) => element.content)
  for (const chapter of ['总览', '实现', '总结']) {
    assert(tocText.includes(chapter), `目录 missing "${chapter}": ${stable(tocText)}`)
  }

  // …and each chapter has exactly one divider, right before its first page.
  const dividers = deck.slides.filter((slide) => slide.kind === 'section')
  assert(dividers.length === 3, `expected 3 dividers, got ${dividers.length}`)
  assert(!warnings.some((warning) => warning.includes('续')), `unexpected 续 warning: ${stable(warnings)}`)

  // Page numbers in the TOC must match where the divider actually is (1-based).
  deck.slides.forEach((slide, index) => {
    if (slide.kind !== 'section') return
    const page = index + 1
    assert(tocText.includes(String(page)), `目录 has no page number ${page} for chapter "${slide.title}"`)
  })

  // A deck with no chapters keeps the flat shape: no 目录, no dividers.
  const flat = composeDeck({
    outline: { title: '无章节', slides: [{ title: '一页', bullets: ['甲'] }, { title: '两页', bullets: ['乙'] }] },
    theme: HW_THEME,
    style: { maxSlides: 10 },
  })
  assert(!flat.deck.slides.some((slide) => slide.kind === 'toc'), 'flat deck grew a 目录')
  assert(!flat.deck.slides.some((slide) => slide.kind === 'section'), 'flat deck grew a divider')
  return `${kinds.length} slides: 目录 + 3 dividers, TOC numbers verified`
})

check('8l · the page and bullet budgets are configuration, not constants', () => {
  // Both used to be module constants (150 chars per page, 90 per bullet). They change what
  // the generator *produces*, so a deployment with another house rule must be able to move
  // them — and a tighter bullet budget must actually shorten the visible copy.
  const bullets = [
    '内容中台把模板解析、卡组编排、内容生成与原生导出串成一条完整链路，每一层都用数据契约连接。',
    '评审阶段直接给出可编辑的 pptx，业务方可以改字改图改配色而不破坏版式。',
    '一次解析可以长期复用，边际成本接近零。',
  ]
  const wide = composeDeck({
    outline: { title: '预算', slides: [{ title: '宽预算', bullets }] },
    theme: HW_THEME,
    style: { maxSlides: 10 },
    bulletCharBudget: 200,
  })
  const narrow = composeDeck({
    outline: { title: '预算', slides: [{ title: '窄预算', bullets }] },
    theme: HW_THEME,
    style: { maxSlides: 10 },
    bulletCharBudget: 30,
  })
  const visibleChars = (result) => {
    const slide = result.deck.slides.find((entry) => entry.kind === 'bullets')
    return textElements(slide)
      .filter((element) => element.id.includes('-body-'))
      .reduce((sum, element) => sum + Array.from(element.content).length, 0)
  }
  assert(visibleChars(narrow) < visibleChars(wide), `narrow ${visibleChars(narrow)} !< wide ${visibleChars(wide)}`)
  // A tighter bullet budget means more detail moves to the notes: the visible copy shrinks
  // and the speaker notes grow to hold the rest (nothing is thrown away).
  const notesLength = (result) => {
    const slide = result.deck.slides.find((entry) => entry.kind === 'bullets')
    return (slide.speakerNotes ?? '').length
  }
  assert(notesLength(narrow) > notesLength(wide), `narrow notes ${notesLength(narrow)} !> wide ${notesLength(wide)}`)
  assert(
    narrow.warnings.some((warning) => warning.includes('细节已移入演讲备注')),
    `no two-track warning: ${stable(narrow.warnings)}`,
  )

  // A tighter page budget must be reported against the configured number, not 150.
  const repaired = repairDeck(wide.deck, HW_THEME, { slideCharBudget: 60 })
  assert(
    repaired.warnings.some((warning) => warning.includes('budget 60')),
    `repair did not use the configured page budget: ${stable(repaired.warnings.slice(0, 3))}`,
  )
  // …and the default stays 150 when nobody says otherwise.
  const defaults = repairDeck(wide.deck, HW_THEME)
  assert(
    !defaults.warnings.some((warning) => warning.includes('budget 60')),
    'the default budget was overridden by accident',
  )
  return `visible chars ${visibleChars(narrow)} (30) vs ${visibleChars(wide)} (200); page budget honoured`
})

check('8m · a 标签： marker names the compare columns instead of becoming copy', () => {
  // Regression: a model writing a compare page often emits `标签：故障现象` followed by that
  // column's bullets. The old splitter treated the marker itself as a bullet, so the literal
  // text "标签：故障现象" appeared on the slide and the columns were halved arbitrarily.
  const bullets = ['标签：故障现象', '中文重影', '启动崩溃', '标签：根因', 'pptgenjs 重复写字型名', '插件把 connection 设为硬依赖']
  const { deck } = composeDeck({
    outline: { title: '对比', slides: [{ title: '两个真实故障', kind: 'compare', bullets }] },
    theme: HW_THEME,
    style: { maxSlides: 10 },
  })
  const slide = deck.slides.find((entry) => entry.kind === 'compare')
  const copy = textElements(slide).map((element) => element.content)
  console.log(`    compare copy: ${stable(copy)}`)
  assert(!copy.some((text) => text.startsWith('标签：')), `the marker leaked into the copy: ${stable(copy)}`)
  assert(copy.includes('故障现象'), 'left column heading missing')
  assert(copy.includes('根因'), 'right column heading missing')
  assert(copy.includes('中文重影') && copy.includes('启动崩溃'), 'left column lost bullets')
  assert(copy.includes('pptgenjs 重复写字型名'), 'right column lost bullets')
  return `headings + ${copy.length} texts, no marker leakage`
})

/* ------------------------------------------------------------------ *
 * Summary
 * ------------------------------------------------------------------ */
const failed = results.filter((result) => !result.ok)
console.log('')
console.log(`${results.length} checks, ${results.length - failed.length} passed, ${failed.length} failed`)
if (failed.length > 0) {
  for (const result of failed) console.log(`  FAILED: ${result.name}`)
  process.exitCode = 1
} else {
  console.log('ALL PASS')
}
