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
 */
import {
  DEFAULT_STYLE,
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
  const bullets = Array.from({ length: 40 }, (_, index) => `第 ${index + 1} 条要点：内容中台把模板解析与原生导出串成一条流水线。`)
  const outline = { title: '上限测试', slides: [{ title: '长清单', bullets }] }
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

check('5b · 40 bullets produce continuation slides titled （续）', () => {
  const bullets = Array.from({ length: 40 }, (_, index) => `第 ${index + 1} 条要点：内容中台把模板解析与原生导出串成一条流水线。`)
  const outline = { title: '续页测试', slides: [{ title: '长清单', bullets }] }
  const { deck, warnings } = composeDeck({ outline, theme: HW_THEME, style: { maxSlides: 60, density: 'standard', bulletLayout: 'list' } })

  const titles = deck.slides.map((slide) => slide.title)
  console.log(`    slide titles: ${titles.join(' | ')}`)
  const continuation = titles.filter((title) => typeof title === 'string' && title.endsWith('（续）'))
  assert(continuation.length > 0, `no continuation slides among ${stable(titles)}`)
  assert(
    titles[1] === '长清单' && titles[2] === '长清单（续）',
    `first continuation title was "${titles[2]}"`,
  )
  // 40 bullets at standard density = 4 per slide = 10 content slides + cover + closing.
  assert(deck.slides.length === 12, `expected 12 slides (cover + 10 + closing), got ${deck.slides.length}`)
  assert(CardDeckSchema.safeParse(deck).success, 'deck is not schema-valid')
  const geometry = repairDeck(deck, HW_THEME).warnings.filter((warning) => GEOMETRY_WARNING.test(warning))
  assert(geometry.length === 0, `continuation deck has geometry warnings: ${geometry.join(' | ')}`)
  assert(warnings.every((warning) => !GEOMETRY_WARNING.test(warning)), `unexpected warnings: ${stable(warnings)}`)
  return `${deck.slides.length} slides, ${continuation.length} continuation slides`
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
