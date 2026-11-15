/**
 * Contract test for src/schema.js — the shared DSL every other module depends on.
 * Pure, no I/O beyond importing fixtures.
 *
 * Run: node test/schema-test.mjs
 */
import { HW_DEMO_DECK } from './fixtures/hw-demo-deck.mjs'
import {
  CANVAS,
  SAFE_AREA,
  CardDeckSchema,
  ThemeTokensSchema,
  parseAndRepairDeck,
  repairDeck,
  resolveColor,
  splitAlpha,
  normalizeHex,
  validateSvg,
  estimateLines,
  estimateTextOverflow,
  innerBox,
} from '../src/schema.js'

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

const results = []
function check(name, fn) {
  try {
    const detail = fn()
    results.push({ name, ok: true })
    console.log(`PASS ${name}${detail === undefined ? '' : ` — ${detail}`}`)
  } catch (error) {
    results.push({ name, ok: false, error })
    console.log(`FAIL ${name}: ${error?.message ?? error}`)
  }
}
const assert = (condition, message) => {
  if (!condition) throw new Error(message)
}

check('ThemeTokensSchema accepts the HW token set', () => {
  const parsed = ThemeTokensSchema.safeParse(HW_THEME)
  assert(parsed.success, JSON.stringify(parsed.error?.issues))
})

check('demo deck matches CardDeckSchema', () => {
  const parsed = CardDeckSchema.safeParse(HW_DEMO_DECK)
  assert(parsed.success, JSON.stringify(parsed.error?.issues?.slice(0, 3), null, 2))
  return `${parsed.data.slides.length} slides, ${parsed.data.slides.reduce((n, s) => n + s.elements.length, 0)} elements`
})

check('repo demo deck repairs cleanly against the HW theme', () => {
  const repaired = parseAndRepairDeck(HW_DEMO_DECK, HW_THEME)
  assert(repaired.ok, 'repair failed')
  const geometryWarnings = repaired.warnings.filter((w) => w.includes('geometry clamped'))
  const overflowWarnings = repaired.warnings.filter((w) => w.includes('overflows'))
  assert(geometryWarnings.length === 0, `unexpected clamping: ${geometryWarnings.join(' | ')}`)
  assert(overflowWarnings.length === 0, `unexpected overflow: ${overflowWarnings.join(' | ')}`)
  // Every element must sit inside the canvas.
  for (const slide of repaired.deck.slides) {
    for (const element of slide.elements) {
      assert(element.x >= 0 && element.y >= 0, `${element.id} negative position`)
      assert(element.x + element.w <= CANVAS.width, `${element.id} exceeds width`)
      assert(element.y + element.h <= CANVAS.height, `${element.id} exceeds height`)
    }
  }
  return `${repaired.warnings.length} warning(s)`
})

check('colors are resolved to hex from tokens', () => {
  const repaired = parseAndRepairDeck(HW_DEMO_DECK, HW_THEME)
  assert(repaired.ok, 'repair failed')
  const texts = repaired.deck.slides.flatMap((s) => s.elements).filter((e) => e.type === 'text')
  for (const text of texts) {
    assert(/^#[0-9A-F]{6}$/.test(text.style.color), `unresolved color ${text.style.color} on ${text.id}`)
  }
  const fills = repaired.deck.slides
    .flatMap((s) => s.elements)
    .filter((e) => e.type === 'shape' && e.style?.fill)
  for (const shape of fills) {
    assert(/^#[0-9A-F]{6}$/.test(shape.style.fill), `unresolved fill ${shape.style.fill} on ${shape.id}`)
  }
  return `${texts.length} text + ${fills.length} filled shapes resolved`
})

check('resolveColor handles tokens, palettes and literals', () => {
  assert(resolveColor('$colors.primary', HW_THEME) === '#C7000A', 'token ref')
  assert(resolveColor('colors.accent', HW_THEME) === '#E9002F', 'unprefixed token ref')
  assert(resolveColor('$colors.chart3', HW_THEME) === '#F4A100', 'palette ref')
  assert(resolveColor('#abc', HW_THEME) === '#AABBCC', 'short hex')
  assert(resolveColor('', HW_THEME, '#123456') === '#123456', 'fallback')
  assert(resolveColor('$colors.nope', HW_THEME, '#123456') === '#123456', 'unknown token falls back')
})

check('alpha hex splits into color + opacity', () => {
  const split = splitAlpha('#80C7000A')
  assert(split.hex === '#C7000A', `hex: ${split.hex}`)
  assert(Math.abs(split.alpha - 0x80 / 255) < 1e-9, `alpha: ${split.alpha}`)
  assert(splitAlpha('#C7000A').alpha === 1, 'opaque')
  assert(normalizeHex('c7000a') === '#C7000A', 'normalize')
})

check('out-of-bounds geometry is clamped and reported', () => {
  const bad = {
    deckId: 'bad',
    title: 'bad',
    templateId: 'ppt-hw',
    slides: [
      {
        slideId: 's1',
        elements: [
          { type: 'shape', id: 'a', shapeType: 'rect', x: -50, y: 700, w: 4000, h: 400, style: { fill: '$colors.primary' } },
        ],
      },
    ],
  }
  const { deck, warnings } = repairDeck(bad, HW_THEME)
  const shape = deck.slides[0].elements[0]
  assert(shape.x >= 0, `x not clamped: ${shape.x}`)
  assert(shape.w <= CANVAS.width, `w not clamped: ${shape.w}`)
  assert(shape.x + shape.w <= CANVAS.width, 'right edge exceeds canvas')
  assert(shape.y + shape.h <= CANVAS.height, 'bottom edge exceeds canvas')
  assert(warnings.some((w) => w.includes('geometry clamped')), 'no clamping warning')
  return warnings[0]
})

check('text under the logo corner is pushed down', () => {
  const deck = {
    deckId: 'd',
    title: 't',
    templateId: 'ppt-hw',
    slides: [
      {
        slideId: 's1',
        elements: [
          {
            type: 'text',
            id: 'headline',
            role: 'title',
            content: '标题',
            x: 1080,
            y: 64,
            w: 130,
            h: 40,
            style: { fontSize: 20, color: '$colors.textPrimary' },
          },
        ],
      },
    ],
  }
  const repaired = repairDeck(deck, HW_THEME)
  const text = repaired.deck.slides[0].elements[0]
  const box = innerBox(HW_THEME)
  assert(box.logo !== null, 'logo should be reserved when the theme has one')
  assert(text.y >= box.logo.height + box.logo.gap, `text stayed in the logo corner: y=${text.y}`)
})

check('svg sanitizer strips foreignObject/script and keeps vectors', () => {
  const dirty = `<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10">
    <foreignObject width="10" height="10"><div>x</div></foreignObject>
    <script>alert(1)</script>
    <rect x="0" y="0" width="10" height="10" fill="#C7000A" onclick="evil()"/>
    <path d="M0 0 L10 10"/>
  </svg>`
  const { sanitized, problems } = validateSvg(dirty)
  assert(!sanitized.includes('foreignObject'), 'foreignObject survived')
  assert(!sanitized.includes('script'), 'script survived')
  assert(!sanitized.includes('onclick'), 'event handler survived')
  assert(sanitized.includes('rect') && sanitized.includes('path'), 'vectors were removed')
  assert(problems.length >= 2, `expected problems, got ${problems.length}`)
  return `${problems.length} problem(s) reported`
})

check('svg without width/height inherits its viewBox size', () => {
  const { sanitized } = validateSvg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 180"><rect width="320" height="180"/></svg>')
  assert(/width="320"/.test(sanitized) && /height="180"/.test(sanitized), sanitized.slice(0, 120))
})

check('text overflow is estimated with CJK width awareness', () => {
  const cjk = estimateLines('品牌智能演示平台内容生产链路', 200, 24)
  const latin = estimateLines('Brand deck content pipeline', 200, 24)
  // 14 full-width glyphs at 24px inside 200px wrap to 2 lines; the same glyph
  // count in Latin is narrower, so it must never need more lines than CJK.
  assert(cjk.lines === 2, `CJK lines: ${cjk.lines}`)
  assert(latin.lines <= cjk.lines, `latin ${latin.lines} > cjk ${cjk.lines}`)
  const slide = {
    slideId: 's',
    elements: [
      {
        type: 'text',
        id: 'tiny',
        role: 'body',
        content: '这是一段明显超出文本框高度的长文本，用于触发溢出预警逻辑。',
        x: 0,
        y: 0,
        w: 160,
        h: 24,
        style: { fontSize: 20, color: '#000000' },
      },
    ],
  }
  const overflow = estimateTextOverflow(slide, HW_THEME)
  assert(overflow.overflowing.length === 1, 'overflowing box not detected')
  return `overflowing: ${overflow.overflowing[0].id}`
})

check('invalid deck produces structured issues, not a throw', () => {
  const result = parseAndRepairDeck({ deckId: 'x', title: '', templateId: 't', slides: [] }, HW_THEME)
  assert(result.ok === false, 'should be rejected')
  assert(Array.isArray(result.issues) && result.issues.length > 0, 'no issues reported')
  return `${result.issues.length} issue(s)`
})

check('safe area reserves the logo only when the theme has one', () => {
  const withoutLogo = { ...HW_THEME, brandAssets: { icons: {} } }
  assert(innerBox(withoutLogo).logo === null, 'logo must not be reserved without an asset')
  assert(innerBox(HW_THEME).logo !== null, 'logo must be reserved with an asset')
  assert(innerBox(HW_THEME).x === SAFE_AREA.left, 'safe area left')
})

const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} schema checks passed`)
if (failed.length > 0) process.exitCode = 1
