/**
 * SVG theme normalisation test — "make model-drawn diagrams belong to the deck".
 *
 * The rule under test is a *judgement* rule, not just a string rewrite:
 *   · foreign palettes are replaced by the template's tokens;
 *   · three same-kind shapes keep ONE colour (a flow has one semantic, not three);
 *   · genuinely multi-series figures can opt into the chart palette;
 *   · semantic values (`none`, `url(#…)`) are never touched;
 *   · the font always comes from the template, never from the model.
 *
 * Run: node test/svg-theme-test.mjs
 */
import { themeSvg, normalizeSvgTheme, classifyColor, luminance, parseColor } from '../src/svg-theme.js'
import { inspect } from 'node:util'

const THEME = {
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
  brandAssets: { icons: {} },
}

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
const assert = (condition, message) => {
  if (!condition) throw new Error(message)
}

const FLOW = `<svg xmlns="http://www.w3.org/2000/svg" width="600" height="200" viewBox="0 0 600 200">
  <rect x="40" y="60" width="140" height="80" rx="12" fill="#8E44AD"/>
  <rect x="230" y="60" width="140" height="80" rx="12" fill="#27AE60"/>
  <rect x="420" y="60" width="140" height="80" rx="12" fill="#2980B9"/>
  <text x="78" y="110" font-family="Comic Sans MS" font-size="22" fill="#FFFFFF">采集</text>
  <text x="268" y="110" font-family="Comic Sans MS" font-size="22" fill="#FFFFFF">清洗</text>
  <text x="458" y="110" font-family="Comic Sans MS" font-size="22" fill="#FFFFFF">入库</text>
</svg>`

/* 1. Foreign palette is gone ------------------------------------------------ */
check('a foreign palette is fully replaced by the template tokens', () => {
  const out = themeSvg(FLOW, THEME)
  for (const foreign of ['#8E44AD', '#27AE60', '#2980B9']) {
    assert(!out.svg.toUpperCase().includes(foreign), `survived: ${foreign}`)
  }
  assert(out.remapped >= 6, `remapped only ${out.remapped}`)
  return `${out.remapped} colours remapped`
})

/* 2. Same-kind shapes keep one colour (the "don't copy blindly" rule) ------- */
check('three same-kind flow shapes all become the brand primary', () => {
  const out = themeSvg(FLOW, THEME)
  const fills = [...out.svg.matchAll(/<rect[^>]*fill="([^"]+)"/g)].map((match) => match[1].toUpperCase())
  assert(fills.length === 3, `got ${fills.length} rects`)
  assert(new Set(fills).size === 1, `colours diverged: ${fills.join(', ')}`)
  assert(fills[0] === '#C7000A', `expected the brand primary, got ${fills[0]}`)
  return `all three -> ${fills[0]}`
})

check('distinctAccents opts into the chart palette for real series', () => {
  const out = normalizeSvgTheme(FLOW, THEME, { distinctAccents: true })
  const fills = [...out.svg.matchAll(/<rect[^>]*fill="([^"]+)"/g)].map((match) => match[1].toUpperCase())
  assert(new Set(fills).size === 3, `expected 3 distinct series colours, got ${fills.join(', ')}`)
  for (const fill of fills) {
    assert(
      THEME.colors.chartPalette.map((entry) => entry.toUpperCase()).includes(fill),
      `${fill} is not from the template palette`,
    )
  }
  return fills.join(' ')
})

/* 3. Semantic values survive ------------------------------------------------ */
check('`none`, `url(#id)` and `currentColor` are treated as semantic', () => {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="g"/></defs>
    <rect fill="url(#g)"/><text fill="none">x</text><path stroke="currentColor"/>
    <rect style="fill:none;stroke:url(#g)"/></svg>`
  const out = themeSvg(svg, THEME)
  assert(out.svg.includes('fill="url(#g)"'), 'url() ref rewritten')
  assert(out.svg.includes('fill="none"'), 'none rewritten')
  assert(out.svg.includes('stroke="currentColor"') || out.svg.includes('currentColor'), 'currentColor rewritten')
  assert(out.svg.includes('fill:none'), 'inline none rewritten')
})

/* 4. Fonts come from the template ------------------------------------------ */
check('the model typeface is replaced by the template typeface', () => {
  const out = themeSvg(FLOW, THEME)
  assert(!/comic sans/i.test(out.svg), 'Comic Sans survived')
  const families = [...out.svg.matchAll(/font-family="([^"]*)"/g)].map((match) => match[1])
  assert(families.length === 3, `expected a family on each text node, got ${families.length}`)
  assert(families.every((family) => family === 'Arial'), `unexpected families: ${families.join(', ')}`)
  return `${families.length} text nodes -> Arial`
})

/* 5. Structure and geometry untouched -------------------------------------- */
check('structure, ids and geometry are not modified', () => {
  const out = themeSvg(FLOW, THEME)
  for (const fragment of ['width="600"', 'height="200"', 'viewBox="0 0 600 200"', 'rx="12"', 'x="230"']) {
    assert(out.svg.includes(fragment), `lost ${fragment}`)
  }
  assert((out.svg.match(/<rect/g) ?? []).length === 3, 'rect count changed')
  assert((out.svg.match(/<text/g) ?? []).length === 3, 'text count changed')
})

/* 6. Theme-driven, not hard-coded ----------------------------------------- */
check('a different template yields a different palette (no hard-coded red)', () => {
  const blue = {
    ...THEME,
    colors: { ...THEME.colors, primary: '#1F4E8C', accent: '#2E7DD1', chartPalette: ['#1F4E8C', '#2E7DD1'] },
  }
  const out = themeSvg(FLOW, blue)
  const fills = [...out.svg.matchAll(/<rect[^>]*fill="([^"]+)"/g)].map((match) => match[1].toUpperCase())
  assert(fills.every((fill) => fill === '#1F4E8C'), `expected the blue primary, got ${fills.join(', ')}`)
  return '#1F4E8C'
})

/* 7. Small helpers --------------------------------------------------------- */
check('colour classification is sane', () => {
  assert(classifyColor('#8E44AD', THEME) === 'primary', 'saturated purple should be a highlight')
  assert(classifyColor('#FFFFFF', THEME) === 'background', 'white should be background')
  assert(classifyColor('#F5F5F5', THEME) === 'surface' || classifyColor('#F5F5F5', THEME) === 'background', 'light grey')
  assert(classifyColor('#1D1D1A', THEME) === 'text', 'near-black should be text')
  assert(luminance('#FFFFFF') > 0.99 && luminance('#000000') < 0.01, 'luminance range')
  assert(JSON.stringify(parseColor('#abc')) === JSON.stringify([170, 187, 204]), 'short hex')
  assert(JSON.stringify(parseColor('rgb(1, 2, 3)')) === JSON.stringify([1, 2, 3]), 'rgb()')
})

check('garbage input never throws', () => {
  for (const input of ['', null, undefined, 42, '<svg>', '<svg fill="not-a-colour"/>']) {
    const out = themeSvg(input, THEME)
    assert(typeof out.svg === 'string', `not a string for ${inspect(input)}`)
  }
})

const failed = results.filter((entry) => !entry.ok)
console.log(`\n${results.length - failed.length}/${results.length} svg-theme checks passed`)
if (failed.length > 0) process.exitCode = 1
