/**
 * Tests for src/mermaid.js — Mermaid flowchart subset → themed SVG.
 *
 * The point of these tests is the *contract with the model*: it writes Mermaid, we must
 * either draw it faithfully or fail with a message that says which line to fix. Silent
 * misdrawing (an `LR` flow stacked vertically, a figure clipped by the viewport) is the
 * failure mode that reaches the customer's slides, so both are pinned here.
 *
 * Run: node test/mermaid-test.mjs
 */
import { MermaidError, mermaidToSvg, parseMermaid } from '../src/mermaid.js'

const COLORS = Object.freeze({
  textPrimary: '#1D1D1A',
  textSecondary: '#666666',
  primary: '#C7000A',
  background: '#FFFFFF',
  surface: '#EBEBEB',
})

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

/** Render with a fixed box so assertions on geometry are readable. */
function render(source, width = 1152, height = 396) {
  return mermaidToSvg(source, { width, height, colors: COLORS })
}

/** Pull the `<text>` contents out of an SVG string. */
function labels(svg) {
  return [...svg.matchAll(/<text[^>]*>([^<]*)<\/text>/g)].map((match) => match[1])
}

/* ------------------------------------------------------------------ *
 * Parsing
 * ------------------------------------------------------------------ */

check('1a · a linear flow parses into nodes and edges in order', () => {
  const graph = parseMermaid(['flowchart LR', '  A[模板解析] --> B[内容生成]', '  B --> C[原生编译]'].join('\n'))
  assert(graph.direction === 'LR', `direction was ${graph.direction}`)
  assert(graph.nodes.length === 3, `nodes ${graph.nodes.length}`)
  assert(graph.edges.length === 2, `edges ${graph.edges.length}`)
  assert(stable(graph.nodes.map((node) => node.text)) === stable(['模板解析', '内容生成', '原生编译']), stable(graph.nodes.map((node) => node.text)))
  return `${graph.nodes.length} nodes / ${graph.edges.length} edges`
})

check('1b · TD and TB both mean top-down; graph is an alias for flowchart', () => {
  for (const header of ['flowchart TD', 'flowchart TB', 'graph TD', 'graph LR']) {
    const graph = parseMermaid(`${header}\n A --> B`)
    const expected = /LR/i.test(header) ? 'LR' : 'TD'
    assert(graph.direction === expected, `${header} → ${graph.direction}`)
  }
  return '4 headers map correctly'
})

check('1c · node shapes and edge labels are read, not guessed', () => {
  const graph = parseMermaid(
    ['flowchart TD', '  A(开始) --> B{判断}', '  B -->|是| C[执行]', '  B -.-> D[跳过]'].join('\n'),
  )
  const byId = new Map(graph.nodes.map((node) => [node.id, node]))
  assert(byId.get('A')?.shape === 'round', `A shape ${byId.get('A')?.shape}`)
  assert(byId.get('B')?.shape === 'diamond', `B shape ${byId.get('B')?.shape}`)
  assert(byId.get('C')?.shape === 'rect', `C shape ${byId.get('C')?.shape}`)
  const labels = graph.edges.map((edge) => edge.label)
  assert(labels.includes('是'), `edge labels ${stable(labels)}`)
  assert(graph.edges.some((edge) => edge.dashed === true), 'dashed edge lost')
  return `shapes round/diamond/rect, labels ${stable(labels)}`
})

check('1d · subgraphs become groups that own their members', () => {
  const graph = parseMermaid(
    ['flowchart TD', '  subgraph 接入层', '    A[网关] --> B[鉴权]', '  end', '  subgraph 存储层', '    C[数据库]', '  end', '  B --> C'].join('\n'),
  )
  assert(graph.groups.length === 2, `groups ${graph.groups.length}`)
  assert(stable(graph.groups[0].members) === stable(['A', 'B']), `group 0 members ${stable(graph.groups[0].members)}`)
  assert(stable(graph.groups[1].members) === stable(['C']), `group 1 members ${stable(graph.groups[1].members)}`)
  return `2 groups: ${graph.groups.map((group) => `${group.title}(${group.members.length})`).join(', ')}`
})

check('1e · chained edges and comments are handled', () => {
  const graph = parseMermaid(['flowchart LR', '  %% 这是注释', '  A --> B --> C', ''].join('\n'))
  assert(graph.edges.length === 2, `edges ${graph.edges.length}`)
  assert(stable(graph.edges.map((edge) => `${edge.from}->${edge.to}`)) === stable(['A->B', 'B->C']), stable(graph.edges))
  return 'A->B->C'
})

/* ------------------------------------------------------------------ *
 * Rejection: a clear error beats a wrong picture
 * ------------------------------------------------------------------ */

check('2a · unsupported diagram types are rejected with the line number', () => {
  const cases = [
    ['sequenceDiagram\n  A->>B: hi', /只支持 flowchart/],
    ['flowchart TD\n  classDef red fill:#f00', /装饰语法/],
    ['flowchart TD\n  A["<b>粗</b>"] --> B', /HTML/],
    ['', /空的图/],
  ]
  for (const [source, pattern] of cases) {
    let message = null
    try {
      parseMermaid(source)
    } catch (error) {
      message = error instanceof MermaidError ? error.message : `wrong error type: ${error}`
    }
    assert(message !== null, `"${source.slice(0, 20)}" was accepted`)
    assert(pattern.test(message), `message was "${message}"`)
  }
  return `${cases.length} inputs rejected with a reason`
})

check('2b · an unbalanced `end` is an error, not a silent flattening', () => {
  let message = null
  try {
    parseMermaid('flowchart TD\n  end')
  } catch (error) {
    message = error.message
  }
  assert(message !== null && /subgraph/.test(message), `message was "${message}"`)
  return 'rejected'
})

/* ------------------------------------------------------------------ *
 * Rendering
 * ------------------------------------------------------------------ */

check('3a · the SVG carries every node label and the theme colours', () => {
  const { svg, nodes, edges } = render(['flowchart LR', '  A[模板解析] --> B[内容生成]', '  B --> C[原生编译]'].join('\n'))
  assert(nodes === 3 && edges === 2, `${nodes}/${edges}`)
  const text = labels(svg)
  for (const label of ['模板解析', '内容生成', '原生编译']) assert(text.includes(label), `label "${label}" missing from ${stable(text)}`)
  assert(svg.includes(COLORS.textPrimary), 'node stroke colour missing')
  assert(svg.startsWith('<svg') && svg.endsWith('</svg>'), 'svg envelope broken')
  assert(!/<(?:foreignObject|script|image)\b/.test(svg), 'svg contains a forbidden element')

  // The brand accent is spent on the *decision* node only: a plain rect flow staying
  // black-and-white is the intended restraint, so this is a separate fixture.
  const withDecision = render(['flowchart TD', '  A[读入] --> B{材料是否长}', '  B -->|是| C[分块]'].join('\n')).svg
  assert(withDecision.includes(COLORS.primary), 'decision node is not using the brand accent')
  return `${text.length} labels, accent on the decision node`
})

check('3b · LR draws along x, TD draws along y', () => {
  const source = ['A[一] --> B[二]', 'B --> C[三]'].join('\n')
  // Node rects are the only `<rect>` elements without rx=0; read their x/y.
  const rects = (svg) => [...svg.matchAll(/<rect x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"/g)].map((m) => ({
    x: Number(m[1]),
    y: Number(m[2]),
    w: Number(m[3]),
  }))
  const lr = rects(render(`flowchart LR\n  ${source}`).svg)
  const td = rects(render(`flowchart TD\n  ${source}`).svg)
  assert(lr.length === 3 && td.length === 3, `rect counts ${lr.length}/${td.length}`)
  const spreadLR = Math.max(...lr.map((r) => r.x)) - Math.min(...lr.map((r) => r.x))
  const spreadTD = Math.max(...td.map((r) => r.x)) - Math.min(...td.map((r) => r.x))
  assert(spreadLR > 100, `LR nodes are not spread horizontally (${spreadLR}px)`)
  assert(spreadLR > spreadTD, `LR spread ${spreadLR} should exceed TD spread ${spreadTD}`)
  return `LR x-spread ${Math.round(spreadLR)}px vs TD ${Math.round(spreadTD)}px`
})

check('3c · the drawing fits the box: nothing is clipped, and it is not a thread', () => {
  const many = Array.from({ length: 9 }, (_, index) => `N${index}[节点${index + 1}]`).join('\n  ')
  const chain = Array.from({ length: 8 }, (_, index) => `N${index} --> N${index + 1}`).join('\n  ')
  const svg = render(`flowchart TD\n  ${many}\n  ${chain}`, 1152, 300).svg
  // The transform is what guarantees fit; it must scale *down* for a tall graph.
  const transform = /<g transform="translate\(([-\d.]+),([-\d.]+)\) scale\(([\d.]+)\)"/.exec(svg)
  assert(transform !== null, 'no fit transform emitted')
  const scale = Number(transform[3])
  assert(scale > 0 && scale <= 1.001, `tall graph should scale down, got ${scale}`)
  assert(scale >= 0.35, `scale ${scale} is too small to read`)
  // …and a small graph must be scaled *up* instead of hugging the middle.
  const small = /scale\(([\d.]+)\)/.exec(render('flowchart LR\n  A[一] --> B[二]', 1152, 300).svg)
  assert(Number(small[1]) > 1, `small graph should scale up, got ${small[1]}`)
  return `tall graph ×${scale}, small graph ×${small[1]}`
})

check('3d · a group box is drawn behind its members and carries its title', () => {
  const { svg, groups } = render(
    ['flowchart TD', '  subgraph 接入层', '    A[网关] --> B[鉴权]', '  end', '  B --> C[数据库]'].join('\n'),
  )
  assert(groups === 1, `groups ${groups}`)
  const text = labels(svg)
  assert(text.includes('接入层'), `group title missing from ${stable(text)}`)
  const groupIndex = svg.indexOf('接入层')
  const nodeIndex = svg.indexOf('网关')
  assert(groupIndex < nodeIndex, 'group box was drawn after its members')
  return 'group box behind nodes'
})

check('3e · the same source always renders the same SVG (determinism)', () => {
  const source = ['flowchart LR', '  A[一] --> B[二]', '  B -->|说明| C{判断}'].join('\n')
  const first = render(source).svg
  const second = render(source).svg
  assert(first === second, 'two renders of the same source differ')
  return `${first.length} bytes, stable`
})

check('3g · a tall TD flow is transposed so it stays readable in a 16:9 band', () => {
  // Seven ranks stacked vertically inside a 382px-tall figure shrink to an unreadable
  // thread; the reading order is preserved by drawing the ranks left-to-right instead.
  const chain = Array.from({ length: 6 }, (_, index) => `N${index}[环节${index + 1}] --> N${index + 1}[环节${index + 2}]`).join('\n  ')
  const svg = render(`flowchart TD\n  ${chain}`, 1152, 382).svg
  const rects = [...svg.matchAll(/<rect x="([\d.]+)" y="([\d.]+)" width="([\d.]+)"/g)].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }))
  const spreadX = Math.max(...rects.map((r) => r.x)) - Math.min(...rects.map((r) => r.x))
  const spreadY = Math.max(...rects.map((r) => r.y)) - Math.min(...rects.map((r) => r.y))
  assert(rects.length === 7, `node rects ${rects.length}`)
  assert(spreadX > spreadY, `transposed flow should run along x (x ${Math.round(spreadX)} vs y ${Math.round(spreadY)})`)
  // A short, wide TD graph keeps its vertical reading order.
  const shortSvg = render('flowchart TD\n  A[一] --> B[二]\n  B --> C[三]', 1152, 382).svg
  const shortRects = [...shortSvg.matchAll(/<rect x="([\d.]+)" y="([\d.]+)"/g)].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }))
  const shortSpreadY = Math.max(...shortRects.map((r) => r.y)) - Math.min(...shortRects.map((r) => r.y))
  assert(shortSpreadY > 0, 'short TD graph lost its vertical order')
  return `tall: x ${Math.round(spreadX)} > y ${Math.round(spreadY)}; short keeps y ${Math.round(shortSpreadY)}`
})

check('3f · a two-node cycle does not hang or throw', () => {
  const { svg, nodes } = render(['flowchart TD', '  A[甲] --> B[乙]', '  B --> A'].join('\n'))
  assert(nodes === 2, `nodes ${nodes}`)
  assert(svg.includes('甲') && svg.includes('乙'), 'cyclic nodes lost')
  return 'cycle rendered'
})

/* ------------------------------------------------------------------ *
 * Summary
 * ------------------------------------------------------------------ */

const failed = results.filter((result) => !result.ok)
console.log(`\n${results.length} checks, ${results.length - failed.length} passed, ${failed.length} failed`)
if (failed.length > 0) {
  console.log(`  FAILED: ${failed.map((result) => result.name).join('; ')}`)
  process.exitCode = 1
} else {
  console.log('ALL PASS')
}
