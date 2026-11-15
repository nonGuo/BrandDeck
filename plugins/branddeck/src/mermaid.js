/**
 * Mermaid (flowchart subset) → themed SVG.
 *
 * WHY THIS EXISTS: an architecture diagram is the one picture a text model can actually
 * author. It knows a pipeline is `A --> B`, it does not know that a box belongs at
 * x=312 and that the arrow has to leave the box's right edge. So the model writes
 * Mermaid — a language it is good at — and this module turns it into deterministic
 * geometry. The result is a plain SVG string, which means the whole existing pipeline
 * (theme-swap, 3x rasterisation, native-PPTX image placement) applies unchanged.
 *
 * WHY A SUBSET: there is no Mermaid/D2/Graphviz runtime in this profile (no mermaid, no
 * elkjs, no viz.js), and pulling a browser-based renderer in would put a Chromium in the
 * plugin's boot path. A flowchart renderer of ~400 lines that is *honest about its
 * limits* beats a dependency that cannot be loaded. Anything outside the subset throws a
 * clear error naming the offending line, so the caller can ask the model again.
 *
 * Supported:
 *   - `flowchart TD|TB|LR` / `graph TD|LR` (direction only changes the axis)
 *   - node shapes: `A[text]` rect, `A(text)` rounded, `A{text}` diamond
 *   - edges: `A --> B`, `A --- B`, `A -.-> B`, `A ==> B`, with `|label|` or `-- label -->`
 *   - `subgraph 名字` … `end` — drawn as a titled group box around its members
 *   - `%%` comments, blank lines
 * Deliberately rejected (loudly): sequence/class/state diagrams, `classDef`/`style`,
 * HTML labels, and anything else that would need a real Mermaid runtime.
 */

/** Thrown for input this renderer cannot draw; the message names the line. */
export class MermaidError extends Error {
  /** @param {string} message @param {number} [line] */
  constructor(message, line) {
    super(line === undefined ? message : `${message}（第 ${line} 行）`)
    this.name = 'MermaidError'
    this.code = 'mermaid_unsupported'
  }
}

/** Layout knobs. Sizes are in SVG units (= canvas px, so 1:1 with the DSL). */
export const MERMAID_GEOMETRY = Object.freeze({
  nodeMinWidth: 96,
  nodeHeight: 44,
  nodePadX: 16,
  /** Horizontal gap between nodes in the same rank (TD) / between ranks (LR). */
  gapMain: 28,
  /** Gap between ranks (TD) / nodes in a rank (LR). */
  gapRank: 56,
  groupPad: 18,
  groupTitleHeight: 22,
  labelSize: 13,
  fontSize: 14,
  arrowSize: 6,
  /** Upper bound when scaling a small graph up to fill the figure box. */
  maxScale: 2.4,
})

/** A shape this renderer understands. */
const SHAPES = Object.freeze({ '[': 'rect', '(': 'round', '{': 'diamond' })

/**
 * Parse Mermaid flowchart source into nodes, edges and groups.
 *
 * Order-preserving and deterministic: nodes are numbered in first-appearance order, which
 * is what makes the layout (and therefore the SVG string) stable across runs.
 *
 * @param {string} source
 * @returns {{
 *   direction: 'TD' | 'LR',
 *   nodes: Array<{id: string, text: string, shape: 'rect'|'round'|'diamond', group: number}>,
 *   edges: Array<{from: string, to: string, label: string | null, dashed: boolean}>,
 *   groups: Array<{title: string, members: string[]}>,
 * }}
 */
export function parseMermaid(source) {
  const text = String(source ?? '').replace(/\r\n?/g, '\n')
  const lines = text.split('\n')

  /** @type {Array<{id: string, text: string, shape: 'rect'|'round'|'diamond', group: number}>} */
  const nodes = []
  /** @type {Map<string, number>} */
  const index = new Map()
  /** @type {Array<{from: string, to: string, label: string | null, dashed: boolean}>} */
  const edges = []
  /** @type {Array<{title: string, members: string[]}>} */
  const groups = []
  /** @type {number[]} */
  const openGroups = []
  let direction = 'TD'
  let sawHeader = false

  /** @param {string} id @param {string | null} label @param {string | null} shape */
  const touchNode = (id, label, shape) => {
    const key = id.trim()
    if (key === '') throw new MermaidError('节点名为空')
    let position = index.get(key)
    if (position === undefined) {
      position = nodes.length
      index.set(key, position)
      nodes.push({
        id: key,
        text: label === null || label === '' ? key : label,
        shape: /** @type {'rect'|'round'|'diamond'} */ (shape ?? 'rect'),
        group: openGroups.length === 0 ? -1 : openGroups[openGroups.length - 1],
      })
      if (openGroups.length > 0) groups[openGroups[openGroups.length - 1]].members.push(key)
      return
    }
    // A later mention may add the label (`A --> B` then `B[文字]`): keep the first text
    // that is more than the bare id.
    const node = nodes[position]
    if ((node.text === node.id || node.text === '') && label !== null && label !== '') node.text = label
    if (shape !== null && shape !== undefined) node.shape = /** @type {'rect'|'round'|'diamond'} */ (shape)
  }

  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    const raw = lines[lineIndex]
    const line = raw.trim()
    if (line === '' || line.startsWith('%%')) continue

    if (sawHeader === false) {
      const header = /^(?:flowchart|graph)\s+(TD|TB|BT|LR|RL)?\s*$/i.exec(line)
      if (header === null) {
        throw new MermaidError('只支持 flowchart/graph 图，请以 `flowchart TD` 或 `flowchart LR` 开头', lineIndex + 1)
      }
      const value = (header[1] ?? 'TD').toUpperCase()
      // BT/RL are valid Mermaid but reverse the reading order; accept them and draw in
      // the same direction so the geometry code stays small.
      direction = value === 'LR' || value === 'RL' ? 'LR' : 'TD'
      sawHeader = true
      continue
    }

    if (/^subgraph\b/i.test(line)) {
      const title = line.replace(/^subgraph\s*/i, '').replace(/^["']|["']$/g, '').trim()
      groups.push({ title: title === '' ? `分组 ${groups.length + 1}` : title, members: [] })
      openGroups.push(groups.length - 1)
      continue
    }
    if (/^end$/i.test(line)) {
      if (openGroups.length === 0) throw new MermaidError('`end` 没有对应的 `subgraph`', lineIndex + 1)
      openGroups.pop()
      continue
    }
    if (/^(?:classDef|class|style|linkStyle|click|accTitle|accDescr)\b/i.test(line)) {
      throw new MermaidError('不支持装饰语法（classDef/style/click 等），请只用节点与连线', lineIndex + 1)
    }
    if (/<[a-z/][^>]*>/i.test(line)) {
      throw new MermaidError('不支持 HTML 标签，节点文字请用纯文本', lineIndex + 1)
    }

    // One line may chain edges: `A --> B --> C`.
    const parts = line.split(/\s*(?:-.->|==>|-->|---|->)\s*/)
    if (parts.length === 1) {
      // A bare node declaration: `A[文字]`.
      const node = parseNodeRef(parts[0], lineIndex + 1)
      touchNode(node.id, node.text, node.shape)
      continue
    }
    const operators = line.match(/-.->|==>|-->|---|->/g) ?? []
    for (let step = 0; step < parts.length; step += 1) {
      const parsed = parseNodeRef(parts[step], lineIndex + 1)
      touchNode(parsed.id, parsed.text, parsed.shape)
    }
    for (let step = 0; step < operators.length; step += 1) {
      const from = parseNodeRef(parts[step], lineIndex + 1).id
      const to = parseNodeRef(parts[step + 1], lineIndex + 1).id
      const operator = operators[step]
      const label = parseEdgeLabel(parts[step + 1], operator)
      edges.push({ from, to, label, dashed: operator === '-.->' })
    }
  }

  if (sawHeader === false) throw new MermaidError('空的图：请写 `flowchart TD` 与至少一条连线')
  if (nodes.length === 0) throw new MermaidError('图里没有节点')
  return { direction, nodes, edges, groups }
}

/**
 * `A[文字]` / `B(文字)` / `C{文字}` / `A` → id + text + shape.
 * @param {string} piece
 * @param {number} line
 */
function parseNodeRef(piece, line) {
  const text = piece.trim().replace(/^\|.*?\|/, '').trim()
  if (text === '') throw new MermaidError('连线缺少节点', line)
  const match = /^([A-Za-z0-9_\u4e00-\u9fa5][\w\u4e00-\u9fa5-]*)\s*(?:([\[({])([\s\S]*?)([\]})]))?$/.exec(text)
  if (match === null) throw new MermaidError(`看不懂的节点写法：${text}`, line)
  const id = match[1]
  const open = match[2]
  const close = match[4]
  if (open !== undefined) {
    const expected = open === '[' ? ']' : open === '(' ? ')' : '}'
    if (close !== expected) throw new MermaidError(`节点括号不匹配：${text}`, line)
  }
  const shape = open === undefined ? null : SHAPES[/** @type {'['|'('|'{'} */ (open)]
  const label = match[3] === undefined ? null : stripQuotes(match[3]).trim()
  return { id, text: label, shape }
}

/** `A -->|说明| B` carries its label on the target side. */
function parseEdgeLabel(piece, operator) {
  const match = /^\s*\|([^|]*)\|/.exec(piece)
  if (match !== null) return stripQuotes(match[1]).trim() || null
  // `A -- 说明 --> B` form.
  const dashed = new RegExp(`--\\s*([^-]+?)\\s*${operator.replace(/[-.]/g, '\\$&')}`).exec(piece)
  return dashed === null ? null : dashed[1].trim() || null
}

/** @param {string} value */
function stripQuotes(value) {
  return value.replace(/^["']|["']$/g, '')
}

/**
 * Assign each node a rank (its longest distance from a source) and place the boxes.
 *
 * Longest-path ranking keeps edges pointing "forward", which is what makes a layered
 * architecture diagram readable; cycles are tolerated by clamping the relaxation count.
 *
 * @param {ReturnType<typeof parseMermaid>} graph
 * @param {{width: number, height: number, fontSize?: number, labelSize?: number}} box
 * @param {{textPrimary: string, primary: string, background: string, surface: string}} colors
 */
function layout(graph, box, colors) {
  const geometry = MERMAID_GEOMETRY
  const fontSize = box.fontSize ?? geometry.fontSize
  const labelSize = box.labelSize ?? geometry.labelSize

  /** @param {string} text */
  const nodeWidth = (text) =>
    Math.max(geometry.nodeMinWidth, Math.round(charWidth(text, fontSize) + geometry.nodePadX * 2))
  /** @param {string} id */
  const byId = new Map(graph.nodes.map((node) => [node.id, node]))

  // Rank = longest path from a source. Relaxation is bounded by the node count so a cycle
  // cannot spin forever.
  /** @type {Map<string, number>} */
  const rank = new Map(graph.nodes.map((node) => [node.id, 0]))
  for (let pass = 0; pass < graph.nodes.length; pass += 1) {
    let changed = false
    for (const edge of graph.edges) {
      const from = rank.get(edge.from) ?? 0
      const to = rank.get(edge.to) ?? 0
      if (to < from + 1) {
        rank.set(edge.to, from + 1)
        changed = true
      }
    }
    if (!changed) break
  }

  /** @type {Map<number, string[]>} */
  const byRank = new Map()
  for (const node of graph.nodes) {
    const level = rank.get(node.id) ?? 0
    if (!byRank.has(level)) byRank.set(level, [])
    byRank.get(level).push(node.id)
  }
  const levels = [...byRank.keys()].sort((a, b) => a - b)
  const ranks = levels.map((level) => byRank.get(level) ?? [])

  // Two adjacent ranks that belong to different `subgraph`s need extra distance: each
  // group draws a titled box around its members, and without the gap the boxes overlap —
  // which reads as one smudged container instead of two layers.
  /** @param {string[]} ids */
  const rankGroup = (ids) => {
    const counts = new Map()
    for (const id of ids) {
      const group = byId.get(id)?.group ?? -1
      counts.set(group, (counts.get(group) ?? 0) + 1)
    }
    let best = -1
    let bestCount = 0
    for (const [group, count] of counts) {
      if (count > bestCount) {
        best = group
        bestCount = count
      }
    }
    return best
  }
  const groupPerRank = ranks.map(rankGroup)

  const heights = ranks.map((ids) =>
    Math.max(...ids.map((id) => (byId.get(id)?.shape === 'diamond' ? geometry.nodeHeight + 18 : geometry.nodeHeight))),
  )
  const widths = ranks.map((ids) => ids.reduce((sum, id) => sum + nodeWidth(byId.get(id)?.text ?? id), 0) + geometry.gapMain * Math.max(0, ids.length - 1))

  // A slide is a 16:9 band, so a *tall* top-down flow cannot be readable: seven ranks
  // stacked vertically shrink to a thread. When the source asks for TD but the graph is
  // long and thin, transpose it — the reading order is unchanged (each rank still follows
  // the previous one), only the axis differs.
  const longAndThin = ranks.length >= 5 && Math.max(...ranks.map((ids) => ids.length)) <= 3
  const vertical = graph.direction === 'TD' && longAndThin === false
  const totalMain = vertical ? heights.reduce((a, b) => a + b, 0) + geometry.gapRank * Math.max(0, heights.length - 1) : Math.max(...widths)
  const totalCross = vertical ? Math.max(...widths) : widths.reduce((a, b) => a + b, 0) + geometry.gapRank * Math.max(0, widths.length - 1)
  void totalMain
  void totalCross

  /** @type {Map<string, {x: number, y: number, w: number, h: number}>} */
  const placed = new Map()

  let mainCursor = 0
  ranks.forEach((ids, rankIndex) => {
    // `main` is the axis the ranks advance along (y for TD, x for LR); `cross` spreads the
    // nodes within one rank. Mixing them up is what drew an `LR` flow as a vertical stack:
    // every rank was centred on the same x.
    const rankWidth = widths[rankIndex]
    let crossCursor = vertical ? (totalCross - rankWidth) / 2 : 0
    const mainStart = mainCursor
    ids.forEach((id) => {
      const node = byId.get(id)
      const text = node?.text ?? id
      const w = nodeWidth(text)
      const h = node?.shape === 'diamond' ? geometry.nodeHeight + 18 : geometry.nodeHeight
      if (vertical) {
        placed.set(id, { x: crossCursor, y: mainStart + (heights[rankIndex] - h) / 2, w, h })
        crossCursor += w + geometry.gapMain
      } else {
        placed.set(id, { x: mainStart, y: crossCursor, w, h })
        crossCursor += h + geometry.gapMain
      }
    })
    const nextIndex = rankIndex + 1
    const switchingGroup =
      nextIndex < ranks.length &&
      groupPerRank[rankIndex] !== -1 &&
      groupPerRank[nextIndex] !== -1 &&
      groupPerRank[rankIndex] !== groupPerRank[nextIndex]
    const extra = switchingGroup
      ? geometry.groupPad * 2 + geometry.groupTitleHeight
      : 0
    mainCursor += (vertical ? heights[rankIndex] : widths[rankIndex]) + geometry.gapRank + extra
  })

  return { placed, labelSize, fontSize }
}

/**
 * Render a Mermaid flowchart as an SVG string, themed with the deck's tokens.
 *
 * The SVG is deliberately plain (rect/rounded rect/polygon/line/path/text) so the
 * existing SVG rasteriser and the theme-swap pass can both read it.
 *
 * @param {string} source Mermaid source
 * @param {{width: number, height: number, colors: {textPrimary: string, primary: string, background: string, surface: string, textSecondary?: string}, fontSize?: number}} options
 * @returns {{svg: string, nodes: number, edges: number, groups: number}}
 */
export function mermaidToSvg(source, options) {
  const graph = parseMermaid(source)
  const box = { width: Math.max(120, options.width), height: Math.max(80, options.height) }
  const colors = options.colors
  const { placed, fontSize, labelSize } = layout(graph, { ...box, fontSize: options.fontSize }, colors)

  /** @type {string[]} */
  const parts = []
  const esc = (value) =>
    String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')

  // Group boxes first (behind everything), sized to their members plus padding.
  graph.groups.forEach((group) => {
    const members = group.members
      .map((id) => placed.get(id))
      .filter((point) => point !== undefined)
    if (members.length === 0) return
    const pad = MERMAID_GEOMETRY.groupPad
    const x = Math.min(...members.map((p) => p.x)) - pad
    const y = Math.min(...members.map((p) => p.y)) - pad - MERMAID_GEOMETRY.groupTitleHeight
    const w = Math.max(...members.map((p) => p.x + p.w)) + pad - x
    const h = Math.max(...members.map((p) => p.y + p.h)) + pad - y
    parts.push(
      `<rect x="${round(x)}" y="${round(y)}" width="${round(w)}" height="${round(h)}" rx="10" ` +
        `fill="${colors.surface}" fill-opacity="0.55" stroke="${colors.surface}" stroke-width="1" />`,
    )
    parts.push(
      `<text x="${round(x + 12)}" y="${round(y + 16)}" font-family="inherit" font-size="${labelSize}" ` +
        `fill="${colors.textSecondary ?? colors.textPrimary}" fill-opacity="0.85">${esc(group.title)}</text>`,
    )
  })

  // Edges, then nodes: an arrowhead must sit on top of the border it lands on.
  graph.edges.forEach((edge) => {
    const from = placed.get(edge.from)
    const to = placed.get(edge.to)
    if (from === undefined || to === undefined) return
    const [start, end] = edgeEndpoints(from, to)
    const dash = edge.dashed ? ' stroke-dasharray="5 4"' : ''
    const midX = (start.x + end.x) / 2
    const midY = (start.y + end.y) / 2
    parts.push(
      `<line x1="${round(start.x)}" y1="${round(start.y)}" x2="${round(end.x)}" y2="${round(end.y)}" ` +
        `stroke="${colors.textSecondary ?? colors.textPrimary}" stroke-width="1.4"${dash} />`,
    )
    parts.push(arrowHead(start, end, colors.textSecondary ?? colors.textPrimary))
    if (edge.label !== null) {
      parts.push(
        `<text x="${round(midX + 6)}" y="${round(midY - 6)}" font-family="inherit" font-size="${labelSize}" ` +
          `fill="${colors.textSecondary ?? colors.textPrimary}">${esc(edge.label)}</text>`,
      )
    }
  })

  for (const node of graph.nodes) {
    const rect = placed.get(node.id)
    if (rect === undefined) continue
    const fill = colors.background
    const stroke = node.shape === 'diamond' ? colors.primary : colors.textPrimary
    if (node.shape === 'diamond') {
      parts.push(
        `<polygon points="${diamondPoints(rect)}" fill="${fill}" stroke="${stroke}" stroke-width="1.4" />`,
      )
    } else {
      const radius = node.shape === 'round' ? Math.round(rect.h / 2) : 6
      parts.push(
        `<rect x="${round(rect.x)}" y="${round(rect.y)}" width="${round(rect.w)}" height="${round(rect.h)}" ` +
          `rx="${radius}" fill="${fill}" stroke="${stroke}" stroke-width="1.4" />`,
      )
    }
    parts.push(
      `<text x="${round(rect.x + rect.w / 2)}" y="${round(rect.y + rect.h / 2 + fontSize * 0.35)}" ` +
        `text-anchor="middle" font-family="inherit" font-size="${fontSize}" fill="${colors.textPrimary}">${esc(node.text)}</text>`,
    )
  }

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.round(box.width)}" height="${Math.round(box.height)}" ` +
    `viewBox="0 0 ${Math.round(box.width)} ${Math.round(box.height)}">` +
    `<g transform="${fitTransform(graph, placed, box)}">${parts.join('')}</g></svg>`
  return { svg, nodes: graph.nodes.length, edges: graph.edges.length, groups: graph.groups.length }
}

/**
 * A `translate + scale` that fits the drawing into the box.
 *
 * Mermaid has no idea how big our slide is, and the model writes 4–9 nodes without
 * counting pixels, so the drawn extent is measured and scaled. Scaling *up* matters as
 * much as down: a linear four-node flow drawn at its natural size is a thin thread in the
 * middle of a 1152px-wide figure. The SVG is vector, so scaling up keeps the type crisp.
 *
 * @param {ReturnType<typeof parseMermaid>} graph
 * @param {Map<string, {x: number, y: number, w: number, h: number}>} placed
 * @param {{width: number, height: number}} box
 */
function fitTransform(graph, placed, box) {
  void graph
  const points = [...placed.values()]
  if (points.length === 0) return 'translate(0,0)'
  const pad = MERMAID_GEOMETRY.groupPad * 2 + MERMAID_GEOMETRY.groupTitleHeight
  const minX = Math.min(...points.map((p) => p.x)) - pad
  const maxX = Math.max(...points.map((p) => p.x + p.w)) + pad
  const minY = Math.min(...points.map((p) => p.y)) - pad
  const maxY = Math.max(...points.map((p) => p.y + p.h)) + pad
  const raw = Math.min(box.width / (maxX - minX), box.height / (maxY - minY))
  const scale = Math.min(MERMAID_GEOMETRY.maxScale, Math.max(0.35, raw))
  const offsetX = (box.width - (maxX - minX) * scale) / 2 - minX * scale
  const offsetY = (box.height - (maxY - minY) * scale) / 2 - minY * scale
  return `translate(${round(offsetX)},${round(offsetY)}) scale(${round(scale * 1000) / 1000})`
}

/**
 * Pick the two points where an edge leaves one box and enters another, so the line does
 * not start in the middle of a rectangle.
 */
function edgeEndpoints(from, to) {
  const fromCenter = { x: from.x + from.w / 2, y: from.y + from.h / 2 }
  const toCenter = { x: to.x + to.w / 2, y: to.y + to.h / 2 }
  const dx = toCenter.x - fromCenter.x
  const dy = toCenter.y - fromCenter.y
  if (Math.abs(dx) >= Math.abs(dy)) {
    const sign = dx >= 0 ? 1 : -1
    return [
      { x: fromCenter.x + (sign * from.w) / 2, y: fromCenter.y },
      { x: toCenter.x - (sign * to.w) / 2, y: toCenter.y },
    ]
  }
  const sign = dy >= 0 ? 1 : -1
  return [
    { x: fromCenter.x, y: fromCenter.y + (sign * from.h) / 2 },
    { x: toCenter.x, y: toCenter.y - (sign * to.h) / 2 },
  ]
}

/** @param {{x: number, y: number}} start @param {{x: number, y: number}} end @param {string} color */
function arrowHead(start, end, color) {
  const size = MERMAID_GEOMETRY.arrowSize
  const angle = Math.atan2(end.y - start.y, end.x - start.x)
  const left = `${round(end.x - size * Math.cos(angle - Math.PI / 7))},${round(end.y - size * Math.sin(angle - Math.PI / 7))}`
  const right = `${round(end.x - size * Math.cos(angle + Math.PI / 7))},${round(end.y - size * Math.sin(angle + Math.PI / 7))}`
  return `<polygon points="${round(end.x)},${round(end.y)} ${left} ${right}" fill="${color}" />`
}

/** @param {{x: number, y: number, w: number, h: number}} rect */
function diamondPoints(rect) {
  const cx = rect.x + rect.w / 2
  const cy = rect.y + rect.h / 2
  return [
    `${round(cx)},${round(rect.y)}`,
    `${round(rect.x + rect.w)},${round(cy)}`,
    `${round(cx)},${round(rect.y + rect.h)}`,
    `${round(rect.x)},${round(cy)}`,
  ].join(' ')
}

/**
 * Rough single-line width of a label. CJK glyphs are ~1 em wide, Latin ~0.55 em; the
 * estimate only has to be good enough to keep boxes from overlapping.
 * @param {string} text @param {number} fontSize
 */
function charWidth(text, fontSize) {
  let width = 0
  for (const char of String(text)) width += /[\u3000-\u9fff\uff00-\uffef]/.test(char) ? fontSize : fontSize * 0.56
  return width
}

/** @param {number} value */
function round(value) {
  return Math.round(value * 10) / 10
}
