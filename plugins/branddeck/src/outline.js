/**
 * BrandDeck Markdown outline projection ("文字版本") and its inverse.
 *
 * `card.json` stays the geometry source of truth; `outline.md` is the content
 * source. A human edits the Markdown and `applyOutline` writes the edited text
 * back onto the deck while preserving every untouched layout box.
 *
 * OUTPUT FORMAT (stable, byte-for-byte, no trailing whitespace, one trailing LF):
 *
 *   # <deck.title>
 *
 *   > deck: <deckId> · template: <templateId> · 画布 1280×720
 *
 *   ## <slide title> {#slideId}
 *
 *   ### 标题 {#cover-title}
 *   AI 驱动的企业内容生产
 *
 *   ### 副标题 {#cover-subtitle}
 *   从企业模板到原生可编辑演示文稿
 *
 *   <!-- 本页另有 3 个非文字元素（图形/图表/SVG），请在卡片视图中编辑 -->
 *
 *   备注：开场说明…
 *
 * Format rules:
 * - `# ` (exactly one level-1 heading) is the deck title. The `> ` block is a
 *   machine header: regenerated on output, ignored on input.
 * - `## <title> {#slideId}` starts a slide; the heading text also becomes
 *   `slide.title`. A slide without `title` is emitted as `## {#slideId}` (empty
 *   heading text), so a round-trip never invents a title.
 * - `### <role label> {#elementId}` starts one text element. Role labels:
 *   标题→title, 副标题→subtitle, 正文→body, 注释→caption; the English role names
 *   are accepted too. An unknown label keeps the matched element's role, or
 *   `body` for a newly added element.
 * - Element content is every line until the next heading, `备注：` or `<!--`.
 *   Lines are joined with `\n`; surrounding blank lines are trimmed.
 * - `备注：<text>` (also `备注:` and `Notes:`) sets `slide.speakerNotes`.
 * - Exactly one `<!-- ... -->` comment per slide, only when the slide has
 *   non-text elements (shapes/charts/SVG/images). Counted, but never editable
 *   here. Comments are ignored on input.
 * - Lines starting with `#`, `备注[:：]`, `Notes:`, `<!--` or `>` are reserved:
 *   inside element content they start a new block (or are dropped), so copy
 *   that begins with those markers must be reworded.
 *
 * APPLY SEMANTICS:
 * - Text elements are matched per slide by `{#id}`. Matched elements only get a
 *   new `content` (and adopt the parsed role when the label maps explicitly);
 *   `x/y/w/h/style` are never touched, so layout is preserved.
 * - An unknown id ADDS one text element with deterministic auto-layout inside
 *   `innerBox(theme)`, below/around existing text boxes with >= 16px vertical
 *   gaps. Ids are `<slideId>-t<k>` with `k` the next free integer. A warning is
 *   recorded per addition.
 * - Deck text elements that the Markdown does not mention are KEPT by default;
 *   `options.prune === true` removes them (and warns per removal). A slide that
 *   would become empty keeps its last element.
 * - The function never throws. Odd input produces warnings and leaves the deck
 *   (or the affected slide) unchanged.
 *
 * Judgement calls (documented deviations from a strictly minimal reading):
 * - The `{#id}` written on a *new* element heading is only a request: the added
 *   element always receives the deterministic `<slideId>-t<k>` id demanded by
 *   the spec, and the warning reports that generated id.
 * - Auto-layout avoids existing *text* boxes (the spec's guarantee), not shapes,
 *   charts or images; decorative backdrops are expected to sit under text.
 * - `changes.untouched` counts text elements that were kept as-is (matched with
 *   identical content, or absent from the outline and kept).
 */
import { CANVAS, estimateLines, innerBox, resolveColor } from './schema.js'
import { makeDefaultTheme } from './defaults.js'

const ROLE_LABELS = Object.freeze({
  zh: Object.freeze({ title: '标题', subtitle: '副标题', body: '正文', caption: '注释' }),
  en: Object.freeze({ title: 'Title', subtitle: 'Subtitle', body: 'Body', caption: 'Caption' }),
})

const ROLE_ALIASES = new Map([
  ['标题', 'title'],
  ['副标题', 'subtitle'],
  ['正文', 'body'],
  ['注释', 'caption'],
  ['title', 'title'],
  ['subtitle', 'subtitle'],
  ['body', 'body'],
  ['caption', 'caption'],
])

/**
 * Font sizes and colours used when the outline adds a brand-new text element.
 * Band offsets are derived from these so the projection stays deterministic.
 */
const AUTO_LAYOUT = Object.freeze({
  title: Object.freeze({ fontSize: 44, bold: true, colorToken: '$colors.textPrimary' }),
  subtitle: Object.freeze({ fontSize: 24, bold: false, colorToken: '$colors.textPrimary' }),
  body: Object.freeze({ fontSize: 20, bold: false, colorToken: '$colors.textPrimary' }),
  caption: Object.freeze({ fontSize: 14, bold: false, colorToken: '$colors.textSecondary' }),
})

const AUTO_ORDER = Object.freeze(['title', 'subtitle', 'body', 'caption'])
const AUTO_GAP = 16
const FALLBACK_TEXT_COLOR = '#1D1D1A'
const FALLBACK_SECONDARY_COLOR = '#666666'

/** Vertical start of each role band, relative to the top of the safe area. */
const AUTO_BANDS = (() => {
  /** @type {Record<string, number>} */
  const bands = {}
  let offset = 0
  for (const role of AUTO_ORDER) {
    bands[role] = offset
    offset += Math.ceil(AUTO_LAYOUT[role].fontSize * 1.35) + AUTO_GAP
  }
  return Object.freeze(bands)
})()

/**
 * Project a deck as Markdown.
 *
 * @param {import('./schema.js').CardDeck} deck
 * @param {import('./schema.js').ThemeTokens} [theme]
 * @param {{labels?: 'zh' | 'en'}} [options]
 * @returns {string} Markdown ending in exactly one newline
 */
export function deckToOutline(deck, theme, options = {}) {
  const labels = options?.labels === 'en' ? ROLE_LABELS.en : ROLE_LABELS.zh
  const title = textOf(deck?.title).trim()
  const deckId = textOf(deck?.deckId)
  const templateId = textOf(deck?.templateId) || textOf(theme?.templateId)
  /** @type {string[]} */
  const lines = []

  lines.push(title === '' ? '#' : `# ${title}`)
  lines.push('')
  lines.push(`> deck: ${deckId} · template: ${templateId} · 画布 ${CANVAS.width}×${CANVAS.height}`)

  const slides = Array.isArray(deck?.slides) ? deck.slides : []
  for (const slide of slides) {
    const slideId = textOf(slide?.slideId)
    const heading = oneLine(textOf(slide?.title).trim())
    lines.push('')
    lines.push(heading === '' ? `## {#${slideId}}` : `## ${heading} {#${slideId}}`)

    for (const element of Array.isArray(slide?.elements) ? slide.elements : []) {
      if (element === null || typeof element !== 'object' || element.type !== 'text') continue
      lines.push('')
      lines.push(`### ${labels[element.role] ?? labels.body} {#${textOf(element.id)}}`)
      const contentLines = trimBlankLines(textOf(element.content).replace(/\r\n?/g, '\n').split('\n'))
      for (const contentLine of contentLines) lines.push(rstrip(contentLine))
    }

    const nonText = nonTextCount(slide)
    if (nonText > 0) {
      lines.push('')
      lines.push(
        `<!-- 本页另有 ${nonText} 个非文字元素（图形/图表/SVG），请在卡片视图中编辑 -->`,
      )
    }

    const notes = textOf(slide?.speakerNotes)
    if (notes.trim() !== '') {
      const noteLines = notes.replace(/\r\n?/g, '\n').split('\n')
      lines.push('')
      lines.push(`备注：${rstrip(noteLines[0])}`)
      for (const noteLine of noteLines.slice(1)) lines.push(rstrip(noteLine))
    }
  }

  while (lines.length > 0 && lines[lines.length - 1].trim() === '') lines.pop()
  return `${lines.map(rstrip).join('\n')}\n`
}

/**
 * Apply an edited Markdown outline onto a deck without touching preserved
 * geometry. Never throws; all oddities are reported as warnings.
 *
 * @param {string} markdown
 * @param {import('./schema.js').CardDeck} deck
 * @param {import('./schema.js').ThemeTokens} [theme]
 * @param {{prune?: boolean}} [options]
 * @returns {{
 *   deck: import('./schema.js').CardDeck,
 *   warnings: string[],
 *   changes: {updated: number, added: number, removed: number, untouched: number},
 * }}
 */
export function applyOutline(markdown, deck, theme, options = {}) {
  /** @type {string[]} */
  const warnings = []
  /** @type {{updated: number, added: number, removed: number, untouched: number}} */
  const changes = { updated: 0, added: 0, removed: 0, untouched: 0 }
  const prune = options?.prune === true
  const tokens = theme !== null && typeof theme === 'object' && theme.colors ? theme : makeDefaultTheme()

  if (typeof markdown !== 'string') {
    warnings.push('deck · outline is not a string; deck left unchanged')
    return { deck, warnings, changes }
  }
  /** @type {import('./schema.js').CardDeck} */
  let next
  try {
    next = structuredClone(deck)
  } catch {
    warnings.push('deck · deck could not be copied; deck left unchanged')
    return { deck, warnings, changes }
  }
  if (next === null || typeof next !== 'object' || !Array.isArray(next.slides)) {
    warnings.push('deck · deck has no slides array; deck left unchanged')
    return { deck, warnings, changes }
  }

  try {
    const parsed = parseOutline(markdown, warnings)
    if (parsed.slides.length === 0) {
      warnings.push(
        markdown.trim() === ''
          ? 'deck · outline is empty; deck left unchanged'
          : 'deck · no slide sections found; deck left unchanged',
      )
      return { deck, warnings, changes }
    }
    if (parsed.title === null) {
      warnings.push('deck · missing "# title" heading; deck title kept')
    }

    /** @type {Map<string, {slide: any, index: number}>} */
    const slidesById = new Map()
    next.slides.forEach((slide, index) => {
      if (slide !== null && typeof slide === 'object' && typeof slide.slideId === 'string') {
        if (!slidesById.has(slide.slideId)) slidesById.set(slide.slideId, { slide, index })
      }
    })

    const seenSections = new Set()
    /** @type {Map<string, any>} */
    const matchedSections = new Map()
    let appliedAny = false
    for (const section of parsed.slides) {
      if (seenSections.has(section.slideId)) {
        warnings.push(`deck · duplicate slide section "${section.slideId}" ignored`)
        continue
      }
      seenSections.add(section.slideId)
      const entry = slidesById.get(section.slideId)
      if (entry === undefined) {
        warnings.push(`deck · unknown slide "${section.slideId}" in outline; section ignored`)
        continue
      }
      matchedSections.set(section.slideId, section)
      appliedAny = true
    }

    if (!appliedAny) {
      warnings.push('deck · no slide section matched the deck; deck left unchanged')
      return { deck, warnings, changes }
    }

    const appliedSections = new Set()
    next.slides.forEach((slide, index) => {
      if (slide === null || typeof slide !== 'object') return
      const section = matchedSections.get(slide.slideId)
      /** @type {Set<string>} */
      let mentioned = new Set()
      /** @type {Set<string>} */
      let addedIds = new Set()
      if (section !== undefined && !appliedSections.has(section.slideId)) {
        appliedSections.add(section.slideId)
        const applied = applySlideSection(section, slide, index, tokens, warnings, changes)
        mentioned = applied.mentioned
        addedIds = applied.addedIds
      }
      // Text elements of slides the outline never mentions are still "absent
      // from the markdown": kept by default, pruned on demand.
      settleTextElements(slide, index, mentioned, addedIds, prune, warnings, changes)
    })

    if (parsed.title !== null) next.title = parsed.title
    return { deck: next, warnings, changes }
  } catch (error) {
    warnings.push(`deck · outline could not be applied (${errorMessage(error)}); deck left unchanged`)
    return { deck, warnings, changes }
  }
}

/**
 * @param {string} markdown
 * @param {string[]} warnings
 */
function parseOutline(markdown, warnings) {
  /** @type {{title: string | null, slides: any[]}} */
  const result = { title: null, slides: [] }
  const source = markdown.replace(/\r\n?/g, '\n')
  const lines = source.split('\n')

  /** @type {any | null} */
  let slide = null
  /** @type {any | null} */
  let block = null
  let ignoredSlide = false
  let inNotes = false
  let inComment = false
  let titleSeen = false
  let outsideWarned = false

  const warnOutside = () => {
    if (outsideWarned) return
    outsideWarned = true
    warnings.push('deck · text outside an element block ignored')
  }

  for (const rawLine of lines) {
    const line = rstrip(rawLine)
    const trimmed = line.trim()

    if (inComment) {
      if (trimmed.includes('-->')) inComment = false
      continue
    }
    if (trimmed.startsWith('<!--')) {
      if (!trimmed.includes('-->')) inComment = true
      continue
    }
    // The `> deck: …` machine header (and any blockquote) is regenerated on
    // output and ignored on input, including unknown variants.
    if (trimmed.startsWith('>')) continue

    if (trimmed === '') {
      if (block !== null) block.lines.push('')
      else if (inNotes && slide !== null) slide.notesLines.push('')
      continue
    }

    const heading = matchHeading(trimmed)
    if (heading !== null && heading.level === 1) {
      block = null
      inNotes = false
      if (!titleSeen && slide === null && heading.text !== '') {
        result.title = heading.text
        titleSeen = true
      } else if (heading.text !== '') {
        warnings.push('deck · extra "# " heading ignored')
      }
      continue
    }

    if (heading !== null && heading.level === 2) {
      block = null
      inNotes = false
      if (heading.id === '') {
        warnings.push(
          `deck · slide heading "${heading.text}" has no {#slideId}; section ignored`,
        )
        slide = null
        ignoredSlide = true
        continue
      }
      slide = {
        slideId: heading.id,
        title: heading.text,
        blocks: [],
        notesLines: [],
        notesSeen: false,
      }
      result.slides.push(slide)
      ignoredSlide = false
      continue
    }

    if (heading !== null && heading.level === 3) {
      inNotes = false
      if (slide === null) {
        block = null
        if (!ignoredSlide) {
          warnings.push(
            `deck · element heading "${heading.text}" appears before any slide section; ignored`,
          )
        }
        continue
      }
      if (heading.id === '') {
        block = null
        warnings.push(
          `slide ${result.slides.length} · element heading "${heading.text}" has no {#id}; section ignored`,
        )
        continue
      }
      block = {
        id: heading.id,
        labelText: heading.text,
        roleHint: roleFromLabel(heading.text),
        lines: [],
        content: '',
      }
      slide.blocks.push(block)
      continue
    }

    if (heading !== null && heading.level >= 4) {
      // A deeper heading is indistinguishable from text; keep it as content.
      if (block !== null) block.lines.push(line)
      else warnOutside()
      continue
    }

    const notesText = matchNotesStart(trimmed)
    if (notesText !== null) {
      block = null
      if (slide === null) {
        warnings.push('deck · "备注：" outside a slide section ignored')
        continue
      }
      if (slide.notesSeen) {
        warnings.push(`slide ${result.slides.length} · duplicate "备注：" block; the last one wins`)
      }
      slide.notesSeen = true
      slide.notesLines = [notesText]
      inNotes = true
      continue
    }

    if (block !== null) {
      block.lines.push(line)
    } else if (inNotes && slide !== null) {
      slide.notesLines.push(line)
    } else {
      warnOutside()
      if (slide === null) continue
      // Plain prose inside a slide heading but before any `###` block: ignore.
    }
  }

  for (const parsedSlide of result.slides) {
    for (const parsedBlock of parsedSlide.blocks) {
      parsedBlock.content = trimBlankLines(parsedBlock.lines).join('\n')
      delete parsedBlock.lines
    }
    parsedSlide.notes = parsedSlide.notesSeen
      ? trimBlankLines(parsedSlide.notesLines).join('\n')
      : null
    delete parsedSlide.notesLines
    delete parsedSlide.notesSeen
  }
  return result
}

/**
 * Apply one parsed slide section: title, speaker notes, matched content and
 * newly added elements. Returns the element ids the markdown accounted for so
 * the caller can settle the remaining text elements.
 *
 * @param {any} section
 * @param {any} slide
 * @param {number} index zero-based deck slide index
 * @param {import('./schema.js').ThemeTokens} theme
 * @param {string[]} warnings
 * @param {{updated: number, added: number, removed: number, untouched: number}} changes
 * @returns {{mentioned: Set<string>, addedIds: Set<string>}}
 */
function applySlideSection(section, slide, index, theme, warnings, changes) {
  const label = `slide ${index + 1}`
  if (!Array.isArray(slide.elements)) slide.elements = []
  const elements = slide.elements

  if (section.title !== '') slide.title = section.title

  if (section.notes !== null) {
    const notes = section.notes.trim()
    if (notes === '') delete slide.speakerNotes
    else slide.speakerNotes = notes
  }

  const box = innerBox(theme)
  /** @type {{x: number, y: number, w: number, h: number}[]} */
  const textBoxes = []
  /** @type {Set<string>} */
  const takenIds = new Set()
  for (const element of elements) {
    if (element === null || typeof element !== 'object') continue
    if (typeof element.id === 'string') takenIds.add(element.id)
    if (element.type === 'text') {
      textBoxes.push({ x: element.x, y: element.y, w: element.w, h: element.h })
    }
  }

  const mentioned = new Set()
  const addedIds = new Set()
  const seenBlocks = new Set()

  for (const parsedBlock of section.blocks) {
    const id = parsedBlock.id
    if (seenBlocks.has(id)) {
      warnings.push(`${label} · duplicate element "${id}" in outline; only the first block was used`)
      continue
    }
    seenBlocks.add(id)

    const target = elements.find((element) => element !== null && element.id === id)
    const labelText = parsedBlock.labelText
    const unknownLabel = labelText !== '' && parsedBlock.roleHint === null

    if (parsedBlock.content === '') {
      warnings.push(`${label} · element "${id}" has no content; skipped`)
      if (target !== undefined) {
        mentioned.add(id)
        changes.untouched += 1
      }
      continue
    }

    if (target !== undefined) {
      if (target.type !== 'text') {
        warnings.push(`${label} · element "${id}" is not a text element; skipped`)
        continue
      }
      if (unknownLabel) {
        warnings.push(
          `${label} · element "${id}": unknown role label "${labelText}"; role "${target.role}" kept`,
        )
      }
      const role = parsedBlock.roleHint ?? target.role
      mentioned.add(id)
      if (target.content !== parsedBlock.content || target.role !== role) {
        target.content = parsedBlock.content
        target.role = role
        changes.updated += 1
      } else {
        changes.untouched += 1
      }
      continue
    }

    const role = parsedBlock.roleHint ?? 'body'
    if (unknownLabel) {
      warnings.push(
        `${label} · element "${id}": unknown role label "${labelText}"; role "body" used`,
      )
    }
    const addedId = nextTextId(slide.slideId, takenIds)
    takenIds.add(addedId)
    const placement = autoLayoutText(role, parsedBlock.content, box, textBoxes)
    const element = {
      type: /** @type {const} */ ('text'),
      id: addedId,
      x: placement.x,
      y: placement.y,
      w: placement.w,
      h: placement.h,
      role,
      content: parsedBlock.content,
      style: autoStyle(role, theme),
    }
    elements.push(element)
    addedIds.add(addedId)
    textBoxes.push({ x: element.x, y: element.y, w: element.w, h: element.h })
    changes.added += 1
    warnings.push(`${label} · added element ${addedId} (auto-layout)`)
    if (placement.crowded) {
      warnings.push(
        `${label} · element "${addedId}" has no free space in the safe area; placed at the bottom`,
      )
    }
  }

  return { mentioned, addedIds }
}

/**
 * Text elements absent from the markdown are kept by default and pruned with
 * `options.prune`. A slide always keeps at least one element (schema floor).
 *
 * @param {any} slide
 * @param {number} index zero-based deck slide index
 * @param {Set<string>} mentioned
 * @param {Set<string>} addedIds
 * @param {boolean} prune
 * @param {string[]} warnings
 * @param {{updated: number, added: number, removed: number, untouched: number}} changes
 */
function settleTextElements(slide, index, mentioned, addedIds, prune, warnings, changes) {
  if (!Array.isArray(slide.elements)) slide.elements = []
  const label = `slide ${index + 1}`
  const elements = slide.elements
  for (const element of [...elements]) {
    if (element === null || typeof element !== 'object' || element.type !== 'text') continue
    if (mentioned.has(element.id) || addedIds.has(element.id)) continue
    if (!prune) {
      changes.untouched += 1
      continue
    }
    if (elements.length <= 1) {
      warnings.push(`${label} · element "${element.id}" kept (last element on the slide)`)
      changes.untouched += 1
      continue
    }
    const at = elements.indexOf(element)
    if (at === -1) continue
    elements.splice(at, 1)
    changes.removed += 1
    warnings.push(`${label} · pruned element "${element.id}"`)
  }
}

/**
 * Deterministic auto-layout for one added text element: role band first, then
 * pushed below any text box it would collide with (>= AUTO_GAP vertical gap),
 * always inside `innerBox(theme)`.
 *
 * @param {string} role
 * @param {string} content
 * @param {ReturnType<typeof innerBox>} box
 * @param {{x: number, y: number, w: number, h: number}[]} textBoxes
 */
function autoLayoutText(role, content, box, textBoxes) {
  const spec = AUTO_LAYOUT[role] ?? AUTO_LAYOUT.body
  const width = autoWidth(role, box)
  const estimated = estimateLines(content, width, spec.fontSize).heightPx
  const height = Math.min(box.h, Math.max(Math.ceil(spec.fontSize * 1.35), Math.ceil(estimated)))
  const bottom = box.y + box.h
  const bandTop = box.y + (AUTO_BANDS[role] ?? AUTO_BANDS.body)

  /** @type {number[]} */
  const candidates = [bandTop]
  for (const textBox of textBoxes) {
    if (!Number.isFinite(textBox.y) || !Number.isFinite(textBox.h)) continue
    candidates.push(textBox.y + textBox.h + AUTO_GAP)
  }
  const ordered = [...new Set(candidates.map((value) => Math.round(value)))]
    .filter((value) => value >= box.y)
    .sort((a, b) => a - b)

  for (const candidate of ordered) {
    if (candidate + height > bottom) continue
    const rect = { x: box.x, y: candidate, w: width, h: height }
    if (!textBoxes.some((textBox) => collidesWithGap(rect, textBox, AUTO_GAP))) {
      return { ...rect, crowded: false }
    }
  }

  const y = Math.max(box.y, bottom - height)
  return { x: box.x, y, w: width, h: height, crowded: true }
}

/** @param {string} role @param {ReturnType<typeof innerBox>} box */
function autoWidth(role, box) {
  const logo = box.logo
  if (role === 'title' && logo !== null && logo !== undefined) {
    return Math.max(160, box.w - logo.width - logo.gap)
  }
  return box.w
}

/**
 * @param {{x: number, y: number, w: number, h: number}} rect
 * @param {{x: number, y: number, w: number, h: number}} other
 * @param {number} gap
 */
function collidesWithGap(rect, other, gap) {
  const horizontal = rect.x < other.x + other.w && rect.x + rect.w > other.x
  if (!horizontal) return false
  const verticalGap = Math.max(
    other.y - (rect.y + rect.h),
    rect.y - (other.y + other.h),
  )
  return verticalGap < gap
}

/**
 * @param {string} role
 * @param {import('./schema.js').ThemeTokens} theme
 */
function autoStyle(role, theme) {
  const spec = AUTO_LAYOUT[role] ?? AUTO_LAYOUT.body
  /** @type {{fontSize: number, color: string, bold?: boolean}} */
  const style = {
    fontSize: spec.fontSize,
    color: resolveColor(spec.colorToken, theme, FALLBACK_TEXT_COLOR),
  }
  if (spec.bold) style.bold = true
  return style
}

/**
 * Deterministic `<slideId>-t<k>` id where `k` is the next free integer.
 * @param {unknown} slideId
 * @param {Set<string>} taken
 */
function nextTextId(slideId, taken) {
  const base = typeof slideId === 'string' && slideId !== '' ? slideId : 'slide'
  for (let k = 1; k <= 10000; k += 1) {
    const candidate = base.length + `-t${k}`.length <= 64 ? `${base}-t${k}` : `t${k}`
    if (!taken.has(candidate)) return candidate
  }
  return `t${taken.size + 1}`
}

/**
 * Parse `## Title {#id}` / `### Label {#id}` / `# Title`.
 * @param {string} line
 */
function matchHeading(line) {
  const match = /^(#{1,6})(?:\s+(.*))?$/.exec(line)
  if (match === null) return null
  let text = (match[2] ?? '').trim()
  let id = ''
  const idMatch = /\{#([^}]*)\}\s*$/.exec(text)
  if (idMatch !== null) {
    id = idMatch[1].trim()
    text = text.slice(0, idMatch.index).trim()
  }
  return { level: match[1].length, text, id }
}

/** @param {string} line @returns {string | null} leading notes text */
function matchNotesStart(line) {
  const zh = /^备注\s*[:：]\s?(.*)$/.exec(line)
  if (zh !== null) return zh[1]
  const en = /^notes\s*:\s?(.*)$/i.exec(line)
  if (en !== null) return en[1]
  return null
}

/** @param {string} label @returns {'title' | 'subtitle' | 'body' | 'caption' | null} */
function roleFromLabel(label) {
  const text = textOf(label).trim()
  if (text === '') return null
  const direct = ROLE_ALIASES.get(text)
  if (direct !== undefined) return direct
  return ROLE_ALIASES.get(text.toLowerCase()) ?? null
}

/** @param {any} slide */
function nonTextCount(slide) {
  const elements = Array.isArray(slide?.elements) ? slide.elements : []
  let count = 0
  for (const element of elements) {
    if (element !== null && typeof element === 'object' && element.type !== 'text') count += 1
  }
  return count
}

/** @param {unknown} value */
function textOf(value) {
  return typeof value === 'string' ? value : ''
}

/** @param {string} value */
function oneLine(value) {
  return value.replace(/\r?\n/g, ' ').trim()
}

/** @param {string} value */
function rstrip(value) {
  return value.replace(/[ \t]+$/, '')
}

/** Drop leading/trailing blank lines, keep interior ones. @param {string[]} lines */
function trimBlankLines(lines) {
  let start = 0
  let end = lines.length
  while (start < end && lines[start].trim() === '') start += 1
  while (end > start && lines[end - 1].trim() === '') end -= 1
  return lines.slice(start, end)
}

/** @param {unknown} error */
function errorMessage(error) {
  if (error instanceof Error) return error.message
  return String(error)
}
