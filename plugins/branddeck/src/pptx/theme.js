/**
 * Theme-token extraction: `ppt/presentation.xml` + the presentation-level theme.
 *
 * This module is where the real-file traps from `research/template-report.md` are
 * handled. Read the decision comments before changing anything:
 *
 * 1. **Inverted / duplicated colour slots.** `PPT模板-HW.pptx` has
 *    `dk1 == lt1 == #1D1D1A` and `dk2 == lt2 == #FFFFFF`, so the naive mapping
 *    "background = lt1" paints a black page. We therefore prefer the *painted*
 *    background (`<p:bg><p:bgPr><a:solidFill>` on the primary master) and only fall
 *    back to the scheme slots with a luminance sanity check.
 * 2. **`sysClr` with `lastClr`.** Office writes `<a:sysClr val="windowText"
 *    lastClr="000000"/>`. `lastClr` is the authoritative value (design doc pitfall #3);
 *    `val` is only a hint (`windowText`/`window` → black/white) used when `lastClr`
 *    is absent. The raw form is preserved per slot for diagnostics.
 * 3. **Slide size is not canonical.** `cx=12196763` (13.3385 in), not 12192000.
 *    Always read `<p:sldSz>`.
 * 4. **Six themes, four masters.** Only the theme reached from
 *    `ppt/_rels/presentation.xml.rels` (`rId14` → `theme1.xml`) defines the deck
 *    tokens; themes 2–6 belong to other masters and must never be averaged in.
 * 5. **Empty `<a:ea typeface=""/>`.** Present-but-empty, not absent: the real CJK face
 *    lives in `<a:font script="Hans">` (major 微软雅黑 / minor 黑体), and the master
 *    `<p:txStyles>` overrides both (master1 title = "Microsoft YaHei").
 */
import { asArray } from './read.js'
import { DEFAULT_FONTS, DEFAULT_THEME, makeDefaultTheme } from '../defaults.js'
import { normalizeHex } from '../schema.js'

/** 914400 EMU = 1 inch (OOXML fixed point). */
export const EMU_PER_INCH = 914400

/** The twelve `<a:clrScheme>` slots we care about, in schema order. */
export const COLOR_SLOTS = Object.freeze([
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
])

/**
 * Fallbacks for `<a:sysClr val="…">` when `lastClr` is missing. These are the
 * *Windows* window-colour defaults, i.e. exactly what PowerPoint shows when the file
 * does not pin a value (design doc pitfall #3).
 */
const SYS_COLOR_HINTS = Object.freeze({
  windowText: '#000000',
  window: '#FFFFFF',
  background: '#FFFFFF',
  highlight: '#0078D7',
  highlightText: '#FFFFFF',
  btnFace: '#F0F0F0',
  btnText: '#000000',
  btnShadow: '#A0A0A0',
  grayText: '#6D6D6D',
  menuText: '#000000',
  infoText: '#000000',
  infoBk: '#FFFFE1',
  activeBorder: '#B4B4B4',
  activeCaption: '#99B4D1',
  captionText: '#000000',
  inactiveCaption: '#BFCDDB',
})

/** Coarse `<a:prstClr val="…">` support (rare, but cheap to handle). */
const PRESET_COLORS = Object.freeze({
  black: '#000000',
  white: '#FFFFFF',
  red: '#FF0000',
  green: '#008000',
  blue: '#0000FF',
  yellow: '#FFFF00',
  gray: '#808080',
  grey: '#808080',
  silver: '#C0C0C0',
  orange: '#FFA500',
})

/** Standard `<p:clrMap>` when a master does not declare one. */
const DEFAULT_CLR_MAP = Object.freeze({
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
})

/* ------------------------------------------------------------------ *
 * Colour helpers
 * ------------------------------------------------------------------ */

/**
 * Strictly validate a hex colour body and normalize it to `#RRGGBB`.
 * @param {unknown} value e.g. `'1D1D1A'` / `'#abc'` / `undefined`
 * @returns {string | undefined}
 */
export function hexColor(value) {
  if (typeof value !== 'string') return undefined
  const raw = value.trim()
  if (!/^#?[0-9a-fA-F]{3}([0-9a-fA-F]{3})?$/.test(raw)) {
    // 8-digit ARGB is legal in OOXML but not a value we hand to PowerPoint as a token.
    if (/^#?[0-9a-fA-F]{8}$/.test(raw)) {
      const body = raw.replace('#', '')
      const normalized = `#${body.slice(2)}`.toUpperCase()
      return normalized
    }
    return undefined
  }
  return normalizeHex(raw)
}

/**
 * WCAG relative luminance (0 = black, 1 = white).
 * @param {string} hex `#RRGGBB`
 * @returns {number}
 */
export function luminance(hex) {
  const body = String(hex ?? '').replace('#', '')
  const channels = [0, 2, 4].map((offset) => parseInt(body.slice(offset, offset + 2), 16) / 255)
  const [r, g, b] = channels.map((value) =>
    value <= 0.03928 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4,
  )
  return 0.2126 * r + 0.7152 * g + 0.0722 * b
}

/**
 * WCAG contrast ratio between two colours (1 … 21).
 * @param {string} a
 * @param {string} b
 * @returns {number}
 */
export function contrastRatio(a, b) {
  const first = luminance(a)
  const second = luminance(b)
  const lighter = Math.max(first, second)
  const darker = Math.min(first, second)
  return (lighter + 0.05) / (darker + 0.05)
}

/**
 * Blend `color` into `base` by `weight` (0 → base, 1 → color), per channel.
 * @param {string} base `#RRGGBB`
 * @param {string} color `#RRGGBB`
 * @param {number} weight 0..1
 * @returns {string} `#RRGGBB`
 */
export function mixColor(base, color, weight) {
  const w = Math.min(1, Math.max(0, weight))
  const read = (hex, offset) => parseInt(String(hex).replace('#', '').slice(offset, offset + 2), 16)
  const channel = (offset) => {
    const value = Math.round(read(base, offset) * (1 - w) + read(color, offset) * w)
    return Math.min(255, Math.max(0, value))
  }
  const toHex = (value) => value.toString(16).padStart(2, '0')
  return `#${toHex(channel(0))}${toHex(channel(2))}${toHex(channel(4))}`.toUpperCase()
}

/**
 * `#RRGGBB` → `RRGGBB` (pptxgenjs colour format).
 * @param {string} hex
 * @returns {string}
 */
export function bareHex(hex) {
  return String(hex).replace('#', '').toUpperCase()
}

/* ------------------------------------------------------------------ *
 * Package structure
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} MasterInfo
 * @property {string} part              e.g. `ppt/slideMasters/slideMaster1.xml`
 * @property {string} index             1-based position in `<p:sldMasterIdLst>`
 * @property {string[]} layouts         layout parts owned by this master
 * @property {string} [themePart]
 * @property {Record<string, string>} clrMap
 */

/**
 * @typedef {object} PackageStructure
 * @property {string} presentationPart
 * @property {MasterInfo[]} masters
 * @property {string[]} layouts
 * @property {string[]} slides
 * @property {string | null} themePart  presentation-level theme
 */

/**
 * Walk master/layout/slide/theme relationships once. Order of `masters` follows
 * `<p:sldMasterIdLst>` (that order is what "primary master" means), falling back to
 * relationship order when the list is missing.
 *
 * NOTE on `r:id`: after `removeNSPrefix` both `id="2147483818"` and `r:id="rId1"`
 * collapse onto the key `@_id`; fast-xml-parser keeps the last attribute, i.e. the
 * relationship id — which is the one we need. The numeric master id is not used.
 *
 * @param {import('./read.js').PackageHandle} pkg
 * @returns {PackageStructure}
 */
export function discoverStructure(pkg) {
  const presentationPart = 'ppt/presentation.xml'
  const rels = pkg.relationships(presentationPart)
  const byId = new Map(rels.map((rel) => [rel.id, rel]))
  const presDoc = pkg.xml(presentationPart)
  const pres = presDoc.presentation ?? presDoc

  /** @type {string[]} */
  const masterOrder = []
  for (const node of asArray(pres?.sldMasterIdLst?.sldMasterId)) {
    const relId = node?.['@_id']
    const rel = relId === undefined ? undefined : byId.get(String(relId))
    if (rel?.resolvedPart !== null && rel?.resolvedPart !== undefined && pkg.has(rel.resolvedPart)) {
      masterOrder.push(rel.resolvedPart)
    }
  }
  if (masterOrder.length === 0) {
    for (const rel of rels) {
      if (rel.type.endsWith('/slideMaster') && rel.resolvedPart !== null && pkg.has(rel.resolvedPart)) {
        masterOrder.push(rel.resolvedPart)
      }
    }
  }

  /** @type {MasterInfo[]} */
  const masters = []
  /** @type {string[]} */
  const layouts = []
  masterOrder.forEach((part, index) => {
    const masterRels = pkg.relationships(part)
    const ownLayouts = []
    for (const rel of masterRels) {
      if (rel.type.endsWith('/slideLayout') && rel.resolvedPart !== null && pkg.has(rel.resolvedPart)) {
        ownLayouts.push(rel.resolvedPart)
        layouts.push(rel.resolvedPart)
      }
    }
    const themeRel = masterRels.find((rel) => rel.type.endsWith('/theme'))
    masters.push({
      part,
      index: index + 1,
      layouts: ownLayouts,
      themePart: themeRel?.resolvedPart ?? undefined,
      clrMap: readClrMap(pkg, part),
    })
  })

  /** @type {string[]} */
  const slides = []
  for (const rel of rels) {
    if (rel.type.endsWith('/slide') && rel.resolvedPart !== null && pkg.has(rel.resolvedPart)) {
      slides.push(rel.resolvedPart)
    }
  }

  const themeRel = rels.find((rel) => rel.type.endsWith('/theme'))
  let themePart = themeRel?.resolvedPart ?? null
  if (themePart === null || !pkg.has(themePart)) {
    themePart = pkg.has('ppt/theme/theme1.xml') ? 'ppt/theme/theme1.xml' : null
  }

  return { presentationPart, masters, layouts, slides, themePart }
}

/**
 * Read a master's `<p:clrMap>` (falls back to the standard Office mapping).
 * @param {import('./read.js').PackageHandle} pkg
 * @param {string} masterPart
 * @returns {Record<string, string>}
 */
function readClrMap(pkg, masterPart) {
  const doc = pkg.xml(masterPart)
  const node = doc?.sldMaster?.clrMap
  /** @type {Record<string, string>} */
  const map = { ...DEFAULT_CLR_MAP }
  if (node !== undefined && node !== null) {
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith('@_') || typeof value !== 'string') continue
      map[key] = value
    }
  }
  return map
}

/* ------------------------------------------------------------------ *
 * Theme part parsing
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} RawColorSlot
 * @property {'srgbClr'|'sysClr'|'scrgbClr'|'prstClr'} kind
 * @property {string | null} val        `val` attribute as written (e.g. `windowText`)
 * @property {string | null} lastClr    `lastClr` attribute as written (e.g. `000000`)
 * @property {string} effective         resolved `#RRGGBB`
 */

/**
 * Resolve one `<a:dk1>…</a:dk1>`-style slot to its effective colour, keeping the raw
 * form for diagnostics.
 * @param {any} node
 * @returns {RawColorSlot | null}
 */
export function readColorSlot(node) {
  if (node === undefined || node === null || typeof node !== 'object') return null

  const srgb = node.srgbClr
  if (srgb !== undefined) {
    const val = attrString(srgb?.['@_val'])
    const effective = hexColor(val)
    if (effective !== undefined) return { kind: 'srgbClr', val, lastClr: null, effective }
  }

  const sys = node.sysClr
  if (sys !== undefined) {
    const val = attrString(sys?.['@_val'])
    const lastClr = attrString(sys?.['@_lastClr'])
    const pinned = hexColor(lastClr)
    const hint = val !== null ? SYS_COLOR_HINTS[val] : undefined
    const effective = pinned ?? hint ?? '#000000'
    return { kind: 'sysClr', val, lastClr, effective }
  }

  // Rare alternatives: keep the attribute for diagnostics, resolve what we can.
  for (const kind of /** @type {const} */ (['scrgbClr', 'prstClr'])) {
    const alternative = node[kind]
    if (alternative === undefined) continue
    const val = attrString(alternative?.['@_val'])
    const effective =
      kind === 'prstClr' && val !== null && PRESET_COLORS[val] !== undefined
        ? PRESET_COLORS[val]
        : '#000000'
    return { kind, val, lastClr: null, effective }
  }

  return null
}

/**
 * @param {any} value
 * @returns {string | null}
 */
function attrString(value) {
  if (value === undefined || value === null) return null
  return String(value)
}

/**
 * @typedef {object} FontFaces
 * @property {string | null} latin
 * @property {string | null} ea
 * @property {string | null} cs
 * @property {string | null} hans  `<a:font script="Hans">` face (the real CJK face)
 */

/**
 * Extract the faces of one `<a:majorFont>` / `<a:minorFont>` element.
 * Empty `typeface=""` is normalized to `null` (trap: present-but-empty).
 * @param {any} fontNode
 * @returns {FontFaces}
 */
export function readFontFaces(fontNode) {
  const face = (node) => {
    const value = node?.['@_typeface']
    if (value === undefined || value === null) return null
    const trimmed = String(value).trim()
    return trimmed.length === 0 ? null : trimmed
  }
  /** @type {Record<string, string | null>} */
  const byScript = {}
  for (const entry of asArray(fontNode?.font)) {
    const script = entry?.['@_script']
    if (script === undefined) continue
    byScript[String(script)] = face(entry)
  }
  return {
    latin: face(fontNode?.latin),
    ea: face(fontNode?.ea),
    cs: face(fontNode?.cs),
    hans: byScript.Hans ?? byScript.hans ?? null,
  }
}

/**
 * @typedef {object} TxStyleFaces
 * @property {string | null} latin
 * @property {string | null} ea
 * @property {string | null} cs
 */

/**
 * Read `a:defRPr` font faces of `titleStyle` / `bodyStyle` level 1 on a master.
 * @param {import('./read.js').PackageHandle} pkg
 * @param {string} masterPart
 * @returns {{ title: TxStyleFaces, body: TxStyleFaces }}
 */
export function readTxStyleFaces(pkg, masterPart) {
  const doc = pkg.xml(masterPart)
  const txStyles = doc?.sldMaster?.txStyles

  const defRPr = (styleKey) => {
    const levelOne = asArray(txStyles?.[styleKey]?.lvl1pPr)[0]
    const level = levelOne !== undefined ? levelOne : asArray(txStyles?.[styleKey]?.lvl2pPr)[0]
    return level?.defRPr ?? null
  }
  const faces = (node) => ({
    latin: attrString(node?.latin?.['@_typeface'])?.trim() || null,
    ea: attrString(node?.ea?.['@_typeface'])?.trim() || null,
    cs: attrString(node?.cs?.['@_typeface'])?.trim() || null,
  })
  return { title: faces(defRPr('titleStyle')), body: faces(defRPr('bodyStyle')) }
}

/**
 * Resolve a `+mj-lt` / `+mn-ea` style reference against the theme font scheme.
 * Literal typefaces pass through untouched.
 * @param {string | null} typeface
 * @param {{ major: FontFaces, minor: FontFaces }} faces
 * @returns {string | null} usable typeface or null
 */
export function resolveTypefaceRef(typeface, faces) {
  if (typeface === null || typeface === undefined) return null
  const value = String(typeface).trim()
  if (value.length === 0) return null
  if (!value.startsWith('+')) return value
  const key = value.slice(1).toLowerCase()
  switch (key) {
    case 'mj-lt':
      return faces.major.latin
    case 'mj-ea':
      return faces.major.ea ?? faces.major.hans
    case 'mj-cs':
      return faces.major.cs
    case 'mn-lt':
      return faces.minor.latin
    case 'mn-ea':
      return faces.minor.ea ?? faces.minor.hans
    case 'mn-cs':
      return faces.minor.cs
    default:
      return null
  }
}

/* ------------------------------------------------------------------ *
 * Background / surface mining
 * ------------------------------------------------------------------ */

/**
 * Read the painted background of a master/layout part.
 * Supports `<p:bg><p:bgPr><a:solidFill><a:srgbClr val/>` plus `<a:schemeClr val="bg1">`.
 * @param {import('./read.js').PackageHandle} pkg
 * @param {string} part
 * @param {Record<string, RawColorSlot>} slots
 * @param {Record<string, string>} clrMap
 * @returns {{ hex: string, kind: 'srgbClr'|'schemeClr' } | null}
 */
export function readPaintedBackground(pkg, part, slots, clrMap) {
  if (!pkg.has(part)) return null
  const doc = pkg.xml(part)
  const root = doc?.sldMaster ?? doc?.sldLayout ?? doc?.sld
  const solidFill = root?.cSld?.bg?.bgPr?.solidFill ?? root?.bg?.bgPr?.solidFill
  if (solidFill === undefined || solidFill === null) return null
  const srgb = solidFill.srgbClr?.['@_val']
  const srgbHex = hexColor(srgb)
  if (srgbHex !== undefined) return { hex: srgbHex, kind: 'srgbClr' }
  const scheme = solidFill.schemeClr?.['@_val']
  if (scheme !== undefined && scheme !== null) {
    const slotName = clrMap[String(scheme)] ?? String(scheme)
    const slot = slots[slotName]
    if (slot !== undefined) return { hex: slot.effective, kind: 'schemeClr' }
  }
  return null
}

/* ------------------------------------------------------------------ *
 * Public API
 * ------------------------------------------------------------------ */

/**
 * @typedef {object} ThemeExtraction
 * @property {import('../schema.js').ThemeTokens} tokens
 * @property {object} raw  diagnostics payload: clrScheme, fontScheme, slideSize, notes
 */

/**
 * Extract BrandDeck design tokens from an opened package.
 *
 * Pure: no disk access, no store access. `tokens.brandAssets` contains only the
 * `icons` map; the logo path/size/frame is filled in by `template.js` once the caller
 * has decided where the logo file will be persisted.
 *
 * @param {import('./read.js').PackageHandle} pkg
 * @param {{templateId?: string, templateName?: string}} [identity]
 * @returns {ThemeExtraction}
 */
export function extractThemeTokens(pkg, identity = {}) {
  const templateId = nonEmptyString(identity.templateId) ?? DEFAULT_THEME.templateId
  const templateName = nonEmptyString(identity.templateName) ?? DEFAULT_THEME.templateName
  /** @type {string[]} */
  const notes = []

  const structure = discoverStructure(pkg)
  const themePart = structure.themePart
  const themeDoc = themePart === null ? null : pkg.xml(themePart)
  const themeRoot = themeDoc?.theme ?? themeDoc?.themeOverride ?? null
  const themeElements = themeRoot?.themeElements ?? null

  /* --- clrScheme ------------------------------------------------------- */
  const clrSchemeNode = themeElements?.clrScheme ?? null
  /** @type {Record<string, RawColorSlot>} */
  const slots = {}
  if (clrSchemeNode === null) {
    notes.push('theme has no <a:clrScheme>; falling back to DEFAULT_THEME colours')
  } else {
    for (const slotName of COLOR_SLOTS) {
      const parsed = readColorSlot(clrSchemeNode[slotName])
      if (parsed !== null) {
        slots[slotName] = parsed
        if (parsed.kind === 'sysClr') {
          notes.push(
            `<a:sysClr> in slot ${slotName}: val="${parsed.val}" lastClr="${parsed.lastClr}" ` +
              `→ ${parsed.effective}` +
              (parsed.lastClr === null ? ' (no lastClr; used the system-colour hint)' : ''),
          )
        }
      } else if (clrSchemeNode[slotName] !== undefined) {
        notes.push(`colour slot ${slotName} uses an unsupported colour form and was ignored`)
      }
    }
  }
  const slotHex = (name, fallback) => slots[name]?.effective ?? fallback

  /* --- slide size ------------------------------------------------------ */
  const presDoc = pkg.xml('ppt/presentation.xml')
  const sldSz = presDoc?.presentation?.sldSz ?? presDoc?.sldSz ?? null
  const cx = toFiniteNumber(sldSz?.['@_cx'])
  const cy = toFiniteNumber(sldSz?.['@_cy'])
  const slideSize = {
    cx: cx ?? 12192000,
    cy: cy ?? 6858000,
    widthIn: round4((cx ?? 12192000) / EMU_PER_INCH),
    heightIn: round4((cy ?? 6858000) / EMU_PER_INCH),
  }
  if (cx !== undefined && cx !== 12192000) {
    notes.push(
      `<p:sldSz cx=${cx} is not the canonical 12192000 (${slideSize.widthIn} in wide); ` +
        'the 1280x720 canvas is mapped to LAYOUT_WIDE at px/96 regardless',
    )
  }
  if (cx === undefined) notes.push('presentation.xml has no <p:sldSz>; assumed 16:9 13.333x7.5in')

  /* --- background ------------------------------------------------------ */
  const primaryMaster = structure.masters[0] ?? null
  const clrMap = primaryMaster?.clrMap ?? { ...DEFAULT_CLR_MAP }

  const painterOrder = [
    ...structure.masters.map((master) => master.part),
    ...structure.layouts,
  ]
  const paintedBackgrounds = []
  for (const part of painterOrder) {
    const painted = readPaintedBackground(pkg, part, slots, clrMap)
    if (painted !== null) paintedBackgrounds.push({ part, ...painted })
  }

  let background
  let backgroundSource
  const masterBg = primaryMaster === null ? null : readPaintedBackground(pkg, primaryMaster.part, slots, clrMap)
  if (masterBg !== null) {
    background = masterBg.hex
    backgroundSource = `painted background of the primary master (${primaryMaster.part})`
  } else {
    // Trap: lt1 may be the *dark* slot (dk1 == lt1 here). Only trust it when it is
    // actually light; otherwise take the lightest of the remaining light-slot candidates.
    const lt1 = slotHex('lt1', DEFAULT_THEME.colors.background)
    const candidates = [lt1, slotHex('lt2', '#FFFFFF'), slotHex('dk2', '#FFFFFF')]
    if (luminance(lt1) >= 0.6) {
      background = lt1
      backgroundSource = 'theme lt1 (no painted background on any master)'
    } else {
      background = candidates.reduce((best, value) => (luminance(value) > luminance(best) ? value : best), candidates[0])
      backgroundSource = 'lightest of lt1/lt2/dk2 (lt1 is dark → inverted theme)'
      notes.push(
        `lt1 resolves to ${lt1} (dark) while dk1 is ${slotHex('dk1', '#000000')}: the light/dark ` +
          'slots are inverted in this theme, so the background came from the lightest light-slot ' +
          `candidate (${background})`,
      )
    }
  }
  if (backgroundSource === undefined) backgroundSource = 'DEFAULT_THEME'

  const darkTheme = luminance(background) < 0.35

  /* --- textPrimary ----------------------------------------------------- */
  const dk1 = slotHex('dk1', DEFAULT_THEME.colors.textPrimary)
  const dk2 = slotHex('dk2', DEFAULT_THEME.colors.textSecondary)
  let textPrimary
  if (darkTheme) {
    // Dark page: the "dark" slots are unusable; take whichever light slot reads best.
    const lightCandidates = [dk2, slotHex('lt1', '#FFFFFF'), slotHex('lt2', '#FFFFFF')]
    textPrimary = lightCandidates.reduce(
      (best, value) => (contrastRatio(background, value) > contrastRatio(background, best) ? value : best),
      lightCandidates[0],
    )
  } else if (contrastRatio(background, dk1) >= 4.5) {
    textPrimary = dk1
  } else if (contrastRatio(background, dk2) >= 4.5) {
    textPrimary = dk2
    notes.push(`dk1 ${dk1} has poor contrast on ${background}; textPrimary came from dk2`)
  } else {
    textPrimary = darkTheme ? '#FFFFFF' : '#000000'
    notes.push('no colour slot reached a 4.5:1 contrast on the background; textPrimary forced')
  }

  /* --- textSecondary --------------------------------------------------- */
  let textSecondary
  let textSecondarySource
  if (!darkTheme && contrastRatio(background, dk2) >= 4.5 && dk2 !== textPrimary) {
    textSecondary = dk2
    textSecondarySource = 'theme dk2'
  } else {
    // dk2 is unusable (here: #FFFFFF on a white page). Fall back to the accent slot
    // that behaves like a secondary text grey: enough contrast to read, not so dark
    // that it duplicates textPrimary. accent6 (#666666 here) is the canonical candidate.
    const ranked = ['accent6', 'accent5', 'accent4', 'accent3', 'hlink', 'accent2', 'accent1', 'folHlink']
      .map((name) => ({ name, hex: slots[name]?.effective }))
      .filter((entry) => entry.hex !== undefined)
    const usable = ranked.find((entry) => {
      const ratio = contrastRatio(background, entry.hex)
      return ratio >= 3 && ratio <= 10 && entry.hex !== textPrimary
    })
    if (usable !== undefined) {
      textSecondary = usable.hex
      textSecondarySource = `theme ${usable.name} (dk2 ${dk2} has no contrast on ${background})`
      notes.push(
        `dk2 resolves to ${dk2}, which is unusable as secondary text on ${background}; ` +
          `textSecondary came from ${usable.name} (${usable.hex})`,
      )
    } else {
      textSecondary = mixColor(background, textPrimary, 0.55)
      textSecondarySource = 'derived: 55% textPrimary over the background'
      notes.push('no accent slot was a plausible secondary text colour; derived a mid tone')
    }
  }

  /* --- surface --------------------------------------------------------- */
  let surface
  let surfaceSource
  const lt2 = slotHex('lt2', DEFAULT_THEME.colors.surface)
  if (luminance(lt2) >= 0.7 && lt2 !== background) {
    surface = lt2
    surfaceSource = 'theme lt2'
  } else {
    // lt2 is white here (identical to the page). The template still *paints* a surface
    // somewhere — look for the most common non-background solidFill among
    // masters/layouts (slideLayout5 = #EBEBEB in this file) before deriving a tint.
    /** @type {Map<string, number>} */
    const counts = new Map()
    for (const entry of paintedBackgrounds) {
      if (entry.hex === background) continue
      if (luminance(entry.hex) < 0.6) continue
      counts.set(entry.hex, (counts.get(entry.hex) ?? 0) + 1)
    }
    const best = [...counts.entries()].sort((a, b) => b[1] - a[1] || luminance(b[0]) - luminance(a[0]))[0]
    if (best !== undefined) {
      surface = best[0]
      surfaceSource = `most common non-page solidFill among masters/layouts (${best[1]} occurrence)`
      notes.push(`lt2 ${lt2} equals the page background; surface came from ${surface} (${surfaceSource})`)
    } else {
      // Documented derivation: a 9% tint of the text colour over the page —
      // #FFFFF F + 9% of #1D1D1A = #EBEBEB, the same value the template paints.
      surface = mixColor(background, textPrimary, 0.09)
      surfaceSource = 'derived: 9% textPrimary tinted into the background'
      notes.push(`lt2 ${lt2} equals the page background and no layout paints a surface; surface derived`)
    }
  }

  /* --- brand accents --------------------------------------------------- */
  const primary = slotHex('accent1', DEFAULT_THEME.colors.primary)
  const accent = slotHex('accent2', DEFAULT_THEME.colors.accent)
  const chartPalette = []
  for (let index = 1; index <= 6; index += 1) {
    const hex = slots[`accent${index}`]?.effective
    if (hex !== undefined) chartPalette.push(hex)
  }
  if (chartPalette.length === 0) {
    chartPalette.push(...DEFAULT_THEME.colors.chartPalette)
    notes.push('no accent slots found; DEFAULT_THEME chart palette used')
  }

  /* --- fonts ----------------------------------------------------------- */
  const majorFaces = readFontFaces(themeElements?.fontScheme?.majorFont)
  const minorFaces = readFontFaces(themeElements?.fontScheme?.minorFont)
  const faces = { major: majorFaces, minor: minorFaces }
  const txStyles = primaryMaster === null ? null : readTxStyleFaces(pkg, primaryMaster.part)

  /**
   * Title face order (CJK-first, because a deck title is the brand display face):
   *   master txStyles ea → master txStyles latin → theme major ea → theme major Hans
   *   → theme major latin → DEFAULT_THEME.
   */
  const title =
    firstNonEmpty(
      resolveTypefaceRef(txStyles?.title.ea, faces),
      resolveTypefaceRef(txStyles?.title.latin, faces),
      majorFaces.ea,
      majorFaces.hans,
      majorFaces.latin,
    ) ?? DEFAULT_FONTS.title

  /**
   * Body face order (Latin-first) — deliberate decision, see
   * `research/template-report.md` §5.3: "the" CJK body face is genuinely ambiguous in
   * this template (master1 bodyStyle ea = Microsoft YaHei, Hans minor = 黑体, content
   * masters inherit Calibri). expected-tokens.json pins the effective *Latin* body face,
   * which is also what PowerPoint uses for digits/latin runs and what the compiler's
   * `fontFace` needs; the CJK body face is still reported in
   * `raw.fontScheme.effective.bodyEa` for the panel/diagnostics.
   */
  const body =
    firstNonEmpty(
      resolveTypefaceRef(txStyles?.body.latin, faces),
      minorFaces.latin,
      resolveTypefaceRef(txStyles?.body.ea, faces),
      minorFaces.ea,
      minorFaces.hans,
      majorFaces.latin,
    ) ?? DEFAULT_FONTS.body

  const latin = firstNonEmpty(minorFaces.latin, majorFaces.latin) ?? DEFAULT_FONTS.latin
  const fontCs = firstNonEmpty(minorFaces.cs, resolveTypefaceRef(txStyles?.body.cs, faces), minorFaces.latin)

  const effectiveBodyEa =
    firstNonEmpty(resolveTypefaceRef(txStyles?.body.ea, faces), minorFaces.ea, minorFaces.hans) ?? null
  const effectiveTitleEa =
    firstNonEmpty(resolveTypefaceRef(txStyles?.title.ea, faces), majorFaces.ea, majorFaces.hans) ?? null

  if (themeElements?.fontScheme === undefined) {
    notes.push('theme has no <a:fontScheme>; DEFAULT_THEME fonts used')
  } else if (majorFaces.ea === null && minorFaces.ea === null) {
    notes.push(
      '<a:ea typeface=""> is present but empty in both font sets: the real CJK faces live ' +
        'in <a:font script="Hans"> and in the master <p:txStyles>',
    )
  }

  /* --- assemble -------------------------------------------------------- */
  const base = makeDefaultTheme({ templateId, templateName })
  /** @type {import('../schema.js').ThemeTokens} */
  const tokens = {
    templateId,
    templateName,
    colors: {
      background,
      surface,
      textPrimary,
      textSecondary,
      primary,
      accent,
      chartPalette,
    },
    fonts: { title, body, latin },
    brandAssets: {
      // Filled in by template.js once the logo file location is known.
      icons: { ...base.brandAssets.icons },
    },
  }

  const raw = {
    themePart,
    primaryMasterPart: primaryMaster?.part ?? null,
    masters: structure.masters.map((master) => ({
      part: master.part,
      index: master.index,
      themePart: master.themePart ?? null,
      layouts: master.layouts,
    })),
    layouts: structure.layouts,
    slides: structure.slides,
    clrScheme: {
      name: attrString(clrSchemeNode?.['@_name']),
      slots,
    },
    fontScheme: {
      name: attrString(themeElements?.fontScheme?.['@_name']),
      major: majorFaces,
      minor: minorFaces,
      txStyles,
      effective: {
        title,
        titleEa: effectiveTitleEa,
        body,
        bodyEa: effectiveBodyEa,
        latin,
        cs: fontCs ?? null,
      },
    },
    slideSize,
    sources: { background: backgroundSource, surface: surfaceSource, textSecondary: textSecondarySource },
    notes,
  }

  return { tokens, raw }
}

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

/** @param {unknown} value */
function nonEmptyString(value) {
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : undefined
}

/** @param {unknown} value */
function toFiniteNumber(value) {
  if (value === undefined || value === null) return undefined
  const number = Number(value)
  return Number.isFinite(number) ? number : undefined
}

/** @param {number} value */
function round4(value) {
  return Math.round(value * 10000) / 10000
}

/**
 * First non-empty string of the argument list.
 * @param {...(string | null | undefined)} values
 * @returns {string | null}
 */
function firstNonEmpty(...values) {
  for (const value of values) {
    if (typeof value === 'string' && value.trim().length > 0) return value.trim()
  }
  return null
}
