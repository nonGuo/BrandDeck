/**
 * theme.js: clrScheme / fontScheme / sldSz extraction, including every trap from
 * `research/template-report.md` §3/§5/§9.
 *
 * The real HW template is parsed and cross-checked field-by-field against
 * `research/expected-tokens.json` (printed as a diff table).
 *
 * Run: node test/engine-theme.mjs
 */
import { existsSync, readFileSync } from 'node:fs'
import JSZip from 'jszip'
import {
  EXPECTED_TOKENS_PATH,
  TEMPLATE_PATH,
  assert,
  assertEqual,
  assertMatch,
  diffTable,
  run,
  test,
} from './engine-helpers.mjs'
import { loadPackage } from '../src/pptx/read.js'
import {
  COLOR_SLOTS,
  contrastRatio,
  discoverStructure,
  extractThemeTokens,
  luminance,
  mixColor,
} from '../src/pptx/theme.js'
import { ThemeTokensSchema } from '../src/schema.js'
import { DEFAULT_THEME } from '../src/defaults.js'

const HEX6 = /^#[0-9A-F]{6}$/

/**
 * Build a synthetic package whose theme/master/presentation are controlled by the
 * caller, so the `sysClr`/inverted-slot paths can be exercised deterministically.
 *
 * @param {object} input
 * @param {string} input.clrScheme inner XML of `<a:clrScheme>`
 * @param {string} [input.fontScheme] inner XML of `<a:fontScheme>`
 * @param {string} [input.masterBg] `<p:bg>` XML for the master (default: none)
 * @param {string} [input.masterTxStyles] `<p:txStyles>` XML for the master
 * @param {string} [input.sldSz] the `<p:sldSz>` element
 * @returns {Promise<Buffer>}
 */
async function syntheticTemplate(input) {
  const zip = new JSZip()
  const sldSz = input.sldSz ?? '<p:sldSz cx="12192000" cy="6858000"/>'
  zip.file(
    'ppt/presentation.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:presentation xmlns:a="a" xmlns:p="p" xmlns:r="r">' +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      sldSz +
      '</p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sldMaster xmlns:a="a" xmlns:p="p" xmlns:r="r"><p:cSld>' +
      (input.masterBg ?? '') +
      '<p:spTree/></p:cSld>' +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
      'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      (input.masterTxStyles ?? '') +
      '</p:sldMaster>',
  )
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/theme/theme1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Synthetic">' +
      '<a:themeElements>' +
      `<a:clrScheme name="Synthetic">${input.clrScheme}</a:clrScheme>` +
      (input.fontScheme === undefined
        ? ''
        : `<a:fontScheme name="Synthetic">${input.fontScheme}</a:fontScheme>`) +
      '</a:themeElements></a:theme>',
  )
  zip.file('ppt/slides/slide1.xml', '<?xml version="1.0"?><p:sld xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:sld>')
  return zip.generateAsync({ type: 'nodebuffer' })
}

const SYS_CLR_SCHEME =
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'

await test('extractThemeTokens: every colour is #RRGGBB and the schema accepts the tokens', async () => {
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const { tokens } = extractThemeTokens(pkg, { templateId: 'hw', templateName: 'HW' })
  const colors = [
    tokens.colors.background,
    tokens.colors.surface,
    tokens.colors.textPrimary,
    tokens.colors.textSecondary,
    tokens.colors.primary,
    tokens.colors.accent,
    ...tokens.colors.chartPalette,
  ]
  for (const color of colors) assertMatch(color, HEX6, `colour ${color} is #RRGGBB`)
  const parsed = ThemeTokensSchema.safeParse(tokens)
  assert(parsed.success, `ThemeTokensSchema accepts the extracted tokens: ${JSON.stringify(parsed.error?.issues)}`)
  const fallback = ThemeTokensSchema.safeParse(DEFAULT_THEME)
  assert(fallback.success, 'DEFAULT_THEME satisfies ThemeTokensSchema')
  return `${colors.length} colours validated`
})

await test('extractThemeTokens: slide size is read, not hardcoded (cx=12196763)', async () => {
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const { raw } = extractThemeTokens(pkg)
  assertEqual(raw.slideSize.cx, 12196763, 'cx from <p:sldSz>')
  assertEqual(raw.slideSize.cy, 6858000, 'cy from <p:sldSz>')
  assertEqual(raw.slideSize.widthIn, 13.3385, 'widthIn = cx / 914400')
  assertEqual(raw.slideSize.heightIn, 7.5, 'heightIn = cy / 914400')
  assert(raw.themePart === 'ppt/theme/theme1.xml', `presentation-level theme is theme1, got ${raw.themePart}`)
  assertEqual(raw.masters.length, 4, 'four masters discovered')
  assertEqual(raw.layouts.length, 7, 'seven layouts discovered')
  assertEqual(raw.slides.length, 5, 'five slides discovered')
  const structure = discoverStructure(pkg)
  assertEqual(structure.masters[0].part, 'ppt/slideMasters/slideMaster1.xml', 'primary master is slideMaster1')
  assertEqual(structure.masters[3].part, 'ppt/slideMasters/slideMaster4.xml', 'master order follows sldMasterIdLst')
  return `${raw.slideSize.widthIn}x${raw.slideSize.heightIn} in`
})

await test('extractThemeTokens: inverted slots (dk1==lt1, dk2==lt2) are handled', async () => {
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const { tokens, raw } = extractThemeTokens(pkg)
  // theme1 really does duplicate the light/dark pairs — assert the trap exists...
  assertEqual(raw.clrScheme.slots.dk1.effective, '#1D1D1A', 'dk1 is 1D1D1A')
  assertEqual(raw.clrScheme.slots.lt1.effective, '#1D1D1A', 'lt1 equals dk1 (inverted theme)')
  assertEqual(raw.clrScheme.slots.dk2.effective, '#FFFFFF', 'dk2 is white')
  assertEqual(raw.clrScheme.slots.lt2.effective, '#FFFFFF', 'lt2 equals dk2')
  // ...and that the parser still produces a light page with dark text.
  assertEqual(tokens.colors.background, '#FFFFFF', 'background is white despite lt1 being dark')
  assert(luminance(tokens.colors.background) > 0.9, 'background is a light page')
  assertEqual(tokens.colors.textPrimary, '#1D1D1A', 'textPrimary came from dk1')
  assert(contrastRatio(tokens.colors.background, tokens.colors.textPrimary) > 10, 'text/background contrast is high')
  assertEqual(tokens.colors.textSecondary, '#666666', 'textSecondary (dk2 is unusable) fell back to accent6')
  assert(
    raw.sources.background.includes('primary master'),
    `background source is documented: ${raw.sources.background}`,
  )
  return raw.sources
})

await test('extractThemeTokens: cross-check against research/expected-tokens.json', async () => {
  if (!existsSync(EXPECTED_TOKENS_PATH)) {
    console.log(`  NOTE ${EXPECTED_TOKENS_PATH} is absent — engine dump is authoritative`)
    return 'expected-tokens.json absent, skipped'
  }
  const expected = JSON.parse(readFileSync(EXPECTED_TOKENS_PATH, 'utf8'))
  const pkg = await loadPackage({ path: TEMPLATE_PATH })
  const { tokens, raw } = extractThemeTokens(pkg, { templateId: 'hw', templateName: 'HW' })

  /** @type {{field: string, expected: unknown, actual: unknown}[]} */
  const rows = [
    { field: 'colors.background', expected: expected.colors.background, actual: tokens.colors.background },
    { field: 'colors.surface', expected: expected.colors.surface, actual: tokens.colors.surface },
    { field: 'colors.textPrimary', expected: expected.colors.textPrimary, actual: tokens.colors.textPrimary },
    { field: 'colors.textSecondary', expected: expected.colors.textSecondary, actual: tokens.colors.textSecondary },
    { field: 'colors.primary', expected: expected.colors.primary, actual: tokens.colors.primary },
    { field: 'colors.accent', expected: expected.colors.accent, actual: tokens.colors.accent },
    { field: 'colors.chartPalette', expected: expected.colors.chartPalette, actual: tokens.colors.chartPalette },
    { field: 'fonts.title', expected: expected.fonts.title, actual: tokens.fonts.title },
    { field: 'fonts.body', expected: expected.fonts.body, actual: tokens.fonts.body },
    { field: 'fonts.latin', expected: expected.fonts.latin, actual: tokens.fonts.latin },
    { field: 'slideSize.cx', expected: expected.slideSize.cx, actual: raw.slideSize.cx },
    { field: 'slideSize.cy', expected: expected.slideSize.cy, actual: raw.slideSize.cy },
    { field: 'slideSize.widthIn', expected: expected.slideSize.widthIn, actual: raw.slideSize.widthIn },
    { field: 'slideSize.heightIn', expected: expected.slideSize.heightIn, actual: raw.slideSize.heightIn },
    { field: 'masters', expected: expected.masters, actual: raw.masters.length },
    { field: 'layouts', expected: expected.layouts, actual: raw.layouts.length },
  ]
  for (const [slot, rawExpected] of Object.entries(expected.colorsRaw ?? {})) {
    rows.push({
      field: `raw.${slot}.effective`,
      expected: String(rawExpected.effective).toUpperCase(),
      actual: raw.clrScheme.slots[slot]?.effective,
    })
  }

  const mismatches = diffTable(rows)
  if (mismatches.length > 0) {
    throw new Error(
      `parser disagrees with research/expected-tokens.json on: ${mismatches.map((row) => row.field).join(', ')}`,
    )
  }
  assertEqual(COLOR_SLOTS.length, 12, 'twelve colour slots are tracked')
  return `${rows.length} fields match`
})

await test('extractThemeTokens: sysClr/lastClr path (theme5-style) resolves from lastClr', async () => {
  const bytes = await syntheticTemplate({
    clrScheme: SYS_CLR_SCHEME,
    fontScheme:
      '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/>' +
      '<a:font script="Hans" typeface="微软雅黑"/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>' +
      '<a:font script="Hans" typeface="黑体"/></a:minorFont>',
  })
  const pkg = await loadPackage({ bytes })
  const { tokens, raw } = extractThemeTokens(pkg, { templateId: 'synthetic', templateName: 'Synthetic' })

  assertEqual(raw.clrScheme.slots.dk1.kind, 'sysClr', 'dk1 is a <a:sysClr>')
  assertEqual(raw.clrScheme.slots.dk1.val, 'windowText', 'sysClr @val is preserved')
  assertEqual(raw.clrScheme.slots.dk1.lastClr, '000000', 'sysClr @lastClr is preserved')
  assertEqual(raw.clrScheme.slots.dk1.effective, '#000000', 'effective colour comes from lastClr')
  assertEqual(raw.clrScheme.slots.lt1.effective, '#FFFFFF', 'lt1 comes from lastClr="FFFFFF"')
  assertEqual(tokens.colors.background, '#FFFFFF', 'background = lt1 (light)')
  assertEqual(tokens.colors.textPrimary, '#000000', 'textPrimary = dk1')
  assertEqual(tokens.colors.surface, '#E7E6E6', 'surface = lt2 (light and distinct)')
  assertEqual(tokens.colors.primary, '#4472C4', 'primary = accent1')
  assertEqual(tokens.colors.accent, '#ED7D31', 'accent = accent2')
  assertEqual(tokens.colors.chartPalette.length, 6, 'palette has six accents')
  assert(
    raw.notes.some((note) => note.includes('sysClr')),
    'diagnostics mention the sysClr slots',
  )
  const parsed = ThemeTokensSchema.safeParse(tokens)
  assert(parsed.success, 'sysClr-derived tokens satisfy ThemeTokensSchema')
  return raw.clrScheme.slots.dk1
})

await test('extractThemeTokens: sysClr without lastClr uses the window-colour hint', async () => {
  const clrScheme =
    '<a:dk1><a:sysClr val="windowText"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="333333"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="C7000A"/></a:accent1><a:accent2><a:srgbClr val="E9002F"/></a:accent2>'
  const bytes = await syntheticTemplate({ clrScheme })
  const pkg = await loadPackage({ bytes })
  const { tokens, raw } = extractThemeTokens(pkg)
  assertEqual(raw.clrScheme.slots.dk1.kind, 'sysClr', 'dk1 is sysClr')
  assertEqual(raw.clrScheme.slots.dk1.lastClr, null, 'lastClr absent')
  assertEqual(raw.clrScheme.slots.dk1.effective, '#000000', 'windowText hint → black')
  assertEqual(raw.clrScheme.slots.lt1.effective, '#FFFFFF', 'window hint → white')
  assertEqual(tokens.colors.background, '#FFFFFF', 'background from the window hint')
  return raw.clrScheme.slots.lt1
})

await test('extractThemeTokens: missing fontScheme / missing \u003ca:ea\u003e falls back safely', async () => {
  const bytes = await syntheticTemplate({
    clrScheme: SYS_CLR_SCHEME,
    masterTxStyles:
      '<p:txStyles><titleStyle><p:lvl1pPr><p:defRPr><a:latin typeface="Microsoft YaHei"/></p:defRPr></p:lvl1pPr></titleStyle>' +
      '<p:bodyStyle><p:lvl1pPr><p:defRPr><a:latin typeface="Arial"/></p:defRPr></p:lvl1pPr></p:bodyStyle></p:txStyles>',
  })
  const pkg = await loadPackage({ bytes })
  const { tokens, raw } = extractThemeTokens(pkg)
  assertEqual(tokens.fonts.title, 'Microsoft YaHei', 'title comes from master titleStyle (latin)')
  assertEqual(tokens.fonts.body, 'Arial', 'body comes from master bodyStyle (latin)')
  assertEqual(tokens.fonts.latin, DEFAULT_THEME.fonts.latin, 'latin falls back to DEFAULT_THEME when there is no theme fontScheme')
  assert(
    raw.notes.some((note) => note.includes('fontScheme')),
    'diagnostics report the missing fontScheme',
  )

  // A theme with fontScheme but empty <a:ea> must not yield empty font names.
  const withEmptyEa = await syntheticTemplate({
    clrScheme: SYS_CLR_SCHEME,
    fontScheme:
      '<a:majorFont><a:latin typeface="Arial Black"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>',
  })
  const pkg2 = await loadPackage({ bytes: withEmptyEa })
  const second = extractThemeTokens(pkg2)
  assert(second.tokens.fonts.title.length > 0, 'title is non-empty')
  assert(second.tokens.fonts.body.length > 0, 'body is non-empty')
  assertEqual(second.tokens.fonts.title, 'Arial Black', 'empty <a:ea> falls through to majorFont latin')
  assertEqual(second.tokens.fonts.body, 'Arial', 'empty <a:ea> falls through to minorFont latin')
  assert(
    second.raw.notes.some((note) => note.includes('typeface=""')),
    'diagnostics explain the empty <a:ea>',
  )
  return second.tokens.fonts
})

await test('mixColor/luminance/contrast helpers are sane', () => {
  assertEqual(mixColor('#FFFFFF', '#000000', 0), '#FFFFFF', 'weight 0 keeps the base')
  assertEqual(mixColor('#FFFFFF', '#000000', 1), '#000000', 'weight 1 applies the colour')
  assertEqual(luminance('#FFFFFF'), 1, 'white luminance is 1')
  assertEqual(luminance('#000000'), 0, 'black luminance is 0')
  assert(Math.abs(contrastRatio('#FFFFFF', '#000000') - 21) < 1e-9, 'black/white contrast is 21')
  return { mix: mixColor('#FFFFFF', '#1D1D1A', 0.09), contrast: contrastRatio('#FFFFFF', '#666666') }
})

run('engine-theme')
