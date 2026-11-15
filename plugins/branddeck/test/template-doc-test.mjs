/**
 * template-doc.js: layouts inventory (`layouts.json`) + template spec (`template.md`).
 *
 * Plain Node ESM, no test framework and no mocks: every check runs against the real
 * corporate template `PPT模板-HW.pptx` (the file the product ships against). Each test
 * prints one PASS/FAIL line; the process exits with code 1 when anything fails.
 *
 * Run: node test/template-doc-test.mjs
 */
import assert from 'node:assert/strict'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { loadPackage } from '../src/pptx/read.js'
import { discoverStructure, extractThemeTokens } from '../src/pptx/theme.js'
import { listMedia } from '../src/pptx/media.js'
import { buildLayoutsInventory, buildTemplateDoc, buildTemplateMarkdown } from '../src/template-doc.js'

/** `plugins/branddeck/test/` → workspace root, where the real template lives. */
const WORKSPACE_DIR = fileURLToPath(new URL('../../../', import.meta.url))
const TEMPLATE_PATH = join(WORKSPACE_DIR, 'PPT模板-HW.pptx')

/** Strings that must never reach the generated markdown. */
const FORBIDDEN = /undefined|NaN|\[object Object\]/

const results = []

/**
 * @param {string} name
 * @param {() => any | Promise<any>} fn
 */
async function test(name, fn) {
  try {
    const detail = await fn()
    results.push({ name, ok: true })
    console.log(`PASS ${name}${detail === undefined ? '' : ` — ${detail}`}`)
  } catch (error) {
    results.push({ name, ok: false })
    console.log(`FAIL ${name}`)
    console.log(
      String(error?.stack ?? error)
        .split('\n')
        .map((line) => `     ${line}`)
        .join('\n'),
    )
  }
}

/* ------------------------------------------------------------------ *
 * Shared real inputs (loaded once, exactly like the Host does)
 * ------------------------------------------------------------------ */

const pkg = await loadPackage({ path: TEMPLATE_PATH })
const structure = discoverStructure(pkg)
const { tokens } = extractThemeTokens(pkg, {
  templateId: 'template-doc-test',
  templateName: 'PPT模板-HW',
})

/** JSON-safe media inventory, mirroring what `parsePptxTemplate` persists. */
const media = listMedia(pkg).map((entry) => ({
  path: entry.path,
  bytes: entry.bytes,
  contentType: entry.contentType,
  format: entry.format,
  widthPx: entry.widthPx,
  heightPx: entry.heightPx,
  referencedBy: entry.referencedBy,
}))

/** Deliberately minimal meta: `buildTemplateDoc` must back-fill counts itself. */
const minimalMeta = {
  templateId: 'template-doc-test',
  templateName: 'PPT模板-HW',
  sourcePath: TEMPLATE_PATH,
}

const inventory = buildLayoutsInventory(pkg, structure)
const doc = buildTemplateDoc({ pkg, structure, theme: tokens, meta: minimalMeta, media })

/* ------------------------------------------------------------------ *
 * 1. Inventory shape / counts / geometry
 * ------------------------------------------------------------------ */

await test('inventory: 4 masters / 7 layouts, every layout has a part path', () => {
  assert.equal(inventory.counts.masters, 4, 'four masters')
  assert.equal(inventory.counts.layouts, 7, 'seven layouts')
  assert.equal(inventory.masters.length, 4, 'masters array length')
  assert.deepEqual(
    inventory.masters.map((master) => master.part),
    [
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/slideMasters/slideMaster2.xml',
      'ppt/slideMasters/slideMaster3.xml',
      'ppt/slideMasters/slideMaster4.xml',
    ],
    'master order follows <p:sldMasterIdLst>',
  )

  const layoutParts = inventory.masters.flatMap((master) => master.layouts.map((layout) => layout.part))
  assert.equal(layoutParts.length, 7, 'seven layouts attached to masters')
  assert.equal(new Set(layoutParts).size, 7, 'layout parts are unique')
  for (const part of layoutParts) {
    assert.equal(typeof part, 'string', `layout part is a string (${part})`)
    assert.match(part, /^ppt\/slideLayouts\/slideLayout\d+\.xml$/, `layout part path is a package part (${part})`)
  }

  // Placeholder counts must match the emitted placeholder arrays.
  for (const master of inventory.masters) {
    for (const layout of master.layouts) {
      assert.equal(
        layout.placeholderCount,
        layout.placeholders.length,
        `${layout.part}: placeholderCount matches placeholders.length`,
      )
    }
  }

  let placeholders = 0
  for (const master of inventory.masters) {
    for (const layout of master.layouts) placeholders += layout.placeholders.length
  }
  return `4 masters · 7 layouts · ${placeholders} layout placeholders`
})

await test('inventory: every geometry value is finite and inside the slide (clamp rule)', () => {
  const { widthIn, heightIn } = inventory.slideSize
  assert.ok(Number.isFinite(widthIn) && widthIn > 0, `widthIn is finite (${widthIn})`)
  assert.ok(Number.isFinite(heightIn) && heightIn > 0, `heightIn is finite (${heightIn})`)
  assert.ok(Number.isInteger(inventory.slideSize.cx), 'cx is an EMU integer')

  /** 1% tolerance, per the contract. */
  const toleranceX = widthIn * 0.01
  const toleranceY = heightIn * 0.01

  /** @type {(rect: any, label: string) => void} */
  const checkRect = (rect, label) => {
    for (const key of ['x', 'y', 'w', 'h']) {
      assert.ok(Number.isFinite(rect?.[key]), `${label}: ${key} is finite (${rect?.[key]})`)
    }
    assert.ok(rect.w >= 0 && rect.h >= 0, `${label}: extent is non-negative`)
    assert.ok(rect.x >= -toleranceX, `${label}: x is not off the left edge (${rect.x})`)
    assert.ok(rect.y >= -toleranceY, `${label}: y is not off the top edge (${rect.y})`)
    assert.ok(
      rect.x + rect.w <= widthIn + toleranceX,
      `${label}: right edge inside the slide (${rect.x} + ${rect.w} > ${widthIn})`,
    )
    assert.ok(
      rect.y + rect.h <= heightIn + toleranceY,
      `${label}: bottom edge inside the slide (${rect.y} + ${rect.h} > ${heightIn})`,
    )
  }

  let checked = 0
  for (const master of inventory.masters) {
    if (master.logoFrame !== undefined) {
      checkRect(master.logoFrame, `${master.part} logoFrame`)
      checked += 1
    }
    for (const layout of master.layouts) {
      for (const [index, placeholder] of layout.placeholders.entries()) {
        checkRect(placeholder, `${layout.part} placeholder[${index}] (${placeholder.type})`)
        checked += 1
      }
    }
  }
  assert.ok(checked > 0, 'at least one rectangle was checked')
  return `${checked} rectangles within ${widthIn}×${heightIn} in`
})

/* ------------------------------------------------------------------ *
 * 2. Markdown contents (real colour / font / logo / layout names)
 * ------------------------------------------------------------------ */

await test('markdown: real primary colour, title font, logo pixels and layout names', () => {
  assert.equal(tokens.colors.primary, '#C7000A', 'extracted primary is #C7000A')
  assert.equal(tokens.fonts.title, 'Microsoft YaHei', 'extracted title font is Microsoft YaHei')

  const markdown = doc.markdown
  assert.ok(markdown.includes('#C7000A'), 'markdown contains the primary colour #C7000A')
  assert.ok(markdown.includes('Microsoft YaHei'), 'markdown contains the title font Microsoft YaHei')
  assert.ok(markdown.includes('529×116'), 'markdown contains the logo pixel size 529×116')

  const layoutNames = inventory.masters.flatMap((master) => master.layouts.map((layout) => layout.name))
  console.log(`     layout names: ${layoutNames.join(' / ')}`)
  for (const name of layoutNames) {
    assert.ok(markdown.includes(name), `markdown lists layout "${name}"`)
  }
  return `${layoutNames.length} layout names: ${layoutNames.join(', ')}`
})

/* ------------------------------------------------------------------ *
 * 3. Determinism
 * ------------------------------------------------------------------ */

await test('determinism: two independent builds produce identical bytes', () => {
  const secondInventory = buildLayoutsInventory(pkg, discoverStructure(pkg))
  const secondDoc = buildTemplateDoc({ pkg, structure, theme: tokens, meta: minimalMeta, media })

  const firstMarkdown = Buffer.from(doc.markdown, 'utf8')
  const secondMarkdown = Buffer.from(secondDoc.markdown, 'utf8')
  assert.ok(firstMarkdown.equals(secondMarkdown), 'markdown bytes are identical across builds')
  assert.equal(
    JSON.stringify(secondInventory),
    JSON.stringify(inventory),
    'layouts.json bytes are identical across builds',
  )
  assert.equal(
    JSON.stringify(secondDoc.layouts),
    JSON.stringify(doc.layouts),
    'buildTemplateDoc().layouts is identical across builds',
  )
  return `${firstMarkdown.length} markdown bytes, ${JSON.stringify(inventory).length} layouts.json bytes`
})

/* ------------------------------------------------------------------ *
 * 4. Robustness (empty media, minimal theme)
 * ------------------------------------------------------------------ */

await test('robustness: empty media + minimal theme renders without throwing', () => {
  const minimalTheme = {
    templateId: 'minimal',
    templateName: 'Minimal Theme',
    colors: {
      background: '#FFFFFF',
      surface: '#F2F2F2',
      textPrimary: '#111111',
      textSecondary: '#777777',
      primary: '#0055AA',
      accent: '#FF8800',
      chartPalette: ['#0055AA', '#FF8800'],
    },
    fonts: { title: 'Arial', body: 'Arial', latin: 'Arial' },
    brandAssets: { icons: {} },
  }

  const markdown = buildTemplateMarkdown({
    theme: minimalTheme,
    meta: { templateId: 'minimal', templateName: 'Minimal Theme' },
    media: [],
    layouts: null,
  })
  assert.equal(typeof markdown, 'string', 'markdown is a string')
  assert.ok(markdown.length > 0, 'markdown is not empty')
  assert.ok(!FORBIDDEN.test(markdown), `minimal markdown has no undefined/NaN/[object Object]\n${markdown}`)
  assert.ok(markdown.includes('#0055AA'), 'minimal primary colour is rendered')
  assert.ok(markdown.includes('Arial'), 'minimal font is rendered')

  // Nothing at all must still not throw.
  const bare = buildTemplateMarkdown({})
  assert.ok(typeof bare === 'string' && bare.length > 0, 'empty input still renders a document')
  assert.ok(!FORBIDDEN.test(bare), 'empty-input markdown has no undefined/NaN/[object Object]')

  // Empty media against the real inventory (the "template without images" case).
  const withoutMedia = buildTemplateMarkdown({
    theme: minimalTheme,
    meta: minimalMeta,
    media: [],
    layouts: inventory,
  })
  assert.ok(!FORBIDDEN.test(withoutMedia), 'real inventory + empty media has no forbidden strings')
  assert.equal(withoutMedia.includes('#C7000A'), false, 'minimal theme does not leak the real primary')
  return `${markdown.split('\n').length} lines from the minimal theme`
})

/* ------------------------------------------------------------------ *
 * 5. Whole-document regex hygiene
 * ------------------------------------------------------------------ */

await test('markdown: no undefined/NaN/[object Object] and no trailing whitespace', () => {
  const markdown = doc.markdown
  assert.equal(FORBIDDEN.test(markdown), false, 'no undefined / NaN / [object Object]')

  const offending = markdown
    .split('\n')
    .map((line, index) => ({ line, index }))
    .filter((entry) => /[ \t]+$/.test(entry.line))
  assert.deepEqual(
    offending,
    [],
    `no trailing whitespace:\n${offending.map((entry) => `${entry.index + 1}: ${JSON.stringify(entry.line)}`).join('\n')}`,
  )
  assert.ok(markdown.endsWith('\n'), 'document ends with exactly one newline')
  assert.equal(markdown.endsWith('\n\n'), false, 'document does not end with a blank line')

  // Structural sanity: every required section is present, in order.
  const sections = ['## 色板', '## 字体', '## 页面', '## 版式清单', '## 品牌素材', '## 复用方式']
  let cursor = -1
  for (const heading of sections) {
    const at = markdown.indexOf(heading)
    assert.ok(at > cursor, `section ${heading} present and after the previous one`)
    cursor = at
  }
  const bullets = markdown.split('## 复用方式')[1].split('\n').filter((line) => line.startsWith('- '))
  assert.equal(bullets.length, 3, `复用方式 has exactly 3 bullet lines (got ${bullets.length})`)
  assert.ok(!/[ \t]+$/m.test(markdown), 'no trailing whitespace anywhere')
  assert.ok(markdown.startsWith('# PPT模板-HW\n'), 'document starts with the template name heading')
  return `${markdown.split('\n').length} lines, ${markdown.length} chars`
})

/* ------------------------------------------------------------------ *
 * Summary
 * ------------------------------------------------------------------ */

const failed = results.filter((entry) => !entry.ok)
console.log(`\ntemplate-doc: ${results.length - failed.length}/${results.length} checks passed`)
if (failed.length > 0) {
  console.log(`failed: ${failed.map((entry) => entry.name).join(', ')}`)
  process.exitCode = 1
}
