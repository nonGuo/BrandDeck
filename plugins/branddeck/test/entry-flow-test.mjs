/**
 * Entry-point test: the "describe it and get slides" path.
 *
 * Mounts the real Host handler on a node:http server and drives it the way the
 * panel does: import a template → save a style profile → generate a deck from
 * pasted text → read the text twin → export a .pptx.
 *
 * Run: node test/entry-flow-test.mjs
 */
import { createServer } from 'node:http'
import { existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const pkgRoot = resolve(here, '..')
const repoRoot = resolve(pkgRoot, '..', '..')
const templatePath = join(repoRoot, 'PPT模板-HW.pptx')
const storeRoot = join(pkgRoot, '.tmp', 'entry-store')

rmSync(storeRoot, { recursive: true, force: true })
mkdirSync(storeRoot, { recursive: true })

const { __internals } = await import('../index.js')
const config = { svgRasterScale: 2, loadSystemFonts: true }
const storeConfig = { root: storeRoot }
const deps = { config, storeConfig }

const server = createServer((req, res) => {
  __internals.handleApiRequest(req, res, deps).catch(() => {
    if (!res.headersSent) {
      res.writeHead(500)
      res.end()
    }
  })
})
await new Promise((done) => server.listen(0, '127.0.0.1', done))
const base = `http://127.0.0.1:${server.address().port}/branddeck-api`

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

async function post(endpoint, payload) {
  const response = await fetch(`${base}/${endpoint}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(payload ?? {}),
  })
  const text = await response.text()
  return { status: response.status, envelope: JSON.parse(text) }
}
const ok = (result, label) => {
  assert(result.envelope.ok === true, `${label}: ${JSON.stringify(result.envelope.error ?? result.envelope).slice(0, 300)}`)
  return result.envelope.value
}

/* 1. Import a template ------------------------------------------------ */
const imported = await post('templates.parse', {
  path: templatePath,
  templateId: 'ppt-hw',
  templateName: 'PPT模板-HW',
})
check('entry step 1 — import a .pptx template', () => {
  const value = ok(imported, 'templates.parse')
  assert(value.theme.colors.primary.toUpperCase() === '#C7000A', `primary ${value.theme.colors.primary}`)
  return `${value.template.templateId} · ${value.theme.fonts.title}`
})

/* 2. Style profiles --------------------------------------------------- */
const stylesList = await post('styles.list')
check('entry step 2 — a default style is always available', () => {
  const value = ok(stylesList, 'styles.list')
  assert(value.styles.length >= 1, 'no style listed')
  assert(value.styles[0].styleId === 'default', `first style ${value.styles[0].styleId}`)
  return `${value.styles.length} style(s)`
})

const saved = await post('styles.save', {
  style: {
    name: '华为红·紧凑卡片',
    description: '华为红强调、信息密度高、卡片式正文',
    templateId: 'ppt-hw',
    density: 'dense',
    tone: 'bold',
    bulletLayout: 'cards',
    maxSlides: 10,
    source: 'chat',
  },
})
check('entry step 2 — a conversational style is persisted and reusable', () => {
  const value = ok(saved, 'styles.save')
  assert(value.style.source === 'chat', `source ${value.style.source}`)
  assert(value.style.bulletLayout === 'cards', 'bulletLayout not stored')
  assert(value.styles.some((entry) => entry.styleId === value.style.styleId), 'style not listed after save')
  return `${value.style.styleId} (${value.style.density}/${value.style.tone}/${value.style.bulletLayout})`
})
const styleId = saved.envelope.value.style.styleId

check('entry step 2 — a style is stored as a readable JSON file', () => {
  const file = join(storeRoot, 'styles', `${styleId}.json`)
  assert(existsSync(file), `missing ${file}`)
  const onDisk = JSON.parse(readFileSync(file, 'utf8'))
  assert(onDisk.name === '华为红·紧凑卡片', 'name not persisted')
  assert(typeof onDisk.description === 'string' && onDisk.description.length > 0, 'description lost')
})

/* 3. Generate from pasted text ---------------------------------------- */
const SOURCE = [
  '# 2026 年度品牌内容规划',
  '',
  '## 业务挑战',
  '- 内容需求增长 3 倍',
  '- 排版与合规返工消耗 60% 交付时间',
  '- 各团队自建模板导致品牌漂移',
  '',
  '备注：数据来自内部复盘。',
  '',
  '## 增长数据',
  '- 内容交付量同比增长 138%',
  '- 返工工时下降 74%',
  '- 单页产出效率提升 3.2 倍',
  '',
  '## 落地路径',
  '- 模板解析一次即可复用',
  '- 卡组编排可视化微调',
  '- 原生导出可编辑 PPT',
].join('\n')

const generated = await post('generate', { text: SOURCE, templateId: 'ppt-hw', styleId })
let deckId = null
check('entry step 3 — pasted text + style produce a deck', () => {
  const value = ok(generated, 'generate')
  deckId = value.deckId
  assert(typeof deckId === 'string' && deckId.length > 0, 'no deckId')
  assert(value.slideCount >= 4, `only ${value.slideCount} slides`)
  assert(Array.isArray(value.warnings), 'warnings must be an array')
  return `${deckId} · ${value.slideCount} slides · ${value.warnings.length} warning(s)`
})

check('generated deck carries the style (cards layout + brand colour)', () => {
  const value = generated.envelope.value
  const elements = value.deck.slides.flatMap((slide) => slide.elements)
  const rects = elements.filter((element) => element.type === 'shape' && element.shapeType === 'roundRect')
  assert(rects.length > 0, 'no card rectangles — bulletLayout:cards not applied')
  const fills = rects.map((rect) => String(rect.style?.fill ?? '').toUpperCase())
  assert(
    fills.some((fill) => fill === '#EBEBEB' || fill === '#C7000A'),
    `unexpected card fills: ${[...new Set(fills)].join(', ')}`,
  )
  return `${rects.length} card(s), fills ${[...new Set(fills)].join('/')}`
})

check('generated deck is inside the canvas and passes the shared contract', () => {
  const value = generated.envelope.value
  for (const slide of value.deck.slides) {
    for (const element of slide.elements) {
      assert(element.x >= 0 && element.y >= 0, `${element.id} negative origin`)
      assert(element.x + element.w <= 1280 + 1, `${element.id} exceeds width`)
      assert(element.y + element.h <= 720 + 1, `${element.id} exceeds height`)
    }
  }
  return `${value.deck.slides.length} slides checked`
})

/* 4. The text twin exists for the generated deck ---------------------- */
const outline = await post('outline.get', { deckId })
check('entry step 3 — the generated deck has an editable text twin', () => {
  const value = ok(outline, 'outline.get')
  assert(value.markdown.startsWith('# '), 'outline has no deck title')
  assert(value.markdown.includes('业务挑战'), 'slide title missing from outline')
  assert(value.markdown.includes('{#'), 'element ids missing from outline')
  return `${value.markdown.length} chars`
})

/* 5. Generation from a file path ------------------------------------- */
const filePath = join(storeRoot, 'input.md')
mkdirSync(dirname(filePath), { recursive: true })
const { writeFileSync } = await import('node:fs')
writeFileSync(filePath, SOURCE, 'utf8')
const fromFile = await post('generate', { path: filePath, templateId: 'ppt-hw', title: '文件来源测试' })
check('entry step 3 — a .md path works as the content source', () => {
  const value = ok(fromFile, 'generate')
  assert(value.deck.title === '文件来源测试', `title override ignored: ${value.deck.title}`)
  return `${value.deckId} · ${value.slideCount} slides`
})

/* 6. Failure paths ---------------------------------------------------- */
const emptySource = await post('generate', { text: '   ', templateId: 'ppt-hw' })
check('entry — empty input fails with a clear code (no crash)', () => {
  assert(emptySource.envelope.ok === false, 'expected failure')
  assert(emptySource.envelope.error.code === 'invalid_body', `code ${emptySource.envelope.error.code}`)
})

const badStyle = await post('generate', { text: SOURCE, styleId: 'no-such-style' })
check('entry — unknown style fails with style_not_found', () => {
  assert(badStyle.envelope.ok === false, 'expected failure')
  assert(badStyle.envelope.error.code === 'style_not_found', `code ${badStyle.envelope.error.code}`)
})

/* 7. Export what was generated --------------------------------------- */
const exported = await post('export', { deckId })
check('entry step 3 — the generated deck exports to .pptx', () => {
  const value = ok(exported, 'export')
  assert(existsSync(value.path), `missing ${value.path}`)
  assert(value.slideCount >= 4, `slideCount ${value.slideCount}`)
  return `${value.path.split(/[\\/]/).pop()} (${value.slideCount} slides)`
})

await new Promise((done) => server.close(done))

const failed = results.filter((entry) => !entry.ok)
console.log(`\n${results.length - failed.length}/${results.length} entry-flow checks passed`)
if (failed.length > 0) process.exitCode = 1
