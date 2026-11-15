/**
 * `/branddeck` command test — the session-native conversation entry.
 *
 * Exercises the pure parser and the handler's non-model branches (help, list,
 * templates, argument errors) against a real store. The model-backed branches are
 * covered live (they need a provider), so this test proves everything that can be
 * proven deterministically: parsing, routing, transcript text, and error copy.
 *
 * Run: node test/command-test.mjs
 */
import { mkdirSync, rmSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const pkgRoot = resolve(here, '..')
const storeRoot = join(pkgRoot, '.tmp', 'command-store')
rmSync(storeRoot, { recursive: true, force: true })
mkdirSync(storeRoot, { recursive: true })

const { __internals, parseBrandDeckCommand } = await import('../index.js')
const store = await import('../src/store.js')
const styles = await import('../src/styles.js')
const { HW_DEMO_DECK } = await import('./fixtures/hw-demo-deck.mjs')

const config = { defaultTemplateId: undefined, commandMaxSlides: 10 }
const storeConfig = { root: storeRoot }

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

/** Invoke the command handler through the same entry the Host uses. */
async function run(line) {
  return __internals.executeBrandDeckCommand(
    { rawInput: line, agent: { session: { id: 'session-test' } }, attachments: [] },
    { ctx: null, config, storeConfig },
  )
}

/* 1. Parser ---------------------------------------------------------------- */
check('parser routes the sub-commands', () => {
  const cases = [
    ['', 'help'],
    ['help', 'help'],
    ['list', 'list'],
    ['templates', 'templates'],
    ['style', 'style-help'],
    ['style 华为红、卡片', 'style'],
    ['deck 给周会做一页', 'deck'],
    ['给周会做一页', 'deck'],
    ['deck --style hw-dense 讲三个要点', 'deck'],
  ]
  for (const [line, expected] of cases) {
    const parsed = parseBrandDeckCommand(line)
    assert(parsed.kind === expected, `"${line}" -> ${parsed.kind}, expected ${expected}`)
  }
  return `${cases.length} forms`
})

check('parser keeps the description intact (中文与标点不丢)', () => {
  const bare = parseBrandDeckCommand('给部门周会做一页：本周做完模板解析，下周启动导出')
  assert(bare.kind === 'deck', bare.kind)
  assert(bare.description === '给部门周会做一页：本周做完模板解析，下周启动导出', bare.description)

  const styled = parseBrandDeckCommand('deck --style hw-dense 讲三个要点')
  assert(styled.styleId === 'hw-dense', `styleId: ${styled.styleId}`)
  assert(styled.description === '讲三个要点', `description: ${styled.description}`)

  const styleCmd = parseBrandDeckCommand('style 华为红、信息密度高')
  assert(styleCmd.description === '华为红、信息密度高', styleCmd.description)
  return '3 shapes'
})

/* 2. Handler: help + status ------------------------------------------------ */
const help = await run('')
check('/branddeck with no argument returns usage plus live status', () => {
  assert(help.kind === 'success', help.kind)
  assert(help.text.includes('/branddeck deck'), 'usage line missing')
  assert(help.text.includes('模板 0 个'), `status missing: ${help.text}`)
  return help.text.split('\n').length + ' lines'
})

/* 3. Templates / list empty states ---------------------------------------- */
const noTemplates = await run('templates')
check('templates lists nothing yet, and says how to fix it', () => {
  assert(noTemplates.kind === 'success', noTemplates.kind)
  assert(noTemplates.text.includes('branddeck_template_parse'), noTemplates.text)
})

const noDecks = await run('list')
check('list lists nothing yet, and suggests the deck command', () => {
  assert(noDecks.kind === 'success', noDecks.kind)
  assert(noDecks.text.includes('/branddeck deck'), noDecks.text)
})

/* 4. Deck without a template is refused with actionable copy -------------- */
const noTemplateDeck = await run('deck 给周会做一页')
check('deck without a parsed template is refused with instructions', () => {
  assert(noTemplateDeck.kind === 'error', noTemplateDeck.kind)
  assert(noTemplateDeck.text.includes('branddeck_template_parse'), noTemplateDeck.text)
  return noTemplateDeck.text.slice(0, 40)
})

/* 5. Missing description on a sub-command -------------------------------- */
const styleNoText = await run('style')
check('style without a description explains the syntax', () => {
  assert(styleNoText.kind === 'error', styleNoText.kind)
  assert(styleNoText.text.includes('/branddeck style'), styleNoText.text)
})

/* 6. Populated store: list / templates render real rows ------------------ */
const theme = {
  templateId: 'ppt-hw',
  templateName: 'PPT模板-HW',
  colors: {
    background: '#FFFFFF',
    surface: '#EBEBEB',
    textPrimary: '#1D1D1A',
    textSecondary: '#666666',
    primary: '#C7000A',
    accent: '#E9002F',
    chartPalette: ['#C7000A'],
  },
  fonts: { title: 'Microsoft YaHei', body: 'Arial', latin: 'Arial' },
  brandAssets: { icons: {} },
}
store.writeTemplate(
  { templateId: 'ppt-hw', templateName: 'PPT模板-HW', tokens: theme, meta: { slideCount: 5 }, media: [] },
  storeConfig,
)
store.writeDeck({ ...HW_DEMO_DECK, templateId: 'ppt-hw' }, storeConfig, { theme })
styles.writeStyle({ name: '华为红·紧凑', density: 'dense', bulletLayout: 'cards', source: 'chat' }, storeConfig)

const listed = await run('list')
check('list renders real decks with page counts', () => {
  assert(listed.kind === 'success', listed.kind)
  assert(listed.text.includes('hw-demo'), listed.text)
  assert(/（\d+ 页/.test(listed.text), `page count missing: ${listed.text}`)
  return listed.text.split('\n').length - 1 + ' deck(s)'
})

const templateList = await run('templates')
check('templates renders real template rows', () => {
  assert(templateList.kind === 'success', templateList.kind)
  assert(templateList.text.includes('ppt-hw'), templateList.text)
  assert(templateList.text.includes('#C7000A'), templateList.text)
})

const helpWithData = await run('')
check('help status reflects the populated store', () => {
  assert(helpWithData.text.includes('模板 1 个'), helpWithData.text)
  assert(helpWithData.text.includes('卡组 1 个'), helpWithData.text)
  assert(helpWithData.text.includes('风格 2 个'), helpWithData.text)
})

/* 7. The model-backed branch without an llm must fail cleanly ------------ */
const noLlm = await run('deck 讲三个要点')
check('deck without an llm service reports a readable error (no crash)', () => {
  assert(noLlm.kind === 'error', noLlm.kind)
  assert(noLlm.text.includes('调用模型失败'), noLlm.text)
  return noLlm.text.slice(0, 60)
})

check('every answer is a renderable transcript result', () => {
  const answers = [help, noTemplates, noDecks, noTemplateDeck, styleNoText, listed, templateList, noLlm]
  for (const answer of answers) {
    assert(answer !== null && typeof answer === 'object', 'answer is not an object')
    assert(answer.kind === 'success' || answer.kind === 'error', `bad kind: ${answer.kind}`)
    assert(typeof answer.text === 'string' && answer.text.length > 0, 'empty text')
  }
  return `${answers.length} answers`
})

check('help text and parser agree on every advertised form', () => {
  // The composer shows COMMAND_USAGE; if it advertised a form the parser rejects the
  // user gets a help screen for a command that cannot work. Derive both from the
  // same source of truth: run each documented example through the parser.
  const examples = [
    ['deck 一句话描述', 'deck'],
    ['deck --style hw-dense 一句话描述', 'deck'],
    ['style 一句话描述', 'style'],
    ['list', 'list'],
    ['templates', 'templates'],
  ]
  for (const [line, kind] of examples) {
    const parsed = parseBrandDeckCommand(line)
    assert(parsed.kind === kind, `"${line}" parsed as ${parsed.kind}, help says ${kind}`)
  }
  assert(help.text.includes('/branddeck deck <一句话描述>'), 'usage line drifted from the parser')
  assert(help.text.includes('/branddeck style <一句话描述>'), 'style usage line missing')
  return `${examples.length} forms`
})

const failed = results.filter((entry) => !entry.ok)
console.log(`\n${results.length - failed.length}/${results.length} command checks passed`)
if (failed.length > 0) process.exitCode = 1
