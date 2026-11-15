/**
 * Transport test for the panel's HTTP surface.
 *
 * The Web panel talks to `POST /branddeck-api/<endpoint>`. This mounts the exact
 * handler the Host registers (`handleApiRequest`) on a real `node:http` server and
 * drives it with real HTTP requests, so the contract the browser sees — method
 * handling, JSON body parsing, status codes, envelope shape — is verified without
 * a browser or a running Harness.
 *
 * Run: node test/http-api-test.mjs
 */
import { createServer } from 'node:http'
import { mkdir, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const pkgRoot = resolve(here, '..')
const repoRoot = resolve(pkgRoot, '..', '..')
const storeRoot = join(pkgRoot, '.tmp', 'http-store')
const templatePath = join(repoRoot, 'PPT模板-HW.pptx')

await rm(storeRoot, { recursive: true, force: true })
await mkdir(storeRoot, { recursive: true })

const { __internals } = await import('../index.js')
const { API_PREFIX } = await import('../index.js')

const config = { svgRasterScale: 2, loadSystemFonts: true }
const storeConfig = { root: storeRoot }
const deps = { config, storeConfig }

// Mount exactly what apply() mounts, minus the Harness web server.
const server = createServer((req, res) => {
  __internals.handleApiRequest(req, res, deps).catch(() => {
    if (!res.headersSent) {
      res.writeHead(500)
      res.end()
    }
  })
})
await new Promise((done) => server.listen(0, '127.0.0.1', done))
const port = server.address().port
const base = `http://127.0.0.1:${port}${API_PREFIX}`

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

/** POST one endpoint the way the browser does. */
async function post(endpoint, payload) {
  const response = await fetch(`${base}/${endpoint}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(payload ?? {}),
  })
  const text = await response.text()
  let envelope
  try {
    envelope = JSON.parse(text)
  } catch {
    throw new Error(`non-JSON response (${response.status}): ${text.slice(0, 120)}`)
  }
  return { status: response.status, envelope }
}

/* 1. Envelope + empty state ------------------------------------------------ */
const state = await post('state')
check('POST /state returns 200 and the { ok, value } envelope', () => {
  assert(state.status === 200, `status ${state.status}`)
  assert(state.envelope.ok === true, JSON.stringify(state.envelope))
  assert(state.envelope.value.deck === null, 'expected empty deck')
  assert(Array.isArray(state.envelope.value.decks), 'decks must be an array')
  return `decks=${state.envelope.value.decks.length}`
})

/* 2. Template parse over HTTP --------------------------------------------- */
const parsed = await post('templates.parse', { path: templatePath, templateId: 'ppt-hw', templateName: 'PPT模板-HW' })
check('POST /templates.parse stores the corporate template', () => {
  assert(parsed.envelope.ok === true, JSON.stringify(parsed.envelope))
  assert(parsed.envelope.value.theme.colors.primary.toUpperCase() === '#C7000A', 'wrong primary color')
  return parsed.envelope.value.theme.colors.primary
})

const logo = await post('templates.logo', { templateId: 'ppt-hw' })
check('POST /templates.logo returns a data URL', () => {
  assert(logo.envelope.ok === true, JSON.stringify(logo.envelope))
  assert(logo.envelope.value.dataUrl.startsWith('data:image/png;base64,'), 'not a png data URL')
  return `${logo.envelope.value.widthPx}x${logo.envelope.value.heightPx}`
})

/* 3. Deck create / save / get over HTTP ----------------------------------- */
const created = await post('decks.create', { title: 'HTTP 测试卡组', templateId: 'ppt-hw' })
check('POST /decks.create persists a starter deck', () => {
  assert(created.envelope.ok === true, JSON.stringify(created.envelope))
  assert(created.envelope.value.deck.slides.length >= 1, 'no slides')
  return created.envelope.value.deck.deckId
})

const saved = await post('decks.save', {
  deck: {
    deckId: 'http-demo',
    title: 'HTTP 测试卡组',
    templateId: 'ppt-hw',
    slides: [
      {
        slideId: 's1',
        background: '$colors.background',
        elements: [
          {
            type: 'text',
            id: 't1',
            role: 'title',
            content: '通过 HTTP 路由保存的标题',
            x: 96,
            y: 120,
            w: 900,
            h: 80,
            style: { fontSize: 44, bold: true, color: '$colors.primary' },
          },
        ],
      },
    ],
  },
})
check('POST /decks.save repairs and stores a deck', () => {
  assert(saved.envelope.ok === true, JSON.stringify(saved.envelope))
  assert(saved.envelope.value.deck.slides.length === 1, 'slide count changed')
  assert(Array.isArray(saved.envelope.value.warnings), 'warnings must be an array')
  return `${saved.envelope.value.warnings.length} warning(s)`
})

const fetched = await post('decks.get', { deckId: 'http-demo' })
check('POST /decks.get returns the deck plus its theme', () => {
  assert(fetched.envelope.ok === true, JSON.stringify(fetched.envelope))
  assert(fetched.envelope.value.deck.deckId === 'http-demo', 'wrong deck')
  assert(fetched.envelope.value.theme.colors.primary.toUpperCase() === '#C7000A', 'theme missing')
})

/* 4. Failure envelopes and status codes ----------------------------------- */
const missing = await post('decks.get', { deckId: 'nope' })
check('missing deck → 404 with deck_not_found', () => {
  assert(missing.status === 404, `status ${missing.status}`)
  assert(missing.envelope.ok === false, 'expected failure envelope')
  assert(missing.envelope.error.code === 'deck_not_found', missing.envelope.error.code)
})

const badDeck = await post('decks.save', { deck: { deckId: 'x', title: '', templateId: 't', slides: [] } })
check('invalid deck → 422 with structured details', () => {
  assert(badDeck.status === 422, `status ${badDeck.status}`)
  assert(badDeck.envelope.error.code === 'invalid_deck', badDeck.envelope.error.code)
  assert(Array.isArray(badDeck.envelope.error.details), 'missing details')
  return `${badDeck.envelope.error.details.length} issue(s)`
})

const unknown = await post('does.not.exist')
check('unknown endpoint → 404 with not_found', () => {
  assert(unknown.status === 404, `status ${unknown.status}`)
  assert(unknown.envelope.error.code === 'not_found', unknown.envelope.error.code)
})

try {
  const wrongMethod = await fetch(`${base}/state`, { method: 'PUT' })
  check('PUT → 405 (only GET/POST are accepted)', () => {
    assert(wrongMethod.status === 405, `status ${wrongMethod.status}`)
  })
} catch (error) {
  check('PUT → 405 (only GET/POST are accepted)', () => {
    throw error
  })
}

/* 5. Export over HTTP ------------------------------------------------------ */
const exported = await post('export', { deckId: 'http-demo' })
check('POST /export compiles a .pptx and reports its path', () => {
  assert(exported.envelope.ok === true, JSON.stringify(exported.envelope))
  assert(typeof exported.envelope.value.path === 'string', 'missing path')
  assert(exported.envelope.value.slideCount === 1, `slideCount ${exported.envelope.value.slideCount}`)
  return exported.envelope.value.path.split(/[\\/]/).pop()
})

/* 6. Payload edge cases ---------------------------------------------------- */
const emptyBody = await fetch(`${base}/state`, { method: 'POST' })
const emptyResult = await emptyBody.text()
check('empty body → 200 with a JSON envelope', () => {
  assert(emptyBody.status === 200, `status ${emptyBody.status}`)
  const envelope = JSON.parse(emptyResult)
  assert(envelope.ok === true, JSON.stringify(envelope))
  return 'ok'
})

await new Promise((done) => server.close(done))

const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} HTTP transport checks passed`)
if (failed.length > 0) process.exitCode = 1