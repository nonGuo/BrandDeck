#!/usr/bin/env node
/**
 * Static + stub-runtime audit of the BrandDeck web client half.
 *
 * No browser is involved: the file is inspected as text and then evaluated in a
 * `node:vm` sandbox whose `window.__ModuleLoader__` captures the plugin
 * definition. A minimal fake React and a stub Cordis context let the test call
 * `apply()` and assert the two slot registrations without a DOM.
 *
 * Run: node plugins/branddeck/test/client-smoke.mjs
 */
import { readFileSync, existsSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import vm from 'node:vm'

const HERE = dirname(fileURLToPath(import.meta.url))
const CLIENT_PATH = resolve(HERE, '..', 'client.js')
const WORKSPACE_ROOT = resolve(HERE, '..', '..', '..')
const THEME_TOKENS_PATH = resolve(
  WORKSPACE_ROOT,
  '.dsh-src/dsh/node_modules/@deepseek-ai/dsh-client-ui-theme/lib/client.js',
)

const source = readFileSync(CLIENT_PATH, 'utf8')

let failures = 0
let passed = 0

/** Run one named check; a thrown error or `false` marks it failed. */
function check(name, body) {
  try {
    const verdict = body()
    if (verdict === false) throw new Error('check returned false')
    passed += 1
    console.log(`PASS  ${name}`)
  } catch (error) {
    failures += 1
    console.log(`FAIL  ${name} :: ${error instanceof Error ? error.message : String(error)}`)
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message ?? 'assertion failed')
}

/**
 * Remove string literals, template literals, comments and regex literals so the
 * remainder is pure code. Used to prove the file holds no raw JSX markup.
 */
function stripNonCode(text) {
  let out = ''
  let index = 0
  const regexPrefix = '(,=:[!&|?{};+-*%<>~^'
  let previousSignificant = ''
  while (index < text.length) {
    const char = text[index]
    const next = text[index + 1]
    if (char === '/' && next === '/') {
      const newline = text.indexOf('\n', index)
      index = newline === -1 ? text.length : newline
      continue
    }
    if (char === '/' && next === '*') {
      const end = text.indexOf('*/', index + 2)
      index = end === -1 ? text.length : end + 2
      continue
    }
    if (char === "'" || char === '"' || char === '`') {
      const quote = char
      index += 1
      while (index < text.length) {
        const inner = text[index]
        if (inner === '\\') {
          index += 2
          continue
        }
        if (inner === quote) {
          index += 1
          break
        }
        index += 1
      }
      previousSignificant = quote
      continue
    }
    if (char === '/' && (previousSignificant === '' || regexPrefix.includes(previousSignificant))) {
      index += 1
      let inClass = false
      while (index < text.length) {
        const inner = text[index]
        if (inner === '\\') {
          index += 2
          continue
        }
        if (inner === '[') inClass = true
        else if (inner === ']') inClass = false
        else if (inner === '/' && inClass === false) {
          index += 1
          break
        } else if (inner === '\n') break
        index += 1
      }
      previousSignificant = '/'
      continue
    }
    if (!/\s/.test(char)) previousSignificant = char
    out += char
    index += 1
  }
  return out
}

/** All string literals of the source that carry CJK text. */
function cjkLiterals(text) {
  const found = []
  const pattern = /'([^'\n]*)'|"([^"\n]*)"/g
  let match
  while ((match = pattern.exec(text)) !== null) {
    const value = match[1] ?? match[2] ?? ''
    if (/[\u4e00-\u9fff]/.test(value)) found.push(value)
  }
  return [...new Set(found)]
}

/** Dict keys of a `const NAME = { ... }` object literal slice. */
function dictKeys(slice) {
  const keys = new Set()
  const pattern = /'([^']+)':/g
  let match
  while ((match = pattern.exec(slice)) !== null) keys.add(match[1])
  return keys
}

/* ------------------------------------------------------------------ *
 * 1. Module format
 * ------------------------------------------------------------------ */

check('loads through window.__ModuleLoader__.load', () => {
  assert(/window\.__ModuleLoader__\.load\(\{/.test(source), 'no __ModuleLoader__.load call')
})

check("registers the id '@local/dsh-branddeck'", () => {
  assert(/id:\s*'@local\/dsh-branddeck'/.test(source), 'definition id missing')
})

check('is a factory(require) definition', () => {
  assert(/factory\s*\(\s*require\s*\)\s*\{/.test(source), 'factory(require) missing')
})

check("holds exactly one require('react') call", () => {
  const calls = source.match(/require\s*\(/g) ?? []
  assert(calls.length === 1, `found ${calls.length} require() calls`)
  assert(/const React = require\('react'\)/.test(source), 'require target is not react')
})

check('has no ES module syntax', () => {
  assert(!/\bimport\s+[\w{*]/.test(source), 'import statement found')
  assert(!/\bexport\s+(default|const|let|var|function|class|\{)/.test(source), 'export statement found')
})

check('has no bundler or CDN dependency', () => {
  assert(!/https?:\/\/(?!www\.w3\.org)/.test(source), 'external URL found')
})

/* ------------------------------------------------------------------ *
 * 2. Slot registrations
 * ------------------------------------------------------------------ */

check('injects both slots through ctx.slots.inject', () => {
  const injects = source.match(/ctx\.slots\.inject\(/g) ?? []
  assert(injects.length === 2, `expected 2 ctx.slots.inject calls, found ${injects.length}`)
  assert(source.includes("ctx.slots.inject('main'"), "'main' slot not injected")
  assert(source.includes("ctx.slots.inject('sidebar.panellist'"), "'sidebar.panellist' slot not injected")
})

check('registers the keyed main page', () => {
  assert(source.includes("name: 'main'"), "main registration name missing")
  assert(source.includes('key: PANEL_ID'), 'main registration key missing')
  assert(source.includes("const PANEL_ID = 'branddeck'"), 'PANEL_ID is not branddeck')
})

check('registers the sidebar.panellist icon', () => {
  assert(source.includes("name: 'sidebar.panellist'"), 'sidebar registration name missing')
  assert(source.includes('id: PANEL_ID'), 'sidebar registration id missing')
  assert(/order:\s*20\b/.test(source), 'sidebar order missing')
  assert(source.includes("label: () => t('panel')"), 'sidebar label missing')
})

check('injects only slots (no unreachable service dependency)', () => {
  // A hard `connection` dependency leaves the client fiber pending forever when
  // the page hides that service, which fails the whole web boot ("<id>: loading").
  assert(/inject:\s*\['slots'\]/.test(source), "inject list must be exactly ['slots']")
  assert(!/inject:\s*\[[^\]]*'connection'/.test(source), "'connection' must not be a hard inject")
  assert(/const API_BASE = '/.test(source), 'same-origin API base missing')
})

/* ------------------------------------------------------------------ *
 * 3. Markup / DOM discipline
 * ------------------------------------------------------------------ */

check('contains no raw JSX markup in code positions', () => {
  const code = stripNonCode(source)
  const tag = /<[A-Za-z/]/.exec(code)
  assert(tag === null, `JSX-like markup in code: ${tag === null ? '' : code.slice(Math.max(0, tag.index - 40), tag.index + 40)}`)
})

check('never writes document.body', () => {
  assert(!/document\s*\.\s*body/.test(source), 'document.body referenced')
})

check('appends exactly one stylesheet through document.head', () => {
  const heads = source.match(/document\.head\.appendChild\(/g) ?? []
  assert(heads.length === 1, `expected 1 appendChild on document.head, found ${heads.length}`)
  assert(/document\.head\.removeChild\(|style\.parentNode\.removeChild\(/.test(source), 'stylesheet is never removed')
})

check('never imports the client UI primitives package', () => {
  assert(!source.includes('@deepseek-ai/dsh-client-ui-primitives'), 'primitives package referenced')
})

/* ------------------------------------------------------------------ *
 * 4. Transport: generic Connection RPC channel
 * ------------------------------------------------------------------ */

check('uses the same-origin Host route only', () => {
  // A bundle client half cannot reach the shipped plugins' `connection` service
  // (it is provided on their own contexts), so the panel talks to its Host half
  // over the `/branddeck-api` prefix the Host registers on the web server.
  assert(source.includes("const API_BASE = '/branddeck-api'"), 'API base constant missing')
  assert(/fetch\(\s*url/.test(source), 'fetch is not routed through the API base')
  assert(!source.includes('ctx.connection'), 'unreachable ctx.connection still read')
  assert(!source.includes('RPC_CHANNEL'), 'stale RPC channel constant present')
})

check('uses only the documented RPC endpoints', () => {
  const allowed = new Set(['state', 'decks.list', 'decks.get', 'decks.save', 'decks.create', 'export', 'templates.parse', 'templates.logo', 'outline.get', 'outline.save', 'decks.materialize', 'templates.spec', 'styles.list', 'styles.save', 'styles.delete', 'generate', 'style.assist', 'content.assist', 'assist.providers'])
  const used = [...source.matchAll(/callRpc\('([^']+)'/g)].map((match) => match[1])
  assert(used.length > 0, 'no callRpc endpoint found')
  for (const endpoint of used) assert(allowed.has(endpoint), `invented endpoint "${endpoint}"`)
})

check('maps deck_not_found to the empty state', () => {
  assert(source.includes('CODE_DECK_NOT_FOUND'), 'deck_not_found handling missing')
})

check('bounds every call with a timeout', () => {
  assert(/REQUEST_TIMEOUT_MS = 8000/.test(source), 'request timeout constant missing')
  assert(/EXPORT_TIMEOUT_MS = 120000/.test(source), 'export timeout constant missing')
})

/* ------------------------------------------------------------------ *
 * 5. Stylesheet discipline
 * ------------------------------------------------------------------ */

check('prefixes every stylesheet selector with branddeck-', () => {
  const start = source.indexOf('/* branddeck:stylesheet:start */')
  const end = source.indexOf('/* branddeck:stylesheet:end */')
  assert(start !== -1 && end > start, 'stylesheet markers missing')
  const css = source.slice(start, end)
  const allowed = ['.branddeck-']
  let match
  const pattern = /(?:^|\n)\s*([^{}@\n][^{}\n]*?)\{/g
  let checked = 0
  while ((match = pattern.exec(css)) !== null) {
    const selector = match[1].trim()
    for (const part of selector.split(',')) {
      const one = part.trim()
      if (one === '') continue
      checked += 1
      const ok = allowed.some((prefix) => one.startsWith(prefix)) || one === 'to' || one === 'from'
      assert(ok, `unprefixed selector "${one}"`)
    }
  }
  assert(checked > 20, `only ${checked} selectors inspected`)
})

check('uses only --dsw-alias-* tokens or --branddeck-* locals', () => {
  const refs = [...source.matchAll(/var\(\s*(--[A-Za-z0-9-]+)/g)].map((match) => match[1])
  assert(refs.length > 0, 'no custom properties referenced')
  for (const name of refs) {
    if (name.startsWith('--branddeck-')) continue
    assert(name.startsWith('--dsw-alias-'), `non-alias custom property "${name}"`)
  }
  assert(!/--dsw-(?!alias-)/.test(source), 'non-alias --dsw-* property used')
})

check('every --dsw-alias-* token is a real theme token', () => {
  const used = new Set([...source.matchAll(/--dsw-alias-[a-zA-Z0-9-]+/g)].map((match) => match[0]))
  assert(used.size > 0, 'no alias tokens used')
  if (!existsSync(THEME_TOKENS_PATH)) {
    console.log('      (theme package not found - token existence check skipped)')
    return
  }
  const theme = readFileSync(THEME_TOKENS_PATH, 'utf8')
  const real = new Set([...theme.matchAll(/--dsw-alias-[a-zA-Z0-9-]+/g)].map((match) => match[0]))
  for (const token of used) assert(real.has(token), `unknown theme token "${token}"`)
})

check('avoid !important overrides', () => {
  assert(!source.includes('!important'), '!important found')
})

/* ------------------------------------------------------------------ *
 * 6. Locale dictionaries
 * ------------------------------------------------------------------ */

const zhStart = source.indexOf('const ZH = {')
const enStart = source.indexOf('const EN = {')
const zhSlice = zhStart === -1 || enStart === -1 ? '' : source.slice(zhStart, enStart)
const enSlice = enStart === -1 ? '' : source.slice(enStart, source.indexOf('const DICT = {', enStart))

check('defines the zh and en dictionaries', () => {
  assert(zhStart !== -1, 'const ZH missing')
  assert(enStart !== -1, 'const EN missing')
})

check('keeps every Chinese string inside the zh dictionary', () => {
  const literals = cjkLiterals(source)
  assert(literals.length > 10, `only ${literals.length} Chinese literals found`)
  for (const literal of literals) assert(zhSlice.includes(literal), `Chinese literal outside ZH: ${literal}`)
})

check('keeps zh and en key sets identical', () => {
  const zhKeys = dictKeys(zhSlice)
  const enKeys = dictKeys(enSlice)
  assert(zhKeys.size > 20, `only ${zhKeys.size} zh keys`)
  for (const key of zhKeys) assert(enKeys.has(key), `key "${key}" missing from EN`)
  for (const key of enKeys) assert(zhKeys.has(key), `key "${key}" missing from ZH`)
})

check('every t() key resolves in both dictionaries', () => {
  const zhKeys = dictKeys(zhSlice)
  const enKeys = dictKeys(enSlice)
  const literalKeys = [...source.matchAll(/\bt\('([^']+)'/g)].map((match) => match[1])
  const keys = new Set(literalKeys)
  /* Composite keys built with a template literal in the render path. */
  for (const state of ['idle', 'saving', 'saved', 'dirty', 'error']) keys.add(`save.${state}`)
  for (const align of ['left', 'center', 'right']) keys.add(`align.${align}`)
  for (const key of keys) {
    assert(zhKeys.has(key) || enKeys.has(key), `unresolved translation key "${key}"`)
  }
  for (const align of ['left', 'center', 'right']) assert(zhKeys.has(`align.${align}`), `align.${align} missing`)
})

check('carries no dead dictionary keys', () => {
  const zhKeys = dictKeys(zhSlice)
  const literalKeys = new Set([...source.matchAll(/\bt\('([^']+)'/g)].map((match) => match[1]))
  /* Keys built with template literals in the render path. */
  const dynamic = new Set([
    'save.idle',
    'save.saving',
    'save.saved',
    'save.dirty',
    'save.error',
    'align.left',
    'align.center',
    'align.right',
  ])
  for (const key of zhKeys) {
    assert(literalKeys.has(key) || dynamic.has(key), `unused dictionary key "${key}"`)
  }
})

check('dispatches every reducer action it defines', () => {
  const defined = new Set([...source.matchAll(/case '([a-z]+\/[a-z]+)':/g)].map((match) => match[1]))
  const dispatched = new Set([...source.matchAll(/type: '([a-z]+\/[a-z]+)'/g)].map((match) => match[1]))
  assert(defined.size > 15, `only ${defined.size} reducer actions found`)
  for (const action of defined) assert(dispatched.has(action), `reducer action "${action}" is never dispatched`)
  for (const action of dispatched) assert(defined.has(action), `dispatch of unknown action "${action}"`)
})

/* ------------------------------------------------------------------ *
 * 7. Behaviour contract markers
 * ------------------------------------------------------------------ */

check('keeps the shared canvas contract constants', () => {
  assert(/const CANVAS = Object\.freeze\(\{ width: 1280, height: 720 \}\)/.test(source), 'CANVAS mismatch')
  assert(/const SAFE_AREA = Object\.freeze\(\{ top: 64, right: 64, bottom: 64, left: 64 \}\)/.test(source), 'SAFE_AREA mismatch')
  assert(/LOGO_RESERVE = Object\.freeze\(\{ width: 150, height: 48, gap: 16 \}\)/.test(source), 'LOGO_RESERVE mismatch')
})

check('scales the 16:9 canvas with transform: scale()', () => {
  assert(/Math\.floor\(\(width \/ CANVAS\.width\) \* 100\) \/ 100/.test(source), 'scale formula missing')
  assert(source.includes('transform: `scale(${scale})`'), 'transform: scale not applied')
  assert(source.includes('ResizeObserver'), 'no ResizeObserver for the container width')
})

check('debounces autosave by 600ms', () => {
  assert(/AUTOSAVE_DEBOUNCE_MS = 600/.test(source), 'debounce constant mismatch')
  assert(/setTimeout\(\(\) => \{\s*\n\s*void saveDeck\(/.test(source), 'autosave timer does not call saveDeck')
})

check('implements the full interaction set', () => {
  const markers = [
    'onDoubleClick: (event) => props.onDoubleClick(event, element)',
    'contentEditable: true',
    'setPointerCapture',
    "case 'element/resize'",
    'NUDGE_STEP_FAST',
    "event.key === 'Delete'",
    "event.key === 'd'",
    "event.key === 's'",
    "event.key === 'Escape'",
    'root.contains(active)',
  ]
  for (const marker of markers) assert(source.includes(marker), `missing behaviour marker: ${marker}`)
})

check('renders the required states', () => {
  const markers = ['branddeck-empty', 'branddeck-spinner', 'branddeck-banner-error', 'branddeck-overflow-badge', 'branddeck-inspector']
  for (const marker of markers) assert(source.includes(marker), `missing UI state: ${marker}`)
  assert(source.includes("background:var(--dsw-alias-bg-mask-1"), 'modal mask missing')
})

/* ------------------------------------------------------------------ *
 * 8. Stub runtime: factory + apply()
 * ------------------------------------------------------------------ */

function createFakeReact(options) {
  const settings = options ?? {}
  class Component {
    constructor(props) {
      this.props = props
    }

    setState(next) {
      this.state = { ...this.state, ...(typeof next === 'function' ? next(this.state) : next) }
    }
  }
  // React semantics: children land in props.children (single child unwrapped).
  const createElement = (type, props, ...children) => {
    const merged = { ...(props ?? {}) }
    if (children.length === 1) merged.children = children[0]
    else if (children.length > 1) merged.children = children
    return { type, props: merged, children }
  }
  return {
    createElement,
    Component,
    Fragment: 'Fragment',
    useState: (initial) => [typeof initial === 'function' ? initial() : initial, () => {}],
    useReducer: (reducer, initialArg, init) => [
      /* `settings.state` seeds the panel reducer so a populated deck can be
       * rendered without mounting React. */
      settings.state ?? (typeof init === 'function' ? init(initialArg) : initialArg),
      () => {},
    ],
    useEffect: () => undefined,
    useLayoutEffect: () => undefined,
    useRef: (initial) => ({ current: initial }),
    useMemo: (factory) => factory(),
    useCallback: (callback) => callback,
  }
}

/** Evaluate the plugin file in a bare sandbox and return the captured definition. */
function loadDefinition() {
  const sandbox = { window: {}, console, setTimeout, clearTimeout }
  sandbox.window.__ModuleLoader__ = { load: (value) => (sandbox.definition = value) }
  vm.runInNewContext(source, sandbox)
  return sandbox.definition
}

const moduleDefinition = loadDefinition()

check('evaluates the definition in a bare sandbox', () => {
  assert(moduleDefinition !== undefined, 'loader did not capture a definition')
  assert(moduleDefinition.id === '@local/dsh-branddeck', `unexpected id ${String(moduleDefinition.id)}`)
  assert(typeof moduleDefinition.factory === 'function', 'factory is not a function')
})

const React = createFakeReact()
const required = []
const moduleExports = moduleDefinition.factory((name) => {
  required.push(name)
  if (name === 'react') return React
  throw new Error(`unexpected require(${name})`)
})

check('requires only react', () => {
  assert(required.length === 1 && required[0] === 'react', `requires: ${required.join(', ')}`)
})

/** Build a stub Cordis context that records slot registrations. */
function createStubContext(options) {
  const settings = options ?? {}
  const registrations = []
  const effects = []
  const ctx = {
    slots: {
      inject(name, factory) {
        factory()
        return { dispose() {} }
      },
      register(opts, component) {
        registrations.push({ opts, component })
        return { dispose() {} }
      },
    },
    effect(factory, label) {
      effects.push(label)
      const dispose = factory()
      return () => {
        if (typeof dispose === 'function') dispose()
      }
    },
  }
  if (settings.locale !== false) {
    ctx.locale = {
      register() {
        if (settings.localeRegisterThrows === true) throw new Error('namespace already registered')
        return () => {}
      },
      bind() {
        return (key) => key
      },
      getLocale() {
        return { active: settings.localeId ?? 'zh' }
      },
    }
  }
  if (settings.connection !== false) {
    ctx.connection = {
      rpc: {
        async call() {
          return { ok: true, value: {} }
        },
      },
    }
  }
  return { ctx, registrations, effects }
}

check('declares slots as the only hard dependency', () => {
  assert(Array.isArray(moduleExports.inject), 'inject is not an array')
  assert(moduleExports.inject.includes('slots'), "'slots' not injected")
  assert(!moduleExports.inject.includes('connection'), "'connection' must not be a hard inject")
  assert(typeof moduleExports.apply === 'function', 'apply is not a function')
})

const stub = createStubContext({ localeId: 'zh' })

check('apply() registers exactly the two slots without throwing', () => {
  moduleExports.apply(stub.ctx)
  assert(stub.registrations.length === 2, `expected 2 registrations, found ${stub.registrations.length}`)
  const main = stub.registrations.filter((entry) => entry.opts.name === 'main')
  const sidebar = stub.registrations.filter((entry) => entry.opts.name === 'sidebar.panellist')
  assert(main.length === 1, `expected 1 main registration, found ${main.length}`)
  assert(sidebar.length === 1, `expected 1 sidebar registration, found ${sidebar.length}`)
  assert(main[0].opts.key === 'branddeck', 'main key is not branddeck')
  assert(main[0].opts.locale === 'branddeck', 'main locale is not the namespace')
  assert(main[0].opts.inject === undefined, 'main must render in root scope (no injected props)')
  assert(typeof main[0].component === 'function', 'main component is not a function')
  assert(sidebar[0].opts.id === 'branddeck', 'sidebar id is not branddeck')
  assert(sidebar[0].opts.order === 20, `sidebar order is ${String(sidebar[0].opts.order)}`)
  assert(sidebar[0].opts.locale === 'branddeck', 'sidebar locale is not the namespace')
  assert(typeof sidebar[0].component === 'function', 'sidebar component is not a function')
})

check('sidebar label resolves through the zh dictionary', () => {
  const sidebar = stub.registrations.find((entry) => entry.opts.name === 'sidebar.panellist')
  const label = sidebar.opts.label()
  assert(typeof label === 'string' && label.length > 0, 'label is empty')
  assert(/品牌/.test(label), `label did not resolve to Chinese: ${label}`)
})

check('renders the sidebar icon from a single svg with currentColor', () => {
  const sidebar = stub.registrations.find((entry) => entry.opts.name === 'sidebar.panellist')
  const tree = sidebar.component({ size: 18 })
  assert(tree !== null && typeof tree === 'object', 'icon returned nothing')
  assert(tree.type === 'svg', `icon root is ${String(tree.type)}`)
  const children = Array.isArray(tree.children) ? tree.children : []
  assert(children.length >= 2, 'icon draws a single shape')
  const usesCurrentColor = children.some((child) => child.props.stroke === 'currentColor')
  assert(usesCurrentColor, 'icon does not use currentColor')
})

check('renders the main page root without props', () => {
  const main = stub.registrations.find((entry) => entry.opts.name === 'main')
  const tree = main.component()
  assert(tree !== null && typeof tree === 'object', 'page returned nothing')
  assert(typeof tree.type === 'function', 'page root is not the error boundary component')
  assert(tree.props.children !== undefined, 'error boundary has no child')
})

check('apply() survives a missing locale service', () => {
  const bare = createStubContext({ locale: false })
  moduleExports.apply(bare.ctx)
  assert(bare.registrations.length === 2, 'slots not registered without locale')
  const sidebar = bare.registrations.find((entry) => entry.opts.name === 'sidebar.panellist')
  assert(typeof sidebar.opts.label() === 'string', 'label failed without locale')
})

check('apply() survives a throwing locale registry', () => {
  const hostile = createStubContext({ localeRegisterThrows: true })
  moduleExports.apply(hostile.ctx)
  assert(hostile.registrations.length === 2, 'slots not registered when register() throws')
})

check('apply() survives a missing connection service', () => {
  const noTransport = createStubContext({ connection: false })
  moduleExports.apply(noTransport.ctx)
  assert(noTransport.registrations.length === 2, 'slots not registered without connection')
})

/* ------------------------------------------------------------------ *
 * 9. Shallow render of the panel body (fake hooks, no DOM)
 * ------------------------------------------------------------------ */

/**
 * Walk a fake-React tree, shallow-rendering function components so their host
 * elements (and therefore their classNames) are visible to the visitor.
 */
function walk(node, visit, depth = 0) {
  if (node === null || node === undefined || typeof node !== 'object' || depth > 80) return
  if (Array.isArray(node)) {
    for (const child of node) walk(child, visit, depth + 1)
    return
  }
  if (typeof node.type === 'function') {
    let output = null
    try {
      output = node.type(node.props)
    } catch (error) {
      /* Class components are functions too but refuse a plain call. */
      const instance = new node.type(node.props)
      output = instance.render()
    }
    walk(output, visit, depth + 1)
    return
  }
  if (node.type !== undefined) visit(node)
  if (node.props !== undefined) walk(node.props.children, visit, depth + 1)
}

/** Render the registered main component down to its element tree. */
function renderMain(seedState) {
  const react = createFakeReact(seedState === undefined ? {} : { state: seedState })
  const plugin = moduleDefinition.factory((name) => {
    if (name === 'react') return react
    throw new Error(`unexpected require(${name})`)
  })
  const stub = createStubContext({ localeId: 'zh' })
  plugin.apply(stub.ctx)
  const main = stub.registrations.find((entry) => entry.opts.name === 'main')
  const boundaryElement = main.component()
  /* The boundary class is the slot root; render it to reach the panel body. */
  const boundary = new boundaryElement.type(boundaryElement.props)
  const pageElement = boundary.render()
  assert(pageElement !== null && typeof pageElement === 'object', 'panel body element missing')
  return pageElement.type()
}

check('renders the loading panel without a deck', () => {
  const tree = renderMain()
  assert(tree !== null && typeof tree === 'object', 'no tree')
  const classes = []
  walk(tree, (node) => {
    if (typeof node.props.className === 'string') classes.push(node.props.className)
  })
  assert(classes.includes('branddeck-root'), 'root class missing')
  assert(classes.includes('branddeck-toolbar'), 'toolbar missing')
  assert(classes.some((name) => name.includes('branddeck-loading')), 'loading state missing')
  assert(classes.some((name) => name.startsWith('branddeck-inspector')), 'inspector missing')
})

check('renders a populated deck and keeps every className branded', () => {
  const theme = {
    templateId: 'tpl',
    templateName: 'Template',
    colors: {
      background: '#FFFFFF',
      surface: '#F4F6F9',
      textPrimary: '#1B1F24',
      textSecondary: '#5B6470',
      primary: '#247BBF',
      accent: '#F2A03D',
      chartPalette: ['#247BBF', '#F2A03D'],
    },
    fonts: { title: 'Title Font', body: 'Body Font', latin: 'Arial' },
    brandAssets: { icons: {} },
  }
  const slide = {
    slideId: 's1',
    background: '$colors.background',
    elements: [
      {
        type: 'text',
        id: 'title',
        role: 'title',
        content: 'Quarterly review with a deliberately long headline that overflows',
        x: 64,
        y: 64,
        w: 200,
        h: 40,
        style: { fontSize: 32, color: '$colors.textPrimary', align: 'left' },
      },
      { type: 'shape', id: 'box', shapeType: 'roundRect', x: 64, y: 200, w: 300, h: 120, style: { fill: '$colors.primary', radius: 12 } },
      { type: 'shape', id: 'arrow', shapeType: 'rightArrow', x: 500, y: 200, w: 120, h: 60, style: { fill: '#F2A03D' } },
      { type: 'shape', id: 'rule', shapeType: 'line', x: 64, y: 400, w: 400, h: 8, style: { stroke: '$colors.accent', strokeWidth: 3 } },
      {
        type: 'chart',
        id: 'bars',
        chartType: 'bar',
        x: 700,
        y: 200,
        w: 480,
        h: 240,
        data: { labels: ['A', 'B', 'C'], series: [{ name: 'S', values: [3, 7, 5] }] },
      },
      {
        type: 'chart',
        id: 'pie',
        chartType: 'pie',
        x: 700,
        y: 470,
        w: 200,
        h: 180,
        data: { labels: ['A', 'B'], series: [{ name: 'S', values: [1, 1] }] },
      },
      { type: 'svg', id: 'diagram', x: 64, y: 450, w: 240, h: 180, svgCode: '<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"/></svg>' },
      { type: 'image', id: 'logo', assetKey: 'logo', x: 1100, y: 64, w: 150, h: 48 },
    ],
  }
  const state = {
    status: 'ready',
    error: 'stale transport message',
    decks: [{ deckId: 'd1', title: 'Deck', templateId: 'tpl', slideCount: 2, updatedAt: '' }],
    templates: [],
    deck: { deckId: 'd1', title: 'Deck', templateId: 'tpl', slides: [slide, { slideId: 's2', elements: [slide.elements[0]] }] },
    theme,
    logo: { templateId: 'tpl', dataUrl: 'data:image/png;base64,AAAA' },
    activeDeckId: 'd1',
    warnings: ['slide 1 · text "title" overflows its box'],
    warningsDismissed: false,
    slideIndex: 0,
    selectionId: 'title',
    editingId: null,
    dirty: true,
    rev: 4,
    saveState: 'dirty',
    saveError: 'save blew up',
    exporting: true,
    exportResult: { ok: true, path: 'out/deck.pptx', slideCount: 2 },
  }
  const tree = renderMain(state)
  const classes = []
  walk(tree, (node) => {
    if (typeof node.props.className === 'string') classes.push(node.props.className)
  })
  for (const name of classes) {
    for (const single of name.split(/\s+/)) {
      assert(single.startsWith('branddeck-'), `unprefixed className "${single}"`)
    }
  }
  const expected = [
    'branddeck-canvas',
    'branddeck-frame',
    'branddeck-thumb',
    'branddeck-overflow-badge',
    'branddeck-inspector',
    'branddeck-banner-warn',
    'branddeck-banner-error',
    'branddeck-banner-success',
    'branddeck-modal-mask',
    'branddeck-handle',
  ]
  for (const name of expected) {
    assert(classes.some((value) => value.includes(name)), `missing rendered element: ${name}`)
  }
  /* The selected element is outlined, overflow is flagged, and the modal shows. */
  const selected = classes.filter((value) => value.includes('branddeck-el-selected'))
  assert(selected.length === 1, `expected 1 selected element, found ${selected.length}`)
  const overflowing = classes.filter((value) => value.includes('branddeck-el-overflow'))
  assert(overflowing.length >= 1, 'overflowing text element not flagged')
})

/* ------------------------------------------------------------------ *
 * Summary
 * ------------------------------------------------------------------ */

console.log('')
console.log(`${passed} passed, ${failures} failed`)
if (failures > 0) {
  console.log('SMOKE FAIL')
  process.exit(1)
}
console.log('SMOKE PASS')
