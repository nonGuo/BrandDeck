/**
 * Tiny test harness shared by `test/engine-*.mjs`.
 *
 * Plain Node ESM, no test framework: each test file imports `test()` / `run()`, runs
 * real files (never mocks the corporate template) and exits with code 1 on failure.
 *
 * Running `node test/engine-helpers.mjs` on its own self-checks the harness.
 */
import { mkdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { randomBytes } from 'node:crypto'

export const TEST_DIR = dirname(fileURLToPath(import.meta.url))
export const PACKAGE_DIR = resolve(TEST_DIR, '..')
export const WORKSPACE_DIR = resolve(PACKAGE_DIR, '..', '..')

/** The real corporate template used by every integration test. */
export const TEMPLATE_PATH = join(WORKSPACE_DIR, 'PPT模板-HW.pptx')

/** Cross-check fixture produced by the parallel analysis teammate. */
export const EXPECTED_TOKENS_PATH = join(WORKSPACE_DIR, 'research', 'expected-tokens.json')

/** All test artefacts stay inside `plugins/branddeck/.tmp`. */
export const TMP_DIR = join(PACKAGE_DIR, '.tmp')
export const OUT_DIR = join(TMP_DIR, 'engine-out')

/** A unique temp directory per test run, inside `.tmp`. */
export function freshTempDir(prefix = 'run') {
  const dir = join(TMP_DIR, `${prefix}-${randomBytes(4).toString('hex')}`)
  mkdirSync(dir, { recursive: true })
  return dir
}

const results = []

/**
 * Register and run one asynchronous test.
 * @param {string} name
 * @param {() => any | Promise<any>} fn
 */
export async function test(name, fn) {
  const started = Date.now()
  try {
    const detail = await fn()
    results.push({ name, ok: true })
    const suffix = detail === undefined ? '' : ` ${formatDetail(detail)}`
    console.log(`PASS ${name} (${Date.now() - started}ms)${suffix}`)
  } catch (error) {
    results.push({ name, ok: false, error })
    console.log(`FAIL ${name} (${Date.now() - started}ms)`)
    console.log(indent(error?.stack ?? String(error)))
  }
}

/**
 * Print the summary and set the process exit code.
 * @param {string} title
 */
export function run(title) {
  const failed = results.filter((entry) => !entry.ok)
  console.log(`\n${title}: ${results.length - failed.length}/${results.length} checks passed`)
  if (failed.length > 0) {
    console.log(`failed: ${failed.map((entry) => entry.name).join(', ')}`)
    process.exitCode = 1
  }
}

/* ------------------------------------------------------------------ *
 * Assertions
 * ------------------------------------------------------------------ */

/**
 * @param {unknown} condition
 * @param {string} message
 */
export function assert(condition, message) {
  if (!condition) throw new Error(`assertion failed: ${message}`)
}

/**
 * @param {unknown} actual
 * @param {unknown} expected
 * @param {string} message
 */
export function assertEqual(actual, expected, message) {
  const a = JSON.stringify(actual)
  const b = JSON.stringify(expected)
  if (a !== b) throw new Error(`assertion failed: ${message}\n  expected: ${b}\n  actual:   ${a}`)
}

/**
 * @param {string} value
 * @param {RegExp} pattern
 * @param {string} message
 */
export function assertMatch(value, pattern, message) {
  if (!pattern.test(String(value))) {
    throw new Error(`assertion failed: ${message}\n  ${pattern} does not match ${JSON.stringify(value)}`)
  }
}

/**
 * Assert that `fn` throws an error carrying `code`.
 * @param {() => any | Promise<any>} fn
 * @param {string} code
 * @param {string} message
 */
export async function assertCode(fn, code, message) {
  try {
    await fn()
  } catch (error) {
    if (/** @type {any} */ (error)?.code !== code) {
      throw new Error(
        `assertion failed: ${message}\n  expected code ${code}, got ${/** @type {any} */ (error)?.code}: ${error?.message}`,
      )
    }
    return error
  }
  throw new Error(`assertion failed: ${message}\n  expected a "${code}" error, nothing was thrown`)
}

/**
 * Print a field-by-field comparison table and return the mismatching rows.
 * @param {{field: string, expected: unknown, actual: unknown}[]} rows
 * @returns {{field: string, expected: unknown, actual: unknown}[]}
 */
export function diffTable(rows) {
  const width = Math.max(...rows.map((row) => row.field.length), 5)
  const mismatches = []
  for (const row of rows) {
    const same = JSON.stringify(row.expected) === JSON.stringify(row.actual)
    if (!same) mismatches.push(row)
    console.log(
      `  ${same ? 'ok  ' : 'DIFF'} ${row.field.padEnd(width)}  expected=${formatCell(row.expected)}  actual=${formatCell(row.actual)}`,
    )
  }
  return mismatches
}

/* ------------------------------------------------------------------ *
 * Internals
 * ------------------------------------------------------------------ */

/** @param {unknown} value */
function formatCell(value) {
  if (Array.isArray(value)) return `[${value.join(',')}]`
  if (value !== null && typeof value === 'object') return JSON.stringify(value)
  return String(value)
}

/** @param {unknown} value */
function formatDetail(value) {
  if (value === null || typeof value !== 'object') return String(value)
  return JSON.stringify(value)
}

/** @param {string} text */
function indent(text) {
  return String(text)
    .split('\n')
    .map((line) => `     ${line}`)
    .join('\n')
}

// Self-check when executed directly (`node test/engine-helpers.mjs`).
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await test('helper: assert/assertEqual behave', () => {
    assert(true, 'true is truthy')
    assertEqual({ a: [1, 2] }, { a: [1, 2] }, 'deep equality via JSON')
    return 'ok'
  })
  await test('helper: assertCode matches error codes', async () => {
    const error = new Error('boom')
    // @ts-expect-error - attaching a code on purpose
    error.code = 'boom_code'
    await assertCode(() => {
      throw error
    }, 'boom_code', 'code matches')
    return 'ok'
  })
  await test('helper: paths resolve', () => {
    assert(TEST_DIR.length > 0, 'TEST_DIR')
    assert(PACKAGE_DIR.endsWith('branddeck'), 'PACKAGE_DIR ends with branddeck')
    assert(TMP_DIR.includes('.tmp'), 'TMP_DIR inside .tmp')
    return PACKAGE_DIR
  })
  run('engine-helpers self-test')
}
