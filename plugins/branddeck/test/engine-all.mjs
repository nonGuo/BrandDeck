/**
 * Runs every `test/engine-*.mjs` file in its own Node process and reports a summary.
 *
 * `node test/engine-*.mjs` relies on the shell globbing the pattern; PowerShell does not
 * expand globs for native commands, so this runner is the portable one-liner:
 *
 *   node test/engine-all.mjs
 *
 * Run: node test/engine-all.mjs   (or: node test/engine-all.mjs read theme)
 */
import { spawnSync } from 'node:child_process'
import { readdirSync } from 'node:fs'
import { join } from 'node:path'
import { TEST_DIR } from './engine-helpers.mjs'

const SELF = 'engine-all.mjs'
const filters = process.argv.slice(2)

const files = readdirSync(TEST_DIR)
  .filter((name) => name.startsWith('engine-') && name.endsWith('.mjs') && name !== SELF)
  .filter((name) => filters.length === 0 || filters.some((filter) => name.includes(filter)))
  .sort()

if (files.length === 0) {
  console.error(`no engine-*.mjs tests matched ${JSON.stringify(filters)}`)
  process.exitCode = 1
} else {
  /** @type {{name: string, code: number | null, ms: number}[]} */
  const results = []
  for (const file of files) {
    const started = Date.now()
    console.log(`\n──────── ${file} ────────`)
    const result = spawnSync(process.execPath, [join(TEST_DIR, file)], {
      stdio: 'inherit',
      cwd: join(TEST_DIR, '..'),
    })
    results.push({ name: file, code: result.status, ms: Date.now() - started })
  }

  const failed = results.filter((entry) => entry.code !== 0)
  console.log('\n════════ engine test summary ════════')
  for (const entry of results) {
    console.log(`${entry.code === 0 ? 'PASS' : 'FAIL'}  ${entry.name.padEnd(22)} ${String(entry.ms).padStart(6)}ms`)
  }
  console.log(`\n${results.length - failed.length}/${results.length} test files passed`)
  if (failed.length > 0) {
    console.log(`failed: ${failed.map((entry) => entry.name).join(', ')}`)
    process.exitCode = 1
  }
}
