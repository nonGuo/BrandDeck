#!/usr/bin/env node
// Usage: node scripts/probe.mjs <file> <regex> [before] [after] [maxMatches]
import { readFileSync } from 'node:fs'

const [file, pattern, beforeStr = '200', afterStr = '400', maxStr = '5'] = process.argv.slice(2)
const before = Number(beforeStr)
const after = Number(afterStr)
const max = Number(maxStr)
const text = readFileSync(file, 'utf8')
const re = new RegExp(pattern, 'g')
let m
let n = 0
while ((m = re.exec(text)) !== null && n < max) {
  n++
  const start = Math.max(0, m.index - before)
  const end = Math.min(text.length, m.index + m[0].length + after)
  console.log(`--- match ${n} @${m.index} ---`)
  console.log(text.slice(start, end).replace(/\n/g, '\\n'))
  if (m.index === re.lastIndex) re.lastIndex++
}
if (n === 0) console.log('no match')
