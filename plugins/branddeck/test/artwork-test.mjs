/**
 * Tests for the template-derived visual system: backdrop art, the content frame, footlines,
 * and the page→master mapping that puts the art on covers/section breaks.
 *
 * WHY THESE EXIST: the deck used to be drawn on a plain white canvas with our own 64px
 * margins, because three separate whitelists (template → store → generate/compile) each
 * rebuilt `brandAssets` and dropped whatever the template had given us. These tests pin the
 * artefacts and — more importantly — pin that they survive the trip to the compiler.
 *
 * Run: node test/artwork-test.mjs
 */
import { readFileSync } from 'node:fs'
import { loadPackage } from '../src/pptx/read.js'
import { discoverStructure } from '../src/pptx/theme.js'
import { readTemplateArtwork } from '../src/pptx/artwork.js'
import { innerBox } from '../src/schema.js'
import { VARIANTS, masterFor } from '../src/compile.js'

const TEMPLATE = 'D:/Projects/ai/BrandDeck/PPT模板-HW.pptx'

const results = []
async function test(name, fn) {
  try {
    const detail = await fn()
    results.push({ name, ok: true })
    console.log(`PASS ${name}${detail === undefined ? '' : ` — ${detail}`}`)
  } catch (error) {
    results.push({ name, ok: false })
    console.log(`FAIL ${name}: ${error?.message ?? error}`)
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message)
}

const stable = (value) => JSON.stringify(value)

let pkg = null
let structure = null
let artwork = null
try {
  pkg = await loadPackage({ path: TEMPLATE })
  structure = discoverStructure(pkg)
  artwork = readTemplateArtwork(pkg, structure)
} catch (error) {
  console.log(`(template unavailable: ${error?.message ?? error})`)
}

/* ------------------------------------------------------------------ *
 * Backdrop art
 * ------------------------------------------------------------------ */

await test('1a · full-bleed layout pictures are found and mapped to their layout', () => {
  assert(artwork !== null, 'template not loaded')
  assert(artwork.backdrops.length >= 4, `expected >=4 backdrops, got ${artwork.backdrops.length}`)
  for (const backdrop of artwork.backdrops) {
    assert(/^ppt\/media\//.test(backdrop.mediaPart), `media part looks wrong: ${backdrop.mediaPart}`)
    assert(/^ppt\/slideLayouts\//.test(backdrop.layoutPart), `layout part looks wrong: ${backdrop.layoutPart}`)
    // The Huawei art is 13.34×6.13in on a 13.34×7.5in slide: full width, 82% height.
    assert(backdrop.fillRatio > 0.6, `fill ratio ${backdrop.fillRatio} is not full-bleed-ish`)
  }
  return `${artwork.backdrops.length} backdrops, fill ${artwork.backdrops.map((b) => b.fillRatio).join('/')}`
})

await test('1b · each backdrop carries the template frame, not a stretched full bleed', () => {
  const first = artwork.backdrops[0]
  const frame = first.frameInches
  assert(frame !== undefined && frame !== null, 'backdrop has no frameInches')
  assert(frame.w > 13 && frame.w < 13.4, `width ${frame.w} should be ~13.34in (the slide width)`)
  // 6.13–6.17in of a 7.5in slide: the footer band is deliberately left uncovered.
  assert(frame.h > 5.9 && frame.h < 6.4, `height ${frame.h} should be ~6.1in, not 7.5in`)
  return `${frame.w}×${frame.h}in at (${frame.x}, ${frame.y})`
})

/* ------------------------------------------------------------------ *
 * Content frame
 * ------------------------------------------------------------------ */

await test('2a · the content frame comes from the template\'s own content page', () => {
  const frame = artwork.frame
  assert(frame !== null, 'no content frame derived')
  // 0.80in → 77px, 0.50in → 48px, width 11.75in → ~1127px, bottom 6.77in → ~650px.
  assert(Math.abs(frame.x - 77) <= 2, `x ${frame.x} should be ~77 (0.80in)`)
  assert(Math.abs(frame.y - 48) <= 2, `y ${frame.y} should be ~48 (0.50in)`)
  assert(Math.abs(frame.w - 1127) <= 4, `w ${frame.w} should be ~1127 (11.75in)`)
  assert(Math.abs(frame.h - 602) <= 6, `h ${frame.h} should be ~602 (title top to body bottom)`)
  return `${frame.x},${frame.y} ${frame.w}×${frame.h}px`
})

await test('2b · innerBox prefers the template frame and clamps a hostile one', () => {
  const theme = { brandAssets: { frame: { x: 77, y: 48, w: 1127, h: 602 } } }
  const box = innerBox(/** @type {any} */ (theme))
  assert(box.x === 77 && box.y === 48, `frame not honoured: ${stable(box)}`)
  assert(box.w === 1127 && box.h === 602, `frame size not honoured: ${stable(box)}`)
  // A frame larger than the canvas is clamped to the sanity floor, never used raw.
  const huge = innerBox(/** @type {any} */ ({ brandAssets: { frame: { x: -50, y: -50, w: 5000, h: 5000 } } }))
  assert(huge.x >= 24 && huge.y >= 24, `negative origin not clamped: ${stable(huge)}`)
  assert(huge.x + huge.w <= 1280 && huge.y + huge.h <= 720, `frame escaped the canvas: ${stable(huge)}`)
  // No frame ⇒ the built-in safe area, unchanged.
  const none = innerBox(/** @type {any} */ ({ brandAssets: {} }))
  assert(none.x === 64 && none.y === 64, `fallback changed: ${stable(none)}`)
  return `frame honoured, hostile frame clamped, fallback ${none.x}/${none.y}`
})

/* ------------------------------------------------------------------ *
 * Footlines
 * ------------------------------------------------------------------ */

await test('3a · compliance footlines are inherited, design annotations are not', () => {
  const lines = artwork.footlines
  assert(lines.length > 0, 'no footlines found')
  assert(lines.every((line) => line.length <= 48), `a footline is suspiciously long: ${stable(lines)}`)
  const joined = lines.join(' | ')
  assert(/Security Level|Confidential|保密|机密/i.test(joined), `no compliance marker in ${stable(lines)}`)
  // "PANTONE186CRGB 200/16/46" sits on the master as a colour spec, not a footer.
  assert(!/PANTONE|RGB\s*\d/i.test(joined), `a design annotation leaked into the footer: ${stable(lines)}`)
  return stable(lines)
})

/* ------------------------------------------------------------------ *
 * Page → master mapping
 * ------------------------------------------------------------------ */

await test('4a · art goes on covers and section breaks, never on content pages', () => {
  const has = { hasBackdrops: true }
  assert(masterFor('cover', false, has) === VARIANTS.cover, 'cover did not get the cover art')
  assert(masterFor('section', false, has) === VARIANTS.section, 'section break did not get art')
  assert(masterFor('bullets', false, has) === VARIANTS.content, 'a content page got art')
  assert(masterFor('toc', false, has) === VARIANTS.content, 'the TOC got art')
  assert(masterFor('stats', false, has) === VARIANTS.content, 'a stats page got art')
  // The last `cover` is the closing page (谢谢), which is drawn as kind 'cover' too.
  assert(masterFor('cover', true, has) === VARIANTS.closing, 'closing page did not get its own art')
  return 'cover→cover art, section→section art, content→content master'
})

await test('4b · a template without art keeps every page on the content master', () => {
  const noArt = { hasBackdrops: false }
  for (const kind of ['cover', 'section', 'bullets', 'toc', 'stats', 'diagram']) {
    assert(masterFor(kind, false, noArt) === VARIANTS.content, `${kind} left the content master`)
    assert(masterFor(kind, true, noArt) === VARIANTS.content, `${kind} (last) left the content master`)
  }
  return 'no art ⇒ content master everywhere (footer intact)'
})

/* ------------------------------------------------------------------ *
 * Summary
 * ------------------------------------------------------------------ */

const failed = results.filter((result) => !result.ok)
console.log(`\n${results.length} checks, ${results.length - failed.length} passed, ${failed.length} failed`)
if (failed.length > 0) {
  console.log(`  FAILED: ${failed.map((result) => result.name).join('; ')}`)
  process.exitCode = 1
} else {
  console.log('ALL PASS')
}
