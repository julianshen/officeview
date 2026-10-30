/**
 * Captures Excel for Mac's own rendering of every sparkline case (spec section 8) through AppleScript
 * `copy picture … appearance screen format bitmap`, read back from the pasteboard as PNG.
 *
 * Probe findings (2026-10-27, Excel for Mac): the bitmap holds the sparklines at a fixed 1 px per point (72 dpi)
 * whatever the sheet zoom, and the pasteboard's PDF only wraps the same bitmaps. `open` returns before the
 * workbook is registered, so each capture waits for the expected workbook name.
 *
 *   bun scripts/sparkline-refs/capture.ts [--stage N] [--case id]
 *
 * Writes tests/fixtures/xlsx-sparklines/refs/<id>[-bordered][-<sheet>].png plus a .json sidecar.
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { capture } from '../xlsx-refs/capture'
import { osa } from '../xlsx-refs/excel'
import { SPARKLINE_CASES, captureRange } from './cases'

// The capture itself moved to the shared harness (scripts/xlsx-refs/capture.ts); re-exported for existing callers.
export { capture }

const FIXTURES = resolve(import.meta.dir, '../../tests/fixtures/xlsx-sparklines')
const REFS = join(FIXTURES, 'refs')

if (import.meta.main) {
  const args = process.argv.slice(2)
  const stage = Number(args[args.indexOf('--stage') + 1] || 3)
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  mkdirSync(REFS, { recursive: true })
  const excelVersion = osa(['tell application "Microsoft Excel" to get version'])
  let count = 0
  for (const c of SPARKLINE_CASES) {
    if (c.stage > stage || (only && c.id !== only)) continue
    for (const variant of ['', '-bordered']) {
      for (const sheet of c.sheets) {
        const range = captureRange(sheet)
        if (!range) continue
        const base = `${c.id}${variant}${c.sheets.length > 1 ? `-${sheet.name}` : ''}`
        const png = join(REFS, `${base}.png`)
        capture(join(FIXTURES, `${c.id}${variant}.xlsx`), sheet.name, range, png)
        writeFileSync(join(REFS, `${base}.json`), JSON.stringify({ excelVersion, route: 'copy-picture', scale: 1, pxPerPt: 1, sheet: sheet.name, range, capturedAt: new Date().toISOString() }, null, 2) + '\n')
        count++
        console.log(`captured ${base} (${range})`)
      }
    }
  }
  console.log(`${count} references (workbooks left open; close Excel windows when done)`)
}
