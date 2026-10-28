/**
 * Captures Excel's rendering of every Stage 1 reference case (pivots plan Task 3, step 3) through the shared
 * `copy picture` harness, in each case's declared appearance.
 *
 * Pivot cases capture from A1 to the bottom-right of the pivot's saved `location`, so page fields above the pivot
 * are included and the range follows the real layout (values-on-rows runs far below the default A1:H24). The
 * `screen` appearance keeps gridlines, which give the cell geometry the sparkline harness got from bordered twins.
 *
 *   bun scripts/xlsx-refs/capture-all.ts [--kind table|pivot|slicer|timeline] [--case id] [--skip-existing]
 *
 * Writes tests/fixtures/xlsx-pivots/refs/<id>.png plus a .json sidecar. Needs an awake, unlocked display.
 */
import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { capture } from './capture'
import { REF_CASES, type RefCase } from './cases'
import { closeWorkbook, osa } from './excel'

const FIXTURES = resolve(import.meta.dir, '../../tests/fixtures/xlsx-pivots')
const REFS = join(FIXTURES, 'refs')

/** A1 through the bottom-right cell of the first pivot's saved `<location ref>`. */
export function pivotCaptureRange(file: string): string {
  const xml = execFileSync('unzip', ['-p', file, 'xl/pivotTables/pivotTable1.xml'], { encoding: 'utf8' })
  const ref = /<location ref="([A-Z]+\d+):([A-Z]+\d+)"/.exec(xml)
  if (!ref) throw new Error(`no pivot location in ${file}`)
  return `A1:${ref[2]}`
}

export const captureRange = (c: RefCase, file: string): string => c.kind === 'pivot' ? pivotCaptureRange(file) : c.range

if (import.meta.main) {
  const args = process.argv.slice(2)
  const arg = (flag: string) => args.includes(flag) ? args[args.indexOf(flag) + 1] : undefined
  const kind = arg('--kind'), only = arg('--case'), skipExisting = args.includes('--skip-existing')
  mkdirSync(REFS, { recursive: true })
  const excelVersion = osa(['tell application "Microsoft Excel" to get version'])
  let count = 0
  for (const c of REF_CASES) {
    if ((kind && c.kind !== kind) || (only && c.id !== only)) continue
    const file = join(FIXTURES, `${c.id}.xlsx`)
    if (!existsSync(file)) { console.log(`missing ${c.id}.xlsx (build it first)`); continue }
    if (skipExisting && existsSync(join(REFS, `${c.id}.json`))) continue
    const range = captureRange(c, file)
    capture(file, c.sheet, range, join(REFS, `${c.id}.png`), { appearance: c.appearance })
    closeWorkbook(`${c.id}.xlsx`, false)
    writeFileSync(join(REFS, `${c.id}.json`), JSON.stringify({ excelVersion, route: 'copy-picture', appearance: c.appearance, pxPerPt: 1, sheet: c.sheet, range, capturedAt: new Date().toISOString() }, null, 2) + '\n')
    count++
    console.log(`captured ${c.id} (${range})`)
  }
  console.log(`${count} references`)
}
