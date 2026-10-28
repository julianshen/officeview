/**
 * Builds the `build: 'excel'` pivot reference cases (pivots plan Task 3, step 1) by reshaping an Excel-made seed.
 *
 * Excel for Mac cannot create a pivot cache through AppleScript (docs/validation/xlsx-pivots.md, probe A), so the
 * seed `recipes/pivot-seed.xlsx` was made once through the Data menu (rows Region > Product, column Year, values
 * Sum of Sales and Sum of Units, compact form, on sheet Pivot). Each case opens the seed, applies its reshape
 * through the AppleScript pivot dictionary, refreshes, and saves as `tests/fixtures/xlsx-pivots/<id>.xlsx`, so
 * every cell, cache and pivot part is Excel's own output.
 *
 *   bun scripts/xlsx-refs/build-pivots.ts [--case id] | --measure
 *
 * Needs an awake, unlocked display. Output paths stay inside the worktree, where Excel already has file access.
 */
import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, rmSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { REF_CASES, type PivotCase, type RefCase } from './cases'
import { closeWorkbook, openWorkbook, q, runScript, saveAs } from './excel'

export const PIVOT_DIR = resolve(import.meta.dir, '../../tests/fixtures/xlsx-pivots')
const SEED = join(PIVOT_DIR, 'recipes/pivot-seed.xlsx')
/** The seed with a hand-written custom pivot style `OfficeViewPivot` (wholeTable, headerRow, totalRow,
 * firstRowSubheading and firstColumn elements) applied; AppleScript cannot define table styles. */
const CUSTOM_SEED = join(PIVOT_DIR, 'recipes/pivot-custom-style-input.xlsx')
const CUSTOM_STYLE = 'OfficeViewPivot'
const savedStyleName = (spec: PivotCase) => spec.options.customStyle ? CUSTOM_STYLE : spec.styleName

/** AppleScript lines that turn the seed pivot into one case; `p` addresses the pivot, `ws` its worksheet. */
export function reshapeLines(spec: PivotCase, p: string, ws: string): string[] {
  const o = spec.options
  // `table style2` is the property Excel saves (pivotTableStyleInfo); `table style` reads back but is not saved
  // (measured 2026-10-29). Setting it to the seed's own style is harmless, so every case sets it.
  const lines = [`set table style2 of ${p} to "${q(savedStyleName(spec))}"`]
  if (spec.layout !== 'compact') lines.push(`row axis layout ${p} layout ${spec.layout} row`)
  if (o.subtotalTop === true) lines.push(`subtotal location ${p} location at top`)
  if (o.subtotalTop === false) lines.push(`subtotal location ${p} location at bottom`)
  if (o.grandTotals === false) lines.push(`set row grand of ${p} to false`, `set column grand of ${p} to false`)
  if (o.dataOnRows) lines.push(`set pivot field orientation of data pivot field of ${p} to orient as row field`)
  if (o.collapsed) lines.push(`set show detail of pivot item "East" of pivot field "Region" of ${p} to false`)
  if (o.pageField) lines.push(`set pivot field orientation of pivot field "Product" of ${p} to orient as page field`)
  // Repeated item labels only show outside compact form, so this case also switches to tabular.
  if (o.repeatLabels) lines.push(`row axis layout ${p} layout tabular row`, `set repeat labels of pivot field "Region" of ${p} to true`)
  // A cell format applied inside the pivot is saved as a pivot `formats` entry (preserve formatting is on).
  if (o.formatsOverride) lines.push(`set color of interior object of range "B7" of ${ws} to {255, 230, 153}`)
  lines.push(`refresh table ${p}`)
  return lines
}

export function buildPivotCase(c: RefCase): string {
  const out = join(PIVOT_DIR, `${c.id}.xlsx`)
  const work = join(PIVOT_DIR, `.${c.id}.seed.xlsx`)
  copyFileSync((c.spec as PivotCase).options.customStyle ? CUSTOM_SEED : SEED, work)
  const name = openWorkbook(work)
  const ws = `worksheet "${q(c.sheet)}" of workbook "${q(name)}"`
  try {
    runScript(reshapeLines(c.spec as PivotCase, `pivot table 1 of ${ws}`, ws))
    if (existsSync(out)) rmSync(out)
    saveAs(name, out)
    closeWorkbook(`${c.id}.xlsx`, false)
  } finally {
    rmSync(work, { force: true })
  }
  // AppleScript drops some settings silently (see reshapeLines), so the saved style is checked, not assumed.
  const saved = execFileSync('unzip', ['-p', out, 'xl/pivotTables/pivotTable1.xml'], { encoding: 'utf8' })
  const style = /<pivotTableStyleInfo name="([^"]+)"/.exec(saved)?.[1]
  const expected = savedStyleName(c.spec as PivotCase)
  if (style !== expected) throw new Error(`${c.id}: saved style ${style}, expected ${expected}`)
  return out
}

/** Measurement workbooks (not reference cases): Excel opens a style input from style-inputs.ts and saves it as is. */
export const MEASURE_INPUTS = [['pivot-precedence-input.xlsx', 'measure-precedence'], ['pivot-cell-override-input.xlsx', 'measure-cell-override']] as const

export function resaveMeasureInput(input: string, id: string): string {
  const out = join(PIVOT_DIR, `${id}.xlsx`)
  const work = join(PIVOT_DIR, `.${id}.input.xlsx`)
  copyFileSync(join(PIVOT_DIR, 'recipes', input), work)
  const name = openWorkbook(work)
  try {
    if (existsSync(out)) rmSync(out)
    saveAs(name, out)
    closeWorkbook(`${id}.xlsx`, false)
  } finally {
    rmSync(work, { force: true })
  }
  return out
}

if (import.meta.main) {
  const args = process.argv.slice(2)
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  if (args.includes('--measure')) {
    for (const [input, id] of MEASURE_INPUTS) console.log(`built ${resaveMeasureInput(input, id)}`)
    process.exit(0)
  }
  let count = 0
  for (const c of REF_CASES) {
    if (c.kind !== 'pivot' || c.build !== 'excel' || (only && c.id !== only)) continue
    console.log(`built ${buildPivotCase(c)}`)
    count++
  }
  console.log(`${count} pivot references`)
}
