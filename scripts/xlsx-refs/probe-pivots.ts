/**
 * Stage 1 Excel probes for the pivot reference harness (pivots plan Task 1). Reproduces the recorded findings in
 * docs/validation/xlsx-pivots.md ("Stage 1 probes"):
 *   A. a pivot cache cannot be created through AppleScript, but an existing pivot can be reshaped and saved;
 *   B. refreshOnLoad is not applied when automation opens the file, while `refresh table` recomputes from source;
 *   C. `copy picture` appearance screen keeps gridlines and printer drops them; neither shows field buttons.
 *
 *   bun scripts/xlsx-refs/probe-pivots.ts <out-dir>
 *
 * Needs corpus/poi-ExcelPivotTableSample.xlsx (bun scripts/fetch-corpus.ts) and an awake, unlocked display.
 */
import { copyFileSync, mkdirSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { capture } from './capture'
import { closeWorkbook, openWorkbook, q, runScript, saveAs } from './excel'

const out = resolve(process.argv[2] ?? 'pivot-probes')
mkdirSync(out, { recursive: true })
const sample = resolve(import.meta.dir, '../../corpus/poi-ExcelPivotTableSample.xlsx')
const work = join(out, 'probe-a-source.xlsx')
copyFileSync(sample, work)
const name = openWorkbook(work)
const wb = `workbook "${q(name)}"`
const pivot = `pivot table 1 of worksheet "Tabelle2" of ${wb}`
const attempt = (label: string, lines: string[]) => {
  try { console.log(`${label}: ${runScript(lines, 60_000) || 'ok'}`) } catch (e) { console.log(`${label}: FAILED ${String((e as Error).message).split('\n').find(l => l.includes('error')) ?? ''}`) }
}
// A. creation (expected to fail: parameter error -50) and reshaping (expected to work).
attempt('A make pivot cache', [`make new pivot cache at ${wb} with properties {source data:"Tabelle1!R1C1:R4C3"}`])
attempt('A table style', [`set table style of ${pivot} to "PivotStyleMedium9"`, `return table style of ${pivot}`])
attempt('A show values as', [`set calculation of data field 1 of ${pivot} to pivot field calculation percent of total`, `refresh table ${pivot}`])
const savedA = join(out, 'probe-a.xlsx')
saveAs(name, savedA)
// C. appearances on the reshaped pivot.
capture(savedA, 'Tabelle2', 'A1:E8', join(out, 'probe-c-screen.png'), { appearance: 'screen' })
capture(savedA, 'Tabelle2', 'A1:E8', join(out, 'probe-c-printer.png'), { appearance: 'printer' })
closeWorkbook('probe-a.xlsx', false)
console.log(`saved ${savedA}; B (refreshOnLoad) is recorded by hand in docs/validation/xlsx-pivots.md`)
