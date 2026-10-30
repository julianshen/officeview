/**
 * Pivot chart references for the chart field buttons (pivots plan Task 11; the plan's "Task 3 measurement" did
 * not exist, ledger Ruling AG).
 *
 * The seed `recipes/pivot-chart-seed.xlsx` was made once in Excel 16.114: `pivot-layout-compact.xlsx` with a
 * clustered column chart whose source is the pivot range A3:C10 (AppleScript `make new chart object` +
 * `set source data`), which Excel saves as a PivotChart (`c:pivotSource`, `c14:pivotOptions`,
 * `c16:pivotOptions16`). Each variant edits only those flags and moves the chart clear of the pivot; `--resave`
 * has Excel open and re-save each variant, so the committed workbook is Excel's own output.
 *
 * There are no captures: Excel for Mac draws no PivotChart field buttons for any flag combination, and copy picture
 * omits them too (docs/validation/xlsx-pivots.md, "Pivot chart field buttons"; ledger Ruling AI).
 *
 *   bun scripts/xlsx-refs/build-pivot-charts.ts [--resave]
 *
 * Needs an awake, unlocked display for --resave.
 */
import { readFileSync, writeFileSync, existsSync, rmSync, copyFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { closeWorkbook, openWorkbook, saveAs } from './excel'

const DIR = resolve(import.meta.dir, '../../tests/fixtures/xlsx-pivots')
const SEED = join(DIR, 'recipes/pivot-chart-seed.xlsx')
const CHART = 'xl/charts/chart1.xml'
const DRAWING = 'xl/drawings/drawing1.xml'
/** Moves the chart clear of the pivot (anchor columns +4: L2:P18). */
const COLUMN_SHIFT = 4

type Zone = 'Filter' | 'Categories' | 'Data' | 'Series'
export interface PivotChartVariant { id: string; zones: Record<Zone, boolean>; visible: boolean; expandCollapse: boolean }

const all = (on: boolean): Record<Zone, boolean> => ({ Filter: on, Categories: on, Data: on, Series: on })
const only = (zone: Zone): Record<Zone, boolean> => ({ ...all(false), [zone]: true })

export const PIVOT_CHART_VARIANTS: readonly PivotChartVariant[] = [
  { id: 'pivot-chart-buttons-all', zones: all(true), visible: true, expandCollapse: true },
  { id: 'pivot-chart-buttons-hidden', zones: all(true), visible: false, expandCollapse: true },
  { id: 'pivot-chart-buttons-filter', zones: only('Filter'), visible: true, expandCollapse: true },
  { id: 'pivot-chart-buttons-categories', zones: only('Categories'), visible: true, expandCollapse: true },
  { id: 'pivot-chart-buttons-data', zones: only('Data'), visible: true, expandCollapse: true },
  { id: 'pivot-chart-buttons-series', zones: only('Series'), visible: true, expandCollapse: true },
  { id: 'pivot-chart-buttons-no-expand', zones: all(true), visible: true, expandCollapse: false },
]

const swap = (xml: string, from: RegExp, to: string, what: string): string => {
  const next = xml.replace(from, to)
  if (next === xml && !from.test(xml)) throw new Error(`seed changed shape: ${what}`)
  return next
}

export async function writeVariant(v: PivotChartVariant): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(SEED))
  let chart = await zip.file(CHART)!.async('text')
  for (const zone of Object.keys(v.zones) as Zone[]) chart = swap(chart, new RegExp(`<c14:dropZone${zone} val="[01]"/>`, 'g'), `<c14:dropZone${zone} val="${v.zones[zone] ? 1 : 0}"/>`, `dropZone${zone}`)
  chart = swap(chart, /<c14:dropZonesVisible val="[01]"\/>/g, `<c14:dropZonesVisible val="${v.visible ? 1 : 0}"/>`, 'dropZonesVisible')
  chart = swap(chart, /<c16:showExpandCollapseFieldButtons val="[01]"\/>/g, `<c16:showExpandCollapseFieldButtons val="${v.expandCollapse ? 1 : 0}"/>`, 'showExpandCollapseFieldButtons')
  zip.file(CHART, chart)
  const drawing = await zip.file(DRAWING)!.async('text')
  zip.file(DRAWING, drawing.replace(/<xdr:col>(\d+)<\/xdr:col>/g, (_, col: string) => `<xdr:col>${Number(col) + COLUMN_SHIFT}</xdr:col>`))
  const out = join(DIR, `${v.id}.xlsx`)
  writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
  return out
}

/** Has Excel open a hand-edited variant and save it, so the committed file is Excel's output. */
function resave(file: string): void {
  const work = file.replace(/([^/]+)\.xlsx$/, '.$1.edit.xlsx')
  copyFileSync(file, work)
  const name = openWorkbook(work)
  try {
    rmSync(file)
    saveAs(name, file)
    closeWorkbook(file.split('/').pop()!, false)
  } finally {
    rmSync(work, { force: true })
  }
}

if (import.meta.main) {
  const args = process.argv.slice(2)
  if (!existsSync(SEED)) throw new Error(`missing seed ${SEED}`)
  for (const v of PIVOT_CHART_VARIANTS) {
    const file = await writeVariant(v)
    if (args.includes('--resave')) resave(file)
    console.log(`built ${file}`)
  }
}
