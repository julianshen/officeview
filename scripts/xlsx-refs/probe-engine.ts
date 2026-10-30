/**
 * Stage 2 Excel probes (pivots plan Task 15, step 1): which pivot reshapes the AppleScript dictionary reaches, so the
 * engine oracle can be built by Excel itself. Each probe opens its own copy of the seed, applies one reshape,
 * refreshes, saves, and checks the saved pivot definition for the marker that proves Excel applied it (Excel can
 * ignore a setting without an error).
 *
 *   bun scripts/xlsx-refs/probe-engine.ts <out-dir inside the worktree> [probe-id …]
 *
 * Needs an awake, unlocked display. Results are recorded in docs/validation/xlsx-pivots.md ("Stage 2 probes").
 */
import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'
import { closeWorkbook, openWorkbook, q, runScript, saveAs, waitFor, workbookNames } from './excel'

const SEED = resolve(import.meta.dir, '../../tests/fixtures/xlsx-pivots/recipes/pivot-seed.xlsx')

interface Probe { id: string; lines: (p: string, ws: string, wb: string) => string[]; marker: RegExp; part?: string }

const PIVOT_PART = 'xl/pivotTables/pivotTable1.xml'
const CACHE_PART = 'xl/pivotCache/pivotCacheDefinition1.xml'
const DATES = Array.from({ length: 36 }, (_, i) => `date "${2019 + (i % 6)}/${(i % 12) + 1}/${(i % 27) + 1}"`)

export const PROBES: Probe[] = [
  { id: 'function-average', lines: p => [`set function of data field "Sum of Sales" of ${p} to do average`], marker: /<dataField [^>]*subtotal="average"/ },
  { id: 'function-stdevp', lines: p => [`set function of data field "Sum of Sales" of ${p} to do standard deviation p`], marker: /<dataField [^>]*subtotal="stdDevp"/ },
  {
    id: 'show-as-difference-previous',
    lines: p => [`set calculation of data field "Sum of Sales" of ${p} to pivot field calculation difference from`,
      `set base field of data field "Sum of Sales" of ${p} to "Region"`, `set base item of data field "Sum of Sales" of ${p} to "(previous)"`],
    marker: /<dataField [^>]*showDataAs="difference"[^>]*baseItem="1048828"/,
  },
  {
    id: 'show-as-percent-of-item',
    lines: p => [`set calculation of data field "Sum of Sales" of ${p} to pivot field calculation percent of`,
      `set base field of data field "Sum of Sales" of ${p} to "Region"`, `set base item of data field "Sum of Sales" of ${p} to "North"`],
    marker: /<dataField [^>]*showDataAs="percent"/,
  },
  // XLSubTotalType enumerators per Excel.sdef: 'sum subtotal', 'count subtotal', 'average subtotal', …
  { id: 'named-subtotals', lines: p => [`set subtotals pivot field "Region" of ${p} subtotal index sum subtotal value true`, `set subtotals pivot field "Region" of ${p} subtotal index count subtotal value true`], marker: /<pivotField [^>]*sumSubtotal="1"/ },
  { id: 'blank-line', lines: p => [`set layout blank line of pivot field "Region" of ${p} to true`], marker: /<pivotField [^>]*insertBlankRow="1"/ },
  { id: 'auto-show-top3', lines: p => [`auto show pivot field "Product" of ${p} type type automatic range position top count 2 field "Sum of Sales"`], marker: /<pivotField [^>]*autoShow="1"/ },
  { id: 'auto-sort-by-value', lines: p => [`auto sort pivot field "Region" of ${p} sort order sort descending sort field "Sum of Sales"`], marker: /<autoSortScope>/ },
  { id: 'sort-label-descending', lines: p => [`auto sort pivot field "Product" of ${p} sort order sort descending sort field "Product"`], marker: /<pivotField [^>]*sortType="descending"/ },
  { id: 'hidden-item', lines: p => [`set visible of pivot item "West" of pivot field "Region" of ${p} to false`], marker: /<item [^>]*h="1"/ },
  {
    id: 'page-multi',
    lines: p => [`set pivot field orientation of pivot field "Product" of ${p} to orient as page field`, `set enable multiple page items of pivot field "Product" of ${p} to true`,
      `set visible of pivot item "Bananas" of pivot field "Product" of ${p} to false`],
    marker: /<pivotField [^>]*multipleItemSelectionAllowed="1"/,
  },
  { id: 'label-filter', lines: p => [`make new pivot filter at pivot field "Product" of ${p} with properties {filter type:pivot caption begins with, value1:"B"}`], marker: /<filter [^>]*type="captionBeginsWith"/ },
  { id: 'value-filter', lines: p => [`make new pivot filter at pivot field "Region" of ${p} with properties {filter type:pivot value is greater than, data field:data field "Sum of Sales" of ${p}, value1:20000}`], marker: /<filter [^>]*type="valueGreaterThan"/ },
  { id: 'top10-filter', lines: p => [`make new pivot filter at pivot field "Region" of ${p} with properties {filter type:pivot top count, data field:data field "Sum of Sales" of ${p}, value1:2}`], marker: /<filter [^>]*type="count"/ },
  { id: 'calculated-field', lines: p => [`make new calculated field at ${p} with properties {name:"Unit Price", formula:"=Sales/Units"}`], marker: /<cacheField [^>]*formula=/, part: CACHE_PART },
  { id: 'calculated-item', lines: p => [`make new calculated item at pivot field "Product" of ${p} with properties {name:"AB", formula:"=Apples+Bananas"}`], marker: /<item [^>]*f="1"|<calculatedItems/ },
  {
    id: 'source-with-dates',
    lines: (p, _ws, wb) => [`set value of range "F1" of worksheet "Data" of ${wb} to "Date"`,
      ...DATES.map((d, i) => `set value of range "F${i + 2}" of worksheet "Data" of ${wb} to ${d}`),
      `set source data of ${p} to "Data!R1C1:R37C6"`],
    marker: /<cacheFields count="6">/, part: CACHE_PART,
  },
  {
    id: 'group-years-numbers',
    // Year is numeric (2022–2024): a number range group over the row/column item cell.
    lines: (p, ws) => [`set pivot field orientation of pivot field "Year" of ${p} to orient as row field`,
      `group (label range of pivot field "Year" of ${p}) start 2022 end 2024 by 2`],
    marker: /<fieldGroup /, part: CACHE_PART,
  },
]

function unzipPart(file: string, part: string): string {
  return execFileSync('unzip', ['-p', file, part], { encoding: 'utf8', maxBuffer: 64 << 20 })
}

if (import.meta.main) {
  const out = resolve(process.argv[2] ?? 'engine-probes')
  const only = new Set(process.argv.slice(3))
  mkdirSync(out, { recursive: true })
  for (const probe of PROBES) {
    if (only.size && !only.has(probe.id)) continue
    const work = join(out, `${probe.id}-in.xlsx`)
    copyFileSync(SEED, work)
    const name = openWorkbook(work)
    const wb = `workbook "${q(name)}"`
    const ws = `worksheet "Pivot" of ${wb}`
    const p = `pivot table 1 of ${ws}`
    const saved = join(out, `${probe.id}.xlsx`)
    let status: string
    try {
      runScript([...probe.lines(p, ws, wb), `refresh table ${p}`], 90_000)
      saveAs(name, saved)
      // Excel writes the file after the script returns.
      waitFor(() => existsSync(saved), `write ${saved}`)
      closeWorkbook(basename(saved), false)
      status = probe.marker.test(unzipPart(saved, probe.part ?? PIVOT_PART)) ? 'REACHED' : 'IGNORED (no marker)'
    } catch (e) {
      status = `FAILED ${String((e as Error).message).split('\n').find(l => /error/i.test(l))?.trim() ?? ''}`
      // The workbook is open under its own name, or under the saved name once `save as` renamed it.
      for (const open of workbookNames() ?? []) if (open === name || open === basename(saved)) closeWorkbook(open, false)
    }
    console.log(`${probe.id}: ${status}`)
  }
}
