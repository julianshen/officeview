/**
 * Measurements read off the Excel-saved oracle files (pivots plan Task 24, step 3), printed as markdown for
 * docs/validation/xlsx-pivots.md ("Stage 3 interaction oracle and measurements"). Read-only; run after
 * `build-engine.ts --interaction`.
 *
 *   bun scripts/xlsx-refs/measure-interaction.ts [--section widths|captions|charts|getpivotdata|refresh]
 *
 * (a) widths: the Pivot sheet's `<col>` widths of every built interaction case against its start file (seed or
 *     base), over the pivot's columns, with the longest saved text per column: the §7.8 autofit rule's input.
 * (b) captions: every implicit-caption cell (the engine's English default, `caption: 'implicit'`) of every built
 *     engine/ and interaction/ file, with the text Excel saved there, grouped by kind into patterns; plus the
 *     data-field names Excel assigned (zh-TW). The caption pattern table's input.
 * (c) charts: per chart case, each series' tx/cat/val refs and point counts, `c:pivotFmts`, and the pivot's
 *     chartFormats (§6.6, CODEX-3).
 * (d) getpivotdata: the cached value of each formula on sheet Formulas per case (§7.5).
 * (e) refresh: location, recordCount and records part of the refresh seeds and cases (§7.7).
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { computePivot } from '../../src/xlsx/pivots/engine/compute'
import { initialViewState } from '../../src/xlsx/pivots/engine/view'
import { GPD_FORMULAS, GPD_SHEET, SEED_FILES, startOf, type InteractionCase } from './interaction-cases'
import type { InteractionManifestEntry } from './build-engine'
import {
  CACHE_RE, CHART_RE, PIVOT_DIR, PIVOT_RE, RECORDS_RE, colWidths, defaultColWidth, fixturePath, loadPivot, loadZip, partNames, partText, sharedStrings,
  sheetCells, sheetXml, type SheetCell,
} from './pivot-files'

const readJson = <T>(file: string): T[] => existsSync(file) ? JSON.parse(readFileSync(file, 'utf8')) as T[] : []
const interaction = readJson<InteractionManifestEntry>(join(PIVOT_DIR, 'interaction/manifest.json')).filter(e => e.status === 'built')
const engine = readJson<{ case: string; status: string }>(join(PIVOT_DIR, 'engine/manifest.json')).filter(e => e.status === 'built')
const caseFile = (e: InteractionManifestEntry) => fixturePath(`interaction/${e.case}`)
const startFile = (e: InteractionManifestEntry) => fixturePath(startOf(e as Pick<InteractionCase, 'seed' | 'base'>))
const colName = (c: number) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s }
const cell = (row: number, col: number) => `${colName(col)}${row + 1}`
const md = (v: unknown) => String(v ?? '').replace(/\|/g, '\\|').replace(/\n/g, ' ')
const table = (head: string[], rows: unknown[][]) => [`| ${head.join(' | ')} |`, `| ${head.map(() => '---').join(' | ')} |`, ...rows.map(r => `| ${r.map(md).join(' | ')} |`)].join('\n')

interface PivotSheet { cells: SheetCell[]; widths: Map<number, { width: number; custom: boolean; bestFit: boolean }>; defaultWidth?: number; ref: string; minCol: number; maxCol: number }
/** The pivot's sheet: cells, explicit widths per column, the location ref and its column span. */
async function pivotSheet(file: string): Promise<PivotSheet | undefined> {
  const zip = await loadZip(file)
  const pivot = await partText(zip, partNames(zip, PIVOT_RE)[0])
  const ref = pivot && /<location [^>]*ref="([^"]+)"/.exec(pivot)?.[1]
  const xml = await sheetXml(zip, 'Pivot')
  if (!ref || !xml) return undefined
  const [a, b] = ref.split(':')
  const col = (r: string) => [.../^([A-Z]+)/.exec(r)![1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
  const widths = new Map<number, { width: number; custom: boolean; bestFit: boolean }>()
  for (const w of colWidths(xml)) for (let c = w.min - 1; c < w.max; c++) widths.set(c, { width: w.width, custom: w.customWidth, bestFit: w.bestFit })
  return { cells: sheetCells(xml, await sharedStrings(zip)), widths, defaultWidth: defaultColWidth(xml), ref, minCol: col(a!), maxCol: col(b ?? a!) }
}
const longest = (s: PivotSheet, col: number) => s.cells.filter(c => c.col === col && c.value !== undefined).reduce<SheetCell | undefined>((m, c) => !m || c.value!.length > m.value!.length ? c : m, undefined)

async function widths(): Promise<string> {
  const out = ['### (a) Column widths after a change (§7.8 autofit input)', '',
    'Pivot sheet `<col>` widths (character units; `—` = no `<col>`, the sheet default) over the union of the start and saved pivot columns. ' +
    '"longest" is the longest saved cell text in that column of the case (numbers as stored, before number formatting).', '']
  for (const e of interaction) {
    const [start, saved] = [await pivotSheet(startFile(e)), await pivotSheet(caseFile(e))]
    if (!start || !saved) { out.push(`- ${e.case}: no Pivot sheet or location`); continue }
    const from = Math.min(start.minCol, saved.minCol), to = Math.max(start.maxCol, saved.maxCol)
    const rows: unknown[][] = []
    for (let c = from; c <= to; c++) {
      const a = start.widths.get(c), b = saved.widths.get(c)
      const delta = a && b ? Math.round((b.width - a.width) * 1000) / 1000 : ''
      const l = longest(saved, c)
      rows.push([colName(c), a ? `${a.width}${a.bestFit ? ' bf' : ''}` : '—', b ? `${b.width}${b.bestFit ? ' bf' : ''}` : '—', delta, l ? `${l.value} (${l.value!.length}, ${l.ref})` : ''])
    }
    const changed = rows.some(r => r[3] !== '' && r[3] !== 0) || rows.some(r => r[1] !== r[2])
    out.push(`**${e.case}** (start ${startOf(e as Pick<InteractionCase, 'seed' | 'base'>)}, location ${start.ref} → ${saved.ref}${e.setup ? `; setup: ${e.setup}` : ''})${changed ? '' : ' — widths unchanged'}`, '')
    out.push(table(['col', 'start', 'saved', 'Δ', 'longest saved text (chars, cell)'], rows), '')
  }
  // Stage 2 files against engine/baseline: every one is the seed plus one change, saved by Excel.
  const base = await pivotSheet(fixturePath('engine/baseline'))
  const rows: unknown[][] = []
  for (const e of engine) {
    const saved = await pivotSheet(fixturePath(`engine/${e.case}`))
    if (!base || !saved || e.case === 'baseline') continue
    for (let c = Math.min(base.minCol, saved.minCol); c <= Math.max(base.maxCol, saved.maxCol); c++) {
      const a = base.widths.get(c)?.width, b = saved.widths.get(c)?.width
      if (a === b) continue
      const lb = longest(base, c), ls = longest(saved, c)
      rows.push([e.case, colName(c), a ?? '—', b ?? '—', a !== undefined && b !== undefined ? Math.round((b - a) * 1000) / 1000 : '', lb ? `${lb.value} (${lb.value!.length})` : '', ls ? `${ls.value} (${ls.value!.length})` : ''])
    }
  }
  out.push('**Stage 2 cases against engine/baseline** (columns whose width differs only)', '',
    table(['case', 'col', 'baseline', 'saved', 'Δ', 'baseline longest', 'saved longest'], rows), '')
  return out.join('\n')
}

// ---- (b) captions ---------------------------------------------------------------------------------------------

const FN = 'Sum|Count|Average|Max|Min|Product|StdDev|StdDevp|Var|Varp'
/** Caption kind of the engine's English default, and the variable part (to turn the saved text into a pattern). */
function captionKind(english: string): { kind: string; vars: Array<[string, string]> } {
  let m: RegExpExecArray | null
  if (english === 'Row Labels') return { kind: 'row labels', vars: [] }
  if (english === 'Column Labels') return { kind: 'column labels', vars: [] }
  if (english === 'Grand Total') return { kind: 'grand total', vars: [] }
  if (english === 'Total') return { kind: 'classic Total', vars: [] }
  if (english === '(blank)') return { kind: '(blank) item', vars: [] }
  if (english === '(All)') return { kind: 'page (All)', vars: [] }
  if (english === '(Multiple Items)') return { kind: 'page (Multiple Items)', vars: [] }
  if ((m = new RegExp(`^(${FN}) of (.+)$`).exec(english))) return { kind: `data field default name (${m[1]})`, vars: [['{field}', m[2]!]] }
  if ((m = /^Total (.+)$/.exec(english))) return { kind: 'grand total on the Values axis', vars: [['{data}', m[1]!]] }
  if ((m = /^(.+) Total$/.exec(english))) return { kind: 'subtotal', vars: [['{item}', m[1]!]] }
  if ((m = new RegExp(`^(.+) (${FN})$`).exec(english))) return { kind: `named subtotal (${m[2]})`, vars: [['{item}', m[1]!]] }
  return { kind: 'other', vars: [] }
}
function pattern(saved: string, vars: Array<[string, string]>): string {
  let p = saved
  for (const [name, value] of vars) if (value && p.includes(value)) p = p.replace(value, name)
  return p
}
const savedText = (sheet: { rows: ReadonlyArray<{ index: number; cells: ReadonlyArray<{ col: number; value: unknown }> }> }, row: number, col: number) =>
  sheet.rows.find(r => r.index === row)?.cells.find(c => c.col === col)?.value

async function captions(): Promise<string> {
  const files = [...engine.map(e => ({ id: `engine/${e.case}`, file: fixturePath(`engine/${e.case}`) })), ...interaction.map(e => ({ id: `interaction/${e.case}`, file: caseFile(e) }))]
  const kinds = new Map<string, Map<string, { english: string; example: string; count: number }>>()
  const refused: string[] = [], names = new Map<string, string>()
  for (const { id, file } of files) {
    let loaded
    try { loaded = await loadPivot(file) } catch (err) { refused.push(`${id}: ${(err as Error).message}`); continue }
    const { def, cache, sheet, store, date1904 } = loaded
    for (const d of def.dataFields) if (d.name && !/^(Sum|Count|Average|Max|Min) of /.test(d.name) && !names.has(d.name)) names.set(d.name, id)
    if (!store) { refused.push(`${id}: no records (${loaded.storeError ?? 'saveData="0"'})`); continue }
    const result = computePivot(store, cache, def, initialViewState(def, cache), { date1904 })
    if (!result.ok) { refused.push(`${id}: engine refuses (${result.diagnostic.feature})`); continue }
    for (const c of result.value.cells) {
      if (c.caption !== 'implicit' || c.value.kind !== 'text') continue
      const english = c.value.value, saved = savedText(sheet, c.row, c.col)
      const { kind, vars } = captionKind(english)
      const p = typeof saved === 'string' ? pattern(saved, vars) : `(saved ${saved === undefined ? 'nothing' : JSON.stringify(saved)})`
      const byPattern = kinds.get(kind) ?? new Map()
      const hit = byPattern.get(p) ?? { english, example: `${id}!${cell(c.row, c.col)} = ${JSON.stringify(saved)}`, count: 0 }
      hit.count++
      byPattern.set(p, hit)
      kinds.set(kind, byPattern)
    }
  }
  const rows: unknown[][] = []
  for (const [kind, byPattern] of [...kinds].sort(([a], [b]) => a.localeCompare(b))) for (const [p, h] of byPattern) rows.push([kind, h.english, p, h.example, h.count])
  return ['### (b) Implicit captions by kind (caption pattern table input)', '',
    'Every cell the engine marks `caption: \'implicit\'` (computed on the saved state with no saved-caption lookup), with the text Excel saved there. ' +
    'Pattern variables: `{item}` the item label, `{data}` the data field name, `{field}` the source field name.', '',
    table(['kind', 'engine default (English)', 'saved pattern', 'first example', 'cells'], rows), '',
    'Data-field names Excel assigned (explicit `dataField@name`, not English defaults):', '',
    table(['name', 'first file'], [...names].map(([n, f]) => [n, f])), '',
    'Not measured (engine refused or no records):', '', ...refused.map(r => `- ${r}`), ''].join('\n')
}

// ---- (c) charts -----------------------------------------------------------------------------------------------

async function charts(): Promise<string> {
  const out = ['### (c) Pivot chart series and categories (§6.6, CODEX-3)', '']
  const entries = [{ case: 'chart-seed (input)', file: fixturePath(SEED_FILES['chart-seed']) }, ...interaction.filter(e => e.seed === 'chart-seed').map(e => ({ case: e.case, file: caseFile(e) }))]
  for (const { case: id, file } of entries) {
    const zip = await loadZip(file)
    const pivot = await partText(zip, partNames(zip, PIVOT_RE)[0]) ?? ''
    for (const name of partNames(zip, CHART_RE)) {
      const xml = (await partText(zip, name))!
      const source = /<c:pivotSource><c:name>([^<]*)<\/c:name>(?:<c:fmtId val="(\d+)"\/>)?/.exec(xml)
      const fmts = (xml.match(/<c:pivotFmt>/g) ?? []).length
      const series = [...xml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)].map(m => {
        const s = m[1]!
        const ref = (tag: string) => new RegExp(`<c:${tag}>[\\s\\S]*?<c:f>([^<]*)<\\/c:f>`).exec(s)?.[1] ?? ''
        const count = (tag: string) => new RegExp(`<c:${tag}>[\\s\\S]*?<c:ptCount val="(\\d+)"`).exec(s)?.[1] ?? ''
        const title = /<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/.exec(s)?.[1] ?? ''
        return [/<c:idx val="(\d+)"/.exec(s)?.[1], title, ref('tx'), ref('cat'), count('cat'), ref('val'), count('val')]
      })
      const chartFormats = /<chartFormats count="(\d+)"/.exec(pivot)?.[1] ?? '0'
      out.push(`**${id}** ${name}: pivotSource ${source?.[1] ?? '—'} fmtId ${source?.[2] ?? '—'}; ${series.length} series; ${fmts} c:pivotFmt; pivot chartFormats ${chartFormats}; location ${/<location [^>]*ref="([^"]+)"/.exec(pivot)?.[1] ?? ''}`, '')
      out.push(table(['idx', 'title', 'tx ref', 'cat ref', 'cat pts', 'val ref', 'val pts'], series), '')
    }
  }
  return out.join('\n')
}

// ---- (d) GETPIVOTDATA -----------------------------------------------------------------------------------------

async function getpivotdata(): Promise<string> {
  const entries = [{ case: 'seed', file: fixturePath(SEED_FILES['getpivotdata-seed']) }, ...interaction.filter(e => e.seed === 'getpivotdata-seed').map(e => ({ case: e.case, file: caseFile(e) }))]
    .filter(e => existsSync(e.file))
  const values = new Map<string, Map<string, string>>()
  for (const e of entries) {
    const zip = await loadZip(e.file)
    const xml = await sheetXml(zip, GPD_SHEET)
    const cells = xml ? sheetCells(xml, await sharedStrings(zip)) : []
    values.set(e.case, new Map(cells.filter(c => c.col === 1).map(c => [c.ref, c.type === 'e' ? `error ${c.value}` : c.value ?? '(no cached value)'])))
  }
  const rows = GPD_FORMULAS.map(([label, formula], i) => [label, formula, ...entries.map(e => values.get(e.case)?.get(`B${i + 1}`) ?? '—')])
  return ['### (d) GETPIVOTDATA and references after a change (§7.5)', '', 'Cached values Excel saved on sheet Formulas, column B.', '',
    table(['formula', 'text', ...entries.map(e => e.case)], rows), ''].join('\n')
}

// ---- (e) refresh ----------------------------------------------------------------------------------------------

async function refresh(): Promise<string> {
  const entries = [
    { case: 'refresh-seed (input)', file: fixturePath(SEED_FILES['refresh-seed']) }, { case: 'refresh-seed-nodata (input)', file: fixturePath('recipes/refresh-seed-nodata') },
    ...interaction.filter(e => e.seed === 'refresh-seed').map(e => ({ case: e.case, file: caseFile(e) })),
  ].filter(e => existsSync(e.file))
  const rows: unknown[][] = []
  for (const e of entries) {
    const zip = await loadZip(e.file)
    const pivot = await partText(zip, partNames(zip, PIVOT_RE)[0]) ?? '', cache = await partText(zip, partNames(zip, CACHE_RE)[0]) ?? ''
    const flag = (name: string) => new RegExp(`<pivotCacheDefinition [^>]*${name}="([^"]*)"`).exec(cache)?.[1] ?? '—'
    rows.push([e.case, /<location [^>]*ref="([^"]+)"/.exec(pivot)?.[1], flag('refreshOnLoad'), flag('saveData'), flag('recordCount'), partNames(zip, RECORDS_RE).length ? 'yes' : 'no'])
  }
  return ['### (e) refreshOnLoad seeds and cases (§7.7)', '', table(['file', 'location', 'refreshOnLoad', 'saveData', 'recordCount', 'records part'], rows), ''].join('\n')
}

if ((import.meta as { main?: boolean }).main) {
  const args = process.argv.slice(2)
  const only = args.includes('--section') ? args[args.indexOf('--section') + 1] : undefined
  const sections: Record<string, () => Promise<string>> = { widths, captions, charts, getpivotdata, refresh }
  if (only && !sections[only]) { console.error(`unknown section ${only}`); process.exit(2) }
  if (!interaction.length) console.log('> interaction/manifest.json has no built cases yet: only the engine/ files are measured.\n')
  for (const [name, run] of Object.entries(sections)) if (!only || only === name) console.log(await run())
}
