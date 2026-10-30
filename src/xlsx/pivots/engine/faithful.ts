/**
 * The faithfulness check (spec §6.4) and the engine gate (§7.6, §7.7). Before any number the engine computes reaches
 * a user, the engine recomputes the SAVED state of the pivot and proves it equals what Excel saved: the lines against
 * `rowItems`/`colItems`, then every cell of the location and its page fields. Anything else leaves the pivot as saved
 * and display-only. Lazy chunk `pivots` (Stage 2: library only). Nothing here mutates its inputs.
 *
 * Rulings recorded here (Task 19):
 * - A saved cell the reader cannot supply at a LABEL position (outside the value block) where the engine places a
 *   label is `cells-missing`: Excel always writes its labels and captions, so the sheet is incomplete (a truncated
 *   save, a failed paged read). At a VALUE position a blank saved cell means "no records", so any disagreement there
 *   is a `mismatch`, which also fails closed (the brief's "other staleness shows as mismatch").
 * - `cells-predate-records`: the cache's `recordCount` differs from the decoded store. No reliable save timestamp
 *   exists to compare with `refreshedDate`.
 * - Absent `rowItems` or `colItems` is unverified whatever the axis holds: the saved layout cannot be laid out
 *   without both (Excel writes `<i/>` for an empty axis).
 * - Implicit captions (Excel's default captions, localised by Excel) are compared for presence as text only.
 */
import type { SheetRect } from '../../tableStyles/model'
import type { ReadOptions, SourceStore } from '../../view/contracts'
import type { XlsxCell, XlsxSheet } from '../../types'
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition } from '../model'
import { savedPivotLayout, type PivotLine } from '../savedLayout'
import { computePivot } from './compute'
import type { PivotCellValue, PivotLayoutCell } from './place'
import type { RecordStore } from './records'
import { initialViewState } from './view'

export type SavedCell = { kind: 'number'; value: number } | { kind: 'text'; value: string } | { kind: 'bool'; value: boolean } | { kind: 'error'; value: string }
/** A saved sheet cell, or undefined for no value. `incomplete` marks a reader that could not read every cell it
 * covers (a failed or cancelled paged read): the check then reports `cells-missing` instead of trusting its blanks. */
export type CellReader = ((row: number, col: number) => SavedCell | undefined) & { readonly incomplete?: true }
export interface CellDiff { row: number; col: number; expected?: SavedCell; actual: PivotCellValue | 'line'; reason: string }
export type FaithfulResult =
  | { status: 'ok' }
  | { status: 'mismatch'; firstDiff: CellDiff }
  | { status: 'unverified'; reason: 'no-row-items' | 'no-col-items' | 'cells-missing' | 'cells-predate-records' | 'unsupported'; detail?: string }
/** The workbook's date system: the cache does not know it (Ruling S2-J). */
export interface FaithfulOptions { date1904: boolean }

type Unverified = Extract<FaithfulResult, { status: 'unverified' }>
const unverified = (reason: Unverified['reason'], detail?: string): Unverified => ({ status: 'unverified', reason, ...(detail !== undefined ? { detail } : {}) })
const mismatch = (diff: CellDiff): FaithfulResult => ({ status: 'mismatch', firstDiff: diff })
const refuse = (feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } })

/**
 * §7.6 gate before any decode: absent records part, unsupported source type or OLAP cache. A `refreshOnLoad` pivot
 * whose in-workbook source the caller can read takes the §7.7 route instead: `refresh: true` tells the caller
 * (Task 30) to build the store from the source range, not from the records part. A `refreshOnLoad` pivot whose source
 * cannot be read is not interactive (§7.7 "Otherwise … records only"): `pivot-records-absent` when it has no records,
 * else `pivot-refresh-on-load`. Stage 2 has no source reader, so its callers pass `sourceReadable: false`.
 */
export function engineEligibility(cache: PivotCacheDefinition, def: PivotTableDefinition, opts: { sourceReadable: boolean }): { ok: true; refresh: boolean } | { ok: false; diagnostic: PivotDiagnostic } {
  const name = `Pivot table ${def.name}`
  if (cache.olap) return refuse('pivot-source-unsupported', `${name} uses an OLAP or data-model cache`)
  if (cache.source.type !== 'worksheet') return refuse('pivot-source-unsupported', `${name} has a ${cache.source.type} source`)
  if (cache.source.worksheet?.relId !== undefined) return refuse('pivot-source-unsupported', `${name} has a source in another workbook`)
  const records = cache.saveData && cache.recordsPartPath !== undefined
  if (cache.refreshOnLoad && opts.sourceReadable) return { ok: true, refresh: true }
  if (!records) return refuse('pivot-records-absent', `${name} was saved without its records (saveData="0")`)
  if (cache.refreshOnLoad) return refuse('pivot-refresh-on-load', `${name} refreshes on load and its source cannot be read`)
  return { ok: true, refresh: false }
}

const sameLine = (a: PivotLine | undefined, b: PivotLine | undefined) => a !== undefined && b !== undefined && a.type === b.type
  && a.first === b.first && a.data === b.data && a.items.length === b.items.length && a.items.every((item, k) => item === b.items[k])
const lineText = (line: PivotLine | undefined) => line ? `${line.type} r=${line.first} i=${line.data} [${line.items.join(',')}]` : 'none'
const inRect = (rect: SheetRect, row: number, col: number) => row >= rect.minRow && row <= rect.maxRow && col >= rect.minCol && col <= rect.maxCol
/** A computed cell with nothing to show: no cell, an empty value or empty text (an item whose label is ''). */
const isBlank = (cell: PivotLayoutCell | undefined) => !cell || cell.value.kind === 'empty' || (cell.value.kind === 'text' && cell.value.value === '')
/** Excel's comparison precision (§9): 15 significant digits. */
const same15 = (a: number, b: number) => a === b || Number(a.toPrecision(15)) === Number(b.toPrecision(15))

/** Why `saved` differs from the computed cell, or undefined when they agree. */
function differs(cell: PivotLayoutCell | undefined, saved: SavedCell | undefined): string | undefined {
  if (isBlank(cell)) return saved === undefined ? undefined : 'the saved cell has a value the engine leaves blank'
  const value = cell!.value
  if (saved === undefined) return 'the engine places a value the saved cell does not have'
  if (cell!.caption === 'implicit') return saved.kind === 'text' ? undefined : 'a default caption over a non-text saved cell'
  switch (value.kind) {
    case 'number': return saved.kind === 'number' && same15(value.value, saved.value) ? undefined : 'the number differs at 15 significant digits'
    case 'text': return saved.kind === 'text' && saved.value === value.value ? undefined : 'the text differs'
    case 'error': return saved.kind === 'error' && saved.value === value.value ? undefined : 'the error differs'
    default: return 'the value differs'
  }
}

/**
 * The faithfulness check (§6.4): compute the saved state, compare its lines with `rowItems`/`colItems`, then every
 * cell in `location` and the page fields: numbers to 15 significant digits, text exactly, implicit captions for
 * presence. An engine refusal is `unverified` / `unsupported` with its feature as `detail`.
 */
export function checkSavedLayout(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition, saved: CellReader, options: FaithfulOptions): FaithfulResult {
  // An unknown construct: computePivot refuses it too, but its own feature names it here.
  if (def.unsupported.length) return unverified('unsupported', 'pivot-definition-unsupported')
  if (!def.rowItems) return unverified('no-row-items')
  if (!def.colItems) return unverified('no-col-items')
  if (cache.recordCount !== undefined && cache.recordCount !== store.count) return unverified('cells-predate-records', `recordCount ${cache.recordCount}, records ${store.count}`)
  const layoutSaved = savedPivotLayout(def)
  // Saved items that do not fill the location, or a layout too deep to lay out (savedPivotLayout's own refusals).
  if (!layoutSaved) return unverified('unsupported', 'pivot-saved-layout')
  const computed = computePivot(store, cache, def, initialViewState(def, cache), { date1904: options.date1904 })
  if (!computed.ok) return unverified('unsupported', computed.diagnostic.feature)
  const layout = computed.value
  const { rect, rects, dataRow0, dataCol0, itemRow0 } = layoutSaved

  // 1. Lines.
  for (const axis of ['rows', 'cols'] as const) {
    const ours = layout.axes[axis], theirs = layoutSaved[axis]
    for (let n = 0; n < Math.max(ours.length, theirs.length); n++) {
      if (sameLine(ours[n], theirs[n])) continue
      const at = axis === 'rows' ? { row: dataRow0 + n, col: rect.minCol } : { row: itemRow0, col: dataCol0 + n }
      return mismatch({ ...at, actual: 'line', reason: `${axis === 'rows' ? 'rowItems' : 'colItems'} line ${n}: computed ${lineText(ours[n])}, saved ${lineText(theirs[n])}` })
    }
  }
  // 2. The location.
  const r = layout.rect
  if (r.minRow !== rect.minRow || r.minCol !== rect.minCol || r.maxRow !== rect.maxRow || r.maxCol !== rect.maxCol) {
    return mismatch({ row: rect.minRow, col: rect.minCol, actual: 'line', reason: `location: computed R${r.minRow + 1}C${r.minCol + 1}:R${r.maxRow + 1}C${r.maxCol + 1}, saved ${def.location.ref}` })
  }
  const at = new Map<string, PivotLayoutCell>()
  for (const cell of layout.cells) at.set(`${cell.row},${cell.col}`, cell)
  // 3. Completeness: a label the engine places must have a saved cell (see the header ruling).
  if (saved.incomplete) return unverified('cells-missing', 'the saved cells could not all be read')
  const isValue = (row: number, col: number) => row >= dataRow0 && row <= rect.maxRow && col >= dataCol0 && col <= rect.maxCol
  for (const area of rects) for (let row = area.minRow; row <= area.maxRow; row++) for (let col = area.minCol; col <= area.maxCol; col++) {
    if (isValue(row, col) || isBlank(at.get(`${row},${col}`))) continue
    if (saved(row, col) === undefined) return unverified('cells-missing', `no saved cell at R${row + 1}C${col + 1}`)
  }
  // 4. Every cell, row-major over the location and page fields; then anything placed outside them.
  const bounds = rects.reduce((box, area) => ({ minRow: Math.min(box.minRow, area.minRow), minCol: Math.min(box.minCol, area.minCol), maxRow: Math.max(box.maxRow, area.maxRow), maxCol: Math.max(box.maxCol, area.maxCol) }))
  for (let row = bounds.minRow; row <= bounds.maxRow; row++) for (let col = bounds.minCol; col <= bounds.maxCol; col++) {
    if (!rects.some(area => inRect(area, row, col))) continue
    const cell = at.get(`${row},${col}`), expected = saved(row, col)
    const reason = differs(cell, expected)
    if (reason) return mismatch({ row, col, ...(expected ? { expected } : {}), actual: cell?.value ?? { kind: 'empty' }, reason })
  }
  for (const cell of layout.cells) {
    if (isBlank(cell) || rects.some(area => inRect(area, cell.row, cell.col))) continue
    return mismatch({ row: cell.row, col: cell.col, actual: cell.value, reason: 'the engine places a cell outside the saved location and page fields' })
  }
  return { status: 'ok' }
}

function savedOf(value: XlsxCell['value'] | undefined, isError: boolean | undefined): SavedCell | undefined {
  if (value === null || value === undefined || value === '') return undefined
  if (isError) return { kind: 'error', value: String(value) }
  if (typeof value === 'number') return { kind: 'number', value }
  if (typeof value === 'boolean') return { kind: 'bool', value }
  return { kind: 'text', value }
}

/** The resident sheet's saved cells. Rows are indexed on first use, each row's cells on its first lookup. */
export function sheetCellReader(sheet: Pick<XlsxSheet, 'rows'>): CellReader {
  let rows: Map<number, XlsxSheet['rows'][number]> | undefined
  const cells = new Map<number, Map<number, XlsxCell>>()
  return (row, col) => {
    rows ??= new Map(sheet.rows.map(entry => [entry.index, entry]))
    let line = cells.get(row)
    if (!line) {
      line = new Map((rows.get(row)?.cells ?? []).map(cell => [cell.col, cell]))
      cells.set(row, line)
    }
    const cell = line.get(col)
    return cell ? savedOf(cell.value, cell.valueIsError) : undefined
  }
}

/** The sheet area the check reads: the location and the page fields' cells (their bounding box). */
function cellArea(def: PivotTableDefinition): SheetRect | undefined {
  const rects = savedPivotLayout(def)?.rects
  if (!rects) return undefined
  return rects.reduce((box, area) => ({ minRow: Math.min(box.minRow, area.minRow), minCol: Math.min(box.minCol, area.minCol), maxRow: Math.max(box.maxRow, area.maxRow), maxCol: Math.max(box.maxCol, area.maxCol) }))
}

/**
 * The paged path's saved cells: the pivot's area read through the worker's `SourceStore.readRows`, page by page, as
 * the paged viewer reads cells, then served synchronously. A page that does not complete, or an `unavailable` cell,
 * makes the reader `incomplete` (never a throw), so the check reports `cells-missing`. A pivot without a saved
 * layout reads nothing: the check refuses it before reading cells.
 */
export async function pagedCellReader(store: Pick<SourceStore, 'readRows'>, def: PivotTableDefinition, options: ReadOptions): Promise<CellReader> {
  const area = cellArea(def)
  const cells = new Map<string, SavedCell>()
  let incomplete = false
  if (area) {
    const range = { r0: area.minRow, r1: area.maxRow, c0: area.minCol, c1: area.maxCol }
    let cursor: string | undefined
    do {
      const page = await store.readRows(range, cursor, options)
      if (page.completeness !== 'complete' || page.error) { incomplete = true; break }
      for (const item of page.items) for (const cell of item.cells) {
        if (cell.unavailable) { incomplete = true; continue }
        const value = savedOf(cell.value, cell.valueIsError)
        if (value) cells.set(`${cell.row},${cell.col}`, value)
      }
      cursor = page.nextCursor
    } while (cursor !== undefined)
  }
  const reader = (row: number, col: number) => cells.get(`${row},${col}`)
  return incomplete ? Object.assign(reader, { incomplete: true as const }) : reader
}
