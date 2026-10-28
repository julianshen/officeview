/**
 * Sparkline data resolution (spec §5.2, §5.3, §6.3, resident path): turns each parsed sparkline's data range into a
 * render-ready point array, once, after every sheet is parsed and values are final. Painting only reads the cache
 * (`XlsxSheet.resolvedSparklines`), so both painters stay synchronous.
 *
 * Every rule below cites its row of the "Measured constants" table in docs/validation/xlsx-sparklines.md.
 * Geometry (flat series, single point, tiny cell, clipping, slot positions) is the painter's concern; this module only
 * yields values.
 */
import { valueAuthority } from '../view/values'
import type { XlsxCell, XlsxRow, XlsxSheet } from '../types'
import type { SparklineArea, SparklineDiagnostic, XlsxSparkline, XlsxSparklineGroup } from './model'

export interface SparklineAxis {
  /** Plot-width fraction of each plotted point's day-slot centre (dateSlotRule: (d − dmin + ½)·slot), parallel to `values`. */
  positions: number[]
  /** One day's slot width as the same fraction: 1/(dmax − dmin + 1). */
  slot: number
}

export interface ResolvedSparkline {
  host: { row: number; col: number }
  /** Canonical worksheet supplying the data (used to apply viewer-owned filters at paint time). */
  dataSheet: string
  /** Source rows corresponding to `values`, including null/error slots. */
  sourceRows: number[]
  group: XlsxSparklineGroup
  /** Plotted order with hidden points removed — and, when `axis` is set, sorted by date (`dateSortRule`), so the
   * first/last plotted slots are the earliest/latest dated points. `null` is an empty slot (a gap, or bridged: see
   * `group.displayEmptyCellsAs` and `bridged`). Slots are kept for empty and error cells, so positions do not close
   * up (`gapRule`, `spanRule`). */
  values: Array<number | null>
  /** Indices into `values` that are null because the cell is an error: always bridged, whatever `displayEmptyCellsAs`
   * says (`cellValueRule`: #N/A is skipped and bridged even with "gap"). */
  bridged: number[]
  /** Indices into `values` (plotted order; for a date axis the plotted order is date order). `-1` when there are no
   * points. First/last are the first/last plotted slot; high/low include every tied point (`tieRule`) and look at
   * this sparkline's own non-null points, never the scale. */
  firstIndex: number
  lastIndex: number
  highIndices: number[]
  lowIndices: number[]
  /** The scale after individual / group / custom, applied per side. */
  min: number
  max: number
  /** `dateSlotRule` (Task 10): set only when the group has dateAxis and a usable date range, i.e. the window maps
   * dates to day slots of the plot box. Absent = even slots (a dateAxis group without a useable range is not drawn). */
  axis?: SparklineAxis
}
/** row → col → resolved sparkline. */
export type ResolvedSparklines = ReadonlyMap<number, ReadonlyMap<number, ResolvedSparkline>>

const MAX_COLS = 16_384

/** Staging gate (spec §10): features from a stage that has not merged in are not drawn. Task 10 delivered the date
 * axis and rightToLeft (they resolve now; geometry mirrors the x), so only a refused group stays staged out. */
function stagedOut(group: XlsxSparklineGroup): boolean {
  return group.refused !== undefined
}

function diagnose(sheet: Pick<SparklineResolveSheet, 'diagnostics' | 'maxDiagnostics'>, diagnostic: SparklineDiagnostic): void {
  const list = sheet.diagnostics ??= []
  if (list.some(item => item.feature === diagnostic.feature && item.message === diagnostic.message)) return
  if (sheet.maxDiagnostics !== undefined && list.length >= sheet.maxDiagnostics) return
  list.push(diagnostic)
}

/** Lazy per-sheet lookups, built only for sheets a sparkline reads. */
class SheetAccess {
  private rowRecords?: Map<number, XlsxRow[]>
  private hiddenCols?: Set<number>
  private readonly cellRows = new Map<number, Map<number, XlsxCell>>()
  constructor(private readonly sheet: XlsxSheet) {}

  private records(row: number): XlsxRow[] {
    if (!this.rowRecords) {
      this.rowRecords = new Map()
      for (const record of this.sheet.rows) {
        const list = this.rowRecords.get(record.index)
        if (list) list.push(record)
        else this.rowRecords.set(record.index, [record])
      }
    }
    return this.rowRecords.get(row) ?? []
  }
  /** `hiddenRowsRule` / `filterHiddenRule`: an autofilter writes its hidden rows as hidden row records, so one flag covers
   * both. As in the viewport, an index is hidden only when every record for it is (a stray duplicate must not hide data). */
  rowHidden(row: number): boolean {
    const records = this.records(row)
    return records.length > 0 && records.every(record => record.hidden === true)
  }
  /** `hiddenColsRule`. */
  colHidden(col: number): boolean {
    if (!this.hiddenCols) {
      this.hiddenCols = new Set()
      for (const spec of this.sheet.cols) {
        if (!spec.hidden) continue
        for (let c = Math.max(0, spec.min); c <= spec.max && c < MAX_COLS; c++) this.hiddenCols.add(c)
      }
    }
    return this.hiddenCols.has(col)
  }
  cell(row: number, col: number): XlsxCell | undefined {
    let cells = this.cellRows.get(row)
    if (!cells) {
      cells = new Map()
      for (const record of this.records(row)) for (const cell of record.cells) cells.set(cell.col, cell)
      this.cellRows.set(row, cells)
    }
    return cells.get(col)
  }
}

/** The synchronous cell/visibility surface used by the one resolver on both
 * resident sheets and worker-prepared paged snapshots. Worker reads populate
 * this bounded snapshot asynchronously before calling the same resolver. */
export interface SparklineAccess {
  rowHidden(row: number): boolean
  colHidden(col: number): boolean
  cell(row: number, col: number): XlsxCell | undefined
}
export interface SparklineResolveSheet {
  name: string
  sparklineGroups?: XlsxSparklineGroup[]
  mergeRanges: XlsxSheet['mergeRanges']
  diagnostics?: SparklineDiagnostic[]
  /** Worker snapshot cap; resident callers omit it to preserve full diagnostics. */
  maxDiagnostics?: number
  resolvedSparklines?: ResolvedSparklines
  access: SparklineAccess
}

const POINT_EMPTY = Symbol('empty')
const POINT_SKIPPED = Symbol('skipped')
type Point = number | typeof POINT_EMPTY | typeof POINT_SKIPPED

/** One cell's contribution to the series, under the document's value policy ("same values as the grid"). */
function pointOf(cell: XlsxCell | undefined, cached: boolean): Point {
  if (!cell) return POINT_EMPTY
  // The same authority the grid reads: under 'cached' a formula without a saved value is unavailable (shown blank), and
  // under 'engine' `cell.value` already holds the recalculated value.
  const authority = valueAuthority(cell, cached ? 'cached' : 'engine')
  if (authority.unavailable) return POINT_EMPTY
  const { value } = authority
  // `cellValueRule`: #N/A is skipped and bridged, excluded from min/max, even with displayEmptyCellsAs="gap". Other error
  // codes were not measured (the table says so); they are treated like #N/A, the only measured error behaviour.
  if (authority.valueIsError) return POINT_SKIPPED
  if (value === null || value === undefined) return POINT_EMPTY
  if (typeof value === 'number') return Number.isFinite(value) ? value : POINT_SKIPPED
  // `cellValueRule`: TRUE -> 1 (measured). FALSE -> 0 is not measured; it is the closest behaviour (booleans are 1/0).
  if (typeof value === 'boolean') return value ? 1 : 0
  // `cellValueRule`: text -> 0 and takes part in min/max.
  return 0
}

function* cellsInReadingOrder(area: SparklineArea): Generator<[number, number]> {
  // A row range runs left to right, a column range top to bottom (spec §5.2); a single cell is both.
  if (area.r0 === area.r1) for (let c = area.c0; c <= area.c1; c++) yield [area.r0, c]
  else for (let r = area.r0; r <= area.r1; r++) yield [r, area.c0]
}

interface Partial_ { sparkline: XlsxSparkline; host: { row: number; col: number }; dataSheet: string; sourceRows: number[]; values: Array<number | null>; bridged: number[]; axis?: SparklineAxis }

function plotted(
  group: XlsxSparklineGroup, area: SparklineArea, access: SparklineAccess, cached: boolean,
): { values: Array<number | null>; bridged: number[]; sourceRows: number[] } {
  const values: Array<number | null> = []
  const bridged: number[] = []
  const sourceRows: number[] = []
  for (const [row, col] of cellsInReadingOrder(area)) {
    // `hiddenRowsRule`/`hiddenColsRule`/`filterHiddenRule`: hidden points are REMOVED and the rest re-spaced, not gaps.
    // `displayHiddenRule`: displayHidden draws them as if visible.
    if (!group.displayHidden && (access.rowHidden(row) || access.colHidden(col))) continue
    sourceRows.push(row)
    const point = pointOf(access.cell(row, col), cached)
    if (point === POINT_SKIPPED) { bridged.push(values.length); values.push(null) }
    // `zeroRule`: an empty cell is 0 and takes part in min/max. `gapRule`/`spanRule`: otherwise a null slot, which the
    // painter treats per group.displayEmptyCellsAs.
    else if (point === POINT_EMPTY) values.push(group.displayEmptyCellsAs === 'zero' ? 0 : null)
    else values.push(point)
  }
  return { values, bridged, sourceRows }
}

/** A date cell's number (Excel serial), or undefined: `dateCellRule` — any non-number (text, empty, error, boolean,
 * or an absent cell) drops the date's point entirely, and a date never converts to 0 the way `pointOf` maps booleans. */
function dateOf(cell: XlsxCell | undefined, cached: boolean): number | undefined {
  if (!cell) return undefined
  const authority = valueAuthority(cell, cached ? 'cached' : 'engine')
  if (authority.unavailable) return undefined
  const { value } = authority
  if (typeof value !== 'number' || !Number.isFinite(value)) return undefined
  return value
}

function plottedDated(
  group: XlsxSparklineGroup, area: SparklineArea, dateArea: SparklineArea,
  access: SparklineAccess, dateAccess: SparklineAccess, cached: boolean,
): { values: Array<number | null>; bridged: number[]; sourceRows: number[]; axis: SparklineAxis } | undefined {
  // `dateLengthRule`: pair data cell k with date cell k in reading order BEFORE any pruning, keeping the first
  // min(nData, nDates) pairs; `dateCellRule`: a pair whose date cell has no number is dropped entirely (out of
  // values and min/max — with no slot left behind, `gapRule`/`spanRule` do not apply to dates).
  const pairs: Array<{ row: number; col: number; date: number }> = []
  const dateCells = cellsInReadingOrder(dateArea)
  for (const [row, col] of cellsInReadingOrder(area)) {
    const next = dateCells.next()
    if (next.done) break
    const [dateRow, dateCol] = next.value
    const date = dateOf(dateAccess.cell(dateRow, dateCol), cached)
    if (date === undefined) continue // pair dropped: the number is simply not in values
    pairs.push({ row, col, date })
  }
  // `hiddenRowsRule`/`hiddenColsRule`/`filterHiddenRule` (on the DATA cell): the hidden point is removed — and with
  // it its date slot, so the span is over the used dates only.
  const visible = pairs.filter(pair => !(!group.displayHidden && (access.rowHidden(pair.row) || access.colHidden(pair.col))))
  if (!visible.length) return undefined
  // `dateSortRule`: plotted order is date order (stable for tied dates), whatever the cell reading order; the
  // first/last slots therefore fall out as the earliest/latest dated points.
  visible.sort((a, b) => a.date - b.date)
  // `dateSlotRule`: the box holds dmax − dmin + 1 equal day slots over the used dates; the point dated d sits at
  // (d − dmin + ½)·slot of the plot width, one day slot each.
  let dmin = Infinity, dmax = -Infinity
  for (const pair of visible) {
    if (pair.date < dmin) dmin = pair.date
    if (pair.date > dmax) dmax = pair.date
  }
  const slot = 1 / (dmax - dmin + 1)
  const values: Array<number | null> = []
  const bridged: number[] = []
  const sourceRows: number[] = []
  const positions: number[] = []
  for (const pair of visible) {
    sourceRows.push(pair.row)
    const point = pointOf(access.cell(pair.row, pair.col), cached)
    if (point === POINT_SKIPPED) { bridged.push(values.length); values.push(null) }
    // `zeroRule` / `gapRule` / `spanRule` exactly as the undated `plotted`: the data cell keeps its meaning.
    else if (point === POINT_EMPTY) values.push(group.displayEmptyCellsAs === 'zero' ? 0 : null)
    else values.push(point)
    positions.push((pair.date - dmin + 0.5) * slot)
  }
  return { values, bridged, sourceRows, axis: { positions, slot } }
}

function extent(values: ReadonlyArray<number | null>): { min: number; max: number } | undefined {
  let min = Infinity, max = -Infinity
  for (const value of values) {
    if (value === null) continue
    if (value < min) min = value
    if (value > max) max = value
  }
  return min <= max ? { min, max } : undefined
}

/** Fill `XlsxSheet.resolvedSparklines` on every sheet. `policy.cached` is the workbook's value policy (parse.ts: `valuePolicy === 'cached'`). */
export function resolveSparklineDataSheets(sheets: SparklineResolveSheet[], policy: { cached: boolean }): void {
  const byName = new Map<string, SparklineResolveSheet>()
  for (const sheet of sheets) {
    const key = sheet.name.toLowerCase()
    if (!byName.has(key)) byName.set(key, sheet)
  }
  for (const sheet of sheets) {
    delete sheet.resolvedSparklines
    if (!sheet.sparklineGroups?.length) continue
    const resolved = new Map<number, Map<number, ResolvedSparkline>>()
    for (const group of sheet.sparklineGroups) {
      if (stagedOut(group)) continue
      const parts: Partial_[] = []
      let droppedMember = group.incomplete === true
      for (const sparkline of group.sparklines) {
        // Parse already diagnosed an unreadable host or reference; nothing to draw.
        if (!sparkline.host || !sparkline.data) { droppedMember = true; continue }
        const merge = sheet.mergeRanges.find(range => sparkline.host!.row >= range.minRow && sparkline.host!.row <= range.maxRow && sparkline.host!.col >= range.minCol && sparkline.host!.col <= range.maxCol)
        if (merge && (sparkline.host.row !== merge.minRow || sparkline.host.col !== merge.minCol)) {
          diagnose(sheet, { kind: 'unsupported', feature: 'sparkline-merged-host', message: 'Sparkline host is covered by a merged cell and is not the merge anchor' })
          droppedMember = true
          continue
        }
        const area = sparkline.data
        const target = area.sheet === undefined ? sheet : byName.get(area.sheet.toLowerCase())
        if (!target) {
          diagnose(sheet, { kind: 'unsupported', feature: 'sparkline-data-reference', message: `Sparkline data sheet not found: ${area.sheet?.slice(0, 80)}` })
          droppedMember = true
          continue
        }
        // Defensive: a hand-built model may carry what the parser already rejects (same feature and message, so it dedupes).
        if (area.r0 !== area.r1 && area.c0 !== area.c1) {
          diagnose(sheet, { kind: 'unsupported', feature: 'sparkline-data-reference', message: 'Sparkline data reference must be one-dimensional' })
          droppedMember = true
          continue
        }
        // `dateMissingRule` (Task 10): with dateAxis the sparkline needs a usable date range — one the parse kept,
        // one-dimensional, on a resolvable sheet (ruling: dateAxisOff/false and an unusable range draw nothing and
        // raise no diagnostic, unlike the data range).
        let axis: SparklineAxis | undefined
        if (group.dateAxis) {
          const dateArea = group.dateRange
          if (!dateArea || (dateArea.r0 !== dateArea.r1 && dateArea.c0 !== dateArea.c1)) continue
          const dateTarget = dateArea.sheet === undefined ? sheet : byName.get(dateArea.sheet.toLowerCase())
          if (!dateTarget) continue
          const dated = plottedDated(group, area, dateArea, target.access, dateTarget.access, policy.cached)
          // No usable date pair, or no numeric point after pairing: nothing to draw or scale, silently. Like the
          // undated singlePointRule, this is not a dropped group member.
          if (!dated || !extent(dated.values)) continue
          axis = dated.axis
          parts.push({ sparkline, host: sparkline.host, dataSheet: target.name, sourceRows: dated.sourceRows, values: dated.values, bridged: dated.bridged, axis })
        } else {
          const { values, bridged, sourceRows } = plotted(group, area, target.access, policy.cached)
          // `singlePointRule`: a range with no drawable point draws nothing, and contributes nothing to a group scale.
          if (!extent(values)) continue
          parts.push({ sparkline, host: sparkline.host, dataSheet: target.name, sourceRows, values, bridged })
        }
      }
      const groupScaled = group.minAxisType === 'group' || group.maxAxisType === 'group'
      if (groupScaled && droppedMember) {
        diagnose(sheet, { kind: 'unsupported', feature: 'sparkline-group-incomplete', message: 'Group-scaled sparkline group omitted because one or more members could not be resolved' })
        continue
      }
      if (!parts.length) continue
      // `groupScaleRule`, measured (docs/validation/xlsx-sparklines.md; line-scale-group, and col-scale-no-zero
      // rows 4–5 for columns): min and max are shared across every sparkline of the group — the min/max over every
      // plotted (non-hidden, non-null) point of every sparkline — and manualMin/manualMax are ignored under group.
      const groupExtent = group.minAxisType === 'group' || group.maxAxisType === 'group' ? extent(parts.flatMap(part => part.values)) : undefined
      for (const part of parts) {
        const own = extent(part.values)!
        // `customScaleRule`: min and max are independent (custom on one side, individual or group on the other).
        // `individualScaleRule`: each sparkline uses its own min and max.
        const min = group.minAxisType === 'custom' ? group.manualMin! : group.minAxisType === 'group' ? groupExtent!.min : own.min
        const max = group.maxAxisType === 'custom' ? group.manualMax! : group.maxAxisType === 'group' ? groupExtent!.max : own.max
        if ((group.minAxisType === 'custom' || group.maxAxisType === 'custom') && min >= max) {
          diagnose(sheet, { kind: 'unsupported', feature: 'sparkline-custom-scale', message: 'Custom sparkline scale must have an effective minimum below its maximum' })
          continue
        }
        // `tieRule`: every tied point is high (low). Flat series therefore make every point both. High/low come from the
        // sparkline's own values, so an all-negative series has its largest negative as the high (LibreOffice tdf#152929).
        const highIndices: number[] = [], lowIndices: number[] = []
        part.values.forEach((value, i) => {
          if (value === null) return
          if (value === own.max) highIndices.push(i)
          if (value === own.min) lowIndices.push(i)
        })
        let row = resolved.get(part.host.row)
        if (!row) resolved.set(part.host.row, row = new Map())
        if (row.has(part.host.col)) {
          diagnose(sheet, { kind: 'unsupported', feature: 'sparkline-duplicate-host', message: 'Multiple sparklines target one host cell; the first is kept' })
          continue
        }
        row.set(part.host.col, { host: part.host, dataSheet: part.dataSheet, sourceRows: part.sourceRows, group, values: part.values, bridged: part.bridged, firstIndex: 0, lastIndex: part.values.length - 1, highIndices, lowIndices, min, max, axis: part.axis })
      }
    }
    if (resolved.size) sheet.resolvedSparklines = resolved
  }
}

/** Resident entry point retained for parser callers. */
export function resolveWorkbookSparklines(sheets: XlsxSheet[], policy: { cached: boolean }): void {
  const access = new Map<XlsxSheet, SheetAccess>()
  const adapters: SparklineResolveSheet[] = sheets.map(sheet => {
    let found = access.get(sheet)
    if (!found) access.set(sheet, found = new SheetAccess(sheet))
    return {
      name: sheet.name, sparklineGroups: sheet.sparklineGroups, mergeRanges: sheet.mergeRanges,
      diagnostics: sheet.diagnostics as SparklineDiagnostic[] | undefined,
      resolvedSparklines: sheet.resolvedSparklines, access: found,
    }
  })
  resolveSparklineDataSheets(adapters, policy)
  for (let i = 0; i < sheets.length; i++) {
    if (adapters[i].resolvedSparklines) sheets[i].resolvedSparklines = adapters[i].resolvedSparklines
    else delete sheets[i].resolvedSparklines
    if (adapters[i].diagnostics) sheets[i].diagnostics = adapters[i].diagnostics as XlsxSheet['diagnostics']
  }
}
