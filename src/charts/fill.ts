import { sourceFormat } from './format'
import { prepareChartAxes } from './prepare'
/**
 * Chart data fill (lazy chunk): uncached references resolved from worksheet
 * cells, defined names, paged-reader lines and the decimal scanner. Moved
 * from `src/xlsx/chart.ts` (A0a) behaviour-identically.
 */
import { CHART_ANALYTIC_POINT_LIMIT, CHART_LABEL_LIMIT, CHART_POINT_LIMIT, CHART_TEXT_BYTES_LIMIT, parseChartRange } from './parse'
import type { ChartModel, ChartPoints, ChartRange, ChartSeries } from './model'
import { POINT_UNAVAILABLE, PointWriter, pointBytes, stringBytes, withHidden } from './points'
import { ChartRefusal, adoptChartModel, chartModelOwner, modelOwner, retainWithTarget, succeedChartModel } from './own'
import type { ChartWorkspace } from './workspace'
import { chartEpoch, convertSerial, isDateFormat, type ChartEpoch } from './dates'

export type ChartCellValue = string | number | boolean | null | undefined
/**
 * Cell lookup for filling uncached references: `sheet` is the resolved sheet
 * name. The point state comes from the cell itself, the same way on every path
 * (spec 5.8): `error` keeps the error code, `unavailable` marks a formula with
 * no saved value, a missing cell is blank.
 */
export interface ChartCell {
  formatCode?: string
  value: ChartCellValue
  /** The cell's formatted text (number/date format), as Excel shows it. */ text?: string
  /** The value could not be read in full (a numeric form past the scanner's exact range). */ truncated?: boolean
  /** The value's numeric reading when the reader computed it without building the text (null: not a number). */ number?: number | null
  /** An error cell's code (`#N/A`, `#DIV/0!`, ...). */ error?: string
  /** The cell has no value to show (a formula without a saved value under the saved-value policy). */ unavailable?: boolean
}
/** The formatter's allocation observer (cellFormat `FormatAllocation`): each working buffer is reported before it
 * is allocated; `allocate` throws to refuse it. */
export interface ChartTextAllocation {
  allocate(bytes: number): void; release(bytes: number): void
  /** Bytes the formatter publishes into a structure that outlives it (`targets[0]`, e.g. a sheet diagnostic). */
  retain?(targets: readonly object[], bytes: number, commit: () => void, kind?: 'diagnostic'): void
}
/** `text`: the caller needs the formatted label (categories, names); values never format. A label's formatting
 * reports its working buffers through `allocation` (the chart workspace admits them, or refuses by throwing). */
export type ChartCellLookup = (sheet: string, row: number, col: number, text: boolean, allocation?: ChartTextAllocation) => ChartCell | undefined
/** Whether a worksheet cell sits in a hidden row or column (provenance for `plotVisOnly`). */
export type ChartHiddenLookup = (sheet: string, row: number, col: number) => boolean

/** The one line of cells a chart consumes from a range (a 2-D range keeps its
 * last row/column, the level nearest the data), bounded by `limit` (the display
 * bound, or the analytic bound for an analytic series' numeric roles).
 * Readers and the fill use this same sequence. */
export interface ChartLine { sheet: string; vertical: boolean; fixed: number; start: number; count: number; truncated: boolean; text: boolean }
export function chartLine(range: ChartRange, ownerSheet: string, text: boolean, limit = CHART_POINT_LIMIT): ChartLine {
  const rows = range.r1 - range.r0 + 1, cols = range.c1 - range.c0 + 1
  const vertical = cols === 1 || rows >= cols
  const length = vertical ? rows : cols
  return { sheet: range.sheet ?? ownerSheet, vertical, fixed: vertical ? range.c1 : range.r1, start: vertical ? range.r0 : range.c0, count: Math.min(length, limit), truncated: length > limit, text }
}
const lineLength = (range: ChartRange) => { const rows = range.r1 - range.r0 + 1, cols = range.c1 - range.c0 + 1; return cols === 1 || rows >= cols ? rows : cols }
/**
 * The acquisition bound of a series' numeric roles (spec 7): an analytic series
 * is acquired in full up to CHART_ANALYTIC_POINT_LIMIT; a reference longer than
 * that refuses the analytic element (never a prefix) and the series is acquired
 * for display only. Readers and the fill decide it the same way.
 */
export function seriesLimit(series: ChartSeries): number {
  // An analytic ChartEx family's series is always acquired to the analytic bound; a longer reference refuses the
  // whole chart in the fill (never a display prefix).
  if (series.analytic === 'chart') return CHART_ANALYTIC_POINT_LIMIT
  if (series.analytic !== true) return CHART_POINT_LIMIT
  const ranges = [series.pending?.values, series.pending?.xValues].filter((range): range is ChartRange => Boolean(range && !range.definedName))
  return ranges.some(range => lineLength(range) > CHART_ANALYTIC_POINT_LIMIT) ? CHART_POINT_LIMIT : CHART_ANALYTIC_POINT_LIMIT
}
/** The category acquisition bound of a series: an analytic ChartEx series pairs every value with its category. */
export const categoryLimit = (series: ChartSeries): number => series.analytic === 'chart' ? CHART_ANALYTIC_POINT_LIMIT : CHART_POINT_LIMIT
export const CHART_TOO_MANY_POINTS = 'This chart has too many data points to summarize'
export const CHART_DATA_NOT_AVAILABLE = "This chart's data is not available"
/** A whole-chart analytic refusal (spec 4.3, 7; chartex-map §2.4): the chart is refused with `message`. */
export class ChartAnalyticRefusal extends ChartRefusal {}
/** Writes the point state of one filled cell (spec 5.8); `onTruncated` hears a value the reader could not read in full.
 * `map` converts a date role's serial into the chart epoch (undefined: no counterpart, the point is unavailable). */
function writeCell(out: PointWriter, i: number, cell: ChartCell | undefined, onTruncated: () => void, map?: (v: number) => number | undefined): void {
  if (!cell) return // blank
  if (cell.unavailable) { out.unavailable(i); return }
  if (cell.error !== undefined) { out.error(i, cell.error); return }
  // A numeric form the scanner cannot read exactly has no value, and the chart says it is truncated.
  if (cell.truncated) { onTruncated(); out.unavailable(i); return }
  const number = cell.number !== undefined ? cell.number ?? undefined : numberOf(cell.value)
  if (number === undefined) return
  const v = map ? map(number) : number
  if (v === undefined) out.unavailable(i); else out.value(i, v, cell.formatCode)
}
/** The cell coordinates of point `i` on a line. */
export const linePoint = (line: ChartLine, i: number): { row: number; col: number } => line.vertical ? { row: line.start + i, col: line.fixed } : { row: line.fixed, col: line.start + i }

export function chartHasPending(model: ChartModel): boolean {
  return Boolean(model.categoryPending) || model.groups.some(g => g.series.some(s => s.pending))
}
/** A cached role keeps its reference for plotVisOnly provenance (spec 5.8): such a chart is visited by the fill too. */
export function chartHasRefs(model: ChartModel): boolean {
  return Boolean(model.categoryRef) || model.groups.some(g => g.series.some(s => s.refs))
}
const numberOf = (value: ChartCellValue): number | undefined => typeof value === 'number' ? (Number.isFinite(value) ? value : undefined)
  : typeof value === 'string' ? scanChartNumber(value.length, i => value.charCodeAt(i)) : undefined
const rawText = (value: ChartCellValue): string => value === null || value === undefined ? '' : typeof value === 'number' ? String(Number(value.toPrecision(15))) : String(value)
const asText = (cell: ChartCell | undefined): string => cell?.text ?? rawText(cell?.value)

/** Resolves a defined name seen from a sheet to its formula (sheet scope first, then workbook). */
export type ChartNameResolver = (name: string, scopeSheet: string) => string | undefined
export interface ChartDefinedName { name: string; source: string; localSheetIndex?: number }
export function chartNameResolver(names: readonly ChartDefinedName[], sheets: ReadonlyArray<{ name: string; workbookIndex?: number }>): ChartNameResolver {
  return (name, scopeSheet) => {
    const upper = name.toUpperCase()
    const scope = sheets.find(sheet => sheet.name.toLowerCase() === scopeSheet.toLowerCase())
    const candidates = names.filter(entry => entry.name.toUpperCase() === upper)
    const local = scope?.workbookIndex !== undefined ? candidates.find(entry => entry.localSheetIndex === scope.workbookIndex) : undefined
    return (local ?? candidates.find(entry => entry.localSheetIndex === undefined))?.source
  }
}
/** Replace defined-name references by the ranges they name; a name that is
 * missing or not a single area turns the chart into a placeholder. */
export function resolveChartNames(model: ChartModel, ownerSheet: string, resolve: ChartNameResolver | undefined): ChartModel {
  let unsupported = model.unsupported
  const fix = (range: ChartRange | undefined): ChartRange | undefined => {
    if (!range?.definedName) return range
    const formula = resolve?.(range.definedName, range.sheet ?? ownerSheet)
    const target = formula ? parseChartRange(formula.replace(/^=/, '')) : undefined
    if (target && !target.definedName) return target
    unsupported ??= `Chart data name “${range.definedName}” ${formula ? 'is not a single worksheet range' : 'is not defined in this workbook'}`
    return undefined
  }
  // A cached role's reference only gives provenance: a name that does not resolve drops it, the cache stays.
  const ref = (range: ChartRange | undefined): ChartRange | undefined => {
    if (!range?.definedName) return range
    const formula = resolve?.(range.definedName, range.sheet ?? ownerSheet)
    const target = formula ? parseChartRange(formula.replace(/^=/, '')) : undefined
    return target && !target.definedName ? target : undefined
  }
  if (!chartHasPending(model) && !chartHasRefs(model)) return model
  const categoryPending = fix(model.categoryPending), categoryRef = ref(model.categoryRef)
  const roles = <T extends Record<string, ChartRange | undefined>>(table: T | undefined, map: (range: ChartRange | undefined) => ChartRange | undefined): T | undefined => {
    if (!table) return undefined
    const out = Object.fromEntries(Object.entries(table).map(([key, range]) => [key, map(range)]).filter(([, range]) => range)) as T
    return Object.keys(out).length ? out : undefined
  }
  const groups = model.groups.map(group => ({ ...group, series: group.series.map(series => {
    if (!series.pending && !series.refs) return series
    const pending = roles(series.pending, fix), refs = roles(series.refs, ref)
    const { pending: _old, refs: _refs, ...rest } = series
    void _old, void _refs
    return { ...rest, ...(pending ? { pending } : {}), ...(refs ? { refs } : {}) }
  }) }))
  const { categoryPending: _cat, categoryRef: _ref, ...rest } = model
  void _cat, void _ref
  const next = { ...rest, groups, ...(categoryPending ? { categoryPending } : {}), ...(categoryRef ? { categoryRef } : {}), ...(unsupported ? { unsupported } : {}) }
  succeedChartModel(model, next)
  return next
}

/** What a fill needs beside the cells: hidden-cell provenance and the chart workspace that owns new storage. */
/** `epoch`: the workbook's date system, which filled values carry (spec 5.1); date roles are converted to the chart epoch. */
export interface ChartFillOptions { hidden?: ChartHiddenLookup; workspace?: ChartWorkspace; epoch?: ChartEpoch }
/**
 * Fill uncached references from worksheet cells (returns a new model that takes
 * over the old one's charge). Every cell is written straight into the role's
 * compact storage (no object per point), the line order is kept (no
 * compaction), and every role with a worksheet reference (filled, or cached
 * with its reference kept) records the indices of its own cells in hidden rows
 * or columns (plotVisOnly provenance, applied per role at draw time).
 *
 * All new storage is admitted to the model's chart workspace (its owner, else
 * `options.workspace`) before it is built. A refusal returns the model as a
 * refused placeholder and refunds what the fill charged.
 */
export function fillChartRanges(model: ChartModel, ownerSheet: string, lookup: ChartCellLookup, options: ChartFillOptions = {}): ChartModel {
  const { hidden } = options
  const owner = chartModelOwner(model) ?? (options.workspace ? modelOwner(options.workspace) : undefined)
  const fillTag = owner ? `${owner.tag}:fill` : '', charged = { bytes: 0 }
  const charge = (bytes: number) => {
    if (!owner) return
    if (!owner.workspace.reserve(bytes, fillTag)) throw new ChartRefusal()
    charged.bytes += bytes
  }
  // A label or name: the formatter's capped output is admitted before any text is formatted or materialized;
  // every working buffer the formatter reports beyond it is admitted (at the live peak) before it is allocated,
  // and a refusal stops the formatter there. The charge is then trimmed to the text actually retained.
  const labelTag = owner ? `${owner.tag}:label` : ''
  const label = (produce: (allocation?: ChartTextAllocation) => string, keepEmpty = true): string => {
    if (!owner) return produce()
    if (!owner.workspace.reserve(LABEL_BOUND_BYTES, labelTag)) throw new ChartRefusal()
    // `admitted`: what the label tag owns (reserved, less what moved to a longer-lived owner); `live`: the
    // formatter's current working set, always covered by `admitted`.
    let live = 0, admitted = LABEL_BOUND_BYTES
    const allocation: ChartTextAllocation = {
      allocate: bytes => {
        live += bytes
        if (live <= admitted) return
        if (!owner.workspace.reserve(live - admitted, labelTag)) throw new ChartRefusal()
        admitted = live
      },
      release: bytes => { live -= bytes },
      // A published diagnostic outlives the label: its admitted bytes move to the sheet's owner (charts/own.ts),
      // kept beyond the label's trim and a later refusal, refunded with the document.
      retain: (targets, bytes, commit) => {
        const target = targets[0]
        if (target && bytes > 0 && bytes <= admitted) { retainWithTarget(target, owner.workspace, labelTag, bytes); admitted -= bytes; live -= bytes }
        commit()
      },
    }
    try {
      const text = produce(allocation)
      const kept = Math.min(text || keepEmpty ? stringBytes(text) : 0, admitted)
      if (kept) { owner.workspace.transfer(labelTag, fillTag, kept); charged.bytes += kept }
      return text
    } finally { owner.workspace.release(labelTag) }
  }
  try {
    const next = prepareChartAxes(fillModel(model, ownerSheet, lookup, hidden, charge, label, options.epoch ?? model.workbookEpoch), owner)
    if (owner) {
      owner.workspace.transfer(fillTag, owner.tag, charged.bytes)
      if (chartModelOwner(model)) succeedChartModel(model, next); else adoptChartModel(next, owner)
    }
    return next
  } catch (error) {
    if (owner) owner.workspace.release(fillTag)
    if (!(error instanceof ChartRefusal)) throw error
    const reason = error instanceof ChartAnalyticRefusal ? error.message : 'This chart needs more memory than is available for charts'
    const refused = { ...model, preparedAxes: undefined, unsupported: reason, refused: true as const }
    succeedChartModel(model, refused)
    return refused
  }
}
/** The bound admitted before a label is built: the formatter's capped output (CHART_LABEL_LIMIT + 1 units,
 * one past the limit marks truncation), which also bounds the retained label. */
const LABEL_BOUND_BYTES = stringBytes(' '.repeat(CHART_LABEL_LIMIT + 1))
function fillModel(model: ChartModel, ownerSheet: string, lookup: ChartCellLookup, hidden: ChartHiddenLookup | undefined, charge: (bytes: number) => void, label: (produce: (allocation?: ChartTextAllocation) => string, keepEmpty?: boolean) => string, workbookEpoch?: ChartEpoch): ChartModel {
  let truncated = model.truncated ?? false
  // Epoch normalization (spec 5.1): filled values carry the workbook's epoch; only date roles (date-axis categories,
  // values on a date-formatted axis) are converted into the chart epoch, through the civil model. A serial without
  // a counterpart is unavailable (chart-date-unrepresentable). Saved data is already in the chart epoch.
  const epoch = chartEpoch(model.epoch, workbookEpoch), source = workbookEpoch ?? epoch
  let unrepresentable = 0
  const toChart = source === epoch ? undefined : (v: number) => { const out = convertSerial(v, source, epoch); if (out === undefined) unrepresentable++; return out }
  const axisById = (id: string | undefined) => id === undefined ? undefined : model.axes.find(axis => axis.id === id)
  const dateFormatted = (id: string | undefined) => isDateFormat(axisById(id)?.formatCode)
  const dateCategories = model.groups.some(g => g.binding.kind === 'cartesian' && axisById(g.binding.category)?.kind === 'date')
  const sheetOf = (range: ChartRange) => range.sheet ?? ownerSheet
  /** A role's own provenance: the sorted indices of its cells in hidden rows or columns (admitted, then built). */
  const provenance = (line: ChartLine): Uint32Array | undefined => {
    if (!hidden) return undefined
    let count = 0
    for (let i = 0; i < line.count; i++) { const at = linePoint(line, i); if (hidden(line.sheet, at.row, at.col)) count++ }
    if (!count) return undefined
    charge(count * 4)
    const out = new Uint32Array(count)
    for (let i = 0, k = 0; i < line.count; i++) { const at = linePoint(line, i); if (hidden(line.sheet, at.row, at.col)) out[k++] = i }
    return out
  }
  // The chart-wide label budget (spec 7: 4 MiB) of analytic ChartEx categories: labels past it are dropped (empty),
  // the chart is marked truncated and diagnosed. Cached labels already in the model count first.
  let labelBytes = 0, labelsDropped = 0, labelsCut = 0
  for (const group of model.groups) for (const series of group.series) if (series.analytic === 'chart') for (const text of series.categories ?? []) labelBytes += stringBytes(text)
  let categories = model.categories, categoryHidden = model.categoryHidden, categoryValues = model.categoryValues
  if (model.categoryPending) {
    const line = chartLine(model.categoryPending, sheetOf(model.categoryPending), true)
    truncated ||= line.truncated
    const labels: string[] = []
    // A date axis keeps its categories' numbers too (the same cells, never re-read from the formatted labels).
    let numbers: PointWriter | undefined
    if (dateCategories) { charge(pointBytes(line.count)); numbers = new PointWriter(line.count, 'lazy', charge) }
    for (let i = 0; i < line.count; i++) {
      const at = linePoint(line, i)
      labels.push(label(allocation => {
        const cell = lookup(line.sheet, at.row, at.col, true, allocation)
        // A text cell stays text even when it looks like a number (spec 5.1): its category is not a date.
        if (numbers && typeof cell?.value !== 'string') writeCell(numbers, i, cell, () => { truncated = true }, toChart)
        const text = asText(cell)
        if (text.length <= CHART_LABEL_LIMIT) return text
        truncated = true
        return text.slice(0, CHART_LABEL_LIMIT)
      }))
    }
    categories = labels
    categoryHidden = provenance(line)
    if (numbers) categoryValues = numbers.done()
  } else if (model.categoryRef && categories.length) categoryHidden = provenance(chartLine(model.categoryRef, sheetOf(model.categoryRef), true, categories.length))
  const groups = model.groups.map(group => ({
    ...group,
    series: group.series.map(series => {
      if (!series.pending && !series.refs) return series
      const { pending = {}, ...rest } = series
      const next: ChartSeries = { ...rest }
      const limit = seriesLimit(series), whole = series.analytic === 'chart'
      if (series.analytic === true && limit === CHART_POINT_LIMIT) next.analytic = 'refused'
      const binding = group.binding
      const valueDates = dateFormatted(binding.kind === 'cartesian' ? binding.value : binding.kind === 'xy' ? binding.y : undefined) ? toChart : undefined
      const xDates = dateFormatted(binding.kind === 'xy' || binding.kind === 'cartesian' ? binding.x : undefined) ? toChart : undefined
      const fill = (range: ChartRange, map?: (v: number) => number | undefined): ChartPoints => {
        const line = chartLine(range, sheetOf(range), false, limit)
        // An analytic ChartEx input longer than the analytic bound refuses the chart before anything is read.
        if (whole && line.truncated) throw new ChartAnalyticRefusal(CHART_TOO_MANY_POINTS)
        // Display truncation (spec 7): an analytic series keeps every acquired point; the drawing is capped.
        if (!whole) truncated ||= line.truncated || line.count > CHART_POINT_LIMIT
        charge(pointBytes(line.count))
        const out = new PointWriter(line.count, 'lazy', charge)
        for (let i = 0; i < line.count; i++) { const at = linePoint(line, i); writeCell(out, i, lookup(line.sheet, at.row, at.col, false), () => { truncated = true }, map) }
        return withHidden(out.done(), provenance(line))
      }
      // A cached role keeps its values; its reference adds only its own provenance.
      const cached = (points: ChartPoints, range: ChartRange) => points.length ? withHidden(points, provenance(chartLine(range, sheetOf(range), false, points.length))) : points
      if (pending.values) next.values = fill(pending.values, valueDates)
      else if (series.refs?.values) next.values = cached(series.values, series.refs.values)
      // An analytic family with a required value point unavailable is refused (chartex-map §2.4).
      if (whole && pending.values && next.values.states.includes(POINT_UNAVAILABLE)) throw new ChartAnalyticRefusal(CHART_DATA_NOT_AVAILABLE)
      if (pending.xValues) next.xValues = fill(pending.xValues, xDates)
      else if (series.refs?.xValues && series.xValues) next.xValues = cached(series.xValues, series.refs.xValues)
      if(pending.categories){
        const line=chartLine(pending.categories,sheetOf(pending.categories),true,categoryLimit(series))
        if (whole && line.truncated) throw new ChartAnalyticRefusal(CHART_TOO_MANY_POINTS)
        truncated ||= line.truncated
        charge(line.count*8)
        next.categories=Array.from({length:line.count},(_,i)=>{const at=linePoint(line,i);return label(allocation=>{
          const text=asText(lookup(line.sheet,at.row,at.col,true,allocation))
          if (!whole) return text.slice(0,CHART_LABEL_LIMIT)
          const kept=text.length>CHART_LABEL_LIMIT?(labelsCut++,text.slice(0,CHART_LABEL_LIMIT)):text
          // The formatter's output is bounded (LABEL_BOUND_BYTES) before it is built; past the chart budget it is dropped.
          if (labelBytes+stringBytes(kept)>CHART_TEXT_BYTES_LIMIT) { labelsDropped++; return '' }
          labelBytes+=stringBytes(kept)
          return kept
        })})
      }
      if (pending.name) {
        const name = pending.name
        const text = label(allocation => {
          const full = asText(lookup(sheetOf(name), name.r0, name.c0, true, allocation))
          if (full.length > CHART_LABEL_LIMIT) truncated = true
          return full.slice(0, CHART_LABEL_LIMIT)
        }, false)
        if (text) next.name = text
      }
      if (next.pointColors && next.values.length > next.pointColors.length) {
        const first = next.pointColors
        next.pointColors = Array.from({ length: Math.min(256, next.values.length) }, (_, i) => first[i] ?? first[i % first.length])
      }
      return next
    }),
  }))
  const diagnostics = model.diagnostics ? [...model.diagnostics] : []
  const note = (code: string, reason: string) => { charge(stringBytes(code) + stringBytes(reason)); diagnostics.push({ code, reason }) }
  if (labelsCut) { truncated = true; note('chart-label-truncated', `${labelsCut} ${labelsCut === 1 ? 'label was' : 'labels were'} cut to ${CHART_LABEL_LIMIT} characters`) }
  if (labelsDropped) { truncated = true; note('chart-label-budget', `${labelsDropped} ${labelsDropped === 1 ? 'label was' : 'labels were'} dropped past the chart's 4 MiB label budget`) }
  if (unrepresentable) {
    const reason = `${unrepresentable} date ${unrepresentable === 1 ? 'value has' : 'values have'} no counterpart in the ${epoch} date system`
    charge(stringBytes(reason))
    diagnostics.push({ code: 'chart-date-unrepresentable', reason })
  }
  const { categoryPending, categoryHidden: _hidden, categoryValues: _values, diagnostics: _diagnostics, ...rest } = model
  void categoryPending, void _hidden, void _values, void _diagnostics
  return { ...rest, categories, ...(categoryValues ? { categoryValues } : {}), ...(categoryHidden ? { categoryHidden } : {}), groups, ...(truncated ? { truncated: true } : {}), ...(diagnostics.length ? { diagnostics } : {}) }
}

/** Resident workbooks: fill every uncached chart reference from parsed cells,
 * after evaluation and semantics. Only the cells on the consumed lines are
 * visited (rows by binary search, cells by binary search), and labels are
 * formatted on demand: the workspace is bounded by the chart point limits. */
export function fillWorkbookCharts<Cell extends { row: number; col: number; value: ChartCellValue; valueIsError?: boolean; style?: {numFmtId:number;formatCode?:string} }, Sheet extends { name: string; workbookIndex?: number; rows: ReadonlyArray<{ index: number; hidden?: boolean; cells: ReadonlyArray<Cell> }>; cols?: ReadonlyArray<{ min: number; max: number; hidden?: boolean }>; drawings?: XlsxDrawingLike[]; semantics?: { dateSystem?: ChartEpoch } }>(
  sheets: readonly Sheet[], options: { names?: readonly ChartDefinedName[]; text?: (sheet: Sheet, cell: Cell, allocation?: ChartTextAllocation) => string; available?: (cell: Cell) => boolean; workspace?: ChartWorkspace } = {},
): void {
  const sheetByName = (name: string) => sheets.find(candidate => candidate.name.toLowerCase() === name.toLowerCase())
  const sortedRows = new WeakMap<object, boolean>()
  const ascending = <T>(items: ReadonlyArray<T>, key: (item: T) => number) => { for (let i = 1; i < items.length; i++) if (key(items[i - 1]) >= key(items[i])) return false; return true }
  const search = <T>(items: ReadonlyArray<T>, key: (item: T) => number, target: number, sorted: boolean): T | undefined => {
    if (!sorted) return items.find(item => key(item) === target)
    let lo = 0, hi = items.length - 1
    while (lo <= hi) { const mid = (lo + hi) >> 1, k = key(items[mid]); if (k === target) return items[mid]; if (k < target) lo = mid + 1; else hi = mid - 1 }
    return undefined
  }
  const isSorted = <T>(owner: object, items: ReadonlyArray<T>, key: (item: T) => number) => { let known = sortedRows.get(owner); if (known === undefined) { known = ascending(items, key); sortedRows.set(owner, known) } return known }
  const rowOf = (sheet: Sheet, rowIndex: number) => search(sheet.rows, row => row.index, rowIndex, isSorted(sheet.rows, sheet.rows, row => row.index))
  const lookup: ChartCellLookup = (name, rowIndex, col, wantText, allocation) => {
    const sheet = sheetByName(name)
    if (!sheet) return undefined
    const row = rowOf(sheet, rowIndex)
    const cell = row && search(row.cells, cell => cell.col, col, isSorted(row.cells, row.cells, cell => cell.col))
    if (!cell) return undefined
    if (options.available && !options.available(cell)) return { value: undefined, unavailable: true }
    let text: string | undefined
    // A refused formatter allocation refuses the chart (never a silent fallback); other formatter failures read raw.
    if (wantText && options.text) { try { text = options.text(sheet, cell, allocation) } catch (error) { if (error instanceof ChartRefusal) throw error; text = undefined } }
    return { value: cell.value, formatCode:sourceFormat(cell.style), ...(cell.valueIsError === true ? { error: String(cell.value) } : {}), ...(text !== undefined ? { text } : {}) }
  }
  const hidden: ChartHiddenLookup = (name, rowIndex, col) => {
    const sheet = sheetByName(name)
    return Boolean(sheet && (rowOf(sheet, rowIndex)?.hidden || sheet.cols?.some(spec => spec.hidden && spec.min <= col && col <= spec.max)))
  }
  const resolve = chartNameResolver(options.names ?? [], sheets)
  for (const sheet of sheets) stampWorkbookEpoch(sheet.drawings, sheet.semantics?.dateSystem)
  // The workbook's date system (every sheet shares it): filled date roles carry it (spec 5.1).
  const epoch = sheets[0]?.semantics?.dateSystem
  for (const sheet of sheets) {
    try { fillDrawings(sheet.drawings, sheet.name, lookup, resolve, { hidden, workspace: options.workspace, ...(epoch ? { epoch } : {}) }) }
    catch { /* a chart that cannot be filled keeps its placeholder state; the workbook stays usable */ }
  }
}
/**
 * Carries a 1904 workbook's epoch on every chart model of a drawing tree (spec 5.1 precedence 2), cached-only
 * charts included, so the transferable model resolves its chart epoch the same way in every paint realm. A new
 * model object takes over the old one's charge. A 1900 workbook stamps nothing (absent means 1900).
 */
export function stampWorkbookEpoch(drawings: XlsxDrawingLike[] | undefined, epoch: ChartEpoch | undefined): void {
  if (epoch !== '1904') return
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && content.chart.workbookEpoch !== epoch) {
      let next: ChartModel
      try { next = prepareChartAxes({ ...content.chart, workbookEpoch: epoch }, chartModelOwner(content.chart)) }
      catch (error) {
        if (!(error instanceof ChartRefusal)) throw error
        next = { ...content.chart, preparedAxes: undefined, workbookEpoch: epoch, unsupported: 'This chart needs more memory than is available for charts', refused: true }
      }
      succeedChartModel(content.chart, next)
      content.chart = next
    }
    if (node.children) stampWorkbookEpoch(node.children, epoch)
  }
}
/** Resolve defined names of a drawing tree in place (paged readers resolve before reading cells). */
export function resolveDrawingNames(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, resolve: ChartNameResolver | undefined): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart) content.chart = resolveChartNames(content.chart, ownerSheet, resolve)
    if (node.children) resolveDrawingNames(node.children, ownerSheet, resolve)
  }
}
/** Whether a drawing tree names any defined name in its chart references. */
export function drawingsNeedNames(drawings: XlsxDrawingLike[] | undefined): boolean {
  return pendingRanges(drawings, '', [], true).some(range => range.definedName)
}
export interface XlsxDrawingLike { content?: { kind: string }; children?: XlsxDrawingLike[] }
/** Fill uncached chart references of a drawing tree in place. */
export function fillDrawings(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, lookup: ChartCellLookup, resolve?: ChartNameResolver, options: ChartFillOptions = {}): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && !content.chart.unsupported && (chartHasPending(content.chart) || (options.hidden && chartHasRefs(content.chart)))) content.chart = fillChartRanges(resolveChartNames(content.chart, ownerSheet, resolve), ownerSheet, lookup, options)
    if (node.children) fillDrawings(node.children, ownerSheet, lookup, resolve, options)
  }
}
/** The lines a drawing tree still needs (defined names resolved first), for paged readers. */
export function pendingLines(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, out: ChartLine[] = []): ChartLine[] {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    const chart = content?.kind === 'xlsxChart' ? content.chart : undefined
    if (chart) {
      const add = (range: ChartRange | undefined, text: boolean, limit?: number) => { if (range && !range.definedName) out.push(chartLine(range, ownerSheet, text, limit)) }
      add(chart.categoryPending, true)
      for (const group of chart.groups) for (const s of group.series) {
        const limit = seriesLimit(s)
        add(s.pending?.values, false, limit); add(s.pending?.xValues, false, limit); add(s.pending?.categories, true, categoryLimit(s))
        if (s.pending?.name) add({ ...s.pending.name, r1: s.pending.name.r0, c1: s.pending.name.c0 }, true)
      }
    }
    if (node.children) pendingLines(node.children, ownerSheet, out)
  }
  return out
}
/** Every reference a drawing tree still needs, for paged readers. */
export function pendingRanges(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, out: Array<ChartRange & { sheet: string }> = [], withNames = false): Array<ChartRange & { sheet: string }> {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    const chart = content?.kind === 'xlsxChart' ? content.chart : undefined
    if (chart) {
      const add = (range: ChartRange | undefined) => { if (range && (withNames || !range.definedName)) out.push({ ...range, sheet: range.sheet ?? ownerSheet }) }
      add(chart.categoryPending)
      for (const group of chart.groups) for (const s of group.series) { add(s.pending?.values); add(s.pending?.xValues); add(s.pending?.name); add(s.pending?.categories) }
      // Names a cached role's provenance reference uses are resolved too (never read as cells).
      if (withNames) { add(chart.categoryRef); for (const group of chart.groups) for (const s of group.series) { add(s.refs?.values); add(s.refs?.xValues) } }
    }
    if (node.children) pendingRanges(node.children, ownerSheet, out, withNames)
  }
  return out
}

/** Charts that still need cells become placeholders with `reason` (a refused read). */
export function refuseCharts(drawings: XlsxDrawingLike[] | undefined, reason: string): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && chartHasPending(content.chart)) { const next = { ...content.chart, unsupported: reason }; succeedChartModel(content.chart, next); content.chart = next }
    if (node.children) refuseCharts(node.children, reason)
  }
}
/** Accounted bytes per chart cell read (map entry, key and value). */
export const CHART_CELL_BYTES = 160

/** Charts with an uncached reference to a sheet `reasonFor` names become placeholders with that reason. */
export function markUnreadableCharts(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, reasonFor: (sheet: string) => string | undefined): boolean {
  let marked = false
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && !content.chart.unsupported) {
      for (const line of pendingLines([node], ownerSheet)) {
        const reason = reasonFor(line.sheet)
        if (reason) { const next = { ...content.chart, unsupported: reason }; succeedChartModel(content.chart, next); content.chart = next; marked = true; break }
      }
    }
    if (node.children && markUnreadableCharts(node.children, ownerSheet, reasonFor)) marked = true
  }
  return marked
}

/** Significant digits kept by the decimal scanner (a sticky digit stands for the rest: exact rounding). */
const SCAN_DIGITS = 800
/**
 * The chart's numeric reading of a text cell, following Excel (not JS
 * Number()): optional surrounding ASCII spaces, an optional sign, digits with
 * an optional decimal point, and an optional exponent. Radix prefixes (0x,
 * 0o, 0b), Infinity, other whitespace or control characters and anything
 * else are not numbers. Streamed from UTF-16 code units without building the
 * text; any-length decimals round exactly (the first SCAN_DIGITS significant
 * digits plus a sticky digit). Resident and paged charts both use it.
 */
export function scanChartNumber(length: number, at: (i: number) => number): number | undefined {
  let start = 0, end = length
  while (start < end && at(start) === 0x20) start++
  while (end > start && at(end - 1) === 0x20) end--
  if (start === end) return undefined
  let i = start, negative = false
  if (at(i) === 0x2b || at(i) === 0x2d) { negative = at(i) === 0x2d; i++ }
  let digits = '', sticky = false, seen = false, exponent = 0, leading = true
  for (; i < end; i++) {
    const c = at(i)
    if (c < 0x30 || c > 0x39) break
    seen = true
    if (leading && c === 0x30) continue
    leading = false
    if (digits.length < SCAN_DIGITS) digits += String.fromCharCode(c); else { exponent++; if (c !== 0x30) sticky = true }
  }
  if (i < end && at(i) === 0x2e) {
    for (i++; i < end; i++) {
      const c = at(i)
      if (c < 0x30 || c > 0x39) break
      seen = true
      if (leading && c === 0x30) { exponent--; continue }
      leading = false
      if (digits.length < SCAN_DIGITS) { digits += String.fromCharCode(c); exponent-- } else if (c !== 0x30) sticky = true
    }
  }
  if (!seen) return undefined
  if (i < end && (at(i) | 0x20) === 0x65) {
    i++
    let expNegative = false, expSeen = false, exp = 0
    if (i < end && (at(i) === 0x2b || at(i) === 0x2d)) { expNegative = at(i) === 0x2d; i++ }
    for (; i < end; i++) {
      const c = at(i)
      if (c < 0x30 || c > 0x39) break
      expSeen = true
      if (exp < 1e7) exp = exp * 10 + (c - 0x30)
    }
    if (!expSeen) return undefined
    exponent += expNegative ? -exp : exp
  }
  if (i !== end) return undefined
  if (!digits) return negative ? -0 : 0
  const value = Number(`${negative ? '-' : ''}${digits}${sticky ? '1' : ''}e${exponent - (sticky ? 1 : 0)}`)
  return Number.isFinite(value) ? value : undefined
}
