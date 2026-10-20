/**
 * Chart data fill (lazy chunk): uncached references resolved from worksheet
 * cells, defined names, paged-reader lines and the decimal scanner. Moved
 * from `src/xlsx/chart.ts` (A0a) behaviour-identically.
 */
import { CHART_ANALYTIC_POINT_LIMIT, CHART_LABEL_LIMIT, CHART_POINT_LIMIT, parseChartRange } from './parse'
import type { ChartModel, ChartPoint, ChartRange, ChartSeries } from './model'
import { BLANK, UNAVAILABLE, errorPoint, valuePoint } from './points'

export type ChartCellValue = string | number | boolean | null | undefined
/**
 * Cell lookup for filling uncached references: `sheet` is the resolved sheet
 * name. The point state comes from the cell itself, the same way on every path
 * (spec 5.8): `error` keeps the error code, `unavailable` marks a formula with
 * no saved value, a missing cell is blank.
 */
export interface ChartCell {
  value: ChartCellValue
  /** The cell's formatted text (number/date format), as Excel shows it. */ text?: string
  /** The value could not be read in full (a numeric form past the scanner's exact range). */ truncated?: boolean
  /** The value's numeric reading when the reader computed it without building the text (null: not a number). */ number?: number | null
  /** An error cell's code (`#N/A`, `#DIV/0!`, ...). */ error?: string
  /** The cell has no value to show (a formula without a saved value under the saved-value policy). */ unavailable?: boolean
}
/** `text`: the caller needs the formatted label (categories, names); values never format. */
export type ChartCellLookup = (sheet: string, row: number, col: number, text: boolean) => ChartCell | undefined
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
  if (series.analytic !== true) return CHART_POINT_LIMIT
  const ranges = [series.pending?.values, series.pending?.xValues].filter((range): range is ChartRange => Boolean(range && !range.definedName))
  return ranges.some(range => lineLength(range) > CHART_ANALYTIC_POINT_LIMIT) ? CHART_POINT_LIMIT : CHART_ANALYTIC_POINT_LIMIT
}
/** The point state of one filled cell (spec 5.8); `onTruncated` hears a value the reader could not read in full. */
export function cellPoint(cell: ChartCell | undefined, onTruncated?: () => void): ChartPoint {
  if (!cell) return BLANK
  if (cell.unavailable) return UNAVAILABLE
  if (cell.error !== undefined) return errorPoint(cell.error)
  // A numeric form the scanner cannot read exactly has no value, and the chart says it is truncated.
  if (cell.truncated) { onTruncated?.(); return UNAVAILABLE }
  const number = cell.number !== undefined ? cell.number ?? undefined : numberOf(cell.value)
  return number === undefined ? BLANK : valuePoint(number)
}
/** The cell coordinates of point `i` on a line. */
export const linePoint = (line: ChartLine, i: number): { row: number; col: number } => line.vertical ? { row: line.start + i, col: line.fixed } : { row: line.fixed, col: line.start + i }

export function chartHasPending(model: ChartModel): boolean {
  return Boolean(model.categoryPending) || model.groups.some(g => g.series.some(s => s.pending))
}
/** The cells a range resolves to, in line order, bounded by `limit`, with the indices of hidden cells. */
function rangeCells(range: ChartRange, sheet: string, lookup: ChartCellLookup, text: boolean, hidden?: ChartHiddenLookup, limit = CHART_POINT_LIMIT): { values: Array<ChartCell | undefined>; truncated: boolean; hidden: number[] } {
  const line = chartLine(range, sheet, text, limit)
  const values: Array<ChartCell | undefined> = [], hiddenAt: number[] = []
  for (let i = 0; i < line.count; i++) {
    const at = linePoint(line, i)
    values.push(lookup(line.sheet, at.row, at.col, text))
    if (hidden?.(line.sheet, at.row, at.col)) hiddenAt.push(i)
  }
  return { values, truncated: line.truncated, hidden: hiddenAt }
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
  if (!chartHasPending(model)) return model
  const categoryPending = fix(model.categoryPending)
  const groups = model.groups.map(group => ({ ...group, series: group.series.map(series => {
    if (!series.pending) return series
    const pending = Object.fromEntries(Object.entries(series.pending).map(([key, range]) => [key, fix(range)]).filter(([, range]) => range)) as NonNullable<ChartSeries['pending']>
    const { pending: _old, ...rest } = series
    void _old
    return Object.keys(pending).length ? { ...rest, pending } : rest
  }) }))
  const { categoryPending: _cat, ...rest } = model
  void _cat
  return { ...rest, groups, ...(categoryPending ? { categoryPending } : {}), ...(unsupported ? { unsupported } : {}) }
}

/**
 * Fill uncached references from worksheet cells (returns a new model). Every
 * cell becomes a point state (`cellPoint`), the line order is kept (no
 * compaction), and a cell in a hidden row or column is recorded by index
 * (`hidden`/`categoryHidden`) for `plotVisOnly`, which applies at draw time.
 */
export function fillChartRanges(model: ChartModel, ownerSheet: string, lookup: ChartCellLookup, hidden?: ChartHiddenLookup): ChartModel {
  let truncated = model.truncated ?? false
  const sheetOf = (range: ChartRange) => range.sheet ?? ownerSheet
  let categories = model.categories, categoryHidden = model.categoryHidden
  if (model.categoryPending) {
    const cells = rangeCells(model.categoryPending, sheetOf(model.categoryPending), lookup, true, hidden)
    truncated ||= cells.truncated
    categories = cells.values.map(cell => { const text = asText(cell); if (text.length > CHART_LABEL_LIMIT) { truncated = true; return text.slice(0, CHART_LABEL_LIMIT) } return text })
    categoryHidden = cells.hidden.length ? cells.hidden : undefined
  }
  const groups = model.groups.map(group => ({
    ...group,
    series: group.series.map(series => {
      if (!series.pending) return series
      const { pending, ...rest } = series
      const next: ChartSeries = { ...rest }
      const limit = seriesLimit(series)
      if (series.analytic === true && limit === CHART_POINT_LIMIT) next.analytic = 'refused'
      const asPoint = (cell: ChartCell | undefined) => cellPoint(cell, () => { truncated = true })
      const hiddenAt = new Set(series.hidden)
      const role = (range: ChartRange) => {
        const cells = rangeCells(range, sheetOf(range), lookup, false, hidden, limit)
        // Display truncation (spec 7): an analytic series keeps every acquired point; the drawing is capped.
        truncated ||= cells.truncated || cells.values.length > CHART_POINT_LIMIT
        for (const i of cells.hidden) hiddenAt.add(i)
        return cells.values.map(asPoint)
      }
      if (pending.values) next.values = role(pending.values)
      if (pending.xValues) next.xValues = role(pending.xValues)
      if (hiddenAt.size) next.hidden = [...hiddenAt].sort((a, b) => a - b)
      if (pending.name) { const full = asText(lookup(sheetOf(pending.name), pending.name.r0, pending.name.c0, true)); const text = full.slice(0, CHART_LABEL_LIMIT); if (full.length > CHART_LABEL_LIMIT) truncated = true; if (text) next.name = text }
      if (next.pointColors && next.values.length > next.pointColors.length) {
        const first = next.pointColors
        next.pointColors = Array.from({ length: Math.min(256, next.values.length) }, (_, i) => first[i] ?? first[i % first.length])
      }
      return next
    }),
  }))
  const { categoryPending, categoryHidden: _hidden, ...rest } = model
  void categoryPending, void _hidden
  return { ...rest, categories, ...(categoryHidden ? { categoryHidden } : {}), groups, ...(truncated ? { truncated: true } : {}) }
}

/** Resident workbooks: fill every uncached chart reference from parsed cells,
 * after evaluation and semantics. Only the cells on the consumed lines are
 * visited (rows by binary search, cells by binary search), and labels are
 * formatted on demand: the workspace is bounded by the chart point limits. */
export function fillWorkbookCharts<Cell extends { row: number; col: number; value: ChartCellValue; valueIsError?: boolean }, Sheet extends { name: string; workbookIndex?: number; rows: ReadonlyArray<{ index: number; hidden?: boolean; cells: ReadonlyArray<Cell> }>; cols?: ReadonlyArray<{ min: number; max: number; hidden?: boolean }>; drawings?: XlsxDrawingLike[] }>(
  sheets: readonly Sheet[], options: { names?: readonly ChartDefinedName[]; text?: (sheet: Sheet, cell: Cell) => string; available?: (cell: Cell) => boolean } = {},
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
  const lookup: ChartCellLookup = (name, rowIndex, col, wantText) => {
    const sheet = sheetByName(name)
    if (!sheet) return undefined
    const row = rowOf(sheet, rowIndex)
    const cell = row && search(row.cells, cell => cell.col, col, isSorted(row.cells, row.cells, cell => cell.col))
    if (!cell) return undefined
    if (options.available && !options.available(cell)) return { value: undefined, unavailable: true }
    let text: string | undefined
    if (wantText && options.text) { try { text = options.text(sheet, cell) } catch { text = undefined } }
    return { value: cell.value, ...(cell.valueIsError === true ? { error: String(cell.value) } : {}), ...(text !== undefined ? { text } : {}) }
  }
  const hidden: ChartHiddenLookup = (name, rowIndex, col) => {
    const sheet = sheetByName(name)
    return Boolean(sheet && (rowOf(sheet, rowIndex)?.hidden || sheet.cols?.some(spec => spec.hidden && spec.min <= col && col <= spec.max)))
  }
  const resolve = chartNameResolver(options.names ?? [], sheets)
  for (const sheet of sheets) {
    try { fillDrawings(sheet.drawings, sheet.name, lookup, resolve, hidden) }
    catch { /* a chart that cannot be filled keeps its placeholder state; the workbook stays usable */ }
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
export function fillDrawings(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, lookup: ChartCellLookup, resolve?: ChartNameResolver, hidden?: ChartHiddenLookup): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && chartHasPending(content.chart)) content.chart = fillChartRanges(resolveChartNames(content.chart, ownerSheet, resolve), ownerSheet, lookup, hidden)
    if (node.children) fillDrawings(node.children, ownerSheet, lookup, resolve, hidden)
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
        add(s.pending?.values, false, limit); add(s.pending?.xValues, false, limit); add(s.pending?.categories, true)
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
    }
    if (node.children) pendingRanges(node.children, ownerSheet, out, withNames)
  }
  return out
}

/** Charts that still need cells become placeholders with `reason` (a refused read). */
export function refuseCharts(drawings: XlsxDrawingLike[] | undefined, reason: string): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: ChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && chartHasPending(content.chart)) content.chart = { ...content.chart, unsupported: reason }
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
        if (reason) { content.chart = { ...content.chart, unsupported: reason }; marked = true; break }
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
