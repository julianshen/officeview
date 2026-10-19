/**
 * Chart data fill (lazy chunk): uncached references resolved from worksheet
 * cells, defined names, paged-reader lines and the decimal scanner. Moved
 * from `src/xlsx/chart.ts` (A0a) behaviour-identically.
 */
import { CHART_LABEL_LIMIT, CHART_POINT_LIMIT, parseChartRange } from './parse'
import type { XlsxChartModel, XlsxChartRange, XlsxChartSeries } from '../xlsx/chartModel'

export type ChartCellValue = string | number | boolean | null | undefined
/** Cell lookup for filling uncached references: `sheet` is the resolved sheet name. */
export interface ChartCell { value: ChartCellValue; /** The cell's formatted text (number/date format), as Excel shows it. */ text?: string; /** The value could not be read in full (a numeric form past the scanner's exact range). */ truncated?: boolean; /** The value's numeric reading when the reader computed it without building the text (null: not a number). */ number?: number | null }
/** `text`: the caller needs the formatted label (categories, names); values never format. */
export type ChartCellLookup = (sheet: string, row: number, col: number, text: boolean) => ChartCell | undefined

/** The one line of cells a chart consumes from a range (a 2-D range keeps its
 * last row/column, the level nearest the data), bounded by CHART_POINT_LIMIT.
 * Readers and the fill use this same sequence. */
export interface ChartLine { sheet: string; vertical: boolean; fixed: number; start: number; count: number; truncated: boolean; text: boolean }
export function chartLine(range: XlsxChartRange, ownerSheet: string, text: boolean): ChartLine {
  const rows = range.r1 - range.r0 + 1, cols = range.c1 - range.c0 + 1
  const vertical = cols === 1 || rows >= cols
  const length = vertical ? rows : cols
  return { sheet: range.sheet ?? ownerSheet, vertical, fixed: vertical ? range.c1 : range.r1, start: vertical ? range.r0 : range.c0, count: Math.min(length, CHART_POINT_LIMIT), truncated: length > CHART_POINT_LIMIT, text }
}
/** The cell coordinates of point `i` on a line. */
export const linePoint = (line: ChartLine, i: number): { row: number; col: number } => line.vertical ? { row: line.start + i, col: line.fixed } : { row: line.fixed, col: line.start + i }

export function chartHasPending(model: XlsxChartModel): boolean {
  return Boolean(model.categoryPending) || model.groups.some(g => g.series.some(s => s.pending))
}
/** The cells a range resolves to, row-major, bounded by CHART_POINT_LIMIT. */
function rangeCells(range: XlsxChartRange, sheet: string, lookup: ChartCellLookup, text: boolean): { values: Array<ChartCell | undefined>; truncated: boolean } {
  const line = chartLine(range, sheet, text)
  const values: Array<ChartCell | undefined> = []
  for (let i = 0; i < line.count; i++) { const at = linePoint(line, i); values.push(lookup(line.sheet, at.row, at.col, text)) }
  return { values, truncated: line.truncated }
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
export function resolveChartNames(model: XlsxChartModel, ownerSheet: string, resolve: ChartNameResolver | undefined): XlsxChartModel {
  let unsupported = model.unsupported
  const fix = (range: XlsxChartRange | undefined): XlsxChartRange | undefined => {
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
    const pending = Object.fromEntries(Object.entries(series.pending).map(([key, range]) => [key, fix(range)]).filter(([, range]) => range)) as NonNullable<XlsxChartSeries['pending']>
    const { pending: _old, ...rest } = series
    void _old
    return Object.keys(pending).length ? { ...rest, pending } : rest
  }) }))
  const { categoryPending: _cat, ...rest } = model
  void _cat
  return { ...rest, groups, ...(categoryPending ? { categoryPending } : {}), ...(unsupported ? { unsupported } : {}) }
}

/** Fill uncached references from worksheet cells (returns a new model). */
export function fillChartRanges(model: XlsxChartModel, ownerSheet: string, lookup: ChartCellLookup): XlsxChartModel {
  let truncated = model.truncated ?? false
  const sheetOf = (range: XlsxChartRange) => range.sheet ?? ownerSheet
  let categories = model.categories
  if (model.categoryPending) {
    const cells = rangeCells(model.categoryPending, sheetOf(model.categoryPending), lookup, true)
    truncated ||= cells.truncated
    categories = cells.values.map(cell => { const text = asText(cell); if (text.length > CHART_LABEL_LIMIT) { truncated = true; return text.slice(0, CHART_LABEL_LIMIT) } return text })
  }
  const groups = model.groups.map(group => ({
    ...group,
    series: group.series.map(series => {
      if (!series.pending) return series
      const { pending, ...rest } = series
      const next: XlsxChartSeries = { ...rest }
      // A numeric form the scanner cannot read exactly is an unavailable point, and the chart says it is truncated.
      const asNumber = (cell: ChartCell | undefined): number | undefined => { if (cell?.truncated) { truncated = true; return undefined } if (cell && cell.number !== undefined) return cell.number ?? undefined; return numberOf(cell?.value) }
      if (pending.values) { const cells = rangeCells(pending.values, sheetOf(pending.values), lookup, false); truncated ||= cells.truncated; next.values = cells.values.map(asNumber) }
      if (pending.xValues) { const cells = rangeCells(pending.xValues, sheetOf(pending.xValues), lookup, false); truncated ||= cells.truncated; next.xValues = cells.values.map(asNumber) }
      if (pending.name) { const full = asText(lookup(sheetOf(pending.name), pending.name.r0, pending.name.c0, true)); const text = full.slice(0, CHART_LABEL_LIMIT); if (full.length > CHART_LABEL_LIMIT) truncated = true; if (text) next.name = text }
      if (next.pointColors && next.values.length > next.pointColors.length) {
        const first = next.pointColors
        next.pointColors = Array.from({ length: Math.min(256, next.values.length) }, (_, i) => first[i] ?? first[i % first.length])
      }
      return next
    }),
  }))
  const { categoryPending, ...rest } = model
  void categoryPending
  return { ...rest, categories, groups, ...(truncated ? { truncated: true } : {}) }
}

/** Resident workbooks: fill every uncached chart reference from parsed cells,
 * after evaluation and semantics. Only the cells on the consumed lines are
 * visited (rows by binary search, cells by binary search), and labels are
 * formatted on demand: the workspace is bounded by the chart point limits. */
export function fillWorkbookCharts<Cell extends { row: number; col: number; value: ChartCellValue }, Sheet extends { name: string; workbookIndex?: number; rows: ReadonlyArray<{ index: number; cells: ReadonlyArray<Cell> }>; drawings?: XlsxDrawingLike[] }>(
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
  const lookup: ChartCellLookup = (name, rowIndex, col, wantText) => {
    const sheet = sheetByName(name)
    if (!sheet) return undefined
    const row = search(sheet.rows, row => row.index, rowIndex, isSorted(sheet.rows, sheet.rows, row => row.index))
    const cell = row && search(row.cells, cell => cell.col, col, isSorted(row.cells, row.cells, cell => cell.col))
    if (!cell) return undefined
    if (options.available && !options.available(cell)) return { value: undefined }
    let text: string | undefined
    if (wantText && options.text) { try { text = options.text(sheet, cell) } catch { text = undefined } }
    return { value: cell.value, ...(text !== undefined ? { text } : {}) }
  }
  const resolve = chartNameResolver(options.names ?? [], sheets)
  for (const sheet of sheets) {
    try { fillDrawings(sheet.drawings, sheet.name, lookup, resolve) }
    catch { /* a chart that cannot be filled keeps its placeholder state; the workbook stays usable */ }
  }
}
/** Resolve defined names of a drawing tree in place (paged readers resolve before reading cells). */
export function resolveDrawingNames(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, resolve: ChartNameResolver | undefined): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
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
export function fillDrawings(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, lookup: ChartCellLookup, resolve?: ChartNameResolver): void {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    if (content?.kind === 'xlsxChart' && content.chart && chartHasPending(content.chart)) content.chart = fillChartRanges(resolveChartNames(content.chart, ownerSheet, resolve), ownerSheet, lookup)
    if (node.children) fillDrawings(node.children, ownerSheet, lookup, resolve)
  }
}
/** The lines a drawing tree still needs (defined names resolved first), for paged readers. */
export function pendingLines(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, out: ChartLine[] = []): ChartLine[] {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    const chart = content?.kind === 'xlsxChart' ? content.chart : undefined
    if (chart) {
      const add = (range: XlsxChartRange | undefined, text: boolean) => { if (range && !range.definedName) out.push(chartLine(range, ownerSheet, text)) }
      add(chart.categoryPending, true)
      for (const group of chart.groups) for (const s of group.series) {
        add(s.pending?.values, false); add(s.pending?.xValues, false); add(s.pending?.categories, true)
        if (s.pending?.name) add({ ...s.pending.name, r1: s.pending.name.r0, c1: s.pending.name.c0 }, true)
      }
    }
    if (node.children) pendingLines(node.children, ownerSheet, out)
  }
  return out
}
/** Every reference a drawing tree still needs, for paged readers. */
export function pendingRanges(drawings: XlsxDrawingLike[] | undefined, ownerSheet: string, out: Array<XlsxChartRange & { sheet: string }> = [], withNames = false): Array<XlsxChartRange & { sheet: string }> {
  for (const node of drawings ?? []) {
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
    const chart = content?.kind === 'xlsxChart' ? content.chart : undefined
    if (chart) {
      const add = (range: XlsxChartRange | undefined) => { if (range && (withNames || !range.definedName)) out.push({ ...range, sheet: range.sheet ?? ownerSheet }) }
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
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
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
    const content = node.content as { kind: string; chart?: XlsxChartModel } | undefined
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
