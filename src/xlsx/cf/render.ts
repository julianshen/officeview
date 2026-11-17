/** Async CF preparation; painters only consume admitted source-space snapshots. */
import { cfSheetNonUniformFontRows, cfRowUniformOfLayout } from '../rowFont'
import { createCfStatistics } from './statistics'
import { sameCfStamp } from './cache'
import { cfRuntimeEvidence, type CfEvidenceGate } from './predicates'
import { createCfService, type CfPaintStatistics } from './service'
import { cfAlignClock, cfEvaluationContextFor, cfRoutePolicy, type CfDocumentEvaluation } from './formula'
export { cfDocumentEvaluationOf, cfDocumentEvaluationOfPaged } from './formula'
import { CF_LIMITS, createCfMemory, createCfWorkBudget, measureCfPayload, type CfMemory, type CfWorkBudget } from './limits'
import { resolveCfCellStyle } from './dxf'
import { barPaintable, barWithheld, CF_BAR_RENDER_MODEL, paintDataBars, type CfBarHost, type CfBarScale } from './dataBarPaint'
import { admitBatchGlyphs, cfHidesText, iconPaintsOn, iconTextInset, paintIconCells, type CfAdmittedGlyphs, type CfIconGate } from './iconPaint'
import type { CfBarDescriptor, CfCellState, CfIconDescriptor, CfLiteral, CfPrepareRequest, CfStamp, CfValueSource, SourceRect } from './model'
import { cfAdmittedSheetNames, cfDependencyGenOf, cfReferencedSheetNames, withCfDependencyGen, type CfWorkbookSheet, type CfWorkbookSource } from './dependencies'
import type { XlsxSheet } from '../types'
import type { RegionStyleLayer } from '../tableStyles/regionLayer'
import { regionStylesFor } from '../tableStyleLoader'
import { cellSourceRow } from '../view/sheetCells'
import { isValueUnavailable, type ValuePolicy } from '../view/values'
import { cfNeedsViewport, cfExplicitBounds, cfRulesUseClock } from '../cfLoader'
import { STATIC_SHEET_BUDGET, computeMetrics, layoutCellBody, spreadsheetColor, textMeasurer, type MeasureCached } from '../render'
import { cellTextWidth, withTextInset, type CellBodyLayout } from '../view/textLayout'
import type { XlsxCell, XlsxCellStyle } from '../types'
import { CfPageFault, cfRulesVersion, pagedStoreCfPages } from '../view/pagedSource'
import { isDateNumberStyle } from '../dateNumberFormat'
import type { SourceStore } from '../view/contracts'
import { hostTask } from '../view/scheduler'

export interface CfPaintBatch {
  readonly stamp: CfStamp
  readonly key: string
  readonly layer: RegionStyleLayer
  /** The bar this batch paints at a source cell (current generation, every paint gate held), else undefined. */
  barAt(row: number, col: number): CfBarDescriptor | undefined
  /** Paints the batch's bars of `visible` in the fill phase, each clipped to its `cellRect` host (Task 22). */
  paintBars(ctx: CanvasRenderingContext2D, cellRect: (row: number, col: number) => CfBarHost | undefined,
    visible: { r0: number; c0: number; r1: number; c1: number }, scale: CfBarScale): boolean
  /** Ruling S3-O: the canvas draws no value text in this host at this scale (a `showValue="0"` bar that paints there,
   * under evidence). */
  hidesTextAt(row: number, col: number, host: CfBarHost, scale: CfBarScale): boolean
  /** The icon this batch paints at a source cell (current generation, every paint gate held, on the resident host),
   * else undefined (Task 29). */
  iconAt(row: number, col: number): CfIconDescriptor | undefined
  /** Paints the batch's icons of `visible` after its bars, each on its `cellRect` host (Task 29). */
  paintIcons(ctx: CanvasRenderingContext2D, cellRect: (row: number, col: number) => CfBarHost | undefined,
    visible: { r0: number; c0: number; r1: number; c1: number }, scale: CfBarScale): boolean
  /** Ruling S4-L: the text inset of this host when its icon paints there, else undefined (no icon, no inset). */
  textInsetAt(row: number, col: number, host: CfBarHost, scale: CfBarScale): { left: number; right: number } | undefined
  /** The static renderer's one call per CF cell (Ruling S4-L, fix round 1): lays the cell out un-inset, builds its one
   * host (box, fill, un-inset `textWidth`, the cell as laid out), and lays it out again inset and/or with its text
   * hidden where this batch's icon or bar says so on that host. Keeps all icon arithmetic out of the eager renderer. */
  layoutCell(ctx: CanvasRenderingContext2D, sheet: XlsxSheet, cell: XlsxCell, x: number, y: number, w: number, h: number, measured: MeasureCached,
    overflow: { left: number; right: number } | undefined, scale: CfBarScale): { layout: CellBodyLayout; host: CfBarHost }
  dispose(): void
}

/** The default measuring context for resident icon hosts (Ruling S4-L): an offscreen or detached canvas, or none (an
 * unknown text width withholds the icon). */
export function defaultLayoutContext(): CanvasRenderingContext2D | undefined {
  try {
    if (typeof OffscreenCanvas !== 'undefined') return (new OffscreenCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D | null) ?? undefined
    if (typeof document !== 'undefined') return document.createElement('canvas').getContext('2d') ?? undefined
  } catch { /* no context: withheld */ }
  return undefined
}
/** Layout work charged per resident icon host whose text is measured: the transient layout and its cached width. */
const ICON_TEXT_MEASURE_BYTES = 512
/** No inferred values: cached formula errors stay errors; missing formula
 * values refuse rather than masquerading as blanks. */
/** Task 9 reuses the borrowed-value resident source for per-batch text scopes. */
export function residentSource(sheet: XlsxSheet, stamp: CfStamp, memory: CfMemory, policy: ValuePolicy): CfValueSource {
  let bounds: SourceRect | undefined, ordered = true
  return {
    // Resident cells carry their resolved style, so every number is classified (data-bar S3-G).
    numberFormats: 'resolved',
    stamp: () => stamp,
    async bounds(_id, budget, signal) {
      if (bounds) return bounds
      let previousRow = -1
      let firstRow = 1048576, lastRow = -1, firstCol = 16384, lastCol = -1
      for (const row of sheet.rows) {
        if (!Number.isSafeInteger(row.index) || row.index < previousRow) ordered = false
        previousRow = row.index
        const rowWork = budget?.spend(1)
        if (rowWork === 'refused') throw new Error('CF work limit')
        if (rowWork === 'yield') { await hostTask(); budget?.resume() }
        for (const cell of row.cells) {
        if (signal?.aborted) throw new Error('CF cancelled')
        const work = budget?.spend(1)
        if (work === 'refused') throw new Error('CF work limit')
        if (work === 'yield') { await hostTask(); budget?.resume() }
        const r = cellSourceRow(row, cell)
        if (r !== row.index) ordered = false
        firstRow = Math.min(firstRow, r); lastRow = Math.max(lastRow, r); firstCol = Math.min(firstCol, cell.col); lastCol = Math.max(lastCol, cell.col)
      } }
      return bounds = lastRow < 0 ? { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 } : { firstRow, lastRow, firstCol, lastCol }
    },
    async *pages(_id, rects, budget, signal) {
      // One-cell borrowed-value pages avoid copying strings or indexing the
      // entire resident sheet. The service owns its admitted result map.
      const hold = memory.reserve(256)
      if (hold.kind !== 'reserved') throw new Error('CF page admission')
      try {
        const rows = new Uint32Array(1), cols = new Uint16Array(1)
        let begin = 0, end = sheet.rows.length
        if (ordered && bounds && rects.length) {
          const first = Math.min(...rects.map(r => r.firstRow)), last = Math.max(...rects.map(r => r.lastRow))
          const lower = (value: number) => { let a = 0, b = sheet.rows.length; while (a < b) { const m = (a + b) >>> 1; if (sheet.rows[m]!.index < value) a = m + 1; else b = m }; return a }
          begin = lower(first); end = lower(last + 1)
        }
        for (let ri = begin; ri < end; ri++) { const row = sheet.rows[ri]!; for (const cell of row.cells) {
          if (signal.aborted) throw new Error('CF cancelled')
          const work = budget.spend(1)
          if (work === 'refused') throw new Error('CF work limit')
          if (work === 'yield') { await hostTask(); budget.resume() }
          const r = cellSourceRow(row, cell)
          if (!rects.some(b => r >= b.firstRow && r <= b.lastRow && cell.col >= b.firstCol && cell.col <= b.lastCol)) continue
          if (isValueUnavailable(cell, policy)) throw new Error('CF unavailable formula value')
          const v = cell.value
          const isErr = cell.valueIsError
          const value: CfLiteral = isErr ? { kind: 'error', value: String(v) } : v === null ? { kind: 'blank' }
            : typeof v === 'number' ? (isDateNumberStyle(cell.style) ? { kind: 'number', value: v, dateSerial: true } : { kind: 'number', value: v }) : typeof v === 'boolean' ? { kind: 'boolean', value: v } : { kind: 'string', value: v }
          rows[0] = r; cols[0] = cell.col
          yield { value: { rows, cols, values: [value] }, bytes: 256, release() {}, transfer() { return { kind: 'unresolved', reason: 'borrowed-page' } } }
        } }
      } finally { hold.release() }
    },
  }
}
/** Read the committed store, including pivot overlays. A borrowed snapshot stops
 * painting synchronously when its source commits, before any async replacement. */
export function storeCfSource(store: SourceStore, stamp: CfStamp, memory: CfMemory): CfValueSource {
  return {
    stamp: () => ({ ...stamp, owner: store.stamp.docId, workbookValueGen: store.stamp.valueGen, structureGen: store.structureGen ?? 0 }),
    bounds: async () => { if (store.completeness !== 'complete') return undefined; const used = store.extent.used; return used
      ? { firstRow: used.r0, lastRow: used.r1, firstCol: used.c0, lastCol: used.c1 }
      : { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 } },
    pages: pagedStoreCfPages(store, memory),
  }
}
/**
 * Task 36 (Ruling S5-K): the workbook seam over canonical, overlay-aware stores (the resident `source.store`, the worker's
 * `PagedXlsxSheetAdapter`s). Each referenced sheet's values (`storeCfSource`) and its generation (`store.stamp.valueGen`)
 * come from the same store object, resolved lazily per name; a store whose import has not completed is pending, one that
 * ended incomplete has failed.
 */
export function cfWorkbookOfStores(sheets: readonly Readonly<{ sheetId: string; name: string }>[], storeOf: (sheetId: string) => SourceStore | undefined, memory: CfMemory): CfWorkbookSource {
  const byName = new Map(sheets.map(s => [s.name.toLowerCase(), s] as const))
  const made = new Map<string, CfWorkbookSheet>()
  return {
    sheetIdOf: name => byName.get(name.toLowerCase())?.sheetId,
    sheet(sheetId) {
      const known = made.get(sheetId)
      if (known) return known
      const entry = sheets.find(s => s.sheetId === sheetId)
      const store = entry ? storeOf(sheetId) : undefined
      if (!entry || !store) return undefined
      const stub: CfStamp = { owner: store.stamp.docId, sheetId, workbookValueGen: 0, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }
      const sheet: CfWorkbookSheet = {
        sheetId, name: entry.name, values: storeCfSource(store, stub, memory),
        state: () => (store.completeness === 'complete' ? 'ready' : store.completeness === 'incomplete' ? 'failed' : 'pending'),
        valueGen: () => store.stamp.valueGen,
      }
      made.set(sheetId, sheet)
      return sheet
    },
  }
}
/**
 * Task 36 (Ruling S5-K, fix rounds 2-3): the resident live key's dependency part, over the preparation's own admitted set
 * (`cfAdmittedSheetNames`: the first 16 distinct referenced names, host excluded). Each admitted sheet's current value
 * generation and completeness join the key; an unknown name reads `-1`. One store read per admitted sheet. Lazy: the
 * interactive view calls it on every render once the CF runtime is loaded, and withholds resident CF until then.
 */
export function cfResidentDependencyKey(rules: import('./model').CfSheetRules | undefined, hostName: string, sheets: readonly Readonly<{ sheetId: string; name: string }>[],
  storeOf: (sheetId: string) => Readonly<{ stamp: Readonly<{ valueGen: number }>; completeness?: string }>): string {
  const parts: Array<string | number> = []
  for (const name of cfAdmittedSheetNames(rules, hostName)) {
    const s = sheets.find(x => x.name.toLowerCase() === name.toLowerCase()), t = s && storeOf(s.sheetId)
    parts.push(t ? `${t.stamp.valueGen}${t.completeness}` : -1)
  }
  return parts.join()
}
/** The canonical stores of a workbook (the resident `WorkbookSource` shape). */
export type CfWorkbookStores = Readonly<{ sheets: readonly Readonly<{ sheetId: string; name: string }>[]; store(sheetId: string): SourceStore }>
/**
 * One preparation's workbook seam (Rulings S5-K, S5-P), or undefined when the host's rules name no other sheet (no
 * cross-sheet work at all). A host store reads its referenced sheets from the same workbook's stores; a host without a
 * store reads the document's sheets; a host store offered only document sheets gets none (cross-sheet rules withhold).
 */
export function cfWorkbookFor(rules: import('./model').CfSheetRules, hostName: string, hostStore: SourceStore | undefined, stores: CfWorkbookStores | undefined,
  documentSheets: readonly XlsxSheet[] | undefined, memory: CfMemory, policy: ValuePolicy): CfWorkbookSource | undefined {
  if (cfReferencedSheetNames(rules, hostName).length === 0) return undefined
  if (stores) return cfWorkbookOfStores(stores.sheets, id => (stores.sheets.some(s => s.sheetId === id) ? stores.store(id) : undefined), memory)
  if (documentSheets && !hostStore) return cfWorkbookOfDocument(documentSheets, memory, policy)
  return undefined
}
/** Static, print and export legs (Ruling S5-K): the document's own sheets, which carry no overlay and never commit, read
 * through the same resident reader the host uses (one value policy). */
export function cfWorkbookOfDocument(sheets: readonly XlsxSheet[], memory: CfMemory, policy: ValuePolicy): CfWorkbookSource {
  const idOf = (sheet: XlsxSheet, index: number): string => sheet.sheetId ?? String(index + 1)
  const made = new Map<string, CfWorkbookSheet>()
  return {
    sheetIdOf: name => { const i = sheets.findIndex(s => s.name.toLowerCase() === name.toLowerCase()); return i < 0 ? undefined : idOf(sheets[i]!, i) },
    sheet(sheetId) {
      const known = made.get(sheetId)
      if (known) return known
      const i = sheets.findIndex((s, j) => idOf(s, j) === sheetId)
      if (i < 0) return undefined
      const stub: CfStamp = { owner: 'static', sheetId, workbookValueGen: 0, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }
      const sheet: CfWorkbookSheet = { sheetId, name: sheets[i]!.name, values: residentSource(sheets[i]!, stub, memory, policy), state: () => 'ready', valueGen: () => 0 }
      made.set(sheetId, sheet)
      return sheet
    },
  }
}
/** One diagnostic surface for initial and late refusals, with the sheet cap. */
function unavailable(sheet: XlsxSheet): undefined {
  const ds = sheet.diagnostics ??= []
  if (ds.length < CF_LIMITS.diagnostics && !ds.some(d => d.feature === 'cf-paint-unavailable')) ds.push({ kind: 'unsupported', feature: 'cf-paint-unavailable', message: 'Conditional formatting could not be prepared within the admitted batch; base formatting is shown.' })
  return undefined
}
function recordDiagnostic(sheet: XlsxSheet, feature: string, message: string): void {
  const ds = sheet.diagnostics ??= []
  if (ds.length < CF_LIMITS.diagnostics && !ds.some(d => d.feature === feature)) {
    ds.push({ kind: 'unsupported', feature, message })
  }
}
// Retained statistics sessions across preparations / viewport scrolls per (sheet, generation) (Ruling I5)
const sheetStatsSessions = new WeakMap<object, { session: import('./statistics').CfStatisticsSession; key: string }>()
const sheetStatsRefs = new WeakMap<object, number>()

export function retainSheetCfStatistics(sheet: object): void {
  sheetStatsRefs.set(sheet, (sheetStatsRefs.get(sheet) ?? 0) + 1)
}

export function disposeSheetCfStatistics(sheet: object): void {
  const refs = (sheetStatsRefs.get(sheet) ?? 1) - 1
  if (refs <= 0) {
    sheetStatsRefs.delete(sheet)
    const retained = sheetStatsSessions.get(sheet)
    if (retained) {
      retained.session.dispose()
      sheetStatsSessions.delete(sheet)
    }
  } else {
    sheetStatsRefs.set(sheet, refs)
  }
}

/**
 * Fix round 2: each resident bar cell's host as the static renderer lays it out (computeMetrics geometry, the fill it
 * paints under the bar: own style through `styleAt`, resolved by `spreadsheetColor`), so `barAt` and `describe` share
 * the painter's own withholding predicate (`barWithheld`). The geometry and the cell index are built once, on first
 * use, and only after `admit` accepts their bytes; a refusal, or a cell outside the static grid, has no host, which
 * withholds (fail closed).
 */
/**
 * Ruling S4-ROW-EXACT (Task 28 round 4): each row's authored height in pt (its `ht`) from the sheet, for the icon gate.
 * A row without `ht` takes the sheet's `defaultRowHeightPt` (Ruling S4-DEFAULT-ROW, round 5); without one it is
 * unknown: undefined, and its icons withhold.
 */
export function cfRowPtOfSheet(sheet: Pick<XlsxSheet, 'rows' | 'defaultRowHeightPt'>): (row: number) => number | undefined {
  let table: Map<number, number> | undefined
  return row => {
    if (!table) {
      table = new Map()
      for (const r of sheet.rows) if (r.heightPt !== undefined && Number.isFinite(r.heightPt) && r.heightPt >= 0) table.set(r.index, r.heightPt)
    }
    // Ruling S4-DEFAULT-ROW: a row without `ht` takes the sheet's default (`sheetFormatPr defaultRowHeight`).
    return table.get(row) ?? sheet.defaultRowHeightPt
  }
}

/** Ruling S4-ROW-UNIFORM: each row's font uniformity from the resident sheet (`rowFont.ts`); rows the sheet does not
 * store hold no cell, so they are uniform (their icons have no cell to sit on). */
export function cfRowUniformOfSheet(sheet: Pick<XlsxSheet, 'rows' | 'rowFontsBlocked' | 'conditionalFormatting' | 'tables' | 'tablesUnknown'>): (row: number) => boolean | undefined {
  let of: ((row: number) => boolean | undefined) | undefined
  // S4-ROW-EFFECTIVE: CF font rules and tables too. A resident sheet without tables has `tables` absent (none); an
  // unknown inventory (`tablesUnknown`) withholds every row, as the paged path does.
  return row => (of ??= cfRowUniformOfLayout(cfSheetNonUniformFontRows({ ...sheet, tables: sheet.tablesUnknown ? undefined : sheet.tables ?? [] })))(row)
}

export function residentBarHosts(sheet: XlsxSheet, styleAt: (row: number, col: number, own: XlsxCellStyle | undefined) => XlsxCellStyle | undefined,
  admit: (bytes: number) => boolean,
  /** Icon hosts (Ruling S4-L, Task 29): also carry the cell as laid out and the un-inset `textWidth` the static renderer's
   * own measurer (`textMeasurer`, `layoutCellBody`) gives it, measured once per cell while the host is built and charged
   * as layout work through `admit`. No context, or a refusal, leaves the width unknown (withheld). */
  text?: { layoutContext?: () => CanvasRenderingContext2D | null | undefined },
): (row: number, col: number) => CfBarHost | undefined {
  let built: { colX: Float64Array; rowY: Float64Array; cells: Map<number, XlsxCell> } | null | undefined
  let measure: { ctx: CanvasRenderingContext2D; measured: MeasureCached } | null | undefined
  const widths = new Map<number, number | undefined>()
  const textWidthOf = (key: number, cell: XlsxCell, boxWidth: number, boxHeight: number): number | undefined => {
    if (widths.has(key)) return widths.get(key)
    if (measure === undefined) {
      const ctx = text?.layoutContext ? text.layoutContext() ?? undefined : defaultLayoutContext()
      measure = ctx ? { ctx, measured: textMeasurer(ctx) } : null
    }
    if (!measure || !admit(ICON_TEXT_MEASURE_BYTES)) return undefined
    let width: number | undefined
    try { width = cellTextWidth(layoutCellBody(measure.ctx, sheet, cell, 0, 0, boxWidth, boxHeight, measure.measured)) } catch { width = undefined }
    widths.set(key, width)
    return width
  }
  return (row, col) => {
    if (built === undefined) {
      let cellCount = 0
      for (const r of sheet.rows) cellCount += r.cells.length
      const m = computeMetrics(sheet)
      if (!admit((m.colWidthsPx.length + m.rowHeightsPx.length + 2) * 8 + cellCount * 64)) built = null
      else {
        const prefix = (sizes: readonly number[]) => { const out = new Float64Array(sizes.length + 1); for (let i = 0; i < sizes.length; i++) out[i + 1] = out[i]! + sizes[i]!; return out }
        const cells = new Map<number, XlsxCell>()
        for (const r of sheet.rows) for (const cell of r.cells) cells.set(cellSourceRow(r, cell) * 16384 + cell.col, cell)
        built = { colX: prefix(m.colWidthsPx), rowY: prefix(m.rowHeightsPx), cells }
      }
    }
    if (!built || row + 1 >= built.rowY.length || col + 1 >= built.colX.length) return undefined
    const own = built.cells.get(row * 16384 + col)
    const style = styleAt(row, col, own?.style), fill = style?.fillColor
    const box = { x: built.colX[col]!, y: built.rowY[row]!, width: built.colX[col + 1]! - built.colX[col]!, height: built.rowY[row + 1]! - built.rowY[row]! }
    const host = { ...box, ...(fill ? { fill: spreadsheetColor(fill) } : {}) }
    if (!text || !own) return host
    const cell = style === own.style ? own : { ...own, ...(style ? { style } : { style: undefined }) }
    const textWidth = textWidthOf(row * 16384 + col, cell, box.width, box.height)
    return { ...host, cell, ...(textWidth !== undefined ? { textWidth } : {}) }
  }
}

/**
 * The frame's own host source for a paged or worker-backed view (Ruling S3-A11Y-PAGED): per cell, the logical box the
 * viewport painter lays it out in (`SheetMetrics.sourceCellToLogical`, as its placements are) and its screen base style
 * (own style through the region layer), or undefined where the frame has no exact host (a cell not loaded yet, hidden,
 * outside the metrics): withheld, fail closed. Nothing is retained per call.
 */
export interface CfFrameHostSource {
  /** `textWidth` (Ruling S4-L) is the un-inset width of the text in the frame's own laid-out record of the cell, and
   * `cell` that cell (value and base style); without them an icon is withheld. */
  cell(row: number, col: number): { readonly box: Readonly<{ x: number; y: number; width: number; height: number }>; readonly base: XlsxCellStyle | undefined
    readonly textWidth?: number; readonly cell?: Pick<XlsxCell, 'value' | 'valueIsError' | 'style'> } | undefined
  /** The frame's own measurement of a cell as it lays it out (Task 29 fix round 1): the un-inset text width of `cell`
   * (carrying the CF-effective style) in `box`, from the viewer's layout host. Used when `cell()` gives no `textWidth`. */
  textWidth?(cell: XlsxCell, box: Readonly<{ width: number; height: number }>): number | undefined
}

/** Each bar cell's host as the viewport frame paints it: the source's box, and the fill the frame resolves under the
 * bar (`styleAt` over the base, by `spreadsheetColor`), so `barAt`/`describe` share the painter's `barWithheld`. */
export function frameBarHosts(source: CfFrameHostSource, styleAt: (row: number, col: number, base: XlsxCellStyle | undefined) => XlsxCellStyle | undefined): (row: number, col: number) => CfBarHost | undefined {
  return (row, col) => {
    const at = source.cell(row, col)
    if (!at) return undefined
    const style = styleAt(row, col, at.base), fill = style?.fillColor
    // The icon host's cell carries the CF-effective style the frame lays it out with (as `fill` does), and its text
    // width is measured on that cell when the source measures.
    const cell = at.cell && { ...at.cell, ...(style ? { style } : { style: undefined }) }
    const textWidth = at.textWidth ?? (cell && 'row' in cell && source.textWidth ? source.textWidth(cell as XlsxCell, at.box) : undefined)
    return { x: at.box.x, y: at.box.y, width: at.box.width, height: at.box.height, ...(fill ? { fill: spreadsheetColor(fill) } : {}),
      ...(textWidth !== undefined ? { textWidth } : {}),
      ...(cell ? { cell: { value: cell.value, ...(cell.valueIsError !== undefined ? { valueIsError: cell.valueIsError } : {}), ...(style ? { style } : {}) } } : {}) }
  }
}

/** Shared by resident assets and paged callers. paintRects scopes whole-axis
 * rules to the visible/printed source coordinates, including unstored blanks. */
export async function prepareCfPaint(sheet: XlsxSheet, options: {
  store?: SourceStore; memory?: CfMemory; request?: CfPrepareRequest; source?: CfValueSource; evidence?: CfEvidenceGate;
  paintRects?: readonly SourceRect[]; valuePolicy?: ValuePolicy; dateContext?: import('./predicates').CfDateContext;
  viewport?: { metrics: import('../view/contracts').SheetMetrics; boxes: readonly import('../view/contracts').LogicalCellBox[] };
  statisticsBudget?: () => CfWorkBudget;
  statisticsMaxUnits?: number;
  /** Stage 5 (Ruling S5-L): the document's evaluation policy (`cfDocumentEvaluationOf(doc)`); this preparation takes one
   * clock snapshot from `now` (default: the wall clock). Absent: formula rules are unresolved, Stage 1-4 rules unchanged. */
  evaluation?: CfDocumentEvaluation; now?: () => number;
  /** The measuring context resident icon hosts use for `iconAt` (default: an offscreen canvas, or none: withheld). */
  layoutContext?: () => CanvasRenderingContext2D | null | undefined;
  /** Task 36 (Ruling S5-K): the workbook's canonical stores (the resident `source`), for a host `store`. */
  workbook?: CfWorkbookStores;
  /** The static leg's document sheets (no overlay), for a host without a store. A host store with only document sheets
   * withholds cross-sheet rules: the two never mix. */
  documentSheets?: readonly XlsxSheet[];
} = {}): Promise<CfPaintBatch | undefined> {
  const rules = sheet.conditionalFormatting
  if (!rules) return undefined
  // An importing prefix cannot certify either absence or range statistics.
  if (options.store && options.store.completeness !== 'complete') return undefined
  let requested = options.request?.sourceRects ?? options.paintRects
  if (rules.rules.length > CF_LIMITS.rules || (requested?.length ?? 0) > CF_LIMITS.rangesPerSheet) return unavailable(sheet)
  let rangeCount = 0, propertyBytes = 0
  for (const rule of rules.rules) {
    if (rule.ranges.length > CF_LIMITS.rangesPerRule || (rangeCount += rule.ranges.length) > CF_LIMITS.rangesPerSheet) return unavailable(sheet)
    for (const key in rule.dxf) { const v = rule.dxf[key]; propertyBytes += 192 + key.length * 4 + (typeof v === 'string' ? v.length * 2 : 16) }
    for (const key of rule.writes) propertyBytes += 128 + key.length * 4
  }
  const memory = options.memory ?? createCfMemory()
  const hold = memory.reserve(2048 + Math.max(rangeCount, requested?.length ?? 0) * 512)
  if (hold.kind !== 'reserved') return unavailable(sheet)
  const evidence = options.evidence ?? cfRuntimeEvidence
  // Ruling S5-L: one value policy per preparation, the document's when the route is given its context.
  const { policy: valuePolicy, evaluation: routed } = cfRoutePolicy(options.evaluation ? cfEvaluationContextFor(options.evaluation, options.now ?? Date.now) : undefined, [options.valuePolicy, options.store?.valuePolicy], 'cached')
  // One clock per preparation (Ruling S5-M): the predicates read the instant the formulas do, or the formulas withhold.
  const { evaluation, dateContext } = cfAlignClock(routed, options.dateContext, () => sheet.semantics ? { dateSystem: sheet.semantics.dateSystem, timeZone: sheet.semantics.timeZone, epochNowMs: Date.now() } : undefined)
  const workbook = cfWorkbookFor(rules, sheet.name, options.store, options.workbook, options.documentSheets, memory, valuePolicy)
  const dependencyGen = cfDependencyGenOf(rules, sheet.name, workbook)
  const stamp = options.request?.stamp ?? {
    owner: options.store?.stamp.docId ?? 'static',
    sheetId: sheet.sheetId ?? sheet.name,
    workbookValueGen: options.store?.stamp.valueGen ?? 0,
    cfGen: options.store ? cfRulesVersion(rules) : 1,
    viewGen: 0,
    structureGen: options.store?.structureGen ?? 0,
    // Ruling S5-M: a sheet with a clock-dependent formula stamps its preparation's one clock (the local day number); every
    // other sheet keeps the Stage 1 value, so a clock-free sheet is never invalidated by the tick.
    clockGen: evaluation && cfRulesUseClock(rules, true) ? evaluation.clock.clockGen : options.dateContext?.clockSerial ? Math.floor(options.dateContext.clockSerial) : 0,
    ...(dependencyGen !== 0 ? { dependencyGen } : {}),
  }
  const signal = options.request?.signal ?? new AbortController().signal
  // The batch's own liveness checks (`styleAt`, bars, icons) read the live `dependencyGen` too (S5-K): a referenced sheet's
  // commit stops this batch painting synchronously, as a host commit does.
  const source = withCfDependencyGen(options.source ?? (options.store ? storeCfSource(options.store, stamp, memory) : residentSource(sheet, stamp, memory, valuePolicy)), rules, sheet.name, workbook)
  const hiddenRows: SourceRect[] = []
  for (const r of sheet.rows) {
    if (r.hidden) hiddenRows.push({ firstRow: r.index, lastRow: r.index, firstCol: 0, lastCol: 16383 })
  }
  const statisticsOptions = hiddenRows.length > 0 ? { hiddenRows } : undefined
  const sheetMerges: SourceRect[] = (sheet.mergeRanges ?? []).map(m => ({ firstRow: m.minRow, lastRow: m.maxRow, firstCol: m.minCol, lastCol: m.maxCol }))
  // Ruling I5: Cache range statistics per (sheet, rule range, data version) across preparations.
  const currentStatsKey = `${stamp.owner}|${stamp.sheetId}|${stamp.workbookValueGen}|${stamp.structureGen}`
  let retainedStats = sheetStatsSessions.get(sheet)
  if (!retainedStats || retainedStats.key !== currentStatsKey) {
    if (retainedStats) retainedStats.session.dispose()
    const session = createCfStatistics(memory, { maxEntries: 64, retainEntries: true })
    retainedStats = { session, key: currentStatsKey }
    sheetStatsSessions.set(sheet, retainedStats)
  }
  const statistics: CfPaintStatistics = { session: retainedStats.session, stamp, results: new Map() }
  let units = 0
  const work = async () => {
    if (signal?.aborted) throw new Error('CF cancelled')
    if (++units > CF_LIMITS.workUnits) throw new Error('CF paint work')
    if (units % 1024 === 0) await hostTask()
  }
  try {
    const scoped = Boolean(options.viewport && cfNeedsViewport(sheet))
    if (options.viewport && scoped) {
      const { metrics, boxes } = options.viewport
      // Union column intervals immediately, so duplicate rules/panes never
      // multiply retained row records. Every new row/interval is admitted first.
      const rowsBySource = new Map<number, Array<[number, number]>>()
      for (const box of boxes) {
        const rows = metrics.rows.visibleBetween(box.y, box.y + box.height), cols = metrics.cols.visibleBetween(box.x, box.x + box.width)
        for (const rr of rows.runs) for (let r = rr.start; r <= rr.end; r++) for (const cc of cols.runs) {
          const sourceRow = metrics.order.displayToSource(r)
          for (const rule of rules.rules) for (const range of rule.ranges) {
            await work()
            if (sourceRow < range.firstRow || sourceRow > range.lastRow) continue
            let from = Math.max(cc.start, range.firstCol), to = Math.min(cc.end, range.lastCol)
            if (from > to) continue
            let intervals = rowsBySource.get(sourceRow)
            if (!intervals) {
              if (rowsBySource.size >= CF_LIMITS.rangesPerSheet || !hold.grow(hold.bytes + 512)) throw new Error('CF viewport framing')
              rowsBySource.set(sourceRow, intervals = [])
            }
            let first = 0
            while (first < intervals.length && intervals[first][1] < from - 1) { await work(); first++ }
            let end = first
            while (end < intervals.length && intervals[end][0] <= to + 1) {
              await work(); from = Math.min(from, intervals[end][0]); to = Math.max(to, intervals[end][1]); end++
            }
            if (end === first && !hold.grow(hold.bytes + 256)) throw new Error('CF viewport interval')
            intervals.splice(first, end - first, [from, to])
          }
        }
      }
      // Sort in source order (including filtered/reordered rows), then extend
      // identical intervals on adjacent rows rather than rescanning row slabs.
      type Slab = { -readonly [K in keyof SourceRect]: SourceRect[K] }
      const visible: Slab[] = [], sourceRows = [...rowsBySource.keys()].sort((a,b)=>a-b)
      let previous = new Map<number, Slab>(), previousRow = -2
      for (const row of sourceRows) {
        const current = new Map<number, Slab>()
        for (const [from,to] of rowsBySource.get(row)!) {
          await work()
          const prior = previousRow === row - 1 ? previous.get(from) : undefined
          if (prior && prior.lastCol === to) { prior.lastRow = row; current.set(from, prior) }
          else {
            if (visible.length >= CF_LIMITS.rangesPerSheet || !hold.grow(hold.bytes + 512)) throw new Error('CF viewport slabs')
            const rect = {firstRow:row,lastRow:row,firstCol:from,lastCol:to}
            visible.push(rect); current.set(from,rect)
          }
        }
        previous = current; previousRow = row
      }
      requested = visible
    }
    const rects: SourceRect[] = []
    if (requested) rects.push(...requested)
    else {
      // computeMetrics includes merges, declarations and finite CF extents (bounded by I3). Open
      // tails do not enlarge it; their coverage is clipped to this paint region.
      const finite = cfExplicitBounds(sheet)
      let bottom = finite && sheet.rows.length === 0 ? finite.r1 : 0
      let right = finite && sheet.rows.length === 0 ? finite.c1 : 0
      for (const row of sheet.rows) {
        await work(); bottom = Math.max(bottom, row.index)
        for (const cell of row.cells) {
          await work(); bottom = Math.max(bottom, cellSourceRow(row, cell)); right = Math.max(right, cell.col)
        }
      }
      for (const m of sheet.mergeRanges ?? []) { await work(); bottom = Math.max(bottom, m.maxRow); right = Math.max(right, m.maxCol) }
      for (const r of regionStylesFor(sheet)?.rects ?? []) { await work(); bottom = Math.max(bottom, r.maxRow); right = Math.max(right, r.maxCol) }
      for (const [row, cols] of sheet.resolvedSparklines ?? []) { bottom = Math.max(bottom, row); for (const col of cols.keys()) { await work(); right = Math.max(right, col) } }
      const declared = sheet.cols.reduce((n, c) => Math.max(n, c.max), right)
      if (declared <= right + 64) right = declared
      if (sheet.drawings?.length) { bottom = Math.max(bottom, sheet.drawingMarkers?.maxRow ?? 0); right = Math.max(right, sheet.drawingMarkers?.maxCol ?? 0) }

      bottom = Math.min(bottom, STATIC_SHEET_BUDGET.maxGridRows - 1)
      right = Math.min(right, STATIC_SHEET_BUDGET.maxGridCols - 1)

      for (const rule of rules.rules) for (const r of rule.ranges) {
        const lastRow = Math.min(r.lastRow, bottom)
        const lastCol = r.lastCol === 16383 ? right : Math.min(r.lastCol, STATIC_SHEET_BUDGET.maxGridCols - 1)
        const rangeWork = Math.max(0, lastRow - r.firstRow + 1) * Math.max(0, lastCol - r.firstCol + 1)
        if (rangeWork > CF_LIMITS.workUnits) {
          recordDiagnostic(sheet, 'cf-resource-limit', 'Conditional formatting range exceeds maximum supported cell work units.')
          continue // Ruling I4: Refuse per range, never whole sheet
        }
        if (lastRow >= r.firstRow && lastCol >= r.firstCol) rects.push({ ...r, lastRow, lastCol })
      }
    }
    for (const r of rects) if (![r.firstRow, r.lastRow, r.firstCol, r.lastCol].every(Number.isSafeInteger) || r.firstRow < 0 || r.lastRow > 1048575 || r.firstCol < 0 || r.lastCol > 16383 || r.lastRow < r.firstRow || r.lastCol < r.firstCol) throw new Error('CF rectangle')
    const base = regionStylesFor(sheet)
    const colors = { theme: sheet.drawingTheme?.palette ?? sheet.tableStyleSource?.colors.theme, indexed: sheet.tableStyleSource?.colors.indexed }
    const states = new Map<number, CfCellState>()
    // Bar cells index the same retained states (no second copy): the painter walks only these.
    const bars = new Map<number, CfBarDescriptor>()
    // Icon cells likewise (Task 29); the 'showValue' token of a bar/icon conflict lives in their unresolved writes.
    const icons = new Map<number, CfIconDescriptor>()
    // Split the union into disjoint row slabs and merged column intervals. No
    // full-sheet cell set, and overlapping rules never multiply cell admission.
    const edges = scoped ? [] : [...new Set(rects.flatMap(r => [r.firstRow, r.lastRow + 1]))].sort((a, b) => a - b)
    const maxCells = Math.max(1, Math.min(256, Math.floor((1024 * 1024 - rules.rules.length * 640 - 2208) / (512 + propertyBytes))))
    const paintCells = options.viewport || options.request ? Math.min(CF_LIMITS.preparedCells, rects.reduce((n, r) => n + (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1), 0)) : 256
    if (!hold.grow(hold.bytes + paintCells * 2048)) throw new Error('CF paint workspace')
    let heldBytes = hold.bytes
    for (let e = 0; e < (scoped ? rects.length : edges.length - 1); e++) {
      const firstRow = scoped ? rects[e].firstRow : edges[e], endRow = scoped ? rects[e].lastRow + 1 : edges[e + 1]
      const intervals: Array<[number, number]> = []
      if (scoped) intervals.push([rects[e].firstCol, rects[e].lastCol])
      else for (const r of rects) { await work(); if (r.firstRow <= firstRow && r.lastRow >= firstRow) intervals.push([r.firstCol, r.lastCol]) }
      intervals.sort((a, b) => a[0] - b[0])
      const merged: Array<[number, number]> = []
      for (const [a, b] of intervals) { const last = merged.at(-1); if (last && a <= last[1] + 1) last[1] = Math.max(last[1], b); else merged.push([a, b]) }
      for (let row = firstRow; row < endRow; row++) for (const [from, to] of merged) for (let col = from; col <= to; col += maxCells) {
        if (signal?.aborted) throw new Error('CF cancelled')
        if (++units % 1024 === 0) await hostTask()
        const rect = { firstRow: row, lastRow: row, firstCol: col, lastCol: Math.min(to, col + maxCells - 1) }
        // Ruling I4: Per-batch work budget so static sheets (e.g. 3000x20) do not exhaust
        // the 1M whole-sheet limit. If a batch exceeds the budget, refusal is per-range.
        // Ruling S3-I1-UNITS: the batch is capped at CF_LIMITS.activeUnits work units, not wall time.
        const batchBudget = createCfWorkBudget({ signal })
        const statsBudgetFactory = options.statisticsBudget ?? (() => createCfWorkBudget({ signal, scan: true, maxUnits: options.statisticsMaxUnits ?? 10_000_000 }))
        const service = createCfService(rules, memory, { evidence, budget: () => batchBudget, statisticsBudget: statsBudgetFactory, statistics, statisticsOptions, dateContext, merges: sheetMerges, ...(evaluation ? { evaluation } : {}),
          sheetName: sheet.name, ...(workbook ? { workbook } : {}) })
        try {
          const lease = await service.prepareCf({ stamp, sourceRects: [rect], signal, grantBytes: options.request?.grantBytes ?? CF_LIMITS.scratchBytes }, source)
          try {
            for (let c = rect.firstCol; c <= rect.lastCol; c++) {
              const result = lease.value.at(row, c)
              if (result.kind !== 'ready' || (!Object.keys(result.value.properties).length && !result.value.visual?.colorScale && !result.value.visual?.dataBar && !result.value.visual?.iconSet)) continue
              const measured = measureCfPayload(result.value, memory)
              if (measured.kind !== 'ready' || !hold.grow(heldBytes + measured.value + 160)) throw new Error('CF retained paint admission')
              heldBytes += measured.value + 160
              // Ownership overlaps the service lease until the batch is drained.
              states.set(row * 16384 + c, result.value)
              if (result.value.visual?.dataBar) bars.set(row * 16384 + c, result.value.visual.dataBar)
              if (result.value.visual?.iconSet) icons.set(row * 16384 + c, result.value.visual.iconSet)
            }
          } finally { lease.release() }
          for (const [, outcome] of statistics.results) {
            if (outcome.kind === 'unresolved' && outcome.reason === 'cf-resource-limit') {
              recordDiagnostic(sheet, 'cf-resource-limit', 'Conditional formatting statistics range exceeded resource limits.')
            }
          }
        } catch (error) {
          if (error instanceof Error && (error.message === 'CF cancelled' || error.name === 'AbortError' || (error as any).code === 'cancelled') || signal?.aborted) {
            throw error
          }
          if (error instanceof Error && (error.message === 'CF work limit' || error.message === 'CF paint work' || (error as any).name === 'PackageError' || error.message.includes('resource-limit'))) {
            // Ruling I4 & M1: Refusal is per-range, not the whole sheet
            recordDiagnostic(sheet, 'cf-resource-limit', 'Conditional formatting batch exceeded work budget.')
            continue
          }
          throw error
        } finally { service.dispose() }
      }
    }
    let disposed = false
    const allows = (cap: string): boolean => evidence.allows(cap)
    // A bar paints only on its own unmerged host: merged hosts are unmeasured (bar:scope-merged-cell).
    const paintableAt = (row: number, col: number): CfBarDescriptor | undefined => {
      if (disposed || !sameCfStamp(source.stamp(stamp.sheetId), stamp)) return undefined
      const bar = bars.get(row * 16384 + col)
      if (!bar || sheet.mergeRanges?.some(m => row >= m.minRow && row <= m.maxRow && col >= m.minCol && col <= m.maxCol)) return undefined
      return barPaintable(bar, CF_BAR_RENDER_MODEL, allows) ? bar : undefined
    }
    // `barAt` reports a bar only where painting would not withhold it (the painter's own predicate).
    const hostAt = residentBarHosts(sheet, (row, col, own) => layer.styleAt(row, col, own), n => hold.grow(hold.bytes + n))
    const barAt = (row: number, col: number): CfBarDescriptor | undefined => {
      const bar = paintableAt(row, col)
      return bar && !barWithheld(bar, hostAt(row, col), allows) ? bar : undefined
    }
    // Ruling S4-R: each glyph class the batch's icons use is decoded and charged once, here; a sheet without icons
    // decodes nothing. A refused class withholds its icons (and their insets) for the batch's whole life.
    const glyphs: CfAdmittedGlyphs = icons.size ? admitBatchGlyphs(icons.values(), memory, () => createCfWorkBudget({ signal })) : { cacheFor: () => undefined, release() {} }
    const merged = (row: number, col: number): boolean => Boolean(sheet.mergeRanges?.some(m => row >= m.minRow && row <= m.maxRow && col >= m.minCol && col <= m.maxCol))
    // An icon paints only on its own unmerged host, in the batch's generation (as bars do).
    const iconCandidate = (row: number, col: number): CfIconDescriptor | undefined => {
      if (disposed || !sameCfStamp(source.stamp(stamp.sheetId), stamp)) return undefined
      const icon = icons.get(row * 16384 + col)
      return icon && !merged(row, col) ? icon : undefined
    }
    const gate: CfIconGate = { admitted: glyphs, hasBar: (row, col) => paintableAt(row, col) !== undefined, check: allows, rowPtOf: cfRowPtOfSheet(sheet), rowUniformOf: cfRowUniformOfSheet(sheet) }
    const showValueConflict = (row: number, col: number): boolean => states.get(row * 16384 + col)?.unresolvedWrites?.has('showValue') === true
    const iconHostAt = icons.size ? residentBarHosts(sheet, (row, col, own) => layer.styleAt(row, col, own), n => hold.grow(hold.bytes + n), { layoutContext: options.layoutContext }) : () => undefined
    const iconAt = (row: number, col: number): CfIconDescriptor | undefined => {
      const icon = iconCandidate(row, col)
      return icon && iconPaintsOn(icon, row, col, iconHostAt(row, col), gate) ? icon : undefined
    }
    const layer: RegionStyleLayer = {
      retainedBytes: hold.bytes,
      rects: [...base?.rects ?? [], ...rects.map(r => ({ minRow: r.firstRow, maxRow: r.lastRow, minCol: r.firstCol, maxCol: r.lastCol }))],
      styleAt(row, col, own, mode) {
        const style = base?.styleAt(row, col, own, mode) ?? own
        const merge = sheet.mergeRanges?.find(m => row >= m.minRow && row <= m.maxRow && col >= m.minCol && col <= m.maxCol)
        if (merge && !evidence.allows('layout:merge-hidden-endpoint') && !evidence.allows('layout:merge-hidden-anchor')) return style
        const state = !disposed && sameCfStamp(source.stamp(stamp.sheetId), stamp) && (states.get(row * 16384 + col) ?? (merge ? states.get(merge.minRow * 16384 + merge.minCol) : undefined))
        return state ? resolveCfCellStyle(style, state, colors) : style
      },
    }
    const batchDecisions = {
      textInsetAt: (row: number, col: number, host: CfBarHost, scale: CfBarScale) => iconTextInset(iconCandidate(row, col), row, col, host, scale, gate),
      hidesTextAt: (row: number, col: number, host: CfBarHost, scale: CfBarScale) => !showValueConflict(row, col)
        && cfHidesText(paintableAt(row, col), iconCandidate(row, col), row, col, host, scale, gate),
    }
    return { stamp, key: JSON.stringify([stamp.owner, stamp.sheetId, stamp.workbookValueGen, stamp.cfGen, stamp.viewGen, stamp.structureGen, stamp.clockGen]), layer,
      barAt,
      paintBars: (ctx, cellRect, visible, scale) => paintDataBars(ctx, bars, paintableAt, cellRect, visible, scale, allows),
      // One predicate over both visuals: a bar/icon showValue conflict (the 'showValue' token) keeps the text (S4-L).
      hidesTextAt: batchDecisions.hidesTextAt,
      iconAt,
      paintIcons: (ctx, cellRect, visible, scale) => paintIconCells(ctx, icons, iconCandidate, cellRect, visible, scale, gate),
      textInsetAt: batchDecisions.textInsetAt,
      layoutCell: (ctx, sheetArg, cell, x, y, w, h, measured, overflow, scale) => layoutCfCell(batchDecisions, ctx, sheetArg, cell, x, y, w, h, measured, overflow, scale),
      dispose() { if (disposed) return; disposed = true; states.clear(); bars.clear(); icons.clear(); glyphs.release(); hold.release() } }
  } catch (error) { hold.release(); return error instanceof CfPageFault && (error.fault === 'incomplete' || error.fault === 'stale') ? undefined : unavailable(sheet) } finally {
    for (const result of statistics.results.values()) if (result.kind === 'ready') result.value.release()
  }
}

export { createCfMemory as createPaintMemory } from './limits'

/** The CF layout of one static or print cell (Ruling S4-L): laid out un-inset, one host built from that layout (box, the
 * fill under it, the un-inset `textWidth`, the cell as laid out), then laid out again inset and/or with its text hidden
 * where `decide` says so on that same host. Shared by the paint batch and the paged scopes. */
export function layoutCfCell(decide: Pick<CfPaintBatch, 'textInsetAt' | 'hidesTextAt'>, ctx: CanvasRenderingContext2D, sheet: XlsxSheet, cell: XlsxCell,
  x: number, y: number, w: number, h: number, measured: MeasureCached, overflow: { left: number; right: number } | undefined, scale: CfBarScale): { layout: CellBodyLayout; host: CfBarHost } {
  const lay = (inset?: { left: number; right: number }, hide?: boolean) =>
    withTextInset(x, w, inset, (bx, bw) => layoutCellBody(ctx, sheet, cell, bx, y, bw, h, measured, overflow, undefined, undefined, hide))
  let layout = lay()
  const textWidth = cellTextWidth(layout)
  const host: CfBarHost = { x, y, width: w, height: h, ...(layout.fill ? { fill: layout.fill } : {}), ...(textWidth !== undefined ? { textWidth } : {}), cell }
  const inset = decide.textInsetAt(cell.row, cell.col, host, scale), hide = decide.hidesTextAt(cell.row, cell.col, host, scale)
  if (inset || hide) layout = lay(inset, hide)
  return { layout, host }
}

/** One facet per receiving document and parent. Parent admission precedes every
 * local allocation, including overlapping viewport and print leases. */
const memories = new WeakMap<object, Map<object | undefined, CfMemory>>()
export function paintMemoryFor(owner: object, parent?: import('../../charts/workspace').ChartParentLedger): CfMemory {
  let map = memories.get(owner)
  if (!map) memories.set(owner, map = new Map())
  let memory = map.get(parent)
  if (!memory) {
    memory = createCfMemory({ context: parent ? {
      limitBytes: CF_LIMITS.workbookMetadataBytes, canReserve: () => true, owns: () => false,
      reserve() {
        let bytes = 0
        return { owner: 'cf-paint', get bytes() { return bytes },
          trySet(next) { if (next > bytes && !parent.reserve(next - bytes, 'cf-paint')) return false; if (next < bytes) parent.release('cf-paint', bytes - next); bytes = next; return true },
          charge(next) { if (!this.trySet(bytes + next)) throw new Error('CF parent admission') },
          release() { parent.release('cf-paint', bytes); bytes = 0 },
          adopt() { throw new Error('CF ownership is held by its paint lease') },
        }
      },
    } : undefined })
    map.set(parent, memory)
  }
  return memory
}
