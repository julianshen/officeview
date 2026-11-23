/**
 * Stage 4b (plan P2, P6–P10): the `sheet-text` pool job. Each op runs a local engine over the text view (Task 1), so
 * parity with local mode holds by construction: `fit` (`autoFitSize`), `find` (`WorkbookFind`) and `copy`
 * (`serializeSelection`) here; `a11y` in Task 8.
 *
 * `textInputs` is the one place the ops get the local viewer's inputs (`InteractiveSheetDoc.tsx` find ~1814–1857 and
 * copy ~2611–2659): the row order, the row and column visibility, the store and the scope-aware display text. Fit
 * (ruling F4) reads only the view's sheet and geometry, with the base style, as the local fit does (it passes no
 * scope). No op reads a clock: the session clock is in the job and the view.
 */
import type { Box, WireSelection } from '../../../src/stream/sheet-wire'
import { autoFitSize, createFitMeasure } from '../../../src/xlsx/view/autoFit'
import type { Measure } from '../../../src/xlsx/view/axisSizes'
import type { RowOrder, RowVisibility, SourceStore } from '../../../src/xlsx/view/contracts'
import { withUnavailableFormulas } from '../../../src/xlsx/view/unavailableFormulas'
import { serializeSelection } from '../../../src/xlsx/view/copy'
import type { GridSelection } from '../../../src/xlsx/view/contracts'
import { WorkbookFind } from '../../../src/xlsx/view/find'
import type { EffectiveFormatScope, FormatBudget } from '../../../src/xlsx/view/formatWorkspace'
import { createObjectWorkbookSource } from '../../../src/xlsx/view/source'
import { initialWorkbookOwner } from '../../../src/worker/viewerProtocol'
import { createIdentityOrder } from '../../../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../../../src/xlsx/view/visibility'
import type { FormatAllocation } from '../../../src/xlsx/formula/format'
import { cellDisplayText } from '../../../src/xlsx/render'
import type { XlsxCell, XlsxDocument } from '../../../src/xlsx/types'
import type { WorkbookSource } from '../../../src/worker/viewerSession'
import { textDocument, type TextView } from './text-view'
import type { WireSizes } from './view-state'

export type SheetTextJob =
  | { op: 'find'; query: string; matchCase: boolean; sheetIndex: number; sizes: WireSizes; limit: number }
  | { op: 'copy'; sheetIndex: number; selection: WireSelection; sizes: WireSizes; maxBytes: number }
  | { op: 'a11y'; sheetIndex: number; rows: Array<[number, number]>; cols: [number, number]; sizes: WireSizes }
  /** `regions` absent (never `[]`): `autoFitSize` then treats the pane as the one visible region, as the local fit does. */
  | { op: 'fit'; sheetIndex: number; axis: 'row' | 'col'; index: number; pane: Box; regions?: Box[]; sizes: WireSizes }
export type SheetTextReply =
  | { op: 'find'; hits: Array<[number, number]>; more: boolean } // [row, col] in local order; more: hits past `limit`
  | { op: 'copy'; text: string; paragraphs: number } | { op: 'copy'; refused: 'too-large'; bytes: number }
  | { op: 'a11y'; cells: Array<[number, number, string]> }
  | { op: 'fit'; size: number }

/** The measurement scratch one fit may hold (`createFitMeasure`'s admitted heap + native bytes); past it the fit is
 * refused (`AutoFitRefusedError`, answered `policy-refused`). */
export const FIT_MEASURE_BYTES = 16 * 1024 * 1024

/** The local viewer's text-op inputs over one text view (`source` is an object source over the text document, so its
 * sheet `view.sheetIndex` is `view.sheet`). */
export interface TextInputs {
  sheetId: string
  /** 4b has no sort: the identity order (display row = source row). */
  order: RowOrder
  /** Display-space row visibility: the authored hidden rows (4b has no filter). */
  visibility: RowVisibility
  /** Column visibility: the collapsed (zero-width) columns of the view's geometry are hidden. */
  columns: RowVisibility
  store(): SourceStore
  /** The store local find reads (`InteractiveSheetDoc.tsx` ~1824): `store()` with a formula lacking any value marked
   * unavailable (`withUnavailableFormulas`). Local copy reads the plain store, so only find uses this one. */
  findStore(): SourceStore
  /** The cell's style as paint and copy use it: a table or pivot region style resolved over its own. */
  styleOf(cell: Pick<XlsxCell, 'row' | 'col'>): XlsxCell['style']
  /** Display text as the local viewer forms it: the region-resolved style through the scope's CF-effective style. */
  textOf(cell: Pick<XlsxCell, 'row' | 'col' | 'value'>, allocation?: FormatAllocation, scope?: EffectiveFormatScope): string
}

function cellOf(view: TextView, row: number, col: number): XlsxCell | undefined {
  if (view.geometry.kind !== 'ready') return undefined
  return view.geometry.cellsOf(row)?.find((c) => c.col === col)
}

export function textInputs(view: TextView, source: WorkbookSource): TextInputs {
  const entry = source.sheets[view.sheetIndex]
  if (!entry) throw new RangeError(`text view sheet ${view.sheetIndex} is not in the source`)
  const sheetId = entry.sheetId
  const order = createIdentityOrder()
  const authored = source.geometry(sheetId).baseHidden.intervals.map(({ r0, r1 }) => ({ r0: order.sourceToDisplay(r0), r1: order.sourceToDisplay(r1) }))
  const hiddenCols: Array<{ r0: number; r1: number }> = []
  let colCount = 0
  if (view.geometry.kind === 'ready') {
    const cols = view.geometry.metrics.cols
    colCount = cols.logicalCount
    for (const o of cols.overrides) if (o.sizeCss === 0) hiddenCols.push({ r0: o.start, r1: o.end })
  }
  const styleOf = (cell: Pick<XlsxCell, 'row' | 'col'>) => {
    const own = cellOf(view, cell.row, cell.col)?.style
    return view.sheet.regionStyles?.styleAt(cell.row, cell.col, own, 'screen') ?? own
  }
  return {
    sheetId, order,
    visibility: visibilityFromHidden(createIntervalSet(authored)),
    columns: visibilityFromHidden(createIntervalSet(hiddenCols), colCount),
    store: () => source.store(sheetId),
    findStore: () => withUnavailableFormulas(source.store(sheetId)),
    styleOf,
    textOf(cell, allocation, scope) {
      const base = styleOf(cell)
      return cellDisplayText(view.sheet, { value: cell.value, style: scope?.effectiveStyle(cell.row, cell.col, base) ?? base }, allocation)
    },
  }
}

/** `box` clamped to `extent` (P9): its origin inside the sheet, its size no past the sheet's far edge. */
export function clampBox(box: Box, extent: { width: number; height: number }): Box {
  const x = Math.min(box.x, extent.width), y = Math.min(box.y, extent.height)
  return { x, y, width: Math.min(box.width, extent.width - x), height: Math.min(box.height, extent.height - y) }
}

/**
 * The fit size (P9): local `autoFitSize` over the view's (blanked) sheet and its sized geometry, with the pane and
 * regions clamped to the sheet's extent and `knownRows` the view's first rows, as on the local resident path.
 * `measureText` is the server backend's text measure; with it injected `createFitMeasure` builds no layout context.
 */
export function fitAnswer(view: TextView, job: Extract<SheetTextJob, { op: 'fit' }>, measureText: Measure): number {
  const geometry = view.geometry
  if (geometry.kind !== 'ready') throw new Error('sheet-fit: sheet geometry unavailable')
  const extent = { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }
  const fit = createFitMeasure({ admit: ({ heap, native }) => heap + native <= FIT_MEASURE_BYTES, measureText, sheet: view.sheet })
  return autoFitSize(job.axis === 'col' ? 'column' : 'row', job.index, {
    // Fix round 1: no regions is `undefined`, never `[]` (`autoFitSize` reads `regions ?? [pane]`; `[]` would mean no
    // visible rows at all).
    geometry, sheet: view.sheet, pane: clampBox(job.pane, extent), ...(job.regions?.length ? { regions: job.regions.map((r) => clampBox(r, extent)) } : {}),
    knownRows: (limit) => view.sheet.rows.slice(0, limit).map((r) => r.index),
    measure: fit.measure, allocation: fit.allocation,
  })
}

/** The formatting one find job may hold at once (per cell, while its display text is formed and matched); past it the
 * job fails, answered `policy-refused`. A hit's retained text is released when its cell is done (no `retain`). */
export const FIND_FORMAT_BYTES = 16 * 1024 * 1024
const FIND_DOC_ID = 'sheet-text'

function jobBudget(max: number): FormatBudget {
  let held = 0
  return {
    reserve(bytes) { if (held + bytes > max) throw new RangeError('sheet-find: formatting exceeds its budget'); held += bytes },
    release(bytes) { held -= bytes },
  }
}

/** `scope` lent to `WorkbookFind`, which disposes each scope `prepareScope` returns: the view owns (and releases) it. */
function borrowed(scope: EffectiveFormatScope): EffectiveFormatScope {
  return new Proxy(scope, {
    get(target, key) {
      if (key === 'dispose') return () => {}
      const v = Reflect.get(target, key, target) as unknown
      return typeof v === 'function' ? (v as (...a: unknown[]) => unknown).bind(target) : v
    },
  })
}

/**
 * One sheet's find (P6): local `WorkbookFind` over an object source of the text document (`workbook` with this sheet
 * replaced by its view), restricted to this sheet, with the local viewer's inputs (`textInputs`): the identity order,
 * the authored row visibility, `isColumnVisible` from the view's column sizes, the CF-effective `textOf`, and the
 * view's CF scope from `prepareScope`. Pages (`FIND_PAGE_HITS` each) are pulled until the find completes or holds more
 * than `limit` hits, so `more` is exact (`limit` 0 asks only whether any hit exists). Hits are `[row, col]`, in local
 * order. A source failure fails the job (never a silently partial answer).
 */
export async function findAnswer(view: TextView, job: Extract<SheetTextJob, { op: 'find' }>, workbook: XlsxDocument): Promise<{ hits: Array<[number, number]>; more: boolean }> {
  const source = createObjectWorkbookSource(textDocument(workbook, new Map([[view.sheetIndex, view]])), FIND_DOC_ID)
  const inputs = textInputs(view, source)
  const { sheetId } = inputs
  const owner = initialWorkbookOwner(FIND_DOC_ID, source.sheets)
  const v = owner.sheetVector.find((e) => e.sheetId === sheetId)!
  const scope = view.scope
  const find = new WorkbookFind({
    owner,
    sheets: [{
      sheetId, store: inputs.findStore(),
      identity: { docId: FIND_DOC_ID, sheetId, valueGen: v.valueGen, fontGen: owner.fontGen, viewGen: v.revision },
      order: inputs.order, visibility: inputs.visibility,
      // Zero-width columns are outside the visible search domain, as locally (`metrics.cols.size(col) > 0`).
      isColumnVisible: (col) => inputs.columns.visible(col),
    }],
  }, { text: job.query, caseSensitive: job.matchCase, sheetIds: [sheetId] }, {
    formatBudget: jobBudget(FIND_FORMAT_BYTES),
    textOf: (cell, _sheetId, allocation, s) => inputs.textOf(cell, allocation, s),
    ...(scope ? { prepareScope: () => Promise.resolve(borrowed(scope)) } : {}),
  })
  const hits: Array<[number, number]> = []
  try {
    for (;;) {
      const page = await find.nextPage()
      if (page.failures.length || page.incomplete.length) throw new Error('sheet-find: the text source failed')
      for (const h of page.hits) hits.push([h.source.row, h.source.col])
      if (page.complete || hits.length > job.limit) break
    }
  } finally { find.close() }
  return { hits: hits.slice(0, job.limit), more: hits.length > job.limit }
}

/** The local selection of a validated wire selection (P7): each range's rows rebuilt as `{kind: 'intervals'}`. */
export function gridSelectionOf(s: WireSelection): GridSelection {
  return {
    active: s.active ? { row: s.active[0], col: s.active[1] } : null,
    anchor: s.anchor ? { row: s.anchor[0], col: s.anchor[1] } : null,
    ranges: s.ranges.map((r) => ({ rows: { kind: 'intervals' as const, intervals: r.rows.map(([r0, r1]) => ({ r0, r1 })) }, cols: r.cols.map(([c0, c1]) => ({ c0, c1 })) })),
  }
}

const COPY_DOC_ID = 'sheet-copy'

/**
 * One sheet's copy (P8): local `serializeSelection` over an object source of the text document, with the local copy's
 * inputs (`textInputs`): the plain store (local copy does not mark unavailable formulas; such a cell refuses the copy
 * as locally), the identity order, the authored row visibility, the zero-width columns hidden, the CF-effective
 * `textOf`, and the view's CF scope (lent: the engine disposes each batch's scope). A withheld cell is blank in the view,
 * so it is an empty field. `paragraphs` counts the non-empty fields (each selected cell is read and formatted once).
 *
 * The local engine's own size refusal (`resource-limit`, before or while serializing) is `{refused: 'too-large', bytes:
 * 0}` (no byte count is known); text over `job.maxBytes` UTF-8 bytes is `{refused: 'too-large', bytes}`. Any other
 * outcome fails the job (answered `policy-refused`).
 */
export async function copyAnswer(view: TextView, job: Extract<SheetTextJob, { op: 'copy' }>, workbook: XlsxDocument): Promise<Extract<SheetTextReply, { op: 'copy' }>> {
  const source = createObjectWorkbookSource(textDocument(workbook, new Map([[view.sheetIndex, view]])), COPY_DOC_ID)
  const inputs = textInputs(view, source)
  const store = inputs.store()
  const owner = initialWorkbookOwner(COPY_DOC_ID, source.sheets)
  const v = owner.sheetVector.find((e) => e.sheetId === inputs.sheetId)!
  const identity = { docId: COPY_DOC_ID, sheetId: inputs.sheetId, valueGen: v.valueGen, fontGen: owner.fontGen, viewGen: v.revision }
  const scope = view.scope
  let paragraphs = 0
  const result = await serializeSelection({
    store, owner: identity, currentOwner: () => identity, allowed: () => true, signal: new AbortController().signal,
    selection: gridSelectionOf(job.selection), order: inputs.order, visibility: inputs.visibility, columns: inputs.columns,
    format: (cell, allocation, s) => {
      const text = inputs.textOf(cell, allocation, s)
      if (text !== '') paragraphs++
      return text
    },
    ...(scope ? { prepareScope: () => Promise.resolve(borrowed(scope)) } : {}),
  })
  if (result.kind === 'refused' && result.code === 'resource-limit') return { op: 'copy', refused: 'too-large', bytes: 0 }
  if (result.kind !== 'ready') throw new Error(`sheet-copy: ${result.kind}${result.kind === 'refused' ? ` (${result.code})` : ''}`)
  const bytes = Buffer.byteLength(result.text, 'utf8')
  if (bytes > job.maxBytes) return { op: 'copy', refused: 'too-large', bytes }
  return { op: 'copy', text: result.text, paragraphs }
}

/** Runs one `sheet-text` op over the text view of `job.sheetIndex` at `job.sizes` (`viewOf`: the host's cached views;
 * `workbook`: the servable workbook the text document is built over). */
export async function answerSheetText(job: SheetTextJob, deps: { viewOf: (sheetIndex: number, sizes: WireSizes) => Promise<TextView>; workbook: () => XlsxDocument; measureText: Measure }): Promise<SheetTextReply> {
  switch (job.op) {
    case 'fit': return { op: 'fit', size: fitAnswer(await deps.viewOf(job.sheetIndex, job.sizes), job, deps.measureText) }
    case 'find': {
      const view = await deps.viewOf(job.sheetIndex, job.sizes)
      return { op: 'find', ...await findAnswer(view, job, deps.workbook()) }
    }
    case 'copy': {
      const view = await deps.viewOf(job.sheetIndex, job.sizes)
      return await copyAnswer(view, job, deps.workbook())
    }
    // Task 8 replaces this.
    case 'a11y':
      throw new Error(`sheet-text: op '${job.op}' is not implemented`)
    default: {
      const never: never = job
      throw new Error(`sheet-text: unknown op ${JSON.stringify((never as { op?: unknown }).op)}`)
    }
  }
}
