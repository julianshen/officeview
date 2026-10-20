/**
 * Paintables metadata for charts (spec 6.6, 6.5). Before a document's pages are handed to a viewer:
 *
 * - each chart's user-shape pictures are decoded in this realm, so the first paint of a model transferred from a
 *   worker realm shows them. The wait is bounded: one shared deadline (`CHART_PICTURE_DEADLINE_MS`, driven by an
 *   injectable clock) and the caller's abort signal. A page whose pictures miss it is published without them and
 *   gets a `chart-user-shapes-deadline` diagnostic; the decode continues (its bytes stay charged until it
 *   settles) and the painter shows the picture at that page's next paint.
 * - each page gets its charts' accessible descriptions with their placement, in reading order.
 *
 * Admission comes before construction, per page: an input-derived bound (placements, one description's working
 * set, the worst-case retained entries) is reserved before anything is built, then reconciled to the exact
 * retained size. A page whose bound is refused gets one readable notice entry and a diagnostic; later pages are
 * unaffected. Over `CHART_DESCRIPTIONS_PER_PAGE` charts, the list ends with an "N more charts not listed" entry
 * and carries the true total.
 */
import type { ChartDocumentContext } from './context'
import type { ChartDiagnostic, ChartModel } from './model'
import { chartDescription, CHART_DESCRIPTION_LIMIT } from './describe'
import { chartUserShapesReady } from './userShapes'

/** A chart as placed on one page or slide, in page px. */
export interface PlacedChart { chart: ChartModel; x: number; y: number; width: number; height: number }
/** A chart's accessible description with its placement on the page (page px); a notice entry has no chart box. */
export interface PaintableChart { description: string; x: number; y: number; width: number; height: number; notice?: true }
/** One page's chart metadata: the listed entries and the page's true chart count. */
export interface PageChartMetadata { charts: PaintableChart[]; total: number }
/** The document's pages, as the engine reads them: an upper bound per page (no allocation), then the placements. */
export interface ChartPageSource {
  readonly pages: number
  /** An upper bound on the page's charts, computed without building them. */
  bound(page: number): number
  /** The page's charts in reading order (at most `bound(page)`). */
  collect(page: number): PlacedChart[]
}
/** Resolves `after(ms)` once that much clock time has passed; `cancel` stops it. Tests drive it by hand. */
export interface ChartClock { after(ms: number): { done: Promise<void>; cancel(): void } }
export interface ChartPaintablesOptions {
  /** Aborts the picture wait (the owner unmounted or navigated). */
  signal?: AbortSignal
  clock?: ChartClock
  /** The shared picture-decode deadline in ms (default `CHART_PICTURE_DEADLINE_MS`). */
  deadlineMs?: number
  /** False: decode pictures only (no descriptions are built or retained). */
  describe?: boolean
}
export interface ChartPaintables { pages: PageChartMetadata[]; diagnostics: Array<ChartDiagnostic & { page: number }>; release(): void }

export const CHART_DESCRIPTIONS_PER_PAGE = 64
export const CHART_PICTURE_DEADLINE_MS = 1500
/** Charged per placement while a page is built (the record, its promise slot). */
const PLACED_BYTES = 96
/** Charged per retained entry beside its string. */
export const ENTRY_BYTES = 64
/** The transient working set of building one description (bounded by describe.ts's limits). */
export const DESCRIPTION_WORK_BYTES = 8 * CHART_DESCRIPTION_LIMIT + 4096
const PAGE_BYTES = 128
const NOTICE_BYTES = ENTRY_BYTES + 2 * 160

const realClock: ChartClock = {
  after(ms) {
    let id: ReturnType<typeof setTimeout> | undefined
    const done = new Promise<void>(resolve => { id = setTimeout(resolve, ms) })
    return { done, cancel: () => clearTimeout(id) }
  },
}
let serial = 0

/** Retained per page record (the record and its entry array). */
export const RECORD_BYTES = 96
/** Retained per described model in the shared description map. */
export const MAP_BYTES = 48
/** Retained per diagnostic beside its reason string. */
const DIAGNOSTIC_BYTES = 64
/** Charged per page whose pictures are awaited (its promise, index and settle flag), for the wait's lifetime. */
const PENDING_BYTES = 64
const stringBytes = (text: string) => 2 * text.length
const diagnosticBytes = (reason: string) => DIAGNOSTIC_BYTES + stringBytes(reason)

/**
 * Static, shared values that need no allocation or charge, used only when the workspace cannot hold even a page
 * record: the page still exposes a readable notice, and the job records one shared diagnostic.
 */
const REFUSED_NOTICE: PaintableChart = Object.freeze({ description: 'Chart descriptions for this page could not be kept in memory.', x: 0, y: 0, width: 0, height: 0, notice: true as const })
const REFUSED_LIST: PaintableChart[] = Object.freeze([REFUSED_NOTICE]) as unknown as PaintableChart[]
/** A refused page whose record could not be charged: its true chart count is unknown to the viewer (`total` 0). */
const REFUSED_PAGE: PageChartMetadata = Object.freeze({ charts: REFUSED_LIST, total: 0 }) as PageChartMetadata
const EMPTY_PAGE: PageChartMetadata = Object.freeze({ charts: Object.freeze([]) as unknown as PaintableChart[], total: 0 }) as PageChartMetadata
const UNRECORDED: ChartDiagnostic & { page: number } = Object.freeze({ page: -1, code: 'chart-description-refused', reason: 'Chart metadata could not be kept in memory for one or more pages' })

export async function prepareChartPaintables(context: ChartDocumentContext, source: ChartPageSource, options: ChartPaintablesOptions = {}): Promise<ChartPaintables> {
  const workspace = context.workspace
  const signals = [options.signal, context.signal].filter((s): s is AbortSignal => !!s)
  const id = ++serial, retained = `chart-descriptions#${id}`, waiting = `${retained}:wait`
  let work = `${retained}:page`
  const describe = options.describe !== false
  const diagnostics: ChartPaintables['diagnostics'] = []
  const described = new Map<ChartModel, string>()
  const pages: PageChartMetadata[] = []
  const pending: Array<Promise<boolean>> = []
  const pendingPages: number[] = []
  /** Records a diagnostic with its retained charge, or the one shared static diagnostic when that is refused. */
  const diagnose = (page: number, code: string, reason: string) => {
    if (workspace.reserve(diagnosticBytes(reason), retained)) diagnostics.push({ page, code, reason })
    else if (!diagnostics.includes(UNRECORDED)) diagnostics.push(UNRECORDED)
  }
  try {
    for (let page = 0; page < source.pages; page++) {
      const bound = Math.max(0, Math.floor(source.bound(page)))
      if (!bound) { pages.push(EMPTY_PAGE); continue }
      const listed = Math.min(bound, CHART_DESCRIPTIONS_PER_PAGE)
      work = `${retained}:page${page}`
      // Admission before construction: placements, one description's working set and the worst-case retention
      // (the page record, its entries and strings, the summary note and its diagnostic).
      const ask = PAGE_BYTES + RECORD_BYTES + bound * PLACED_BYTES
        + (describe ? DESCRIPTION_WORK_BYTES + listed * (ENTRY_BYTES + MAP_BYTES + 2 * CHART_DESCRIPTION_LIMIT) + NOTICE_BYTES + DIAGNOSTIC_BYTES + 2 * 160 : 0)
      if (!workspace.reserve(ask, work)) {
        if (!describe) { pages.push(EMPTY_PAGE); continue }
        const reason = `Chart descriptions for this ${source.pages > 1 ? 'page' : 'document'} could not be kept in memory (${bound} ${bound === 1 ? 'chart' : 'charts'}).`
        // A charged record keeps the true total; otherwise the static notice still makes the omission readable.
        if (workspace.reserve(RECORD_BYTES, retained)) pages.push({ charts: REFUSED_LIST, total: bound })
        else pages.push(REFUSED_PAGE)
        diagnose(page, 'chart-description-refused', reason)
        continue
      }
      const placed = source.collect(page).slice(0, bound)
      const total = placed.length
      // The page's pictures start decoding now; the wait below is shared and bounded (its bookkeeping charged).
      let images = false
      for (const p of placed) if (p.chart.userShapes?.images.length) { images = true; break }
      if (images && workspace.reserve(PENDING_BYTES, waiting)) {
        pending.push(Promise.all(placed.filter(p => p.chart.userShapes?.images.length).map(p => chartUserShapesReady(p.chart, workspace).catch(() => {}))).then(() => true))
        pendingPages.push(page)
      }
      const charts: PaintableChart[] = []
      let exact = RECORD_BYTES
      if (describe) {
        for (const p of placed.slice(0, CHART_DESCRIPTIONS_PER_PAGE)) {
          let description = described.get(p.chart)
          if (description === undefined) { description = chartDescription(p.chart); described.set(p.chart, description); exact += stringBytes(description) + MAP_BYTES }
          exact += ENTRY_BYTES
          charts.push({ description, x: p.x, y: p.y, width: p.width, height: p.height })
        }
        if (total > CHART_DESCRIPTIONS_PER_PAGE) {
          const more = total - CHART_DESCRIPTIONS_PER_PAGE
          const note = `${more} more ${more === 1 ? 'chart' : 'charts'} not listed.`
          charts.push({ description: note, x: 0, y: 0, width: 0, height: 0, notice: true })
          const reason = `${total} charts on the page; the first ${CHART_DESCRIPTIONS_PER_PAGE} are described`
          exact += ENTRY_BYTES + stringBytes(note) + diagnosticBytes(reason)
          diagnostics.push({ page, code: 'chart-description-limit', reason })
        }
      }
      // Reconcile: keep the exact retained bytes (record, entries, strings, diagnostic), refund the transient rest.
      workspace.transfer(work, retained, Math.min(exact, ask))
      workspace.release(work)
      pages.push({ charts, total })
    }
    if (pending.length) {
      const settled = pending.map(() => false)
      pending.forEach((p, i) => { void p.then(() => { settled[i] = true }) })
      const timer = (options.clock ?? realClock).after(options.deadlineMs ?? CHART_PICTURE_DEADLINE_MS)
      let onAbort = () => {}
      const aborted = new Promise<void>(resolve => {
        onAbort = () => resolve()
        for (const signal of signals) { if (signal.aborted) resolve(); else signal.addEventListener('abort', onAbort, { once: true }) }
      })
      try { await Promise.race([Promise.all(pending), timer.done, aborted]) }
      finally {
        timer.cancel()
        for (const signal of signals) signal.removeEventListener('abort', onAbort)
      }
      // A microtask turn lets decodes that finished with the deadline record themselves.
      await Promise.resolve()
      const wasAborted = signals.some(s => s.aborted)
      if (!wasAborted) settled.forEach((done, i) => {
        if (!done) diagnose(pendingPages[i], 'chart-user-shapes-deadline', `Chart pictures on this page were not decoded within ${options.deadlineMs ?? CHART_PICTURE_DEADLINE_MS} ms; they appear when the page is next painted`)
      })
    }
  } catch (error) {
    // Every reservation this job owns is released; in-flight decodes keep their own charges until they settle.
    workspace.release(work)
    workspace.release(retained)
    throw error
  } finally {
    workspace.release(waiting)
  }
  return { pages, diagnostics, release: () => workspace.release(retained) }
}
