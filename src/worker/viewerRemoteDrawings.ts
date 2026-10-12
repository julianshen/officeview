/**
 * Main-thread drawing content of worker-backed sheets (S13.2; spec RENDER-02,
 * AC-RENDER-01, §10.1). Once a sheet's structure is complete (its row/column
 * metadata is final), one bounded `read { op: 'drawings' }` brings the sheet's
 * drawings, sheet-local pictures and theme; the paged geometry then paints them
 * exactly like resident drawings and the picture lease admits their images
 * through the same budgets.
 *
 * Ownership: a reply is adopted only for the request this cache still waits
 * for, from the sheet's current source owner (document, sheet, value
 * generation); anything else is dropped and, while the sheet is open, asked
 * again. Adoption bumps the sheet's structure generation (the host's
 * `changed`), so geometry and tiles keyed by it rebuild. Retained payloads are
 * held to one byte budget; the least recently used sheets give theirs back
 * (they are read again when shown). What could not be shown is a diagnostic
 * the viewer displays, never an empty drawing layer.
 */
import type { CacheIdentity, SheetDrawings, WorkerDiagnostic, WorkerReply } from '../xlsx/view/contracts'
import type { XlsxDocument, XlsxDrawing, XlsxImage, XlsxSheet } from '../xlsx/types'
import type { ThemeContext } from '../drawing/style'
import { drawingQuery, type SheetGeometry } from '../xlsx/view/viewport'
import { DRAWINGS_UNAVAILABLE } from '../xlsx/view/drawingPayload'
import type { CommandOf } from './viewerProtocol'

/** Re-asks after a stale/cancelled/still-loading answer, a bounded number of times. */
const RETRY_MS = 250
const MAX_ATTEMPTS = 20
/** After the attempts ran out on transient answers (e.g. a long burst of view
 * commits), showing the sheet again asks once more, at most this often, until
 * the sheet's total attempts reach MAX_TOTAL_ATTEMPTS (then the failure stays). */
const REARM_MS = 5000
const MAX_TOTAL_ATTEMPTS = 40

export interface DrawingHost {
  identity(sheetId: string): CacheIdentity
  loadState(sheetId: string): { completeness: 'complete' | 'loading' | 'incomplete' }
  /** Resolves a reply, or a non-object marker once the document closed or the request was cancelled. */
  request(command: CommandOf<'read'>, signal?: AbortSignal): Promise<WorkerReply<'read'> | symbol>
  nextRequestId(): string
  /** The sheet's drawing content changed: bump its structure generation and re-render. */
  changed(sheetId: string): void
}

/** One sheet's drawing content as the viewer reads it (fields change in place; `gen` tells when). */
export interface PagedSheetDrawings {
  readonly state: 'waiting' | 'loading' | 'ready' | 'failed'
  readonly gen: number
  readonly drawings: readonly XlsxDrawing[]
  readonly images: readonly XlsxImage[]
  readonly theme?: ThemeContext
  readonly diagnostics: readonly WorkerDiagnostic[]
  readonly bytes: number
  drawingsIn: SheetGeometry['drawingsIn']
}

interface Entry {
  state: PagedSheetDrawings['state']
  gen: number
  drawings: readonly XlsxDrawing[]
  images: readonly XlsxImage[]
  theme?: ThemeContext
  diagnostics: readonly WorkerDiagnostic[]
  bytes: number
  drawingsIn: SheetGeometry['drawingsIn']
  /** Identifies the request whose reply may be adopted. */
  seq: number
  /** Reads sent since the last reset (never reset by a re-arm). */
  attempts: number
  /** Retries continue while attempts are below this. */
  limit: number
  /** Set when the attempts ran out on transient answers: when get() may re-arm. */
  rearmAt?: number
  controller?: AbortController
  timer?: ReturnType<typeof setTimeout>
  /** Derived sheet/document shells, memoized per generation. */
  shells?: { gen: number; base: XlsxSheet; sheet: XlsxSheet }
}
const none: SheetGeometry['drawingsIn'] = () => []
const sameOwner = (a: CacheIdentity, b: CacheIdentity) => a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen

export class RemoteDrawingCache {
  private readonly entries = new Map<string, Entry>()
  private closed = false
  constructor(private readonly host: DrawingHost, readonly budgetBytes: number) {}

  /** The sheet's drawing content; asks the worker once the sheet's structure is complete. */
  get(sheetId: string): PagedSheetDrawings {
    let entry = this.entries.get(sheetId)
    if (entry) {
      // Recency order for eviction.
      this.entries.delete(sheetId)
      this.entries.set(sheetId, entry)
    } else {
      entry = { state: 'waiting', gen: 0, drawings: [], images: [], diagnostics: [], bytes: 0, drawingsIn: none, seq: 0, attempts: 0, limit: MAX_ATTEMPTS }
      this.entries.set(sheetId, entry)
    }
    if (entry.state === 'failed' && entry.rearmAt !== undefined && Date.now() >= entry.rearmAt) {
      // A transient failure is not final: the notice stays until a new answer replaces it.
      entry.state = 'waiting'
      entry.limit = entry.attempts + 1
      entry.rearmAt = undefined
    }
    if (entry.state === 'waiting' && !entry.timer && this.host.loadState(sheetId).completeness === 'complete') void this.load(sheetId, entry)
    return entry
  }
  /** Bytes of the retained payloads. */
  get retainedBytes(): number {
    let total = 0
    for (const entry of this.entries.values()) total += entry.bytes
    return total
  }
  /** The sheet's structure completed: a waiting request may go now. */
  structureComplete(sheetId: string): void {
    const entry = this.entries.get(sheetId)
    if (entry?.state === 'waiting' && !entry.timer) void this.load(sheetId, entry)
  }
  /** Placement depends on the layout: a newer layout reads the content again. */
  invalidate(sheetId: string): void {
    const entry = this.entries.get(sheetId)
    if (!entry || entry.state === 'waiting') return
    this.reset(entry)
    entry.gen++
    this.host.changed(sheetId)
    this.structureComplete(sheetId)
  }
  close(): void {
    this.closed = true
    for (const entry of this.entries.values()) this.reset(entry)
    this.entries.clear()
  }

  private reset(entry: Entry): void {
    entry.seq++
    entry.controller?.abort()
    entry.controller = undefined
    if (entry.timer) clearTimeout(entry.timer)
    entry.timer = undefined
    entry.state = 'waiting'
    entry.attempts = 0
    entry.limit = MAX_ATTEMPTS
    entry.rearmAt = undefined
    entry.drawings = []
    entry.images = []
    entry.theme = undefined
    entry.diagnostics = []
    entry.bytes = 0
    entry.drawingsIn = none
    entry.shells = undefined
  }

  private async load(sheetId: string, entry: Entry): Promise<void> {
    if (this.closed) return
    const seq = ++entry.seq
    const owner = this.host.identity(sheetId)
    const controller = entry.controller = new AbortController()
    entry.state = 'loading'
    entry.attempts++
    const out = await this.host.request({ kind: 'read', requestId: this.host.nextRequestId(), owner, query: { op: 'drawings' } }, controller.signal)
    // Superseded (reset, close) or no longer this sheet's source owner: never adopted.
    if (this.closed || entry.seq !== seq || this.entries.get(sheetId) !== entry) return
    entry.controller = undefined
    if (typeof out !== 'object') { entry.state = 'waiting'; return }
    const result = out.kind === 'complete' && out.result.kind === 'read' && out.result.result.op === 'drawings' ? out.result.result.result : undefined
    // Content of another owner (or of a source generation that moved on meanwhile) is never adopted.
    const owned = sameOwner(out.owner, owner) && sameOwner(owner, this.host.identity(sheetId))
    if (result && owned && result.completeness === 'complete') {
      this.adopt(sheetId, entry, result)
      return
    }
    // A foreign/stale owner, a cancelled read or a sheet still loading is asked
    // again; anything else (or too many attempts) is a visible failure.
    const retryable = !owned || (result ? result.error?.retryable === true : out.kind === 'cancelled' || (out.kind === 'error' && out.retryable))
    if (retryable && entry.attempts < entry.limit) {
      entry.state = 'waiting'
      entry.timer = setTimeout(() => { entry.timer = undefined; if (entry.seq === seq) void this.load(sheetId, entry) }, RETRY_MS)
      return
    }
    entry.state = 'failed'
    entry.rearmAt = retryable && entry.attempts < MAX_TOTAL_ATTEMPTS ? Date.now() + REARM_MS : undefined
    entry.diagnostics = result?.diagnostics.length ? result.diagnostics
      : [{ code: DRAWINGS_UNAVAILABLE, message: `Drawings could not be loaded${out.kind === 'error' && out.diagnostics[0] ? `: ${out.diagnostics[0].message}` : ''}` }]
    entry.gen++
    this.host.changed(sheetId)
  }

  private adopt(sheetId: string, entry: Entry, result: SheetDrawings): void {
    entry.state = 'ready'
    entry.drawings = result.drawings
    entry.images = result.images
    entry.theme = result.theme
    entry.diagnostics = result.diagnostics
    entry.bytes = result.bytes
    entry.drawingsIn = drawingQuery(result.drawings)
    entry.shells = undefined
    entry.gen++
    // Least recently used sheets give their payloads back first; the one just adopted stays.
    for (const [id, other] of this.entries) {
      if (this.retainedBytes <= this.budgetBytes) break
      if (other === entry || other.state !== 'ready') continue
      this.reset(other)
      other.gen++
      this.host.changed(id)
    }
    this.host.changed(sheetId)
  }
}

/** The paged shell sheet with its drawing content (painting theme, picture lease), stable per generation. */
export function sheetWithDrawings(base: XlsxSheet, content: PagedSheetDrawings): XlsxSheet {
  const entry = content as Entry
  if (entry.state !== 'ready') return base
  if (entry.shells?.gen === entry.gen && entry.shells.base === base) return entry.shells.sheet
  const sheet: XlsxSheet = { ...base, drawings: entry.drawings as XlsxDrawing[], ...(entry.theme ? { drawingTheme: entry.theme } : {}) }
  entry.shells = { gen: entry.gen, base, sheet }
  return sheet
}
/** A document whose sheet `index` carries its drawing content and sheet-local pictures (for the picture lease). */
export function documentWithDrawings(doc: XlsxDocument, index: number, content: PagedSheetDrawings): XlsxDocument {
  if (content.state !== 'ready' || !doc.sheets[index]) return doc
  return { ...doc, sheets: doc.sheets.map((sheet, i) => i === index ? sheetWithDrawings(sheet, content) : sheet), images: content.images as XlsxImage[] }
}
/** Status notices for what the large-file view could not show (bounded, deduplicated). */
export function drawingNotices(content: PagedSheetDrawings): string[] {
  return [...new Set(content.diagnostics.map(d => d.message))].slice(0, 4)
}
