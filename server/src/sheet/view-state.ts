/**
 * Stage 4b (S4 §6.3, plan P11–P12): a session's per-sheet view state — the user's row and column sizes and the sheet's
 * monotonic `viewGen` — and the process ledger that caps it.
 *
 * - Sizes are clamped by `clampSize` (`RESIZE_LIMITS`), keyed by source row and source column, as local mode keys them.
 * - Caps (P12), both refused `view-state-limit`, never silently evicted: at most `MAX_SESSION_SIZES` (4,096) entries
 *   per axis per sheet, and at most `ViewStateLedger.maxBytes` (16 MiB) per process at `SESSION_SIZE_ENTRY_BYTES`
 *   (128 B) per entry. Only a new index is charged; a reset and `dispose` return the bytes.
 * - A refusal changes nothing, and `viewGen` advances only when the sizes change.
 *
 * Stage 4c (plan P2, P11, F4): each sheet also carries its canonical sort and filter view.
 * - A sheet is "ordered" when its view sorts or filters; at most `MAX_ORDERED_SHEETS` per session. A third is
 *   `'view-state-limit'`; sizes are not capped by it.
 * - An ordered view is admitted by the caller's `hold`, which returns its release, or `null` when refused. The new
 *   hold is taken before the replaced one is released; `dispose` releases every hold.
 * - `viewGen` advances when the sizes or the view change.
 */
import { clampSize, MAX_SESSION_SIZES, SESSION_SIZE_ENTRY_BYTES, type SessionSizes } from '../../../src/xlsx/view/axisSizes'
import { canonicalView, canonicalViewHash, IDENTITY_VIEW, isIdentityView, type CanonicalView } from './view-spec'

/** P11: at most two sheets per session hold a sort or a filter (S4 §6.3). */
export const MAX_ORDERED_SHEETS = 2

/** The canonical wire form of one sheet's sizes: axis `'col'` before `'row'`, each ascending by index. */
export type WireSizes = Array<{ axis: 'row' | 'col'; index: number; px: number }>

/** The process cap on view-state bytes (P12): 16 MiB by default. */
export class ViewStateLedger {
  private used = 0
  constructor(readonly maxBytes = 16 * 1024 * 1024) {}
  /** Charges `n` bytes; false (nothing charged) when they do not fit. */
  take(n: number): boolean {
    if (this.used + n > this.maxBytes) return false
    this.used += n
    return true
  }
  give(n: number): void { this.used = Math.max(0, this.used - n) }
  get bytes(): number { return this.used }
}

interface SheetView { gen: number; rows: Map<number, number>; cols: Map<number, number>; view: CanonicalView; viewHash: string; release: (() => void) | null }
const IDENTITY_HASH = canonicalViewHash(IDENTITY_VIEW)
const emptySheet = (): SheetView => ({ gen: 0, rows: new Map(), cols: new Map(), view: IDENTITY_VIEW, viewHash: IDENTITY_HASH, release: null })
export type ViewChange = { axis: 'row' | 'col'; index: number; size: number } | { reset: true }

/** One session's view state, per sheet index. Survives a resume; `dispose` at teardown. */
export class SessionViews {
  private readonly sheets = new Map<number, SheetView>()
  private charged = 0
  constructor(private readonly ledger: ViewStateLedger) {}

  viewGen(sheetIndex: number): number { return this.sheets.get(sheetIndex)?.gen ?? 0 }

  sizes(sheetIndex: number): SessionSizes {
    const v = this.sheets.get(sheetIndex)
    return Object.freeze({ rows: new Map(v?.rows ?? []), cols: new Map(v?.cols ?? []) })
  }

  canonical(sheetIndex: number): WireSizes {
    const v = this.sheets.get(sheetIndex)
    return v ? wireSizesOf(v) : []
  }

  /**
   * Applies one change. A size is clamped; `applied` is the clamped size. A new index past the axis cap or the
   * process ledger is `'view-state-limit'` and changes nothing. `viewGen` advances iff the sizes changed.
   */
  apply(sheetIndex: number, change: ViewChange): { viewGen: number; applied?: number } | 'view-state-limit' {
    let v = this.sheets.get(sheetIndex)
    if ('reset' in change) {
      if (!v || (!v.rows.size && !v.cols.size)) return { viewGen: v?.gen ?? 0 }
      const n = (v.rows.size + v.cols.size) * SESSION_SIZE_ENTRY_BYTES
      v.rows.clear(); v.cols.clear()
      this.release(n)
      return { viewGen: ++v.gen }
    }
    const applied = clampSize(change.axis, change.size)
    const map = v ? (change.axis === 'row' ? v.rows : v.cols) : undefined
    if (map?.get(change.index) === applied) return { viewGen: v!.gen, applied }
    if (!map?.has(change.index)) {
      if ((map?.size ?? 0) >= MAX_SESSION_SIZES) return 'view-state-limit'
      if (!this.ledger.take(SESSION_SIZE_ENTRY_BYTES)) return 'view-state-limit'
      this.charged += SESSION_SIZE_ENTRY_BYTES
    }
    if (!v) this.sheets.set(sheetIndex, v = emptySheet())
    ;(change.axis === 'row' ? v.rows : v.cols).set(change.index, applied)
    return { viewGen: ++v.gen, applied }
  }

  /** The sheet's canonical sort and filter view (the identity view when it has none). */
  view(sheetIndex: number): CanonicalView { return this.sheets.get(sheetIndex)?.view ?? IDENTITY_VIEW }

  /** How many sheets hold a sort or a filter. */
  orderedSheets(): number {
    let n = 0
    for (const v of this.sheets.values()) if (!isIdentityView(v.view)) n++
    return n
  }

  /**
   * Commits `next` as the sheet's view and returns the new `viewGen`. An equal view changes nothing. An ordered view
   * past `MAX_ORDERED_SHEETS`, or one whose `hold` is refused, is `'view-state-limit'` and changes nothing. The
   * identity view takes no hold. The replaced view's hold is released after the new one is taken.
   */
  commit(sheetIndex: number, next: CanonicalView, hold: () => (() => void) | null): number | 'view-state-limit' {
    const view = canonicalView(next), viewHash = canonicalViewHash(view)
    const v = this.sheets.get(sheetIndex)
    if ((v?.viewHash ?? IDENTITY_HASH) === viewHash) return v?.gen ?? 0
    const ordered = !isIdentityView(view)
    if (ordered && (!v || isIdentityView(v.view)) && this.orderedSheets() >= MAX_ORDERED_SHEETS) return 'view-state-limit'
    const release = ordered ? hold() : null
    if (ordered && !release) return 'view-state-limit'
    const s = v ?? emptySheet()
    if (!v) this.sheets.set(sheetIndex, s)
    const old = s.release
    s.view = view; s.viewHash = viewHash; s.release = release
    old?.()
    return ++s.gen
  }

  /** Returns every charged byte to the ledger and releases every view hold (idempotent). */
  dispose(): void {
    this.release(this.charged)
    for (const v of this.sheets.values()) { const r = v.release; v.release = null; r?.() }
    this.sheets.clear()
  }

  private release(n: number) {
    const m = Math.min(n, this.charged)
    this.charged -= m
    this.ledger.give(m)
  }
}

export function sessionSizesOf(w: WireSizes): SessionSizes {
  const rows = new Map<number, number>(), cols = new Map<number, number>()
  for (const e of w) (e.axis === 'row' ? rows : cols).set(e.index, e.px)
  return Object.freeze({ rows, cols })
}

/** Canonical order: every column, then every row, each ascending by index. */
export function wireSizesOf(s: Pick<SessionSizes, 'rows' | 'cols'>): WireSizes {
  const out: WireSizes = []
  for (const [index, px] of [...s.cols].sort((a, b) => a[0] - b[0])) out.push({ axis: 'col', index, px })
  for (const [index, px] of [...s.rows].sort((a, b) => a[0] - b[0])) out.push({ axis: 'row', index, px })
  return out
}
