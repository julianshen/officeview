/**
 * Stage 4b (S4 §6.7, plan P10, P12a, Open item 2): the a11y grid's text on a stream, fetched from the server's
 * `sheet-a11y` by the manifest's `caps.a11y` mode and cached per `(sheetId, viewGen, cell)`.
 *
 * - **unit:** once the shown window has rested 250 ms, its uncached cells are asked for, in rects of at most 2,000
 *   cells (rows grouped by their uncached column span, one group per column run). At most 30 requests go out per
 *   60 s (all modes), one at a time; a newer rest replaces the requests not yet sent. A `rate-limited` refusal is
 *   retried once, 1 s later.
 * - **focused:** once the active cell has rested 400 ms, that one cell is asked for (1 × 1), unless it is cached.
 *   The server counts a non-empty cell against the session's copy quota, so a cached cell is never asked again.
 * - **off:** nothing is fetched (the server would refuse `text-unavailable`).
 *
 * Every asked cell is cached, `''` for one the server left out (blank or withheld: "no text"), so it is not asked
 * again. Only the newest viewGen of each sheet is kept: a reply filed under an older one is dropped, and a
 * `stale-view` refusal drops the reply (the next rest under the new viewGen asks again). `text-unavailable`,
 * `copy-quota-exhausted` and `closed` stop all fetching; any other refusal stops those cells loading until the next
 * rest. Nothing reaches the server outside `a11y` (the caller passes `StreamWorkbook.a11y`, which re-fetches the
 * layout on `stale-view`).
 */
import type { RemoteSheets } from '../../stream/remote/remote-doc'
import type { SourceCell } from '../../xlsx/view/coordinates'

export type StreamA11yMode = 'unit' | 'focused' | 'off'
export type StreamA11yFetch = NonNullable<RemoteSheets['view']>['a11y']

/** P10: one `sheet-a11y` asks at most this many cells. */
export const STREAM_A11Y_MAX_CELLS = 2000
/** P12a: the unit-mode window rest. */
export const STREAM_A11Y_UNIT_REST_MS = 250
/** Open item 2: the focused-mode active-cell rest. */
export const STREAM_A11Y_FOCUSED_REST_MS = 400
/** P12a: at most this many requests per `STREAM_A11Y_RATE_WINDOW_MS`. */
export const STREAM_A11Y_RATE_REQUESTS = 30
export const STREAM_A11Y_RATE_WINDOW_MS = 60_000
/** P12a: a `rate-limited` refusal is retried once, this long after. */
export const STREAM_A11Y_RETRY_MS = 1000
/** Unit mode keeps at most this many cached cells per sheet (the oldest go first). */
const MAX_CACHED_CELLS = 100_000
const COLS = 16_384

export interface StreamA11yText {
  /** The cached text of a cell under `viewGen` (`''`: asked, no text), or `undefined` when not cached. */
  textAt(sheetId: string, viewGen: number, at: SourceCell): string | undefined
  /** The cell is uncached and a fetch for it is waiting or in flight. */
  pending(sheetId: string, viewGen: number, at: SourceCell): boolean
  /** The shown window: source row intervals, and one column range or several disjoint column runs. Unit mode only. */
  showWindow(sheetId: string, viewGen: number, rows: Array<[number, number]>, cols: [number, number] | ReadonlyArray<readonly [number, number]>): void
  /** The active cell. Focused mode only. */
  focus(sheetId: string, viewGen: number, at: SourceCell): void
  onChange(cb: () => void): () => void
  dispose(): void
}

interface Request { sheetId: string; viewGen: number; rows: Array<[number, number]>; cols: [number, number]; keys: Set<number>; retried: boolean }
interface SheetCache { viewGen: number; cells: Map<number, string>; failed: Set<number> }

const keyOf = (row: number, col: number) => row * COLS + col
/** Ascending rows as inclusive intervals. */
function intervals(rows: readonly number[]): Array<[number, number]> {
  const out: Array<[number, number]> = []
  for (const r of rows) {
    const last = out[out.length - 1]
    if (last && last[1] === r - 1) last[1] = r
    else out.push([r, r])
  }
  return out
}

export function createStreamA11yText(o: { mode: StreamA11yMode; a11y: StreamA11yFetch; now?: () => number; setTimer?: typeof setTimeout }): StreamA11yText {
  const { mode, a11y } = o
  const now = () => (o.now ?? Date.now)()
  const later = (fn: () => void, ms: number) => { (o.setTimer ?? setTimeout)(fn, ms) }
  const listeners = new Set<() => void>()
  const caches = new Map<string, SheetCache>()
  const sent: number[] = []
  let queue: Request[] = []
  let inflight: Request | null = null
  let disposed = false
  let halted = false
  let window: { sheetId: string; viewGen: number; rows: Array<[number, number]>; cols: Array<readonly [number, number]>; sig: string } | null = null
  let focused: { sheetId: string; viewGen: number; key: number } | null = null
  // Debounce by generation: a timer whose token is no longer current does nothing.
  let restToken = 0, pumpToken = 0
  let resting = false
  let holdUntil = 0

  const notify = () => { if (!disposed) for (const cb of [...listeners]) cb() }
  /** The sheet's cache under `viewGen`; a newer viewGen replaces an older one, an older one gets none. */
  const cacheFor = (sheetId: string, viewGen: number, create: boolean): SheetCache | undefined => {
    const cache = caches.get(sheetId)
    if (cache && cache.viewGen === viewGen) return cache
    if (!create || (cache && cache.viewGen > viewGen)) return undefined
    const fresh = { viewGen, cells: new Map<number, string>(), failed: new Set<number>() }
    caches.set(sheetId, fresh)
    return fresh
  }
  const inFlight = (sheetId: string, viewGen: number, key: number) => !!inflight && inflight.sheetId === sheetId && inflight.viewGen === viewGen && inflight.keys.has(key)
  const queued = (sheetId: string, viewGen: number, key: number) => queue.some(r => r.sheetId === sheetId && r.viewGen === viewGen && r.keys.has(key))

  const request = (sheetId: string, viewGen: number, rows: number[], cols: [number, number]): Request => {
    const keys = new Set<number>()
    for (const r of rows) for (let c = cols[0]; c <= cols[1]; c++) keys.add(keyOf(r, c))
    return { sheetId, viewGen, rows: intervals(rows), cols, keys, retried: false }
  }
  /** The shown window's uncached cells as requests of at most `STREAM_A11Y_MAX_CELLS` cells. */
  const plan = (): Request[] => {
    if (!window) return []
    const { sheetId, viewGen } = window
    const cache = cacheFor(sheetId, viewGen, true)!
    cache.failed.clear()
    const out: Request[] = []
    for (const [a, b] of window.cols) {
      // Rows grouped by their uncached column span: a move asks only what it uncovered.
      const groups = new Map<string, { lo: number; hi: number; rows: number[] }>()
      for (const [r0, r1] of window.rows) for (let r = r0; r <= r1; r++) {
        let lo = -1, hi = -1
        for (let c = a; c <= b; c++) {
          const k = keyOf(r, c)
          if (cache.cells.has(k) || inFlight(sheetId, viewGen, k)) continue
          if (lo < 0) lo = c
          hi = c
        }
        if (lo < 0) continue
        const id = `${lo}:${hi}`
        let group = groups.get(id)
        if (!group) groups.set(id, group = { lo, hi, rows: [] })
        group.rows.push(r)
      }
      for (const { lo, hi, rows } of groups.values()) {
        for (let c0 = lo; c0 <= hi; c0 += STREAM_A11Y_MAX_CELLS) {
          const c1 = Math.min(hi, c0 + STREAM_A11Y_MAX_CELLS - 1)
          const per = Math.max(1, Math.floor(STREAM_A11Y_MAX_CELLS / (c1 - c0 + 1)))
          for (let i = 0; i < rows.length; i += per) out.push(request(sheetId, viewGen, rows.slice(i, i + per), [c0, c1]))
        }
      }
    }
    return out
  }

  const schedulePump = (ms: number) => {
    const token = ++pumpToken
    later(() => { if (token === pumpToken) pump() }, ms)
  }
  const pump = () => {
    if (disposed || halted || inflight || !queue.length) return
    const t = now()
    while (sent.length && sent[0]! <= t - STREAM_A11Y_RATE_WINDOW_MS) sent.shift()
    if (t < holdUntil) { schedulePump(holdUntil - t); return }
    if (sent.length >= STREAM_A11Y_RATE_REQUESTS) { schedulePump(sent[0]! + STREAM_A11Y_RATE_WINDOW_MS - t); return }
    send(queue.shift()!)
  }
  const file = (req: Request, cells: ReadonlyArray<{ row: number; col: number; text: string }>) => {
    const cache = cacheFor(req.sheetId, req.viewGen, false)
    if (!cache) return
    for (const k of req.keys) cache.cells.set(k, '')
    for (const c of cells) {
      const k = keyOf(c.row, c.col)
      if (req.keys.has(k)) cache.cells.set(k, c.text)
    }
    if (mode === 'unit' && cache.cells.size > MAX_CACHED_CELLS) {
      for (const k of cache.cells.keys()) { if (cache.cells.size <= MAX_CACHED_CELLS) break; cache.cells.delete(k) }
    }
  }
  const fail = (req: Request) => {
    const cache = cacheFor(req.sheetId, req.viewGen, false)
    if (cache) for (const k of req.keys) cache.failed.add(k)
  }
  const send = (req: Request) => {
    inflight = req
    sent.push(now())
    let retry = false
    void (async () => {
      try {
        const cells = await a11y(req.sheetId, req.viewGen, req.rows, req.cols)
        if (!disposed) file(req, cells)
      } catch (error) {
        if (disposed) return
        const code = (error as { code?: unknown })?.code
        if (code === 'rate-limited' && !req.retried) { retry = true; req.retried = true; holdUntil = now() + STREAM_A11Y_RETRY_MS }
        else if (code === 'text-unavailable' || code === 'copy-quota-exhausted' || code === 'closed') { halted = true; queue = [] }
        else if (code === 'stale-view') { queue = queue.filter(r => r.sheetId !== req.sheetId || r.viewGen !== req.viewGen); fail(req) }
        else fail(req)
      } finally {
        if (inflight === req) inflight = null
        if (!disposed) {
          // A retry goes again only while its cells are still wanted (the window or the focus that asked them).
          if (retry && wanted(req)) queue.unshift(req)
          else if (retry) fail(req)
          notify()
          pump()
        }
      }
    })()
  }
  const wanted = (req: Request): boolean => mode === 'unit'
    ? !!window && window.sheetId === req.sheetId && window.viewGen === req.viewGen && !resting
    : !!focused && focused.sheetId === req.sheetId && focused.viewGen === req.viewGen && req.keys.has(focused.key)

  return {
    textAt(sheetId, viewGen, at) {
      return caches.get(sheetId)?.viewGen === viewGen ? caches.get(sheetId)!.cells.get(keyOf(at.row, at.col)) : undefined
    },
    pending(sheetId, viewGen, at) {
      if (mode === 'off' || disposed || halted) return false
      const k = keyOf(at.row, at.col)
      const cache = caches.get(sheetId)
      if (cache && cache.viewGen === viewGen && (cache.cells.has(k) || cache.failed.has(k))) return false
      if (cache && cache.viewGen > viewGen) return false
      if (inFlight(sheetId, viewGen, k) || queued(sheetId, viewGen, k)) return true
      if (mode === 'focused') return !!focused && focused.sheetId === sheetId && focused.viewGen === viewGen && focused.key === k
      return resting && !!window && window.sheetId === sheetId && window.viewGen === viewGen
        && window.rows.some(([a, b]) => at.row >= a && at.row <= b) && window.cols.some(([a, b]) => at.col >= a && at.col <= b)
    },
    showWindow(sheetId, viewGen, rows, cols) {
      if (mode !== 'unit' || disposed || halted) return
      const runs: Array<readonly [number, number]> = typeof cols[0] === 'number' ? [cols as [number, number]] : [...cols as ReadonlyArray<readonly [number, number]>]
      const sig = JSON.stringify([sheetId, viewGen, rows, runs])
      if (window?.sig === sig) return
      cacheFor(sheetId, viewGen, true)
      window = { sheetId, viewGen, rows: rows.map(([a, b]) => [a, b]), cols: runs, sig }
      queue = []
      resting = true
      const token = ++restToken
      later(() => {
        if (disposed || token !== restToken) return
        resting = false
        queue = plan()
        notify()
        pump()
      }, STREAM_A11Y_UNIT_REST_MS)
      notify()
    },
    focus(sheetId, viewGen, at) {
      if (mode !== 'focused' || disposed || halted) return
      const key = keyOf(at.row, at.col)
      if (focused && focused.sheetId === sheetId && focused.viewGen === viewGen && focused.key === key) return
      focused = { sheetId, viewGen, key }
      const cache = cacheFor(sheetId, viewGen, true)
      const token = ++restToken
      queue = []
      if (!cache || cache.cells.has(key)) { notify(); return }
      cache.failed.delete(key)
      later(() => {
        if (disposed || halted || token !== restToken) return
        const current = cacheFor(sheetId, viewGen, false)
        if (!current || current.cells.has(key) || inFlight(sheetId, viewGen, key)) return
        queue = [request(sheetId, viewGen, [at.row], [at.col, at.col])]
        pump()
        notify()
      }, STREAM_A11Y_FOCUSED_REST_MS)
      notify()
    },
    onChange(cb) {
      listeners.add(cb)
      return () => { listeners.delete(cb) }
    },
    dispose() {
      disposed = true
      listeners.clear()
      queue = []
      caches.clear()
    },
  }
}
