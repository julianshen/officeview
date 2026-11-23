/**
 * Stage 4b (S4 §6.7; P6, P7): the client side of stream text that needs the viewer's types — the selection's wire form
 * and the stream find session. There is no local text index on a stream (parent §8.3): hits are server positions, and
 * a hit carries no text.
 *
 * Reachable only from the lazy stream viewer (`StreamWorkbook`), never from the package index.
 */
import { SELECTION_LIMITS, encodeWireSelection, type SheetFindAnswer, type SheetFindHit, type WireSelection } from '../../stream/sheet-wire'
import type { RemoteSheetOps } from '../../stream/remote/remote-doc'
import { rowSetIntervals } from '../../worker/viewerSession'
import type { GridSelection, SearchQuery, SourceColInterval, WorkbookOwner } from '../../xlsx/view/contracts'
import { cellId } from '../../xlsx/view/coordinates'
import { FIND_PAGE_HITS, FindError, type FindHit, type FindPage, type FindSession } from '../../xlsx/view/find'

/** A selection's encoded form must leave room for the request envelope inside the 64 KiB request cap (P7). */
const SELECTION_BYTES = 64 * 1024 - 1024
const utf8 = new TextEncoder()

/** Ascending, disjoint, merged column intervals (a part's may arrive in gesture order). */
function colIntervals(cols: readonly SourceColInterval[]): Array<[number, number]> {
  const sorted = [...cols].sort((a, b) => a.c0 - b.c0)
  const out: Array<[number, number]> = []
  for (const c of sorted) {
    if (!(c.c0 <= c.c1)) throw new RangeError('Invalid column interval')
    const last = out[out.length - 1]
    if (last && c.c0 <= last[1] + 1) last[1] = Math.max(last[1], c.c1)
    else out.push([c.c0, c.c1])
  }
  return out
}

/**
 * The selection as it travels (P7): each part's rows through `rowSetIntervals`, its columns merged. `'too-large'` (the
 * client refuses locally, 'Selection too large to copy.') when a part's rows are a worker `ref` (it cannot occur on a
 * stream), the set is malformed, more than 256 ranges remain, or the encoded selection would not fit the request cap.
 * A part with no rows or no columns selects nothing and is dropped (the server refuses an empty interval list).
 */
export function wireSelection(s: GridSelection): WireSelection | 'too-large' {
  const ranges: WireSelection['ranges'] = []
  try {
    for (const part of s.ranges) {
      const rows = rowSetIntervals(part.rows)
      if (rows === null) return 'too-large'
      const cols = colIntervals(part.cols)
      if (!rows.length || !cols.length) continue
      if (ranges.length >= SELECTION_LIMITS.ranges) return 'too-large'
      ranges.push({ rows: rows.map((r): [number, number] => [r.r0, r.r1]), cols })
    }
  } catch { return 'too-large' }
  const out: WireSelection = {
    active: s.active ? [s.active.row, s.active.col] : null,
    anchor: s.anchor ? [s.anchor.row, s.anchor.col] : null,
    ranges,
  }
  return utf8.encode(JSON.stringify(encodeWireSelection(out))).length > SELECTION_BYTES ? 'too-large' : out
}

/** A stream find hit: no text (`text: ''`, `start = end = 0`); `hidden: 'filter'` when the filter hides its row. */
export type StreamFindHit = FindHit & { hidden?: 'filter' }
/** F15: every page carries the answer's `truncated` (more than 5,000 matches existed), so the viewer can say so. */
export interface StreamFindPage extends FindPage { hits: readonly StreamFindHit[]; truncated: boolean }
export interface StreamFindSession extends FindSession { nextPage(): Promise<StreamFindPage> }

/**
 * One workbook find over the session's `sheet-find`, paged at `FIND_PAGE_HITS` in the server's (`WorkbookFind`'s)
 * order. The find is sent on the first `nextPage`; `close()` aborts it. `sheetIds` maps the wire sheet index to the
 * sheet id; a hit on a sheet outside `query.sheetIds` (when given) is dropped. `displaySlot` is the hit's display row,
 * which is its source row under 4b's identity order. `totalKnown` is set only when the answer is not truncated.
 */
export function streamFindSession(view: Pick<RemoteSheetOps, 'find'>, owner: WorkbookOwner, sheetIds: readonly string[], query: SearchQuery, signal?: AbortSignal): StreamFindSession {
  const abort = new AbortController()
  const onAbort = () => abort.abort(signal?.reason)
  signal?.addEventListener('abort', onAbort, { once: true })
  let answer: Promise<{ hits: StreamFindHit[]; truncated: boolean }> | undefined
  let page = 0, closed = false
  const scope = query.sheetIds ? new Set(query.sheetIds) : undefined
  const toHit = (h: SheetFindHit): StreamFindHit | undefined => {
    const sheetId = sheetIds[h.sheetIndex]
    if (sheetId === undefined || (scope && !scope.has(sheetId))) return undefined
    return { sheetId, cellId: cellId(h.row, h.col), source: { row: h.row, col: h.col }, text: '', start: 0, end: 0, displaySlot: h.row, ...(h.hidden ? { hidden: h.hidden } : {}) }
  }
  const load = () => answer ??= view.find({ query: query.text, matchCase: query.caseSensitive }, abort.signal).then((a: SheetFindAnswer) => {
    const hits: StreamFindHit[] = []
    for (const h of a.hits) { const hit = toHit(h); if (hit) hits.push(hit) }
    return { hits, truncated: a.truncated }
  })
  return {
    async nextPage() {
      if (closed) throw new FindError('closed', 'Find session closed')
      if (abort.signal.aborted) throw new FindError('cancelled', 'Find cancelled')
      let a: Awaited<ReturnType<typeof load>>
      try { a = await load() } catch (error) {
        // As the local session: closed, aborted or superseded (a newer find, `closed`) end the run quietly; any other
        // refusal (rate-limited, policy-refused, …) is an outcome the viewer shows.
        if (closed || (error as { code?: unknown })?.code === 'closed') throw new FindError('closed', 'Find session closed')
        if (abort.signal.aborted) throw new FindError('cancelled', 'Find cancelled')
        throw error
      }
      if (closed) throw new FindError('closed', 'Find session closed')
      const from = page * FIND_PAGE_HITS, hits = a.hits.slice(from, from + FIND_PAGE_HITS)
      page++
      const complete = from + FIND_PAGE_HITS >= a.hits.length
      return {
        owner, hits, complete, truncated: a.truncated, unavailableCount: 0, incomplete: [], failures: [],
        ...(complete && !a.truncated ? { totalKnown: a.hits.length } : {}),
      }
    },
    close() {
      if (closed) return
      closed = true
      signal?.removeEventListener('abort', onAbort)
      abort.abort()
    },
  }
}
