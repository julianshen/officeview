/**
 * Stage 4b (P10): the `sheet-a11y` turn on the main thread. It is engine-free, as `copy-turn.ts` is: `ws.ts` imports it
 * at runtime, while the text itself (`a11yAnswer`, local `formatCellPresentation`) runs in the pool worker through
 * `sheet/text-ops.ts`, which the main thread imports type-only.
 *
 * Requests queue on the sid's Stage 3 `TextSlot.copyQueue` and share its `copiedParagraphs` counter with copy, so a
 * focused read and a copy draw on one session quota, measured against the caps in force when the turn starts.
 */
import type { A11yMode, CopyPolicy } from '../hooks'
import { AdmissionRefused } from '../pool'
import { type CopyRefusal, TEXT_MESSAGE_BYTES, type TextSlot } from '../text/ops'
import type { SourceRectRecord } from './copy-turn'
import type { SheetTextReply } from './text-ops'
import type { WireSizes } from './view-state'

export type SheetA11yRefusal = CopyRefusal | 'stale-view'
/** A served cell's text is a bounded label (`boundedSelectionText`: at most 256 UTF-16 units). */
const MAX_CELL_TEXT_UNITS = 256

export interface SheetA11yContext {
  slot: TextSlot
  reqId: number
  /** The validated request (`a11yArgs`). */
  sheetIndex: number
  sheetId: string
  viewGen: number
  rows: Array<[number, number]>
  cols: [number, number]
  /** The sheet's current viewGen and its sizes, read together (one tick). */
  view: () => { viewGen: number; sizes: WireSizes }
  /** The sid's a11y mode and copy caps in force now (read when the turn starts). */
  mode: () => A11yMode
  caps: () => CopyPolicy
  /** The pool's `sheet-text` a11y job at `sizes`. */
  job: (sizes: WireSizes) => Promise<SheetTextReply>
  /** Durable audit of one served request; throws ⇒ policy-refused, nothing sent. */
  record: (r: { paragraphs: number; fromUnit: number; toUnit: number; sheetId: string; rect: SourceRectRecord }) => void
  alive: () => boolean
  send: (msg: object) => void
  refuse: (code: SheetA11yRefusal, nums?: { limit?: number; requested?: number }) => void
}

export interface SheetA11yResultMessage { reqId: number; type: 'sheet-a11y-result'; seq: number; cells: Array<[number, number, string]>; last: boolean }

const bytesOf = (v: unknown) => Buffer.byteLength(JSON.stringify(v), 'utf8')

/** `cells` as `sheet-a11y-result` chunks of at most `limit` UTF-8 bytes each, always at least one (the last); undefined
 * when one entry cannot fit a message. */
export function chunkA11yCells(reqId: number, cells: ReadonlyArray<[number, number, string]>, limit = TEXT_MESSAGE_BYTES): SheetA11yResultMessage[] | undefined {
  const baseOf = (seq: number) => bytesOf({ reqId, type: 'sheet-a11y-result', seq, cells: [], last: false })
  const out: SheetA11yResultMessage[] = []
  let cur: Array<[number, number, string]> = [], base = baseOf(0), size = base
  for (const c of cells) {
    const n = bytesOf(c)
    if (cur.length && size + 1 + n > limit) {
      out.push({ reqId, type: 'sheet-a11y-result', seq: out.length, cells: cur, last: false })
      cur = []; base = baseOf(out.length); size = base
    }
    if (base + n > limit) return undefined
    size += (cur.length ? 1 : 0) + n
    cur.push(c)
  }
  out.push({ reqId, type: 'sheet-a11y-result', seq: out.length, cells: cur, last: true })
  return out
}

const inWindow = (rows: ReadonlyArray<[number, number]>, cols: [number, number], row: number, col: number) =>
  col >= cols[0] && col <= cols[1] && rows.some(([r0, r1]) => row >= r0 && row <= r1)

/** The worker's cells rebuilt field by field (fail closed): integer positions inside the requested window, strictly
 * ascending (row, then column), each with non-empty text of at most 256 UTF-16 units; undefined otherwise. */
function rebuild(raw: unknown, rows: ReadonlyArray<[number, number]>, cols: [number, number]): Array<[number, number, string]> | undefined {
  if (!Array.isArray(raw)) return undefined
  const out: Array<[number, number, string]> = []
  let prev: [number, number] | undefined
  for (const e of raw as unknown[]) {
    if (!Array.isArray(e) || e.length !== 3) return undefined
    const [row, col, text] = e as unknown[]
    if (!Number.isSafeInteger(row) || !Number.isSafeInteger(col) || typeof text !== 'string') return undefined
    const r = row as number, c = col as number
    if (!inWindow(rows, cols, r, c) || text === '' || text.length > MAX_CELL_TEXT_UNITS) return undefined
    if (prev && !(r > prev[0] || (r === prev[0] && c > prev[1]))) return undefined
    prev = [r, c]
    out.push([r, c, text])
  }
  return out
}

/** Queues one validated sheet a11y request behind the sid's earlier copies and a11y reads; resolves when it has been
 * answered (or dropped because its owner is gone). Never rejects. */
export function runSheetA11y(c: SheetA11yContext): Promise<void> {
  const turn = (c.slot.copyQueue ?? Promise.resolve()).then(() => sheetA11yTurn(c))
  c.slot.copyQueue = turn
  void turn.then(() => { if (c.slot.copyQueue === turn) c.slot.copyQueue = undefined })
  return turn
}

/**
 * One sheet a11y request, in `a11yTurn`'s order with P10's modes: mode (off ⇒ text-unavailable) → a `focused` request
 * must be 1×1 (else policy-refused) and needs copy caps (none ⇒ text-unavailable) → `stale-view` when the sheet's
 * viewGen is no longer the request's (its sizes are read in the same tick; ruling F8: before any count or audit) → the
 * worker job → the cells rebuilt and checked → every entry fits one message (else copy-too-large) → `focused` with text:
 * session quota (`used + 1`, else copy-quota-exhausted {limit, requested: 1}) → audit (one row per served request with
 * text: `unit` the served cells, uncounted; `focused` 1) → counter (`focused` only) → `sheet-a11y-result` chunks.
 * Nothing awaits after the job returns.
 */
async function sheetA11yTurn(c: SheetA11yContext): Promise<void> {
  let answered = false
  const refuse = (code: SheetA11yRefusal, nums?: { limit?: number; requested?: number }) => { answered = true; c.refuse(code, nums) }
  try {
    if (!c.alive()) return
    const mode = c.mode()
    if (mode === 'off') return refuse('text-unavailable')
    const focused = mode === 'focused'
    if (focused && !(c.rows.length === 1 && c.rows[0]![0] === c.rows[0]![1] && c.cols[0] === c.cols[1])) return refuse('policy-refused')
    const caps = c.caps()
    if (focused && !caps) return refuse('text-unavailable')
    const view = c.view()
    if (view.viewGen !== c.viewGen) return refuse('stale-view')
    let reply: SheetTextReply
    try { reply = await c.job(view.sizes) } catch (e) {
      if (!c.alive()) return
      return refuse(e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused')
    }
    if (!c.alive()) return
    if (!reply || reply.op !== 'a11y') return refuse('policy-refused')
    const cells = rebuild(reply.cells, c.rows, c.cols)
    if (!cells || (focused && cells.length > 1)) return refuse('policy-refused')
    const messages = chunkA11yCells(c.reqId, cells)
    if (!messages) return refuse('copy-too-large', { limit: TEXT_MESSAGE_BYTES })
    if (focused && cells.length) {
      const cap = caps ? caps.maxSessionParagraphs : null
      if (cap !== null && !(c.slot.copiedParagraphs + 1 <= cap)) return refuse('copy-quota-exhausted', { limit: cap, requested: 1 })
    }
    if (cells.length) {
      const rect = { r0: c.rows[0]![0], c0: c.cols[0], r1: c.rows[c.rows.length - 1]![1], c1: c.cols[1] }
      try { c.record({ paragraphs: focused ? 1 : cells.length, fromUnit: c.sheetIndex, toUnit: c.sheetIndex, sheetId: c.sheetId, rect }) } catch { return refuse('policy-refused') }
      if (focused) c.slot.copiedParagraphs += 1
    }
    answered = true
    for (const m of messages) c.send(m)
  } catch {
    // A failure inside the turn never breaks the sid's queue.
    if (!answered) try { if (c.alive()) c.refuse('policy-refused') } catch { /* the socket is gone */ }
  }
}
