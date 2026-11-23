/**
 * Stage 4b (P8): the `sheet-copy` turn on the main thread. It is engine-free: `ws.ts` imports it at runtime, while the
 * copy itself (`copyAnswer`, local `serializeSelection`) runs in the pool worker through `sheet/text-ops.ts`, which the
 * main thread imports type-only.
 *
 * Copies queue on the sid's Stage 3 `TextSlot.copyQueue` and share its `copiedParagraphs` counter (with focused a11y),
 * so the session quota survives a resume and is measured against the caps in force when the turn starts.
 */
import type { WireSelection } from '../../../src/stream/sheet-wire'
import type { CopyPolicy } from '../hooks'
import { AdmissionRefused } from '../pool'
import { chunkText, COPY_MAX_BYTES, type CopyRefusal, type TextSlot } from '../text/ops'
import type { SheetTextReply } from './text-ops'
import type { WireSizes } from './view-state'

export type SheetCopyRefusal = CopyRefusal | 'stale-view'
export interface SourceRectRecord { r0: number; c0: number; r1: number; c1: number }

export interface SheetCopyContext {
  slot: TextSlot
  reqId: number
  /** The validated request (`copyArgs`). */
  sheetIndex: number
  sheetId: string
  viewGen: number
  selection: WireSelection
  /** The sheet's current viewGen and its sizes, read together (one tick). */
  view: () => { viewGen: number; sizes: WireSizes }
  /** The sid's copy caps in force now (read when the turn starts). */
  caps: () => CopyPolicy
  /** The pool's `sheet-text` copy job at `sizes`, with `maxBytes` UTF-8 bytes. */
  job: (sizes: WireSizes, maxBytes: number) => Promise<SheetTextReply>
  /** Durable audit of one counted copy; throws ⇒ policy-refused, nothing sent. */
  record: (r: { paragraphs: number; fromUnit: number; toUnit: number; sheetId: string; rect?: SourceRectRecord }) => void
  alive: () => boolean
  send: (msg: object) => void
  refuse: (code: SheetCopyRefusal, nums?: { limit?: number; requested?: number }) => void
}

/** The bounding source rect of `s`'s ranges (undefined for no ranges). */
export function selectionRect(s: WireSelection): SourceRectRecord | undefined {
  if (!s.ranges.length) return undefined
  let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1
  for (const r of s.ranges) {
    r0 = Math.min(r0, r.rows[0]![0]); r1 = Math.max(r1, r.rows[r.rows.length - 1]![1])
    c0 = Math.min(c0, r.cols[0]![0]); c1 = Math.max(c1, r.cols[r.cols.length - 1]![1])
  }
  return { r0, c0, r1, c1 }
}

/** Queues one validated sheet copy behind the sid's earlier copies (and a11y reads); resolves when it has been
 * answered (or dropped because its owner is gone). Never rejects. */
export function runSheetCopy(c: SheetCopyContext): Promise<void> {
  const turn = (c.slot.copyQueue ?? Promise.resolve()).then(() => sheetCopyTurn(c))
  c.slot.copyQueue = turn
  void turn.then(() => { if (c.slot.copyQueue === turn) c.slot.copyQueue = undefined })
  return turn
}

/**
 * One sheet copy, in `copyTurn`'s order (P8): caps (off ⇒ text-unavailable) → `stale-view` when the sheet's viewGen is
 * no longer the request's (its sizes are read in the same tick) → resolve and serialize in the worker (the 4 MiB UTF-8
 * cap: copy-too-large {limit: 4194304, requested?}) → count (n = the non-empty fields) → per-copy cap
 * (`n > maxParagraphs` ⇒ copy-too-large) → session cap (`used + n > maxSessionParagraphs` ⇒ copy-quota-exhausted) →
 * audit (n > 0 only, as Stage 3; failure ⇒ policy-refused) → counter → `copy-result` chunks. Never truncated. Nothing
 * awaits after the job returns, so the checks, record, count and send are one step. A copy of only blank or withheld
 * cells (n = 0) still sends its TSV shape (tabs and newlines), uncounted and unaudited.
 */
async function sheetCopyTurn(c: SheetCopyContext): Promise<void> {
  let answered = false
  const refuse = (code: SheetCopyRefusal, nums?: { limit?: number; requested?: number }) => { answered = true; c.refuse(code, nums) }
  try {
    if (!c.alive()) return
    const caps = c.caps()
    if (!caps) return refuse('text-unavailable')
    const view = c.view()
    if (view.viewGen !== c.viewGen) return refuse('stale-view')
    let reply: SheetTextReply
    try { reply = await c.job(view.sizes, COPY_MAX_BYTES) } catch (e) {
      if (!c.alive()) return
      return refuse(e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused')
    }
    if (!c.alive()) return
    if (!reply || reply.op !== 'copy') return refuse('policy-refused')
    if ('refused' in reply) {
      const bytes = reply.bytes
      return refuse('copy-too-large', { limit: COPY_MAX_BYTES, ...(Number.isSafeInteger(bytes) && bytes > COPY_MAX_BYTES ? { requested: bytes } : {}) })
    }
    const { text, paragraphs: n } = reply
    if (typeof text !== 'string' || !Number.isSafeInteger(n) || n < 0) return refuse('policy-refused')
    // Re-measured here (fail closed): never more than 4 MiB of UTF-8 travels.
    const bytes = Buffer.byteLength(text, 'utf8')
    if (!(bytes <= COPY_MAX_BYTES)) return refuse('copy-too-large', { limit: COPY_MAX_BYTES, requested: bytes })
    if (caps.maxParagraphs !== null && !(n <= caps.maxParagraphs)) return refuse('copy-too-large', { limit: caps.maxParagraphs, requested: n })
    const cap = caps.maxSessionParagraphs
    if (cap !== null && !(c.slot.copiedParagraphs + n <= cap)) return refuse('copy-quota-exhausted', { limit: cap, requested: n })
    if (n > 0) {
      const rect = selectionRect(c.selection)
      try { c.record({ paragraphs: n, fromUnit: c.sheetIndex, toUnit: c.sheetIndex, sheetId: c.sheetId, ...(rect ? { rect } : {}) }) } catch { return refuse('policy-refused') }
      c.slot.copiedParagraphs += n
    }
    answered = true
    for (const m of chunkText(c.reqId, text)) c.send(m)
  } catch {
    // A failure inside the turn never breaks the sid's queue.
    if (!answered) try { if (c.alive()) c.refuse('policy-refused') } catch { /* the socket is gone */ }
  }
}
