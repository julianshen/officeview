/**
 * Stage 3a — main-thread text operations over the wire (Task 4: search;
 * Task 5: select and copy, with the copy caps and the session quota; Task 8:
 * a11y text, sharing copy's queue and session counter).
 *
 * `ws.ts` authenticates and dispatches; this module owns argument validation
 * (R6), the batched unit walk, the ≤ 48 KiB message chunking and per-sid
 * supersession. It takes callbacks, never a socket, so it is unit-testable.
 */
import type { CopyPolicy } from '../hooks'
import { AdmissionRefused, TextTooLarge } from '../pool'
import type { TextJob, TextRange, TextReply, UnitPoint } from './served'

/** R6: a query is 1..256 UTF-16 units. */
export const MAX_QUERY_UNITS = 256
/** R6: at most this many matches are sent for one search (`truncated: true` beyond). */
export const MAX_SEARCH_MATCHES = 5000
/** Units per `Pool.text` search job. */
export const SEARCH_BATCH_UNITS = 8
/** Every server→client text message is ≤ 48 KiB of UTF-8. */
export const TEXT_MESSAGE_BYTES = 48 * 1024
/** R6: one copy is at most this many bytes of UTF-8 (`copy-too-large` beyond, even under 'full'). */
export const COPY_MAX_BYTES = 4 * 1024 * 1024

export type SearchJob = Extract<TextJob, { op: 'search' }>
export interface SearchMatch { unit: number; rects: number[] }
export interface SearchResultMessage { reqId: number; type: 'search-result'; matches: SearchMatch[] }
export type TextRefusal = 'rate-limited' | 'policy-refused'

const enc = new TextEncoder()
const bytesOf = (v: unknown) => enc.encode(JSON.stringify(v)).length

/** The search arguments of a client message, or undefined (⇒ policy-refused). */
export function searchArgs(msg: Record<string, unknown>): { query: string; caseSensitive: boolean } | undefined {
  const { query, caseSensitive } = msg
  if (typeof query !== 'string' || query.length === 0 || query.length > MAX_QUERY_UNITS) return undefined
  if (caseSensitive !== undefined && typeof caseSensitive !== 'boolean') return undefined
  return { query, caseSensitive: caseSensitive === true }
}

/** `search-result` messages for `matches`, in order, each ≤ `limit` bytes of
 * UTF-8 when JSON-encoded. One wire entry per match, always: a match whose
 * entry alone cannot fit one message is omitted (counted in `omitted`). */
export function chunkMatches(reqId: number, matches: readonly SearchMatch[], limit = TEXT_MESSAGE_BYTES): { messages: SearchResultMessage[]; omitted: number } {
  const base = bytesOf({ reqId, type: 'search-result', matches: [] })
  const messages: SearchResultMessage[] = []
  let cur: SearchMatch[] = [], size = base, omitted = 0
  for (const raw of matches) {
    const m = { unit: raw.unit, rects: raw.rects }
    const n = bytesOf(m)
    if (base + n > limit) { omitted++; continue }
    if (cur.length && size + 1 + n > limit) { messages.push({ reqId, type: 'search-result', matches: cur }); cur = []; size = base }
    size += (cur.length ? 1 : 0) + n
    cur.push(m)
  }
  if (cur.length) messages.push({ reqId, type: 'search-result', matches: cur })
  return { messages, omitted }
}

/** The in-flight search of one sid. `send`/`alive` are bound to the socket
 * that issued it, so a superseded run on a replaced socket sends nothing. */
export interface SearchRun {
  readonly reqId: number
  /** Matches sent so far. */
  sent: number
  /** Finished, superseded, refused or torn down: the run sends nothing more. */
  stopped: boolean
  /** Resolves once the run's last pool job has returned. */
  readonly settled: Promise<void>
  readonly send: (msg: object) => void
  readonly alive: () => boolean
}
/** Where a sid keeps its in-flight search (the per-sid session entry). */
export interface SearchSlot { search?: SearchRun }

export interface SearchContext {
  slot: SearchSlot
  reqId: number
  query: string
  caseSensitive: boolean
  unitCount: number
  /** One worker job over `fromUnit..toUnit` (Pool.text). */
  batch: (job: SearchJob) => Promise<TextReply>
  /** The issuing socket still owns a live session in `stream` (re-checked after every await). */
  alive: () => boolean
  send: (msg: object) => void
  refuse: (code: TextRefusal) => void
}

/** Stops the sid's in-flight search, answering its reqId with what was sent
 * so far and `truncated: true` (R6 supersession). */
export function supersede(slot: SearchSlot): SearchRun | undefined {
  const prev = slot.search
  if (!prev || prev.stopped) return prev
  prev.stopped = true
  if (prev.alive()) prev.send({ reqId: prev.reqId, type: 'search-done', total: prev.sent, truncated: true })
  return prev
}

/**
 * Runs one validated search: supersedes the sid's running one, then walks
 * the units in order, SEARCH_BATCH_UNITS per job, sending each batch's
 * matches as `search-result` messages and finally `search-done`. Each batch
 * asks for one match more than may still be sent, so truncation is exact; a
 * match too large for one message is omitted and marks the answer truncated. A
 * pool failure is a refusal (AdmissionRefused ⇒ rate-limited, anything else
 * ⇒ policy-refused) and no `search-done` follows it.
 */
export async function runSearch(c: SearchContext): Promise<void> {
  const prev = supersede(c.slot)
  let settle!: () => void
  const run: SearchRun = { reqId: c.reqId, sent: 0, stopped: false, settled: new Promise<void>((r) => { settle = r }), send: c.send, alive: c.alive }
  c.slot.search = run
  const current = () => !run.stopped && c.slot.search === run && c.alive()
  /** A match omitted for size makes the answer incomplete. */
  let omitted = false
  const done = (truncated: boolean) => { run.stopped = true; c.send({ reqId: c.reqId, type: 'search-done', total: run.sent, truncated: truncated || omitted }) }
  try {
    // One search in flight per sid: wait for the superseded run's pool job.
    if (prev) await prev.settled
    for (let from = 0; from < c.unitCount; from += SEARCH_BATCH_UNITS) {
      if (!current()) return
      const remaining = MAX_SEARCH_MATCHES - run.sent
      let reply: TextReply
      try {
        reply = await c.batch({ op: 'search', query: c.query, caseSensitive: c.caseSensitive, fromUnit: from, toUnit: Math.min(from + SEARCH_BATCH_UNITS, c.unitCount) - 1, maxMatches: remaining + 1 })
      } catch (e) {
        if (!current()) return
        run.stopped = true
        return c.refuse(e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused')
      }
      if (!current()) return
      if (!('op' in reply) || reply.op !== 'search') { run.stopped = true; return c.refuse('policy-refused') }
      const send = reply.matches.slice(0, remaining)
      const chunks = chunkMatches(c.reqId, send)
      for (const m of chunks.messages) c.send(m)
      // `total` counts the entries actually sent; an omitted match frees its budget slot.
      run.sent += send.length - chunks.omitted
      if (chunks.omitted) omitted = true
      if (reply.matches.length > remaining) return done(true)
    }
    if (current()) done(false)
  } finally {
    run.stopped = true
    if (c.slot.search === run) c.slot.search = undefined
    settle()
  }
}

// ── Task 5: select and copy ────────────────────────────────────────────────

export type SelectJob = Extract<TextJob, { op: 'select' }>
export type CopyJob = Extract<TextJob, { op: 'copy' }>
export type CopyRefusal = TextRefusal | 'text-unavailable' | 'copy-too-large' | 'copy-quota-exhausted'
export interface SelectUnit { unit: number; rects: number[] }
export interface SelectResultMessage { reqId: number; type: 'select-result'; units: SelectUnit[]; paragraphs: number; last?: true }
export interface CopyResultMessage { reqId: number; type: 'copy-result'; seq: number; text: string; last: boolean }

/** The in-flight select of one sid (a newer one supersedes it). */
export interface SelectRun { readonly reqId: number; stopped: boolean }
/** Per-sid text state on the session entry: survives resume, gone on teardown. */
export interface TextSlot extends SearchSlot {
  select?: SelectRun
  /** Paragraphs copied so far by this sid (the session quota's counter). */
  copiedParagraphs: number
  /** Tail of the sid's copy queue: copies run one at a time. */
  copyQueue?: Promise<void>
}

const isPoint = (p: unknown, unitCount: number): p is UnitPoint => {
  if (!p || typeof p !== 'object' || Array.isArray(p)) return false
  const { unit, x, y } = p as Record<string, unknown>
  return Number.isInteger(unit) && (unit as number) >= 0 && (unit as number) < unitCount && Number.isFinite(x) && Number.isFinite(y)
}

/** The range of a client message, rebuilt from its validated fields only, or
 * undefined (⇒ policy-refused): `{ all: true }`, or anchor and focus points
 * whose `unit` is an integer in range and whose `x`/`y` are finite numbers. */
export function rangeArg(v: unknown, unitCount: number): TextRange | undefined {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return undefined
  const r = v as Record<string, unknown>
  if ('all' in r) return r.all === true ? { all: true } : undefined
  const { anchor, focus } = r
  if (!isPoint(anchor, unitCount) || !isPoint(focus, unitCount)) return undefined
  return { anchor: { unit: anchor.unit, x: anchor.x, y: anchor.y }, focus: { unit: focus.unit, x: focus.x, y: focus.y } }
}

/** `select-result` messages, each ≤ `limit` bytes of JSON UTF-8. Whole
 * `{unit, rects}` entries share a message where they fit; a unit whose entry
 * alone cannot fit is split across messages at rect (4-number) boundaries.
 * Every message carries `paragraphs`; only the last carries `last: true`; an
 * empty selection is one message with no units. */
export function chunkSelect(reqId: number, units: readonly SelectUnit[], paragraphs: number, limit = TEXT_MESSAGE_BYTES): SelectResultMessage[] {
  // Budgeted with `last: false` (longer than `last: true`): holds for every chunk.
  const base = bytesOf({ reqId, type: 'select-result', units: [], paragraphs, last: false })
  const out: SelectUnit[][] = []
  let cur: SelectUnit[] = [], size = base
  const add = (u: SelectUnit, n: number) => {
    if (cur.length && size + 1 + n > limit) { out.push(cur); cur = []; size = base }
    size += (cur.length ? 1 : 0) + n
    cur.push(u)
  }
  for (const raw of units) {
    const entry = { unit: raw.unit, rects: raw.rects }
    const n = bytesOf(entry)
    if (base + n <= limit) { add(entry, n); continue }
    // Too large alone: split this unit's rects.
    if (cur.length) { out.push(cur); cur = []; size = base }
    const empty = bytesOf({ unit: raw.unit, rects: [] })
    let rects: number[] = [], rs = empty
    for (let i = 0; i < raw.rects.length; i += 4) {
      const quad = raw.rects.slice(i, i + 4)
      const q = bytesOf(quad) - 2
      if (rects.length && base + rs + 1 + q > limit) { out.push([{ unit: raw.unit, rects }]); rects = []; rs = empty }
      rs += (rects.length ? 1 : 0) + q
      rects.push(...quad)
    }
    add({ unit: raw.unit, rects }, rs)
  }
  if (cur.length || !out.length) out.push(cur)
  return out.map((u, i) => ({ reqId, type: 'select-result', units: u, paragraphs, ...(i === out.length - 1 ? { last: true as const } : {}) }))
}

/** UTF-8 bytes of one code point inside a JSON string (as JSON.stringify writes it). */
function jsonCost(cp: number): number {
  if (cp === 0x22 || cp === 0x5c) return 2
  if (cp < 0x20) return cp === 8 || cp === 9 || cp === 10 || cp === 12 || cp === 13 ? 2 : 6
  if (cp < 0x80) return 1
  if (cp < 0x800) return 2
  if (cp >= 0xd800 && cp <= 0xdfff) return 6 // a lone surrogate is written escaped
  return cp < 0x10000 ? 3 : 4
}

/** `copy-result` messages for `text`, each ≤ `limit` bytes of JSON UTF-8,
 * split on code-point boundaries (never inside a surrogate pair), `seq` from
 * 0, `last: true` on the final one. Empty text is one empty message. */
export function chunkText(reqId: number, text: string, limit = TEXT_MESSAGE_BYTES): CopyResultMessage[] {
  const out: CopyResultMessage[] = []
  const baseOf = (seq: number) => bytesOf({ reqId, type: 'copy-result', seq, text: '', last: false })
  let start = 0, i = 0, base = baseOf(0), size = base
  for (const ch of text) {
    const n = jsonCost(ch.codePointAt(0)!)
    if (i > start && size + n > limit) {
      out.push({ reqId, type: 'copy-result', seq: out.length, text: text.slice(start, i), last: false })
      start = i; base = baseOf(out.length); size = base
    }
    size += n; i += ch.length
  }
  out.push({ reqId, type: 'copy-result', seq: out.length, text: text.slice(start), last: true })
  return out
}

export interface SelectContext {
  slot: TextSlot
  reqId: number
  range: TextRange
  job: (job: SelectJob) => Promise<TextReply>
  alive: () => boolean
  send: (msg: object) => void
  refuse: (code: TextRefusal) => void
}

/** One validated select: supersedes the sid's in-flight one (whose reqId then
 * gets no further message), runs one worker job and sends its rects chunked.
 * AdmissionRefused ⇒ rate-limited; any other failure (TextTooLarge included)
 * ⇒ policy-refused. */
export async function runSelect(c: SelectContext): Promise<void> {
  if (c.slot.select) c.slot.select.stopped = true
  const run: SelectRun = { reqId: c.reqId, stopped: false }
  c.slot.select = run
  const current = () => !run.stopped && c.slot.select === run && c.alive()
  try {
    let reply: TextReply
    try { reply = await c.job({ op: 'select', range: c.range }) } catch (e) {
      if (current()) c.refuse(e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused')
      return
    }
    if (!current()) return
    if (!('op' in reply) || reply.op !== 'select') return c.refuse('policy-refused')
    for (const m of chunkSelect(c.reqId, reply.units, reply.paragraphs)) c.send(m)
  } finally {
    run.stopped = true
    if (c.slot.select === run) c.slot.select = undefined
  }
}

export interface CopyContext {
  slot: TextSlot
  reqId: number
  range: TextRange
  unitCount: number
  /** The sid's copy caps in force now (read when the copy's turn starts). */
  caps: () => CopyPolicy
  job: (job: CopyJob) => Promise<TextReply>
  /** Durable audit of one copy; throws ⇒ refused, nothing sent. */
  record: (r: { paragraphs: number; fromUnit: number; toUnit: number }) => void
  alive: () => boolean
  send: (msg: object) => void
  refuse: (code: CopyRefusal, nums?: { limit?: number; requested?: number }) => void
}

/** Queues one validated copy behind the sid's earlier ones; resolves when it
 * has been answered (or dropped because its owner is gone). Never rejects. */
export function runCopy(c: CopyContext): Promise<void> {
  const turn = (c.slot.copyQueue ?? Promise.resolve()).then(() => copyTurn(c))
  c.slot.copyQueue = turn
  void turn.then(() => { if (c.slot.copyQueue === turn) c.slot.copyQueue = undefined })
  return turn
}

/**
 * One copy, in order: caps (off ⇒ text-unavailable) → worker job with
 * `maxParagraphs` and the 4 MiB byte cap → per-copy caps (re-checked here, fail
 * closed) → session quota (`used + n > maxSessionParagraphs` ⇒ refused) →
 * audit record (failure ⇒ policy-refused) → counter → chunks. Nothing awaits
 * after the job returns, so the checks, record, count and send are one step.
 */
async function copyTurn(c: CopyContext): Promise<void> {
  let answered = false
  const refuse = (code: CopyRefusal, nums?: { limit?: number; requested?: number }) => { answered = true; c.refuse(code, nums) }
  try {
    if (!c.alive()) return
    const caps = c.caps()
    if (!caps) return refuse('text-unavailable')
    let reply: TextReply
    try { reply = await c.job({ op: 'copy', range: c.range, maxParagraphs: caps.maxParagraphs, maxBytes: COPY_MAX_BYTES }) } catch (e) {
      if (!c.alive()) return
      if (e instanceof TextTooLarge) return refuse('copy-too-large', { limit: COPY_MAX_BYTES })
      return refuse(e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused')
    }
    if (!c.alive()) return
    if (!('op' in reply) || reply.op !== 'copy') return refuse('policy-refused')
    const n = reply.paragraphs
    const overParagraphs = caps.maxParagraphs !== null && !(n <= caps.maxParagraphs)
    if (overParagraphs) return refuse('copy-too-large', { limit: caps.maxParagraphs!, requested: n })
    if ('refused' in reply) return refuse('copy-too-large', { limit: COPY_MAX_BYTES, ...(reply.bytes > 0 ? { requested: reply.bytes } : {}) })
    if (n === 0) { answered = true; return c.send({ reqId: c.reqId, type: 'copy-result', seq: 0, text: '', last: true }) }
    const bytes = Buffer.byteLength(reply.text, 'utf8')
    if (!(bytes <= COPY_MAX_BYTES)) return refuse('copy-too-large', { limit: COPY_MAX_BYTES, requested: bytes })
    const cap = caps.maxSessionParagraphs
    if (cap !== null && !(c.slot.copiedParagraphs + n <= cap)) return refuse('copy-quota-exhausted', { limit: cap, requested: n })
    const r = c.range
    const units = 'all' in r ? { fromUnit: 0, toUnit: Math.max(0, c.unitCount - 1) } : { fromUnit: Math.min(r.anchor.unit, r.focus.unit), toUnit: Math.max(r.anchor.unit, r.focus.unit) }
    try { c.record({ paragraphs: n, ...units }) } catch { return refuse('policy-refused') }
    c.slot.copiedParagraphs += n
    answered = true
    for (const m of chunkText(c.reqId, reply.text)) c.send(m)
  } catch {
    // A failure inside the turn never breaks the sid's queue.
    if (!answered) try { if (c.alive()) c.refuse('policy-refused') } catch { /* the socket is gone */ }
  }
}

// ── Task 8: a11y text ──────────────────────────────────────────────────────

export type A11yJob = Extract<TextJob, { op: 'a11y' }>
export type A11yMode = 'off' | 'unit' | 'focused'
export interface A11yParagraph { rect: number[]; text?: string }
export interface A11yResultMessage { reqId: number; type: 'a11y-result'; seq: number; paragraphs: A11yParagraph[]; last: boolean }
/** The client's per-request a11y limits: paragraphs, and UTF-8 bytes of text (= COPY_MAX_BYTES). */
export const A11Y_MAX_PARAGRAPHS = 20_000

/** The a11y arguments of a client message, rebuilt from its validated fields
 * only, or undefined (⇒ policy-refused): `unit` an integer in range,
 * `paragraph` absent or a non-negative safe integer. */
export function a11yArgs(msg: Record<string, unknown>, unitCount: number): { unit: number; paragraph?: number } | undefined {
  const { unit, paragraph } = msg
  if (!Number.isInteger(unit) || (unit as number) < 0 || (unit as number) >= unitCount) return undefined
  if (paragraph !== undefined && !(Number.isSafeInteger(paragraph) && (paragraph as number) >= 0)) return undefined
  return { unit: unit as number, ...(paragraph !== undefined ? { paragraph: paragraph as number } : {}) }
}

/** `a11y-result` messages for `paragraphs`, in order, each ≤ `limit` bytes of
 * JSON UTF-8: `seq` from 0, `last` on every message (true on the final one
 * only), never an empty non-final message; no paragraphs is one empty final
 * message. One wire entry per paragraph: undefined when an entry alone cannot
 * fit one message (the fixed wire shape cannot split a paragraph). */
export function chunkA11y(reqId: number, paragraphs: readonly A11yParagraph[], limit = TEXT_MESSAGE_BYTES): A11yResultMessage[] | undefined {
  const baseOf = (seq: number) => bytesOf({ reqId, type: 'a11y-result', seq, paragraphs: [], last: false })
  const out: A11yResultMessage[] = []
  let cur: A11yParagraph[] = [], base = baseOf(0), size = base
  for (const p of paragraphs) {
    const n = bytesOf(p)
    if (cur.length && size + 1 + n > limit) {
      out.push({ reqId, type: 'a11y-result', seq: out.length, paragraphs: cur, last: false })
      cur = []; base = baseOf(out.length); size = base
    }
    if (base + n > limit) return undefined
    size += (cur.length ? 1 : 0) + n
    cur.push(p)
  }
  out.push({ reqId, type: 'a11y-result', seq: out.length, paragraphs: cur, last: true })
  return out
}

export interface A11yContext {
  slot: TextSlot
  reqId: number
  unit: number
  paragraph?: number
  /** The sid's a11y mode and copy caps in force now (read when the turn starts). */
  mode: () => A11yMode
  caps: () => CopyPolicy
  job: (job: A11yJob) => Promise<TextReply>
  /** Durable audit of one answered a11y request; throws ⇒ refused, nothing sent. */
  record: (r: { paragraphs: number }) => void
  alive: () => boolean
  send: (msg: object) => void
  refuse: (code: CopyRefusal, nums?: { limit?: number; requested?: number }) => void
}

/** Queues one validated a11y request on the sid's copy queue (the session
 * quota's counter is shared with copy); resolves when answered. Never rejects. */
export function runA11y(c: A11yContext): Promise<void> {
  const turn = (c.slot.copyQueue ?? Promise.resolve()).then(() => a11yTurn(c))
  c.slot.copyQueue = turn
  void turn.then(() => { if (c.slot.copyQueue === turn) c.slot.copyQueue = undefined })
  return turn
}

/** A paragraph's box as exactly four numbers (a paragraph with no rects gets a
 * zero box: dropping it would shift the indices `paragraph` names). */
const box = (r: unknown): number[] => Array.isArray(r) && r.length === 4 && r.every((v) => Number.isFinite(v)) && r[2] >= 0 && r[3] >= 0 ? [r[0], r[1], r[2], r[3]] : [0, 0, 0, 0]

/**
 * One a11y request, in order: mode (off ⇒ text-unavailable; a paragraph in
 * 'unit' mode ⇒ policy-refused) → worker job → the answer rebuilt for the
 * mode ('unit': every text; 'focused': rects only, or paragraph k's text
 * only — k out of range ⇒ policy-refused) → client limits (≤ 20,000
 * paragraphs, ≤ 4 MiB of text, every entry fits one message; else
 * copy-too-large) → 'focused' k: session quota (`used + 1`) → audit record
 * (failure ⇒ policy-refused; rects-only and empty answers write none) →
 * counter ('focused' k only) → chunks. Nothing awaits after the job returns.
 */
async function a11yTurn(c: A11yContext): Promise<void> {
  let answered = false
  const refuse = (code: CopyRefusal, nums?: { limit?: number; requested?: number }) => { answered = true; c.refuse(code, nums) }
  try {
    if (!c.alive()) return
    const mode = c.mode()
    if (mode === 'off') return refuse('text-unavailable')
    if (mode === 'unit' && c.paragraph !== undefined) return refuse('policy-refused')
    const k = mode === 'focused' ? c.paragraph : undefined
    const caps = c.caps()
    if (k !== undefined && !caps) return refuse('text-unavailable')
    let reply: TextReply
    try { reply = await c.job({ op: 'a11y', unit: c.unit, ...(k !== undefined ? { paragraph: k } : {}) }) } catch (e) {
      if (!c.alive()) return
      if (e instanceof TextTooLarge) return refuse('copy-too-large', { limit: COPY_MAX_BYTES })
      return refuse(e instanceof AdmissionRefused ? 'rate-limited' : 'policy-refused')
    }
    if (!c.alive()) return
    if (!('op' in reply) || reply.op !== 'a11y' || !Array.isArray(reply.paragraphs)) return refuse('policy-refused')
    const raw = reply.paragraphs
    // Rebuilt field by field: text only where the mode allows it.
    const withText = (i: number) => mode === 'unit' || i === k
    const paragraphs: A11yParagraph[] = raw.map((p, i) => withText(i) && typeof p.text === 'string' ? { rect: box(p.rect), text: p.text } : { rect: box(p.rect) })
    if (mode === 'unit' && paragraphs.some((p) => p.text === undefined)) return refuse('policy-refused')
    if (k !== undefined && paragraphs[k]?.text === undefined) return refuse('policy-refused')
    if (paragraphs.length > A11Y_MAX_PARAGRAPHS) return refuse('copy-too-large', { limit: COPY_MAX_BYTES })
    const bytes = paragraphs.reduce((n, p) => n + (p.text === undefined ? 0 : Buffer.byteLength(p.text, 'utf8')), 0)
    if (bytes > COPY_MAX_BYTES) return refuse('copy-too-large', { limit: COPY_MAX_BYTES, requested: bytes })
    const messages = chunkA11y(c.reqId, paragraphs)
    if (!messages) return refuse('copy-too-large', { limit: TEXT_MESSAGE_BYTES })
    if (k !== undefined) {
      const cap = caps ? caps.maxSessionParagraphs : null
      if (cap !== null && !(c.slot.copiedParagraphs + 1 <= cap)) return refuse('copy-quota-exhausted', { limit: cap, requested: 1 })
    }
    const counted = mode === 'unit' ? paragraphs.length : k !== undefined ? 1 : 0
    if (counted > 0) try { c.record({ paragraphs: counted }) } catch { return refuse('policy-refused') }
    if (k !== undefined) c.slot.copiedParagraphs += 1
    answered = true
    for (const m of messages) c.send(m)
  } catch {
    // A failure inside the turn never breaks the sid's queue.
    if (!answered) try { if (c.alive()) c.refuse('policy-refused') } catch { /* the socket is gone */ }
  }
}
