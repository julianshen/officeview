/**
 * Stage 3a — main-thread text operations over the wire (Task 4: search).
 *
 * `ws.ts` authenticates and dispatches; this module owns argument validation
 * (R6), the batched unit walk, the ≤ 48 KiB message chunking and per-sid
 * supersession. It takes callbacks, never a socket, so it is unit-testable.
 */
import { AdmissionRefused } from '../pool'
import type { TextJob, TextReply } from './served'

/** R6: a query is 1..256 UTF-16 units. */
export const MAX_QUERY_UNITS = 256
/** R6: at most this many matches are sent for one search (`truncated: true` beyond). */
export const MAX_SEARCH_MATCHES = 5000
/** Units per `Pool.text` search job. */
export const SEARCH_BATCH_UNITS = 8
/** Every server→client text message is ≤ 48 KiB of UTF-8. */
export const TEXT_MESSAGE_BYTES = 48 * 1024

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

/** One match whose JSON exceeds `room` bytes, split into pieces (same unit,
 * whole x,y,w,h quads) that each fit. */
function splitMatch(m: SearchMatch, room: number): SearchMatch[] {
  const base = bytesOf({ unit: m.unit, rects: [] })
  const out: SearchMatch[] = []
  let rects: number[] = [], size = base
  for (let i = 0; i < m.rects.length; i += 4) {
    const quad = m.rects.slice(i, i + 4)
    const n = enc.encode(quad.map((x) => JSON.stringify(x)).join(',')).length
    if (base + n > room) throw new Error('a single rect exceeds the message limit')
    if (rects.length && size + 1 + n > room) { out.push({ unit: m.unit, rects }); rects = []; size = base }
    size += (rects.length ? 1 : 0) + n
    rects.push(...quad)
  }
  if (rects.length) out.push({ unit: m.unit, rects })
  return out
}

/** `search-result` messages for `matches`, in order, each ≤ `limit` bytes of
 * UTF-8 when JSON-encoded. A match is split only when it alone cannot fit. */
export function chunkMatches(reqId: number, matches: readonly SearchMatch[], limit = TEXT_MESSAGE_BYTES): SearchResultMessage[] {
  const base = bytesOf({ reqId, type: 'search-result', matches: [] })
  const out: SearchResultMessage[] = []
  let cur: SearchMatch[] = [], size = base
  const add = (m: SearchMatch, n: number) => {
    if (cur.length && size + 1 + n > limit) { out.push({ reqId, type: 'search-result', matches: cur }); cur = []; size = base }
    size += (cur.length ? 1 : 0) + n
    cur.push(m)
  }
  for (const raw of matches) {
    const m = { unit: raw.unit, rects: raw.rects }
    const n = bytesOf(m)
    if (base + n <= limit) add(m, n)
    else for (const piece of splitMatch(m, limit - base)) add(piece, bytesOf(piece))
  }
  if (cur.length) out.push({ reqId, type: 'search-result', matches: cur })
  return out
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
 * asks for one match more than may still be sent, so truncation is exact. A
 * pool failure is a refusal (AdmissionRefused ⇒ rate-limited, anything else
 * ⇒ policy-refused) and no `search-done` follows it.
 */
export async function runSearch(c: SearchContext): Promise<void> {
  const prev = supersede(c.slot)
  let settle!: () => void
  const run: SearchRun = { reqId: c.reqId, sent: 0, stopped: false, settled: new Promise<void>((r) => { settle = r }), send: c.send, alive: c.alive }
  c.slot.search = run
  const current = () => !run.stopped && c.slot.search === run && c.alive()
  const done = (truncated: boolean) => { run.stopped = true; c.send({ reqId: c.reqId, type: 'search-done', total: run.sent, truncated }) }
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
      for (const m of chunkMatches(c.reqId, send)) c.send(m)
      run.sent += send.length
      if (reply.matches.length > remaining) return done(true)
    }
    if (current()) done(false)
  } finally {
    run.stopped = true
    if (c.slot.search === run) c.slot.search = undefined
    settle()
  }
}
