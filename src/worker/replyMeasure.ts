/**
 * The worker's one payload walker (spec §10.2; FC-T2 pattern). A single
 * admitted, checkpointed traversal that, in one pass:
 *  - measures a reply exactly under measurePayloadBytes' rules ('clone'), so
 *    the kernel stamps `replyBytes` and main charges the received clone
 *    without walking it, and collects the reply's transferables (every
 *    distinct ArrayBuffer behind a typed array, or bare), as transferablesOf
 *    does, so posting does not walk again; or
 *  - accounts a query page under queryPayloadBytes' rules ('query', with its
 *    early stop past the limit) for the bounded query facade.
 * Every visited value is one work unit of a CheckpointBudget: its check
 * (cancellation and owner) runs at every checkpoint and the slice yields at
 * 4,096 units or when its time budget is spent. A walk that needs no pause
 * completes synchronously (small replies keep their posting order); one that
 * does continues in the same iterator after the yield, never from scratch.
 * Worker-only: main never imports this module.
 */
import type { CheckpointBudget } from '../xlsx/view/scheduler'
import { MEASURE_CHUNK_ENTRIES, MEASURE_ENTRY_BYTES, MEASURE_FIRST_CHUNK_ENTRIES } from '../xlsx/view/payloadBytes'
import { isArrayBuffer } from './viewerProtocol'

const hasOwn = Object.prototype.hasOwnProperty

export type WalkOutcome =
  | { kind: 'measured'; bytes: number; transfer: ArrayBuffer[] }
  /** The document pool refused the walk's workspace: nothing measured. */
  | { kind: 'refused' }
  /** The budget's check threw StopWalk: the result will not be published. */
  | { kind: 'stopped' }
/** Thrown by a walk budget's check to abandon the walk without an error. */
export class StopWalk extends Error {}
export interface WalkOptions {
  /** 'clone' (default): measurePayloadBytes rules plus transfers; 'query': queryPayloadBytes rules. */
  model?: 'clone' | 'query'
  /** 'query' only: stop once the bytes exceed this. */
  limit?: number
}

interface Tally { bytes: number; transfer: Set<ArrayBuffer> }

/** One pause point per visited value; returns false when admission is refused. */
function* walk(value: unknown, admit: (bytes: number) => boolean, options: WalkOptions, tally: Tally): Generator<void, boolean, void> {
  const query = options.model === 'query', limit = options.limit ?? Number.POSITIVE_INFINITY
  const buffers = new Set<ArrayBufferLike>(), visited = new Set<object>()
  const stack: unknown[] = []
  let room = 0
  // Admits one more distinct object's workspace (visited slot, stack slot, buffer slot).
  const enter = (v: object): boolean => {
    if (room === 0) {
      const chunk = Math.min(MEASURE_CHUNK_ENTRIES, Math.max(MEASURE_FIRST_CHUNK_ENTRIES, visited.size))
      if (!admit(chunk * MEASURE_ENTRY_BYTES)) return false
      room = chunk
    }
    room--
    visited.add(v)
    return true
  }
  const buffer = (b: ArrayBufferLike) => {
    if (buffers.has(b)) return
    buffers.add(b)
    tally.bytes += b.byteLength
    if (!query && isArrayBuffer(b)) tally.transfer.add(b)
  }
  if (query) {
    // queryPayloadBytes: every value is pushed; objects are counted once.
    stack.push(value)
    while (stack.length && tally.bytes <= limit) {
      const item = stack.pop()
      yield
      if (typeof item === 'string') { tally.bytes += 16 + 2 * item.length; continue }
      if (typeof item !== 'object' || item === null) { tally.bytes += 8; continue }
      if (visited.has(item)) continue
      if (!enter(item)) return false
      if (ArrayBuffer.isView(item)) { stack.push(item.buffer); tally.bytes += 24; continue }
      if (Object.prototype.toString.call(item) === '[object ArrayBuffer]') { tally.bytes += (item as ArrayBuffer).byteLength; continue }
      const values = Array.isArray(item) ? item : Object.values(item)
      tally.bytes += 24 + values.length * 16
      if (tally.bytes > limit) break
      for (const child of values) stack.push(child)
    }
    return true
  }
  // measurePayloadBytes: primitives inline, each distinct object queued once.
  const add = (v: unknown): boolean => {
    switch (typeof v) {
      case 'string': tally.bytes += 16 + v.length * 2; return true
      case 'object': break
      default: tally.bytes += 8; return true
    }
    if (v === null) { tally.bytes += 8; return true }
    if (visited.has(v)) return true
    if (!enter(v)) return false
    stack.push(v)
    return true
  }
  if (!add(value)) return false
  while (stack.length) {
    const v = stack.pop() as object
    yield
    if (ArrayBuffer.isView(v)) { buffer(v.buffer); continue }
    if (v instanceof ArrayBuffer) { buffer(v); continue }
    if (Array.isArray(v)) {
      tally.bytes += 16 + v.length * 8
      for (let i = 0; i < v.length; i++) { if (!add(v[i])) return false; yield }
      continue
    }
    if (v instanceof Map) {
      tally.bytes += 32 + v.size * 24
      for (const [k, item] of v) { if (!add(k) || !add(item)) return false; yield }
      continue
    }
    if (v instanceof Set) {
      tally.bytes += 32 + v.size * 16
      for (const item of v) { if (!add(item)) return false; yield }
      continue
    }
    tally.bytes += 24
    // A structured clone holds own data properties only.
    for (const key in v) {
      if (!hasOwn.call(v, key)) continue
      tally.bytes += 16
      if (!add((v as Record<string, unknown>)[key])) return false
      yield
    }
  }
  return true
}

/** Runs the walk under `budget`: synchronously while no pause is due, then in
 * the same iterator after each yield. A StopWalk from the budget's check ends it
 * as 'stopped'; any other error propagates. */
export function walkPayload(value: unknown, admit: (bytes: number) => boolean, budget: CheckpointBudget, options: WalkOptions = {}): WalkOutcome | Promise<WalkOutcome> {
  const tally: Tally = { bytes: 0, transfer: new Set() }
  const steps = walk(value, admit, options, tally)
  const stopped = (cause: unknown): WalkOutcome => { if (cause instanceof StopWalk) return { kind: 'stopped' }; throw cause }
  const drive = (): WalkOutcome | Promise<WalkOutcome> => {
    try {
      for (;;) {
        const step = steps.next()
        if (step.done) {
          if (!step.value) return { kind: 'refused' }
          budget.checkpoint()
          return { kind: 'measured', bytes: tally.bytes, transfer: [...tally.transfer] }
        }
        if (budget.spend(1)) return budget.pause().then(drive, stopped)
      }
    } catch (cause) { return stopped(cause) }
  }
  return drive()
}
