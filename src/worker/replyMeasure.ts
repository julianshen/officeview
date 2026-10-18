/**
 * The one payload walker (spec §10.2; FC-T2 pattern). A single admitted,
 * checkpointed traversal that, in one pass:
 *  - measures a reply exactly under measurePayloadBytes' rules ('clone'), so
 *    the kernel stamps `replyBytes` and main charges the received clone
 *    without walking it, and collects the reply's transferables (every
 *    distinct ArrayBuffer behind a typed array, or bare), as transferablesOf
 *    does, so posting does not walk again; or
 *  - accounts a payload under queryPayloadBytes' rules ('query', with its
 *    early stop past the limit).
 * The walk keeps one frame per open container (a cursor over its children),
 * never a list of pending children: primitives are counted as the cursor reads
 * them, and an object is entered (its workspace admitted) only when reached.
 * Every visited value is one work unit of a CheckpointBudget: its check
 * (cancellation and owner) runs at every checkpoint and the slice yields at
 * 4,096 units or when its time budget is spent. A walk that needs no pause
 * completes synchronously; one that does continues in the same frames after
 * the yield, never from scratch. Main-safe (main uses it for its own
 * accounting with the main ledger as the admitting pool).
 */
import type { CheckpointBudget } from '../xlsx/view/scheduler'
import { MEASURE_CHUNK_ENTRIES, MEASURE_ENTRY_BYTES, MEASURE_FIRST_CHUNK_ENTRIES } from '../xlsx/view/payloadBytes'
import { isArrayBuffer } from './viewerProtocol'

const hasOwn = Object.prototype.hasOwnProperty
/** Workspace of one open-container frame (its cursor/generator object). */
export const WALK_FRAME_BYTES = 256
/** Frames a walk without deduplication may hold (its constant workspace); deeper payloads refuse. */
const UNSHARED_MAX_FRAMES = 256

export type WalkOutcome =
  | { kind: 'measured'; bytes: number; transfer: ArrayBuffer[]; peakFrames: number; admittedFrames: number }
  /** The pool refused the walk's workspace: nothing measured. */
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
  /** false: count shared objects at every occurrence (an over-estimate, never
   * less than the exact measure) with no visited set, in constant workspace:
   * at most UNSHARED_MAX_FRAMES frames, admitted to nobody. Acyclic payloads only. */
  dedupe?: boolean
}

interface Tally { bytes: number; transfer: Set<ArrayBuffer>; peak: number; admittedFrames: number }

/** One pause point per visited value; returns false when admission is refused. */
function* walk(value: unknown, admit: (bytes: number) => boolean, options: WalkOptions, tally: Tally): Generator<void, boolean, void> {
  const query = options.model === 'query', dedupe = options.dedupe !== false
  const limit = options.limit ?? Number.POSITIVE_INFINITY
  const buffers = new Set<ArrayBufferLike>(), visited = new Set<object>()
  const frames: Iterator<unknown>[] = []
  let room = 0
  // Admits one more distinct object's workspace (visited slot, stack slot, buffer slot).
  const enter = (v: object): boolean => {
    if (!dedupe) return true
    if (room === 0) {
      const chunk = Math.min(MEASURE_CHUNK_ENTRIES, Math.max(MEASURE_FIRST_CHUNK_ENTRIES, visited.size))
      if (!admit(chunk * MEASURE_ENTRY_BYTES)) return false
      room = chunk
    }
    room--
    visited.add(v)
    return true
  }
  const seen = (v: object): boolean => dedupe && visited.has(v)
  // Frames are admitted as the depth first exceeds what was admitted.
  const open = (frame: Iterator<unknown>): boolean => {
    if (frames.length >= tally.admittedFrames) {
      if (!dedupe) {
        if (frames.length >= UNSHARED_MAX_FRAMES) return false
        tally.admittedFrames = frames.length + 1
      } else {
        const chunk = Math.max(4, tally.admittedFrames)
        if (!admit(chunk * WALK_FRAME_BYTES)) return false
        tally.admittedFrames += chunk
      }
    }
    frames.push(frame)
    tally.peak = Math.max(tally.peak, frames.length)
    return true
  }
  const buffer = (b: ArrayBufferLike) => {
    if (dedupe && buffers.has(b)) return
    if (dedupe) buffers.add(b)
    tally.bytes += b.byteLength
    if (!query && isArrayBuffer(b)) tally.transfer.add(b)
  }
  function* items(v: readonly unknown[]): Generator<unknown, void, void> { for (let i = 0; i < v.length; i++) yield v[i] }
  function* entries(v: Map<unknown, unknown>): Generator<unknown, void, void> { for (const [k, item] of v) { yield k; yield item } }
  function* members(v: Set<unknown>): Generator<unknown, void, void> { for (const item of v) yield item }
  // A structured clone (and Object.values) holds own enumerable data properties.
  function* values(v: object): Generator<unknown, void, void> {
    for (const key in v) {
      if (!hasOwn.call(v, key)) continue
      tally.bytes += 16
      yield (v as Record<string, unknown>)[key]
    }
  }
  // Counts one reached value; an object not yet seen is entered and, when it
  // is a container, opened as a frame. False: admission refused.
  const reach = (v: unknown): boolean => {
    if (typeof v === 'string') { tally.bytes += 16 + 2 * v.length; return true }
    if (typeof v !== 'object' || v === null) { tally.bytes += 8; return true }
    if (seen(v)) return true
    if (!enter(v)) return false
    if (query) {
      if (ArrayBuffer.isView(v)) {
        tally.bytes += 24
        const b = v.buffer
        if (seen(b)) return true
        if (!enter(b)) return false
        tally.bytes += b.byteLength
        return true
      }
      if (Object.prototype.toString.call(v) === '[object ArrayBuffer]') { tally.bytes += (v as ArrayBuffer).byteLength; return true }
      if (Array.isArray(v)) { tally.bytes += 24 + v.length * 16; return open(items(v)) }
      tally.bytes += 24
      return open(values(v))
    }
    if (ArrayBuffer.isView(v)) { buffer(v.buffer); return true }
    if (v instanceof ArrayBuffer) { buffer(v); return true }
    if (Array.isArray(v)) { tally.bytes += 16 + v.length * 8; return open(items(v)) }
    if (v instanceof Map) { tally.bytes += 32 + v.size * 24; return open(entries(v)) }
    if (v instanceof Set) { tally.bytes += 32 + v.size * 16; return open(members(v)) }
    tally.bytes += 24
    return open(values(v))
  }
  if (!reach(value)) return false
  while (frames.length && tally.bytes <= limit) {
    // The unit is charged before the cursor reads the value.
    yield
    const next = frames[frames.length - 1].next()
    if (next.done) { frames.pop(); continue }
    if (!reach(next.value)) return false
  }
  return true
}

/** Workspace a walk with no pool to admit to may use in all: replies of
 * commands with no live document (a kernel-built close), never a document's
 * payload. */
export const UNPOOLED_WALK_BYTES = 64 * 1024
/** The admitting function of a walk: chunks admitted to the pool share
 * `workspace` (cumulative), or at most UNPOOLED_WALK_BYTES without one. */
export function walkAdmission(workspace: { set(bytes: number): boolean } | undefined): (bytes: number) => boolean {
  let held = 0
  return more => workspace ? workspace.set(held += more) : (held += more) <= UNPOOLED_WALK_BYTES
}
/** Runs the walk under `budget`: synchronously while no pause is due, then in
 * the same frames after each yield. A StopWalk from the budget's check ends it
 * as 'stopped'; any other error propagates. */
export function walkPayload(value: unknown, admit: (bytes: number) => boolean, budget: CheckpointBudget, options: WalkOptions = {}): WalkOutcome | Promise<WalkOutcome> {
  const tally: Tally = { bytes: 0, transfer: new Set(), peak: 0, admittedFrames: 0 }
  const steps = walk(value, admit, options, tally)
  const stopped = (cause: unknown): WalkOutcome => { if (cause instanceof StopWalk) return { kind: 'stopped' }; throw cause }
  const drive = (): WalkOutcome | Promise<WalkOutcome> => {
    try {
      for (;;) {
        const step = steps.next()
        if (step.done) {
          if (!step.value) return { kind: 'refused' }
          budget.checkpoint()
          return { kind: 'measured', bytes: tally.bytes, transfer: [...tally.transfer], peakFrames: tally.peak, admittedFrames: tally.admittedFrames }
        }
        if (budget.spend(1)) return budget.pause().then(drive, stopped)
      }
    } catch (cause) { return stopped(cause) }
  }
  return drive()
}
