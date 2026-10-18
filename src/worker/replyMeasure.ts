/**
 * Worker-side reply sizing (spec §10.2; FC-T2 pattern): the worker measures a
 * provider result's reply before posting it and stamps the exact size on it as
 * `replyBytes`, so main charges the received clone without walking it.
 * Worker-only: main never imports this module.
 */
import { CheckpointBudget, YIELD_WORK_UNITS } from '../xlsx/view/scheduler'
import { MEASURE_CHUNK_ENTRIES, MEASURE_ENTRY_BYTES, MEASURE_FIRST_CHUNK_ENTRIES } from '../xlsx/view/payloadBytes'

const hasOwn = Object.prototype.hasOwnProperty

/** measurePayloadBytes as a walk that pauses (yields to its driver) after
 * every visited value: the same exact value and the same admitted workspace. */
function* walk(value: unknown, admit: (bytes: number) => boolean): Generator<void, number | undefined, void> {
  const buffers = new Set<ArrayBufferLike>(), visited = new Set<object>()
  const stack: object[] = []
  let bytes = 0, room = 0
  const add = (v: unknown): boolean => {
    switch (typeof v) {
      case 'string': bytes += 16 + v.length * 2; return true
      case 'object': break
      default: bytes += 8; return true
    }
    if (v === null) { bytes += 8; return true }
    if (visited.has(v)) return true
    if (room === 0) {
      const chunk = Math.min(MEASURE_CHUNK_ENTRIES, Math.max(MEASURE_FIRST_CHUNK_ENTRIES, visited.size))
      if (!admit(chunk * MEASURE_ENTRY_BYTES)) return false
      room = chunk
    }
    room--
    visited.add(v)
    stack.push(v)
    return true
  }
  if (!add(value)) return undefined
  while (stack.length) {
    const v = stack.pop()!
    yield
    if (ArrayBuffer.isView(v)) { if (!buffers.has(v.buffer)) { buffers.add(v.buffer); bytes += v.buffer.byteLength } continue }
    if (v instanceof ArrayBuffer) { if (!buffers.has(v)) { buffers.add(v); bytes += v.byteLength } continue }
    if (Array.isArray(v)) {
      bytes += 16 + v.length * 8
      for (let i = 0; i < v.length; i++) { if (!add(v[i])) return undefined; yield }
      continue
    }
    if (v instanceof Map) {
      bytes += 32 + v.size * 24
      for (const [k, item] of v) { if (!add(k) || !add(item)) return undefined; yield }
      continue
    }
    if (v instanceof Set) {
      bytes += 32 + v.size * 16
      for (const item of v) { if (!add(item)) return undefined; yield }
      continue
    }
    bytes += 24
    // A structured clone holds own data properties only.
    for (const key in v) {
      if (!hasOwn.call(v, key)) continue
      bytes += 16
      if (!add((v as Record<string, unknown>)[key])) return undefined
      yield
    }
  }
  return bytes
}

/** The exact size when the walk takes at most YIELD_WORK_UNITS visits (one
 * slice, measured synchronously: small replies keep their posting order), or
 * null when it needs more and must be measured in steps. `undefined` bytes: the
 * pool refused the walk's workspace. */
export function measureWithinSlice(value: unknown, admit: (bytes: number) => boolean): { bytes: number | undefined } | null {
  const steps = walk(value, admit)
  for (let units = 0; ; units++) {
    const step = steps.next()
    if (step.done) return { bytes: step.value }
    if (units >= YIELD_WORK_UNITS) return null
  }
}

/** measurePayloadBytes in budget steps: at most 4,096 visited values between
 * real yields and a check at every checkpoint (the budget's check throws to
 * abandon the walk). */
export async function measurePayloadBytesStepped(value: unknown, admit: (bytes: number) => boolean, budget: CheckpointBudget): Promise<number | undefined> {
  const steps = walk(value, admit)
  for (;;) {
    const step = steps.next()
    if (step.done) { await budget.finish(); return step.value }
    if (budget.spend(1)) await budget.pause()
  }
}
