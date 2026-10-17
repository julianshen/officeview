/**
 * One byte measure for structured-clone payloads, shared by both threads
 * (FC-T2): the worker checks a reply against its receiving grant with it right
 * before posting, and main charges the received clone with it. A structured
 * clone keeps property and element order, Map/Set order, graph sharing and each
 * view's FULL backing buffer, so both sides compute the same value.
 *
 * Conservative byte estimate of a structured-clone payload (a worker reply):
 * ArrayBuffer views by their whole buffer's byteLength (each buffer once), strings at 2 B per code
 * unit plus a header, numbers/booleans 8 B, objects/arrays a header plus their
 * slots. Walks at most `maxNodes` values; the remainder is extrapolated from
 * the average so huge payloads stay bounded in cost.
 */
export function estimatePayloadBytes(value: unknown, maxNodes = 200_000): number {
  const seen = new Set<ArrayBufferLike>()
  const stack: unknown[] = [value]
  let bytes = 0, nodes = 0, pending = 0
  while (stack.length) {
    const v = stack.pop()
    if (++nodes > maxNodes) { pending = stack.length + 1; break }
    if (v === null || v === undefined) { bytes += 8; continue }
    switch (typeof v) {
      case 'string': bytes += 16 + v.length * 2; continue
      case 'number': case 'boolean': case 'bigint': bytes += 8; continue
      case 'object': break
      default: bytes += 8; continue
    }
    if (ArrayBuffer.isView(v)) { if (!seen.has(v.buffer)) { seen.add(v.buffer); bytes += v.buffer.byteLength } continue }
    if (v instanceof ArrayBuffer) { if (!seen.has(v)) { seen.add(v); bytes += v.byteLength } continue }
    if (Array.isArray(v)) { bytes += 16 + v.length * 8; for (const item of v) stack.push(item); continue }
    if (v instanceof Map) { bytes += 32 + v.size * 24; for (const [k, item] of v) { stack.push(k); stack.push(item) }; continue }
    if (v instanceof Set) { bytes += 32 + v.size * 16; for (const item of v) stack.push(item); continue }
    const entries = Object.entries(v as Record<string, unknown>)
    bytes += 24 + entries.length * 16
    for (const [, item] of entries) stack.push(item)
  }
  return pending ? Math.ceil(bytes + (bytes / nodes) * pending) : bytes
}

/**
 * The exact measure of a structured-clone payload under the same per-value
 * costs as estimatePayloadBytes, for transport admission (FC-T2): never an
 * extrapolation. Each object is counted once (a clone keeps graph sharing);
 * strings count in O(1); primitives inside arrays, Maps and Sets are counted
 * inline, so the walk's workspace (stack + visited set) grows only with the
 * number of distinct objects. That workspace is bounded by `maxObjects`: a
 * payload with more objects is UNMEASURABLE (undefined), and the caller must
 * treat it as not admitted.
 */
export function measurePayloadBytes(value: unknown, maxObjects = 200_000): number | undefined {
  const buffers = new Set<ArrayBufferLike>(), visited = new Set<object>()
  const stack: object[] = []
  let bytes = 0
  // Adds a value's own cost; an object not yet seen is queued for its slots.
  const add = (v: unknown): boolean => {
    switch (typeof v) {
      case 'string': bytes += 16 + v.length * 2; return true
      case 'object': break
      default: bytes += 8; return true
    }
    if (v === null) { bytes += 8; return true }
    if (visited.has(v)) return true
    if (visited.size >= maxObjects) return false
    visited.add(v)
    stack.push(v)
    return true
  }
  if (!add(value)) return undefined
  while (stack.length) {
    const v = stack.pop()!
    if (ArrayBuffer.isView(v)) { if (!buffers.has(v.buffer)) { buffers.add(v.buffer); bytes += v.buffer.byteLength } continue }
    if (v instanceof ArrayBuffer) { if (!buffers.has(v)) { buffers.add(v); bytes += v.byteLength } continue }
    if (Array.isArray(v)) {
      bytes += 16 + v.length * 8
      for (let i = 0; i < v.length; i++) if (!add(v[i])) return undefined
      continue
    }
    if (v instanceof Map) {
      bytes += 32 + v.size * 24
      for (const [k, item] of v) if (!add(k) || !add(item)) return undefined
      continue
    }
    if (v instanceof Set) {
      bytes += 32 + v.size * 16
      for (const item of v) if (!add(item)) return undefined
      continue
    }
    const keys = Object.keys(v)
    bytes += 24 + keys.length * 16
    for (const key of keys) if (!add((v as Record<string, unknown>)[key])) return undefined
  }
  return bytes
}
