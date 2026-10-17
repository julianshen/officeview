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

/** Workspace the exact walk admits per distinct object: a visited-Set slot
 * (64 B, hash table included), a stack reference (8 B) and, for a typed view,
 * a buffer-dedup Set slot (64 B). Every entry is admitted before the walk
 * creates it, in chunks that start at MEASURE_FIRST_CHUNK_ENTRIES and double
 * up to MEASURE_CHUNK_ENTRIES (the admitted workspace stays within twice what
 * the walk holds). Small kernel-built replies are sized by smallTreeBytes. */
export const MEASURE_ENTRY_BYTES = 136
export const MEASURE_FIRST_CHUNK_ENTRIES = 64
export const MEASURE_CHUNK_ENTRIES = 4096
const hasOwn = Object.prototype.hasOwnProperty

/**
 * The exact measure of a structured-clone payload under the same per-value
 * costs as estimatePayloadBytes, for transport admission (FC-T2): never an
 * extrapolation, and no object-count cap. Each object is counted once (a clone
 * keeps graph sharing); strings count in O(1); primitives inside arrays, Maps
 * and Sets are counted inline and plain objects are walked with for...in, so
 * the walk's only growing workspace is its visited set, buffer set and stack,
 * one entry per distinct object. Before that workspace grows by a chunk,
 * `admit` must accept the chunk's bytes (entries x MEASURE_ENTRY_BYTES); a refusal makes the payload UNMEASURABLE now
 * (undefined). Callers walking a large untrusted graph pass an admitting pool.
 */
export function measurePayloadBytes(value: unknown, admit: (bytes: number) => boolean = () => true): number | undefined {
  const buffers = new Set<ArrayBufferLike>(), visited = new Set<object>()
  const stack: object[] = []
  let bytes = 0, room = 0
  // Adds a value's own cost; an object not yet seen is queued for its slots.
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
    bytes += 24
    // A structured clone holds own data properties only.
    for (const key in v) {
      if (!hasOwn.call(v, key)) continue
      bytes += 16
      if (!add((v as Record<string, unknown>)[key])) return undefined
    }
  }
  return bytes
}

const SMALL_TREE_NODES = 64
const SMALL_TREE_DEPTH = 8
/**
 * measurePayloadBytes of a small TREE of plain objects, arrays and primitives
 * (a kernel-built payload-free reply: refusal, cancellation, error), computed
 * by bounded recursion with no Sets or arrays: its workspace is a constant
 * (at most SMALL_TREE_DEPTH frames), so it needs no admission and a full pool
 * can still size a refusal. Anything else (more than SMALL_TREE_NODES
 * objects, deeper, buffers, Maps, Sets) is undefined: use the admitted walk.
 * Without sharing, a tree's size equals the exact measure.
 */
export function smallTreeBytes(value: unknown): number | undefined {
  let nodes = 0
  const size = (v: unknown, depth: number): number => {
    switch (typeof v) {
      case 'string': return 16 + v.length * 2
      case 'object': break
      default: return 8
    }
    if (v === null) return 8
    if (++nodes > SMALL_TREE_NODES || depth > SMALL_TREE_DEPTH || ArrayBuffer.isView(v) || v instanceof ArrayBuffer || v instanceof Map || v instanceof Set) return NaN
    let bytes = 0
    if (Array.isArray(v)) {
      bytes = 16 + v.length * 8
      for (let i = 0; i < v.length; i++) bytes += size(v[i], depth + 1)
      return bytes
    }
    bytes = 24
    for (const key in v) if (hasOwn.call(v, key)) bytes += 16 + size((v as Record<string, unknown>)[key], depth + 1)
    return bytes
  }
  const bytes = size(value, 0)
  return Number.isNaN(bytes) ? undefined : bytes
}
