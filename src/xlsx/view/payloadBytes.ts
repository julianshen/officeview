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
