/** `seen` may be shared across calls so parts of one payload measured
 * separately count each shared object once, as one whole-payload call would.
 * Bounded, cycle-safe clone estimate with an early stop, including UTF-16 text
 * and complete backing buffers (not just a small typed-array view). */
export function queryPayloadBytes(value: unknown, limit: number, seen = new Set<object>()): number {
  const stack: unknown[] = [value]
  let bytes = 0
  while (stack.length && bytes <= limit) {
    const item = stack.pop()
    if (typeof item === 'string') { bytes += 16 + 2 * item.length; continue }
    if (typeof item !== 'object' || item === null) { bytes += 8; continue }
    if (seen.has(item)) continue
    seen.add(item)
    if (ArrayBuffer.isView(item)) { stack.push(item.buffer); bytes += 24; continue }
    if (Object.prototype.toString.call(item) === '[object ArrayBuffer]') { bytes += (item as ArrayBuffer).byteLength; continue }
    const values = Array.isArray(item) ? item : Object.values(item)
    bytes += 24 + values.length * 16
    if (bytes > limit) break
    for (const child of values) stack.push(child)
  }
  return bytes
}
