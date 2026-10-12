import type { CooperativeSlice } from './scheduler'

export const NUMERIC_ASCII = ' -.0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz'
export interface TextKeyMemory {
  charge(bytes: number): void
  release(bytes: number): void
  allocate<T>(create: () => T): T
}
const digit = (unit: number) => unit >= 48 && unit <= 57
const allowed = (unit: number) => digit(unit) || unit === 32 || unit === 45 || unit === 46 || unit === 95
  || unit >= 65 && unit <= 90 || unit >= 97 && unit <= 122

/** Check the host's numeric primary order, including every claimed prefix unit.
 * No character weights, case folding or locale contractions are approximated. */
async function certified<Phase extends string>(collator: Intl.Collator, slice: CooperativeSlice<Phase>, phase: Phase): Promise<boolean> {
  if (!collator.resolvedOptions().numeric) return false
  for (let a = 0; a < 100; a++) for (let b = 0; b < 100; b++) {
    if (Math.sign(collator.compare(String(a), String(b))) !== Math.sign(a - b)) return false
    if (slice.due()) await slice.pause(phase)
  }
  for (const prefix of NUMERIC_ASCII) for (let a = 0; a < 10; a++) for (let b = 0; b < 10; b++) {
    // Digit prefixes are part of the same numeric run; the same leading digit
    // preserves this order. Unequal numeric primary weights dominate suffixes.
    if (a !== b && Math.sign(collator.compare(prefix + a + 'z', prefix + b + 'a')) !== Math.sign(a - b)) return false
    if (slice.due()) await slice.pause(phase)
  }
  return true
}

/** Only discriminate unequal <=15-digit ASCII runs in strings of at most 64 units, after identical prefixes.
 * With numeric:true that run supplies the first different PRIMARY weight; later
 * weights cannot reverse it. Equal numbers need host secondary/tertiary weights.
 * Outside the claimed alphabet, or ICU's long-number segmentation, use the host. */
export async function numericTextKeys<Phase extends string>(
  values: readonly string[], collator: Intl.Collator, memory: TextKeyMemory,
  slice: CooperativeSlice<Phase>, phase: Phase,
): Promise<{ compare(a: number, b: number): number; dispose(): void }> {
  const host = (a: number, b: number) => collator.compare(values[a], values[b])
  if (!values.length || !(await certified(collator, slice, phase))) return { compare: host, dispose: () => {} }
  const bytes = values.length * 12
  memory.charge(bytes)
  const starts = memory.allocate(() => new Uint32Array(values.length))
  const numbers = memory.allocate(() => new Float64Array(values.length))
  for (let id = 0; id < values.length; id++) {
    const value = values[id]
    if (value.length > 64) {
      if (slice.due()) await slice.pause(phase)
      continue
    }
    let start = -1, number = 0, digits = 0, closed = false, valid = true
    for (let at = 0; at < value.length; at++) {
      const unit = value.charCodeAt(at)
      if (!allowed(unit)) { valid = false; break }
      if (digit(unit) && !closed) {
        if (start < 0) start = at
        number = number * 10 + unit - 48
        if (++digits > 15) { valid = false; break }
      } else if (start >= 0) closed = true
      if (slice.due()) await slice.pause(phase)
    }
    if (valid && start >= 0) { starts[id] = start + 1; numbers[id] = number }
    if (slice.due()) await slice.pause(phase)
  }
  let disposed = false
  return {
    compare(a, b) {
      const start = starts[a]
      if (start && start === starts[b] && numbers[a] !== numbers[b]) {
        const first = values[a], second = values[b]
        for (let at = 0; at < start - 1; at++) if (first.charCodeAt(at) !== second.charCodeAt(at)) return host(a, b)
        return numbers[a] - numbers[b]
      }
      return host(a, b)
    },
    dispose() { if (!disposed) { memory.release(bytes); disposed = true } },
  }
}
