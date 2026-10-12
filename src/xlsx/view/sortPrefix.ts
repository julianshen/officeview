import type { CooperativeSlice } from './scheduler'
import type { SortLane } from './sort'

export interface PrefixMemory {
  canCharge(bytes: number): boolean
  charge(bytes: number): void
  release(bytes: number): void
  allocate<T>(create: () => T): T
}
export interface SortPrefix { codes: Uint32Array; keys: number; bytes: number }
/** Order-preserving prefix of private direction-normalized integer key lanes.
 * Tags are the major digit; values occupy disjoint ranges within each tag.
 * Multiplication is bounded BEFORE storage, never rounded or bitwise packed. */
export async function buildSortPrefix<Phase extends string>(
  lanes: readonly SortLane[], length: number, memory: PrefixMemory,
  slice: CooperativeSlice<Phase>, phase: Phase,
): Promise<SortPrefix | undefined> {
  if (lanes.length < 4 || !length || !memory.canCharge(512)) return undefined
  memory.charge(512)
  let retained = false, charged = 512
  try {
    const ranges = memory.allocate(() => new Array<{ min: number; tag: number; span: number; radix: number }>())
    let product = 1
    for (let key = 0; key < Math.min(4, lanes.length); key++) {
      const lane = lanes[key]
      let min = Infinity, max = -Infinity, tag = Infinity, lastTag = -Infinity, valid = true
      for (let at = 0; at < length; at++) {
        const value = lane.numbers[at], type = lane.tags[at]
        if (slice.due()) await slice.pause(phase)
        if (!Number.isSafeInteger(value)) { valid = false; break }
        if (value < min) min = value
        if (value > max) max = value
        if (type < tag) tag = type
        if (type > lastTag) lastTag = type
      }
      const span = max - min + 1, radix = span * (lastTag - tag + 1)
      if (!valid || !Number.isSafeInteger(radix) || radix < 1 || radix > 4294967296
        || product > Math.floor(4294967296 / radix)) break
      ranges.push({ min, tag, span, radix })
      product *= radix
    }
    if (ranges.length < 2 || !memory.canCharge(length * 4)) return undefined
    memory.charge(length * 4)
    charged += length * 4
    const codes = memory.allocate(() => new Uint32Array(length))
    for (let at = 0; at < length; at++) {
      let code = 0
      for (let key = 0; key < ranges.length; key++) {
        const range = ranges[key], lane = lanes[key]
        code = code * range.radix + (lane.tags[at] - range.tag) * range.span + (lane.numbers[at] - range.min)
      }
      codes[at] = code
      if (slice.due()) await slice.pause(phase)
    }
    retained = true
    return { codes, keys: ranges.length, bytes: length * 4 + 512 }
  } finally { if (!retained) memory.release(charged) }
}
