/**
 * Pull-style raw-deflate inflater port (design note §2.4.1, §4.1).
 *
 * The v1 implementation is pako 1.0.11's low-level zlib port driven one step
 * at a time into a caller-owned output buffer: a step can never produce more
 * than out.length bytes, reports exactly how much input it consumed, and
 * distinguishes Z_STREAM_END from "ran out of input". The high-level
 * pako.Inflate class reports truncated and trailing-junk streams as success,
 * and native DecompressionStream accepts junk after the final block on some
 * engines, so neither is used (§2.4.2).
 */
import { inflate, inflateEnd, inflateInit2 } from 'pako/lib/zlib/inflate.js'
import ZStream from 'pako/lib/zlib/zstream.js'
import { PackageError } from './package-error'

export type InflateStatus = 'progress' | 'need-input' | 'need-output' | 'end'
export interface InflateStep { status: InflateStatus; inConsumed: number; outProduced: number }
export interface InflateTotals { ended: boolean; totalIn: number; totalOut: number }
/** Consume input[inPos..], write AT MOST out.length bytes into out. The
 * signature allows promises so an async double can run the same suite. */
export interface InflaterPort {
  step(input: Uint8Array, inPos: number, out: Uint8Array): InflateStep | Promise<InflateStep>
  /** Called once when the entry's compressed range is exhausted (or after end). */
  finish(): InflateTotals | Promise<InflateTotals>
  dispose(): void | Promise<void>
}
export type InflaterFactory = () => InflaterPort

// zlib ABI return codes (pako/lib/zlib/constants.js; Z_MEM_ERROR is only in inflate.js)
const Z_NO_FLUSH = 0
const Z_OK = 0
const Z_STREAM_END = 1
const Z_NEED_DICT = 2
const Z_MEM_ERROR = -4
const Z_BUF_ERROR = -5

/** Raw deflate (windowBits -15): ZIP entries carry no zlib header. */
export function createPakoInflater(): InflaterPort {
  let strm: ZStream | null = new ZStream()
  let ended = false
  let totalIn = 0
  let totalOut = 0
  const init = inflateInit2(strm, -15)
  if (init !== Z_OK) throw new PackageError(init === Z_MEM_ERROR ? 'resource-limit' : 'parse-error', `inflateInit2 failed (${init})`)
  return {
    step(input, inPos, out) {
      if (strm === null) throw new PackageError('parse-error', 'inflater already disposed')
      if (!Number.isSafeInteger(inPos) || inPos < 0 || inPos > input.length) {
        throw new PackageError('parse-error', `inflater input position ${inPos} out of range [0, ${input.length}]`)
      }
      strm.input = input
      strm.next_in = inPos
      strm.avail_in = input.length - inPos
      strm.output = out
      strm.next_out = 0
      strm.avail_out = out.length
      const ret = inflate(strm, Z_NO_FLUSH)
      const inConsumed = strm.next_in - inPos
      const outProduced = strm.next_out
      totalIn += inConsumed
      totalOut += outProduced
      // never keep references to caller buffers between steps
      const availIn = strm.avail_in
      const availOut = strm.avail_out
      strm.input = null
      strm.output = null
      switch (ret) {
        case Z_OK: return { status: 'progress', inConsumed, outProduced }
        case Z_STREAM_END:
          ended = true
          return { status: 'end', inConsumed, outProduced }
        case Z_BUF_ERROR:
          // No progress possible: a request, not corruption (§2.4.1).
          if (availIn === 0) return { status: 'need-input', inConsumed, outProduced }
          if (availOut === 0) return { status: 'need-output', inConsumed, outProduced }
          throw new PackageError('parse-error', 'inflater made no progress')
        case Z_MEM_ERROR: throw new PackageError('resource-limit', 'inflater allocation failed')
        case Z_NEED_DICT: throw new PackageError('parse-error', 'raw deflate stream asked for a dictionary')
        default: throw new PackageError('parse-error', `corrupt deflate stream: ${strm.msg || ret}`)
      }
    },
    finish() {
      return { ended, totalIn, totalOut }
    },
    dispose() {
      if (strm === null) return
      inflateEnd(strm)
      strm = null
    },
  }
}

/** Production factory. Never constructs DecompressionStream (§2.4.2). */
export const defaultInflaterFactory: InflaterFactory = createPakoInflater
