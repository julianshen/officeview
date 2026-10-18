// Local declarations for the pako 1.0.11 low-level zlib modules (no bundled
// types, no @types/pako). Only what src/core/inflater.ts and its conformance
// test use. Keep in sync with the exact pin in package.json.

declare module 'pako/lib/zlib/zstream.js' {
  class ZStream {
    input: Uint8Array | null
    next_in: number
    avail_in: number
    total_in: number
    output: Uint8Array | null
    next_out: number
    avail_out: number
    total_out: number
    msg: string
    state: unknown
    data_type: number
    adler: number
  }
  export default ZStream
}

declare module 'pako/lib/zlib/inflate.js' {
  import type ZStream from 'pako/lib/zlib/zstream.js'
  export function inflateInit2(strm: ZStream, windowBits: number): number
  export function inflate(strm: ZStream, flush: number): number
  export function inflateEnd(strm: ZStream): number
}

declare module 'pako/lib/zlib/constants.js' {
  const constants: {
    Z_NO_FLUSH: number
    Z_OK: number
    Z_STREAM_END: number
    Z_NEED_DICT: number
    Z_STREAM_ERROR: number
    Z_DATA_ERROR: number
    Z_BUF_ERROR: number
  }
  export default constants
}
