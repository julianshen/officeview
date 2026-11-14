import { crc32 } from './crc32'

export const enum Kind {
  MANIFEST = 1, ASSET_DEFINE = 2, UNIT_RASTER = 3, UNIT_COMMANDS = 4,
  XLSX_TILE = 5, SHOW_PLAN = 6, NOTES_RASTER = 7, PRINT_RASTER = 8,
}
export const FRAGMENT_CAP = 1 << 20, MAX_FRAGMENTS = 256, REASSEMBLY_MS = 5000, HEADER_BYTES = 23, EXT01_BYTES = 10
export type RejectReason = 'magic' | 'version' | 'kind' | 'unexpected-reqId' | 'flags' | 'invariant-mismatch' | 'sequence' | 'final-count' | 'missing-ext01' | 'fragment-cap' | 'length' | 'fragment-count' | 'fragment-crc' | 'total-len' | 'encoded-cap' | 'decoded-cap' | 'inflate' | 'whole-crc' | 'timeout' | 'cancelled'
export type Rejection = { ok: false; reqId?: number; reason: RejectReason }
export type FrameResult = { ok: true; kind: Kind; reqId: number; payload: Uint8Array } | Rejection
const MAGIC = 0x4453564f
const data = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength)
const validInteger = (n: number, max: number) => Number.isInteger(n) && n >= 0 && n <= max

export function encodeMessage(opts: {
  ver: number; kind: Kind; reqId: number; payload: Uint8Array;
  compress: (b: Uint8Array) => Uint8Array; fragmentBytes?: number; ext01?: boolean;
}): Uint8Array[] {
  const cap = opts.fragmentBytes ?? FRAGMENT_CAP
  if (!validInteger(opts.ver, 255) || !validInteger(opts.reqId, 0xffffffff) || !validInteger(opts.kind, 8) || opts.kind < 1 || !validInteger(cap, FRAGMENT_CAP) || cap < HEADER_BYTES + EXT01_BYTES + 1) throw new RangeError('Invalid OVSD encoder options')
  const compressed = opts.compress(opts.payload)
  const total = compressed.byteLength + 4
  const ext = !!opts.ext01 || total + HEADER_BYTES > cap
  const lengths: number[] = []
  let left = total
  while (left) {
    const room = cap - HEADER_BYTES - (lengths.length === 0 && ext ? EXT01_BYTES : 0)
    let length = Math.min(left, room)
    if (left > length && left - length < 4) length -= 4 - (left - length)
    lengths.push(length); left -= length
    if (lengths.length > MAX_FRAGMENTS) throw new RangeError('OVSD fragment-count')
  }
  const wholeCrc = crc32(opts.payload)
  let offset = 0
  return lengths.map((length, seq) => {
    const firstExt = seq === 0 && ext
    const start = HEADER_BYTES + (firstExt ? EXT01_BYTES : 0)
    const b = new Uint8Array(start + length), v = data(b)
    v.setUint32(0, MAGIC, true); v.setUint8(4, opts.ver); v.setUint8(5, opts.kind)
    v.setUint32(6, opts.reqId, true); v.setUint32(10, seq, true)
    v.setUint8(18, 4 | (firstExt ? 2 : 0) | (seq === lengths.length - 1 ? 1 : 0)); v.setUint32(19, length, true)
    if (firstExt) { v.setBigUint64(23, BigInt(total), true); v.setUint16(31, lengths.length, true) }
    b.set(compressed.subarray(offset, Math.min(offset + length, compressed.length)), start)
    if (seq === lengths.length - 1) v.setUint32(b.length - 4, wholeCrc, true)
    v.setUint32(14, crc32(b.subarray(start)), true); offset += length
    return b
  })
}

type Pending = {
  ver: number; kind: Kind; compressed: number; count: number; total: number; next: number;
  received: number; began: number; chunks: Uint8Array<ArrayBuffer>[]; inflating: boolean;
  rejected?: Rejection; cancel?: () => void;
}
type Options = {
  acceptedVersions: ReadonlySet<number>; expects: (reqId: number, kind: Kind) => boolean;
  decodedCap: (kind: Kind, reqId: number) => number; encodedCap: (kind: Kind) => number; now?: () => number;
}

/** Owns only admitted fragments. No caller-visible payload exists until all checks pass. */
export class Reassembler {
  private readonly pending = new Map<number, Pending>()
  private readonly now: () => number
  constructor(private readonly opts: Options) { this.now = opts.now ?? (() => performance.now()) }
  /** Retained encoded bytes (not declared size, nor the caller's input buffer). */
  get bufferedBytes(): number { let bytes = 0; for (const p of this.pending.values()) bytes += p.received; return bytes }
  private reject(reqId: number | undefined, reason: RejectReason): Rejection {
    const rejection: Rejection = reqId === undefined ? { ok: false, reason } : { ok: false, reqId, reason }
    if (reqId !== undefined) {
      const p = this.pending.get(reqId)
      if (p) { p.rejected = rejection; p.cancel?.(); p.chunks.length = 0; this.pending.delete(reqId) }
    }
    return rejection
  }
  /** Drops `reqId`'s pending message, if any: its buffered fragments are
   * released and an inflation in progress is cancelled (that push resolves
   * `cancelled`). Idempotent; later fragments of it are out of sequence. */
  /** A message for `reqId` is in partial reassembly or inflation. */
  has(reqId: number): boolean { return this.pending.has(reqId) }
  drop(reqId: number): void { if (this.pending.has(reqId)) this.reject(reqId, 'cancelled') }
  sweep(): Rejection[] {
    const rejections: Rejection[] = [], now = this.now()
    for (const [id, p] of this.pending) if (!p.inflating && now - p.began >= REASSEMBLY_MS) rejections.push(this.reject(id, 'timeout'))
    return rejections
  }
  async push(fragment: Uint8Array): Promise<FrameResult | null> {
    // Until the envelope is structurally valid, its claimed reqId cannot
    // authorize mutating any existing message (including cancelling inflation).
    if (fragment.length < HEADER_BYTES) return this.reject(undefined, 'length')
    const v = data(fragment)
    if (v.getUint32(0, true) !== MAGIC) return this.reject(undefined, 'magic')
    const ver = fragment[4], kind = fragment[5] as Kind
    if (!this.opts.acceptedVersions.has(ver)) return this.reject(undefined, 'version')
    if (kind < 1 || kind > 8) return this.reject(undefined, 'kind')
    const seq = v.getUint32(10, true), flags = fragment[18], compressed = flags & 4
    const final = !!(flags & 1), ext = !!(flags & 2)
    if ((flags & ~7) || (ext && seq !== 0)) return this.reject(undefined, 'flags')
    const start = HEADER_BYTES + (ext ? EXT01_BYTES : 0), length = v.getUint32(19, true)
    if (fragment.length > FRAGMENT_CAP) return this.reject(undefined, 'fragment-cap')
    if (fragment.length < start || length !== fragment.length - start) return this.reject(undefined, 'length')
    const id = v.getUint32(6, true)
    let p = this.pending.get(id)
    if (p && (p.ver !== ver || p.kind !== kind || p.compressed !== compressed)) return this.reject(id, 'invariant-mismatch')
    if (!compressed) return this.reject(undefined, 'flags')
    if (!this.opts.expects(id, kind)) return this.reject(id, 'unexpected-reqId')
    if (p && !p.inflating && this.now() - p.began >= REASSEMBLY_MS) return this.reject(id, 'timeout')
    if (p ? p.inflating || seq !== p.next : seq !== 0) return this.reject(id, 'sequence')
    let total = p?.total ?? length, count = p?.count ?? 1
    let totalBig = BigInt(total)
    if (!p && ext) { totalBig = v.getBigUint64(23, true); count = v.getUint16(31, true) }
    if (count < 1 || count > MAX_FRAGMENTS) return this.reject(id, 'fragment-count')
    if (seq >= count || final !== (seq === count - 1)) return this.reject(id, 'final-count')
    if (!p && !ext && !final) return this.reject(id, 'missing-ext01')
    const cap = this.opts.encodedCap(kind)
    if (!Number.isSafeInteger(cap) || cap < 0 || totalBig > BigInt(cap)) return this.reject(id, 'encoded-cap')
    total = Number(totalBig)
    const received = (p?.received ?? 0) + length
    if (received > total || (final && received !== total)) return this.reject(id, 'total-len')
    if (final && length < 4) return this.reject(id, 'total-len')
    if (crc32(fragment.subarray(start)) !== v.getUint32(14, true)) return this.reject(id, 'fragment-crc')
    // Every length, invariant, expectation and encoded admission check precedes the copy.
    if (!p) {
      p = { ver, kind, compressed, count, total, next: 0, received: 0, began: this.now(), chunks: [], inflating: false }
      this.pending.set(id, p)
    }
    p.chunks.push(fragment.slice(start)); p.received = received; p.next++
    if (!final) return null
    // Reassembly is complete; decoded-cap admission, not the fragment clock,
    // bounds the remaining inflation work.
    p.inflating = true
    return this.inflate(id, p)
  }
  private async inflate(id: number, p: Pending): Promise<FrameResult> {
    const last = p.chunks[p.chunks.length - 1], expectedCrc = data(last).getUint32(last.length - 4, true)
    const cap = this.opts.decodedCap(p.kind, id)
    if (!Number.isSafeInteger(cap) || cap < 0) return this.reject(id, 'decoded-cap')
    // Small compressed pulls bound the native inflater's output ahead of our admission.
    // In particular, never enqueue a complete compressed space bomb in one write.
    let part = 0, offset = 0
    const input = new ReadableStream<BufferSource>({
      pull(controller) {
        while (part < p.chunks.length) {
          const chunk = p.chunks[part], end = chunk.length - (part === p.chunks.length - 1 ? 4 : 0)
          if (offset === end) { part++; offset = 0; continue }
          const next = Math.min(end, offset + 1024)
          controller.enqueue(chunk.subarray(offset, next)); offset = next; return
        }
        controller.close()
      },
    })
    let reader: ReadableStreamDefaultReader<Uint8Array> | undefined
    try {
      reader = input.pipeThrough(new DecompressionStream('gzip')).getReader()
      const activeReader = reader
      p.cancel = () => { void activeReader.cancel().catch(() => {}) }
      const chunks: Uint8Array[] = []
      let size = 0, checksum = 0
      while (true) {
        const { done, value } = await reader.read()
        if (p.rejected) return p.rejected
        if (done) break
        if (value.length > cap - size) return this.reject(id, 'decoded-cap')
        size += value.length; checksum = crc32(value, checksum); chunks.push(value)
      }
      if (checksum !== expectedCrc) return this.reject(id, 'whole-crc')
      const payload = new Uint8Array(size)
      let at = 0; for (const chunk of chunks) { payload.set(chunk, at); at += chunk.length }
      return { ok: true, reqId: id, kind: p.kind, payload }
    } catch {
      return p.rejected ?? this.reject(id, 'inflate')
    } finally {
      await reader?.cancel().catch(() => {})
      reader?.releaseLock()
      p.cancel = undefined; p.chunks.length = 0
      if (this.pending.get(id) === p) this.pending.delete(id)
    }
  }
}
