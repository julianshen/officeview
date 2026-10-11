/** Programmatic ZIP writer for the S1.1 reader tests. It returns the byte
 * layout (offsets of every record) so the §2.5 matrix can forge a single
 * field of an otherwise valid archive with setU16/setU32/setU64. */
import { deflateRawSync } from 'node:zlib'
import { crc32 } from '../../src/core/crc32'

const encoder = new TextEncoder()

export interface BuildEntry {
  name: string
  data?: Uint8Array | string
  method?: number
  /** Use these bytes as the entry's compressed data instead of compressing `data`. */
  raw?: Uint8Array
  flags?: number
  crc?: number
  compressedSize?: number
  uncompressedSize?: number
  /** Override the local header name (parser-differential tests). */
  localName?: string
  localExtra?: Uint8Array
  centralExtra?: Uint8Array
  /** Default true: bit 3 set, zero sizes in the local header, descriptor after the data. */
  dataDescriptor?: boolean
}

export interface BuildOptions {
  zip64?: boolean
  comment?: Uint8Array
  /** Bytes appended after the EOCD (then the comment length no longer matches). */
  trailer?: Uint8Array
}

export interface EntryLayout { localOffset: number; dataStart: number; dataLength: number; centralOffset: number }
export interface ZipLayout {
  bytes: Uint8Array
  entries: EntryLayout[]
  cdOffset: number
  cdSize: number
  eocdOffset: number
  zip64EocdOffset?: number
  locatorOffset?: number
}

export function setU16(bytes: Uint8Array, at: number, value: number): void {
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint16(at, value, true)
}
export function setU32(bytes: Uint8Array, at: number, value: number): void {
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setUint32(at, value >>> 0, true)
}
export function setU64(bytes: Uint8Array, at: number, value: number | bigint): void {
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setBigUint64(at, BigInt(value), true)
}
export function getU32(bytes: Uint8Array, at: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(at, true)
}

class Writer {
  private parts: Uint8Array[] = []
  length = 0
  push(bytes: Uint8Array): number {
    const at = this.length
    this.parts.push(bytes)
    this.length += bytes.length
    return at
  }
  bytes(): Uint8Array {
    const out = new Uint8Array(this.length)
    let at = 0
    for (const p of this.parts) { out.set(p, at); at += p.length }
    return out
  }
}

function record(size: number, fill: (view: DataView) => void): Uint8Array {
  const out = new Uint8Array(size)
  fill(new DataView(out.buffer))
  return out
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) { out.set(p, at); at += p.length }
  return out
}

function zip64Extra(values: number[]): Uint8Array {
  return record(4 + values.length * 8, (v) => {
    v.setUint16(0, 0x0001, true)
    v.setUint16(2, values.length * 8, true)
    values.forEach((value, i) => v.setBigUint64(4 + i * 8, BigInt(value), true))
  })
}

export function toBytes(data: Uint8Array | string | undefined): Uint8Array {
  if (data === undefined) return new Uint8Array(0)
  return typeof data === 'string' ? encoder.encode(data) : data
}

export function buildZip(entries: BuildEntry[], options: BuildOptions = {}): ZipLayout {
  const w = new Writer()
  const zip64 = options.zip64 === true
  const layouts: EntryLayout[] = []
  const prepared = entries.map((entry) => {
    const data = toBytes(entry.data)
    const method = entry.method ?? 8
    const compressed = entry.raw ?? (method === 8 ? new Uint8Array(deflateRawSync(data)) : data)
    const descriptor = entry.dataDescriptor ?? true
    const flags = (entry.flags ?? 0) | (descriptor ? 0x08 : 0) | 0x800
    return {
      entry, method, compressed, descriptor, flags,
      crc: entry.crc ?? crc32(data),
      csize: entry.compressedSize ?? compressed.length,
      usize: entry.uncompressedSize ?? data.length,
      name: encoder.encode(entry.name),
      localName: encoder.encode(entry.localName ?? entry.name),
    }
  })
  for (const p of prepared) {
    const localExtra = p.entry.localExtra ?? (zip64 ? zip64Extra([p.usize, p.csize]) : new Uint8Array(0))
    const header = record(30, (v) => {
      v.setUint32(0, 0x04034b50, true)
      v.setUint16(4, zip64 ? 45 : 20, true)
      v.setUint16(6, p.flags, true)
      v.setUint16(8, p.method, true)
      v.setUint32(14, p.descriptor ? 0 : p.crc, true)
      v.setUint32(18, zip64 ? 0xffffffff : p.descriptor ? 0 : p.csize, true)
      v.setUint32(22, zip64 ? 0xffffffff : p.descriptor ? 0 : p.usize, true)
      v.setUint16(26, p.localName.length, true)
      v.setUint16(28, localExtra.length, true)
    })
    const localOffset = w.push(concat(header, p.localName, localExtra))
    const dataStart = w.push(p.compressed)
    if (p.descriptor) {
      w.push(zip64
        ? record(24, (v) => { v.setUint32(0, 0x08074b50, true); v.setUint32(4, p.crc, true); v.setBigUint64(8, BigInt(p.csize), true); v.setBigUint64(16, BigInt(p.usize), true) })
        : record(16, (v) => { v.setUint32(0, 0x08074b50, true); v.setUint32(4, p.crc, true); v.setUint32(8, p.csize, true); v.setUint32(12, p.usize, true) }))
    }
    layouts.push({ localOffset, dataStart, dataLength: p.compressed.length, centralOffset: 0 })
  }
  const cdOffset = w.length
  prepared.forEach((p, i) => {
    const extra = p.entry.centralExtra ?? (zip64 ? zip64Extra([p.usize, p.csize, layouts[i].localOffset]) : new Uint8Array(0))
    const header = record(46, (v) => {
      v.setUint32(0, 0x02014b50, true)
      v.setUint16(4, zip64 ? 45 : 20, true)
      v.setUint16(6, zip64 ? 45 : 20, true)
      v.setUint16(8, p.flags, true)
      v.setUint16(10, p.method, true)
      v.setUint32(16, p.crc, true)
      v.setUint32(20, zip64 ? 0xffffffff : p.csize, true)
      v.setUint32(24, zip64 ? 0xffffffff : p.usize, true)
      v.setUint16(28, p.name.length, true)
      v.setUint16(30, extra.length, true)
      v.setUint32(42, zip64 ? 0xffffffff : layouts[i].localOffset, true)
    })
    layouts[i].centralOffset = w.push(concat(header, p.name, extra))
  })
  const cdSize = w.length - cdOffset
  let zip64EocdOffset: number | undefined
  let locatorOffset: number | undefined
  if (zip64) {
    zip64EocdOffset = w.push(record(56, (v) => {
      v.setUint32(0, 0x06064b50, true)
      v.setBigUint64(4, 44n, true)
      v.setUint16(12, 45, true)
      v.setUint16(14, 45, true)
      v.setBigUint64(24, BigInt(entries.length), true)
      v.setBigUint64(32, BigInt(entries.length), true)
      v.setBigUint64(40, BigInt(cdSize), true)
      v.setBigUint64(48, BigInt(cdOffset), true)
    }))
    const z = zip64EocdOffset
    locatorOffset = w.push(record(20, (v) => {
      v.setUint32(0, 0x07064b50, true)
      v.setBigUint64(8, BigInt(z), true)
      v.setUint32(16, 1, true)
    }))
  }
  const comment = options.comment ?? new Uint8Array(0)
  const eocdOffset = w.push(record(22, (v) => {
    v.setUint32(0, 0x06054b50, true)
    v.setUint16(8, zip64 ? 0xffff : entries.length, true)
    v.setUint16(10, zip64 ? 0xffff : entries.length, true)
    v.setUint32(12, zip64 ? 0xffffffff : cdSize, true)
    v.setUint32(16, zip64 ? 0xffffffff : cdOffset, true)
    v.setUint16(20, comment.length, true)
  }))
  w.push(comment)
  if (options.trailer) w.push(options.trailer)
  return { bytes: w.bytes(), entries: layouts, cdOffset, cdSize, eocdOffset, zip64EocdOffset, locatorOffset }
}

/** Stored deflate block: BFINAL bit, BTYPE=00, LEN, NLEN, data. */
export function storedBlock(data: Uint8Array, final: boolean): Uint8Array {
  const out = new Uint8Array(5 + data.length)
  out[0] = final ? 1 : 0
  out[1] = data.length & 0xff
  out[2] = (data.length >>> 8) & 0xff
  out[3] = ~data.length & 0xff
  out[4] = (~data.length >>> 8) & 0xff
  out.set(data, 5)
  return out
}

export { concat as concatBytes }
