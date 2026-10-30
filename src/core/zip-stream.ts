/**
 * Bounded ZIP reading for large OPC packages (design note §2.5, §4.1).
 *
 * Only the end-of-central-directory tail, the ZIP64 locator/record, the
 * central directory and each requested part's own byte range are ever read,
 * through a lazy RangeSource (Blob.slice().arrayBuffer(); jsdom has no
 * Blob.stream). Parts are pulled through a caller-owned output buffer, so
 * one step never expands past it, and a part completes only if the inflater
 * reached Z_STREAM_END having consumed exactly the declared compressed size,
 * produced exactly the declared size, and (by default) matched the CRC.
 * No XLSX knowledge lives here; JSZip/OfficePackage are untouched.
 */
import { crc32Update } from './crc32'
import { defaultInflaterFactory, type InflaterFactory, type InflateStep } from './inflater'
import { PackageError, throwIfAborted, type PackageErrorCode, type PackageErrorReason } from './package-error'

export { PackageError, isPackageError, type PackageErrorCode } from './package-error'
export type { InflateStatus, InflateStep, InflateTotals, InflaterPort, InflaterFactory } from './inflater'

/** Lazy random-access bytes. */
export interface RangeSource { readonly size: number; read(start: number, end: number, signal?: AbortSignal): Promise<Uint8Array> }

export interface ZipLimits {
  maxEntries: number; maxCentralDirectoryBytes: number; maxNameBytes: number
  maxPartExpandedBytes: number; maxTotalExpandedBytes: number
  /** Declared uncompressed/compressed ratio, applied to parts above 1 MiB. */
  maxRatio: number
}
export interface ZipEntryInfo {
  name: string; method: 0 | 8; flags: number; crc32: number
  compressedSize: number; uncompressedSize: number; localHeaderOffset: number; zip64: boolean
  /** Upper bound for this entry's local header and data: the next local header or the central directory. */
  rangeEnd: number
}
export interface ZipInventory { readonly entries: readonly ZipEntryInfo[]; find(name: string): ZipEntryInfo | undefined }
/** Charged before expanded bytes are handed out; throws PackageError('resource-limit'). */
export interface ExpandedBudget { charge(bytes: number): void }
export interface PartReadOptions {
  signal: AbortSignal; budget: ExpandedBudget; verifyCrc?: boolean /* default true */
  rangeReadBytes?: number /* default 262144 (spec 256 KiB input chunk) */
  outputBufferBytes?: number /* default 65536, max 262144; bounds expansion per step */
  /** Internal prefix mode: stop after this many expanded bytes when the part is larger. */
  stopAfterOutputBytes?: number
}

const SIG_LOCAL = 0x04034b50
const SIG_CENTRAL = 0x02014b50
const SIG_EOCD = 0x06054b50
const SIG_ZIP64_EOCD = 0x06064b50
const SIG_ZIP64_LOCATOR = 0x07064b50
const EOCD_BYTES = 22
const MAX_COMMENT = 65_535
const LOCATOR_BYTES = 20
const ZIP64_EOCD_BYTES = 56
const CENTRAL_BYTES = 46
const LOCAL_BYTES = 30
const RATIO_FLOOR = 1024 * 1024
const DEFAULT_RANGE_READ = 262_144
const DEFAULT_OUTPUT = 65_536
const MAX_OUTPUT = 262_144
// encrypted (bit 0), strong encryption (bit 6), encrypted central directory (bit 13)
const ENCRYPTION_FLAGS = 0x0001 | 0x0040 | 0x2000

const nameDecoder = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true })
const nameEncoder = new TextEncoder()
const EMPTY = new Uint8Array(0)

function fail(code: PackageErrorCode, message: string, partId?: string, reason?: PackageErrorReason): never {
  throw new PackageError(code, message, partId, reason)
}

function u16(b: Uint8Array, at: number): number {
  return b[at] | (b[at + 1] << 8)
}
function u32(b: Uint8Array, at: number): number {
  return (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16)) + b[at + 3] * 0x1000000
}
/** ZIP64 fields are 64-bit; anything at or beyond 2^53 cannot be a real offset or size here. */
function u64(b: Uint8Array, at: number): number {
  const hi = u32(b, at + 4)
  if (hi > 0x1fffff) fail('parse-error', 'ZIP64 value beyond 2^53')
  return hi * 0x100000000 + u32(b, at)
}

function checkRange(start: number, end: number, size: number): void {
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > size) {
    fail('parse-error', `read [${start}, ${end}) outside a ${size}-byte source`)
  }
}

/** Blob adapter: slice(a, b).arrayBuffer() per range, never Blob.stream() or a whole-file read. */
export function blobRangeSource(blob: Blob): RangeSource {
  const size = blob.size
  return {
    size,
    async read(start, end, signal) {
      checkRange(start, end, size)
      throwIfAborted(signal)
      const buffer = await blob.slice(start, end).arrayBuffer()
      throwIfAborted(signal)
      if (buffer.byteLength !== end - start) fail('storage-error', `short read: ${buffer.byteLength} of ${end - start} bytes`)
      return new Uint8Array(buffer)
    },
  }
}

/** Already-in-memory input; reads are views, no copy. */
export function bytesRangeSource(bytes: Uint8Array): RangeSource {
  const size = bytes.length
  return {
    size,
    async read(start, end, signal) {
      checkRange(start, end, size)
      throwIfAborted(signal)
      return bytes.subarray(start, end)
    },
  }
}

/** Reads inside the already-fetched tail are served from it. */
class TailReader {
  constructor(private readonly src: RangeSource, private readonly tail: Uint8Array, private readonly tailStart: number, private readonly signal?: AbortSignal) {}
  async read(start: number, end: number): Promise<Uint8Array> {
    checkRange(start, end, this.src.size)
    if (start >= this.tailStart) return this.tail.subarray(start - this.tailStart, end - this.tailStart)
    const bytes = await this.src.read(start, end, this.signal)
    throwIfAborted(this.signal)
    return bytes
  }
}

/** EOCD tail + ZIP64 locator (read separately) + central directory only. */
export async function readZipInventory(src: RangeSource, limits: ZipLimits, signal?: AbortSignal): Promise<ZipInventory> {
  throwIfAborted(signal)
  const size = src.size
  if (size < EOCD_BYTES) fail('parse-error', 'not a ZIP archive: too short')
  const tailStart = Math.max(0, size - (EOCD_BYTES + MAX_COMMENT))
  const tail = await src.read(tailStart, size, signal)
  throwIfAborted(signal)
  // Backward scan; the comment length must end the record exactly at EOF.
  const candidates: number[] = []
  for (let i = tail.length - EOCD_BYTES; i >= 0; i--) {
    if (u32(tail, i) === SIG_EOCD && u16(tail, i + 20) === tail.length - i - EOCD_BYTES) candidates.push(i)
  }
  if (candidates.length === 0) fail('parse-error', 'end of central directory record not found')
  // An EOCD-like record inside a comment can match too (a parser differential no
  // reader fully closes); take the last candidate whose directory is plausible.
  const reader = new TailReader(src, tail, tailStart, signal)
  let firstError: unknown
  for (const i of candidates) {
    try {
      return await parseFromEocd(reader, tail.subarray(i, i + EOCD_BYTES), tailStart + i, limits)
    } catch (error) {
      if (!(error instanceof PackageError) || error.code === 'cancelled') throw error
      firstError ??= error
    }
  }
  throw firstError
}

async function parseFromEocd(reader: TailReader, eocd: Uint8Array, eocdPos: number, limits: ZipLimits): Promise<ZipInventory> {
  const diskNo = u16(eocd, 4)
  const cdDisk = u16(eocd, 6)
  const entriesThisDisk = u16(eocd, 8)
  const entriesTotal = u16(eocd, 10)
  const cdSize32 = u32(eocd, 12)
  const cdOffset32 = u32(eocd, 16)
  const sentinel = diskNo === 0xffff || cdDisk === 0xffff || entriesThisDisk === 0xffff || entriesTotal === 0xffff
    || cdSize32 === 0xffffffff || cdOffset32 === 0xffffffff

  let locator: Uint8Array | undefined
  const locatorPos = eocdPos - LOCATOR_BYTES
  if (locatorPos >= 0) {
    const bytes = await reader.read(locatorPos, eocdPos)
    if (u32(bytes, 0) === SIG_ZIP64_LOCATOR) locator = bytes
  }

  let total: number
  let cdSize: number
  let cdOffset: number
  let cdEndLimit: number
  if (locator !== undefined) {
    if (u32(locator, 4) !== 0) fail('unsupported', 'multi-disk archive: ZIP64 end record on another disk', undefined, 'multi-disk')
    if (u32(locator, 16) !== 1) fail('unsupported', `multi-disk archive: ZIP64 locator declares ${u32(locator, 16)} disks`, undefined, 'multi-disk')
    if ((diskNo !== 0 && diskNo !== 0xffff) || (cdDisk !== 0 && cdDisk !== 0xffff)) fail('unsupported', 'multi-disk archive', undefined, 'multi-disk')
    const z64Offset = u64(locator, 8)
    if (z64Offset + ZIP64_EOCD_BYTES > locatorPos) fail('parse-error', 'ZIP64 end record overlaps its locator')
    const z = await reader.read(z64Offset, z64Offset + ZIP64_EOCD_BYTES)
    if (u32(z, 0) !== SIG_ZIP64_EOCD) fail('parse-error', 'ZIP64 end record signature not found')
    const recordSize = u64(z, 4)
    if (recordSize < ZIP64_EOCD_BYTES - 12 || z64Offset + 12 + recordSize > locatorPos) fail('parse-error', `ZIP64 end record size ${recordSize} out of bounds`)
    if (u32(z, 16) !== 0 || u32(z, 20) !== 0) fail('unsupported', 'multi-disk archive: ZIP64 disk numbers', undefined, 'multi-disk')
    const entries64ThisDisk = u64(z, 24)
    total = u64(z, 32)
    if (entries64ThisDisk !== total) fail('parse-error', 'ZIP64 entries on this disk differ from the total')
    cdSize = u64(z, 40)
    cdOffset = u64(z, 48)
    if ((entriesThisDisk !== 0xffff && entriesThisDisk !== total) || (entriesTotal !== 0xffff && entriesTotal !== total)
      || (cdSize32 !== 0xffffffff && cdSize32 !== cdSize) || (cdOffset32 !== 0xffffffff && cdOffset32 !== cdOffset)) {
      fail('parse-error', 'classic and ZIP64 end records disagree')
    }
    cdEndLimit = z64Offset
  } else {
    if (sentinel) fail('parse-error', 'ZIP64 sentinels without a ZIP64 locator')
    if (diskNo !== 0 || cdDisk !== 0) fail('unsupported', 'multi-disk archive', undefined, 'multi-disk')
    if (entriesThisDisk !== entriesTotal) fail('parse-error', 'entries on this disk differ from the total')
    total = entriesTotal
    cdSize = cdSize32
    cdOffset = cdOffset32
    cdEndLimit = eocdPos
  }
  if (cdOffset + cdSize > cdEndLimit) fail('parse-error', 'central directory overlaps the end records')
  if (total > limits.maxEntries) fail('resource-limit', `${total} entries exceed the limit of ${limits.maxEntries}`, undefined, 'entry-limit')
  if (cdSize > limits.maxCentralDirectoryBytes) fail('resource-limit', `central directory of ${cdSize} bytes exceeds the limit of ${limits.maxCentralDirectoryBytes}`, undefined, 'central-directory-limit')
  if (cdSize < total * CENTRAL_BYTES) fail('parse-error', 'central directory too small for its entry count')
  const cd = await reader.read(cdOffset, cdOffset + cdSize)
  return parseCentralDirectory(cd, total, cdOffset, limits)
}

function parseCentralDirectory(cd: Uint8Array, total: number, cdOffset: number, limits: ZipLimits): ZipInventory {
  const entries: ZipEntryInfo[] = []
  const nameLengths: number[] = []
  const byName = new Map<string, ZipEntryInfo>()
  const folded = new Set<string>()
  let declaredTotal = 0
  let p = 0
  for (let k = 0; k < total; k++) {
    if (p + CENTRAL_BYTES > cd.length) fail('parse-error', 'central directory truncated')
    if (u32(cd, p) !== SIG_CENTRAL) fail('parse-error', 'central directory entry signature not found')
    const flags = u16(cd, p + 8)
    const method = u16(cd, p + 10)
    const crc = u32(cd, p + 16)
    let compressedSize = u32(cd, p + 20)
    let uncompressedSize = u32(cd, p + 24)
    const nameLength = u16(cd, p + 28)
    const extraLength = u16(cd, p + 30)
    const commentLength = u16(cd, p + 32)
    let diskStart = u16(cd, p + 34)
    let localHeaderOffset = u32(cd, p + 42)
    const next = p + CENTRAL_BYTES + nameLength + extraLength + commentLength
    if (next > cd.length) fail('parse-error', 'central directory entry overruns the directory')
    if (nameLength > limits.maxNameBytes) fail('resource-limit', `part name of ${nameLength} bytes exceeds the limit of ${limits.maxNameBytes}`, undefined, 'name-limit')
    let name: string
    try {
      name = nameDecoder.decode(cd.subarray(p + CENTRAL_BYTES, p + CENTRAL_BYTES + nameLength))
    } catch {
      fail('parse-error', 'part name is not valid UTF-8', undefined, 'name-encoding')
    }
    validateName(name)
    let zip64 = false
    if (uncompressedSize === 0xffffffff || compressedSize === 0xffffffff || localHeaderOffset === 0xffffffff || diskStart === 0xffff) {
      zip64 = true
      const extra = cd.subarray(p + CENTRAL_BYTES + nameLength, p + CENTRAL_BYTES + nameLength + extraLength)
      const values = zip64Extra(extra, name)
      // only the fields whose central value is a sentinel, in this fixed order
      if (uncompressedSize === 0xffffffff) uncompressedSize = values.u64()
      if (compressedSize === 0xffffffff) compressedSize = values.u64()
      if (localHeaderOffset === 0xffffffff) localHeaderOffset = values.u64()
      if (diskStart === 0xffff) diskStart = values.u32()
    }
    if ((flags & ENCRYPTION_FLAGS) !== 0) fail('unsupported', 'encrypted entry', name, 'encrypted')
    if (method !== 0 && method !== 8) fail('unsupported', `compression method ${method}`, name, 'method')
    if (diskStart !== 0) fail('unsupported', 'multi-disk archive: entry on another disk', name, 'multi-disk')
    if (method === 0 && compressedSize !== uncompressedSize) fail('parse-error', 'stored entry sizes differ', name)
    if (localHeaderOffset + LOCAL_BYTES + nameLength + compressedSize > cdOffset) fail('parse-error', 'entry data outside the data area', name)
    if (uncompressedSize > limits.maxPartExpandedBytes) fail('resource-limit', `declared size ${uncompressedSize} exceeds the part limit of ${limits.maxPartExpandedBytes}`, name)
    declaredTotal += uncompressedSize
    if (declaredTotal > limits.maxTotalExpandedBytes) fail('resource-limit', `declared expanded total exceeds the limit of ${limits.maxTotalExpandedBytes}`, name)
    if (uncompressedSize > RATIO_FLOOR && uncompressedSize > limits.maxRatio * compressedSize) {
      fail('resource-limit', `declared ratio exceeds ${limits.maxRatio}:1`, name)
    }
    // OPC part names compare ASCII case-insensitively (ECMA-376 Part 2 §9.1.1.1)
    const key = name.replace(/[A-Z]/g, (c) => c.toLowerCase())
    if (byName.has(name)) fail('parse-error', `duplicate part name ${name}`, undefined, 'duplicate-name')
    if (folded.has(key)) fail('parse-error', `duplicate part name ${name}`, undefined, 'case-duplicate-name')
    folded.add(key)
    const entry: ZipEntryInfo = {
      name, method, flags, crc32: crc, compressedSize, uncompressedSize, localHeaderOffset, zip64, rangeEnd: cdOffset,
    }
    byName.set(name, entry)
    entries.push(entry)
    nameLengths.push(nameLength)
    p = next
  }
  if (p !== cd.length) fail('parse-error', 'central directory size does not match its entries')
  // Entries must not overlap (overlapping-file bombs reuse one compressed stream).
  const order = entries.map((_, i) => i).sort((a, b) => entries[a].localHeaderOffset - entries[b].localHeaderOffset)
  for (let i = 0; i < order.length; i++) {
    const entry = entries[order[i]]
    const nextStart = i + 1 < order.length ? entries[order[i + 1]].localHeaderOffset : cdOffset
    if (entry.localHeaderOffset + LOCAL_BYTES + nameLengths[order[i]] + entry.compressedSize > nextStart) {
      fail('parse-error', 'entries overlap', entry.name)
    }
    entry.rangeEnd = nextStart
  }
  return { entries, find: (name) => byName.get(name) }
}

function validateName(name: string): void {
  if (name.length === 0) fail('parse-error', 'empty part name', undefined, 'name-invalid')
  if (name.includes('\u0000') || name.includes('\\') || name.startsWith('/') || /^[A-Za-z]:/.test(name)
    || name.split('/').includes('..')) {
    fail('parse-error', `escaping or invalid part name ${JSON.stringify(name)}`, undefined, 'name-invalid')
  }
}

function zip64Extra(extra: Uint8Array, name: string): { u64(): number; u32(): number } {
  let q = 0
  while (q + 4 <= extra.length) {
    const id = u16(extra, q)
    const length = u16(extra, q + 2)
    if (q + 4 + length > extra.length) fail('parse-error', 'extra field overruns', name)
    if (id === 0x0001) {
      let at = q + 4
      const end = at + length
      return {
        u64() {
          if (at + 8 > end) fail('parse-error', 'ZIP64 extra field too short', name)
          const v = u64(extra, at)
          at += 8
          return v
        },
        u32() {
          if (at + 4 > end) fail('parse-error', 'ZIP64 extra field too short', name)
          const v = u32(extra, at)
          at += 4
          return v
        },
      }
    }
    q += 4 + length
  }
  fail('parse-error', 'ZIP64 extra field missing', name)
}

function checkBufferSize(value: number, max: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > max) throw new RangeError(`${label} must be an integer in [1, ${max}]`)
  return value
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

function checkStep(step: InflateStep, inPos: number, inputLength: number, outLength: number, part: string): void {
  const okIn = Number.isSafeInteger(step.inConsumed) && step.inConsumed >= 0 && inPos + step.inConsumed <= inputLength
  const okOut = Number.isSafeInteger(step.outProduced) && step.outProduced >= 0 && step.outProduced <= outLength
  if (!okIn || !okOut) fail('parse-error', `inflater broke the port contract (in ${step.inConsumed}, out ${step.outProduced} of ${outLength})`, part)
}

/**
 * Pull model. Yields subarrays of ONE reused output buffer, valid until the
 * next pull (copy to retain). Reads only the local header and
 * [dataStart, dataStart + compressedSize). Every byte is checked against the
 * declared size and charged to the budget before it is yielded.
 */
export async function* openPart(src: RangeSource, entry: ZipEntryInfo, opts: PartReadOptions, makeInflater: InflaterFactory = defaultInflaterFactory): AsyncGenerator<Uint8Array, void, void> {
  const part = entry.name
  const signal = opts.signal
  const rangeReadBytes = checkBufferSize(opts.rangeReadBytes ?? DEFAULT_RANGE_READ, Number.MAX_SAFE_INTEGER, 'rangeReadBytes')
  const outputBufferBytes = checkBufferSize(opts.outputBufferBytes ?? DEFAULT_OUTPUT, MAX_OUTPUT, 'outputBufferBytes')
  const stopAfter = opts.stopAfterOutputBytes
  if (stopAfter !== undefined) checkBufferSize(stopAfter, Number.MAX_SAFE_INTEGER, 'stopAfterOutputBytes')
  const prefixMode = stopAfter !== undefined && entry.uncompressedSize > stopAfter
  const out = new Uint8Array(Math.min(outputBufferBytes, prefixMode ? stopAfter! : outputBufferBytes))
  const verifyCrc = opts.verifyCrc !== false
  throwIfAborted(signal, part)

  const nameBytes = nameEncoder.encode(entry.name)
  const headerEnd = entry.localHeaderOffset + LOCAL_BYTES + nameBytes.length
  if (headerEnd > entry.rangeEnd) fail('parse-error', 'local header overruns its range', part)
  const header = await src.read(entry.localHeaderOffset, headerEnd, signal)
  throwIfAborted(signal, part)
  if (u32(header, 0) !== SIG_LOCAL) fail('parse-error', 'local header signature not found', part)
  if ((u16(header, 6) & ENCRYPTION_FLAGS) !== 0) fail('unsupported', 'encrypted entry', part, 'encrypted')
  if (u16(header, 8) !== entry.method) fail('parse-error', 'local and central compression methods differ', part)
  if (u16(header, 26) !== nameBytes.length || !sameBytes(header.subarray(LOCAL_BYTES), nameBytes)) {
    fail('parse-error', 'local and central part names differ', part)
  }
  // Data start from the LOCAL header (its extra may differ); sizes from the central directory.
  const dataStart = headerEnd + u16(header, 28)
  const dataEnd = dataStart + entry.compressedSize
  if (dataEnd > entry.rangeEnd || dataEnd > src.size) fail('parse-error', 'entry data overruns its range', part)

  let produced = 0
  let crc = 0
  const accept = (n: number): void => {
    if (produced + n > entry.uncompressedSize) fail('parse-error', `part inflates beyond its declared size of ${entry.uncompressedSize} bytes`, part, 'size-mismatch')
    opts.budget.charge(n)
    if (verifyCrc) crc = crc32Update(crc, out, 0, n)
    produced += n
  }

  if (entry.method === 0) {
    for (let pos = dataStart; pos < dataEnd;) {
      throwIfAborted(signal, part)
      const end = Math.min(dataEnd, pos + rangeReadBytes)
      const chunk = await src.read(pos, end, signal)
      throwIfAborted(signal, part)
      pos = end
      for (let at = 0; at < chunk.length; at += out.length) {
        const remaining = prefixMode ? stopAfter! - produced : out.length
        const n = Math.min(out.length, chunk.length - at, remaining)
        out.set(chunk.subarray(at, at + n))
        accept(n)
        yield out.subarray(0, n)
        if (prefixMode && produced >= stopAfter!) return
        throwIfAborted(signal, part)
      }
    }
  } else {
    const inflater = makeInflater()
    try {
      let input: Uint8Array = EMPTY
      let inPos = 0
      let readPos = dataStart
      let consumed = 0
      // The next range is read only when the inflater asks for input, so
      // need-input (Z_BUF_ERROR, avail_in 0) is the one refill signal.
      let wantsInput = true
      for (;;) {
        throwIfAborted(signal, part)
        if (wantsInput && inPos === input.length && readPos < dataEnd) {
          const end = Math.min(dataEnd, readPos + rangeReadBytes)
          input = await src.read(readPos, end, signal)
          throwIfAborted(signal, part)
          readPos = end
          inPos = 0
        }
        const available = prefixMode ? Math.min(out.length, stopAfter! - produced) : out.length
        const step = await inflater.step(input, inPos, out.subarray(0, available))
        throwIfAborted(signal, part)
        checkStep(step, inPos, input.length, out.length, part)
        inPos += step.inConsumed
        consumed += step.inConsumed
        if (step.outProduced > 0) {
          accept(step.outProduced)
          yield out.subarray(0, step.outProduced)
          if (prefixMode && produced >= stopAfter!) return
        }
        if (step.status === 'end') break
        wantsInput = step.status === 'need-input'
        if (wantsInput) {
          if (inPos < input.length) fail('parse-error', 'inflater asked for input while input was pending', part)
          // Truncation is decided only here, at the end of the entry's compressed range.
          if (readPos >= dataEnd) fail('parse-error', 'compressed data ends before the deflate stream does', part, 'truncated-deflate')
        } else if (step.inConsumed === 0 && step.outProduced === 0) {
          fail('parse-error', 'inflater made no progress', part)
        }
      }
      if (inPos < input.length || readPos < dataEnd) fail('parse-error', 'trailing data after the end of the deflate stream', part)
      const totals = await inflater.finish()
      if (!totals.ended || consumed !== entry.compressedSize || totals.totalIn !== entry.compressedSize) {
        fail('parse-error', `deflate stream consumed ${consumed} of ${entry.compressedSize} declared bytes`, part, 'truncated-deflate')
      }
    } finally {
      await inflater.dispose()
    }
  }
  if (produced !== entry.uncompressedSize) fail('parse-error', `part produced ${produced} of ${entry.uncompressedSize} declared bytes`, part, 'size-mismatch')
  if (verifyCrc && crc !== entry.crc32) fail('parse-error', 'CRC-32 mismatch', part)
}

/** For small parts (rels, workbook, content types): bounded, never cached by this module. */
export async function readPartBytes(src: RangeSource, entry: ZipEntryInfo, maxBytes: number, opts: PartReadOptions): Promise<Uint8Array> {
  if (entry.uncompressedSize > maxBytes) fail('resource-limit', `declared size ${entry.uncompressedSize} exceeds ${maxBytes} bytes`, entry.name)
  const result = new Uint8Array(entry.uncompressedSize)
  let at = 0
  for await (const chunk of openPart(src, entry, opts)) {
    result.set(chunk, at)
    at += chunk.length
  }
  return result
}

/** Inflate at most maxBytes and stop immediately when the prefix is complete. */
export async function readPartPrefixBytes(src: RangeSource, entry: ZipEntryInfo, maxBytes: number, opts: PartReadOptions): Promise<Uint8Array> {
  checkBufferSize(maxBytes, Number.MAX_SAFE_INTEGER, 'maxBytes')
  if (entry.uncompressedSize <= maxBytes) return readPartBytes(src, entry, maxBytes, opts)
  const parts: Uint8Array[] = []
  let length = 0
  for await (const chunk of openPart(src, entry, { ...opts, stopAfterOutputBytes: maxBytes })) {
    parts.push(chunk.slice())
    length += chunk.length
  }
  const out = new Uint8Array(length)
  let offset = 0
  for (const part of parts) { out.set(part, offset); offset += part.length }
  return out
}
