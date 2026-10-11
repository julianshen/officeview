/** S1.1 bounded ZIP reader (design note §2.5, §4.1; RED tests 1-8, 5a). */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { createDeflateRaw, deflateRawSync } from 'node:zlib'
import { describe, expect, it, vi } from 'vitest'
import {
  blobRangeSource, bytesRangeSource, openPart, readPartBytes, readZipInventory,
  type ExpandedBudget, type InflaterFactory, type InflaterPort, type RangeSource, type ZipEntryInfo, type ZipLimits,
} from '../src/core/zip-stream'
import { createPakoInflater } from '../src/core/inflater'
import { PackageError, type PackageErrorCode } from '../src/core/package-error'
import { crc32 } from '../src/core/crc32'
import { SOURCE_ERROR_CODES, type SourceErrorCode } from '../src/xlsx/view/contracts'
import { buildZip, concatBytes, getU32, setU16, setU32, setU64, storedBlock, type BuildEntry, type BuildOptions } from './helpers/zip-builder'
import { probeTermination } from './helpers/inflater-conformance'

const LARGE = process.env.IV_LARGE === '1'
const enc = new TextEncoder()
const MiB = 1024 * 1024

const LIMITS: ZipLimits = {
  maxEntries: 4096, maxCentralDirectoryBytes: 4 * MiB, maxNameBytes: 1024,
  maxPartExpandedBytes: 2048 * MiB, maxTotalExpandedBytes: 4096 * MiB, maxRatio: 200,
}

interface SpySource extends RangeSource { reads: Array<[number, number]>; bytesRead(): number }
function spy(inner: RangeSource): SpySource {
  const reads: Array<[number, number]> = []
  return {
    size: inner.size,
    reads,
    bytesRead: () => reads.reduce((n, [a, b]) => n + (b - a), 0),
    read(start, end, signal) { reads.push([start, end]); return inner.read(start, end, signal) },
  }
}

interface BudgetDouble extends ExpandedBudget { total: number; peakCharge: number; failed: boolean }
function budget(limit = Number.POSITIVE_INFINITY): BudgetDouble {
  const b: BudgetDouble = {
    total: 0, peakCharge: 0, failed: false,
    charge(bytes: number) {
      if (b.failed) throw new Error('charged after the budget failed')
      if (b.total + bytes > limit) { b.failed = true; throw new PackageError('resource-limit', `expanded budget ${limit} exceeded`) }
      b.total += bytes
      b.peakCharge = Math.max(b.peakCharge, bytes)
    },
  }
  return b
}

function opts(extra: Partial<Parameters<typeof openPart>[2]> = {}) {
  return { signal: new AbortController().signal, budget: budget(), ...extra }
}

async function collect(gen: AsyncGenerator<Uint8Array, void, void>): Promise<Uint8Array> {
  const parts: Uint8Array[] = []
  for await (const chunk of gen) parts.push(chunk.slice())
  return concatBytes(...parts)
}

async function codeOf(p: Promise<unknown>): Promise<PackageErrorCode | 'ok' | string> {
  try { await p; return 'ok' } catch (e) { return e instanceof PackageError ? e.code : `untyped: ${String(e)}` }
}

function inventory(bytes: Uint8Array, limits: ZipLimits = LIMITS) {
  return readZipInventory(bytesRangeSource(bytes), limits)
}

const text = (b: Uint8Array) => new TextDecoder().decode(b)

const SAMPLE: BuildEntry[] = [
  { name: '[Content_Types].xml', data: '<Types/>', method: 0 },
  { name: 'xl/workbook.xml', data: '<workbook>' + 'x'.repeat(5000) + '</workbook>' },
  { name: 'xl/worksheets/sheet1.xml', data: '<worksheet>' + '<row/>'.repeat(3000) + '</worksheet>' },
]

async function expectAllPartsVerify(bytes: Uint8Array): Promise<void> {
  const src = bytesRangeSource(bytes)
  const inv = await readZipInventory(src, LIMITS)
  for (const spec of SAMPLE) {
    const entry = inv.find(spec.name)
    expect(entry, spec.name).toBeDefined()
    expect(text(await collect(openPart(src, entry!, opts())))).toBe(spec.data)
  }
}

describe('error codes', () => {
  it('every PackageErrorCode is a frozen SOURCE_ERROR_CODES member', () => {
    const codes: PackageErrorCode[] = ['cancelled', 'resource-limit', 'unsupported', 'parse-error', 'storage-error', 'quota']
    // compile-time: PackageErrorCode must be assignable to SourceErrorCode
    const asSource: SourceErrorCode[] = codes
    for (const c of asSource) expect(SOURCE_ERROR_CODES).toContain(c)
  })
})

describe('ZIP64 and EOCD validation matrix (§2.5, RED test 2)', () => {
  const comment = new Uint8Array(65_535).fill(0x20)

  it.each<[string, BuildOptions]>([
    ['valid classic', {}],
    ['valid classic with a 65,535-byte comment', { comment }],
    ['valid ZIP64 (sentinels, ZIP64 extras, local extra of a different layout)', { zip64: true }],
    ['valid ZIP64 with a 65,535-byte comment', { zip64: true, comment }],
  ])('accepts %s and verifies every entry', async (_name, options) => {
    const layout = buildZip(SAMPLE, options)
    const inv = await inventory(layout.bytes)
    expect(inv.entries.map((e) => e.name)).toEqual(SAMPLE.map((e) => e.name))
    expect(inv.entries.every((e) => e.zip64 === (options.zip64 === true))).toBe(true)
    await expectAllPartsVerify(layout.bytes)
  })

  type Forge = (bytes: Uint8Array, l: ReturnType<typeof buildZip>) => void
  const zip64Cases: Array<[string, Forge, PackageErrorCode]> = [
    ['locator total disks = 2', (b, l) => setU32(b, l.locatorOffset! + 16, 2), 'unsupported'],
    ['locator total disks = 0', (b, l) => setU32(b, l.locatorOffset! + 16, 0), 'unsupported'],
    ['locator disk = 1', (b, l) => setU32(b, l.locatorOffset! + 4, 1), 'unsupported'],
    ['ZIP64 EOCD disk = 1', (b, l) => setU32(b, l.zip64EocdOffset! + 16, 1), 'unsupported'],
    ['ZIP64 EOCD cd-disk = 3', (b, l) => setU32(b, l.zip64EocdOffset! + 20, 3), 'unsupported'],
    ['classic diskNo = 1 next to sentinels', (b, l) => setU16(b, l.eocdOffset + 4, 1), 'unsupported'],
    ['classic cdDisk = 1 next to sentinels', (b, l) => setU16(b, l.eocdOffset + 6, 1), 'unsupported'],
    ['non-sentinel classic total that disagrees (5 vs 3)', (b, l) => { setU16(b, l.eocdOffset + 8, 5); setU16(b, l.eocdOffset + 10, 5) }, 'parse-error'],
    ['entries-this-disk != total', (b, l) => setU64(b, l.zip64EocdOffset! + 24, 2), 'parse-error'],
    ['record size 10', (b, l) => setU64(b, l.zip64EocdOffset! + 4, 10), 'parse-error'],
    ['record size 2^30', (b, l) => setU64(b, l.zip64EocdOffset! + 4, 2 ** 30), 'parse-error'],
    ['ZIP64 EOCD offset beyond the locator', (b, l) => setU64(b, l.locatorOffset! + 8, l.locatorOffset! + 4), 'parse-error'],
    ['central directory overlapping the ZIP64 EOCD', (b, l) => setU64(b, l.zip64EocdOffset! + 40, l.cdSize + 10), 'parse-error'],
    ['ZIP64 EOCD signature damaged', (b, l) => setU32(b, l.zip64EocdOffset!, 0x12345678), 'parse-error'],
    ['offset beyond 2^53', (b, l) => setU64(b, l.zip64EocdOffset! + 48, 2n ** 60n), 'parse-error'],
  ]
  it.each(zip64Cases)('ZIP64 forged: %s -> %s', async (_name, forge, code) => {
    const layout = buildZip(SAMPLE, { zip64: true })
    forge(layout.bytes, layout)
    expect(await codeOf(inventory(layout.bytes))).toBe(code)
  })

  it('accepts ZIP64 with a non-sentinel classic total that agrees', async () => {
    const layout = buildZip(SAMPLE, { zip64: true })
    setU16(layout.bytes, layout.eocdOffset + 8, 3)
    setU16(layout.bytes, layout.eocdOffset + 10, 3)
    expect((await inventory(layout.bytes)).entries).toHaveLength(3)
  })

  it('rejects ZIP64 sentinels without a locator', async () => {
    const layout = buildZip(SAMPLE, { zip64: true })
    setU32(layout.bytes, layout.locatorOffset!, 0)
    expect(await codeOf(inventory(layout.bytes))).toBe('parse-error')
  })

  const classicCases: Array<[string, Forge, PackageErrorCode]> = [
    ['disk 1', (b, l) => setU16(b, l.eocdOffset + 4, 1), 'unsupported'],
    ['cd-disk 1', (b, l) => setU16(b, l.eocdOffset + 6, 1), 'unsupported'],
    ['entries-this-disk != total', (b, l) => setU16(b, l.eocdOffset + 8, 2), 'parse-error'],
    ['total entries more than present', (b, l) => { setU16(b, l.eocdOffset + 8, 4); setU16(b, l.eocdOffset + 10, 4) }, 'parse-error'],
    ['total entries fewer than present', (b, l) => { setU16(b, l.eocdOffset + 8, 2); setU16(b, l.eocdOffset + 10, 2) }, 'parse-error'],
    ['cdOff beyond EOCD', (b, l) => setU32(b, l.eocdOffset + 16, l.eocdOffset + 1), 'parse-error'],
    ['central signature damaged', (b, l) => setU32(b, l.cdOffset, 0x01020304), 'parse-error'],
    ['entry disk start = 1', (b, l) => setU16(b, l.entries[1].centralOffset + 34, 1), 'unsupported'],
    ['encrypted entry (bit 0)', (b, l) => setU16(b, l.entries[1].centralOffset + 8, 0x0809), 'unsupported'],
    ['method 12 (bzip2)', (b, l) => setU16(b, l.entries[1].centralOffset + 10, 12), 'unsupported'],
    ['stored entry whose sizes differ', (b, l) => setU32(b, l.entries[0].centralOffset + 24, 9), 'parse-error'],
    ['local header offset inside the central directory', (b, l) => setU32(b, l.entries[2].centralOffset + 42, l.cdOffset + 4), 'parse-error'],
    ['overlapping entries (shared local header)', (b, l) => setU32(b, l.entries[2].centralOffset + 42, l.entries[1].localOffset), 'parse-error'],
  ]
  it.each(classicCases)('classic forged: %s -> %s', async (_name, forge, code) => {
    const layout = buildZip(SAMPLE)
    forge(layout.bytes, layout)
    expect(await codeOf(inventory(layout.bytes))).toBe(code)
  })

  it('rejects input with no end-of-central-directory record, or one that does not end the file', async () => {
    expect(await codeOf(inventory(new Uint8Array(10)))).toBe('parse-error')
    expect(await codeOf(inventory(new Uint8Array(4096)))).toBe('parse-error')
    const layout = buildZip(SAMPLE, { trailer: enc.encode('JUNK') })
    expect(await codeOf(inventory(layout.bytes))).toBe('parse-error')
  })

  it('an EOCD-looking record inside the comment does not hide the real one when it is implausible', async () => {
    const fake = new Uint8Array(22)
    setU32(fake, 0, 0x06054b50)
    setU16(fake, 8, 1)
    setU16(fake, 10, 1)
    setU32(fake, 12, 46)
    setU32(fake, 16, 999_999) // central directory beyond the file
    setU16(fake, 20, 0) // exact comment length: it is the last record
    const layout = buildZip(SAMPLE, { comment: fake })
    expect((await inventory(layout.bytes)).entries).toHaveLength(3)
  })

  it('accepts a local extra field that differs from the central one', async () => {
    const extra = new Uint8Array(40)
    setU16(extra, 0, 0xcafe)
    setU16(extra, 2, 36)
    const layout = buildZip(SAMPLE.map((e) => ({ ...e, localExtra: extra })))
    await expectAllPartsVerify(layout.bytes)
  })

  it('reads data-descriptor entries (bit 3, zero local sizes), stored ones included, from the central directory', async () => {
    const layout = buildZip(SAMPLE)
    for (const e of layout.entries) expect(getU32(layout.bytes, e.localOffset + 18)).toBe(0)
    const inv = await inventory(layout.bytes)
    expect(inv.entries.every((e) => (e.flags & 8) === 8)).toBe(true)
    await expectAllPartsVerify(layout.bytes)
  })

  it('reads the committed generator fixtures (stored + descriptors) with CRC and size verified', async () => {
    for (const file of ['f5-small.xlsx', 'f8-small.xlsx']) {
      const bytes = new Uint8Array(readFileSync(join(process.cwd(), 'tests/fixtures/xlsx-viewer', file)))
      const src = bytesRangeSource(bytes)
      const inv = await readZipInventory(src, LIMITS)
      expect(inv.find('xl/worksheets/sheet1.xml')).toBeDefined()
      for (const entry of inv.entries) {
        const data = await collect(openPart(src, entry, opts()))
        expect(data.length).toBe(entry.uncompressedSize)
      }
    }
  })
})

describe('names and limits (§2.5, RED test 2)', () => {
  it.each([
    ['exact duplicate', ['a.xml', 'a.xml']],
    ['duplicate [Content_Types].xml (disguised spreadsheet chain)', ['[Content_Types].xml', 'x.xml', '[Content_Types].xml']],
    ['ASCII case-insensitive duplicate (OPC part-name equivalence)', ['xl/Workbook.xml', 'xl/workbook.xml']],
    ['.. segment', ['xl/../evil.xml']],
    ['absolute name', ['/etc/passwd']],
    ['backslash', ['xl\\sheet.xml']],
    ['NUL byte', ['a\u0000.xml']],
    ['drive letter', ['C:/x.xml']],
    ['empty name', ['']],
  ])('rejects %s with parse-error', async (_name, names) => {
    const layout = buildZip(names.map((name, i) => ({ name, data: `part ${i}` })))
    const err = await readZipInventory(bytesRangeSource(layout.bytes), LIMITS).catch((e: unknown) => e)
    expect(err).toBeInstanceOf(PackageError)
    expect((err as PackageError).code).toBe('parse-error')
  })

  it('names the duplicate part', async () => {
    const layout = buildZip([{ name: '[Content_Types].xml', data: 'a' }, { name: '[Content_Types].xml', data: 'b' }])
    await expect(inventory(layout.bytes)).rejects.toThrow(/duplicate part name \[Content_Types\]\.xml/)
  })

  it('rejects names that are not valid UTF-8', async () => {
    const layout = buildZip([{ name: 'ab.xml', data: 'x' }])
    // overwrite the central name 'a' with 0xff
    layout.bytes[layout.entries[0].centralOffset + 46] = 0xff
    expect(await codeOf(inventory(layout.bytes))).toBe('parse-error')
  })

  it.each<[string, Partial<ZipLimits>, BuildEntry[]]>([
    ['entry count over limit', { maxEntries: 2 }, SAMPLE],
    ['central directory bytes over limit', { maxCentralDirectoryBytes: 100 }, SAMPLE],
    ['name bytes over limit', { maxNameBytes: 10 }, SAMPLE],
    ['declared part size over limit', { maxPartExpandedBytes: 1000 }, SAMPLE],
    ['declared total over limit', { maxTotalExpandedBytes: 10_000 }, SAMPLE],
    ['declared ratio over limit for a part above 1 MiB', { maxRatio: 200 }, [{ name: 'z.xml', data: new Uint8Array(4 * MiB) }]],
  ])('%s -> resource-limit', async (_name, limits, entries) => {
    const layout = buildZip(entries)
    expect(await codeOf(inventory(layout.bytes, { ...LIMITS, ...limits }))).toBe('resource-limit')
  })

  it('the ratio limit does not apply to parts at or below 1 MiB', async () => {
    const layout = buildZip([{ name: 'z.xml', data: new Uint8Array(MiB) }])
    expect((await inventory(layout.bytes, { ...LIMITS, maxRatio: 2 })).entries).toHaveLength(1)
  })
})

describe('range reads (RED test 1)', () => {
  it('reads only the tail, the central directory and each part range, never the whole file, over a jsdom Blob', async () => {
    expect((Blob.prototype as { stream?: unknown }).stream).toBeUndefined()
    // F2-shaped: one large deflated sheet plus small parts, data descriptors everywhere
    const rows: string[] = []
    for (let r = 1; r <= 40_000; r++) rows.push(`<row r="${r}"><c r="A${r}"><v>${(r * 7919) % 100_003}</v></c><c r="B${r}" t="s"><v>${r % 97}</v></c></row>`)
    const sheet = `<worksheet><sheetData>${rows.join('')}</sheetData></worksheet>`
    const layout = buildZip([...SAMPLE.slice(0, 2), { name: 'xl/worksheets/sheet1.xml', data: sheet }])
    const src = spy(blobRangeSource(new Blob([layout.bytes as BlobPart])))
    const inv = await readZipInventory(src, LIMITS)
    const rangeReadBytes = 262_144
    let expanded = 0
    for (const entry of inv.entries) {
      for await (const chunk of openPart(src, entry, opts({ rangeReadBytes }))) expanded += chunk.length
    }
    expect(expanded).toBe(inv.entries.reduce((n, e) => n + e.uncompressedSize, 0))
    const total = layout.bytes.length
    const largest = Math.max(...src.reads.map(([a, b]) => b - a))
    expect(src.reads.some(([a, b]) => a === 0 && b === total)).toBe(false)
    expect(largest).toBeLessThanOrEqual(Math.max(rangeReadBytes, layout.cdSize, 65_557))
    const compressed = inv.entries.reduce((n, e) => n + e.compressedSize, 0)
    const localHeaders = inv.entries.reduce((n, e) => n + 30 + enc.encode(e.name).length, 0)
    expect(src.bytesRead()).toBeLessThanOrEqual(compressed + Math.min(total, 65_557) + 20 + layout.cdSize + localHeaders)
  })

  it('range sources reject out-of-bounds reads', async () => {
    const src = bytesRangeSource(new Uint8Array(10))
    expect(await codeOf(src.read(5, 11))).toBe('parse-error')
    expect(await codeOf(src.read(-1, 2))).toBe('parse-error')
    expect(await codeOf(blobRangeSource(new Blob(['abc'])).read(2, 1))).toBe('parse-error')
    expect(text(await blobRangeSource(new Blob(['abcdef'])).read(2, 4))).toBe('cd')
  })
})

/** Wrap the pako port: record every step and make it async (§5.3 test 5 "fake async port"). */
function recordingFactory(log: { steps: Array<{ in: number; out: number; status: string }>; disposed: number; finished: Array<{ ended: boolean; totalIn: number }> }, asyncSteps = false): InflaterFactory {
  return () => {
    const inner = createPakoInflater()
    const port: InflaterPort = {
      step(input, inPos, out) {
        const s = inner.step(input, inPos, out) as ReturnType<InflaterPort['step']> & { status: string; inConsumed: number; outProduced: number }
        log.steps.push({ in: s.inConsumed, out: s.outProduced, status: s.status })
        return asyncSteps ? Promise.resolve(s) : s
      },
      finish() {
        const t = inner.finish() as { ended: boolean; totalIn: number; totalOut: number }
        log.finished.push({ ended: t.ended, totalIn: t.totalIn })
        return asyncSteps ? Promise.resolve(t) : t
      },
      dispose() { log.disposed++; return inner.dispose() },
    }
    return port
  }
}
const newLog = () => ({ steps: [] as Array<{ in: number; out: number; status: string }>, disposed: 0, finished: [] as Array<{ ended: boolean; totalIn: number }> })

async function openSingle(entry: BuildEntry, partOpts: Partial<Parameters<typeof openPart>[2]> = {}, factory?: InflaterFactory, limits: ZipLimits = LIMITS) {
  const layout = buildZip([entry])
  const src = spy(bytesRangeSource(layout.bytes))
  const inv = await readZipInventory(src, limits)
  const info = inv.entries[0]
  return { layout, src, info, run: () => collect(openPart(src, info, { ...opts(), ...partOpts }, factory)) }
}

describe('stream termination through openPart (RED test 5)', () => {
  const hello = new Uint8Array(deflateRawSync(enc.encode('hello world')))
  const cases: Array<[string, BuildEntry]> = [
    ['non-final stored block only (truncated)', { name: 'p', data: 'ABC', raw: storedBlock(enc.encode('ABC'), false) }],
    ['final block followed by trailing bytes inside the entry range', { name: 'p', data: 'hello world', raw: concatBytes(hello, enc.encode('GARBAGE')) }],
    ['declared compressed size larger than the stream (consumed < compressedSize)', { name: 'p', data: 'hello world', raw: concatBytes(hello, new Uint8Array(1)) }],
    ['declared compressed size smaller than the stream (consumed would exceed it)', { name: 'p', data: 'hello world', raw: hello.subarray(0, hello.length - 1) }],
    ['empty compressed data', { name: 'p', data: '', raw: new Uint8Array(0) }],
    ['truncated deflate', { name: 'p', data: 'hello world', raw: hello.subarray(0, 4) }],
    ['random garbage', { name: 'p', data: 'hello world', raw: Uint8Array.of(0x07, 0xff, 0x00, 0x13) }],
  ]
  for (const asyncSteps of [false, true]) {
    const label = asyncSteps ? 'fake async port' : 'pako port'
    it.each(cases)(`${label}: %s -> parse-error`, async (_name, entry) => {
      const { run } = await openSingle(entry, {}, recordingFactory(newLog(), asyncSteps))
      expect(await codeOf(run())).toBe('parse-error')
    })

    it(`${label}: a valid stream ends with totalIn === compressedSize`, async () => {
      const log = newLog()
      const { run, info } = await openSingle({ name: 'p', data: 'hello world', raw: hello }, {}, recordingFactory(log, asyncSteps))
      expect(text(await run())).toBe('hello world')
      expect(log.finished).toEqual([{ ended: true, totalIn: info.compressedSize }])
      expect(log.disposed).toBe(1)
    })

    it(`${label}: Z_BUF_ERROR is not corruption: 1-byte range reads and a 1-byte output buffer complete`, async () => {
      const data = 'The quick brown fox jumps over the lazy dog. '.repeat(50)
      const log = newLog()
      const { run } = await openSingle({ name: 'p', data }, { rangeReadBytes: 1, outputBufferBytes: 1 }, recordingFactory(log, asyncSteps))
      expect(text(await run())).toBe(data)
      expect(log.steps.some((s) => s.status === 'need-input')).toBe(true)
      expect(Math.max(...log.steps.map((s) => s.out))).toBe(1)
    })
  }

  it('the high-level pako.Inflate answer is not the oracle: it reports err 0 on a truncated stream that openPart rejects', async () => {
    const { Inflate } = await import('pako/lib/inflate.js')
    const raw = hello.subarray(0, 8)
    const hl = new Inflate({ raw: true })
    hl.push(raw, true)
    expect(hl.err).toBe(0)
    const { run } = await openSingle({ name: 'p', data: 'hello world', raw })
    expect(await codeOf(run())).toBe('parse-error')
  })
})

describe('native DecompressionStream exclusion (RED test 5a)', () => {
  it('the production path never constructs DecompressionStream', async () => {
    const ctor = vi.fn()
    const original = (globalThis as { DecompressionStream?: unknown }).DecompressionStream
    ;(globalThis as { DecompressionStream?: unknown }).DecompressionStream = ctor
    try {
      const { run } = await openSingle({ name: 'p', data: 'native must not be used '.repeat(100) })
      await run()
      expect(ctor).not.toHaveBeenCalled()
    } finally {
      ;(globalThis as { DecompressionStream?: unknown }).DecompressionStream = original
    }
  })

  it('the conformance probe passes for the pako port and fails for an engine that accepts trailing bytes', async () => {
    expect((await probeTermination(createPakoInflater)).pass).toBe(true)
    // Simulated Bun/Node-like engine: at the final block it swallows whatever input remains.
    const lenient: InflaterFactory = () => {
      const inner = createPakoInflater()
      let claimedIn = 0
      return {
        step(input, inPos, out) {
          const s = inner.step(input, inPos, out) as { status: 'end' | 'progress' | 'need-input' | 'need-output'; inConsumed: number; outProduced: number }
          const reported = s.status === 'end' ? { ...s, inConsumed: input.length - inPos } : s
          claimedIn += reported.inConsumed
          return reported
        },
        finish: () => ({ ...(inner.finish() as { ended: boolean; totalIn: number; totalOut: number }), totalIn: claimedIn }),
        dispose: () => inner.dispose(),
      }
    }
    const result = await probeTermination(lenient)
    expect(result.pass).toBe(false)
    expect(result.rows.filter((r) => r.accepted !== r.expectAccept).map((r) => r.name)).toEqual(['final stored block + 3 junk bytes', 'deflate("hello") + GARBAGE'])
  })
})

describe('declared-size and budget bounds (RED tests 3, 4, 6, 8)', () => {
  it('bogus declared size: 1 KiB declared, 10 MiB inflated, fails at the first byte beyond, before it is yielded', async () => {
    const big = new Uint8Array(10 * MiB).fill(0x41)
    const b = budget()
    const layout = buildZip([{ name: 'p', data: big }])
    // forge the central uncompressed size to 1 KiB
    const forged = layout.bytes.slice()
    setU32(forged, getCentral(layout) + 24, 1024)
    const src = bytesRangeSource(forged)
    const inv = await readZipInventory(src, { ...LIMITS, maxRatio: 100_000 })
    let yielded = 0
    const err = await (async () => { for await (const c of openPart(src, inv.entries[0], opts({ budget: b }))) yielded += c.length })().catch((e: unknown) => e)
    expect((err as PackageError).code).toBe('parse-error')
    expect(yielded).toBeLessThanOrEqual(1024)
    expect(b.total).toBeLessThanOrEqual(1024)
  })

  it('declared size larger than produced is a size mismatch parse-error', async () => {
    const { layout } = await openSingle({ name: 'p', data: 'hello world' })
    const forged = layout.bytes.slice()
    setU32(forged, getCentral(layout) + 24, 12)
    const src = bytesRangeSource(forged)
    const inv = await readZipInventory(src, LIMITS)
    expect(await codeOf(collect(openPart(src, inv.entries[0], opts())))).toBe('parse-error')
  })

  it(`limit crossing while decoding: ${LARGE ? '1 GiB' : '64 MiB'} of zeros under an 8 MiB budget stops at limit + one output buffer`, async () => {
    const size = LARGE ? 1024 * MiB : 64 * MiB
    const raw = await deflateZeros(size)
    const { src, info } = await openSingle({ name: 'bomb.xml', raw, data: new Uint8Array(0), uncompressedSize: size, crc: 0 }, {}, undefined, { ...LIMITS, maxRatio: 2000 })
    const b = budget(8 * MiB)
    const readsBefore = src.reads.length
    let yielded = 0
    let readsAtError = -1
    const err = await (async () => {
      for await (const c of openPart(src, info, opts({ budget: b, rangeReadBytes: 262_144 }))) yielded += c.length
    })().catch((e: unknown) => { readsAtError = src.reads.length; return e })
    expect((err as PackageError).code).toBe('resource-limit')
    expect(yielded).toBeLessThanOrEqual(8 * MiB + 65_536)
    expect(b.peakCharge).toBeLessThanOrEqual(65_536)
    // compressed data read: one 256 KiB range at most, although the entry is ~1 KiB per MiB of zeros
    const [header, ...data] = src.reads.slice(readsBefore)
    expect(header[0]).toBe(info.localHeaderOffset)
    expect(data.reduce((n, [a, e]) => n + (e - a), 0)).toBeLessThanOrEqual(262_144)
    expect(src.reads.length).toBe(readsAtError)
  })

  it('a port that produces more than the output buffer is rejected; the real port never exceeds it, whatever the input chunk size', async () => {
    const bursty: InflaterFactory = () => {
      const inner = createPakoInflater()
      return {
        step(input, inPos, out) {
          const s = inner.step(input, inPos, out) as { status: 'end' | 'progress' | 'need-input' | 'need-output'; inConsumed: number; outProduced: number }
          return { ...s, outProduced: s.outProduced > 0 ? out.length * 1029 : 0 }
        },
        finish: () => inner.finish(),
        dispose: () => inner.dispose(),
      }
    }
    const data = new Uint8Array(4 * MiB)
    const a = await openSingle({ name: 'p', data }, {}, bursty, { ...LIMITS, maxRatio: 100_000 })
    expect(await codeOf(a.run())).toBe('parse-error')
    for (const rangeReadBytes of [262_144, 4 * MiB]) {
      for (const outputBufferBytes of [65_536, 16_384]) {
        const log = newLog()
        const b = await openSingle({ name: 'p', data }, { rangeReadBytes, outputBufferBytes }, recordingFactory(log), { ...LIMITS, maxRatio: 100_000 })
        expect((await b.run()).length).toBe(data.length)
        expect(Math.max(...log.steps.map((s) => s.out))).toBe(outputBufferBytes)
      }
    }
  })

  it('a port that makes no progress is rejected instead of spinning', async () => {
    const stuck: InflaterFactory = () => ({
      step: () => ({ status: 'progress', inConsumed: 0, outProduced: 0 }),
      finish: () => ({ ended: false, totalIn: 0, totalOut: 0 }),
      dispose() {},
    })
    const { run } = await openSingle({ name: 'p', data: 'abc' }, {}, stuck)
    expect(await codeOf(run())).toBe('parse-error')
  })

  it('CRC mismatch is a parse-error by default and skipped only when verifyCrc is false', async () => {
    const entry: BuildEntry = { name: 'p', data: 'hello world', crc: 0x12345678 }
    expect(await codeOf((await openSingle(entry)).run())).toBe('parse-error')
    expect(text(await (await openSingle(entry, { verifyCrc: false })).run())).toBe('hello world')
    expect(await codeOf((await openSingle({ ...entry, method: 0 })).run())).toBe('parse-error')
  })

  it('a correct file verifies CRC, exact size and exact end of stream', async () => {
    const data = new Uint8Array(300_000).map((_, i) => (i * 31) & 0xff)
    const log = newLog()
    const { run, info } = await openSingle({ name: 'p', data }, {}, recordingFactory(log))
    const out = await run()
    expect(crc32(out)).toBe(info.crc32)
    expect(out.length).toBe(info.uncompressedSize)
    expect(log.finished[0]).toEqual({ ended: true, totalIn: info.compressedSize })
  })

  it('rejects a local header whose name differs from the central one, or whose data overruns the next entry', async () => {
    const a = buildZip([{ name: 'a.xml', localName: 'b.xml', data: 'x' }])
    const src = bytesRangeSource(a.bytes)
    expect(await codeOf(collect(openPart(src, (await readZipInventory(src, LIMITS)).entries[0], opts())))).toBe('parse-error')
    const extra = new Uint8Array(60)
    setU16(extra, 0, 0xcafe)
    setU16(extra, 2, 56)
    const b = buildZip([{ name: 'a.xml', data: 'x', localExtra: extra, dataDescriptor: false }, { name: 'b.xml', data: 'y' }])
    // b's central offset now points inside a's local extra: it passes the central-directory
    // span check (which cannot know a's local extra length) but a's data overruns it
    const a0 = b.entries[0]
    setU32(b.bytes, b.entries[1].centralOffset + 42, a0.localOffset + 30 + 5 + a0.dataLength + 10)
    const bsrc = bytesRangeSource(b.bytes)
    const inv = await readZipInventory(bsrc, LIMITS)
    const aInfo: ZipEntryInfo = inv.find('a.xml')!
    expect(aInfo.rangeEnd).toBe(a0.localOffset + 30 + 5 + a0.dataLength + 10)
    expect(await codeOf(collect(openPart(bsrc, aInfo, opts())))).toBe('parse-error')
  })

  it('rejects a local header whose method differs from the central one', async () => {
    const layout = buildZip([{ name: 'a.xml', data: 'x' }])
    setU16(layout.bytes, layout.entries[0].localOffset + 8, 0)
    const src = bytesRangeSource(layout.bytes)
    expect(await codeOf(collect(openPart(src, (await readZipInventory(src, LIMITS)).entries[0], opts())))).toBe('parse-error')
  })
})

function getCentral(layout: ReturnType<typeof buildZip>): number {
  return layout.entries[0].centralOffset
}

async function deflateZeros(size: number): Promise<Uint8Array> {
  if (size <= 64 * MiB) return new Uint8Array(deflateRawSync(new Uint8Array(size), { level: 9 }))
  const d = createDeflateRaw({ level: 9 })
  const parts: Uint8Array[] = []
  d.on('data', (c: Uint8Array) => parts.push(c))
  const block = new Uint8Array(4 * MiB)
  for (let done = 0; done < size; done += block.length) {
    if (!d.write(block)) await new Promise((r) => d.once('drain', r))
  }
  await new Promise((r) => d.end(r))
  return concatBytes(...parts)
}

/** Incompressible bytes, so a part spans many range reads. */
function noise(size: number, seed: number): Uint8Array {
  const out = new Uint8Array(size)
  let x = seed >>> 0
  for (let i = 0; i < size; i++) { x = (Math.imul(x, 1664525) + 1013904223) >>> 0; out[i] = x >>> 24 }
  return out
}

describe('cancellation (RED test 7)', () => {
  it('abort mid-part stops range reads, disposes the inflater and surfaces cancelled', async () => {
    const data = noise(3 * MiB, 7)
    const log = newLog()
    const ac = new AbortController()
    const { src, info } = await openSingle({ name: 'p', data }, {}, undefined, { ...LIMITS, maxRatio: 100_000 })
    const before = src.reads.length
    let pulls = 0
    let readsAtAbort = 0
    const err = await (async () => {
      for await (const c of openPart(src, info, { ...opts(), signal: ac.signal, rangeReadBytes: 65_536 }, recordingFactory(log))) {
        void c
        if (++pulls === 3) { ac.abort(); readsAtAbort = src.reads.length }
      }
    })().catch((e: unknown) => e)
    expect((err as PackageError).code).toBe('cancelled')
    expect(src.reads.length).toBe(readsAtAbort)
    expect(readsAtAbort).toBeGreaterThan(before)
    expect(log.disposed).toBe(1)
  })

  it('generator return() stops range reads and disposes the inflater', async () => {
    const data = noise(3 * MiB, 13)
    const log = newLog()
    const { src, info } = await openSingle({ name: 'p', data }, {}, undefined, { ...LIMITS, maxRatio: 100_000 })
    const gen = openPart(src, info, { ...opts(), rangeReadBytes: 65_536 }, recordingFactory(log))
    await gen.next()
    await gen.next()
    const reads = src.reads.length
    await gen.return()
    expect(log.disposed).toBe(1)
    expect((await gen.next()).done).toBe(true)
    expect(src.reads.length).toBe(reads)
  })

  it('an already-aborted signal fails the inventory read with cancelled', async () => {
    const ac = new AbortController()
    ac.abort()
    expect(await codeOf(readZipInventory(bytesRangeSource(buildZip(SAMPLE).bytes), LIMITS, ac.signal))).toBe('cancelled')
  })
})

describe('readPartBytes', () => {
  it('returns the whole small part and refuses a declared size above maxBytes before reading data', async () => {
    const layout = buildZip(SAMPLE)
    const src = spy(bytesRangeSource(layout.bytes))
    const inv = await readZipInventory(src, LIMITS)
    const wb = inv.find('xl/workbook.xml')!
    expect(text(await readPartBytes(src, wb, 1 << 20, opts()))).toBe(SAMPLE[1].data)
    const reads = src.reads.length
    expect(await codeOf(readPartBytes(src, wb, 100, opts()))).toBe('resource-limit')
    expect(src.reads.length).toBe(reads)
  })
})
