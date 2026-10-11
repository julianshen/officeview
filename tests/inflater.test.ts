/** S1.1 inflater port conformance (design note §2.4.1, RED tests 5, 6). */
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import pakoConstants from 'pako/lib/zlib/constants.js'
import { createPakoInflater, type InflaterPort } from '../src/core/inflater'
import { PackageError } from '../src/core/package-error'

const enc = new TextEncoder()

/** Stored block: BFINAL bit, BTYPE=00, LEN, NLEN, data. */
function storedBlock(data: Uint8Array, final: boolean): Uint8Array {
  const out = new Uint8Array(5 + data.length)
  out[0] = final ? 1 : 0
  out[1] = data.length & 0xff
  out[2] = (data.length >>> 8) & 0xff
  out[3] = ~data.length & 0xff
  out[4] = (~data.length >>> 8) & 0xff
  out.set(data, 5)
  return out
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) { out.set(p, at); at += p.length }
  return out
}

interface DriveResult { output: Uint8Array; ended: boolean; consumed: number; trailing: number; statuses: string[]; maxStep: number }

/** Drive a port over `input` the way openPart does, in `inChunk`-byte slices. */
async function drive(port: InflaterPort, input: Uint8Array, inChunk: number, outSize: number, lazyRead = false): Promise<DriveResult> {
  const out = new Uint8Array(outSize)
  const produced: number[] = []
  const statuses: string[] = []
  let readPos = 0
  let chunk: Uint8Array = new Uint8Array(0)
  let last = 'need-input'
  let pos = 0
  let consumed = 0
  let maxStep = 0
  let ended = false
  for (let guard = 0; guard < 10_000_000; guard++) {
    // lazyRead: only refill after the port asked for input (exercises need-input)
    if (pos === chunk.length && readPos < input.length && (!lazyRead || last === 'need-input')) {
      chunk = input.subarray(readPos, Math.min(input.length, readPos + inChunk))
      readPos += chunk.length
      pos = 0
    }
    const step = await port.step(chunk, pos, out)
    statuses.push(step.status)
    last = step.status
    pos += step.inConsumed
    consumed += step.inConsumed
    maxStep = Math.max(maxStep, step.outProduced)
    for (let i = 0; i < step.outProduced; i++) produced.push(out[i])
    if (step.status === 'end') { ended = true; break }
    if (step.status === 'need-input' && readPos >= input.length && pos === chunk.length) break
  }
  const totals = await port.finish()
  expect(totals.ended).toBe(ended)
  await port.dispose()
  return { output: Uint8Array.from(produced), ended, consumed, trailing: input.length - consumed, statuses, maxStep }
}

const text = (bytes: Uint8Array) => new TextDecoder().decode(bytes)

describe('pako low-level inflater port: termination table (§2.4.1)', () => {
  it('a valid stream ends with every compressed byte consumed', async () => {
    const data = enc.encode('hello world, '.repeat(200))
    const raw = deflateRawSync(data)
    const r = await drive(createPakoInflater(), raw, 262_144, 65_536)
    expect(r.ended).toBe(true)
    expect(r.consumed).toBe(raw.length)
    expect(text(r.output)).toBe(text(data))
  })

  it('a non-final stored block alone is truncated (no end), not success', async () => {
    const r = await drive(createPakoInflater(), storedBlock(enc.encode('ABC'), false), 262_144, 65_536)
    expect(r.ended).toBe(false)
    expect(text(r.output)).toBe('ABC')
  })

  it('a final stored block followed by 3 bytes ends with 3 trailing bytes', async () => {
    const r = await drive(createPakoInflater(), concat(storedBlock(enc.encode('ABC'), true), enc.encode('xyz')), 262_144, 65_536)
    expect(r.ended).toBe(true)
    expect(r.trailing).toBe(3)
  })

  it('deflate("hello") + GARBAGE ends with 7 trailing bytes', async () => {
    const r = await drive(createPakoInflater(), concat(deflateRawSync(enc.encode('hello')), enc.encode('GARBAGE')), 262_144, 65_536)
    expect(r.ended).toBe(true)
    expect(r.trailing).toBe(7)
    expect(text(r.output)).toBe('hello')
  })

  it('truncated deflate and empty input never end', async () => {
    const raw = deflateRawSync(enc.encode('hello world hello world'))
    expect((await drive(createPakoInflater(), raw.subarray(0, 8), 262_144, 65_536)).ended).toBe(false)
    expect((await drive(createPakoInflater(), new Uint8Array(0), 262_144, 65_536)).ended).toBe(false)
  })

  it('random garbage throws a typed parse-error (invalid block type)', async () => {
    // BFINAL=1, BTYPE=11 (reserved)
    const err = await drive(createPakoInflater(), Uint8Array.of(0x07, 0x12, 0x34), 262_144, 65_536).catch((e: unknown) => e)
    expect(err).toBeInstanceOf(PackageError)
    expect((err as PackageError).code).toBe('parse-error')
  })
})

describe('pako low-level inflater port: status mapping', () => {
  it('Z_BUF_ERROR with no input mid-stream is need-input, not corruption', async () => {
    const raw = deflateRawSync(enc.encode('abcdefgh'.repeat(1000)))
    const out = new Uint8Array(65_536)
    // feed half, then an empty step
    const half = raw.subarray(0, raw.length >> 1)
    const p2 = createPakoInflater()
    const a = await p2.step(half, 0, out)
    expect(a.inConsumed).toBe(half.length)
    const b = await p2.step(half, half.length, out)
    expect(b).toEqual({ status: 'need-input', inConsumed: 0, outProduced: 0 })
    const c = await p2.step(raw, half.length, out)
    expect(c.status).not.toBe('need-input')
    await p2.dispose()
  })

  it('a zero-length output buffer reports need-output, and empty steps after end repeat end', async () => {
    const raw = deflateRawSync(enc.encode('xyz'.repeat(100)))
    const port = createPakoInflater()
    const r0 = await port.step(raw, 0, new Uint8Array(0))
    expect(['need-output', 'progress']).toContain(r0.status)
    expect(r0.outProduced).toBe(0)
    let pos = r0.inConsumed
    const out = new Uint8Array(65_536)
    const r1 = await port.step(raw, pos, out)
    pos += r1.inConsumed
    expect(r1.status).toBe('end')
    expect(pos).toBe(raw.length)
    expect((await port.step(raw, pos, out)).status).toBe('end')
    expect(await port.finish()).toEqual({ ended: true, totalIn: raw.length, totalOut: 300 })
  })

  it('1-byte input and 1-byte output stepping completes with the right bytes', async () => {
    const data = enc.encode('The quick brown fox jumps over the lazy dog. '.repeat(40))
    const raw = deflateRawSync(data)
    const r = await drive(createPakoInflater(), raw, 1, 1, true)
    expect(r.ended).toBe(true)
    expect(r.consumed).toBe(raw.length)
    expect(text(r.output)).toBe(text(data))
    expect(r.statuses).toContain('need-input')
    expect(r.maxStep).toBe(1)
  })

  it('uses the zlib status constants of the pinned pako build', () => {
    expect(pakoConstants.Z_OK).toBe(0)
    expect(pakoConstants.Z_STREAM_END).toBe(1)
    expect(pakoConstants.Z_NEED_DICT).toBe(2)
    expect(pakoConstants.Z_STREAM_ERROR).toBe(-2)
    expect(pakoConstants.Z_DATA_ERROR).toBe(-3)
    expect(pakoConstants.Z_BUF_ERROR).toBe(-5)
    expect(pakoConstants.Z_NO_FLUSH).toBe(0)
  })
})

describe('pako low-level inflater port: contract checks', () => {
  it.each([-1, 4, 999, 1.5, Number.NaN])('an input position %s outside [0, input.length] is a typed parse-error', (inPos) => {
    const port = createPakoInflater()
    let err: unknown
    try { port.step(Uint8Array.of(1, 2, 3), inPos, new Uint8Array(16)) } catch (e) { err = e }
    expect(err).toBeInstanceOf(PackageError)
    expect((err as PackageError).code).toBe('parse-error')
    expect((err as PackageError).message).toMatch(/input position/)
    port.dispose()
  })
})

describe('pako low-level inflater port: structural output bound (RED test 6)', () => {
  it('one step never produces more than the caller-owned buffer, whatever the input slice', async () => {
    const zeros = deflateRawSync(new Uint8Array(16 * 1024 * 1024))
    for (const outSize of [65_536, 16_384]) {
      const port = createPakoInflater()
      const out = new Uint8Array(outSize)
      let pos = 0
      let max = 0
      let total = 0
      // whole compressed stream offered in one slice (far beyond 256 KiB-equivalent expansion)
      for (;;) {
        const step = await port.step(zeros, pos, out)
        pos += step.inConsumed
        max = Math.max(max, step.outProduced)
        total += step.outProduced
        if (step.status === 'end') break
      }
      expect(max).toBe(outSize)
      expect(total).toBe(16 * 1024 * 1024)
      await port.dispose()
    }
  })
})
