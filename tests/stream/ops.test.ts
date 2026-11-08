// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { InvalidOperand, Opcode, OpReader, OpWriter, Truncated, UnknownOpcode, type Op } from '../../src/stream/ops'
import { decodeUnit, encodeUnit } from '../../src/stream/unit'

const cases: Op[] = [
  { code: Opcode.STATE_PUSH }, { code: Opcode.STATE_POP },
  { code: Opcode.TRANSFORM, mode: 0, matrix: [-0, NaN, Infinity, -Infinity, Number.MIN_VALUE, 1] },
  { code: Opcode.TRANSFORM, mode: 1, matrix: [1, 0, 0, 1, 3, 4] },
  { code: Opcode.STYLE_COLOR, target: 2, color: '' },
  { code: Opcode.STYLE_COLOR, target: 0, color: '🎨red' },
  { code: Opcode.STYLE_OBJ, target: 1, assetId: 0xffffffff },
  { code: Opcode.STATE_SETUP, fields: {
    lineWidth: -0, lineCap: 'square', lineJoin: 'bevel', miterLimit: Infinity,
    lineDash: [0, Number.MIN_VALUE, 2], lineDashOffset: -Infinity, globalAlpha: NaN,
    globalCompositeOperation: 'source-over', imageSmoothingEnabled: false,
    imageSmoothingQuality: 'high', letterSpacing: '', wordSpacing: '2px', direction: 'rtl',
    lang: 'zh-Hant-😀', fontKerning: 'none', fontVariantCaps: 'small-caps', shadowBlur: 3,
    shadowOffsetX: -4, shadowOffsetY: 5, font: '12px "🦊"', textAlign: 'end',
    textBaseline: 'ideographic', fontStretch: 'expanded', textRendering: 'geometricPrecision',
  } },
  { code: Opcode.STATE_SETUP, fields: {} },
  ...([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10] as const).map(sub => ({
    code: Opcode.PATH as const, sub,
    args: Array.from({ length: [0, 2, 2, 4, 6, 5, 5, 7, 4, 4, 0][sub] }, (_, i) => i ? i : -0),
    ...(sub === 5 || sub === 7 ? { ccw: true } : {}),
    ...(sub === 9 ? { radii: [1, 2, 3, 4] } : {}),
  })),
  { code: Opcode.PATH_END, mode: 3, rule: 1 },
  { code: Opcode.RECT, mode: 2, rect: [NaN, Infinity, -Infinity, Number.MIN_VALUE] },
  { code: Opcode.GLYPH_RUN, mode: 0, text: 'A🦊', x: -0, y: 2, maxWidth: Infinity, advances: [1, NaN, Number.MIN_VALUE] },
  { code: Opcode.GLYPH_RUN, mode: 1, text: '', x: 0, y: 0, advances: [] },
  ...([3, 5, 9] as const).map(argc => ({ code: Opcode.IMAGE_DRAW as const, assetId: 2, args: Array.from({ length: argc - 1 }, (_, i) => i + .5) })),
  { code: Opcode.TILE, w: 1, h: 2, dx: -0, dy: Infinity, rgba: Uint8Array.of(1, 2, 3, 4, 5, 6, 7, 8) },
]
function write(op: Op) { const w = new OpWriter(); w.op(op.code, op as never); return w }

describe('Stage 1b op codec', () => {
  it.each(cases.map((op, i) => [i, op] as const))('round-trips case %s with IEEE-754 and UTF-8 fidelity', (_i, op) => {
    const w = write(op)
    expect(w.opCount).toBe(1)
    expect([...new OpReader(w.bytes())]).toEqual([op])
  })
  it.each(cases.map((op, i) => [i, op] as const))('rejects every incomplete byte prefix of case %s', (_i, op) => {
    const bytes = write(op).bytes()
    // Empty bytes are a complete empty stream. Every nonempty partial op must fail.
    for (let n = 1; n < bytes.length; n++) {
      const iterator = new OpReader(bytes.subarray(0, n))[Symbol.iterator]()
      expect(() => iterator.next(), `length ${n}/${bytes.length}`).toThrow(Truncated)
    }
  })
  it('writes LE f64 and u32 fields, not host-endian numbers', () => {
    const w = new OpWriter().op(Opcode.IMAGE_DRAW, { assetId: 0x12345678, args: [1, -0] })
    expect([...w.bytes().slice(0, 6)]).toEqual([11, 0x78, 0x56, 0x34, 0x12, 3])
    expect([...w.bytes().slice(6, 14)]).toEqual([0, 0, 0, 0, 0, 0, 0xf0, 0x3f])
  })
  it.each([0, 13, 255])('rejects unknown/reserved opcode %s', code => {
    expect(() => [...new OpReader(Uint8Array.of(code))]).toThrow(UnknownOpcode)
  })
  it('rejects a forged string length before decoding and invalid UTF-8', () => {
    expect(() => [...new OpReader(Uint8Array.of(4, 0, 255, 255, 255, 255))]).toThrow(Truncated)
    expect(() => [...new OpReader(Uint8Array.of(4, 0, 1, 0, 0, 0, 255))]).toThrow(InvalidOperand)
  })
  it('rejects unknown field-mask bits and enum values', () => {
    expect(() => [...new OpReader(Uint8Array.of(6, 0, 0, 0, 1))]).toThrow(InvalidOperand)
    expect(() => [...new OpReader(Uint8Array.of(4, 3, 0, 0, 0, 0))]).toThrow(InvalidOperand)
  })
  it('rejects oversized array counts before allocating', () => {
    const bytes = write({ code: Opcode.GLYPH_RUN, mode: 0, text: '', x: 0, y: 0, advances: [] }).bytes()
    new DataView(bytes.buffer).setUint32(bytes.length - 4, 0xffffffff, true)
    expect(() => [...new OpReader(bytes)]).toThrow(Truncated)
  })
  it('rejects a tile byte length inconsistent with its dimensions', () => {
    const bytes = write(cases.at(-1)!).bytes()
    new DataView(bytes.buffer).setUint32(1, 2, true)
    expect(() => [...new OpReader(bytes)]).toThrow(InvalidOperand)
  })
  it('rejects invalid writer input atomically', () => {
    const w = new OpWriter().op(Opcode.STATE_PUSH)
    const before = w.bytes()
    expect(() => w.op(Opcode.IMAGE_DRAW, { assetId: 1, args: [1] })).toThrow(InvalidOperand)
    expect(w.bytes()).toEqual(before)
    expect(w.opCount).toBe(1)
  })
})

describe('unit body', () => {
  it('derives command opCount and byteLen, then validates them on decode', () => {
    const body = new OpWriter().op(Opcode.STATE_PUSH).op(Opcode.STATE_POP).bytes()
    const bytes = encodeUnit({ unitId: 0xffffffff, discriminator: 1, body })
    expect([...bytes.slice(0, 13)]).toEqual([255, 255, 255, 255, 2, 0, 0, 0, 2, 0, 0, 0, 1])
    expect(decodeUnit(bytes)).toEqual({ unitId: 0xffffffff, byteLen: 2, opCount: 2, discriminator: 1, body })
    bytes[8] = 1
    expect(() => decodeUnit(bytes)).toThrow(InvalidOperand)
  })
  it.each([2, 3] as const)('round-trips opaque body discriminator %s with zero opCount', discriminator => {
    const body = Uint8Array.of(0, 13, 255)
    expect(decodeUnit(encodeUnit({ unitId: 7, discriminator, body }))).toEqual({ unitId: 7, byteLen: 3, opCount: 0, discriminator, body })
  })
  it('rejects every truncated unit prefix and trailing bytes', () => {
    const bytes = encodeUnit({ unitId: 7, discriminator: 2, body: Uint8Array.of(1, 2, 3) })
    for (let i = 0; i < bytes.length; i++) expect(() => decodeUnit(bytes.slice(0, i))).toThrow()
    expect(() => decodeUnit(Uint8Array.from([...bytes, 0]))).toThrow(InvalidOperand)
  })
  it('rejects invalid discriminator and byteLen mismatch', () => {
    const bytes = encodeUnit({ unitId: 1, discriminator: 1, body: new Uint8Array() })
    bytes[12] = 0
    expect(() => decodeUnit(bytes)).toThrow(InvalidOperand)
    bytes[12] = 1; bytes[4] = 1
    expect(() => decodeUnit(bytes)).toThrow(Truncated)
  })
})
