/** Stage 1b op grammar. All numeric fields are little-endian; opcode is u8. */
export enum Opcode {
  STATE_PUSH = 1, STATE_POP, TRANSFORM, STYLE_COLOR, STYLE_OBJ, STATE_SETUP,
  PATH, PATH_END, RECT, GLYPH_RUN, IMAGE_DRAW, TILE, RASTER_BLOCK,
}
export class UnknownOpcode extends Error {
  constructor(readonly opcode: number) { super(`Unknown opcode ${opcode}`); this.name = 'UnknownOpcode' }
}
export class Truncated extends Error {
  constructor() { super('Truncated stream'); this.name = 'Truncated' }
}
export class InvalidOperand extends Error {
  constructor(message: string) { super(message); this.name = 'InvalidOperand' }
}

/** Wire enum indices are the positions in these frozen lists. */
export const STATE_ENUMS = {
  lineCap: ['butt', 'round', 'square'], lineJoin: ['round', 'bevel', 'miter'],
  imageSmoothingQuality: ['low', 'medium', 'high'], direction: ['inherit', 'ltr', 'rtl'],
  fontKerning: ['auto', 'normal', 'none'], textAlign: ['left', 'right', 'center', 'start', 'end'],
  textBaseline: ['top', 'hanging', 'middle', 'alphabetic', 'ideographic', 'bottom'],
  textRendering: ['auto', 'optimizeSpeed', 'optimizeLegibility', 'geometricPrecision'],
} as const
export interface StateFields {
  lineWidth?: number; lineCap?: typeof STATE_ENUMS.lineCap[number]; lineJoin?: typeof STATE_ENUMS.lineJoin[number]
  miterLimit?: number; lineDash?: number[]; lineDashOffset?: number; globalAlpha?: number
  globalCompositeOperation?: string; imageSmoothingEnabled?: boolean
  imageSmoothingQuality?: typeof STATE_ENUMS.imageSmoothingQuality[number]
  letterSpacing?: string; wordSpacing?: string; direction?: typeof STATE_ENUMS.direction[number]; lang?: string
  fontKerning?: typeof STATE_ENUMS.fontKerning[number]; fontVariantCaps?: string
  shadowBlur?: number; shadowOffsetX?: number; shadowOffsetY?: number; font?: string
  textAlign?: typeof STATE_ENUMS.textAlign[number]; textBaseline?: typeof STATE_ENUMS.textBaseline[number]
  fontStretch?: string; textRendering?: typeof STATE_ENUMS.textRendering[number]
}
export type Matrix = [number, number, number, number, number, number]
export interface Operands {
  [Opcode.STATE_PUSH]: Record<never, never>
  [Opcode.STATE_POP]: Record<never, never>
  [Opcode.TRANSFORM]: { mode: 0 | 1; matrix: Matrix }
  [Opcode.STYLE_COLOR]: { target: 0 | 1 | 2; color: string }
  [Opcode.STYLE_OBJ]: { target: 0 | 1; assetId: number }
  [Opcode.STATE_SETUP]: { fields: StateFields }
  [Opcode.PATH]: { sub: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10; args: number[]; ccw?: boolean; radii?: number[] }
  [Opcode.PATH_END]: { mode: 0 | 1 | 2 | 3; rule: 0 | 1 }
  [Opcode.RECT]: { mode: 0 | 1 | 2; rect: [number, number, number, number] }
  [Opcode.GLYPH_RUN]: { mode: 0 | 1; text: string; x: number; y: number; maxWidth?: number; advances: number[] }
  [Opcode.IMAGE_DRAW]: { assetId: number; args: number[] }
  [Opcode.TILE]: { w: number; h: number; dx: number; dy: number; rgba: Uint8Array }
}
export type OpCode = keyof Operands
export type Op = { [K in OpCode]: { code: K } & Operands[K] }[OpCode]
export type Asset = { id: number; hash: string } & (
  { type: 'bitmap'; w: number; h: number; rgba: Uint8Array } |
  { type: 'linear'; x0: number; y0: number; x1: number; y1: number; stops: Array<[number, string]> } |
  { type: 'radial'; x0: number; y0: number; r0: number; x1: number; y1: number; r1: number; stops: Array<[number, string]> } |
  { type: 'pattern'; bitmap: number; repetition: string; transform?: Matrix }
)

const stateKeys = [
  'lineWidth', 'lineCap', 'lineJoin', 'miterLimit', 'lineDash', 'lineDashOffset', 'globalAlpha',
  'globalCompositeOperation', 'imageSmoothingEnabled', 'imageSmoothingQuality', 'letterSpacing',
  'wordSpacing', 'direction', 'lang', 'fontKerning', 'fontVariantCaps', 'shadowBlur', 'shadowOffsetX',
  'shadowOffsetY', 'font', 'textAlign', 'textBaseline', 'fontStretch', 'textRendering',
] as const
const strings = new Set<string>(['globalCompositeOperation', 'letterSpacing', 'wordSpacing', 'lang', 'fontVariantCaps', 'font', 'fontStretch'])
const pathArgs = [0, 2, 2, 4, 6, 5, 5, 7, 4, 4, 0]
const encoder = new TextEncoder()
const decoder = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true })
function valid(condition: boolean, message: string): asserts condition { if (!condition) throw new InvalidOperand(message) }
function uint(n: number, max = 0xffffffff) { valid(Number.isInteger(n) && n >= 0 && n <= max, 'Invalid unsigned integer'); return n }
function utf8Length(s: string) {
  valid(typeof s === 'string', 'Expected string')
  let n = 0
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    if (c < 128) n++
    else if (c < 2048) n += 2
    else if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length && s.charCodeAt(i + 1) >= 0xdc00 && s.charCodeAt(i + 1) <= 0xdfff) { n += 4; i++ }
    else { valid(c < 0xd800 || c > 0xdfff, 'Unpaired UTF-16 surrogate'); n += 3 }
  }
  return uint(n)
}
/** First pass sizes and validates; second writes into an exactly-sized buffer. */
class Sink {
  at = 0
  private view?: DataView
  constructor(private bytes?: Uint8Array) { if (bytes) this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength) }
  u8(n: number) { uint(n, 255); this.view?.setUint8(this.at, n); this.at++ }
  u16(n: number) { uint(n, 65535); this.view?.setUint16(this.at, n, true); this.at += 2 }
  u32(n: number) { uint(n); this.view?.setUint32(this.at, n, true); this.at += 4 }
  f64(n: number) { valid(typeof n === 'number', 'Expected number'); this.view?.setFloat64(this.at, n, true); this.at += 8 }
  raw(b: Uint8Array) { this.bytes?.set(b, this.at); this.at += b.length }
  string(s: string) { const n = utf8Length(s); this.u32(n); if (this.bytes) encoder.encodeInto(s, this.bytes.subarray(this.at, this.at + n)); this.at += n }
  enum(n: number, max: number) { uint(n, max); this.u8(n) }
  bool(b: boolean) { valid(typeof b === 'boolean', 'Expected boolean'); this.u8(b ? 1 : 0) }
  numbers(a: number[], count?: number) { valid(Array.isArray(a) && (count === undefined || a.length === count), 'Invalid numeric arity'); for (const n of a) this.f64(n) }
}
function writeOp(w: Sink, op: Op) {
  w.u8(op.code)
  switch (op.code) {
    case Opcode.STATE_PUSH: case Opcode.STATE_POP: return
    case Opcode.TRANSFORM: w.enum(op.mode, 1); w.numbers(op.matrix, 6); return
    case Opcode.STYLE_COLOR: w.enum(op.target, 2); w.string(op.color); return
    case Opcode.STYLE_OBJ: w.enum(op.target, 1); w.u32(op.assetId); return
    case Opcode.STATE_SETUP: {
      const fields = op.fields
      valid(fields !== null && typeof fields === 'object', 'Invalid state')
      for (const key of Object.keys(fields)) valid((stateKeys as readonly string[]).includes(key), 'Unknown state field')
      let mask = 0
      stateKeys.forEach((key, i) => { if (fields[key] !== undefined) mask |= 1 << i })
      w.u32(mask)
      for (const key of stateKeys) {
        const value = fields[key]
        if (value === undefined) continue
        if (key in STATE_ENUMS) { const list = STATE_ENUMS[key as keyof typeof STATE_ENUMS] as readonly string[]; w.enum(list.indexOf(value as string), list.length - 1) }
        else if (key === 'lineDash') { w.u16((value as number[]).length); w.numbers(value as number[]) }
        else if (key === 'imageSmoothingEnabled') w.bool(value as boolean)
        else if (strings.has(key)) w.string(value as string)
        else w.f64(value as number)
      }
      return
    }
    case Opcode.PATH:
      w.enum(op.sub, 10); w.numbers(op.args, pathArgs[op.sub])
      if (op.sub === 5 || op.sub === 7) w.bool(op.ccw ?? false)
      else valid(op.ccw === undefined, 'Unexpected ccw')
      if (op.sub === 9) { valid(!!op.radii && op.radii.length >= 1 && op.radii.length <= 4, 'Invalid roundRect radii'); w.u8(op.radii.length); w.numbers(op.radii) }
      else valid(op.radii === undefined, 'Unexpected radii')
      return
    case Opcode.PATH_END: w.enum(op.mode, 3); w.enum(op.rule, 1); return
    case Opcode.RECT: w.enum(op.mode, 2); w.numbers(op.rect, 4); return
    case Opcode.GLYPH_RUN:
      w.enum(op.mode, 1); w.string(op.text); w.f64(op.x); w.f64(op.y); w.bool(op.maxWidth !== undefined)
      if (op.maxWidth !== undefined) w.f64(op.maxWidth)
      w.u32(op.advances.length); w.numbers(op.advances); return
    case Opcode.IMAGE_DRAW:
      w.u32(op.assetId); valid([2, 4, 8].includes(op.args.length), 'Invalid image argc'); w.u8(op.args.length + 1); w.numbers(op.args); return
    case Opcode.TILE:
      w.u32(op.w); w.u32(op.h); w.f64(op.dx); w.f64(op.dy)
      valid(op.rgba instanceof Uint8Array && op.w * op.h * 4 === op.rgba.length, 'Invalid RGBA length')
      w.u32(op.rgba.length); w.raw(op.rgba); return
    default: throw new UnknownOpcode((op as { code: number }).code)
  }
}
export class OpWriter {
  private chunks: Uint8Array[] = []
  private length = 0
  private count = 0
  get opCount() { return this.count }
  op<K extends OpCode>(code: K, ...args: K extends Opcode.STATE_PUSH | Opcode.STATE_POP ? [operands?: Operands[K]] : [operands: Operands[K]]): this {
    const op = { ...args[0], code } as Op
    const size = new Sink(); writeOp(size, op)
    uint(this.length + size.at); uint(this.count + 1)
    const bytes = new Uint8Array(size.at); writeOp(new Sink(bytes), op)
    this.chunks.push(bytes); this.length += bytes.length; this.count++
    return this
  }
  bytes(): Uint8Array { const bytes = new Uint8Array(this.length); let at = 0; for (const chunk of this.chunks) { bytes.set(chunk, at); at += chunk.length } return bytes }
}

class Cursor {
  at = 0
  private view: DataView
  constructor(readonly bytes: Uint8Array) { this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength) }
  need(n: number) { if (n > this.bytes.length - this.at) throw new Truncated() }
  u8() { this.need(1); return this.view.getUint8(this.at++) }
  u16() { this.need(2); const n = this.view.getUint16(this.at, true); this.at += 2; return n }
  u32() { this.need(4); const n = this.view.getUint32(this.at, true); this.at += 4; return n }
  f64() { this.need(8); const n = this.view.getFloat64(this.at, true); this.at += 8; return n }
  enum(max: number) { const n = this.u8(); valid(n <= max, 'Unknown enum value'); return n }
  bool() { return this.enum(1) === 1 }
  numbers(n: number) { this.need(n * 8); const a = new Array<number>(n); for (let i = 0; i < n; i++) a[i] = this.f64(); return a }
  raw(n: number) { this.need(n); const b = this.bytes.subarray(this.at, this.at + n); this.at += n; return b }
  string() { const b = this.raw(this.u32()); try { return decoder.decode(b) } catch { throw new InvalidOperand('Invalid UTF-8') } }
}
function readOp(r: Cursor): Op {
  const code = r.u8()
  switch (code) {
    case Opcode.STATE_PUSH: case Opcode.STATE_POP: return { code }
    case Opcode.TRANSFORM: return { code, mode: r.enum(1) as 0 | 1, matrix: r.numbers(6) as Matrix }
    case Opcode.STYLE_COLOR: return { code, target: r.enum(2) as 0 | 1 | 2, color: r.string() }
    case Opcode.STYLE_OBJ: return { code, target: r.enum(1) as 0 | 1, assetId: r.u32() }
    case Opcode.STATE_SETUP: {
      const mask = r.u32(); valid(mask <= 0xffffff, 'Unknown state mask')
      const fields: Record<string, unknown> = {}
      for (let i = 0; i < stateKeys.length; i++) if (mask & (1 << i)) {
        const key = stateKeys[i]
        if (key in STATE_ENUMS) { const list = STATE_ENUMS[key as keyof typeof STATE_ENUMS]; fields[key] = list[r.enum(list.length - 1)] }
        else if (key === 'lineDash') fields[key] = r.numbers(r.u16())
        else if (key === 'imageSmoothingEnabled') fields[key] = r.bool()
        else if (strings.has(key)) fields[key] = r.string()
        else fields[key] = r.f64()
      }
      return { code, fields: fields as StateFields }
    }
    case Opcode.PATH: {
      const sub = r.enum(10) as Operands[Opcode.PATH]['sub']; const args = r.numbers(pathArgs[sub])
      if (sub === 5 || sub === 7) return { code, sub, args, ccw: r.bool() }
      if (sub === 9) { const count = r.u8(); valid(count >= 1 && count <= 4, 'Invalid roundRect radii'); return { code, sub, args, radii: r.numbers(count) } }
      return { code, sub, args }
    }
    case Opcode.PATH_END: return { code, mode: r.enum(3) as 0 | 1 | 2 | 3, rule: r.enum(1) as 0 | 1 }
    case Opcode.RECT: return { code, mode: r.enum(2) as 0 | 1 | 2, rect: r.numbers(4) as [number, number, number, number] }
    case Opcode.GLYPH_RUN: {
      const mode = r.enum(1) as 0 | 1, text = r.string(), x = r.f64(), y = r.f64(), max = r.bool()
      const maxWidth = max ? r.f64() : undefined; const advances = r.numbers(r.u32())
      return { code, mode, text, x, y, ...(max ? { maxWidth } : {}), advances }
    }
    case Opcode.IMAGE_DRAW: {
      const assetId = r.u32(), argc = r.u8(); valid([3, 5, 9].includes(argc), 'Invalid image argc')
      return { code, assetId, args: r.numbers(argc - 1) }
    }
    case Opcode.TILE: {
      const w = r.u32(), h = r.u32(), dx = r.f64(), dy = r.f64(), n = r.u32()
      valid(w * h * 4 === n, 'Invalid RGBA length'); return { code, w, h, dx, dy, rgba: r.raw(n) }
    }
    default: throw new UnknownOpcode(code)
  }
}
export class OpReader implements Iterable<Op> {
  constructor(private readonly bytes: Uint8Array) {}
  *[Symbol.iterator](): Iterator<Op> {
    const r = new Cursor(this.bytes)
    while (r.at < this.bytes.length) yield readOp(r)
  }
}
