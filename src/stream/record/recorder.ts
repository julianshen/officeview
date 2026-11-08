/** Server/test-only canvas recorder. Never import this module from the client runtime. */
import { createHash } from 'node:crypto'
import { RECORD_TEXT } from '../../core/text-recording'
import { Opcode, OpWriter, type Asset, type Matrix, type StateFields } from '../ops'
import { RECORDED_SUBSET } from './subset'

type WithoutIdentity<T> = T extends Asset ? Omit<T, 'id' | 'hash'> : never
export type AssetInput = WithoutIdentity<Asset>
export type RecordedUnit = { discriminator: 1; ops: Uint8Array; opCount: number; assets: Asset[] } | { discriminator: 3; reason: string }
export type ImageSnapshot = { w: number; h: number; rgba: Uint8Array }

/** Canonical tagged encoding: sorted object keys, UTF-8 lengths, LE f64s, raw bytes.
 * IDs are excluded; pattern references hash the referenced bitmap's content hash. */
function hashValue(value: unknown): string {
  const hash = createHash('sha256')
  const number = (n: number) => { const b = new Uint8Array(8); new DataView(b.buffer).setFloat64(0, n, true); hash.update(b) }
  const visit = (v: unknown): void => {
    if (typeof v === 'number') { hash.update('n'); number(v) }
    else if (typeof v === 'string') { const b = new TextEncoder().encode(v); hash.update('s'); number(b.length); hash.update(b) }
    else if (v instanceof Uint8Array) { hash.update('b'); number(v.length); hash.update(v) }
    else if (Array.isArray(v)) { hash.update('a'); number(v.length); v.forEach(visit) }
    else if (v && typeof v === 'object') { const keys = Object.keys(v).sort(); hash.update('o'); number(keys.length); for (const k of keys) { visit(k); visit((v as Record<string, unknown>)[k]) } }
    else throw new Error('Unsupported canonical asset value')
  }
  visit(value); return hash.digest('hex')
}
function bitmap(b: ImageSnapshot): void {
  if (!Number.isInteger(b.w) || !Number.isInteger(b.h) || b.w < 1 || b.h < 1 || b.w * b.h * 4 !== b.rgba.length) throw new Error('Invalid bitmap')
}
export class AssetTable {
  private readonly entries: Asset[] = []
  private readonly hashes = new Map<string, number>()
  add(input: AssetInput): number {
    let owned: AssetInput
    if (input.type === 'bitmap') {
      bitmap(input)
      // Hash before the retained copy, so identical snapshots need no second allocation.
      owned = { type: input.type, source: input.source, w: input.w, h: input.h, rgba: input.rgba }
    } else if (input.type === 'linear') owned = { type: input.type, x0: input.x0, y0: input.y0, x1: input.x1, y1: input.y1, stops: input.stops.map(([at, c]) => [at, c]) }
    else if (input.type === 'radial') owned = { type: input.type, x0: input.x0, y0: input.y0, r0: input.r0, x1: input.x1, y1: input.y1, r1: input.r1, stops: input.stops.map(([at, c]) => [at, c]) }
    else owned = { type: 'pattern', bitmap: input.bitmap, repetition: input.repetition, ...(input.transform ? { transform: [...input.transform] as Matrix } : {}) }
    const reference = owned.type === 'pattern' ? this.entries[owned.bitmap] : undefined
    if (owned.type === 'pattern' && reference?.type !== 'bitmap') throw new Error('Unknown pattern bitmap')
    const hash = hashValue(owned.type === 'pattern' ? { ...owned, bitmap: reference!.hash } : owned)
    const previous = this.hashes.get(hash); if (previous !== undefined) return previous
    if (owned.type === 'bitmap') owned = { ...owned, rgba: owned.rgba.slice() }
    const id = this.entries.length
    this.entries.push({ ...owned, id, hash } as Asset); this.hashes.set(hash, id); return id
  }
  all(): Asset[] { return this.entries.slice() }
}

type Native = Record<PropertyKey, any>
type Wrapped = { native: object; facade: object; input: AssetInput; matrix?: Matrix; source?: CanvasImageSource }
const matrixOf = (m: DOMMatrix2DInit = {}): Matrix => [m.a ?? m.m11 ?? 1, m.b ?? m.m12 ?? 0, m.c ?? m.m21 ?? 0, m.d ?? m.m22 ?? 1, m.e ?? m.m41 ?? 0, m.f ?? m.m42 ?? 0]
const pathSubs: Record<string, number> = { beginPath: 0, moveTo: 1, lineTo: 2, quadraticCurveTo: 3, bezierCurveTo: 4, arc: 5, arcTo: 6, ellipse: 7, rect: 8, roundRect: 9, closePath: 10 }
const pathCounts = [0, 2, 2, 4, 6, 5, 5, 7, 4, 4, 0]
const paints = new Set(['fill', 'stroke', 'fillRect', 'strokeRect', 'fillText', 'strokeText'])

export function createRecorder(backing: CanvasRenderingContext2D, opts: {
  scale: number; assets?: AssetTable;
  partialAlphaSnapshots?: 'refuse' | 'allow';
  /** Match native pattern ownership: Cairo uses live canvases; napi snapshots at creation. */
  patternSources?: 'live' | 'snapshot';
  /** Synchronous backend reader for decoded image sources; canvases read their own pixels. */
  readImage?: (source: CanvasImageSource) => ImageSnapshot;
}): { ctx: CanvasRenderingContext2D; finish(): RecordedUnit } {
  if (!Number.isFinite(opts.scale) || opts.scale <= 0) throw new RangeError('Invalid recorder scale')
  // Scale is supplied by the caller's paint transform; never inferred from browser DPR.
  const writer = new OpWriter(), assets = opts.assets ?? new AssetTable(), native = backing as unknown as Native
  const wrappers = new WeakMap<object, Wrapped>()
  let reason: string | undefined
  const refuse = (name: string) => { reason ??= name }
  const record = (name: string, action: () => void) => { if (!reason) try { action() } catch { refuse(name) } }
  function exactSnapshot(image: ImageSnapshot): ImageSnapshot {
    bitmap(image)
    // Straight-RGBA readback/re-upload is not lossless for premultiplied
    // partial alpha on all backends (node-canvas floors the conversion).
    // Keep the whole unit as a tile rather than claim a zero-diff command unit.
    if (opts.partialAlphaSnapshots !== 'allow') for (let i = 3; i < image.rgba.length; i += 4) if (image.rgba[i] !== 0 && image.rgba[i] !== 255) throw new Error('Non-opaque snapshot requires raster fallback')
    return image
  }
  function snapshot(source: CanvasImageSource): ImageSnapshot & { source: 'canvas' | 'image' } {
    const canvas = source as unknown as { width: number; height: number; naturalWidth?: number; naturalHeight?: number; getContext?: (type: string) => CanvasRenderingContext2D | null }
    if (typeof canvas.getContext === 'function') {
      const ctx = canvas.getContext('2d'); if (!ctx) throw new Error('Unreadable canvas')
      const rgba = ctx.getImageData(0, 0, canvas.width, canvas.height).data
      return { ...exactSnapshot({ w: canvas.width, h: canvas.height, rgba: new Uint8Array(rgba.buffer, rgba.byteOffset, rgba.byteLength) }), source: 'canvas' }
    }
    if (!opts.readImage) throw new Error('Decoded image requires readImage')
    const image = opts.readImage(source)
    if (image.w !== (canvas.naturalWidth ?? canvas.width) || image.h !== (canvas.naturalHeight ?? canvas.height)) throw new Error('Mismatched image dimensions')
    return { ...exactSnapshot(image), source: 'image' }
  }
  const styles: Record<string, any> = { fillStyle: native.fillStyle, strokeStyle: native.strokeStyle, shadowColor: native.shadowColor }
  const savedStyles: Array<Record<string, any>> = []
  const styleIds: Record<string, number | undefined> = {}
  const savedStyleIds: Array<Record<string, number | undefined>> = []
  function style(name: 'fillStyle' | 'strokeStyle' | 'shadowColor') {
    const value = styles[name], target = name === 'fillStyle' ? 0 : name === 'strokeStyle' ? 1 : 2
    if (typeof value === 'string') { styleIds[name] = undefined; writer.op(Opcode.STYLE_COLOR, { target, color: value }) }
    else {
      const wrapper = wrappers.get(value); if (!wrapper || target === 2) throw new Error('Unrecorded style')
      if (wrapper.input.type === 'pattern' && wrapper.source && opts.patternSources !== 'snapshot') wrapper.input.bitmap = assets.add({ type: 'bitmap', ...snapshot(wrapper.source) })
      const assetId = assets.add(wrapper.input)
      // Rebinding an unchanged gradient unnecessarily round-trips the path's
      // transform on native backends. Its immutable snapshot is already active.
      if (styleIds[name] === assetId) return
      styleIds[name] = assetId
      const current = matrixOf(backing.getTransform())
      if (wrapper.matrix) writer.op(Opcode.TRANSFORM, { mode: 0, matrix: wrapper.matrix })
      writer.op(Opcode.STYLE_OBJ, { target, assetId })
      if (wrapper.matrix) writer.op(Opcode.TRANSFORM, { mode: 0, matrix: current })
    }
  }
  function wrap(out: object, input: AssetInput, matrix?: Matrix, source?: CanvasImageSource): object {
    const wrapper: Wrapped = { native: out, facade: {}, input, matrix, source }
    wrapper.facade = new Proxy({}, {
      get(_target, key) {
        const value = Reflect.get(out, key, out)
        if (typeof value !== 'function') return value
        return (...args: any[]) => {
          const result = Reflect.apply(value, out, args)
          record(String(key), () => {
            if (key === 'addColorStop' && (input.type === 'linear' || input.type === 'radial')) input.stops.push([Number(args[0]), String(args[1])])
            else if (key === 'setTransform' && input.type === 'pattern') input.transform = matrixOf(args[0])
            else refuse(String(key))
          })
          return result
        }
      },
      has: (_target, key) => Reflect.has(out, key),
      getPrototypeOf: () => Object.getPrototypeOf(out),
      set(_target, key, value) { Reflect.set(out, key, value, out); refuse(String(key)); return true },
    })
    wrappers.set(out, wrapper); wrappers.set(wrapper.facade, wrapper); return wrapper.facade
  }
  // Capture caller-supplied initial state (including a pre-existing scale transform).
  record('initial-state', () => {
    writer.op(Opcode.TRANSFORM, { mode: 0, matrix: matrixOf(backing.getTransform()) })
    for (const key of RECORDED_SUBSET.setters) {
      if (key === 'fillStyle' || key === 'strokeStyle' || key === 'shadowColor') {
        // Exact transparent zero is safe (and is the default shadow); fractional
        // initial alpha may already have lost precision in native readback.
        const color = styles[key]
        if (typeof color === 'string') {
          const rgba = /^rgba\([^,]+,[^,]+,[^,]+,\s*([\d.]+)%?\s*\)$/i.exec(color)
          const hex = /^#(?:[0-9a-f]{6})([0-9a-f]{2})$/i.exec(color)
          const short = /^#[0-9a-f]{3}([0-9a-f])$/i.exec(color)
          const alpha = rgba ? Number(rgba[1]) : hex ? parseInt(hex[1], 16) / 255 : short ? parseInt(short[1], 16) / 15 : 1
          if (alpha > 0 && alpha < 1) throw new Error('Ambiguous initial alpha')
        }
        style(key)
      }
      else if (native[key] !== undefined) writer.op(Opcode.STATE_SETUP, { fields: { [key]: native[key] } as StateFields })
    }
    writer.op(Opcode.STATE_SETUP, { fields: { lineDash: backing.getLineDash() } })
  })
  const canvasView = new Proxy({}, {
    get(_target, key) {
      if (key === 'width' || key === 'height' || key === 'ownerDocument' || key === 'constructor') return Reflect.get(backing.canvas, key, backing.canvas)
      const value = Reflect.get(backing.canvas, key, backing.canvas)
      if (value === undefined) return undefined
      if (typeof value === 'function') return (...args: any[]) => { refuse(`canvas.${String(key)}`); return Reflect.apply(value, backing.canvas, args) }
      refuse(`canvas.${String(key)}`); return value
    },
    set(_target, key, value) { refuse(`canvas.${String(key)}`); return Reflect.set(backing.canvas, key, value, backing.canvas) },
  })
  const ctx = new Proxy({}, {
    get(_target, prop) {
      if (prop === RECORD_TEXT) return undefined
      if (prop === 'canvas') return canvasView
      const value = Reflect.get(backing, prop, backing)
      if (value && typeof value === 'object' && wrappers.has(value)) return wrappers.get(value)!.facade
      if (typeof value !== 'function') return value
      return (...args: any[]) => {
        const name = String(prop)
        let image: number | undefined
        // TODO(§7.7): crop source-rect snapshots only after edge-filter fidelity is proven; commands are relaxed-session only.
        if (name === 'drawImage' || name === 'createPattern') record(name, () => { image = assets.add({ type: 'bitmap', ...snapshot(args[0]) }) })
        // Wrapped styles can change after assignment. Snapshot their current state at each paint.
        if (paints.has(name)) record(name, () => { for (const key of ['fillStyle', 'strokeStyle'] as const) if (typeof styles[key] === 'object') style(key) })
        const result = Reflect.apply(value, backing, args)
        if (!RECORDED_SUBSET.methods.has(name)) { refuse(name); return result }
        if (name === 'createLinearGradient' || name === 'createRadialGradient') {
          const input: AssetInput = name === 'createLinearGradient'
            ? { type: 'linear', x0: +args[0], y0: +args[1], x1: +args[2], y1: +args[3], stops: [] }
            : { type: 'radial', x0: +args[0], y0: +args[1], r0: +args[2], x1: +args[3], y1: +args[4], r1: +args[5], stops: [] }
          if (!result || typeof result !== 'object') { refuse(name); return result }
          return wrap(result, input, matrixOf(backing.getTransform()))
        }
        if (name === 'createPattern') return result && image !== undefined ? wrap(result, { type: 'pattern', bitmap: image, repetition: args[1] || 'repeat' }, undefined, args[0]) : result
        record(name, () => {
          if (['measureText', 'getTransform', 'getLineDash'].includes(name)) return
          if (name === 'save' || name === 'restore') {
            if (name === 'save') { savedStyles.push({ ...styles }); savedStyleIds.push({ ...styleIds }) }
            else if (savedStyles.length) {
              Object.assign(styles, savedStyles.pop())
              for (const key of Object.keys(styleIds)) delete styleIds[key]
              Object.assign(styleIds, savedStyleIds.pop())
            }
            writer.op(name === 'save' ? Opcode.STATE_PUSH : Opcode.STATE_POP); return }
          if (['setTransform', 'transform', 'translate', 'scale', 'rotate'].includes(name)) { writer.op(Opcode.TRANSFORM, { mode: 0, matrix: matrixOf(backing.getTransform()) }); return }
          if (name === 'setLineDash') { writer.op(Opcode.STATE_SETUP, { fields: { lineDash: backing.getLineDash() } }); return }
          if (name in pathSubs) {
            const sub = pathSubs[name] as 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10
            const operands = args.slice(0, pathCounts[sub]).map(Number)
            const radii = name === 'roundRect' ? (Array.isArray(args[4]) ? args[4] : [args[4] ?? 0]) : undefined
            writer.op(Opcode.PATH, { sub, args: operands, ...(sub === 5 || sub === 7 ? { ccw: !!args[pathCounts[sub]] } : {}), ...(radii ? { radii } : {}) }); return
          }
          if (name === 'fill' || name === 'stroke' || name === 'clip') {
            if (args.length && typeof args[0] !== 'string') throw new Error('Path2D is not an op-list path')
            writer.op(Opcode.PATH_END, { mode: name === 'fill' ? 0 : name === 'stroke' ? 1 : 3, rule: args[0] === 'evenodd' ? 1 : 0 }); return
          }
          if (['fillRect', 'strokeRect', 'clearRect'].includes(name)) { writer.op(Opcode.RECT, { mode: name === 'fillRect' ? 0 : name === 'strokeRect' ? 1 : 2, rect: args.slice(0, 4).map(Number) as [number, number, number, number] }); return }
          if (name === 'fillText' || name === 'strokeText') {
            const text = String(args[0])
            writer.op(Opcode.GLYPH_RUN, { mode: name === 'fillText' ? 0 : 1, text, x: +args[1], y: +args[2], ...(args[3] !== undefined ? { maxWidth: +args[3] } : {}), advances: [backing.measureText(text).width] }); return
          }
          if (name === 'drawImage' && image !== undefined) { writer.op(Opcode.IMAGE_DRAW, { assetId: image, args: args.slice(1).map(Number) }); return }
          refuse(name) // e.g. reserved rasterBlock has no Stage-1b opcode yet.
        })
        return result
      }
    },
    has: (_target, prop) => prop !== RECORD_TEXT && Reflect.has(backing, prop),
    getPrototypeOf: () => Object.getPrototypeOf(backing),
    set(_target, prop, value) {
      if (prop === RECORD_TEXT) return true
      const wrapper = value && typeof value === 'object' ? wrappers.get(value) : undefined
      Reflect.set(backing, prop, wrapper?.native ?? value, backing)
      const name = String(prop)
      if (!RECORDED_SUBSET.setters.has(name)) refuse(name)
      else record(name, () => {
        if (name === 'fillStyle' || name === 'strokeStyle' || name === 'shadowColor') { styles[name] = typeof value === 'string' ? value : wrapper?.native ?? value; style(name) }
        else writer.op(Opcode.STATE_SETUP, { fields: { [name]: native[name] } as StateFields })
      })
      return true
    },
  }) as CanvasRenderingContext2D
  return { ctx, finish: () => reason ? { discriminator: 3, reason } : { discriminator: 1, ops: writer.bytes(), opCount: writer.opCount, assets: assets.all() } }
}
