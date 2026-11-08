/**
 * ASSET_DEFINE (§7.1 kind 2) payload codec. Assets are unit-local (ids are the
 * recorder's table indices) and precede their UNIT_COMMANDS message under the
 * same reqId. Bitmaps travel as raw RGBA inside the once-gzipped message
 * (metadata-free by construction, §7.7). Every decoded asset's hash is
 * recomputed with the recorder's canonical encoding; a mismatch fails the
 * whole payload.
 */
import type { Asset, Matrix, TransformOperands } from './ops'

export class AssetWireError extends Error { constructor(m: string) { super(m); this.name = 'AssetWireError' } }
export class AssetTooLarge extends Error { constructor(readonly assetId: number, readonly bytes: number) { super(`asset ${assetId} needs ${bytes} B`); this.name = 'AssetTooLarge' } }

/** §14 ASSET_DEFINE decoded cap. */
export const ASSET_DEFINE_DECODED_CAP = 8 * 1024 * 1024
const TYPES = ['bitmap', 'linear', 'radial', 'pattern'] as const
const enc = new TextEncoder(), dec = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true })

class Out {
  private chunks: Uint8Array[] = []; size = 0
  private push(b: Uint8Array) { this.chunks.push(b); this.size += b.length }
  u8(n: number) { this.push(Uint8Array.of(n)) }
  u16(n: number) { const b = new Uint8Array(2); new DataView(b.buffer).setUint16(0, n, true); this.push(b) }
  u32(n: number) { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); this.push(b) }
  f64(n: number) { const b = new Uint8Array(8); new DataView(b.buffer).setFloat64(0, n, true); this.push(b) }
  str(s: string) { const b = enc.encode(s); this.u32(b.length); this.push(b) }
  raw(b: Uint8Array) { this.push(b) }
  bytes() { const o = new Uint8Array(this.size); let at = 0; for (const c of this.chunks) { o.set(c, at); at += c.length } return o }
}
class In {
  at = 0; private v: DataView
  constructor(private b: Uint8Array) { this.v = new DataView(b.buffer, b.byteOffset, b.byteLength) }
  private need(n: number) { if (this.at + n > this.b.length) throw new AssetWireError('truncated') }
  u8() { this.need(1); return this.v.getUint8(this.at++) }
  u16() { this.need(2); const n = this.v.getUint16(this.at, true); this.at += 2; return n }
  u32() { this.need(4); const n = this.v.getUint32(this.at, true); this.at += 4; return n }
  f64() { this.need(8); const n = this.v.getFloat64(this.at, true); this.at += 8; return n }
  str() { const n = this.u32(); this.need(n); try { return dec.decode(this.b.subarray(this.at, this.at += n)) } catch { throw new AssetWireError('utf8') } }
  raw(n: number) { this.need(n); return this.b.slice(this.at, this.at += n) }
  get done() { return this.at === this.b.length }
}

const ARITY = [6, 6, 2, 2, 1]
function writeTransforms(o: Out, ts: TransformOperands[] | undefined) { o.u16(ts?.length ?? 0); for (const t of ts ?? []) { o.u8(t.mode); for (const n of t.matrix) o.f64(n) } }
function readTransforms(i: In): TransformOperands[] | undefined {
  const n = i.u16(); if (n === 0) return undefined
  return Array.from({ length: n }, () => { const mode = i.u8(); if (mode > 4) throw new AssetWireError('transform mode'); return { mode, matrix: Array.from({ length: ARITY[mode]! }, () => i.f64()) } as TransformOperands })
}
function writeStops(o: Out, stops: Array<[number, string]>) { o.u16(stops.length); for (const [at, c] of stops) { o.f64(at); o.str(c) } }
const readStops = (i: In): Array<[number, string]> => Array.from({ length: i.u16() }, () => [i.f64(), i.str()] as [number, string])

function writeAsset(o: Out, a: Asset) {
  o.u32(a.id); o.str(a.hash); o.u8(TYPES.indexOf(a.type))
  if (a.type === 'bitmap') { o.u8(a.source === 'image' ? 0 : 1); o.u32(a.w); o.u32(a.h); o.raw(a.rgba) }
  else if (a.type === 'linear') { for (const n of [a.x0, a.y0, a.x1, a.y1]) o.f64(n); writeStops(o, a.stops); writeTransforms(o, a.transforms) }
  else if (a.type === 'radial') { for (const n of [a.x0, a.y0, a.r0, a.x1, a.y1, a.r1]) o.f64(n); writeStops(o, a.stops); writeTransforms(o, a.transforms) }
  else { o.u32(a.bitmap); o.str(a.repetition); o.u8(a.transform ? 1 : 0); for (const n of a.transform ?? []) o.f64(n) }
}

/** Encodes `assets` into ≥ 1 payloads, each ≤ the decoded cap. A single
 * asset over the cap throws `AssetTooLarge` (the caller ships a raster). */
export function encodeAssets(assets: readonly Asset[], cap = ASSET_DEFINE_DECODED_CAP): Uint8Array[] {
  const out: Uint8Array[] = []
  let batch: Uint8Array[] = [], size = 4
  const flush = () => { if (!batch.length) return; const o = new Out(); o.u32(batch.length); for (const b of batch) o.raw(b); out.push(o.bytes()); batch = []; size = 4 }
  for (const a of assets) {
    const o = new Out(); writeAsset(o, a); const b = o.bytes()
    if (b.length + 4 > cap) throw new AssetTooLarge(a.id, b.length + 4)
    if (size + b.length > cap) flush()
    batch.push(b); size += b.length
  }
  flush()
  return out
}

/** Canonical asset bytes: identical to the recorder's `hashValue` stream
 * (tags n/s/b/a/o, sorted keys, LE f64 lengths), so SHA-256 over them equals
 * the recorder's hash. Pattern bitmaps hash by the referenced asset's hash. */
export function canonicalAssetBytes(value: unknown): Uint8Array {
  const o = new Out()
  const tag = (t: string) => o.raw(enc.encode(t))
  const visit = (v: unknown): void => {
    if (typeof v === 'number') { tag('n'); o.f64(v) }
    else if (typeof v === 'string') { const b = enc.encode(v); tag('s'); o.f64(b.length); o.raw(b) }
    else if (v instanceof Uint8Array) { tag('b'); o.f64(v.length); o.raw(v) }
    else if (Array.isArray(v)) { tag('a'); o.f64(v.length); v.forEach(visit) }
    else if (v && typeof v === 'object') { const keys = Object.keys(v).sort(); tag('o'); o.f64(keys.length); for (const k of keys) { visit(k); visit((v as Record<string, unknown>)[k]) } }
    else throw new AssetWireError('unsupported canonical value')
  }
  visit(value)
  return o.bytes()
}
const hexOf = (b: ArrayBuffer) => Array.from(new Uint8Array(b), (x) => x.toString(16).padStart(2, '0')).join('')

/** Decodes and verifies one ASSET_DEFINE payload. `known` resolves pattern
 * bitmap references to assets of the same unit defined earlier. */
export async function decodeAssets(payload: Uint8Array, known: ReadonlyMap<number, Asset> = new Map()): Promise<Asset[]> {
  const i = new In(payload), out: Asset[] = [], byId = new Map(known)
  const count = i.u32()
  if (count === 0) throw new AssetWireError('empty')
  for (let k = 0; k < count; k++) {
    const id = i.u32(), hash = i.str(), type = TYPES[i.u8()]
    if (!/^[0-9a-f]{64}$/.test(hash)) throw new AssetWireError('hash format')
    let a: Asset
    if (type === 'bitmap') {
      const s = i.u8(); if (s > 1) throw new AssetWireError('bitmap source')
      const w = i.u32(), h = i.u32()
      if (w < 1 || h < 1 || w * h * 4 > ASSET_DEFINE_DECODED_CAP) throw new AssetWireError('bitmap size')
      a = { id, hash, type, source: s === 0 ? 'image' : 'canvas', w, h, rgba: i.raw(w * h * 4) }
    } else if (type === 'linear') {
      const [x0, y0, x1, y1] = [i.f64(), i.f64(), i.f64(), i.f64()]
      const stops = readStops(i), transforms = readTransforms(i)
      a = { id, hash, type, x0, y0, x1, y1, stops, ...(transforms ? { transforms } : {}) }
    } else if (type === 'radial') {
      const [x0, y0, r0, x1, y1, r1] = [i.f64(), i.f64(), i.f64(), i.f64(), i.f64(), i.f64()]
      const stops = readStops(i), transforms = readTransforms(i)
      a = { id, hash, type, x0, y0, r0, x1, y1, r1, stops, ...(transforms ? { transforms } : {}) }
    } else if (type === 'pattern') {
      const bitmap = i.u32(), repetition = i.str(), has = i.u8()
      if (has > 1) throw new AssetWireError('pattern flag')
      const transform = has ? (Array.from({ length: 6 }, () => i.f64()) as Matrix) : undefined
      a = { id, hash, type, bitmap, repetition, ...(transform ? { transform } : {}) }
    } else throw new AssetWireError('asset type')
    if (byId.has(id)) throw new AssetWireError('duplicate id')
    const { id: _id, hash: _hash, ...content } = a as Asset & Record<string, unknown>
    void _id; void _hash
    let canonical: unknown = content
    if (a.type === 'pattern') {
      const ref = byId.get(a.bitmap)
      if (ref?.type !== 'bitmap') throw new AssetWireError('pattern bitmap')
      canonical = { ...content, bitmap: ref.hash }
    }
    if (hexOf(await crypto.subtle.digest('SHA-256', canonicalAssetBytes(canonical) as BufferSource)) !== hash) throw new AssetWireError('hash mismatch')
    byId.set(id, a); out.push(a)
  }
  if (!i.done) throw new AssetWireError('trailing bytes')
  return out
}
