import type { CanvasBackend, CanvasLike, ImageLike } from '../core/canvas-backend'
import { Opcode, OpReader, type Asset, type Op, type TransformOperands, type Matrix } from './ops'

export class ReplayError extends Error {
  constructor(message: string) { super(message); this.name = 'ReplayError' }
}
const fail = (message: string): never => { throw new ReplayError(message) }
const finite = (...n: number[]) => n.every(Number.isFinite)
const pathMethods = ['beginPath', 'moveTo', 'lineTo', 'quadraticCurveTo', 'bezierCurveTo', 'arc', 'arcTo', 'ellipse', 'rect', 'roundRect', 'closePath'] as const
const styleNames = ['fillStyle', 'strokeStyle', 'shadowColor'] as const
function matrix(ctx: CanvasRenderingContext2D, values: Matrix): DOMMatrix {
  // Use the backend's own matrix class: native implementations can reject a
  // foreign DOMMatrix even though its six fields are structurally identical.
  const m = ctx.getTransform()
  ;[m.a, m.b, m.c, m.d, m.e, m.f] = values
  return m
}
function gradient(ctx: CanvasRenderingContext2D, a: Extract<Asset, { type: 'linear' | 'radial' }>): CanvasGradient {
  const g = a.type === 'linear' ? ctx.createLinearGradient(a.x0, a.y0, a.x1, a.y1) : ctx.createRadialGradient(a.x0, a.y0, a.r0, a.x1, a.y1, a.r1)
  for (const [at, color] of a.stops) g.addColorStop(at, color)
  return g
}

function applyTransform(ctx: CanvasRenderingContext2D, op: TransformOperands) {
  if (op.mode === 0) ctx.setTransform(...op.matrix)
  else if (op.mode === 1) ctx.transform(...op.matrix)
  else if (op.mode === 2) ctx.translate(...op.matrix)
  else if (op.mode === 3) ctx.scale(...op.matrix)
  else ctx.rotate(op.matrix[0])
}

/** Replay at the recorded scale. Validate and prepare every referenced asset
 * before touching destination pixels; a malformed late op cannot half-paint. */
export async function replayUnit(
  ctx: CanvasRenderingContext2D,
  unit: { ops: Uint8Array; assets: ReadonlyMap<number, Asset> },
  backend: Pick<CanvasBackend, 'createCanvas' | 'imageFromPixels'>,
): Promise<void> {
  try {
    const ops = [...new OpReader(unit.ops)]
    const assets = new Map<number, Asset>()
    const requireAsset = (id: number, bitmapOnly: boolean): Asset => {
      let a = assets.get(id)
      if (!a) {
        const original = unit.assets.get(id)
        if (!original || original.id !== id) return fail(`Unknown asset ${id}`)
        if (bitmapOnly && original.type !== 'bitmap') return fail(`Wrong asset type ${id}`)
        switch (original.type) {
          case 'bitmap':
            if (!['image', 'canvas'].includes(original.source) || !Number.isInteger(original.w) || !Number.isInteger(original.h) || original.w <= 0 || original.h <= 0 || !(original.rgba instanceof Uint8Array) || original.w * original.h * 4 !== original.rgba.length) return fail(`Invalid bitmap ${id}`)
            a = { ...original, rgba: original.rgba.slice() }; break
          case 'linear': case 'radial':
            if (!finite(original.x0, original.y0, original.x1, original.y1) || (original.type === 'radial' && (!finite(original.r0, original.r1) || original.r0 < 0 || original.r1 < 0))) return fail(`Invalid gradient ${id}`)
            if (!Array.isArray(original.stops) || original.stops.some(([n, color]) => !Number.isFinite(n) || n < 0 || n > 1 || typeof color !== 'string')) return fail(`Invalid stops ${id}`)
            if (original.transforms?.some(t => ![0,1,2,3,4].includes(t.mode) || t.matrix.length !== (t.mode < 2 ? 6 : t.mode < 4 ? 2 : 1) || !finite(...t.matrix))) return fail(`Invalid gradient transforms ${id}`)
            a = { ...original, stops: original.stops.map(([n, color]) => [n, color]), ...(original.transforms ? { transforms: original.transforms.map(t => ({ ...t, matrix: [...t.matrix] } as TransformOperands)) } : {}) }
            // Native color validation happens before any destination draw.
            gradient(ctx, a); break
          case 'pattern':
            if (!['repeat', 'repeat-x', 'repeat-y', 'no-repeat'].includes(original.repetition) || (original.transform && (original.transform.length !== 6 || !finite(...original.transform)))) return fail(`Invalid pattern ${id}`)
            a = { ...original, ...(original.transform ? { transform: [...original.transform] as Matrix } : {}) }
            requireAsset(a.bitmap, true); break
          default: return fail(`Unknown asset type ${id}`)
        }
        assets.set(id, a)
      }
      if (bitmapOnly ? a.type !== 'bitmap' : a.type === 'bitmap') return fail(`Wrong asset type ${id}`)
      return a
    }
    // Entire stream/reference validation is synchronous, before the first await.
    for (const op of ops) {
      if (op.code === Opcode.TRANSFORM && !finite(...op.matrix)) fail('Nonfinite transform')
      if (op.code === Opcode.GLYPH_RUN && (!finite(op.x, op.y) || (op.maxWidth !== undefined && (!Number.isFinite(op.maxWidth) || op.maxWidth < 0)))) fail('Invalid text coordinates')
      if (op.code === Opcode.STATE_SETUP && op.fields.lineDash?.some(n => !Number.isFinite(n) || n < 0)) fail('Invalid dash')
      if (op.code === Opcode.PATH && (!finite(...op.args) || (op.radii && !finite(...op.radii)))) fail('Nonfinite path')
      if (op.code === Opcode.IMAGE_DRAW) requireAsset(op.assetId, true)
      if (op.code === Opcode.STYLE_OBJ) requireAsset(op.assetId, false)
      if (op.code === Opcode.PATH && typeof ctx[pathMethods[op.sub]] !== 'function') fail(`Unsupported path ${pathMethods[op.sub]}`)
      if (op.code === Opcode.PATH && ((op.sub === 5 && op.args[2] < 0) || (op.sub === 6 && op.args[4] < 0) || (op.sub === 7 && (op.args[2] < 0 || op.args[3] < 0)) || (op.sub === 9 && op.radii!.some(r => r < 0)))) fail('Negative path radius')
      if (op.code === Opcode.TILE) op.rgba = op.rgba.slice()
    }
    const bitmaps = new Map<number, CanvasLike | ImageLike>(), tiles = new Map<Op, CanvasLike>(), patterns = new Map<number, CanvasPattern>()
    const makeBitmap = (w: number, h: number, rgba: Uint8Array) => {
      const canvas = backend.createCanvas(w, h), c = canvas.getContext('2d')
      if (!c) return fail('Bitmap context unavailable')
      const image = c.createImageData(w, h); image.data.set(rgba); c.putImageData(image, 0, 0)
      return canvas
    }
    for (const [id, asset] of assets) if (asset.type === 'bitmap') bitmaps.set(id, asset.source === 'image' ? await backend.imageFromPixels(asset.w, asset.h, asset.rgba) : makeBitmap(asset.w, asset.h, asset.rgba))
    for (const op of ops) if (op.code === Opcode.TILE && op.w && op.h) tiles.set(op, makeBitmap(op.w, op.h, op.rgba))
    for (const [id, asset] of assets) if (asset.type === 'pattern') {
      const pattern = ctx.createPattern(bitmaps.get(asset.bitmap)! as unknown as CanvasImageSource, asset.repetition)
      if (!pattern) fail(`Pattern unavailable ${id}`)
      if (asset.transform) pattern!.setTransform(matrix(ctx, asset.transform))
      patterns.set(id, pattern!)
    }
    // Prepare recorder gradients on an isolated context. Temporarily moving the
    // destination CTM can round/reproject an already-built native path.
    const gradients = new Map<number, CanvasGradient>()
    let gradientHost: CanvasRenderingContext2D | null | undefined
    for (const [id, a] of assets) if ((a.type === 'linear' || a.type === 'radial') && a.transforms) {
      const host = gradientHost ??= backend.createCanvas(1, 1).getContext('2d')
      if (!host) return fail('Gradient context unavailable')
      host.setTransform(1, 0, 0, 1, 0, 0)
      for (const t of a.transforms) applyTransform(host, t)
      gradients.set(id, gradient(host, a))
    }
    for (const op of ops) {
      switch (op.code) {
        case Opcode.STATE_PUSH: ctx.save(); break
        case Opcode.STATE_POP: ctx.restore(); break
        case Opcode.TRANSFORM:
          applyTransform(ctx, op)
          break
        case Opcode.STYLE_COLOR: ctx[styleNames[op.target]] = op.color; break
        case Opcode.STYLE_OBJ: {
          const a = assets.get(op.assetId)!
          const style = a.type === 'linear' || a.type === 'radial' ? gradients.get(a.id) ?? gradient(ctx, a) : patterns.get(a.id)!
          if (op.target === 0) ctx.fillStyle = style; else ctx.strokeStyle = style
          break
        }
        case Opcode.STATE_SETUP:
          for (const [key, value] of Object.entries(op.fields)) {
            if (key === 'lineDash') ctx.setLineDash(value as number[])
            else (ctx as unknown as Record<string, unknown>)[key] = value
          }
          break
        case Opcode.PATH: {
          const args: unknown[] = [...op.args]
          if (op.sub === 5 || op.sub === 7) args.push(op.ccw)
          if (op.sub === 9) args.push(op.radii)
          Reflect.apply(ctx[pathMethods[op.sub]], ctx, args); break
        }
        case Opcode.PATH_END: {
          const rule = op.rule === 0 ? 'nonzero' : 'evenodd'
          if (op.mode === 3) ctx.clip(rule)
          else { if (op.mode === 0 || op.mode === 2) ctx.fill(rule); if (op.mode === 1 || op.mode === 2) ctx.stroke() }
          break
        }
        case Opcode.RECT: Reflect.apply([ctx.fillRect, ctx.strokeRect, ctx.clearRect][op.mode], ctx, op.rect); break
        case Opcode.GLYPH_RUN:
          // Preserve the native shaped run, kerning and maxWidth behavior. The
          // advances accompany the run; splitting it into glyph draws changes pixels.
          Reflect.apply(op.mode === 0 ? ctx.fillText : ctx.strokeText, ctx, [op.text, op.x, op.y, ...(op.maxWidth === undefined ? [] : [op.maxWidth])]); break
        case Opcode.IMAGE_DRAW: Reflect.apply(ctx.drawImage, ctx, [bitmaps.get(op.assetId), ...op.args]); break
        case Opcode.TILE: { const bitmap = tiles.get(op); if (bitmap) ctx.drawImage(bitmap as unknown as CanvasImageSource, op.dx, op.dy); break }
      }
    }
  } catch (error) {
    if (error instanceof ReplayError) throw error
    throw new ReplayError(error instanceof Error ? error.message : String(error))
  }
}
