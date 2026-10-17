/**
 * One shared, admitted measurement context per viewer (spec §§5.6, 12.4).
 *
 * The canonical layout measures text on a context and on the local identity
 * context core/text-metrics derives from it. Laying out on each tile's own
 * context would create (and grow) that pair per tile, outside admission. The
 * viewer instead lays every record out on one shared 1×1 context:
 * - both contexts are admitted (TILE_CANVAS_BYTES each) before the canvas exists;
 * - every distinct (font, text) either context shapes is admitted with the
 *   shapeCacheBytes model before it is measured, in O(1);
 * - a refusal throws LayoutAdmissionError, which the painter turns into its
 *   declared coarse tile (or the find overlay into a whole-cell highlight);
 * - trim() drops the pair (returning its charge) once its caches pass a cap,
 *   so the retained shaping caches stay bounded.
 * The admitted bytes are retained under the canvas-text measure charge
 * (LayoutAdmission.measureChanged) while the contexts live.
 */
import type { XlsxSheet } from '../../xlsx/types'
import { createCellLayoutHost } from '../../xlsx/view/renderViewport'
import { LayoutAdmissionError, type CellLayoutHost, type LayoutAdmission } from '../../xlsx/view/textLayout'
import { TILE_CANVAS_BYTES, shapeCacheBytes } from '../../xlsx/view/mainLedger'

/** Retained shaping caches beyond this are dropped by trim(). */
export const MEASURE_CONTEXT_CAP = 4 * 1024 * 1024

export class SharedMeasureContext {
  private ctx: CanvasRenderingContext2D | null = null
  private generation = 0
  private charged = 0
  private readonly caches = new Map<string, { seen: Set<string>; units: number }>()
  constructor(private readonly admission: () => LayoutAdmission, private readonly cap = MEASURE_CONTEXT_CAP) {}

  /** Bytes retained by the shared contexts and their shaping caches. */
  get bytes(): number { return this.charged }

  private grow(bytes: number): void {
    if (bytes <= 0) return
    const admission = this.admission()
    if (!admission.admit(bytes)) throw new LayoutAdmissionError()
    admission.release(bytes)
    admission.measureChanged?.(bytes)
    this.charged += bytes
  }
  /** Reserves one distinct text in the named context's shaping cache. */
  private shape(cache: string, font: string, text: string): void {
    const key = `${cache}\u0000${font}`
    let entry = this.caches.get(key)
    if (entry?.seen.has(text)) return
    entry ??= { seen: new Set(), units: 0 }
    this.grow(shapeCacheBytes(entry.seen.size + 1, entry.units + text.length) - shapeCacheBytes(entry.seen.size, entry.units))
    entry.seen.add(text)
    entry.units += text.length
    this.caches.set(key, entry)
  }
  private context(): CanvasRenderingContext2D | null {
    if (this.ctx) return this.ctx
    // The base context and the local identity context the layout derives from it.
    this.grow(2 * TILE_CANVAS_BYTES)
    let ctx: CanvasRenderingContext2D | null = null
    try {
      const canvas = globalThis.document?.createElement('canvas')
      if (canvas) { canvas.width = 1; canvas.height = 1 }
      ctx = canvas?.getContext('2d') ?? null
    } catch { ctx = null }
    if (!ctx) return null
    const native = ctx.measureText.bind(ctx), base = ctx
    ctx.measureText = (text: string) => { this.shape('base', base.font, text); return native(text) }
    this.ctx = ctx
    this.generation++
    return ctx
  }
  /** Drops the contexts once their caches pass the cap (between tiles). */
  trim(): void {
    if (this.charged <= this.cap) return
    this.admission().measureChanged?.(-this.charged)
    this.charged = 0
    this.ctx = null
    this.caches.clear()
  }
  /** A layout host on the shared contexts (created lazily, admitted first). */
  host(sheet: Partial<XlsxSheet>): CellLayoutHost {
    // Formatting and overflow certification need no context.
    const pure = createCellLayoutHost(undefined as unknown as CanvasRenderingContext2D, sheet)
    let made: { host: CellLayoutHost; generation: number } | null = null
    const current = (): CellLayoutHost => {
      const ctx = this.context()
      if (!ctx) throw new LayoutAdmissionError()
      if (!made || made.generation !== this.generation) made = { host: createCellLayoutHost(ctx, sheet), generation: this.generation }
      return made.host
    }
    return {
      layoutBody: (cell, width, height, overflow, preformatted) => {
        const host = current(), base = this.ctx!
        // The local identity context is created inside the layout and measures
        // through the shared 2D prototype: observe exactly this synchronous call.
        const proto = Object.getPrototypeOf(base) as CanvasRenderingContext2D
        const original = proto.measureText
        const shape = (font: string, text: string) => this.shape('local', font, text)
        proto.measureText = function (this: CanvasRenderingContext2D, text: string) {
          if (this !== base) shape(this.font, text)
          return original.call(this, text)
        }
        try { return host.layoutBody(cell, width, height, overflow, preformatted) } finally { proto.measureText = original }
      },
      canReuseOverflow: (body, overflow) => pure.canReuseOverflow!(body, overflow),
      paintedText: (cell, width, admission) => pure.paintedText!(cell, width, admission),
    }
  }
}
