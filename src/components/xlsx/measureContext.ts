/**
 * One shared, admitted measurement context per viewer (spec §§5.6, 12.4).
 *
 * The canonical layout measures text on a context and on a local identity
 * context that core/text-metrics derives from `base.canvas.ownerDocument`.
 * Laying out on each tile's own context would create (and grow) that pair per
 * tile, outside admission. The viewer instead lays records out on ONE admitted
 * pair it owns (createAdmittedLayoutContext):
 * - both contexts are admitted (TILE_CANVAS_BYTES each) before they exist;
 * - each measures only through its own instance wrapper, which admits every
 *   distinct (font, text) with the shapeCacheBytes model before shaping it;
 * - the layout's local identity context IS the admitted local context: the
 *   base is presented through a facade whose canvas hands it out, so no other
 *   context is created and no global (prototype) state is ever touched;
 * - a refusal throws LayoutAdmissionError (the painter's coarse tile, the find
 *   overlay's whole-cell highlight);
 * - trim() retires the pair once its caches pass a cap: every host built on it
 *   is dropped with it (hosts handed out never capture a context), so its
 *   charge is returned only when nothing references the pair;
 * - a font publication (setFontGen) drops the hosts and their width caches.
 */
import type { XlsxSheet } from '../../xlsx/types'
import { createCellLayoutHost } from '../../xlsx/view/renderViewport'
import { LayoutAdmissionError, type CellLayoutHost, type LayoutAdmission } from '../../xlsx/view/textLayout'
import { TILE_CANVAS_BYTES, shapeCacheBytes } from '../../xlsx/view/mainLedger'

/** Retained shaping caches beyond this are dropped by trim(). */
export const MEASURE_CONTEXT_CAP = 4 * 1024 * 1024

/** A base measurement context (and the local identity context the canonical
 * layout derives from it), each measuring only through `admitShape` first.
 * The caller admits both contexts (2 × TILE_CANVAS_BYTES) before calling.
 * Null without a canvas-capable document. */
export function createAdmittedLayoutContext(admitShape: (cache: 'base' | 'local', font: string, text: string) => void): CanvasRenderingContext2D | null {
  const doc = globalThis.document
  if (!doc) return null
  const make = (cache: 'base' | 'local'): CanvasRenderingContext2D | null => {
    let ctx: CanvasRenderingContext2D | null = null
    try {
      const canvas = doc.createElement('canvas')
      // Measurement inherits the same language as the painting tiles
      // (createCanvasTile), so generic/fallback faces resolve alike.
      const lang = doc.documentElement?.getAttribute('lang')
      if (lang) canvas.setAttribute('lang', lang)
      canvas.width = 1
      canvas.height = 1
      ctx = canvas.getContext('2d')
    } catch { ctx = null }
    if (!ctx) return null
    const owned = ctx, native = owned.measureText.bind(owned)
    owned.measureText = (text: string) => { admitShape(cache, owned.font, text); return native(text) }
    return owned
  }
  const base = make('base'), local = make('local')
  if (!base || !local) return null
  // core/text-metrics creates its local context from base.canvas.ownerDocument:
  // the facade's canvas hands it the admitted local context instead.
  const localCanvas = { width: 4, height: 4, getContext: () => local }
  // `closest` lets core/text-metrics inherit the document language, as a
  // painted page canvas would (visual review #16).
  const canvas = { width: 1, height: 1, ownerDocument: { createElement: () => localCanvas }, closest: (selector: string) => doc.documentElement?.closest?.(selector) ?? null }
  return new Proxy(base, {
    get(target, property) {
      if (property === 'canvas') return canvas
      const value = Reflect.get(target, property, target) as unknown
      return typeof value === 'function' ? (value as (...args: unknown[]) => unknown).bind(target) : value
    },
    set(target, property, value) { return Reflect.set(target, property, value, target) },
  })
}

export class SharedMeasureContext {
  private ctx: CanvasRenderingContext2D | null = null
  private hosts = new WeakMap<object, CellLayoutHost>()
  private fontGen: number | undefined
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
  private context(): CanvasRenderingContext2D {
    if (this.ctx) return this.ctx
    this.grow(2 * TILE_CANVAS_BYTES)
    const ctx = createAdmittedLayoutContext((cache, font, text) => this.shape(cache, font, text))
    if (!ctx) throw new LayoutAdmissionError()
    return this.ctx = ctx
  }
  /** A font publication: hosts (and their measured-width caches) rebuild. */
  setFontGen(fontGen: number): void {
    if (fontGen === this.fontGen) return
    this.fontGen = fontGen
    this.hosts = new WeakMap()
  }
  /** Retires the pair once its caches pass the cap (between frames). */
  trim(): void {
    if (this.charged <= this.cap) return
    // Hosts built on the pair go with it: nothing references it afterwards.
    this.hosts = new WeakMap()
    this.ctx = null
    this.caches.clear()
    this.admission().measureChanged?.(-this.charged)
    this.charged = 0
  }
  private inner(sheet: Partial<XlsxSheet>): CellLayoutHost {
    const ctx = this.context()
    let host = this.hosts.get(sheet)
    if (!host) this.hosts.set(sheet, host = createCellLayoutHost(ctx, sheet))
    return host
  }
  /** A layout host for `sheet`. It never captures a context: each layout uses
   * the current admitted pair (created and admitted on first use). */
  host(sheet: Partial<XlsxSheet>): CellLayoutHost {
    // Formatting and overflow certification need no context.
    const pure = createCellLayoutHost(undefined as unknown as CanvasRenderingContext2D, sheet)
    return {
      layoutBody: (cell, width, height, overflow, preformatted, allocation) => this.inner(sheet).layoutBody(cell, width, height, overflow, preformatted, allocation),
      canReuseOverflow: (body, overflow) => pure.canReuseOverflow!(body, overflow),
      paintedText: (cell, width, admission) => pure.paintedText!(cell, width, admission),
    }
  }
}
