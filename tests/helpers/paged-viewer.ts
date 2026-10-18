/**
 * jsdom helpers for the paged viewer tests: a layout mock for the scroll
 * container (jsdom has no layout), a coarse-pointer stub, synthetic pages for
 * a mocked getPaintables, and a hand-built DocxDocument.
 */
import type { DocxDocument } from '../../src/docx/types'
import type { Paintable, PaintableArray } from '../../src/render/paint'

export interface FakePage { w: number; h: number; text?: string }

/** Synthetic paintables: each page paints its text (so search finds it). */
export function fakePaintables(pages: readonly FakePage[]): PaintableArray {
  const list: Paintable[] = pages.map((p) => ({
    spec: { widthPx: p.w, heightPx: p.h },
    paint: (ctx: CanvasRenderingContext2D) => { if (p.text) { ctx.font = '16px Arial'; ctx.fillText(p.text, 40, 60) } },
  }))
  return Object.assign(list, { dispose: () => {}, fontDiagnostics: [] })
}

/** The region's clientWidth/clientHeight (all other elements report 0). */
export function mockViewerLayout(size: { width: number; height: number }): () => void {
  const proto = HTMLElement.prototype
  const saved = ['clientWidth', 'clientHeight'].map((k) => [k, Object.getOwnPropertyDescriptor(proto, k)] as const)
  const isRegion = (el: HTMLElement) => el.getAttribute?.('aria-label') === 'Document viewer'
  Object.defineProperty(proto, 'clientWidth', { configurable: true, get(this: HTMLElement) { return isRegion(this) ? size.width : 0 } })
  Object.defineProperty(proto, 'clientHeight', { configurable: true, get(this: HTMLElement) { return isRegion(this) ? size.height : 0 } })
  return () => { for (const [k, d] of saved) { if (d) Object.defineProperty(proto, k, d); else delete (proto as unknown as Record<string, unknown>)[k] } }
}

/** A matchMedia stub; `coarse` decides (pointer: coarse). */
export function stubPointer(coarse: boolean): { set(v: boolean): void; restore(): void } {
  const listeners = new Set<() => void>()
  let matches = coarse
  const had = 'matchMedia' in window
  const prev = window.matchMedia
  Object.defineProperty(window, 'matchMedia', {
    configurable: true,
    writable: true,
    value: (q: string) => ({
      get matches() { return q.includes('coarse') ? matches : false },
      media: q,
      addEventListener: (_: string, cb: () => void) => listeners.add(cb),
      removeEventListener: (_: string, cb: () => void) => listeners.delete(cb),
    }),
  })
  return {
    set(v) { matches = v; for (const l of [...listeners]) l() },
    restore() { if (had) window.matchMedia = prev; else delete (window as unknown as Record<string, unknown>).matchMedia },
  }
}

/** A hand-built Word model: one paragraph per entry. */
export function docxModel(paragraphs: readonly string[]): DocxDocument {
  return {
    sections: [{
      margins: {}, pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' }, paragraphs: [],
      blocks: paragraphs.map((text) => ({ kind: 'p', paragraph: { runs: [{ text }], images: [], align: 'left' } })),
    }],
    defaultFontFamily: 'Arial', defaultFontSizePt: 11, styleDefaults: new Map(),
  } as unknown as DocxDocument
}
