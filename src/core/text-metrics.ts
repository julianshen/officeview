/** Identity-coordinate text metrics shared by drawing layout and text recording. */
const measurers = new WeakMap<CanvasRenderingContext2D, (text: string) => TextMetrics>()

/**
 * Cairo-backed Canvas measureText depends on its active transform. Measure on a
 * separate identity Canvas so layout uses local font metrics before placement.
 * The caller's font, transform and painting state are never changed.
 */
export function createLocalTextMeasurer(base: CanvasRenderingContext2D): (text: string) => TextMetrics {
  const cached = measurers.get(base)
  if (cached) return cached
  let canvas: HTMLCanvasElement
  const ownerDocument = base.canvas.ownerDocument ?? (typeof document !== 'undefined' ? document : undefined)
  if (ownerDocument) canvas = ownerDocument.createElement('canvas')
  else {
    // Both OffscreenCanvas and node-canvas accept width/height without a DOM.
    const Canvas = base.canvas.constructor as unknown as new (width: number, height: number) => HTMLCanvasElement
    canvas = new Canvas(4, 4)
  }
  canvas.width = 4
  canvas.height = 4
  let context = canvas.getContext('2d') as unknown as CanvasRenderingContext2D | null
  if (!context) {
    // A host `document` can produce a canvas whose 2d context is null
    // (jsdom-without-canvas vs a real node-canvas base). Fall back to the
    // ACTUAL constructor of the base canvas, which always measures with the
    // same engine as the painting context.
    const Canvas = base.canvas.constructor as unknown as new (width: number, height: number) => HTMLCanvasElement
    canvas = new Canvas(4, 4)
    canvas.width = 4
    canvas.height = 4
    context = canvas.getContext('2d') as unknown as CanvasRenderingContext2D | null
  }
  if (!context) throw new Error('No local text measurement context available')
  const measure = (text: string): TextMetrics => {
    context.font = base.font
    // Copy current font shaping/spacing settings supported by this implementation.
    for (const name of ['lang', 'direction', 'fontKerning', 'fontStretch', 'fontVariantCaps', 'letterSpacing', 'wordSpacing', 'textRendering']) {
      const value = (base as unknown as Record<string, unknown>)[name]
      if (typeof value === 'string' && name in context) {
        // Some Canvas hosts serialize effective zero as an empty string (or
        // "normal") but ignore those values when assigned to a reused context.
        // Explicit zero clears tracking left by the previous measured run.
        const spacing = name === 'letterSpacing' || name === 'wordSpacing'
        ;(context as unknown as Record<string, unknown>)[name] = spacing && (value === '' || value === 'normal') ? '0px' : value
      }
    }
    return context.measureText(text)
  }
  measurers.set(base, measure)
  return measure
}
