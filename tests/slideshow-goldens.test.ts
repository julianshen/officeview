/**
 * Slideshow frame goldens (slideshow spec 10; plan S1-15): selected 2-D
 * transitions rendered through `transitionDraw` and the Canvas 2D compositor
 * (node-canvas) at 0/25/50/75/100 %, diffed against
 * tests/goldens/slideshow/*.png with the catalogue row's tolerance.
 *
 * Mac only, like the other pixel goldens: CI (Linux) skips the file, because
 * node-canvas rasterises anti-aliased edges differently there. A missing
 * golden is recorded on first run; OFFICEVIEW_STRICT_GOLDEN=1 fails instead.
 * The frames are "Office-shaped synthetic", never Office evidence. GPU
 * (WebGL2) frames are compared against Canvas 2D in Chromium by
 * tests/browser/slideshow-gpu.spec.ts.
 */
import { describe, expect, it } from 'vitest'
import { existsSync, mkdirSync } from 'fs'
import { dirname, join } from 'path'
import { fileURLToPath } from 'url'
import { createCanvas } from 'canvas'
import { rowForTransition } from '../src/slideshow/catalogue'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createRng } from '../src/slideshow/rng'
import { createCanvas2DCompositor } from '../src/slideshow/render/canvas2d'
import { transitionDraw } from '../src/slideshow/render/transitions2d'
import { frameDiff, withinTolerance } from '../src/debug/slideshow-parity'
import { loadPng, savePng } from '../src/test/pixel-diff'
import type { CanvasFactory, LayerHandle, Lease, SlideShowWorkspace, TransitionSpec } from '../src/slideshow/types'

const GOLDEN_DIR = join(dirname(fileURLToPath(import.meta.url)), 'goldens', 'slideshow')
const strict = process.env.OFFICEVIEW_STRICT_GOLDEN === '1'

// A 160 × 90 slide on a 176 × 110 output, letterboxed by 8 / 10 device px.
const SLIDE_W = 160, SLIDE_H = 90, DEV_W = 176, DEV_H = 110
const VIEW = [1, 0, 0, 1, 8, 10] as const

function workspace(): SlideShowWorkspace {
  let used = 0
  const lease = (bytes: number, tag: string): Lease => {
    used += bytes
    let held = bytes, live = true
    const l: Lease = {
      tag, get bytes() { return live ? held : 0 },
      grow(n) { used += n; held += n; return true },
      shrink(n) { const k = Math.min(n, held); held -= k; used -= k },
      transfer: () => l,
      release() { if (live) { live = false; used -= held } },
    }
    return l
  }
  return { limit: 1 << 30, lease, used: () => used, reserveTerminal: () => undefined }
}
const factory: CanvasFactory = (w, h) => createCanvas(w, h) as unknown as HTMLCanvasElement

/** A slide frame with asymmetric content, so a flip or a wrong direction shows. */
function slideFrame(kind: 'from' | 'to') {
  const c = createCanvas(SLIDE_W, SLIDE_H)
  const x = c.getContext('2d')
  const g = x.createLinearGradient(0, 0, SLIDE_W, SLIDE_H)
  g.addColorStop(0, kind === 'from' ? '#204080' : '#a02020')
  g.addColorStop(1, kind === 'from' ? '#80c0ff' : '#ffd060')
  x.fillStyle = g
  x.fillRect(0, 0, SLIDE_W, SLIDE_H)
  x.fillStyle = '#10e010'
  x.fillRect(12, 10, 40, 18)
  x.fillStyle = '#ffffff'
  x.beginPath(); x.arc(112, 64, 16, 0, Math.PI * 2); x.fill()
  return c
}

function render(spec: TransitionSpec, progress: number): Uint8ClampedArray {
  const ws = workspace()
  const out = createCanvas(DEV_W, DEV_H)
  const sink = createDiagnosticSink()
  const comp = createCanvas2DCompositor(out as unknown as HTMLCanvasElement, factory, ws, sink)
  comp.setViewport({ matrix: VIEW, deviceWidth: DEV_W, deviceHeight: DEV_H })
  const upload = (kind: 'from' | 'to'): LayerHandle => {
    const r = comp.upload({ canvas: slideFrame(kind) as unknown as HTMLCanvasElement, bounds: { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H }, scale: 1, bytes: SLIDE_W * SLIDE_H * 4 }, ws.lease(SLIDE_W * SLIDE_H * 4, 'layer')!)
    if ('refused' in r) throw new Error(r.refused)
    return r
  }
  const from = upload('from'), to = upload('to')
  const rng = createRng(2026)
  comp.frame([], { from, to, effect: spec.element, progress, params: transitionDraw(spec, progress, rng, SLIDE_W / SLIDE_H), rng })
  expect(sink.list()).toEqual([])
  // Over an opaque backdrop, so the PNG holds no premultiplied alpha.
  const flat = createCanvas(DEV_W, DEV_H)
  const fx = flat.getContext('2d')
  fx.fillStyle = '#202020'
  fx.fillRect(0, 0, DEV_W, DEV_H)
  fx.drawImage(out, 0, 0)
  comp.dispose()
  return fx.getImageData(0, 0, DEV_W, DEV_H).data
}

const spec = (element: string, attrs: Record<string, string> = {}): TransitionSpec => ({ element, ns: 'p', attrs, durationMs: 500, advClick: true })
const CASES: Array<{ name: string; spec: TransitionSpec }> = [
  { name: 'fade', spec: spec('fade') },
  { name: 'push', spec: spec('push', { dir: 'u' }) },
  { name: 'wipe', spec: spec('wipe', { dir: 'r' }) },
  { name: 'split', spec: spec('split', { orient: 'vert', dir: 'in' }) },
  { name: 'circle', spec: spec('circle') },
  { name: 'dissolve', spec: spec('dissolve') },
  { name: 'wheel', spec: spec('wheel', { spokes: '3' }) },
]
const PROGRESS = [0, 25, 50, 75, 100]

describe.skipIf(process.platform !== 'darwin')('slideshow transition frame goldens (Mac)', () => {
  for (const c of CASES) {
    for (const pct of PROGRESS) {
      it(`${c.name} at ${pct} %`, async () => {
        const pixels = render(c.spec, pct / 100)
        const path = join(GOLDEN_DIR, `transition-${c.name}.${pct}.png`)
        if (!existsSync(path)) {
          if (strict) throw new Error(`missing golden ${path} (strict mode); record it first`)
          mkdirSync(GOLDEN_DIR, { recursive: true })
          await savePng({ width: DEV_W, height: DEV_H, data: pixels }, path)
          return
        }
        const golden = await loadPng(path)
        expect([golden.width, golden.height]).toEqual([DEV_W, DEV_H])
        const tolerance = rowForTransition('p', c.spec.element)!.tolerance
        const d = frameDiff(pixels, golden.data, DEV_W, DEV_H)
        expect(withinTolerance(d, tolerance), `${c.name} @${pct}: meanAbs ${d.meanAbs.toFixed(4)}, rowMax ${d.rowMax.toFixed(4)}`).toBe(true)
      })
    }
  }
})
