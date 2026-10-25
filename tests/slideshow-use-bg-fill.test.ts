/**
 * A `p:sp@useBgFill` shape shows the slide background in the show, as it does in the viewer.
 *
 * tdf93868's master lays a white full-slide rectangle under a rounded rectangle filled with the
 * slide background (a black-to-grey gradient), and the slide's text is white. The show painted the
 * rounded rectangle without the background, so the frame came out white (D4 compare, PR #30).
 */
import { existsSync, readFileSync } from 'node:fs'
import { createCanvas } from 'canvas'
import { afterEach, describe, expect, it } from 'vitest'
import { paintDirect } from '../src/slideshow/render/direct'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

const DECK = 'corpus/lo-tdf93868.pptx'
let seams: ShowSeams | undefined
afterEach(() => { seams?.restore(); seams = undefined; document.body.replaceChildren() })

describe.skipIf(!existsSync(DECK))('slideshow useBgFill', () => {
  it('tdf93868: the master shape filled with the slide background paints the background', async () => {
    const size = { width: 960, height: 720 }
    seams = installShowSeams(size)
    const r = await startShow(await loadDeck(new Uint8Array(readFileSync(DECK))), { seams, size })
    await r.tick(100)
    const c = r.frameCanvas()!
    const px = (x: number, y: number) => Array.from(c.getContext('2d')!.getImageData(x, y, 1, 1).data)
    // The lower half of the slide is the gradient's black stop, inside the rounded rectangle.
    const [rr, gg, bb] = px(480, 600)
    expect(Math.max(rr, gg, bb)).toBeLessThan(16)
    // Outside the rounded rectangle's corner the white rectangle beneath still shows.
    expect(px(1, 1).slice(0, 3).every(v => v > 200)).toBe(true)
  }, 30_000)

  it('tdf93868: the direct painter, at a device scale, paints the background through the shape', async () => {
    const slide = (await loadDeck(new Uint8Array(readFileSync(DECK)))).slides[0]
    // 720 × 540 slide px painted at scale 2 on a context already translated by (10, 10).
    const c = createCanvas(1460, 1100)
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.translate(10, 10)
    paintDirect(ctx, slide, new Map(), { images: [], resolveFont: f => f, scale: 2, scratch: { lease: () => undefined } }) // the deck has no images
    const [rr, gg, bb] = Array.from(c.getContext('2d').getImageData(730, 910, 1, 1).data)
    expect(Math.max(rr, gg, bb)).toBeLessThan(16)
    // The gradient's grey stop at the top, also inside the shape.
    const top = Array.from(c.getContext('2d').getImageData(730, 40, 1, 1).data)
    expect(top[0]).toBeGreaterThan(40)
    expect(top[0]).toBeLessThan(80)
  }, 30_000)
})
