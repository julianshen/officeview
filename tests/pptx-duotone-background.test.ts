/**
 * tdf123684: the theme background is a 64 × 64 black-and-white picture recoloured with `a:duotone`
 * between two shades of the background colour (bg2, a dark teal). PowerPoint shows a dark navy slide
 * with a teal-green chamfered panel (D4 reference: about (9, 48, 62) and (81, 148, 123)); without
 * the duotone the picture painted black and white.
 */
import { existsSync, readFileSync } from 'node:fs'
import { afterEach, describe, expect, it } from 'vitest'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

const DECK = 'corpus/lo-tdf123684.pptx'
let seams: ShowSeams | undefined
afterEach(() => { seams?.restore(); seams = undefined; document.body.replaceChildren() })

describe.skipIf(!existsSync(DECK))('duotone picture background', () => {
  it('tdf123684 paints its theme picture through the duotone', async () => {
    const size = { width: 960, height: 540 }
    seams = installShowSeams(size)
    const r = await startShow(await loadDeck(new Uint8Array(readFileSync(DECK))), { seams, size })
    await r.tick(100)
    const c = r.frameCanvas()!
    const px = (fx: number, fy: number) => Array.from(c.getContext('2d')!.getImageData(Math.round(fx * size.width), Math.round(fy * size.height), 1, 1).data)
    // The panel (white in the picture), also seen through the useBgFill text box: teal-green.
    for (const [fx, fy] of [[0.45, 0.55], [0.3, 0.3]]) {
      const [rr, gg, bb] = px(fx, fy)
      expect(gg, `${fx},${fy}`).toBeGreaterThan(rr + 30)
      expect(gg, `${fx},${fy}`).toBeGreaterThan(bb)
      expect(Math.max(rr, gg, bb), `${fx},${fy}`).toBeLessThan(220)
    }
    // Outside the panel (black in the picture): dark navy, not black.
    const [rr, gg, bb] = px(0.02, 0.04)
    expect(bb).toBeGreaterThan(rr + 20)
    expect(Math.max(rr, gg, bb)).toBeLessThan(110)
  }, 30_000)
})
