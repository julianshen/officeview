/**
 * tdf146223: both slides' backgrounds are stretched 2 × 2 JPEGs (slide 2 under a master picture that is one
 * too). PowerPoint shows each slide as one flat colour (D4 reference: slide 2 is (35, 35, 35) everywhere);
 * we smoothed the pictures into gradients across the slide.
 */
import { existsSync, readFileSync } from 'node:fs'
import { afterEach, describe, expect, it } from 'vitest'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

const DECK = 'corpus/lo-tdf146223.pptx'
let seams: ShowSeams | undefined
afterEach(() => { seams?.restore(); seams = undefined; document.body.replaceChildren() })

describe.skipIf(!existsSync(DECK))('stretched 2 × 2 pictures', () => {
  it('tdf146223 paints its 2 × 2 picture backgrounds as flat colours', async () => {
    const size = { width: 960, height: 720 }
    seams = installShowSeams(size)
    const r = await startShow(await loadDeck(new Uint8Array(readFileSync(DECK))), { seams, size })
    await r.tick(5000)
    const sample = () => {
      const ctx = r.frameCanvas()!.getContext('2d')!
      // Clear of the title, rule and bullets.
      return [[0.5, 0.6], [0.95, 0.95], [0.05, 0.95], [0.95, 0.45]].map(([fx, fy]) => Array.from(ctx.getImageData(Math.round(fx * (size.width - 1)), Math.round(fy * (size.height - 1)), 1, 1).data).slice(0, 3))
    }
    const one = sample()
    for (const p of one) expect(p).toEqual(one[0])
    await r.key('ArrowRight')
    await r.tick(5000)
    const two = sample()
    for (const p of two) expect(p).toEqual(two[0])
    // The master's grey picture, JPEG-weighted: about 37 (PowerPoint 35).
    for (const v of two[0]) expect(Math.abs(v - 36)).toBeLessThanOrEqual(3)
  }, 30_000)
})
