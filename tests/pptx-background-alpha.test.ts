/**
 * tdf146223: slide 1's background is a stretched 2 × 2 picture with `a:alphaModFix amt="70000"`.
 * PowerPoint composites it at 70% over the white slide (probe deck: the same picture is 206 green
 * at full opacity and 221 at 70%, i.e. 0.7 × 206 + 0.3 × 255); we painted it fully opaque.
 */
import { existsSync, readFileSync } from 'node:fs'
import { afterEach, describe, expect, it } from 'vitest'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

const DECK = 'corpus/lo-tdf146223.pptx'
let seams: ShowSeams | undefined
afterEach(() => { seams?.restore(); seams = undefined; document.body.replaceChildren() })

describe.skipIf(!existsSync(DECK))('translucent picture background', () => {
  it('tdf146223 composites its 70% background picture over white', async () => {
    const size = { width: 960, height: 720 }
    seams = installShowSeams(size)
    const r = await startShow(await loadDeck(new Uint8Array(readFileSync(DECK))), { seams, size })
    await r.tick(5000)
    const c = r.frameCanvas()!
    // The picture's bottom-right pixel is pure red; at 70% over white no channel drops below 0.3 × 255.
    const [rr, gg, bb] = Array.from(c.getContext('2d')!.getImageData(size.width - 2, size.height - 2, 1, 1).data)
    expect(rr).toBeGreaterThan(240)
    expect(gg).toBeGreaterThanOrEqual(70)
    expect(bb).toBeGreaterThanOrEqual(70)
  }, 30_000)
})
