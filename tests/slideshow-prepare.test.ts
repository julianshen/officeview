/**
 * Piece 4 Task 1: the local stage prepares its slides through the shared,
 * DOM-free `prepareSlide` (src/slideshow/prepare). A passthrough mock counts
 * the calls; the real show runs through the driven rig.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'

vi.mock('../src/slideshow/prepare/slide', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/prepare/slide')>()
  return { ...real, prepareSlide: vi.fn(real.prepareSlide) }
})

import { prepareSlide } from '../src/slideshow/prepare/slide'
import { flyInDeck } from '../src/testdata/show-decks'
import { installShowSeams, loadDeck, startShow, type ShowRig, type ShowSeams } from './helpers/slideshow-show-rig'

let seams: ShowSeams | undefined
let rig: ShowRig | undefined
afterEach(() => {
  rig?.exit(); rig = undefined
  seams?.restore(); seams = undefined
  document.body.replaceChildren()
  vi.mocked(prepareSlide).mockClear()
})

describe('slideshow prepare', () => {
  it('stage.prepare is prepareSlide', async () => {
    seams = installShowSeams({ width: 960, height: 720 })
    rig = await startShow(await loadDeck(await flyInDeck()), { seams, size: { width: 960, height: 720 } })
    await rig.tick(0)
    expect(vi.mocked(prepareSlide)).toHaveBeenCalledTimes(1)
    expect(vi.mocked(prepareSlide).mock.calls[0][0]).toMatchObject({ index: 0, watermark: false })
    // The box flies in on click 1: hidden until then, then shown.
    expect(rig.state().get('sp:3')?.visible).toBe(false)
    await rig.key('ArrowRight')
    await rig.tick(1000)
    expect(rig.state().get('sp:3')?.visible).not.toBe(false)
  })
})
