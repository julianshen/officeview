/**
 * Piece 5 Task 11 (spec §4.4 "Presenter View"): audio plays only in the audience window, as in local Presenter View.
 * The remote show adopts a presenter popup (a fake window over a second document, as
 * `slideshow-presenter-window.test.tsx`) and plays the recorded narration deck: the one media element is in the
 * audience surface, never in the presenter's document, in the window and in the single-screen swap alike.
 */
import { afterEach, beforeEach, describe, expect } from 'vitest'
import { elements, it, mediaDoc, mediaFixture, run, setUp, show, tearDown } from './remote-show-media-rig'

interface FakePopup extends Window { closeCalls: number }
function fakePopup(): FakePopup {
  const t = new EventTarget() as unknown as FakePopup & { closed: boolean }
  Object.assign(t, {
    document: document.implementation.createHTMLDocument(''),
    closed: false,
    closeCalls: 0,
    close() {
      t.closeCalls++
      if (t.closed) return
      t.closed = true
      t.dispatchEvent(new Event('pagehide'))
    },
  })
  return t
}

beforeEach(() => setUp())
afterEach(() => tearDown())

const media = (d: ParentNode) => [...d.querySelectorAll('audio, video')]

describe('remote show media and Presenter View (Task 11)', () => {
  it('presenter window creates no audio element; audio plays only in the audience window', async () => {
    const fd = mediaDoc(mediaFixture('narration'), { notes: true })
    const popup = fakePopup()
    const { s, root } = await show(fd, { popup, presenterView: true } as never)
    await run(400)
    expect(s.presenter).toBe('window')
    expect(popup.document.querySelector('[data-officeview-presenter-view]')).not.toBeNull()
    expect(media(popup.document)).toEqual([])
    const [el] = media(root)
    expect(el).toBeDefined()
    expect(el!.ownerDocument).toBe(document)
    expect(elements().playing()).toEqual([el])
    expect(elements().plays.every((p) => p.el.ownerDocument === document && root.contains(p.el))).toBe(true)
    s.exit()
  }, 15_000)

  it('single-screen swap: the audio stays on the hidden audience surface; Presenter View holds none', async () => {
    const fd = mediaDoc(mediaFixture('narration'), { notes: true })
    const { s, root, host } = await show(fd, { popup: null })
    await run(400)
    expect(s.presenter).toBe('swap')
    const swap = host.querySelector('[data-officeview-presenter-swap]')!
    expect(swap).not.toBeNull()
    expect(media(swap)).toEqual([])
    expect(media(root)).toHaveLength(1)
    expect(elements().playing()).toEqual(media(root))
    s.exit()
  }, 15_000)
})
