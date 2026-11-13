/**
 * Stage 3c Task 8: presenting a streamed deck in a real browser against
 * server/test/app.ts (`?role=presenter` makes the viewer a presenter;
 * `?notes=off` sets the session's notes policy to false; absent = the demo's
 * audience role). Deck: lo-tdf104788.pptx (6 slides, each with speaker notes).
 * The show starts from the harness page's own buttons (src/debug/stream-viewer.tsx),
 * which call the ref's present() inside a real click; the outcome lands in
 * window.__ovStream.present ('started' | 'ended' | an error code).
 *
 * An init script makes the browser parts the test cannot drive deterministic:
 * - `window.open` is recorded and returns null (a blocked popup), so Presenter
 *   View runs single-screen, swapped into the show host (`presenter: 'swap'`);
 * - `Element.prototype.requestFullscreen` is recorded and rejects, so the show
 *   runs in the window (headless Chromium's full screen is not the product's);
 * - every WebSocket's close code is recorded (Playwright's own websocket
 *   `close` event carries no code), so "no 4008" is a real assertion.
 * Frames sent are counted from page load: the server's 60 unit requests/min
 * window is per session and includes the viewer's own paints.
 */
import { expect, test, type Page } from '@playwright/test'

const P = Number(process.env.OV_STREAM_PORT ?? 4181)
const CONTROL = `http://127.0.0.1:${P + 1}`
const DECK = 'doc=lo-tdf104788.pptx'
const open = (page: Page, q: string) => page.goto(`/src/debug/stream-viewer.html?port=${P}&${DECK}${q ? `&${q}` : ''}`)
/** The audience surface's root (Presenter View's own canvases carry `data-officeview-show-surface="presenter"`). */
const surface = (page: Page) => page.locator('[data-officeview-show-surface=""]')
const live = (page: Page) => page.locator('[data-officeview-show-live]')
const swap = (page: Page) => page.locator('[data-officeview-presenter-swap]')
const presenterView = (page: Page) => page.locator('[data-officeview-presenter-view]')
const notesCanvas = (page: Page) => page.locator('canvas[aria-label="Speaker notes (image)"]')
const presentState = (page: Page) => page.evaluate(() => (window as unknown as { __ovStream: { present?: string } }).__ovStream.present)
const streamError = (page: Page) => page.evaluate(() => (window as unknown as { __ovStream: { error?: string } }).__ovStream.error)

type Rec = { opens: number; fullscreen: number; closes: number[] }
const rec = (page: Page) => page.evaluate(() => (window as unknown as { __ovPresent: Rec }).__ovPresent)

/** Client → server ops by type, with their send times; refusals received. */
function wire(page: Page) {
  const sent: Array<{ type: string; at: number }> = []
  const refused: string[] = []
  const texts: string[] = []
  page.on('websocket', (ws) => {
    ws.on('framesent', (f) => {
      if (typeof f.payload !== 'string') return
      const m = /"type":"([a-z-]+)"/.exec(f.payload)
      if (m) sent.push({ type: m[1]!, at: Date.now() })
    })
    ws.on('framereceived', (f) => {
      if (typeof f.payload !== 'string') return
      texts.push(f.payload)
      if (/"type":"(refused|error)"/.test(f.payload)) refused.push(f.payload)
    })
  })
  return {
    sent,
    refused,
    /** Every text message received (binary frames are the marked PNG rasters). */
    texts,
    count: (type: string) => sent.filter((s) => s.type === type).length,
    /** The most `type` ops sent inside any `ms` window. */
    peak(type: string, ms: number) {
      const t = sent.filter((s) => s.type === type).map((s) => s.at)
      let best = 0
      for (let i = 0, j = 0; i < t.length; i++) {
        while (t[i]! - t[j]! >= ms) j++
        best = Math.max(best, i - j + 1)
      }
      return best
    },
  }
}

test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    const r = { opens: 0, fullscreen: 0, closes: [] as number[] }
    ;(window as unknown as { __ovPresent: typeof r }).__ovPresent = r
    window.open = () => { r.opens++; return null }
    Element.prototype.requestFullscreen = function () { r.fullscreen++; return Promise.reject(new Error('full screen is stubbed')) }
    const WS = window.WebSocket
    window.WebSocket = class extends WS {
      constructor(url: string | URL, protocols?: string | string[]) {
        super(url, protocols)
        this.addEventListener('close', (e) => { r.closes.push(e.code) })
      }
    }
  })
})

/** The stream viewer is up: every unit laid out, the first painted. */
async function ready(page: Page) {
  await expect(page.locator('[data-unit-wrap]')).toHaveCount(6)
  await expect(page.locator('canvas[data-painted="1"]').first()).toBeVisible()
}

/** The audience frame shows streamed slide `n` (1-based) as an image. */
async function onSlide(page: Page, n: number) {
  await expect(live(page)).toHaveText(`Slide ${n} of 6`)
  await expect(surface(page)).toHaveAttribute('data-slide', String(n - 1))
  await expect(surface(page)).toHaveAttribute('data-officeview-show-frame', 'image')
}

test('audience: present, step all 6 slides with →, go to a slide by number + Enter, Esc exits; no notes, no popup', async ({ page }) => {
  const w = wire(page)
  await open(page, '')
  await ready(page)
  await page.locator('[data-debug-present]').click()
  await expect.poll(() => presentState(page)).toBe('started')
  // The surface opened inside the click: full screen was asked for in the gesture; no presenter window for an audience viewer.
  expect((await rec(page)).fullscreen).toBe(1)
  expect((await rec(page)).opens).toBe(0)
  await onSlide(page, 1)
  for (let n = 2; n <= 6; n++) {
    await page.keyboard.press('ArrowRight')
    await onSlide(page, n)
  }
  await page.keyboard.press('ArrowRight')
  await expect(page.locator('[data-officeview-show-end]')).toBeVisible()
  await expect(live(page)).toHaveText('End of slide show, click to exit.')
  // goTo by number: digits then Enter (1-based on screen).
  await page.keyboard.press('3')
  await page.keyboard.press('Enter')
  await onSlide(page, 3)
  // Alt+F5 swaps in Presenter View single-screen even for an audience viewer: it shows no notes and asks for none.
  await page.keyboard.press('Alt+F5')
  await expect(presenterView(page)).toBeVisible()
  await expect(page.locator('[data-officeview-presenter-status]')).toHaveText('Slide 3 of 6')
  await expect(page.locator('[data-officeview-presenter-notes-empty]')).toHaveText('No notes for this slide')
  await page.waitForTimeout(800)
  await expect(notesCanvas(page)).toHaveCount(0)
  await page.keyboard.press('Alt+F5')
  await expect(presenterView(page)).toHaveCount(0)
  await page.keyboard.press('Escape')
  await expect.poll(() => presentState(page)).toBe('ended')
  await expect(surface(page)).toHaveCount(0)
  expect(w.count('notes')).toBe(0)
  expect(await streamError(page)).toBeUndefined()
  expect((await rec(page)).closes).toEqual([])
})

test('presenter: Presenter View swaps in single-screen with a painted notes image; the audience surface never holds notes', async ({ page }) => {
  const w = wire(page)
  await open(page, 'role=presenter')
  await ready(page)
  await page.locator('[data-debug-presenter]').click()
  await expect.poll(() => presentState(page)).toBe('started')
  // presenterView: true opens the provisional window in the gesture; the stub blocks it, so Presenter View swaps in.
  expect((await rec(page)).opens).toBe(1)
  await expect(swap(page).locator('[data-officeview-presenter-view]')).toBeVisible()
  await expect(surface(page)).toBeHidden()
  await expect(page.locator('[data-officeview-presenter-status]')).toHaveText('Slide 1 of 6')
  // Slide 1's notes image: 720 CSS px × scale 1 (DPR 1) wide, in Presenter View (this slide's notes body is empty:
  // the image is white with the marks only, one empty-notes height).
  const notes = swap(page).locator('canvas[aria-label="Speaker notes (image)"]')
  await expect(notes).toHaveCount(1)
  await expect.poll(() => notes.evaluate((c: HTMLCanvasElement) => c.width)).toBe(720)
  expect(w.count('notes')).toBe(1)
  // The audience surface holds its one canvas only.
  await expect(surface(page).locator('canvas')).toHaveCount(1)
  await expect(surface(page).locator('canvas[aria-label="Speaker notes (image)"]')).toHaveCount(0)
  // Slide 5 by number from Presenter View: its notes (several lines of text) are asked for after the 300 ms
  // settle and painted, taller than the empty image and with dark text pixels.
  const emptyH = await notes.evaluate((c: HTMLCanvasElement) => c.height)
  await presenterView(page).focus()
  await page.keyboard.press('5')
  await page.keyboard.press('Enter')
  await expect(page.locator('[data-officeview-presenter-status]')).toHaveText('Slide 5 of 6')
  await expect.poll(() => w.count('notes')).toBe(2)
  await expect.poll(() => notes.evaluate((c: HTMLCanvasElement) => c.height)).toBeGreaterThan(emptyH)
  expect(await notes.evaluate((c: HTMLCanvasElement) => c.width)).toBe(720)
  const dark = await notes.evaluate((c: HTMLCanvasElement) => {
    const d = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data
    let n = 0
    for (let k = 0; k < d.length; k += 4) if (d[k]! < 128 && d[k + 1]! < 128 && d[k + 2]! < 128) n++
    return n
  })
  expect(dark).toBeGreaterThan(500)
  await expect(surface(page).locator('canvas[aria-label="Speaker notes (image)"]')).toHaveCount(0)
  // Back to the audience surface: it shows slide 5, and no notes canvas exists anywhere.
  await page.keyboard.press('Alt+F5')
  await expect(swap(page)).toHaveCount(0)
  await onSlide(page, 5)
  await expect(notesCanvas(page)).toHaveCount(0)
  await page.keyboard.press('Escape')
  await expect.poll(() => presentState(page)).toBe('ended')
  expect(w.refused).toEqual([])
  // Wire audit, browser side: slide 5's notes text (in the package only) is in no text message received.
  for (const words of ['Small Basic', 'different elements of the environment', 'multiple Editor']) {
    expect(w.texts.filter((t) => t.includes(words))).toEqual([])
  }
  expect(await streamError(page)).toBeUndefined()
  expect((await rec(page)).closes).toEqual([])
})

test('holding the arrow keys for 10 s keeps the socket open (no 4008) and the show on the slide stopped on', async ({ page }) => {
  test.setTimeout(60_000)
  const w = wire(page)
  await open(page, '')
  await ready(page)
  await page.locator('[data-debug-present]').click()
  await expect.poll(() => presentState(page)).toBe('started')
  await onSlide(page, 1)
  // Key-repeat cadence (~30 keydowns/s, `repeat` set while held). A held → past the last slide would
  // end the show (no loop on this deck), so → is held to slide 6, then ← back to 1, for 10 s.
  const t0 = Date.now()
  let at = 0, key: 'ArrowRight' | 'ArrowLeft' = 'ArrowRight', presses = 0
  while (Date.now() - t0 < 10_000) {
    if (key === 'ArrowRight' && at === 5) { await page.keyboard.up(key); key = 'ArrowLeft' }
    else if (key === 'ArrowLeft' && at === 0) { await page.keyboard.up(key); key = 'ArrowRight' }
    await page.keyboard.down(key)
    at += key === 'ArrowRight' ? 1 : -1
    presses++
    await page.waitForTimeout(33)
  }
  await page.keyboard.up(key)
  // Stop on slide 4: back to 1, then three →.
  for (; at > 0; at--) await page.keyboard.press('ArrowLeft')
  for (; at < 3; at++) await page.keyboard.press('ArrowRight')
  await onSlide(page, 4)
  await expect(surface(page)).toHaveAttribute('data-unit', '3')
  expect(presses).toBeGreaterThan(200)
  expect(await presentState(page)).toBe('started')
  expect(await streamError(page)).toBeUndefined()
  expect((await rec(page)).closes).toEqual([])
  expect(w.refused.filter((r) => r.includes('rate-limited'))).toEqual([])
  // The server's unit window is 60/min per session; the show budgets 48/min of it.
  expect(w.peak('unit', 60_000)).toBeLessThanOrEqual(60)
  console.log(`hold: ${presses} key events in 10 s; unit requests from page load ${w.count('unit')} (peak in any 60 s ${w.peak('unit', 60_000)}); rate refusals 0; close codes none`)
  await page.keyboard.press('Escape')
  await expect.poll(() => presentState(page)).toBe('ended')
})

test('notes=off presenter: presenterView: true is refused before anything opens; a plain show works and asks for no notes', async ({ page }) => {
  const w = wire(page)
  await open(page, 'role=presenter&notes=off')
  await ready(page)
  await page.locator('[data-debug-presenter]').click()
  await expect.poll(() => presentState(page)).toBe('not-presentable')
  expect(await rec(page)).toEqual({ opens: 0, fullscreen: 0, closes: [] })
  await expect(surface(page)).toHaveCount(0)
  await page.locator('[data-debug-present]').click()
  await expect.poll(() => presentState(page)).toBe('started')
  expect((await rec(page)).opens).toBe(0)
  await onSlide(page, 1)
  await page.keyboard.press('ArrowRight')
  await onSlide(page, 2)
  await page.keyboard.press('Alt+F5')
  await expect(page.locator('[data-officeview-presenter-notes-empty]')).toHaveText('No notes for this slide')
  await page.waitForTimeout(800)
  await expect(notesCanvas(page)).toHaveCount(0)
  await page.keyboard.press('Escape')
  await expect.poll(() => presentState(page)).toBe('ended')
  expect(w.count('notes')).toBe(0)
  expect((await rec(page)).closes).toEqual([])
})

test('revocation mid-show exits the show cleanly: ended resolves, the surface is gone, nothing renders locally', async ({ page }) => {
  await open(page, 'role=presenter')
  await ready(page)
  await page.locator('[data-debug-present]').click()
  await expect.poll(() => presentState(page)).toBe('started')
  await onSlide(page, 1)
  // Live sessions only: earlier pages leave detached ones for the resume grace.
  const sids = (await (await page.request.get(`${CONTROL}/sessions`)).json() as Array<{ sid: string; status: string }>).filter((s) => s.status === 'stream')
  expect(sids.length).toBeGreaterThan(0)
  for (const { sid } of sids) await page.request.post(`${CONTROL}/revoke?sid=${sid}`)
  await expect.poll(() => presentState(page)).toBe('ended')
  await expect(surface(page)).toHaveCount(0)
  await expect(page.locator('[data-officeview-presenter-view]')).toHaveCount(0)
  await expect.poll(async () => (await rec(page)).closes).toEqual([4005])
})
