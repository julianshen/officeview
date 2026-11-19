/**
 * Task 11b (owner decision 2026-11-19; plan D20): a still slide that sounds, in Chromium against server/test/app.ts
 * (`?deck=stillsound`: `stillSoundDeck('audio-end')`, a narration on shape 3 begun 250 ms after entry, a box revealed
 * on its end, a click-started effect sound). The slide plays its piece-3 still; the narration is asked for and plays
 * on the show clock; the click sound never is.
 *
 * D20: open-source Chromium may lack AAC (`canPlayType('audio/mp4; codecs="mp4a.40.2"') === ''`). Then the media-codec
 * path is asserted instead: the lease refuses `media-codec` before any byte is requested (no `media-request` on the
 * wire), the player raises its typed notice, and the still shows as without media.
 */
import { expect, test, type Page } from '@playwright/test'

const P = Number(process.env.OV_STREAM_PORT ?? 4181)
const surface = (page: Page) => page.locator('[data-officeview-show-surface=""]')
type Stream = { present?: string; showDiagnostics: Array<{ code: string; message: string }> }
const stream = (page: Page) => page.evaluate(() => (window as unknown as { __ovStream: Stream }).__ovStream)

test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    window.open = () => null
    Element.prototype.requestFullscreen = function () { return Promise.reject(new Error('full screen is stubbed')) }
    const closes: number[] = []
    ;(window as unknown as { __ovCloses: number[] }).__ovCloses = closes
    const WS = window.WebSocket
    window.WebSocket = class extends WS {
      constructor(url: string | URL, protocols?: string | string[]) { super(url, protocols); this.addEventListener('close', (e) => { closes.push(e.code) }) }
    }
  })
})

test('a still with audio plays its narration (or takes the media-codec path without AAC); a click advances', async ({ page }) => {
  const sent: string[] = []
  page.on('websocket', (ws) => ws.on('framesent', (f) => { if (typeof f.payload === 'string') { const m = /"type":"([a-z-]+)"/.exec(f.payload); if (m) sent.push(m[1]!) } }))
  await page.goto(`/src/debug/stream-viewer.html?port=${P}&deck=stillsound&backend=canvas2d`)
  await expect(page.locator('[data-unit-wrap]')).toHaveCount(1)
  await expect(page.locator('canvas[data-painted="1"]').first()).toBeVisible()
  const aac = await page.evaluate(() => document.createElement('audio').canPlayType('audio/mp4; codecs="mp4a.40.2"'))
  test.info().annotations.push({ type: 'D20', description: aac === '' ? 'no AAC: the media-codec path' : `AAC: ${aac}` })

  await page.locator('[data-debug-present]').click()
  await expect.poll(async () => (await stream(page)).present).toBe('started')
  await expect(surface(page)).toHaveAttribute('data-slide', '0')
  // The piece-3 still: drawn as an image, and diagnosed as a still (the server's sound-only plan names it).
  await expect(surface(page)).toHaveAttribute('data-officeview-show-frame', 'image')
  await expect.poll(async () => (await stream(page)).showDiagnostics.some((d) => d.code === 'still')).toBe(true)
  expect(sent).toContain('show-plan')
  expect(sent).toContain('media-show')

  const narration = surface(page).locator('audio[data-officeview-show-media-spid="3"]')
  if (aac !== '') {
    // The narration is asked for once and its element is made in the audience surface at its begin.
    await expect.poll(() => sent.filter((t) => t === 'media-request').length, { timeout: 10_000 }).toBe(1)
    await expect(narration).toHaveCount(1, { timeout: 10_000 })
  } else {
    // D20: refused `media-codec` before any request; the typed notice; no element plays.
    await expect.poll(async () => (await stream(page)).showDiagnostics.some((d) => /media-codec/.test(`${d.code} ${d.message}`)), { timeout: 10_000 }).toBe(true)
    expect(sent.filter((t) => t === 'media-request')).toEqual([])
  }
  await expect(surface(page)).toHaveAttribute('data-officeview-show-frame', 'image')

  // A click on the still advances (no step, no trigger): the click-started effect sound is never asked for.
  const before = sent.filter((t) => t === 'media-request').length
  // Fix round 1 (minor 6): a pointer click on the surface (the driver's hit test and advClick path), not a key.
  await surface(page).click({ position: { x: 20, y: 20 } })
  await expect(surface(page)).toHaveAttribute('data-slide', 'end')
  expect(sent.filter((t) => t === 'media-request').length).toBe(before)
  await expect(surface(page).locator('audio[data-officeview-show-effect-sound]')).toHaveCount(0)
  expect(await page.evaluate(() => (window as unknown as { __ovCloses: number[] }).__ovCloses)).not.toContain(4008)
})
