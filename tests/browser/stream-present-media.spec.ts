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
type Stream = { present?: string; showDiagnostics: Array<{ code: string; message: string }>; mediaEnds: Array<{ slide: number; nodeId: number; t: number; wall: number }>; mediaTrace: boolean }
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

// ---------------------------------------------------------------------------
// Task 14 (plan D20, D22): the media smoke, on `?deck=narration|transition-sound|video` (server/test/app.ts).
// ---------------------------------------------------------------------------

const CONTROL = `http://127.0.0.1:${P + 1}`
/** Opens `deck`, waits for the painted unit, and records `canPlayType` for AAC (D20): `''` is the no-codec branch. */
async function openDeck(page: Page, deck: string, sent: string[], extra = ''): Promise<string> {
  page.on('websocket', (ws) => ws.on('framesent', (f) => { if (typeof f.payload === 'string') { const m = /"type":"([a-z-]+)"/.exec(f.payload); if (m) sent.push(m[1]!) } }))
  await page.goto(`/src/debug/stream-viewer.html?port=${P}&deck=${deck}&backend=canvas2d${extra ? `&${extra}` : ''}`)
  await expect(page.locator('[data-unit-wrap]')).toHaveCount(1)
  await expect(page.locator('canvas[data-painted="1"]').first()).toBeVisible()
  const aac = await page.evaluate(() => document.createElement('audio').canPlayType('audio/mp4; codecs="mp4a.40.2"'))
  test.info().annotations.push({ type: 'D20', description: aac === '' ? 'no AAC: the media-codec path' : `AAC: ${aac}` })
  return aac
}
async function startPresent(page: Page): Promise<void> {
  await page.locator('[data-debug-present]').click()
  await expect.poll(async () => (await stream(page)).present).toBe('started')
  await expect(surface(page)).toHaveAttribute('data-slide', '0')
}
/** The largest `currentTime` (s) of the elements `selector` matches in the document (0 when none). */
const playedSeconds = (page: Page, selector: string) =>
  page.evaluate((sel) => Math.max(0, ...[...document.querySelectorAll<HTMLMediaElement>(sel)].map((e) => e.currentTime)), selector)

/** `narrationDeck`: the narration is `a.m4a` (probed `durationMs` 2000, plan D17) on `cTn` 3, begun with the slide. */
const NARRATION = { nodeId: 3, durationMs: 2000 }

test('no-codec Chromium: poster + media-codec notice + clock end', async ({ page }) => {
  test.setTimeout(60_000)
  const sent: string[] = []
  // `?trace=media`: the debug page records each computed media end (node, show-clock time, wall time seen).
  const aac = await openDeck(page, 'narration', sent, 'trace=media')
  test.skip(aac !== '', `this Chromium decodes AAC (canPlayType ${JSON.stringify(aac)}): the playback cases run instead`)
  await expect.poll(async () => (await stream(page)).mediaTrace).toBe(true)
  await startPresent(page)
  const started = await page.evaluate(() => performance.now())
  // Refused before any byte is asked for; the typed notice; no element plays.
  await expect.poll(async () => (await stream(page)).showDiagnostics.some((d) => /media-codec/.test(`${d.code} ${d.message}`)), { timeout: 10_000 }).toBe(true)
  expect(sent.filter((t) => t === 'media-request')).toEqual([])
  expect(await playedSeconds(page, 'audio, video')).toBe(0)
  // D4: a client-side refusal keeps the clock end, begin + durationMs. Not yet at 1 s of the slide…
  await page.waitForTimeout(Math.max(0, 1000 - (await page.evaluate(() => performance.now())) + started))
  expect((await stream(page)).mediaEnds.filter((e) => e.nodeId === NARRATION.nodeId)).toEqual([])
  // …and recorded at exactly 2000 ms of show time, never seen before 2000 ms of wall time after the show started.
  await expect.poll(async () => (await stream(page)).mediaEnds.filter((e) => e.nodeId === NARRATION.nodeId).length, { timeout: 10_000 }).toBe(1)
  const end = (await stream(page)).mediaEnds.find((e) => e.nodeId === NARRATION.nodeId)!
  expect(end).toMatchObject({ slide: 0, t: NARRATION.durationMs })
  expect(end.wall - started).toBeGreaterThanOrEqual(NARRATION.durationMs - 250)
  expect(await playedSeconds(page, 'audio, video')).toBe(0)
  await surface(page).click({ position: { x: 20, y: 20 } })
  await expect(surface(page)).toHaveAttribute('data-slide', 'end')
})

test('transition sound and narration play (media element currentTime advances)', async ({ page }) => {
  test.setTimeout(60_000)
  const sent: string[] = []
  const aac = await openDeck(page, 'transition-sound', sent)
  test.skip(aac === '', 'D20: this Chromium has no AAC; the no-codec case covers it')
  await startPresent(page)
  // The start sound (`a.m4a`) is asked for once and plays in the audience surface.
  await expect.poll(() => sent.filter((t) => t === 'media-request').length, { timeout: 10_000 }).toBe(1)
  await expect.poll(() => playedSeconds(page, '[data-officeview-show-surface=""] audio[data-officeview-show-sound]'), { timeout: 10_000 }).toBeGreaterThan(0.2)

  const sent2: string[] = []
  await openDeck(page, 'narration', sent2)
  await startPresent(page)
  const narration = surface(page).locator('audio[data-officeview-show-media-spid="3"]')
  await expect.poll(() => sent2.filter((t) => t === 'media-request').length, { timeout: 10_000 }).toBe(1)
  await expect(narration).toHaveCount(1, { timeout: 10_000 })
  const a = await playedSeconds(page, 'audio[data-officeview-show-media-spid="3"]')
  await expect.poll(() => playedSeconds(page, 'audio[data-officeview-show-media-spid="3"]'), { timeout: 10_000 }).toBeGreaterThan(a + 0.2)
  expect(await page.evaluate(() => (window as unknown as { __ovCloses: number[] }).__ovCloses)).not.toContain(4008)
})

test('a video slide shows its piece-4 still, and no media-request is sent for it', async ({ page }) => {
  const sent: string[] = []
  await openDeck(page, 'video', sent)
  await startPresent(page)
  // D22: the still (drawn as the unit image, diagnosed `still`); no video element, no request, with or without AAC.
  await expect(surface(page)).toHaveAttribute('data-officeview-show-frame', 'image')
  await expect.poll(async () => (await stream(page)).showDiagnostics.some((d) => d.code === 'still')).toBe(true)
  await page.waitForTimeout(1500)
  expect(sent.filter((t) => t === 'media-request')).toEqual([])
  await expect(page.locator('video, audio')).toHaveCount(0)
  await surface(page).click({ position: { x: 20, y: 20 } })
  await expect(surface(page)).toHaveAttribute('data-slide', 'end')
})

test('revocation mid-narration exits; no element plays after ended', async ({ page }) => {
  test.setTimeout(60_000)
  const sent: string[] = []
  const aac = await openDeck(page, 'narration', sent)
  test.skip(aac === '', 'D20: this Chromium has no AAC; nothing plays to revoke mid-way')
  await startPresent(page)
  // Mid-narration: hold the media elements themselves (they may be detached later), and see them playing: each
  // narration element unpaused with its currentTime advancing between two samples.
  await expect.poll(() => playedSeconds(page, 'audio[data-officeview-show-media-spid="3"]'), { timeout: 10_000 }).toBeGreaterThan(0.1)
  await page.evaluate(() => { (window as unknown as { __ovHeld: HTMLMediaElement[] }).__ovHeld = [...document.querySelectorAll<HTMLMediaElement>('audio, video')] })
  const held = () => page.evaluate(() => (window as unknown as { __ovHeld: HTMLMediaElement[] }).__ovHeld.map((e) => ({ narration: e.getAttribute('data-officeview-show-media-spid') === '3', paused: e.paused, t: e.currentTime, connected: e.isConnected })))
  const s1 = await held()
  await page.waitForTimeout(250)
  const s2 = await held()
  expect(s1.filter((x) => x.narration)).toHaveLength(1)
  for (const [i, x] of s2.entries()) if (x.narration) { expect(x.paused).toBe(false); expect(x.t).toBeGreaterThan(s1[i]!.t) }
  const sids = (await (await page.request.get(`${CONTROL}/sessions`)).json() as Array<{ sid: string; status: string }>).filter((s) => s.status === 'stream')
  expect(sids.length).toBeGreaterThan(0)
  for (const { sid } of sids) await page.request.post(`${CONTROL}/revoke?sid=${sid}`)
  await expect.poll(async () => (await stream(page)).present).toBe('ended')
  await expect(surface(page)).toHaveCount(0)
  // After ended: the same elements, attached or not, are paused and do not advance.
  const a1 = await held()
  await page.waitForTimeout(500)
  const a2 = await held()
  expect(a1.length).toBe(s1.length)
  for (const [i, x] of a1.entries()) { expect(x.paused, `element ${i}`).toBe(true); expect(a2[i]!.t, `element ${i}`).toBe(x.t) }
  await expect.poll(() => page.evaluate(() => (window as unknown as { __ovCloses: number[] }).__ovCloses)).toContain(4005)
})
