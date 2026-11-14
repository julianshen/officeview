/**
 * Piece 4 Task 14 — the streamed-show browser smoke (spec §8): animated streamed decks in Chromium against
 * server/test/app.ts, under both compositor backends (`?backend=webgl2`, the browser's own; `?backend=canvas2d`,
 * where the harness page makes WebGL2 unavailable). Decks: `?deck=fly|morph|cube` (src/testdata/show-decks.ts).
 *
 * An init script records, per page:
 * - every WebSocket close code (Playwright's own websocket `close` event carries no code), so "no 4008" is real;
 * - every draw on a canvas that has been inside the audience show surface (Canvas 2D draws and WebGL2
 *   draws/clears/uploads), with its time, so "no canvas mutation after ended" is measured from the page's own
 *   `ended` time (`__ovStream.endedAt`), and is proven non-vacuous by the draws while the show animates.
 * `window.open` returns null and `requestFullscreen` rejects, as in stream-present.spec.ts.
 */
import { expect, test, type Page } from '@playwright/test'

const P = Number(process.env.OV_STREAM_PORT ?? 4181)
const CONTROL = `http://127.0.0.1:${P + 1}`
const BACKENDS = ['webgl2', 'canvas2d'] as const
const SLIDES = { fly: 1, morph: 2, cube: 2 } as const
type Deck = keyof typeof SLIDES

const open = (page: Page, deck: Deck, backend: string) => page.goto(`/src/debug/stream-viewer.html?port=${P}&deck=${deck}&backend=${backend}`)
const surface = (page: Page) => page.locator('[data-officeview-show-surface=""]')
const presentState = (page: Page) => page.evaluate(() => (window as unknown as { __ovStream: { present?: string } }).__ovStream.present)
const endedAt = (page: Page) => page.evaluate(() => (window as unknown as { __ovStream: { endedAt?: number } }).__ovStream.endedAt)
const streamError = (page: Page) => page.evaluate(() => (window as unknown as { __ovStream: { error?: string } }).__ovStream.error)
type Rec = { closes: number[]; draws: number[] }
const rec = (page: Page) => page.evaluate(() => (window as unknown as { __ovAnim: Rec }).__ovAnim)
const drawsSince = async (page: Page, t: number) => (await rec(page)).draws.filter((d) => d > t).length
const now = (page: Page) => page.evaluate(() => performance.now())

/** Client → server ops by type, with their send times; refusals received. */
function wire(page: Page) {
  const sent: Array<{ type: string; at: number }> = []
  const refused: string[] = []
  const received: string[] = []
  page.on('websocket', (ws) => {
    ws.on('framesent', (f) => {
      if (typeof f.payload !== 'string') return
      const m = /"type":"([a-z-]+)"/.exec(f.payload)
      if (m) sent.push({ type: m[1]!, at: Date.now() })
    })
    ws.on('framereceived', (f) => {
      if (typeof f.payload !== 'string') return
      const m = /"type":"([a-z-]+)"/.exec(f.payload)
      if (m) received.push(m[1]!)
      if (/"type":"(refused|error)"/.test(f.payload)) refused.push(f.payload)
    })
  })
  return {
    sent, refused, received,
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
    const r = { closes: [] as number[], draws: [] as number[] }
    ;(window as unknown as { __ovAnim: typeof r }).__ovAnim = r
    window.open = () => null
    Element.prototype.requestFullscreen = function () { return Promise.reject(new Error('full screen is stubbed')) }
    const WS = window.WebSocket
    window.WebSocket = class extends WS {
      constructor(url: string | URL, protocols?: string | string[]) {
        super(url, protocols)
        this.addEventListener('close', (e) => { r.closes.push(e.code) })
      }
    }
    // A canvas is a show canvas once it has been inside the audience surface; it stays one after the surface goes.
    const shown = new WeakSet<object>()
    const isShow = (c: unknown) => {
      if (!c || shown.has(c as object)) return !!c && shown.has(c as object)
      const el = c as HTMLCanvasElement
      if (typeof el.closest === 'function' && el.closest('[data-officeview-show-surface=""]')) { shown.add(el); return true }
      return false
    }
    const wrap = (proto: object | undefined, names: string[]) => {
      if (!proto) return
      for (const n of names) {
        const f = (proto as Record<string, unknown>)[n]
        if (typeof f !== 'function') continue
        ;(proto as Record<string, unknown>)[n] = function (this: { canvas: unknown }, ...a: unknown[]) {
          if (isShow(this.canvas)) r.draws.push(performance.now())
          return (f as (...x: unknown[]) => unknown).apply(this, a)
        }
      }
    }
    wrap(window.CanvasRenderingContext2D?.prototype, ['drawImage', 'putImageData', 'fillRect', 'clearRect', 'fill', 'stroke', 'fillText'])
    wrap((window as unknown as { WebGL2RenderingContext?: { prototype: object } }).WebGL2RenderingContext?.prototype, ['drawArrays', 'drawElements', 'clear', 'texImage2D', 'texSubImage2D'])
  })
})

/** The stream viewer is up: every unit laid out, the first painted. */
async function ready(page: Page, deck: Deck) {
  await expect(page.locator('[data-unit-wrap]')).toHaveCount(SLIDES[deck])
  await expect(page.locator('canvas[data-painted="1"]').first()).toBeVisible()
}

async function present(page: Page) {
  await page.locator('[data-debug-present]').click()
  await expect.poll(() => presentState(page)).toBe('started')
  await expect(surface(page)).toHaveAttribute('data-slide', '0')
}

/** The show composites on the backend the page was opened with (never silently the other). */
async function expectBackend(page: Page, backend: string) {
  const gl = surface(page).locator('canvas[data-slideshow-backend="webgl2"]')
  const c2d = surface(page).locator('canvas[data-slideshow-backend="canvas2d"]')
  if (backend === 'webgl2') { await expect(gl).toHaveCount(1); await expect(c2d).toHaveCount(0) } else { await expect(c2d).toHaveCount(1); await expect(gl).toHaveCount(0) }
}

/** A hash of the audience output canvas (the surface's own 2-D canvas), read in the page. */
const hashOf = (page: Page, handle: Awaited<ReturnType<Page['evaluateHandle']>>) => page.evaluate((c) => {
  const cv = c as HTMLCanvasElement
  if (!(cv.width > 0 && cv.height > 0)) return `empty:${cv.width}x${cv.height}`
  const ctx = cv.getContext('2d')
  if (!ctx) return 'no-2d'
  const d = ctx.getImageData(0, 0, cv.width, cv.height).data
  let h = 0x811c9dc5
  for (let i = 0; i < d.length; i += 4) { h ^= d[i]! ^ (d[i + 1]! << 8) ^ (d[i + 2]! << 16); h = Math.imul(h, 0x01000193) >>> 0 }
  return `${cv.width}x${cv.height}:${h.toString(16)}`
}, handle)

for (const backend of BACKENDS) {
  test.describe(`backend=${backend}`, () => {
    test('fly-in, Morph and 3-D decks play end to end', async ({ page }) => {
      test.setTimeout(90_000)
      for (const deck of ['fly', 'morph', 'cube'] as const) {
        const w = wire(page)
        await open(page, deck, backend)
        await ready(page, deck)
        await present(page)
        await expectBackend(page, backend)
        // The first slide's bundle has arrived (its plan reply) before the first →: what is measured is the show,
        // not a press during the load (see the Task 14 report).
        if (deck === 'fly') await expect.poll(() => w.received.includes('show-plan-start')).toBe(true)
        await page.waitForTimeout(1000)
        // → through every build and slide to end-of-show (two presses on each deck), with frames drawn while it animates.
        const t0 = await now(page)
        await page.keyboard.press('ArrowRight')
        // fly: its one build (500 ms); morph / cube: slide 2 by Morph (1.2 s) / Cube (1 s).
        await expect(surface(page)).toHaveAttribute('data-slide', deck === 'fly' ? '0' : '1')
        await page.waitForTimeout(1800)
        await expect(page.locator('[data-officeview-show-end]')).toBeHidden()
        await page.keyboard.press('ArrowRight')
        await expect(page.locator('[data-officeview-show-end]')).toBeVisible()
        expect(await drawsSince(page, t0), `${deck}: frames drawn while animating`).toBeGreaterThan(20)
        expect(w.count('show-plan'), deck).toBeGreaterThan(0)
        if (deck === 'morph') expect(w.count('show-morph')).toBeGreaterThan(0)
        expect(w.refused, deck).toEqual([])
        await page.keyboard.press('Escape')
        await expect.poll(() => presentState(page)).toBe('ended')
        expect(await streamError(page), deck).toBeUndefined()
        expect((await rec(page)).closes, deck).toEqual([])
      }
    })

    test('holding → for 10 s: socket stays open (no 4008)', async ({ page }) => {
      test.setTimeout(60_000)
      const w = wire(page)
      // The Morph deck: every revisit re-charges its bundle and its Morph pair (Task 11 ruling).
      await open(page, 'morph', backend)
      await ready(page, 'morph')
      await present(page)
      await expectBackend(page, backend)
      const t0 = Date.now()
      let at = 0, presses = 0
      while (Date.now() - t0 < 10_000) {
        const key = at === 0 ? 'ArrowRight' : 'ArrowLeft'
        await page.keyboard.down(key)
        await page.keyboard.up(key)
        at = 1 - at
        presses++
        await page.waitForTimeout(33)
      }
      await page.waitForTimeout(500)
      expect(presses).toBeGreaterThan(150)
      expect(await presentState(page)).toBe('started')
      expect(await streamError(page)).toBeUndefined()
      expect((await rec(page)).closes).toEqual([])
      expect(w.refused.filter((r) => r.includes('rate-limited'))).toEqual([])
      const peaks = Object.fromEntries(['unit', 'show-plan', 'show-track', 'show-morph'].map((t) => [t, w.peak(t, 60_000)]))
      // The server's charged window is 60/min per session; the show budgets 48/min of it (Task 12).
      expect(peaks.unit! + peaks['show-plan']! + peaks['show-track']! + peaks['show-morph']!).toBeLessThanOrEqual(60)
      console.log(`hold [${backend}]: ${presses} presses in 10 s; peaks per 60 s ${JSON.stringify(peaks)}; totals ${JSON.stringify(Object.fromEntries(['unit', 'show-plan', 'show-track', 'show-morph'].map((t) => [t, w.count(t)])))}; close codes none`)
      await page.keyboard.press('Escape')
      await expect.poll(() => presentState(page)).toBe('ended')
    })

    test('revocation mid-animation exits; no canvas mutation after ended', async ({ page }) => {
      test.setTimeout(60_000)
      await open(page, 'fly', backend)
      await ready(page, 'fly')
      await present(page)
      await expectBackend(page, backend)
      const output = await page.evaluateHandle(() => document.querySelector('[data-officeview-show-surface=""] > canvas'))
      // Control: the fly-in animates — the surface's canvases are drawn on while it plays.
      const before = await now(page)
      await page.keyboard.press('ArrowRight')
      await page.waitForTimeout(120)
      expect(await drawsSince(page, before), 'draws while the fly-in plays').toBeGreaterThan(3)
      // Revoke mid-flight (the fly-in runs 500 ms).
      const sids = (await (await page.request.get(`${CONTROL}/sessions`)).json() as Array<{ sid: string; status: string }>).filter((s) => s.status === 'stream')
      expect(sids.length).toBeGreaterThan(0)
      for (const { sid } of sids) await page.request.post(`${CONTROL}/revoke?sid=${sid}`)
      await expect.poll(() => presentState(page)).toBe('ended')
      const end = (await endedAt(page))!
      expect(end).toBeGreaterThan(before)
      const h0 = await hashOf(page, output)
      await page.waitForTimeout(250)
      const h1 = await hashOf(page, output)
      await page.waitForTimeout(250)
      const h2 = await hashOf(page, output)
      expect([h1, h2]).toEqual([h0, h0])
      expect(await drawsSince(page, end), 'draws on a show canvas after ended').toBe(0)
      await expect(surface(page)).toHaveCount(0)
      await expect.poll(async () => (await rec(page)).closes).toEqual([4005])
    })
  })
}
