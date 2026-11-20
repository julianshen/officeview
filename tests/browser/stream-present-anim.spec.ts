/**
 * Piece 4 Task 14 — the streamed-show browser smoke (spec §8): animated streamed decks in Chromium against
 * server/test/app.ts, under both compositor backends (`?backend=webgl2`, the browser's own; `?backend=canvas2d`,
 * where the harness page makes WebGL2 unavailable). Decks: `?deck=fly|morph|cube|nav` (src/testdata/show-decks.ts).
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
const SLIDES = { fly: 1, morph: 2, cube: 2, nav: 3 } as const
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

    test('holding → (300 key events, ~10 s): socket stays open (no 4008)', async ({ page }) => {
      // A fixed volume of input, not a wall-clock window (PR #78 round 4: under CPU load a 10 s window counted
      // fewer presses, a false failure); a loaded run takes longer instead.
      test.setTimeout(120_000)
      const w = wire(page)
      // The Morph deck: every revisit re-charges its bundle and its Morph pair (Task 11 ruling).
      await open(page, 'morph', backend)
      await ready(page, 'morph')
      await present(page)
      await expectBackend(page, backend)
      // Every slide change the show makes while held (the surface's `data-slide`, observed in the page).
      await page.evaluate(() => {
        const el = document.querySelector('[data-officeview-show-surface=""]')!
        const w = window as unknown as { __ovSlideChanges: string[] }
        w.__ovSlideChanges = []
        new MutationObserver(() => w.__ovSlideChanges.push(el.getAttribute('data-slide') ?? '')).observe(el, { attributes: true, attributeFilter: ['data-slide'] })
      })
      const HOLD_EVENTS = 300
      const t0 = Date.now()
      let at = 0, presses = 0
      while (presses < HOLD_EVENTS) {
        const key = at === 0 ? 'ArrowRight' : 'ArrowLeft'
        await page.keyboard.down(key)
        await page.keyboard.up(key)
        at = 1 - at
        presses++
        await page.waitForTimeout(33)
      }
      await page.waitForTimeout(500)
      // The input was not ignored: the presses navigated back and forth between slides 1 and 2 (a press that lands
      // during the Morph seeks it instead, so not every press is a change).
      const changes = await page.evaluate(() => (window as unknown as { __ovSlideChanges: string[] }).__ovSlideChanges)
      expect(changes.length).toBeGreaterThanOrEqual(HOLD_EVENTS / 2)
      expect(new Set(changes)).toEqual(new Set(['0', '1']))
      // Task 13 (8435707c): once settled, ← / → asks for no Morph pair again. The deck has one Morph pair (slide 1 →
      // slide 2), so show-morph is bounded by the pairs, not the presses: one fetch per pair, plus one re-ask in case
      // the first press aborts it in flight. Before 8435707c every revisit re-asked it (40 in 300 presses).
      const MORPH_PAIRS = 1
      expect(w.count('show-morph')).toBeLessThanOrEqual(2 * MORPH_PAIRS)
      expect(w.count('show-morph')).toBeGreaterThanOrEqual(MORPH_PAIRS)
      expect(await surface(page).getAttribute('data-slide')).toMatch(/^[01]$/)
      expect(await presentState(page)).toBe('started')
      expect(await streamError(page)).toBeUndefined()
      expect((await rec(page)).closes).toEqual([])
      expect(w.refused.filter((r) => r.includes('rate-limited'))).toEqual([])
      const peaks = Object.fromEntries(['unit', 'show-plan', 'show-track', 'show-morph'].map((t) => [t, w.peak(t, 60_000)]))
      // The server's charged window is 60/min per session; the show budgets 48/min of it (Task 12).
      expect(peaks.unit! + peaks['show-plan']! + peaks['show-track']! + peaks['show-morph']!).toBeLessThanOrEqual(60)
      console.log(`hold [${backend}]: ${presses} presses, ${changes.length} slide changes, in ${((Date.now() - t0) / 1000).toFixed(1)} s; peaks per 60 s ${JSON.stringify(peaks)}; totals ${JSON.stringify(Object.fromEntries(['unit', 'show-plan', 'show-track', 'show-morph'].map((t) => [t, w.count(t)])))}; close codes none`)
      await page.keyboard.press('Escape')
      await expect.poll(() => presentState(page)).toBe('ended')
    })

    test('slide changes (cut and fly-in, bundles prefetched) never show a mostly black frame (PR #78 round 2)', async ({ page }) => {
      test.setTimeout(60_000)
      // A WebGL2 canvas read outside the task that drew it is its cleared buffer (transparent, so the black 2-D
      // output under it shows): with the buffer preserved, a read at any time is the last presented frame, which
      // is what the screen holds. Only the probe's read changes; the show draws as it does.
      await page.addInitScript(() => {
        const get = HTMLCanvasElement.prototype.getContext as (this: HTMLCanvasElement, id: string, o?: object) => RenderingContext | null
        ;(HTMLCanvasElement.prototype as unknown as { getContext: typeof get }).getContext = function (id, o) {
          return get.call(this, id, id === 'webgl2' ? { ...o, preserveDrawingBuffer: true } : o)
        }
      })
      const w = wire(page)
      await open(page, 'nav', backend)
      await ready(page, 'nav')
      await present(page)
      await expectBackend(page, backend)
      // Both animated slides' bundles have arrived before the first press (slide 1's, then slide 3's as the +2
      // prefetch): what is measured is the change itself. A slide whose bundle is still out shows the piece-3
      // loading frame (black, with its text), which is not the flash this pins. The server can be slow here
      // after the hold test's renders (observed: slide 3's reply ~6 s late), hence the long poll.
      await expect.poll(() => w.received.filter((t) => t === 'show-plan-start').length, { timeout: 30_000 }).toBeGreaterThanOrEqual(2)
      await page.waitForTimeout(1000)
      type Sample = { slide: string | null; kind: string | null; black: number; control?: 'opaque' | 'clear' }
      // Each animation frame: the share of what the audience sees that is black. The probe starts from the show
      // host's own opaque backdrop (its computed background, black here: the letterbox), then draws each visible
      // canvas of the surface in DOM order where it is displayed (its box, object-fit honoured), so a transparent
      // or cleared canvas shows the backdrop through it, as on screen.
      const geometry = await page.evaluate(() => {
        const layers = () => [...document.querySelectorAll<HTMLCanvasElement>('[data-officeview-show-surface=""] canvas')]
          .filter((c) => c.width > 0 && c.height > 0 && getComputedStyle(c).display !== 'none' && getComputedStyle(c).visibility !== 'hidden' && !c.closest('[hidden]'))
        const surfaceEl = document.querySelector<HTMLElement>('[data-officeview-show-surface=""]')!
        /** The nearest element at or above the surface with an opaque background: what shows through transparency. */
        let back: HTMLElement | null = surfaceEl
        while (back && !/^rgb\(|rgba\(.*,\s*1\)$/.test(getComputedStyle(back).backgroundColor)) back = back.parentElement
        if (!back) throw new Error('no opaque backdrop above the show surface')
        const backdrop = getComputedStyle(back).backgroundColor
        const host = back.getBoundingClientRect()
        const PW = 64, PH = Math.max(1, Math.round(64 * host.height / host.width))
        const probe = document.createElement('canvas')
        probe.width = PW; probe.height = PH
        const ctx = probe.getContext('2d', { willReadFrequently: true })!
        /** Where `c`'s pixels are displayed, in probe px (object-fit contain / cover / fill). */
        const placed = (c: HTMLCanvasElement) => {
          const r = c.getBoundingClientRect()
          const fit = getComputedStyle(c).objectFit
          let w = r.width, h = r.height
          if (fit === 'contain' || fit === 'cover') {
            const k = (fit === 'contain' ? Math.min : Math.max)(r.width / c.width, r.height / c.height)
            w = c.width * k; h = c.height * k
          }
          const x = r.left + (r.width - w) / 2 - host.left, y = r.top + (r.height - h) / 2 - host.top
          const sx = PW / host.width, sy = PH / host.height
          return [x * sx, y * sy, w * sx, h * sy] as const
        }
        /** The black share of the frame: as shown, or a control (the top canvas opaque black; no canvas drawn at all). */
        const measure = (mode?: 'opaque' | 'clear') => {
          ctx.globalCompositeOperation = 'copy'
          ctx.fillStyle = backdrop
          ctx.fillRect(0, 0, PW, PH)
          ctx.globalCompositeOperation = 'source-over'
          const ls = mode === 'clear' ? [] : layers()
          for (const c of ls) ctx.drawImage(c, ...placed(c))
          const top = ls.at(-1)
          if (mode === 'opaque' && top) { ctx.fillStyle = '#000'; ctx.fillRect(...placed(top)) }
          const d = ctx.getImageData(0, 0, PW, PH).data
          let black = 0
          for (let i = 0; i < d.length; i += 4) if (Math.max(d[i]!, d[i + 1]!, d[i + 2]!) < 32) black++
          return black / (PW * PH)
        }
        const st = { samples: [] as Sample[], on: true, control: false }
        ;(window as unknown as { __ovBlack: typeof st }).__ovBlack = st
        const tick = () => {
          if (!st.on) return
          const root = document.querySelector('[data-officeview-show-surface=""]')
          const at = { slide: root?.getAttribute('data-slide') ?? null, kind: root?.getAttribute('data-officeview-show-frame') ?? null }
          st.samples.push({ ...at, black: measure() })
          if (st.control) {
            // The controls, on the same frame: what the oracle scores if the stage drew opaque black, or drew nothing.
            st.control = false
            st.samples.push({ ...at, black: measure('opaque'), control: 'opaque' }, { ...at, black: measure('clear'), control: 'clear' })
          }
          requestAnimationFrame(tick)
        }
        requestAnimationFrame(tick)
        return { backdrop, w: host.width, h: host.height }
      })
      // → through the fly-in, the cut to slide 2 and the fly-in on slide 3, then ← back to the first slide.
      for (const key of ['ArrowRight', 'ArrowRight', 'ArrowRight', 'ArrowRight', 'ArrowLeft', 'ArrowLeft', 'ArrowLeft', 'ArrowLeft']) {
        await page.keyboard.press(key)
        await page.waitForTimeout(900)
      }
      await page.evaluate(() => { (window as unknown as { __ovBlack: { control: boolean } }).__ovBlack.control = true })
      await page.waitForTimeout(300)
      const all = await page.evaluate(() => {
        const st = (window as unknown as { __ovBlack: { on: boolean; samples: Sample[] } }).__ovBlack
        st.on = false
        return st.samples
      })
      const samples = all.filter((x) => !x.control)
      // The deck's slides are 4:3 (buildPptxDeck's 9144000 × 6858000 EMU), letterboxed in the host: the backdrop's bars
      // are this share of the frame. A frame is "mostly black" when, besides the bars, over half the slide is black.
      const bars = 1 - Math.min(geometry.w / geometry.h, 4 / 3) / Math.max(geometry.w / geometry.h, 4 / 3)
      const mostly = bars + (1 - bars) / 2
      // Non-vacuous: every frame for ~7 s, across all three slides; the backdrop is the host's (opaque black);
      // both controls score as mostly black.
      expect(geometry.backdrop).toBe('rgb(0, 0, 0)')
      expect(samples.length).toBeGreaterThan(200)
      expect(new Set(samples.map((x) => x.slide))).toEqual(new Set(['0', '1', '2']))
      const controls = Object.fromEntries(all.filter((x) => x.control).map((x) => [x.control!, x.black]))
      expect(Object.keys(controls).sort()).toEqual(['clear', 'opaque'])
      expect(controls.opaque).toBeGreaterThan(mostly)
      expect(controls.clear).toBeGreaterThan(mostly)
      // Prefetched: every frame is a slide's picture, never the loading or error frame.
      expect([...new Set(samples.map((x) => x.kind))], 'frame kinds').toEqual(['image'])
      const max = Math.max(...samples.map((x) => x.black))
      console.log(`no-black [${backend}]: ${samples.length} frames; host ${geometry.w}×${geometry.h}, bars ${bars.toFixed(3)}; max black share ${max.toFixed(3)}; controls opaque ${controls.opaque!.toFixed(3)} clear ${controls.clear!.toFixed(3)}`)
      // The bars are seen (the backdrop is in the measure), and no frame is mostly black.
      expect(max).toBeGreaterThanOrEqual(bars - 0.03)
      const dark = samples.map((x, i) => ({ i, ...x })).filter((x) => x.black > mostly)
      expect(dark, `mostly black frames [${backend}]`).toEqual([])
      expect(w.refused).toEqual([])
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
