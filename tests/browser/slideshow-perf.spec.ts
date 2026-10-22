/**
 * Slideshow plan S1-22: the spec 7 frame benchmark. Runs only with
 * OV_SLIDESHOW_BENCH=1 (it is timing, so it never runs in CI; the CI-shaped
 * check that the recorder works on WebGL2 is in slideshow-corpus.spec.ts).
 *
 *   OV_SLIDESHOW_BENCH=1 IV_HARNESS_PORT=4391 bunx playwright test \
 *     --config playwright.viewer.config.ts --project slideshow slideshow-perf
 *
 * Benchmark deck tdf104788 at 1920 × 1080 CSS px, DPR 2 (a 3840 × 2160
 * backing store), Chromium. Each of slide 6's six click steps is measured
 * from its click until it settles: frame work is `performance.now()` around
 * the session's tick (JavaScript plus WebGL2 command submission), and with
 * `sync` a 1-px readback after it, so the GPU's work is in the number too.
 * A dropped frame is a rAF gap above 1.5 × the run's median gap. Target:
 * p95 ≤ 12 ms (spec 7). A miss is recorded (expect.soft) and the numbers are
 * still written, never hidden.
 *
 * The same run profiles the S1-14 concern: an animating transition mask
 * re-traces and re-uploads a device-sized texture per frame. tdf44223's
 * slide 2 `split` (mask, 1500 ms) is measured against tdf115394's slide 5
 * `fade` (transform, 4250 ms) at the same size.
 *
 * The project's launch options force SwiftShader (CPU rasterisation). With
 * OV_SLIDESHOW_GPU=1 the browser is launched on the machine's GPU instead
 * (ANGLE Metal on macOS); the renderer string is recorded with every result.
 * OV_SLIDESHOW_BENCH_OUT names a JSON-lines file the results are appended
 * to, one per measurement (a failed expect.soft restarts the worker, so the
 * results are written as they are taken).
 */
import { expect, test, type Page } from '@playwright/test'
import { appendFileSync, existsSync, mkdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'
import type { SessionController } from '../../src/slideshow/player/session'

export type { SlideshowStart } from '../../src/debug/slideshow-start'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-start.html'
const CORPUS = join(ROOT, 'corpus')
const GPU = process.env.OV_SLIDESHOW_GPU === '1'
export const FRAME_BUDGET_MS = 12

test.skip(process.env.OV_SLIDESHOW_BENCH !== '1', 'timing benchmark: set OV_SLIDESHOW_BENCH=1 (never in CI)')
test.skip(!existsSync(join(CORPUS, 'lo-tdf104788.pptx')), 'corpus/ not fetched (bun scripts/fetch-corpus.ts)')
test.use({
  viewport: { width: 1920, height: 1080 },
  deviceScaleFactor: 2,
  ...(GPU ? { launchOptions: { args: ['--use-angle=metal', '--enable-gpu', '--ignore-gpu-blocklist'] } } : {}),
})

interface Stats { frames: number; p50: number; p95: number; max: number; dropped: number; medianGapMs: number }
interface Result extends Stats { name: string; sync: boolean; renderer: string; backing: { w: number; h: number; dpr: number }; backend: string }
const results: Result[] = []
const OUT = process.env.OV_SLIDESHOW_BENCH_OUT

test.beforeAll(async () => {
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity').toBeNull()
})
test.beforeEach(async ({ page }) => {
  await page.route('**/__corpus/*.pptx', (route) => {
    const name = new URL(route.request().url()).pathname.split('/').pop()!
    return route.fulfill({ body: readFileSync(join(CORPUS, name)) })
  })
})

const quantile = (sorted: number[], q: number) => sorted.length ? sorted[Math.min(sorted.length - 1, Math.ceil(q * sorted.length) - 1)] : NaN
function stats(work: number[], stamps: number[]): Stats {
  const sorted = [...work].sort((a, b) => a - b)
  const gaps = stamps.slice(1).map((t, i) => t - stamps[i]).filter((g) => g > 0)
  const medianGap = quantile([...gaps].sort((a, b) => a - b), 0.5)
  return {
    frames: work.length, p50: quantile(sorted, 0.5), p95: quantile(sorted, 0.95), max: sorted[sorted.length - 1] ?? NaN,
    dropped: gaps.filter((g) => g > 1.5 * medianGap).length, medianGapMs: medianGap,
  }
}

async function present(page: Page, deck: string, sync: boolean, query = ''): Promise<void> {
  await page.goto(`${PAGE}?deck=corpus&name=${deck}&bench=1${sync ? '&sync=1' : ''}${query}`)
  await page.waitForFunction(() => window.__slideshowStart?.ready === true)
  await page.click('#present-a')
  await expect.poll(() => page.evaluate(() => [...window.__slideshowStart!.log])).toContain('a:started')
}
const info = (page: Page) => page.evaluate(() => {
  const s = window.__slideshowStart!.sessions.a as unknown as SessionController
  return { slide: s.state.slide, build: s.state.build, backend: s.stage?.backend ?? '', level: s.stage?.level, backing: s.stage!.backingStore, transition: s.stage?.transition }
})
const samples = (page: Page) => page.evaluate(() => ({ work: [...window.__slideshowStart!.bench!.work], stamps: [...window.__slideshowStart!.bench!.stamps] }))

async function record(page: Page, name: string, sync: boolean, work: number[], stamps: number[]): Promise<Result> {
  const i = await info(page)
  const renderer = await page.evaluate(() => window.__slideshowStart!.bench!.renderer)
  const r: Result = { name, sync, renderer, backing: i.backing, backend: i.backend, ...stats(work, stamps) }
  results.push(r)
  return r
}
function write(r: Result): void {
  const line = JSON.stringify({ gpu: GPU, budgetMs: FRAME_BUDGET_MS, ...r })
  console.log(line)
  if (OUT) { mkdirSync(dirname(OUT), { recursive: true }); appendFileSync(OUT, line + '\n') }
}


/** tdf44223 slide 2 splits in (a mask, re-traced and re-uploaded every frame); tdf115394 slide 5 fades in (a transform). */
const TRANSITIONS = [
  { deck: 'tdf44223', slide: 1, ms: 1500, effect: 'split' },
  { deck: 'tdf115394', slide: 4, ms: 4250, effect: 'fade' },
] as const

/** Advances into the slide's transition and measures it; `cut` holds the diagnostic when the transition was refused. */
async function transitionRun(page: Page, t: typeof TRANSITIONS[number], sync: boolean, query = ''): Promise<Result & { cut?: string }> {
  await present(page, t.deck, sync, query)
  await expect.poll(async () => (await info(page)).level).toBe('full')
  expect((await info(page)).backend).toBe('webgl2')
  if (t.slide > 1) await page.evaluate((n) => (window.__slideshowStart!.sessions.a as unknown as SessionController).goTo(n), t.slide - 1)
  await page.waitForTimeout(500)
  // Next until the slide changes (tdf44223 slide 1 has a click step first); the samples are of the last Next only.
  for (let n = 0; n < 4 && (await info(page)).slide !== t.slide; n++) {
    await page.evaluate(() => window.__slideshowStart!.bench!.reset())
    await page.keyboard.press('ArrowRight')
    await page.waitForTimeout(30)
  }
  await expect.poll(async () => (await info(page)).slide).toBe(t.slide)
  await page.waitForTimeout(t.ms + 100)
  const s = await samples(page)
  const refused = await page.evaluate(() => window.__slideshowStart!.diagnostics.filter((d) => d.code === 'transition-workspace').map((d) => d.message))
  const ws = await page.evaluate(() => window.__slideshowStart!.bench!.workspace())
  const r: Result & { cut?: string; workspace?: unknown } = await record(page, `${t.deck} slide ${t.slide + 1} ${t.effect} transition${query ? ` (${query.slice(1)})` : ''}`, sync, s.work, s.stamps)
  if (refused.length) r.cut = `${refused.join('; ')} (workspace ${ws ? `${ws.used} of ${ws.limit} B` : 'unknown'})`
  r.workspace = ws
  write(r)
  await page.keyboard.press('Escape')
  return r
}

for (const sync of [false, true]) {
  test(`tdf104788 frame work at 1920 × 1080, DPR 2 (${sync ? 'through a GPU readback' : 'submit only'})`, async ({ page }) => {
    test.setTimeout(180_000)
    await present(page, 'tdf104788', sync)
    await page.evaluate(() => (window.__slideshowStart!.sessions.a as unknown as SessionController).goTo(5))
    await expect.poll(async () => (await info(page)).slide).toBe(5)
    await expect.poll(async () => (await info(page)).level).toBe('full')
    const i = await info(page)
    // WebGL2 is required, and the backing store must be the benchmark's: 1920 × 1080 CSS px at DPR 2.
    expect(i.backend).toBe('webgl2')
    expect(i.backing).toEqual({ w: 3840, h: 2160, dpr: 2 })
    const spans = [1000, 1000, 500, 1000, 1000, 2000]
    const all = { work: [] as number[], stamps: [] as number[] }
    for (const [k, span] of spans.entries()) {
      await page.evaluate(() => window.__slideshowStart!.bench!.reset())
      await page.keyboard.press('ArrowRight')
      await expect.poll(async () => (await info(page)).build).toBe(k + 1)
      await page.waitForTimeout(span + 150)
      const s = await samples(page)
      write(await record(page, `tdf104788 slide 6 step ${k + 1}`, sync, s.work, s.stamps))
      all.work.push(...s.work)
      // Gaps between steps are not frames of steady animation.
      all.stamps.push(...s.stamps)
    }
    const r = await record(page, 'tdf104788 slide 6 all steps', sync, all.work, [])
    const steps = results.filter((x) => x.sync === sync && /step \d/.test(x.name))
    r.dropped = steps.reduce((n, x) => n + x.dropped, 0)
    write(r)
    expect.soft(r.p95, `p95 frame work (${r.renderer})`).toBeLessThanOrEqual(FRAME_BUDGET_MS)
    expect.soft(r.dropped, `dropped frames (${r.renderer})`).toBe(0)
  })

  test(`slide transitions at 1920 × 1080, DPR 2: split (mask) and fade (transform) (${sync ? 'through a GPU readback' : 'submit only'})`, async ({ page }) => {
    test.setTimeout(180_000)
    for (const t of TRANSITIONS) {
      const r = await transitionRun(page, t, sync)
      // At the benchmark size a refused transition frame is a diagnosed cut: recorded and escalated, not hidden.
      expect.soft(r.cut, `${t.effect} at 3840 × 2160: ${r.cut ?? ''}`).toBeUndefined()
      if (!r.cut) expect.soft(r.p95, `${t.effect} p95 (${r.renderer})`).toBeLessThanOrEqual(FRAME_BUDGET_MS)
    }
  })
}

// The S1-14 mask question needs a transition that runs: at DPR 1 (a 1920 × 1080 backing store) the frames fit.
test.describe('transition mask re-upload at 1920 × 1080, DPR 1', () => {
  test.use({ deviceScaleFactor: 1 })
  for (const sync of [false, true]) {
    test(`split (mask, re-uploaded per frame) vs fade (transform) (${sync ? 'through a GPU readback' : 'submit only'})`, async ({ page }) => {
      test.setTimeout(180_000)
      for (const t of TRANSITIONS) {
        const r = await transitionRun(page, t, sync)
        expect(r.cut).toBeUndefined()
        expect(r.backing).toEqual({ w: 1920, h: 1080, dpr: 1 })
        expect.soft(r.p95, `${t.effect} p95 (${r.renderer})`).toBeLessThanOrEqual(FRAME_BUDGET_MS)
      }
    })
  }
})

// Profiling only (not the product budget): with a 1 GiB workspace the 3840 × 2160 transition frames fit, so the
// per-frame cost of re-tracing and re-uploading a device-sized split mask can be measured at the benchmark size.
test.describe('transition mask re-upload at 1920 × 1080, DPR 2, 1 GiB workspace (profiling)', () => {
  for (const sync of [false, true]) {
    test(`split (mask) vs fade (transform) at 3840 × 2160 (${sync ? 'through a GPU readback' : 'submit only'})`, async ({ page }) => {
      test.setTimeout(180_000)
      for (const t of TRANSITIONS) {
        const r = await transitionRun(page, t, sync, '&budget=wide')
        expect(r.cut).toBeUndefined()
        expect(r.backing).toEqual({ w: 3840, h: 2160, dpr: 2 })
        expect.soft(r.p95, `${t.effect} p95 (${r.renderer})`).toBeLessThanOrEqual(FRAME_BUDGET_MS)
      }
    })
  }
})
