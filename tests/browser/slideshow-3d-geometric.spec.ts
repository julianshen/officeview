/**
 * Slideshow plan S3-2: the geometric 3-D transitions in Chromium (the
 * `slideshow` Playwright project: SwiftShader WebGL2), driven through
 * `window.__slideshow3dGeometric` (src/debug/slideshow-3d-geometric.ts).
 *
 * Behaviour on both paths, never GPU frame times: every row's ends match the
 * flat frames, its middle is not the look-alike and puts the top piece where
 * the model does (the unit tests hold the model to the sheet's oracles),
 * attributes change the frame, and Canvas 2D draws the catalogued look-alike
 * with a diagnostic.
 *
 * Frame goldens at 0/25/50/75/100 % (tests/goldens/slideshow-3d/) are Mac
 * only and never compared in CI: a missing golden is recorded,
 * OFFICEVIEW_STRICT_GOLDEN=1 fails instead.
 *
 * `OV_SLIDESHOW_BENCH=1 OV_SLIDESHOW_GPU=1` adds an informal frame-time
 * measurement on the host GPU (its own Chromium, without SwiftShader). It is
 * never part of a normal run and asserts nothing about the numbers.
 */
import { chromium, expect, test } from '@playwright/test'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-3d-geometric.html'
const GOLDEN_DIR = join(ROOT, 'tests', 'goldens', 'slideshow-3d')
const STRICT = process.env.OFFICEVIEW_STRICT_GOLDEN === '1'

export type { Slideshow3DGeometric } from '../../src/debug/slideshow-3d-geometric'

const ELEMENTS = ['prism', 'flip', 'gallery', 'doors', 'window', 'switch', 'conveyor', 'ferris']
/** Every attribute variant with its own geometry. */
const VARIANTS: Array<[string, Record<string, string>]> = [
  ['prism', {}], ['prism', { dir: 'r' }], ['prism', { dir: 'u' }], ['prism', { dir: 'd' }], ['prism', { isInverted: '1' }], ['prism', { isContent: '1' }],
  ['flip', {}], ['flip', { dir: 'r' }],
  ['gallery', {}], ['gallery', { dir: 'r' }],
  ['doors', {}], ['doors', { dir: 'horz' }],
  ['window', {}], ['window', { dir: 'horz' }],
  ['switch', {}], ['switch', { dir: 'r' }],
  ['conveyor', {}], ['conveyor', { dir: 'r' }],
  ['ferris', {}], ['ferris', { dir: 'r' }],
]

const errors: string[] = []

test.beforeAll(async () => {
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity').toBeNull()
})

test.beforeEach(async ({ page }) => {
  errors.length = 0
  page.on('pageerror', e => errors.push(e.message))
  await page.goto(PAGE)
  await page.waitForFunction(() => window.__slideshow3dGeometric !== undefined)
})

test.afterEach(() => {
  expect(errors, 'uncaught page errors').toEqual([])
})

test('every geometric row reproduces the outgoing and incoming frames at its ends', async ({ page }) => {
  expect(await page.evaluate(() => window.__slideshow3dGeometric!.elements())).toEqual(ELEMENTS)
  for (const [element, attrs] of VARIANTS) {
    const rows = await page.evaluate(([e, a]) => window.__slideshow3dGeometric!.ends(e, a), [element, attrs] as const)
    console.log(`[slideshow-3d-geometric] ends ${element} ${JSON.stringify(attrs)} ${JSON.stringify(rows.map(r => [r.progress, +r.meanAbs.toFixed(5), +r.rowMax.toFixed(5)]))}`)
    for (const r of rows) {
      expect(r.ok, `${element} ${JSON.stringify(attrs)} @${r.progress} vs ${r.which}: meanAbs ${r.meanAbs}, rowMax ${r.rowMax}`).toBe(true)
      expect(r.diagnostics).toEqual([])
      expect(r.shaderCalls).toBe(r.progress === 0 || r.progress === 1 ? 0 : 1)
    }
  }
})

test('mid-way every row draws its 3-D model, not the look-alike, and attributes change the frame', async ({ page }) => {
  const digests = new Map<string, string>()
  for (const [element, attrs] of VARIANTS) {
    const rows = await page.evaluate(([e, a]) => window.__slideshow3dGeometric!.middle(e, a), [element, attrs] as const)
    console.log(`[slideshow-3d-geometric] middle ${element} ${JSON.stringify(attrs)} ${JSON.stringify(rows.map(r => ({ p: r.progress, vs: +r.vsLookalike.toFixed(3), probes: r.probes, matched: r.matched, worst: r.worst })))}`)
    for (const r of rows) {
      const what = `${element} ${JSON.stringify(attrs)} @${r.progress}`
      expect(r.shaderCalls, what).toBe(1)
      expect(r.diagnostics, what).toEqual([])
      expect(r.letterbox, what).toEqual([32, 32, 32, 255])
      expect(r.vsLookalike, what).toBeGreaterThan(0.02)
      // The top piece shows its frame where the model projects it (most probes; texture filtering and shading blur a few).
      if (element === 'flip' && r.progress === 0.5) {
        // The card is exactly edge-on: nothing to probe, and the slide is the black backdrop.
        expect(r.probes, what).toBe(0)
        expect(r.centre, what).toEqual([0, 0, 0, 255])
        continue
      }
      expect(r.probes, what).toBeGreaterThanOrEqual(4)
      expect(r.matched / r.probes, `${what}: ${r.matched}/${r.probes}, worst ${r.worst}`).toBeGreaterThanOrEqual(0.8)
    }
    digests.set(`${element} ${JSON.stringify(attrs)}`, rows.map(r => r.digest).join('|'))
  }
  // Each attribute variant draws a different frame from its row's default.
  for (const [element, attrs] of VARIANTS) {
    if (Object.keys(attrs).length === 0) continue
    expect(digests.get(`${element} ${JSON.stringify(attrs)}`), `${element} ${JSON.stringify(attrs)}`).not.toBe(digests.get(`${element} {}`))
  }
})

test('Canvas 2D draws each row\'s catalogued look-alike with a diagnostic', async ({ page }) => {
  for (const element of ELEMENTS) {
    const r = await page.evaluate(e => window.__slideshow3dGeometric!.canvas2d(e), element)
    expect(r.route, element).toBe('lookalike')
    expect(r.glRoute, `${element}: the WebGL2 compositor (blocking; escalate to the owner if this fails in CI)`).toBe('shader')
    expect(r.glDiagnostics, element).toBe(0)
    expect(r.lookalike, element).toBe(element === 'doors' || element === 'window' ? 'split' : 'push')
    expect(r.diagnostics, element).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 4, effect: element })])
    expect(r.diagnostics[0].message, element).toContain(r.lookalike!)
    for (const x of r.results) {
      expect(x.shaderCalls, element).toBe(0)
      expect(x.ok, `${element} @${x.progress}: meanAbs ${x.meanAbs}, rowMax ${x.rowMax}`).toBe(true)
    }
  }
})

test('frame goldens at 0/25/50/75/100 % (Mac only, never in CI)', async ({ page }) => {
  test.skip(process.platform !== 'darwin' || !!process.env.CI, 'GPU frame goldens are Mac-only and never compared in CI')
  mkdirSync(GOLDEN_DIR, { recursive: true })
  const recorded: string[] = []
  for (const element of ELEMENTS) {
    for (const p of [0, 0.25, 0.5, 0.75, 1]) {
      const file = join(GOLDEN_DIR, `geometric-${element}-${Math.round(p * 100)}.png`)
      if (!existsSync(file)) {
        expect(STRICT, `missing golden ${file}`).toBe(false)
        const f = await page.evaluate(([e, q]) => window.__slideshow3dGeometric!.frame(e, q), [element, p] as const)
        writeFileSync(file, Buffer.from(f.png.split(',')[1], 'base64'))
        recorded.push(file)
        continue
      }
      const golden = `data:image/png;base64,${readFileSync(file).toString('base64')}`
      const d = await page.evaluate(([e, q, g]) => window.__slideshow3dGeometric!.compare(e, q, g), [element, p, golden] as const)
      expect(d.size, file).toEqual([160, 98])
      expect(d.meanAbs, `${file}: meanAbs ${d.meanAbs}, rowMax ${d.rowMax}`).toBeLessThanOrEqual(0.004)
      expect(d.rowMax, `${file}: meanAbs ${d.meanAbs}, rowMax ${d.rowMax}`).toBeLessThanOrEqual(0.02)
    }
  }
  if (recorded.length) console.log(`[slideshow-3d-geometric] recorded ${recorded.length} goldens:\n${recorded.join('\n')}`)
})

test('informal frame timings on the host GPU (OV_SLIDESHOW_BENCH=1 OV_SLIDESHOW_GPU=1 only)', async () => {
  test.skip(process.env.OV_SLIDESHOW_BENCH !== '1' || process.env.OV_SLIDESHOW_GPU !== '1', 'runs only on request, on a real GPU')
  // Its own Chromium without the project's SwiftShader flags, so WebGL2 runs on the host GPU.
  const browser = await chromium.launch({ args: ['--enable-gpu', '--ignore-gpu-blocklist', '--use-angle=metal'] })
  try {
    const page = await browser.newPage()
    await page.goto(`${BASE}${PAGE}`)
    await page.waitForFunction(() => window.__slideshow3dGeometric !== undefined)
    const rows = []
    for (const [w, h] of [[1920, 1080], [3840, 2160]]) {
      for (const element of ELEMENTS) rows.push(await page.evaluate(([e, ww, hh]) => window.__slideshow3dGeometric!.bench(e, ww, hh, 240), [element, w, h] as const))
    }
    for (const r of rows) console.log(`[slideshow-3d-geometric] bench ${r.devW}x${r.devH} ${r.element} (${r.route}) p50 ${r.p50.toFixed(2)} ms, p95 ${r.p95.toFixed(2)} ms, max ${r.max.toFixed(2)} ms on ${r.renderer}`)
    for (const r of rows) expect(r.route).toBe('shader')
  } finally {
    await browser.close()
  }
})
