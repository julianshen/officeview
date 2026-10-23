/**
 * Slideshow plan S3-4: the deformation 3-D transitions in Chromium (the
 * `slideshow` Playwright project: SwiftShader WebGL2), driven through
 * `window.__slideshowDeform` (src/debug/slideshow-3d-deform.ts), which
 * registers the group in that page only.
 *
 * It checks the GLSL against the CPU models that tests/slideshow-3d-deform.test.ts
 * pins to the algorithm sheet: every shader compiles and draws without a
 * fallback, matches the flat frames just inside its ends, reproduces the
 * sheet's 50 % oracles in pixels (fold fractions; edge positions), changes
 * with its attributes and seed, and routes to its look-alike on Canvas 2D.
 * Behaviour only, never frame times. The frame goldens (0 / 25 / 50 / 75 /
 * 100 %) are Mac-only, behind OFFICEVIEW_STRICT_GOLDEN=1.
 */
import { expect, test } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'
import { DEFAULT_FOV_DEG, mul4, project, slideCamera, slidePoint, type Mat4 } from '../../src/slideshow/render/three/mesh'
import { WIND, airplaneModel, airplaneSilhouette, curtainVertex, drapeVertex, fallOverModel, origamiModel, prestigeModel, warpK, windTile } from '../../src/slideshow/render/three/deform'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-3d-deform.html'
const A = 16 / 9
const C = slideCamera(DEFAULT_FOV_DEG, A)

export type { SlideshowDeform } from '../../src/debug/slideshow-3d-deform'

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
  await page.waitForFunction(() => window.__slideshowDeform !== undefined)
})

test.afterEach(() => {
  expect(errors, 'uncaught page errors').toEqual([])
})

const NAMES = ['pageCurlSingle', 'pageCurlDouble', 'peelOff', 'airplane', 'origami', 'curtains', 'drape', 'wind', 'fallOver', 'prestige', 'warp']
type Rgba = number[]
const isBlue = (c: Rgba) => c[2] > 60 && c[2] > 2 * c[0] && c[1] < 40
const isRed = (c: Rgba) => c[0] > 60 && c[0] > 2 * c[2] && c[1] < 40
const ndc = (model: Mat4, u: number, v: number) => project(mul4(C, model), slidePoint(u, v, A))
const worldNdc = (p: readonly number[]) => project(C, [p[0], p[1], p[2]])

test('every deformation shader compiles, draws, and matches the flat frames just inside its ends', async ({ page }) => {
  expect(await page.evaluate(() => window.__slideshowDeform!.rows())).toEqual(NAMES)
  const cases: Array<[string, Record<string, string>]> = [...NAMES.map(n => [n, {}] as [string, Record<string, string>]), ['warp', { dir: 'out' }], ['wind', { invX: '1' }], ['pageCurlSingle', { invX: '1', invY: '1' }]]
  for (const [name, attrs] of cases) {
    const rows = await page.evaluate(([n, a]) => window.__slideshowDeform!.ends(n, a), [name, attrs] as const)
    console.log(`[slideshow-3d-deform] ends ${name} ${JSON.stringify(attrs)} ${JSON.stringify(rows.map(r => [r.progress, r.meanAbs.toFixed(4), r.rowMax.toFixed(4)]))}`)
    for (const r of rows) {
      expect(r.diagnostics, `${name} @${r.progress}`).toEqual([])
      expect(r.shaderCalls, `${name} @${r.progress}`).toBe(1)
      expect(r.ok, `${name} ${r.which} @${r.progress}: meanAbs ${r.meanAbs}, rowMax ${r.rowMax}`).toBe(true)
    }
  }
})

test('the fold rows reproduce the sheet’s 50 % fractions in pixels', async ({ page }) => {
  for (const [name, attrs, to, back] of [
    ['pageCurlSingle', {}, 0.5, 0.3291], ['pageCurlSingle', { invX: '1' }, 0.5, 0.3291],
    ['pageCurlDouble', {}, 0.5, 0.4475], ['peelOff', {}, 0.5, 0.3291],
  ] as Array<[string, Record<string, string>, number, number]>) {
    const r = await page.evaluate(([n, a]) => window.__slideshowDeform!.fractions(n, a, 0.5), [name, attrs] as const)
    console.log(`[slideshow-3d-deform] fractions ${name} ${JSON.stringify(attrs)} ${JSON.stringify(r)}`)
    expect(r.diagnostics).toEqual([])
    expect(Math.abs(r.to - to), `${name} to ${r.to}`).toBeLessThan(0.01)
    expect(Math.abs(r.back - back), `${name} back ${r.back}`).toBeLessThan(0.01)
    expect(r.black).toBeLessThan(0.002)
  }
})

/** Probes `pairs` of [inside-from point, inside-to point] in NDC at `p`. */
async function edges(page: import('@playwright/test').Page, name: string, attrs: Record<string, string>, p: number, pairs: Array<[[number, number], [number, number]]>, fromTest = isBlue) {
  const r = await page.evaluate(([n, a, pp, pts]) => window.__slideshowDeform!.probe(n, a, pp, pts), [name, attrs, p, pairs.flat()] as const)
  expect(r.diagnostics).toEqual([])
  pairs.forEach(([f, t], i) => {
    expect(fromTest(r.colors[2 * i]), `${name} @${p}: ${JSON.stringify(f)} shows from (${r.colors[2 * i]})`).toBe(true)
    expect(isRed(r.colors[2 * i + 1]), `${name} @${p}: ${JSON.stringify(t)} shows to (${r.colors[2 * i + 1]})`).toBe(true)
  })
}

test('the rigid and deformed rows put their edges where the sheet’s 50 % oracles say', async ({ page }) => {
  const d = 0.03
  // curtains: the halves' inner edges at x = ∓0.3306 (sheet), z = 0 there.
  const inner = worldNdc(curtainVertex(0.5, 0.5, 0.5, A))[0]
  expect(curtainVertex(0.5, 0.5, 0.5, A)[0]).toBeCloseTo(-0.3306, 3)
  await edges(page, 'curtains', {}, 0.5, [[[inner - d, 0], [inner + d, 0]], [[-inner + d, 0], [-inner - d, 0]]])
  // drape: the top edge at y = −0.06 (sheet).
  const top = worldNdc(drapeVertex(0.5, 0.5, 0, A))[1]
  expect(drapeVertex(0.5, 0.5, 0, A)[1]).toBeCloseTo(-0.06, 3)
  await edges(page, 'drape', {}, 0.5, [[[0, top - d], [0, top + d]]])
  // warp in: the right edge pulled to world x = 0.1296 (sheet); to is flat at 0.85 behind it.
  const right = worldNdc([(A / 2) * warpK(0.5, A / 2), 0, 0])[0]
  expect(right).toBeCloseTo(0.1296 / (A / 2), 3)
  await edges(page, 'warp', {}, 0.5, [[[right - d, 0], [right + d, 0]]])
  // fallOver: from's top edge at NDC (−0.0401, 0.2584) (sheet).
  const fo = ndc(fallOverModel(0.5, false).from, 0.5, 0)
  expect(fo[1]).toBeCloseTo(0.2584, 3)
  await edges(page, 'fallOver', {}, 0.5, [[[fo[0], fo[1] - d], [fo[0], fo[1] + d]]])
  // prestige: to's top edge at NDC y = −0.1 (sheet), from at half opacity above it.
  const pt = ndc(prestigeModel(0.5).to, 0.5, 0)
  expect(pt[1]).toBeCloseTo(-0.1, 3)
  const halfBlue = (c: Rgba) => c[2] > 100 && c[2] < 160 && c[0] < 20
  await edges(page, 'prestige', {}, 0.5, [[[0, pt[1] + d], [0, pt[1] - d]]], halfBlue)
  // origami: panel 3's free edge at mid-height (its top corner is the sheet's NDC (0.5208, 1)).
  expect(ndc(origamiModel(0.5, A, false).panels[3], 1, 0)[0]).toBeCloseTo(0.5208, 3)
  const oe = ndc(origamiModel(0.5, A, false).panels[3], 1, 0.5)
  await edges(page, 'origami', {}, 0.5, [[[oe[0] - d, 0], [oe[0] + d, 0]]])
  // airplane at 50 %: a folded paper dart, nose at the bottom centre (sheet: NDC (0, −1)); its pixels cover its outline (sheet area 0.93 of 4).
  const nose = ndc(airplaneModel(0.5, A, false).panels.bodyL, 0.5, 1)
  expect(nose[1]).toBeCloseTo(-1, 3)
  await edges(page, 'airplane', {}, 0.5, [[[0, -0.85], [-0.3, -0.92]], [[0, -0.6], [0.3, -0.92]]])
  const dart = await page.evaluate(() => window.__slideshowDeform!.fractions('airplane', {}, 0.5))
  const outline = airplaneSilhouette(0.5, A, false).area / 4
  expect(Math.abs(dart.from - outline), `dart ${dart.from} vs outline ${outline}`).toBeLessThan(0.02)
  // In flight (75 %) it is still a visible dart, never an edge-on sliver.
  const flying = await page.evaluate(() => window.__slideshowDeform!.fractions('airplane', {}, 0.75))
  expect(flying.from).toBeGreaterThan(0.01)
})

test('wind’s tiles land where the CPU tile model puts them', async ({ page }) => {
  const p = 0.5
  const { seed } = await page.evaluate(() => window.__slideshowDeform!.probe('wind', {}, 0.5, []))
  expect(typeof seed).toBe('number')
  // Every tile's projected quad and opacity at p.
  const quads: Array<{ pts: number[][]; opacity: number }> = []
  for (let j = 0; j < WIND.rows; j++) for (let i = 0; i < WIND.cols; i++) {
    const t = windTile(p, i, j, seed!, A, false)
    if (t.opacity <= 0) continue
    const u0 = i / WIND.cols, u1 = (i + 1) / WIND.cols, v0 = j / WIND.rows, v1 = (j + 1) / WIND.rows
    quads.push({ pts: [[u0, v0], [u1, v0], [u1, v1], [u0, v1]].map(([u, v]) => ndc(t.model, u, v)), opacity: t.opacity })
  }
  const inside = (q: number[][], x: number, y: number) => {
    let sign = 0
    for (let k = 0; k < 4; k++) {
      const [ax, ay] = q[k], [bx, by] = q[(k + 1) % 4]
      const c = (bx - ax) * (y - ay) - (by - ay) * (x - ax)
      if (Math.abs(c) < 1e-12) continue
      if (sign === 0) sign = Math.sign(c)
      else if (Math.sign(c) !== sign) return false
    }
    return true
  }
  const expectAt = (x: number, y: number): 'from' | 'to' | undefined => {
    const cover = quads.filter(q => inside(q.pts, x, y))
    if (cover.length === 0) return 'to'
    if (cover.length === 1 && cover[0].opacity > 0.95) return 'from'
    return undefined
  }
  // Sample points whose expectation is the same 1.5 px around them (no edge anti-aliasing).
  const px = 2 / 320, py = 2 / 180
  const points: Array<[number, number]> = [], want: string[] = []
  for (let gy = 0; gy < 24; gy++) for (let gx = 0; gx < 40; gx++) {
    const x = -1 + (gx + 0.5) * (2 / 40), y = 1 - (gy + 0.5) * (2 / 24)
    const e = expectAt(x, y)
    if (!e || [[1.5, 0], [-1.5, 0], [0, 1.5], [0, -1.5]].some(([dx, dy]) => expectAt(x + dx * px, y + dy * py) !== e)) continue
    points.push([x, y]); want.push(e)
  }
  expect(want.filter(w => w === 'from').length).toBeGreaterThan(20)
  expect(want.filter(w => w === 'to').length).toBeGreaterThan(20)
  const r = await page.evaluate(([pts]) => window.__slideshowDeform!.probe('wind', {}, 0.5, pts), [points] as const)
  expect(r.seed).toBe(seed)
  const got = r.colors.map(c => (c[2] > c[0] ? 'from' : 'to'))
  const mismatches = got.filter((g, i) => g !== want[i]).length
  console.log(`[slideshow-3d-deform] wind ${points.length} points, ${mismatches} mismatches`)
  expect(mismatches / points.length).toBeLessThan(0.02)
})

test('direction, option attributes and the effect seed change the output', async ({ page }) => {
  for (const [name, a, b, p, seeds] of [
    ['pageCurlSingle', {}, { invX: '1' }, 0.5], ['pageCurlSingle', {}, { invY: '1' }, 0.5], ['pageCurlDouble', {}, { invY: '1' }, 0.5],
    ['peelOff', {}, { invX: '1' }, 0.5], ['airplane', {}, { invX: '1' }, 0.75], ['origami', {}, { invX: '1' }, 0.5],
    ['wind', {}, { invX: '1' }, 0.5], ['fallOver', {}, { invX: '1' }, 0.5], ['warp', {}, { dir: 'out' }, 0.3],
    ['wind', {}, {}, 0.5, [21, 22]],
  ] as Array<[string, Record<string, string>, Record<string, string>, number, [number, number]?]>) {
    const r = await page.evaluate(([n, x, y, pp, s]) => window.__slideshowDeform!.differs(n, x, y, pp, s), [name, a, b, p, seeds ?? [21, 21]] as const)
    expect(r.identical, `${name} ${JSON.stringify(a)} vs ${JSON.stringify(b)} ${seeds ?? ''}`).toBe(false)
    // airplane at 75 % is an edge-on sliver, so its mirror moves few pixels.
    expect(r.meanAbs, `${name} ${JSON.stringify(b)}`).toBeGreaterThan(name === 'airplane' ? 0.001 : 0.005)
  }
  // The same seed draws the identical frame.
  const same = await page.evaluate(() => window.__slideshowDeform!.differs('wind', {}, {}, 0.5, [21, 21]))
  expect(same.identical).toBe(true)
})

test('Canvas 2D draws each deformation row’s catalogued look-alike with a diagnostic', async ({ page }) => {
  for (const name of NAMES) {
    const r = await page.evaluate(n => window.__slideshowDeform!.canvas2d(n), name)
    expect(r.route).toBe('lookalike')
    expect(r.diagnostics).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 2 })])
    expect(r.diagnostics[0].message).toContain(name === 'curtains' ? 'look-alike split' : 'look-alike fade')
  }
})

test.describe('frame goldens (Mac, OFFICEVIEW_STRICT_GOLDEN=1)', () => {
  test.skip(process.platform !== 'darwin' || process.env.OFFICEVIEW_STRICT_GOLDEN !== '1', 'GPU goldens are Mac-only and opt-in')
  for (const name of NAMES) {
    test(`${name} at 0 / 25 / 50 / 75 / 100 %`, async ({ page }) => {
      expect(await page.evaluate(n => window.__slideshowDeform!.strip(n), name)).toBe(5)
      await expect(page.locator('#strip')).toHaveScreenshot(`deform-${name}.png`, { maxDiffPixelRatio: 0.01 })
    })
  }
})
