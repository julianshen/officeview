/**
 * Slideshow plan S3-3: the tile and particle 3-D transitions in Chromium (the
 * `slideshow` Playwright project: SwiftShader WebGL2), driven through
 * `window.__slideshowParticles` (src/debug/slideshow-3d-particles.ts).
 *
 * What it holds the shaders to:
 * - the algorithm sheet's reference models (scripts/slideshow-3d-oracles.ts):
 *   each row's own vertex shader runs under transform feedback on every cell
 *   (centre and corner) and must give the model's time, opacity, shade and
 *   world position, and the sheet's literal 50 % numbers;
 * - the picture: at 50 % a tile whose centre no other tile covers must show
 *   the tile's source (or the backdrop, by its opacity) at the projected
 *   centre, which proves the draw, the textures, the blending and the camera;
 * - the ends: just inside 0 and 1 the frame matches the flat frames;
 * - attributes, seeds and the Canvas 2D look-alike, never frame times (those
 *   are measured only with OV_SLIDESHOW_BENCH=1 OV_SLIDESHOW_GPU=1).
 *
 * `flash` and `reveal` are exact 2-D rows and are not part of this group.
 */
import { expect, test, chromium, type Page } from '@playwright/test'
import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'
import {
  ASPECT, ROWS, SEED, crushTile, fractureTile, glitterTile, honeycombTile, rippleZ, shredTile, vortexTile,
} from '../../scripts/slideshow-3d-oracles'
import { DEFAULT_FOV_DEG, ease, mul4, project, slideCamera, slidePoint } from '../../src/slideshow/render/three/mesh'
import { particlePlan } from '../../src/slideshow/render/three/particles'
import { frameDiff, withinTolerance } from '../../src/debug/slideshow-parity'
import { createCanvas, loadImage } from 'canvas'
import type { ProbeResult, ProbeVertex, RenderResult } from '../../src/debug/slideshow-3d-particles'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-3d-particles.html'
const GOLDENS = join(dirname(fileURLToPath(import.meta.url)), '..', 'goldens', 'slideshow-3d-particles')

type Attrs = Record<string, string>
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
  await open(page)
})

test.afterEach(() => {
  expect(errors, 'uncaught page errors').toEqual([])
})

/** PNG IO through node-canvas directly (src/test/pixel-diff.ts needs a CommonJS `require`, which the Playwright loader lacks). */
async function savePng(data: Uint8ClampedArray, path: string) {
  const c = createCanvas(DW, DH)
  const ctx = c.getContext('2d')
  const img = ctx.createImageData(DW, DH)
  img.data.set(data)
  ctx.putImageData(img, 0, 0)
  writeFileSync(path, c.toBuffer('image/png'))
}
async function loadPng(path: string): Promise<Uint8ClampedArray> {
  const img = await loadImage(path)
  const c = createCanvas(DW, DH)
  const ctx = c.getContext('2d')
  ctx.drawImage(img, 0, 0)
  return ctx.getImageData(0, 0, DW, DH).data as unknown as Uint8ClampedArray
}

async function open(page: Page) {
  await page.goto(PAGE)
  await page.waitForFunction(() => window.__slideshowParticles !== undefined)
}

const KEYS = ['vortex', 'shred', 'fracture', 'crush', 'glitter', 'honeycomb', 'ripple'] as const
type Key = (typeof KEYS)[number]

/** The attribute variants of each row. */
const VARIANTS: Record<Key, Attrs[]> = {
  vortex: [{ dir: 'l' }, { dir: 'u' }, { dir: 'r' }, { dir: 'd' }],
  shred: [{ pattern: 'strip' }, { pattern: 'rectangle' }, { pattern: 'strip', dir: 'in' }, { pattern: 'rectangle', dir: 'in' }],
  fracture: [{}],
  crush: [{}],
  glitter: [{ dir: 'l' }, { dir: 'u' }, { dir: 'r' }, { dir: 'd' }, { pattern: 'hexagon', dir: 'l' }],
  honeycomb: [{}],
  ripple: [{ dir: 'center' }, { dir: 'lu' }, { dir: 'ru' }, { dir: 'ld' }, { dir: 'rd' }],
}
const ELEMENT: Record<Key, { element: string; preset?: string }> = {
  vortex: { element: 'vortex' }, shred: { element: 'shred' }, fracture: { element: 'prstTrans', preset: 'fracture' },
  crush: { element: 'prstTrans', preset: 'crush' }, glitter: { element: 'glitter' }, honeycomb: { element: 'honeycomb' }, ripple: { element: 'ripple' },
}
const ROW_ID: Record<Key, string> = {
  vortex: 'transition:p14:vortex', shred: 'transition:p14:shred', fracture: 'transition:p15:prstTrans:fracture', crush: 'transition:p15:prstTrans:crush',
  glitter: 'transition:p14:glitter', honeycomb: 'transition:p14:honeycomb', ripple: 'transition:p14:ripple',
}
const LOOKALIKE: Record<Key, string> = { vortex: 'dissolve', shred: 'dissolve', fracture: 'fade', crush: 'fade', glitter: 'dissolve', honeycomb: 'checker', ripple: 'dissolve' }
const label = (k: string, a: Attrs) => `${k} ${JSON.stringify(a)}`

// The harness geometry: a 192 × 108 slide at (8, 8) on a 208 × 124 output.
const SW = 192, SH = 108, DW = 208, DH = 124, OX = 8, OY = 8
const CAM = slideCamera(DEFAULT_FOV_DEG, ASPECT)

async function render(page: Page, key: Key, p: number, attrs: Attrs = {}, extra: { backend?: 'webgl2' | 'canvas2d'; rngSeed?: number } = {}): Promise<RenderResult & { data: Uint8ClampedArray }> {
  const r = await page.evaluate(o => window.__slideshowParticles!.render(o), { key, p, attrs, ...extra })
  return { ...r, data: Uint8ClampedArray.from(r.pixels) }
}
async function probe(page: Page, key: Key, attrs: Attrs, p: number, vertices: ProbeVertex[], seed = SEED, pass = 0): Promise<ProbeResult[]> {
  const shader = key
  const { element, preset } = ELEMENT[key]
  return page.evaluate(o => window.__slideshowParticles!.probe(o), { key: shader, element, preset, attrs, p, aspect: ASPECT, seed, vertices, pass })
}
const diff = (a: Uint8ClampedArray, b: Uint8ClampedArray) => frameDiff(a, b, DW, DH)

// ---------------------------------------------------------------------------
// The reference models, per row
// ---------------------------------------------------------------------------
interface Ref { i: number; j: number; t: number; model: Float32Array; opacity: number; shade: number; showsTo: boolean }

function refCells(key: Key, attrs: Attrs, p: number, seed: number): { grid: readonly [number, number]; cells: Ref[] } {
  const { element, preset } = ELEMENT[key]
  const plan = particlePlan(element, preset, attrs, p, ASPECT, seed)
  const [cols, rows] = plan.grid
  const cells: Ref[] = []
  for (let j = 0; j < rows; j++) {
    for (let i = 0; i < cols; i++) {
      let r: { t: number; model: Float32Array; opacity?: number; shade?: number; showsTo?: boolean }
      switch (key) {
        case 'vortex': r = vortexTile(p, i, j, seed, attrs.dir ?? 'l'); break
        case 'shred': r = shredTile(attrs.dir === 'in' ? 1 - p : p, i, j, attrs.pattern ?? 'strip', seed); break
        case 'fracture': r = fractureTile(p, i, j, seed); break
        case 'crush': r = crushTile(p, i, j, seed); break
        case 'glitter': r = glitterTile(p, i, j, attrs.pattern ?? 'diamond', attrs.dir ?? 'l', seed); break
        case 'honeycomb': r = honeycombTile(p, i, j, seed); break
        default: throw new Error(`${key} has no tiles`)
      }
      cells.push({ i, j, t: r.t, model: r.model, opacity: r.opacity ?? 1, shade: r.shade ?? 1, showsTo: r.showsTo ?? false })
    }
  }
  return { grid: plan.grid, cells }
}

const apply = (m: Float32Array, q: readonly number[]): [number, number, number] =>
  [0, 1, 2].map(r => m[r] * q[0] + m[4 + r] * q[1] + m[8 + r] * q[2] + m[12 + r]) as [number, number, number]
const ndcOf = (r: { clip: readonly number[] }): [number, number] => [r.clip[0] / r.clip[3], r.clip[1] / r.clip[3]]
const dev = (n: readonly number[]): [number, number] => [OX + ((n[0] + 1) / 2) * SW, OY + ((1 - n[1]) / 2) * SH]

/** Per cell: a probe vertex at the cell centre, then one at its top-left corner (both with the cell's centre as `a_cell`). */
function cellVertices(grid: readonly [number, number]): ProbeVertex[] {
  const [cols, rows] = grid
  const out: ProbeVertex[] = []
  for (let j = 0; j < rows; j++) {
    for (let i = 0; i < cols; i++) {
      const cell: [number, number] = [(i + 0.5) / cols, (j + 0.5) / rows]
      out.push({ uv: cell, cell }, { uv: [i / cols, j / rows], cell })
    }
  }
  return out
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------
test('every row in the group has a registered shader and a lookalike', async ({ page }) => {
  for (const key of KEYS) {
    const gl = await render(page, key, 0.5)
    expect(gl.route, key).toBe('shader')
    expect(gl.diagnostics, key).toEqual([])
    const c2 = await render(page, key, 0.5, {}, { backend: 'canvas2d' })
    expect(c2.route, key).toBe('lookalike')
    expect(c2.diagnostics, key).toEqual([expect.objectContaining({ code: 'transition-approximate' })])
    expect(c2.diagnostics[0].message, key).toContain(LOOKALIKE[key])
    // The shader is not the look-alike.
    const d = diff(gl.data, c2.data)
    expect(d.meanAbs, `${key} differs from its look-alike`).toBeGreaterThan(0.01)
    expect(ROW_ID[key]).toBeTruthy()
  }
})

/** Far tighter than the catalogue's: a mirrored back face or a stray tile moves the mean by 0.005 and more. */
const ENDS = { meanAbs: 0.002, rowMax: 0.03 }

test('the ends match the flat frames just inside 0 and 1', async ({ page }) => {
  const rows: string[] = []
  for (const key of KEYS) {
    for (const attrs of VARIANTS[key]) {
      const from = await render(page, key, 0, attrs), to = await render(page, key, 1, attrs)
      const near0 = await render(page, key, 0.001, attrs), near1 = await render(page, key, 0.999, attrs)
      expect(near0.diagnostics, label(key, attrs)).toEqual([])
      const d0 = diff(near0.data, from.data), d1 = diff(near1.data, to.data)
      rows.push(`${label(key, attrs)} @0.001 ${d0.meanAbs.toFixed(5)}/${d0.rowMax.toFixed(5)} @0.999 ${d1.meanAbs.toFixed(5)}/${d1.rowMax.toFixed(5)}`)
      // The exact flat frames at 0 and 1 are the two textures.
      expect(diff(from.data, to.data).meanAbs, `${label(key, attrs)} flat frames differ`).toBeGreaterThan(0.1)
      expect(withinTolerance(d0, ENDS), `${label(key, attrs)} @0.001 ${JSON.stringify(d0)}`).toBe(true)
      expect(withinTolerance(d1, ENDS), `${label(key, attrs)} @0.999 ${JSON.stringify(d1)}`).toBe(true)
    }
  }
  console.log(`[slideshow-3d-particles] ends\n${rows.join('\n')}`)
})

test('direction and option attributes change the output', async ({ page }) => {
  for (const key of KEYS) {
    const frames: Array<[Attrs, Uint8ClampedArray]> = []
    for (const attrs of VARIANTS[key]) frames.push([attrs, (await render(page, key, 0.5, attrs)).data])
    for (let a = 0; a < frames.length; a++) {
      for (let b = a + 1; b < frames.length; b++) {
        const d = diff(frames[a][1], frames[b][1])
        expect(d.meanAbs, `${label(key, frames[a][0])} vs ${label(key, frames[b][0])}`).toBeGreaterThan(0.002)
      }
    }
  }
  // Defaults: omitted attributes are the default value, and unknown values fall back to it.
  const same = async (key: Key, a: Attrs, b: Attrs) => {
    const x = await render(page, key, 0.5, a), y = await render(page, key, 0.5, b)
    expect(diff(x.data, y.data).meanAbs, `${label(key, a)} vs ${label(key, b)}`).toBe(0)
  }
  await same('vortex', {}, { dir: 'l' })
  await same('vortex', { dir: 'sideways' }, { dir: 'l' })
  await same('shred', {}, { pattern: 'strip', dir: 'out' })
  await same('glitter', {}, { pattern: 'diamond', dir: 'l' })
  await same('ripple', {}, { dir: 'center' })
  await same('honeycomb', {}, { dir: 'r' })
})

test('seeded effects are deterministic', async ({ page }) => {
  for (const key of KEYS) {
    const attrs = VARIANTS[key][VARIANTS[key].length > 1 ? 1 : 0]
    const a = await render(page, key, 0.5, attrs, { rngSeed: 5 }), b = await render(page, key, 0.5, attrs, { rngSeed: 5 })
    expect(a.seed, key).toBe(b.seed)
    expect(diff(a.data, b.data).meanAbs, `${key} repeats`).toBe(0)
    if (key === 'ripple') continue // no seeded parameters
    const c = await render(page, key, 0.5, attrs, { rngSeed: 6 })
    expect(c.seed, key).not.toBe(a.seed)
    expect(diff(a.data, c.data).meanAbs, `${key} with another seed`).toBeGreaterThan(0.002)
  }
})

/** The sheet's 50 % numbers (docs/validation/slideshow-3d-algorithms.md). */
const ndcClose = (got: readonly number[], want: readonly [number, number], what: string) => {
  expect(Math.abs(got[0] - want[0]), `${what} x ${got[0]} vs ${want[0]}`).toBeLessThan(1e-3)
  expect(Math.abs(got[1] - want[1]), `${what} y ${got[1]} vs ${want[1]}`).toBeLessThan(1e-3)
}

test('each shader matches its algorithm-sheet oracle at 50 %', async ({ page }) => {
  const summary: string[] = []
  // 1. Every cell of every variant at three progresses against the reference models.
  for (const key of KEYS.filter(k => k !== 'ripple')) {
    for (const attrs of VARIANTS[key]) {
      for (const p of [0.25, 0.5, 0.75]) {
        const { grid, cells } = refCells(key, attrs, p, SEED)
        const got = await probe(page, key, attrs, p, cellVertices(grid), SEED, key === 'glitter' ? 1 : 0)
        let worstT = 0, worstOpacity = 0, worstShade = 0, worstPos = 0
        cells.forEach((ref, k) => {
          const c = got[2 * k], corner = got[2 * k + 1]
          worstT = Math.max(worstT, Math.abs(c.t - ref.t))
          worstOpacity = Math.max(worstOpacity, Math.abs(c.opacity - ref.opacity))
          worstShade = Math.max(worstShade, Math.abs(corner.shade - ref.shade))
          const want = apply(ref.model, slidePoint(ref.i / grid[0], ref.j / grid[1], ASPECT))
          for (let a = 0; a < 3; a++) worstPos = Math.max(worstPos, Math.abs(corner.world[a] - want[a]))
        })
        const what = `${label(key, attrs)} @${p}`
        expect(worstT, `${what} local time`).toBeLessThan(1e-4)
        expect(worstOpacity, `${what} opacity`).toBeLessThan(2e-3)
        expect(worstShade, `${what} shade`).toBeLessThan(2e-3)
        expect(worstPos, `${what} corner position`).toBeLessThan(2e-3)
        if (p === 0.5) summary.push(`${what}: t ${worstT.toExponential(1)} opacity ${worstOpacity.toExponential(1)} shade ${worstShade.toExponential(1)} pos ${worstPos.toExponential(1)}`)
      }
    }
  }
  console.log(`[slideshow-3d-particles] oracle worst errors\n${summary.join('\n')}`)

  // 2. The sheet's literal numbers at 50 %, read from the GPU.
  const at = async (key: Key, attrs: Attrs) => {
    const { grid } = refCells(key, attrs, 0.5, SEED)
    const got = await probe(page, key, attrs, 0.5, cellVertices(grid), SEED, key === 'glitter' ? 1 : 0)
    const centres = got.filter((_v, k) => k % 2 === 0)
    const frac = (f: (t: number) => boolean) => Math.round((centres.filter(c => f(c.t)).length / centres.length) * 1e4) / 1e4
    return { grid, centres, rest: frac(t => t === 0), done: frac(t => t === 1), toShown: frac(t => ease(t) > 0.5) }
  }
  const sheet = (key: Key) => ROWS.find(r => r.id === ROW_ID[key])!.oracle(0.5)
  const near = (got: number, want: number | string, what: string, tol = 0.01) => expect(Math.abs(got - Number(want)), `${what}: ${got} vs ${want}`).toBeLessThanOrEqual(tol)

  const vortex = await at('vortex', { dir: 'l' })
  near(vortex.rest, sheet('vortex').rest, 'vortex rest'); near(vortex.done, sheet('vortex').done, 'vortex done')
  near(vortex.rest, 0.059, 'vortex rest (sheet)'); near(vortex.done, 0.059, 'vortex done (sheet)')
  ndcClose(ndcOf(vortex.centres[0]), [-4.2778, 0.8061], 'vortex cell (0,0) centre NDC')
  near(vortex.centres[0].t, 0.8158, 'vortex cell (0,0) t', 1e-3)

  const strip = await at('shred', { pattern: 'strip' })
  near(strip.rest, 0, 'shred strip rest'); near(strip.done, 0, 'shred strip done')
  near(strip.centres[0].t, 0.8404, 'shred strip 0 t', 1e-3)
  near(strip.centres[0].world[1], 0.9181, 'shred strip 0 dy', 1e-3)
  const rect = await at('shred', { pattern: 'rectangle' })
  near(rect.rest, 0, 'shred rectangle rest'); near(rect.done, 0, 'shred rectangle done')

  const fracture = await at('fracture', {})
  near(fracture.rest, 0.0714, 'fracture rest', 0.005); near(fracture.done, 0.0238, 'fracture done', 0.005)
  ndcClose(ndcOf(fracture.centres[0]), [-0.9167, 0.8571], 'fracture cell (0,0) centre NDC')

  const crush = await at('crush', {})
  near(crush.rest, 0, 'crush rest'); near(crush.done, 0, 'crush done')
  ndcClose(ndcOf(crush.centres[0]), [-0.6738, 0.6388], 'crush cell (0,0) centre NDC')
  near(crush.centres[0].opacity, 1, 'crush opacity', 1e-3)

  const glitter = await at('glitter', { dir: 'l' })
  near(glitter.rest, 0.1372, 'glitter rest'); near(glitter.done, 0.1319, 'glitter done')
  near(glitter.toShown, 0.4896, 'glitter cells showing to')
  near(glitter.centres[0].t, 0.8183, 'glitter cell (0,0) t', 1e-3)

  const honey = await at('honeycomb', {})
  near(honey.rest, 0.0455, 'honeycomb rest', 0.005); near(honey.done, 0.0318, 'honeycomb done', 0.005)
  near(honey.centres[5 * 20 + 10].t, 1, 'honeycomb cell (10,5) t', 1e-3)
  near(honey.centres[0].t, 0, 'honeycomb cell (0,0) t', 1e-3)

  // 3. Ripple: z by the vertex's uv (and the outgoing frame's opacity), for every origin on a lattice.
  for (const dir of ['center', 'lu', 'ru', 'ld', 'rd']) {
    const verts: ProbeVertex[] = []
    for (let j = 0; j <= 18; j++) for (let i = 0; i <= 32; i++) verts.push({ uv: [i / 32, j / 18], cell: [i / 32, j / 18] })
    for (const p of [0.25, 0.5, 0.75]) {
      const got = await probe(page, 'ripple', { dir }, p, verts)
      let worst = 0
      verts.forEach((v, k) => { worst = Math.max(worst, Math.abs(got[k].world[2] - rippleZ(p, v.uv[0], v.uv[1], dir)), Math.abs(got[k].opacity - (1 - ease(p)))) })
      expect(worst, `ripple ${dir} @${p}`).toBeLessThan(1e-4)
    }
  }
  const probeUv = async (dir: string, u: number, v: number) => (await probe(page, 'ripple', { dir }, 0.5, [{ uv: [u, v], cell: [u, v] }]))[0]
  near((await probeUv('center', 0.75, 0.5)).world[2], 0.0303, 'ripple z at uv(0.75,0.5)', 1e-4)
  near((await probeUv('center', 0.5, 0.5)).world[2], 0, 'ripple z at the centre', 1e-4)
  near((await probeUv('lu', 0.25, 0.25)).world[2], -0.0069, 'ripple lu z at uv(0.25,0.25)', 1e-4)
  near((await probeUv('center', 0.5, 0.5)).opacity, 0.5, 'ripple from opacity', 1e-4)
})

// ---------------------------------------------------------------------------
// The picture: a tile no other tile covers shows its source at its projected centre
// ---------------------------------------------------------------------------
/** A cell's corners in slide uv. */
const cornerUv = (c: { i: number; j: number }, grid: readonly [number, number]): Array<[number, number]> =>
  [[c.i, c.j], [c.i + 1, c.j], [c.i + 1, c.j + 1], [c.i, c.j + 1]].map(([i, j]) => [i / grid[0], j / grid[1]])
const rgb = (data: Uint8ClampedArray, x: number, y: number): number[] => {
  const xi = Math.min(DW - 1, Math.max(0, Math.floor(x))), yi = Math.min(DH - 1, Math.max(0, Math.floor(y)))
  return [data[(yi * DW + xi) * 4], data[(yi * DW + xi) * 4 + 1], data[(yi * DW + xi) * 4 + 2]]
}

test('at 50 % uncovered tiles show their source at the projected centre', async ({ page }) => {
  const report: string[] = []
  const cases: Array<[Key, Attrs]> = [
    ['vortex', { dir: 'l' }], ['vortex', { dir: 'u' }], ['shred', { pattern: 'strip' }], ['shred', { pattern: 'rectangle' }], ['shred', { pattern: 'rectangle', dir: 'in' }],
    ['fracture', {}], ['crush', {}], ['glitter', { dir: 'l' }], ['glitter', { dir: 'd' }], ['honeycomb', {}],
  ]
  for (const [key, attrs] of cases) {
    // `in` runs time backwards, so its early frames are the late ones of `out`.
    for (const p of attrs.dir === 'in' ? [0.5, 0.65] : [0.5, 0.35]) {
      const frame = await render(page, key, p, attrs)
      const flatFrom = (await render(page, key, 0, attrs)).data, flatTo = (await render(page, key, 1, attrs)).data
      const { element, preset } = ELEMENT[key]
      const plan = particlePlan(element, preset, attrs, p, ASPECT, frame.seed)
      const { grid, cells } = refCells(key, attrs, p, frame.seed)
      const [cols, rows] = grid
      // Each cell's projected outline in device px.
      const quads = cells.map((c) => {
        const mvp = mul4(CAM, c.model)
        return cornerUv(c, grid).map(([u, v]) => dev(project(mvp, slidePoint(u, v, ASPECT))))
      })
      const area = (q: Array<[number, number]>) => Math.abs(q.reduce((s, a, k) => { const b = q[(k + 1) % 4]; return s + (a[0] * b[1] - b[0] * a[1]) }, 0)) / 2
      /** Whether (x, y) lies in the convex quad `q` grown by `pad` px; a quad with no area contains nothing. */
      const inside = (q: Array<[number, number]>, x: number, y: number, pad: number) => {
        const signed = q.reduce((s, a, k) => { const b = q[(k + 1) % 4]; return s + (a[0] * b[1] - b[0] * a[1]) }, 0)
        if (Math.abs(signed) < 0.1) return false
        const orient = Math.sign(signed)
        for (let k = 0; k < 4; k++) {
          const a = q[k], b = q[(k + 1) % 4]
          const cr = (b[0] - a[0]) * (y - a[1]) - (b[1] - a[1]) * (x - a[0])
          if ((orient * cr) / (Math.hypot(b[0] - a[0], b[1] - a[1]) || 1) < -pad) return false
        }
        return true
      }
      // A cell is sampled at its centre; a strip that has left the slide is sampled lower or higher within itself (strips are 8 x 108 px).
      const samples: Array<[number, number]> = key === 'shred' && attrs.pattern !== 'rectangle' ? [[0.5, 0.5], [0.5, 0.92], [0.5, 0.08], [0.5, 0.8], [0.5, 0.2]] : [[0.5, 0.5]]
      // Few tiles are left uncovered by the others (crush, fracture) or still visible on the slide (shred strips at 50 %: all are mostly gone).
      const minChecked = key === 'crush' || key === 'fracture' ? 3 : key === 'shred' && attrs.pattern !== 'rectangle' ? 2 : 8
      let checked = 0, worst = 0
      for (const [k, c] of cells.entries()) {
        // A sliver (a tile nearly edge-on or shrunk away) can miss the pixel; require a fair share of the cell's own area.
        if (area(quads[k]) < 0.3 * ((SW / cols) * (SH / rows))) continue
        if (key === 'glitter' ? Math.abs(ease(c.t) - 0.5) < 0.08 : c.opacity < 0.15) continue
        const mvp = mul4(CAM, c.model)
        for (const [fu, fv] of samples) {
          const u = (c.i + fu) / cols, v = (c.j + fv) / rows
          const [cx, cy] = dev(project(mvp, slidePoint(u, v, ASPECT)))
          if (!(cx > OX + 3 && cx < OX + SW - 3 && cy > OY + 3 && cy < OY + SH - 3)) continue
          if (quads.some((q, m) => m !== k && inside(q, cx, cy, 1.5))) continue
          const rest = dev(project(CAM, slidePoint(u, v, ASPECT)))
          let want: number[]
          if (key === 'glitter') {
            const src = rgb(c.showsTo ? flatTo : flatFrom, rest[0], rest[1])
            want = src.map(x => Math.min(255, x * c.shade))
          } else {
            const tile = rgb(plan.passes[0].texture === 'to' ? flatTo : flatFrom, rest[0], rest[1])
            const back = rgb(plan.flat === 'to' ? flatTo : flatFrom, cx, cy)
            want = tile.map((x, a) => c.opacity * x + (1 - c.opacity) * back[a])
          }
          const got = rgb(frame.data, cx, cy)
          const err = Math.max(...got.map((x, a) => Math.abs(x - want[a])))
          worst = Math.max(worst, err)
          checked++
          expect(err, `${label(key, attrs)} @${p} cell (${c.i},${c.j}) at ${cx.toFixed(1)},${cy.toFixed(1)}: got ${got} want ${want.map(Math.round)}`).toBeLessThanOrEqual(10)
          break
        }
      }
      report.push(`${label(key, attrs)} @${p}: ${checked} tiles, worst ${worst.toFixed(1)}`)
      expect(checked, `${label(key, attrs)} @${p} checked tiles`).toBeGreaterThanOrEqual(minChecked)
    }
  }
  console.log(`[slideshow-3d-particles] pixels\n${report.join('\n')}`)
})

test('ripple at 50 % blends the displaced frame over the incoming one by e(p)', async ({ page }) => {
  const frame = (await render(page, 'ripple', 0.5)).data
  const from = (await render(page, 'ripple', 0)).data, to = (await render(page, 'ripple', 1)).data
  // At the centre z = 0 and the outgoing opacity is 1 - e(1/2) = 1/2; sample a few pixels around it.
  for (const [dx, dy] of [[0, 0], [1, 0], [0, 1], [-1, -1]]) {
    const x = OX + SW / 2 + dx, y = OY + SH / 2 + dy
    const a = rgb(from, x, y), b = rgb(to, x, y), got = rgb(frame, x, y)
    const err = Math.max(...got.map((v, k) => Math.abs(v - (a[k] + b[k]) / 2)))
    expect(err, `centre ${dx},${dy}: got ${got} from ${a} to ${b}`).toBeLessThanOrEqual(10)
  }
  // Away from the ripple origin the displaced frame stays near the flat one: the colours move, they are not black.
  expect(Math.max(...rgb(frame, OX + 4, OY + 4))).toBeGreaterThan(20)
})

// ---------------------------------------------------------------------------
// Frame goldens (Mac only, like the other pixel goldens): 0/25/50/75/100 % through the SwiftShader WebGL2 compositor
// ---------------------------------------------------------------------------
test('frame goldens at 0, 25, 50, 75 and 100 %', async ({ page }, info) => {
  test.skip(process.platform !== 'darwin' || !!process.env.CI, 'Mac-only goldens (SwiftShader on the developer machine)')
  mkdirSync(GOLDENS, { recursive: true })
  const strict = process.env.OFFICEVIEW_STRICT_GOLDEN === '1'
  const recorded: string[] = []
  for (const key of KEYS) {
    const attrs = VARIANTS[key][0]
    for (const pct of [0, 25, 50, 75, 100]) {
      const frame = await render(page, key, pct / 100, attrs)
      const path = join(GOLDENS, `${key}.${pct}.png`)
      if (!existsSync(path)) {
        if (strict) throw new Error(`missing golden ${path}`)
        await savePng(frame.data, path)
        recorded.push(path)
        continue
      }
      const golden = await loadPng(path)
      const d = frameDiff(frame.data, golden, DW, DH)
      expect(withinTolerance(d, { meanAbs: 0.01, rowMax: 0.05 }), `${key} @${pct} vs golden: ${JSON.stringify(d)}`).toBe(true)
    }
  }
  if (recorded.length) info.annotations.push({ type: 'recorded', description: `${recorded.length} goldens` })
})

// ---------------------------------------------------------------------------
// Frame time on a real GPU (not CI): OV_SLIDESHOW_BENCH=1 OV_SLIDESHOW_GPU=1
// ---------------------------------------------------------------------------
test('p95 frame time at 3840 x 2160 on the machine GPU', async () => {
  test.skip(process.env.OV_SLIDESHOW_BENCH !== '1' || process.env.OV_SLIDESHOW_GPU !== '1', 'set OV_SLIDESHOW_BENCH=1 OV_SLIDESHOW_GPU=1 (real GPU, not CI)')
  test.setTimeout(600_000)
  const browser = await chromium.launch({ args: ['--use-angle=metal', '--ignore-gpu-blocklist', '--enable-gpu-rasterization'] })
  try {
    const page = await browser.newPage()
    page.on('pageerror', e => errors.push(e.message))
    await open(page)
    const renderer = await page.evaluate(() => window.__slideshowParticles!.renderer())
    console.log(`[slideshow-3d-particles] renderer ${renderer}`)
    expect(renderer, 'a hardware GPU, not SwiftShader').not.toMatch(/swiftshader|llvmpipe|software/i)
    const rows: string[] = []
    for (const key of KEYS) {
      for (const attrs of (key === 'glitter' ? [{}, { pattern: 'hexagon' }] : key === 'shred' ? [{}, { pattern: 'rectangle' }] : [{}]) as Attrs[]) {
        const times = await page.evaluate(o => window.__slideshowParticles!.bench(o), { key, attrs, devW: 3840, devH: 2160, frames: 120 })
        const sorted = [...times].sort((a, b) => a - b)
        const p95 = sorted[Math.min(sorted.length - 1, Math.ceil(sorted.length * 0.95) - 1)]
        rows.push(`${label(key, attrs)}: p50 ${sorted[Math.floor(sorted.length / 2)].toFixed(2)} ms, p95 ${p95.toFixed(2)} ms, max ${sorted[sorted.length - 1].toFixed(2)} ms`)
        expect(p95, `${label(key, attrs)} p95`).toBeLessThanOrEqual(12)
      }
    }
    console.log(`[slideshow-3d-particles] bench (${renderer})\n${rows.join('\n')}`)
  } finally { await browser.close() }
})
