/**
 * Stage 4a Task 15 — T2 on tiles, coverage pixels and marks over the wire (S4 §6.2, §6.5; parent §3, T2). Every tile
 * here comes over the whole server path through the harness `Client` (a tiles session: `hello` with `sheet-tiles@1`,
 * `ready`, `sheet-layout`, then `tiles`), at (zoomStep 5, dpr 1) and (zoomStep 5, dpr 2), forensic marking off.
 *
 * - **T2**: for each XLSX hidden fixture, every tile of `withHidden`'s used extent is PNG-byte-equal to the same tile
 *   of `withoutHidden`.
 * - **Coverage pixels**: a base-covered cell paints empty under visible-only (and as authored under 'full', the
 *   only 'full' use); a rotated box blanks the cells it covers; unknown coverage paints as authored (a local-path
 *   `renderSheetViewport` of the projected sheet); tiles paint saved values. Each "equals" has a same-policy
 *   "differs" control, so a comparison of a document with itself (or a cache that conflates documents) cannot pass.
 * - **Marks over the wire**: with the fixed visible mark, tiles (0,0) and (1,0) composited equal one
 *   `paintMarkRegion` over the 1024 × 512 region of the unmarked tiles (Task 7's recipe, at zoomStep 5, dpr 2).
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { OfficePackage } from '../../src/core/zip'
import { prepareSheetAssets } from '../../src/components/xlsx/sheetAssets'
import { decodeSheetLayout, TILE_PX, TILE_ZOOM_STEPS } from '../../src/stream/sheet-wire'
import type { LogicalViewport } from '../../src/xlsx/view/contracts'
import { renderSheetViewport } from '../../src/xlsx/view/renderViewport'
import { streamSheetGeometry } from '../../src/xlsx/view/streamGeometry'
import { LayoutCache } from '../../src/xlsx/view/textLayout'
import { tileLogicalBox } from '../../src/xlsx/view/tileCache'
import { buildSheetGeometry, frameBudgetFor, planViewport } from '../../src/xlsx/view/viewport'
import { parseXlsx } from '../../src/xlsx/parse'
import type { AuthDecision } from '../src/hooks'
import { decodePng } from '../src/mark/png'
import { markOptions, paintMarkRegion } from '../src/mark/visible'
import { projectPackage } from '../src/policy'
import { DocumentHost } from '../src/render'
import { clockEpochMs, type SheetClock } from '../src/sheet/clock'
import { TILE_PALETTE_COLORS, tileLimit } from '../src/sheet/tile'
import { augmentedFixtures } from './fixtures'
import { Client, disposePool, start } from './harness'
import { blankSheet, patchSheetXml, patchXlsx, plainSeven, rotatedBox, savedFormula, unfilledShapeOverText, unresolvedAnchor } from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

type Policy = AuthDecision['policy']
const TILES = ['sheet-tiles@1']
const VISIBLE: Policy = { contentPolicy: 'visible-only', visibleMark: 'off', forensicMark: false }
const PAIRS = [[5, 1], [5, 2]] as const
const T = TILE_PX

/** One sheet's tile coordinates at one pair. */
interface PlanEntry { sheetId: string; zoomStep: number; dpr: 1 | 2; coords: Array<{ col: number; row: number }> }
const tileKey = (e: { sheetId: string; zoomStep: number; dpr: number }, t: { col: number; row: number }) => `${e.sheetId}|${e.zoomStep}x${e.dpr}|${t.col},${t.row}`

/**
 * A tiles session over `doc` under `policy`: unless `plan` is given, the plan is every tile of each visible sheet's
 * used extent (from its wire `sheet-layout`) at both pairs. Returns the plan, every tile's PNG by key, and the session
 * clock (for a local reference).
 */
async function wireTiles(doc: Uint8Array, policy: Policy, plan?: PlanEntry[]) {
  const h = await start({ doc, policy })
  const c = new Client(h.url)
  try {
    const m = await c.hello({ features: TILES })
    await c.ready()
    expect(m.sheetMode).toBe('tiles')
    let reqId = 10
    if (!plan) {
      plan = []
      for (const { sheetId } of m.sheets!) {
        const lr = await c.sheetLayout(++reqId, sheetId)
        expect(lr.end).toBeUndefined()
        const layout = decodeSheetLayout(lr.parts.map((p) => p.text))
        const mt = streamSheetGeometry(layout).metrics
        const right = layout.used ? mt.cols.start(layout.used.c1) + mt.cols.size(layout.used.c1) : 0
        const bottom = layout.used ? mt.rows.start(layout.used.r1) + mt.rows.size(layout.used.r1) : 0
        for (const [zoomStep, dpr] of PAIRS) {
          const zoom = TILE_ZOOM_STEPS[zoomStep]!, cols = tileLimit(right, zoom, dpr), rows = tileLimit(bottom, zoom, dpr)
          plan.push({ sheetId, zoomStep, dpr, coords: Array.from({ length: cols * rows }, (_, i) => ({ col: i % cols, row: Math.floor(i / cols) })) })
        }
      }
    }
    const pngs = new Map<string, Uint8Array>()
    for (const e of plan) {
      for (let i = 0; i < e.coords.length; i += 64) {
        const r = await c.tiles(++reqId, { sheetId: e.sheetId, viewGen: 0, zoomStep: e.zoomStep, dpr: e.dpr, palette: 'light', tiles: e.coords.slice(i, i + 64) }, { ms: 60_000 })
        expect(r.end).toEqual({ reqId, type: 'tiles-done', refused: [] })
        for (const t of r.tiles) pngs.set(tileKey(e, t), t.png)
      }
    }
    expect(pngs.size).toBe(plan.reduce((n, e) => n + e.coords.length, 0))
    return { plan, pngs, clock: h.stream.sheetClockOf(m.session.sid)! }
  } finally { c.close(); await h.stop() }
}

const sameBytes = (a: Uint8Array, b: Uint8Array) => Buffer.compare(Buffer.from(a), Buffer.from(b)) === 0
/** The keys whose PNGs are not byte-equal between `a` and `b` (every key of `a`). */
const unequal = (a: Map<string, Uint8Array>, b: Map<string, Uint8Array>) => [...a.keys()].filter((k) => !b.has(k) || !sameBytes(a.get(k)!, b.get(k)!))
/** Differing pixels of two RGBA surfaces. */
function diffCount(a: Uint8ClampedArray, b: Uint8ClampedArray): number {
  expect(a.length).toBe(b.length)
  let n = 0
  for (let i = 0; i < a.length; i += 4) if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2] || a[i + 3] !== b[i + 3]) n++
  return n
}

/** `bytes` with the cells `refs` of sheet 0 deleted (throws if one is absent: a twin must differ from its fixture). */
function withoutCells(bytes: Uint8Array, refs: string[]): Promise<Uint8Array> {
  return patchXlsx(bytes, (zip) => patchSheetXml(zip, 0, (xml) => {
    for (const ref of refs) {
      const next = xml.replace(new RegExp(`<c r="${ref}"(?:[^>]*/>|[^>]*>.*?</c>)`), '')
      if (next === xml) throw new Error(`t2: no cell ${ref}`)
      xml = next
    }
    return xml
  }))
}

/**
 * The local path's tile of sheet 0 (no servable snapshot: the projected package parsed with saved values, then
 * `renderTile`'s own paint calls, line for line), as RGBA.
 */
async function localTile(bytes: Uint8Array, clock: SheetClock, zoomStep: number, dpr: 1 | 2, col: number, row: number): Promise<Uint8ClampedArray> {
  const host = new DocumentHost(); hosts.push(host)
  const projected = await projectPackage(await OfficePackage.load(bytes), 'xlsx', 'visible-only')
  const doc = await parseXlsx(projected.pkg, { valuePolicy: 'cached' })
  const sheet = doc.sheets[0]!
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('local geometry is a diagnostic')
  const lease = await prepareSheetAssets(doc, { sheetIndex: 0, registerFont: () => { throw new Error('no embedded fonts') }, cfClock: { epochNowMs: clockEpochMs(clock), timeZone: clock.timeZone } })
  try {
    const zoom = TILE_ZOOM_STEPS[zoomStep]!, s = zoom * dpr
    const canvas = host.backend.createCanvas(T, T)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    const viewport: LogicalViewport = { box: tileLogicalBox(col, row, zoom, dpr), zoom, deviceScale: dpr }
    const frameBudget = frameBudgetFor('D')
    const plan = planViewport(geometry, viewport, 'D', { frameBudget })
    ctx.setTransform(s, 0, 0, s, 0, 0)
    renderSheetViewport(sheet, ctx, geometry, viewport, lease.assets, { colors: TILE_PALETTE_COLORS.light, preparedPlan: plan, profile: 'D', cache: new LayoutCache(), frameBudget })
    return decodePng(new Uint8Array((canvas as unknown as { toBuffer(m: 'image/png'): Buffer }).toBuffer('image/png'))).rgba
  } finally { lease.dispose() }
}

describe('T2 on tiles', () => {
  test('every withHidden tile equals the withoutHidden tile (each XLSX hidden fixture)', async () => {
    const xlsx = (await augmentedFixtures()).filter((f) => f.kind === 'xlsx')
    expect(xlsx.length).toBe(5)
    for (const f of xlsx) {
      const withH = await wireTiles(f.withHidden, VISIBLE)
      expect([f.id, withH.plan.map((e) => `${e.zoomStep}x${e.dpr}`)]).toEqual([f.id, ['5x1', '5x2']])
      const withoutH = await wireTiles(f.withoutHidden, VISIBLE, withH.plan)
      expect([f.id, unequal(withH.pngs, withoutH.pngs)]).toEqual([f.id, []])
    }
  }, 300_000)
})

describe('coverage pixels', () => {
  test('base-covered paints empty under visible-only', async () => {
    // "The tile of the fixture with B2 deleted" is that fixture painted as authored: the local path (no snapshot). Its
    // wire tile is not the reference: with B2 gone, A2's conservative spill rect (R8: general text spills right into
    // the contiguous run of empty cells) reaches the box inside B2, so the twin's own snapshot also blanks A2.
    const bytes = await unfilledShapeOverText(), deleted = await withoutCells(bytes, ['B2'])
    const a = await wireTiles(bytes, VISIBLE)
    const authored = new Map<string, Uint8ClampedArray>()
    for (const e of a.plan) for (const t of e.coords) authored.set(tileKey(e, t), await localTile(deleted, a.clock, e.zoomStep, e.dpr, t.col, t.row))
    expect(authored.size).toBeGreaterThan(0)
    for (const [k, ref] of authored) expect([k, diffCount(decodePng(a.pngs.get(k)!).rgba, ref)]).toEqual([k, 0])
    // …and as authored under full: the covered text paints, so every tile differs from the B2-deleted one.
    const af = await wireTiles(bytes, { ...VISIBLE, contentPolicy: 'full' }, a.plan)
    for (const [k, ref] of authored) expect([k, diffCount(decodePng(af.pngs.get(k)!).rgba, ref) > 0]).toEqual([k, true])
  }, 120_000)

  test('rotated box', async () => {
    const bytes = await rotatedBox()
    const a = await wireTiles(bytes, VISIBLE)
    const covered = await wireTiles(await withoutCells(bytes, ['A1', 'A2', 'A3']), VISIBLE, a.plan)
    expect(unequal(a.pngs, covered.pngs)).toEqual([])
    // Control: deleting an uncovered cell does change the tile.
    const uncovered = await wireTiles(await withoutCells(bytes, ['B1']), VISIBLE, a.plan)
    expect(unequal(a.pngs, uncovered.pngs).length).toBeGreaterThan(0)
  }, 120_000)

  test('unknown coverage paints as authored', async () => {
    const bytes = await unresolvedAnchor()
    const a = await wireTiles(bytes, VISIBLE)
    for (const e of a.plan) for (const t of e.coords) {
      const wire = decodePng(a.pngs.get(tileKey(e, t))!).rgba
      const local = await localTile(bytes, a.clock, e.zoomStep, e.dpr, t.col, t.row)
      expect([tileKey(e, t), diffCount(wire, local)]).toEqual([tileKey(e, t), 0])
    }
    // Control: the tile is not a blank sheet's.
    const blank = await wireTiles(await blankSheet(), VISIBLE, a.plan)
    expect(unequal(a.pngs, blank.pngs).length).toBeGreaterThan(0)
  }, 120_000)

  test('saved values', async () => {
    const a = await wireTiles(await savedFormula(), VISIBLE)
    const seven = await wireTiles(await plainSeven(), VISIBLE, a.plan)
    expect(unequal(a.pngs, seven.pngs)).toEqual([])
    // Control: the recomputed value (2) would paint differently.
    const two = await patchXlsx(await plainSeven(), (zip) => patchSheetXml(zip, 0, (xml) => {
      const next = xml.replace('<v>7</v>', '<v>2</v>')
      if (next === xml) throw new Error('t2: no <v>7</v>')
      return next
    }))
    const twos = await wireTiles(two, VISIBLE, a.plan)
    expect(unequal(a.pngs, twos.pngs).length).toBeGreaterThan(0)
  }, 120_000)
})

describe('marks over the wire', () => {
  test('wire tiles keep the mark continuous', async () => {
    // Task 7's recipe (F9) at zoomStep 5, dpr 2: the reference is the unmarked wire tiles composited, then one
    // `paintMarkRegion` over the whole 1024 × 512 region.
    const plan: PlanEntry[] = [{ sheetId: '0', zoomStep: 5, dpr: 2, coords: [{ col: 0, row: 0 }, { col: 1, row: 0 }] }]
    const doc = await blankSheet()
    const marked = await wireTiles(doc, { ...VISIBLE, visibleMark: 'fixed' }, plan)
    const unmarked = await wireTiles(doc, VISIBLE, plan)
    const W = 2 * T, H = T
    const composite = (pngs: Map<string, Uint8Array>) => {
      const out = new Uint8ClampedArray(W * H * 4)
      plan[0]!.coords.forEach((t, i) => {
        const d = decodePng(pngs.get(tileKey(plan[0]!, t))!)
        expect([d.w, d.h]).toEqual([T, T])
        for (let y = 0; y < T; y++) out.set(d.rgba.subarray(y * T * 4, (y + 1) * T * 4), (y * W + i * T) * 4)
      })
      return out
    }
    const base = composite(unmarked.pngs), got = composite(marked.pngs)
    const host = new DocumentHost(); hosts.push(host)
    const canvas = host.backend.createCanvas(W, H)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    const img = ctx.createImageData(W, H)
    img.data.set(base)
    ctx.putImageData(img, 0, 0)
    paintMarkRegion(ctx, markOptions({ kind: 'fixed' })!, { x: 0, y: 0, w: W, h: H, scale: TILE_ZOOM_STEPS[5]! * 2 })
    const reference = ctx.getImageData(0, 0, W, H).data
    expect(diffCount(base, reference)).toBeGreaterThan(1000)
    expect(diffCount(got, reference)).toBe(0)
  }, 120_000)
})
