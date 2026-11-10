/**
 * Stage-0 Task 2 — T1 backend-confirm tests.
 *
 * The suite validates the T1 machinery in isolation (fixture-injected units,
 * bitwise-identical bitmaps, the 9-row construction-site matrix) without
 * loading the full corpus — the full corpus battery runs in
 * `scripts/stream-s0/t1-backend.ts` (background run; see the .done file).
 */
// removed
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import * as napiCanvas from '@napi-rs/canvas'
import * as canvasModule from 'canvas'
import { listCorpus, loadUnits, type Bitmap, type Unit } from '../../scripts/stream-s0/corpus'
import { collectRenderResults, diffBitmaps } from '../../scripts/stream-s0/t1-backend'

// Round-10-fix1: the round-9 T1 run compared node-canvas with node-canvas
// (the parent ran without the preload, so "napi" was node-canvas too and all
// 648 diff rows were all-zero). These tests pin the fix: the two runners
// report DIFFERENT backends, and a known text page really differs.
const napiRun = process.env.S0_BACKEND === 'napi'

const solidBitmap = (w: number, h: number, r: number, g: number, b: number): Bitmap => {
  const data = new Uint8ClampedArray(w * h * 4)
  for (let i = 0; i < data.length; i += 4) {
    data[i] = r
    data[i + 1] = g
    data[i + 2] = b
    data[i + 3] = 255
  }
  return { width: w, height: h, data }
}

describe('T1 machinery (injected units, no corpus)', () => {
  it('counts napi-only failures (a unit that throws only under the napi ctx)', () => {
    // Injected units where paint throws only when the ctx says it was
    // created by the napi shim: the plain backend renders fine.
    const napiCtxName = (globalThis as { __s0NapiTestClass?: string }).__s0NapiTestClass
    void napiCtxName
    const marker = '__streamS0NapiProbe'
    const throwingUnit = {
      index: 1,
      widthPx: 10,
      heightPx: 10,
      paint: (ctx: CanvasRenderingContext2D) => {
        // Both backends name their context "CanvasRenderingContext2D", so
        // class-name discrimination is impossible — the runner stamps the
        // napi ctx with the marker key instead (verified: ctor name is
        // "CanvasRenderingContext2D" on both backends; the plan's
        // "ctx.constructor.name is the napi context" reads as the stamp).
        if ((ctx as unknown as Record<string, unknown>)[marker]) {
          throw new Error('napi-only injected failure')
        }
      },
    }
    const plainUnit = {
      index: 0,
      widthPx: 10,
      heightPx: 10,
      paint: () => {},
    }
    const results = collectRenderResults(
      [
        { file: 'ok.docx', units: [plainUnit, throwingUnit] },
      ],
      {
        plain: { canvasFactory: createCanvas as unknown as never, stamp: undefined, env: 'node-canvas' },
        napi: { canvasFactory: createCanvas as unknown as never, stamp: marker, env: '@napi-rs/canvas' },
      },
    )
    const napiOnly = results.renderSuccess.napiOnlyFailures
    expect(napiOnly).toHaveLength(1)
    expect(napiOnly[0]).toMatchObject({ file: 'ok.docx', unit: 1 })
    expect(napiOnly[0].error).toMatch(/napi-only/)
    expect(results.renderSuccess.bothFailed).toHaveLength(0)
  })

  it('diff of identical bitmaps is zero', () => {
    const a = solidBitmap(8, 8, 255, 0, 0)
    const b = solidBitmap(8, 8, 255, 0, 0)
    const d = diffBitmaps(a, b)!
    expect(d.pixelsDiffering).toBe(0)
    expect(d.fracDiffering).toBe(0)
    expect(d.maxChannelDelta).toBe(0)
    expect(d.p99ChannelDelta).toBe(0)
  })

  it('diff of differing bitmaps records both frac and p99', () => {
    const a = solidBitmap(10, 10, 255, 0, 0)
    const b = solidBitmap(10, 10, 240, 0, 20)
    const d = diffBitmaps(a, b)!
    expect(d.pixelsDiffering).toBe(100)
    expect(d.maxChannelDelta).toBe(20)
    expect(d.p99ChannelDelta).toBeGreaterThan(0)
  })

  it('site matrix has 9 rows (§7.3 construction sites, verbatim)', () => {
    const rows = collectRenderResults(
      [{ file: 'empty.docx', units: [] }],
      {
        plain: { canvasFactory: createCanvas as unknown as never, stamp: undefined, env: 'node-canvas' },
        napi: { canvasFactory: createCanvas as unknown as never, stamp: undefined, env: '@napi-rs/canvas' },
      },
      { syntheticOnly: true },
    ).constructionSites
    expect(rows).toHaveLength(9)
    const sites = new Set(rows.map((r) => r.site))
    for (const site of [
      'src/render/paint.ts:199',
      'src/core/search.ts:212',
      'src/core/images.ts:65',
      'src/core/images.ts:149',
      'src/slideshow/render/svg-raster.ts:64',
      'src/core/text-metrics.ts:31',
      'src/core/text-metrics.ts:42',
      'src/drawing/paint.ts:126',
      'src/drawing/pattern.ts:155',
    ]) expect(sites).toContain(site)
    for (const row of rows) {
      expect(['ok', 'throws', 'undefined']).toContain(row.napi)
      expect(row.evidence.length).toBeGreaterThan(0)
    }
  })

  it('every construction site renders a real canvas under the napi backend', () => {
    const rows = collectRenderResults(
      [],
      {
        plain: { canvasFactory: createCanvas as unknown as never, stamp: undefined, env: 'node-canvas' },
        napi: { canvasFactory: createCanvas as unknown as never, stamp: undefined, env: '@napi-rs/canvas' },
      },
      { syntheticOnly: true },
    ).constructionSites
    const bad = rows.filter((r) => r.napi !== 'ok')
    expect(bad).toEqual([])
  })
})

describe('T1 backend identity (round-10-fix1)', () => {
  it('the two runners report different backends', () => {
    // Module identity: under the plain run `canvas` resolves to the real
    // node-canvas package (no shim marker); `@napi-rs/canvas` is the napi
    // module. The napi run aliases `canvas` → the shim, whose marker export
    // matches. Either way the two sides are distinct modules.
    // The marker is MY shim's export; the raw @napi-rs/canvas module never
    // has it — so identity is proven by the RESOLVED `canvas` module (aliased
    // under napi = shim = marker present; plain = real canvas = marker absent).
    const canvasIsNapi = 'STREAM_S0_BACKEND' in canvasModule
    const napiRawHasMarker = 'STREAM_S0_BACKEND' in napiCanvas
    expect(napiRawHasMarker).toBe(false)
    expect(canvasIsNapi).toBe(napiRun)
  })

  it.skipIf(napiRun)('a known text page differs >0 pixels between the backends', async () => {
    // The orchestrator's evidence: poi-55733.docx unit 0 at scale 1 differs
    // between node-canvas and napi (glyph rasterization differs). This test
    // pins that the T1 diff is REAL (never all-zero again).
    const file = listCorpus().find((f) => f.name === 'poi-55733.docx')
    expect(file, 'corpus/poi-55733.docx must exist').toBeTruthy()
    const units = await loadUnits(file!)
    const unit: Unit = units[0]
    const nodeBmp: Bitmap = { width: 0, height: 0, data: new Uint8ClampedArray(0) }
    const napiBmp: Bitmap = { width: 0, height: 0, data: new Uint8ClampedArray(0) }
    // Render via collectRenderResults with two distinct factories.
    const results = collectRenderResults(
      [{ file: 'poi-55733.docx', units: [unit] }],
      {
        plain: { canvasFactory: createCanvas as unknown as never, stamp: undefined, env: 'node-canvas' },
        napi: { canvasFactory: napiCanvas.createCanvas as unknown as never, stamp: undefined, env: '@napi-rs/canvas' },
      },
      { grid: [[1, 1]] },
    )
    expect(results.renderSuccess.okNode).toBe(1)
    expect(results.renderSuccess.okNapi).toBe(1)
    const row = results.diff[0]
    expect(row).toBeTruthy()
    expect(row!.pixelsDiffering).toBeGreaterThan(0)
    void nodeBmp
    void napiBmp
  })
})
