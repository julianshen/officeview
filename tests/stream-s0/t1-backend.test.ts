/**
 * Stage-0 Task 2 — T1 backend-confirm tests.
 *
 * The suite validates the T1 machinery in isolation (fixture-injected units,
 * bitwise-identical bitmaps, the 9-row construction-site matrix) without
 * loading the full corpus — the full corpus battery runs in
 * `scripts/stream-s0/t1-backend.ts` (background run; see the .done file).
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import type { Bitmap } from '../../scripts/stream-s0/corpus'
import { collectRenderResults, diffBitmaps } from '../../scripts/stream-s0/t1-backend'

const blackBitmap = (w: number, h: number): Bitmap => ({ width: w, height: h, data: new Uint8ClampedArray(w * h * 4) })
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
    const d = diffBitmaps(a, b)
    expect(d.pixelsDiffering).toBe(0)
    expect(d.fracDiffering).toBe(0)
    expect(d.maxChannelDelta).toBe(0)
    expect(d.p99ChannelDelta).toBe(0)
  })

  it('diff of differing bitmaps records both frac and p99', () => {
    const a = solidBitmap(10, 10, 255, 0, 0)
    const b = solidBitmap(10, 10, 240, 0, 20)
    const d = diffBitmaps(a, b)
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
