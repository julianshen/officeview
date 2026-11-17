/**
 * Ruling S3-A11Y-PAGED (docs/validation/xlsx-cf.md §11): the forced-paged view's grid semantics read a wire scope that
 * never paints. Given the frame's own host source (each cell's logical box from the sheet metrics and its screen base
 * style), that scope decides whether a bar paints with the painter's own predicate (`barWithheld` at the capture
 * scale), so it announces exactly what the resident and worker legs announce: the bar descriptor where the bar paints,
 * "unavailable" where painting withholds it, and never a bar the frame would not paint.
 */
import { describe, expect, it } from 'vitest'
import { prepareCfWireScope, type CfFrameHostSource } from '../src/xlsx/cf/text'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import type { CfStamp } from '../src/xlsx/cf/model'
import type { CfReadResult } from '../src/xlsx/view/contracts'
import type { XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'
import { barSheet, withBars } from './helpers/cf-bars'

const UNAVAILABLE = { key: 'cf-unavailable', text: 'Conditional formatting unavailable' }

const wireOf = async (sheet: XlsxSheet, rows: number): Promise<CfReadResult> => {
  const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId!
  const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
  const memory = createCfMemory({})
  const wire = await prepareCfRemote({ rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, { budget: () => createCfWorkBudget({ now: () => 0 }) })
  expect(wire.kind).toBe('ready')
  return wire as CfReadResult
}
const scopeOver = (wire: CfReadResult, rows: number, locale: string, hosts?: CfFrameHostSource) =>
  prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }],
    { budget: { reserve() {}, release() {} }, locale, ...(hosts ? { hosts } : {}) })
/** The frame's host source: a `width` × 20 px box per row, the given screen base style. */
const frameHosts = (width = 200, base?: XlsxCellStyle): CfFrameHostSource => ({
  cell: (row, col) => ({ box: { x: col * width, y: row * 20, width, height: 20 }, base }),
})

describe('the forced-paged semantic scope decides paintability from the frame host (S3-A11Y-PAGED)', () => {
  it('announces a painted bar with its value position, without any paint having run (en and zh-TW)', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([0, 72, 100]), 3)
      const en = (await scopeOver(wire, 3, 'en-US', frameHosts()))!
      expect([0, 1, 2].map(r => en.describe(r, 0))).toEqual([
        { key: 'cf:dataBar', text: 'Data bar, 0% of range' },
        { key: 'cf:dataBar', text: 'Data bar, 72% of range' },
        { key: 'cf:dataBar', text: 'Data bar, 100% of range' },
      ])
      expect(en.barAt?.(1, 0)?.valueFraction).toBeCloseTo(0.72, 10)
      en.dispose()
      const zh = (await scopeOver(wire, 3, 'zh-TW', frameHosts()))!
      expect(zh.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: '資料橫條 72%' })
      zh.dispose()
    })
  })

  it('a host too small to lay the bar out at the capture scale withholds it: unavailable', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([0, 72, 100]), 3)
      // 3 px: the 2 pt + 1 pt insets leave no span at the capture scale (commit b85ddb88).
      const scope = (await scopeOver(wire, 3, 'en-US', frameHosts(3)))!
      for (const row of [0, 1, 2]) {
        expect(scope.describe(row, 0), `row ${row}`).toEqual(UNAVAILABLE)
        expect(scope.barAt?.(row, 0), `row ${row}`).toBeUndefined()
      }
      scope.dispose()
    })
  })

  it('a cell the frame has no exact host for (unknown, hidden, outside the metrics) withholds: unavailable', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([0, 72, 100]), 3)
      const scope = (await scopeOver(wire, 3, 'en-US', { cell: (row, col) => row === 1 ? undefined : frameHosts().cell(row, col) }))!
      expect(scope.describe(1, 0)).toEqual(UNAVAILABLE)
      expect(scope.describe(2, 0)).toEqual({ key: 'cf:dataBar', text: 'Data bar, 100% of range' })
      scope.dispose()
    })
  })

  it('a non-white axis over the host fill the frame paints withholds: unavailable', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([-50, 50, 100]), 3)
      const filled = (await scopeOver(wire, 3, 'en-US', frameHosts(200, { numFmtId: 0, fillColor: 'FFFFFF00' })))!
      expect(filled.describe(1, 0)).toEqual(UNAVAILABLE)
      filled.dispose()
      const plain = (await scopeOver(wire, 3, 'en-US', frameHosts()))!
      expect(plain.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: 'Data bar, 67% of range' })
      plain.dispose()
    })
  })

  it('unverified paint capabilities withhold the bar on the paged path: unavailable', async () => {
    for (const options of [{ paint: false }, { omit: ['bar:geometry-narrow-clip'] }] as const) {
      await withBars(async () => {
        const wire = await wireOf(barSheet([0, 72, 100]), 3)
        const scope = (await scopeOver(wire, 3, 'en-US', frameHosts()))!
        expect(scope.describe(1, 0), JSON.stringify(options)).toEqual(UNAVAILABLE)
        expect(scope.barAt?.(1, 0)).toBeUndefined()
        scope.dispose()
      }, options)
    }
  })

  it('agrees with what the painter does: each hosted descriptor equals a host-less scope\'s after painting the same hosts', async () => {
    const { createCanvas } = await import('canvas')
    await withBars(async () => {
      const cases: ReadonlyArray<readonly [string, XlsxSheet, number, XlsxCellStyle | undefined]> = [
        ['plain', barSheet([0, 72, 100]), 200, undefined],
        ['too small', barSheet([0, 72, 100]), 3, undefined],
        ['sub-pixel share', barSheet([0, 0.001, 1000]), 200, undefined],
        ['filled axis', barSheet([-50, 50, 100]), 200, { numFmtId: 0, fillColor: 'FFFFFF00' }],
        ['axis control', barSheet([-50, 50, 100]), 200, undefined],
      ]
      for (const [name, sheet, width, base] of cases) {
        const wire = await wireOf(sheet, 3)
        const hosts = frameHosts(width, base)
        const hosted = (await scopeOver(wire, 3, 'en-US', hosts))!
        const painted = (await scopeOver(wire, 3, 'en-US'))!
        const ctx = createCanvas(width, 60).getContext('2d') as unknown as CanvasRenderingContext2D
        painted.paintBars?.(ctx, (row, col) => {
          const at = hosts.cell(row, col)!
          return { ...at.box, ...(at.base?.fillColor ? { fill: '#' + at.base.fillColor.slice(2) } : {}) }
        }, { r0: 0, c0: 0, r1: 2, c1: 0 }, { zoom: 1, dpr: 1 })
        for (const row of [0, 1, 2]) {
          expect(hosted.describe(row, 0), `${name} row ${row}`).toEqual(painted.describe(row, 0))
          expect(hosted.barAt?.(row, 0), `${name} row ${row}`).toEqual(painted.barAt?.(row, 0))
        }
        hosted.dispose(); painted.dispose()
      }
    })
  })
})

/**
 * Fix round 1 (review Important 3): the host fill the frame hands the bar painter. `renderSheetViewport` styles every
 * placement through the CF paint layer (`paintLayer.styleAt` = the scope's `effectiveStyle` over the region/own base),
 * and `record.body.fill` is that CF-effective fill. So a CF fill under a bar is part of the painter's host, and the
 * semantic host must resolve its fill the same way (`frameBarHosts` over `effectiveStyle`), not from the base alone.
 */
describe('the semantic host fill is the frame painter\'s host fill (fix round 1)', () => {
  it('a CF fill under an automatic-axis bar withholds it in the frame, and the semantic scope agrees', async () => {
    const { createCanvas } = await import('canvas')
    const { renderSheetViewport } = await import('../src/xlsx/view/renderViewport')
    const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
    const { LayoutCache } = await import('../src/xlsx/view/textLayout')
    const { frameBarHosts } = await import('../src/xlsx/cf/render')
    await withBars(async () => {
      // [-50, 50, 100] under the automatic black axis; a second rule fills every positive cell yellow.
      const bars = barSheet([-50, 50, 100])
      const sheet: XlsxSheet = { ...bars, conditionalFormatting: { ...bars.conditionalFormatting!, rules: [...bars.conditionalFormatting!.rules, {
        id: 2, priority: 2, documentOrder: 1, ranges: [{ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }], stopIfTrue: false,
        writes: new Set(['fill']), dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FFFFFF00' },
        definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
      }] } }
      const wire = await wireOf(sheet, 3)
      const paint = (await scopeOver(wire, 3, 'en-US'))!
      const geometry = buildSheetGeometry(sheet)
      expect(geometry.kind).toBe('ready')
      const metrics = (geometry as Extract<typeof geometry, { kind: 'ready' }>).metrics
      // What the frame hands the painter, per cell.
      const painted = new Map<number, { x: number; y: number; width: number; height: number; fill?: string }>()
      const batch = {
        stamp: wire.stamp, key: 'paged', layer: paint.paintLayer!,
        barAt: (r: number, c: number) => paint.barAt?.(r, c),
        paintBars: (ctx: CanvasRenderingContext2D, cellRect: (r: number, c: number) => { x: number; y: number; width: number; height: number; fill?: string } | undefined, visible: { r0: number; c0: number; r1: number; c1: number }, scale: { zoom: number; dpr: number; output?: 'logical' | 'device' }) => {
          for (let r = visible.r0; r <= visible.r1; r++) { const host = cellRect(r, 0); if (host) painted.set(r, host) }
          return paint.paintBars?.(ctx, cellRect, visible, scale) ?? true
        },
        hidesTextAt: (r: number, c: number, host: { x: number; y: number; width: number; height: number; fill?: string }, scale: { zoom: number; dpr: number; output?: 'logical' | 'device' }) => paint.hidesTextAt?.(r, c, host, scale) ?? false,
        iconAt: (r: number, c: number) => paint.iconAt?.(r, c),
        paintIcons: () => true,
        textInsetAt: () => undefined,
        dispose: () => {},
      }
      const ctx = createCanvas(400, 200).getContext('2d') as unknown as CanvasRenderingContext2D
      renderSheetViewport(sheet, ctx, geometry, { box: { x: 0, y: 0, width: 400, height: 200 }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch } as never, { cache: new LayoutCache() })
      expect(painted.get(1)?.fill, 'the frame host carries the CF fill').toBe('#FFFF00')
      expect(painted.get(0)?.fill, 'an unfilled cell has no host fill').toBeUndefined()
      // The semantic scope over the frame host source the viewer builds (metrics box, base style without CF).
      const source: CfFrameHostSource = { cell: (row, col) => { const box = metrics.sourceCellToLogical({ row, col }); return box ? { box, base: undefined } : undefined } }
      const semantic = (await scopeOver(wire, 3, 'en-US', source))!
      const hostAt = frameBarHosts(source, semantic.effectiveStyle)
      for (const row of [0, 1, 2]) {
        // The bar's host fields (Task 29 widened the frame host with the icon's text width and cell; bars ignore them).
        const barHost = (h: { x: number; y: number; width: number; height: number; fill?: string } | undefined) => h && { x: h.x, y: h.y, width: h.width, height: h.height, ...(h.fill ? { fill: h.fill } : {}) }
        expect(barHost(hostAt(row, 0)), `row ${row}: the semantic host is the frame's host`).toEqual(barHost(painted.get(row)))
        expect(semantic.describe(row, 0), `row ${row}: announces what the frame painted`).toEqual(paint.describe(row, 0))
      }
      expect(semantic.describe(1, 0)).toEqual(UNAVAILABLE)
      expect(semantic.describe(0, 0)?.key).toBe('cf:dataBar')
      // A base-only fill (no CF) would differ from the frame's host here, and so announce the bar the frame withholds.
      expect(frameBarHosts(source, (_r, _c, base) => base)(1, 0)?.fill).toBeUndefined()
      semantic.dispose(); paint.dispose()
    })
  })
})

/**
 * Fix round 1 (review Important 1): which legs reach the hosted semantic scope. `InteractiveSheetDoc` passes hosts only
 * on its `remote` branch (a paged workbook source); the `?worker=1` leg without `paged=1` routes a workbook to the
 * paged worker only when admission says `paged` (`largeSource.loadSeekable`), and admits the Stage 3 browser fixture as
 * an object workbook: a resident sheet, `remote` null, the resident text scope, unchanged. A workbook admission does
 * route paged is the same paged view the forced-paged leg exercises, which needs (and gets) the fix.
 */
describe('the worker leg keeps the resident scope for the Stage 3 fixture (fix round 1)', () => {
  it('admission keeps stage3-bar-lengths.xlsx on the object route under both profiles', async () => {
    const { readFileSync } = await import('node:fs')
    const { blobRangeSource } = await import('../src/core/zip-stream')
    const { admitSpreadsheet } = await import('../src/xlsx/storage/admission')
    const { importLimitsFor } = await import('../src/xlsx/storage/importSession')
    const bytes = readFileSync('tests/fixtures/xlsx-cf/stage3-bar-lengths.xlsx')
    for (const profile of ['D', 'M'] as const) {
      const admission = await admitSpreadsheet(blobRangeSource(new Blob([bytes as BlobPart])), profile, importLimitsFor(profile), new AbortController().signal)
      expect(admission.mode, profile).toBe('object')
    }
  })
})

/**
 * Task 29 (S4-N, Lesson 10): the forced-paged semantic scope announces an icon exactly where the frame painter paints
 * it. It decides through `frameBarHosts` + `iconWithheld` on the frame's own host: the box, the CF-effective style, and
 * the `textWidth` of the frame's laid-out record. A source without `textWidth` withholds. Icons paint under injected
 * stubs only (tests/helpers/cf-icons.ts).
 */
describe('the forced-paged semantic scope announces an icon exactly where the frame paints it (S4-N)', () => {
  const ICON_ROWS = 3
  const iconFrameHosts = async (sheet: XlsxSheet, options: { width?: boolean } = {}): Promise<CfFrameHostSource> => {
    const { createCanvas } = await import('canvas')
    const { createCellLayoutHost } = await import('../src/xlsx/view/renderViewport')
    const { computeMetrics } = await import('../src/xlsx/render')
    const metrics = computeMetrics(sheet)
    const layout = createCellLayoutHost(createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D)
    return {
      cell: (row, col) => {
        const cell = sheet.rows[row]?.cells.find(c => c.col === col)
        if (!cell) return undefined
        const box = { x: 0, y: metrics.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0), width: metrics.colWidthsPx[col]!, height: metrics.rowHeightsPx[row]! }
        // The frame's laid-out record of this cell (un-inset): its text width.
        const body = layout.layoutBody(cell, box.width, box.height)
        const textWidth = body.text?.kind === 'plain' ? body.text.width : undefined
        return { box, base: cell.style, cell, ...(options.width === false || textWidth === undefined ? {} : { textWidth }) }
      },
    }
  }
  it('resident, worker and forced-paged legs announce the same labels for the same cells', async () => {
    const { iconSheet, withIcons } = await import('./helpers/cf-icons')
    const { prepareCfTextScope } = await import('../src/xlsx/cf/text')
    const { createCanvas } = await import('canvas')
    await withIcons(async () => {
      for (const [name, sheet] of [['plain', iconSheet([10, 50, 90])], ['narrow', iconSheet([1, 5, 9], { widthChars: 10 / 7 })]] as const) {
        const wire = await wireOf(sheet, ICON_ROWS)
        const semantic = (await scopeOver(wire, ICON_ROWS, 'zh-TW', await iconFrameHosts(sheet)))!
        const blind = (await scopeOver(wire, ICON_ROWS, 'zh-TW', await iconFrameHosts(sheet, { width: false })))!
        const resident = (await prepareCfTextScope(sheet, { rects: [{ r0: 0, r1: ICON_ROWS - 1, c0: 0, c1: 0 }], signal: new AbortController().signal, locale: 'zh-TW', budget: { reserve() {}, release() {} },
          layoutContext: () => createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D }))!
        // The worker leg (the host-less paged paint scope) after painting the frame's hosts.
        const painted = (await scopeOver(wire, ICON_ROWS, 'zh-TW'))!
        const hosts = await iconFrameHosts(sheet)
        const ctx = createCanvas(400, 100).getContext('2d') as unknown as CanvasRenderingContext2D
        painted.paintIcons?.(ctx, (row, col) => { const at = hosts.cell(row, col); return at ? { ...at.box, textWidth: at.textWidth, cell: at.cell } : undefined }, { r0: 0, c0: 0, r1: ICON_ROWS - 1, c1: 0 }, { zoom: 1, dpr: 1 })
        for (const row of [0, 1, 2]) {
          const want = resident.describe(row, 0)
          expect(semantic.describe(row, 0), `${name} row ${row}: forced-paged = resident`).toEqual(want)
          expect(painted.describe(row, 0), `${name} row ${row}: worker = resident`).toEqual(want)
          expect(semantic.iconAt?.(row, 0), `${name} row ${row}`).toEqual(painted.iconAt?.(row, 0))
          expect(blind.iconAt?.(row, 0), `${name} row ${row}: a source without textWidth withholds`).toBeUndefined()
          expect(blind.describe(row, 0)?.key, `${name} row ${row}`).toBe('cf-unavailable')
        }
        if (name === 'plain') expect(semantic.describe(1, 0)).toEqual({ key: 'cf:iconSet', text: '圖示 2/3' })
        else expect(semantic.describe(1, 0)?.key).toBe('cf-unavailable')
        for (const s of [semantic, blind, resident, painted]) s.dispose()
      }
    })
  })
})

/**
 * Task 29 fix round 1 (the S3-A11Y-PAGED lesson for icons): the viewer's own frame host source
 * (`createFrameHostSource`, which InteractiveSheetDoc builds for its forced-paged grid semantics) supplies each cell and
 * its text width from the projection metrics, the viewer's layout host and the CF-effective style the frame lays it out
 * with. So the forced-paged semantic leg decides each icon exactly as the resident and worker legs do.
 */
describe('the viewer frame host source announces icons as the other legs do (Task 29 fix round 1)', () => {
  const run = async (glyphs: boolean) => {
    const { iconSheet, withIcons } = await import('./helpers/cf-icons')
    const { prepareCfTextScope } = await import('../src/xlsx/cf/text')
    const { createCanvas } = await import('canvas')
    const { createFrameHostSource } = await import('../src/xlsx/view/frameHosts')
    const { createCellLayoutHost, renderSheetViewport } = await import('../src/xlsx/view/renderViewport')
    const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
    const { LayoutCache } = await import('../src/xlsx/view/textLayout')
    const ctx = () => createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
    return withIcons(async () => {
      const sheet = iconSheet([10, 50, 90])
      const geometry = buildSheetGeometry(sheet)
      if (geometry.kind !== 'ready') throw new Error('geometry')
      const layoutHost = createCellLayoutHost(ctx(), sheet)
      // The inputs InteractiveSheetDoc reads from its sheet entry: metrics, stored cells, screen base style, layout host.
      const hosts = createFrameHostSource(() => ({
        metrics: geometry.metrics,
        cellAt: at => geometry.cellsOf(at.row)?.find(c => c.col === at.col),
        styleOf: cell => cell.style,
        layoutHost: () => layoutHost,
      }))
      const wire = await wireOf(sheet, 3)
      const semantic = (await scopeOver(wire, 3, 'zh-TW', hosts))!
      const resident = (await prepareCfTextScope(sheet, { rects: [{ r0: 0, r1: 2, c0: 0, c1: 0 }], signal: new AbortController().signal, locale: 'zh-TW', budget: { reserve() {}, release() {} }, layoutContext: ctx }))!
      // The worker leg: the host-less paged paint scope after the real frame painted it.
      const worker = (await scopeOver(wire, 3, 'zh-TW'))!
      const batch = { stamp: wire.stamp, key: 'paged', layer: worker.paintLayer!,
        barAt: (r: number, c: number) => worker.barAt?.(r, c), paintBars: () => true,
        hidesTextAt: (r: number, c: number, h: never, s: never) => worker.hidesTextAt?.(r, c, h, s) ?? false,
        iconAt: (r: number, c: number) => worker.iconAt?.(r, c),
        paintIcons: (cx: CanvasRenderingContext2D, rect: never, v: { r0: number; c0: number; r1: number; c1: number }, s: never) => worker.paintIcons?.(cx, rect, v, s) ?? true,
        textInsetAt: (r: number, c: number, h: never, s: never) => worker.textInsetAt?.(r, c, h, s), dispose: () => {} }
      renderSheetViewport(sheet, createCanvas(400, 100).getContext('2d') as unknown as CanvasRenderingContext2D, geometry, { box: { x: 0, y: 0, width: 400, height: 100 }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch } as never, { cache: new LayoutCache(), layoutHost })
      const out = [0, 1, 2].map(row => ({ resident: resident.describe(row, 0), worker: worker.describe(row, 0), semantic: semantic.describe(row, 0) }))
      for (const s of [semantic, resident, worker]) s.dispose()
      return out
    }, glyphs ? {} : { glyphs: () => undefined })
  }
  it('all three legs announce the same 「圖示 b/k」; a withheld icon announces unavailable on all three', async () => {
    const painted = await run(true)
    painted.forEach((legs, row) => {
      const want = { key: 'cf:iconSet', text: `圖示 ${row + 1}/3` }
      expect(legs, `row ${row}`).toEqual({ resident: want, worker: want, semantic: want })
    })
    const withheld = await run(false)
    const unavailable = { key: 'cf-unavailable', text: '無法使用條件式格式設定' }
    withheld.forEach((legs, row) => expect(legs, `row ${row}`).toEqual({ resident: unavailable, worker: unavailable, semantic: unavailable }))
  })
  it('InteractiveSheetDoc builds its forced-paged semantic hosts with createFrameHostSource', async () => {
    const { readFileSync } = await import('node:fs')
    const source = readFileSync('src/components/xlsx/InteractiveSheetDoc.tsx', 'utf8')
    expect(source).toMatch(/createFrameHostSource\(/)
    expect(source).not.toMatch(/const frameHostSource = \(sheetId: string\)[^]*?return box \? \{ box, base: styleForCell/)
  })
})

describe('a formula-bounded bar and scale announce on the forced-paged scope as on the resident one (Task 37 follow-up, S5-P)', () => {
  it('a host the frame lays out announces the position under $F$1; a 3 px host withholds, on both scopes alike', async () => {
    const { formulaBoundsBytes, load, textLeg, workerLeg, withPaint, frameHosts, firstSelections } = await import('./helpers/cf-views')
    const { injectCfFormulaSelections, VERIFIED_CAPABILITIES } = await import('../src/xlsx/cf/evidence')
    const { REGISTERED } = await import('./helpers/cf-views')
    const { withIcons } = await import('./helpers/cf-icons')
    await withIcons(async () => {
      injectCfFormulaSelections(firstSelections())
      try {
        const l = await load(formulaBoundsBytes(), 'cached')
        const refs = ['B2', 'B5', 'B8', 'C5', 'C8']
        for (const [width, expectBar] of [[200, true], [3, false]] as const) {
          const hosts = frameHosts(width)
          const resident = await textLeg(l, refs, { hosts, resident: true }), paged = (await workerLeg(l, refs, { hosts })).scope
          expect(Object.fromEntries([...paged].map(([k, v]) => [k, v.desc]))).toEqual(Object.fromEntries([...resident].map(([k, v]) => [k, v.desc])))
          expect(resident.get('B8')!.desc, `width ${width}`).toBe(expectBar ? 'cf:dataBar|Data bar, 100% of range' : 'cf-unavailable|Conditional formatting unavailable')
          expect(paged.get('C8')!.desc).toBe('cf:colorScale|Colour scale, 100% of range')
        }
      } finally { injectCfFormulaSelections(null) }
    }, { extra: [...VERIFIED_CAPABILITIES, ...REGISTERED, 'bar:participation-dateserial'] })
    void withPaint
  })
})
