// @vitest-environment node
/**
 * Task 22: data bars paint in every view through one lazy painter (`src/xlsx/cf/dataBarPaint.ts`), reached by the eager
 * renderers only through `CfPaintBatch.barAt`/`paintBars`/`hidesTextAt`.
 *
 * Every bar here is painted under INJECTED evidence (tests/helpers/cf-bars.ts): nothing Stage 3 is published, so in
 * production every bar stays withheld until Task 23. The model constants are checked against the Excel sidecars; the
 * Lesson 1 geometry conformance against every measured cell lives in tests/xlsx-cf-oracle-conformance.test.ts.
 */
import { readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { computeMetrics, computePrintMetrics, renderPrintPage, renderSheet } from '../src/xlsx/render'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { prepareCfPaint, residentSource, type CfPaintBatch } from '../src/xlsx/cf/render'
import { prepareCfWireScope } from '../src/xlsx/cf/text'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_BAR_PAINT_CAPABILITIES, CF_BAR_RENDER_MODEL, barGradientEndArgb, layoutDataBar, paintDataBars } from '../src/xlsx/cf/dataBarPaint'
import { injectCfBarGradient, injectCfBarGradientBorder } from '../src/xlsx/cf/evidence'
import { predictBarGradientBorderPixel } from '../scripts/cf-refs/verify'
import { getPaintables } from '../src/render/paint'
import { decodePngSync } from '../scripts/cf-refs/verify'
import { auditStage3Manifest } from '../scripts/cf-refs/audit'
import type { CfBarDescriptor, CfStamp } from '../src/xlsx/cf/model'
import type { CfManifest } from '../scripts/cf-refs/cases'
import type { XlsxSheet } from '../src/xlsx/types'
import { BAR_SHOWVALUE_EVIDENCE, BLUE, GREEN, barSheet, withBars } from './helpers/cf-bars'

const REPO_ROOT = resolve(__dirname, '..')
const REFS = join(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs')
type Sidecar = { colWidthsPt: number[]; rowHeightsPt: number[]; observedBars: Array<{ ref: string; cellLeftPx: number; cellWidthPx: number; bar: { leftPx: number; rightPx: number; topPx: number; bottomPx: number } | null; axisPx?: number }> }
const sidecar = (id: string): Sidecar => JSON.parse(readFileSync(join(REFS, `${id}.json`), 'utf8')) as Sidecar
const rowOf = (ref: string): number => Number(/\d+/.exec(ref)![0]) - 1
const rowTop = (s: Sidecar, ref: string): number => s.rowHeightsPt.slice(0, rowOf(ref)).reduce((a, b) => a + b, 0)
/** At zoom × DPR = 3/4 one point is one device px: the capture scale (Excel zoom 100, 1 px/pt). */
const AT_CAPTURE = { zoom: 0.75, dpr: 1, output: 'device' as const }
const ptCell = (left: number, top: number, width: number, height: number) => ({ x: left * 4 / 3, y: top * 4 / 3, width: width * 4 / 3, height: height * 4 / 3 })
const ALL = (): boolean => true

const bar = (overrides: Partial<CfBarDescriptor> = {}): CfBarDescriptor => ({
  ruleId: 1, axisFraction: 0, startFraction: 0, endFraction: 0.5, fillColor: BLUE, isGradient: false, isNegative: false,
  showValue: true, direction: 'leftToRight', ...overrides,
})

describe('CF_BAR_RENDER_MODEL equals the measured sidecars', () => {
  it('insets, border and axis widths come from the Stage 3 captures; unmeasured values are absent', () => {
    // Insets: every measured bar of the geometry and rendering captures (three widths, three row heights) starts 2 px
    // right of its cell's left gridline (left-to-right bars), keeps 2 px below the row top and 2 px above the next row;
    // the full-length bars stop 1 px short of the next column's gridline. No expected value comes from the painter.
    const left = new Set<number>(), top = new Set<number>(), bottom = new Set<number>(), right: number[] = []
    for (const id of ['stage3-bar-geometry', 'stage3-bar-rendering']) {
      const s = sidecar(id)
      for (const o of s.observedBars) {
        if (!o.bar) continue
        const col = o.ref.replace(/\d+/, '')
        if (id === 'stage3-bar-rendering' && col === 'F') { right.push(o.cellLeftPx + o.cellWidthPx - (o.bar.rightPx + 1)); continue } // right-to-left
        left.add(o.bar.leftPx - o.cellLeftPx)
        const y0 = rowTop(s, o.ref), h = s.rowHeightsPt[rowOf(o.ref)]!
        top.add(o.bar.topPx - y0)
        bottom.add(y0 + h - (o.bar.bottomPx + 1))
        right.push(o.cellLeftPx + o.cellWidthPx - (o.bar.rightPx + 1))
      }
    }
    expect([...left]).toEqual([CF_BAR_RENDER_MODEL.insetsPt.left])
    expect([...top]).toEqual([CF_BAR_RENDER_MODEL.insetsPt.top])
    expect([...bottom]).toEqual([CF_BAR_RENDER_MODEL.insetsPt.bottom])
    expect(Math.min(...right)).toBe(CF_BAR_RENDER_MODEL.insetsPt.right)
    expect(right.filter(r => r === CF_BAR_RENDER_MODEL.insetsPt.right).length).toBeGreaterThanOrEqual(3)

    // Border: the bordered column D of the rendering capture draws a frame exactly borderPt wide inside the bar rect.
    const png = decodePngSync(new Uint8Array(readFileSync(join(REFS, 'stage3-bar-rendering.png'))))!
    const px = (x: number, y: number) => [0, 1, 2].map(k => png.data[(y * png.width + x) * 4 + k]!)
    const near = (p: number[], argb: number) => [16, 8, 0].every((sh, k) => Math.abs(p[k]! - ((argb >> sh) & 0xff)) <= 8)
    const rendering = sidecar('stage3-bar-rendering')
    let framed = 0
    for (const o of rendering.observedBars.filter(o => o.ref.startsWith('D') && o.bar)) {
      const mid = Math.floor((o.bar!.topPx + o.bar!.bottomPx) / 2), midX = Math.floor((o.bar!.leftPx + o.bar!.rightPx) / 2)
      let w = 0; while (near(px(o.bar!.leftPx + w, mid), GREEN)) w++
      let h = 0; while (near(px(midX, o.bar!.topPx + h), GREEN)) h++
      expect(w, o.ref).toBe(CF_BAR_RENDER_MODEL.borderPt)
      expect(h, o.ref).toBe(CF_BAR_RENDER_MODEL.borderPt)
      expect(near(px(o.bar!.leftPx + w, mid), BLUE), o.ref).toBe(true)
      framed++
    }
    expect(framed).toBeGreaterThanOrEqual(3)

    // Axis: one column at axisPx, its neighbours carry no axis ink, over the zero cell with no bar (axis-negative B4).
    const axisPng = decodePngSync(new Uint8Array(readFileSync(join(REFS, 'stage3-bar-axis-negative.png'))))!
    const apx = (x: number, y: number) => [0, 1, 2].map(k => axisPng.data[(y * axisPng.width + x) * 4 + k]!)
    const axisNeg = sidecar('stage3-bar-axis-negative')
    const zero = axisNeg.observedBars.find(o => o.ref === 'B4')!
    expect(zero.bar).toBeNull()
    const y0 = rowTop(axisNeg, 'B4')
    const inked = (x: number) => Array.from({ length: axisNeg.rowHeightsPt[3]! }, (_, i) => apx(x, y0 + i)).filter(p => p.every(c => c < 64)).length
    expect(inked(zero.axisPx!)).toBeGreaterThan(axisNeg.rowHeightsPt[3]! / 2)
    expect(inked(zero.axisPx! - 1)).toBe(0)
    expect(inked(zero.axisPx! + CF_BAR_RENDER_MODEL.axisPt)).toBe(0)

    // Unmeasured: the minimum visible width carries no value and the harness does not verify it. The gradient is not a
    // model constant: it is the oracle's verified export (Task 22b), read through barGradientEvidence.
    expect(CF_BAR_RENDER_MODEL.gradient).toBeUndefined()
    expect(CF_BAR_RENDER_MODEL.minVisiblePt).toBeUndefined()
    const manifest = JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
    const audit = auditStage3Manifest(manifest, { rootDir: REPO_ROOT })
    const disposition = (cap: string) => audit.perCapability.find(p => p.capabilityId === cap)?.disposition
    expect(disposition('bar:render-gradient-stops')).toBe('verified')
    expect(audit.barGradient).toEqual({ 'bar-length': { endBlend: 0.847, settledBy: ['bar:render-gradient-stops'] } })
    expect(disposition('bar:geometry-min-visible-width')).not.toBe('verified')
    for (const cap of [...CF_BAR_PAINT_CAPABILITIES, 'bar:render-border']) expect(disposition(cap), cap).toBe('verified')
  }, 30_000)

  it('without a verified capability, the matching branch of layoutDataBar returns undefined', () => {
    const cell = ptCell(84, 30, 180, 30)
    expect(layoutDataBar(bar(), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)?.fill).toBeDefined()
    for (const cap of CF_BAR_PAINT_CAPABILITIES) expect(layoutDataBar(bar(), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, c => c !== cap), cap).toBeUndefined()
    // Gradient: no verified gradient export (none is published), so never drawn (and never drawn solid instead).
    expect(layoutDataBar(bar({ isGradient: true }), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
    // Border: needs bar:render-border; with it the frame is borderPt wide.
    expect(layoutDataBar(bar({ borderColor: GREEN }), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, c => c !== 'bar:render-border')).toBeUndefined()
    const framed = layoutDataBar(bar({ borderColor: GREEN }), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)!
    expect(framed.border!.map(r => Math.min(r.w, r.h))).toEqual([1, 1, 1, 1])
    // Minimum visible width: a non-zero share that rounds to no pixel is the unmeasured branch; zero is no bar.
    expect(layoutDataBar(bar({ endFraction: 0.001 }), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
    expect(layoutDataBar(bar({ endFraction: 0 }), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toEqual({})
  })
})

/** Records every fill and text call with the transform it ran under, in order. */
function recorder(width: number, height: number) {
  const canvas = createCanvas(Math.max(1, Math.ceil(width)), Math.max(1, Math.ceil(height)))
  const ctx = canvas.getContext('2d')
  type Op = { op: 'rect' | 'text'; style: string; x: number; y: number; w: number; h: number; text?: string }
  const ops: Op[] = []
  const fillRect = ctx.fillRect.bind(ctx), fillText = ctx.fillText.bind(ctx)
  ctx.fillRect = (x: number, y: number, w: number, h: number) => {
    const t = ctx.getTransform()
    ops.push({ op: 'rect', style: String(ctx.fillStyle).toLowerCase(), x: t.a * x + t.e, y: t.d * y + t.f, w: t.a * w, h: t.d * h })
    fillRect(x, y, w, h)
  }
  ctx.fillText = (text: string, x: number, y: number) => {
    const t = ctx.getTransform()
    ops.push({ op: 'text', style: String(ctx.fillStyle).toLowerCase(), x: t.a * x + t.e, y: t.d * y + t.f, w: 0, h: 0, text })
    fillText(text, x, y)
  }
  return { ctx: ctx as unknown as CanvasRenderingContext2D, ops, canvas }
}

/** The paged main thread's batch: the wire scope (production `prepareCfRemote` → structured clone → `prepareCfWireScope`)
 * handed to the painter exactly as InteractiveSheetDoc does. */
async function pagedBatch(sheet: XlsxSheet, rows: number, cols: number): Promise<CfPaintBatch> {
  const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId ?? sheet.name
  const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
  const memory = createCfMemory({})
  const rects = [{ r0: 0, r1: rows - 1, c0: 0, c1: cols - 1 }]
  const wire = await prepareCfRemote({ rects, stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, { budget: () => createCfWorkBudget({ now: () => 0 }) })
  expect(wire.kind, JSON.stringify(wire.reason)).toBe('ready')
  const scope = (await prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), rects, { budget: { reserve() {}, release() {} } }))!
  return {
    stamp: st, key: 'paged', layer: scope.paintLayer!,
    barAt: (r, c) => scope.barAt?.(r, c),
    paintBars: (ctx, cellRect, visible, scale) => scope.paintBars?.(ctx, cellRect, visible, scale) ?? true,
    hidesTextAt: (r, c, host, scale) => scope.hidesTextAt?.(r, c, host, scale) ?? false,
    dispose: () => scope.dispose(),
  }
}

/** Static yellow fills on every bar cell, so the record shows fill → bar → content for one host. The axis is the
 * measured automatic white: an explicit (non-white) axis over a filled host withholds (its dash gap is unmeasured). */
function fivesSheet(): XlsxSheet {
  const sheet = barSheet([-50, 0, 50, 100], { def: { axisArgb: 0xffffffff } })
  for (const row of sheet.rows) for (const cell of row.cells) cell.style = { numFmtId: 0, fillColor: 'FFFFFF00' }
  return sheet
}

describe('bars paint in the fill phase in all five views', () => {
  it('fill → bar → content in renderSheet, renderSheetViewport (resident and paged), renderPrintPage at PRINT_SCALE = 2 and export', async () => {
    await withBars(async () => {
      const sheet = fivesSheet()
      const metrics = computeMetrics(sheet)
      const cellX = 0, cellY = metrics.rowHeightsPx[0]! + metrics.rowHeightsPx[1]!, cellW = metrics.colWidthsPx[0]!, cellH = metrics.rowHeightsPx[2]!
      expect(cellW).toBe(200)
      const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      const paged = await pagedBatch(sheet, 4, 1)
      const S = 2 // device px per logical px in every view, so the source-space rects are comparable exactly
      const views: Record<string, ReturnType<typeof recorder>> = {}
      {
        const r = views.static = recorder(metrics.widthPx * S, metrics.heightPx * S)
        r.ctx.scale(S, S)
        renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: resident })
      }
      const geometry = buildSheetGeometry(sheet)
      for (const [name, batch] of [['resident', resident], ['paged', paged]] as const) {
        const r = views[name] = recorder(metrics.widthPx * S, metrics.heightPx * S)
        r.ctx.scale(S, S)
        renderSheetViewport(sheet, r.ctx, geometry, { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: S }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
      }
      {
        const print = computePrintMetrics(sheet, metrics)
        expect(print.scale).toBe(1)
        const r = views.print = recorder(print.paperPx.width * S, print.paperPx.height * S)
        r.ctx.scale(S, S) // PRINT_SCALE = 2
        renderPrintPage(sheet, r.ctx, metrics, print, 0, 0, undefined, { conditionalFormatting: resident })
      }
      const pages = await getPaintables({ sheets: [sheet] })
      try {
        const r = views.export = recorder(pages[0]!.spec.widthPx * S, pages[0]!.spec.heightPx * S)
        r.ctx.scale(S, S)
        pages[0]!.paint(r.ctx)
      } finally { pages.dispose() }

      // The 50 of [-50, 100] under an automatic axis: the bar spans [1/3, 2/3] of the 200 px cell (the paged batch
      // reports it once it has painted it).
      for (const b of [resident, paged]) {
        const d = b.barAt(2, 0)!
        expect(d.axisFraction).toBeCloseTo(1 / 3, 12); expect(d.startFraction).toBeCloseTo(1 / 3, 12); expect(d.endFraction).toBeCloseTo(2 / 3, 12)
      }
      const blue = '#0000ff', yellow = '#ffff00'
      const sourceRects: Record<string, { x: number; y: number; w: number; h: number }> = {}
      const axisRects: Record<string, string[]> = {}
      for (const [name, { ops }] of Object.entries(views)) {
        // Calibrate device → source from the host's own static fill (its source rect is known).
        const fill = ops.filter(o => o.op === 'rect' && o.style === yellow)[2]
        expect(fill, `${name}: host fill`).toBeDefined()
        const fillIndex = ops.indexOf(fill!), k = fill!.w / cellW, ox = fill!.x - cellX * k, oy = fill!.y - cellY * k
        expect(fill!.h / k).toBeCloseTo(cellH, 9)
        const barIndex = ops.findIndex(o => o.op === 'rect' && o.style === blue && (o.y - oy) / k >= cellY && (o.y - oy) / k < cellY + cellH)
        const textIndex = ops.findIndex(o => o.op === 'text' && o.text === '50')
        expect(barIndex, `${name}: bar painted`).toBeGreaterThan(fillIndex)
        expect(textIndex, `${name}: text after the bar`).toBeGreaterThan(barIndex)
        // No fill of any cell lands after its row's bar in this view (the fill phase ends before bars).
        expect(ops.slice(barIndex).some(o => o.op === 'rect' && o.style === yellow), `${name}: fills precede bars`).toBe(false)
        const b = ops[barIndex]!
        sourceRects[name] = { x: (b.x - ox) / k, y: (b.y - oy) / k, w: b.w / k, h: b.h / k }
        // The axis column (gap underlay and dashes) of the host, in source space.
        axisRects[name] = ops.filter(o => o.op === 'rect' && o.style === '#ffffff' && o.w / k < 2 && (o.y - oy) / k >= cellY - 1e-9 && (o.y - oy) / k < cellY + cellH)
          .map(o => [(o.x - ox) / k, (o.y - oy) / k, o.w / k, o.h / k].map(v => v.toFixed(6)).join(','))
      }
      expect(axisRects.static!.length, 'the axis paints its gap column and its dashes').toBeGreaterThan(2)
      for (const [name, rects] of Object.entries(axisRects)) expect(rects, `${name} axis`).toEqual(axisRects.static)
      const reference = sourceRects.static!
      // [1/3, 2/3] of the 200 px host, inside the measured insets.
      expect(reference.x).toBeGreaterThan(cellX + 66); expect(reference.x + reference.w).toBeLessThan(cellX + 136)
      for (const [name, r] of Object.entries(sourceRects)) for (const k of ['x', 'y', 'w', 'h'] as const) expect(r[k], `${name}.${k}`).toBeCloseTo(reference[k], 9)
      resident.dispose(); paged.dispose()
    })
  })
})

/** The gradient model the audit verifies (Task 22b), injected as Task 23's export would publish it. */
const GRADIENT_EXPORT = { 'bar-length': { endBlend: 0.847, settledBy: ['bar:render-gradient-stops'] } }

describe('gradient bars (Task 22b): one painter, the verified model only, the measured shape only', () => {
  it('layoutDataBar draws the gradient from the export alone: fill at the left edge, the blend toward white at the right', () => {
    const cell = ptCell(264, 150, 180, 30)
    const g = bar({ isGradient: true, endFraction: 1 })
    expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL), 'no export: withheld').toBeUndefined()
    injectCfBarGradient(GRADIENT_EXPORT)
    try {
      const p = layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)!
      expect(p.fill).toEqual({ x: 266, y: 152, w: 177, h: 26, argb: BLUE, gradientEnd: barGradientEndArgb(BLUE, 0.847) })
      // 0.847 of the way to white: 0 + 255 x 0.847 = 215.985 -> 216 on red and green, blue stays 255.
      expect(barGradientEndArgb(BLUE, 0.847)).toBe(0xffd8d8ff)
      expect(barGradientEndArgb(0xff638ec6, 0.847)).toBe(0xffe7eef6)
      // The export's capability gates it.
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, c => c !== 'bar:render-gradient-stops')).toBeUndefined()
      // Only the measured shape paints: no right-to-left, border, axis, negative bar or filled host was ever measured.
      expect(layoutDataBar({ ...g, direction: 'rightToLeft' }, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      expect(layoutDataBar({ ...g, borderColor: GREEN }, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      expect(layoutDataBar({ ...g, axisColor: 0xff000000, axisFraction: 0.25, startFraction: 0.25 }, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      expect(layoutDataBar({ ...g, isNegative: true, axisFraction: 1, startFraction: 0.5 }, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      expect(layoutDataBar(g, { ...cell, fill: '#ffff00' }, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      expect(layoutDataBar(g, { ...cell, fill: '#ffffff' }, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)?.fill?.gradientEnd).toBeDefined()
      // A solid bar is unaffected by the export.
      expect(layoutDataBar(bar({ endFraction: 1 }), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)!.fill!.gradientEnd).toBeUndefined()
      // A malformed export never paints.
      injectCfBarGradient({ 'bar-length': { endBlend: 1.5, settledBy: ['bar:render-gradient-stops'] } })
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      injectCfBarGradient({ 'bar-length': { endBlend: 0.847, settledBy: [] } })
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
    } finally {
      injectCfBarGradient(null)
    }
  })

  it('every view paints the same gradient bar through the one painter: blue at the left edge, the blend at the right', async () => {
    const sheet = barSheet([0, 50, 100], { def: { gradient: true } })
    const paintAll = async () => withBars(async () => {
      const metrics = computeMetrics(sheet)
      const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      const paged = await pagedBatch(sheet, 3, 1)
      const S = 2
      const views: Record<string, ReturnType<typeof recorder>> = {}
      { const r = views.static = recorder(metrics.widthPx * S, metrics.heightPx * S); r.ctx.scale(S, S); renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: resident }) }
      const geometry = buildSheetGeometry(sheet)
      for (const [name, batch] of [['resident', resident], ['paged', paged]] as const) {
        const r = views[name] = recorder(metrics.widthPx * S, metrics.heightPx * S)
        r.ctx.scale(S, S)
        renderSheetViewport(sheet, r.ctx, geometry, { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: S }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
      }
      {
        const print = computePrintMetrics(sheet, metrics)
        const r = views.print = recorder(print.paperPx.width * S, print.paperPx.height * S)
        r.ctx.scale(S, S)
        renderPrintPage(sheet, r.ctx, metrics, print, 0, 0, undefined, { conditionalFormatting: resident })
      }
      const pages = await getPaintables({ sheets: [sheet] })
      try { const r = views.export = recorder(pages[0]!.spec.widthPx * S, pages[0]!.spec.heightPx * S); r.ctx.scale(S, S); pages[0]!.paint(r.ctx) } finally { pages.dispose() }
      resident.dispose(); paged.dispose()
      return views
    }, { extra: ['bar:render-gradient-stops'] })

    injectCfBarGradient(GRADIENT_EXPORT)
    let views: Awaited<ReturnType<typeof paintAll>>
    try { views = await paintAll() } finally { injectCfBarGradient(null) }
    const end = [216, 216, 255]
    const rects: Record<string, string> = {}
    for (const [name, { ops, canvas }] of Object.entries(views)) {
      // The bar of the 100 (row 2) is the one fill whose style is a gradient object, not a colour string.
      const grads = ops.filter(o => o.op === 'rect' && !o.style.startsWith('#') && !o.style.startsWith('rgb'))
      expect(grads.length, `${name}: two gradient bars (the 0 draws none)`).toBe(2)
      const b = grads[1]!
      rects[name] = [b.w, b.h].map(v => v.toFixed(6)).join(',')
      const img = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height)
      const at = (x: number, y: number) => [0, 1, 2].map(k => img.data[(Math.floor(y) * img.width + Math.floor(x)) * 4 + k]!)
      const midY = b.y + b.h / 2
      // One column in from each edge: print's page origin is fractional, so its edge columns are partly covered.
      const first = at(b.x + 1.5, midY), last = at(b.x + b.w - 1.5, midY)
      expect(first.every((c, k) => Math.abs(c - [0, 0, 255][k]!) <= 8), `${name}: first column ${first}`).toBe(true)
      expect(last.every((c, k) => Math.abs(c - end[k]!) <= 8), `${name}: last column ${last}`).toBe(true)
    }
    for (const [name, r] of Object.entries(rects)) expect(r, `${name} bar size`).toEqual(rects.static)

    // Without the export no view paints a gradient bar (never solid instead).
    for (const [name, { ops }] of Object.entries(await paintAll())) {
      expect(ops.some(o => o.op === 'rect' && (o.style === '#0000ff' || (!o.style.startsWith('#') && !o.style.startsWith('rgb')))), `${name}: withheld`).toBe(false)
    }
  })
})

describe('bordered gradient bars (Ruling S3-GRAD-BORDER): the verified shape only, painted as the harness predicts', () => {
  const cell = ptCell(264, 150, 180, 30)
  const shapes = ['frame-over-ramp', 'frame-inner-ramp', 'gradient-frame', 'no-frame', 'frame-solid-fill'] as const
  const exportOf = (shape: string) => ({ [shape]: { ...(shape === 'frame-solid-fill' ? {} : { endBlend: 0.847 }), settledBy: ['bar:render-gradient-border'] } })

  it('withholds without the bordered export, even with the borderless gradient export', () => {
    const g = bar({ isGradient: true, endFraction: 0.5, borderColor: GREEN })
    injectCfBarGradient(GRADIENT_EXPORT)
    try {
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      injectCfBarGradientBorder(exportOf('frame-over-ramp'))
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)?.fill).toBeDefined()
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, c => c !== 'bar:render-gradient-border')).toBeUndefined()
      // An export naming two shapes, an unknown one or no blend never paints.
      injectCfBarGradientBorder({ ...exportOf('frame-over-ramp'), ...exportOf('no-frame') })
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      injectCfBarGradientBorder({ 'frame-around': { endBlend: 0.847, settledBy: ['bar:render-gradient-border'] } })
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      injectCfBarGradientBorder({ 'frame-over-ramp': { settledBy: ['bar:render-gradient-border'] } })
      expect(layoutDataBar(g, cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
      // The bordered model never paints a borderless gradient, which keeps its own export.
      injectCfBarGradient(null)
      injectCfBarGradientBorder(exportOf('frame-over-ramp'))
      expect(layoutDataBar(bar({ isGradient: true, endFraction: 0.5 }), cell, AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
    } finally {
      injectCfBarGradient(null)
      injectCfBarGradientBorder(null)
    }
  })

  for (const shape of shapes) {
    it(`${shape}: the painter reproduces the harness prediction pixel by pixel, at bar lengths 5 to 177`, () => {
      injectCfBarGradientBorder(exportOf(shape))
      try {
        for (const [L, border] of [[5, GREEN], [9, BLUE], [18, GREEN], [89, BLUE], [177, GREEN]] as const) {
          const b = bar({ isGradient: true, endFraction: L / 177, borderColor: border })
          const canvas = createCanvas(450, 190)
          const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
          ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 450, 190)
          ctx.setTransform(0.75, 0, 0, 0.75, 0, 0)
          paintDataBars(ctx, new Map([[5 * 16384 + 2, b]]), () => b, () => cell, { r0: 5, c0: 2, r1: 5, c1: 2 }, AT_CAPTURE, ALL)
          const img = (ctx as unknown as { getImageData(x: number, y: number, w: number, h: number): ImageData }).getImageData(0, 0, 450, 190)
          const rgb = (n: number) => [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff] as [number, number, number]
          const sample = { ref: 'C6', fill: rgb(BLUE), border: rgb(border), leftPx: 266, topPx: 152, bottomPx: 177, length: L, span: 177, pixels: [] }
          let worst = 0
          for (let y = 152; y <= 177; y++) for (let i = 0; i < L; i++) {
            const want = predictBarGradientBorderPixel(shape, sample, i, y, 0.847).map(Math.round)
            const k = (y * 450 + 266 + i) * 4
            worst = Math.max(worst, ...[0, 1, 2].map(j => Math.abs(img.data[k + j]! - want[j]!)))
          }
          expect(worst, `${shape} L=${L}`).toBeLessThanOrEqual(2)
          // Nothing outside the bar's rect.
          expect([0, 1, 2].map(j => img.data[(160 * 450 + 266 + L) * 4 + j])).toEqual([255, 255, 255])
        }
      } finally {
        injectCfBarGradientBorder(null)
      }
    })
  }
})

/** Two 200 px columns; the bar host B crosses the 256 px tile seam. */
function seamSheet(): XlsxSheet {
  const sheet = barSheet([10, 40, 70, 100], { col: 1 })
  sheet.cols = [{ min: 0, max: 1, widthChars: 200 / 7 }]
  return sheet
}

type DeviceRect = { style: string; x: number; y: number; w: number; h: number }
async function tiledBars(sheet: XlsxSheet, batchFor: (tx: number) => CfPaintBatch | undefined, zoom: number, dpr: number, tile = 256): Promise<Map<number, DeviceRect[]>> {
  const metrics = computeMetrics(sheet), geometry = buildSheetGeometry(sheet), cache = new LayoutCache()
  const out = new Map<number, DeviceRect[]>()
  for (let tx = 0; tx < metrics.widthPx; tx += tile) for (let ty = 0; ty < metrics.heightPx; ty += tile) {
    const r = recorder(tile * zoom * dpr, tile * zoom * dpr)
    r.ctx.scale(zoom * dpr, zoom * dpr)
    const batch = batchFor(tx)
    renderSheetViewport(sheet, r.ctx, geometry, { box: { x: tx, y: ty, width: tile, height: tile }, zoom, deviceScale: dpr }, batch ? { conditionalFormatting: batch } : undefined, { cache })
    const rects = out.get(tx) ?? []
    // Back to the sheet-global device grid: add the tile's own device origin.
    for (const o of r.ops) if (o.op === 'rect' && o.style === '#0000ff') rects.push({ style: o.style, x: o.x + tx * zoom * dpr, y: o.y + ty * zoom * dpr, w: o.w, h: o.h })
    out.set(tx, rects)
  }
  return out
}

describe('tile and full render carry identical bar primitives at six scales', () => {
  it('a bar host straddling a tile edge has the same device-snapped rects and colours in full and tiled renders', async () => {
    await withBars(async () => {
      const sheet = seamSheet()
      const metrics = computeMetrics(sheet)
      expect(metrics.colWidthsPx[0]! < 256 && metrics.colWidthsPx[0]! + metrics.colWidthsPx[1]! > 256).toBe(true)
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      let checked = 0
      for (const zoom of [0.5, 1, 1.75]) for (const dpr of [1, 2]) {
        const s = zoom * dpr
        const full = recorder(metrics.widthPx * s, metrics.heightPx * s)
        full.ctx.scale(s, s)
        renderSheet(sheet, full.ctx, metrics, undefined, { conditionalFormatting: batch })
        const fullBars = full.ops.filter(o => o.op === 'rect' && o.style === '#0000ff').map(({ style, x, y, w, h }) => ({ style, x, y, w, h }))
        // The minimum (10) has a zero-length bar: three rects.
        expect(fullBars, `zoom ${zoom} dpr ${dpr}`).toHaveLength(3)
        // Snapped: every edge on the device grid.
        for (const r of fullBars) for (const v of [r.x, r.y, r.w, r.h]) expect(Math.abs(v - Math.round(v)), `zoom ${zoom} dpr ${dpr}`).toBeLessThan(1e-9)
        const tiles = await tiledBars(sheet, () => batch, zoom, dpr)
        const first = tiles.get(0)!, seam = tiles.get(256)!
        // The longest bar (100) crosses x = 256: both tiles paint it, identically to the full render.
        const longest = fullBars.reduce((a, b) => (b.w > a.w ? b : a))
        expect(longest.x / s < 256 && (longest.x + longest.w) / s > 256).toBe(true)
        for (const [name, rects] of [['first tile', first], ['seam tile', seam]] as const) {
          const match = rects.find(r => Math.abs(r.y - longest.y) < 1e-9)
          expect(match, `${name} at zoom ${zoom} dpr ${dpr}`).toBeDefined()
          for (const k of ['x', 'y', 'w', 'h'] as const) expect(match![k], `${name}.${k} at zoom ${zoom} dpr ${dpr}`).toBeCloseTo(longest[k], 9)
          expect(match!.style).toBe(longest.style)
        }
        // Every bar of the first tile equals its full-render bar.
        for (const r of first) expect(fullBars.some(f => (['x', 'y', 'w', 'h'] as const).every(k => Math.abs(f[k] - r[k]) < 1e-9))).toBe(true)
        checked++
      }
      expect(checked).toBe(6)
      batch.dispose()
    })
  })

  it('RED mutant proof: dropping bars on non-first tiles fails the seam check', async () => {
    await withBars(async () => {
      const sheet = seamSheet()
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      const tiles = await tiledBars(sheet, tx => (tx === 0 ? batch : undefined), 1, 1)
      expect(tiles.get(0)!.length, 'mutant: the first tile still paints its bars').toBeGreaterThan(0)
      expect(tiles.get(256)!, 'mutant: the seam tile has no bar, so the parity check above would fail').toHaveLength(0)
      batch.dispose()
    })
  })
})

describe('narrow cell clips the bar to the host rect', () => {
  it('with bar:geometry-narrow-clip a 2-character column keeps every bar pixel inside the cell; without it barAt is undefined', async () => {
    const sheet = barSheet([0, 100], { widthChars: 2 })
    const metrics = computeMetrics(sheet)
    const w = metrics.colWidthsPx[0]!
    expect(w).toBeLessThan(20)
    await withBars(async () => {
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      expect(batch.barAt(1, 0)).toBeDefined()
      for (const s of [0.75, 1, 2]) {
        const r = recorder(metrics.widthPx * s, metrics.heightPx * s)
        r.ctx.scale(s, s)
        renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: batch })
        const bars = r.ops.filter(o => o.op === 'rect' && o.style === '#0000ff')
        expect(bars, `scale ${s}`).toHaveLength(1)
        const y1 = metrics.rowHeightsPx[0]!
        for (const b of bars) {
          expect(b.x).toBeGreaterThanOrEqual(0); expect(b.x + b.w).toBeLessThanOrEqual(w * s)
          expect(b.y).toBeGreaterThanOrEqual(y1 * s); expect(b.y + b.h).toBeLessThanOrEqual((y1 + metrics.rowHeightsPx[1]!) * s)
        }
        // The canvas agrees: no bar ink right of the host.
        const data = r.canvas.getContext('2d').getImageData(Math.ceil(w * s), 0, Math.max(1, Math.floor(metrics.widthPx * s - w * s)), Math.ceil(metrics.heightPx * s)).data
        for (let i = 0; i < data.length; i += 4) expect(data[i] === 0 && data[i + 1] === 0 && data[i + 2] === 255, `scale ${s}`).toBe(false)
      }
      batch.dispose()
    })
    await withBars(async () => {
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      expect(batch.barAt(1, 0)).toBeUndefined()
      batch.dispose()
    }, { omit: ['bar:geometry-narrow-clip'] })
  })
})

describe('CF bars never expand static, print or export extents', () => {
  it('a whole-column bar on A:A leaves the print canvas and export image at the data extent', async () => {
    await withBars(async () => {
      const dataRows = 10
      const sheet = barSheet(Array.from({ length: dataRows }, (_, r) => (r + 1) * 10))
      ;(sheet.conditionalFormatting!.rules[0]!.ranges as { firstRow: number; lastRow: number; firstCol: number; lastCol: number }[]).splice(0, 1, { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 })
      const m = computeMetrics(sheet)
      expect(m.rowHeightsPx.length).toBe(dataRows)
      const print = computePrintMetrics(sheet, m)
      expect(print.pagesWide * print.pagesTall).toBe(1)
      expect(print.paperPx.height).toBeLessThan(2000)
      // The bars do paint (a positive control), yet the export page stays at the data extent.
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      expect(batch.barAt(dataRows - 1, 0)).toBeDefined()
      expect(batch.barAt(dataRows + 5, 0)).toBeUndefined()
      batch.dispose()
      const pages = await getPaintables({ sheets: [sheet] })
      try {
        expect(pages).toHaveLength(1)
        expect(pages[0]!.spec.heightPx).toBe(m.heightPx)
        expect(pages[0]!.spec.widthPx).toBe(m.widthPx)
      } finally { pages.dispose() }
    })
  })
})

describe('showValue=0 hides canvas text only under evidence (S3-O)', () => {
  const texts = async (sheet: XlsxSheet, view: 'static' | 'viewport', batch: CfPaintBatch) => {
    const metrics = computeMetrics(sheet)
    const r = recorder(metrics.widthPx, metrics.heightPx)
    if (view === 'static') renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: batch })
    else renderSheetViewport(sheet, r.ctx, buildSheetGeometry(sheet), { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
    return { texts: r.ops.filter(o => o.op === 'text').map(o => o.text), bars: r.ops.filter(o => o.op === 'rect' && o.style === '#0000ff').length }
  }
  it('with bar:showvalue-canvas-suppression the content phase draws no glyph run for the cell, and the bar paints', async () => {
    await withBars(async () => {
      const sheet = barSheet([10, 40, 70], { def: { showValue: false } })
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      for (const view of ['static', 'viewport'] as const) {
        const got = await texts(sheet, view, batch)
        expect(got.bars, view).toBe(2) // the minimum's bar has zero length
        expect(got.texts, view).toEqual([])
      }
      expect(batch.hidesTextAt(1, 0, { x: 0, y: 20, width: 200, height: 20 }, { zoom: 1, dpr: 1 })).toBe(true)
      batch.dispose()
    }, { extra: [BAR_SHOWVALUE_EVIDENCE] })
  })
  it('without it the text draws', async () => {
    await withBars(async () => {
      const sheet = barSheet([10, 40, 70], { def: { showValue: false } })
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      for (const view of ['static', 'viewport'] as const) {
        const got = await texts(sheet, view, batch)
        expect(got.bars, view).toBe(2)
        expect(got.texts, view).toEqual(['10', '40', '70'])
      }
      batch.dispose()
    })
  })
  it('an unpainted (withheld) bar never suppresses text', async () => {
    // A gradient bar evaluates under injected bar:render-gradient-stops but has no measured stops, so it never paints.
    for (const [def, options] of [
      [{ showValue: false, gradient: true }, { extra: [BAR_SHOWVALUE_EVIDENCE, 'bar:render-gradient-stops'] }],
      [{ showValue: false }, { extra: [BAR_SHOWVALUE_EVIDENCE], paint: false }],
    ] as const) {
      await withBars(async () => {
        const sheet = barSheet([10, 40, 70], { def })
        const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
        expect(batch.barAt(1, 0)).toBeUndefined()
        for (const view of ['static', 'viewport'] as const) {
          const got = await texts(sheet, view, batch)
          expect(got.bars, view).toBe(0)
          expect(got.texts, view).toEqual(['10', '40', '70'])
        }
        batch.dispose()
      }, options)
    }
  })
})

describe('the axis is Excel\'s dotted column (fix round 1)', () => {
  const axisBar = bar({ axisFraction: 1 / 3, startFraction: 1 / 3, endFraction: 2 / 3, axisColor: 0xff000000 })
  /** The colour each device row of the axis column paints: the gap underlay, overpainted by the dashes. */
  const column = (p: NonNullable<ReturnType<typeof layoutDataBar>>): Map<number, number> => {
    const rows = new Map<number, number>()
    for (const r of [p.axis!, ...p.axisDashes!]) for (let y = r.y; y < r.y + r.h; y++) rows.set(y, r.argb)
    return rows
  }
  it('3 on, 1 off, the gap starting where y ≡ x (mod 4) at the capture scale, white in the gap', () => {
    for (const [left, top] of [[84, 30], [264, 60], [444, 150], [85, 31]] as const) {
      const p = layoutDataBar(axisBar, ptCell(left, top, 180, 30), AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)!
      expect(p.axis!.w).toBe(1)
      const x = p.axis!.x, rows = column(p)
      expect([...rows.keys()].sort((a, b) => a - b)).toEqual(Array.from({ length: 30 }, (_, i) => top + i))
      for (const [y, argb] of rows) expect(argb, `x ${x} y ${y}`).toBe((((y - x) % 4) + 4) % 4 === 0 ? 0xffffffff : 0xff000000)
    }
  })
  it('scales the dash with the point geometry and snaps it to device px (OD-8)', () => {
    // 2 device px per pt: an 8 px period with 2 px gaps.
    const p = layoutDataBar(axisBar, ptCell(84, 30, 180, 30), { zoom: 1.5, dpr: 1, output: 'device' }, CF_BAR_RENDER_MODEL, ALL)!
    expect(p.axis!.w).toBe(2)
    const gaps = [...column(p)].filter(([, argb]) => argb === 0xffffffff).map(([y]) => y)
    expect(gaps.length).toBeGreaterThanOrEqual(14) // 7.5 periods in a 60 px host: 7 or 8 two-pixel gaps
    for (const y of gaps) expect((((y - p.axis!.x) % 8) + 8) % 8 < 2, `gap ${y}`).toBe(true)
  })
  it('a non-white axis over a filled host withholds the bar; a white host or the white axis paints', () => {
    const host = (fill?: string) => ({ ...ptCell(84, 30, 180, 30), ...(fill ? { fill } : {}) })
    expect(layoutDataBar(axisBar, host('#FFFF00'), AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)).toBeUndefined()
    expect(layoutDataBar(axisBar, host('#FFFFFF'), AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)?.axisDashes?.length).toBeGreaterThan(0)
    expect(layoutDataBar(axisBar, host(), AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)?.axisDashes?.length).toBeGreaterThan(0)
    expect(layoutDataBar({ ...axisBar, axisColor: 0xffffffff }, host('#A6A6A6'), AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)?.axis).toBeDefined()
    // No axis drawn: the host fill does not matter.
    expect(layoutDataBar(bar(), host('#FFFF00'), AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL)?.fill).toBeDefined()
  })
})

describe('text is hidden only under a bar that paints (S3-O, fix round 1)', () => {
  it('a bar withheld by the unmeasured minimum visible width keeps its text, at every scale', async () => {
    await withBars(async () => {
      // 0.001 of [0, 1000]: a non-zero share that rounds to no pixel at the capture scale (withheld at every zoom).
      const sheet = barSheet([0, 0.001, 1000], { def: { showValue: false } })
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      const metrics = computeMetrics(sheet)
      for (const s of [0.75, 1, 2]) {
        const r = recorder(metrics.widthPx * s, metrics.heightPx * s)
        r.ctx.scale(s, s)
        renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: batch })
        // 0 (a zero-length bar, laid out) and 1000 hide their text; the withheld 0.001 shows it.
        expect(r.ops.filter(o => o.op === 'text').map(o => o.text), `scale ${s}`).toEqual(['0.001'])
        expect(batch.hidesTextAt(1, 0, { x: 0, y: metrics.rowHeightsPx[0]!, width: 200, height: metrics.rowHeightsPx[1]! }, { zoom: s, dpr: 1 })).toBe(false)
        expect(batch.hidesTextAt(2, 0, { x: 0, y: 0, width: 200, height: 20 }, { zoom: s, dpr: 1 })).toBe(true)
      }
      const viewport = recorder(metrics.widthPx, metrics.heightPx)
      renderSheetViewport(sheet, viewport.ctx, buildSheetGeometry(sheet), { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
      expect(viewport.ops.filter(o => o.op === 'text').map(o => o.text)).toEqual(['0.001'])
      batch.dispose()
    }, { extra: [BAR_SHOWVALUE_EVIDENCE] })
  })
  it('a non-white axis over a filled host withholds the bar and keeps the text', async () => {
    await withBars(async () => {
      const sheet = barSheet([-50, 50, 100], { def: { showValue: false } })
      for (const row of sheet.rows) for (const cell of row.cells) cell.style = { numFmtId: 0, fillColor: 'FFFFFF00' }
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      const metrics = computeMetrics(sheet)
      const r = recorder(metrics.widthPx, metrics.heightPx)
      renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: batch })
      expect(r.ops.filter(o => o.op === 'rect' && (o.style === '#0000ff' || o.style === '#000000'))).toEqual([])
      expect(r.ops.filter(o => o.op === 'text').map(o => o.text)).toEqual(['-50', '50', '100'])
      batch.dispose()
    }, { extra: [BAR_SHOWVALUE_EVIDENCE] })
  })
})

describe('fix round 2: host fills parse strictly; text hides only under a painted fill or a zero-length bar', () => {
  const axisBar = bar({ axisFraction: 1 / 3, startFraction: 1 / 3, endFraction: 2 / 3, axisColor: 0xff000000 })
  it('any fill that is not exactly opaque #FFFFFF counts as filled, so a black axis over it withholds', () => {
    const host = (fill?: string) => ({ ...ptCell(84, 30, 180, 30), ...(fill !== undefined ? { fill } : {}) })
    for (const white of [undefined, '#FFFFFF', '#ffffff']) expect(layoutDataBar(axisBar, host(white), AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL), String(white)).toBeDefined()
    for (const other of ['ffffff', ' #FFFFFF', '#FFFFFF80', '#ffffff00', 'white', 'transparent', 'rgba(255,255,255,1)', '#fff', 'FFFFFFFF', 'not-a-colour', ''])
      expect(layoutDataBar(axisBar, host(other), AT_CAPTURE, CF_BAR_RENDER_MODEL, ALL), other).toBeUndefined()
  })
  it('a non-zero share that rounds away at the render scale shows its text; a zero-length bar hides it at every scale', async () => {
    await withBars(async () => {
      // 5 of [0, 1000] in a 200 px column: about 0.7 pt at the capture scale (one pixel, painted), nothing at 1/4 zoom.
      const sheet = barSheet([0, 5, 1000], { def: { showValue: false } })
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      const metrics = computeMetrics(sheet)
      const texts = (s: number) => {
        const r = recorder(metrics.widthPx * s, metrics.heightPx * s)
        r.ctx.scale(s, s)
        renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: batch })
        return r.ops.filter(o => o.op === 'text').map(o => o.text)
      }
      expect(texts(1)).toEqual([])
      expect(texts(0.75)).toEqual([])
      expect(texts(0.25)).toEqual(['5'])
      batch.dispose()
    }, { extra: [BAR_SHOWVALUE_EVIDENCE] })
  })
})
