// @vitest-environment node
/**
 * Task 29: icon sets paint in every view through one lazy painter (`src/xlsx/cf/iconPaint.ts`), reached by the eager
 * renderers only through `CfPaintBatch.iconAt`/`paintIcons`/`textInsetAt`/`hidesTextAt`.
 *
 * Nothing Stage 4 is captured and no glyph table exists yet (Ruling S4-ORDER), so production withholds every icon.
 * Every icon painted here is painted under INJECTED stubs (tests/helpers/cf-icons.ts): exact evidence, a hypothetical
 * render model and a stub glyph table. Each "withholds" test carries a painted control under the same stubs, so it
 * cannot pass vacuously.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { computeMetrics, computePrintMetrics, layoutCellBody, renderPrintPage, renderSheet, textMeasurer } from '../src/xlsx/render'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { createCellLayoutHost, renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache, withTextInset, type CellLayoutHost, type PlacedCell } from '../src/xlsx/view/textLayout'
import { prepareCfPaint, residentSource, type CfPaintBatch } from '../src/xlsx/cf/render'
import { cfRowPtOfOverrides, prepareCfTextScope, prepareCfWireScope, type CfFrameHostSource } from '../src/xlsx/cf/text'
import { buildGeometry } from '../src/xlsx/view/source'
import { cfNonUniformFontRows, cfRowUniformOfLayout } from '../src/xlsx/rowFont'
import { SHEET_ZOOM_STEPS } from '../src/xlsx/view/sheetZoom'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { CF_ICON_PAINT_CAPABILITIES, CF_ICON_RENDER_MODEL, CF_ICON_SCALES, iconClearsRightTextAt, CF_ICON_TEXT_PADDING_PX, iconWithheld, layoutIcon } from '../src/xlsx/cf/iconPaint'
import { admitIconGlyphs, cfIconGlyphHeader, iconGlyph, injectCfIconGlyphs, type CfGlyphOp } from '../src/xlsx/cf/iconGlyphs'
import { getPaintables } from '../src/render/paint'
import type { CfIconDescriptor, CfStamp } from '../src/xlsx/cf/model'
import type { XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'
import type { EffectiveFormatScope } from '../src/xlsx/view/formatWorkspace'
import { ICON_COLOURS, ICON_PAINT_EVIDENCE, ICON_SHOWVALUE_EVIDENCE, STUB_INSET_PX, STUB_MODEL, iconRecorder, iconSheet, stubGlyph, withIcons } from './helpers/cf-icons'
import { barRule } from './helpers/cf-bars'

const REFS = join(resolve(__dirname, '..'), 'tests/fixtures/xlsx-cf/refs')
const ALL = (): boolean => true
const AT_CAPTURE = { zoom: 0.75, dpr: 1, output: 'device' as const }
const AT_100 = { zoom: 1, dpr: 1 }
const recorder = (w: number, h: number) => iconRecorder(w, h, createCanvas as never)
const layoutContext = () => createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
const icon = (overrides: Partial<CfIconDescriptor> = {}): CfIconDescriptor => ({ ruleId: 1, set: '3Arrows', iconId: 1, bucket: 1, buckets: 3, showValue: true, ...overrides })
/** A 200 × 20 px host holding a short right-aligned number. */
const host = (overrides: Record<string, unknown> = {}) => ({ x: 0, y: 0, width: 200, height: 20, textWidth: 12, cell: { value: 50 }, ...overrides })
const LEFT: XlsxCellStyle = { numFmtId: 0, horizontal: 'left' }
const isIcon = (o: { op: string; style: string }) => o.op === 'icon' && (ICON_COLOURS as readonly string[]).includes(o.style)

/** The paged main thread's batch, built as InteractiveSheetDoc does (wire → structured clone → wire scope). */
async function pagedScope(sheet: XlsxSheet, rows: number, cols = 1, hosts?: CfFrameHostSource, locale = 'en-US'): Promise<EffectiveFormatScope> {
  const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId ?? sheet.name
  const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
  const memory = createCfMemory({})
  const rects = [{ r0: 0, r1: rows - 1, c0: 0, c1: cols - 1 }]
  const wire = await prepareCfRemote({ rects, stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, { budget: () => createCfWorkBudget({ now: () => 0 }) })
  expect(wire.kind, JSON.stringify(wire.reason)).toBe('ready')
  // S4-ROW-EXACT: the row heights the paged main thread holds, as InteractiveSheetDoc passes them: the worker's authored
  // row overrides (view/source.ts `buildGeometry`), read back as pt.
  const geometry = buildGeometry(sheet.rows, sheet.cols ?? [], [], { defaultRowPt: sheet.defaultRowHeightPt, nonUniformFontRows: cfNonUniformFontRows(sheet.rows, sheet.rowFontsBlocked) })
  const rowPtOf = cfRowPtOfOverrides(geometry.rowOverrides, geometry.defaultRowPt)
  // S4-ROW-UNIFORM: the rows' font uniformity travels with the layout, as the row heights do.
  const rowUniformOf = cfRowUniformOfLayout(geometry.nonUniformFontRows)
  return (await prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), rects, { budget: { reserve() {}, release() {} }, locale, rowPtOf, rowUniformOf, ...(hosts ? { hosts } : {}) }))!
}
const asBatch = (scope: EffectiveFormatScope): CfPaintBatch => ({
  stamp: { owner: 'paged', sheetId: 'x', workbookValueGen: 1, cfGen: scope.gen, viewGen: 1, structureGen: 1, clockGen: 1 }, key: 'paged', layer: scope.paintLayer!,
  barAt: (r, c) => scope.barAt?.(r, c),
  paintBars: (ctx, cellRect, visible, scale) => scope.paintBars?.(ctx, cellRect, visible, scale) ?? true,
  hidesTextAt: (r, c, h, scale) => scope.hidesTextAt?.(r, c, h, scale) ?? false,
  iconAt: (r, c) => scope.iconAt?.(r, c),
  paintIcons: (ctx, cellRect, visible, scale) => scope.paintIcons?.(ctx, cellRect, visible, scale) ?? true,
  textInsetAt: (r, c, h, scale) => scope.textInsetAt?.(r, c, h, scale),
  layoutCell: (...args) => scope.layoutCell!(...args),
  dispose: () => scope.dispose(),
})
const paint = (sheet: XlsxSheet, batch: CfPaintBatch, view: 'static' | 'viewport', options: { coarse?: boolean; active?: { row: number; col: number }; cache?: LayoutCache } = {}) => {
  const metrics = computeMetrics(sheet)
  const r = recorder(metrics.widthPx, metrics.heightPx)
  if (view === 'static') renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: batch })
  else renderSheetViewport(sheet, r.ctx, buildSheetGeometry(sheet), { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, {
    cache: options.cache ?? new LayoutCache(), ...(options.active ? { activeCell: options.active } : {}), ...(options.coarse ? { frameBudget: { records: 0, gridlines: 0 } } : {}),
  })
  const rowY = (y: number) => Math.floor(y / 20)
  return {
    icons: r.ops.filter(isIcon).map(o => ({ row: rowY(o.y), style: o.style, x: o.x, y: o.y, w: o.w, h: o.h })),
    texts: r.ops.filter(o => o.op === 'text' && !/^Zoom in /.test(o.text!) && !/^Loading/.test(o.text!)).map(o => ({ row: rowY(o.y - 1), text: o.text!, x: o.x })),
  }
}

describe('CF_ICON_RENDER_MODEL equals the measured sidecars', () => {
  it('the Stage 4 geometry and inset sidecars are captured, and the production model carries only the values they determine (S4-PROD-ALIGN)', () => {
    // The captures have landed (c63de53a) and record raw px, but no geometry capability can verify until the box
    // hypotheses exist (Task 29): every geometry claim is unverified. When one verifies, this fails and the comparison
    // against the sidecar is written here.
    const manifest = JSON.parse(readFileSync(join(REFS, '..', 'manifest.json'), 'utf8')) as { cases: Array<{ id: string; expected: { cells: Array<{ iconEvidence?: { geometry: string } }> } }> }
    for (const id of ['stage4-icon-geometry', 'stage4-icon-inset']) {
      expect(existsSync(join(REFS, `${id}.json`)), id).toBe(true)
      const sidecar = JSON.parse(readFileSync(join(REFS, `${id}.json`), 'utf8')) as { observedIcons?: Array<{ icon: unknown }> }
      expect(sidecar.observedIcons?.some(o => o.icon !== null), `${id} records measured icons`).toBe(true)
      // Fix round 7 gives these capabilities readings, so cells may be pixel-measured; the model stays empty until Task 30.
      const cells = manifest.cases.find(c => c.id === id)!.expected.cells.filter(e => e.iconEvidence !== undefined)
      expect(cells.length, id).toBeGreaterThan(0)
      for (const e of cells) expect(['pixel-measured', 'non-discriminating', 'unverified'], id).toContain(e.iconEvidence!.geometry)
    }
    // S4-PROD-ALIGN: the measured box variant and the 11 pt left inset; no box size or offsets (ink is not the box).
    // Task 28: the glyph-key frame, exact at the key geometry only (tests/xlsx-cf-icon-glyphs.test.ts).
    expect(CF_ICON_RENDER_MODEL).toEqual({ box: { size: 'font', vertical: 'cell-align' }, leftInsetPt: { 11: 12 }, rightInsetPt: { 11: 0 }, keyFrame: { fontPt: 11, fontName: 'Calibri', widthPt: 13, heightPt: 12, leftPt: 2, rows: CF_ICON_RENDER_MODEL.keyFrame!.rows } })
    expect(Object.isFrozen(CF_ICON_RENDER_MODEL)).toBe(true)
    // Under every capability, the production model lays nothing out off the key geometry (a 10 pt font, a 15 px row).
    expect(layoutIcon(icon(), host(), { sizePt: 10, vertical: 'bottom', horizontal: 'right' }, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)).toBeUndefined()
  })

  it('without a verified capability, the matching branch of layoutIcon returns undefined', () => {
    const font = { sizePt: 10, vertical: 'bottom' as const, horizontal: 'left' as const }
    const laid = layoutIcon(icon(), host(), font, AT_CAPTURE, STUB_MODEL, ALL)!
    // Control: the stub model at the capture scale: a 9 px box 2 px in, centred in the 15 px row; a 13 pt text inset.
    expect(laid.box).toEqual({ x: 2, y: 3, w: 9, h: 9 })
    expect(laid.inset.left).toBeCloseTo(STUB_INSET_PX, 12)
    for (const cap of [...CF_ICON_PAINT_CAPABILITIES, 'icon:inset-left']) {
      expect(layoutIcon(icon(), host(), font, AT_CAPTURE, STUB_MODEL, c => c !== cap), cap).toBeUndefined()
    }
    for (const [horizontal, cap] of [['right', 'icon:inset-right-aligned'], ['center', 'icon:inset-center']] as const) {
      expect(layoutIcon(icon(), host(), { ...font, horizontal }, AT_CAPTURE, STUB_MODEL, ALL), horizontal).toBeDefined()
      expect(layoutIcon(icon(), host(), { ...font, horizontal }, AT_CAPTURE, STUB_MODEL, c => c !== cap), horizontal).toBeUndefined()
    }
    // A box the host cannot hold: a narrow (8 px wide) host needs its clip measured; a short (6 px tall) row never
    // clips, because Excel shrinks the box by an unmeasured law (S4-PROD-ALIGN, tests/xlsx-cf-icon-align.test.ts).
    const narrow = host({ width: 8 }), short = host({ height: 6 })
    expect(layoutIcon(icon(), narrow, font, AT_CAPTURE, STUB_MODEL, c => c !== 'icon:geometry-narrow-clip')).toBeUndefined()
    expect(layoutIcon(icon(), narrow, font, AT_CAPTURE, STUB_MODEL, ALL)).toBeDefined()
    expect(layoutIcon(icon(), short, font, AT_CAPTURE, STUB_MODEL, ALL)).toBeUndefined()
    // Each model value the variant needs: absent, withheld.
    for (const key of ['box', 'sizePt', 'offsetsPt', 'leftInsetPt'] as const) {
      const { [key]: _drop, ...partial } = STUB_MODEL
      expect(layoutIcon(icon(), host(), font, AT_CAPTURE, partial, ALL), key).toBeUndefined()
    }
    // An id outside its family (3Arrows has 3 icons) and NoIcons never lay out.
    expect(layoutIcon(icon({ iconId: 3 }), host(), font, AT_CAPTURE, STUB_MODEL, ALL)).toBeUndefined()
    expect(layoutIcon(icon({ set: 'NoIcons', iconId: 0 }), host(), font, AT_CAPTURE, STUB_MODEL, ALL)).toBeUndefined()
  })

  it('iconWithheld: the painter withholds an unknown host, a missing glyph, an out-of-family id and an unknown text width', async () => {
    await withIcons(() => {
      expect(iconWithheld(icon(), host(), ALL)).toBe(false) // control
      expect(iconWithheld(icon(), undefined, ALL)).toBe(true)
      expect(iconWithheld(icon({ iconId: 3 }), host(), ALL)).toBe(true)
      expect(iconWithheld(icon({ set: '4Arrows', iconId: 3, buckets: 4 }), host(), ALL)).toBe(false)
      expect(iconWithheld(icon(), host({ textWidth: undefined }), ALL)).toBe(true)
      expect(iconWithheld(icon(), host({ cell: undefined }), ALL)).toBe(true)
      // Wrapped or rotated text: its inset is unmeasured.
      expect(iconWithheld(icon(), host({ cell: { value: 50, style: { numFmtId: 0, wrapText: true } } }), ALL)).toBe(true)
    })
    await withIcons(() => { expect(iconWithheld(icon(), host(), ALL)).toBe(true) }, { glyphs: () => undefined })
  })
})

/** Static yellow fills on every icon cell, so the record shows fill → (bar →) icon → content for one host. */
const yellow = (sheet: XlsxSheet): XlsxSheet => { for (const row of sheet.rows) for (const cell of row.cells) cell.style = { numFmtId: 0, fillColor: 'FFFFFF00' }; return sheet }

describe('icons paint in the fill phase in all five views', () => {
  it('fill → icon → content in renderSheet, renderSheetViewport (resident and paged), renderPrintPage at PRINT_SCALE = 2 and export; one source rect', async () => {
    await withIcons(async () => {
      const sheet = yellow(iconSheet([10, 50, 90]))
      const metrics = computeMetrics(sheet)
      const cellY = metrics.rowHeightsPx[0]!, cellW = metrics.colWidthsPx[0]!
      expect(cellW).toBe(200)
      const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      const paged = asBatch(await pagedScope(sheet, 3))
      const S = 2
      const views: Record<string, ReturnType<typeof recorder>> = {}
      { const r = views.static = recorder(metrics.widthPx * S, metrics.heightPx * S); r.ctx.scale(S, S); renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: resident }) }
      const geometry = buildSheetGeometry(sheet)
      for (const [name, batch] of [['resident', resident], ['paged', paged]] as const) {
        const r = views[name] = recorder(metrics.widthPx * S, metrics.heightPx * S); r.ctx.scale(S, S)
        renderSheetViewport(sheet, r.ctx, geometry, { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: S }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
      }
      { const print = computePrintMetrics(sheet, metrics); const r = views.print = recorder(print.paperPx.width * S, print.paperPx.height * S); r.ctx.scale(S, S); renderPrintPage(sheet, r.ctx, metrics, print, 0, 0, undefined, { conditionalFormatting: resident }) }
      const pages = await getPaintables({ sheets: [sheet] })
      try { const r = views.export = recorder(pages[0]!.spec.widthPx * S, pages[0]!.spec.heightPx * S); r.ctx.scale(S, S); pages[0]!.paint(r.ctx) } finally { pages.dispose() }
      // The resident batch and the paged scope (once it has painted) report the bucket-1 icon.
      expect(resident.iconAt(1, 0)).toMatchObject({ set: '3Arrows', iconId: 1, bucket: 1, buckets: 3 })
      expect(paged.iconAt(1, 0)).toMatchObject({ set: '3Arrows', iconId: 1, bucket: 1, buckets: 3 })
      const sourceRects: Record<string, { x: number; y: number; w: number; h: number }> = {}
      for (const [name, { ops }] of Object.entries(views)) {
        const fill = ops.filter(o => o.op === 'rect' && o.style === '#ffff00')[1]
        expect(fill, `${name}: host fill`).toBeDefined()
        const k = fill!.w / cellW, ox = fill!.x, oy = fill!.y - cellY * k
        const iconIndex = ops.findIndex(o => isIcon(o) && o.style === ICON_COLOURS[1])
        const textIndex = ops.findIndex(o => o.op === 'text' && o.text === '50')
        expect(iconIndex, `${name}: icon painted after the fill`).toBeGreaterThan(ops.indexOf(fill!))
        expect(textIndex, `${name}: text after the icon`).toBeGreaterThan(iconIndex)
        expect(ops.slice(iconIndex).some(o => o.op === 'rect' && o.style === '#ffff00'), `${name}: fills precede icons`).toBe(false)
        const o = ops[iconIndex]!
        sourceRects[name] = { x: (o.x - ox) / k, y: (o.y - oy) / k, w: o.w / k, h: o.h / k }
        expect(ops.filter(isIcon).map(i => i.style).sort(), `${name}: one icon per bucket`).toEqual([ICON_COLOURS[0], ICON_COLOURS[1], ICON_COLOURS[2]].sort())
      }
      for (const [name, r] of Object.entries(sourceRects)) for (const key of ['x', 'y', 'w', 'h'] as const) expect(r[key], `${name}.${key}`).toBeCloseTo(sourceRects.static![key], 9)
      resident.dispose(); paged.dispose()
    })
  })

  it('a bar and an icon on one cell: fill → bar → icon → content once icon:composition-with-bar settles z-order; without it no icon paints on the bar cell', async () => {
    const sheet = (): XlsxSheet => yellow(iconSheet([10, 50, 90], { extraRules: [{ ...barRule({ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }, { axisArgb: 0xffffffff }, 2), priority: 2, documentOrder: 1 }] }))
    await withIcons(async () => {
      const s = sheet()
      const batch = (await prepareCfPaint(s, { evidence: cfRuntimeEvidence, layoutContext }))!
      expect(batch.barAt(1, 0)).toBeDefined()
      const metrics = computeMetrics(s)
      const r = recorder(metrics.widthPx, metrics.heightPx)
      renderSheet(s, r.ctx, metrics, undefined, { conditionalFormatting: batch })
      const bar = r.ops.findIndex(o => o.op === 'rect' && o.style === '#0000ff')
      const ic = r.ops.findIndex(o => isIcon(o) && o.style === ICON_COLOURS[1])
      const text = r.ops.findIndex(o => o.op === 'text' && o.text === '50')
      expect(bar).toBeGreaterThan(-1)
      expect(ic).toBeGreaterThan(bar)
      expect(text).toBeGreaterThan(ic)
      batch.dispose()
    })
    // The evaluator itself withholds an icon meeting a bar without the composition capability; the painter's own
    // bar-cell gate is checked through the scope, whose icon state carries the bar beside it.
    await withIcons(async () => {
      const s = sheet()
      const batch = (await prepareCfPaint(s, { evidence: cfRuntimeEvidence, layoutContext }))!
      expect(batch.barAt(1, 0)).toBeDefined()
      expect(batch.iconAt(1, 0)).toBeUndefined()
      expect(paint(s, batch, 'static').icons).toEqual([])
      batch.dispose()
    }, { omit: ['icon:composition-with-bar'] })
  })
})

describe('inset and icon come from one predicate on one host (S4-L, Lesson 11)', () => {
  // Left-aligned numbers: the inset moves the text by 13 pt when the icon paints.
  const left = (values: readonly number[] = [10, 50, 90], options: Parameters<typeof iconSheet>[1] = {}) => iconSheet(values, { style: LEFT, ...options })
  const textX = (got: ReturnType<typeof paint>, row: number) => got.texts.find(t => t.row === row)?.x
  const unInset = async (name: string, sheet: XlsxSheet, row: number, options: Parameters<typeof withIcons>[1], view: 'static' | 'viewport' = 'static') => {
    // The un-inset position: the same sheet with no conditional formatting at all.
    const plainSheet = { ...sheet, conditionalFormatting: undefined } as XlsxSheet
    const plainX = (v: 'static' | 'viewport') => {
      const metrics = computeMetrics(plainSheet), r = recorder(metrics.widthPx, metrics.heightPx)
      if (v === 'static') renderSheet(plainSheet, r.ctx, metrics)
      else renderSheetViewport(plainSheet, r.ctx, buildSheetGeometry(plainSheet), { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, undefined, { cache: new LayoutCache() })
      return r.ops.find(o => o.op === 'text' && Math.floor((o.y - 1) / 20) === row)?.x
    }
    await withIcons(async () => {
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      for (const v of view === 'static' ? ['static', 'viewport'] as const : [view]) {
        const got = paint(sheet, batch, v)
        expect(got.icons.filter(i => i.row === row), `${name} ${v}: no icon`).toEqual([])
        expect(plainX(v), `${name} ${v}: baseline text`).toBeDefined()
        expect(textX(got, row), `${name} ${v}: text visible, un-inset`).toBeCloseTo(plainX(v)!, 6)
      }
      expect(batch.iconAt(row, 0), name).toBeUndefined()
      batch.dispose()
    }, options)
  }

  it('control: under the full stubs every cell is inset by 13 pt and paints its icon, in both views', async () => {
    await withIcons(async () => {
      const sheet = left()
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      for (const view of ['static', 'viewport'] as const) {
        const got = paint(sheet, batch, view)
        expect(got.icons.map(i => i.row), view).toEqual([0, 1, 2])
        for (const row of [0, 1, 2]) expect(textX(got, row), `${view} row ${row}`).toBeCloseTo(3 + STUB_INSET_PX, 6)
      }
      batch.dispose()
    })
  })

  it('no glyph: un-inset, visible, no icon', async () => { await unInset('no glyph', left(), 1, { glyphs: () => undefined }) })

  it('a glyph class refused at batch preparation (admitIconGlyphs undefined): un-inset, visible, no icon, for that class only', async () => {
    // Bucket 1's class has an op list too large for the batch's work budget: admission refuses it, the others paint.
    const huge = new Array<CfGlyphOp>(2_000_000).fill(['Z'])
    const glyphs = (set: Parameters<typeof stubGlyph>[0], iconId: number) => (iconId === 1 ? { ...stubGlyph(set, iconId), path: huge } : stubGlyph(set, iconId))
    await withIcons(() => {
      expect(admitIconGlyphs(new Set(['3Arrows:1']), createCfMemory({}), createCfWorkBudget({}), () => glyphs('3Arrows', 1))).toBeUndefined()
      expect(admitIconGlyphs(new Set(['3Arrows:0']), createCfMemory({}), createCfWorkBudget({}), () => glyphs('3Arrows', 0))).toBeDefined()
    })
    await unInset('refused class', left(), 1, { glyphs })
    await withIcons(async () => {
      const sheet = left()
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      expect(paint(sheet, batch, 'static').icons.map(i => i.row)).toEqual([0, 2])
      batch.dispose()
    }, { glyphs })
  })

  it('a NoIcons bucket: un-inset, visible, no icon', async () => {
    const custom = [{ set: '3Arrows' as const, iconId: 0 }, { set: 'NoIcons' as const, iconId: 0 }, { set: '3Arrows' as const, iconId: 2 }]
    await unInset('NoIcons', left([10, 50, 90], { def: { custom } }), 1, {})
  })

  it('a narrow cell without icon:geometry-narrow-clip: un-inset, visible, no icon', async () => {
    // 10 px: narrower than the 2 pt offset plus the 9 pt box.
    await unInset('narrow', left([1, 5, 9], { widthChars: 10 / 7, def: { thresholds: [{ bound: { type: 'num', value: 0 }, gte: true }, { bound: { type: 'num', value: 4 }, gte: true }, { bound: { type: 'num', value: 7 }, gte: true }] } }), 1, {})
  })

  it('a coarse frame: the retained active cell alone is inset and painted', async () => {
    await withIcons(async () => {
      const sheet = left()
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      const got = paint(sheet, batch, 'viewport', { coarse: true, active: { row: 1, col: 0 } })
      expect(got.icons.map(i => i.row)).toEqual([1])
      expect(got.texts.map(t => t.row)).toEqual([1])
      expect(textX(got, 1)).toBeCloseTo(3 + STUB_INSET_PX, 6)
      batch.dispose()
    })
  })

  it('the inset is decided on the host, not from iconAt alone (mutant: an inset from iconAt would inset this host)', async () => {
    await withIcons(async () => {
      const sheet = left()
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      expect(batch.iconAt(1, 0)).toBeDefined()
      const cell = sheet.rows[1]!.cells[0]!
      expect(batch.textInsetAt(1, 0, { x: 0, y: 20, width: 200, height: 20, textWidth: 12, cell }, AT_100)?.left).toBeCloseTo(STUB_INSET_PX, 9)
      // The same cell in a 10 px host (another frame's geometry): no inset, whatever iconAt says.
      expect(batch.textInsetAt(1, 0, { x: 0, y: 20, width: 10, height: 20, textWidth: 12, cell }, AT_100)).toBeUndefined()
      expect(batch.textInsetAt(1, 0, { x: 0, y: 20, width: 200, height: 20, cell }, AT_100), 'unknown text width').toBeUndefined()
      batch.dispose()
    })
  })
})

describe('inset overflow decides identically on every leg', () => {
  it('a left-aligned number that fits only without the inset gets no icon on the resident, worker and forced-paged legs; an unknown width withholds', async () => {
    await withIcons(async () => {
      // The number's un-inset width, as the static renderer lays it out.
      const probe = createCellLayoutHost(layoutContext())
      const body = probe.layoutBody({ ref: 'A1', row: 0, col: 0, value: 123456, styleIndex: 0, style: LEFT }, 500, 20)
      const width = body.text?.kind === 'layout' ? body.text.layout.lines[0]!.segments.reduce((n, s) => n + s.width, 0) : NaN
      expect(width).toBeGreaterThan(20)
      // A column the number fits without the inset (w − 6 ≥ width) but not with it (w − 6 − 17.3 < width).
      const widthPx = Math.ceil(width + 6 + STUB_INSET_PX / 2)
      for (const [name, colPx, want] of [['tight', widthPx, false], ['control', 300, true]] as const) {
        const sheet = iconSheet([123456], { style: LEFT, widthChars: colPx / 7, def: { thresholds: [{ bound: { type: 'num', value: 0 }, gte: true }, { bound: { type: 'num', value: 40 }, gte: true }, { bound: { type: 'num', value: 70 }, gte: true }] } })
        expect(computeMetrics(sheet).colWidthsPx[0], name).toBe(colPx)
        const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
        const text = (await prepareCfTextScope(sheet, { rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], signal: new AbortController().signal, budget: { reserve() {}, release() {} }, layoutContext }))!
        // Worker leg: the host-less paged paint scope learns from the frame it paints.
        const worker = asBatch(await pagedScope(sheet, 1))
        const got = paint(sheet, worker, 'viewport')
        // Forced-paged semantic leg: the frame's own laid-out record supplies textWidth.
        const metrics = computeMetrics(sheet)
        const record = createCellLayoutHost(layoutContext()).layoutBody(sheet.rows[0]!.cells[0]!, metrics.colWidthsPx[0]!, metrics.rowHeightsPx[0]!)
        const recordWidth = record.text?.kind === 'layout' ? record.text.layout.lines[0]!.segments.reduce((n, s) => n + s.width, 0) : undefined
        const semantic = await pagedScope(sheet, 1, 1, { cell: (row, col) => ({ box: { x: 0, y: 0, width: metrics.colWidthsPx[0]!, height: metrics.rowHeightsPx[0]! }, base: LEFT, textWidth: recordWidth, cell: sheet.rows[row]!.cells[col]! }) })
        const blind = await pagedScope(sheet, 1, 1, { cell: (row, col) => ({ box: { x: 0, y: 0, width: metrics.colWidthsPx[0]!, height: metrics.rowHeightsPx[0]! }, base: LEFT, cell: sheet.rows[row]!.cells[col]! }) })
        expect(Boolean(resident.iconAt(0, 0)), `${name}: resident`).toBe(want)
        expect(Boolean(text.iconAt?.(0, 0)), `${name}: resident text scope`).toBe(want)
        expect(Boolean(worker.iconAt(0, 0)), `${name}: worker`).toBe(want)
        expect(got.icons.length > 0, `${name}: worker frame paints`).toBe(want)
        expect(Boolean(semantic.iconAt?.(0, 0)), `${name}: forced-paged semantic`).toBe(want)
        expect(blind.iconAt?.(0, 0), `${name}: a source without textWidth withholds`).toBeUndefined()
        const label = (s: EffectiveFormatScope | undefined) => s?.describe(0, 0)?.text
        expect(label(semantic), name).toBe(label(text))
        for (const s of [text, semantic, blind]) s.dispose()
        resident.dispose(); worker.dispose()
      }
    })
  })
})

describe('inset enters the layout key and the ### width', () => {
  const cellOf = (value: number, style: XlsxCellStyle = LEFT) => ({ ref: 'A1', row: 0, col: 0, value, styleIndex: 0, style })
  it('layoutKey of two otherwise equal placements differs when textInset differs', () => {
    const cache = new LayoutCache()
    const ctx = layoutContext()
    const layoutHost: CellLayoutHost = createCellLayoutHost(ctx)
    const placed = (textInset?: { left: number; right: number }): PlacedCell => ({ cell: cellOf(5), source: { row: 0, col: 0 }, x: 0, y: 0, width: 200, height: 20, ...(textInset ? { textInset } : {}) })
    const owner = { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    cache.layout(owner, placed(), layoutHost)
    cache.layout(owner, placed(), layoutHost)
    expect(cache.stats).toMatchObject({ built: 1, reused: 1 })
    const inset = cache.layout(owner, placed({ left: 20, right: 0 }), layoutHost)
    expect(cache.stats.built, 'an inset placement misses').toBe(2)
    const plain = cache.layout(owner, placed(), layoutHost)
    expect(cache.stats.built, 'and the un-inset one is still its own entry').toBe(2)
    expect(inset).not.toBe(plain)
  })

  it('a left-aligned number in a 40 px cell with a 20 px inset is laid out from x = 20 (the hypothesis run); one that no longer fits shows ###', () => {
    const ctx = layoutContext(), measured = textMeasurer(ctx)
    const sheet = { name: 's', rows: [], cols: [], merges: [], mergeRanges: [] } as unknown as XlsxSheet
    const plain = layoutCellBody(ctx, sheet, cellOf(1), 0, 0, 40, 20, measured)
    const inset = withTextInset(0, 40, { left: 20, right: 0 }, (x, w) => layoutCellBody(ctx, sheet, cellOf(1), x, 0, w, 20, measured))
    const x = (b: typeof plain) => (b.text?.kind === 'layout' ? b.text.layout.lines[0]!.segments[0]!.x + (b.text.clip.x) : b.text?.kind === 'plain' ? b.text.tx : NaN)
    expect(plain.formattedText).toBe('1')
    expect(inset.formattedText).toBe('1')
    expect(x(inset) - x(plain)).toBeCloseTo(20, 9)
    // The fill and box stay the cell's own.
    expect([inset.x, inset.w]).toEqual([0, 40])
    // A number that fits 40 − 6 px un-inset, not the 14 px the inset leaves: the ### width is the inset width.
    const font = plain.text!.font
    const value = [12345, 1234, 123, 12].find(v => measured(font, String(v)) <= 34 && measured(font, String(v)) > 14)!
    expect(value, 'a number between 14 and 34 px wide').toBeDefined()
    const fits = layoutCellBody(ctx, sheet, cellOf(value), 0, 0, 40, 20, measured)
    expect(fits.formattedText).toBe(String(value))
    const hashed = withTextInset(0, 40, { left: 20, right: 0 }, (x, w) => layoutCellBody(ctx, sheet, cellOf(value), x, 0, w, 20, measured))
    expect(hashed.formattedText).toMatch(/^#+$/)
    expect(measured(hashed.text!.font, hashed.formattedText)).toBeLessThanOrEqual(14)
  })

  it('toggling a rule from no icon to icon within one session misses the layout cache, never a stale hit', async () => {
    const cache = new LayoutCache()
    const sheet = iconSheet([10, 50, 90], { style: LEFT })
    let off: ReturnType<typeof paint> | undefined, on: ReturnType<typeof paint> | undefined
    await withIcons(async () => {
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      off = paint(sheet, batch, 'viewport', { cache })
      batch.dispose()
    }, { glyphs: () => undefined })
    const built = cache.stats.built
    await withIcons(async () => {
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      on = paint(sheet, batch, 'viewport', { cache })
      batch.dispose()
    })
    expect(off!.icons).toEqual([])
    expect(on!.icons.map(i => i.row)).toEqual([0, 1, 2])
    expect(cache.stats.built - built, 'each icon cell lays out again, inset').toBeGreaterThanOrEqual(3)
    for (const row of [0, 1, 2]) {
      expect(off!.texts.find(t => t.row === row)!.x).toBeCloseTo(3, 6)
      expect(on!.texts.find(t => t.row === row)!.x).toBeCloseTo(3 + STUB_INSET_PX, 6)
    }
  })
})

describe('showValue=0 hides canvas text only where an icon paints (S4-L)', () => {
  const run = async (options: Parameters<typeof withIcons>[1], sheet = () => iconSheet([10, 50, 90], { def: { showValue: false } })) => withIcons(async () => {
    const s = sheet()
    const batch = (await prepareCfPaint(s, { evidence: cfRuntimeEvidence, layoutContext }))!
    const out = { static: paint(s, batch, 'static'), viewport: paint(s, batch, 'viewport') }
    batch.dispose()
    return out
  }, options)
  it('with icon:showvalue-canvas-suppression the content phase draws no glyph run, and the icon paints', async () => {
    const got = await run({ extra: [ICON_SHOWVALUE_EVIDENCE] })
    for (const view of ['static', 'viewport'] as const) {
      expect(got[view].icons.map(i => i.row), view).toEqual([0, 1, 2])
      expect(got[view].texts, view).toEqual([])
    }
  })
  it('without it the text draws', async () => {
    const got = await run({})
    for (const view of ['static', 'viewport'] as const) {
      expect(got[view].icons.length, view).toBe(3)
      expect(got[view].texts.map(t => t.text), view).toEqual(['10', '50', '90'])
    }
  })
  it('a withheld icon never suppresses text', async () => {
    const got = await run({ extra: [ICON_SHOWVALUE_EVIDENCE], glyphs: () => undefined })
    for (const view of ['static', 'viewport'] as const) {
      expect(got[view].icons, view).toEqual([])
      expect(got[view].texts.map(t => t.text), view).toEqual(['10', '50', '90'])
    }
  })
  it('a bar and an icon with conflicting showValue keep the text until icon:composition-bar-showvalue is verified', async () => {
    const sheet = () => iconSheet([10, 50, 90], { def: { showValue: false }, extraRules: [{ ...barRule({ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }, { axisArgb: 0xffffffff, showValue: true }, 2), priority: 2, documentOrder: 1 }] })
    const got = await run({ extra: [ICON_SHOWVALUE_EVIDENCE, 'bar:showvalue-canvas-suppression'] }, sheet)
    for (const view of ['static', 'viewport'] as const) {
      expect(got[view].icons.length, `${view}: the icons do paint`).toBe(3)
      expect(got[view].texts.map(t => t.text), view).toEqual(['10', '50', '90'])
    }
  })
})

describe('tile and full render carry identical icon primitives at six scales', () => {
  /** Column A is 250 px, so column B's icon box (2 pt in, 9 pt wide) straddles the x = 256 tile edge. */
  const seamSheet = (): XlsxSheet => { const s = iconSheet([10, 50, 90], { col: 1 }); s.cols = [{ min: 0, max: 0, widthChars: 250 / 7 }, { min: 1, max: 1, widthChars: 200 / 7 }]; return s }
  const tiled = (sheet: XlsxSheet, batchFor: (tx: number) => CfPaintBatch | undefined, zoom: number, dpr: number, tile = 256) => {
    const metrics = computeMetrics(sheet), geometry = buildSheetGeometry(sheet), cache = new LayoutCache()
    const out = new Map<number, Array<{ style: string; x: number; y: number; w: number; h: number }>>()
    for (let tx = 0; tx < metrics.widthPx; tx += tile) {
      const r = recorder(tile * zoom * dpr, tile * zoom * dpr)
      r.ctx.scale(zoom * dpr, zoom * dpr)
      const batch = batchFor(tx)
      renderSheetViewport(sheet, r.ctx, geometry, { box: { x: tx, y: 0, width: tile, height: tile }, zoom, deviceScale: dpr }, batch ? { conditionalFormatting: batch } : undefined, { cache })
      out.set(tx, r.ops.filter(isIcon).map(o => ({ style: o.style, x: o.x + tx * zoom * dpr, y: o.y, w: o.w, h: o.h })))
    }
    return out
  }
  it('an icon cell straddling a tile edge has identical device-snapped rects in the full and tiled renders', async () => {
    await withIcons(async () => {
      const sheet = seamSheet()
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      let checked = 0
      for (const zoom of [0.5, 1, 1.75]) for (const dpr of [1, 2]) {
        const s = zoom * dpr, metrics = computeMetrics(sheet)
        const full = recorder(metrics.widthPx * s, metrics.heightPx * s)
        full.ctx.scale(s, s)
        renderSheet(sheet, full.ctx, metrics, undefined, { conditionalFormatting: batch })
        const fullIcons = full.ops.filter(isIcon).map(({ style, x, y, w, h }) => ({ style, x, y, w, h }))
        expect(fullIcons, `zoom ${zoom} dpr ${dpr}`).toHaveLength(3)
        for (const r of fullIcons) for (const v of [r.x, r.y, r.w, r.h]) expect(Math.abs(v - Math.round(v)), `snapped at zoom ${zoom} dpr ${dpr}`).toBeLessThan(1e-6)
        expect(fullIcons.every(r => r.x / s < 256 && (r.x + r.w) / s > 256), 'each box straddles the tile edge').toBe(true)
        const tiles = tiled(sheet, () => batch, zoom, dpr)
        for (const tx of [0, 256]) {
          const rects = tiles.get(tx)!
          expect(rects, `tile ${tx} at zoom ${zoom} dpr ${dpr}`).toHaveLength(3)
          for (const r of rects) expect(fullIcons.some(f => f.style === r.style && (['x', 'y', 'w', 'h'] as const).every(k => Math.abs(f[k] - r[k]) < 1e-6)), `tile ${tx} rect ${JSON.stringify(r)}`).toBe(true)
        }
        checked++
      }
      expect(checked).toBe(6)
      batch.dispose()
    })
  })
  it('RED mutant proof: dropping icons on non-first tiles fails the seam check', async () => {
    await withIcons(async () => {
      const sheet = seamSheet()
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      const tiles = tiled(sheet, tx => (tx === 0 ? batch : undefined), 1, 1)
      expect(tiles.get(0)!.length).toBe(3)
      expect(tiles.get(256)!, 'the mutant seam tile paints no icon, so the parity check above would fail').toHaveLength(0)
      batch.dispose()
    })
  })
})

describe('CF icons never expand static, print or export extents', () => {
  it('a whole-column icon set on A:A leaves the print canvas and export dimensions at the data extent', async () => {
    await withIcons(async () => {
      const dataRows = 10
      const sheet = iconSheet(Array.from({ length: dataRows }, (_, r) => (r + 1) * 10))
      ;(sheet.conditionalFormatting!.rules[0]!.ranges as { firstRow: number; lastRow: number; firstCol: number; lastCol: number }[]).splice(0, 1, { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 })
      const m = computeMetrics(sheet)
      expect(m.rowHeightsPx.length).toBe(dataRows)
      const print = computePrintMetrics(sheet, m)
      expect(print.pagesWide * print.pagesTall).toBe(1)
      // A positive control: the icons do paint, yet nothing grows.
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      expect(batch.iconAt(dataRows - 1, 0)).toBeDefined()
      expect(batch.iconAt(dataRows + 5, 0)).toBeUndefined()
      expect(paint(sheet, batch, 'static').icons.length).toBe(dataRows)
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

describe('production withholds every icon off the admitted geometry (S4-ORDER, Task 30)', () => {
  it('with the published glyph table and every capability, a host without the key font or a measured row paints nothing and is not inset', async () => {
    // Task 30 published the glyph (OD-9(b)); the host (no font name, no row height) is not the admitted geometry.
    expect(iconGlyph('3Arrows', 0)).toBeDefined()
    await withIcons(async () => {
      const sheet = iconSheet([10, 50, 90], { style: LEFT })
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      for (const row of [0, 1, 2]) expect(batch.iconAt(row, 0)).toBeUndefined()
      const got = paint(sheet, batch, 'viewport')
      expect(got.icons).toEqual([])
      expect(got.texts.map(t => t.x)).toEqual([3, 3, 3])
      batch.dispose()
    }, { glyphs: false, extra: [...ICON_PAINT_EVIDENCE] })
  })
})

describe('the eager renderer reaches icon layout only through the batch (Task 29 fix round 1)', () => {
  it('render.ts holds no icon host, text-width or inset code: the static leg calls CfPaintBatch.layoutCell', async () => {
    const { readFileSync } = await import('node:fs')
    const src = readFileSync('src/xlsx/render.ts', 'utf8')
    expect(src).not.toMatch(/cellTextWidth|textInsetAt|inset\.left|export const PADDING_[LR]/)
    expect(src).toMatch(/cf\.layoutCell\(/)
  })
  it('the lazy text padding equals the static layout padding (plain left text starts 3 px in, plain numbers end 3 px short)', () => {
    const ctx = layoutContext(), measured = textMeasurer(ctx)
    const sheet = { name: 's', rows: [], cols: [], merges: [], mergeRanges: [] } as unknown as XlsxSheet
    const text = layoutCellBody(ctx, sheet, { ref: 'A1', row: 0, col: 0, value: 'ab', styleIndex: 0 }, 0, 0, 100, 20, measured)
    const num = layoutCellBody(ctx, sheet, { ref: 'A1', row: 0, col: 0, value: 5, styleIndex: 0 }, 0, 0, 100, 20, measured)
    expect(text.text?.kind === 'plain' && text.text.tx).toBe(3)
    expect(num.text?.kind === 'plain' && num.text.tx + num.text.width).toBeCloseTo(97, 9)
    expect(CF_ICON_TEXT_PADDING_PX).toBe(6)
  })
})

/**
 * Task 28 round 2, Ruling S4-RIGHTINSET-ZERO: `icon:inset-right-aligned` reads `unshifted` (§12), so a right-aligned
 * number keeps its text box (inset 0) wherever it fits right of the icon frame and its measured gap (the left inset);
 * one that reaches the frame would take the overflow path, unmeasured for right-aligned text: withheld. Production
 * model and production glyph table, 11 pt Calibri on admitted 15 pt rows, centred (round 4: S4-ROW-EXACT).
 */
describe('S4-RIGHTINSET-ZERO: right-aligned numbers at the key geometry', () => {
  const KEY_STYLE: XlsxCellStyle = { numFmtId: 0, fontSizePt: 11, fontName: 'Calibri' }
  const GLYPH_CAPS = ['icon:glyph-3Arrows-0', 'icon:glyph-3Arrows-1', 'icon:glyph-3Arrows-2', 'icon:geometry-row-offset']
  const keySheet = (values: number[], widthChars?: number): XlsxSheet => {
    // S4-ROW-EXACT (round 4): the key's own 20 pt row withholds; right-aligned numbers are measured centred at 15 pt.
    const sheet = iconSheet(values, { style: { ...KEY_STYLE, vertical: 'center' }, ...(widthChars !== undefined ? { widthChars } : {}) })
    for (const row of sheet.rows) { row.heightPt = 15; row.customHeight = true }
    return sheet
  }
  const production = { model: false as const, glyphs: false as const }
  async function views(sheet: XlsxSheet) {
    const metrics = computeMetrics(sheet)
    const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
    const paged = asBatch(await pagedScope(sheet, sheet.rows.length))
    const out: Record<string, ReturnType<typeof recorder>> = {}
    { const r = out.static = recorder(metrics.widthPx, metrics.heightPx); renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: resident }) }
    const geometry = buildSheetGeometry(sheet)
    for (const [name, batch] of [['resident', resident], ['paged', paged]] as const) {
      const r = out[name] = recorder(metrics.widthPx, metrics.heightPx)
      renderSheetViewport(sheet, r.ctx, geometry, { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
    }
    { const print = computePrintMetrics(sheet, metrics); const r = out.print = recorder(print.paperPx.width, print.paperPx.height); renderPrintPage(sheet, r.ctx, metrics, print, 0, 0, undefined, { conditionalFormatting: resident }) }
    const pages = await getPaintables({ sheets: [sheet] })
    try { const r = out.export = recorder(pages[0]!.spec.widthPx, pages[0]!.spec.heightPx); pages[0]!.paint(r.ctx) } finally { pages.dispose() }
    const result = Object.fromEntries(Object.entries(out).map(([name, r]) => [name, {
      icons: r.ops.filter(o => o.op === 'icon').length,
      texts: r.ops.filter(o => o.op === 'text' && /^\d+$/.test(o.text!)).map(o => ({ text: o.text!, x: o.x })),
    }]))
    const iconAt = [resident.iconAt(0, 0) !== undefined, paged.iconAt(0, 0) !== undefined]
    resident.dispose(); paged.dispose()
    return { result, iconAt, metrics }
  }

  it('Task 28 round 3: a default 15 pt row, centred (Excel 16.115\'s Normal style), paints identically in static, resident, worker, forced-paged and print', async () => {
    const sheet = iconSheet([10, 50, 90], { style: { ...KEY_STYLE, vertical: 'center' } })
    for (const row of sheet.rows) { row.heightPt = 15; row.customHeight = true }
    expect(computeMetrics(sheet).rowHeightsPx[0]).toBe(20)
    const run = async () => {
      const metrics = computeMetrics(sheet)
      const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      const worker = asBatch(await pagedScope(sheet, 3))
      const host = createCellLayoutHost(layoutContext())
      const forced = asBatch(await pagedScope(sheet, 3, 1, { cell: (row, col) => {
        const cell = sheet.rows[row]!.cells[col]!
        const body = host.layoutBody(cell, metrics.colWidthsPx[0]!, metrics.rowHeightsPx[row]!)
        return { box: { x: 0, y: metrics.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0), width: metrics.colWidthsPx[0]!, height: metrics.rowHeightsPx[row]! }, base: cell.style!, textWidth: body.text?.kind === 'plain' ? body.text.width : undefined, cell }
      } }))
      const out: Record<string, Array<{ x: number; y: number; w: number; h: number }>> = {}
      { const r = recorder(metrics.widthPx, metrics.heightPx); renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: resident }); out.static = r.ops.filter(o => o.op === 'icon') }
      const geometry = buildSheetGeometry(sheet)
      for (const [name, batch] of [['resident', resident], ['worker', worker], ['forced-paged', forced]] as const) {
        const r = recorder(metrics.widthPx, metrics.heightPx)
        renderSheetViewport(sheet, r.ctx, geometry, { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
        out[name] = r.ops.filter(o => o.op === 'icon')
      }
      { const print = computePrintMetrics(sheet, metrics); const r = recorder(print.paperPx.width, print.paperPx.height); renderPrintPage(sheet, r.ctx, metrics, print, 0, 0, undefined, { conditionalFormatting: resident }); out.print = r.ops.filter(o => o.op === 'icon') }
      resident.dispose(); worker.dispose(); forced.dispose()
      return out
    }
    const got = await withIcons(run, { ...production, extra: GLYPH_CAPS })
    for (const [name, ops] of Object.entries(got)) expect(ops.length, name).toBe(6)
    // Same frame on every leg, relative to the first icon's origin (print adds a page offset).
    const r6 = (v: number) => Math.round(v * 1e6) / 1e6
    const rel = (ops: Array<{ x: number; y: number; w: number; h: number }>) => ops.map(o => ({ x: r6(o.x - ops[0]!.x), y: r6(o.y - ops[0]!.y), w: r6(o.w), h: r6(o.h) }))
    for (const name of ['resident', 'worker', 'forced-paged', 'print']) expect(rel(got[name]!), name).toEqual(rel(got.static!))
    expect(got.static![0]!.x).toBeCloseTo(0 + 2 * 4 / 3, 0)
    // The control: an unmeasured 15.75 pt row withholds on every leg.
    for (const rowPt of [15.75, 25, 20]) {
      // Unmeasured (15.75 pt), and measured but with viewer px Excel does not draw (25 pt: 33 or 33.33 px; 20 pt).
      for (const row of sheet.rows) row.heightPt = rowPt
      const none = await withIcons(run, { ...production, extra: GLYPH_CAPS })
      for (const [name, ops] of Object.entries(none)) expect(ops.length, `${rowPt} pt: ${name}`).toBe(0)
    }
  })

  it('Task 28 round 5 (S4-DEFAULT-ROW): rows without ht on a defaultRowHeight="15" sheet paint in every view; without defaultRowHeight they withhold', async () => {
    const sheet = iconSheet([10, 50, 90], { style: { ...KEY_STYLE, vertical: 'center' } })
    for (const row of sheet.rows) { delete row.heightPt; delete row.customHeight }
    const run = async () => {
      const metrics = computeMetrics(sheet)
      const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
      const worker = asBatch(await pagedScope(sheet, 3))
      const host = createCellLayoutHost(layoutContext())
      const forced = asBatch(await pagedScope(sheet, 3, 1, { cell: (row, col) => {
        const cell = sheet.rows[row]!.cells[col]!
        const body = host.layoutBody(cell, metrics.colWidthsPx[0]!, metrics.rowHeightsPx[row]!)
        return { box: { x: 0, y: metrics.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0), width: metrics.colWidthsPx[0]!, height: metrics.rowHeightsPx[row]! }, base: cell.style!, textWidth: body.text?.kind === 'plain' ? body.text.width : undefined, cell }
      } }))
      const out: Record<string, number> = {}
      { const r = recorder(metrics.widthPx, metrics.heightPx); renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: resident }); out.static = r.ops.filter(o => o.op === 'icon').length }
      const geometry = buildSheetGeometry(sheet)
      for (const [name, batch] of [['resident', resident], ['worker', worker], ['forced-paged', forced]] as const) {
        const r = recorder(metrics.widthPx, metrics.heightPx)
        renderSheetViewport(sheet, r.ctx, geometry, { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
        out[name] = r.ops.filter(o => o.op === 'icon').length
      }
      { const print = computePrintMetrics(sheet, metrics); const r = recorder(print.paperPx.width, print.paperPx.height); renderPrintPage(sheet, r.ctx, metrics, print, 0, 0, undefined, { conditionalFormatting: resident }); out.print = r.ops.filter(o => o.op === 'icon').length }
      const describedIcons = [0, 1, 2].filter(row => resident.iconAt(row, 0) !== undefined).length
      resident.dispose(); worker.dispose(); forced.dispose()
      return { out, describedIcons }
    }
    sheet.defaultRowHeightPt = 15
    const painted = await withIcons(run, { ...production, extra: GLYPH_CAPS })
    for (const [name, n] of Object.entries(painted.out)) expect(n, name).toBe(6)
    expect(painted.describedIcons).toBe(3)
    delete sheet.defaultRowHeightPt
    const none = await withIcons(run, { ...production, extra: GLYPH_CAPS })
    for (const [name, n] of Object.entries(none.out)) expect(n, `no defaultRowHeight: ${name}`).toBe(0)
    expect(none.describedIcons).toBe(0)
  })

  it('Task 30 fix round (S4-ROW-UNIFORM): a bottom-aligned icon paints in every view on a row whose every cell is Calibri 11 regular, and withholds in every view when one cell of the row is 20 pt', async () => {
    // Excel lifts a bottom-aligned icon onto its row's shared baseline, which a larger font anywhere in the row moves
    // (stage4-icon-geometry C3/M3..O3 and stage4-icon-geometry-2 C2/G2/I2/J2: 2 px above the measured frame in rows that
    // hold a 20 pt cell). The measured bottom frame stands only on rows like stage4-icon-rows-bottom's, where every cell
    // is the default Calibri 11 regular; every other row withholds bottom placement. Centred placement is unaffected.
    const count = async (vertical: 'bottom' | 'center', rowPt: number, big = false) => {
      const sheet = iconSheet([10, 50, 90], { style: { ...KEY_STYLE, horizontal: 'left', vertical } })
      for (const row of sheet.rows) { row.heightPt = rowPt; row.customHeight = true }
      // One 20 pt cell beside the icon (column C, outside the rule and the requested rect) in every row.
      if (big) for (const row of sheet.rows) row.cells.push({ ref: `C${row.index + 1}`, row: row.index, col: 2, value: 'x', styleIndex: 0, style: { ...KEY_STYLE, fontSizePt: 20 } })
      return withIcons(async () => {
        const metrics = computeMetrics(sheet)
        const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
        const worker = asBatch(await pagedScope(sheet, 3))
        const host = createCellLayoutHost(layoutContext())
        const forced = asBatch(await pagedScope(sheet, 3, 1, { cell: (row, col) => {
          const cell = sheet.rows[row]!.cells[col]!
          const body = host.layoutBody(cell, metrics.colWidthsPx[0]!, metrics.rowHeightsPx[row]!)
          return { box: { x: 0, y: metrics.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0), width: metrics.colWidthsPx[0]!, height: metrics.rowHeightsPx[row]! }, base: cell.style!, textWidth: body.text?.kind === 'plain' ? body.text.width : undefined, cell }
        } }))
        const out: Record<string, number> = {}
        { const r = recorder(metrics.widthPx, metrics.heightPx); renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: resident }); out.static = r.ops.filter(o => o.op === 'icon').length }
        const geometry = buildSheetGeometry(sheet)
        for (const [name, batch] of [['resident', resident], ['worker', worker], ['forced-paged', forced]] as const) {
          const r = recorder(metrics.widthPx, metrics.heightPx)
          renderSheetViewport(sheet, r.ctx, geometry, { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
          out[name] = r.ops.filter(o => o.op === 'icon').length
        }
        { const print = computePrintMetrics(sheet, metrics); const r = recorder(print.paperPx.width, print.paperPx.height); renderPrintPage(sheet, r.ctx, metrics, print, 0, 0, undefined, { conditionalFormatting: resident }); out.print = r.ops.filter(o => o.op === 'icon').length }
        out.described = [0, 1, 2].filter(row => resident.iconAt(row, 0) !== undefined).length
        out.inset = [0, 1, 2].filter(row => resident.textInsetAt(row, 0, { x: 0, y: 0, width: metrics.colWidthsPx[0]!, height: metrics.rowHeightsPx[row]!, textWidth: 10, rowPt, cell: sheet.rows[row]!.cells[0]! }, { zoom: 1, dpr: 1 }) !== undefined).length
        resident.dispose(); worker.dispose(); forced.dispose()
        return out
      }, { ...production, extra: [...GLYPH_CAPS, 'icon:inset-left'] })
    }
    for (const rowPt of [15, 18, 30]) {
      for (const vertical of ['center', 'bottom'] as const) {
        const uniform = await count(vertical, rowPt)
        for (const name of ['static', 'resident', 'worker', 'forced-paged', 'print']) expect(uniform[name], `${vertical} ${rowPt} pt uniform: ${name}`).toBe(6)
        expect(uniform.described, `${vertical} ${rowPt} pt uniform`).toBe(3)
        expect(uniform.inset, `${vertical} ${rowPt} pt uniform`).toBe(3)
      }
      // A 20 pt cell in the row: bottom withholds in every view; centred still paints.
      expect(await count('bottom', rowPt, true), `bottom ${rowPt} pt with a 20 pt cell`).toEqual({ static: 0, resident: 0, worker: 0, 'forced-paged': 0, print: 0, described: 0, inset: 0 })
      expect((await count('center', rowPt, true)).worker, `centre ${rowPt} pt with a 20 pt cell`).toBe(6)
    }
  })

  it('fix round P3: right-aligned text clears the positioned frame\'s right edge (4Rating:3, 15 pt row, 100 px cell)', async () => {
    // The reproduction: 78 px of text in a 100 px cell starts at 100 - 3 - 78 = 19 px, inside the frame (2 pt + 13 pt =
    // 20 px). Round 8 (S4-CLEAR-ALLZOOM): admission needs a 10 px clearance (every zoom and DPR >= 1): 67 px paints.
    const desc: CfIconDescriptor = { ruleId: 1, set: '4Rating', iconId: 3, bucket: 3, buckets: 4, showValue: true }
    const host = (textWidth: number) => ({ x: 0, y: 0, width: 100, height: 20, textWidth, rowPt: 15, cell: { value: 1234, style: { ...KEY_STYLE, vertical: 'center' as const } } })
    const caps = new Set([...ICON_PAINT_EVIDENCE, 'icon:geometry-row-offset'])
    injectCfIconGlyphs((set, id) => cfIconGlyphHeader(`${set}:${id}`))
    try {
      expect(iconWithheld(desc, host(78), c => caps.has(c), CF_ICON_RENDER_MODEL)).toBe(true)
      expect(iconWithheld(desc, host(77), c => caps.has(c), CF_ICON_RENDER_MODEL)).toBe(true)
      expect(iconWithheld(desc, host(67), c => caps.has(c), CF_ICON_RENDER_MODEL)).toBe(false)
    } finally { injectCfIconGlyphs(null) }
  })

  it('Task 30 (S4-DPR-FLOOR): below device scale 0.25 the paint rechecks the snapped clearance and withholds an intruding icon', () => {
    // A right-aligned number whose text starts at x = 30 (logical) in a 100 px host: at device scale s its first column
    // is floor(30 s). A frame whose snapped right edge passes it intrudes; at or after it, it clears.
    const host = { x: 0, y: 0, width: 100, height: 20, textWidth: 67, rowPt: 15, cell: { value: 1234, style: { ...KEY_STYLE, vertical: 'center' as const } } }
    const at = (zoom: number, dpr: number, right: number) => iconClearsRightTextAt(host, { x: 0, y: 0, w: right, h: 2 }, { zoom, dpr, output: 'device' })
    expect(at(0.25, 0.5, 3)).toBe(true) // floor(30 x 0.125) = 3: the frame ends at the text's first column
    expect(at(0.25, 0.5, 4)).toBe(false)
    // Logical output: the box is in logical px, snapped to the device grid (edge 4 device px at s = 0.125 is 32 px).
    expect(iconClearsRightTextAt(host, { x: 0, y: 0, w: 32, h: 2 }, { zoom: 0.25, dpr: 0.5 })).toBe(false)
    expect(iconClearsRightTextAt(host, { x: 0, y: 0, w: 24, h: 2 }, { zoom: 0.25, dpr: 0.5 })).toBe(true)
    // Left-aligned and centred text are not rechecked (their inset is laid out beside the frame).
    expect(iconClearsRightTextAt({ ...host, cell: { value: 'a', style: { ...KEY_STYLE, horizontal: 'left' as const, vertical: 'center' as const } } }, { x: 0, y: 0, w: 99, h: 2 }, { zoom: 0.25, dpr: 0.5, output: 'device' })).toBe(true)
    // The production frame at the tightest admitted clearance (10 px) clears at every scale from 0.05 to 4 (sweep).
    const desc: CfIconDescriptor = { ruleId: 1, set: '3Arrows', iconId: 2, bucket: 2, buckets: 3, showValue: true }
    const caps = new Set([...ICON_PAINT_EVIDENCE, 'icon:geometry-row-offset'])
    let swept = 0
    for (let x = 0; x < 8; x += 0.125) for (let s = 0.05; s <= 4; s += 0.01) {
      const h = { ...host, x }
      const laid = layoutIcon(desc, h, { sizePt: 11, vertical: 'center', horizontal: 'right' }, { zoom: s, dpr: 1, output: 'device' }, CF_ICON_RENDER_MODEL, c => caps.has(c))
      if (!laid) continue
      expect(iconClearsRightTextAt(h, laid.box, { zoom: s, dpr: 1, output: 'device' }), `x ${x} s ${s}`).toBe(true)
      swept++
    }
    expect(swept).toBeGreaterThan(10000)
  })

  it('round 8 (S4-CLEAR-ALLZOOM): right-aligned clearance holds at every supported zoom (steps and continuous) and DPR', () => {
    const desc: CfIconDescriptor = { ruleId: 1, set: '3Arrows', iconId: 2, bucket: 2, buckets: 3, showValue: true }
    const host = (x: number, textWidth = 77) => ({ x, y: 0, width: 100, height: 20, textWidth, rowPt: 15, cell: { value: 1234, style: { ...KEY_STYLE, vertical: 'center' as const } } })
    const caps = new Set([...ICON_PAINT_EVIDENCE, 'icon:geometry-row-offset'])
    injectCfIconGlyphs((set, id) => cfIconGlyphHeader(`${set}:${id}`))
    try {
      // Codex's reproduction: x = 0, 77 px of text (it starts at 20 px, the frame's exact edge). At zoom 0.33 the
      // snapped frame ends at device 7 while the text starts in column 6; at 0.67, 14 against 13. Withheld.
      expect(iconWithheld(desc, host(0), c => caps.has(c), CF_ICON_RENDER_MODEL)).toBe(true)
      // The round-7 case (x = 0.5) as well.
      expect(iconWithheld(desc, host(0.5), c => caps.has(c), CF_ICON_RENDER_MODEL)).toBe(true)
      // 10 px of clearance (67 px of text) paints; 9.9 px does not.
      expect(iconWithheld(desc, host(0, 67), c => caps.has(c), CF_ICON_RENDER_MODEL)).toBe(false)
      expect(iconWithheld(desc, host(0, 67.1), c => caps.has(c), CF_ICON_RENDER_MODEL)).toBe(true)
      // Property: wherever the predicate admits, no keyboard zoom step x supported DPR, nor any continuous wheel zoom
      // in [0.25, 4], lays the frame over the text's first device column.
      const scales: Array<readonly [number, number]> = [...CF_ICON_SCALES]
      for (let k = 0; k <= 60; k++) for (const dpr of [1, 1.25, 1.5, 2, 3]) scales.push([0.25 * Math.pow(16, k / 60), dpr])
      expect(CF_ICON_SCALES.map(([z]) => z)).toEqual(expect.arrayContaining([...SHEET_ZOOM_STEPS]))
      for (let i = 0; i < 300; i++) {
        const x = (i * 7.31) % 50, tw = 50 + (i * 3.17) % 30
        const h = host(x, tw)
        if (iconWithheld(desc, h, c => caps.has(c), CF_ICON_RENDER_MODEL)) continue
        for (const [zoom, dpr] of scales) {
          const laid = layoutIcon(desc, h, { sizePt: 11, vertical: 'center', horizontal: 'right' }, { zoom, dpr, output: 'device' }, CF_ICON_RENDER_MODEL, c => caps.has(c))
          if (!laid) continue
          const textStart = Math.floor((x + 100 - 3 - tw) * zoom * dpr)
          expect(laid.box.x + laid.box.w, `x=${x} tw=${tw} ${zoom}/${dpr}`).toBeLessThanOrEqual(textStart)
        }
      }
    } finally { injectCfIconGlyphs(null) }
  })

  it('fix round P3: the ±1 px boundary of the frame clearance holds in every view', async () => {
    const probe = createCellLayoutHost(layoutContext())
    const value = 1234567
    const body = probe.layoutBody({ ref: 'A1', row: 0, col: 0, value, styleIndex: 0, style: { ...KEY_STYLE, vertical: 'center' } }, 500, 20)
    const width = body.text?.kind === 'plain' ? body.text.width : body.text?.kind === 'layout' ? body.text.layout.lines[0]!.segments.reduce((n, g) => n + g.width, 0) : NaN
    // The text starts at col - 3 - width; it must start at least 10 px after the frame's right edge, 20 px (round 8).
    const clear = Math.ceil(3 + width + 20 + 10)
    for (const [name, px, want] of [['clear', clear, 2], ['one px in', clear - 1, 0]] as const) {
      const sheet = keySheet([value], px / 7)
      expect(computeMetrics(sheet).colWidthsPx[0], name).toBe(px)
      const got = await withIcons(() => views(sheet), { ...production, extra: GLYPH_CAPS })
      for (const view of ['static', 'resident', 'paged', 'print', 'export']) expect(got.result[view]!.icons, `${name}: ${view}`).toBe(want)
      expect(got.iconAt, name).toEqual([want > 0, want > 0])
    }
  })

  it('the production model reads rightInsetPt 0 at 11 pt', () => {
    expect(CF_ICON_RENDER_MODEL.rightInsetPt).toEqual({ 11: 0 })
  })

  it('a short right-aligned number paints its icon in all five views with its text exactly where it is without one', async () => {
    const sheet = keySheet([10, 50, 90])
    expect(computeMetrics(sheet).rowHeightsPx[0]).toBe(20)
    const withIcon = await withIcons(() => views(sheet), { ...production, extra: GLYPH_CAPS })
    const control = await withIcons(() => views(sheet), production)
    for (const name of ['static', 'resident', 'paged', 'print', 'export']) {
      // Each 3Arrows glyph is two fills: three icons, six fills; the control (no glyph evidence) paints none.
      expect(withIcon.result[name]!.icons, `${name}: icons`).toBe(6)
      expect(control.result[name]!.icons, `${name}: control`).toBe(0)
      expect(withIcon.result[name]!.texts, `${name}: text unshifted`).toEqual(control.result[name]!.texts)
      expect(withIcon.result[name]!.texts.map(t => t.text), name).toEqual(['10', '50', '90'])
    }
    expect(withIcon.iconAt).toEqual([true, true])
  })

  it('a right-aligned number that reaches the icon frame withholds on every view (overflow unmeasured)', async () => {
    const probe = createCellLayoutHost(layoutContext())
    const value = 1234567
    const body = probe.layoutBody({ ref: 'A1', row: 0, col: 0, value, styleIndex: 0, style: { ...KEY_STYLE, vertical: 'center' } }, 500, 20)
    const width = body.text?.kind === 'plain' ? body.text.width : body.text?.kind === 'layout' ? body.text.layout.lines[0]!.segments.reduce((n, s) => n + s.width, 0) : NaN
    expect(width).toBeGreaterThan(20)
    // Fits the cell (w - 6 >= width) but not right of the 12 pt (16 px) frame and gap.
    const colPx = Math.ceil(width + CF_ICON_TEXT_PADDING_PX + 8)
    for (const [name, px, want] of [['tight', colPx, 0], ['control', colPx + 30, 2]] as const) {
      const sheet = keySheet([value], px / 7)
      expect(computeMetrics(sheet).colWidthsPx[0], name).toBe(px)
      const got = await withIcons(() => views(sheet), { ...production, extra: [...GLYPH_CAPS, 'icon:inset-overflow'] })
      for (const view of ['static', 'resident', 'paged', 'print', 'export']) expect(got.result[view]!.icons, `${name}: ${view}`).toBe(want)
      expect(got.iconAt, name).toEqual([want > 0, want > 0])
    }
  })
})
