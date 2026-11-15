// @vitest-environment node
/**
 * Task 29 fix round 3, Ruling S4-PROD-ALIGN: production icon painting adopts exactly the reading Excel 16.115 was
 * measured to show for each Stage 4 geometry, inset and showValue capability (docs/validation/xlsx-cf.md §12, rounds
 * 3–10; the matched readings are recomputed from the committed sidecars in tests/xlsx-cf-oracle-conformance.test.ts).
 *
 * The painting itself is exercised under the injected stubs of tests/helpers/cf-icons.ts (hypothetical box geometry,
 * never Excel facts), because the measured box is not determined by the evidence (ink is not the box) and no glyph
 * table exists. The readings, and the production model's measured values, are pinned against production directly.
 */
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { computeMetrics, computePrintMetrics, renderPrintPage, renderSheet } from '../src/xlsx/render'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { prepareCfPaint, residentSource, type CfPaintBatch } from '../src/xlsx/cf/render'
import { prepareCfWireScope, type CfFrameHostSource } from '../src/xlsx/cf/text'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import {
  CF_ICON_PAINT_CAPABILITIES, CF_ICON_PAINT_READINGS, CF_ICON_RENDER_MODEL, iconHidesText, iconTextInset, iconWithheld, layoutIcon,
  type CfIconGate, type CfIconRenderModel,
} from '../src/xlsx/cf/iconPaint'
import type { CfIconDescriptor, CfStamp } from '../src/xlsx/cf/model'
import type { XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'
import type { EffectiveFormatScope } from '../src/xlsx/view/formatWorkspace'
import { ICON_COLOURS, ICON_SHOWVALUE_EVIDENCE, STUB_INSET_PX, STUB_MODEL, iconRecorder, iconSheet, withIcons } from './helpers/cf-icons'
import { barRule } from './helpers/cf-bars'

const ALL = (): boolean => true
const except = (...caps: string[]) => (c: string): boolean => !caps.includes(c)
const AT_CAPTURE = { zoom: 0.75, dpr: 1, output: 'device' as const }
const PT_PX = 4 / 3
const layoutContext = () => createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
const icon = (overrides: Partial<CfIconDescriptor> = {}): CfIconDescriptor => ({ ruleId: 1, set: '3Arrows', iconId: 1, bucket: 1, buckets: 3, showValue: true, ...overrides })
const NO_ICONS = (overrides: Partial<CfIconDescriptor> = {}): CfIconDescriptor => icon({ set: 'NoIcons', iconId: 0, ...overrides })
const LEFT: XlsxCellStyle = { numFmtId: 0, horizontal: 'left' }
const host = (overrides: Record<string, unknown> = {}) => ({ x: 0, y: 0, width: 200, height: 20, textWidth: 12, cell: { value: 50, style: LEFT }, ...overrides })
const gate = (check: (cap: string) => boolean): CfIconGate => ({ admitted: { cacheFor: () => ({ classes: new Set<string>(), release() {} }) as never, release() {} }, hasBar: () => false, check })
const isIcon = (o: { op: string; style: string }) => o.op === 'icon' && (ICON_COLOURS as readonly string[]).includes(o.style)

describe('S4-PROD-ALIGN: the production model and the readings it adopts', () => {
  it('CF_ICON_RENDER_MODEL carries only what the sidecars determine: the font-scaled, cell-aligned box variant and the 12 pt left inset at 11 pt', () => {
    expect(CF_ICON_RENDER_MODEL.box).toEqual({ size: 'font', vertical: 'cell-align' })
    expect(CF_ICON_RENDER_MODEL.leftInsetPt).toEqual({ 11: 12 })
    // Undetermined by the evidence (ink is not the box; Task 28's glyphs place the ink): withheld.
    for (const key of ['sizePt', 'offsetsPt', 'rightInsetPt'] as const) expect(CF_ICON_RENDER_MODEL[key], key).toBeUndefined()
    expect(Object.isFrozen(CF_ICON_RENDER_MODEL)).toBe(true)
    // So under every capability the production model still lays no icon out.
    expect(layoutIcon(icon(), host(), { sizePt: 11, vertical: 'bottom', horizontal: 'left' }, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)).toBeUndefined()
  })

  it('the readings production implements are the ones Excel matched', () => {
    expect(CF_ICON_PAINT_READINGS).toEqual({
      'icon:geometry-box': 'font',
      'icon:geometry-vertical-align': 'cell-align',
      'icon:geometry-narrow-clip': 'clip',
      'icon:geometry-short-row-clip': 'shrink',
      'icon:inset-left': 'inset',
      'icon:inset-right-aligned': 'unshifted',
      'icon:inset-center': 'full-cell',
      'icon:inset-overflow': 'icon-and-text',
      'icon:noicons-inset': 'inset',
      'icon:noicons-showvalue': 'text-hidden',
      'icon:showvalue-canvas-suppression': 'text-hidden',
      'icon:composition-bar-showvalue': 'priority-wins',
    })
  })

  it('item 9 (S4-SNAP-SHARED): the paint gates use the verified bar:geometry-snapping, not icon:geometry-snapping', () => {
    expect(CF_ICON_PAINT_CAPABILITIES).toEqual(['icon:geometry-box', 'icon:geometry-vertical-align', 'bar:geometry-snapping'])
    const font = { sizePt: 10, vertical: 'bottom' as const, horizontal: 'left' as const }
    expect(layoutIcon(icon(), host(), font, AT_CAPTURE, STUB_MODEL, except('icon:geometry-snapping'))).toBeDefined()
    expect(layoutIcon(icon(), host(), font, AT_CAPTURE, STUB_MODEL, except('bar:geometry-snapping'))).toBeUndefined()
  })
})

describe('S4-PROD-ALIGN: each capability behaves as its matched reading', () => {
  it('item 1 (icon:geometry-box = font): the box side is the model scale times the font size', () => {
    const model: CfIconRenderModel = { ...STUB_MODEL, box: { size: 'font', vertical: 'bottom' }, sizePt: 1 }
    for (const pt of [8, 11, 20]) {
      const laid = layoutIcon(icon(), host({ height: 60 }), { sizePt: pt, vertical: 'bottom', horizontal: 'left' }, AT_CAPTURE, model, ALL)!
      expect(laid.box.w, `${pt} pt`).toBe(pt)
      expect(laid.box.h, `${pt} pt`).toBe(pt)
    }
  })

  it('item 2 (icon:geometry-vertical-align = cell-align): the box follows the cell\'s vertical alignment', () => {
    const model: CfIconRenderModel = { ...STUB_MODEL, box: { size: 'fixed', vertical: 'cell-align' }, offsetsPt: { left: 2, top: 1, bottom: 1 } }
    const y = (vertical: 'top' | 'center' | 'bottom') => layoutIcon(icon(), host({ height: 40 }), { sizePt: 10, vertical, horizontal: 'left' }, AT_CAPTURE, model, ALL)!.box.y
    // 40 logical px = 30 device px at the capture scale; a 9 px box.
    expect(y('top')).toBe(1)
    expect(y('bottom')).toBe(30 - 1 - 9)
    expect(y('center')).toBe(Math.floor((30 - 9) / 2 + 0.5))
  })

  it('item 3 (icon:geometry-short-row-clip = shrink): a row shorter than the box is never clipped; the shrink law is unmeasured, so it withholds', () => {
    const font = { sizePt: 10, vertical: 'bottom' as const, horizontal: 'left' as const }
    // Control: the same box in a row that holds it.
    expect(layoutIcon(icon(), host(), font, AT_CAPTURE, STUB_MODEL, ALL)).toBeDefined()
    // 6 logical px = 4.5 device px: the 9 px box does not fit. Excel shrinks it (L5: 5×4 ink against 11); production
    // has no measured shrink law and must not clip.
    expect(layoutIcon(icon(), host({ height: 6 }), font, AT_CAPTURE, STUB_MODEL, ALL)).toBeUndefined()
    expect(iconWithheld(icon(), host({ height: 6 }), ALL, STUB_MODEL)).toBe(true)
  })

  it('item 4 (icon:inset-center = full-cell): centred text is not inset', () => {
    const laid = layoutIcon(icon(), host(), { sizePt: 10, vertical: 'bottom', horizontal: 'center' }, AT_CAPTURE, STUB_MODEL, ALL)!
    expect(laid).toBeDefined()
    expect(laid.inset).toEqual({ left: 0, right: 0 })
    expect(layoutIcon(icon(), host(), { sizePt: 10, vertical: 'bottom', horizontal: 'center' }, AT_CAPTURE, STUB_MODEL, except('icon:inset-center'))).toBeUndefined()
  })

  it('item 5 (icon:inset-overflow = icon-and-text): a number that does not fit beside the inset keeps its icon; text overflow stays withheld', async () => {
    await withIcons(() => {
      const tight = { width: 40, textWidth: 30 } // 40 − 6 − 17.3 < 30
      expect(iconWithheld(icon(), host({ ...tight }), ALL)).toBe(false)
      expect(iconWithheld(icon(), host({ ...tight }), except('icon:inset-overflow'))).toBe(true)
      // Excel's capture holds numbers only (stage4-icon-inset E, a '#' fill beside the icon): a string is unmeasured.
      expect(iconWithheld(icon(), host({ ...tight, cell: { value: 'abcdef', style: LEFT } }), ALL)).toBe(true)
      // An unknown width still withholds.
      expect(iconWithheld(icon(), host({ ...tight, textWidth: undefined }), ALL)).toBe(true)
    })
  })

  it('item 6 (icon:noicons-inset = inset): a NoIcons cell\'s text is inset as if an icon were present', async () => {
    await withIcons(() => {
      const g = gate(ALL)
      expect(iconTextInset(NO_ICONS(), 0, 0, host(), { zoom: 1, dpr: 1 }, g)?.left).toBeCloseTo(STUB_INSET_PX, 9)
      expect(iconTextInset(NO_ICONS(), 0, 0, host(), { zoom: 1, dpr: 1 }, gate(except('icon:noicons-inset')))).toBeUndefined()
      // The inset itself is the alignment's measured one: centred text keeps the whole cell.
      expect(iconTextInset(NO_ICONS(), 0, 0, host({ cell: { value: 50, style: { numFmtId: 0, horizontal: 'center' } } }), { zoom: 1, dpr: 1 }, g)).toEqual({ left: 0, right: 0 })
    })
    // The production model: 12 pt at 11 pt (stage4-icon-inset F2..F4, +12 px at 1 px/pt), nothing at other sizes.
    const at = (fontSizePt: number) => iconTextInset(NO_ICONS(), 0, 0, host({ cell: { value: 50, style: { ...LEFT, fontSizePt } } }), { zoom: 1, dpr: 1 }, gate(ALL))
    expect(at(11)?.left).toBeCloseTo(12 * PT_PX, 9)
    expect(at(12)).toBeUndefined()
  })

  it('item 7 (icon:noicons-showvalue = text-hidden): NoIcons with showValue="0" hides the text', async () => {
    await withIcons(() => {
      expect(iconHidesText(NO_ICONS({ showValue: false }), 0, 0, host(), gate(ALL))).toBe(true)
      expect(iconHidesText(NO_ICONS({ showValue: false }), 0, 0, host(), gate(except('icon:noicons-showvalue')))).toBe(false)
      expect(iconHidesText(NO_ICONS({ showValue: true }), 0, 0, host(), gate(ALL))).toBe(false)
    })
  })

  it('item 10: icon:inset-left insets left text by the model inset; icon:geometry-narrow-clip clips a narrow host', () => {
    const font = { sizePt: 10, vertical: 'bottom' as const, horizontal: 'left' as const }
    expect(layoutIcon(icon(), host(), font, AT_CAPTURE, STUB_MODEL, ALL)!.inset.left).toBeCloseTo(STUB_INSET_PX, 9)
    expect(layoutIcon(icon(), host(), font, AT_CAPTURE, STUB_MODEL, except('icon:inset-left'))).toBeUndefined()
    const narrow = host({ width: 8 })
    expect(layoutIcon(icon(), narrow, font, AT_CAPTURE, STUB_MODEL, ALL)).toBeDefined()
    expect(layoutIcon(icon(), narrow, font, AT_CAPTURE, STUB_MODEL, except('icon:geometry-narrow-clip'))).toBeUndefined()
  })
})

/** The paged main thread's batch, built as InteractiveSheetDoc does (wire → structured clone → wire scope). */
async function pagedScope(sheet: XlsxSheet, rows: number, hosts?: CfFrameHostSource): Promise<EffectiveFormatScope> {
  const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId ?? sheet.name
  const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
  const memory = createCfMemory({})
  const rects = [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }]
  const wire = await prepareCfRemote({ rects, stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, {
    budget: () => createCfWorkBudget({ now: () => 0 }),
    // As PagedSheetSource passes them (pagedSource.ts `merges`).
    merges: (sheet.mergeRanges ?? []).map(m => ({ firstRow: m.minRow, lastRow: m.maxRow, firstCol: m.minCol, lastCol: m.maxCol })),
  })
  expect(wire.kind, JSON.stringify(wire.reason)).toBe('ready')
  return (await prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), rects, { budget: { reserve() {}, release() {} }, locale: 'en-US', ...(hosts ? { hosts } : {}) }))!
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

type Got = { icons: number[]; texts: Array<{ row: number; text: string; x: number }> }
/** Static, resident viewport, worker (paged wire scope), forced-paged (semantic wire scope with frame hosts) and print. */
async function everyView(sheet: XlsxSheet, rows: number): Promise<Record<string, Got>> {
  const metrics = computeMetrics(sheet)
  const rowOf = (y: number) => Math.floor(y / 20)
  const read = (r: ReturnType<typeof iconRecorder>): Got => ({
    icons: r.ops.filter(isIcon).map(o => rowOf(o.y)),
    texts: r.ops.filter(o => o.op === 'text' && !/^Zoom in /.test(o.text!) && !/^Loading/.test(o.text!)).map(o => ({ row: rowOf(o.y - 1), text: o.text!, x: o.x })),
  })
  const rec = () => iconRecorder(metrics.widthPx, metrics.heightPx, createCanvas as never)
  const viewport = (batch: CfPaintBatch) => {
    const r = rec()
    renderSheetViewport(sheet, r.ctx, buildSheetGeometry(sheet), { box: { x: 0, y: 0, width: metrics.widthPx, height: metrics.heightPx }, zoom: 1, deviceScale: 1 }, { conditionalFormatting: batch }, { cache: new LayoutCache() })
    return read(r)
  }
  const resident = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext }))!
  const out: Record<string, Got> = {}
  { const r = rec(); renderSheet(sheet, r.ctx, metrics, undefined, { conditionalFormatting: resident }); out.static = read(r) }
  out.resident = viewport(resident)
  const worker = asBatch(await pagedScope(sheet, rows))
  out.worker = viewport(worker)
  const box = (row: number) => ({ x: 0, y: metrics.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0), width: metrics.colWidthsPx[0]!, height: metrics.rowHeightsPx[row]! })
  const forced = asBatch(await pagedScope(sheet, rows, { cell: (row, col) => ({ box: box(row), base: sheet.rows[row]!.cells[col]!.style ?? { numFmtId: 0 }, textWidth: 12, cell: sheet.rows[row]!.cells[col]! }) }))
  out.forced = viewport(forced)
  {
    const print = computePrintMetrics(sheet, metrics)
    const r = iconRecorder(print.paperPx.width, print.paperPx.height, createCanvas as never)
    renderPrintPage(sheet, r.ctx, metrics, print, 0, 0, undefined, { conditionalFormatting: resident })
    // The page has its own origin: callers compare text order and x offsets against the same view's control only.
    out.print = read(r)
  }
  resident.dispose(); worker.dispose(); forced.dispose()
  return out
}

describe('S4-PROD-ALIGN: every view agrees', () => {
  const custom = [{ set: 'NoIcons' as const, iconId: 0 }, { set: 'NoIcons' as const, iconId: 0 }, { set: 'NoIcons' as const, iconId: 0 }]
  it('item 6: a NoIcons cell is inset by the left inset in the static, resident, worker, forced-paged and print views', async () => {
    const plain = async () => everyView(iconSheet([10, 50, 90], { style: LEFT, def: { custom } }), 3)
    const on = await withIcons(plain, { extra: ['icon:noicons-inset', 'icon:scope-merged-cell'] })
    const off = await withIcons(plain)
    for (const [view, got] of Object.entries(on)) {
      expect(got.icons, view).toEqual([])
      const base = off[view]!.texts
      expect(got.texts.map(t => t.text), view).toEqual(['10', '50', '90'])
      for (const [i, t] of got.texts.entries()) expect(t.x - base[i]!.x, `${view} ${t.text}`).toBeCloseTo(STUB_INSET_PX, 6)
    }
  })

  it('item 7: NoIcons with showValue="0" draws no text in any view, only with icon:noicons-showvalue', async () => {
    const sheet = () => iconSheet([10, 50, 90], { style: LEFT, def: { custom, showValue: false } })
    const on = await withIcons(() => everyView(sheet(), 3), { extra: ['icon:noicons-showvalue'] })
    const off = await withIcons(() => everyView(sheet(), 3))
    for (const view of Object.keys(on)) {
      expect(on[view]!.texts, view).toEqual([])
      expect(off[view]!.texts.map(t => t.text), view).toEqual(['10', '50', '90'])
    }
  })

  it('item 8: an icon(1) ahead of a bar(0) shows the text (priority-wins; any-hides would hide it) in every view', async () => {
    // The icon rule is priority 1, the bar priority 2 (stage4-icon-geometry-2 J).
    const sheet = () => iconSheet([10, 50, 90], { def: { showValue: true }, extraRules: [{ ...barRule({ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }, { axisArgb: 0xffffffff, showValue: false }, 2), priority: 2, documentOrder: 1 }] })
    const sheetH = () => iconSheet([10, 50, 90], { def: { showValue: false }, extraRules: [{ ...barRule({ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }, { axisArgb: 0xffffffff, showValue: true }, 2), priority: 2, documentOrder: 1 }] })
    const caps = [ICON_SHOWVALUE_EVIDENCE, 'bar:showvalue-canvas-suppression', 'icon:composition-bar-showvalue']
    const shown = await withIcons(() => everyView(sheet(), 3), { extra: caps })
    const hidden = await withIcons(() => everyView(sheetH(), 3), { extra: caps })
    for (const view of Object.keys(shown)) {
      expect(shown[view]!.texts.map(t => t.text), `${view}: icon(1) decides`).toEqual(['10', '50', '90'])
      expect(hidden[view]!.texts, `${view}: icon(0) decides`).toEqual([])
    }
  })

  it('item 10: a right-aligned number keeps its position under the inset (icon:inset-right-aligned = unshifted)', async () => {
    const sheet = () => iconSheet([10, 50, 90])
    const on = await withIcons(() => everyView(sheet(), 3))
    const off = await withIcons(() => everyView(sheet(), 3), { paint: false })
    for (const view of Object.keys(on)) {
      expect(on[view]!.icons.length, `${view}: icons paint`).toBe(3)
      expect(off[view]!.icons.length, `${view}: control paints none`).toBe(0)
      for (const [i, t] of on[view]!.texts.entries()) expect(t.x, `${view} ${t.text}`).toBeCloseTo(off[view]!.texts[i]!.x, 6)
    }
  })
})
