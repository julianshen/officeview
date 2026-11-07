// @vitest-environment node
import { createCanvas } from 'canvas'
import { expect, it } from 'vitest'
import { getPaintables } from '../src/render/paint'
import type { XlsxSheet } from '../src/xlsx/types'
it('print scale consumes the same prepared CF fill', async () => {
  const sheet: XlsxSheet = { name: 'CF', merges: [], mergeRanges: [], rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 2, styleIndex: 0 }] }], cols: [], conditionalFormatting: { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], writes: new Set(['fill']), dxf: { 'fill.bg': 'FF00FF00' }, definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] } }] } }
  const pages = await getPaintables({ sheets: [sheet] })
  try { for (const scale of [1, 2]) {
    const canvas = createCanvas(pages[0].spec.widthPx * scale, pages[0].spec.heightPx * scale), ctx = canvas.getContext('2d'); ctx.scale(scale, scale)
    pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
    expect([...ctx.getImageData(10 * scale, 5 * scale, 1, 1).data]).toEqual([0, 255, 0, 255])
  } } finally { pages.dispose() }
})
it('interactive/print sheet assets wait for CF preparation', async () => {
  const { prepareSheetAssets } = await import('../src/components/xlsx/sheetAssets')
  const held = await prepareSheetAssets({ sheets: [{ name: 'CF', merges: [], mergeRanges: [], rows: [], cols: [], conditionalFormatting: { rules: [], diagnostics: [] } }] })
  try { expect(held.assets.conditionalFormatting).toBeDefined() } finally { held.dispose() }
})
it('delayed optional renderer prevents publishing cacheable sheet assets', async () => {
  const { vi } = await import('vitest')
  const loader = await import('../src/xlsx/cfLoader')
  const { prepareSheetAssets } = await import('../src/components/xlsx/sheetAssets')
  const module = await import('../src/xlsx/cf/render')
  let finish!: () => void
  const gate = new Promise<void>(r => { finish = r })
  const load = vi.spyOn(loader, 'ensureConditionalFormatting').mockImplementation(async () => { await gate; return module })
  let published = false
  const pending = prepareSheetAssets({ sheets: [{ name: 'CF', rows: [], cols: [], merges: [], mergeRanges: [], conditionalFormatting: { rules: [], diagnostics: [] } }] }).then(held => { published = true; return held })
  try {
    await new Promise(r => setTimeout(r, 10))
    expect(load).toHaveBeenCalledTimes(1); expect(published).toBe(false)
    finish(); const held = await pending
    expect(held.assets.conditionalFormatting).toBeDefined(); held.dispose()
    const plain = await prepareSheetAssets({ sheets: [{ name: 'plain', rows: [], cols: [], merges: [], mergeRanges: [] }] })
    plain.dispose(); expect(load).toHaveBeenCalledTimes(1)
  } finally { finish(); load.mockRestore() }
})
it('actual print pages share the prepared fill, border and number format with screen', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { computeMetrics, computePrintMetrics, renderPrintPage, renderSheet } = await import('../src/xlsx/render')
  const sheet: XlsxSheet = { name: 'CF', rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 1.25, styleIndex: 0, style: { numFmtId: 0, formatCode: '0.0' } }] }], cols: [], merges: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], writes: new Set(['fill', 'numFmt', 'border.bottom']), dxf: { 'fill.bg': 'FF00FF00', 'numFmt.code': '0.00', 'border.bottom.style': 'thick', 'border.bottom.color': 'FFFF0000' }, definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] } }] } }
  const batch = await prepareCfPaint(sheet), m = computeMetrics(sheet), print = computePrintMetrics(sheet, m)
  try { for (const printMode of [false, true]) {
    const canvas = createCanvas(print.paperPx.width, print.paperPx.height), ctx = canvas.getContext('2d'), texts: string[] = [], strokes: string[] = [], fills: string[] = []
    const fillText = ctx.fillText.bind(ctx), stroke = ctx.stroke.bind(ctx), fillRect = ctx.fillRect.bind(ctx)
    ctx.fillText = (text, x, y) => { texts.push(text); fillText(text, x, y) }
    ctx.stroke = () => { strokes.push(String(ctx.strokeStyle)); stroke() }
    ctx.fillRect = (x, y, w, h) => { fills.push(String(ctx.fillStyle)); fillRect(x, y, w, h) }
    if (printMode) renderPrintPage(sheet, ctx as unknown as CanvasRenderingContext2D, m, print, 0, 0, undefined, { conditionalFormatting: batch })
    else renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m, undefined, { conditionalFormatting: batch })
    expect(texts).toContain('1.25'); expect(texts).not.toContain('1.3')
    expect(fills).toContain('#00ff00'); expect(strokes).toContain('#ff0000')
  } } finally { batch?.dispose() }
})
