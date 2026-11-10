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
it('print output at screen and PRINT_SCALE = 2 carries the prepared scale fill', async () => {
  const { withInjectedCfEvidence } = await import('../src/xlsx/cf/evidence')
  const { cfRuntimeEvidence } = await import('../src/xlsx/cf/predicates')
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { computeMetrics, computePrintMetrics, renderPrintPage } = await import('../src/xlsx/render')
  const scaleEvidence = [
    'scale:interpolation-color-space',
    'scale:interpolation-rounding-ladder',
    'scale:thresholds-min-max',
  ]
  await withInjectedCfEvidence(scaleEvidence, async () => {
    const sheet: XlsxSheet = {
      name: 'CFScale',
      rows: [
        { index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 10, styleIndex: 0 }] },
        { index: 1, cells: [{ row: 1, col: 0, ref: 'A2', value: 70, styleIndex: 0 }] },
      ],
      cols: [],
      merges: [],
      mergeRanges: [],
      conditionalFormatting: {
        diagnostics: [],
        rules: [
          {
            id: 1,
            priority: 1,
            documentOrder: 0,
            stopIfTrue: false,
            ranges: [{ firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 0 }],
            writes: new Set(['colorScale', 'fill']),
            definition: {
              kind: 'colorScale',
              points: [
                { bound: { type: 'min' }, argb: 0xffff0000 },
                { bound: { type: 'max' }, argb: 0xff00ff00 },
              ],
              dialect: 'base',
              formulas: [],
            },
          },
        ],
      },
    }
    const batch = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence })
    expect(batch).toBeDefined()
    const m = computeMetrics(sheet)
    const print = computePrintMetrics(sheet, m)

    // Interactive screen fill
    const screenFillA1 = batch!.layer.styleAt(0, 0, undefined)?.fillColor
    const screenFillA2 = batch!.layer.styleAt(1, 0, undefined)?.fillColor
    expect(screenFillA1).toBe('FFFF0000')
    expect(screenFillA2).toBe('FF00FF00')

    // At PRINT_SCALE = 2:
    const PRINT_SCALE = 2
    const canvas = createCanvas(print.paperPx.width * PRINT_SCALE, print.paperPx.height * PRINT_SCALE)
    const ctx = canvas.getContext('2d')
    ctx.scale(PRINT_SCALE, PRINT_SCALE)

    const fills: string[] = []
    const origFillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x, y, w, h) => {
      fills.push(String(ctx.fillStyle).toLowerCase())
      origFillRect(x, y, w, h)
    }

    renderPrintPage(sheet, ctx as unknown as CanvasRenderingContext2D, m, print, 0, 0, undefined, { conditionalFormatting: batch })

    // Exact ARGB matches interactive state
    expect(fills).toContain('#ff0000')
    expect(fills).toContain('#00ff00')

    batch?.dispose()
  })
})
it('print uses verified measured printer-appearance variant (stage2-scale-interpolation-printer.json)', async () => {
  const fs = await import('node:fs')
  const path = await import('node:path')
  const { withInjectedCfEvidence } = await import('../src/xlsx/cf/evidence')
  const { cfRuntimeEvidence } = await import('../src/xlsx/cf/predicates')
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { parseXlsx } = await import('../src/xlsx/parse')
  const { OfficePackage } = await import('../src/core/zip')

  const scaleEvidence = [
    'scale:interpolation-color-space',
    'scale:interpolation-rounding-ladder',
    'scale:thresholds-min-max',
    'scale:thresholds-num-bounds',
    'scale:thresholds-percentile-inc-exc',
  ]
  await withInjectedCfEvidence(scaleEvidence, async () => {
    const fixturePath = path.resolve(__dirname, 'fixtures/xlsx-cf/refs/stage2-scale-interpolation-printer.json')
    const xlsxPath = path.resolve(__dirname, 'fixtures/xlsx-cf/stage2-scale-interpolation.xlsx')
    const fixture = JSON.parse(fs.readFileSync(fixturePath, 'utf8'))
    expect(fixture.appearance).toBe('printer')

    const bytes = new Uint8Array(fs.readFileSync(xlsxPath))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const sheet = doc.sheets[0]!
    const batch = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence })
    expect(batch).toBeDefined()

    // B2 is min (10) -> red FFFF0000
    const b2Observed = fixture.observedCells.find((c: any) => c.ref === 'B2')
    expect(b2Observed.fillRgb).toBe('FFFF0000')
    const styleB2 = batch!.layer.styleAt(1, 1, undefined)
    expect(styleB2?.fillColor).toBe('FFFF0000')

    batch?.dispose()
  })
})
