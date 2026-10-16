/**
 * Visual-review stream B (PR #11): XLSX cell-rendering corrections toward
 * Excel. Paint is recorded through a fixed-metric context so assertions do not
 * depend on host fonts.
 */
import { describe, expect, it } from 'vitest'
import { cellDisplayText, computeMetrics, renderSheet, type GridMetrics } from '../src/xlsx/render'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { buildGeometry } from '../src/xlsx/view/source'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { workbookFormatting } from '../src/xlsx/storage/pagedFormatting'
import { parseXml } from '../src/core/xml'
import JSZip from 'jszip'

interface PaintOp { op: string; args: unknown[]; fillStyle: unknown; strokeStyle: unknown; lineWidth: unknown; dash: number[]; font: string }
/** A canvas stand-in: every glyph is 7px wide; state at each call is recorded. */
function recordingContext(glyph: (text: string, font: string) => number = text => text.length * 7) {
  const ops: PaintOp[] = []
  const state: Record<string, unknown> = { fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1, font: '10px sans-serif', textBaseline: 'alphabetic' }
  let dash: number[] = []
  const stack: Array<{ state: Record<string, unknown>; dash: number[] }> = []
  const record = (op: string, args: unknown[]) => ops.push({ op, args, fillStyle: state.fillStyle, strokeStyle: state.strokeStyle, lineWidth: state.lineWidth, dash: [...dash], font: String(state.font) })
  const methods: Record<string, (...args: unknown[]) => unknown> = {
    save: () => { stack.push({ state: { ...state }, dash: [...dash] }) },
    restore: () => { const top = stack.pop(); if (top) { Object.assign(state, top.state); dash = top.dash } },
    setLineDash: (d: unknown) => { dash = [...(d as number[])] },
    getLineDash: () => [...dash],
    measureText: (text: unknown) => ({ width: glyph(String(text), String(state.font)) }),
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
  }
  const ctx = new Proxy(state, {
    get(target, key: string) {
      if (key in methods) return methods[key]
      if (key in target) return target[key]
      return (...args: unknown[]) => { record(key, args) }
    },
    set(target, key: string, value) { target[key] = value; return true },
  }) as unknown as CanvasRenderingContext2D
  return { ctx, ops }
}

function sheetOf(cells: Array<Partial<XlsxCell> & { row: number; col: number }>, extra: Partial<XlsxSheet> = {}): XlsxSheet {
  const rows = new Map<number, XlsxCell[]>()
  for (const c of cells) {
    const cell: XlsxCell = { ref: '', styleIndex: 0, value: null, ...c }
    rows.set(c.row, [...(rows.get(c.row) ?? []), cell])
  }
  return { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [...rows].sort((a, b) => a[0] - b[0]).map(([index, cs]) => ({ index, cells: cs })), ...extra }
}
const grid = (cols: number[], rows: number[]): GridMetrics => ({ colWidthsPx: cols, rowHeightsPx: rows, widthPx: cols.reduce((a, b) => a + b, 0), heightPx: rows.reduce((a, b) => a + b, 0) })
const style = (s: Partial<XlsxCellStyle>): XlsxCellStyle => ({ numFmtId: 0, ...s })

describe('#1 SpreadsheetML ARGB alpha byte is ignored (Excel treats 00RRGGBB as opaque)', () => {
  it('paints an openpyxl 00RRGGBB fill and font colour opaque in the static path', () => {
    const { ctx, ops } = recordingContext()
    const sheet = sheetOf([{ row: 0, col: 0, value: 'x', style: style({ fillColor: '00FFF2CC', color: '00FFFFFF' }) }])
    renderSheet(sheet, ctx, grid([64], [20]))
    expect(ops.some(o => o.op === 'fillRect' && o.fillStyle === '#FFF2CC')).toBe(true)
    expect(ops.find(o => o.op === 'fillText')?.fillStyle).toBe('#FFFFFF')
  })
})

describe('#11 worker-served object documents (compatibility mode) keep their cell styles', () => {
  it('the object source publishes the workbook style table, indexed by cellXfs', async () => {
    const bytes = await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', t: 'str', v: 'head', style: 1 }, { ref: 'B1', v: 3.14159, style: 2 }] }] }])
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const source = createObjectWorkbookSource(doc, 'd')
    const formatting = source.formatting?.()
    // The main-thread remote source styles cells only through this table
    // (viewerRemoteSource styleOf); without it fills, fonts and borders vanish.
    expect(formatting?.styles[1]).toMatchObject({ fillColor: 'FFFFFF00', bold: true, borders: { left: 'thin' } })
    expect(formatting?.styles[2]).toMatchObject({ numFmtId: 2 })
    expect(formatting?.styles[1]).toEqual(doc.sheets[0].rows[0].cells[0].style)
  })
})

describe('#10 a stored row height applies without customHeight (Excel honours ht)', () => {
  // poi-56295: <row ht="12.75"> with no customHeight is 17px in Excel.
  const sheet = (): XlsxSheet => ({ name: 'S', cols: [], merges: [], mergeRanges: [], rows: [0, 1, 2].map(index => ({ index, heightPt: 12.75, cells: [{ ref: '', row: index, col: 0, styleIndex: 0, value: `r${index}` }] })) })
  it('resident interactive, paged and static geometry agree on 17px', () => {
    const resident = buildSheetGeometry(sheet())
    if (resident.kind !== 'ready') throw new Error('geometry')
    expect(resident.metrics.rows.start(1)).toBe(17)
    expect(buildGeometry(sheet().rows, []).rows.start(1)).toBe(17)
    expect(computeMetrics(sheet()).rowHeightsPx.slice(0, 3)).toEqual([17, 17, 17])
  })
})

describe('#4 Excel border styles and colours', () => {
  const strokes = (ops: PaintOp[]) => ops.filter(o => o.op === 'stroke' && o.strokeStyle !== '#d0d0d0')
  const paint = (borders: XlsxCellStyle['borders'], borderColors?: XlsxCellStyle['borderColors']) => {
    const { ctx, ops } = recordingContext()
    renderSheet(sheetOf([{ row: 1, col: 1, value: 'b', style: style({ borders, borderColors }) }]), ctx, grid([64, 64, 64], [20, 20, 20]))
    return ops
  }
  it('thick red is a 3px #C00000 stroke (openpyxl 00C00000)', () => {
    const s = strokes(paint({ top: 'thick' }, { top: '00C00000' }))
    expect(s).toHaveLength(1)
    expect(s[0]).toMatchObject({ strokeStyle: '#C00000', lineWidth: 3, dash: [] })
  })
  it('medium is 2px; hair, dotted, dashed and the dash-dot families set a dash pattern', () => {
    expect(strokes(paint({ left: 'medium' }))[0]).toMatchObject({ lineWidth: 2, dash: [] })
    for (const name of ['hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot']) expect(strokes(paint({ left: name }))[0], name).toMatchObject({ lineWidth: 1, dash: expect.arrayContaining([expect.any(Number)]) })
    for (const name of ['mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot']) expect(strokes(paint({ left: name }))[0], name).toMatchObject({ lineWidth: 2, dash: expect.arrayContaining([expect.any(Number)]) })
    expect(strokes(paint({ left: 'dashDot' }))[0].dash).not.toEqual(strokes(paint({ left: 'dashed' }))[0].dash)
  })
  it('double draws two 1px lines a pixel apart in its colour', () => {
    const ops = paint({ bottom: 'double' }, { bottom: '0000B050' })
    const s = strokes(ops)
    expect(s).toHaveLength(2)
    expect(s.every(o => o.strokeStyle === '#00B050' && o.lineWidth === 1)).toBe(true)
    const ys = ops.filter(o => o.op === 'moveTo').map(o => (o.args as number[])[1]).filter(y => y !== undefined)
    expect(ys).toEqual(expect.arrayContaining([40.5 - 1, 40.5 + 1]))
  })
  it('thin with no colour stays the exact 1px black line; the dash state never leaks', () => {
    const { ctx, ops } = recordingContext()
    renderSheet(sheetOf([{ row: 0, col: 0, value: 'a', style: style({ borders: { left: 'dashed', right: 'thin' } }) }]), ctx, grid([64, 64], [20, 20]))
    const s = strokes(ops)
    expect(s.map(o => o.dash.length > 0)).toEqual([true, false])
    expect(s[1]).toMatchObject({ strokeStyle: '#000000', lineWidth: 1 })
    expect(ctx.getLineDash()).toEqual([])
  })
})

describe('#4 border colours parse from styles.xml', () => {
  it('records each side colour next to its style, in both the object and the paged resolution', async () => {
    const zip = await JSZip.loadAsync(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', t: 'str', v: 'x', style: 1 }] }] }]))
    const styles = `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="2"><border/><border><left style="thick"><color rgb="FFC00000"/></left><right style="dashed"><color rgb="000070C0"/></right><top style="double"><color indexed="64"/></top><bottom style="thin"/><diagonal/></border></borders>
      <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1"/></cellXfs></styleSheet>`
    zip.file('xl/styles.xml', styles)
    const doc = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const expected = { borders: { left: 'thick', right: 'dashed', top: 'double', bottom: 'thin' }, borderColors: { left: 'FFC00000', right: '000070C0' } }
    expect(doc.sheets[0].rows[0].cells[0].style).toMatchObject(expected)
    const paged = await workbookFormatting(parseXml('<workbook><workbookPr/></workbook>'), styles)
    expect(paged.styles[1]).toMatchObject(expected)
  })
})

describe('#8 gridlines are hidden under text spilling into empty neighbours', () => {
  /** Vertical gridline segments [y0, y1] at device x (gridline stroke only). */
  const verticalAt = (ops: PaintOp[], x: number) => {
    const out: Array<[number, number]> = []
    for (let i = 0; i < ops.length - 1; i++) {
      const a = ops[i], b = ops[i + 1]
      if (a.op === 'moveTo' && b.op === 'lineTo' && a.strokeStyle === '#d0d0d0' && (a.args as number[])[0] === x && (b.args as number[])[0] === x) out.push([(a.args as number[])[1], (b.args as number[])[1]])
    }
    return out
  }
  const covers = (segs: Array<[number, number]>, y: number) => segs.some(([a, b]) => Math.min(a, b) < y && y < Math.max(a, b))
  // A1 spills right across B and C (30 glyphs x 7px = 210px); A2 is blocked by B2;
  // D3 is right-aligned and spills left across C.
  const cells = () => sheetOf([
    { row: 0, col: 0, value: 'x'.repeat(30) },
    { row: 1, col: 0, value: 'y'.repeat(30) }, { row: 1, col: 1, value: 'b' },
    { row: 2, col: 3, value: 'z'.repeat(14), style: style({ horizontal: 'right' }) },
  ])
  const check = (ops: PaintOp[]) => {
    // Row 1 (y 0..20): the B|C and C|D lines vanish under the spill, D|E stays (text ends at 213).
    expect(covers(verticalAt(ops, 64.5), 10)).toBe(false)
    expect(covers(verticalAt(ops, 128.5), 10)).toBe(false)
    expect(covers(verticalAt(ops, 256.5), 10)).toBe(true)
    // Row 2: blocked by B2, so A|B keeps its line; rows elsewhere keep theirs.
    expect(covers(verticalAt(ops, 64.5), 30)).toBe(true)
    expect(covers(verticalAt(ops, 128.5), 30)).toBe(true)
    // Row 3: right-aligned D3 (98px) spills left over the C|D line only.
    expect(covers(verticalAt(ops, 192.5), 50)).toBe(false)
    expect(covers(verticalAt(ops, 128.5), 50)).toBe(true)
  }
  it('static path', () => {
    const { ctx, ops } = recordingContext()
    renderSheet(cells(), ctx, grid([64, 64, 64, 64, 64], [20, 20, 20]))
    check(ops)
  })
  it('interactive viewport path', () => {
    const sheet = cells()
    const geometry = buildSheetGeometry(sheet)
    if (geometry.kind !== 'ready') throw new Error('geometry')
    const { ctx, ops } = recordingContext()
    const result = renderSheetViewport(sheet, ctx, geometry, { box: { x: 0, y: 0, width: 320, height: 60 }, zoom: 1, deviceScale: 1 })
    expect(result.kind).toBe('detail')
    check(ops)
  })
})

describe('#7 cell number formats: scientific, fractions and colour sections (Excel display)', () => {
  const sheet = sheetOf([])
  const shown = (value: number | string, formatCode: string, numFmtId = 164) => cellDisplayText(sheet, { value, style: style({ numFmtId, formatCode }) })
  it.each([
    [123456789, '0.00E+00', '1.23E+08'],
    [123456789, '##0.0E+0', '123.5E+6'],
    [0.000123, '0.00E+00', '1.23E-04'],
    [0, '0.00E+00', '0.00E+00'],
    [-123456789, '0.00E+00', '-1.23E+08'],
    [99999, '0.0E+00', '1.0E+05'],
    [1234, '0.00E-00', '1.23E03'],
  ] as const)('scientific %s with %s shows %s', (value, code, text) => expect(shown(value, code)).toBe(text))
  it.each([
    [1.75, '# ?/?', '1 3/4'],
    [0.5, '?/?', '1/2'],
    [3.14159, '# ??/??', '3 14/99'],
    [0.3333, '# ??/??', '  1/3 '],
    [2.625, '# ?/8', '2 5/8'],
    [2.3, '# ?/4', '2 1/4'],
    [5, '# ?/?', '5    '],
    [-1.75, '# ?/?', '-1 3/4'],
  ] as const)('fraction %s with %s shows %j', (value, code, text) => expect(shown(value, code)).toBe(text))
  it('selects sections like Excel without doubling the authored minus', () => {
    const code = '"$"#,##0.00;[Red]-"$"#,##0.00'
    expect(shown(-1234.5, code)).toBe('-$1,234.50')
    expect(shown(1234.5, code)).toBe('$1,234.50')
    expect(shown(-5, '#,##0_);[Red](#,##0)')).toBe('(5)')
    expect(shown(5, '#,##0_);[Red](#,##0)')).toBe('5 ')
    expect(shown(5, '[$€-407]#,##0.00')).toBe('€5.00')
    expect(shown(5, '[$-409]0.0')).toBe('5.0')
    expect(shown(-3, 'General;[Blue]-General')).toBe('-3')
  })
  it('built-in ids 11, 12 and 48 render through the same layer', () => {
    expect(cellDisplayText(sheet, { value: 123456789, style: style({ numFmtId: 11 }) })).toBe('1.23E+08')
    expect(cellDisplayText(sheet, { value: 1.75, style: style({ numFmtId: 12 }) })).toBe('1 3/4')
    expect(cellDisplayText(sheet, { value: 123456789, style: style({ numFmtId: 48 }) })).toBe('123.5E+6')
  })
  it('paints the selected colour section, including negative and [Color n] sections', () => {
    const paintColor = (value: number | string, formatCode: string, color?: string) => {
      const { ctx, ops } = recordingContext()
      renderSheet(sheetOf([{ row: 0, col: 0, value, style: style({ numFmtId: 164, formatCode, color }) }]), ctx, grid([200], [20]))
      return ops.find(o => o.op === 'fillText')?.fillStyle
    }
    const code = '"$"#,##0.00;[Red]-"$"#,##0.00'
    expect(paintColor(-1234.5, code)).toBe('#FF0000')
    expect(paintColor(1234.5, code, 'FF00B050')).toBe('#00B050')
    expect(paintColor(7, '[Blue]0;[Red]-0')).toBe('#0000FF')
    expect(paintColor(7, '[Color10]0')).toBe('#008000')
    expect(paintColor('t', '0;0;0;[Magenta]@')).toBe('#FF00FF')
  })
  it('conditions stay diagnosed with the existing fallback', () => {
    const s2 = sheetOf([])
    expect(cellDisplayText(s2, { value: 3.14159, style: style({ numFmtId: 164, formatCode: '[>100]0.00' }) })).toBe('3.14')
    expect(s2.diagnostics?.some(d => d.feature === 'format-unsupported')).toBe(true)
  })
})
