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
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { MAX_COLS } from '../src/xlsx/view/coordinates'
import { defaultColumnWidthPx, maxDigitWidthPx } from '../src/xlsx/columnWidth'
import { readFileSync } from 'node:fs'
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
    measureText: (text: unknown) => { record('measureText', [text]); return { width: glyph(String(text), String(state.font)) } },
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

describe('#4 border geometry is identical in the static and interactive paths', () => {
  // Interactive bodies are laid out at the origin and replayed under a
  // translation, so no clamp may depend on the box sitting at x/y = 0.
  const xs = (ops: PaintOp[]) => ops.filter((o, i) => o.op === 'moveTo' && ops[i + 1]?.op === 'lineTo' && o.strokeStyle !== '#d0d0d0').map(o => o.args as number[])
  const sheet = () => sheetOf([{ row: 1, col: 1, value: 'b', style: style({ borders: { left: 'double', top: 'double', right: 'thick' } }) }])
  it('double straddles the grid line and thick centres on it in both paths', () => {
    const fixed = recordingContext()
    renderSheet(sheet(), fixed.ctx, grid([64, 64, 64], [20, 20, 20]))
    const geometry = buildSheetGeometry(sheet())
    if (geometry.kind !== 'ready') throw new Error('geometry')
    // The recording ctx ignores translate, so interactive strokes stay box-local.
    const live = recordingContext()
    renderSheetViewport(sheet(), live.ctx, geometry, { box: { x: 0, y: 0, width: 192, height: 60 }, zoom: 1, deviceScale: 1 })
    expect(xs(fixed.ops).filter(([x]) => x < 70).map(([x]) => x)).toEqual(expect.arrayContaining([63.5, 65.5]))
    // Box-local interactive coordinates: the same strokes relative to the cell.
    const local = xs(live.ops).map(([x, y]) => [x + 64, y + 20])
    expect(local.filter(([x]) => x < 70).map(([x]) => x)).toEqual(expect.arrayContaining([63.5, 65.5]))
    expect(local.some(([x]) => x === 128.5)).toBe(true)
    expect(xs(fixed.ops).some(([x]) => x === 128.5)).toBe(true)
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
  it('built-in elapsed and fractional-second ids 45-47 (openpyxl writes [h]:mm:ss as id 46)', () => {
    expect(cellDisplayText(sheet, { value: 1.5, style: style({ numFmtId: 46 }) })).toBe('36:00:00')
    expect(cellDisplayText(sheet, { value: 0.0423611111, style: style({ numFmtId: 45 }) })).toBe('01:00')
    expect(cellDisplayText(sheet, { value: 0.000706018, style: style({ numFmtId: 47 }) })).toBe('01:01.0')
    expect(shown(0.5000011574, 'hh:mm:ss.00')).toBe('12:00:00.10')
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
  it('conditional sections choose by condition, then the catch-all section', () => {
    const code = '[>100][Red]0;[<=0][Blue]0.0;0.00'
    expect(shown(150, code)).toBe('150')
    expect(shown(-2, code)).toBe('-2.0')
    expect(shown(5, code)).toBe('5.00')
    expect(shown(150.4, '[>=100]#,##0;0.0')).toBe('150')
    expect(shown(7, '[>=100]#,##0;0.0')).toBe('7.0')
    expect(shown(-7, '[>=100]#,##0;0.0')).toBe('-7.0')
    // A single conditional section the value fails: Excel fills the cell with #.
    expect(shown(3.14159, '[>100]0.00')).toMatch(/^#+$/)
  })
  it('conditional sections paint their colour', () => {
    const paintColor = (value: number) => {
      const { ctx, ops } = recordingContext()
      renderSheet(sheetOf([{ row: 0, col: 0, value, style: style({ numFmtId: 164, formatCode: '[>100][Red]0;[<=0][Blue]0.0;0.00' }) }]), ctx, grid([200], [20]))
      return ops.find(o => o.op === 'fillText')?.fillStyle
    }
    expect(paintColor(150)).toBe('#FF0000')
    expect(paintColor(-2)).toBe('#0000FF')
    expect(paintColor(5)).toBe('#000000')
  })
  it.each([
    [1.5, '[h]:mm:ss', '36:00:00'],
    [2.25, '[hh]:mm', '54:00'],
    [0.0423611111, '[mm]:ss', '61:00'],
    [0.000694444, '[ss]', '60'],
    [1.0000578704, '[h]:mm:ss.00', '24:00:05.00'],
    [0.0000116, '[h]:mm:ss.0', '0:00:01.0'],
    [-0.5, '[h]:mm', '-12:00'],
    [0.99999999, '[h]:mm', '24:00'],
  ] as const)('elapsed %s with %s shows %s', (value, code, text) => expect(shown(value, code)).toBe(text))
  it('unknown bracket tokens stay diagnosed with the existing fallback', () => {
    const s2 = sheetOf([])
    expect(cellDisplayText(s2, { value: 3.14159, style: style({ numFmtId: 164, formatCode: '[Purpleish]0.00' }) })).toBe('3.14')
    expect(s2.diagnostics?.some(d => d.feature === 'format-unsupported')).toBe(true)
  })
})

describe('#2/#14 one Calibri fallback stack and Excel-model #### decisions', () => {
  /** Arial-like advances (the common sans fallback when Calibri/Carlito are absent). */
  const arial = (text: string, font: string) => {
    const px = Number(/([\d.]+)pt/.exec(font)?.[1] ?? 10) * 96 / 72
    let em = 0
    for (const ch of text) em += /[0-9#]/.test(ch) ? 0.556 : ch === '-' ? 0.333 : ch === ':' || ch === ' ' ? 0.278 : 0.5
    return em * px
  }
  it('plain and laid-out cell text both paint with the explicit stack (no context-dependent default)', async () => {
    const { createCanvas } = await import('canvas')
    const ctx = createCanvas(64, 80).getContext('2d') as unknown as CanvasRenderingContext2D
    const used: Array<[string, string, string]> = []
    const measure = ctx.measureText.bind(ctx), fill = ctx.fillText.bind(ctx)
    ctx.measureText = (text: string) => { used.push(['measure', text, ctx.font]); return measure(text) }
    ctx.fillText = (text: string, x: number, y: number) => { used.push(['fill', text, ctx.font]); fill(text, x, y) }
    renderSheet(sheetOf([
      { row: 0, col: 0, value: 'plain' },
      { row: 1, col: 0, value: 'wrapped words here', style: style({ wrapText: true }) },
    ]), ctx, grid([64], [20, 60]))
    // Every measurement and paint of both cells (the wrapped one goes through
    // the shared text-body layout) resolves Calibri to the same explicit stack.
    const cellText = used.filter(([, text]) => /plain|wrapped|words|here/.test(text))
    expect(cellText.some(([op, text]) => op === 'fill' && /wrapped|words|here/.test(text))).toBe(true)
    expect(cellText.some(([op, text]) => op === 'fill' && text === 'plain')).toBe(true)
    // node-canvas normalises the shorthand; the family list must survive.
    for (const [, , font] of cellText) expect(font).toMatch(/Carlito/)
  })
  it('native-1904 A2: a date Excel shows in full at width 19.5 is not ####, and is fitted to the cell', () => {
    // 12pt "yyyy-mm-dd hh:mm:ss" in round(19.5 * 7 + 5) = 142px: Calibri needs ~135.5px
    // of the 136px interior; an Arial fallback would need ~148.6px.
    const { ctx, ops } = recordingContext(arial)
    const sheet = sheetOf([{ row: 0, col: 0, value: 43844.5, style: style({ numFmtId: 176, formatCode: 'yyyy\\-mm\\-dd\\ hh:mm:ss', fontSizePt: 12 }) }], { semantics: { dateSystem: '1904', unicode: { version: 1, source: 'workbook-default' }, locale: 'en-US', timeZone: 'UTC', epochNowMs: 0 } })
    renderSheet(sheet, ctx, grid([142], [20]))
    const painted = ops.find(o => o.op === 'fillText')!
    expect(painted.args[0]).toBe('2024-01-15 12:00:00')
    expect(arial(String(painted.args[0]), painted.font)).toBeLessThanOrEqual(136 + 1e-9)
  })
  it('styled B26 (width 18) shows its date with no shrink when the fallback fits', () => {
    const { ctx, ops } = recordingContext(arial)
    renderSheet(sheetOf([{ row: 0, col: 0, value: 45360.336805555555, style: style({ numFmtId: 164, formatCode: 'yyyy-mm-dd hh:mm', fontSizePt: 11 }) }]), ctx, grid([131], [20]))
    const painted = ops.find(o => o.op === 'fillText')!
    expect(painted.args[0]).toBe('2024-03-09 08:05')
    expect(painted.font).toMatch(/^11pt /)
  })
  it('a number Excel cannot fit either still shows ####', () => {
    const { ctx, ops } = recordingContext(arial)
    renderSheet(sheetOf([{ row: 0, col: 0, value: 123456789.123, style: style({ numFmtId: 4, fontSizePt: 11 }) }]), ctx, grid([33], [20]))
    expect(String(ops.find(o => o.op === 'fillText')!.args[0])).toMatch(/^#+$/)
  })
})

describe('indexed and theme colours (fills, fonts, borders; object and paged)', () => {
  const THEME = `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="T">
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
    <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
    <a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
    <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`
  const styles = (colors = '') => `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${colors}
    <fonts count="3"><font><sz val="11"/><color theme="1"/></font><font><sz val="11"/><color theme="0"/></font><font><sz val="11"/><color indexed="10"/></font></fonts>
    <fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
      <fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill>
      <fill><patternFill patternType="solid"><fgColor theme="0" tint="-4.9989318521683403E-2"/></patternFill></fill>
      <fill><patternFill patternType="solid"><fgColor indexed="13"/></patternFill></fill></fills>
    <borders count="2"><border/><border><left style="thin"><color indexed="17"/></left><right style="thin"><color theme="5"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color auto="1"/></bottom></border></borders>
    <cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="0"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0"/></cellXfs></styleSheet>`
  async function both(stylesXml: string, theme?: string) {
    const zip = await JSZip.loadAsync(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', t: 'str', v: 'a', style: 1 }, { ref: 'B1', t: 'str', v: 'b', style: 2 }, { ref: 'C1', t: 'str', v: 'c', style: 3 }] }] }]))
    zip.file('xl/styles.xml', stylesXml)
    if (theme) {
      zip.file('xl/theme/theme1.xml', theme)
      const rels = (await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).replace('</Relationships>', '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>')
      zip.file('xl/_rels/workbook.xml.rels', rels)
    }
    const doc = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const paged = await workbookFormatting(parseXml('<workbook><workbookPr/></workbook>'), stylesXml, theme)
    return { object: doc.cellStyles!, paged: paged.styles }
  }
  // Excel's swatches round its tint; float HLS lands within 2 per channel.
  const near = (actual: string | undefined, hex: string) => {
    expect(actual, `${actual} vs ${hex}`).toMatch(/^FF[0-9A-F]{6}$/)
    for (let i = 0; i < 3; i++) expect(Math.abs(parseInt(actual!.slice(2 + 2 * i, 4 + 2 * i), 16) - parseInt(hex.slice(2 * i, 2 * i + 2), 16)), `${actual} vs ${hex}`).toBeLessThanOrEqual(2)
  }
  it('resolves theme slots (0=lt1, 1=dk1, 4=accent1) with Excel HLS tint, and the legacy indexed palette', async () => {
    for (const table of Object.values(await both(styles(), THEME))) {
      expect(table[0].color).toBe('FF000000')
      expect(table[1].color).toBe('FFFFFFFF')
      near(table[1].fillColor, 'D9E1F2') // accent1, lighter 80%
      expect(table[1].borderColors).toEqual({ left: 'FF008000', right: 'FFED7D31' }) // indexed 17; accent2; 64/auto keep black
      expect(table[2].color).toBe('FFFF0000') // indexed 10
      near(table[2].fillColor, 'F2F2F2') // lt1 darker 5%
      expect(table[3].fillColor).toBe('FFFFFF00') // indexed 13
    }
  })
  it('falls back to the Office theme without a theme part and honours indexedColors overrides', async () => {
    const override = '<colors><indexedColors>' + Array.from({ length: 64 }, (_, i) => `<rgbColor rgb="${i === 10 ? 'FF123456' : 'FF000000'}"/>`).join('') + '</indexedColors></colors>'
    for (const table of Object.values(await both(styles(override)))) {
      near(table[1].fillColor, 'D9E1F2')
      expect(table[2].color).toBe('FF123456')
    }
  })
  it('paints a theme-tinted fill in the static path', async () => {
    const { object } = await both(styles(), THEME)
    const { ctx, ops } = recordingContext()
    renderSheet(sheetOf([{ row: 0, col: 0, value: 'x', style: object[1] }]), ctx, grid([64], [20]))
    expect(ops.some(o => o.op === 'fillRect' && typeof o.fillStyle === 'string' && /^#D[9A]E[1-3]F[1-3]$/.test(o.fillStyle))).toBe(true)
  })
})

describe('Excel column widths: max digit width, baseColWidth/defaultColWidth', () => {
  const withCols = (extra: Partial<XlsxSheet> = {}): XlsxSheet => ({ name: 'S', merges: [], mergeRanges: [], rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, styleIndex: 0, value: 'a' }, { ref: 'D1', row: 0, col: 3, styleIndex: 0, value: 'd' }] }],
    cols: [{ min: 0, max: 0, widthChars: 26 }, { min: 1, max: 1, widthChars: 9.140625 }], ...extra })
  it('converts a stored width with pixels = trunc((256w + trunc(128/mdw))/256 * mdw) (Calibri 11: mdw 7)', () => {
    // 26 -> 182 (was round(26*7+5) = 187), 9.140625 (Excel's 8.43 default) -> 64, unset -> 64.
    expect(computeMetrics(withCols()).colWidthsPx).toEqual([182, 64, 64, 64])
  })
  it('uses the workbook max digit width and the sheet default width in every path', () => {
    const sheet = withCols({ columnMetrics: { maxDigitWidthPx: 8, defaultWidthPx: 88 } })
    // mdw 8: 26 -> 208, 9.140625 -> 73; unset columns take the default.
    expect(computeMetrics(sheet).colWidthsPx).toEqual([208, 73, 88, 88])
    const resident = buildSheetGeometry(sheet)
    if (resident.kind !== 'ready') throw new Error('geometry')
    expect([1, 2, 3, 4].map(c => resident.metrics.cols.start(c))).toEqual([208, 281, 369, 457])
    const paged = buildGeometry(sheet.rows, sheet.cols, [], { columns: sheet.columnMetrics })
    expect([1, 2, 3, 4].map(c => paged.cols.start(c))).toEqual([208, 281, 369, 457])
    // The worker publishes only overrides: the default must survive in them.
    const remote = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64, overrides: paged.colOverrides })
    expect([1, 2, 3, 4, 1000].map(c => remote.start(c))).toEqual([208, 281, 369, 457, 208 + 73 + 88 * 998])
  })
  it('derives the default width from baseColWidth (rounded up to 8px) or defaultColWidth', () => {
    expect(defaultColumnWidthPx({}, 7)).toBe(64)
    expect(defaultColumnWidthPx({ baseColWidth: 10 }, 8)).toBe(88)
    expect(defaultColumnWidthPx({ baseColWidth: 10 }, 7)).toBe(80)
    expect(defaultColumnWidthPx({ defaultColWidth: 9.140625 }, 7)).toBe(64)
    expect(maxDigitWidthPx('Calibri', 11)).toBe(7)
    expect(maxDigitWidthPx('Arial', 10)).toBe(7)
    expect(maxDigitWidthPx('新細明體', 12)).toBe(8)
  })
  it('native-1904.xlsx (新細明體 12pt, baseColWidth 10, A width 19.5) parses to Excel widths in the object and paged paths', async () => {
    const bytes = new Uint8Array(readFileSync('tests/fixtures/xlsx-viewer/native-1904.xlsx'))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    expect(doc.sheets[0].columnMetrics).toEqual({ maxDigitWidthPx: 8, defaultWidthPx: 88 })
    expect(computeMetrics(doc.sheets[0]).colWidthsPx.slice(0, 2)).toEqual([156, 88])
    const zip = await JSZip.loadAsync(bytes)
    const paged = await workbookFormatting(parseXml(await zip.file('xl/workbook.xml')!.async('string')), await zip.file('xl/styles.xml')!.async('string'))
    expect(paged.maxDigitWidthPx).toBe(8)
  })
})
