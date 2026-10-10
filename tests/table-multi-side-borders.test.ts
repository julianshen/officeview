import { describe, test, expect } from 'vitest'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide } from '../src/pptx/render'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages, type MeasureFn } from '../src/docx/layout'
import { buildPptx, buildDocx, type PptxShapeSpec } from '../src/testdata/ooxml-builders'

interface RecordedStroke {
  points: Array<{ x: number; y: number; cmd: 'moveTo' | 'lineTo' }>
  strokeStyle: any
  lineWidth: number
  lineDash: number[]
}

function recordStrokes(ctx: any): RecordedStroke[] {
  const strokes: RecordedStroke[] = []
  let points: Array<{ x: number; y: number; cmd: 'moveTo' | 'lineTo' }> = []
  const origMoveTo = ctx.moveTo.bind(ctx)
  const origLineTo = ctx.lineTo.bind(ctx)
  const origStroke = ctx.stroke.bind(ctx)
  const origBeginPath = ctx.beginPath.bind(ctx)

  ctx.beginPath = function() {
    points = []
    return origBeginPath()
  }
  ctx.moveTo = function(x: number, y: number) {
    points.push({ x, y, cmd: 'moveTo' })
    return origMoveTo(x, y)
  }
  ctx.lineTo = function(x: number, y: number) {
    points.push({ x, y, cmd: 'lineTo' })
    return origLineTo(x, y)
  }
  ctx.stroke = function() {
    strokes.push({
      points: [...points],
      strokeStyle: ctx.strokeStyle,
      lineWidth: ctx.lineWidth,
      lineDash: typeof ctx.getLineDash === 'function' ? ctx.getLineDash() : [],
    })
    return origStroke()
  }
  return strokes
}

const measureFixed: MeasureFn = (text, style) => {
  return text.length * style.fontSizePt * 0.6 * (96 / 72)
}

describe('multi-side table cell borders', () => {
  test('parses separate per-side border styles (lnL, lnR, lnT, lnB, lnTlToBr, lnBlToTr) in DrawingML table cells', async () => {
    const bordersXml = `
      <a:lnL w="9525" cmpd="s"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL>
      <a:lnR w="19050" cmpd="s"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnR>
      <a:lnT w="28575" cmpd="s"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:lnT>
      <a:lnB w="38100" cmpd="s"><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:lnB>
      <a:lnTlToBr w="14287" cmpd="s"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:lnTlToBr>
      <a:lnBlToTr w="23812" cmpd="s"><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill></a:lnBlToTr>
    `
    const shape: PptxShapeSpec = {
      off: ['0', '0'],
      ext: ['1000000', '1000000'],
      table: {
        colWidths: ['1000000'],
        rows: [{
          cells: [{
            borders: bordersXml,
            paragraphs: [{ runs: [{ text: 'Cell' }] }],
          }],
        }],
      },
    }

    const doc = await parsePptx(await OfficePackage.load(await buildPptx([shape])))
    const cell = doc.slides[0].shapes[0].table!.rows[0].cells[0]

    expect(cell.drawingBorders).toBeDefined()
    expect(cell.drawingBorders!.left).toBeDefined()
    expect(cell.drawingBorders!.right).toBeDefined()
    expect(cell.drawingBorders!.top).toBeDefined()
    expect(cell.drawingBorders!.bottom).toBeDefined()
    expect(cell.drawingBorders!.tlToBr).toBeDefined()
    expect(cell.drawingBorders!.blToTr).toBeDefined()

    expect(cell.drawingBorders!.left?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } })
    expect(cell.drawingBorders!.left?.width).toBe(1)

    expect(cell.drawingBorders!.right?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 255, b: 0, a: 1 } })
    expect(cell.drawingBorders!.right?.width).toBe(2)

    expect(cell.drawingBorders!.top?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 0, b: 255, a: 1 } })
    expect(cell.drawingBorders!.top?.width).toBe(3)

    expect(cell.drawingBorders!.bottom?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 255, b: 0, a: 1 } })
    expect(cell.drawingBorders!.bottom?.width).toBe(4)

    expect(cell.drawingBorders!.tlToBr?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 0, b: 255, a: 1 } })
    expect(cell.drawingBorders!.tlToBr?.width).toBeCloseTo(1.5, 2)

    expect(cell.drawingBorders!.blToTr?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 255, b: 255, a: 1 } })
    expect(cell.drawingBorders!.blToTr?.width).toBeCloseTo(2.5, 2)
  })

  test('parses separate per-side border properties (w:top, w:bottom, w:left, w:right, diagonals) in DOCX table cells', async () => {
    const bordersXml1 = `
      <w:top w:val="single" w:sz="8" w:color="FF0000"/>
      <w:left w:val="double" w:sz="16" w:color="00FF00"/>
      <w:bottom w:val="dashed" w:sz="24" w:color="0000FF"/>
      <w:right w:val="dotted" w:sz="32" w:color="FFFF00"/>
      <w:tl2br w:val="single" w:sz="12" w:color="FF00FF"/>
      <w:tr2bl w:val="single" w:sz="20" w:color="00FFFF"/>
    `
    const bordersXml2 = `
      <w:left w:val="wave" w:sz="10" w:color="123456"/>
      <w:right w:val="thick" w:sz="18" w:color="654321"/>
    `
    const docxBuf = await buildDocx([], [{
      gridCols: ['5000', '5000'],
      rows: [{
        cells: [
          {
            borders: bordersXml1,
            paragraphs: [{ runs: [{ text: 'Cell 1' }] }],
          },
          {
            borders: bordersXml2,
            paragraphs: [{ runs: [{ text: 'Cell 2' }] }],
          },
        ],
      }],
    }])

    const doc = await parseDocx(await OfficePackage.load(docxBuf))
    const tbl = (doc.sections[0].blocks[0] as { kind: 'table'; table: any }).table
    const cell1 = tbl.rows[0].cells[0]
    const cell2 = tbl.rows[0].cells[1]

    expect(cell1.borders).toBeDefined()
    expect(cell1.borders!.top).toEqual({ style: 'single', widthPt: 1, color: 'FF0000' })
    expect(cell1.borders!.left).toEqual({ style: 'double', widthPt: 2, color: '00FF00' })
    expect(cell1.borders!.bottom).toEqual({ style: 'dashed', widthPt: 3, color: '0000FF' })
    expect(cell1.borders!.right).toEqual({ style: 'dotted', widthPt: 4, color: 'FFFF00' })
    expect(cell1.borders!.tl2br).toEqual({ style: 'single', widthPt: 1.5, color: 'FF00FF' })
    expect(cell1.borders!.tr2bl).toEqual({ style: 'single', widthPt: 2.5, color: '00FFFF' })

    expect(cell2.borders).toBeDefined()
    expect(cell2.borders!.left).toEqual({ style: 'wave', widthPt: 1.25, color: '123456' })
    expect(cell2.borders!.right).toEqual({ style: 'thick', widthPt: 2.25, color: '654321' })
  })

  test('table cell painter renders distinct strokes per cell edge with proper joins and offsets', async () => {
    // 1. PPTX rendering verification
    const pptxBordersXml = `
      <a:lnL w="9525" cmpd="s"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL>
      <a:lnR w="19050" cmpd="s"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnR>
      <a:lnT w="28575" cmpd="s"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:lnT>
      <a:lnB w="38100" cmpd="s"><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:lnB>
      <a:lnTlToBr w="14287" cmpd="s"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:lnTlToBr>
      <a:lnBlToTr w="23812" cmpd="s"><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill></a:lnBlToTr>
    `
    const shape: PptxShapeSpec = {
      off: ['0', '0'],
      ext: ['1905000', '952500'], // 200px x 100px
      table: {
        colWidths: ['1905000'],
        rows: [{
          h: '952500',
          cells: [{
            borders: pptxBordersXml,
            paragraphs: [{ runs: [{ text: 'PPTX Cell' }] }],
          }],
        }],
      },
    }
    const pptxDoc = await parsePptx(await OfficePackage.load(await buildPptx([shape])))
    const pptxCanvas = createCanvas(300, 200)
    const pptxCtx = pptxCanvas.getContext('2d')
    const pptxStrokes = recordStrokes(pptxCtx)
    renderSlide(pptxDoc.slides[0], pptxCtx as unknown as CanvasRenderingContext2D)

    // Verify diagonals were drawn
    const diagTlBr = pptxStrokes.find(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && p1.x < p2.x && p1.y < p2.y
    })
    expect(diagTlBr).toBeDefined()
    expect(diagTlBr!.lineWidth).toBeCloseTo(1.5, 1)

    const diagBlTr = pptxStrokes.find(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && p1.x < p2.x && p1.y > p2.y
    })
    expect(diagBlTr).toBeDefined()
    expect(diagBlTr!.lineWidth).toBeCloseTo(2.5, 1)

    // 2. DOCX rendering verification
    const docxBordersXml = `
      <w:top w:val="single" w:sz="24" w:color="0000FF"/>
      <w:bottom w:val="single" w:sz="32" w:color="FFFF00"/>
      <w:left w:val="dashed" w:sz="8" w:color="FF0000"/>
      <w:right w:val="dotted" w:sz="16" w:color="00FF00"/>
      <w:tl2br w:val="single" w:sz="12" w:color="FF00FF"/>
      <w:tr2bl w:val="single" w:sz="20" w:color="00FFFF"/>
    `
    const docxBuf = await buildDocx([], [{
      gridCols: ['2880'], // 2 inches = 144px
      rows: [{
        heightTwips: '1440', // 1 inch = 72px
        cells: [{
          borders: docxBordersXml,
          paragraphs: [{ runs: [{ text: 'DOCX Cell' }] }],
        }],
      }],
    }])
    const docxDoc = await parseDocx(await OfficePackage.load(docxBuf))
    const pages = layoutDocx(docxDoc, measureFixed)
    const docxCanvas = createCanvas(816, 1056)
    const docxCtx = docxCanvas.getContext('2d')
    const docxStrokes = recordStrokes(docxCtx)
    renderPages(pages, docxCtx as unknown as CanvasRenderingContext2D)

    // Check dashed left border
    const dashedStroke = docxStrokes.find(s => s.lineDash.length > 0 && s.strokeStyle === '#ff0000')
    expect(dashedStroke).toBeDefined()

    // Check dotted right border
    const dottedStroke = docxStrokes.find(s => s.lineDash.length > 0 && s.strokeStyle === '#00ff00')
    expect(dottedStroke).toBeDefined()

    // Check DOCX diagonal tl2br and tr2bl
    const docxDiagTlBr = docxStrokes.find(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && p1.x < p2.x && p1.y < p2.y && s.strokeStyle === '#ff00ff'
    })
    expect(docxDiagTlBr).toBeDefined()
    expect(docxDiagTlBr!.lineWidth).toBeCloseTo(2, 1) // 1.5pt * 4/3 = 2px

    const docxDiagTrBl = docxStrokes.find(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && p1.x > p2.x && p1.y < p2.y && s.strokeStyle === '#00ffff'
    })
    expect(docxDiagTrBl).toBeDefined()
    expect(docxDiagTrBl!.lineWidth).toBeCloseTo(3.33, 1) // 2.5pt * 4/3 = 3.33px
  })

  test('adjacent cell border collapse resolves conflicting widths and styles consistently', async () => {
    // 1. PPTX: test conflicting adjacent border resolution
    // Shape A: Left cell has 30000 EMU (~3.15px) red right border, Right cell has 10000 EMU (~1.05px) black left border
    const shapeA: PptxShapeSpec = {
      off: ['0', '0'],
      ext: ['1905000', '952500'],
      table: {
        colWidths: ['952500', '952500'],
        rows: [{
          h: '952500',
          cells: [
            {
              borders: '<a:lnR w="30000" cmpd="s"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnR>',
              paragraphs: [{ runs: [{ text: 'Left' }] }],
            },
            {
              borders: '<a:lnL w="10000" cmpd="s"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>',
              paragraphs: [{ runs: [{ text: 'Right' }] }],
            },
          ],
        }],
      },
    }
    const docA = await parsePptx(await OfficePackage.load(await buildPptx([shapeA])))
    const canvasA = createCanvas(300, 200)
    const ctxA = canvasA.getContext('2d')
    const strokesA = recordStrokes(ctxA)
    renderSlide(docA.slides[0], ctxA as unknown as CanvasRenderingContext2D)

    // The shared internal boundary at x ≈ 100px must resolve to the wider red border (30000 EMU = 3.15px)
    const sharedStrokeA = strokesA.find(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && Math.abs(p1.x - 100) < 5 && Math.abs(p2.x - 100) < 5 && p1.y !== p2.y
    })
    expect(sharedStrokeA).toBeDefined()
    expect(sharedStrokeA!.lineWidth).toBeCloseTo(30000 / 9525, 1)
    expect(sharedStrokeA!.strokeStyle).toBe('#ff0000')

    // Shape B: Left cell has 10000 EMU black right border, Right cell has 30000 EMU red left border
    const shapeB: PptxShapeSpec = {
      off: ['0', '0'],
      ext: ['1905000', '952500'],
      table: {
        colWidths: ['952500', '952500'],
        rows: [{
          h: '952500',
          cells: [
            {
              borders: '<a:lnR w="10000" cmpd="s"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnR>',
              paragraphs: [{ runs: [{ text: 'Left' }] }],
            },
            {
              borders: '<a:lnL w="30000" cmpd="s"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL>',
              paragraphs: [{ runs: [{ text: 'Right' }] }],
            },
          ],
        }],
      },
    }
    const docB = await parsePptx(await OfficePackage.load(await buildPptx([shapeB])))
    const canvasB = createCanvas(300, 200)
    const ctxB = canvasB.getContext('2d')
    const strokesB = recordStrokes(ctxB)
    renderSlide(docB.slides[0], ctxB as unknown as CanvasRenderingContext2D)

    // The shared internal boundary must ALSO resolve to the wider red border (30000 EMU)
    const sharedStrokeB = strokesB.find(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && Math.abs(p1.x - 100) < 5 && Math.abs(p2.x - 100) < 5 && p1.y !== p2.y
    })
    expect(sharedStrokeB).toBeDefined()
    expect(sharedStrokeB!.lineWidth).toBeCloseTo(30000 / 9525, 1)
    expect(sharedStrokeB!.strokeStyle).toBe('#ff0000')

    // 2. DOCX: test conflicting adjacent border resolution
    // Doc 1: Cell 0 has 3pt red right border, Cell 1 has 1pt black left border
    const docxBuf1 = await buildDocx([], [{
      gridCols: ['1440', '1440'],
      rows: [{
        heightTwips: '1440',
        cells: [
          {
            borders: '<w:right w:val="single" w:sz="24" w:color="FF0000"/>',
            paragraphs: [{ runs: [{ text: 'Cell 0' }] }],
          },
          {
            borders: '<w:left w:val="single" w:sz="8" w:color="000000"/>',
            paragraphs: [{ runs: [{ text: 'Cell 1' }] }],
          },
        ],
      }],
    }])
    const docxDoc1 = await parseDocx(await OfficePackage.load(docxBuf1))
    const pages1 = layoutDocx(docxDoc1, measureFixed)
    const docxCanvas1 = createCanvas(816, 1056)
    const docxCtx1 = docxCanvas1.getContext('2d')
    const docxStrokes1 = recordStrokes(docxCtx1)
    renderPages(pages1, docxCtx1 as unknown as CanvasRenderingContext2D)

    const cell0 = pages1[0].tables[0].rows[0].cells[0]
    const sharedX1 = pages1[0].tables[0].xPx + cell0.xPx + cell0.widthPx

    // The shared vertical boundary must resolve to the wider red 3pt (4px) border
    const sharedDocx1 = docxStrokes1.filter(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && Math.abs(p1.x - sharedX1) < 2 && Math.abs(p2.x - sharedX1) < 2 && p1.y !== p2.y
    })
    // Only one stroke for this boundary (collapsed)
    expect(sharedDocx1).toHaveLength(1)
    expect(sharedDocx1[0].strokeStyle).toBe('#ff0000')
    expect(sharedDocx1[0].lineWidth).toBe(4) // 3pt * 4/3

    // Doc 2: Cell 0 has 1pt black right border, Cell 1 has 3pt red left border
    const docxBuf2 = await buildDocx([], [{
      gridCols: ['1440', '1440'],
      rows: [{
        heightTwips: '1440',
        cells: [
          {
            borders: '<w:right w:val="single" w:sz="8" w:color="000000"/>',
            paragraphs: [{ runs: [{ text: 'Cell 0' }] }],
          },
          {
            borders: '<w:left w:val="single" w:sz="24" w:color="FF0000"/>',
            paragraphs: [{ runs: [{ text: 'Cell 1' }] }],
          },
        ],
      }],
    }])
    const docxDoc2 = await parseDocx(await OfficePackage.load(docxBuf2))
    const pages2 = layoutDocx(docxDoc2, measureFixed)
    const docxCanvas2 = createCanvas(816, 1056)
    const docxCtx2 = docxCanvas2.getContext('2d')
    const docxStrokes2 = recordStrokes(docxCtx2)
    renderPages(pages2, docxCtx2 as unknown as CanvasRenderingContext2D)

    const sharedDocx2 = docxStrokes2.filter(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && Math.abs(p1.x - sharedX1) < 2 && Math.abs(p2.x - sharedX1) < 2 && p1.y !== p2.y
    })
    expect(sharedDocx2).toHaveLength(1)
    expect(sharedDocx2[0].strokeStyle).toBe('#ff0000')
    expect(sharedDocx2[0].lineWidth).toBe(4)
  })

  test('DOCX merged-cell conflict resolves per segment along unequal-height neighbors', async () => {
    // Tall vMerge anchor (3pt red right border) beside two stacked cells with
    // 1pt black left borders: every shared-gridline stretch must show the
    // conflict winner (red), with no gaps and no double-drawn overlaps.
    const p = (t: string) => ({ runs: [{ text: t }] })
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([], [{
      gridCols: ['2000', '2000'],
      rows: [
        { cells: [
          { paragraphs: [p('A')], vMerge: 'restart', borders: '<w:right w:val="single" w:sz="24" w:color="FF0000"/>' },
          { paragraphs: [p('B')], borders: '<w:left w:val="single" w:sz="8" w:color="000000"/>' },
        ] },
        { cells: [
          { vMerge: 'continue' },
          { paragraphs: [p('C')], borders: '<w:left w:val="single" w:sz="8" w:color="000000"/>' },
        ] },
      ],
    }])))
    const pages = layoutDocx(doc, measureFixed)
    const table = pages[0].tables[0]
    const canvas = createCanvas(816, 1056)
    const ctx = canvas.getContext('2d') as any
    const strokes = recordStrokes(ctx)
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 816, 1056)
    renderPages(pages, ctx as unknown as CanvasRenderingContext2D)
    const gx = table.xPx + table.rows[0].cells[0].xPx + table.rows[0].cells[0].widthPx
    const shared = strokes.filter(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && Math.abs(p1.x - gx) < 3 && Math.abs(p2.x - gx) < 3 && p1.y !== p2.y
    })
    // One stroke per shared stretch (upper + lower), both the red winner.
    expect(shared).toHaveLength(2)
    for (const s of shared) {
      expect(s.strokeStyle).toBe('#ff0000')
      expect(s.lineWidth).toBe(4)
    }
    // Stretches tile the anchor edge contiguously: the first ends where the
    // second begins (within rounding), covering the full shared gridline.
    const spans = shared.map(s => [s.points[0].y, s.points[1].y].sort((a, b) => a - b) as [number, number])
      .sort((a, b) => a[0] - b[0])
    expect(Math.abs(spans[0][1] - spans[1][0])).toBeLessThan(2)
    const anchor = table.rows[0].cells[0]
    expect(spans[0][0]).toBeGreaterThanOrEqual(table.yPx + anchor.yPx - 2)
    expect(spans[1][1]).toBeLessThanOrEqual(table.yPx + anchor.yPx + anchor.heightPx + 2)
  })

  test('PPTX table-style diagonal paints with no direct cell override', async () => {
    // Style-level tl2br with no direct lnTlToBr must still render (theme
    // table styles define diagonals; cells need not repeat them).
    const shape: PptxShapeSpec = {
      off: ['0', '0'],
      ext: ['1905000', '952500'],
      table: {
        colWidths: ['952500', '952500'],
        styleId: 'DiagOnlyStyle',
        customStyleXml: `<a:tblStyle styleId="DiagOnlyStyle" styleName="DiagOnlyStyle">` +
          `<a:wholeTbl><a:tcStyle><a:tcBdr>` +
          `<a:tl2br><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:tl2br>` +
          `</a:tcBdr></a:tcStyle></a:wholeTbl>` +
          `</a:tblStyle>`,
        rows: [{
          h: '952500',
          cells: [
            { paragraphs: [{ runs: [{ text: 'A' }] }] },
            { paragraphs: [{ runs: [{ text: 'B' }] }] },
          ],
        }],
      },
    }
    const doc = await parsePptx(await OfficePackage.load(await buildPptx([shape])))
    const table = doc.slides[0].shapes[0].table!
    expect(table.styleBorders?.tl2br?.color).toBe('#FF0000')
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    const strokes = recordStrokes(ctx as any)
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    // Diagonal strokes run corner-to-corner (x and y both change).
    const diags = strokes.filter(s => {
      const p1 = s.points[0], p2 = s.points[1]
      return p1 && p2 && p1.x !== p2.x && p1.y !== p2.y
    })
    expect(diags.length).toBeGreaterThanOrEqual(2)
    for (const d of diags) expect(String(d.strokeStyle).toLowerCase()).toBe('#ff0000')
  })
})
