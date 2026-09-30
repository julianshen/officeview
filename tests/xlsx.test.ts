import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx, parseRef } from '../src/xlsx/parse'
import { computeMetrics, formatValue, renderSheet } from '../src/xlsx/render'
import { buildXlsx } from '../src/testdata/ooxml-builders'

describe('xlsx parse', () => {
  test('parses refs', () => {
    expect(parseRef('A1')).toEqual([0, 0])
    expect(parseRef('BC23')).toEqual([22, 54])
  })

  test('parses cells, shared strings, styles, merges', async () => {
    const buf = await buildXlsx([
      {
        name: 'Data',
        rows: [
          { r: 1, cells: [{ ref: 'A1', t: 's', v: 0, style: 1 }, { ref: 'B1', v: 3.14159, style: 2 }] },
          { r: 2, cells: [{ ref: 'A2', t: 's', v: 1 }, { ref: 'B2', v: 42 }] },
        ],
        cols: '<col min="1" max="1" width="18" customWidth="1"/>',
        merges: ['A1:B1'],
      },
    ], ['Region', 'Total Sales'])
    const pkg = await OfficePackage.load(buf)
    const doc = await parseXlsx(pkg)
    expect(doc.sheets).toHaveLength(1)
    const sheet = doc.sheets[0]
    expect(sheet.name).toBe('Data')
    expect(sheet.merges).toEqual(['A1:B1'])
    expect(sheet.cols[0].widthChars).toBe(18)
    const a1 = sheet.rows[0].cells[0]
    expect(a1.value).toBe('Region')
    expect(a1.style?.bold).toBe(true)
    expect(a1.style?.fillColor).toBe('FFFFFF00')
    expect(a1.style?.borders?.left).toBe('thin')
    const b1 = sheet.rows[0].cells[1]
    expect(b1.value).toBeCloseTo(3.14159)
    expect(b1.style?.numFmtId).toBe(2)
    expect(sheet.rows[1].cells[0].value).toBe('Total Sales')
  })
})

describe('xlsx render', () => {
  test('formats numbers', () => {
    expect(formatValue(3.14159, 2)).toBe('3.14')
    expect(formatValue(0.42, 9)).toBe('42%')
    expect(formatValue(4500.5, 3)).toBe('4,501')
    expect(formatValue(45000, 14)).toMatch(/^\d{4}-\d{2}-\d{2}$/)
    expect(formatValue('text', 0)).toBe('text')
    expect(formatValue(true, 0)).toBe('TRUE')
  })

  test('computes metrics from custom widths', async () => {
    const buf = await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }], cols: '<col min="1" max="2" width="10"/>' }])
    const doc = await parseXlsx(await OfficePackage.load(buf))
    const m = computeMetrics(doc.sheets[0])
    expect(m.colWidthsPx[0]).toBe(Math.round(10 * 7 + 5))
    expect(m.colWidthsPx[1]).toBe(m.colWidthsPx[0])
  })

  test('paints a real canvas grid', async () => {
    const buf = await buildXlsx([
      {
        name: 'Data',
        rows: [
          { r: 1, cells: [{ ref: 'A1', t: 's', v: 0, style: 1 }] },
          { r: 2, cells: [{ ref: 'A2', v: 3.14159, style: 2 }] },
        ],
      },
    ], ['Header Cell'])
    const doc = await parseXlsx(await OfficePackage.load(buf))
    const sheet = doc.sheets[0]
    const m = computeMetrics(sheet)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')
    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m)
    // yellow header fill present at top-left region
    const px = ctx.getImageData(4, 4, 1, 1).data
    expect(px[0]).toBe(255)
    expect(px[1]).toBe(255)
    expect(px[2]).toBe(0)
    // dark text pixels exist in header band (blue bold text on yellow)
    const band = ctx.getImageData(0, 0, m.widthPx, m.rowHeightsPx ? 20 : 20)
    let blue = 0
    for (let i = 0; i < band.data.length; i += 4) {
      if (band.data[i + 2] > 150 && band.data[i] < 100) blue++
    }
    expect(blue).toBeGreaterThan(10)
  })
})

describe('xlsx grid sizing limits', () => {
  test('ignores an implausibly wide <col> declaration (found via corpus poi-56295.xlsx)', async () => {
    // declares 1025 columns but only holds data in A1:C1
    const buf = await buildXlsx([
      { name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }, { ref: 'C1', v: 2 }] }] },
    ])
    // patch the fixture: widen the declared column range
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(buf)
    let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    xml = xml.replace('<cols>', '<cols>').replace('</cols>', '<col min="1" max="1025" width="8.5"/></cols>')
    zip.file('xl/worksheets/sheet1.xml', xml)
    const patched = await zip.generateAsync({ type: 'uint8array' })

    const doc = await parseXlsx(await OfficePackage.load(patched))
    const m = computeMetrics(doc.sheets[0])
    // sized to the used range (C), not 1025 declared columns
    expect(m.colWidthsPx.length).toBeLessThanOrEqual(4)
    expect(m.widthPx).toBeLessThan(2000)
  })

  test('keeps a modestly wider declared range (real empty-column formatting)', async () => {
    const buf = await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }], cols: '<col min="1" max="2" width="10"/>' }])
    const doc = await parseXlsx(await OfficePackage.load(buf))
    const m = computeMetrics(doc.sheets[0])
    expect(m.colWidthsPx).toHaveLength(2)
  })
})
