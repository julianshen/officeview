import { describe, expect, test } from 'vitest'
import { createCanvas, type Canvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { buildXlsx } from '../src/testdata/ooxml-builders'

async function bigSheet() {
  const rows = []
  for (let r = 1; r <= 200; r++) {
    const cells = []
    for (let c = 0; c < 20; c++) cells.push({ ref: `${String.fromCharCode(65 + c)}${r}`, v: `R${r}C${c}` })
    rows.push({ r, cells })
  }
  const bytes = await buildXlsx([{ name: 'Big', rows, merges: ['B50:D60'] }])
  const doc = await parseXlsx(await OfficePackage.load(bytes))
  const sheet = doc.sheets[0]
  return { sheet, m: computeMetrics(sheet) }
}
const pixels = (canvas: Canvas) => {
  const ctx = canvas.getContext('2d')
  const d = ctx.getImageData(0, 0, canvas.width, canvas.height).data
  return { data: d, width: canvas.width, height: canvas.height }
}

describe('xlsx viewport culling', () => {
  test('viewport paint equals the matching full-sheet crop', async () => {
    const { sheet, m } = await bigSheet()
    const full = createCanvas(m.widthPx, m.heightPx)
    renderSheet(sheet, full.getContext('2d') as never, m, undefined, {})
    // Viewport over rows ~50-60 (row height default 15px => y 735..900).
    const vx = 0, vy = 735, vw = m.widthPx, vh = 165
    const view = createCanvas(vw, vh)
    renderSheet(sheet, view.getContext('2d') as never, m, undefined, {}, { x: vx, y: vy, width: vw, height: vh })
    const a = pixels(full)
    const b = pixels(view)
    expect([b.width, b.height]).toEqual([vw, vh])
    let diff = 0
    for (let y = 0; y < vh; y++) {
      for (let x = 0; x < vw; x++) {
        const i = (y * vw + x) * 4, j = ((vy + y) * a.width + (vx + x)) * 4
        for (let k = 0; k < 3; k++) if (Math.abs(b.data[i + k] - a.data[j + k]) > 0) { diff++; break }
      }
    }
    expect(diff).toBe(0)
  })
  test('merged range paints when its anchor sits above the viewport', async () => {
    const { sheet, m } = await bigSheet()
    // B50:D52 merge anchor (row 50, y~735) sits above a viewport starting at
    // row 53 (y 795): the range reaches in and must paint identically.
    const vx = 0, vy = 53 * 15, vw = m.widthPx, vh = 120
    const full = createCanvas(m.widthPx, m.heightPx)
    renderSheet(sheet, full.getContext('2d') as never, m, undefined, {})
    const view = createCanvas(vw, vh)
    renderSheet(sheet, view.getContext('2d') as never, m, undefined, {}, { x: vx, y: vy, width: vw, height: vh })
    const a = pixels(full)
    const b = pixels(view)
    let diff = 0
    for (let y = 0; y < vh; y++) {
      for (let x = 0; x < vw; x++) {
        const i = (y * vw + x) * 4, j = ((vy + y) * a.width + (vx + x)) * 4
        for (let k = 0; k < 3; k++) if (Math.abs(b.data[i + k] - a.data[j + k]) > 0) { diff++; break }
      }
    }
    expect(diff).toBe(0)
    // And the merged text is really in the window (not vacuously white).
    expect(Array.from(b.data).some(v => v < 250)).toBe(true)
  })
  test('omitted viewport paints the whole sheet as before', async () => {
    const { sheet, m } = await bigSheet()
    const a = createCanvas(400, 300), b = createCanvas(400, 300)
    renderSheet(sheet, a.getContext('2d') as never, m, undefined, {})
    renderSheet(sheet, b.getContext('2d') as never, m, undefined, {}, undefined)
    expect(Buffer.from(a.toBuffer())).toEqual(Buffer.from(b.toBuffer()))
  })
})
