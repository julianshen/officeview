/**
 * Golden-image regression suite. Renders fixture documents through the real
 * pipeline and pixel-diffs against tests/goldens/*.png.
 *
 * First run auto-records missing goldens. Delete a golden to re-record it
 * after an intentional layout change. Set OFFICEVIEW_STRICT_GOLDEN=1 to fail
 * instead of auto-recording (e.g. in CI).
 */
import { describe, test, expect } from 'vitest'
import { existsSync, mkdirSync } from 'fs'
import { join, dirname } from 'path'
import { fileURLToPath } from 'url'
import { getPaintables } from '../src/render/paint'
import { diffBitmaps, savePng, loadPng, renderPaintables } from '../src/test/pixel-diff'
import { buildDocx, buildXlsx, buildPptx, type DocxParaSpec, type DocxTableSpec } from './ooxml-fixtures'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { parsePptx } from '../src/pptx/parse'
import { OfficePackage } from '../src/core/zip'

const GOLDEN_DIR = join(dirname(fileURLToPath(import.meta.url)), 'goldens')
const THRESHOLD = 8 // per-channel AA forgiveness
const MAX_RATIO = 0.005 // 0.5% of pixels may differ

const p = (text: string): DocxParaSpec => ({ runs: [{ text }] })

async function docxFixture() {
  const table: DocxTableSpec = {
    gridCols: ['4320', '4320'],
    borders: '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/>',
    rows: [
      { cells: [{ paragraphs: [{ align: 'center', runs: [{ text: 'Quarter', bold: true }] }], fill: 'FFCC00' }, { paragraphs: [{ runs: [{ text: 'Revenue', bold: true }] }], fill: 'D9D9D9' }] },
      { cells: [{ paragraphs: [p('Q1')] }, { paragraphs: [{ runs: [{ text: '$1,204.50', color: '0070C0' }] }] }] },
      { cells: [{ paragraphs: [p('Q2')] }, { paragraphs: [{ runs: [{ text: '$2,918.25', color: 'C00000' }] }] }] },
    ],
  }
  return parseDocx(await OfficePackage.load(await buildDocx([
    { align: 'center', runs: [{ text: 'Quarterly Report', bold: true, size: 56 }] },
    { runs: [{ text: 'Sales summary with ', }, { text: 'highlighted figures', bold: true }, { text: ' and a table below.' }] },
  ], [table])))
}

async function xlsxFixture() {
  return parseXlsx(await OfficePackage.load(await buildXlsx([
    {
      name: 'Budget',
      rows: [
        { r: 1, cells: [{ ref: 'A1', t: 's', v: 0, style: 1 }, { ref: 'B1', t: 's', v: 1, style: 1 }, { ref: 'C1', t: 's', v: 2, style: 1 }] },
        { r: 2, cells: [{ ref: 'A2', t: 's', v: 3 }, { ref: 'B2', v: 1250.5, style: 2 }, { ref: 'C2', v: 0.185, style: 3 }] },
        { r: 3, cells: [{ ref: 'A3', t: 's', v: 4 }, { ref: 'B3', v: 980.25, style: 2 }, { ref: 'C3', v: 0.42, style: 3 }] },
      ],
      cols: '<col min="1" max="1" width="12" customWidth="1"/><col min="2" max="3" width="10"/>',
    },
  ], ['Category', 'Amount', 'Share', 'Hardware', 'Marketing'])))
}

async function pptxInspectable() {
  return parsePptx(await OfficePackage.load(await buildPptx([
    {
      prst: 'rect',
      off: ['914400', '914400'],
      ext: ['3657600', '1828800'],
      fill: 'FFCC00',
      paragraphs: [{ align: 'ctr', runs: [{ text: 'Pixel Harness', b: true, sz: '4400' }] }],
    },
    { prst: 'ellipse', off: ['5486400', '914400'], ext: ['1828800', '1828800'] },
  ])))
}

const fixtures: Array<{ name: string; build: () => Promise<unknown> }> = [
  { name: 'docx-report', build: docxFixture },
  { name: 'xlsx-budget', build: xlsxFixture },
  { name: 'pptx-slides', build: pptxInspectable },
]

const strict = process.env.OFFICEVIEW_STRICT_GOLDEN === '1'

describe('golden pixel regression', () => {
  for (const { name, build } of fixtures) {
    test(`${name} matches golden`, async () => {
      const doc = await build()
      const paintables = await getPaintables(doc as never)
      const actuals = await renderPaintables(paintables as never)
      const results: string[] = []
      for (let i = 0; i < actuals.length; i++) {
        const goldenPath = join(GOLDEN_DIR, `${name}.${i}.png`)
        if (!existsSync(goldenPath)) {
          if (strict) throw new Error(`missing golden ${goldenPath} (strict mode) — record it first`)
          mkdirSync(GOLDEN_DIR, { recursive: true })
          await savePng(actuals[i], goldenPath)
          results.push(`${i}: recorded`)
          continue
        }
        const golden = await loadPng(goldenPath)
        const d = diffBitmaps(golden, actuals[i], THRESHOLD)
        results.push(`${i}: ${(d.ratio * 100).toFixed(3)}%`)
        expect(d.ratio, `unit ${i}: ${d.changedPixels.toLocaleString()} px changed — inspect tests/goldens/${name}.${i}.png vs actual`).toBeLessThanOrEqual(MAX_RATIO)
      }
      if (results.some((r) => r.endsWith('recorded'))) console.warn(`[golden] auto-recorded ${name}:`, results.join(' '))
    })
  }
})
