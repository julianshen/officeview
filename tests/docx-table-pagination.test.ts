import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, type MeasureFn } from '../src/docx/layout'
import { buildDocx, type DocxTableSpec } from '../src/testdata/ooxml-builders'

const measureFixed: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.6 * (96 / 72)

const p = (text: string) => ({ runs: [{ text }] })

const BORDERS =
  '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/>'

/** US Letter (12240 x 15840 twips) with 1" margins => content bottom at 960 px. */
const PAGE_H_PX = 15840 / 15
const CONTENT_BOTTOM_PX = PAGE_H_PX - 96

function tallTable(rowCount: number, opts: { isHeader?: boolean } = {}): DocxTableSpec {
  return {
    gridCols: ['4320', '4320'],
    borders: BORDERS,
    rows: Array.from({ length: rowCount }, (_, i) => ({
      isHeader: opts.isHeader && i === 0,
      cells: [{ paragraphs: [p(`R${i}A`)] }, { paragraphs: [p(`R${i}B`)] }],
    })),
  }
}

async function layoutTableSpec(table: DocxTableSpec) {
  const doc = await parseDocx(await OfficePackage.load(await buildDocx([], [table])))
  return layoutDocx(doc, measureFixed)
}

describe('docx table row pagination', () => {
  test('table taller than a page splits at row boundaries', async () => {
    const pages = await layoutTableSpec(tallTable(80))
    expect(pages.length).toBeGreaterThan(1)

    // every source row's text is present exactly once, in order
    const texts = pages.flatMap((pg) => pg.lines.map((l) => l.segs.map((s) => s.text).join('')))
    for (let i = 0; i < 80; i++) {
      expect(texts).toContain(`R${i}A`)
      expect(texts).toContain(`R${i}B`)
    }

    // no row box crosses the bottom of its page's content area, and rows
    // never repeat across pages
    const seen = new Set<number>()
    for (const pg of pages) {
      for (const table of pg.tables) {
        for (const row of table.rows) {
          const bottom = table.yPx + row.yPx + row.heightPx
          expect(bottom).toBeLessThanOrEqual(CONTENT_BOTTOM_PX + 0.5)
          const firstText = pg.lines.find((l) => Math.abs(l.yPx - (table.yPx + row.yPx)) < 0.6)?.segs[0]?.text ?? ''
          const m = /^R(\d+)A$/.exec(firstText)
          if (m) {
            const idx = Number(m[1])
            expect(seen.has(idx)).toBe(false)
            seen.add(idx)
          }
        }
      }
    }
    expect(seen.size).toBe(80)
  })

  test('header row repeats on every continuation page', async () => {
    const pages = await layoutTableSpec(tallTable(80, { isHeader: true }))
    expect(pages.length).toBeGreaterThan(1)
    for (const pg of pages) {
      const table = pg.tables[0]
      expect(table).toBeDefined()
      const headerTexts = pg.lines
        .filter((l) => l.yPx < table.yPx + table.rows[0].heightPx)
        .map((l) => l.segs.map((s) => s.text).join(''))
      expect(headerTexts).toContain('R0A')
      expect(headerTexts).toContain('R0B')
    }
    // body rows still appear exactly once each
    const bodyFirsts = pages.flatMap((pg) =>
      pg.lines
        .filter((l) => /^R\d+A$/.test(l.segs[0]?.text ?? '') && l.segs[0].text !== 'R0A')
        .map((l) => l.segs[0].text),
    )
    expect(new Set(bodyFirsts).size).toBe(79)
  })

  test('a single row taller than the page still places (no infinite split)', async () => {
    // 1 row of 3" height inside a 1.5" table area is impossible only if the
    // row exceeds the page; use a row taller than a full page
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([], [
      {
        gridCols: ['9360'],
        rows: [{ heightTwips: '30000', heightRule: 'exact', cells: [{ paragraphs: [p('huge')] }] }],
      },
    ])))
    const pages = layoutDocx(doc, measureFixed)
    expect(pages).toHaveLength(1)
    expect(pages[0].tables[0].rows[0].heightPx).toBeCloseTo(30000 / 15, 0)
  })

  test('table starting mid-page uses only the remaining space', async () => {
    // two paragraphs then a table: the first page should hold paragraphs +
    // as many rows as fit, not overflow
    const doc = await parseDocx(await OfficePackage.load(await buildDocx(
      [p('intro one'), p('intro two')],
      [tallTable(70)],
    )))
    const pages = layoutDocx(doc, measureFixed)
    for (const pg of pages) {
      for (const table of pg.tables) {
        for (const row of table.rows) {
          expect(table.yPx + row.yPx + row.heightPx).toBeLessThanOrEqual(CONTENT_BOTTOM_PX + 0.5)
        }
      }
    }
    // table starts on page 1 (after the intro paragraphs) and continues
    expect(pages.length).toBeGreaterThan(1)
    expect(pages[0].tables.length).toBe(1)
    expect(pages[0].tables[0].yPx).toBeGreaterThan(96)
  })
})

describe('docx table cell vertical alignment', () => {
  test('vAlign center shifts text within a taller row', async () => {
    const table: DocxTableSpec = {
      gridCols: ['4320'],
      borders: BORDERS,
      rows: [
        { heightTwips: '1200', heightRule: 'exact', cells: [{ vAlign: 'top', paragraphs: [p('topCell')] }] },
        { heightTwips: '1200', heightRule: 'exact', cells: [{ vAlign: 'center', paragraphs: [p('midCell')] }] },
        { heightTwips: '1200', heightRule: 'exact', cells: [{ vAlign: 'bottom', paragraphs: [p('botCell')] }] },
      ],
    }
    const pages = await layoutTableSpec(table)
    const page = pages[0]
    const at = (text: string): number => {
      const line = page.lines.find((l) => l.segs[0]?.text === text)
      expect(line, `line for ${text}`).toBeDefined()
      return line!.yPx
    }
    const topY = at('topCell')
    const midY = at('midCell')
    const botY = at('botCell')
    // rows are 80px tall (1200 twips); centering/bottom push text down
    expect(midY).toBeGreaterThan(topY)
    expect(botY).toBeGreaterThan(midY)
    // bottom-aligned text bottom edge lands on the row's bottom
    const table0 = page.tables[0]
    const row3 = table0.rows[2]
    const bottomLine = page.lines.find((l) => l.segs[0]?.text === 'botCell')!
    // row boxes are table-relative; lines are page-absolute
    expect(bottomLine.yPx + bottomLine.heightPx).toBeLessThanOrEqual(table0.yPx + row3.yPx + row3.heightPx + 0.5)
  })
})
