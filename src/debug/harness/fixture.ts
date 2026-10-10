/**
 * Fixture loading for the dev harness page. The real F1..F8 fixtures arrive
 * with S0.2; until then the page builds a tiny in-page XLSX placeholder, which
 * is recorded as adapter `xlsx-placeholder` and never as a spec fixture.
 */
import JSZip from 'jszip'

export const PLACEHOLDER_ROWS = 120
export const PLACEHOLDER_COLS = 10

export interface FixtureInfo {
  adapter: string
  bytes: number
  rows: number | null
  cols: number | null
  note?: string
}

export interface LoadedFixture {
  info: FixtureInfo
  grid: string[][] | null
}

/** A real minimal XLSX, generated deterministically in-page. */
async function buildPlaceholderXlsx(rows: number, cols: number): Promise<{ bytes: Uint8Array; grid: string[][] }> {
  const grid: string[][] = []
  const sheetRows: string[] = []
  for (let r = 0; r < rows; r++) {
    const row: string[] = []
    const cells: string[] = []
    for (let c = 0; c < cols; c++) {
      const value = r === 0 ? `Column ${c + 1}` : String((r * 31 + c * 17) % 997)
      row.push(value)
      cells.push(`<c r="${columnName(c)}${r + 1}" t="inlineStr"><is><t>${value}</t></is></c>`)
    }
    grid.push(row)
    sheetRows.push(`<row r="${r + 1}">${cells.join('')}</row>`)
  }
  const sheet =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<sheetData>${sheetRows.join('')}</sheetData></worksheet>`
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Placeholder" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  zip.file('xl/worksheets/sheet1.xml', sheet)
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  return { bytes, grid }
}

function columnName(index: number): string {
  let name = ''
  let n = index
  while (n >= 0) {
    name = String.fromCharCode(65 + (n % 26)) + name
    n = Math.floor(n / 26) - 1
  }
  return name
}

export const fixtureLoader = {
  async loadPlaceholder(): Promise<LoadedFixture> {
    const { bytes, grid } = await buildPlaceholderXlsx(PLACEHOLDER_ROWS, PLACEHOLDER_COLS)
    return {
      info: {
        adapter: 'xlsx-placeholder',
        bytes: bytes.byteLength,
        rows: PLACEHOLDER_ROWS,
        cols: PLACEHOLDER_COLS,
        note: 'tiny synthetic XLSX generated in-page; real fixtures arrive with S0.2',
      },
      grid,
    }
  },
  async load(url: string): Promise<LoadedFixture> {
    const response = await fetch(url)
    if (!response.ok) throw new Error(`fixture fetch failed with status ${response.status}`)
    const buffer = await response.arrayBuffer()
    return this.loadBuffer(new Uint8Array(buffer))
  },
  async loadBuffer(bytes: Uint8Array): Promise<LoadedFixture> {
    const isZip = bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b
    return {
      info: {
        adapter: isZip ? 'xlsx' : 'unknown',
        bytes: bytes.byteLength,
        rows: null,
        cols: null,
        note: isZip
          ? 'source bytes recorded; parsing/rendering awaits the S1 importer and S5 viewport'
          : 'unrecognized fixture format',
      },
      grid: null,
    }
  },
}
