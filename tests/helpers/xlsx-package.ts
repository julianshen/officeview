/** Programmatic SpreadsheetML packages for the S1.1b admission and entry tests. */
import { buildZip, type BuildEntry } from './zip-builder'

export const SS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const CT = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`
const ROOT_RELS = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
const WB_RELS = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`

export interface PackageSpec {
  sheetData?: string
  afterSheetData?: string
  beforeSheetData?: string
  sharedStrings?: string[]
  styles?: string
  fullCalcOnLoad?: boolean
  extra?: BuildEntry[]
}

export function xlsxEntries(spec: PackageSpec): BuildEntry[] {
  const entries: BuildEntry[] = [
    { name: '[Content_Types].xml', data: CT },
    { name: '_rels/.rels', data: ROOT_RELS },
    { name: 'xl/workbook.xml', data: `<?xml version="1.0"?><workbook xmlns="${SS_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>${spec.fullCalcOnLoad ? '<calcPr fullCalcOnLoad="1"/>' : ''}</workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', data: WB_RELS },
    { name: 'xl/worksheets/sheet1.xml', data: `<?xml version="1.0"?><worksheet xmlns="${SS_NS}">${spec.beforeSheetData ?? ''}<sheetData>${spec.sheetData ?? ''}</sheetData>${spec.afterSheetData ?? ''}</worksheet>` },
  ]
  if (spec.sharedStrings) {
    entries.push({ name: 'xl/sharedStrings.xml', data: `<?xml version="1.0"?><sst xmlns="${SS_NS}" uniqueCount="${spec.sharedStrings.length}">${spec.sharedStrings.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>` })
  }
  entries.push({ name: 'xl/styles.xml', data: `<?xml version="1.0"?><styleSheet xmlns="${SS_NS}">${spec.styles ?? '<cellXfs count="1"><xf numFmtId="0"/></cellXfs>'}</styleSheet>` })
  return [...entries, ...(spec.extra ?? [])]
}

/** Deflated entries (the builder default). Compression ratio alone never refuses a package. */
export function xlsxBytes(spec: PackageSpec): Uint8Array {
  return buildZip(xlsxEntries(spec)).bytes
}

export const repeat = (n: number, f: (i: number) => string) => { const out: string[] = []; for (let i = 0; i < n; i++) out.push(f(i)); return out.join('') }
export const metaRows = (n: number) => repeat(n, (i) => `<row r="${i + 1}"/>`)
export const colName = (c: number) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s }
export const valueCells = (n: number) => `<row r="1">${repeat(n, (i) => `<c r="${colName(i % 16384)}${Math.floor(i / 16384) + 1}"><v>1</v></c>`)}</row>`
export const formulaRows = (n: number, f: (i: number) => string, attrs = '', cached = false) =>
  repeat(n, (i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c><c r="B${i + 1}"${attrs}><f>${f(i)}</f>${cached ? '<v>1</v>' : ''}</c></row>`)

