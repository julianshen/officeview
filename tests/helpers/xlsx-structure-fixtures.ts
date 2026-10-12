/** Real XLSX packages carrying structural metadata (S3.3): DrawingML anchors,
 * saved/engine dynamic-array spills, fixed arrays, tables and giant merges.
 * Built on the shared SpreadsheetML package writer so the same bytes feed the
 * object parser and the paged importer. */
import { buildZip, type BuildEntry } from './zip-builder'
import { xlsxEntries } from './xlsx-package'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
export const EMU_PER_PX = 9525
export const DRAWING_PART = 'xl/drawings/drawing1.xml'

export const DYNAMIC_METADATA = '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:da="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><da:dynamicArrayProperties fDynamic="1"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'

const shape = (id: number) => `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="shape${id}"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${10 * EMU_PER_PX}" cy="${10 * EMU_PER_PX}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:sp>`
const marker = (col: number, row: number, colOffPx = 0, rowOffPx = 0) =>
  `<xdr:col>${col}</xdr:col><xdr:colOff>${colOffPx * EMU_PER_PX}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOffPx * EMU_PER_PX}</xdr:rowOff>`
/** Anchors (zero-based markers, sizes in CSS px). The object is a plain shape: placement is all that matters. */
export const oneCell = (col: number, row: number, wPx: number, hPx: number, id = 2, colOffPx = 0, rowOffPx = 0) =>
  `<xdr:oneCellAnchor><xdr:from>${marker(col, row, colOffPx, rowOffPx)}</xdr:from><xdr:ext cx="${wPx * EMU_PER_PX}" cy="${hPx * EMU_PER_PX}"/>${shape(id)}<xdr:clientData/></xdr:oneCellAnchor>`
export const twoCell = (from: [number, number], to: [number, number], id = 3) =>
  `<xdr:twoCellAnchor><xdr:from>${marker(from[0], from[1])}</xdr:from><xdr:to>${marker(to[0], to[1])}</xdr:to>${shape(id)}<xdr:clientData/></xdr:twoCellAnchor>`
export const absolute = (xPx: number, yPx: number, wPx: number, hPx: number, id = 4) =>
  `<xdr:absoluteAnchor><xdr:pos x="${xPx * EMU_PER_PX}" y="${yPx * EMU_PER_PX}"/><xdr:ext cx="${wPx * EMU_PER_PX}" cy="${hPx * EMU_PER_PX}"/>${shape(id)}<xdr:clientData/></xdr:absoluteAnchor>`
/** A top-level anchor wrapped in mc:AlternateContent (neither parser reads its placement). */
export const wrappedAnchor = (inner: string) => `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice Requires="a14">${inner}</mc:Choice></mc:AlternateContent>`

export interface StructureSpec {
  sheetData: string
  /** Worksheet children between sheetData and drawing (mergeCells, ...). */
  afterSheetData?: string
  beforeSheetData?: string
  /** wsDr children; omitted means no drawing part. */
  drawing?: string
  /** `<drawing r:id>` without a resolvable relationship. */
  danglingDrawing?: boolean
  /** Full `<table ...>` part bodies, attached through tableParts. */
  tables?: string[]
  /** Adds xl/metadata.xml so cm="1" marks dynamic arrays. */
  dynamic?: boolean
  fullCalcOnLoad?: boolean
}

export function structureBytes(spec: StructureSpec): Uint8Array {
  const rels: string[] = []
  const extra: BuildEntry[] = []
  let after = spec.afterSheetData ?? ''
  if (spec.drawing !== undefined || spec.danglingDrawing) {
    after += `<drawing xmlns:r="${R_NS}" r:id="rD1"/>`
    if (!spec.danglingDrawing) {
      rels.push(`<Relationship Id="rD1" Type="${R_NS}/drawing" Target="../drawings/drawing1.xml"/>`)
      extra.push({ name: DRAWING_PART, data: `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">${spec.drawing}</xdr:wsDr>` })
    }
  }
  if (spec.tables?.length) {
    after += `<tableParts count="${spec.tables.length}">${spec.tables.map((_, i) => `<tablePart xmlns:r="${R_NS}" r:id="rT${i + 1}"/>`).join('')}</tableParts>`
    spec.tables.forEach((body, i) => {
      rels.push(`<Relationship Id="rT${i + 1}" Type="${R_NS}/table" Target="../tables/table${i + 1}.xml"/>`)
      extra.push({ name: `xl/tables/table${i + 1}.xml`, data: `<?xml version="1.0"?>${body}` })
    })
  }
  if (rels.length) {
    extra.push({ name: 'xl/worksheets/_rels/sheet1.xml.rels', data: `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>` })
  }
  const entries = xlsxEntries({ sheetData: spec.sheetData, afterSheetData: after, beforeSheetData: spec.beforeSheetData, fullCalcOnLoad: spec.fullCalcOnLoad, extra })
  if (spec.dynamic) {
    const wbRels = entries.find(e => e.name === 'xl/_rels/workbook.xml.rels')!
    wbRels.data = String(wbRels.data).replace('</Relationships>', '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>')
    entries.push({ name: 'xl/metadata.xml', data: DYNAMIC_METADATA })
  }
  return buildZip(entries).bytes
}

/** Rows 1..n with numbers in A and B (n descending in A so a sort really permutes). */
export const dataRows = (n: number, extra: (r: number) => string = () => '', attrs: (r: number) => string = () => '') =>
  Array.from({ length: n }, (_, r) => `<row r="${r + 1}"${attrs(r)}><c r="A${r + 1}"><v>${n - r}</v></c><c r="B${r + 1}"><v>${r}</v></c>${extra(r)}</row>`).join('')
