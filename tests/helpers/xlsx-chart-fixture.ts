import JSZip from 'jszip'
import { OfficePackage } from '../../src/core/zip'
import { buildXlsx } from '../../src/testdata/ooxml-builders'

export const emu = (px: number) => px * 9525
export const NS = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
/** openpyxl's chart frame: an empty xfrm, placed only by its one-cell anchor. */
export const openpyxlFrame = (name = 'Chart 1', descr?: string) => `<xdr:oneCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${emu(320)}" cy="${emu(200)}"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="${name}"${descr ? ` descr="${descr}"` : ''}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`

export const CACHED_COLUMN = '<chartSpace><chart><title><tx><rich><p><r><t>Sales</t></r></p></rich></tx></title><plotArea><barChart><barDir val="col"/><grouping val="clustered"/><ser><idx val="0"/><tx><v>2024</v></tx><cat><strLit><ptCount val="2"/><pt idx="0"><v>A</v></pt><pt idx="1"><v>B</v></pt></strLit></cat><val><numLit><ptCount val="2"/><pt idx="0"><v>3</v></pt><pt idx="1"><v>5</v></pt></numLit></val></ser></barChart><catAx/><valAx/></plotArea><legend/></chart></chartSpace>'

export async function chartWorkbook(chartXml: string, options: { frame?: string; cells?: Array<{ ref: string; v: string | number; style?: number }>; patch?: (zip: JSZip) => Promise<void> | void } = {}): Promise<OfficePackage> {
  return OfficePackage.load(await chartWorkbookBytes(chartXml, options))
}
export async function chartWorkbookBytes(chartXml: string, options: { frame?: string; cells?: Array<{ ref: string; v: string | number; style?: number }>; patch?: (zip: JSZip) => Promise<void> | void } = {}): Promise<Uint8Array> {
  const cells = options.cells ?? [{ ref: 'A1', v: 1 }]
  const rows = new Map<number, Array<{ ref: string; v: string | number; t?: 'str'; style?: number }>>()
  for (const cell of cells) {
    const r = Number(/\d+/.exec(cell.ref)![0])
    rows.set(r, [...rows.get(r) ?? [], typeof cell.v === 'string' ? { ...cell, t: 'str' as const } : cell])
  }
  const base = await buildXlsx([{ name: 'Styled', rows: [...rows].sort((a, b) => a[0] - b[0]).map(([r, list]) => ({ r, cells: list })) }])
  const zip = await JSZip.loadAsync(base)
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  sheet = sheet.replace('</worksheet>', '<drawing r:id="rDraw"/></worksheet>')
  if (!sheet.includes('xmlns:r=')) sheet = sheet.replace('<worksheet', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
  zip.file('xl/worksheets/sheet1.xml', sheet)
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="rDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
  zip.file('xl/drawings/drawing1.xml', `<xdr:wsDr ${NS}>${options.frame ?? openpyxlFrame()}</xdr:wsDr>`)
  zip.file('xl/drawings/_rels/drawing1.xml.rels', '<Relationships><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>')
  zip.file('xl/charts/chart1.xml', chartXml)
  await options.patch?.(zip)
  return zip.generateAsync({ type: 'uint8array' })
}

/** Adds an xf with the given built-in number format; returns its style index (3). */
export async function addNumberFormat(zip: JSZip, numFmtId: number): Promise<void> {
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file('xl/styles.xml', styles.replace('<cellXfs count="3">', '<cellXfs count="4">').replace('</cellXfs>', `<xf numFmtId="${numFmtId}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>`))
}
/** Adds workbook defined names. */
export async function addDefinedNames(zip: JSZip, names: string): Promise<void> {
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', workbook.replace('</sheets>', `</sheets><definedNames>${names}</definedNames>`))
}
