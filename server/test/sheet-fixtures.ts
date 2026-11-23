/**
 * Stage 4a sheet fixtures (Task 1; every later 4a task adds to this file).
 *
 * `buildXlsx` (src/testdata/ooxml-builders.ts) writes cells, columns and merges only. Everything else a tile
 * test needs (drawings, conditional formatting, row heights, freeze panes, autoFilter, extra styles) is
 * patched into its output with the JSZip helpers below, which are exported for later fixtures.
 *
 * Every fixture's bytes are deterministic: `finishXlsx` rewrites every entry with one fixed date, so the
 * same call returns the same bytes in every run (JSZip otherwise stamps the wall clock).
 */
import JSZip from 'jszip'
import { buildXlsx, type XlsxCellSpec, type XlsxSheetSpec } from '../../src/testdata/ooxml-builders'
import { encodePng } from '../src/mark/png'

/** The fixed entry date of every fixture zip (JSZip stores DOS time; any constant works). */
export const FIXTURE_ZIP_DATE = new Date(Date.UTC(2026, 0, 1, 0, 0, 0))
/** EMU per CSS px (914400 EMU per inch, 96 px per inch). */
export const EMU_PER_PX = 9525

export const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
export const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const DRAWING_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const DRAWING_CT = 'application/vnd.openxmlformats-officedocument.drawing+xml'

/** CT_Worksheet child order (ECMA-376 Part 1 §18.3.1.99); `insertSheetElement` keeps it. */
export const WORKSHEET_ORDER: readonly string[] = Object.freeze([
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
  'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
  'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
  'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts',
  'extLst',
])

const sheetPath = (sheetIndex: number) => `xl/worksheets/sheet${sheetIndex + 1}.xml`

async function textOf(zip: JSZip, path: string): Promise<string> {
  const f = zip.file(path)
  if (!f) throw new Error(`sheet-fixtures: no part ${path}`)
  return f.async('string')
}

/** Patch an xlsx package: `edit` mutates the loaded zip; the result is re-generated deterministically. */
export async function patchXlsx(bytes: Uint8Array, edit: (zip: JSZip) => void | Promise<void>): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  await edit(zip)
  return finishXlsx(zip)
}

/** Re-generate a zip with every entry at FIXTURE_ZIP_DATE (deterministic bytes, sorted entry order). */
export async function finishXlsx(zip: JSZip): Promise<Uint8Array> {
  const out = new JSZip()
  const names = Object.keys(zip.files).filter((n) => !zip.files[n]!.dir).sort()
  for (const name of names) out.file(name, await zip.files[name]!.async('uint8array'), { date: FIXTURE_ZIP_DATE, createFolders: false })
  return out.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } })
}

/** Rewrite one worksheet's XML. */
export async function patchSheetXml(zip: JSZip, sheetIndex: number, edit: (xml: string) => string): Promise<void> {
  const path = sheetPath(sheetIndex)
  zip.file(path, edit(await textOf(zip, path)))
}

/** Rewrite xl/styles.xml. */
export async function patchStylesXml(zip: JSZip, edit: (xml: string) => string): Promise<void> {
  zip.file('xl/styles.xml', edit(await textOf(zip, 'xl/styles.xml')))
}

/** Insert `xml` (one `<name …>` element) into a worksheet at its CT_Worksheet position. */
export function insertSheetElement(sheetXml: string, name: string, xml: string): string {
  const at = WORKSHEET_ORDER.indexOf(name)
  if (at < 0) throw new Error(`sheet-fixtures: unknown worksheet element ${name}`)
  for (const later of WORKSHEET_ORDER.slice(at + 1)) {
    const m = new RegExp(`<${later}[\\s/>]`).exec(sheetXml)
    if (m) return sheetXml.slice(0, m.index) + xml + sheetXml.slice(m.index)
  }
  const end = sheetXml.lastIndexOf('</worksheet>')
  if (end < 0) throw new Error('sheet-fixtures: no </worksheet>')
  return sheetXml.slice(0, end) + xml + sheetXml.slice(end)
}

/** Add attributes to `<row r="N">` for each listed 1-based row (e.g. `{ ht: '40', customHeight: '1' }`). */
export function setRowAttrs(sheetXml: string, rows: Iterable<number>, attrs: Record<string, string>): string {
  const extra = Object.entries(attrs).map(([k, v]) => ` ${k}="${v}"`).join('')
  for (const r of rows) {
    const tag = `<row r="${r}">`
    if (!sheetXml.includes(tag)) throw new Error(`sheet-fixtures: row ${r} is not a bare <row r="${r}">`)
    sheetXml = sheetXml.replace(tag, `<row r="${r}"${extra}>`)
  }
  return sheetXml
}

/** Freeze `cols` columns and `rows` rows (a `sheetViews` pane). */
export function freezeXml(cols: number, rows: number): string {
  const topLeft = `${colName(cols)}${rows + 1}`
  const pane = rows > 0 && cols > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight'
  return `<sheetViews><sheetView workbookViewId="0"><pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${topLeft}" activePane="${pane}" state="frozen"/></sheetView></sheetViews>`
}

/** 0-based column index → letters (0 → A, 99 → CV). */
export function colName(c: number): string {
  let s = ''
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}

/**
 * Add a drawing part to worksheet `sheetIndex`: writes `xl/drawings/drawing<n>.xml` (the `<xdr:wsDr>` root
 * wraps `anchorsXml`), the sheet's relationship, the content type and the `<drawing r:id>` element.
 * Returns the drawing part path.
 */
export async function addDrawing(zip: JSZip, sheetIndex: number, anchorsXml: string): Promise<string> {
  let n = 1
  while (zip.file(`xl/drawings/drawing${n}.xml`)) n++
  const part = `xl/drawings/drawing${n}.xml`
  zip.file(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${REL_NS}">${anchorsXml}</xdr:wsDr>`)
  const relsPath = `xl/worksheets/_rels/sheet${sheetIndex + 1}.xml.rels`
  const existing = zip.file(relsPath) ? await textOf(zip, relsPath) : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
  let id = 1
  while (existing.includes(`Id="rIdD${id}"`)) id++
  const rel = `<Relationship Id="rIdD${id}" Type="${DRAWING_REL}" Target="../drawings/drawing${n}.xml"/>`
  zip.file(relsPath, existing.replace('</Relationships>', `${rel}</Relationships>`))
  const ct = await textOf(zip, '[Content_Types].xml')
  zip.file('[Content_Types].xml', ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${DRAWING_CT}"/></Types>`))
  await patchSheetXml(zip, sheetIndex, (xml) => {
    const withNs = xml.includes('xmlns:r=') ? xml : xml.replace(`<worksheet xmlns="${SPREADSHEET_NS}"`, `<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}"`)
    if (withNs === xml && !xml.includes('xmlns:r=')) throw new Error('sheet-fixtures: worksheet root not recognised')
    return insertSheetElement(withNs, 'drawing', `<drawing r:id="rIdD${id}"/>`)
  })
  return part
}

/** A rectangle shape (`<xdr:sp>`) for an anchor body. `fill: null` is `<a:noFill/>`; `rotDeg` sets `xfrm rot`. */
export function rectShapeXml(opts: { id: number; name: string; offPx: [number, number]; extPx: [number, number]; fill: string | null; line?: string | null; rotDeg?: number }): string {
  const [x, y] = opts.offPx.map((v) => Math.round(v * EMU_PER_PX)) as [number, number]
  const [cx, cy] = opts.extPx.map((v) => Math.round(v * EMU_PER_PX)) as [number, number]
  const rot = opts.rotDeg ? ` rot="${Math.round(opts.rotDeg * 60000)}"` : ''
  const fill = opts.fill === null ? '<a:noFill/>' : `<a:solidFill><a:srgbClr val="${opts.fill}"/></a:solidFill>`
  const line = opts.line === undefined ? '' : opts.line === null ? '<a:ln><a:noFill/></a:ln>' : `<a:ln w="12700"><a:solidFill><a:srgbClr val="${opts.line}"/></a:solidFill></a:ln>`
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${opts.id}" name="${opts.name}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm${rot}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}${line}</xdr:spPr></xdr:sp>`
}

/** A two-cell anchor over 0-based cells (from) .. (to); offsets into each cell in CSS px (default 0). */
export interface AnchorMarker { col: number; row: number; colOffPx?: number; rowOffPx?: number }
export function twoCellAnchorXml(from: AnchorMarker, to: AnchorMarker, body: string): string {
  const m = (p: AnchorMarker) => `<xdr:col>${p.col}</xdr:col><xdr:colOff>${Math.round((p.colOffPx ?? 0) * EMU_PER_PX)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${Math.round((p.rowOffPx ?? 0) * EMU_PER_PX)}</xdr:rowOff>`
  return `<xdr:twoCellAnchor><xdr:from>${m(from)}</xdr:from><xdr:to>${m(to)}</xdr:to>${body}<xdr:clientData/></xdr:twoCellAnchor>`
}

/** An absolute anchor at CSS px (x, y), size (w, h). */
export function absoluteAnchorXml(posPx: [number, number], extPx: [number, number], body: string): string {
  const [x, y] = posPx.map((v) => Math.round(v * EMU_PER_PX))
  const [cx, cy] = extPx.map((v) => Math.round(v * EMU_PER_PX))
  return `<xdr:absoluteAnchor><xdr:pos x="${x}" y="${y}"/><xdr:ext cx="${cx}" cy="${cy}"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`
}

/** Build a one-sheet workbook (sheet `S`), with text cells through shared strings, then patch it. */
async function sheetWorkbook(rows: XlsxSheetSpec['rows'], strings: string[], edit?: (zip: JSZip) => void | Promise<void>, extra: Partial<XlsxSheetSpec> = {}): Promise<Uint8Array> {
  return patchXlsx(await buildXlsx([{ name: 'S', rows, ...extra }], strings), async (zip) => { await edit?.(zip) })
}

/** Shared-string table builder: `s(text)` returns a `t: 's'` cell value index. */
function stringTable() {
  const list: string[] = [], index = new Map<string, number>()
  const s = (text: string) => { let i = index.get(text); if (i === undefined) { i = list.length; list.push(text); index.set(text, i) } return i }
  return { list, s }
}
const textCell = (st: ReturnType<typeof stringTable>, ref: string, text: string): XlsxCellSpec => ({ ref, t: 's', v: st.s(text) })

/**
 * 500 rows × 100 columns (A1:CV500) on sheet `S`. Mixed: a cell whose 1-based column is ≡ 1 (mod 3) holds the
 * text `R{r}C{c}` (1-based row and column); every other cell holds the number `r * 1000 + c`. Default sizes
 * (64 × 20 CSS px), so at zoom 1 the used range is 6400 × 10000 CSS px: 13 × 20 tiles at dpr 1, 25 × 40 at dpr 2.
 */
export async function grid500x100(): Promise<Uint8Array> {
  const st = stringTable()
  const rows = Array.from({ length: 500 }, (_, i) => {
    const r = i + 1
    return { r, cells: Array.from({ length: 100 }, (_, j): XlsxCellSpec => {
      const c = j + 1, ref = `${colName(j)}${r}`
      return c % 3 === 1 ? textCell(st, ref, `R${r}C${c}`) : { ref, v: r * 1000 + c }
    }) }
  })
  return sheetWorkbook(rows, st.list)
}

/** Rows whose height is authored (customHeight) in `millionRows`. */
export const MILLION_ROWS_BAND = Object.freeze({ first: 10, last: 20, heightPt: 30 })
/**
 * A1:D2000 populated (A: text `M{r}`, B–D: numbers), plus A1048576 = `LAST-ROW`, on sheet `S`. Rows 10–20
 * (1-based) have an authored height of 30 pt (40 CSS px), so the last row's y is not 20 × 1,048,575.
 */
export async function millionRows(): Promise<Uint8Array> {
  const st = stringTable()
  const rows: XlsxSheetSpec['rows'] = Array.from({ length: 2000 }, (_, i) => {
    const r = i + 1
    return { r, cells: [textCell(st, `A${r}`, `M${r}`), { ref: `B${r}`, v: r }, { ref: `C${r}`, v: r * 2.5 }, { ref: `D${r}`, v: r % 97 }] }
  })
  rows.push({ r: 1_048_576, cells: [textCell(st, 'A1048576', 'LAST-ROW')] })
  const band = Array.from({ length: MILLION_ROWS_BAND.last - MILLION_ROWS_BAND.first + 1 }, (_, i) => MILLION_ROWS_BAND.first + i)
  return sheetWorkbook(rows, st.list, (zip) => patchSheetXml(zip, 0, (xml) => setRowAttrs(xml, band, { ht: String(MILLION_ROWS_BAND.heightPt), customHeight: '1' })))
}

/** Sheet `S` with no cells. */
export async function blankSheet(): Promise<Uint8Array> {
  return sheetWorkbook([], [])
}

/** The date serials in `cfToday` A1:A10: 2026-11-19 .. 2026-11-28 (46345 .. 46354); A5 is 2026-11-23 (46349). */
export const CF_TODAY_SERIALS = Object.freeze(Array.from({ length: 10 }, (_, i) => 46345 + i))
/**
 * A1:A10 hold the dates CF_TODAY_SERIALS (style 3: numFmtId 14, m/d/yyyy), under one
 * `timePeriod today` rule (dxf 0: solid red fill). Which cell fills depends only on the injected clock.
 */
export async function cfToday(): Promise<Uint8Array> {
  const rows = CF_TODAY_SERIALS.map((v, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, v, style: 3 }] }))
  return sheetWorkbook(rows, [], async (zip) => {
    await patchStylesXml(zip, (xml) => xml
      .replace('<cellXfs count="3">', '<cellXfs count="4">')
      .replace('</cellXfs>', '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>'))
    await patchSheetXml(zip, 0, (xml) => insertSheetElement(xml, 'conditionalFormatting',
      '<conditionalFormatting sqref="A1:A10"><cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="today"><formula>FLOOR(A1,1)=TODAY()</formula></cfRule></conditionalFormatting>'))
  })
}

/**
 * The `TODAY()` formula twin of `cfToday` (Task 5, F2): the same A1:A10 dates under one `expression` rule
 * `A1=TODAY()` (dxf 0: solid red fill). It evaluates only when the CF evaluation context survives
 * `cfAlignClock`, so a clock that reaches the predicates but not the evaluation leaves every cell unfilled.
 */
export async function cfTodayFormula(): Promise<Uint8Array> {
  const rows = CF_TODAY_SERIALS.map((v, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, v, style: 3 }] }))
  return sheetWorkbook(rows, [], async (zip) => {
    await patchStylesXml(zip, (xml) => xml
      .replace('<cellXfs count="3">', '<cellXfs count="4">')
      .replace('</cellXfs>', '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>'))
    await patchSheetXml(zip, 0, (xml) => insertSheetElement(xml, 'conditionalFormatting',
      '<conditionalFormatting sqref="A1:A10"><cfRule type="expression" dxfId="0" priority="1"><formula>A1=TODAY()</formula></cfRule></conditionalFormatting>'))
  })
}

/** The unfilled box of `unfilledShapeOverText`: strictly inside B2 (x 64..128, y 20..40 CSS px) at x 72..120,
 * y 24..36, so its 1 pt outline (half width 0.67 px) stays clear of every neighbouring cell. */
const insideB2 = () => twoCellAnchorXml({ col: 1, row: 1, colOffPx: 8, rowOffPx: 4 }, { col: 1, row: 1, colOffPx: 56, rowOffPx: 16 },
  rectShapeXml({ id: 2, name: 'Unfilled', offPx: [72, 24], extPx: [48, 12], fill: null, line: '404040' }))

/**
 * B2 holds `COVERED-B2` (A1:D4 hold `CELL-r-c` around it) under an unfilled rectangle (`<a:noFill/>`, a
 * 1 pt outline) anchored two-cell strictly inside B2 (Task 5: it was A1..C3, which covered nine cells, while
 * S4a Tasks 5 and 15 expect B2 alone). Coverage ignores fill (S4 §6.2 step 4): B2 is the only covered cell.
 */
export async function unfilledShapeOverText(): Promise<Uint8Array> {
  const st = stringTable()
  const rows = [1, 2, 3, 4].map((r) => ({ r, cells: [0, 1, 2, 3].map((j) => textCell(st, `${colName(j)}${r}`, r === 2 && j === 1 ? 'COVERED-B2' : `CELL-${r}-${j + 1}`)) }))
  return sheetWorkbook(rows, st.list, async (zip) => { await addDrawing(zip, 0, insideB2()) })
}

/** B1:B3 of `unfilledShapeOverDataBar`; B2 (20) sits under the box. */
export const DATA_BAR_VALUES = Object.freeze([10, 20, 30])
/**
 * The data-bar variant of `unfilledShapeOverText` (Task 5): B1:B3 hold DATA_BAR_VALUES under one `dataBar`
 * rule (min..max), A1:A3 and C1:C3 hold text, and the same unfilled box sits strictly inside B2.
 */
export async function unfilledShapeOverDataBar(): Promise<Uint8Array> {
  const st = stringTable()
  const rows = [1, 2, 3].map((r) => ({ r, cells: [textCell(st, `A${r}`, `BAR-${r}-1`), { ref: `B${r}`, v: DATA_BAR_VALUES[r - 1]! }, textCell(st, `C${r}`, `BAR-${r}-3`)] }))
  return sheetWorkbook(rows, st.list, async (zip) => {
    await addDrawing(zip, 0, insideB2())
    await patchSheetXml(zip, 0, (xml) => insertSheetElement(xml, 'conditionalFormatting',
      '<conditionalFormatting sqref="B1:B3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>'))
  })
}

/**
 * Fix round 1 (a): two visible sheets. Sheet `S` is the data-bar variant's grid with B2 = 20 under the box strictly
 * inside B2; sheet `T` holds A1 = 1 under one `expression` rule `S!$B$2=20` (dxf 0: solid red fill). Under
 * `visible-only` S!B2 is covered, so the rule must read it blank and T!A1 stays unfilled; under `'full'` it fills.
 */
export async function crossSheetCoveredRef(): Promise<Uint8Array> {
  const st = stringTable()
  const s = [1, 2, 3].map((r) => ({ r, cells: [textCell(st, `A${r}`, `X-${r}-1`), { ref: `B${r}`, v: DATA_BAR_VALUES[r - 1]! }] }))
  return patchXlsx(await buildXlsx([{ name: 'S', rows: s }, { name: 'T', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }], st.list), async (zip) => {
    await addDrawing(zip, 0, insideB2())
    await patchStylesXml(zip, (xml) => xml.replace('</cellXfs>', '</cellXfs><dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>'))
    await patchSheetXml(zip, 1, (xml) => insertSheetElement(xml, 'conditionalFormatting',
      '<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>S!$B$2=20</formula></cfRule></conditionalFormatting>'))
  })
}

/**
 * Fix round 1 (b): two line sparklines over D1:D4 (numbers 1..4), located at B2 (under the box strictly inside B2,
 * so covered) and at C2 (uncovered, the control). A1 holds text so the sheet is not otherwise empty.
 */
export async function sparklineUnderShape(): Promise<Uint8Array> {
  const st = stringTable()
  const rows = [1, 2, 3, 4].map((r) => ({ r, cells: [...(r === 1 ? [textCell(st, 'A1', 'SPARK')] : []), { ref: `D${r}`, v: r }] }))
  const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', xm = 'http://schemas.microsoft.com/office/excel/2006/main'
  const line = (at: string) => `<x14:sparkline><xm:f>S!D1:D4</xm:f><xm:sqref>${at}</xm:sqref></x14:sparkline>`
  return sheetWorkbook(rows, st.list, async (zip) => {
    await addDrawing(zip, 0, insideB2())
    await patchSheetXml(zip, 0, (xml) => insertSheetElement(xml, 'extLst',
      `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${x14}"><x14:sparklineGroups xmlns:xm="${xm}"><x14:sparklineGroup displayEmptyCellsAs="gap"><x14:colorSeries rgb="FF376092"/><x14:sparklines>${line('B2')}${line('C2')}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`))
  })
}

/** The text of `spillUnderShape` A3: far wider than A3 (64 px), so it runs on into the empty B3 and C3. */
export const SPILL_TEXT = 'SPILL-A3 long text that overflows into the empty cells to its right'
/**
 * R8's spill case (Task 5): A3 holds SPILL_TEXT, aligned `left` or `right` (style 3), B3 and C3 are empty,
 * and a filled box sits strictly inside C3 (x 128..192, y 40..60 → box x 136..184, y 44..56). Left-aligned
 * text spills right, under the box: A3 is covered. Right-aligned text spills left (nothing there): it is not.
 */
export async function spillUnderShape(align: 'left' | 'right'): Promise<Uint8Array> {
  const st = stringTable()
  const rows = [{ r: 3, cells: [{ ...textCell(st, 'A3', SPILL_TEXT), style: 3 }] }]
  return sheetWorkbook(rows, st.list, async (zip) => {
    await patchStylesXml(zip, (xml) => xml
      .replace('<cellXfs count="3">', '<cellXfs count="4">')
      .replace('</cellXfs>', `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="${align}"/></xf></cellXfs>`))
    await addDrawing(zip, 0, twoCellAnchorXml({ col: 2, row: 2, colOffPx: 8, rowOffPx: 4 }, { col: 2, row: 2, colOffPx: 56, rowOffPx: 16 },
      rectShapeXml({ id: 2, name: 'OverC3', offPx: [136, 44], extPx: [48, 12], fill: 'FF8800' })))
  })
}

/**
 * S4 §6.2 step 4's rotated-box adversary: an absolute anchor at (0, 0) of 100 × 20 CSS px holding a filled
 * rectangle rotated 90° (`xfrm rot="5400000"`). Rotated about its centre (50, 10), the painted box spans
 * **x 40..60, y −40..60** (F19; the spec's "y 30..50" is a subset; y < 0 is off-sheet). A1:C6 hold
 * `ROT-r-c` (1-based, default 64 × 20 CSS px cells): A1, A2 and A3 (x 0..64, y 0..60) meet it; column B
 * (x ≥ 64) and rows 4+ (y ≥ 60) are clear by grid rect (their spill rects are the caller's to add).
 */
export async function rotatedBox(): Promise<Uint8Array> {
  const st = stringTable()
  const rows = Array.from({ length: 6 }, (_, i) => ({ r: i + 1, cells: [0, 1, 2].map((j) => textCell(st, `${colName(j)}${i + 1}`, `ROT-${i + 1}-${j + 1}`)) }))
  return sheetWorkbook(rows, st.list, async (zip) => {
    await addDrawing(zip, 0, absoluteAnchorXml([0, 0], [100, 20], rectShapeXml({ id: 2, name: 'Rotated', offPx: [0, 0], extPx: [100, 20], fill: '4472C4', rotDeg: 90 })))
  })
}

/**
 * A two-cell anchor with no `<xdr:to>`: its placement cannot be read, so the parse counts
 * `drawingAnchors.unresolved = 1` (fail-closed: every cell counts as blank for text answers). A1:B2 hold text.
 */
export async function unresolvedAnchor(): Promise<Uint8Array> {
  const st = stringTable()
  const rows = [1, 2].map((r) => ({ r, cells: [0, 1].map((j) => textCell(st, `${colName(j)}${r}`, `UNRES-${r}-${j + 1}`)) }))
  return sheetWorkbook(rows, st.list, async (zip) => {
    await addDrawing(zip, 0, `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>${rectShapeXml({ id: 2, name: 'NoTo', offPx: [0, 0], extPx: [100, 40], fill: 'FF8800' })}<xdr:clientData/></xdr:twoCellAnchor>`)
  })
}

/**
 * A1 = `=1+1` with saved value 7: saved values paint 7, a recompute paints 2. The workbook sets
 * `calcPr fullCalcOnLoad="1"` (Task 5, F1): without it the engine keeps a saved value, so the two policies
 * could not be told apart; with it the engine parse recomputes A1 to 2.
 */
export async function savedFormula(): Promise<Uint8Array> {
  return sheetWorkbook([{ r: 1, cells: [{ ref: 'A1', formula: '1+1', v: 7 }] }], [], async (zip) => {
    const wb = await textOf(zip, 'xl/workbook.xml')
    if (!wb.includes('</workbook>')) throw new Error('sheet-fixtures: no </workbook>')
    zip.file('xl/workbook.xml', wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>'))
  })
}

/** A1 = the value 7 (the saved-values twin of `savedFormula`). */
export async function plainSeven(): Promise<Uint8Array> {
  return sheetWorkbook([{ r: 1, cells: [{ ref: 'A1', v: 7 }] }], [])
}

/** `frozenMerges` (Task 6, F6): its merges, as authored. */
export const FROZEN_MERGES = Object.freeze(['A5:C6', 'D10:E20', 'B30:B31', 'H2:J2'])
/**
 * Task 6 (F6): frozen panes (2 columns, 3 rows; scrolling pane from C4) plus FROZEN_MERGES on sheet `S`, with
 * authored sizes on both axes: columns B:C are 20 characters wide and rows 7 and 25 are 30 pt high. A1:A40 hold
 * numbers so the used range is A1:A40 (J2 is inside a merge, not a cell).
 */
export async function frozenMerges(): Promise<Uint8Array> {
  const rows = Array.from({ length: 40 }, (_, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, v: i + 1 }] }))
  return sheetWorkbook(rows, [], async (zip) => {
    await patchSheetXml(zip, 0, (xml) => insertSheetElement(setRowAttrs(xml, [7, 25], { ht: '30', customHeight: '1' }), 'sheetViews', freezeXml(2, 3)))
  }, { cols: '<col min="2" max="3" width="20" customWidth="1"/>', merges: [...FROZEN_MERGES] })
}

/** `autoFilterCollapsed` (Task 6, F6): the declared autoFilter range and the two hidden rows (1-based). */
export const AUTO_FILTER = Object.freeze({ ref: 'A1:B8', hiddenInside: 4, hiddenOutside: 10, lastRow: 11 })
/**
 * Task 6 (F6): A1:B11 hold `H-r-c` text (row 1 is the header), under a worksheet `<autoFilter ref="A1:B8"/>`.
 * Row 4 (inside the filter range) and row 10 (outside it, but inside the used range) are hidden, as a saved
 * filter or the author leaves them.
 */
export async function autoFilterCollapsed(): Promise<Uint8Array> {
  const st = stringTable()
  const rows = Array.from({ length: AUTO_FILTER.lastRow }, (_, i) => ({ r: i + 1, cells: [0, 1].map((j) => textCell(st, `${colName(j)}${i + 1}`, `H-${i + 1}-${j + 1}`)) }))
  return sheetWorkbook(rows, st.list, async (zip) => {
    await patchSheetXml(zip, 0, (xml) => insertSheetElement(setRowAttrs(xml, [AUTO_FILTER.hiddenInside, AUTO_FILTER.hiddenOutside], { hidden: '1' }), 'autoFilter', `<autoFilter ref="${AUTO_FILTER.ref}"/>`))
  })
}

/**
 * Task 6 (F6): `n` two-column merges on sheet `S`, and no cells. By default one per row (`A{r}:B{r}`, r = 1..n);
 * with `perRow`, `perRow` side by side per row (merge i at row ⌊i / perRow⌋ + 1, columns 2(i mod perRow) and the
 * next), which keeps a few thousand merges inside an 8192 px static unit. 20,000 chunk a layout into several parts;
 * 70,000 are over the 65,536-merge cap (R14).
 */
export async function manyMerges(n: number, o: { perRow?: number } = {}): Promise<Uint8Array> {
  const per = o.perRow ?? 1
  const merges = Array.from({ length: n }, (_, i) => {
    const r = Math.floor(i / per) + 1, c = 2 * (i % per)
    return `${colName(c)}${r}:${colName(c + 1)}${r}`
  })
  return sheetWorkbook([], [], undefined, { merges })
}

/** The strings `sentinels()` hides from the wire (Task 15): none may occur in any decoded server→client text. */
export const SENTINELS = Object.freeze({
  /** A1's formula is `=LEN("SENTINEL_F")` (saved value 10). */
  formula: 'SENTINEL_F', formulaText: '=LEN', formulaSaved: 10,
  /** B1 holds this number under number format `0` (it paints `1235`). */
  number: '1234.5678',
  /** A workbook-scope defined name (`S!$B$1`). */
  definedName: 'SENTINEL_N',
  /** The comment on C1. */
  comment: 'SENTINEL_C',
  /** The second sheet is hidden; its name and its A1. */
  hiddenSheet: 'SENTINEL_HS', hiddenCell: 'SENTINEL_HC',
  /** A visible text cell (A2): a cell value never reaches the client either. */
  visibleText: 'SENTINEL_VIS',
  /** The hidden picture's `cNvPr name`. */
  drawingName: 'SENTINEL_D',
})

/** The hidden picture of `sentinels()`: a 16 × 16 RGBA PNG of deterministic noise (xorshift32), 1,024 pixel bytes that
 * deflate cannot shrink, so its bytes (and its IDAT) are unique on any wire. */
export function sentinelPng(): Uint8Array {
  const rgba = new Uint8Array(16 * 16 * 4)
  let x = 0x9e3779b9
  for (let i = 0; i < rgba.length; i++) { x ^= x << 13; x >>>= 0; x ^= x >>> 17; x ^= x << 5; x >>>= 0; rgba[i] = x & 0xff }
  return encodePng(16, 16, rgba)
}

/**
 * Task 15's wire-audit fixture: visible sheet `S` and hidden sheet `SENTINEL_HS` (`state="hidden"`, A1 =
 * `SENTINEL_HC`). On `S`: A1 = `=LEN("SENTINEL_F")` with saved value 10; B1 = 1234.5678 under number format `0`
 * (style 3, numFmtId 1); A2 = `SENTINEL_VIS`; a comment on C1 (`SENTINEL_C`, a comments part the visible-only
 * projection drops); a workbook defined name `SENTINEL_N` → `S!$B$1`; and a hidden picture (`cNvPr hidden="1"`,
 * named `SENTINEL_D`) over E2:H8 whose image is `sentinelPng()`; I10 = 1 extends the used range past the picture.
 * No sentinel sits in a covered cell.
 */
export async function sentinels(): Promise<Uint8Array> {
  const S = SENTINELS
  const rows: XlsxSheetSpec['rows'] = [
    { r: 1, cells: [{ ref: 'A1', formula: `LEN("${S.formula}")`, v: S.formulaSaved }, { ref: 'B1', v: Number(S.number), style: 3 }] },
    { r: 2, cells: [{ ref: 'A2', t: 's', v: 0 }] },
    // Extends the used range past the picture's box (E2:H8, x 256..512, y 20..160 CSS px), so the tiles of the used
    // extent cover it at every pair.
    { r: 10, cells: [{ ref: 'I10', v: 1 }] },
  ]
  const hidden: XlsxSheetSpec['rows'] = [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 1 }] }]
  return patchXlsx(await buildXlsx([{ name: 'S', rows }, { name: S.hiddenSheet, rows: hidden }], [S.visibleText, S.hiddenCell]), async (zip) => {
    const wb = await textOf(zip, 'xl/workbook.xml')
    const marked = wb.replace(`<sheet name="${S.hiddenSheet}" sheetId="2"`, `<sheet name="${S.hiddenSheet}" sheetId="2" state="hidden"`)
      .replace('</sheets>', `</sheets><definedNames><definedName name="${S.definedName}">S!$B$1</definedName></definedNames>`)
    if (marked === wb || !marked.includes('state="hidden"') || !marked.includes('definedNames')) throw new Error('sheet-fixtures: workbook not recognised')
    zip.file('xl/workbook.xml', marked)
    await patchStylesXml(zip, (xml) => xml
      .replace('<cellXfs count="3">', '<cellXfs count="4">')
      .replace('</cellXfs>', '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'))
    // The hidden picture: its drawing, the drawing's image relationship and the media part.
    const pic = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="${S.drawingName}" hidden="1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="${256 * EMU_PER_PX}" y="${20 * EMU_PER_PX}"/><a:ext cx="${256 * EMU_PER_PX}" cy="${140 * EMU_PER_PX}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`
    const part = await addDrawing(zip, 0, twoCellAnchorXml({ col: 4, row: 1 }, { col: 8, row: 8 }, pic))
    const name = part.slice(part.lastIndexOf('/') + 1)
    zip.file(`xl/drawings/_rels/${name}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`)
    zip.file('xl/media/image1.png', sentinelPng())
    // The comment on C1.
    zip.file('xl/comments1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${SPREADSHEET_NS}"><authors><author>a</author></authors><commentList><comment ref="C1" authorId="0"><text><t>${S.comment}</t></text></comment></commentList></comments>`)
    const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels'
    const rels = await textOf(zip, relsPath)
    zip.file(relsPath, rels.replace('</Relationships>', '<Relationship Id="rIdC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>'))
    const ct = await textOf(zip, '[Content_Types].xml')
    zip.file('[Content_Types].xml', ct
      .replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"')
      .replace('</Types>', '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>'))
  })
}

/** The long label of `labelBesideShape`: at Calibri 11 it runs well past D2's left edge (192 CSS px). */
export const LONG_LABEL = 'LONG-A2 label text that runs on across the empty cells into the shape in D2'
/**
 * Task 15 fix round 1 (measured spill): A2 holds `text` (general alignment; `rotated` adds style 3, `textRotation`
 * 45), B2:E2 are empty, and a filled box sits strictly inside D2 (x 192..256, y 20..40 → box x 200..248, y 24..36).
 * A short label fits A2 and never reaches the box; `LONG_LABEL` does.
 */
export async function labelBesideShape(text: string, o: { rotated?: boolean } = {}): Promise<Uint8Array> {
  const st = stringTable()
  const rows = [{ r: 2, cells: [{ ...textCell(st, 'A2', text), ...(o.rotated ? { style: 3 } : {}) }] }]
  return sheetWorkbook(rows, st.list, async (zip) => {
    if (o.rotated) await patchStylesXml(zip, (xml) => xml
      .replace('<cellXfs count="3">', '<cellXfs count="4">')
      .replace('</cellXfs>', '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="45"/></xf></cellXfs>'))
    await addDrawing(zip, 0, twoCellAnchorXml({ col: 3, row: 1, colOffPx: 8, rowOffPx: 4 }, { col: 3, row: 1, colOffPx: 56, rowOffPx: 16 },
      rectShapeXml({ id: 2, name: 'OverD2', offPx: [200, 24], extPx: [48, 12], fill: 'FF8800' })))
  })
}
