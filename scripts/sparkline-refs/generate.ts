/**
 * Writes the XLSX sparkline reference workbooks (spec section 8) to tests/fixtures/xlsx-sparklines/.
 * Run: bun scripts/sparkline-refs/generate.ts
 *
 * Raw SpreadsheetML on purpose: Excel opens these to capture reference images, so the parts, element order and the
 * x14 extension must be what Excel itself writes (no repair prompt). Output is deterministic (fixed zip dates, no
 * timestamps), so re-running leaves git clean.
 */
import { mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { SPARKLINE_CASES, colIndex, colName, type CaseCell, type CaseGroup, type CaseSheet, type CaseValue, type ColourKey, type ColourSpec, type SparklineCase } from './cases'

/** Host cell size, identical in every case (cases.ts header). Width is in characters, height in points. */
export const HOST_WIDTH_CHARS = 40
export const HOST_HEIGHT_PT = 60

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const NS_X14AC = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const ZIP_DATE = new Date(Date.UTC(2026, 0, 1))

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
/** OOXML column width for N characters of the default font (Calibri 11, max digit width 7 px): 40 -> 40.7109375. */
const colWidth = (chars: number) => Math.trunc(((chars * 7 + 5) / 7) * 256) / 256
const DEFAULT_COL_WIDTH = 9.140625
const parseRef = (ref: string) => { const m = /^([A-Z]+)(\d+)$/.exec(ref)!; return { col: colIndex(m[1]), row: Number(m[2]) } }
const range = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i)

interface HostInfo { rows: Set<number>; cols: Set<number>; cells: Set<string> }
/** Host rows, host columns (a merged host widens every column it covers), and every cell that is part of a host. */
function hostInfo(sheet: CaseSheet): HostInfo {
  const info: HostInfo = { rows: new Set(), cols: new Set(), cells: new Set() }
  const merges = (sheet.merges ?? []).map(m => { const [a, b] = m.split(':').map(parseRef); return { a, b } })
  for (const g of sheet.groups) for (const sp of g.sparklines) {
    const h = parseRef(sp.sqref)
    const m = merges.find(x => x.a.col === h.col && x.a.row === h.row)
    const [c0, c1, r0, r1] = m ? [m.a.col, m.b.col, m.a.row, m.b.row] : [h.col, h.col, h.row, h.row]
    for (const r of range(r0, r1)) { info.rows.add(r); for (const c of range(c0, c1)) { info.cols.add(c); info.cells.add(`${colName(c)}${r}`) } }
  }
  return info
}

const COLOUR_ORDER: ColourKey[] = ['colorSeries', 'colorNegative', 'colorAxis', 'colorMarkers', 'colorFirst', 'colorLast', 'colorHigh', 'colorLow']
const colourXml = (key: ColourKey, c: ColourSpec, p: string) => {
  const attrs = 'rgb' in c ? `rgb="${c.rgb}"` : 'theme' in c ? `theme="${c.theme}"${c.tint !== undefined ? ` tint="${c.tint}"` : ''}` : `indexed="${c.indexed}"`
  return `<${p}:${key} ${attrs}/>`
}

function groupXml(g: CaseGroup, x: string, m: string): string {
  const attrs = Object.entries(g.attrs).map(([k, v]) => ` ${k}="${esc(String(v))}"`).join('')
  const colours = COLOUR_ORDER.filter(k => g.colours[k]).map(k => colourXml(k, g.colours[k]!, x)).join('')
  const dates = g.dateRange ? `<${m}:f>${esc(g.dateRange)}</${m}:f>` : ''
  const sparks = g.sparklines.map(s => `<${x}:sparkline><${m}:f>${esc(s.f)}</${m}:f><${m}:sqref>${esc(s.sqref)}</${m}:sqref></${x}:sparkline>`).join('')
  return `<${x}:sparklineGroup${attrs}>${colours}${dates}<${x}:sparklines>${sparks}</${x}:sparklines></${x}:sparklineGroup>`
}

function extLstXml(sheet: CaseSheet): string {
  if (sheet.groups.length === 0) return ''
  const x = sheet.groups[0].prefixes?.x14 ?? 'x14', m = sheet.groups[0].prefixes?.xm ?? 'xm'
  for (const g of sheet.groups) if ((g.prefixes?.x14 ?? 'x14') !== x || (g.prefixes?.xm ?? 'xm') !== m) throw new Error('groups of one sheet must share namespace prefixes')
  return `<extLst><ext uri="${EXT_URI}" xmlns:${x}="${NS_X14}"><${x}:sparklineGroups xmlns:${m}="${NS_XM}">${sheet.groups.map(g => groupXml(g, x, m)).join('')}</${x}:sparklineGroups></ext></extLst>`
}

/** Style indices (styles.xml cellXfs): 0 plain, 1 border, 2 fill, 3 fill + border. */
const styleOf = (fill: boolean | undefined, border: boolean) => (fill ? 2 : 0) + (border ? 1 : 0)

function sheetXml(sheet: CaseSheet, bordered: boolean, sst: (s: string) => number): string {
  const hosts = hostInfo(sheet)
  const filter = sheet.autoFilter
  const filterRows = new Set<number>()
  if (filter) {
    const [a, b] = filter.ref.split(':').map(parseRef)
    const values = new Map(sheet.cells.map(c => [c.ref, c.v]))
    for (let r = a.row + 1; r <= b.row; r++) if (!filter.keep.includes(values.get(`${colName(a.col)}${r}`) as number)) filterRows.add(r)
  }
  const hiddenRows = new Set([...(sheet.hiddenRows ?? []), ...filterRows])

  // Host cells exist (styled, empty) when they carry a border, so the border has a cell to live on.
  const cells = new Map<string, CaseCell>(sheet.cells.map(c => [c.ref, c]))
  if (bordered) for (const ref of hosts.cells) if (!cells.has(ref)) cells.set(ref, { ref, v: null })
  const byRow = new Map<number, Array<{ col: number; cell: CaseCell }>>()
  for (const cell of cells.values()) {
    const p = parseRef(cell.ref)
    if (!byRow.has(p.row)) byRow.set(p.row, [])
    byRow.get(p.row)!.push({ col: p.col, cell })
  }
  const rows = [...new Set([...byRow.keys(), ...hosts.rows, ...hiddenRows])].sort((a, b) => a - b)

  const size = sheet.hostSize ?? { widthChars: HOST_WIDTH_CHARS, heightPt: HOST_HEIGHT_PT }
  const cellXml = (ref: string, s: number, v: CaseValue): string => {
    const sAttr = s ? ` s="${s}"` : ''
    if (v === null) return s ? `<c r="${ref}"${sAttr}/>` : ''
    if (typeof v === 'number') return `<c r="${ref}"${sAttr}><v>${v}</v></c>`
    if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`
    if (typeof v === 'string') return `<c r="${ref}"${sAttr} t="s"><v>${sst(v)}</v></c>`
    return `<c r="${ref}"${sAttr} t="e"><v>${esc(v.error)}</v></c>`
  }
  const rowsXml = rows.map(r => {
    const content = (byRow.get(r) ?? []).sort((a, b) => a.col - b.col).map(({ cell }) => cellXml(cell.ref, styleOf(cell.fill, bordered && hosts.cells.has(cell.ref)), cell.v)).join('')
    const attrs = `${hosts.rows.has(r) ? ` ht="${size.heightPt}" customHeight="1"` : ''}${hiddenRows.has(r) ? ' hidden="1"' : ''}`
    return `<row r="${r}"${attrs}>${content}</row>`
  }).join('')

  const colSpecs = new Map<number, { width: number; hidden?: boolean }>()
  for (const c of hosts.cols) colSpecs.set(c, { width: colWidth(size.widthChars) })
  for (const c of sheet.hiddenCols ?? []) colSpecs.set(c, { width: colSpecs.get(c)?.width ?? DEFAULT_COL_WIDTH, hidden: true })
  const colsXml = colSpecs.size
    ? `<cols>${[...colSpecs].sort((a, b) => a[0] - b[0]).map(([c, s]) => `<col min="${c}" max="${c}" width="${s.width}"${s.hidden ? ' hidden="1"' : ''} customWidth="1"/>`).join('')}</cols>`
    : ''

  const refs = [...cells.keys(), ...hosts.cells].map(parseRef)
  let dim = 'A1'
  if (refs.length) {
    const a = `${colName(Math.min(...refs.map(p => p.col)))}${Math.min(...refs.map(p => p.row))}`
    const b = `${colName(Math.max(...refs.map(p => p.col)))}${Math.max(...refs.map(p => p.row))}`
    dim = a === b ? a : `${a}:${b}`
  }
  const filterXml = filter
    ? `<autoFilter ref="${filter.ref}"><filterColumn colId="0"><filters>${filter.keep.map(v => `<filter val="${v}"/>`).join('')}</filters></filterColumn></autoFilter>`
    : ''
  const mergeXml = sheet.merges?.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''
  return `${XML_DECL}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" mc:Ignorable="x14ac" xmlns:x14ac="${NS_X14AC}">`
    + `${filter ? '<sheetPr filterMode="1"/>' : ''}<dimension ref="${dim}"/>`
    + `<sheetViews><sheetView showGridLines="${sheet.gridlines ? 1 : 0}" workbookViewId="0"/></sheetViews>`
    + `<sheetFormatPr defaultRowHeight="15"/>${colsXml}<sheetData>${rowsXml}</sheetData>${filterXml}${mergeXml}`
    + `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${extLstXml(sheet)}</worksheet>`
}

const STYLES_XML = `${XML_DECL}<styleSheet xmlns="${NS_MAIN}">
<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`

/** The default Office theme (Office 2013-2022). */
const THEME_XML = `${XML_DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>
<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`

const absolute = (ref: string) => ref.replace(/([A-Z]+)(\d+)/, '$$$1$$$2')

/** One case as an xlsx package (`bordered`: a thin black border on every host cell, for the comparison crop). */
export async function buildCaseWorkbook(c: SparklineCase, bordered: boolean): Promise<Uint8Array> {
  const strings: string[] = []
  const sst = (s: string) => { const i = strings.indexOf(s); if (i >= 0) return i; strings.push(s); return strings.length - 1 }
  const sheets = c.sheets.map(s => sheetXml(s, bordered, sst))
  const n = sheets.length
  const filters = c.sheets
    .map((s, i) => s.autoFilter ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${esc(s.name)}!${s.autoFilter.ref.split(':').map(absolute).join(':')}</definedName>` : '')
    .join('')
  const files: Record<string, string> = {
    '[Content_Types].xml': `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
    '_rels/.rels': `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `${XML_DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><workbookPr defaultThemeVersion="166925"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="17500"/></bookViews><sheets>${c.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${filters ? `<definedNames>${filters}</definedNames>` : ''}<calcPr calcId="191029"/></workbook>`,
    'xl/_rels/workbook.xml.rels': `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${n + 1}" Type="${REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rId${n + 2}" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId${n + 3}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
    'xl/theme/theme1.xml': THEME_XML,
    'xl/styles.xml': STYLES_XML,
    'xl/sharedStrings.xml': `${XML_DECL}<sst xmlns="${NS_MAIN}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map(s => `<si><t>${esc(s)}</t></si>`).join('')}</sst>`,
  }
  sheets.forEach((x, i) => { files[`xl/worksheets/sheet${i + 1}.xml`] = x })
  const zip = new JSZip()
  for (const [name, data] of Object.entries(files)) zip.file(name, data, { date: ZIP_DATE, createFolders: false })
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

export const FIXTURE_DIR = fileURLToPath(new URL('../../tests/fixtures/xlsx-sparklines/', import.meta.url))

async function main() {
  mkdirSync(FIXTURE_DIR, { recursive: true })
  const wanted = new Set<string>()
  for (const c of SPARKLINE_CASES) for (const bordered of [false, true]) {
    const name = `${c.id}${bordered ? '-bordered' : ''}.xlsx`
    wanted.add(name)
    writeFileSync(resolve(FIXTURE_DIR, name), await buildCaseWorkbook(c, bordered))
  }
  for (const f of readdirSync(FIXTURE_DIR)) if (f.endsWith('.xlsx') && !wanted.has(f)) rmSync(resolve(FIXTURE_DIR, f))
  console.log(`wrote ${wanted.size} workbooks to ${FIXTURE_DIR}`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await main()
