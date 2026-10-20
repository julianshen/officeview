/**
 * Office-shaped synthetic DOCX/PPTX with charts whose data lives in an
 * embedded workbook (`c:externalData` -> `../embeddings/*.xlsx`). These are
 * synthetic builder fixtures, never Office evidence.
 *
 * The outer package is written with `zip-builder`, so a test can forge one
 * entry's declared sizes or replace its compressed bytes (`entry`), which
 * JSZip's own writer never allows.
 */
import { buildZip, type BuildEntry } from './zip-builder'
import { buildXlsx, type XlsxCellSpec } from '../../src/testdata/ooxml-builders'

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PACKAGE_REL = `${REL}/package`
const CHART_REL = `${REL}/chart`

export interface ChartEmbedSpec {
  /** The chart part; default `officeChartXml(rId)`. */
  chartXml?: string
  /** The relationship id `c:externalData` names (default rId1). */
  rId?: string
  /** The embedded workbook; omitted -> no embedding part is written (the target dangles). */
  workbook?: Uint8Array
  /** Relationship target, relative to the chart part (default `../embeddings/Microsoft_Excel_Worksheet{n}.xlsx`). */
  target?: string
  /** TargetMode="External". */
  external?: boolean
  /** Overrides for the embedding's zip entry: forged sizes, raw compressed bytes, method. */
  entry?: Partial<BuildEntry>
}

export interface ChartFixture {
  bytes: Uint8Array
  /** The chart part path of each chart, in order. */
  chartPaths: string[]
  /** The embedding part path each chart's relationship targets (resolved; undefined when external). */
  embeddingPaths: (string | undefined)[]
}

/** Office's bar chart with a reference series, no caches, and `c:externalData` after `c:chart`. */
export function officeChartXml(rId = 'rId1', values = 'Sheet1!$B$2:$B$3', categories = 'Sheet1!$A$2:$A$3'): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>${categories}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>${values}</c:f></c:numRef></c:val></c:ser><c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:externalData r:id="${rId}"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`
}

/** The embedded workbook Office writes for a new chart: Sheet1 with a header row and two data rows. */
export function embeddedXlsx(rows: Array<{ r: number; cells: XlsxCellSpec[] }> = [
  { r: 1, cells: [{ ref: 'A1', v: 'Category', t: 'str' }, { ref: 'B1', v: 'Series 1', t: 'str' }] },
  { r: 2, cells: [{ ref: 'A2', v: 'A', t: 'str' }, { ref: 'B2', v: 4.3 }] },
  { r: 3, cells: [{ ref: 'A3', v: 'B', t: 'str' }, { ref: 'B3', v: 2.5 }] },
]): Promise<Uint8Array> {
  return buildXlsx([{ name: 'Sheet1', rows }])
}

const embeddingName = (i: number) => `Microsoft_Excel_Worksheet${i ? i : ''}.xlsx`
const resolve = (owner: string, target: string) => {
  const bits = owner.split('/').slice(0, -1)
  for (const bit of target.split('/')) { if (bit === '..') bits.pop(); else if (bit && bit !== '.') bits.push(bit) }
  return bits.join('/')
}

function chartEntries(root: 'word' | 'ppt', charts: ChartEmbedSpec[]): { entries: BuildEntry[]; chartPaths: string[]; embeddingPaths: (string | undefined)[] } {
  const entries: BuildEntry[] = [], chartPaths: string[] = [], embeddingPaths: (string | undefined)[] = []
  charts.forEach((spec, i) => {
    const rId = spec.rId ?? 'rId1', chartPath = `${root}/charts/chart${i + 1}.xml`
    const target = spec.target ?? `../embeddings/${embeddingName(i)}`
    chartPaths.push(chartPath)
    entries.push({ name: chartPath, data: spec.chartXml ?? officeChartXml(rId) })
    entries.push({ name: `${root}/charts/_rels/chart${i + 1}.xml.rels`, data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${rId}" Type="${PACKAGE_REL}" Target="${target}"${spec.external ? ' TargetMode="External"' : ''}/></Relationships>` })
    const embedding = spec.external ? undefined : resolve(chartPath, target)
    embeddingPaths.push(embedding)
    if (embedding && (spec.workbook || spec.entry?.raw)) entries.push({ name: embedding, data: spec.workbook, ...spec.entry })
  })
  return { entries, chartPaths, embeddingPaths }
}

const contentTypes = (main: string, chartOverrides: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>${main}${chartOverrides}</Types>`
const chartOverride = (path: string) => `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`

/** A DOCX with one inline chart per spec. */
export function buildChartDocx(charts: ChartEmbedSpec[]): ChartFixture {
  const { entries, chartPaths, embeddingPaths } = chartEntries('word', charts)
  const inline = (i: number) => `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="${i + 1}" name="Chart ${i + 1}"/><a:graphic><a:graphicData uri="${NS_C}"><c:chart r:id="rChart${i + 1}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="${NS_A}" xmlns:c="${NS_C}" xmlns:r="${NS_R}"><w:body>${charts.map((_, i) => inline(i)).join('')}<w:sectPr/></w:body></w:document>`
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${chartPaths.map((p, i) => `<Relationship Id="rChart${i + 1}" Type="${CHART_REL}" Target="${p.slice('word/'.length)}"/>`).join('')}</Relationships>`
  const bytes = buildZip([
    { name: '[Content_Types].xml', data: contentTypes('<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>', chartPaths.map(chartOverride).join('')) },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/document.xml', data: document },
    { name: 'word/_rels/document.xml.rels', data: rels },
    ...entries,
  ]).bytes
  return { bytes, chartPaths, embeddingPaths }
}

/** A PPTX with one slide holding one chart graphic frame per spec. */
export function buildChartPptx(charts: ChartEmbedSpec[]): ChartFixture {
  const { entries, chartPaths, embeddingPaths } = chartEntries('ppt', charts)
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
  const frame = (i: number) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${i + 2}" name="Chart ${i + 1}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${838200 + i * 100}" y="1825625"/><a:ext cx="5486400" cy="3200400"/></p:xfrm><a:graphic><a:graphicData uri="${NS_C}"><c:chart r:id="rChart${i + 1}"/></a:graphicData></a:graphic></p:graphicFrame>`
  const slide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="${P}" xmlns:a="${NS_A}" xmlns:c="${NS_C}" xmlns:r="${NS_R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${charts.map((_, i) => frame(i)).join('')}</p:spTree></p:cSld></p:sld>`
  const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${chartPaths.map((p, i) => `<Relationship Id="rChart${i + 1}" Type="${CHART_REL}" Target="../charts/${p.slice('ppt/charts/'.length)}"/>`).join('')}</Relationships>`
  const bytes = buildZip([
    { name: '[Content_Types].xml', data: contentTypes('<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>', chartPaths.map(chartOverride).join('')) },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>` },
    { name: 'ppt/presentation.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="${P}" xmlns:r="${NS_R}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>` },
    { name: 'ppt/_rels/presentation.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="${REL}/slide" Target="slides/slide1.xml"/></Relationships>` },
    { name: 'ppt/slides/slide1.xml', data: slide },
    { name: 'ppt/slides/_rels/slide1.xml.rels', data: slideRels },
    ...entries,
  ]).bytes
  return { bytes, chartPaths, embeddingPaths }
}

/** One worksheet of an embedded workbook: `data` is the inner XML of `sheetData` (see `row`/`cell`). */
export interface EmbeddedSheetSpec {
  name: string
  data?: string
  /** Inner XML of `cols`. */
  cols?: string
  /** A complete worksheet part, replacing the generated one. */
  raw?: string
}
/** An Office-shaped synthetic embedded workbook with cells, shared strings, styles, names and epoch. */
export interface EmbeddedWorkbookSpec {
  sheets: EmbeddedSheetSpec[]
  sharedStrings?: string[]
  /** A complete sharedStrings part, replacing the generated one. */
  sst?: string
  definedNames?: Array<{ name: string; formula: string; localSheetIndex?: number }>
  date1904?: boolean
  numFmts?: Array<{ id: number; code: string }>
  /** One `xf` per entry: its numFmtId. */
  cellXfs?: number[]
  /** Workbook rels Target of each sheet (default `worksheets/sheet{n}.xml`). */
  sheetTargets?: string[]
}
const esc = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
/** A worksheet cell: a number, `t` typed text (`s` index, `str`, `e`, `b`), an optional formula and style. */
export function cell(ref: string, v?: string | number, options: { t?: 's' | 'str' | 'e' | 'b' | 'n' | 'inlineStr'; f?: string; s?: number } = {}): string {
  const body = (options.f !== undefined ? `<f>${esc(options.f)}</f>` : '') + (options.t === 'inlineStr' ? `<is><t>${esc(String(v))}</t></is>` : v !== undefined ? `<v>${esc(String(v))}</v>` : '')
  return `<c r="${ref}"${options.t ? ` t="${options.t}"` : ''}${options.s !== undefined ? ` s="${options.s}"` : ''}>${body}</c>`
}
export const row = (r: number, ...cells: string[]): string => `<row r="${r}">${cells.join('')}</row>`
export function embeddedWorkbook(spec: EmbeddedWorkbookSpec): Uint8Array {
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  const target = (i: number) => spec.sheetTargets?.[i] ?? `worksheets/sheet${i + 1}.xml`
  const names = spec.definedNames?.length ? `<definedNames>${spec.definedNames.map(n => `<definedName name="${esc(n.name)}"${n.localSheetIndex !== undefined ? ` localSheetId="${n.localSheetIndex}"` : ''}>${esc(n.formula)}</definedName>`).join('')}</definedNames>` : ''
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${S}" xmlns:r="${NS_R}"><fileVersion appName="xl"/><workbookPr${spec.date1904 ? ' date1904="1"' : ''} defaultThemeVersion="166925"/><bookViews><workbookView xWindow="0" yWindow="0"/></bookViews><sheets>${spec.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${names}<calcPr calcId="191029"/></workbook>`
  const n = spec.sheets.length
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${spec.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="${target(i)}"/>`).join('')}<Relationship Id="rId${n + 1}" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId${n + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`
  const numFmts = spec.numFmts?.length ? `<numFmts count="${spec.numFmts.length}">${spec.numFmts.map(f => `<numFmt numFmtId="${f.id}" formatCode="${esc(f.code)}"/>`).join('')}</numFmts>` : ''
  const xfs = spec.cellXfs ?? [0]
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${S}">${numFmts}<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.map(id => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"${id ? ' applyNumberFormat="1"' : ''}/>`).join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`
  const sheet = (s: EmbeddedSheetSpec) => s.raw ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${S}" xmlns:r="${NS_R}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>${s.cols ? `<cols>${s.cols}</cols>` : ''}<sheetData>${s.data ?? ''}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
  const strings = spec.sharedStrings ?? []
  const sst = spec.sst ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${S}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map(s => `<si><t>${esc(s)}</t></si>`).join('')}</sst>`
  const resolveSheet = (i: number) => { const t = target(i); return t.startsWith('/') ? t.slice(1) : `xl/${t}` }
  return buildZip([
    { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>` },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: 'xl/workbook.xml', data: workbook },
    { name: 'xl/_rels/workbook.xml.rels', data: rels },
    { name: 'xl/styles.xml', data: styles },
    { name: 'xl/sharedStrings.xml', data: sst },
    ...spec.sheets.map((s, i) => ({ name: resolveSheet(i), data: sheet(s) })),
  ]).bytes
}
/** Office's default chart workbook (Sheet1: header row, categories in A, values in B), shared strings as Office writes them. */
export function officeChartWorkbook(values: Array<number | string> = [4.3, 2.5], categories: string[] = ['A', 'B'], extra: Partial<EmbeddedWorkbookSpec> = {}): Uint8Array {
  const strings = ['Category', 'Series 1', ...categories]
  const rows = [row(1, cell('A1', 0, { t: 's' }), cell('B1', 1, { t: 's' }))]
  values.forEach((v, i) => rows.push(row(i + 2, cell(`A${i + 2}`, i + 2, { t: 's' }), typeof v === 'number' ? cell(`B${i + 2}`, v) : v)))
  return embeddedWorkbook({ sheets: [{ name: 'Sheet1', data: rows.join('') }], sharedStrings: strings, ...extra })
}
