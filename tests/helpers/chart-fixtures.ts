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
