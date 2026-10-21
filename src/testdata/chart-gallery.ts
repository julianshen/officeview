/**
 * The synthetic chart gallery (spec 8, Goldens): one chart per natively drawn family and major element, in one
 * DOCX, one PPTX and one XLSX document. Every part here is **Office-shaped synthetic** (written by hand to the
 * shapes Office writes), never Office evidence.
 *
 * `buildChartDocument` is the shared multi-object package writer: it places chart frames (classic `c:` and ChartEx
 * `cx:` parts, optionally inside `mc:AlternateContent` with a fallback picture), model3D and OMML Choices, embedded
 * workbooks and external links in the format's own carrier. The robustness suite (tests/charts-robustness.test.ts)
 * builds its torture documents with it; the golden suite (tests/golden.test.ts) renders `chartGallery(format)`.
 *
 * Test data only: nothing in the library imports this module.
 */
import JSZip from 'jszip'
import { buildXlsx, type XlsxCellSpec } from './ooxml-builders'

export type ChartDocFormat = 'docx' | 'pptx' | 'xlsx'

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const NS_CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const NS_CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_A14 = 'http://schemas.microsoft.com/office/drawing/2010/main'
const NS_AM3D = 'http://schemas.microsoft.com/office/drawing/2017/model3d'
const NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const EMU = 9525

/** One placed object of a chart document. */
export interface ChartDocObject {
  /** The object's name (`cNvPr@name` / `docPr@name`). */
  name: string
  /** A classic `c:chartSpace` or ChartEx `cx:chartSpace` part (the namespace picks the relationship type). */
  part?: string
  /** DOCX/PPTX: the chart's embedded workbook (`c:externalData`/`cx:externalData` rId1 targets it). */
  workbook?: Uint8Array
  /** DOCX/PPTX: rId1 is an external link to this target instead (TargetMode="External"). */
  externalWorkbook?: string
  /** Wrap the frame in `mc:AlternateContent` with this Requires (e.g. `cx1`); `fallback` is its picture. */
  requires?: string
  /** PNG bytes of the Fallback picture (only with `requires` or `choice`). */
  fallback?: Uint8Array
  /** A non-chart Choice body instead of a chart frame: a 3-D model or an OMML equation (Requires am3d / a14). */
  choice?: 'model3D' | 'omml'
}

export interface ChartDocOptions {
  /** Objects per row (DOCX paragraph, PPTX slide row, XLSX grid row). Default 2. */
  columns?: number
  /** PPTX: rows per slide. Default 2. */
  rows?: number
  /** Object size in px (96 dpi). Default 320 x 220. */
  width?: number
  height?: number
  /** XLSX: worksheet rows (the hierarchy data the ChartEx names point at, for instance). */
  cells?: Array<{ r: number; cells: XlsxCellSpec[] }>
  /** XLSX: workbook defined names. */
  definedNames?: Array<{ name: string; formula: string }>
  /** DOCX: a paragraph of text (and an inline OMML equation when `omml`) before the objects. */
  heading?: string
  omml?: boolean
}

const esc = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const isChartEx = (part: string) => part.includes(NS_CX)
const rels = (items: Array<{ id: string; type: string; target: string; external?: boolean }>) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(r => `<Relationship Id="${r.id}" Type="${r.type.startsWith('http') ? r.type : `${REL}/${r.type}`}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`
const NAMESPACES = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:c="${NS_C}" xmlns:cx="${NS_CX}" xmlns:cx1="${NS_CX1}" xmlns:mc="${NS_MC}" xmlns:pic="${NS_PIC}" xmlns:a14="${NS_A14}" xmlns:am3d="${NS_AM3D}" xmlns:m="${NS_M}"`

/** The graphic of one object: a chart, ChartEx, model3D or picture payload. */
function graphic(object: ChartDocObject, index: number, picture = false): string {
  if (picture) return `<a:graphic><a:graphicData uri="${NS_PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="${1000 + index}" name="${esc(object.name)} fallback"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rImg${index}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`
  if (object.choice === 'model3D') return `<a:graphic><a:graphicData uri="${NS_AM3D}"><am3d:model3d r:embed="rModel${index}"><am3d:spPr/><am3d:camera><am3d:pos x="0" y="0" z="1000"/></am3d:camera></am3d:model3d></a:graphicData></a:graphic>`
  const chartEx = object.part !== undefined && isChartEx(object.part)
  return `<a:graphic><a:graphicData uri="${chartEx ? NS_CX : NS_C}">${chartEx ? `<cx:chart r:id="rChart${index}"/>` : `<c:chart r:id="rChart${index}"/>`}</a:graphicData></a:graphic>`
}
const OMML = `<m:oMathPara><m:oMath><m:r><m:t>E</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>`
const requiresOf = (object: ChartDocObject) => object.choice === 'model3D' ? 'am3d' : object.choice === 'omml' ? 'a14' : object.requires
const alternate = (choice: string, fallback: string | undefined, requires: string) =>
  `<mc:AlternateContent><mc:Choice Requires="${requires}">${choice}</mc:Choice>${fallback !== undefined ? `<mc:Fallback>${fallback}</mc:Fallback>` : ''}</mc:AlternateContent>`

/** Package parts every object needs (its chart part, rels, embedding, fallback image, model), under `root`. */
function objectParts(zip: JSZip, root: 'word' | 'ppt' | 'xl', objects: ChartDocObject[], ownerRels: Array<{ id: string; type: string; target: string }>, ownerPrefix: string, base = 0): string[] {
  const overrides: string[] = []
  objects.forEach((object, k) => {
    const i = base + k
    if (object.part !== undefined) {
      const ex = isChartEx(object.part), file = ex ? `chartEx${i + 1}.xml` : `chart${i + 1}.xml`
      zip.file(`${root}/charts/${file}`, object.part)
      overrides.push(`<Override PartName="/${root}/charts/${file}" ContentType="${ex ? 'application/vnd.ms-office.chartex+xml' : 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'}"/>`)
      ownerRels.push({ id: `rChart${i + 1}`, type: ex ? 'http://schemas.microsoft.com/office/2014/relationships/chartEx' : 'chart', target: `${ownerPrefix}charts/${file}` })
      if (object.workbook || object.externalWorkbook) {
        const target = object.externalWorkbook ?? `../embeddings/Microsoft_Excel_Worksheet${i + 1}.xlsx`
        zip.file(`${root}/charts/_rels/${file}.rels`, rels([{ id: 'rId1', type: 'package', target, external: object.externalWorkbook !== undefined }]))
        if (object.workbook) zip.file(`${root}/embeddings/Microsoft_Excel_Worksheet${i + 1}.xlsx`, object.workbook)
      }
    }
    if (object.fallback) {
      zip.file(`${root}/media/fallback${i + 1}.png`, object.fallback)
      ownerRels.push({ id: `rImg${i + 1}`, type: 'image', target: `${ownerPrefix}media/fallback${i + 1}.png` })
    }
    if (object.choice === 'model3D') {
      zip.file(`${root}/media/model${i + 1}.glb`, new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]))
      ownerRels.push({ id: `rModel${i + 1}`, type: 'http://schemas.microsoft.com/office/2017/06/relationships/model3d', target: `${ownerPrefix}media/model${i + 1}.glb` })
    }
  })
  return overrides
}

const contentTypes = (main: string, overrides: string[]) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="glb" ContentType="model/gltf-binary"/><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>${main}${overrides.join('')}</Types>`

function docxObject(object: ChartDocObject, i: number, w: number, h: number): string {
  const drawing = (inner: string, id: number) => `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${w * EMU}" cy="${h * EMU}"/><wp:docPr id="${id}" name="${esc(object.name)}"/>${inner}</wp:inline></w:drawing>`
  if (object.choice === 'omml') {
    const shape = `<w:drawing><wp:inline><wp:extent cx="${w * EMU}" cy="${h * EMU}"/><wp:docPr id="${i + 1}" name="${esc(object.name)}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><a14:m>${OMML}</a14:m></a:graphicData></a:graphic></wp:inline></w:drawing>`
    return `<w:r>${alternate(shape, object.fallback ? drawing(graphic(object, i + 1, true), 2000 + i) : undefined, 'a14')}</w:r>`
  }
  const frame = drawing(graphic(object, i + 1), i + 1)
  const requires = requiresOf(object)
  if (!requires) return `<w:r>${frame}</w:r>`
  return `<w:r>${alternate(frame, object.fallback ? drawing(graphic(object, i + 1, true), 2000 + i) : undefined, requires)}</w:r>`
}

function pptxObject(object: ChartDocObject, i: number, x: number, y: number, w: number, h: number): string {
  const xfrm = `<a:off x="${x * EMU}" y="${y * EMU}"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/>`
  const picture = `<p:pic><p:nvPicPr><p:cNvPr id="${2000 + i}" name="${esc(object.name)} fallback"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rImg${i + 1}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm>${xfrm}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  if (object.choice === 'omml') {
    const shape = `<p:sp><p:nvSpPr><p:cNvPr id="${i + 2}" name="${esc(object.name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>${xfrm}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a14:m>${OMML}</a14:m></a:p></p:txBody></p:sp>`
    return alternate(shape, object.fallback ? picture : undefined, 'a14')
  }
  const frame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${i + 2}" name="${esc(object.name)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm>${xfrm}</p:xfrm>${graphic(object, i + 1)}</p:graphicFrame>`
  const requires = requiresOf(object)
  return requires ? alternate(frame, object.fallback ? picture : undefined, requires) : frame
}

function xlsxObject(object: ChartDocObject, i: number, x: number, y: number, w: number, h: number): string {
  const xfrm = `<a:off x="0" y="0"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/>`
  const picture = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${2000 + i}" name="${esc(object.name)} fallback"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rImg${i + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm>${xfrm}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`
  let body: string
  if (object.choice === 'omml') {
    const shape = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${i + 2}" name="${esc(object.name)}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm>${xfrm}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a14:m>${OMML}</a14:m></a:p></xdr:txBody></xdr:sp>`
    body = alternate(shape, object.fallback ? picture : undefined, 'a14')
  } else {
    const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="${esc(object.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm>${xfrm}</xdr:xfrm>${graphic(object, i + 1)}</xdr:graphicFrame>`
    const requires = requiresOf(object)
    body = requires ? alternate(frame, object.fallback ? picture : undefined, requires) : frame
  }
  return `<xdr:absoluteAnchor><xdr:pos x="${x * EMU}" y="${y * EMU}"/><xdr:ext cx="${w * EMU}" cy="${h * EMU}"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`
}

/** Writes a document of `format` holding `objects`, laid out in a grid. */
export async function buildChartDocument(format: ChartDocFormat, objects: ChartDocObject[], options: ChartDocOptions = {}): Promise<Uint8Array> {
  const columns = options.columns ?? 2, w = options.width ?? 320, h = options.height ?? 220, gap = 12
  if (format === 'docx') {
    const zip = new JSZip(), documentRels: Array<{ id: string; type: string; target: string }> = []
    const overrides = objectParts(zip, 'word', objects, documentRels, '')
    const paragraphs: string[] = []
    if (options.heading !== undefined || options.omml) paragraphs.push(`<w:p>${options.heading !== undefined ? `<w:r><w:t xml:space="preserve">${esc(options.heading)}</w:t></w:r>` : ''}${options.omml ? `<m:oMath><m:r><m:t>a+b</m:t></m:r></m:oMath>` : ''}</w:p>${options.omml ? `<w:p>${OMML}</w:p>` : ''}`)
    for (let i = 0; i < objects.length; i += columns) {
      const row = objects.slice(i, i + columns).map((object, k) => docxObject(object, i + k, w, h)).join('<w:r><w:t xml:space="preserve"> </w:t></w:r>')
      paragraphs.push(`<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr>${row}</w:p>`)
    }
    zip.file('[Content_Types].xml', contentTypes('<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>', overrides))
    zip.file('_rels/.rels', rels([{ id: 'rId1', type: 'officeDocument', target: 'word/document.xml' }]))
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${NS_W}" xmlns:wp="${NS_WP}" ${NAMESPACES}><w:body>${paragraphs.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', rels(documentRels))
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  }
  if (format === 'pptx') {
    const zip = new JSZip(), rowsPerSlide = options.rows ?? 2, perSlide = columns * rowsPerSlide
    const slideW = columns * (w + gap) + gap, slideH = rowsPerSlide * (h + gap) + gap
    const slides: string[] = [], overrides: string[] = []
    for (let s = 0; s * perSlide < objects.length || (s === 0 && !objects.length); s++) {
      const items = objects.slice(s * perSlide, (s + 1) * perSlide)
      const slideRels: Array<{ id: string; type: string; target: string }> = []
      // Parts and relationship ids are numbered by their document-wide index, so slides never collide.
      overrides.push(...objectParts(zip, 'ppt', items, slideRels, '../', s * perSlide))
      const shapes = items.map((object, k) => pptxObject(object, s * perSlide + k, gap + (k % columns) * (w + gap), gap + Math.floor(k / columns) * (h + gap), w, h)).join('')
      zip.file(`ppt/slides/slide${s + 1}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:p="${NS_P}" ${NAMESPACES}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`)
      zip.file(`ppt/slides/_rels/slide${s + 1}.xml.rels`, rels(slideRels))
      overrides.push(`<Override PartName="/ppt/slides/slide${s + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`)
      slides.push(`slides/slide${s + 1}.xml`)
    }
    zip.file('[Content_Types].xml', contentTypes('<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>', overrides))
    zip.file('_rels/.rels', rels([{ id: 'rId1', type: 'officeDocument', target: 'ppt/presentation.xml' }]))
    zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentation xmlns:p="${NS_P}" xmlns:r="${NS_R}"><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rSlide${i + 1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${slideW * EMU}" cy="${slideH * EMU}"/></p:presentation>`)
    zip.file('ppt/_rels/presentation.xml.rels', rels(slides.map((target, i) => ({ id: `rSlide${i + 1}`, type: 'slide', target }))))
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  }
  const base = await buildXlsx([{ name: 'Sheet1', rows: options.cells ?? [], cols: '<col min="1" max="3" width="10" customWidth="1"/>' }])
  const zip = await JSZip.loadAsync(base), drawingRels: Array<{ id: string; type: string; target: string }> = []
  const overrides = objectParts(zip, 'xl', objects, drawingRels, '../')
  // Charts sit to the right of the data columns (A:C, about 210 px wide).
  const left = 230
  const anchors = objects.map((object, i) => xlsxObject(object, i, left + (i % columns) * (w + gap), gap + Math.floor(i / columns) * (h + gap), w, h)).join('')
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  sheet = sheet.replace('<worksheet ', `<worksheet xmlns:r="${NS_R}" `).replace('</worksheet>', '<drawing r:id="rDrawing"/></worksheet>')
  zip.file('xl/worksheets/sheet1.xml', sheet)
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', rels([{ id: 'rDrawing', type: 'drawing', target: '../drawings/drawing1.xml' }]))
  zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS_XDR}" ${NAMESPACES}>${anchors}</xdr:wsDr>`)
  zip.file('xl/drawings/_rels/drawing1.xml.rels', rels(drawingRels))
  if (options.definedNames?.length) {
    const workbook = await zip.file('xl/workbook.xml')!.async('string')
    zip.file('xl/workbook.xml', workbook.replace('</sheets>', `</sheets><definedNames>${options.definedNames.map(n => `<definedName name="${esc(n.name)}" hidden="1">${esc(n.formula)}</definedName>`).join('')}</definedNames>`))
  }
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file('[Content_Types].xml', types.replace('</Types>', `<Default Extension="png" ContentType="image/png"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>${overrides.join('')}</Types>`))
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

// ---- Chart part builders (Office-shaped synthetic) ---------------------------------------------------------------

const C_SPACE = `xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"`
const titleXml = (text: string) => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1200" b="1"/><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
/** A classic chart space: `plot` is the plot area's groups and axes. */
export function chartSpace(title: string | undefined, plot: string, options: { legend?: 'r' | 'b' | 't' | 'l' | null; titleXml?: string; extra?: string; blanks?: 'gap' | 'zero' | 'span' } = {}): string {
  const legend = options.legend === undefined ? 'b' : options.legend
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace ${C_SPACE}><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${options.titleXml ?? (title !== undefined ? titleXml(title) : '')}<c:autoTitleDeleted val="${title === undefined && !options.titleXml ? 1 : 0}"/><c:plotArea><c:layout/>${plot}</c:plotArea>${legend ? `<c:legend><c:legendPos val="${legend}"/><c:overlay val="0"/></c:legend>` : ''}<c:plotVisOnly val="1"/><c:dispBlanksAs val="${options.blanks ?? 'gap'}"/>${options.extra ?? ''}</c:chart></c:chartSpace>`
}
/** A cached string reference (`c:strRef`) of `labels`. */
export const strRef = (f: string, labels: string[]) => `<c:strRef><c:f>${f}</c:f><c:strCache><c:ptCount val="${labels.length}"/>${labels.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`
/** A cached numeric reference (`c:numRef`); `undefined` leaves the point out (blank). */
export const numRef = (f: string, values: Array<number | string | undefined>, format = 'General') => `<c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>${esc(format)}</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => v === undefined ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`
const col = (k: number) => String.fromCharCode(66 + k)
/** One series with a name, categories and values (cached references into Sheet1). */
export function series(k: number, name: string, categories: string[] | undefined, values: Array<number | string | undefined>, extra = ''): string {
  const n = values.length
  return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:tx>${strRef(`Sheet1!$${col(k)}$1`, [name])}</c:tx>${extra}${categories ? `<c:cat>${strRef(`Sheet1!$A$2:$A$${n + 1}`, categories)}</c:cat>` : ''}<c:val>${numRef(`Sheet1!$${col(k)}$2:$${col(k)}$${n + 1}`, values)}</c:val></c:ser>`
}
const axis = (kind: 'catAx' | 'valAx' | 'dateAx' | 'serAx', id: number, cross: number, pos: string, inner = '', options: { deleted?: boolean; gridlines?: boolean; scaling?: string } = {}) =>
  `<c:${kind}><c:axId val="${id}"/><c:scaling>${options.scaling ?? ''}<c:orientation val="minMax"/></c:scaling><c:delete val="${options.deleted ? 1 : 0}"/><c:axPos val="${pos}"/>${options.gridlines ? '<c:majorGridlines/>' : ''}${inner}<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/>${kind === 'valAx' ? '<c:crosses val="autoZero"/><c:crossBetween val="between"/>' : kind === 'catAx' ? '<c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>' : ''}</c:${kind}>`
/** The usual category (bottom) and value (left, gridlines) pair, ids 1 and 2. */
const catVal = (horizontal = false) => axis('catAx', 1, 2, horizontal ? 'l' : 'b') + axis('valAx', 2, 1, horizontal ? 'b' : 'l', '', { gridlines: true })
const valVal = () => axis('valAx', 1, 2, 'b').replace('<c:crossBetween val="between"/>', '<c:crossBetween val="midCat"/>') + axis('valAx', 2, 1, 'l', '', { gridlines: true }).replace('<c:crossBetween val="between"/>', '<c:crossBetween val="midCat"/>')
const labels = (position?: string, show: { val?: boolean; percent?: boolean; cat?: boolean } = { val: true }) =>
  `<c:dLbls>${position ? `<c:dLblPos val="${position}"/>` : ''}<c:showLegendKey val="0"/><c:showVal val="${show.val ? 1 : 0}"/><c:showCatName val="${show.cat ? 1 : 0}"/><c:showSerName val="0"/><c:showPercent val="${show.percent ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`
const lineFill = (color: string, width = 19050, dash?: string) => `<c:spPr><a:ln w="${width}" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${dash ? `<a:prstDash val="${dash}"/>` : ''}<a:round/></a:ln></c:spPr>`
const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4']
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']

/** ChartEx part wrapper. */
const cxSpace = (title: string, body: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cx:chartSpace xmlns:cx="${NS_CX}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">${body.replace('<cx:chart>', `<cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>${esc(title)}</cx:v></cx:txData></cx:tx></cx:title>`)}</cx:chartSpace>`
const cxLevel = (values: Array<number | string>) => `<cx:lvl ptCount="${values.length}">${values.map((v, i) => `<cx:pt idx="${i}">${typeof v === 'string' ? esc(v) : v}</cx:pt>`).join('')}</cx:lvl>`

/** The hierarchy the treemap and sunburst draw: [level 1, level 2, weight] rows (Sheet1!A2:C9). */
export const GALLERY_HIERARCHY: ReadonlyArray<readonly [string, string, number]> = [
  ['Fruit', 'Apple', 12], ['Fruit', 'Pear', 7], ['Fruit', 'Plum', 4],
  ['Veg', 'Kale', 9], ['Veg', 'Leek', 5],
  ['Nuts', 'Almond', 6], ['Nuts', 'Pecan', 3], ['Nuts', 'Cashew', 4],
]
const HIERARCHY_NAMES = [
  { name: '_xlchart.v1.0', formula: 'Sheet1!$A$2:$B$9' },
  { name: '_xlchart.v1.1', formula: 'Sheet1!$C$2:$C$9' },
]
const hierarchyCells = (): Array<{ r: number; cells: XlsxCellSpec[] }> => [
  { r: 1, cells: [{ ref: 'A1', t: 'str', v: 'Group' }, { ref: 'B1', t: 'str', v: 'Item' }, { ref: 'C1', t: 'str', v: 'Weight' }] },
  ...GALLERY_HIERARCHY.map(([a, b, v], i) => ({ r: i + 2, cells: [{ ref: `A${i + 2}`, t: 'str' as const, v: a }, { ref: `B${i + 2}`, t: 'str' as const, v: b }, { ref: `C${i + 2}`, v }] })),
]
/** The workbook a DOCX/PPTX hierarchy chart embeds (Office's `_xlchart` names over the hierarchy cells). */
export async function galleryHierarchyWorkbook(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildXlsx([{ name: 'Sheet1', rows: hierarchyCells() }]))
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', workbook.replace('</sheets>', `</sheets><definedNames>${HIERARCHY_NAMES.map(n => `<definedName name="${n.name}" hidden="1">${n.formula}</definedName>`).join('')}</definedNames>`))
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
const hierarchyPart = (layout: 'treemap' | 'sunburst', title: string, embedded: boolean) => cxSpace(title,
  `<cx:chartData>${embedded ? '<cx:externalData r:id="rId1" cx:autoUpdate="0"/>' : ''}<cx:data id="0"><cx:strDim type="cat"><cx:f dir="col">_xlchart.v1.0</cx:f></cx:strDim><cx:numDim type="size"><cx:f>_xlchart.v1.1</cx:f></cx:numDim></cx:data></cx:chartData>` +
  `<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layout}"><cx:tx><cx:txData><cx:v>Weight</cx:v></cx:txData></cx:tx><cx:dataLabels pos="ctr"><cx:visibility seriesName="0" categoryName="1" value="${layout === 'treemap' ? 1 : 0}"/></cx:dataLabels><cx:dataId val="0"/>${layout === 'treemap' ? '<cx:layoutPr><cx:parentLabelLayout val="banner"/></cx:layoutPr>' : ''}</cx:series></cx:plotAreaRegion></cx:plotArea><cx:legend pos="r" align="ctr" overlay="0"/></cx:chart>`)

/** One gallery entry: its title is drawn by the chart, so a golden shows which family each picture is. */
export interface GalleryChart { name: string; part: string; hierarchy?: boolean }

/** The classic (`c:`) gallery: every natively drawn DrawingML family and the major elements of spec 5. */
export function classicGallery(): GalleryChart[] {
  const sales = [[42, 55, 61, 48], [30, 38, 35, 44], [18, 22, 30, 27]]
  const names = ['North', 'South', 'West']
  const titleRuns = `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1200" b="1"/><a:t>Column </a:t></a:r><a:r><a:rPr lang="en-US" sz="1200" i="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr><a:t>with labels</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
  const ohlc = { open: [44, 47, 45, 50, 53], high: [49, 51, 50, 56, 58], low: [41, 44, 42, 47, 50], close: [47, 45, 49, 54, 51], volume: [1200, 900, 1500, 1800, 1100] }
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
  const serials = [45658, 45689, 45717, 45748, 45778, 45809]
  return [
    { name: 'Clustered column, title runs and data labels', part: chartSpace(undefined, `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series(0, names[0], QUARTERS, sales[0], labels('outEnd'))}${series(1, names[1], QUARTERS, sales[1])}${series(2, names[2], QUARTERS, sales[2])}<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${catVal()}`, { titleXml: titleRuns, legend: 't' }) },
    { name: 'Stacked bar', part: chartSpace('Stacked bar', `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>${names.map((n, k) => series(k, n, QUARTERS, sales[k])).join('')}<c:gapWidth val="80"/><c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${catVal(true)}`) },
    { name: 'Stacked area', part: chartSpace('Stacked area', `<c:areaChart><c:grouping val="stacked"/><c:varyColors val="0"/>${names.map((n, k) => series(k, n, MONTHS, [12 + k * 3, 18 - k, 15 + k * 2, 22 - k * 3, 19 + k, 25 - k])).join('')}<c:axId val="1"/><c:axId val="2"/></c:areaChart>${catVal()}`) },
    { name: 'Line: smoothing and a linear trendline', part: chartSpace('Line, smooth, trendline', `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series(0, 'Actual', MONTHS, [12, 15, 13, 19, 21, 24], `<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker><c:trendline><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr><c:trendlineType val="linear"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>`)}${series(1, 'Smoothed', MONTHS, [8, 14, 9, 16, 11, 18], '<c:marker><c:symbol val="none"/></c:marker><c:smooth val="1"/>')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${catVal()}`) },
    { name: 'Pie with percentages', part: chartSpace('Pie', `<c:pieChart><c:varyColors val="1"/>${series(0, 'Share', ['Rent', 'Food', 'Travel', 'Fun', 'Other'], [35, 25, 18, 12, 10], labels('bestFit', { percent: true }))}<c:firstSliceAng val="0"/></c:pieChart>`, { legend: 'r' }) },
    { name: 'Doughnut', part: chartSpace('Doughnut', `<c:doughnutChart><c:varyColors val="1"/>${series(0, 'Share', ['A', 'B', 'C', 'D'], [40, 30, 20, 10])}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`, { legend: 'r' }) },
    { name: 'Pie of pie', part: chartSpace('Pie of pie', `<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>${series(0, 'Share', ['A', 'B', 'C', 'D', 'E', 'F'], [40, 25, 15, 9, 7, 4])}<c:gapWidth val="100"/><c:splitType val="pos"/><c:splitPos val="3"/><c:secondPieSize val="60"/><c:serLines/></c:ofPieChart>`, { legend: 'r' }) },
    { name: 'Scatter: error bars and a polynomial trendline', part: chartSpace('Scatter, error bars', `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef('Sheet1!$B$1', ['Samples'])}</c:tx><c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="diamond"/><c:size val="7"/></c:marker><c:trendline><c:trendlineType val="poly"/><c:order val="2"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline><c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="0"/><c:val val="1.5"/></c:errBars><c:xVal>${numRef('Sheet1!$A$2:$A$7', [1, 2, 3, 4, 5, 6])}</c:xVal><c:yVal>${numRef('Sheet1!$B$2:$B$7', [2.1, 3.9, 6.2, 7.8, 8.4, 8.9])}</c:yVal><c:smooth val="0"/></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>${valVal()}`, { legend: null }) },
    { name: 'Bubble', part: chartSpace('Bubble', `<c:bubbleChart><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef('Sheet1!$B$1', ['Markets'])}</c:tx><c:invertIfNegative val="0"/><c:xVal>${numRef('Sheet1!$A$2:$A$5', [1, 2, 3, 4])}</c:xVal><c:yVal>${numRef('Sheet1!$B$2:$B$5', [10, 30, 20, 40])}</c:yVal><c:bubbleSize>${numRef('Sheet1!$C$2:$C$5', [4, 9, 2, 6])}</c:bubbleSize><c:bubble3D val="0"/></c:ser><c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>${valVal()}`, { legend: null }) },
    { name: 'Radar', part: chartSpace('Radar', `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${series(0, 'Team A', ['Speed', 'Power', 'Range', 'Cost', 'Comfort'], [4, 3, 5, 2, 4])}${series(1, 'Team B', ['Speed', 'Power', 'Range', 'Cost', 'Comfort'], [3, 5, 2, 4, 3])}<c:axId val="1"/><c:axId val="2"/></c:radarChart>${axis('catAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l', '', { gridlines: true })}`, { legend: 'r' }) },
    { name: 'Contour (2-D surface)', part: chartSpace('Contour', `<c:surfaceChart><c:wireframe val="0"/>${[0, 1, 2, 3].map(k => series(k, `Row ${k + 1}`, ['A', 'B', 'C', 'D', 'E'], [0, 1, 2, 3, 4].map(j => Math.round(10 * Math.sin((j + k) / 2) + 10)))).join('')}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surfaceChart>${axis('catAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l', '', { gridlines: true })}${axis('serAx', 3, 2, 'b', '', { deleted: true })}`, { legend: 'r' }) },
    { name: 'Stock: volume-open-high-low-close', part: chartSpace('Stock (volume, OHLC)', `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series(0, 'Volume', days, ohlc.volume, '<c:spPr><a:solidFill><a:srgbClr val="BDD7EE"/></a:solidFill></c:spPr>')}<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart><c:stockChart>${series(1, 'Open', days, ohlc.open)}${series(2, 'High', days, ohlc.high)}${series(3, 'Low', days, ohlc.low)}${series(4, 'Close', days, ohlc.close)}<c:hiLowLines/><c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars><c:axId val="3"/><c:axId val="4"/></c:stockChart>${axis('catAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l', '', { gridlines: true })}${axis('catAx', 3, 4, 'b', '', { deleted: true })}${axis('valAx', 4, 3, 'r', '', { scaling: '<c:max val="60"/><c:min val="30"/>' }).replace('<c:crosses val="autoZero"/>', '<c:crosses val="max"/>')}`, { legend: null }) },
    { name: 'Combination on a secondary axis', part: chartSpace('Combo, secondary axis', `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series(0, 'Revenue', QUARTERS, [120, 150, 170, 160])}<c:gapWidth val="120"/><c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series(1, 'Margin %', QUARTERS, [0.12, 0.18, 0.21, 0.16], `${lineFill('ED7D31', 28575)}<c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>`)}<c:marker val="1"/><c:axId val="3"/><c:axId val="4"/></c:lineChart>${catVal()}${axis('catAx', 3, 4, 'b', '', { deleted: true })}${axis('valAx', 4, 3, 'r', '<c:numFmt formatCode="0%" sourceLinked="0"/>').replace('<c:crosses val="autoZero"/>', '<c:crosses val="max"/>')}`) },
    { name: 'Date axis', part: chartSpace('Date axis (months)', `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef('Sheet1!$B$1', ['Visitors'])}</c:tx>${lineFill('4472C4', 22225)}<c:marker><c:symbol val="none"/></c:marker><c:cat>${numRef('Sheet1!$A$2:$A$7', serials, 'mmm\\-yy')}</c:cat><c:val>${numRef('Sheet1!$B$2:$B$7', [320, 410, 380, 520, 610, 580])}</c:val><c:smooth val="0"/></c:ser><c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axis('dateAx', 1, 2, 'b', '<c:numFmt formatCode="mmm\\-yy" sourceLinked="0"/>').replace('<c:crossAx val="2"/>', '<c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblOffset val="100"/><c:baseTimeUnit val="months"/>')}${axis('valAx', 2, 1, 'l', '', { gridlines: true })}`, { legend: null }) },
    { name: 'Logarithmic value axis', part: chartSpace('Log axis (base 10)', `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series(0, 'Count', ['A', 'B', 'C', 'D', 'E'], [3, 40, 250, 1800, 9000])}<c:gapWidth val="80"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axis('catAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l', '', { gridlines: true, scaling: '<c:logBase val="10"/>' })}`, { legend: null }) },
    { name: 'Multi-level categories', part: chartSpace('Multi-level categories', `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strRef('Sheet1!$C$1', ['Units'])}</c:tx><c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$7</c:f><c:multiLvlStrCache><c:ptCount val="6"/><c:lvl>${['Jan', 'Feb', 'Mar', 'Jan', 'Feb', 'Mar'].map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:lvl><c:lvl><c:pt idx="0"><c:v>2025</c:v></c:pt><c:pt idx="3"><c:v>2026</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat><c:val>${numRef('Sheet1!$C$2:$C$7', [14, 18, 11, 17, 22, 19])}</c:val></c:ser><c:gapWidth val="60"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${catVal()}`, { legend: null }) },
  ]
}

/** The ChartEx (`cx:`) gallery: the families the engine paints natively (spec 4.2). */
export function chartExGallery(embedded: boolean): GalleryChart[] {
  const steps = ['Start', 'Sales', 'Returns', 'Subtotal', 'Costs', 'End']
  return [
    { name: 'Waterfall', part: cxSpace('Waterfall', `<cx:chartData><cx:data id="0"><cx:strDim type="cat">${cxLevel(steps)}</cx:strDim><cx:numDim type="val">${cxLevel([100, 30, -20, 110, -40, 70])}</cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:tx><cx:txData><cx:v>Cash</cx:v></cx:txData></cx:tx><cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels><cx:dataId val="0"/><cx:layoutPr><cx:subtotals><cx:idx val="3"/><cx:idx val="5"/></cx:subtotals></cx:layoutPr></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea><cx:legend pos="b" align="ctr" overlay="0"/></cx:chart>`) },
    { name: 'Funnel', part: cxSpace('Funnel', `<cx:chartData><cx:data id="0"><cx:strDim type="cat">${cxLevel(['Leads', 'Qualified', 'Proposal', 'Negotiation', 'Won'])}</cx:strDim><cx:numDim type="val">${cxLevel([1000, 720, 480, 300, 120])}</cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"><cx:tx><cx:txData><cx:v>Pipeline</cx:v></cx:txData></cx:tx><cx:dataLabels pos="ctr"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels><cx:dataId val="0"/></cx:series></cx:plotAreaRegion><cx:axis id="1"><cx:catScaling gapWidth="0.3"/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart>`) },
    { name: 'Treemap', part: hierarchyPart('treemap', 'Treemap', embedded), hierarchy: true },
    { name: 'Sunburst', part: hierarchyPart('sunburst', 'Sunburst', embedded), hierarchy: true },
  ]
}

/** The gallery document of one format: every classic family, then the painted ChartEx families. */
export async function chartGallery(format: ChartDocFormat): Promise<Uint8Array> {
  const embedded = format !== 'xlsx'
  const workbook = embedded ? await galleryHierarchyWorkbook() : undefined
  const objects: ChartDocObject[] = [...classicGallery(), ...chartExGallery(embedded)].map(chart => ({ name: chart.name, part: chart.part, ...(chart.hierarchy && workbook ? { workbook } : {}) }))
  if (format === 'docx') return buildChartDocument('docx', objects, { columns: 2, width: 336, height: 228 })
  if (format === 'pptx') return buildChartDocument('pptx', objects, { columns: 2, rows: 2, width: 400, height: 260 })
  return buildChartDocument('xlsx', objects, { columns: 3, width: 360, height: 240, cells: hierarchyCells(), definedNames: HIERARCHY_NAMES })
}
