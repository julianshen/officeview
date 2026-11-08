/**
 * Stage-0 Task 5 — T2 hidden-content fixture matrix (Stage-0 subset).
 */
import JSZip from 'jszip'

export interface Fixture {
  id: string
  kind: 'docx' | 'xlsx' | 'pptx'
  withHidden: Uint8Array
  withoutHidden: Uint8Array
  marker: string
}

const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const CT_DOCX = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
const ROOT_RELS = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`

async function zipBytes(entries: Record<string, string>): Promise<Uint8Array> {
  const zip = new JSZip()
  for (const [p, v] of Object.entries(entries)) zip.file(p, v)
  return (await zip.generateAsync({ type: 'uint8array' })) as Uint8Array
}

function docxBytes(bodyInner: string): Promise<Uint8Array> {
  return zipBytes({
    '[Content_Types].xml': CT_DOCX,
    '_rels/.rels': ROOT_RELS,
    'word/document.xml': `<w:document ${WNS}><w:body>${bodyInner}</w:body></w:document>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
  })
}

const SS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const CT_XLSX = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
const ROOT_RELS_X = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
const WB_RELS_X = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`

function xlsxBytes(sheetData: string, workbookInner = '<sheet name="S" sheetId="1" r:id="rId1"/>', sheetExtra = ''): Promise<Uint8Array> {
  return zipBytes({
    '[Content_Types].xml': CT_XLSX,
    '_rels/.rels': ROOT_RELS_X,
    'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns="${SS_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${workbookInner}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': WB_RELS_X,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="${SS_NS}">${sheetExtra}<sheetData>${sheetData}</sheetData></worksheet>`,
    'xl/styles.xml': `<?xml version="1.0"?><styleSheet xmlns="${SS_NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`,
  })
}

const CT_PPTX = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`
const ROOT_RELS_P = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`
const PRES_RELS = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`
const PRES = `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="s"/></p:sldIdLst></p:presentation>`

function pptxBytes(slideInner: string, extra: Record<string,string> = {}): Promise<Uint8Array> {
  return zipBytes({
    '[Content_Types].xml': CT_PPTX,
    '_rels/.rels': ROOT_RELS_P,
    'ppt/presentation.xml': PRES,
    'ppt/_rels/presentation.xml.rels': PRES_RELS,
    'ppt/slides/slide1.xml': `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>${slideInner}</p:spTree></p:cSld></p:sld>`,
    ...extra,
  })
}

const SHAPE = (text: string, extra = '') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s" ${extra}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`

export async function hiddenFixtures(): Promise<Fixture[]> {
  const out: Fixture[] = []
  const _dbg=(x:string)=>console.error("DBG",x)
  const mk = (id: string, kind: Fixture['kind'], withHidden: Uint8Array, withoutHidden: Uint8Array, marker: string) =>
    out.push({ id, kind, withHidden, withoutHidden, marker })

  // xlsx hidden sheet
  {
    const marker = 'MARKER-xlsx-hidden-sheet'
    const hiddenWb = `<sheets><sheet name="S" sheetId="1" r:id="rId1"/><sheet name="H" sheetId="2" state="hidden" r:id="rId2"/></sheets>`
    // need rels for sheet2; reuse single-sheet rels by inlining marker into sheet1? Instead: hidden sheet is a SECOND sheet; build a 2-sheet package.
    const zipH = new (await import('jszip')).default()
    zipH.file('[Content_Types].xml', CT_XLSX.replace('/xl/worksheets/sheet1.xml', '/xl/worksheets/sheet1.xml').replace('sheet1.xml', 'sheet1.xml'))
    // Add content type for sheet2
    zipH.file('_rels/.rels', ROOT_RELS_X)
    zipH.file('xl/workbook.xml', `<?xml version="1.0"?><workbook xmlns="${SS_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/><sheet name="H" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`)
    zipH.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`)
    zipH.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/worksheets/sheet2.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${marker}</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/styles.xml', `<?xml version="1.0"?><styleSheet xmlns="${SS_NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`)
    const withH = (await zipH.generateAsync({ type: 'uint8array' })) as Uint8Array
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row>`)
    mk('xlsx-hidden-sheet', 'xlsx', withH, withoutH, marker)
  }

  // xlsx veryhidden
  {
    const marker = 'MARKER-xlsx-veryhidden-sheet'
    const zipH = new (await import('jszip')).default()
    zipH.file('[Content_Types].xml', CT_XLSX.replace('sheet1.xml', 'sheet1.xml'))
    zipH.file('_rels/.rels', ROOT_RELS_X)
    zipH.file('xl/workbook.xml', `<?xml version="1.0"?><workbook xmlns="${SS_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/><sheet name="H" sheetId="2" state="veryHidden" r:id="rId2"/></sheets></workbook>`)
    zipH.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`)
    zipH.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/worksheets/sheet2.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${marker}</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/styles.xml', `<?xml version="1.0"?><styleSheet xmlns="${SS_NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`)
    const withH = (await zipH.generateAsync({ type: 'uint8array' })) as Uint8Array
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row>`)
    mk('xlsx-veryhidden-sheet', 'xlsx', withH, withoutH, marker)
  }

  // xlsx hidden rows
  {
    const marker = 'MARKER-xlsx-hidden-rows'
    const withH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row><row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>${marker}</t></is></c></row>`)
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row>`)
    mk('xlsx-hidden-rows', 'xlsx', withH, withoutH, marker)
  }

  // xlsx hidden cols
  {
    const marker = 'MARKER-xlsx-hidden-cols'
    const withH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c><c r="B1" t="inlineStr"><is><t>${marker}</t></is></c></row>`, '', `<cols><col min="2" max="2" width="10" hidden="1"/></cols>`)
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row>`)
    mk('xlsx-hidden-cols', 'xlsx', withH, withoutH, marker)
  }

  // docx vanish
  {
    const marker = 'MARKER-docx-vanish'
    const withH = await docxBytes(`<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>${marker}</w:t></w:r></w:p><w:p><w:r><w:t>VIS</w:t></w:r></w:p>`)
    const withoutH = await docxBytes(`<w:p><w:r><w:t>VIS</w:t></w:r></w:p>`)
    mk('docx-vanish', 'docx', withH, withoutH, marker)
  }

  // docx hidden shape docPr
  {
    const marker = 'MARKER-docx-hidden-shape-docPr'
    const hiddenShape = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="h" hidden="1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><a:txBody><a:p><a:r><a:t>${marker}</a:t></a:r></a:p></a:txBody></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const withH = await docxBytes(hiddenShape + `<w:p><w:r><w:t>VIS</w:t></w:r></w:p>`)
    const withoutH = await docxBytes(`<w:p><w:r><w:t>VIS</w:t></w:r></w:p>`)
    mk('docx-hidden-shape-docPr', 'docx', withH, withoutH, marker)
  }

  // docx hidden shape in group — model as a second paragraph marker inside a group shape? Keep simple: vanish in group not supported; use a docPr hidden shape inside a group container (w:group). We'll just reuse docPr hidden inside a group wrapper paragraph.
  {
    const marker = 'MARKER-docx-hidden-shape-in-group'
    const grp = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:extent cx="1828800" cy="914400"/><wp:docPr id="3" name="g"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wps:wsp><wps:cNvPr id="4" name="inner" hidden="1"/><wps:spPr/><wps:txbx><w:txbxContent><w:p><w:r><w:t>${marker}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const withH = await docxBytes(grp + `<w:p><w:r><w:t>VIS</w:t></w:r></w:p>`)
    const withoutH = await docxBytes(`<w:p><w:r><w:t>VIS</w:t></w:r></w:p>`)
    mk('docx-hidden-shape-in-group', 'docx', withH, withoutH, marker)
  }

  // pptx hidden shape cNvPr
  {
    const marker = 'MARKER-pptx-hidden-shape-cNvPr'
    const withH = await pptxBytes(SHAPE(marker, 'hidden="1"') + SHAPE('VIS'))
    const withoutH = await pptxBytes(SHAPE('VIS'))
    mk('pptx-hidden-shape-cNvPr', 'pptx', withH, withoutH, marker)
  }

  // pptx hidden shape in group
  {
    const marker = 'MARKER-pptx-hidden-shape-in-group'
    const grp = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${SHAPE(marker, 'hidden="1"')}</p:grpSp>`
    const withH = await pptxBytes(grp + SHAPE('VIS'))
    const withoutH = await pptxBytes(SHAPE('VIS'))
    mk('pptx-hidden-shape-in-group', 'pptx', withH, withoutH, marker)
  }

  // pptx hidden slide
  {
    const marker = 'MARKER-pptx-hidden-slide'
    const presHidden = `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="s" show="0"/></p:sldIdLst></p:presentation>`
    const slideWithMarker = `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:p><a:r><a:t>${marker}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
    const zipH = new (await import('jszip')).default()
    zipH.file('[Content_Types].xml', CT_PPTX)
    zipH.file('_rels/.rels', ROOT_RELS_P)
    zipH.file('ppt/presentation.xml', presHidden)
    zipH.file('ppt/_rels/presentation.xml.rels', PRES_RELS)
    zipH.file('ppt/slides/slide1.xml', slideWithMarker)
    const withH = (await zipH.generateAsync({ type: 'uint8array' })) as Uint8Array
    const withoutH = await pptxBytes(SHAPE('VIS'))
    mk('pptx-hidden-slide', 'pptx', withH, withoutH, marker)
  }

  // pptx notes
  {
    const marker = 'MARKER-pptx-notes'
    const notesXml = `<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="n"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:p><a:r><a:t>${marker}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`
    const extra = { 'ppt/notesSlides/notesSlide1.xml': notesXml }
    const withH = await pptxBytes(SHAPE('VIS'), extra)
    const withoutH = await pptxBytes(SHAPE('VIS'))
    mk('pptx-notes', 'pptx', withH, withoutH, marker)
  }

  // xlsx comment
  {
    const marker = 'MARKER-xlsx-comment'
    const withH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row>`, '', `<sheetComments><comment ref="A1" authorId="0"><text><r><t>${marker}</t></r></text></comment></sheetComments>`)
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>VIS</t></is></c></row>`)
    mk('xlsx-comment', 'xlsx', withH, withoutH, marker)
  }

  // docx comment
  {
    const marker = 'MARKER-docx-comment'
    const withH = await docxBytes(`<w:p><w:r><w:t>VIS</w:t></w:r></w:p>`, {
      'word/comments.xml': `<?xml version="1.0"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="a" w:date="2020-01-01T00:00:00Z"><w:p><w:r><w:t>${marker}</w:t></w:r></w:p></w:comment></w:comments>`,
    })
    const withoutH = await docxBytes(`<w:p><w:r><w:t>VIS</w:t></w:r></w:p>`)
    mk('docx-comment', 'docx', withH, withoutH, marker)
  }

  return out
}
