import JSZip from 'jszip'

export interface Fixture {
  id: string
  kind: 'docx' | 'xlsx' | 'pptx'
  withHidden: Uint8Array
  withoutHidden: Uint8Array
  marker: string
  control: string
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
function pptxBytes(slideInner: string, extra: Record<string, string> = {}): Promise<Uint8Array> {
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
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s" ${extra}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`

export async function hiddenFixtures(): Promise<Fixture[]> {
  const out: Fixture[] = []
  const mk = (id: string, kind: Fixture['kind'], withHidden: Uint8Array, withoutHidden: Uint8Array, marker: string) =>
    out.push({ id, kind, withHidden, withoutHidden, marker, control: `VIS-${id}` })

  // xlsx hidden sheet
  {
    const id = 'xlsx-hidden-sheet'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const zipH = new (await import('jszip')).default()
    zipH.file('[Content_Types].xml', CT_XLSX)
    zipH.file('_rels/.rels', ROOT_RELS_X)
    zipH.file('xl/workbook.xml', `<?xml version="1.0"?><workbook xmlns="${SS_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/><sheet name="H" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`)
    zipH.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`)
    zipH.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/worksheets/sheet2.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${marker}</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/styles.xml', `<?xml version="1.0"?><styleSheet xmlns="${SS_NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`)
    const withH = (await zipH.generateAsync({ type: 'uint8array' })) as Uint8Array
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row>`)
    mk(id, 'xlsx', withH, withoutH, marker)
  }

  // xlsx veryhidden
  {
    const id = 'xlsx-veryhidden-sheet'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const zipH = new (await import('jszip')).default()
    zipH.file('[Content_Types].xml', CT_XLSX)
    zipH.file('_rels/.rels', ROOT_RELS_X)
    zipH.file('xl/workbook.xml', `<?xml version="1.0"?><workbook xmlns="${SS_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/><sheet name="H" sheetId="2" state="veryHidden" r:id="rId2"/></sheets></workbook>`)
    zipH.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`)
    zipH.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/worksheets/sheet2.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${marker}</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/styles.xml', `<?xml version="1.0"?><styleSheet xmlns="${SS_NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`)
    const withH = (await zipH.generateAsync({ type: 'uint8array' })) as Uint8Array
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row>`)
    mk(id, 'xlsx', withH, withoutH, marker)
  }

  // xlsx hidden rows
  {
    const id = 'xlsx-hidden-rows'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const withH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row><row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>${marker}</t></is></c></row>`)
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row>`)
    mk(id, 'xlsx', withH, withoutH, marker)
  }

  // xlsx hidden cols
  {
    const id = 'xlsx-hidden-cols'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const withH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c><c r="B1" t="inlineStr"><is><t>${marker}</t></is></c></row>`, '<sheet name="S" sheetId="1" r:id="rId1"/>', `<cols><col min="2" max="2" width="10" hidden="1"/></cols>`)
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row>`)
    mk(id, 'xlsx', withH, withoutH, marker)
  }

  // docx vanish
  {
    const id = 'docx-vanish'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const withH = await docxBytes(`<w:p><w:r><w:t>${control}</w:t></w:r></w:p><w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>${marker}</w:t></w:r></w:p>`)
    const withoutH = await docxBytes(`<w:p><w:r><w:t>${control}</w:t></w:r></w:p>`)
    mk(id, 'docx', withH, withoutH, marker)
  }

  // docx hidden shape docPr
  {
    const id = 'docx-hidden-shape-docPr'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const hiddenShape = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="h" hidden="1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><a:txBody><a:p><a:r><a:t>${marker}</a:t></a:r></a:p></a:txBody></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const withH = await docxBytes(`<w:p><w:r><w:t>${control}</w:t></w:r></w:p>` + hiddenShape)
    const withoutH = await docxBytes(`<w:p><w:r><w:t>${control}</w:t></w:r></w:p>`)
    mk(id, 'docx', withH, withoutH, marker)
  }

  // docx hidden shape in group
  {
    const id = 'docx-hidden-shape-in-group'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const grp = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wp:extent cx="1828800" cy="914400"/><wp:docPr id="3" name="g"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wps:wsp><wps:cNvPr id="4" name="inner" hidden="1"/><wps:spPr/><wps:txbx><w:txbxContent><w:p><w:r><w:t>${marker}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const withH = await docxBytes(`<w:p><w:r><w:t>${control}</w:t></w:r></w:p>` + grp)
    const withoutH = await docxBytes(`<w:p><w:r><w:t>${control}</w:t></w:r></w:p>`)
    mk(id, 'docx', withH, withoutH, marker)
  }

  // pptx hidden shape cNvPr
  {
    const id = 'pptx-hidden-shape-cNvPr'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const withH = await pptxBytes(SHAPE(control) + SHAPE(marker, 'hidden="1"'))
    const withoutH = await pptxBytes(SHAPE(control))
    mk(id, 'pptx', withH, withoutH, marker)
  }

  // pptx hidden shape in group
  {
    const id = 'pptx-hidden-shape-in-group'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const grp = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${SHAPE(marker, 'hidden="1"')}</p:grpSp>`
    const withH = await pptxBytes(SHAPE(control) + grp)
    const withoutH = await pptxBytes(SHAPE(control))
    mk(id, 'pptx', withH, withoutH, marker)
  }

  // pptx hidden slide
  {
    const id = 'pptx-hidden-slide'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const presHidden = `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="s"/><p:sldId id="257" r:id="s2" show="0"/></p:sldIdLst></p:presentation>`
    const presRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="s2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>`
    const zipH = new (await import('jszip')).default()
    zipH.file('[Content_Types].xml', CT_PPTX)
    zipH.file('_rels/.rels', ROOT_RELS_P)
    zipH.file('ppt/presentation.xml', presHidden)
    zipH.file('ppt/_rels/presentation.xml.rels', presRels)
    zipH.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>${SHAPE(control)}</p:spTree></p:cSld></p:sld>`)
    zipH.file('ppt/slides/slide2.xml', `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>${SHAPE(marker)}</p:spTree></p:cSld></p:sld>`)
    const withH = (await zipH.generateAsync({ type: 'uint8array' })) as Uint8Array
    const withoutH = await pptxBytes(SHAPE(control))
    mk(id, 'pptx', withH, withoutH, marker)
  }

  // pptx notes
  {
    const id = 'pptx-notes'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const extra = { 'ppt/notesSlides/notesSlide1.xml': `<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="n"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:p><a:r><a:t>${marker}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>` }
    const withH = await pptxBytes(SHAPE(control), extra)
    const withoutH = await pptxBytes(SHAPE(control))
    mk(id, 'pptx', withH, withoutH, marker)
  }

  // xlsx comment (comments part + rels + VML drawing)
  {
    const id = 'xlsx-comment'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const zipH = new (await import('jszip')).default()
    zipH.file('[Content_Types].xml', CT_XLSX)
    zipH.file('_rels/.rels', ROOT_RELS_X)
    zipH.file('xl/workbook.xml', `<?xml version="1.0"?><workbook xmlns="${SS_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`)
    zipH.file('xl/_rels/workbook.xml.rels', WB_RELS_X)
    zipH.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="${SS_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row></sheetData></worksheet>`)
    zipH.file('xl/styles.xml', `<?xml version="1.0"?><styleSheet xmlns="${SS_NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`)
    const withH = (await zipH.generateAsync({ type: 'uint8array' })) as Uint8Array
    const withoutH = await xlsxBytes(`<row r="1"><c r="A1" t="inlineStr"><is><t>${control}</t></is></c></row>`)
    mk(id, 'xlsx', withH, withoutH, marker)
  }

  // docx comment
  {
    const id = 'docx-comment'; const marker = `MARKER-${id}`; const control = `VIS-${id}`
    const withH = await docxBytes(`<w:p><w:r><w:t>${control}</w:t></w:r></w:p>`)
    const withoutH = await docxBytes(`<w:p><w:r><w:t>${control}</w:t></w:r></w:p>`)
    mk(id, 'docx', withH, withoutH, marker)
  }

  return out
}
