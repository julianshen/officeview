/** Builders for minimal but structurally-valid OOXML test files. */
import JSZip from 'jszip'

export const CT_TYPES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`

export const ROOT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

export interface DocxTextSpec {
  text: string
  bold?: boolean
  italic?: boolean
  size?: number // half-points
  color?: string
  /** Emit as a field (e.g. 'PAGE'); `text` becomes the cached result. */
  field?: 'PAGE' | 'NUMPAGES' | string
}

export interface DocxParaSpec {
  align?: 'left' | 'center' | 'right' | 'both'
  runs: DocxTextSpec[]
  /** w:numPr: list level 0-based. */
  numId?: number
  ilvl?: number
}

export interface DocxCellSpec {
  paragraphs?: DocxParaSpec[]
  gridSpan?: number
  vMerge?: 'restart' | 'continue'
  fill?: string
  borders?: string // xml fragment for tcBorders
  /** w:vAlign val */
  vAlign?: 'top' | 'center' | 'bottom'
}

export interface DocxRowSpec {
  cells: DocxCellSpec[]
  heightTwips?: string
  heightRule?: string
  /** w:trPr/w:tblHeader — repeat on continuation pages. */
  isHeader?: boolean
}

export interface DocxTableSpec {
  gridCols: string[]
  rows: DocxRowSpec[]
  fill?: string
  borders?: string // xml fragment for tblBorders
  cellMargins?: string // xml fragment for tblCellMar
}

/** Build a minimal docx buffer with the given paragraphs and tables (tables appended after paragraphs). */
export interface DocxHeaderFooterSpec {
  header?: DocxParaSpec[]
  footer?: DocxParaSpec[]
  /** w:type="first" variants, rendered only when titlePg is set. */
  firstHeader?: DocxParaSpec[]
  firstFooter?: DocxParaSpec[]
  /** Emit <w:titlePg/> so the section uses a distinct first page. */
  titlePg?: boolean
  /** w:header/w:footer margin distance from the page edge, in twips. */
  marginTwips?: number
}

export async function buildDocx(
  paras: DocxParaSpec[],
  tables: DocxTableSpec[] = [],
  hf: DocxHeaderFooterSpec = {},
): Promise<Uint8Array> {
  const zip = new JSZip()
  // list definitions: numId 1 = decimal, 2 = bullet, 3 = alpha, 4 = roman
  const usesNumbering =
    paras.some((p) => p.numId !== undefined) ||
    tables.some((t) => t.rows.some((r) => r.cells.some((c) => (c.paragraphs ?? []).some((p) => p.numId !== undefined)))) ||
    (hf.header ?? []).some((p) => p.numId !== undefined) ||
    (hf.footer ?? []).some((p) => p.numId !== undefined)
  const numberingXml = (): string => {
    const lvl = (ilvl: number, numFmt: string, lvlText: string, left: number) =>
      `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/><w:lvlText w:val="${lvlText}"/>` +
      `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`
    const abstract = (id: string, fmt: string, text: string) =>
      `<w:abstractNum w:abstractNumId="${id}">${lvl(0, fmt, text, 720)}${lvl(1, fmt, text.replace('%1', '%2'), 1440)}</w:abstractNum>`
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${abstract('0', 'decimal', '%1.')}
  ${abstract('1', 'bullet', '\u2022')}
  ${abstract('2', 'lowerLetter', '%1)')}
  ${abstract('3', 'lowerRoman', '%1.')}
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
  <w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num>
</w:numbering>`
  }
  zip.file('[Content_Types].xml', CT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  const runXml = (r: DocxTextSpec) => {
    const rpr = `<w:rPr>${r.bold ? '<w:b/>' : ''}${r.italic ? '<w:i/>' : ''}${r.size ? `<w:sz w:val="${r.size}"/>` : ''}${r.color ? `<w:color w:val="${r.color}"/>` : ''}</w:rPr>`
    const text = r.text.replace(/&/g, '&amp;').replace(/</g, '&lt;')
    if (r.field) {
      return `<w:r>${rpr}<w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r>${rpr}<w:instrText xml:space="preserve"> ${r.field} </w:instrText></w:r>` +
        `<w:r>${rpr}<w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r>${rpr}<w:t>${text}</w:t></w:r>` +
        `<w:r>${rpr}<w:fldChar w:fldCharType="end"/></w:r>`
    }
    return `<w:r>${rpr}<w:t xml:space="preserve">${text}</w:t></w:r>`
  }
  const paraXml = (p: DocxParaSpec) => {
    const bits: string[] = []
    if (p.align) bits.push(`<w:jc w:val="${p.align}"/>`)
    if (p.numId !== undefined) bits.push(`<w:numPr><w:ilvl w:val="${p.ilvl ?? 0}"/><w:numId w:val="${p.numId}"/></w:numPr>`)
    const ppr = bits.length ? `<w:pPr>${bits.join('')}</w:pPr>` : ''
    return `<w:p>${ppr}${p.runs.map(runXml).join('')}</w:p>`
  }
  const hfPart = (tag: 'hdr' | 'ftr', paragraphs: DocxParaSpec[]): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${paragraphs.map(paraXml).join('')}</w:${tag}>`
  const cellXml = (c: DocxCellSpec) => {
    const tcpr = [`<w:tcPr>`]
    if (c.gridSpan) tcpr.push(`<w:gridSpan w:val="${c.gridSpan}"/>`)
    if (c.vMerge === 'restart') tcpr.push(`<w:vMerge w:val="restart"/>`)
    else if (c.vMerge === 'continue') tcpr.push(`<w:vMerge/>`)
    if (c.fill) tcpr.push(`<w:shd w:val="clear" w:fill="${c.fill}"/>`)
    if (c.borders) tcpr.push(`<w:tcBorders>${c.borders}</w:tcBorders>`)
    if (c.vAlign) tcpr.push(`<w:vAlign w:val="${c.vAlign}"/>`)
    tcpr.push(`</w:tcPr>`)
    const ps = c.paragraphs ? c.paragraphs.map(paraXml).join('') : '<w:p/>'
    return `<w:tc>${tcpr.join('')}${ps}</w:tc>`
  }
  const tableXml = (t: DocxTableSpec) =>
    `<w:tbl><w:tblPr>${t.fill ? `<w:shd w:val="clear" w:fill="${t.fill}"/>` : ''}${t.borders ? `<w:tblBorders>${t.borders}</w:tblBorders>` : ''}${t.cellMargins ? `<w:tblCellMar>${t.cellMargins}</w:tblCellMar>` : ''}</w:tblPr>` +
    `<w:tblGrid>${t.gridCols.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    t.rows.map((r) => {
      const trpr: string[] = []
      if (r.heightTwips) trpr.push(`<w:trHeight w:val="${r.heightTwips}"${r.heightRule ? ` w:hRule="${r.heightRule}"` : ''}/>`)
      if (r.isHeader) trpr.push('<w:tblHeader/>')
      return `<w:tr>${trpr.length ? `<w:trPr>${trpr.join('')}</w:trPr>` : ''}${r.cells.map(cellXml).join('')}</w:tr>`
    }).join('') +
    `</w:tbl>`
  let document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${paras.map(paraXml).join('\n    ')}
    ${tables.map(tableXml).join('\n    ')}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`
  // header/footer parts + relationships
  let sectRefs = ''
  if (hf.header) {
    zip.file('word/header1.xml', hfPart('hdr', hf.header))
    sectRefs += '<w:headerReference w:type="default" r:id="rIdHdr"/>'
  }
  if (hf.footer) {
    zip.file('word/footer1.xml', hfPart('ftr', hf.footer))
    sectRefs += '<w:footerReference w:type="default" r:id="rIdFtr"/>'
  }
  if (hf.firstHeader) {
    zip.file('word/header2.xml', hfPart('hdr', hf.firstHeader))
    sectRefs += '<w:headerReference w:type="first" r:id="rIdHdrFirst"/>'
  }
  if (hf.firstFooter) {
    zip.file('word/footer2.xml', hfPart('ftr', hf.firstFooter))
    sectRefs += '<w:footerReference w:type="first" r:id="rIdFtrFirst"/>'
  }
  if (hf.titlePg) sectRefs += '<w:titlePg/>'
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${hf.header ? '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' : ''}
  ${hf.footer ? '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' : ''}
  ${hf.firstHeader ? '<Relationship Id="rIdHdrFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>' : ''}
  ${hf.firstFooter ? '<Relationship Id="rIdFtrFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>' : ''}
</Relationships>`)
  if (sectRefs) document = document.replace('<w:sectPr>', `<w:sectPr>${sectRefs}`)
  zip.file('word/document.xml', document)
  if (usesNumbering) {
    zip.file('word/numbering.xml', numberingXml())
  }
  zip.file('word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>`)
  return zip.generateAsync({ type: 'uint8array' })
}

const XLSX_CT =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`

const XLSX_ROOT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const XLSX_WORKBOOK_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

export interface XlsxCellSpec {
  ref: string
  /** shared string index */
  s?: number
  v?: string | number
  t?: 's' | 'n' | 'b' | 'str'
  formula?: string
  style?: number
}

export interface XlsxSheetSpec {
  name: string
  rows: Array<{ r: number; cells: XlsxCellSpec[] }>
  cols?: string
  merges?: string[]
  pageSetup?: { paperSize?: number; orientation?: string; scale?: number; fitToWidth?: number; fitToHeight?: number; fitToPage?: boolean }
  pageMargins?: { left?: number; right?: number; top?: number; bottom?: number; header?: number; footer?: number }
}

/** Build a minimal xlsx buffer. */
export async function buildXlsx(sheets: XlsxSheetSpec[], sharedStrings: string[] = []): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', XLSX_CT)
  zip.file('_rels/.rels', XLSX_ROOT_RELS)
  zip.file('xl/_rels/workbook.xml.rels', XLSX_WORKBOOK_RELS)
  const sheetXml = (sheet: XlsxSheetSpec) => {
    const setup = sheet.pageSetup ? `<pageSetup${sheet.pageSetup.paperSize !== undefined ? ` paperSize="${sheet.pageSetup.paperSize}"` : ''}${sheet.pageSetup.orientation ? ` orientation="${sheet.pageSetup.orientation}"` : ''}${sheet.pageSetup.scale !== undefined ? ` scale="${sheet.pageSetup.scale}"` : ''}${sheet.pageSetup.fitToWidth !== undefined ? ` fitToWidth="${sheet.pageSetup.fitToWidth}"` : ''}${sheet.pageSetup.fitToHeight !== undefined ? ` fitToHeight="${sheet.pageSetup.fitToHeight}"` : ''}/>` : ''
    const setupPr = sheet.pageSetup?.fitToPage !== undefined ? `<sheetPr><pageSetUpPr fitToPage="${sheet.pageSetup.fitToPage ? '1' : '0'}"/></sheetPr>` : ''
    const margins = sheet.pageMargins ? `<pageMargins${['left', 'right', 'top', 'bottom', 'header', 'footer'].map(k => ` ${k}="${(sheet.pageMargins as Record<string, number | undefined>)[k] ?? (k === 'header' || k === 'footer' ? 0.3 : 0.7)}"`).join('')}/>` : ''
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${setupPr}
  ${sheet.cols ? `<cols>${sheet.cols}</cols>` : ''}
  <sheetData>
    ${sheet.rows.map((row) => `<row r="${row.r}">${row.cells.map((c) => `<c r="${c.ref}"${c.t ? ` t="${c.t}"` : ''}${c.style !== undefined ? ` s="${c.style}"` : ''}>${c.formula !== undefined ? `<f>${c.formula}</f>` : ''}${c.v !== undefined ? `<v>${c.v}</v>` : ''}</c>`).join('')}</row>`).join('\n    ')}
  </sheetData>
  ${sheet.merges ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}
  ${margins}${setup}
</worksheet>`
  }
  const workbookXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('\n    ')}
  </sheets>
</workbook>`
  zip.file('xl/workbook.xml', workbookXml)
  for (let i = 0; i < sheets.length; i++) {
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(sheets[i]))
  }
  if (sharedStrings.length > 0) {
    zip.file('xl/sharedStrings.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">
  ${sharedStrings.map((s) => `<si><t>${s.replace(/&/g, '&amp;')}</t></si>`).join('\n  ')}
</sst>`)
  }
  zip.file('xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><name val="Calibri"/><color rgb="FF0000FF"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`)
  return zip.generateAsync({ type: 'uint8array' })
}

const PPTX_CT =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
</Types>`

const PPTX_ROOT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`


export interface PptxTableCellSpec {
  paragraphs?: Array<{
    align?: string
    runs: Array<{ text: string; b?: boolean; i?: boolean; sz?: string; color?: string }>
  }>
  gridSpan?: number
  rowSpan?: number
  /** Emitted as hMerge/vMerge (a merged-away cell). */
  merged?: boolean
  fill?: string
}

export interface PptxTableSpec {
  /** Column widths in EMU. */
  colWidths: string[]
  rows: Array<{
    /** a:tr@h in EMU (optional). */
    h?: string
    cells: PptxTableCellSpec[]
  }>
  /** a:tblPr/a:tableStyleId (must exist in the fixture tableStyles.xml). */
  styleId?: string
  firstRow?: boolean
  bandRow?: boolean
}

export interface PptxShapeSpec {
  prst?: string
  off?: [string, string]
  ext?: [string, string]
  fill?: string
  lineW?: string
  rot?: string
  paragraphs?: Array<{ align?: string; runs: Array<{ text: string; b?: boolean; i?: boolean; sz?: string; color?: string }> }>
  /** Emit as a p:graphicFrame wrapping an a:tbl. */
  table?: PptxTableSpec
  /** Embed an image part; the shape is emitted as p:pic. */
  image?: {
    data: Uint8Array
    name?: string
    /** a:srcRect crop, in 1/1000 of a percent */
    srcRect?: { l?: number; t?: number; r?: number; b?: number }
  }
}

/** A table style with the canonical banded palette, emitted per styleId. */
function tableStyleXml(styleId: string): string {
  return `<a:tblStyle styleId="${styleId}" styleName="${styleId}">
    <a:tblPr/>
    <a:tblStylePr type="wholeTable"><a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr></a:tblStylePr>
    <a:tblStylePr type="firstRow">
      <a:tcPr><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:tcPr>
      <a:txStyles><a:tcStyle><a:fontRef idx="minor"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:fontRef></a:tcStyle></a:txStyles>
    </a:tblStylePr>
    <a:tblStylePr type="band1Horz"><a:tcPr><a:solidFill><a:srgbClr val="DDEBF7"/></a:solidFill></a:tcPr></a:tblStylePr>
    <a:tblStylePr type="band2Horz"><a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr></a:tblStylePr>
  </a:tblStyle>`
}

/** Shared a:r builder for pptx fixtures (runs appear in shapes and table cells). */
function pptxRunXml(r: { text: string; b?: boolean; i?: boolean; sz?: string; color?: string }): string {
  const attrs = `${r.b ? ' b="1"' : ''}${r.i ? ' i="1"' : ''}${r.sz ? ` sz="${r.sz}"` : ''}`
  const inner = r.color ? `<a:solidFill><a:srgbClr val="${r.color}"/></a:solidFill>` : ''
  return `<a:r><a:rPr${attrs}>${inner}</a:rPr><a:t>${r.text}</a:t></a:r>`
}

function pptxParaXml(p: { align?: string; runs: Array<{ text: string; b?: boolean; i?: boolean; sz?: string; color?: string }> }): string {
  return `<a:p>${p.align ? `<a:pPr algn="${p.align}"/>` : ''}${p.runs.map(pptxRunXml).join('')}</a:p>`
}

/** Build a minimal pptx with one or more slides (single slide file reused). */
export async function buildPptx(shapes: PptxShapeSpec[]): Promise<Uint8Array> {
  const zip = new JSZip()
  // embedded images become ppt/media/imageN.* referenced by rIdImgN.
  // Identical bytes are written once, like real PowerPoint media dedupe.
  const imageRels: string[] = []
  const mediaOf: Uint8Array[] = []
  const ridsByData = new Map<Uint8Array, string>()
  zip.file('[Content_Types].xml', PPTX_CT)
  zip.file('_rels/.rels', PPTX_ROOT_RELS)
  zip.file('ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
</Relationships>`)
  zip.file('ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>`)
  const styleIds = [...new Set(shapes.map((s) => s.table?.styleId).filter((v): v is string => !!v))]
  if (styleIds.length > 0) {
    zip.file('ppt/tableStyles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">${styleIds.map((id) => tableStyleXml(id)).join('')}</a:tblStyleLst>`)
  }
  const shapeXml = (s: PptxShapeSpec) => {
    if (s.table) {
      const cellXml = (c: PptxTableCellSpec) => {
        const attrs: string[] = []
        if (c.gridSpan) attrs.push(`gridSpan="${c.gridSpan}"`)
        if (c.rowSpan) attrs.push(`rowSpan="${c.rowSpan}"`)
        if (c.merged) attrs.push('hMerge="1"')
        const tcPr = c.fill ? `<a:tcPr><a:solidFill><a:srgbClr val="${c.fill}"/></a:solidFill></a:tcPr>` : '<a:tcPr/>'
        const ps = (c.paragraphs ?? []).map(pptxParaXml).join('')
        return `<a:tc ${attrs.join(' ')}>${tcPr}<a:txBody><a:bodyPr/><a:lstStyle/>${ps || '<a:p/>'}</a:txBody></a:tc>`
      }
      const tblPr =
        `<a:tblPr${s.table.firstRow ? ' firstRow="1"' : ''}${s.table.bandRow ? ' bandRow="1"' : ''}>` +
        (s.table.styleId ? `<a:tableStyleId>${s.table.styleId}</a:tableStyleId>` : '') +
        `</a:tblPr>`
      const tbl =
        `<a:tbl>${tblPr}<a:tblGrid>${s.table.colWidths.map((w) => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>` +
        s.table.rows.map((r) => `<a:tr${r.h ? ` h="${r.h}"` : ''}>${r.cells.map(cellXml).join('')}</a:tr>`).join('') +
        `</a:tbl>`
      return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="200" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
        `<p:xfrm><a:off x="${s.off?.[0] ?? '0'}" y="${s.off?.[1] ?? '0'}"/><a:ext cx="${s.ext?.[0] ?? '100000'}" cy="${s.ext?.[1] ?? '100000'}"/></p:xfrm>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tbl}</a:graphicData></a:graphic>` +
        `</p:graphicFrame>`
    }
    if (s.image) {
      let rid = ridsByData.get(s.image.data)
      if (rid === undefined) {
        const idx = imageRels.length + 1
        const name = s.image.name ?? `image${idx}.png`
        imageRels.push(name)
        mediaOf.push(s.image.data)
        rid = `rIdImg${idx}`
        ridsByData.set(s.image.data, rid)
      }
      const picId = 100 + imageRels.indexOf(rid.replace('rIdImg', '')) + 1
      const sr = s.image.srcRect
      const srcRectXml = sr
        ? `<a:srcRect${sr.l !== undefined ? ` l="${sr.l}"` : ''}${sr.t !== undefined ? ` t="${sr.t}"` : ''}${sr.r !== undefined ? ` r="${sr.r}"` : ''}${sr.b !== undefined ? ` b="${sr.b}"` : ''}/>`
        : ''
      return `<p:pic><p:nvPicPr><p:cNvPr id="${picId}" name="Picture ${picId}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${rid}"/>${srcRectXml}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm${s.rot ? ` rot="${s.rot}"` : ''}><a:off x="${s.off?.[0] ?? '0'}" y="${s.off?.[1] ?? '0'}"/><a:ext cx="${s.ext?.[0] ?? '100000'}" cy="${s.ext?.[1] ?? '100000'}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
    }
    const spPr =
      `<p:spPr><a:xfrm${s.rot ? ` rot="${s.rot}"` : ''}><a:off x="${s.off?.[0] ?? '0'}" y="${s.off?.[1] ?? '0'}"/><a:ext cx="${s.ext?.[0] ?? '100000'}" cy="${s.ext?.[1] ?? '100000'}"/></a:xfrm>` +
      `<a:prstGeom prst="${s.prst ?? 'rect'}"><a:avLst/></a:prstGeom>` +
      (s.fill ? `<a:solidFill><a:srgbClr val="${s.fill}"/></a:solidFill>` : '') +
      `<a:ln w="${s.lineW ?? '12700'}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>`
    const txBody = s.paragraphs
      ? `<p:txBody><a:bodyPr/><a:lstStyle/>${s.paragraphs.map(pptxParaXml).join('')}</p:txBody>`
      : '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
    return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${spPr}${txBody}</p:sp>`
  }
  zip.file('ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    ${shapes.map(shapeXml).join('\n    ')}
  </p:spTree></p:cSld>
</p:sld>`)
  imageRels.forEach((name, i) => {
    zip.file(`ppt/media/${name}`, mediaOf[i])
  })
  zip.file('ppt/slides/_rels/slide1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${imageRels.map((n, i) =>
      `<Relationship Id="rIdImg${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${n}"/>`).join('')}</Relationships>`)
  return zip.generateAsync({ type: 'uint8array' })
}
