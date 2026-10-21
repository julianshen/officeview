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

const XLSX_ROOT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

export interface XlsxCellSpec {
  ref: string
  v?: string | number
  t?: 's' | 'n' | 'b' | 'str' | 'e'
  formula?: string
  sharedFormula?: { si: number; ref?: string }
  ca?: boolean
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
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n  ')}
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
  zip.file('[Content_Types].xml', contentTypesXml)
  zip.file('_rels/.rels', XLSX_ROOT_RELS)

  const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('\n  ')}
  <Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId${sheets.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  zip.file('xl/_rels/workbook.xml.rels', workbookRelsXml)
  const escXml = (f: string) => f.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  const sheetXml = (sheet: XlsxSheetSpec) => {
    const setup = sheet.pageSetup ? `<pageSetup${sheet.pageSetup.paperSize !== undefined ? ` paperSize="${sheet.pageSetup.paperSize}"` : ''}${sheet.pageSetup.orientation ? ` orientation="${sheet.pageSetup.orientation}"` : ''}${sheet.pageSetup.scale !== undefined ? ` scale="${sheet.pageSetup.scale}"` : ''}${sheet.pageSetup.fitToWidth !== undefined ? ` fitToWidth="${sheet.pageSetup.fitToWidth}"` : ''}${sheet.pageSetup.fitToHeight !== undefined ? ` fitToHeight="${sheet.pageSetup.fitToHeight}"` : ''}/>` : ''
    const setupPr = sheet.pageSetup?.fitToPage !== undefined ? `<sheetPr><pageSetUpPr fitToPage="${sheet.pageSetup.fitToPage ? '1' : '0'}"/></sheetPr>` : ''
    const margins = sheet.pageMargins ? `<pageMargins${['left', 'right', 'top', 'bottom', 'header', 'footer'].map(k => ` ${k}="${(sheet.pageMargins as Record<string, number | undefined>)[k] ?? (k === 'header' || k === 'footer' ? 0.3 : 0.7)}"`).join('')}/>` : ''
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${setupPr}
  ${sheet.cols ? `<cols>${sheet.cols}</cols>` : ''}
  <sheetData>
    ${sheet.rows.map((row) => `<row r="${row.r}">${row.cells.map((c) => {
      const fXml = c.sharedFormula !== undefined
        ? `<f t="shared"${c.sharedFormula.ref ? ` ref="${c.sharedFormula.ref}"` : ''} si="${c.sharedFormula.si}">${c.formula !== undefined ? escXml(c.formula) : ''}</f>`
        : c.formula !== undefined
          ? `<f>${escXml(c.formula)}</f>`
          : ''
      return `<c r="${c.ref}"${c.ca ? ' ca="1"' : ''}${c.t ? ` t="${c.t}"` : ''}${c.style !== undefined ? ` s="${c.style}"` : ''}>${fXml}${c.v !== undefined ? `<v>${typeof c.v === 'string' ? escXml(c.v) : c.v}</v>` : ''}</c>`
    }).join('')}</row>`).join('\n    ')}
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
  borders?: string
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
  /** Raw <a:tblStyle> XML overriding the canned style entry for styleId (test-only). */
  customStyleXml?: string
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
  /** p:cNvPr id; defaults to today's values (2 for sp, 200 for tables). */
  id?: number
  /** p:cNvPr name; defaults to today's values ("Shape", "Table", "Picture N"). */
  name?: string
  /** Mark the shape as a p:nvPr/p:ph placeholder of the layout/master. */
  placeholder?: { type: string; idx?: number }
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

/** The graphicFrame body of a table shape (shared by buildPptx and buildPptxDeck). */
function pptxTableXml(
  table: PptxTableSpec,
  off: [string, string] | undefined,
  ext: [string, string] | undefined,
  id: number | string = 200,
  name = 'Table',
): string {
  const cellXml = (c: PptxTableCellSpec) => {
    const attrs: string[] = []
    if (c.gridSpan) attrs.push(`gridSpan="${c.gridSpan}"`)
    if (c.rowSpan) attrs.push(`rowSpan="${c.rowSpan}"`)
    if (c.merged) attrs.push('hMerge="1"')
    const tcPrContent = `${c.fill ? `<a:solidFill><a:srgbClr val="${c.fill}"/></a:solidFill>` : ''}${c.borders ?? ''}`
    const tcPr = tcPrContent ? `<a:tcPr>${tcPrContent}</a:tcPr>` : '<a:tcPr/>'
    const ps = (c.paragraphs ?? []).map(pptxParaXml).join('')
    return `<a:tc ${attrs.join(' ')}>${tcPr}<a:txBody><a:bodyPr/><a:lstStyle/>${ps || '<a:p/>'}</a:txBody></a:tc>`
  }
  const tblPr =
    `<a:tblPr${table.firstRow ? ' firstRow="1"' : ''}${table.bandRow ? ' bandRow="1"' : ''}>` +
    (table.styleId ? `<a:tableStyleId>${table.styleId}</a:tableStyleId>` : '') +
    `</a:tblPr>`
  const tbl =
    `<a:tbl>${tblPr}<a:tblGrid>${table.colWidths.map((w) => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>` +
    table.rows.map((r) => `<a:tr${r.h ? ` h="${r.h}"` : ''}>${r.cells.map(cellXml).join('')}</a:tr>`).join('') +
    `</a:tbl>`
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${off?.[0] ?? '0'}" y="${off?.[1] ?? '0'}"/><a:ext cx="${ext?.[0] ?? '100000'}" cy="${ext?.[1] ?? '100000'}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tbl}</a:graphicData></a:graphic>` +
    `</p:graphicFrame>`
}

/** A p:pic shape over an already-allocated media relationship (shared by buildPptx and buildPptxDeck). */
function pptxPicXml(
  rid: string,
  picId: number | string,
  pic: { srcRect?: { l?: number; t?: number; r?: number; b?: number }; rot?: string; off?: [string, string]; ext?: [string, string] },
  name?: string,
): string {
  const sr = pic.srcRect
  const srcRectXml = sr
    ? `<a:srcRect${sr.l !== undefined ? ` l="${sr.l}"` : ''}${sr.t !== undefined ? ` t="${sr.t}"` : ''}${sr.r !== undefined ? ` r="${sr.r}"` : ''}${sr.b !== undefined ? ` b="${sr.b}"` : ''}/>`
    : ''
  return `<p:pic><p:nvPicPr><p:cNvPr id="${picId}" name="${name ?? `Picture ${picId}`}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rid}"/>${srcRectXml}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm${pic.rot ? ` rot="${pic.rot}"` : ''}><a:off x="${pic.off?.[0] ?? '0'}" y="${pic.off?.[1] ?? '0'}"/><a:ext cx="${pic.ext?.[0] ?? '100000'}" cy="${pic.ext?.[1] ?? '100000'}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
}

/** A plain p:sp shape (shared by buildPptx and buildPptxDeck). */
function pptxSpXml(s: PptxShapeSpec, id: number | string, name: string): string {
  const spPr =
    `<p:spPr><a:xfrm${s.rot ? ` rot="${s.rot}"` : ''}><a:off x="${s.off?.[0] ?? '0'}" y="${s.off?.[1] ?? '0'}"/><a:ext cx="${s.ext?.[0] ?? '100000'}" cy="${s.ext?.[1] ?? '100000'}"/></a:xfrm>` +
    `<a:prstGeom prst="${s.prst ?? 'rect'}"><a:avLst/></a:prstGeom>` +
    (s.fill ? `<a:solidFill><a:srgbClr val="${s.fill}"/></a:solidFill>` : '') +
    `<a:ln w="${s.lineW ?? '12700'}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>`
  const txBody = s.paragraphs
    ? `<p:txBody><a:bodyPr/><a:lstStyle/>${s.paragraphs.map(pptxParaXml).join('')}</p:txBody>`
    : '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${spPr}${txBody}</p:sp>`
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
  const customStyles = new Map<string, string>()
  for (const s of shapes) {
    if (s.table?.styleId && s.table.customStyleXml) customStyles.set(s.table.styleId, s.table.customStyleXml)
  }
  if (styleIds.length > 0) {
    zip.file('ppt/tableStyles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">${styleIds.map((id) => customStyles.get(id) ?? tableStyleXml(id)).join('')}</a:tblStyleLst>`)
  }
  const shapeXml = (s: PptxShapeSpec) => {
    if (s.table) {
      return pptxTableXml(s.table, s.off, s.ext)
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
      return pptxPicXml(rid, picId, { srcRect: s.image.srcRect, rot: s.rot, off: s.off, ext: s.ext })
    }
    return pptxSpXml(s, 2, 'Shape')
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

export type PptxBgSpec =
  | { solid: string }
  | { gradient: { /** lin angle in degrees (converted to 1/60000 deg in the XML). */ angle: number; stops: Array<[pos: number, rgb: string]> } }
  | { picture: { data: Uint8Array; tile?: boolean } }
  | { pattern: { prst: string; fg: string; bg: string } }
  | { bgRef: { idx: number; color: string } }

export interface PptxLayoutSpec {
  bg?: PptxBgSpec
  shapes?: PptxShapeSpec[]
  showMasterSp?: boolean
  placeholders?: Array<{ type: string; idx?: number; off: [string, string]; ext: [string, string] }>
}

export interface PptxMasterSpec {
  bg?: PptxBgSpec
  shapes?: PptxShapeSpec[]
  layouts: PptxLayoutSpec[]
}

export interface PptxSlideSpec {
  layout?: [master: number, layout: number]
  bg?: PptxBgSpec
  shapes: PptxShapeSpec[]
  showMasterSp?: boolean
  extraXml?: { beforeSpTree?: string; afterCSld?: string }
}

/** The raw a:bgFillStyleLst children that theme.bgFillStyles is parsed from. */
export interface PptxThemeSpec {
  bgFillStyles?: string[]
}

export interface PptxDeckSpec {
  /** p:sldSz cx/cy in EMU; defaults to buildPptx's 9144000x6858000. */
  slideSize?: [number, number]
  masters?: PptxMasterSpec[]
  slides: PptxSlideSpec[]
  /** Whole ppt/presProps.xml part content, written verbatim. */
  presProps?: string
  theme?: PptxThemeSpec
}

const PPTX_REL = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  presProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
  tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
}

/** p:bg from a PptxBgSpec; `allocImage` registers media and returns the part-local rId. */
function pptxBgXml(bg: PptxBgSpec | undefined, allocImage: (data: Uint8Array, name?: string) => string): string {
  if (!bg) return ''
  const bgPr = (children: string) => `<p:bg><p:bgPr>${children}<a:effectLst/></p:bgPr></p:bg>`
  if ('solid' in bg) return bgPr(`<a:solidFill><a:srgbClr val="${bg.solid}"/></a:solidFill>`)
  if ('gradient' in bg) {
    const gs = bg.gradient.stops.map(([pos, rgb]) => `<a:gs pos="${pos}"><a:srgbClr val="${rgb}"/></a:gs>`).join('')
    return bgPr(`<a:gradFill><a:gsLst>${gs}</a:gsLst><a:lin ang="${Math.round(bg.gradient.angle * 60000)}" scaled="1"/></a:gradFill>`)
  }
  if ('picture' in bg) {
    const fill = bg.picture.tile ? '<a:tile/>' : '<a:stretch><a:fillRect/></a:stretch>'
    return bgPr(`<a:blipFill><a:blip r:embed="${allocImage(bg.picture.data)}"/>${fill}</a:blipFill>`)
  }
  if ('pattern' in bg) {
    return bgPr(`<a:pattFill prst="${bg.pattern.prst}"><a:fgClr><a:srgbClr val="${bg.pattern.fg}"/></a:fgClr><a:bgClr><a:srgbClr val="${bg.pattern.bg}"/></a:bgClr></a:pattFill>`)
  }
  const color = /^[0-9A-Fa-f]{6}$/.test(bg.bgRef.color)
    ? `<a:srgbClr val="${bg.bgRef.color}"/>`
    : `<a:schemeClr val="${bg.bgRef.color}"/>`
  return `<p:bg><p:bgRef idx="${bg.bgRef.idx}">${color}</p:bgRef></p:bg>`
}

/** A placeholder p:sp; xfrm is emitted only when the shape carries an explicit position. */
function pptxPlaceholderXml(s: PptxShapeSpec, id: number, name: string): string {
  const ph = s.placeholder!
  const phTag = `<p:ph type="${ph.type}"${ph.idx !== undefined ? ` idx="${ph.idx}"` : ''}/>`
  const xfrm = s.off && s.ext ? `<a:xfrm><a:off x="${s.off[0]}" y="${s.off[1]}"/><a:ext cx="${s.ext[0]}" cy="${s.ext[1]}"/></a:xfrm>` : ''
  const body = s.paragraphs ? s.paragraphs.map(pptxParaXml).join('') : '<a:p/>'
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${phTag}</p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`
}

/** A layout placeholder p:sp with the geometry the layout defines for it. */
function pptxLayoutPlaceholderXml(
  ph: { type: string; idx?: number; off: [string, string]; ext: [string, string] },
  id: number,
): string {
  const phTag = `<p:ph type="${ph.type}"${ph.idx !== undefined ? ` idx="${ph.idx}"` : ''}/>`
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${ph.type}"/><p:cNvSpPr/><p:nvPr>${phTag}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${ph.off[0]}" y="${ph.off[1]}"/><a:ext cx="${ph.ext[0]}" cy="${ph.ext[1]}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`
}

/** Per-part image relationship bookkeeping (rIds are part-local, media files are deck-global). */
interface PptxPartRels {
  allocImage: (data: Uint8Array, name?: string) => string
  idFor: (s: PptxShapeSpec) => number
  nextId: () => number
  imageRelsXml: () => string
}

function pptxPart(
  allocGlobalMedia: (data: Uint8Array, name?: string) => string,
  explicitIds: number[] = [],
): PptxPartRels {
  const rels: string[] = []
  const byRef = new Map<Uint8Array, string>()
  const used = new Set<number>(explicitIds)
  let counter = 1
  function nextId(): number {
    do { counter += 1 } while (used.has(counter))
    used.add(counter)
    return counter
  }
  return {
    allocImage: (data, name) => {
      const hit = byRef.get(data)
      if (hit) return hit
      const media = allocGlobalMedia(data, name)
      const rid = `rIdImg${rels.length + 1}`
      rels.push(`<Relationship Id="${rid}" Type="${PPTX_REL.image}" Target="../media/${media}"/>`)
      byRef.set(data, rid)
      return rid
    },
    idFor: (s) => {
      if (s.id !== undefined) {
        used.add(s.id)
        return s.id
      }
      return nextId()
    },
    nextId,
    imageRelsXml: () => rels.join(''),
  }
}

const PPTX_MEDIA_CT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
}

/** Build a multi-slide deck with layouts, masters, backgrounds and a theme. */
export async function buildPptxDeck(spec: PptxDeckSpec): Promise<Uint8Array> {
  const zip = new JSZip()
  const masters = spec.masters ?? []
  const slides = spec.slides
  const mediaFiles: Array<{ name: string; ext: string; data: Uint8Array }> = []
  const mediaByRef = new Map<Uint8Array, string>()
  const allocGlobalMedia = (data: Uint8Array, name?: string): string => {
    const hit = mediaByRef.get(data)
    if (hit) return hit
    const ext = name && /\.[^./]+$/.test(name) ? name.slice(name.lastIndexOf('.') + 1).toLowerCase() : 'png'
    const entry = `image${mediaFiles.length + 1}.${ext}`
    mediaFiles.push({ name: entry, ext, data })
    mediaByRef.set(data, entry)
    return entry
  }

  // layouts are numbered globally across masters
  const layoutGlobals: Array<{ master: number; index: number; number: number }> = []
  masters.forEach((m, mi) => m.layouts.forEach((_, li) => layoutGlobals.push({ master: mi, index: li, number: layoutGlobals.length + 1 })))
  const layoutNumberOf = (mi: number, li: number): number | undefined =>
    layoutGlobals.find((l) => l.master === mi && l.index === li)?.number

  const spTreeOpen = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>\n    <p:grpSpPr/>'
  const deckShapeXml = (s: PptxShapeSpec, part: PptxPartRels): string => {
    if (s.table) return pptxTableXml(s.table, s.off, s.ext, part.idFor(s), s.name ?? 'Table')
    if (s.image) {
      const rid = part.allocImage(s.image.data, s.image.name)
      return pptxPicXml(rid, part.idFor(s), { srcRect: s.image.srcRect, rot: s.rot, off: s.off, ext: s.ext }, s.name)
    }
    if (s.placeholder) return pptxPlaceholderXml(s, part.idFor(s), s.name ?? s.placeholder.type)
    return pptxSpXml(s, part.idFor(s), s.name ?? 'Shape')
  }

  const masterXmls = masters.map((m, mi) => {
    const part = pptxPart(allocGlobalMedia, (m.shapes ?? []).map((s) => s.id).filter((v): v is number => v !== undefined))
    const bg = pptxBgXml(m.bg, part.allocImage)
    const shapes = (m.shapes ?? []).map((s) => deckShapeXml(s, part)).join('\n    ')
    const layoutIds = layoutGlobals.filter((l) => l.master === mi)
    return {
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>${bg}<p:spTree>
    ${spTreeOpen}
    ${shapes}
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>${layoutIds.map((l, j) => `<p:sldLayoutId id="${2147483648 + masters.length + l.number}" r:id="rId${j + 2}"/>`).join('')}</p:sldLayoutIdLst>
</p:sldMaster>`,
      rels: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${PPTX_REL.theme}" Target="../theme/theme1.xml"/>${layoutIds.length ? '\n  ' + layoutIds.map((l, j) => `<Relationship Id="rId${j + 2}" Type="${PPTX_REL.slideLayout}" Target="../slideLayouts/slideLayout${l.number}.xml"/>`).join('\n  ') : ''}${part.imageRelsXml()}
</Relationships>`,
    }
  })

  const layoutXmls = layoutGlobals.map((l) => {
    const layoutSpec = masters[l.master].layouts[l.index]
    const part = pptxPart(allocGlobalMedia, (layoutSpec.shapes ?? []).map((s) => s.id).filter((v): v is number => v !== undefined))
    const bg = pptxBgXml(layoutSpec.bg, part.allocImage)
    const phs = (layoutSpec.placeholders ?? []).map((ph) => pptxLayoutPlaceholderXml(ph, part.nextId())).join('\n    ')
    const shapes = (layoutSpec.shapes ?? []).map((s) => deckShapeXml(s, part)).join('\n    ')
    return {
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"${layoutSpec.showMasterSp === false ? ' showMasterSp="0"' : ''}>
  <p:cSld>${bg}<p:spTree>
    ${spTreeOpen}
    ${phs}
    ${shapes}
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`,
      rels: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${PPTX_REL.slideMaster}" Target="../slideMasters/slideMaster${l.master + 1}.xml"/>${part.imageRelsXml()}
</Relationships>`,
    }
  })

  const slideXmls = slides.map((sl, si) => {
    const part = pptxPart(allocGlobalMedia, sl.shapes.map((s) => s.id).filter((v): v is number => v !== undefined))
    const bg = pptxBgXml(sl.bg, part.allocImage)
    const shapes = sl.shapes.map((s) => deckShapeXml(s, part)).join('\n    ')
    const [mi, li] = sl.layout ?? [0, 0]
    const layoutNumber = masters.length ? layoutNumberOf(mi, li) : undefined
    if (layoutNumber === undefined && (masters.length || sl.layout)) {
      throw new Error(`buildPptxDeck: slide ${si} references layout [${mi}, ${li}] which does not exist`)
    }
    const before = sl.extraXml?.beforeSpTree ?? ''
    const after = sl.extraXml?.afterCSld ?? ''
    return {
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"${sl.showMasterSp === false ? ' showMasterSp="0"' : ''}>
  <p:cSld>${bg}${before}<p:spTree>
    ${spTreeOpen}
    ${shapes}
  </p:spTree></p:cSld>${after}
</p:sld>`,
      rels: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${layoutNumber !== undefined ? `<Relationship Id="rId1" Type="${PPTX_REL.slideLayout}" Target="../slideLayouts/slideLayout${layoutNumber}.xml"/>` : ''}${part.imageRelsXml()}
</Relationships>`,
    }
  })

  const bgFillStyleLst = spec.theme?.bgFillStyles?.join('')
    ?? '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr" alpha="50000"/></a:solidFill><a:solidFill><a:schemeClr val="phClr" alpha="20000"/></a:solidFill>'
  zip.file('ppt/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>
        <a:solidFill><a:schemeClr val="phClr" alpha="60000"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000" alpha="35000"/></a:outerShdw></a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>${bgFillStyleLst}</a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`)

  masterXmls.forEach((m, i) => {
    zip.file(`ppt/slideMasters/slideMaster${i + 1}.xml`, m.xml)
    zip.file(`ppt/slideMasters/_rels/slideMaster${i + 1}.xml.rels`, m.rels)
  })
  layoutXmls.forEach((l, i) => {
    zip.file(`ppt/slideLayouts/slideLayout${i + 1}.xml`, l.xml)
    zip.file(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, l.rels)
  })
  slideXmls.forEach((s, i) => {
    zip.file(`ppt/slides/slide${i + 1}.xml`, s.xml)
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, s.rels)
  })

  const deckShapes = [
    ...slides.flatMap((s) => s.shapes),
    ...masters.flatMap((m) => [...(m.shapes ?? []), ...m.layouts.flatMap((l) => l.shapes ?? [])]),
  ]
  const styleIds = [...new Set(deckShapes.map((s) => s.table?.styleId).filter((v): v is string => !!v))]
  const customStyles = new Map<string, string>()
  for (const s of deckShapes) {
    if (s.table?.styleId && s.table.customStyleXml) customStyles.set(s.table.styleId, s.table.customStyleXml)
  }
  if (styleIds.length > 0) {
    zip.file('ppt/tableStyles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">${styleIds.map((id) => customStyles.get(id) ?? tableStyleXml(id)).join('')}</a:tblStyleLst>`)
  }
  if (spec.presProps) zip.file('ppt/presProps.xml', spec.presProps)

  const presRelRows: string[] = []
  slides.forEach((_, i) => presRelRows.push(`<Relationship Id="rId${i + 1}" Type="${PPTX_REL.slide}" Target="slides/slide${i + 1}.xml"/>`))
  masters.forEach((_, m) => presRelRows.push(`<Relationship Id="rId${slides.length + m + 1}" Type="${PPTX_REL.slideMaster}" Target="slideMasters/slideMaster${m + 1}.xml"/>`))
  let presRelSeq = slides.length + masters.length
  if (spec.presProps) presRelRows.push(`<Relationship Id="rId${++presRelSeq}" Type="${PPTX_REL.presProps}" Target="presProps.xml"/>`)
  if (styleIds.length > 0) presRelRows.push(`<Relationship Id="rId${++presRelSeq}" Type="${PPTX_REL.tableStyles}" Target="tableStyles.xml"/>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${presRelRows.join('\n  ')}
</Relationships>`)
  const slideSizeXml = spec.slideSize ? `<p:sldSz cx="${spec.slideSize[0]}" cy="${spec.slideSize[1]}"/>` : '<p:sldSz cx="9144000" cy="6858000"/>'
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${masters.length ? `<p:sldMasterIdLst>${masters.map((_, m) => `<p:sldMasterId id="${2147483648 + m}" r:id="rId${slides.length + m + 1}"/>`).join('')}</p:sldMasterIdLst>` : ''}
  <p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join('')}</p:sldIdLst>
  ${slideSizeXml}
</p:presentation>`)

  const imgDefaults = [...new Set(mediaFiles.map((m) => m.ext))]
    .map((ext) => `<Default Extension="${ext}" ContentType="${PPTX_MEDIA_CT[ext] ?? `image/${ext}`}"/>`)
    .join('\n  ')
  const ctRows = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    imgDefaults,
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    ...slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`),
    ...layoutGlobals.map((_, i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`),
    ...masters.map((_, m) => `<Override PartName="/ppt/slideMasters/slideMaster${m + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`),
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    spec.presProps ? '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' : '',
    styleIds.length > 0 ? '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' : '',
  ].filter(Boolean)
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  ${ctRows.join('\n  ')}
</Types>`)
  zip.file('_rels/.rels', PPTX_ROOT_RELS)
  for (const media of mediaFiles) zip.file(`ppt/media/${media.name}`, media.data)
  return zip.generateAsync({ type: 'uint8array' })
}
