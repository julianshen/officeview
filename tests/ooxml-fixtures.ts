/** Builders for minimal but structurally-valid OOXML test files. */
import JSZip from 'jszip'

export const CT_TYPES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
}

export interface DocxParaSpec {
  align?: 'left' | 'center' | 'right' | 'both'
  runs: DocxTextSpec[]
}

/** Build a minimal docx buffer with the given paragraphs. */
export async function buildDocx(paras: DocxParaSpec[]): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  const runXml = (r: DocxTextSpec) => {
    const rpr = `<w:rPr>${r.bold ? '<w:b/>' : ''}${r.italic ? '<w:i/>' : ''}${r.size ? `<w:sz w:val="${r.size}"/>` : ''}${r.color ? `<w:color w:val="${r.color}"/>` : ''}</w:rPr>`
    const text = r.text.replace(/&/g, '&amp;').replace(/</g, '&lt;')
    return `<w:r>${rpr}<w:t xml:space="preserve">${text}</w:t></w:r>`
  }
  const paraXml = (p: DocxParaSpec) => {
    const ppr = p.align ? `<w:pPr><w:jc w:val="${p.align}"/></w:pPr>` : ''
    return `<w:p>${ppr}${p.runs.map(runXml).join('')}</w:p>`
  }
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${paras.map(paraXml).join('\n    ')}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`
  zip.file('word/document.xml', document)
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
  style?: number
}

export interface XlsxSheetSpec {
  name: string
  rows: Array<{ r: number; cells: XlsxCellSpec[] }>
  cols?: string
  merges?: string[]
}

/** Build a minimal xlsx buffer. */
export async function buildXlsx(sheets: XlsxSheetSpec[], sharedStrings: string[] = []): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', XLSX_CT)
  zip.file('_rels/.rels', XLSX_ROOT_RELS)
  zip.file('xl/_rels/workbook.xml.rels', XLSX_WORKBOOK_RELS)
  const sheetXml = (sheet: XlsxSheetSpec) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheet.cols ? `<cols>${sheet.cols}</cols>` : ''}
  <sheetData>
    ${sheet.rows.map((row) => `<row r="${row.r}">${row.cells.map((c) => `<c r="${c.ref}"${c.t ? ` t="${c.t}"` : ''}${c.style !== undefined ? ` s="${c.style}"` : ''}>${c.v !== undefined ? `<v>${c.v}</v>` : ''}</c>`).join('')}</row>`).join('\n    ')}
  </sheetData>
  ${sheet.merges ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}
</worksheet>`
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
</Types>`

const PPTX_ROOT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`

const PPTX_PRES_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`

export interface PptxShapeSpec {
  prst?: string
  off?: [string, string]
  ext?: [string, string]
  fill?: string
  lineW?: string
  rot?: string
  paragraphs?: Array<{ align?: string; runs: Array<{ text: string; b?: boolean; i?: boolean; sz?: string; color?: string }> }>
}

/** Build a minimal pptx with one or more slides (single slide file reused). */
export async function buildPptx(shapes: PptxShapeSpec[]): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', PPTX_CT)
  zip.file('_rels/.rels', PPTX_ROOT_RELS)
  zip.file('ppt/_rels/presentation.xml.rels', PPTX_PRES_RELS)
  zip.file('ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>`)
  const shapeXml = (s: PptxShapeSpec) => {
    const spPr =
      `<p:spPr><a:xfrm${s.rot ? ` rot="${s.rot}"` : ''}><a:off x="${s.off?.[0] ?? '0'}" y="${s.off?.[1] ?? '0'}"/><a:ext cx="${s.ext?.[0] ?? '100000'}" cy="${s.ext?.[1] ?? '100000'}"/></a:xfrm>` +
      `<a:prstGeom prst="${s.prst ?? 'rect'}"><a:avLst/></a:prstGeom>` +
      (s.fill ? `<a:solidFill><a:srgbClr val="${s.fill}"/></a:solidFill>` : '') +
      `<a:ln w="${s.lineW ?? '12700'}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>`
    const txBody = s.paragraphs
      ? `<p:txBody><a:bodyPr/><a:lstStyle/>${s.paragraphs.map((p) =>
          `<a:p>${p.align ? `<a:pPr algn="${p.align}"/>` : ''}${p.runs.map((r) => {
            const attrs = `${r.b ? ' b="1"' : ''}${r.i ? ' i="1"' : ''}${r.sz ? ` sz="${r.sz}"` : ''}`
            const inner = r.color ? `<a:solidFill><a:srgbClr val="${r.color}"/></a:solidFill>` : ''
            return `<a:r><a:rPr${attrs}>${inner}</a:rPr><a:t>${r.text}</a:t></a:r>`
          }).join('')}</a:p>`).join('')}</p:txBody>`
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
  return zip.generateAsync({ type: 'uint8array' })
}
