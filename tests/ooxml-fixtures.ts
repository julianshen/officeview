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
