import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { parseXmlOrdered } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parseTextBody } from '../src/drawing/text-parse'
import { prepareDrawingContent } from '../src/drawing/content'

describe('Phase 14: Nondefault textPlain Adjustments & Diagnostic Cleanup', () => {
  describe('textPlain nondefault adjustment diagnostics', () => {
    test('textPlain with default or omitted adjustment emits zero warp diagnostics', () => {
      const xmlDefault = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr>` +
        `    <a:prstTxWarp prst="textPlain"/>` +
        `  </a:bodyPr>` +
        `  <a:p><a:r><a:t>Hello</a:t></a:r></a:p>` +
        `</p:txBody>`
      )
      const bodyDefault = parseTextBody(xmlDefault)
      expect(bodyDefault.diagnostics?.some(d => d.kind === 'unsupported-text-warp')).toBeFalsy()

      const xmlAdj50000 = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr>` +
        `    <a:prstTxWarp prst="textPlain">` +
        `      <a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst>` +
        `    </a:prstTxWarp>` +
        `  </a:bodyPr>` +
        `  <a:p><a:r><a:t>Hello</a:t></a:r></a:p>` +
        `</p:txBody>`
      )
      const body50000 = parseTextBody(xmlAdj50000)
      expect(body50000.diagnostics?.some(d => d.kind === 'unsupported-text-warp')).toBeFalsy()
    })

    test('textPlain with nondefault adjustment emits explicit unsupported-text-warp diagnostic', () => {
      const xmlAdj30000 = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr>` +
        `    <a:prstTxWarp prst="textPlain">` +
        `      <a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst>` +
        `    </a:prstTxWarp>` +
        `  </a:bodyPr>` +
        `  <a:p><a:r><a:t>Hello</a:t></a:r></a:p>` +
        `</p:txBody>`
      )
      const body30000 = parseTextBody(xmlAdj30000)
      expect(body30000.diagnostics?.some(d => d.kind === 'unsupported-text-warp' && d.feature === 'textPlain')).toBe(true)
      expect(body30000.diagnostics?.find(d => d.feature === 'textPlain')?.message).toContain('30000')

      const xmlAdj70000 = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr>` +
        `    <a:prstTxWarp prst="textPlain">` +
        `      <a:avLst><a:gd name="adj" fmla="val 70000"/></a:avLst>` +
        `    </a:prstTxWarp>` +
        `  </a:bodyPr>` +
        `  <a:p><a:r><a:t>Hello</a:t></a:r></a:p>` +
        `</p:txBody>`
      )
      const body70000 = parseTextBody(xmlAdj70000)
      expect(body70000.diagnostics?.some(d => d.kind === 'unsupported-text-warp' && d.feature === 'textPlain')).toBe(true)
    })
  })

  describe('overridden default fill diagnostic pruning', () => {
    test('overridden default fill does not report fallback degradation warning when run fill is supported', () => {
      const xmlOverridden = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr/>` +
        `  <a:p>` +
        `    <a:pPr>` +
        `      <a:defRPr>` +
        `        <a:pattFill prst="unsupportedPreset"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr></a:pattFill>` +
        `      </a:defRPr>` +
        `    </a:pPr>` +
        `    <a:r>` +
        `      <a:rPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr>` +
        `      <a:t>Overridden</a:t>` +
        `    </a:r>` +
        `  </a:p>` +
        `</p:txBody>`
      )
      const body = parseTextBody(xmlOverridden)
      expect(body.paragraphs[0].runs[0].color).toBe('#00FF00')
      // Since the run overrode the fill with supported solidFill, the defRPr pattern warning must not be reported
      expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance')).toBeFalsy()
    })

    test('unoverridden default fill reports fallback degradation warning when run inherits it', () => {
      const xmlInherited = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr/>` +
        `  <a:p>` +
        `    <a:pPr>` +
        `      <a:defRPr>` +
        `        <a:pattFill prst="unsupportedPreset"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr></a:pattFill>` +
        `      </a:defRPr>` +
        `    </a:pPr>` +
        `    <a:r>` +
        `      <a:t>Inherited</a:t>` +
        `    </a:r>` +
        `  </a:p>` +
        `</p:txBody>`
      )
      const body = parseTextBody(xmlInherited)
      // Run inherited the unsupported pattern fill, so warning MUST be reported
      expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance')).toBe(true)
    })

    test('mixed paragraph where one run overrides fill and one run inherits reports warning for inherited run', () => {
      const xmlMixed = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr/>` +
        `  <a:p>` +
        `    <a:pPr>` +
        `      <a:defRPr>` +
        `        <a:pattFill prst="unsupportedPreset"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr></a:pattFill>` +
        `      </a:defRPr>` +
        `    </a:pPr>` +
        `    <a:r>` +
        `      <a:rPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr>` +
        `      <a:t>Overridden</a:t>` +
        `    </a:r>` +
        `    <a:r>` +
        `      <a:t>Inherited</a:t>` +
        `    </a:r>` +
        `  </a:p>` +
        `</p:txBody>`
      )
      const body = parseTextBody(xmlMixed)
      expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance')).toBe(true)
    })

    test('invalid, empty, or whitespace fmla attributes in textPlain avLst do not emit false-positive warp warnings', () => {
      const xmlBadFmla = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr>` +
        `    <a:prstTxWarp prst="textPlain">` +
        `      <a:avLst><a:gd name="adj" fmla="notANumber"/></a:avLst>` +
        `    </a:prstTxWarp>` +
        `  </a:bodyPr>` +
        `  <a:p><a:r><a:t>Hello</a:t></a:r></a:p>` +
        `</p:txBody>`
      )
      const bodyBad = parseTextBody(xmlBadFmla)
      expect(bodyBad.diagnostics?.some(d => d.kind === 'unsupported-text-warp')).toBeFalsy()

      const xmlWhitespaceFmla = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr>` +
        `    <a:prstTxWarp prst="textPlain">` +
        `      <a:avLst><a:gd name="adj" fmla="val   "/></a:avLst>` +
        `    </a:prstTxWarp>` +
        `  </a:bodyPr>` +
        `  <a:p><a:r><a:t>Hello</a:t></a:r></a:p>` +
        `</p:txBody>`
      )
      const bodyWhitespace = parseTextBody(xmlWhitespaceFmla)
      expect(bodyWhitespace.diagnostics?.some(d => d.kind === 'unsupported-text-warp')).toBeFalsy()
    })

    test('adjustment formula without val prefix is parsed correctly', () => {
      const xmlPrefixFree = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr>` +
        `    <a:prstTxWarp prst="textPlain">` +
        `      <a:avLst><a:gd name="adj" fmla="30000"/></a:avLst>` +
        `    </a:prstTxWarp>` +
        `  </a:bodyPr>` +
        `  <a:p><a:r><a:t>Hello</a:t></a:r></a:p>` +
        `</p:txBody>`
      )
      const body = parseTextBody(xmlPrefixFree)
      expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-warp' && d.feature === 'textPlain')).toBe(true)
      expect(body.diagnostics?.find(d => d.feature === 'textPlain')?.message).toContain('30000')
    })

    test('multi-line paragraph with <a:br/> between overridden runs retains diagnostic pruning', () => {
      const xmlBr = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr/>` +
        `  <a:p>` +
        `    <a:pPr>` +
        `      <a:defRPr>` +
        `        <a:pattFill prst="unsupportedPreset"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr></a:pattFill>` +
        `      </a:defRPr>` +
        `    </a:pPr>` +
        `    <a:r>` +
        `      <a:rPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr>` +
        `      <a:t>Line 1</a:t>` +
        `    </a:r>` +
        `    <a:br/>` +
        `    <a:r>` +
        `      <a:rPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr>` +
        `      <a:t>Line 2</a:t>` +
        `    </a:r>` +
        `  </a:p>` +
        `</p:txBody>`
      )
      const body = parseTextBody(xmlBr)
      expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance')).toBeFalsy()
    })

    test('overridden default outline does not report warning when all runs specify direct outline', () => {
      const xmlOutline = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr/>` +
        `  <a:p>` +
        `    <a:pPr>` +
        `      <a:defRPr>` +
        `        <a:ln w="0"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>` +
        `      </a:defRPr>` +
        `    </a:pPr>` +
        `    <a:r>` +
        `      <a:rPr><a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:rPr>` +
        `      <a:t>Outlined</a:t>` +
        `    </a:r>` +
        `  </a:p>` +
        `</p:txBody>`
      )
      const body = parseTextBody(xmlOutline)
      expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'ln')).toBeFalsy()
    })

    test('multi-layer inheritance where paragraph defRPr overrides master unsupported fill prunes master warning', () => {
      const slideXml = parseXmlOrdered(
        `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
        `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr/>` +
        `  <a:p>` +
        `    <a:pPr>` +
        `      <a:defRPr>` +
        `        <a:solidFill><a:srgbClr val="00AA00"/></a:solidFill>` +
        `      </a:defRPr>` +
        `    </a:pPr>` +
        `    <a:r><a:t>Inherits paragraph solid fill</a:t></a:r>` +
        `  </a:p>` +
        `</p:txBody>`
      )
      const masterBody = parseXmlOrdered(
        `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:bodyPr/>` +
        `  <a:lstStyle><a:lvl1pPr><a:defRPr>` +
        `    <a:pattFill prst="unsupportedPreset"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr></a:pattFill>` +
        `  </a:defRPr></a:lvl1pPr></a:lstStyle>` +
        `  <a:p><a:r><a:t/></a:r></a:p>` +
        `</a:txBody>`
      )
      const body = parseTextBody(slideXml, undefined, undefined, {}, [{ origin: 'master', body: masterBody }])
      expect(body.paragraphs[0].runs[0].color).toBe('#00AA00')
      expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance')).toBeFalsy()
    })

    test('Word textbox via prepareDrawingContent handles textPlain nondefault adjustment in content.ts', async () => {
      const zip = new JSZip()
      zip.file('word/document.xml', '<w:document/>')
      const bytes = await zip.generateAsync({ type: 'uint8array' })
      const pkg = await OfficePackage.load(bytes)
      const theme = { colors: new Map(), fonts: new Map() }
      const drawingTheme: any = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }

      const graphicNonDefault = parseXmlOrdered(
        `<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ` +
        `  uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
        `  <wps:wsp>` +
        `    <wps:bodyPr>` +
        `      <a:prstTxWarp prst="textPlain">` +
        `        <a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>` +
        `      </a:prstTxWarp>` +
        `    </wps:bodyPr>` +
        `    <wps:txbx><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
        `  </wps:wsp>` +
        `</a:graphicData>`
      )
      const content = await prepareDrawingContent(pkg, graphicNonDefault, 'word/document.xml', theme, drawingTheme, {
        parseParagraph: () => ({ runs: [{ text: 'Hello' }], align: 'left' } as any),
      })
      expect(content).toBeDefined()
      if (content?.kind === 'textbox') {
        expect(content.diagnostics?.some(d => d.kind === 'unsupported-text-warp' && d.feature === 'textPlain')).toBe(true)
        expect(content.diagnostics?.find(d => d.feature === 'textPlain')?.message).toContain('25000')
      }

      const graphicDefault = parseXmlOrdered(
        `<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ` +
        `  uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
        `  <wps:wsp>` +
        `    <wps:bodyPr>` +
        `      <a:prstTxWarp prst="textPlain"/>` +
        `    </wps:bodyPr>` +
        `    <wps:txbx><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
        `  </wps:wsp>` +
        `</a:graphicData>`
      )
      const contentDef = await prepareDrawingContent(pkg, graphicDefault, 'word/document.xml', theme, drawingTheme, {
        parseParagraph: () => ({ runs: [{ text: 'Hello' }], align: 'left' } as any),
      })
      expect(contentDef).toBeDefined()
      if (contentDef?.kind === 'textbox') {
        expect(contentDef.diagnostics?.some(d => d.kind === 'unsupported-text-warp')).toBeFalsy()
      }
    })
  })
})
