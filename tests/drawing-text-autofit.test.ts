import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'
import { layoutTextBody, type MeasureText } from '../src/drawing/text-layout'
import { graphemes } from '../src/core/text-recording'
import { paintTextBody } from '../src/drawing/text-paint'

function makeTxBody(bodyPrContent: string = '', bodyPrAttrs: string = '') {
  return parseXmlOrdered(
    `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:bodyPr ${bodyPrAttrs}>${bodyPrContent}</a:bodyPr>` +
    `<a:p><a:r><a:rPr sz="2400"/><a:t>Hello World</a:t></a:r></a:p>` +
    `</a:txBody>`
  )
}

describe('normal autofit horizontal bounds', () => {
  const measure: MeasureText = (text, style) => {
    const size = (style.fontSizePt ?? 12) * 96 / 72
    return { width: graphemes(text).length * size / 2, ascent: size * .8, descent: size * .2, normalHeight: size }
  }
  const body = (text: string, align: 'left' | 'center' | 'right', wrap = false, marginLeftPx = 0) => ({
    anchor: 't' as const, wrap,
    insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0,
    autofit: { kind: 'normal' as const },
    paragraphs: [{ align, level: 0, marginLeftEmu: marginLeftPx * 9525, runs: [{ text, fontSizePt: 24 }] }],
  })

  test.each(['center', 'right'] as const)('fits the entire %s aligned line', align => {
    const layout = layoutTextBody(body('12345678901234567890', align), 100, 500, measure)
    const segment = layout.lines[0].segments[0]
    expect(segment.width).toBeCloseTo(100)
    expect(segment.x).toBeCloseTo(0)
    expect(segment.style.fontSizePt).toBeCloseTo(7.5)
    expect(layout.overflow).toEqual({ horizontal: false, vertical: false })
  })

  test.each(['center', 'right'] as const)('reports left-side overflow at the shrink floor for %s alignment', align => {
    const layout = layoutTextBody(body('A'.repeat(100), align), 100, 500, measure)
    expect(layout.lines[0].segments[0].style.fontSizePt).toBeCloseTo(4.8)
    expect(layout.lines[0].segments[0].x).toBeLessThan(0)
    expect(layout.overflow?.horizontal).toBe(true)
  })

  test.each(['left', 'center'] as const)('fits %s aligned text beside a fixed paragraph margin', align => {
    const withMargin = body('1234567890', align, false, 50)
    const layout = layoutTextBody(withMargin, 100, 500, measure)
    const segment = layout.lines[0].segments[0]
    expect(segment.x).toBeGreaterThanOrEqual(50 - 1e-4)
    expect(segment.x + segment.width).toBeLessThanOrEqual(100 + 1e-4)
    expect(segment.style.fontSizePt).toBeCloseTo(7.5, 3)
    expect(layout.overflow).toEqual({ horizontal: false, vertical: false })
  })

  test.each(['left', 'center'] as const)('terminates at the floor when a fixed %s paragraph margin cannot fit', align => {
    const withMargin = body('1234567890', align, false, 150)
    const measuredScales = new Set<number>()
    const layout = layoutTextBody(withMargin, 100, 500, (text, style) => {
      measuredScales.add(style.fontSizePt!)
      return measure(text, style)
    })
    expect(layout.lines[0].segments[0].style.fontSizePt).toBeCloseTo(4.8)
    expect(layout.overflow?.horizontal).toBe(true)
    expect(measuredScales.size).toBeLessThanOrEqual(27)
  })

  test('terminates at the floor when fixed margin plus minimum glyph widths cannot fit', () => {
    const withMargin = body('A'.repeat(100), 'left', false, 50)
    const layout = layoutTextBody(withMargin, 100, 500, measure)
    expect(layout.lines[0].segments[0].style.fontSizePt).toBeCloseTo(4.8)
    expect(layout.overflow?.horizontal).toBe(true)
  })

  test.each(['A', '漢', '👨‍👩‍👧‍👦'])('shrinks an indivisible wrapped grapheme %s', text => {
    const layout = layoutTextBody(body(text, 'left', true), 8, 500, measure)
    const segment = layout.lines[0].segments[0]
    expect(segment.text).toBe(text)
    expect(segment.width).toBeCloseTo(8)
    expect(segment.style.fontSizePt).toBeCloseTo(12)
    expect(layout.overflow).toEqual({ horizontal: false, vertical: false })
  })

  test.each(['A', '漢', '👨‍👩‍👧‍👦'])('reports wrapped grapheme %s overflow at the shrink floor', text => {
    const layout = layoutTextBody(body(text, 'left', true), 1, 500, measure)
    expect(layout.lines[0].segments[0].width).toBeCloseTo(3.2)
    expect(layout.overflow?.horizontal).toBe(true)
  })

  test('wrapping a long word does not shrink glyphs that already fit', () => {
    const layout = layoutTextBody(body('ABCD', 'left', true), 16, 500, measure)
    expect(layout.lines).toHaveLength(4)
    expect(layout.lines[0].segments[0].style.fontSizePt).toBe(24)
    expect(layout.overflow).toEqual({ horizontal: false, vertical: false })
  })
})

describe('DrawingML Text Autofit parsing', () => {
  test('parses <a:noAutofit/> into autofit kind none', () => {
    const txBody = makeTxBody('<a:noAutofit/>')
    const body = parseTextBody(txBody)
    expect(body.autofit).toEqual({ kind: 'none' })
  })

  test('parses <a:spAutoFit/> into autofit kind shape', () => {
    const txBody = makeTxBody('<a:spAutoFit/>')
    const body = parseTextBody(txBody)
    expect(body.autofit).toEqual({ kind: 'shape' })
  })

  test('parses <a:normAutofit fontScale="80000" lnSpcReduction="20000"/> into normal autofit with scale factors', () => {
    const txBody = makeTxBody('<a:normAutofit fontScale="80000" lnSpcReduction="20000"/>')
    const body = parseTextBody(txBody)
    expect(body.autofit).toBeDefined()
    expect(body.autofit?.kind).toBe('normal')
    if (body.autofit?.kind === 'normal') {
      expect(body.autofit.fontScale).toBeCloseTo(0.8, 4)
      expect(body.autofit.lnSpcReduction).toBeCloseTo(0.2, 4)
    }
  })

  test('parses <a:normAutofit/> without attributes into normal autofit with undefined scale and reduction', () => {
    const txBody = makeTxBody('<a:normAutofit/>')
    const body = parseTextBody(txBody)
    expect(body.autofit).toEqual({ kind: 'normal' })
  })

  test('defaults to undefined autofit when no autofit element is present', () => {
    const txBody = makeTxBody('')
    const body = parseTextBody(txBody)
    expect(body.autofit).toBeUndefined()
  })

  test('parses percentage strings for fontScale and lnSpcReduction', () => {
    const txBody = makeTxBody('<a:normAutofit fontScale="75%" lnSpcReduction="15%"/>')
    const body = parseTextBody(txBody)
    expect(body.autofit?.kind).toBe('normal')
    if (body.autofit?.kind === 'normal') {
      expect(body.autofit.fontScale).toBeCloseTo(0.75, 4)
      expect(body.autofit.lnSpcReduction).toBeCloseTo(0.15, 4)
    }
  })

  test('clamps fontScale and lnSpcReduction to engine safety bounds', () => {
    // Under 1% clamped to 0.01, over 100% clamped to 1.0
    const txLow = makeTxBody('<a:normAutofit fontScale="100" lnSpcReduction="-5000"/>')
    const bodyLow = parseTextBody(txLow)
    if (bodyLow.autofit?.kind === 'normal') {
      expect(bodyLow.autofit.fontScale).toBe(0.01)
      expect(bodyLow.autofit.lnSpcReduction).toBe(0)
    }

    const txHigh = makeTxBody('<a:normAutofit fontScale="150000" lnSpcReduction="150000"/>')
    const bodyHigh = parseTextBody(txHigh)
    if (bodyHigh.autofit?.kind === 'normal') {
      expect(bodyHigh.autofit.fontScale).toBe(1.0)
      expect(bodyHigh.autofit.lnSpcReduction).toBe(1.0)
    }
  })

  test('handles invalid non-numeric strings safely', () => {
    const txInvalid = makeTxBody('<a:normAutofit fontScale="abc" lnSpcReduction="xyz"/>')
    const body = parseTextBody(txInvalid)
    expect(body.autofit).toEqual({ kind: 'normal' })
  })

  test('parses wps:bodyPr autofit in DOCX textboxes', async () => {
    const { OfficePackage } = await import('../src/core/zip')
    const { parseDocx } = await import('../src/docx/parse')
    const JSZip = (await import('jszip')).default

    const zip = new JSZip()
    zip.file('[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`)
    zip.file('_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`)

    const docxXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:r>
              <w:drawing>
                <wp:inline distT="0" distB="0" distL="0" distR="0">
                  <wp:extent cx="1905000" cy="1905000"/>
                  <wp:docPr id="1" name="Shape 1"/>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                      <wps:wsp>
                        <wps:cNvSpPr txBox="1"/>
                        <wps:spPr>
                          <a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></a:xfrm>
                          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                        </wps:spPr>
                        <wps:txbx id="1">
                          <w:txbxContent>
                            <w:p>
                              <w:r>
                                <w:rPr><w:sz w:val="24"/></w:rPr>
                                <w:t>Autofit Test</w:t>
                              </w:r>
                            </w:p>
                          </w:txbxContent>
                        </wps:txbx>
                        <wps:bodyPr>
                          <a:normAutofit fontScale="70000" lnSpcReduction="10000"/>
                        </wps:bodyPr>
                      </wps:wsp>
                    </a:graphicData>
                  </a:graphic>
                </wp:inline>
              </w:drawing>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`

    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const doc = await parseDocx(pkg)
    const tb = doc.sections[0].paragraphs[0].images[0].drawing as any
    expect(tb?.kind).toBe('textbox')
    expect(tb.autofit).toBeDefined()
    expect(tb.autofit?.kind).toBe('normal')
    expect(tb.autofit?.fontScale).toBeCloseTo(0.7, 4)
    expect(tb.autofit?.lnSpcReduction).toBeCloseTo(0.1, 4)
  })

  test('layout applies explicit fontScale reduction to run font sizes in normAutofit', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return {
        width: text.length * size * 0.6,
        ascent: size * 0.8,
        descent: size * 0.2,
        normalHeight: size,
      }
    }

    const unscaledBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [
        {
          runs: [{ text: 'Testing font scaling', fontSizePt: 20 }],
          align: 'left' as const,
          level: 0,
        },
      ],
    }

    const unscaledLayout = layoutTextBody(unscaledBody, 500, 500, mockMeasure)
    expect(unscaledLayout.lines[0].segments[0].style.fontSizePt).toBe(20)

    const scaledBody = {
      ...unscaledBody,
      autofit: { kind: 'normal' as const, fontScale: 0.75 },
    }

    const scaledLayout = layoutTextBody(scaledBody, 500, 500, mockMeasure)
    expect(scaledLayout.lines[0].segments[0].style.fontSizePt).toBe(15) // 20 * 0.75 = 15
    expect(scaledLayout.lines[0].height).toBeLessThan(unscaledLayout.lines[0].height)
  })

  test('layout applies explicit lnSpcReduction to line spacing in normAutofit', () => {
    const mockMeasure = (_text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: 50, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    const unreducedBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [
        {
          runs: [{ text: 'Line 1\nLine 2', fontSizePt: 20 }],
          lineSpacing: { kind: 'percent' as const, value: 2.0 },
          align: 'left' as const,
          level: 0,
        },
      ],
    }

    const unreducedLayout = layoutTextBody(unreducedBody, 500, 500, mockMeasure)

    const reducedBody = {
      ...unreducedBody,
      autofit: { kind: 'normal' as const, fontScale: 1.0, lnSpcReduction: 0.5 },
    }

    const reducedLayout = layoutTextBody(reducedBody, 500, 500, mockMeasure)
    expect(reducedLayout.height).toBeLessThan(unreducedLayout.height)
  })

  test('layout dynamically shrinks overflowing text to fit available height under normAutofit', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * 10, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    // 4 lines of 20pt text -> total height around ~100px+
    const overflowingBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [
        {
          runs: [{ text: 'Line 1\nLine 2\nLine 3\nLine 4', fontSizePt: 24 }],
          align: 'left' as const,
          level: 0,
        },
      ],
    }

    // Box height is restricted to 60px
    const boxHeight = 60
    const unscaledLayout = layoutTextBody(overflowingBody, 300, boxHeight, mockMeasure)
    expect(unscaledLayout.height).toBeGreaterThan(boxHeight)

    const autofitBody = {
      ...overflowingBody,
      autofit: { kind: 'normal' as const }, // dynamic autofit without authored fontScale
    }

    const fittedLayout = layoutTextBody(autofitBody, 300, boxHeight, mockMeasure)
    expect(fittedLayout.height).toBeLessThanOrEqual(boxHeight + 1e-4)
    expect(fittedLayout.lines[0].segments[0].style.fontSizePt).toBeLessThan(24)
  })

  test('layout respects vertical text flow under normAutofit', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * 15, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    const verticalBody = {
      direction: 'vert' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [
        {
          runs: [{ text: 'Vertical text', fontSizePt: 20 }],
          align: 'left' as const,
          level: 0,
        },
      ],
      autofit: { kind: 'normal' as const, fontScale: 0.5 },
    }

    const layout = layoutTextBody(verticalBody, 200, 200, mockMeasure)
    expect(layout.lines.length).toBeGreaterThan(0)
    expect(layout.lines[0].segments[0].style.fontSizePt).toBe(10)
  })

  test('zero-scale and extreme overflow boundary conditions degrade safely', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = Math.max(0, (style.fontSizePt ?? 12) * 96 / 72)
      return { width: text.length * 10, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    const boundaryBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [{ runs: [{ text: 'Boundary test', fontSizePt: 20 }], align: 'left' as const, level: 0 }],
      autofit: { kind: 'normal' as const },
    }

    // Zero height box
    expect(() => layoutTextBody(boundaryBody, 200, 0, mockMeasure)).not.toThrow()
    // Negative height box
    expect(() => layoutTextBody(boundaryBody, 200, -50, mockMeasure)).not.toThrow()
    // fontScale 0 or negative
    const zeroScaleBody = {
      ...boundaryBody,
      autofit: { kind: 'normal' as const, fontScale: 0 },
    }
    expect(() => layoutTextBody(zeroScaleBody, 200, 200, mockMeasure)).not.toThrow()
  })

  test('paintTextBody renders scaled text ink on canvas with autofit', () => {
    const renderInkCount = (autofit?: any): number => {
      const canvas = createCanvas(200, 100)
      const ctx = canvas.getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, 200, 100)

      const body = {
        direction: 'horz' as const,
        anchor: 't' as const,
        wrap: true,
        insetLeftEmu: 0,
        insetRightEmu: 0,
        insetTopEmu: 0,
        insetBottomEmu: 0,
        paragraphs: [
          {
            runs: [{ text: 'Autofit Canvas Ink', fontSizePt: 24, color: '#000000' }],
            align: 'left' as const,
            level: 0,
          },
        ],
        ...(autofit ? { autofit } : {}),
      }

      paintTextBody(body, ctx as never, 0, 0, 200, 100, f => f)
      const data = ctx.getImageData(0, 0, 200, 100).data
      let ink = 0
      for (let i = 0; i < data.length; i += 4) {
        if (data[i] < 200 || data[i + 1] < 200 || data[i + 2] < 200) ink++
      }
      return ink
    }

    const unscaledInk = renderInkCount()
    const scaledInk = renderInkCount({ kind: 'normal', fontScale: 0.5 })

    expect(unscaledInk).toBeGreaterThan(0)
    expect(scaledInk).toBeGreaterThan(0)
    // Scaled text glyphs occupy significantly fewer pixels than unscaled 24pt glyphs
    expect(scaledInk).toBeLessThan(unscaledInk * 0.7)
  })

  test('M1: dynamic floor never inflates an authored tiny scale under overflow', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * 10, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    const tinyBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [
        {
          runs: [{ text: 'Line 1\nLine 2\nLine 3\nLine 4\nLine 5', fontSizePt: 24 }],
          align: 'left' as const,
          level: 0,
        },
      ],
      autofit: { kind: 'normal' as const, fontScale: 0.05 },
    }

    const layout = layoutTextBody(tinyBody, 200, 2, mockMeasure)
    expect(layout.lines[0].segments[0].style.fontSizePt).toBeLessThanOrEqual(24 * 0.05)
  })

  test('M3: characterSpacingPt and textOutline width scale proportionally with fontScale', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * 10, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    const styledBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [
        {
          runs: [
            {
              text: 'Outline & Spacing',
              fontSizePt: 20,
              characterSpacingPt: 4,
              textOutline: { color: '#0000FF', widthPx: 2 },
            },
          ],
          align: 'left' as const,
          level: 0,
        },
      ],
      autofit: { kind: 'normal' as const, fontScale: 0.5 },
    }

    const layout = layoutTextBody(styledBody, 300, 200, mockMeasure)
    const seg = layout.lines[0].segments[0]
    expect(seg.style.fontSizePt).toBe(10)
    expect(seg.style.characterSpacingPt).toBe(2)
    expect(seg.style.textOutline?.widthPx).toBe(1)
  })

  test('S1: unwrapped text (wrap: false) dynamically shrinks on width overflow under normAutofit', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * size * 0.5, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    const unwrappedBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: false,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [
        {
          runs: [{ text: 'This is a very long unwrapped line of text that exceeds box width', fontSizePt: 20 }],
          align: 'left' as const,
          level: 0,
        },
      ],
    }

    const boxWidth = 300
    const unscaledLayout = layoutTextBody(unwrappedBody, boxWidth, 200, mockMeasure)
    const unscaledWidth = unscaledLayout.lines[0].segments[0].width
    expect(unscaledWidth).toBeGreaterThan(boxWidth)

    const autofitBody = {
      ...unwrappedBody,
      autofit: { kind: 'normal' as const },
    }

    const fittedLayout = layoutTextBody(autofitBody, boxWidth, 200, mockMeasure)
    const fittedWidth = fittedLayout.lines[0].segments[0].width
    expect(fittedWidth).toBeLessThanOrEqual(boxWidth + 1e-4)
    expect(fittedLayout.lines[0].segments[0].style.fontSizePt).toBeLessThan(20)
  })

  test('S2: vertical text direction dynamically shrinks columns on width overflow under normAutofit', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * size * 0.5, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    // Multiple vertical paragraphs that stack into columns
    const verticalMultiColBody = {
      direction: 'vert' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [
        { runs: [{ text: 'Col 1', fontSizePt: 24 }], align: 'left' as const, level: 0 },
        { runs: [{ text: 'Col 2', fontSizePt: 24 }], align: 'left' as const, level: 0 },
        { runs: [{ text: 'Col 3', fontSizePt: 24 }], align: 'left' as const, level: 0 },
        { runs: [{ text: 'Col 4', fontSizePt: 24 }], align: 'left' as const, level: 0 },
      ],
    }

    // Narrow box width (40px) where 4 columns of 24pt text will overflow transverse width
    const narrowWidth = 40
    const unscaledLayout = layoutTextBody(verticalMultiColBody, narrowWidth, 200, mockMeasure)
    expect(unscaledLayout.lines[0].segments[0].style.fontSizePt).toBe(24)

    const autofitBody = {
      ...verticalMultiColBody,
      autofit: { kind: 'normal' as const },
    }

    const fittedLayout = layoutTextBody(autofitBody, narrowWidth, 200, mockMeasure)
    expect(fittedLayout.lines[0].segments[0].style.fontSizePt).toBeLessThan(24)
  })

  test('S2 (transverse width): vertical text direction dynamically shrinks when column length overflows physical height with wrap: false', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * size * 0.5, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }
    const tallVerticalBody = {
      direction: 'vert' as const,
      anchor: 't' as const,
      wrap: false,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [{ runs: [{ text: 'A'.repeat(50), fontSizePt: 24 }], align: 'left' as const, level: 0 }],
      autofit: { kind: 'normal' as const },
    }
    // Physical box width 200, physical height 50.
    // In transverse space: width is physical height (50), height is physical width (200).
    // The 50-char line exceeds 50px transverse width when wrap: false, triggering transverse width shrink!
    const unscaledLayout = layoutTextBody({ ...tallVerticalBody, autofit: undefined }, 200, 50, mockMeasure)
    expect(unscaledLayout.lines[0].segments[0].style.fontSizePt).toBe(24)

    const fittedLayout = layoutTextBody(tallVerticalBody, 200, 50, mockMeasure)
    expect(fittedLayout.lines[0].segments[0].style.fontSizePt).toBeLessThan(24)
  })

  test('S4: layoutTextBody surfaces residual overflow as TextLayout.overflow { horizontal, vertical } (autofit branch only)', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * size * 0.5, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }

    // 1. Non-autofit bodies never carry the overflow key
    const plainBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [{ runs: [{ text: 'Hello', fontSizePt: 20 }], align: 'left' as const, level: 0 }],
    }
    const plainLayout = layoutTextBody(plainBody, 200, 200, mockMeasure)
    expect(plainLayout.overflow).toBeUndefined()

    // 2. Fitting text under normAutofit reports { horizontal: false, vertical: false }
    const fittingAutofitBody = {
      ...plainBody,
      autofit: { kind: 'normal' as const },
    }
    const fittingLayout = layoutTextBody(fittingAutofitBody, 200, 200, mockMeasure)
    expect(fittingLayout.overflow).toEqual({ horizontal: false, vertical: false })

    // 3. Text that still overflows vertically past the 0.2 floor reports { horizontal: false, vertical: true }
    // 20 lines of 20pt text in a tiny 5px box: even after shrinking to 0.2 floor (4pt), 20 lines > 5px!
    const tallLines = Array.from({ length: 20 }, (_, i) => `Line ${i + 1}`).join('\n')
    const overflowingVerticalBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [{ runs: [{ text: tallLines, fontSizePt: 20 }], align: 'left' as const, level: 0 }],
      autofit: { kind: 'normal' as const },
    }
    const overflowVerticalLayout = layoutTextBody(overflowingVerticalBody, 200, 5, mockMeasure)
    expect(overflowVerticalLayout.overflow).toEqual({ horizontal: false, vertical: true })

    // 4. Text that still overflows horizontally past the 0.2 floor reports { horizontal: true, vertical: false }
    // An enormous 1000-char unwrapped line in a 10px box: at 0.2 floor (4pt), width >> 10px!
    const wideText = 'A'.repeat(1000)
    const overflowingHorizontalBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: false,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [{ runs: [{ text: wideText, fontSizePt: 20 }], align: 'left' as const, level: 0 }],
      autofit: { kind: 'normal' as const },
    }
    const overflowHorizontalLayout = layoutTextBody(overflowingHorizontalBody, 10, 200, mockMeasure)
    expect(overflowHorizontalLayout.overflow).toEqual({ horizontal: true, vertical: false })

    // 5. Vertical text direction swaps transverse axes back to physical coordinates
    // In vertical text flow with wrap: false:
    // A long line that exceeds the physical height (transverse width):
    // transverse horizontal overflow -> physical vertical overflow!
    const tallVerticalBody = {
      direction: 'vert' as const,
      anchor: 't' as const,
      wrap: false,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [{ runs: [{ text: 'A'.repeat(500), fontSizePt: 20 }], align: 'left' as const, level: 0 }],
      autofit: { kind: 'normal' as const },
    }
    const vertLayout1 = layoutTextBody(tallVerticalBody, 200, 10, mockMeasure)
    expect(vertLayout1.overflow).toEqual({ horizontal: false, vertical: true })

    // And multiple lines that exceed physical width (transverse height):
    // transverse vertical overflow -> physical horizontal overflow!
    const wideVerticalBody = {
      direction: 'vert' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [{ runs: [{ text: Array.from({ length: 30 }, (_, i) => `Col${i}`).join('\n'), fontSizePt: 20 }], align: 'left' as const, level: 0 }],
      autofit: { kind: 'normal' as const },
    }
    const vertLayout2 = layoutTextBody(wideVerticalBody, 10, 200, mockMeasure)
    expect(vertLayout2.overflow).toEqual({ horizontal: true, vertical: false })
  })

  test('M6: layoutTextBody never mutates caller text body and repeated layout of same body is stable', () => {
    const mockMeasure = (text: string, style: any) => {
      const size = (style.fontSizePt ?? 12) * 96 / 72
      return { width: text.length * size * 0.5, ascent: size * 0.8, descent: size * 0.2, normalHeight: size }
    }
    const tallLines = Array.from({ length: 20 }, (_, i) => `Line ${i + 1}`).join('\n')
    const originalBody = {
      direction: 'horz' as const,
      anchor: 't' as const,
      wrap: true,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      paragraphs: [{ runs: [{ text: tallLines, fontSizePt: 24 }], align: 'left' as const, level: 0 }],
      autofit: { kind: 'normal' as const },
    }

    const firstLayout = layoutTextBody(originalBody, 200, 100, mockMeasure)
    // 1. Caller body object must NOT have its autofit property stripped/mutated
    expect(originalBody.autofit).toEqual({ kind: 'normal' })

    // 2. Repeated layout of the same body instance must produce identical results
    const secondLayout = layoutTextBody(originalBody, 200, 100, mockMeasure)
    expect(secondLayout.height).toBe(firstLayout.height)
    expect(secondLayout.lines[0].segments[0].style.fontSizePt).toBe(firstLayout.lines[0].segments[0].style.fontSizePt)
  })
})
