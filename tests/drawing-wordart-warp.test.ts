import { describe, expect, it } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'

const txBodyXml = (bodyPrInner: string, text = 'WordArt Warp') => `
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr>
    ${bodyPrInner}
  </a:bodyPr>
  <a:p>
    <a:r>
      <a:t>${text}</a:t>
    </a:r>
  </a:p>
</a:txBody>`

describe('WordArt Preset Text Warp Parsing & Modeling (<a:prstTxWarp>) - Phase 4', () => {
  it('parses preset text warp type (prst) from <a:bodyPr><a:prstTxWarp> into DrawingTextBody', () => {
    const xml = txBodyXml('<a:prstTxWarp prst="textArchUp"/>')
    const parsed = parseTextBody(parseXmlOrdered(xml))

    expect(parsed.textWarp).toBeDefined()
    expect(parsed.textWarp?.preset).toBe('textArchUp')
  })

  it('parses adjustment values (<a:avLst><a:gd>) for preset text warps with appropriate units (angles vs percentage)', () => {
    // Angle adjustment: 5400000 (90 degrees in 60000ths of a degree)
    const archXml = txBodyXml(`
      <a:prstTxWarp prst="textArchUp">
        <a:avLst>
          <a:gd name="adj" fmla="val 5400000"/>
        </a:avLst>
      </a:prstTxWarp>
    `)
    const parsedArch = parseTextBody(parseXmlOrdered(archXml))
    expect(parsedArch.textWarp?.preset).toBe('textArchUp')
    expect(parsedArch.textWarp?.adjustments?.adj).toBe(5400000)

    // Percentage adjustment: 25000 (25% in 100000ths)
    const inflateXml = txBodyXml(`
      <a:prstTxWarp prst="textInflate">
        <a:avLst>
          <a:gd name="adj" fmla="val 25000"/>
        </a:avLst>
      </a:prstTxWarp>
    `)
    const parsedInflate = parseTextBody(parseXmlOrdered(inflateXml))
    expect(parsedInflate.textWarp?.preset).toBe('textInflate')
    expect(parsedInflate.textWarp?.adjustments?.adj).toBe(25000)

    // Multi-guide adjustment (e.g. wave presets with adj1 and adj2)
    const waveXml = txBodyXml(`
      <a:prstTxWarp prst="textWave1">
        <a:avLst>
          <a:gd name="adj1" fmla="val 10000"/>
          <a:gd name="adj2" fmla="val 35000"/>
        </a:avLst>
      </a:prstTxWarp>
    `)
    const parsedWave = parseTextBody(parseXmlOrdered(waveXml))
    expect(parsedWave.textWarp?.preset).toBe('textWave1')
    expect(parsedWave.textWarp?.adjustments?.adj1).toBe(10000)
    expect(parsedWave.textWarp?.adjustments?.adj2).toBe(35000)
  })

  it('applies default adjustment table for presets when <a:avLst> is omitted', () => {
    // textArchUp default adj is 10800000 (180 degrees)
    const archXml = txBodyXml('<a:prstTxWarp prst="textArchUp"/>')
    const parsedArch = parseTextBody(parseXmlOrdered(archXml))
    expect(parsedArch.textWarp?.adjustments?.adj).toBe(10800000)

    // textCircle default adj is 10800000 (180 degrees)
    const circleXml = txBodyXml('<a:prstTxWarp prst="textCircle"/>')
    const parsedCircle = parseTextBody(parseXmlOrdered(circleXml))
    expect(parsedCircle.textWarp?.adjustments?.adj).toBe(10800000)

    // textWave1 defaults: adj1 = 0, adj2 = 50000 (50%)
    const waveXml = txBodyXml('<a:prstTxWarp prst="textWave1"/>')
    const parsedWave = parseTextBody(parseXmlOrdered(waveXml))
    expect(parsedWave.textWarp?.adjustments?.adj1).toBe(0)
    expect(parsedWave.textWarp?.adjustments?.adj2).toBe(50000)

    // textInflate default adj is 50000 (50%)
    const inflateXml = txBodyXml('<a:prstTxWarp prst="textInflate"/>')
    const parsedInflate = parseTextBody(parseXmlOrdered(inflateXml))
    expect(parsedInflate.textWarp?.adjustments?.adj).toBe(50000)

    // textSlantUp default adj is 25000
    const slantXml = txBodyXml('<a:prstTxWarp prst="textSlantUp"/>')
    const parsedSlant = parseTextBody(parseXmlOrdered(slantXml))
    expect(parsedSlant.textWarp?.adjustments?.adj).toBe(25000)
  })

  it('handles textNoShape and textPlain as unwarped standard text', () => {
    const noShapeXml = txBodyXml('<a:prstTxWarp prst="textNoShape"/>', 'Standard Plain Text')
    const parsedNoShape = parseTextBody(parseXmlOrdered(noShapeXml))
    expect(parsedNoShape.textWarp).toBeUndefined()
    expect(parsedNoShape.diagnostics).toBeUndefined()

    const plainXml = txBodyXml('<a:prstTxWarp prst="textPlain"/>', 'Standard Plain Text')
    const parsedPlain = parseTextBody(parseXmlOrdered(plainXml))
    expect(parsedPlain.textWarp).toBeUndefined()
    expect(parsedPlain.diagnostics).toBeUndefined()
  })

  it('emits diagnostic for unknown warp presets and falls back to unwarped text rendering without throwing', () => {
    const unknownXml = txBodyXml('<a:prstTxWarp prst="textSuperUnknownPreset"/>', 'Fallback Text')
    const parsed = parseTextBody(parseXmlOrdered(unknownXml))

    // Must not crash and fall back to unwarped text
    expect(parsed.textWarp).toBeUndefined()
    expect(parsed.paragraphs[0].runs[0].text).toBe('Fallback Text')

    // Must emit diagnostic with kind 'unsupported-text-warp'
    expect(parsed.diagnostics).toBeDefined()
    expect(parsed.diagnostics).toContainEqual({
      kind: 'unsupported-text-warp',
      feature: 'textSuperUnknownPreset',
      message: expect.stringContaining('textSuperUnknownPreset'),
    })
  })
})
