import { describe, expect, test } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'

// Mirrors the retained a11 VVVVXXXX run (54pt bold, 1pt outline, dkUpDiag
// pattern, hard offset shadow) with srgbClr colors so no theme is needed.
const WORDART_RPR =
  `<a:rPr sz="5400" b="1"><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>` +
  `<a:pattFill prst="dkUpDiag"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="00FF00"/></a:bgClr></a:pattFill>` +
  `<a:effectLst><a:outerShdw dist="38100" dir="2640000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst></a:rPr>`
const txBody = (rpr: string, text: string) =>
  parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:p><a:r>${rpr}<a:t>${text}</a:t></a:r></a:p></a:txBody>`)

describe('wordart run appearance parse', () => {
  test('pattern, outline and shadow land on the model with resolved colors', () => {
    const body = parseTextBody(txBody(WORDART_RPR, 'W'))
    const run = body.paragraphs[0].runs[0]
    expect(run.fontSizePt).toBe(54)
    expect(run.bold).toBe(true)
    expect(run.textFill).toMatchObject({ kind: 'pattern', preset: 'dkUpDiag', fg: '#FF0000', bg: '#00FF00' })
    expect(run.textOutline).toMatchObject({ color: '#0000FF' })
    expect(run.textOutline?.widthPx).toBeCloseTo((12700 * 96) / (12700 * 72), 6)
    expect(run.textShadow).toMatchObject({ color: '#000000', blurPx: 0 })
    // dist 38100 EMU = 4px at 44 degrees (dir 2640000/60000): exact offsets.
    expect(run.textShadow?.offsetX).toBeCloseTo(Math.cos((2640000 * Math.PI) / 10800000) * 4, 6)
    expect(run.textShadow?.offsetY).toBeCloseTo(Math.sin((2640000 * Math.PI) / 10800000) * 4, 6)
  })
  test('blur radius converts from EMU', () => {
    const rpr = `<a:rPr><a:effectLst><a:outerShdw dist="0" dir="0" bluRad="19050"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst></a:rPr>`
    const run = parseTextBody(txBody(rpr, 'B')).paragraphs[0].runs[0]
    expect(run.textShadow?.blurPx).toBe(2)
  })
  test('linear gradient stops resolve with angle', () => {
    const rpr = `<a:rPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></a:rPr>`
    const run = parseTextBody(txBody(rpr, 'G')).paragraphs[0].runs[0]
    expect(run.textFill?.kind).toBe('gradient')
    expect(run.textFill).toMatchObject({ stops: [{ position: 0, color: '#FF0000' }, { position: 1, color: '#0000FF' }] })
    if (run.textFill?.kind === 'gradient') expect(run.textFill.angle).toBeCloseTo(Math.PI / 2, 9)
  })
  test('unsupported pattern preset and picture fill diagnose and fall back', () => {
    const rpr = `<a:rPr><a:pattFill prst="noSuchPattern"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr></a:pattFill></a:rPr>`
    const body = parseTextBody(txBody(rpr, 'P'))
    expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance')).toBe(true)
    expect(body.paragraphs[0].runs[0].textFill).toBeUndefined()
    const blip = `<a:rPr><a:blipFill><a:blip r:embed="rId1"/></a:blipFill></a:rPr>`
    const blipBody = parseTextBody(txBody(blip, 'B'))
    expect(blipBody.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'blipFill')).toBe(true)
    expect(blipBody.paragraphs[0].runs[0].textFill).toBeUndefined()
  })
  test('unsupported path gradient diagnoses and falls back to first stop color', () => {
    const rpr = `<a:rPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:path path="circle"/></a:gradFill></a:rPr>`
    const body = parseTextBody(txBody(rpr, 'P'))
    expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'gradFill:path')).toBe(true)
    expect(body.paragraphs[0].runs[0].textFill).toBeUndefined()
    expect(body.paragraphs[0].runs[0].color).toBe('#FF0000')
  })
  test('outline-only run keeps noFill with a stroked outline', () => {
    const rpr = `<a:rPr><a:noFill/><a:ln w="25400"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:rPr>`
    const run = parseTextBody(txBody(rpr, 'O')).paragraphs[0].runs[0]
    expect(run.noFill).toBe(true)
    expect(run.textOutline).toMatchObject({ color: '#00FF00' })
  })
  test('run gradFill clears style-layer noFill (most specific wins)', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:noFill/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></a:rPr><a:t>G</a:t></a:r></a:p></a:txBody>`)
    const run = parseTextBody(xml).paragraphs[0].runs[0]
    expect(run.noFill).toBe(false)
    expect(run.textFill?.kind).toBe('gradient')
  })
  test('appearance colors resolve through theme scheme colors', async () => {
    const { parseThemeContext } = await import('../src/drawing/style')
    const theme = parseThemeContext(
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>` +
      `<a:clrScheme><a:dk2><a:srgbClr val="FF0000"/></a:dk2></a:clrScheme></a:themeElements></a:theme>`)
    const rpr =
      `<a:rPr><a:pattFill prst="dkUpDiag"><a:fgClr><a:schemeClr val="tx2"/></a:fgClr>` +
      `<a:bgClr><a:schemeClr val="tx2"><a:lumMod val="50000"/></a:schemeClr></a:bgClr></a:pattFill>` +
      `<a:ln w="12700"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:ln></a:rPr>`
    const body = parseTextBody(txBody(rpr, 'T'), theme)
    const run = body.paragraphs[0].runs[0]
    expect(run.textFill).toMatchObject({ kind: 'pattern', fg: '#FF0000' })
    expect(run.textFill && run.textFill.kind === 'pattern' && run.textFill.bg).not.toBe('#FF0000')
    expect(run.textOutline).toMatchObject({ color: '#FF0000' })
  })
  test('run solidFill clears style-layer noFill (most specific wins)', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:noFill/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>G</a:t></a:r></a:p></a:txBody>`)
    const run = parseTextBody(xml).paragraphs[0].runs[0]
    expect(run.noFill).toBe(false)
    expect(run.color).toBe('#00FF00')
    expect(run.textFill).toBeUndefined()
  })
  test('run ln/noFill clears an inherited outline silently', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:ln><a:noFill/></a:ln></a:rPr><a:t>N</a:t></a:r></a:p></a:txBody>`)
    const body = parseTextBody(xml)
    expect(body.paragraphs[0].runs[0].textOutline).toBeUndefined()
    expect(body.diagnostics ?? []).toEqual([])
  })
  test('width-less ln diagnoses and clears the inherited outline', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:ln><a:prstDash val="solid"/></a:ln></a:rPr><a:t>N</a:t></a:r></a:p></a:txBody>`)
    const body = parseTextBody(xml)
    expect(body.paragraphs[0].runs[0].textOutline).toBeUndefined()
    expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'ln')).toBe(true)
  })
  test('malformed ln width diagnoses and clears instead of coercing to zero', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:ln w="NaN"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:rPr><a:t>N</a:t></a:r></a:p></a:txBody>`)
    const body = parseTextBody(xml)
    expect(body.paragraphs[0].runs[0].textOutline).toBeUndefined()
    expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'ln')).toBe(true)
  })
})
