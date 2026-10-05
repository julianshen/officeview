import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { getChildren, namespaceUri, parseXml, parseXmlOrdered } from '../src/core/xml'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide } from '../src/pptx/render'
import { parseDocx } from '../src/docx/parse'
import { paintDrawing } from '../src/docx/drawing'
import { parseXlsx } from '../src/xlsx/parse'
import { drawingReport, pptxCoverage } from '../src/drawing/coverage'
import { contentDiagnostic, drawingPartContext, reserveDrawingNode } from '../src/drawing/parts'

const rels = (body: string) => `<Relationships>${body}</Relationships>`
const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`
async function packageOf(parts: Record<string, string | Uint8Array>) {
  const zip = new JSZip()
  for (const [path, value] of Object.entries(parts)) zip.file(path, value)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}
const xfrm = '<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>'
const shape = (id: number, geometry = '<a:prstGeom prst="rect"/>', text = '') => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="shape${id}"/></p:nvSpPr><p:spPr>${xfrm}${geometry}</p:spPr>${text ? `<p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>` : ''}</p:sp>`
const frame = (id: number, uri: string, payload = '') => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="frame${id}"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></p:xfrm><a:graphic><a:graphicData uri="${uri}">${payload}</a:graphicData></a:graphic></p:graphicFrame>`
const picture = (id: number, rid: string) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="pic${id}"/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/></p:blipFill><p:spPr>${xfrm}</p:spPr></p:pic>`
const alternative = (choice: string, fallback: string, requires: string) => `<mc:AlternateContent><mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
const inkXml = (units = 'px') => `<ink><traceFormat><channel name="X" units="${units}"/><channel name="Y" units="px"/></traceFormat><trace>0 0, 100 100</trace></ink>`
const horizontalChart = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
async function pptx(body: string, parts: Record<string, string | Uint8Array> = {}) {
  return parsePptx(await packageOf({
    'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': rels(rel('s1', 'slide', 'slides/slide1.xml')),
    'ppt/slides/slide1.xml': `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`,
    ...parts,
  }))
}

describe('machine drawing coverage', () => {
  test.each([
    ['unsupported', inkXml('unsupported'), 'unsupported', 'none'],
    ['missing', undefined, 'malformed', 'none'],
    ['malformed', '<ink><trace></ink>', 'malformed', 'none'],
    ['valid', inkXml(), 'native', 'inkml'],
  ] as const)('XLSX direct contentPart %s reports actual consumed payload', async (_mode, payload, status, selectedRepresentation) => {
    const parts: Record<string, string> = {
      'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><xdr:contentPart r:id="ink"/><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>',
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('ink', 'customXml', '../ink/ink1.xml')),
    }
    if (payload) parts['xl/ink/ink1.xml'] = payload
    const doc = await parseXlsx(await packageOf(parts))
    expect(doc.drawingCoverage![0]).toMatchObject({ element: 'contentPart', referenceId: 'ink', feature: expect.stringMatching(/ink|contentPart/i), status, selectedRepresentation })
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.sheets[0].drawings?.[0].content?.kind).toBe(payload === inkXml() ? 'ink' : undefined)
  })

  test('Word repeated unsupported InkML keeps owners and no orphan diagnostics', async () => {
    const drawing = (id: number) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData uri="ink"><p:contentPart r:id="ink"/></a:graphicData></a:graphic></wp:inline></w:drawing>`
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r>${drawing(1)}${drawing(2)}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('ink', 'customXml', '../assets/ink.xml')),
      'assets/ink.xml': inkXml('unsupported'),
    }))
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 2, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original').map(e => [e.referenceId, e.status, e.selectedRepresentation])).toEqual([['ink', 'unsupported', 'none'], ['ink', 'unsupported', 'none']])
  })

  test('PPTX unsupported horizontal chart retains owners without inventing missing relationships', async () => {
    const doc = await pptx(frame(1, 'chart', '<c:chart r:id="bar"/>') + frame(2, 'chart', '<c:chart r:id="bar"/>'), {
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('bar', 'chart', '../charts/bar.xml')),
      'ppt/charts/bar.xml': horizontalChart,
    })
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 2, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original').every(e => e.status === 'unsupported' && e.reason !== 'relationship-not-found')).toBe(true)
  })

  test('PPTX parsed unsupported chart document without plot area is not interpreted as another relationship', async () => {
    const doc = await pptx(frame(1, 'chart', '<c:chart r:id="bar"/>'), {
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('bar', 'chart', '../charts/bar.xml')),
      'ppt/charts/bar.xml': '<c:chartSpace><c:chart><c:title/></c:chart></c:chartSpace>',
    })
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage![0]).toMatchObject({ status: 'unsupported', reason: 'unsupported-chart', selectedRepresentation: 'none' })
  })

  test('PPTX direct contentPart retains its root relationship for malformed XML', async () => {
    const direct = '<p:contentPart r:id="ink"><p:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></p:xfrm></p:contentPart>'
    const doc = await pptx(direct, {
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('ink', 'customXml', '../ink/ink1.xml')),
      'ppt/ink/ink1.xml': '<ink><trace></ink>',
    })
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage![0]).toMatchObject({ element: 'contentPart', referenceId: 'ink', status: 'malformed', reason: 'invalid-content-xml', selectedRepresentation: 'none' })
  })

  test.each([
    ['valid', '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>', 'native', undefined],
    ['missing', undefined, 'malformed', 'part-not-found'],
    ['malformed', '<c:chartSpace><', 'malformed', 'invalid-content-xml'],
    ['external', undefined, 'unsupported', 'external-relationship'],
  ] as const)('PPTX Choice-wrapped direct contentPart %s retains wrapper reference and concrete outcome', async (mode, payload, status, reason) => {
    const cp = '<p:contentPart r:id="ch"><p:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></p:xfrm></p:contentPart>'
    const body = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p">${cp}</mc:Choice></mc:AlternateContent>`
    const parts: Record<string, string> = {
      'ppt/slides/_rels/slide1.xml.rels': `<Relationships><Relationship Id="ch" Target="${mode === 'external' ? 'https://example.invalid/chart' : '../charts/chart.xml'}"${mode === 'external' ? ' TargetMode="External"' : ''}/></Relationships>`,
    }
    if (payload) parts['ppt/charts/chart.xml'] = payload
    const doc = await pptx(body, parts)
    const original = doc.drawingCoverage!.find(e => e.scope === 'original')!
    expect(original).toMatchObject({ element: 'AlternateContent', referenceId: 'ch', status, ...(reason ? { reason } : {}) })
    expect(drawingReport(doc).counts).toMatchObject({ originalObjects: 1, selectedDescendants: mode === 'valid' ? 1 : 0 })
    if (mode !== 'valid') expect(doc.drawingCoverage!.find(e => e.scope === 'diagnostic')).toMatchObject({ element: 'contentPart', referenceId: 'ch', reason })
  })

  test('Word successful selected vector retains actual relationship and inline textbox invents none', async () => {
    const drawing = (id: number, payload: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData>${payload}</a:graphicData></a:graphic></wp:inline></w:drawing>`
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r>${drawing(1, '<p:contentPart r:id="ink"/>')}${drawing(2, '<a:wsp><a:txbx><w:txbxContent><w:p><w:r><w:t>text</w:t></w:r></w:p></w:txbxContent></a:txbx></a:wsp>')}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('ink', 'customXml', '../assets/ink.xml')),
      'assets/ink.xml': inkXml(),
    }))
    expect(doc.drawingCoverage!.find(e => e.id === '1')).toMatchObject({ referenceId: 'ink', status: 'native', selectedRepresentation: 'inkml' })
    expect(doc.drawingCoverage!.find(e => e.id === '2')?.referenceId).toBeUndefined()
  })
  test('PPTX selected fallback discards unsupported content diagnostics from its Choice', async () => {
    const choice = frame(1, 'ink', '<p:contentPart r:id="bad"/>')
    const body = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:ink="http://www.w3.org/2003/InkML"><mc:Choice Requires="ink">${choice}</mc:Choice><mc:Fallback>${shape(9)}</mc:Fallback></mc:AlternateContent>`
    const doc = await pptx(body, {
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('bad', 'customXml', '../ink/bad.xml')),
      'ppt/ink/bad.xml': inkXml('unsupported'),
    })
    expect(doc.drawingCoverage!.some(e => e.scope === 'diagnostic')).toBe(false)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toMatchObject([{ status: 'fallback', selectedRepresentation: 'native-shape' }])
  })

  test('XLSX recursive content failure retains the original relationship and depth limit once', async () => {
    const parts: Record<string, string> = {
      'xl/workbook.xml': '<workbook><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': '<xdr:wsDr><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><xdr:contentPart r:id="ink"/><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>',
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('ink', 'customXml', '../ink/ink1.xml')),
    }
    for (let n = 1; n <= 35; n++) {
      parts[`xl/ink/ink${n}.xml`] = n === 35 ? inkXml() : '<root><contentPart r:id="next"/></root>'
      if (n < 35) parts[`xl/ink/_rels/ink${n}.xml.rels`] = rels(rel('next', 'customXml', `ink${n + 1}.xml`))
    }
    const doc = await parseXlsx(await packageOf(parts))
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage![0]).toMatchObject({ referenceId: 'ink', status: 'unsupported', selectedRepresentation: 'none', reason: 'reference-depth', limit: 32 })
  })

  test.each([
    ['missing', undefined, false, 'malformed', 'part-not-found'],
    ['malformed', '<ink><trace></ink>', false, 'malformed', 'invalid-content-xml'],
    ['external', undefined, true, 'unsupported', 'external-relationship'],
    ['malformed relationships', undefined, false, 'malformed', 'invalid-relationship-xml'],
  ] as const)('XLSX shared wrapper %s failure reaches both original owners', async (_mode, target, external, status, reason) => {
    const anchor = (rid: string) => `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><xdr:contentPart r:id="${rid}"/><xdr:clientData/></xdr:absoluteAnchor>`
    const neighbor = '<xdr:absoluteAnchor><xdr:pos x="952500" y="0"/><xdr:ext cx="952500" cy="952500"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9" name="green"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:sp><xdr:clientData/></xdr:absoluteAnchor>'
    const parts: Record<string, string> = {
      'xl/workbook.xml': '<workbook><sheets><sheet name="One" sheetId="1" r:id="s1"/><sheet name="Two" sheetId="2" r:id="s2"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml') + rel('s2', 'worksheet', 'worksheets/sheet2.xml')),
      'xl/ink/wrapper.xml': '<root><contentPart r:id="inner"/></root>',
      'xl/ink/_rels/wrapper.xml.rels': _mode === 'malformed relationships' ? '<Relationships><' : `<Relationships><Relationship Id="inner" Target="${external ? 'https://example.invalid/ink.xml' : 'end.xml'}"${external ? ' TargetMode="External"' : ''}/></Relationships>`,
    }
    for (const n of [1, 2]) {
      parts[`xl/worksheets/sheet${n}.xml`] = `<worksheet><sheetData/><drawing r:id="dr${n}"/></worksheet>`
      parts[`xl/worksheets/_rels/sheet${n}.xml.rels`] = rels(rel(`dr${n}`, 'drawing', `../drawings/drawing${n}.xml`))
      parts[`xl/drawings/drawing${n}.xml`] = `<xdr:wsDr>${anchor(`outer${n}`)}${neighbor}</xdr:wsDr>`
      parts[`xl/drawings/_rels/drawing${n}.xml.rels`] = rels(rel(`outer${n}`, 'customXml', '../ink/wrapper.xml'))
    }
    if (target) parts['xl/ink/end.xml'] = target
    const doc = await parseXlsx(await packageOf(parts))
    const originals = doc.drawingCoverage!.filter(e => e.scope === 'original')
    expect(originals).toHaveLength(4)
    expect(originals.filter(e => e.element === 'contentPart').map(e => [e.referenceId, e.status, e.reason, e.selectedRepresentation])).toEqual([
      ['outer1', status, reason, 'none'], ['outer2', status, reason, 'none'],
    ])
    expect(originals.filter(e => e.element === 'sp').every(e => e.status === 'native')).toBe(true)
    expect(drawingReport(doc).counts.diagnosticEntries).toBe(0)
  })

  test('PPTX two slides sharing a malformed terminal part keep their distinct source references', async () => {
    const parts: Record<string, string> = {
      'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s1"/><p:sldId r:id="s2"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s1', 'slide', 'slides/slide1.xml') + rel('s2', 'slide', 'slides/slide2.xml')),
      'ppt/ink/wrapper.xml': '<root><contentPart r:id="inner"/></root>',
      'ppt/ink/_rels/wrapper.xml.rels': rels(rel('inner', 'customXml', 'bad.xml')),
      'ppt/ink/bad.xml': '<ink><trace></ink>',
    }
    for (const n of [1, 2]) {
      parts[`ppt/slides/slide${n}.xml`] = `<p:sld><p:cSld><p:spTree>${frame(n, 'ink', `<p:contentPart r:id="outer${n}"/>`)}${shape(n + 10)}</p:spTree></p:cSld></p:sld>`
      parts[`ppt/slides/_rels/slide${n}.xml.rels`] = rels(rel(`outer${n}`, 'customXml', '../ink/wrapper.xml'))
    }
    const doc = await parsePptx(await packageOf(parts))
    const originals = doc.drawingCoverage!.filter(e => e.scope === 'original')
    expect(originals.filter(e => e.element === 'graphicFrame').map(e => [e.partPath, e.referenceId, e.status, e.reason])).toEqual([
      ['ppt/slides/slide1.xml', 'outer1', 'malformed', 'invalid-content-xml'],
      ['ppt/slides/slide2.xml', 'outer2', 'malformed', 'invalid-content-xml'],
    ])
    expect(originals.filter(e => e.element === 'sp').every(e => e.status === 'native')).toBe(true)
    expect(drawingReport(doc).counts.diagnosticEntries).toBe(0)
  })

  test.each(['selected', 'fallback', 'unowned'] as const)('Word preloaded malformed owner relationships %s have no surplus failure facet', async mode => {
    const drawing = (id: number, payload: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData uri="chart">${payload}</a:graphicData></a:graphic></wp:inline></w:drawing>`
    const selected = drawing(1, '<c:chart r:id="one"/>') + drawing(2, '<c:chart r:id="two"/>')
    const recovered = `<mc:AlternateContent><mc:Choice Requires="c">${drawing(1, '<c:chart r:id="one"/>')}</mc:Choice><mc:Fallback>${drawing(9, '<a:wsp><a:txbx><w:txbxContent><w:p><w:r><w:t>fallback</w:t></w:r></w:p></w:txbxContent></a:txbx></a:wsp>')}</mc:Fallback></mc:AlternateContent>`
    const body = mode === 'selected' ? selected : mode === 'fallback' ? recovered : '<w:t>text</w:t>'
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><w:p><w:r>${body}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': '<Relationships><',
    }))
    const entries = doc.drawingCoverage!
    expect(drawingReport(doc).counts.diagnosticEntries).toBe(mode === 'unowned' ? 1 : 0)
    if (mode === 'selected') expect(entries.filter(e => e.scope === 'original').map(e => [e.status, e.reason])).toEqual([['malformed', 'invalid-relationship-xml'], ['malformed', 'invalid-relationship-xml']])
    if (mode === 'fallback') expect(entries.find(e => e.scope === 'original')).toMatchObject({ status: 'fallback', selectedRepresentation: 'textbox' })
  })

  test.each(['selected', 'fallback', 'unowned'] as const)('XLSX preloaded malformed owner relationships %s have no surplus failure facet', async mode => {
    const anchor = (body: string) => `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`
    const native = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:sp>'
    const selected = anchor('<xdr:contentPart r:id="one"/>') + anchor('<xdr:contentPart r:id="two"/>') + anchor(native)
    const recovered = anchor(`<mc:AlternateContent><mc:Choice Requires="xdr"><xdr:contentPart r:id="one"/></mc:Choice><mc:Fallback>${native}</mc:Fallback></mc:AlternateContent>`)
    const body = mode === 'selected' ? selected : mode === 'fallback' ? recovered : anchor(native)
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/a.xml')),
      'xl/drawings/a.xml': `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${body}</xdr:wsDr>`,
      'xl/drawings/_rels/a.xml.rels': '<Relationships><',
    }))
    const entries = doc.drawingCoverage!
    expect(drawingReport(doc).counts.diagnosticEntries).toBe(mode === 'unowned' ? 1 : 0)
    if (mode === 'selected') expect(entries.filter(e => e.element === 'contentPart').map(e => [e.status, e.reason])).toEqual([['malformed', 'invalid-relationship-xml'], ['malformed', 'invalid-relationship-xml']])
    if (mode === 'fallback') expect(entries.find(e => e.scope === 'original')).toMatchObject({ status: 'fallback', selectedRepresentation: 'native-shape' })
  })
  test.each([false, true])('Word drawing budget marks only the rejected %s chart placement', async realBoundary => {
    const chart = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    const drawing = (id: number) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData uri="chart"><c:chart r:id="ch"/></a:graphicData></a:graphic></wp:inline></w:drawing>`
    const count = realBoundary ? 10001 : 2
    const pkg = await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r>${Array.from({ length: count }, (_, i) => drawing(i + 1)).join('')}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('ch', 'chart', 'charts/good.xml')),
      'word/charts/good.xml': chart,
    })
    if (!realBoundary) for (let i = 0; i < 9999; i++) reserveDrawingNode(drawingPartContext(pkg), 'earlier/part.xml')
    const doc = await parseDocx(pkg)
    expect(doc.sections.flatMap(s => s.paragraphs).flatMap(p => p.images)).toHaveLength(realBoundary ? 10000 : 1)
    const originals = doc.drawingCoverage!.filter(e => e.scope === 'original')
    expect(originals).toHaveLength(count)
    expect(originals.slice(0, -1).every(e => e.status === 'native' && e.selectedRepresentation === 'column-chart')).toBe(true)
    expect(originals.at(-1)).toMatchObject({ status: 'unsupported', reason: 'drawing node budget exceeded', selectedRepresentation: 'none' })
    expect(originals.at(-1)?.limit).toBe(10000)
    expect(drawingReport(doc).counts.diagnosticEntries).toBe(0)
  }, 15000)

  test('Word repeated malformed chart targets attach to both original placements', async () => {
    const drawing = (id: number) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData uri="chart"><c:chart r:id="broken"/></a:graphicData></a:graphic></wp:inline></w:drawing>`
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r>${drawing(1)}${drawing(2)}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('broken', 'chart', 'charts/broken.xml')),
      'word/charts/broken.xml': '<c:chartSpace><',
    }))
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 2, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original').map(e => [e.status, e.reason])).toEqual([['malformed', 'invalid-content-xml'], ['malformed', 'invalid-content-xml']])
  })

  test('PPTX malformed chart targets attach across repeated placements and slides', async () => {
    const badFrame = (id: number) => frame(id, 'chart', '<c:chart r:id="broken"/>')
    const doc = await parsePptx(await packageOf({
      'ppt/presentation.xml': '<p:presentation><p:sldIdLst><p:sldId r:id="s1"/><p:sldId r:id="s2"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s1', 'slide', 'slides/slide1.xml') + rel('s2', 'slide', 'slides/slide2.xml')),
      'ppt/slides/slide1.xml': `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree>${badFrame(1)}${badFrame(2)}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/slide2.xml': `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree>${badFrame(3)}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('broken', 'chart', '../charts/broken.xml')),
      'ppt/slides/_rels/slide2.xml.rels': rels(rel('broken', 'chart', '../charts/broken.xml')),
      'ppt/charts/broken.xml': '<c:chartSpace><',
    }))
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 3, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original').map(e => [e.status, e.reason])).toEqual(Array(3).fill(['malformed', 'invalid-content-xml']))
  })

  test('PPTX same relationship ID on another slide cannot inherit a malformed target', async () => {
    const chart = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    const slide = (id: number) => `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree>${frame(id, 'chart', '<c:chart r:id="shared"/>')}</p:spTree></p:cSld></p:sld>`
    const doc = await parsePptx(await packageOf({
      'ppt/presentation.xml': '<p:presentation><p:sldIdLst><p:sldId r:id="s1"/><p:sldId r:id="s2"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s1', 'slide', 'slides/slide1.xml') + rel('s2', 'slide', 'slides/slide2.xml')),
      'ppt/slides/slide1.xml': slide(1), 'ppt/slides/slide2.xml': slide(2),
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('shared', 'chart', '../charts/good.xml')),
      'ppt/slides/_rels/slide2.xml.rels': rels(rel('shared', 'chart', '../charts/broken.xml')),
      'ppt/charts/good.xml': chart, 'ppt/charts/broken.xml': '<c:chartSpace><',
    }))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original').map(e => [e.status, e.reason])).toEqual([['native', undefined], ['malformed', 'invalid-content-xml']])
    expect(drawingReport(doc).counts.diagnosticEntries).toBe(0)
  })

  test('unowned malformed part remains a diagnostic facet', async () => {
    const pkg = await packageOf({ 'word/document.xml': '<w:document><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>' })
    contentDiagnostic(drawingPartContext(pkg), 'malformed-part', 'word/document.xml', 'chart', { identity: 'unowned', reason: 'invalid-content-xml' })
    const doc = await parseDocx(pkg)
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 0, selectedDescendants: 0, diagnosticEntries: 1 })
    expect(doc.drawingCoverage![0]).toMatchObject({ scope: 'diagnostic', status: 'malformed', reason: 'invalid-content-xml' })
  })
  test('ordered and standard XML retain inherited and locally shadowed namespace bindings', () => {
    const xml = '<root xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><Choice Requires="d"/><wrapper xmlns:d="urn:unsupported"><Choice Requires="d"/></wrapper></root>'
    for (const parse of [parseXmlOrdered, parseXml]) {
      const root = parse(xml)
      expect(namespaceUri(getChildren(root, 'Choice')[0], 'd')).toBe('http://schemas.openxmlformats.org/drawingml/2006/main')
      expect(namespaceUri(getChildren(getChildren(root, 'wrapper')[0], 'Choice')[0], 'd')).toBe('urn:unsupported')
      expect(namespaceUri(getChildren(root, 'Choice')[0], 'missing')).toBeUndefined()
    }
  })
  test('native shapes and geometry failures remain distinct while text survives', async () => {
    const invalid = '<a:custGeom><a:pathLst><a:path><a:moveTo><a:pt x="missing" y="0"/></a:moveTo></a:path></a:pathLst></a:custGeom>'
    const doc = await pptx(shape(2) + shape(3, '<a:prstGeom prst="unknownFuturePreset"/>', 'Retained') + shape(4, invalid, 'Still here'))
    const entries = doc.drawingCoverage!
    expect(entries.filter(e => e.scope === 'original')).toHaveLength(3)
    expect(entries.find(e => e.id === '2')).toMatchObject({ status: 'native', feature: 'shape', representation: 'native', selectedRepresentation: 'native-shape', partPath: 'ppt/slides/slide1.xml' })
    expect(entries.find(e => e.id === '3')).toMatchObject({ status: 'unsupported', feature: 'geometry', reason: expect.stringContaining('unknown-preset') })
    expect(entries.find(e => e.id === '4')).toMatchObject({ status: 'malformed', feature: 'geometry', reason: expect.stringContaining('invalid') })
    expect(doc.slides[0].shapes[1].textBody?.paragraphs[0].runs[0].text).toBe('Retained')
    expect(doc.slides[0].shapes[2].textBody?.paragraphs[0].runs[0].text).toBe('Still here')
    expect(JSON.parse(JSON.stringify(drawingReport(doc))).counts.originalObjects).toBe(3)
  })

  test.each([['ChartEx', 'cx1', frame(3, 'http://schemas.microsoft.com/office/drawing/2014/chartex')], ['OMML', 'a14', `${shape(3)}<m:oMath/>`], ['model3D', 'am3d', `${shape(3)}<am3d:model3D/>`]])('%s uses one fallback source object and selected picture', async (feature, requires, choice) => {
    const doc = await pptx(alternative(choice, picture(3, 'im'), requires), {
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('im', 'image', '../media/fallback.png')),
      'ppt/media/fallback.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ status: 'fallback', feature, representation: 'fallback', selectedRepresentation: 'raster-fallback', id: '3' })
    expect(doc.slides[0].shapes).toHaveLength(1)
    expect(doc.images).toHaveLength(1)
  })

  test('unsupported graphicData and missing picture stay in source inventory', async () => {
    const doc = await pptx(frame(7, 'urn:future:graphic') + picture(8, 'absent'))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(2)
    expect(doc.drawingCoverage!.find(e => e.id === '7')).toMatchObject({ status: 'unsupported', feature: 'graphicData' })
    expect(doc.drawingCoverage!.find(e => e.id === '8')).toMatchObject({ status: 'malformed', feature: 'image', reason: expect.stringContaining('missing') })
  })

  test('PPTX selected fallback excludes the unselected missing part diagnostic', async () => {
    const doc = await pptx(alternative(frame(13, 'chart', '<c:chart r:id="missing"/>'), picture(13, 'im'), 'c'), {
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('missing', 'chart', '../charts/absent.xml') + rel('im', 'image', '../media/fallback.png')),
      'ppt/media/fallback.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '13', status: 'fallback', selectedRepresentation: 'raster-fallback' })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'diagnostic')).toHaveLength(0)
  })

  test('PPTX Requires resolves an alias to its inherited namespace URI', async () => {
    const body = alternative(shape(61), picture(62, 'im'), 'alias')
    const doc = await pptx(body, {
      'ppt/slides/slide1.xml': `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:alias="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('im', 'image', '../media/f.png')),
      'ppt/media/f.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    })
    expect(doc.slides[0].shapes.map(s => s.source?.id)).toEqual(['61'])
    expect(doc.drawingCoverage![0]).toMatchObject({ representation: 'choice', status: 'native' })
    expect(doc.images).toHaveLength(0)
  })

  test('PPTX local prefix shadow rejects a known spelling bound to an unknown URI', async () => {
    const body = alternative(shape(63), picture(64, 'im'), 'a').replace('<mc:Choice Requires="a">', '<mc:Choice xmlns:a="urn:unsupported" Requires="a">')
    const doc = await pptx(body, {
      'ppt/slides/slide1.xml': `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('im', 'image', '../media/f.png')),
      'ppt/media/f.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    })
    expect(doc.slides[0].shapes.map(s => s.source?.id)).toEqual(['64'])
    expect(doc.drawingCoverage![0]).toMatchObject({ representation: 'fallback', status: 'fallback' })
  })

  test('supported empty Choice is blank and does not promote a fallback image', async () => {
    const doc = await pptx(alternative('', picture(9, 'im'), 'a'))
    expect(doc.slides[0].shapes).toHaveLength(0)
    expect(doc.images).toHaveLength(0)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ status: 'native', representation: 'choice', selectedRepresentation: 'blank' })
  })

  test('later empty Choice cannot hide an earlier selected unsupported graphicData', async () => {
    const body = `<mc:AlternateContent><mc:Choice Requires="a">${frame(301, 'urn:unsupported:graphic')}</mc:Choice><mc:Choice Requires="a"/><mc:Fallback>${shape(302)}</mc:Fallback></mc:AlternateContent>`
    const doc = await pptx(body)
    expect(doc.slides[0].shapes).toHaveLength(0)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '301', feature: 'graphicData', status: 'unsupported', selectedRepresentation: 'none', representation: 'choice' })
    expect(doc.drawingCoverage![0].reason).not.toContain('empty Choice')
  })

  test('later empty Choice is blank when an earlier unknown namespace is skipped', async () => {
    const body = `<mc:AlternateContent><mc:Choice Requires="future">${frame(311, 'urn:unsupported:graphic')}</mc:Choice><mc:Choice Requires="a"/><mc:Fallback>${shape(312)}</mc:Fallback></mc:AlternateContent>`
    const doc = await pptx(body)
    expect(doc.slides[0].shapes).toHaveLength(0)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ feature: 'empty-choice', status: 'native', selectedRepresentation: 'blank', representation: 'choice' })
  })

  test.each([false, true])('PPTX selected Choice retains scene-less sibling source objects (reverse=%s)', async reverse => {
    const children = [shape(401), frame(402, 'urn:unsupported:graphic')]
    const doc = await pptx(alternative((reverse ? children.reverse() : children).join(''), '', 'a'))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage!.find(e => e.id === '402')).toMatchObject({ status: 'unsupported', feature: 'graphicData', selectedRepresentation: 'none' })
    expect(doc.drawingCoverage!.some(e => e.id === '401' && e.status === 'native')).toBe(true)
  })

  test('PPTX selected missing-chart Choice without fallback retains source identity and failure', async () => {
    const doc = await pptx(`<mc:AlternateContent><mc:Choice Requires="c">${frame(403, 'chart', '<c:chart r:id="missing"/>')}</mc:Choice></mc:AlternateContent>`, {
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('missing', 'chart', '../charts/absent.xml')),
    })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage!.some(e => e.id === '403' && e.status === 'malformed' && e.selectedRepresentation === 'none')).toBe(true)
  })

  test.each([false, true])('PPTX repeated missing chart references classify every %s placement', async wrapped => {
    const body = [1, 2].map(id => wrapped ? `<mc:AlternateContent><mc:Choice Requires="c">${frame(id, 'chart', '<c:chart r:id="missing"/>')}</mc:Choice></mc:AlternateContent>` : frame(id, 'chart', '<c:chart r:id="missing"/>')).join('')
    const doc = await pptx(body, { 'ppt/slides/_rels/slide1.xml.rels': rels(rel('missing', 'chart', '../charts/absent.xml')) })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(2)
    for (const id of ['1', '2']) expect(doc.drawingCoverage!.find(e => e.id === id && e.scope === 'original')).toMatchObject({ status: 'malformed', referenceId: 'missing', reason: 'part-not-found', selectedRepresentation: 'none' })
  })

  test('a later supported Choice does not inherit an unselected ChartEx feature', async () => {
    const choices = `<mc:AlternateContent><mc:Choice Requires="cx1">${frame(10, 'http://schemas.microsoft.com/office/drawing/2014/chartex')}</mc:Choice><mc:Choice Requires="a">${shape(11)}</mc:Choice><mc:Fallback>${picture(12, 'unused')}</mc:Fallback></mc:AlternateContent>`
    const doc = await pptx(choices)
    expect(doc.slides[0].shapes.map(s => s.source?.id)).toEqual(['11'])
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '11', feature: 'shape', status: 'native', representation: 'choice' })
    expect(doc.images).toHaveLength(0)
  })

  test('each additional selected Choice shape reports its own failed geometry and retained text', async () => {
    const doc = await pptx(alternative(shape(81) + shape(82, '<a:prstGeom prst="unknownFuturePreset"/>', 'Retained'), '', 'a'))
    expect(doc.slides[0].shapes).toHaveLength(2)
    expect(doc.slides[0].shapes[1].textBody?.paragraphs[0].runs[0].text).toBe('Retained')
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage!.find(e => e.id === '82')).toMatchObject({ scope: 'descendant', feature: 'geometry', status: 'unsupported', selectedRepresentation: 'text-only', reason: expect.stringContaining('unknown-preset') })
  })

  test('nested PPTX fallback chart reports its actual compatibility selection on the owning frame', async () => {
    const columns = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    const nested = '<mc:AlternateContent><mc:Choice Requires="future"><future:unknown/></mc:Choice><mc:Fallback><a:graphicData uri="chart"><c:chart r:id="chart"/></a:graphicData></mc:Fallback></mc:AlternateContent>'
    const doc = await pptx(frame(91, 'unused', nested), {
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('chart', 'chart', '../charts/chart1.xml')),
      'ppt/charts/chart1.xml': columns,
    })
    expect(doc.slides[0].shapes[0].content?.kind).toBe('chart')
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '91', feature: 'chart', status: 'fallback', representation: 'fallback', selectedRepresentation: 'column-chart', reason: expect.stringContaining('fallback') })
  })

  test('nested PPTX supported Choice reports choice provenance, including intentional blank', async () => {
    const columns = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    const nested = (choice: string) => frame(91, 'chart', `<mc:AlternateContent><mc:Choice Requires="c">${choice}</mc:Choice><mc:Fallback><c:chart r:id="good"/></mc:Fallback></mc:AlternateContent>`)
    const parts = { 'ppt/slides/_rels/slide1.xml.rels': rels(rel('good', 'chart', '../charts/chart1.xml')), 'ppt/charts/chart1.xml': columns }
    const chosen = await pptx(nested('<c:chart r:id="good"/>'), parts)
    expect(chosen.drawingCoverage![0]).toMatchObject({ status: 'native', representation: 'choice', selectedRepresentation: 'column-chart' })
    const blank = await pptx(nested(''), parts)
    expect(blank.slides[0].shapes).toHaveLength(0)
    expect(blank.drawingCoverage![0]).toMatchObject({ feature: 'empty-choice', status: 'native', representation: 'choice', selectedRepresentation: 'blank' })
  })

  test.each([
    ['unknown preset', '<a:prstGeom prst="unknownFuturePreset"/>', 'unsupported', 'unknown-preset'],
    ['invalid custom path', '<a:custGeom><a:pathLst><a:path><a:moveTo><a:pt x="missing" y="0"/></a:moveTo></a:path></a:pathLst></a:custGeom>', 'malformed', 'invalid'],
  ])('selected fallback descendant preserves %s failure and retained text', async (_label, geometry, status, reason) => {
    const doc = await pptx(alternative('', shape(101) + shape(102, geometry, 'Retained'), 'future'))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage!.find(e => e.id === '102')).toMatchObject({ scope: 'descendant', feature: 'geometry', status,
      representation: 'fallback', selectedRepresentation: 'text-only', reason: expect.stringContaining(reason) })
    expect(doc.slides[0].shapes[1].textBody?.paragraphs[0].runs[0].text).toBe('Retained')
  })

  test.each([false, true])('XLSX %s drawing-part reuse keeps failures tied to the owning relationship', async reusePart => {
    const pic = '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="im"/></xdr:blipFill><xdr:spPr/></xdr:pic>'
    const anchor = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${pic}<xdr:clientData/></xdr:absoluteAnchor>`
    const drawing = `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchor}</xdr:wsDr>`
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="s1"/><sheet name="Two" sheetId="2" r:id="s2"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml') + rel('s2', 'worksheet', 'worksheets/sheet2.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/sheet2.xml': '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/worksheets/_rels/sheet2.xml.rels': rels(rel('dr', 'drawing', reusePart ? '../drawings/drawing1.xml' : '../drawings/drawing2.xml')),
      'xl/drawings/drawing1.xml': drawing, 'xl/drawings/drawing2.xml': drawing,
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('im', 'image', '../media/absent.png')),
      'xl/drawings/_rels/drawing2.xml.rels': rels(rel('im', 'image', '../media/good.png')),
      'xl/media/good.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    }))
    const entries = doc.drawingCoverage!
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 2, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(entries[0]).toMatchObject({ unit: 0, referenceId: 'im', status: 'malformed', reason: 'part-not-found' })
    expect(entries[1]).toMatchObject({ unit: 1, referenceId: 'im', status: reusePart ? 'malformed' : 'native', selectedRepresentation: reusePart ? 'none' : 'picture' })
  })

  test('XLSX recovered fallback stays clean beside a selected missing chart sharing its rejected reference', async () => {
    const frame = (id: number) => `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></xdr:xfrm><a:graphic><a:graphicData uri="chart"><c:chart r:id="missing"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    const shape = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:sp>'
    const anchor = (body: string) => `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`
    const recovered = `<mc:AlternateContent><mc:Choice Requires="c">${frame(1)}</mc:Choice><mc:Fallback>${shape}</mc:Fallback></mc:AlternateContent>`
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${anchor(recovered)}${anchor(frame(3))}</xdr:wsDr>`,
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('missing', 'chart', '../charts/absent.xml')),
    }))
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 2, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage!.find(entry => entry.id === '2')).toMatchObject({ status: 'fallback', selectedRepresentation: 'native-shape' })
    expect(doc.drawingCoverage!.find(entry => entry.id === '3')).toMatchObject({ referenceId: 'missing', status: 'malformed', reason: 'part-not-found' })
  })

  test('XLSX malformed referenced chart part belongs to its selected placement', async () => {
    const frame = '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></xdr:xfrm><a:graphic><a:graphicData uri="chart"><c:chart r:id="broken"/></a:graphicData></a:graphic></xdr:graphicFrame>'
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${frame}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('broken', 'chart', '../charts/broken.xml')),
      'xl/charts/broken.xml': '<c:chartSpace><c:chart><',
    }))
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage![0]).toMatchObject({ referenceId: 'broken', status: 'malformed', reason: 'invalid-content-xml' })
  })

  test.each([false, true])('XLSX valid raster fallback remains successful beside a failed chart with the same target, reversed=%s', async reversed => {
    const pngCanvas = createCanvas(2, 2)
    pngCanvas.getContext('2d').fillRect(0, 0, 2, 2)
    const png = pngCanvas.toBuffer('image/png')
    const frame = (id: number) => `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></xdr:xfrm><a:graphic><a:graphicData uri="chart"><c:chart r:id="shared"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    const pic = '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="shared"/></xdr:blipFill><xdr:spPr/></xdr:pic>'
    const recovered = `<mc:AlternateContent><mc:Choice Requires="c">${frame(1)}</mc:Choice><mc:Fallback>${pic}</mc:Fallback></mc:AlternateContent>`
    const anchor = (body: string) => `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`
    const bodies = reversed ? [frame(3), recovered] : [recovered, frame(3)]
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${bodies.map(anchor).join('')}</xdr:wsDr>`,
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('shared', 'image', '../media/image.png')),
      'xl/media/image.png': png,
    }))
    expect(doc.images).toHaveLength(1)
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 2, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage!.find(e => e.id === '2')).toMatchObject({ status: 'fallback', selectedRepresentation: 'raster-fallback', representation: 'fallback' })
    expect(doc.drawingCoverage!.find(e => e.id === '3')).toMatchObject({ status: 'malformed', reason: 'invalid-content-xml' })
  })

  test('nested fallback on an owning shape preserves its geometry failure', async () => {
    const body = shape(201, '<a:prstGeom prst="unknownFuturePreset"/>', 'Retained')
      .replace('<a:t>Retained</a:t>', '<mc:AlternateContent><mc:Choice Requires="future"><future:unknown/></mc:Choice><mc:Fallback><a:t>Retained</a:t></mc:Fallback></mc:AlternateContent>')
    const doc = await pptx(body)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '201', feature: 'geometry', status: 'unsupported', representation: 'fallback',
      selectedRepresentation: 'text-only', reason: expect.stringContaining('unknown-preset') })
  })

  test('group depth limit retains source identity and exceeded limit', async () => {
    const groupXfrm = '<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm>'
    let nested = shape(100)
    for (let id = 65; id >= 1; id--) nested = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/></p:nvGrpSpPr><p:grpSpPr>${groupXfrm}</p:grpSpPr>${nested}</p:grpSp>`
    const doc = await pptx(nested)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage!.find(e => e.id === '65')).toMatchObject({ status: 'unsupported', reason: 'group-depth', limit: 64 })
  })

  test('Word retains one source placement with selected native content and missing payload', async () => {
    const drawing = (id: number, payload: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}" name="Drawing ${id}"/><a:graphic><a:graphicData>${payload}</a:graphicData></a:graphic></wp:inline></w:drawing>`
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document><w:body><w:p><w:r>${drawing(1, '<pic:pic><pic:blipFill><a:blip r:embed="im"/></pic:blipFill></pic:pic>')}</w:r><w:r>${drawing(2, '<c:chart r:id="missing"/>')}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('im', 'image', 'media/a.png') + rel('missing', 'chart', 'charts/absent.xml')),
      'word/media/a.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    }))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(2)
    expect(doc.drawingCoverage!.find(e => e.id === '1')).toMatchObject({ status: 'native', feature: 'picture' })
    expect(doc.drawingCoverage!.find(e => e.id === '2')).toMatchObject({ status: 'malformed', feature: 'chart' })
    expect(doc.drawingCoverage!.find(e => e.id === '1')?.unit).toBeUndefined()
  })

  test('Word supported empty Choice remains blank and records its source wrapper', async () => {
    const fallback = '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="20"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="im"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="w"/><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('im', 'image', 'media/a.png')),
      'word/media/a.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    }))
    expect(doc.sections[0].paragraphs[0].images).toHaveLength(0)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ feature: 'empty-choice', status: 'native', selectedRepresentation: 'blank', representation: 'choice' })
  })

  test('Word fallback excludes diagnostics from an unselected broken Choice', async () => {
    const placed = (id: number, payload: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData>${payload}</a:graphicData></a:graphic></wp:inline></w:drawing>`
    const xml = `<mc:AlternateContent><mc:Choice Requires="c">${placed(40, '<c:chart r:id="missing"/>')}</mc:Choice><mc:Fallback>${placed(41, '<pic:pic><pic:blipFill><a:blip r:embed="im"/></pic:blipFill></pic:pic>')}</mc:Fallback></mc:AlternateContent>`
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document><w:body><w:p><w:r>${xml}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('missing', 'chart', 'charts/absent.xml') + rel('im', 'image', 'media/a.png')),
      'word/media/a.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    }))
    expect(doc.sections[0].paragraphs[0].images).toHaveLength(1)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '41', status: 'fallback', representation: 'fallback', selectedRepresentation: 'raster-fallback' })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'diagnostic')).toHaveLength(0)
  })

  test('Word unknown namespace Choice does not enter selected picture inventory', async () => {
    const placed = (id: number, rid: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
    const xml = `<mc:AlternateContent><mc:Choice Requires="future">${placed(50, 'choice')}</mc:Choice><mc:Fallback>${placed(51, 'fallback')}</mc:Fallback></mc:AlternateContent>`
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document><w:body><w:p><w:r>${xml}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('choice', 'image', 'media/c.png') + rel('fallback', 'image', 'media/f.png')),
      'word/media/c.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
      'word/media/f.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    }))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '51', status: 'fallback', selectedRepresentation: 'raster-fallback' })
  })

  test('Word preloads the selected chart fallback after rejecting an unknown Choice', async () => {
    const placed = (id: number, rid: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}" name="source${id}"/><a:graphic><a:graphicData><c:chart r:id="${rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing>`
    const chart = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="future">${placed(1, 'c1')}</mc:Choice><mc:Fallback>${placed(2, 'c2')}</mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('c1', 'chart', 'charts/c1.xml') + rel('c2', 'chart', 'charts/c2.xml')),
      'word/charts/c1.xml': chart, 'word/charts/c2.xml': chart,
    }))
    expect(doc.sections[0].paragraphs[0].images[0].drawing?.kind).toBe('chart')
    const ctx = createCanvas(100, 100).getContext('2d')
    paintDrawing(doc.sections[0].paragraphs[0].images[0].drawing!, ctx as unknown as CanvasRenderingContext2D, 100, 100)
    expect(ctx.getImageData(0, 0, 100, 100).data.some(value => value !== 0)).toBe(true)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '2', status: 'fallback', selectedRepresentation: 'column-chart', representation: 'fallback' })
  })

  test('Word inventories every rejected drawing placement without a fallback', async () => {
    const placed = (id: number) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}" name="source${id}"/><a:graphic><a:graphicData><future:graphic/></a:graphicData></a:graphic></wp:inline></w:drawing>`
    const doc = await parseDocx(await packageOf({ 'word/document.xml': `<w:document><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="future">${placed(1)}${placed(2)}</mc:Choice></mc:AlternateContent></w:r></w:p></w:body></w:document>` }))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original').map(e => e.id)).toEqual(['1', '2'])
    expect(doc.drawingCoverage!.every(e => e.status === 'unsupported' && e.selectedRepresentation === 'none')).toBe(true)
  })

  test.each([false, true])('Word repeated missing chart references classify every %s placement', async wrapped => {
    const placed = (id: number) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}" name="source${id}"/><a:graphic><a:graphicData><c:chart r:id="missing"/></a:graphicData></a:graphic></wp:inline></w:drawing>`
    const body = [1, 2].map(id => wrapped ? `<mc:AlternateContent><mc:Choice Requires="c">${placed(id)}</mc:Choice></mc:AlternateContent>` : placed(id)).join('')
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r>${body}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('missing', 'chart', 'charts/absent.xml')),
    }))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(2)
    for (const id of ['1', '2']) expect(doc.drawingCoverage!.find(e => e.id === id)).toMatchObject({ status: 'malformed', referenceId: 'missing', reason: 'part-not-found', selectedRepresentation: 'none' })
  })

  test('Word selected empty Choice excludes a later missing-chart diagnostic', async () => {
    const drawing = '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="44"/><a:graphic><a:graphicData><c:chart r:id="missing"/></a:graphicData></a:graphic></wp:inline></w:drawing>'
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="c"/><mc:Fallback>${drawing}</mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('missing', 'chart', 'charts/absent.xml')),
    }))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ feature: 'empty-choice', selectedRepresentation: 'blank' })
    expect(doc.drawingCoverage!.filter(e => e.scope === 'diagnostic')).toHaveLength(0)
  })

  test('Word nested empty Choice remains blank and nested failed chart does not leak after fallback', async () => {
    const chart = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    const make = async (choice: string) => parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="71"/><a:graphic><a:graphicData uri="chart"><mc:AlternateContent><mc:Choice Requires="c">${choice}</mc:Choice><mc:Fallback><c:chart r:id="good"/></mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('missing', 'chart', 'charts/absent.xml') + rel('good', 'chart', 'charts/good.xml')),
      'word/charts/good.xml': chart,
    }))
    const blank = await make('')
    expect(blank.sections[0].paragraphs[0].images).toHaveLength(0)
    expect(blank.drawingCoverage![0]).toMatchObject({ id: '71', feature: 'empty-choice', status: 'native', representation: 'choice', selectedRepresentation: 'blank' })
    const recovered = await make('<c:chart r:id="missing"/>')
    expect(recovered.sections[0].paragraphs[0].images[0].drawing?.kind).toBe('chart')
    expect(recovered.drawingCoverage![0]).toMatchObject({ id: '71', status: 'fallback', representation: 'fallback', selectedRepresentation: 'column-chart' })
    expect(recovered.drawingCoverage!.filter(e => e.scope === 'diagnostic')).toHaveLength(0)
  })

  test('Word outer Choice containing an intentional nested blank does not promote its fallback', async () => {
    const blank = '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="80"/><a:graphic><a:graphicData><mc:AlternateContent><mc:Choice Requires="c"/><mc:Fallback><c:chart r:id="missing"/></mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></wp:inline></w:drawing>'
    const picture = '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="81"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="im"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="w">${blank}</mc:Choice><mc:Fallback>${picture}</mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('missing', 'chart', 'charts/absent.xml') + rel('im', 'image', 'media/a.png')),
      'word/media/a.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    }))
    expect(doc.sections[0].paragraphs[0].images).toHaveLength(0)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '80', feature: 'empty-choice', selectedRepresentation: 'blank' })
  })

  test.each(['chart-then-empty', 'empty-then-chart', 'outer-choice'])('Word %s retains the neighboring selected chart', async order => {
    const chart = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    const empty = '<mc:AlternateContent><mc:Choice Requires="c"/><mc:Fallback><c:chart r:id="c2"/></mc:Fallback></mc:AlternateContent>'
    const selected = order === 'empty-then-chart' ? `${empty}<c:chart r:id="c1"/>` : `<c:chart r:id="c1"/>${empty}`
    const payload = order === 'outer-choice' ? `<mc:AlternateContent><mc:Choice Requires="c">${selected}</mc:Choice><mc:Fallback><c:chart r:id="c2"/></mc:Fallback></mc:AlternateContent>` : selected
    const drawing = `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="81"/><a:graphic><a:graphicData uri="chart">${payload}</a:graphicData></a:graphic></wp:inline></w:drawing>`
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r>${drawing}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('c1', 'chart', 'charts/c1.xml') + rel('c2', 'chart', 'charts/c2.xml')),
      'word/charts/c1.xml': chart, 'word/charts/c2.xml': chart,
    }))
    expect(doc.sections[0].paragraphs[0].images[0]?.drawing?.kind).toBe('chart')
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '81', feature: 'chart', selectedRepresentation: 'column-chart' })
  })

  test.each([
    ['unsupported', 'future', '<future:unknown/>', 'unsupported'],
    ['missing chart', 'c', '<c:chart r:id="absent"/>', 'malformed'],
  ])('Word %s Choice without fallback retains original source placement', async (_label, requires, payload, status) => {
    const drawing = `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="99" name="source 99"/><a:graphic><a:graphicData>${payload}</a:graphicData></a:graphic></wp:inline></w:drawing>`
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="${requires}">${drawing}</mc:Choice></mc:AlternateContent></w:r></w:p></w:body></w:document>`,
    }))
    expect(doc.sections[0].paragraphs[0].images).toHaveLength(0)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage!.find(e => e.scope === 'original')).toMatchObject({ id: '99', name: 'source 99', status, selectedRepresentation: 'none', partPath: 'word/document.xml', treePath: expect.stringContaining('AlternateContent') })
  })

  test('reused Word header part contributes one original source placement', async () => {
    const headerRef = '<w:sectPr><w:headerReference r:id="h" type="default"/></w:sectPr>'
    const drawing = '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="30" name="header image"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="im"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    const doc = await parseDocx(await packageOf({
      'word/document.xml': `<w:document><w:body><w:p><w:pPr>${headerRef}</w:pPr></w:p><w:p><w:pPr>${headerRef}</w:pPr></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('h', 'header', 'header1.xml')),
      'word/header1.xml': `<w:hdr><w:p><w:r>${drawing}</w:r></w:p></w:hdr>`,
      'word/_rels/header1.xml.rels': rels(rel('im', 'image', 'media/a.png')),
      'word/media/a.png': new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    }))
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original' && e.partPath === 'word/header1.xml')).toHaveLength(1)
    expect(doc.drawingCoverage!.find(e => e.partPath === 'word/header1.xml')?.treePath).toMatch(/^header\//)
  })

  test('XLSX counts anchors separately from group descendants', async () => {
    const zip = await packageOf({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="1" name="group"/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm></xdr:grpSpPr><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="child"/></xdr:nvSpPr><xdr:spPr>${xfrm}</xdr:spPr></xdr:sp></xdr:grpSp><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
    })
    const doc = await parseXlsx(zip)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'descendant')).toHaveLength(1)
    expect(drawingReport(doc).counts).toMatchObject({ originalObjects: 1, selectedDescendants: 1, diagnosticEntries: 0 })
  })

  test.each([
    ['two native pictures', ['pic', 'pic'], [1, 2], ['badim', 'badim'], 2, 0],
    ['two native charts', ['chart', 'chart'], [1, 2], ['missing', 'missing'], 2, 0],
    ['two fallback pictures', ['fallback-pic', 'fallback-pic'], [2, 4], ['badim', 'badim'], 2, 0],
    ['grouped pictures with valid neighbor', ['group', 'shape'], [1, 2], ['badim', 'badim'], 2, 2],
    ['distinct chart and picture relationships', ['chart', 'pic'], [1, 2], ['missing', 'badim'], 2, 0],
  ] as const)('XLSX associates %s failures with every selected source object', async (_label, kinds, expectedIds, references, originals, descendants) => {
    const pic = (id: number) => `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="source${id}"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="badim"/></xdr:blipFill><xdr:spPr>${xfrm}</xdr:spPr></xdr:pic>`
    const chart = (id: number) => `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="source${id}"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></xdr:xfrm><a:graphic><a:graphicData uri="chart"><c:chart r:id="missing"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    const shape = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="999"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:sp>'
    const group = `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="10"/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm></xdr:grpSpPr>${pic(1)}${pic(2)}</xdr:grpSp>`
    const fallback = (id: number) => `<mc:AlternateContent><mc:Choice Requires="future">${shape}</mc:Choice><mc:Fallback>${pic(id)}</mc:Fallback></mc:AlternateContent>`
    const object = (kind: string, i: number) => kind === 'pic' ? pic(i + 1) : kind === 'chart' ? chart(i + 1) : kind === 'fallback-pic' ? fallback(i * 2 + 2) : kind === 'group' ? group : shape
    const anchor = (body: string) => `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:unsupported">${kinds.map((kind, i) => anchor(object(kind, i))).join('')}</xdr:wsDr>`,
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('missing', 'chart', '../charts/absent.xml') + rel('badim', 'image', '../media/absent.png')),
    }))
    const entries = doc.drawingCoverage!
    expect(drawingReport(doc).counts).toEqual({ originalObjects: originals, selectedDescendants: descendants, diagnosticEntries: 0 })
    for (const [index, id] of expectedIds.entries()) expect(entries.find(entry => entry.id === String(id))).toMatchObject({
      partPath: 'xl/drawings/drawing1.xml', referenceId: references[index], status: 'malformed', reason: 'part-not-found', selectedRepresentation: 'none',
    })
    if (kinds[0] === 'fallback-pic') expect(entries.filter(entry => entry.id === '2' || entry.id === '4').every(entry => entry.representation === 'fallback')).toBe(true)
    if (kinds[0] === 'group') expect(entries.find(entry => entry.id === '999')).toMatchObject({ status: 'native' })
  })

  test('XLSX viewport limit records requested and retained pixels with worksheet provenance', async () => {
    const zip = await packageOf({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr><xdr:absoluteAnchor><xdr:pos x="${200000 * 9525}" y="${200000 * 9525}"/><xdr:ext cx="95250" cy="95250"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="11"/></xdr:nvSpPr><xdr:spPr>${xfrm}</xdr:spPr></xdr:sp><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
    })
    const doc = await parseXlsx(zip)
    const viewport = doc.drawingCoverage!.find(e => e.feature === 'viewport')
    expect(viewport).toMatchObject({ partPath: 'xl/worksheets/sheet1.xml', scope: 'diagnostic', limit: 16384,
      requestedExtent: { width: expect.any(Number), height: expect.any(Number) },
      retainedExtent: { width: expect.any(Number), height: expect.any(Number) } })
    expect(viewport!.requestedExtent!.width).toBeGreaterThan(viewport!.retainedExtent!.width)
  })

  test('XLSX supported empty Choice stays blank in coverage', async () => {
    const zip = await packageOf({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><mc:AlternateContent><mc:Choice Requires="a"/><mc:Fallback><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="12"/></xdr:nvSpPr><xdr:spPr>${xfrm}</xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
    })
    const doc = await parseXlsx(zip)
    expect(doc.sheets[0].drawings ?? []).toHaveLength(0)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ status: 'native', representation: 'choice', feature: 'empty-choice' })
  })

  test('XLSX alias and local shadow select exactly one compatibility object', async () => {
    const make = async (shadow: boolean) => {
      const selected = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="71"/></xdr:nvSpPr><xdr:spPr>${xfrm}</xdr:spPr></xdr:sp>`
      const fallback = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="72"/></xdr:nvSpPr><xdr:spPr>${xfrm}</xdr:spPr></xdr:sp>`
      const choice = `<mc:Choice ${shadow ? 'xmlns:alias="urn:unsupported" ' : ''}Requires="alias">${selected}</mc:Choice>`
      return parseXlsx(await packageOf({
        'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="s1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
        'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
        'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
        'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:alias="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><mc:AlternateContent>${choice}<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
      }))
    }
    const aliased = await make(false), shadowed = await make(true)
    expect(aliased.drawingCoverage![0]).toMatchObject({ id: '71', representation: 'choice' })
    expect(shadowed.drawingCoverage![0]).toMatchObject({ id: '72', representation: 'fallback' })
    expect(aliased.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(shadowed.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
  })

  test.each([
    ['missing picture', '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="501"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="absent"/></xdr:blipFill><xdr:spPr/></xdr:pic>', 'malformed'],
    ['unsupported graphicData', '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="502"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></xdr:xfrm><a:graphic><a:graphicData uri="urn:unsupported"/></a:graphic></xdr:graphicFrame>', 'unsupported'],
  ])('XLSX fallback %s preserves payload failure classification', async (_label, fallback, status) => {
    const zip = await packageOf({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><mc:AlternateContent><mc:Choice Requires="future">${shape(599)}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('absent', 'image', '../media/absent.png')),
    })
    const doc = await parseXlsx(zip)
    expect(doc.drawingCoverage!.filter(e => e.scope === 'original')).toHaveLength(1)
    expect(doc.drawingCoverage![0]).toMatchObject({ status, representation: 'fallback', selectedRepresentation: 'none' })
  })

  test('XLSX discards failed Choice diagnostics after selecting fallback and reports nested empty Choice', async () => {
    const frameWith = (inside: string) => `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="601"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></xdr:xfrm><a:graphic><a:graphicData uri="chart">${inside}</a:graphicData></a:graphic></xdr:graphicFrame>`
    const make = async (payload: string) => parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${payload}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('missing', 'chart', '../charts/absent.xml')),
    }))
    const outer = await make(`<mc:AlternateContent><mc:Choice Requires="c">${frameWith('<c:chart r:id="missing"/>')}</mc:Choice><mc:Fallback><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="602"/></xdr:nvSpPr><xdr:spPr>${xfrm}</xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent>`)
    expect(outer.drawingCoverage!.filter(e => e.scope === 'diagnostic')).toHaveLength(0)
    expect(outer.drawingCoverage![0]).toMatchObject({ id: '602', representation: 'fallback' })
    const blank = await make(frameWith('<mc:AlternateContent><mc:Choice Requires="c"/><mc:Fallback><c:chart r:id="missing"/></mc:Fallback></mc:AlternateContent>'))
    expect(blank.drawingCoverage![0]).toMatchObject({ id: '601', feature: 'empty-choice', status: 'native', representation: 'choice', selectedRepresentation: 'blank' })
  })

  test('XLSX outer Choice retains a chart beside its nested empty Choice', async () => {
    const chart = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    const inner = '<mc:AlternateContent><mc:Choice Requires="c"/><mc:Fallback><c:chart r:id="c2"/></mc:Fallback></mc:AlternateContent>'
    const frame = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="701"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></xdr:xfrm><a:graphic><a:graphicData uri="chart"><mc:AlternateContent><mc:Choice Requires="c"><c:chart r:id="c1"/>${inner}</mc:Choice><mc:Fallback><c:chart r:id="c2"/></mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></xdr:graphicFrame>`
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/sheet1.xml')),
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('dr', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${frame}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
      'xl/drawings/_rels/drawing1.xml.rels': rels(rel('c1', 'chart', '../charts/c1.xml') + rel('c2', 'chart', '../charts/c2.xml')),
      'xl/charts/c1.xml': chart, 'xl/charts/c2.xml': chart,
    }))
    expect(doc.sheets[0].drawings?.[0].content?.kind).toBe('chart')
    expect(doc.drawingCoverage![0]).toMatchObject({ id: '701', selectedRepresentation: 'column-chart', feature: 'chart', representation: 'choice' })
  })

  test('malformed owner relationships follow every selected Word picture and discard a failed picture Choice', async () => {
    const picture = (id: number, rid: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
    const xml = (body: string) => `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r>${body}</w:r></w:p></w:body></w:document>`
    const make = async (body: string) => parseDocx(await packageOf({ 'word/document.xml': xml(body), 'word/_rels/document.xml.rels': '<Relationships><' }))
    const selected = await make(picture(1, 'one') + picture(2, 'two'))
    expect(selected.drawingCoverage!.filter(e => ['1', '2'].includes(e.id ?? '')).map(e => [e.referenceId, e.status, e.reason, e.selectedRepresentation])).toEqual([
      ['one', 'malformed', 'invalid-relationship-xml', 'none'], ['two', 'malformed', 'invalid-relationship-xml', 'none'],
    ])
    expect(drawingReport(selected).counts.diagnosticEntries).toBe(0)
    const textbox = '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="8"/><a:graphic><a:graphicData><a:wsp><a:txbx><w:txbxContent><w:p><w:r><w:t>kept</w:t></w:r></w:p></w:txbxContent></a:txbx></a:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>'
    const fallback = `<mc:AlternateContent><mc:Choice Requires="pic">${picture(1, 'one')}</mc:Choice><mc:Fallback>${textbox}</mc:Fallback></mc:AlternateContent>`
    const recovered = await make(fallback)
    expect(recovered.drawingCoverage!.some(e => e.id === '1')).toBe(false)
    expect(recovered.drawingCoverage!.find(e => e.id === '8')?.status).toBe('fallback')
    expect(drawingReport(recovered).counts.diagnosticEntries).toBe(0)
  })

  test('malformed owner relationships follow selected XLSX pictures without surplus facets', async () => {
    const pic = (id: number, rid: string) => `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:absoluteAnchor>`
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/s1.xml')),
      'xl/worksheets/s1.xml': '<worksheet><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/s1.xml.rels': rels(rel('dr', 'drawing', '../drawings/a.xml')),
      'xl/drawings/a.xml': `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${pic(1, 'one')}${pic(2, 'two')}</xdr:wsDr>`,
      'xl/drawings/_rels/a.xml.rels': '<Relationships><',
    }))
    expect(doc.drawingCoverage!.filter(e => ['1', '2'].includes(e.id ?? '')).map(e => [e.referenceId, e.status, e.reason, e.selectedRepresentation])).toEqual([
      ['one', 'malformed', 'invalid-relationship-xml', 'none'], ['two', 'malformed', 'invalid-relationship-xml', 'none'],
    ])
    expect(drawingReport(doc).counts.diagnosticEntries).toBe(0)
  })

  test('PPTX retains concrete malformed picture owners and an actually unowned relationship facet', async () => {
    const p = (id: number, rid: string) => picture(id, rid)
    const parts = { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><' }
    const selected = await pptx(p(1, 'one') + p(2, 'two') + shape(9), parts)
    expect(selected.drawingCoverage!.filter(e => ['1', '2'].includes(e.id ?? '')).map(e => [e.referenceId, e.status, e.reason, e.selectedRepresentation])).toEqual([
      ['one', 'malformed', 'invalid-relationship-xml', 'none'], ['two', 'malformed', 'invalid-relationship-xml', 'none'],
    ])
    expect(drawingReport(selected).counts.diagnosticEntries).toBe(0)
    const imageFilledShape = shape(7).replace('</p:spPr>', '<a:blipFill><a:blip r:embed="fill"/></a:blipFill><a:ln w="28575"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></p:spPr>')
    const filled = await pptx(imageFilledShape + shape(9), parts)
    const canvas = createCanvas(120, 120)
    renderSlide(filled.slides[0], canvas.getContext('2d') as never)
    const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
    let greenPixels = 0
    for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) greenPixels++
    expect(greenPixels).toBeGreaterThan(0)
    expect(filled.drawingCoverage!.find(e => e.id === '7')).toMatchObject({ referenceId: 'fill', status: 'malformed', reason: 'invalid-relationship-xml', selectedRepresentation: 'native-shape' })
    expect(drawingReport(filled).counts.diagnosticEntries).toBe(0)
    const zeroWidth = await pptx(imageFilledShape.replace('cx="952500"', 'cx="0"'), parts)
    const lineCanvas = createCanvas(120, 120)
    renderSlide(zeroWidth.slides[0], lineCanvas.getContext('2d') as never)
    const linePixels = lineCanvas.getContext('2d').getImageData(0, 0, 120, 120).data
    let degenerateGreen = 0
    for (let i = 0; i < linePixels.length; i += 4) if (linePixels[i] < 20 && linePixels[i + 1] > 230 && linePixels[i + 2] < 20 && linePixels[i + 3] > 230) degenerateGreen++
    expect(degenerateGreen).toBeGreaterThan(0)
    expect(zeroWidth.drawingCoverage!.find(e => e.id === '7')).toMatchObject({ status: 'malformed', selectedRepresentation: 'native-shape' })
    const noStroke = await pptx(imageFilledShape.replace('w="28575"', 'w="0"'), parts)
    const blankCanvas = createCanvas(120, 120)
    renderSlide(noStroke.slides[0], blankCanvas.getContext('2d') as never)
    const blankPixels = blankCanvas.getContext('2d').getImageData(0, 0, 120, 120).data
    let visibleGreen = 0
    for (let i = 0; i < blankPixels.length; i += 4) if (blankPixels[i] < 20 && blankPixels[i + 1] > 230 && blankPixels[i + 2] < 20 && blankPixels[i + 3] > 230) visibleGreen++
    expect(visibleGreen).toBe(0)
    expect(noStroke.drawingCoverage!.find(e => e.id === '7')).toMatchObject({ status: 'malformed', selectedRepresentation: 'none' })
    const unowned = await pptx(shape(9), parts)
    expect(drawingReport(unowned).counts.diagnosticEntries).toBe(1)
    expect(unowned.drawingCoverage!.find(e => e.scope === 'diagnostic')?.reason).toBe('invalid-relationship-xml')
  })

  test.each([
    ['preset solid stroke', '<a:prstGeom prst="rect"/>', '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>', true],
    ['preset gradient stroke', '<a:prstGeom prst="rect"/>', '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>', true],
    ['preset gradient stroke zero width', '<a:prstGeom prst="rect"/>', '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>', true],
    ['preset radial stroke zero width', '<a:prstGeom prst="rect"/>', '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:path path="circle"/></a:gradFill>', false],
    ['preset no stroke', '<a:prstGeom prst="rect"/>', '<a:noFill/>', false],
    ['custom solid enabled', '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="none" stroke="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:lnTo><a:pt x="80" y="80"/></a:lnTo><a:lnTo><a:pt x="0" y="80"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>', '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>', true],
    ['custom solid disabled', '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="none" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:lnTo><a:pt x="80" y="80"/></a:lnTo><a:lnTo><a:pt x="0" y="80"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>', '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>', false],
    ['custom gradient disabled', '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="none" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:lnTo><a:pt x="80" y="80"/></a:lnTo><a:lnTo><a:pt x="0" y="80"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>', '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>', false],
  ] as const)('PPTX partial image-fill geometry follows actual paint: %s', async (_case, geometry, lineFill, painted) => {
    const body = `<p:sp xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:cNvPr id="7" name="partial"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="${_case.includes('zero width') ? 0 : 762000}" cy="762000"/></a:xfrm>${geometry}<a:blipFill><a:blip r:embed="missing"/></a:blipFill><a:ln w="28575">${lineFill}</a:ln></p:spPr></p:sp>`
    const doc = await pptx(body, { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><' })
    const canvas = createCanvas(120, 120)
    renderSlide(doc.slides[0], canvas.getContext('2d') as never)
    const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
    let green = 0
    for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
    expect(painted ? green > 0 : green === 0).toBe(true)
    expect(doc.drawingCoverage!.find(e => e.id === '7')).toMatchObject({ referenceId: 'missing', status: 'malformed', reason: 'invalid-relationship-xml', selectedRepresentation: painted ? 'native-shape' : 'none' })
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
  })

  const pathStroke = {
    solid: '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>',
    gradient: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>',
  }
  const pathCases = [
    ['move + close, butt', '<a:close/>', 'flat', false],
    ['zero line, butt', '<a:lnTo><a:pt x="0" y="0"/></a:lnTo><a:close/>', 'flat', false],
    ['move + close, square', '<a:close/>', 'sq', false],
    ['zero line, square', '<a:lnTo><a:pt x="0" y="0"/></a:lnTo><a:close/>', 'sq', false],
    ['move + close, round', '<a:close/>', 'rnd', true],
    ['zero line, round', '<a:lnTo><a:pt x="0" y="0"/></a:lnTo><a:close/>', 'rnd', true],
    ['nonzero line, butt', '<a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:close/>', 'flat', true],
    ['zero quadratic, butt', '<a:quadBezTo><a:pt x="0" y="0"/><a:pt x="0" y="0"/></a:quadBezTo><a:close/>', 'flat', false],
    ['zero quadratic, round', '<a:quadBezTo><a:pt x="0" y="0"/><a:pt x="0" y="0"/></a:quadBezTo><a:close/>', 'rnd', true],
    ['quadratic loop, butt', '<a:quadBezTo><a:pt x="40" y="80"/><a:pt x="0" y="0"/></a:quadBezTo><a:close/>', 'flat', true],
    ['zero cubic, butt', '<a:cubicBezTo><a:pt x="0" y="0"/><a:pt x="0" y="0"/><a:pt x="0" y="0"/></a:cubicBezTo><a:close/>', 'flat', false],
    ['cubic loop, butt', '<a:cubicBezTo><a:pt x="0" y="80"/><a:pt x="80" y="80"/><a:pt x="0" y="0"/></a:cubicBezTo><a:close/>', 'flat', true],
    ['arc, butt', '<a:moveTo><a:pt x="40" y="40"/></a:moveTo><a:arcTo wR="20" hR="20" stAng="0" swAng="5400000"/>', 'flat', true],
  ] as const
  for (const [styleName, lineFill] of Object.entries(pathStroke)) {
    test.each(pathCases)(`PPTX retained path semantics ${styleName}: %s`, async (_case, commands, cap, painted) => {
      const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="none" stroke="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${commands}</a:path></a:pathLst></a:custGeom>`
      const body = `<p:sp xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:cNvPr id="7" name="path"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}<a:blipFill><a:blip r:embed="missing"/></a:blipFill><a:ln w="28575" cap="${cap}">${lineFill}</a:ln></p:spPr></p:sp>`
      const doc = await pptx(body, { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><' })
      const canvas = createCanvas(120, 120)
      const ctx = canvas.getContext('2d')
      ctx.fillStyle = '#fff'
      ctx.fillRect(0, 0, 120, 120)
      renderSlide(doc.slides[0], ctx as never)
      const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
      let green = 0, nonwhite = 0
      for (let i = 0; i < pixels.length; i += 4) {
        if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
        if (pixels[i] !== 255 || pixels[i + 1] !== 255 || pixels[i + 2] !== 255) nonwhite++
      }
      expect(painted ? green > 0 && nonwhite > 0 : green === 0 && nonwhite === 0).toBe(true)
      expect(doc.drawingCoverage!.find(e => e.id === '7')).toMatchObject({ referenceId: 'missing', status: 'malformed', reason: 'invalid-relationship-xml', selectedRepresentation: painted ? 'native-shape' : 'none' })
      expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
    })
  }

  for (const [styleName, lineFill] of Object.entries({ ...pathStroke, none: '<a:noFill/>' })) {
    test.each([
      ['flat', false], ['sq', styleName !== 'none'], ['rnd', styleName !== 'none'],
    ] as const)(`PPTX zero-length dash ${styleName} with %s cap`, async (cap, painted) => {
      const geometry = '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="none" stroke="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>'
      const body = `<p:sp xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:cNvPr id="7" name="dashed"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}<a:blipFill><a:blip r:embed="missing"/></a:blipFill><a:ln w="28575" cap="${cap}">${lineFill}<a:custDash><a:ds d="0" sp="100000"/></a:custDash></a:ln></p:spPr></p:sp>`
      const doc = await pptx(body, { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><' })
      const canvas = createCanvas(120, 120)
      renderSlide(doc.slides[0], canvas.getContext('2d') as never)
      const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
      let green = 0
      for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
      expect(painted ? green > 0 : green === 0).toBe(true)
      expect(doc.drawingCoverage!.find(e => e.id === '7')).toMatchObject({ referenceId: 'missing', status: 'malformed', reason: 'invalid-relationship-xml', selectedRepresentation: painted ? 'native-shape' : 'none' })
      expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
    })
  }

  for (const [styleName, lineFill] of Object.entries(pathStroke)) {
    test.each([
      ['no arrow', '<a:lnTo><a:pt x="60" y="40"/></a:lnTo>', '', false],
      ['tail arrow', '<a:lnTo><a:pt x="60" y="40"/></a:lnTo>', '<a:tailEnd type="triangle"/>', true],
      ['head arrow', '<a:lnTo><a:pt x="60" y="40"/></a:lnTo>', '<a:headEnd type="triangle"/>', true],
      ['zero tangent arrow', '<a:lnTo><a:pt x="20" y="40"/></a:lnTo><a:close/>', '<a:tailEnd type="triangle"/>', false],
    ] as const)(`PPTX zero-dash arrow ${styleName}: %s`, async (_case, commands, arrow, painted) => {
      const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="none" stroke="1"><a:moveTo><a:pt x="20" y="40"/></a:moveTo>${commands}</a:path></a:pathLst></a:custGeom>`
      const body = `<p:sp xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:cNvPr id="7" name="arrow"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}<a:blipFill><a:blip r:embed="missing"/></a:blipFill><a:ln w="28575" cap="flat">${lineFill}<a:custDash><a:ds d="0" sp="100000"/></a:custDash>${arrow}</a:ln></p:spPr></p:sp>`
      const doc = await pptx(body, { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><' })
      const canvas = createCanvas(120, 120)
      renderSlide(doc.slides[0], canvas.getContext('2d') as never)
      const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
      let green = 0
      for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
      expect(painted ? green > 0 : green === 0).toBe(true)
      expect(doc.drawingCoverage!.find(e => e.id === '7')).toMatchObject({ referenceId: 'missing', status: 'malformed', reason: 'invalid-relationship-xml', selectedRepresentation: painted ? 'native-shape' : 'none' })
    })
  }

  for (const [styleName, lineFill] of Object.entries({ ...pathStroke, none: '<a:noFill/>' })) {
    test.each([
      ['long', 'flat', '', false], ['short', 'flat', '', styleName !== 'none'],
      ['long', 'sq', '', styleName !== 'none'], ['long', 'rnd', '', styleName !== 'none'],
      ['long', 'flat', '<a:headEnd type="triangle"/>', styleName !== 'none'],
      ['long', 'flat', '<a:tailEnd type="triangle"/>', styleName !== 'none'],
    ] as const)(`PPTX delayed dash ${styleName}: %s gap %s cap %s arrow`, async (gap, cap, arrow, painted) => {
      const geometry = '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="none" stroke="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>'
      const dash = `<a:custDash><a:ds d="0" sp="${gap === 'long' ? '10000000' : '100000'}"/><a:ds d="100000" sp="100000"/></a:custDash>`
      const body = `<p:sp xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:cNvPr id="7" name="delayed dash"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}<a:blipFill><a:blip r:embed="missing"/></a:blipFill><a:ln w="28575" cap="${cap}">${lineFill}${dash}${arrow}</a:ln></p:spPr></p:sp>`
      const doc = await pptx(body, { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><' })
      const canvas = createCanvas(120, 120)
      const ctx = canvas.getContext('2d')
      ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 120, 120)
      renderSlide(doc.slides[0], ctx as never)
      const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
      let green = 0, nonwhite = 0
      for (let i = 0; i < pixels.length; i += 4) {
        if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
        if (pixels[i] !== 255 || pixels[i + 1] !== 255 || pixels[i + 2] !== 255) nonwhite++
      }
      expect(painted ? green > 0 && nonwhite > 0 : green === 0 && nonwhite === 0).toBe(true)
      const entry = doc.drawingCoverage!.find(e => e.id === '7')!
      expect(entry).toMatchObject({ referenceId: 'missing', status: 'malformed', reason: 'invalid-relationship-xml', selectedRepresentation: gap === 'long' && cap === 'flat' && !arrow && styleName !== 'none' ? 'unverified' : painted ? 'native-shape' : 'none' })
      if (gap === 'long' && cap === 'flat' && !arrow && styleName !== 'none') expect(entry.assessmentReason).toBe('dash-paint-inconclusive')
      expect(drawingReport(doc).counts).toEqual({ originalObjects: 1, selectedDescendants: 0, diagnosticEntries: 0 })
    })
  }

  const fillPoint = (x: number, y: number) => `<a:lnTo><a:pt x="${x}" y="${y}"/></a:lnTo>`
  const fillMove = (x: number, y: number) => `<a:moveTo><a:pt x="${x}" y="${y}"/></a:moveTo>`
  const fillCases = [
    ['retrace angle', fillPoint(80, 0) + fillPoint(80, 80) + fillPoint(80, 0) + '<a:close/>', false],
    ['opposite triangle', fillPoint(80, 0) + fillPoint(80, 80) + fillPoint(0, 0) + fillPoint(80, 80) + fillPoint(80, 0) + '<a:close/>', false],
    ['opposite subpaths', fillPoint(80, 0) + fillPoint(80, 80) + '<a:close/>' + fillMove(0, 0) + fillPoint(80, 80) + fillPoint(80, 0) + '<a:close/>', false],
    ['same winding subpaths', fillPoint(80, 0) + fillPoint(80, 80) + '<a:close/>' + fillMove(0, 0) + fillPoint(80, 0) + fillPoint(80, 80) + '<a:close/>', true],
    ['bowtie', fillPoint(80, 80) + fillPoint(0, 80) + fillPoint(80, 0) + '<a:close/>', true],
    ['quadratic retrace', '<a:quadBezTo><a:pt x="40" y="80"/><a:pt x="80" y="0"/></a:quadBezTo><a:quadBezTo><a:pt x="40" y="80"/><a:pt x="0" y="0"/></a:quadBezTo><a:close/>', true],
  ] as const
  for (const [styleName, fillMarkup] of Object.entries(pathStroke)) {
    test.each(fillCases)(`PPTX completed fill winding ${styleName}: %s`, async (_case, commands, painted) => {
      const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="norm" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${commands}</a:path></a:pathLst></a:custGeom>`
      const body = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="filled"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}${fillMarkup}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
      const doc = await pptx(body)
      const canvas = createCanvas(120, 120)
      renderSlide(doc.slides[0], canvas.getContext('2d') as never)
      const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
      let green = 0
      for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
      expect(painted ? green > 0 : green === 0).toBe(true)
      const source = doc.slides[0].shapes[0]
      source.diagnostics!.push({ kind: 'missing-image', message: 'Simulated unavailable referenced image', source: source.source })
      const tree = parseXmlOrdered(`<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</p:spTree>`)
      expect(pptxCoverage(doc.slides[0], tree, 'ppt/slides/slide1.xml')[0]).toMatchObject({ status: 'malformed', selectedRepresentation: _case === 'quadratic retrace' ? 'unverified' : painted ? 'native-shape' : 'none', ...(_case === 'quadratic retrace' ? { assessmentReason: 'curve-paint-inconclusive' } : {}) })
    })
  }

  for (const [styleName, fillMarkup] of Object.entries(pathStroke)) {
    test.each([['opposite contours', false], ['same winding contours', true]] as const)(`PPTX long completed fill ${styleName}: %s`, async (_case, painted) => {
      const clockwise = fillMove(0, 0) + fillPoint(80, 0) + fillPoint(80, 80) + '<a:close/>'
      const reverse = fillMove(0, 0) + fillPoint(80, 80) + fillPoint(80, 0) + '<a:close/>'
      const commands = Array.from({ length: 12 }, () => clockwise + (painted ? clockwise : reverse)).join('')
      const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="norm" stroke="0">${commands}</a:path></a:pathLst></a:custGeom>`
      const body = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="long fill"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}${fillMarkup}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
      const doc = await pptx(body)
      const canvas = createCanvas(120, 120)
      renderSlide(doc.slides[0], canvas.getContext('2d') as never)
      const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
      let green = 0
      for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
      expect(painted ? green > 0 : green === 0).toBe(true)
      const source = doc.slides[0].shapes[0]
      source.diagnostics!.push({ kind: 'missing-image', message: 'Simulated unavailable referenced image', source: source.source })
      const tree = parseXmlOrdered(`<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</p:spTree>`)
      expect(pptxCoverage(doc.slides[0], tree, 'ppt/slides/slide1.xml')[0]).toMatchObject({ status: 'malformed', selectedRepresentation: painted ? 'native-shape' : 'none' })
    })
  }

  for (const [styleName, fillMarkup] of Object.entries(pathStroke)) {
    test.each([[24, false], [24, true], [32, false], [32, true], [40, false], [40, true]] as const)(`PPTX split winding ${styleName} at %i vertices, same direction=%s`, async (vertices, sameDirection) => {
      const forward = Array.from({ length: vertices }, (_, i) => [i * 2, i === 0 ? 0 : i % 2 ? 40 : 80] as const)
      const draw = (points: readonly (readonly [number, number])[]) => fillMove(...points[0]) + points.slice(1).map(([x, y]) => fillPoint(x, y)).join('') + '<a:close/>'
      const reverse = [...forward].reverse().flatMap((point, i, all) => i + 1 < all.length
        ? [point, [(point[0] + all[i + 1][0]) / 2, (point[1] + all[i + 1][1]) / 2] as const] : [point])
      const commands = draw(forward) + draw(sameDirection ? forward : reverse)
      const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="norm" stroke="0">${commands}</a:path></a:pathLst></a:custGeom>`
      const body = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="split winding"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}${fillMarkup}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
      const doc = await pptx(body)
      const canvas = createCanvas(120, 120)
      renderSlide(doc.slides[0], canvas.getContext('2d') as never)
      const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
      let green = 0
      for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
      expect(sameDirection ? green > 0 : green === 0).toBe(true)
      const source = doc.slides[0].shapes[0]
      source.diagnostics!.push({ kind: 'missing-image', message: 'Simulated unavailable referenced image', source: source.source })
      const tree = parseXmlOrdered(`<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</p:spTree>`)
      expect(pptxCoverage(doc.slides[0], tree, 'ppt/slides/slide1.xml')[0]).toMatchObject(sameDirection
        ? { status: 'malformed', reason: 'missing image part or relationship', selectedRepresentation: 'native-shape' }
        : { status: 'malformed', reason: 'missing image part or relationship', selectedRepresentation: 'unverified', assessmentReason: 'winding-analysis-limit', analysisLimit: 64 })
    })
  }

  for (const [styleName, fillMarkup] of Object.entries(pathStroke)) {
    test.each([
      ['quadratic', '<a:quadBezTo><a:pt x="0" y="0"/><a:pt x="0" y="0"/></a:quadBezTo>'],
      ['cubic', '<a:cubicBezTo><a:pt x="0" y="0"/><a:pt x="0" y="0"/><a:pt x="0" y="0"/></a:cubicBezTo>'],
      ['arc', '<a:arcTo wR="0" hR="0" stAng="0" swAng="0"/>'],
    ] as const)(`PPTX canceled fill plus no-op ${styleName} %s remains none`, async (_case, noop) => {
      const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="norm" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${fillPoint(80, 0)}${fillPoint(80, 80)}${fillPoint(80, 0)}<a:close/>${noop}</a:path></a:pathLst></a:custGeom>`
      const body = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="no-op curve"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}${fillMarkup}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
      const doc = await pptx(body)
      const canvas = createCanvas(120, 120)
      renderSlide(doc.slides[0], canvas.getContext('2d') as never)
      const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
      let green = 0
      for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
      expect(green).toBe(0)
      const source = doc.slides[0].shapes[0]
      source.diagnostics!.push({ kind: 'missing-image', message: 'Simulated unavailable referenced image', source: source.source })
      const tree = parseXmlOrdered(`<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</p:spTree>`)
      expect(pptxCoverage(doc.slides[0], tree, 'ppt/slides/slide1.xml')[0]).toMatchObject({ status: 'malformed', selectedRepresentation: 'none' })
    })
  }

  test.each([false, true])('PPTX proven second path dominates inconclusive fill, proven first=%s', async provenFirst => {
    const unknown = '<a:path w="80" h="80" fill="norm" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:quadBezTo><a:pt x="40" y="80"/><a:pt x="80" y="0"/></a:quadBezTo><a:quadBezTo><a:pt x="40" y="80"/><a:pt x="0" y="0"/></a:quadBezTo><a:close/></a:path>'
    const proven = '<a:path w="80" h="80" fill="norm" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:lnTo><a:pt x="0" y="80"/></a:lnTo><a:close/></a:path>'
    const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${provenFirst ? proven + unknown : unknown + proven}</a:pathLst></a:custGeom>`
    const body = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="aggregate"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
    const doc = await pptx(body)
    const canvas = createCanvas(120, 120)
    renderSlide(doc.slides[0], canvas.getContext('2d') as never)
    const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
    let green = 0
    for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
    expect(green).toBeGreaterThan(0)
    const source = doc.slides[0].shapes[0]
    source.diagnostics!.push({ kind: 'missing-image', message: 'Simulated unavailable referenced image', source: source.source })
    const tree = parseXmlOrdered(`<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</p:spTree>`)
    const entry = pptxCoverage(doc.slides[0], tree, 'ppt/slides/slide1.xml')[0]
    expect(entry).toMatchObject({ status: 'malformed', selectedRepresentation: 'native-shape' })
    expect(entry.assessmentReason).toBeUndefined()
  })

  test.each([
    ['move + close', '<a:close/>', false],
    ['zero line', '<a:lnTo><a:pt x="0" y="0"/></a:lnTo><a:close/>', false],
    ['single line', '<a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:close/>', false],
    ['triangle', '<a:lnTo><a:pt x="80" y="0"/></a:lnTo><a:lnTo><a:pt x="0" y="80"/></a:lnTo><a:close/>', true],
    ['quadratic bulge', '<a:quadBezTo><a:pt x="40" y="80"/><a:pt x="80" y="0"/></a:quadBezTo><a:close/>', true],
    ['cubic bulge', '<a:cubicBezTo><a:pt x="0" y="80"/><a:pt x="80" y="80"/><a:pt x="80" y="0"/></a:cubicBezTo><a:close/>', true],
  ] as const)('PPTX retained fill path has actual area: %s', async (_case, commands, painted) => {
    const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="80" h="80" fill="norm" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${commands}</a:path></a:pathLst></a:custGeom>`
    const body = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="fill path"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="762000" cy="762000"/></a:xfrm>${geometry}<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
    const doc = await pptx(body)
    const canvas = createCanvas(120, 120)
    renderSlide(doc.slides[0], canvas.getContext('2d') as never)
    const pixels = canvas.getContext('2d').getImageData(0, 0, 120, 120).data
    let green = 0
    for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] > 230 && pixels[i + 2] < 20 && pixels[i + 3] > 230) green++
    expect(painted ? green > 0 : green === 0).toBe(true)
    const shape = doc.slides[0].shapes[0]
    shape.diagnostics!.push({ kind: 'missing-image', message: 'Simulated unavailable referenced image', source: shape.source })
    const tree = parseXmlOrdered(`<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</p:spTree>`)
    expect(pptxCoverage(doc.slides[0], tree, 'ppt/slides/slide1.xml')[0]).toMatchObject({ status: 'malformed', selectedRepresentation: _case === 'quadratic bulge' || _case === 'cubic bulge' ? 'unverified' : painted ? 'native-shape' : 'none', ...(_case === 'quadratic bulge' || _case === 'cubic bulge' ? { assessmentReason: 'curve-paint-inconclusive' } : {}) })
  })

  test.each([false, true])('cached malformed wrapper survives discarded XLSX Choice, reversed=%s', async reversed => {
    const anchor = (body: string) => `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${body}<xdr:clientData/></xdr:absoluteAnchor>`
    const native = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:sp>'
    const recovered = `<mc:AlternateContent><mc:Choice Requires="xdr"><xdr:contentPart r:id="bad1"/></mc:Choice><mc:Fallback>${native}</mc:Fallback></mc:AlternateContent>`
    const selected = '<xdr:contentPart r:id="bad2"/>'
    const bodies = reversed ? [selected, recovered] : [recovered, selected]
    const doc = await parseXlsx(await packageOf({
      'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': rels(rel('s1', 'worksheet', 'worksheets/s1.xml')),
      'xl/worksheets/s1.xml': '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="dr"/></worksheet>',
      'xl/worksheets/_rels/s1.xml.rels': rels(rel('dr', 'drawing', '../drawings/a.xml')),
      'xl/drawings/a.xml': `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${bodies.map(anchor).join('')}</xdr:wsDr>`,
      'xl/drawings/_rels/a.xml.rels': rels(rel('bad1', 'customXml', '../ink/wrapper.xml') + rel('bad2', 'customXml', '../ink/wrapper.xml')),
      'xl/ink/wrapper.xml': '<root xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><contentPart r:id="inner"/></root>',
      'xl/ink/_rels/wrapper.xml.rels': '<Relationships><',
    }))
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 2, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage!.find(e => e.referenceId === 'bad2')).toMatchObject({ status: 'malformed', reason: 'invalid-relationship-xml', selectedRepresentation: 'none' })
    expect(doc.drawingCoverage!.find(e => e.id === '9')).toMatchObject({ status: 'fallback', selectedRepresentation: 'native-shape' })
  })

  test.each(['picture', 'chart'] as const)('reused PPTX slide part keeps concrete %s failure on both units', async kind => {
    const content = kind === 'picture' ? picture(1, 'one') : frame(1, 'chart', '<c:chart r:id="one"/>')
    const doc = await parsePptx(await packageOf({
      'ppt/presentation.xml': '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="s0"/><p:sldId id="257" r:id="s1"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s0', 'slide', 'slides/slide0.xml') + rel('s1', 'slide', 'slides/slide0.xml')),
      'ppt/slides/slide0.xml': `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>${content}${shape(9)}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide0.xml.rels': '<Relationships><',
    }))
    expect(drawingReport(doc).counts).toEqual({ originalObjects: 4, selectedDescendants: 0, diagnosticEntries: 0 })
    expect(doc.drawingCoverage!.filter(e => e.referenceId === 'one').map(e => [e.unit, e.status, e.reason, e.selectedRepresentation])).toEqual([
      [0, 'malformed', 'invalid-relationship-xml', 'none'], [1, 'malformed', 'invalid-relationship-xml', 'none'],
    ])
    expect(doc.drawingCoverage!.filter(e => e.id === '9').map(e => e.status)).toEqual(['native', 'native'])
  })

  test('CLI emits one JSON result per path with explicit format and errors', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'officeview-drawing-report-'))
    try {
      const zip = new JSZip()
      zip.file('ppt/presentation.xml', '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>')
      zip.file('ppt/_rels/presentation.xml.rels', rels(rel('s1', 'slide', 'slides/slide1.xml')))
      zip.file('ppt/slides/slide1.xml', `<p:sld><p:cSld><p:spTree>${shape(1)}</p:spTree></p:cSld></p:sld>`)
      const good = join(dir, 'good.pptx'), missing = join(dir, 'missing.docx')
      writeFileSync(good, Buffer.from(await zip.generateAsync({ type: 'uint8array' })))
      const run = spawnSync('bun', ['scripts/drawing-report.ts', good, missing], { cwd: join(import.meta.dirname, '..'), encoding: 'utf8' })
      const report = JSON.parse(run.stdout)
      expect(run.status).toBe(1)
      expect(report.files).toHaveLength(2)
      expect(report.files[0]).toMatchObject({ format: 'pptx', unit: 'slide', unitCount: 1, counts: { originalObjects: 1 } })
      expect(report.files[1]).toMatchObject({ path: missing, error: expect.any(String) })
    } finally { rmSync(dir, { recursive: true, force: true }) }
  })
})
