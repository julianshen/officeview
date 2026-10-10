import { describe, expect, test } from 'vitest'
import { spawnSync } from 'node:child_process'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { prepareDrawingContent, type DrawingContentShape } from '../src/drawing/content'
import { drawingPartContext } from '../src/drawing/parts'

const drawingPath = 'word/diagrams/drawing.xml'
const theme = { colors: new Map<string, string>(), fonts: new Map<string, string>() }
const drawingTheme = { colors: {}, palette: {}, colorMap: {}, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
const graphic = { contentPart: { '@attrs': { id: 'cache' } } }
const shape = '<dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"/></dsp:spPr></dsp:sp>'
const picture = (id: string) => `<dsp:pic><dsp:blipFill><a:blip r:embed="${id}"/></dsp:blipFill><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></dsp:spPr></dsp:pic>`
const group = (tree: string) => `<dsp:grpSp><dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></dsp:grpSpPr>${tree}</dsp:grpSp>`
const cell = (attributes = '') => `<a:tc ${attributes}><a:tcPr/><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody></a:tc>`
const table = (rows: string, cols = [100, 200]) => `<dsp:graphicFrame><dsp:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="100"/></dsp:xfrm><a:graphic><a:graphicData><a:tbl><a:tblGrid>${cols.map(w => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></dsp:graphicFrame>`
const row = (cells: string, h = 40) => `<a:tr h="${h}">${cells}</a:tr>`

function parts(tree: string): Record<string, string> {
  return {
    'word/_rels/document.xml.rels': '<Relationships><Relationship Id="cache" Target="diagrams/drawing.xml"/></Relationships>',
    [drawingPath]: `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dsp:spTree>${tree}</dsp:spTree></dsp:drawing>`,
    'word/diagrams/_rels/drawing.xml.rels': '<Relationships><Relationship Id="img1" Target="../media/one.png"/><Relationship Id="img2" Target="../media/two.png"/></Relationships>',
    'word/media/one.png': 'first picture',
    'word/media/two.png': 'second picture',
  }
}

async function load(tree: string) {
  const zip = new JSZip()
  for (const [path, data] of Object.entries(parts(tree))) zip.file(path, data)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  const content = await prepareDrawingContent(pkg, graphic, 'word/document.xml', theme, drawingTheme)
  expect(content?.kind).toBe('diagram')
  if (content?.kind !== 'diagram') throw new Error('Expected cached diagram')
  return { shapes: content.shapes, diagnostics: drawingPartContext(pkg).diagnostics }
}

function images(shapes: DrawingContentShape[]): string[] {
  return shapes.flatMap(s => [...(s.image ? [new TextDecoder().decode(s.image.data)] : []), ...images(s.children ?? [])])
}

describe('cached content bounded materialization', () => {
  test.each([
    ['shape first', shape + picture('img1') + picture('img2')],
    ['picture first', picture('img1') + shape + picture('img2')],
    ['table first', table(row(cell())) + picture('img1') + picture('img2')],
    ['mixed groups', shape + group(shape + picture('img1') + group(picture('img2') + shape))],
  ])('preloads every referenced picture: %s', async (_name, tree) => {
    const result = await load(tree)
    expect(images(result.shapes)).toEqual(['first picture', 'second picture'])
    expect(result.diagnostics.some(d => d.reason === 'cached-picture-unsupported')).toBe(false)
  })

  test.each(['gridSpan', 'rowSpan'])('extreme finite %s completes in a bounded subprocess', span => {
    // A synchronous infinite/huge span loop cannot be stopped by Vitest's
    // timeout; isolate this hostile payload so regressions fail safely.
    const source = `
      import JSZip from ${JSON.stringify(resolve('node_modules/jszip/lib/index.js'))};
      import { OfficePackage } from ${JSON.stringify(resolve('src/core/zip.ts'))};
      import { prepareDrawingContent } from ${JSON.stringify(resolve('src/drawing/content.ts'))};
      import { drawingPartContext } from ${JSON.stringify(resolve('src/drawing/parts.ts'))};
      const zip = new JSZip();
      for (const [path, data] of Object.entries(${JSON.stringify(parts(table(row(cell(`${span}="1000000000000"`)))))})) zip.file(path, data);
      const pkg = await OfficePackage.load(await zip.generateAsync({type:'uint8array'}));
      const content = await prepareDrawingContent(pkg, ${JSON.stringify(graphic)}, 'word/document.xml', {colors:new Map(),fonts:new Map()}, ${JSON.stringify(drawingTheme)});
      console.log(JSON.stringify({shapes:content.shapes,diagnostics:drawingPartContext(pkg).diagnostics}));
    `
    const result = spawnSync('bun', ['--eval', source], { encoding: 'utf8', timeout: 2000 })
    expect(result.error, 'span expansion must not time out').toBeUndefined()
    expect(result.status, result.stderr).toBe(0)
    const parsed = JSON.parse(result.stdout)
    expect(parsed.shapes[0].widthEmu).toBe(100)
    expect(parsed.shapes[0].heightEmu).toBe(40)
    expect(parsed.diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', feature: span, reason: 'invalid-table-span', identity: '1000000000000' }))
  })

  test.each(['0', '-1', '1.5', 'NaN', 'Infinity', '3', ''])('diagnoses invalid spans %j and retains cell text', async value => {
    const { shapes, diagnostics } = await load(table(row(cell(`gridSpan="${value}" rowSpan="${value}"`))))
    expect(shapes[0].widthEmu).toBe(100)
    expect(shapes[0].heightEmu).toBe(40)
    expect(shapes[0].paragraphs[0].runs[0].text).toBe('Cell')
    for (const feature of ['gridSpan', 'rowSpan']) {
      expect(diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', partPath: drawingPath, feature, reason: 'invalid-table-span', identity: value }))
    }
  })

  test('valid merged table spans respect remaining grid and row extents', async () => {
    const { shapes, diagnostics } = await load(table(row(cell('gridSpan="2" rowSpan="2"')) + row(cell() + cell(), 60)))
    expect(shapes[0]).toMatchObject({ xEmu: 10, yEmu: 20, widthEmu: 300, heightEmu: 100 })
    expect(shapes[1]).toMatchObject({ xEmu: 10, yEmu: 60, widthEmu: 100, heightEmu: 60 })
    expect(shapes[2]).toMatchObject({ xEmu: 110, yEmu: 60, widthEmu: 200, heightEmu: 60 })
    expect(diagnostics).toEqual([])
  })

  test('valid physical merge placeholders do not paint or shift later cells', async () => {
    const { shapes, diagnostics } = await load(table(
      row(cell('gridSpan="2" rowSpan="2"') + cell('hMerge="1" rowSpan="2"') + cell()) +
      row(cell('vMerge="1" gridSpan="2"') + cell('hMerge="1" vMerge="1"') + cell(), 60),
      [100, 200, 300],
    ))
    expect(shapes).toHaveLength(3)
    expect(shapes[0]).toMatchObject({ xEmu: 10, yEmu: 20, widthEmu: 300, heightEmu: 100 })
    expect(shapes[1]).toMatchObject({ xEmu: 310, yEmu: 20, widthEmu: 300, heightEmu: 40 })
    expect(shapes[2]).toMatchObject({ xEmu: 310, yEmu: 60, widthEmu: 300, heightEmu: 60 })
    expect(diagnostics).toEqual([])
  })

  test('contradictory physical spans retain every cell without overlapping siblings', async () => {
    const { shapes, diagnostics } = await load(table(
      row(cell('gridSpan="2"') + cell() + cell()), [100, 100, 100],
    ))
    expect(shapes.map(s => [s.xEmu, s.widthEmu])).toEqual([[10, 100], [110, 100], [210, 100]])
    expect(shapes.every(s => s.paragraphs[0].runs[0].text === 'Cell')).toBe(true)
    expect(diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', reason: 'missing-merge-placeholder', feature: 'gridSpan' }))
  })

  test('bounds spans at the current column and row, preserving later siblings', async () => {
    const { shapes, diagnostics } = await load(table(row(cell() + cell('gridSpan="2"')) + row(cell('rowSpan="2"'), 60)) + shape)
    expect(shapes).toHaveLength(4)
    expect(shapes[1].widthEmu).toBe(200)
    expect(shapes[2].heightEmu).toBe(60)
    expect(diagnostics.filter(d => d.reason === 'invalid-table-span')).toHaveLength(2)
    expect(diagnostics.every(d => d.limit === 1)).toBe(true)
  })
})
