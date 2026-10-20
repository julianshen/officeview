/**
 * Task A7 (spec 6.6): accessible chart descriptions in the paged viewer. Each page's charts travel with the
 * paintables (description plus placement, charged to the document's chart workspace) and the lazy chrome exposes
 * them as an offscreen per-page list of role="img" elements in reading order. Also the A4d carry-over: a model
 * transferred from a worker realm shows its user-shape pictures at the first paint. Fixtures are Office-shaped
 * synthetic, never Office evidence.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { render, waitFor, within } from '@testing-library/react'
import { createCanvas } from 'canvas'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { getPaintables } from '../src/render/paint'
import { chartDescription } from '../src/charts/describe'
import { cardChart } from '../src/drawing/content'
import { buildChartDocx, buildChartPptx, officeChartWorkbook, officeChartXml } from './helpers/chart-fixtures'

const titled = (title: string, kind: 'bar' | 'line' = 'bar') => {
  let xml = officeChartXml().replace('<c:autoTitleDeleted val="1"/>', `<c:title><c:tx><c:rich><a:p><a:r><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`)
  if (kind === 'line') xml = xml.replace('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>', '<c:lineChart><c:grouping val="standard"/>').replace('<c:gapWidth val="219"/><c:overlap val="-27"/>', '').replace('</c:barChart>', '</c:lineChart>')
  return xml
}
const SURFACE = officeChartXml().replace(/<c:barChart>.*<\/c:barChart>/, '<c:surface3DChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:surface3DChart>')
const lists = (container: HTMLElement) => [...container.querySelectorAll<HTMLElement>('[data-testid="officeview-page-charts"] ul')]

afterEach(() => { vi.restoreAllMocks() })

describe('accessible chart descriptions (spec 6.6)', () => {
  it('two charts on one page expose two role=img entries in reading order', async () => {
    const docx = await parseDocx(await OfficePackage.load(buildChartDocx([
      { chartXml: titled('Revenue'), workbook: officeChartWorkbook([4.3, 2.5]) },
      { chartXml: titled('Costs', 'line'), workbook: officeChartWorkbook([1, 2]) },
    ]).bytes))
    const pages = await getPaintables(docx)
    expect(pages[0].charts?.map(c => c.description.split(',')[0])).toEqual(['Chart: Revenue. Column chart', 'Chart: Costs. Line chart'])
    // The placement is the chart's box on the page, the first above the second.
    const [first, second] = pages[0].charts!
    expect(first.width).toBeCloseTo(5486400 / 9525, 0)
    expect(second.y).toBeGreaterThan(first.y + first.height - 1)
    pages.dispose()

    const view = render(<OfficeDoc document={docx} />)
    await waitFor(() => expect(view.container.querySelectorAll('[role="img"]').length).toBe(2))
    const [list] = lists(view.container)
    expect(list.getAttribute('aria-label')).toBe('Charts on page 1 (2)')
    expect(within(list).getAllByRole('img').map(el => el.getAttribute('aria-label')!.split(',')[0])).toEqual(['Chart: Revenue. Column chart', 'Chart: Costs. Line chart'])
    view.unmount()

    // PPTX: slide z-order.
    const pptx = await parsePptx(await OfficePackage.load(buildChartPptx([
      { chartXml: titled('Back'), workbook: officeChartWorkbook([1, 2]) },
      { chartXml: titled('Front', 'line'), workbook: officeChartWorkbook([3, 4]) },
    ]).bytes))
    const slides = render(<OfficeDoc document={pptx} />)
    await waitFor(() => expect(slides.container.querySelectorAll('[role="img"]').length).toBe(2))
    expect(lists(slides.container)[0].getAttribute('aria-label')).toBe('Charts on slide 1 (2)')
    expect(within(lists(slides.container)[0]).getAllByRole('img').map(el => el.getAttribute('aria-label')!.split('.')[0])).toEqual(['Chart: Back', 'Chart: Front'])
    slides.unmount()
  })

  it('a header chart repeated on each page appears once per page', async () => {
    const docx = await parseDocx(await OfficePackage.load(buildChartDocx([
      { chartXml: titled('Body'), workbook: officeChartWorkbook([4.3, 2.5]) },
      { chartXml: titled('Masthead', 'line'), workbook: officeChartWorkbook([1, 2]) },
    ], { headerCharts: 1, paragraphs: 120 }).bytes))
    const pages = await getPaintables(docx)
    expect(pages.length).toBeGreaterThan(2)
    for (const [index, page] of pages.entries()) {
      const names = page.charts?.map(c => c.description.split('.')[0]) ?? []
      // Reading order: the header first, then the body chart (page 1 only).
      expect(names, `page ${index + 1}`).toEqual(index === 0 ? ['Chart: Masthead', 'Chart: Body'] : ['Chart: Masthead'])
      expect(page.charts![0].y, `page ${index + 1}: in the header band`).toBeLessThan(page.spec.heightPx / 4)
    }
    pages.dispose()
    const view = render(<OfficeDoc document={docx} />)
    await waitFor(() => expect(lists(view.container).length).toBe(pages.length))
    for (const [index, list] of lists(view.container).entries()) expect(within(list).getAllByRole('img').filter(el => el.getAttribute('aria-label')!.startsWith('Chart: Masthead')), `page ${index + 1}`).toHaveLength(1)
    view.unmount()
  })

  it('cards expose placeholder: reason', async () => {
    expect(chartDescription(cardChart('This chart type is not supported', 'Future'))).toBe('Chart: Future. placeholder: This chart type is not supported.')
    const docx = await parseDocx(await OfficePackage.load(buildChartDocx([{ chartXml: SURFACE }]).bytes))
    const view = render(<OfficeDoc document={docx} />)
    await waitFor(() => expect(view.container.querySelectorAll('[role="img"]').length).toBe(1))
    expect(view.container.querySelector('[role="img"]')!.getAttribute('aria-label')).toBe('Chart. placeholder: 3-D surface charts are shown as a placeholder.')
    view.unmount()
  })

  it('allowCopy false keeps descriptions readable, not selectable', async () => {
    const docx = await parseDocx(await OfficePackage.load(buildChartDocx([{ chartXml: titled('Revenue'), workbook: officeChartWorkbook([4.3, 2.5]) }]).bytes))
    for (const allowCopy of [true, false]) {
      const view = render(<OfficeDoc document={docx} allowCopy={allowCopy} />)
      await waitFor(() => expect(view.container.querySelectorAll('[role="img"]').length).toBe(1))
      const layer = view.container.querySelector<HTMLElement>('[data-testid="officeview-page-charts"]')!
      const item = within(layer).getByRole('img', { name: /^Chart: Revenue\. Column chart/ })
      // Read aloud either way: named, never hidden from assistive technology.
      expect(layer.closest('[aria-hidden="true"]')).toBeNull()
      if (allowCopy) {
        expect(layer.style.userSelect).toBe('')
        expect(item.textContent).toMatch(/^Chart: Revenue/)
      } else {
        expect(layer.style.userSelect).toBe('none')
        expect(item.textContent).toBe('')
      }
      view.unmount()
    }
  })

  it('descriptions are charged to the chart workspace and refunded with the pages', async () => {
    const { documentChartContext } = await import('../src/charts/context')
    const context = documentChartContext('desktop')
    const docx = await parseDocx(await OfficePackage.load(buildChartDocx([{ chartXml: titled('Revenue'), workbook: officeChartWorkbook([4.3, 2.5]) }]).bytes))
    const before = context.workspace.used
    const pages = await getPaintables(docx, { chartContext: context })
    expect(pages[0].charts).toHaveLength(1)
    expect(context.workspace.used).toBeGreaterThanOrEqual(before + 2 * pages[0].charts![0].description.length)
    pages.dispose()
    expect(context.workspace.used).toBe(before)
  })
})

describe('A4d carry-over: transferred user-shape pictures', () => {
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const CDR = 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing'
  const png = () => { const c = createCanvas(4, 4); const g = c.getContext('2d'); g.fillStyle = '#00ff00'; g.fillRect(0, 0, 4, 4); return new Uint8Array(c.toBuffer('image/png')) }
  /** A grey column chart with a user-shape picture over its top-left quarter. */
  async function docxWithPicture(): Promise<Uint8Array> {
    const JSZip = (await import('jszip')).default
    const chart = `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="808080"/></a:solidFill></c:spPr><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart><c:userShapes r:id="rU"/></c:chartSpace>`
    const shapes = `<c:userShapes xmlns:c="${C}" xmlns:cdr="${CDR}" xmlns:a="${A}" xmlns:r="${R}"><cdr:relSizeAnchor><cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from><cdr:to><cdr:x>0.25</cdr:x><cdr:y>0.25</cdr:y></cdr:to><cdr:pic><cdr:nvPicPr><cdr:cNvPr id="2" name="Picture 1"/><cdr:cNvPicPr/></cdr:nvPicPr><cdr:blipFill><a:blip r:embed="rI"/><a:stretch><a:fillRect/></a:stretch></cdr:blipFill><cdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></cdr:spPr></cdr:pic></cdr:relSizeAnchor></c:userShapes>`
    const zip = await JSZip.loadAsync(buildChartDocx([{ chartXml: chart }]).bytes)
    zip.file('word/charts/_rels/chart1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rU" Type="${R}/chartUserShapes" Target="../drawings/drawing1.xml"/></Relationships>`)
    zip.file('word/drawings/drawing1.xml', shapes)
    zip.file('word/drawings/_rels/drawing1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rI" Type="${R}/image" Target="../media/image1.png"/></Relationships>`)
    zip.file('word/media/image1.png', png())
    return zip.generateAsync({ type: 'uint8array' })
  }
  const greenPixels = (pages: Awaited<ReturnType<typeof getPaintables>>) => {
    const { widthPx, heightPx } = pages[0].spec
    const canvas = createCanvas(widthPx, heightPx)
    pages[0].paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)
    const data = canvas.getContext('2d').getImageData(0, 0, widthPx, heightPx).data
    let green = 0
    for (let i = 0; i < data.length; i += 4) if (data[i] < 40 && data[i + 1] > 200 && data[i + 2] < 40) green++
    return green
  }

  it('a model parsed in another realm shows its user-shape pictures at the first paint', async () => {
    const bytes = await docxWithPicture()
    // The parse realm: decoded there, painted there.
    const parsed = await parseDocx(await OfficePackage.load(bytes))
    const local = await getPaintables(parsed)
    const expected = greenPixels(local)
    local.dispose()
    expect(expected, 'the picture is painted in the parse realm').toBeGreaterThan(1000)
    // A fresh realm receiving a structured clone (the worker hand-off): the decode sidecar is empty.
    const transferred = structuredClone(parsed)
    vi.resetModules()
    const fresh = await import('../src/render/paint')
    const pages = await fresh.getPaintables(transferred)
    expect(greenPixels(pages), 'the first paint already shows the picture').toBe(expected)
    pages.dispose()
  })
})
