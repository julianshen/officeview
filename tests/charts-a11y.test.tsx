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

afterEach(() => { vi.restoreAllMocks(); vi.doUnmock('../src/core/images') })

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

/** A hand-driven clock: `fire()` ends every pending wait; nothing depends on wall time. */
function manualClock() {
  const waits: Array<{ resolve(): void; ms: number }> = []
  return {
    waits,
    clock: { after(ms: number) { let resolve = () => {}; const done = new Promise<void>(r => { resolve = r }); waits.push({ resolve, ms }); return { done, cancel: () => {} } } },
    fire() { for (const w of waits.splice(0)) w.resolve() },
  }
}
const flush = async (turns = 20) => { for (let i = 0; i < turns; i++) await Promise.resolve() }

describe('bounded, cancellable picture wait (A7 review I1)', () => {
  async function transferredRealm() {
    const bytes = await docxWithPictureForWait()
    const parsed = await parseDocx(await OfficePackage.load(bytes))
    const transferred = structuredClone(parsed)
    vi.resetModules()
    // The fresh realm's decoder never finishes until the test says so.
    let finish: (bitmap: unknown) => void = () => {}
    const decodes: number[] = []
    vi.doMock('../src/core/images', async () => {
      const actual = await vi.importActual<typeof import('../src/core/images')>('../src/core/images')
      return { ...actual, decodeImage: vi.fn(async (data: Uint8Array, mime?: string) => { decodes.push(data.length); await new Promise(r => { finish = r }); return actual.decodeImage(data, mime) }) }
    })
    const paint = await import('../src/render/paint')
    const { documentChartContext } = await import('../src/charts/context')
    return { transferred, paint, documentChartContext, decodes, finish: () => finish(undefined) }
  }

  it('a picture that misses the deadline publishes the page without it, with a diagnostic; it appears at the next paint', async () => {
    const realm = await transferredRealm()
    const context = realm.documentChartContext('desktop')
    const { clock, fire, waits } = manualClock()
    let done = false
    const pending = realm.paint.getPaintables(realm.transferred, { chartContext: context, chartClock: clock }).then(p => { done = true; return p })
    await vi.waitFor(() => expect(realm.decodes.length).toBe(1))
    await flush()
    expect(done, 'waiting for the picture, within the deadline').toBe(false)
    expect(waits.map(w => w.ms)).toEqual([1500])
    fire()
    const pages = await pending
    expect(pages.chartDiagnostics).toEqual([expect.objectContaining({ page: 0, code: 'chart-user-shapes-deadline' })])
    expect(greenOf(pages), 'published without the picture').toBe(0)
    // The in-flight decode stays charged until it settles, then the painter shows it at the next paint.
    const charged = context.workspace.used
    realm.finish()
    await vi.waitFor(() => expect(context.workspace.used).toBeLessThan(charged))
    await flush()
    expect(greenOf(pages), 'the next paint shows it').toBeGreaterThan(1000)
    pages.dispose()
  })

  it('aborting (the viewer unmounted) ends the wait at once and holds nothing past the decode', async () => {
    const realm = await transferredRealm()
    const context = realm.documentChartContext('desktop')
    const { clock, waits } = manualClock()
    const abort = new AbortController()
    const before = context.workspace.used
    const pending = realm.paint.getPaintables(realm.transferred, { chartContext: context, chartClock: clock, signal: abort.signal })
    await vi.waitFor(() => expect(realm.decodes.length).toBe(1))
    abort.abort()
    const pages = await pending
    expect(waits).toHaveLength(1)
    expect(pages.chartDiagnostics ?? [], 'an abort is not a missed deadline').toEqual([])
    pages.dispose()
    // Only the in-flight decode (its surface and transient) is still charged; it is released as it settles.
    const { releaseChartModel } = await import('../src/charts/own')
    for (const image of collectCharts(realm.transferred)) releaseChartModel(image)
    realm.finish()
    await vi.waitFor(() => expect(context.workspace.used).toBe(before))
  })
})

/** The transferred-picture fixture of the A4d test, reused. */
async function docxWithPictureForWait(): Promise<Uint8Array> {
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart', A = 'http://schemas.openxmlformats.org/drawingml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', CDR = 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing'
  const JSZip = (await import('jszip')).default
  const c = createCanvas(4, 4); const g = c.getContext('2d'); g.fillStyle = '#00ff00'; g.fillRect(0, 0, 4, 4)
  const chart = `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="808080"/></a:solidFill></c:spPr><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart><c:userShapes r:id="rU"/></c:chartSpace>`
  const shapes = `<c:userShapes xmlns:c="${C}" xmlns:cdr="${CDR}" xmlns:a="${A}" xmlns:r="${R}"><cdr:relSizeAnchor><cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from><cdr:to><cdr:x>0.25</cdr:x><cdr:y>0.25</cdr:y></cdr:to><cdr:pic><cdr:nvPicPr><cdr:cNvPr id="2" name="Picture 1"/><cdr:cNvPicPr/></cdr:nvPicPr><cdr:blipFill><a:blip r:embed="rI"/><a:stretch><a:fillRect/></a:stretch></cdr:blipFill><cdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></cdr:spPr></cdr:pic></cdr:relSizeAnchor></c:userShapes>`
  const zip = await JSZip.loadAsync(buildChartDocx([{ chartXml: chart }]).bytes)
  zip.file('word/charts/_rels/chart1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rU" Type="${R}/chartUserShapes" Target="../drawings/drawing1.xml"/></Relationships>`)
  zip.file('word/drawings/drawing1.xml', shapes)
  zip.file('word/drawings/_rels/drawing1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rI" Type="${R}/image" Target="../media/image1.png"/></Relationships>`)
  zip.file('word/media/image1.png', new Uint8Array(c.toBuffer('image/png')))
  return zip.generateAsync({ type: 'uint8array' })
}
function greenOf(pages: Awaited<ReturnType<typeof getPaintables>>): number {
  const { widthPx, heightPx } = pages[0].spec
  const canvas = createCanvas(widthPx, heightPx)
  pages[0].paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)
  const data = canvas.getContext('2d').getImageData(0, 0, widthPx, heightPx).data
  let green = 0
  for (let i = 0; i < data.length; i += 4) if (data[i] < 40 && data[i + 1] > 200 && data[i + 2] < 40) green++
  return green
}
function collectCharts(doc: { sections: Array<{ paragraphs: Array<{ images?: Array<{ drawing?: { kind: string; chart?: unknown } }> }> }> }): Array<import('../src/charts/model').ChartModel> {
  return doc.sections.flatMap(s => s.paragraphs.flatMap(p => (p.images ?? []).flatMap(i => i.drawing?.kind === 'xlsxChart' ? [i.drawing.chart as import('../src/charts/model').ChartModel] : [])))
}

describe('admission and complete lists (A7 review I2, I3, M1)', () => {
  const model = (title: string) => {
    let reads = 0
    const chart = { ...cardChart('reason', title) }
    Object.defineProperty(chart, 'title', { get() { reads++; return title } })
    return { chart, reads: () => reads }
  }

  it('admits each page before building it; a refused page gets a notice and a diagnostic, later pages are unaffected', async () => {
    const { prepareChartPaintables, DESCRIPTION_WORK_BYTES } = await import('../src/charts/paintables')
    const { createChartWorkspace } = await import('../src/charts/workspace')
    // The parent refuses only page 0's admission.
    const workspace = createChartWorkspace('mobile', { reserve: (_bytes, tag) => !tag.endsWith(':page0'), release: () => {} })
    const a = model('Refused'), b = model('Kept')
    const observed: number[] = []
    const result = await prepareChartPaintables({ profile: 'mobile', workspace }, {
      pages: 3,
      bound: page => page === 2 ? 0 : 1,
      collect: page => { observed.push(workspace.used); return [{ chart: page === 0 ? a.chart : b.chart, x: 1, y: 2, width: 3, height: 4 }] },
    })
    expect(a.reads(), 'nothing of the refused page was built').toBe(0)
    expect(observed, 'only page 1 was collected').toHaveLength(1)
    expect(observed[0], 'its bound was charged before construction').toBeGreaterThanOrEqual(DESCRIPTION_WORK_BYTES)
    expect(result.page(0)).toEqual({ total: 1, charts: [expect.objectContaining({ notice: true, description: 'Chart descriptions for this page could not be kept in memory.' })] })
    expect(result.page(1).charts.map(c => c.description)).toEqual(['Chart: Kept. placeholder: reason.'])
    expect(result.page(2)).toEqual({ charts: [], total: 0 })
    expect(result.diagnostics).toEqual([expect.objectContaining({ page: 0, code: 'chart-description-refused', reason: expect.stringMatching(/\(1 chart\)/) })])
    // Reconciled to the exact retention, then refunded.
    expect(workspace.used).toBeLessThan(DESCRIPTION_WORK_BYTES)
    result.release()
    expect(workspace.used).toBe(0)
  })

  it('repeated charts over many pages: each page is bounded at construction, the string is retained once', async () => {
    const { prepareChartPaintables, DESCRIPTION_WORK_BYTES, RECORD_BYTES, ENTRY_BYTES, MAP_BYTES } = await import('../src/charts/paintables')
    const { createChartWorkspace } = await import('../src/charts/workspace')
    const workspace = createChartWorkspace('mobile')
    const shared = model('Masthead')
    const peaks: number[] = []
    const result = await prepareChartPaintables({ profile: 'mobile', workspace }, { pages: 200, bound: () => 1, collect: () => { peaks.push(workspace.used); return [{ chart: shared.chart, x: 0, y: 0, width: 1, height: 1 }] } })
    expect(Array.from({ length: 200 }, (_, i) => result.page(i).charts.length).every(n => n === 1)).toBe(true)
    const retained = workspace.used
    // The construction peak never exceeds the retained set plus one page's bound.
    expect(Math.max(...peaks)).toBeLessThanOrEqual(retained + DESCRIPTION_WORK_BYTES + 64 * 4096)
    // Exactly what stays live: 200 page records and entries, one description string and its map entry.
    expect(retained).toBe(200 * (RECORD_BYTES + ENTRY_BYTES) + 2 * 'Chart: Masthead. placeholder: reason.'.length + MAP_BYTES)
    result.release()
    expect(workspace.used).toBe(0)
  })

  it('65 charts on one slide: 64 described, a "1 more chart not listed" note, and the true total', async () => {
    const pptx = await parsePptx(await OfficePackage.load(buildChartPptx(Array.from({ length: 65 }, (_, i) => ({ chartXml: titled(`C${i + 1}`), workbook: officeChartWorkbook([1, 2]) }))).bytes))
    expect((pptx.drawingCoverage ?? []).filter(e => e.scope === 'original' && e.selectedRepresentation === 'chart')).toHaveLength(65)
    const pages = await getPaintables(pptx)
    expect(pages[0].chartTotal).toBe(65)
    expect(pages[0].charts).toHaveLength(65)
    expect(pages[0].charts!.filter(c => !c.notice)).toHaveLength(64)
    expect(pages[0].charts!.at(-1)).toMatchObject({ notice: true, description: '1 more chart not listed.' })
    expect(pages.chartDiagnostics).toEqual([expect.objectContaining({ page: 0, code: 'chart-description-limit' })])
    pages.dispose()
    const view = render(<OfficeDoc document={pptx} />)
    await waitFor(() => expect(view.container.querySelectorAll('[role="img"]').length).toBe(64))
    expect(lists(view.container)[0].getAttribute('aria-label')).toBe('Charts on slide 1 (65)')
    expect(within(lists(view.container)[0]).getByRole('note').getAttribute('aria-label')).toBe('1 more chart not listed.')
    view.unmount()
  }, 60_000)

  it('grouped charts are placed at their own page-space boxes', async () => {
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(buildChartPptx([{ chartXml: titled('Left'), workbook: officeChartWorkbook([1, 2]) }, { chartXml: titled('Right'), workbook: officeChartWorkbook([3, 4]) }]).bytes)
    let slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
    slide = slide.replace('<a:off x="838300"', '<a:off x="6000000"')
    // A group whose child space is twice its own box (children are drawn at half size), placed at (952500, 952500).
    slide = slide.replace(/(<p:graphicFrame>.*<\/p:graphicFrame>)/, '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="6096000" cy="3429000"/><a:chOff x="0" y="0"/><a:chExt cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr>$1</p:grpSp>')
    zip.file('ppt/slides/slide1.xml', slide)
    const pptx = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const pages = await getPaintables(pptx)
    const [left, right] = pages[0].charts!
    const px = (emu: number) => emu / 9525
    expect(left.x).toBeCloseTo(px(952500 + 838200 / 2), 0)
    expect(left.y).toBeCloseTo(px(952500 + 1825625 / 2), 0)
    expect(left.width).toBeCloseTo(px(5486400 / 2), 0)
    expect(right.x).toBeCloseTo(px(952500 + 6000000 / 2), 0)
    expect(right.width).toBeCloseTo(left.width, 3)
    pages.dispose()
  })
})

describe('A7 review round 2 (R1, R2, R3, M1)', () => {
  it('a fully refused workspace still exposes a readable notice per page, with a diagnostic and nothing charged', async () => {
    const { prepareChartPaintables } = await import('../src/charts/paintables')
    const { createChartWorkspace } = await import('../src/charts/workspace')
    const workspace = createChartWorkspace('mobile', { reserve: () => false, release: () => {} })
    const chart = cardChart('reason', 'Never built')
    const result = await prepareChartPaintables({ profile: 'mobile', workspace }, { pages: 3, bound: () => 2, collect: () => { throw new Error('never collected') } })
    for (const page of [0, 1, 2].map(i => result.page(i))) expect(page.charts).toEqual([expect.objectContaining({ notice: true, description: 'Chart descriptions for this page could not be kept in memory.' })])
    expect(result.diagnostics).toEqual([expect.objectContaining({ code: 'chart-description-refused' })])
    expect(workspace.used).toBe(0)
    void chart
    // End to end: the mounted list keeps the notice (a note), labelled without a false count.
    const { documentChartContext } = await import('../src/charts/context')
    const docx = await parseDocx(await OfficePackage.load(buildChartDocx([{ chartXml: titled('Revenue'), workbook: officeChartWorkbook([4.3, 2.5]) }]).bytes))
    const context = documentChartContext('mobile')
    const pages = await getPaintables(docx, { chartContext: context })
    // Fill the workspace after the parse so only the metadata is refused.
    pages.dispose()
    context.workspace.reserve(context.workspace.limit - context.workspace.used, 'full')
    const refused = await getPaintables(docx, { chartContext: context })
    expect(refused[0].charts).toEqual([expect.objectContaining({ notice: true })])
    expect(refused[0].chartTotal).toBeUndefined()
    const { PageCharts } = await import('../src/components/paged/PageCharts')
    const view = render(<PageCharts pages={refused} unit="Page" selectable={false} />)
    expect(view.getByRole('list', { name: 'Charts on page 1' })).toBeTruthy()
    expect(view.getByRole('note').getAttribute('aria-label')).toBe('Chart descriptions for this page could not be kept in memory.')
    view.unmount()
    refused.dispose()
  })

  it('retained diagnostics and records stay charged; a refused page with room for its record keeps the true total', async () => {
    const { prepareChartPaintables, RECORD_BYTES } = await import('../src/charts/paintables')
    const { createChartWorkspace } = await import('../src/charts/workspace')
    const workspace = createChartWorkspace('mobile', { reserve: (_bytes, tag) => !tag.endsWith(':page0'), release: () => {} })
    const result = await prepareChartPaintables({ profile: 'mobile', workspace }, { pages: 1, bound: () => 3, collect: () => [] })
    expect(result.page(0).total).toBe(3)
    const reason = result.diagnostics[0].reason
    expect(workspace.used).toBe(RECORD_BYTES + 64 + 2 * reason.length)
    result.release()
    expect(workspace.used).toBe(0)
  })

  it('an exception mid-construction releases everything the job admitted (first page and after a successful page)', async () => {
    const { prepareChartPaintables } = await import('../src/charts/paintables')
    const { createChartWorkspace } = await import('../src/charts/workspace')
    for (const failAt of [0, 2]) {
      const workspace = createChartWorkspace('mobile')
      const chart = cardChart('reason', 'Kept')
      await expect(prepareChartPaintables({ profile: 'mobile', workspace }, {
        pages: 4, bound: () => 1,
        collect: page => { if (page === failAt) throw new Error('layout failed'); return [{ chart, x: 0, y: 0, width: 1, height: 1 }] },
      }), `fail at page ${failAt}`).rejects.toThrow('layout failed')
      expect(workspace.used, `fail at page ${failAt}`).toBe(0)
    }
  })

  it('rotated and flipped groups place each chart at its painted bounding box', async () => {
    const JSZip = (await import('jszip')).default
    const E = 9525
    const build = async (xfrm: string) => {
      const zip = await JSZip.loadAsync(buildChartPptx([{ chartXml: titled('Inner'), workbook: officeChartWorkbook([1, 2]) }]).bytes)
      let slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
      slide = slide.replace(/<a:off x="838200" y="1825625"\/><a:ext cx="5486400" cy="3200400"\/>/, `<a:off x="0" y="0"/><a:ext cx="${100 * E}" cy="${50 * E}"/>`)
      slide = slide.replace(/(<p:graphicFrame>.*<\/p:graphicFrame>)/, `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm${xfrm}><a:off x="${100 * E}" y="${100 * E}"/><a:ext cx="${200 * E}" cy="${100 * E}"/><a:chOff x="0" y="0"/><a:chExt cx="${200 * E}" cy="${100 * E}"/></a:xfrm></p:grpSpPr>$1</p:grpSp>`)
      zip.file('ppt/slides/slide1.xml', slide)
      const pages = await getPaintables(await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))))
      const box = pages[0].charts![0]
      pages.dispose()
      return [box.x, box.y, box.width, box.height].map(v => Math.round(v))
    }
    // A 90-degree group turns its child (100..200, 100..150) about the group centre (200, 150).
    expect(await build(' rot="5400000"')).toEqual([200, 50, 50, 100])
    // A horizontally flipped group mirrors it about x = 200.
    expect(await build(' flipH="1"')).toEqual([200, 100, 100, 50])
    expect(await build('')).toEqual([100, 100, 100, 50])
  })
})

describe('A7 review round 3 (R2: the result never grows uncharged)', () => {
  it('100,000 refused or empty pages store nothing and charge nothing, yet every refused page reads its notice', async () => {
    const { prepareChartPaintables } = await import('../src/charts/paintables')
    const { createChartWorkspace } = await import('../src/charts/workspace')
    const refusing = createChartWorkspace('mobile', { reserve: () => false, release: () => {} })
    const refused = await prepareChartPaintables({ profile: 'mobile', workspace: refusing }, { pages: 100_000, bound: page => page % 2, collect: () => { throw new Error('never collected') } })
    expect(refused.stored).toBe(0)
    expect(refusing.used).toBe(0)
    expect(refused.diagnostics).toHaveLength(1)
    expect(refused.page(99_999).charts).toEqual([expect.objectContaining({ notice: true })])
    expect(refused.page(99_998).charts).toEqual([])
    const open = createChartWorkspace('mobile')
    const empty = await prepareChartPaintables({ profile: 'mobile', workspace: open }, { pages: 100_000, bound: () => 0, collect: () => [] })
    expect(empty.stored).toBe(0)
    expect(open.used).toBe(0)
  })

  it('stored pages are exactly the charged ones', async () => {
    const { prepareChartPaintables, RECORD_BYTES, ENTRY_BYTES, MAP_BYTES } = await import('../src/charts/paintables')
    const { createChartWorkspace } = await import('../src/charts/workspace')
    const workspace = createChartWorkspace('mobile')
    const chart = cardChart('reason', 'Sparse')
    const result = await prepareChartPaintables({ profile: 'mobile', workspace }, { pages: 10_000, bound: page => page % 1000 === 0 ? 1 : 0, collect: () => [{ chart, x: 0, y: 0, width: 1, height: 1 }] })
    expect(result.stored).toBe(10)
    expect(workspace.used).toBe(10 * (RECORD_BYTES + ENTRY_BYTES) + 2 * 'Chart: Sparse. placeholder: reason.'.length + MAP_BYTES)
    result.release()
    expect(workspace.used).toBe(0)
  })
})
