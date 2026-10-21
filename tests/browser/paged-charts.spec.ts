/**
 * Task A7 (spec 6.6, 8): charts in the paged viewer, in a real browser. Every "renders natively" test rejects
 * fallbacks: it records the series bars as they are painted on the page canvas (native geometry whose heights
 * keep the data's ratio; a card paints no series bars) and requires the chart's accessible description to name
 * a native type, never a placeholder. Dark mode must leave the page pixels identical (paper rule), and charts
 * loading must not shift layout. Fixtures are Office-shaped synthetic (tests/helpers/chart-fixtures.ts).
 *
 * Page: src/debug/viewer-e2e.html (the shipped <OfficeFile>). The Layout Instability API is Chromium-only.
 */
import { expect, test, type Page } from '@playwright/test'
import { buildChartDocx, buildChartPptx, officeChartWorkbook, officeChartXml } from '../helpers/chart-fixtures'

interface Bar { w: number; h: number }
declare global { interface Window { __bars?: Bar[]; __chartShifts?: Array<{ value: number; recent: boolean; t: number; chrome?: boolean; sources?: string[] }> } }

/** A column chart whose single series is cached as literal values (no embedded workbook needed). */
const cachedColumn = (values: number[]) => officeChartXml().replace(/<c:val>.*?<\/c:val>/, `<c:val><c:numLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit></c:val>`)
  .replace(/<c:cat>.*?<\/c:cat>/, `<c:cat><c:strLit><c:ptCount val="${values.length}"/>${values.map((_, i) => `<c:pt idx="${i}"><c:v>K${i}</c:v></c:pt>`).join('')}</c:strLit></c:cat>`)
  .replace(/<c:tx><c:strRef>.*?<\/c:strRef><\/c:tx>/, '<c:tx><c:v>Cached</c:v></c:tx>')

async function open(page: Page, name: string, bytes: Uint8Array, query: Record<string, string> = {}): Promise<void> {
  await page.addInitScript(() => {
    const bars: Bar[] = []
    window.__bars = bars
    for (const proto of [CanvasRenderingContext2D.prototype, ...(typeof OffscreenCanvasRenderingContext2D === 'function' ? [OffscreenCanvasRenderingContext2D.prototype] : [])]) {
      const fillRect = proto.fillRect
      proto.fillRect = function (this: CanvasRenderingContext2D, x: number, y: number, w: number, h: number) {
        if (String(this.fillStyle).toLowerCase() === '#4472c4') bars.push({ w: Math.abs(w * this.getTransform().a), h: Math.abs(h * this.getTransform().d) })
        return fillRect.call(this, x, y, w, h)
      } as typeof proto.fillRect
    }
  })
  await page.route(url => url.pathname === `/__fixtures/${name}`, route => route.fulfill({ body: Buffer.from(bytes), contentType: 'application/octet-stream' }))
  await page.goto(`/src/debug/viewer-e2e.html?${new URLSearchParams({ src: `/__fixtures/${name}`, ...query })}`)
}
/** The first paint's series bars (a repaint may record them again). */
async function bars(page: Page, count = 2): Promise<Bar[]> {
  await page.locator('canvas[aria-label^="page "]').first().waitFor({ timeout: 60_000 })
  await expect.poll(() => page.evaluate(() => window.__bars?.length ?? 0), { timeout: 30_000 }).toBeGreaterThanOrEqual(count)
  return (await page.evaluate(() => window.__bars ?? [])).slice(0, count)
}
/** The accessible chart descriptions of the page chart layer, by list. */
const descriptions = (page: Page) => page.evaluate(() => [...document.querySelectorAll('[data-testid="officeview-page-charts"] ul')].map(ul => ({
  name: ul.getAttribute('aria-label'), items: [...ul.querySelectorAll('[role="img"]')].map(li => li.getAttribute('aria-label') ?? ''),
})))
const frames = (page: Page) => page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))))
function expectNative(found: Bar[], ratio: number, label: string): void {
  const [a, b] = found
  expect(a.w, `${label}: equal bar widths`).toBeCloseTo(b.w, 1)
  expect(a.h, `${label}: a real bar`).toBeGreaterThan(10)
  expect(a.h / b.h, `${label}: heights keep the data ratio`).toBeCloseTo(ratio, 1)
}
async function expectNativeDescription(page: Page, unit: 'page' | 'slide', label: string): Promise<void> {
  await expect.poll(() => descriptions(page).then(lists => lists.length), { timeout: 30_000 }).toBe(1)
  const [list] = await descriptions(page)
  expect(list.name, label).toBe(`Charts on ${unit} 1 (1)`)
  expect(list.items, label).toHaveLength(1)
  expect(list.items[0], `${label}: names a native type`).toMatch(/^Chart\. Column chart, 1 series/)
  expect(list.items[0], `${label}: never a card`).not.toContain('placeholder')
  // The named element is exposed to assistive technology.
  await expect(page.getByRole('img', { name: /^Chart\. Column chart/ })).toHaveCount(1)
}

test.describe('paged charts (spec 6.6, 8)', () => {
  test('docx and pptx charts render natively in the paged viewer (browser)', async ({ page }) => {
    for (const format of ['docx', 'pptx'] as const) {
      const fixture = (format === 'docx' ? buildChartDocx : buildChartPptx)([{ chartXml: cachedColumn([3, 6]) }])
      await open(page, `cached.${format}`, fixture.bytes)
      expectNative(await bars(page), 3 / 6, format)
      await expectNativeDescription(page, format === 'docx' ? 'page' : 'slide', format)
      expect((await descriptions(page))[0].items[0], format).toContain('Cached: 3, 6')
      await page.unrouteAll({ behavior: 'ignoreErrors' })
    }
  })

  test('cacheless embedded-workbook chart renders natively (browser)', async ({ page }) => {
    for (const format of ['docx', 'pptx'] as const) {
      const fixture = (format === 'docx' ? buildChartDocx : buildChartPptx)([{ workbook: officeChartWorkbook([4.3, 2.5]) }])
      await open(page, `cacheless.${format}`, fixture.bytes)
      expectNative(await bars(page), 4.3 / 2.5, format)
      await expectNativeDescription(page, format === 'docx' ? 'page' : 'slide', format)
      expect((await descriptions(page))[0].items[0], format).toContain('4.3, 2.5')
      await page.unrouteAll({ behavior: 'ignoreErrors' })
    }
  })

  test('dark mode leaves chart pixels identical (browser)', async ({ page }) => {
    const fixture = buildChartDocx([{ workbook: officeChartWorkbook([4.3, 2.5]) }])
    const pixels: Record<string, string> = {}
    for (const theme of ['light', 'dark'] as const) {
      await page.emulateMedia({ colorScheme: theme })
      await open(page, `theme-${theme}.docx`, fixture.bytes, { theme })
      expectNative(await bars(page), 4.3 / 2.5, theme)
      await expect(page.locator('[data-ov-theme-root]').first()).toHaveAttribute('data-ov-theme', theme)
      await frames(page)
      pixels[theme] = await page.locator('canvas[aria-label^="page "]').first().evaluate(canvas => (canvas as HTMLCanvasElement).toDataURL('image/png'))
      await page.unrouteAll({ behavior: 'ignoreErrors' })
    }
    expect(pixels.dark.length).toBeGreaterThan(1000)
    expect(pixels.dark === pixels.light, 'the dark page canvas is pixel-identical to the light one').toBe(true)
  })

  test('CLS 0 while charts load (browser)', async ({ page, browserName }) => {
    test.skip(browserName !== 'chromium', 'the Layout Instability API is Chromium-only')
    await page.addInitScript(() => {
      window.__chartShifts = []
      // Chrome only (toolbar, status line, notices): the viewer's own mount, measured by paged-toolbar.spec.ts.
      const CHROME = '[data-testid="officeview-status-slot"], [data-testid="officeview-toolbar"], [data-testid="officeview-toolbar-placeholder"], [data-officeview-chrome]'
      new PerformanceObserver(list => {
        for (const e of list.getEntries() as unknown as Array<{ value: number; hadRecentInput: boolean; startTime: number; sources?: Array<{ node?: Node | null }> }>) {
          const elements = (e.sources ?? []).map(src => src.node && src.node.nodeType === 1 ? src.node as Element : src.node?.parentElement ?? null)
          window.__chartShifts!.push({ value: e.value, recent: e.hadRecentInput, t: e.startTime, chrome: elements.length > 0 && elements.every(el => !!el?.closest(CHROME)),
            sources: elements.map(el => el ? `${el.tagName.toLowerCase()}${el.getAttribute('data-testid') ? `[${el.getAttribute('data-testid')}]` : ''}` : 'text') })
        }
      }).observe({ type: 'layout-shift', buffered: true })
    })
    // Every shift over the whole load, from navigation: the parse (which loads the chart engine), the pages with
    // their charts (pictures decoded before the first paint), the chart descriptions layer and the lazy chrome.
    const fixture = buildChartDocx([{ workbook: officeChartWorkbook([4.3, 2.5]) }, { chartXml: cachedColumn([3, 6]) }])
    await open(page, 'cls.docx', fixture.bytes)
    await bars(page, 4)
    await expect.poll(() => descriptions(page).then(lists => lists[0]?.items.length ?? 0), { timeout: 30_000 }).toBe(2)
    await frames(page)
    await frames(page)
    const shifts = await page.evaluate(() => (window.__chartShifts ?? []).filter(s => !s.recent))
    // Shifts that move anything but the viewer chrome: pages, charts, the description layer, the host page.
    const content = shifts.filter(s => !s.chrome)
    expect(content.reduce((sum, s) => sum + s.value, 0), `layout shifts outside the chrome while the charts loaded: ${JSON.stringify(content)}`).toBe(0)
    expect(shifts.flatMap(s => s.sources ?? []).filter(source => /canvas|officeview-page-charts|officeview-content/.test(source)), 'no page canvas or chart layer ever shifts').toEqual([])
  })
})
