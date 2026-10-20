/**
 * A3b / Review Focus 1 in a real browser: a DOCX/PPTX chart whose values have
 * no cache renders natively from its embedded workbook (spec 6.4). The chart's
 * bars are recorded as they are painted on the page canvas: two series-colour
 * bars whose heights keep the workbook values' ratio (never a card, which
 * paints no series bars). Fixtures are Office-shaped synthetic.
 */
import { expect, test, type Page } from '@playwright/test'
import { buildChartDocx, buildChartPptx, officeChartWorkbook } from '../helpers/chart-fixtures'

interface Bar { w: number; h: number }
declare global { interface Window { __bars?: Bar[] } }

async function openDocument(page: Page, name: string, bytes: Uint8Array): Promise<Bar[]> {
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
  await page.goto(`/src/debug/viewer-e2e.html?${new URLSearchParams({ src: `/__fixtures/${name}` })}`)
  await page.locator('canvas[aria-label^="page "]').first().waitFor({ timeout: 60_000 })
  await expect.poll(() => page.evaluate(() => window.__bars?.length ?? 0), { timeout: 30_000 }).toBeGreaterThanOrEqual(2)
  return page.evaluate(() => window.__bars ?? [])
}

for (const format of ['docx', 'pptx'] as const) {
  test(`${format}: a cacheless chart renders natively from its embedded workbook`, async ({ page }) => {
    const fixture = (format === 'docx' ? buildChartDocx : buildChartPptx)([{ workbook: officeChartWorkbook([4.3, 2.5]) }])
    const bars = await openDocument(page, `chart.${format}`, fixture.bytes)
    // The first paint's two bars (a repaint may record them again).
    const [a, b] = bars
    expect(a.w).toBeCloseTo(b.w, 1)
    expect(a.h / b.h).toBeCloseTo(4.3 / 2.5, 1)
  })
}
