/**
 * View-only row/column resizing in the real viewer (spec §5.6 amendment), on
 * Chromium and Playwright WebKit (`webkit-playwright`, never Safari evidence):
 * a real mouse drag on a column header's right edge and a row header's bottom
 * edge (guide line, one commit on release), double-click auto-fit, the
 * repainted canvas, the status announcement and Reset sizes.
 */
import { expect, test, type Page } from '@playwright/test'
import { ensureFixtures, openViewer, status } from './viewer-e2e'

test.use({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 })
test.beforeAll(async () => { await ensureFixtures() })

const columnHeader = (page: Page, letter: string) => page.locator(`[role=columnheader][aria-label^="Column ${letter}"]`).first()
const width = async (page: Page, letter: string) => (await columnHeader(page, letter).boundingBox())!.width
const rowHeight = async (page: Page, label: string) => (await page.getByRole('rowheader', { name: label, exact: true }).boundingBox())!.height
async function center(page: Page, selector: string): Promise<{ x: number; y: number }> {
  const box = (await page.locator(selector).first().boundingBox())!
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 }
}
const paneShot = (page: Page) => page.getByTestId('officeview-sheet-base').screenshot()

test.describe('view-only resize (real viewer)', () => {
  test('drag a column edge: guide while dragging, applied on release, canvas repainted; double-click fits; Reset restores', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const before = await width(page, 'B')
    const shot = await paneShot(page)
    const edge = await center(page, '[data-resize-column="1"]')
    await page.mouse.move(edge.x, edge.y)
    await page.mouse.down()
    await page.mouse.move(edge.x + 30, edge.y, { steps: 3 })
    await page.mouse.move(edge.x + 60, edge.y, { steps: 3 })
    await expect(page.locator('[data-officeview-ui="resize-guide"]')).toContainText(`Width: ${Math.round(before + 60)} px`)
    expect(await width(page, 'B')).toBeCloseTo(before, 0)
    await page.mouse.up()
    await expect(page.locator('[data-officeview-ui="resize-guide"]')).toHaveCount(0)
    await expect.poll(() => width(page, 'B')).toBeCloseTo(before + 60, 0)
    await expect(status(page)).toContainText(`Column B width ${Math.round(before + 60)} pixels`)
    await expect.poll(async () => Buffer.compare(await paneShot(page), shot) !== 0).toBe(true)

    // Double-click the same edge: auto-fit to the widest visible/known value.
    const fitEdge = await center(page, '[data-resize-column="1"]')
    await page.mouse.dblclick(fitEdge.x, fitEdge.y)
    await expect(status(page)).toContainText(/Column B width \d+ pixels, fitted to content/)
    const fitted = await width(page, 'B')
    expect(fitted).not.toBeCloseTo(before + 60, 0)
    expect(fitted).toBeLessThan(before + 60)

    // A row's bottom edge.
    const rowBefore = await rowHeight(page, 'Row 3')
    const rowEdge = await center(page, '[data-resize-row="2"]')
    await page.mouse.move(rowEdge.x, rowEdge.y)
    await page.mouse.down()
    await page.mouse.move(rowEdge.x, rowEdge.y + 24, { steps: 4 })
    await page.mouse.up()
    await expect.poll(() => rowHeight(page, 'Row 3')).toBeCloseTo(rowBefore + 24, 0)
    await expect(status(page)).toContainText(`Row 3 height ${Math.round(rowBefore + 24)} pixels`)

    await page.getByRole('button', { name: 'Reset sizes' }).click()
    await expect.poll(() => width(page, 'B')).toBeCloseTo(before, 0)
    await expect.poll(() => rowHeight(page, 'Row 3')).toBeCloseTo(rowBefore, 0)
    await expect(status(page)).toContainText('restored to the workbook sizes')
    await expect(page.getByRole('button', { name: 'Reset sizes' })).toBeDisabled()
  })

  test('keyboard: Alt+Shift+Arrow on the grid resizes the active column, announced', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const before = await width(page, 'A')
    await page.getByRole('grid').click({ position: { x: 44 + 10, y: 60 } })
    await page.keyboard.press('Alt+Shift+ArrowRight')
    await expect.poll(() => width(page, 'A')).toBeCloseTo(before + 8, 0)
    await expect(status(page)).toContainText(`Column A width ${Math.round(before + 8)} pixels`)
  })

  test('worker-backed sheet: a drag applies on the main-thread geometry and repaints; Reset restores', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx', { worker: true, paged: true })
    const before = await width(page, 'C')
    const shot = await paneShot(page)
    const edge = await center(page, '[data-resize-column="2"]')
    await page.mouse.move(edge.x, edge.y)
    await page.mouse.down()
    await page.mouse.move(edge.x + 80, edge.y, { steps: 4 })
    await page.mouse.up()
    await expect.poll(() => width(page, 'C')).toBeCloseTo(before + 80, 0)
    await expect.poll(async () => Buffer.compare(await paneShot(page), shot) !== 0).toBe(true)
    await page.getByRole('button', { name: 'Reset sizes' }).click()
    await expect.poll(() => width(page, 'C')).toBeCloseTo(before, 0)
  })
})
