/**
 * PR #11 stream F (iOS Simulator run), real-browser evidence at phone and
 * tablet sizes with touch emulation (Playwright emulation, never Safari or
 * device evidence):
 * - item 2: the column-header row stays one label row plus one 44px actions
 *   row (it was 4 x 44 = 176px), and every column action stays reachable
 *   with 44px targets;
 * - item 3: the zoom controls never cover the grid (they dock in the toolbar).
 */
import { expect, test, type Page } from '@playwright/test'
import { ensureFixtures, openViewer } from './viewer-e2e'

type Box = { left: number; right: number; top: number; bottom: number; width: number; height: number }
const rect = (page: Page, selector: string) => page.evaluate(s => document.querySelector(s)!.getBoundingClientRect().toJSON() as Box, selector)
const intersects = (a: Box, b: Box) => a.left < b.right - 0.5 && b.left < a.right - 0.5 && a.top < b.bottom - 0.5 && b.top < a.bottom - 0.5
const TARGET = 44

for (const viewport of [{ width: 390, height: 844 }, { width: 834, height: 1194 }]) {
  test.describe(`${viewport.width}x${viewport.height} touch viewport`, () => {
    test.use({ viewport, hasTouch: true, isMobile: true, deviceScaleFactor: 2 })
    test.beforeAll(async () => { await ensureFixtures() })

    test('compact column headers keep every action reachable with 44px targets; zoom never covers the grid', async ({ page }) => {
      await openViewer(page, 'f2-1k.xlsx')
      test.skip(!(await page.evaluate(() => matchMedia('(pointer: coarse)').matches)), 'this engine does not emulate a coarse pointer')
      const headers = await rect(page, '[data-officeview-ui=grid-headers]')
      expect(headers.height).toBeLessThanOrEqual(2 * TARGET + 0.5)

      // Column B (default width): its actions are inline or behind its disclosure.
      const disclosure = page.getByRole('button', { name: 'Actions for column B' })
      if (await disclosure.count()) {
        const box = (await disclosure.boundingBox())!
        expect(box.width).toBeGreaterThanOrEqual(TARGET - 0.5)
        expect(box.height).toBeGreaterThanOrEqual(TARGET - 0.5)
        await page.touchscreen.tap(box.x + box.width / 2, box.y + box.height / 2)
        await expect(page.getByRole('group', { name: 'Column B actions' })).toBeVisible()
      }
      for (const name of ['Cycle sort column B', /^Filter column B/, 'More sort options for B'] as const) {
        const action = page.getByRole('button', { name })
        await expect(action).toBeVisible()
        const box = (await action.boundingBox())!
        expect(box.width, String(name)).toBeGreaterThanOrEqual(TARGET - 0.5)
        expect(box.height, String(name)).toBeGreaterThanOrEqual(TARGET - 0.5)
      }
      // The sort action works from the group (a tap sorts column B and closes it).
      const sort = (await page.getByRole('button', { name: 'Cycle sort column B' }).boundingBox())!
      await page.touchscreen.tap(sort.x + sort.width / 2, sort.y + sort.height / 2)
      await expect(page.getByRole('status').filter({ hasText: 'Visible body' })).toContainText('View-only sort applied', { timeout: 60_000 })
      await expect(page.getByRole('group', { name: 'Column B actions' })).toHaveCount(0)

      // Zoom controls: docked outside the cell area, never over the active cell or data.
      const zoom = await rect(page, '[data-testid=officeview-zoom-controls]'), pane = await rect(page, '[data-testid=officeview-sheet-pane]')
      expect(intersects(zoom, pane), `zoom ${JSON.stringify(zoom)} covers the grid ${JSON.stringify(pane)}`).toBe(false)
      for (const name of ['Zoom in', 'Zoom out', 'Reset zoom']) {
        const box = (await page.getByRole('button', { name }).boundingBox())!
        expect(box.width, name).toBeGreaterThanOrEqual(TARGET - 0.5)
        expect(box.height, name).toBeGreaterThanOrEqual(TARGET - 0.5)
      }
      await page.getByRole('button', { name: 'Zoom in' }).tap()
      await expect(page.getByRole('button', { name: 'Reset zoom' })).not.toHaveText('100%')
    })
  })
}
