/**
 * Task 8 fix round 2 (visual review, agy + codex): on a 375x667 touch screen
 * the sheet toolbar's action buttons (Copy selection, Print, Reset sizes) are
 * full 44px targets, visible box and hit area alike, and the toolbar takes
 * fewer rows. The canvas-mirrored geometry is unchanged: an 88px touch header
 * (label row + 44px actions row) and the 24px status bar.
 */
import { expect, test } from '@playwright/test'
import { ensureFixtures, openViewer } from './viewer-e2e'

const TARGET = 44

test.describe('375x667 touch sheet toolbar', () => {
  test.use({ viewport: { width: 375, height: 667 }, hasTouch: true, isMobile: true, deviceScaleFactor: 2 })
  test.beforeAll(async () => { await ensureFixtures() })

  test('sheet action buttons are 44px touch targets; mirrored header geometry unchanged (browser)', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    test.skip(!(await page.evaluate(() => matchMedia('(pointer: coarse)').matches)), 'this engine does not emulate a coarse pointer')
    const actions = page.locator('[data-officeview-ui=sheet-actions]')
    for (const name of ['Copy selection', 'Print', 'Reset sizes']) {
      const button = actions.getByRole('button', { name, exact: true })
      const box = (await button.boundingBox())!
      expect(box.height, `${name} height`).toBeGreaterThanOrEqual(TARGET - 0.5)
      expect(box.width, `${name} width`).toBeGreaterThanOrEqual(TARGET - 0.5)
      // The hit area is the visible box: the button itself is hit at its centre.
      expect(await page.evaluate(([x, y]) => document.elementFromPoint(x, y)?.closest('button')?.textContent ?? '', [box.x + box.width / 2, box.y + box.height / 2]), `${name} hit`).toBe(name)
    }
    // Mirrored geometry: an 88px touch header.
    const header = await page.locator('[data-officeview-ui=grid-headers]').boundingBox()
    expect(header!.height).toBeCloseTo(2 * TARGET, 0)
    // The toolbar leaves the grid in the upper half of the screen.
    const pane = await page.getByTestId('officeview-sheet-pane').boundingBox()
    expect(pane!.y, 'sheet pane top').toBeLessThan(667 / 2)
  })
})
