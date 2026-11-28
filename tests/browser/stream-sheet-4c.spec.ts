/**
 * Stage 4c Task 11, the `stream-4c` project (playwright.viewer.config.ts): these specs run against their own
 * server/test/app.ts instance on OV_STREAM_4C_PORT, so their sessions count toward that server's per-IP gate (§14:
 * 30 new sessions/min), not the `stream` or `stream-4b` project's.
 *
 * The default client (`?sheets=view`: HELLO `['sheet-view@1']`) sorts and filters a streamed sheet through the
 * server's `sheetSort` (P1), under the demo's protected (`visible-only`) policy. Fixtures (`?doc=`, server/test/
 * sheet-fixtures.ts): `sort-grid` (`sortGrid`: A1:F40, row 1 headers, B `sortGridScore(r)`, row 9 `hidden="1"`, row 12
 * `ht="0"`, column E `width="0"`) and `filter-auto-hidden` (`filterAutoHidden`: A1:C20 under a saved autoFilter that
 * hid rows 4 and 7; B holds `r`).
 *
 * The resume case of the brief is not here: `page.context().setOffline(true)` does not close the open stream socket
 * in Chromium (no close within 30 s), so it cannot drop the session. Task 9's fake-socket test covers the resume.
 *
 * Helpers as in stream-sheet-4b.spec.ts: `window.__ovSheetStats` is the harness page's frame stats.
 */
import { expect, test, type Page } from '@playwright/test'

const P = Number(process.env.OV_STREAM_4C_PORT ?? Number(process.env.OV_STREAM_PORT ?? 4181) + 20)

type Stats = { frames: number; basePaints: number; basePaintsTotal: number; pendingTiles: number; tileHits: number; zoom: number; activeSheetId: string }
const stats = (page: Page) => page.evaluate(() => (window as unknown as { __ovSheetStats?: Stats }).__ovSheetStats ?? null)
/** Settled: the latest frame shows server tiles and nothing is pending. */
async function settled(page: Page) {
  await expect.poll(async () => {
    const s = await stats(page)
    return s ? s.frames > 0 && s.pendingTiles === 0 && s.tileHits > 0 : false
  }, { timeout: 60_000 }).toBe(true)
}

test.beforeEach(async ({ page }) => {
  page.on('pageerror', (e) => { throw e })
})

test.describe('4c default client (sheet-view@1 + sheetSort)', () => {
  const openView = (page: Page, q: string) => page.goto(`/src/debug/stream-viewer.html?port=${P}&sheets=view&${q}`)
  /** `cellId` (src/xlsx/view/coordinates.ts): row × MAX_COLS + col, the suffix of a gridcell's id. */
  const MAX_COLS = 16_384
  /** The 0-based source rows of the rendered gridcells of the row at display slot `slot` (0-based). */
  const sourceRowsAt = (page: Page, slot: number) => page.evaluate((s) => [...document.querySelectorAll(`[role=row][aria-rowindex="${s + 2}"] [role=gridcell]`)]
    .map((cell) => Math.floor(Number(cell.id.replace(/^.*-cell-/, '')) / 16_384)), slot)
  /** The 0-based source rows of `filter-auto-hidden`'s A1:C20 that have a rendered gridcell in column A, sorted. */
  const shownRows = (page: Page) => page.evaluate((max) => [...new Set([...document.querySelectorAll('[role=gridcell]')]
    .map((cell) => Number(cell.id.replace(/^.*-cell-/, ''))).filter((id) => id % max === 0).map((id) => id / max))].filter((r) => r < 20).sort((a, b) => a - b), MAX_COLS)
  const statusText = (page: Page) => page.evaluate(() => document.body.textContent ?? '')
  const columnHeader = (page: Page, letter: string) => page.locator(`[role=columnheader][aria-label^="Column ${letter}"]`).first()
  /** Clicks the centre of a cell, from its column header's x and its row header's y. */
  async function click(page: Page, letter: string, row: number, shift = false) {
    const c = (await columnHeader(page, letter).boundingBox())!
    const r = (await page.getByRole('rowheader', { name: `Row ${row}`, exact: true }).boundingBox())!
    if (shift) await page.keyboard.down('Shift')
    await page.mouse.click(c.x + c.width / 2, r.y + r.height / 2)
    if (shift) await page.keyboard.up('Shift')
  }
  /** Selects `filter-auto-hidden`'s A1:C20 (18 shown rows × 3 columns), ending on A1: the extending press makes its cell
   * the active one, and the dropdown moves the active cell to its column in the same row (B1, inside the range). */
  async function selectFilterRange(page: Page) {
    await click(page, 'C', 20)
    await click(page, 'A', 1, true)
    await expect.poll(() => page.locator('[role=gridcell][aria-selected=true]').count()).toBe(54)
  }
  /** Opens column `letter`'s filter dropdown and waits until its sort buttons are live (the detection answered). */
  async function openFilter(page: Page, letter: string) {
    await page.getByRole('button', { name: new RegExp(`^Filter column ${letter} `) }).first().click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await expect(dialog.getByRole('button', { name: 'Sort ascending' })).toBeEnabled({ timeout: 30_000 })
    return dialog
  }
  /** Sorts column B descending from its filter dropdown, then waits for the sorted frame to settle. */
  async function sortBDescending(page: Page) {
    const frames = (await stats(page))!.frames
    await selectFilterRange(page)
    const dialog = await openFilter(page, 'B')
    await dialog.getByRole('button', { name: 'Sort descending' }).click()
    await expect.poll(() => statusText(page), { timeout: 30_000 }).toContain('View-only sort applied; formulas unchanged.')
    await expect.poll(async () => (await stats(page))?.frames ?? 0).toBeGreaterThan(frames)
    await settled(page)
  }
  /** `filterAutoHidden`'s largest B: r 20 (B20 = 20), 0-based. */
  const LARGEST_B_ROW = 19

  test('sort descending moves the tiles', async ({ page }) => {
    await openView(page, 'doc=filter-auto-hidden')
    await settled(page)
    expect(new Set(await sourceRowsAt(page, 1))).toEqual(new Set([1]))
    await sortBDescending(page)
    expect((await stats(page))!.pendingTiles).toBe(0)
    await expect.poll(async () => [...new Set(await sourceRowsAt(page, 1))]).toEqual([LARGEST_B_ROW])
    await expect(page.getByRole('columnheader', { name: /^Column B/ })).toHaveAttribute('aria-sort', 'descending')
    // C1: the rows the saved autoFilter hid (0-based 3 and 6) stay hidden in the sorted view.
    await expect.poll(() => shownRows(page)).toEqual(Array.from({ length: 20 }, (_, r) => r).filter((r) => r !== 3 && r !== 6))
  })

  test('filter by values hides rows and keeps author-hidden rows hidden', async ({ page }) => {
    await openView(page, 'doc=filter-auto-hidden')
    await settled(page)
    // The saved autoFilter hid rows 4 and 7 (0-based 3 and 6).
    const all = Array.from({ length: 20 }, (_, r) => r)
    await expect.poll(() => shownRows(page)).toEqual(all.filter((r) => r !== 3 && r !== 6))
    await selectFilterRange(page)
    const dialog = await openFilter(page, 'B')
    for (const v of ['10', '15']) {
      const box = dialog.getByRole('checkbox', { name: v, exact: true })
      await box.click()
      await expect(box).not.toBeChecked()
    }
    // Withheld cells list only as "(Blanks)": rows 4 and 7's B values are not listed.
    for (const v of ['4', '7']) await expect(dialog.getByRole('checkbox', { name: v, exact: true })).toHaveCount(0)
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => statusText(page), { timeout: 30_000 }).toContain('Filter applied.')
    await settled(page)
    // C1: rows 10 and 15 (0-based 9 and 14) are filtered out; rows 4 and 7 stay hidden.
    await expect.poll(() => shownRows(page)).toEqual(all.filter((r) => ![3, 6, 9, 14].includes(r)))
    expect((await stats(page))!.pendingTiles).toBe(0)
  })

  test('number filters are greyed under protection', async ({ page }) => {
    await openView(page, 'doc=sort-grid')
    await settled(page)
    const dialog = await openFilter(page, 'B')
    const option = dialog.getByRole('option', { name: 'Number filters' })
    await expect(option).toHaveAttribute('disabled', '')
    await expect(option).toHaveAttribute('title', "This filter isn't available in this viewing mode.")
  })
})
