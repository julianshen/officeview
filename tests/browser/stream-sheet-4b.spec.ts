/**
 * Stage 4b Task 12, the `stream-4b` project (playwright.viewer.config.ts): these specs run against their own
 * server/test/app.ts instance on OV_STREAM_4B_PORT, so their sessions count toward that server's per-IP gate (§14:
 * 30 new sessions/min), not the `stream` project's. One browser runs the whole `stream` project from one IP, and its
 * busiest minute already holds 30 sessions without these specs.
 *
 * Helpers as in stream-sheet.spec.ts: `window.__ovSheetStats` is the harness page's frame stats.
 */
import { expect, test, type Page } from '@playwright/test'

const P = Number(process.env.OV_STREAM_4B_PORT ?? Number(process.env.OV_STREAM_PORT ?? 4181) + 10)

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

/**
 * Stage 4b Task 12: the default client (`?sheets=view`: no `streamSheets` option, so the HELLO carries `sheet-view@1`)
 * selects, finds, copies and resizes on a streamed sheet. Fixtures (`?doc=`): `text-grid` (server/test/sheet-fixtures.ts
 * `textGrid`: 30 × 8 of mixed content, row 7 and column F hidden, A5:C6 merged, frozen 1 × 1, a picture covering
 * `NEEDLE-D20`) and `slid-under-anchor` (`SLID-E1` in E1, which a 150 px shape at B1 covers once column B is 8 px).
 */
test.describe('4b default client (sheet-view@1)', () => {
  const openView = (page: Page, q: string) => page.goto(`/src/debug/stream-viewer.html?port=${P}&sheets=view&${q}`)
  /** `cellId` (src/xlsx/view/coordinates.ts): row × MAX_COLS + col, the suffix of a gridcell's id. */
  const cellIdOf = (row: number, col: number) => String(row * 16_384 + col)
  const activeId = (page: Page) => page.evaluate(() => document.querySelector('[role=grid]')?.getAttribute('aria-activedescendant')?.replace(/^.*-cell-/, '') ?? null)
  const selectedCells = (page: Page) => page.locator('[role=gridcell][aria-selected=true]').count()
  const columnHeader = (page: Page, letter: string) => page.locator(`[role=columnheader][aria-label^="Column ${letter}"]`).first()
  const width = async (page: Page, letter: string) => (await columnHeader(page, letter).boundingBox())!.width
  const searchStatus = (page: Page) => page.getByTestId('officeview-search-status')
  /** The centre of a cell, from its column header's x and its row header's y. */
  async function cellCenter(page: Page, letter: string, row: number) {
    const c = (await columnHeader(page, letter).boundingBox())!
    const r = (await page.getByRole('rowheader', { name: `Row ${row}`, exact: true }).boundingBox())!
    return { x: c.x + c.width / 2, y: r.y + r.height / 2 }
  }
  async function click(page: Page, letter: string, row: number, shift = false) {
    const at = await cellCenter(page, letter, row)
    if (shift) await page.keyboard.down('Shift')
    await page.mouse.click(at.x, at.y)
    if (shift) await page.keyboard.up('Shift')
  }
  /** A real mouse drag on column `index`'s right edge, by `dx` CSS px. */
  async function dragColumnEdge(page: Page, index: number, dx: number) {
    const box = (await page.locator(`[data-resize-column="${index}"]`).first().boundingBox())!
    const edge = { x: box.x + box.width / 2, y: box.y + box.height / 2 }
    await page.mouse.move(edge.x, edge.y)
    await page.mouse.down()
    await page.mouse.move(edge.x + dx / 2, edge.y, { steps: 3 })
    await page.mouse.move(edge.x + dx, edge.y, { steps: 3 })
    await page.mouse.up()
  }

  test('shift+click selects a range', async ({ page }) => {
    await openView(page, 'doc=text-grid')
    await settled(page)
    // The grid has no press-and-move selection, locally or on streams. B2:D4 avoids hidden row 7, column F and A5:C6.
    await click(page, 'B', 2)
    await expect.poll(() => selectedCells(page)).toBe(1)
    await click(page, 'D', 4, true)
    await expect.poll(() => selectedCells(page)).toBe(9)
    // As locally, the extending press makes its cell the active one.
    expect(await activeId(page)).toBe(cellIdOf(3, 3))
  })

  test('find moves to a match and shows no hit text', async ({ page }) => {
    await openView(page, 'doc=text-grid&a11y=off')
    await settled(page)
    const before = await activeId(page)
    const box = page.getByRole('searchbox', { name: 'Search document' })
    await box.fill('needle')
    await expect(searchStatus(page)).toHaveText(/^1 of \d+$/, { timeout: 30_000 })
    // The first hit in WorkbookFind's order on Grid1, never the covered NEEDLE-D20.
    const servable = [[1, 1], [9, 2], [4, 0], [24, 4], [29, 7]].map(([r, c]) => cellIdOf(r!, c!))
    await expect.poll(() => activeId(page)).not.toBe(before)
    expect(servable).toContain(await activeId(page))
    // Positions only: no element's text or attribute names the hit, except the search box's own value.
    const leaks = await page.evaluate((q) => {
      const out: string[] = []
      for (const el of document.body.querySelectorAll('*')) {
        if (el instanceof HTMLInputElement && el.type === 'search') continue
        if ((el.textContent ?? '').toLowerCase().includes(q)) { if (![...el.children].some((c) => (c.textContent ?? '').toLowerCase().includes(q))) out.push(el.outerHTML.slice(0, 120)) }
        for (const a of el.attributes) if (a.value.toLowerCase().includes(q)) out.push(`${el.tagName}[${a.name}]`)
      }
      return out
    }, 'needle')
    expect(leaks).toEqual([])
  })

  test('copy puts servable TSV on the clipboard', async ({ page, context, baseURL }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: baseURL })
    await openView(page, 'doc=text-grid&copy=full')
    await settled(page)
    await page.evaluate(() => navigator.clipboard.writeText('SENTINEL'))
    // D19:E21: D holds r × 100, E holds `E{r}`; D20 (`NEEDLE-D20`) is under the picture, so its field is empty.
    await click(page, 'D', 19)
    await click(page, 'E', 21, true)
    await expect.poll(() => selectedCells(page)).toBe(6)
    await page.getByRole('button', { name: 'Copy selection' }).first().click()
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText()), { timeout: 30_000 }).toBe('1900\tE19\n\tE20\n2100\tE21')
  })

  test('resizing a column moves the tiles', async ({ page }) => {
    await openView(page, 'doc=text-grid')
    await settled(page)
    const before = await width(page, 'B')
    const frames = (await stats(page))!.frames
    await dragColumnEdge(page, 1, 60)
    await expect.poll(() => width(page, 'B'), { timeout: 30_000 }).toBeCloseTo(before + 60, 0)
    await expect.poll(async () => (await stats(page))?.frames ?? 0).toBeGreaterThan(frames)
    await settled(page)
    expect((await stats(page))!.pendingTiles).toBe(0)
    expect(await width(page, 'B')).toBeCloseTo(before + 60, 0)
  })

  test('a slid-under cell stops answering after resize', async ({ page }) => {
    await openView(page, 'doc=slid-under-anchor')
    await settled(page)
    const box = page.getByRole('searchbox', { name: 'Search document' })
    await box.fill('SLID-E1')
    await expect(searchStatus(page)).toHaveText('1 of 1', { timeout: 30_000 })
    await expect.poll(() => activeId(page)).toBe(cellIdOf(0, 4))
    await box.fill('')
    await click(page, 'A', 1)
    await expect.poll(() => activeId(page)).toBe(cellIdOf(0, 0))
    // Column B to 8 px: E1 slides to x 200..264, under the shape (x 64..214).
    await dragColumnEdge(page, 1, 8 - await width(page, 'B'))
    await expect.poll(() => width(page, 'B'), { timeout: 30_000 }).toBeCloseTo(8, 0)
    await settled(page)
    await box.fill('SLID-E1')
    await expect(searchStatus(page)).toHaveText('no matches', { timeout: 30_000 })
    expect(await activeId(page)).toBe(cellIdOf(0, 0))
  })
})
