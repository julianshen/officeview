/**
 * PR #11 visual review, stream A: real-browser layout evidence for the
 * interactive viewer (host fill #6, popup/toolbar #13, headers #9/#21, status
 * line #18, chrome font #19, print #12). Runs on the viewer-e2e page, whose
 * host (#root) is a 100%-height flex column.
 */
import { expect, test, type Page } from '@playwright/test'
import { ensureFixtures, goTo, openViewer } from './viewer-e2e'

type Box = { left: number; right: number; top: number; bottom: number; width: number; height: number }
const rect = (page: Page, selector: string) => page.evaluate(s => document.querySelector(s)!.getBoundingClientRect().toJSON() as Box, selector)
const overlaps = (a: Box, b: Box) => a.left < b.right - 0.5 && b.left < a.right - 0.5 && a.top < b.bottom - 0.5 && b.top < a.bottom - 0.5

test.describe('visual review: interactive viewer layout', () => {
  test.beforeAll(async () => { await ensureFixtures() })

  test('#6/#18 fills a sized flex host and shows the status line inside it', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    const wrapper = await rect(page, '[data-officeview-sheet-instance]'), pane = await rect(page, '[data-testid=officeview-sheet-pane]')
    expect(wrapper.bottom).toBeGreaterThan(800 - 1)
    expect(pane.bottom).toBeGreaterThan(800 - 1)
    const status = await rect(page, '[data-officeview-ui=viewer-status]')
    expect(status.height).toBeGreaterThan(10)
    expect(status.top).toBeGreaterThanOrEqual(pane.top)
    expect(status.bottom).toBeLessThanOrEqual(pane.bottom + 0.5)
    await expect(page.locator('[data-officeview-ui=viewer-status]')).toBeVisible()
  })

  test('#19 chrome controls and labels use the viewer font, never the host serif', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const fonts = await page.evaluate(() => ['[data-officeview-ui=sheet-navigator] label', '[data-officeview-ui=sheet-navigator] button', '[data-officeview-ui=sheet-navigator] select', '[data-testid=officeview-search] input']
      .map(s => getComputedStyle(document.querySelector(s)!).fontFamily))
    for (const font of fonts) expect(font).toMatch(/^system-ui/)
  })

  test('#13 the filter popup opens under its column, inside the cell area, with OK and Cancel visible', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    await page.getByRole('button', { name: /^Filter column D/ }).click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await expect(dialog.getByRole('checkbox', { name: 'alpha', exact: true })).toBeVisible()
    const popup = await rect(page, '[data-officeview-ui=grid-popup]'), header = await page.getByRole('columnheader', { name: /^Column D/ }).boundingBox()
    const headers = await rect(page, '[data-officeview-ui=grid-headers]'), status = await rect(page, '[data-officeview-ui=viewer-status]')
    expect(Math.abs(popup.left - header!.x)).toBeLessThan(1)
    expect(popup.top).toBeGreaterThanOrEqual(headers.bottom - 0.5)
    expect(popup.bottom).toBeLessThanOrEqual(status.top + 0.5)
    for (const name of ['OK', 'Cancel']) {
      const button = await dialog.getByRole('button', { name, exact: true }).boundingBox()
      expect(button!.y + button!.height).toBeLessThanOrEqual(popup.bottom + 0.5)
      expect(button!.y).toBeGreaterThanOrEqual(popup.top)
    }
    await expect(page.getByTestId('officeview-zoom-controls')).toHaveCount(0)
  })

  test.describe('375px touch viewport', () => {
    test.use({ viewport: { width: 375, height: 667 }, hasTouch: true, isMobile: true, deviceScaleFactor: 2 })
    test('#13 toolbar controls never overlap; the filter sheet is not covered by the zoom buttons', async ({ page }) => {
      await openViewer(page, 'f2-1k.xlsx')
      const boxes = await page.evaluate(() => ['[aria-label=Worksheet]', '[aria-label="Address space"]', '[aria-label="Go To cell"]', '[data-officeview-ui=sheet-navigator] button', '[data-officeview-ui=sheet-actions]', '[data-testid=officeview-search]']
        .map(s => ({ s, ...document.querySelector(s)!.getBoundingClientRect().toJSON() })))
      for (const [i, a] of boxes.entries()) for (const b of boxes.slice(i + 1)) expect(overlaps(a, b), `${a.s} overlaps ${b.s}`).toBe(false)
      const coarse = await page.evaluate(() => matchMedia('(pointer: coarse)').matches)
      if (coarse) for (const s of ['[aria-label="Search document"]', '[aria-label="Next match"]', '[aria-label=Worksheet]']) {
        const box = await rect(page, s)
        expect(box.height, s).toBeGreaterThanOrEqual(44)
      }
      await goTo(page, 'D2')
      await page.getByRole('grid').focus()
      await page.keyboard.press('Alt+ArrowDown')
      // The bottom sheet is fixed to the viewport (its popup container has no box of its own).
      const sheet = page.locator('[data-officeview-ui=filter-dropdown]')
      await expect(sheet).toBeVisible()
      await expect(sheet.getByRole('button', { name: 'OK', exact: true })).toBeInViewport()
      await expect(page.getByTestId('officeview-zoom-controls')).toHaveCount(0)
    })
  })

  test('#9 header labels never truncate and action glyphs never crowd them (100%, 50%, far corner)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    const check = async () => {
      const result = await page.evaluate(() => {
        const labels = [...document.querySelectorAll<HTMLElement>('[data-header-column] > span')].filter(s => s.getBoundingClientRect().width > 0)
        const truncated = labels.filter(s => s.scrollWidth > s.clientWidth + 0.5).map(s => s.textContent)
        // A visible action glyph never overlaps its column's label text.
        const crowded = labels.filter(s => {
          const header = s.closest('[role=columnheader]')!, label = s.getBoundingClientRect()
          const text = document.createRange(); text.selectNodeContents(s); const ink = text.getBoundingClientRect()
          return [...header.querySelectorAll<HTMLElement>('[data-action-column]')].some(b => { const r = b.getBoundingClientRect(); return r.width > 1 && r.left < ink.right - 0.5 && ink.left < r.right && label.width > 0 })
        }).map(s => s.textContent)
        const rows = [...document.querySelectorAll<HTMLElement>('[data-header-row] > span')].filter(s => s.scrollWidth > s.clientWidth + 0.5).map(s => s.textContent)
        return { count: labels.length, truncated, crowded, rows }
      })
      expect(result.count).toBeGreaterThan(0)
      expect(result.truncated).toEqual([])
      expect(result.crowded).toEqual([])
      expect(result.rows).toEqual([])
    }
    await check()
    for (let i = 0; i < 3; i++) await page.getByRole('button', { name: 'Zoom out' }).click()
    await expect(page.getByRole('button', { name: 'Reset zoom' })).toHaveText(/50%|25%/)
    await check()
    // Keyboard reach: F2 control mode still reaches a compact column's actions.
    await page.getByRole('grid').focus()
    await page.keyboard.press('F2')
    const sort = page.getByRole('button', { name: 'Cycle sort column A', exact: true })
    await expect(sort).toHaveCount(1)
    await sort.focus()
    expect((await sort.boundingBox())!.width).toBeGreaterThan(4) // revealed on focus
    await page.getByRole('button', { name: 'Reset zoom' }).click()
    await goTo(page, 'XFD1048576')
    await expect(page.getByRole('rowheader', { name: 'Row 1048576' })).toBeVisible()
    await check()
  })

  test('#21 header labels meet WCAG AA against the header strip', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const ratios = await page.evaluate(() => {
      const parse = (c: string) => (c.match(/[\d.]+/g) ?? []).slice(0, 3).map(Number)
      const lum = (c: number[]) => { const [r, g, b] = c.map(v => v / 255).map(v => v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4); return 0.2126 * r + 0.7152 * g + 0.0722 * b }
      const ratio = (a: number[], b: number[]) => { const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x); return (hi + 0.05) / (lo + 0.05) }
      const strip = parse(getComputedStyle(document.querySelector('[data-officeview-ui=grid-headers]')!).backgroundColor)
      return [...document.querySelectorAll('[data-header-column], [data-header-row]')].slice(0, 20).map(el => ratio(parse(getComputedStyle(el).color), strip))
    })
    expect(ratios.length).toBeGreaterThan(0)
    for (const r of ratios) expect(r).toBeGreaterThanOrEqual(4.5)
  })

  test('#12 print: white page, 2x canvases, a titled page per sheet, controls hidden', async ({ page, browserName }) => {
    await page.addInitScript(() => { (window as unknown as { print: () => void }).print = () => { (window as unknown as { __printed: number }).__printed = 1 } })
    await openViewer(page, 'f7-two-sheets.xlsx')
    await page.getByRole('button', { name: 'Print', exact: true }).click()
    await page.waitForFunction(() => (window as unknown as { __printed?: number }).__printed === 1)
    await page.emulateMedia({ media: 'print' })
    await expect(page.getByTestId('officeview-search')).toBeHidden()
    await expect(page.getByTestId('officeview-zoom-controls')).toBeHidden()
    const facts = await page.evaluate(() => ({
      background: getComputedStyle(document.querySelector('[data-officeview-root]')!).backgroundColor,
      canvases: [...document.querySelectorAll<HTMLCanvasElement>('[data-testid=officeview-sheet-print-pages] canvas')].map(c => ({ ratio: c.width / c.getBoundingClientRect().width, css: c.getBoundingClientRect().width })),
      titles: [...document.querySelectorAll('[data-officeview-print-sheet-title]')].map(h => h.textContent),
    }))
    expect(facts.background).toMatch(/rgb\(255, 255, 255\)|rgba\(0, 0, 0, 0\)/)
    expect(facts.titles).toEqual(['A', 'B'])
    for (const c of facts.canvases) expect(c.ratio).toBeGreaterThanOrEqual(2 - 1e-6)
    if (browserName === 'chromium') {
      const pdf = await page.pdf({ format: 'A4' })
      const pages = (pdf.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
      expect(pages).toBeGreaterThanOrEqual(2)
    }
    // Browser-menu print (nothing prepared): the notice, with no floating control over it.
    await page.emulateMedia({ media: 'screen' })
    await page.reload()
    await expect(page.getByRole('grid')).toBeVisible({ timeout: 60_000 })
    await page.emulateMedia({ media: 'print' })
    await expect(page.getByTestId('officeview-sheet-print-notice')).toBeVisible()
    await expect(page.getByTestId('officeview-search')).toBeHidden()
  })
})
