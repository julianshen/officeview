/**
 * S10.2 (round 4): the viewer must not add layout shift to its host page.
 * Chrome reports every element whose layout position changed between frames
 * (Layout Instability API); CLS sums those without recent input. Viewport
 * movement, header strips and the semantic grid rows/cells follow the scroll
 * window, so they are positioned with transforms (which are not layout
 * shifts) inside fixed-size containers.
 *
 * Bound: the summed shift score of shifts without recent input over open,
 * vertical scroll, header sort, filter dialog, horizontal scroll, zoom, sheet
 * switch and close stays at or below 0.001, in the light and dark themes
 * (Task 6: the chrome motion on those moments animates transform and opacity
 * only, so it adds no shift). Chrome rates CLS up to 0.1 "good"; a viewer that only
 * moves content by transform should register no shift at all, and 0.001
 * leaves room for sub-pixel rounding only (one 20 px row moving 1 px in a
 * 1440x900 page scores ~2e-5 per frame).
 */
import { expect, test, type Page } from '@playwright/test'
import { ensureFixtures, fixturePath } from './viewer-e2e'
import type { ViewerPerf } from '../../src/debug/viewer-perf'

type W = Window & { __viewerPerf: ViewerPerf; __shifts: { value: number; recent: boolean; phase: string; sources: string[] }[]; __phase: string }

async function boot(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const w = window as unknown as W
    w.__shifts = []
    w.__phase = 'boot'
    new PerformanceObserver(list => {
      for (const entry of list.getEntries() as unknown as { value: number; hadRecentInput: boolean; sources?: { node?: Node | null }[] }[]) {
        w.__shifts.push({ value: entry.value, recent: entry.hadRecentInput, phase: w.__phase, sources: (entry.sources ?? []).map(s => { const n = s.node, el = n && n.nodeType === 1 ? n as Element : n?.parentElement; return el ? `${el.tagName.toLowerCase()}${el.getAttribute('role') ? `[role=${el.getAttribute('role')}]` : ''}` : String(n) }) })
      }
    }).observe({ type: 'layout-shift', buffered: true })
  })
  await page.goto('/src/debug/viewer-perf.html')
  await page.waitForFunction(() => (window as unknown as W).__viewerPerf?.booted === true)
}

/** Zoom in twice and back out through the chrome; each step waits past the
 * 500 ms recent-input window so any late chrome movement would count. */
async function zoomRun(page: Page): Promise<void> {
  for (const name of ['Zoom in', 'Zoom in', 'Zoom out', 'Zoom out']) {
    const readout = await page.getByRole('button', { name: 'Reset zoom' }).textContent()
    await page.getByRole('button', { name, exact: true }).click()
    await expect(page.getByRole('button', { name: 'Reset zoom' })).not.toHaveText(readout ?? '')
    await page.waitForTimeout(600)
  }
}
const totals = (page: Page) => page.evaluate(() => {
  const counted = (window as unknown as W).__shifts.filter(s => !s.recent)
  const byPhase: Record<string, number> = {}
  for (const s of counted) byPhase[s.phase] = (byPhase[s.phase] ?? 0) + s.value
  return { cls: counted.reduce((sum, s) => sum + s.value, 0), byPhase, sources: counted.slice(0, 3).map(s => s.sources) }
})

for (const theme of ['light', 'dark'] as const) test.describe(`S10.2 viewer layout stability (${theme})`, () => {
  test.beforeAll(async () => { await ensureFixtures() })
  test.beforeEach(async ({ page }) => { await page.emulateMedia({ colorScheme: theme }) })

  test(`open, scroll, sort, filter, zoom and close add no layout shift to the host page (${theme})`, async ({ page, browserName }) => {
    test.skip(browserName !== 'chromium', 'The Layout Instability API is Chromium-only')
    await boot(page)
    const phase = (name: string) => page.evaluate(n => { (window as unknown as W).__phase = n }, name)
    await phase('open')
    await page.setInputFiles('#perf-file', fixturePath('f2-paged.xlsx'))
    await page.waitForFunction(() => (window as unknown as W).__viewerPerf.marks.completeAt !== null, undefined, { timeout: 90_000 })
    await phase('scroll-vertical')
    await page.evaluate(() => (window as unknown as W).__viewerPerf.scrollRun({ warmup: 0, frames: 30, dx: 0, dy: 60 }))
    await phase('sort')
    const header = page.getByRole('columnheader', { name: /^Column B/ })
    const before = await header.getAttribute('aria-sort')
    await page.getByRole('button', { name: 'Cycle sort column B' }).click()
    await expect.poll(() => header.getAttribute('aria-sort'), { timeout: 60_000 }).not.toBe(before)
    await phase('filter')
    // The value list is held back until well after the 500 ms recent-input
    // window (a deterministic test hook on the perf page's worker port), so
    // whatever the dialog moves when its values land counts as layout shift.
    await page.evaluate(() => (window as unknown as W).__viewerPerf.holdReplies(['valueQuery']))
    const clickedAt = Date.now()
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await dialog.waitFor({ state: 'visible' })
    await expect(dialog.getByRole('progressbar')).toHaveCount(1)
    await page.waitForTimeout(Math.max(0, 900 - (Date.now() - clickedAt)))
    const released = await page.evaluate(() => (window as unknown as W).__viewerPerf.holdReplies(null))
    expect(released, 'value-list replies were held and released').toBeGreaterThan(0)
    expect(Date.now() - clickedAt).toBeGreaterThan(500)
    // Values actually arrived: entries are listed and Select All is usable.
    await expect.poll(() => dialog.locator('[data-value-ordinal]').count(), { timeout: 60_000 }).toBeGreaterThan(0)
    await expect(dialog.getByRole('button', { name: 'Select All', exact: true })).toBeEnabled()
    await page.waitForTimeout(300)
    const filterShift = await page.evaluate(() => (window as unknown as W).__shifts.filter(s => s.phase === 'filter' && !s.recent).reduce((sum, s) => sum + s.value, 0))
    expect(filterShift, 'shift when the held value list landed').toBeLessThanOrEqual(0.001)
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await dialog.waitFor({ state: 'hidden' })
    await phase('scroll-horizontal')
    await page.evaluate(() => (window as unknown as W).__viewerPerf.scrollRun({ warmup: 0, frames: 20, dx: 80, dy: 0 }))
    await phase('zoom')
    await zoomRun(page)
    await phase('close')
    await page.evaluate(() => (window as unknown as W).__viewerPerf.close())
    await page.waitForTimeout(500)
    const { cls, byPhase, sources } = await totals(page)
    expect(cls, `layout shift by phase ${JSON.stringify(byPhase)}; first sources ${JSON.stringify(sources)}`).toBeLessThanOrEqual(0.001)
  })

  test(`sheet switch and zoom add no layout shift to the host page (${theme})`, async ({ page, browserName }) => {
    test.skip(browserName !== 'chromium', 'The Layout Instability API is Chromium-only')
    await boot(page)
    const phase = (name: string) => page.evaluate(n => { (window as unknown as W).__phase = n }, name)
    await phase('open')
    await page.setInputFiles('#perf-file', fixturePath('f7-two-sheets.xlsx'))
    await page.waitForFunction(() => (window as unknown as W).__viewerPerf.marks.completeAt !== null, undefined, { timeout: 90_000 })
    const picker = page.getByRole('combobox', { name: 'Worksheet' })
    // The theme really applies: the chrome is dark only in the dark run.
    const chrome = await page.locator('[data-officeview-ui="sheet-toolbar"]').evaluate(el => Number(/\d+/.exec(getComputedStyle(el).backgroundColor)![0]))
    expect(theme === 'dark' ? chrome < 100 : chrome > 150, `toolbar red channel ${chrome}`).toBe(true)
    for (const sheet of ['B', 'A', 'B']) {
      await phase(`sheet-${sheet}`)
      const label = await picker.locator('option', { hasText: sheet }).getAttribute('value')
      await picker.selectOption(label!)
      await expect(page.getByRole('grid')).toHaveAttribute('aria-label', `Worksheet ${sheet}`, { timeout: 30_000 })
      await page.waitForTimeout(600)
    }
    await phase('zoom')
    await zoomRun(page)
    await phase('close')
    await page.evaluate(() => (window as unknown as W).__viewerPerf.close())
    await page.waitForTimeout(500)
    const { cls, byPhase, sources } = await totals(page)
    expect(cls, `layout shift by phase ${JSON.stringify(byPhase)}; first sources ${JSON.stringify(sources)}`).toBeLessThanOrEqual(0.001)
  })
})
