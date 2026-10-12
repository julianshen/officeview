/**
 * S10.2 (round 4): the viewer must not add layout shift to its host page.
 * Chrome reports every element whose layout position changed between frames
 * (Layout Instability API); CLS sums those without recent input. Viewport
 * movement, header strips and the semantic grid rows/cells follow the scroll
 * window, so they are positioned with transforms (which are not layout
 * shifts) inside fixed-size containers.
 *
 * Bound: the summed shift score of shifts without recent input over open,
 * vertical scroll, header sort, filter dialog, horizontal scroll and close stays
 * at or below 0.001. Chrome rates CLS up to 0.1 "good"; a viewer that only
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

test.describe('S10.2 viewer layout stability', () => {
  test.beforeAll(async () => { await ensureFixtures() })

  test('open, scroll, sort, filter and close add no layout shift to the host page', async ({ page, browserName }) => {
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
    // A slowed main thread makes the value list arrive well after the 500 ms
    // recent-input window, as a 1M-row column does at full scale: whatever the
    // dialog moves when its values land then counts as layout shift.
    const cdp = await page.context().newCDPSession(page)
    await cdp.send('Emulation.setCPUThrottlingRate', { rate: 20 })
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await dialog.waitFor({ state: 'visible' })
    await expect(dialog.getByRole('progressbar')).toHaveCount(0, { timeout: 60_000 })
    await page.waitForTimeout(600)
    await cdp.send('Emulation.setCPUThrottlingRate', { rate: 1 })
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await dialog.waitFor({ state: 'hidden' })
    await phase('scroll-horizontal')
    await page.evaluate(() => (window as unknown as W).__viewerPerf.scrollRun({ warmup: 0, frames: 20, dx: 80, dy: 0 }))
    await phase('close')
    await page.evaluate(() => (window as unknown as W).__viewerPerf.close())
    await page.waitForTimeout(500)
    const shifts = await page.evaluate(() => (window as unknown as W).__shifts)
    const counted = shifts.filter(s => !s.recent)
    const cls = counted.reduce((sum, s) => sum + s.value, 0)
    const byPhase: Record<string, number> = {}
    for (const s of counted) byPhase[s.phase] = (byPhase[s.phase] ?? 0) + s.value
    expect(cls, `layout shift by phase ${JSON.stringify(byPhase)}; first sources ${JSON.stringify(counted.slice(0, 3).map(s => s.sources))}`).toBeLessThanOrEqual(0.001)
  })
})
