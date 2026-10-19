/**
 * Task 6 (spec §6.4): chrome motion honours prefers-reduced-motion. With
 * reduced motion emulated, the grid popups, the paged More menu and the
 * mobile filter sheet open in their final state at once: no running
 * animation lasts longer than 1 ms, every motion element is fully opaque and
 * untransformed on the first frame, and the CSS-transition controls
 * (ov:motion-reduce:transition-none) compute transition-property: none.
 */
import { expect, test, type Locator, type Page } from '@playwright/test'
import { buildDocx } from '../../src/testdata/ooxml-builders'
import { ensureFixtures, openViewer } from './viewer-e2e'

/** One frame after opening: every motion element is final and nothing animates. */
async function settledAtOnce(page: Page, popup: Locator): Promise<void> {
  await popup.waitFor({ state: 'visible' })
  const state = await page.evaluate(() => new Promise<{ animations: number[]; elements: { name: string; opacity: string; transform: string }[]; transitions: string[] }>((resolve) => requestAnimationFrame(() => resolve({
    animations: document.getAnimations().map((a) => Number(a.effect?.getComputedTiming().duration ?? 0)),
    elements: [...document.querySelectorAll<HTMLElement>('[data-ov-motion]')].map((el) => ({ name: el.getAttribute('data-ov-motion')!, opacity: getComputedStyle(el).opacity, transform: getComputedStyle(el).transform })),
    transitions: [...document.querySelectorAll<HTMLElement>('[class*="ov:motion-reduce:transition-none"]')].map((el) => getComputedStyle(el).transitionProperty),
  }))))
  for (const duration of state.animations) expect(duration).toBeLessThanOrEqual(1)
  expect(state.elements.length).toBeGreaterThan(0)
  for (const el of state.elements) {
    expect(el.opacity, el.name).toBe('1')
    expect(['none', 'matrix(1, 0, 0, 1, 0, 0)'], `${el.name} transform`).toContain(el.transform)
  }
  for (const property of state.transitions) expect(property).toBe('none')
}

test.describe('reduced motion', () => {
  test.beforeAll(async () => { await ensureFixtures() })
  test.beforeEach(async ({ page }) => { await page.emulateMedia({ reducedMotion: 'reduce' }) })

  test('reduced motion: popups, More menu and mobile sheet open without animation', async ({ page }) => {
    // Desktop grid popup (column filter dropdown).
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    await settledAtOnce(page, page.getByRole('dialog', { name: 'Column filter' }))
    await page.getByRole('dialog', { name: 'Column filter' }).getByRole('button', { name: 'Cancel' }).click()
    await expect(page.getByRole('dialog', { name: 'Column filter' })).toBeHidden()

    // The sort dialog.
    await page.getByRole('button', { name: 'More sort options for A' }).click()
    await settledAtOnce(page, page.getByRole('dialog', { name: 'Sort worksheet' }))
    await page.keyboard.press('Escape')
    await expect(page.getByRole('dialog', { name: 'Sort worksheet' })).toBeHidden()

    // Mobile filter sheet: a narrow pane turns the dropdown into a bottom sheet.
    await page.setViewportSize({ width: 600, height: 800 })
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    const sheet = page.locator('[data-officeview-ui="filter-dropdown"]')
    await settledAtOnce(page, sheet)
    expect(await sheet.evaluate((el) => getComputedStyle(el).position)).toBe('fixed')
    expect(await sheet.getAttribute('data-ov-motion')).toBe('sheet')

    // The paged viewer's compact More menu.
    await page.setViewportSize({ width: 375, height: 667 })
    const docx = await buildDocx(Array.from({ length: 40 }, (_, i) => ({ runs: [{ text: `Line ${i + 1} of the motion test document.` }] })))
    await page.route((url) => url.pathname === '/__fixtures/Motion.docx', (route) => route.fulfill({ body: Buffer.from(docx), contentType: 'application/octet-stream' }))
    await page.goto(`/src/debug/viewer-e2e.html?${new URLSearchParams({ src: '/__fixtures/Motion.docx', name: '1' })}`)
    const bar = page.getByRole('toolbar', { name: 'Document toolbar' })
    await bar.waitFor({ timeout: 60_000 })
    await bar.getByLabel('More options').click()
    await settledAtOnce(page, page.getByRole('menu', { name: 'More options' }))
  })
})
