/**
 * Task 8 fix round 1: sheet dialog and toolbar control polish, asserted with
 * the real stylesheet (Chromium + WebKit).
 *  - native selects draw a dropdown chevron and reserve room for it;
 *  - a disabled primary button (OK) carries no accent fill;
 *  - neighbouring dialog buttons do not touch;
 *  - the focused dialog input shows the accent focus ring in the dark theme.
 */
import { expect, test, type Locator, type Page } from '@playwright/test'
import { ensureFixtures, openViewer } from './viewer-e2e'

const css = (el: Locator, prop: string) => el.evaluate((node, p) => getComputedStyle(node).getPropertyValue(p), prop)
const accent = (page: Page) => page.locator('[data-ov-theme-root]').first().evaluate((root) => {
  const probe = document.createElement('span')
  probe.style.color = 'var(--ov-color-accent)'
  root.append(probe)
  const value = getComputedStyle(probe).color
  probe.remove()
  return value
})

test.describe('sheet control polish (browser)', () => {
  test.beforeAll(async () => { await ensureFixtures() })

  test('selects draw a chevron with room for it (browser)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    for (const select of [page.getByRole('combobox', { name: 'Worksheet' }), page.getByRole('combobox', { name: 'Address space' }), dialog.getByRole('combobox', { name: 'Filter method' })]) {
      expect(await css(select, 'background-image')).toContain('linear-gradient')
      expect(parseFloat(await css(select, 'padding-right'))).toBeGreaterThanOrEqual(20)
      expect(await css(select, 'background-repeat')).toContain('no-repeat')
    }
  })

  test('disabled OK has no accent fill and dialog buttons keep a gap (browser)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    const ok = dialog.getByRole('button', { name: 'OK' })
    const cancel = dialog.getByRole('button', { name: 'Cancel' })
    await expect(ok).toBeVisible()
    const enabledBg = await ok.evaluate((el) => (el as HTMLButtonElement).disabled ? null : getComputedStyle(el).backgroundColor)
    // Force the disabled state the dialog shows while busy or invalid.
    await ok.evaluate((el) => { (el as HTMLButtonElement).disabled = true })
    const disabledBg = await css(ok, 'background-color')
    expect(disabledBg).not.toBe(await accent(page))
    if (enabledBg) expect(disabledBg).not.toBe(enabledBg)
    const [a, b] = [await ok.boundingBox(), await cancel.boundingBox()]
    expect(b!.x - (a!.x + a!.width), 'OK to Cancel gap').toBeGreaterThanOrEqual(4)
    const asc = await dialog.getByRole('button', { name: 'Sort ascending' }).boundingBox()
    const desc = await dialog.getByRole('button', { name: 'Sort descending' }).boundingBox()
    expect(desc!.x - (asc!.x + asc!.width), 'sort buttons gap').toBeGreaterThanOrEqual(4)
  })

  test('the focused dialog input shows the accent ring in the dark theme (browser)', async ({ page }) => {
    await page.emulateMedia({ colorScheme: 'dark' })
    // A narrow pane turns the filter dropdown into the bottom sheet.
    await page.setViewportSize({ width: 600, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    await page.locator('[data-ov-theme-root]').first().evaluate((root) => root.setAttribute('data-ov-theme', 'dark'))
    await page.getByRole('button', { name: /^Filter column A/ }).click()
    expect(await page.locator('[data-officeview-ui="filter-dropdown"]').evaluate((el) => getComputedStyle(el).position)).toBe('fixed')
    const search = page.getByRole('textbox', { name: 'Search all values' })
    await search.focus()
    await expect.poll(() => css(search, 'outline-style')).toBe('solid')
    expect(await css(search, 'outline-color')).toBe(await accent(page))
  })
})
