/**
 * Task 8 fix round 1 (browser twin of tests/ui-icon-fill.test.tsx): every
 * Lucide icon in the paged toolbar and the sheet chrome computes fill: none
 * with the real stylesheet. The only filled icon is the active filter funnel,
 * which computes the button's currentColor.
 */
import { expect, test, type Page } from '@playwright/test'
import { buildDocx } from '../../src/testdata/ooxml-builders'
import { ensureFixtures, openViewer, status } from './viewer-e2e'

interface IconFill { name: string; fill: string; color: string; label: string }
const iconFills = (page: Page) => page.evaluate(() => [...document.querySelectorAll<SVGElement>('svg[data-icon]')].map((svg) => ({
  name: svg.getAttribute('data-icon') ?? '',
  fill: getComputedStyle(svg).fill,
  color: getComputedStyle(svg).color,
  label: svg.closest('[aria-label]')?.getAttribute('aria-label') ?? '',
})) as IconFill[])

test.describe('chrome icons are outlines (browser)', () => {
  test.beforeAll(async () => { await ensureFixtures() })

  test('paged toolbar icons compute fill none (desktop and compact More menu) (browser)', async ({ page }) => {
    const docx = await buildDocx(Array.from({ length: 40 }, (_, i) => ({ runs: [{ text: `Line ${i + 1} of the icon test document.` }] })))
    await page.route((url) => url.pathname === '/__fixtures/Icons.docx', (route) => route.fulfill({ body: Buffer.from(docx), contentType: 'application/octet-stream' }))
    for (const viewport of [{ width: 1280, height: 800 }, { width: 375, height: 667 }]) {
      await page.setViewportSize(viewport)
      await page.goto(`/src/debug/viewer-e2e.html?${new URLSearchParams({ src: '/__fixtures/Icons.docx', name: '1' })}`)
      const bar = page.getByRole('toolbar', { name: 'Document toolbar' })
      await bar.waitFor({ timeout: 60_000 })
      if (viewport.width < 600) {
        await bar.getByLabel('More options').click()
        await page.getByRole('menu', { name: 'More options' }).waitFor()
      }
      const icons = await iconFills(page)
      expect(icons.length, `${viewport.width}px icons`).toBeGreaterThan(4)
      for (const icon of icons) expect(icon.fill, `${viewport.width}px ${icon.name} (${icon.label})`).toBe('none')
    }
  })

  test('sheet chrome icons compute fill none; the active filter funnel is currentColor (browser)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    const before = await iconFills(page)
    expect(before.length).toBeGreaterThan(10)
    for (const icon of before) expect(icon.fill, `${icon.name} (${icon.label})`).toBe('none')

    // Apply a value filter on D so its funnel turns active.
    await page.getByRole('button', { name: /^Filter column D/ }).click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await dialog.getByRole('button', { name: 'Deselect All', exact: true }).click()
    await dialog.getByRole('checkbox', { name: 'alpha', exact: true }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()
    await expect(status(page)).not.toContainText('Visible body 999 /', { timeout: 60_000 })

    const after = await iconFills(page)
    const active = after.filter((icon) => icon.name === 'filter' && /^Filter column D/.test(icon.label))
    expect(active).toHaveLength(1)
    expect(active[0].fill).toBe(active[0].color)
    expect(active[0].fill).not.toBe('none')
    for (const icon of after.filter((i) => !active.includes(i))) expect(icon.fill, `${icon.name} (${icon.label})`).toBe('none')
  })
})
