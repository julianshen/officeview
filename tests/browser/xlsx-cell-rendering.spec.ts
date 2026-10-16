/**
 * Visual review stream B (PR #11) in installed Google Chrome: cell text Excel
 * shows in full must be painted, not '#'. Branded Chrome resolves a bare
 * "Calibri" on the viewer's detached tile canvases to a wide fallback face, so
 * native-1904.xlsx A2:A7 (width 19.5, bestFit) painted ############## in the
 * interactive views while the static view showed the dates (#2/#14).
 * Bundled Chromium resolves generic faces differently and never showed it, so
 * this spec runs only on channel 'chrome'.
 */
import { expect, test } from '@playwright/test'
import { ensureFixtures, openViewer } from './viewer-e2e'

test.use({ channel: 'chrome' })
test.skip(({ browserName }) => browserName !== 'chromium', 'channel chrome is Chromium-only')
test.beforeAll(async () => { await ensureFixtures() })

for (const options of [{}, { worker: true, paged: true }] as const) {
  test(`native-1904.xlsx ${'paged' in options ? 'worker paged' : 'resident'}: dates Excel shows in full are painted, not ####`, async ({ page }) => {
    await page.addInitScript(() => {
      const painted: string[] = ((window as unknown as { __painted: string[] }).__painted = [])
      const fill = CanvasRenderingContext2D.prototype.fillText
      CanvasRenderingContext2D.prototype.fillText = function (text: string, x: number, y: number, max?: number) {
        painted.push(text)
        return max === undefined ? fill.call(this, text, x, y) : fill.call(this, text, x, y, max)
      }
    })
    await openViewer(page, 'native-1904.xlsx', options)
    const painted = () => page.evaluate(() => (window as unknown as { __painted: string[] }).__painted)
    await expect.poll(async () => (await painted()).some(text => text.startsWith('2024-01-15') || /^#{10,}$/.test(text)), { timeout: 60_000 }).toBe(true)
    const texts = await painted()
    expect(texts.filter(text => /^#{10,}$/.test(text))).toEqual([])
    for (const text of ['2024-01-15 12:00:00', '2024-01-16 00:00:00', '2024-02-29 06:00:00', '1904-01-01 00:00:00', '2023-12-31 12:00:00']) expect(texts).toContain(text)
  })
}
