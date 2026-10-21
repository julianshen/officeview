/**
 * Slideshow S0-7 (spec 8, 11 D2): slide backgrounds and master artwork reach the shipped paged viewer and its
 * print output in a real browser, through getPaintables' lazy appearance chunk. Real LibreOffice decks from the
 * fetched corpus (`bun scripts/fetch-corpus.ts`); each test skips when its deck is absent.
 *
 * Page: src/debug/viewer-e2e.html (the shipped <OfficeFile>). Pixels are probed at slide fractions, so the
 * checks hold at any zoom or DPR. Chromium evidence only.
 */
import { expect, test, type Page } from '@playwright/test'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const CORPUS = join(fileURLToPath(new URL('../..', import.meta.url)), 'corpus')
type Rgb = [number, number, number]
const near = (actual: number[], expected: Rgb, tolerance: number) => expected.every((v, i) => Math.abs(actual[i] - v) <= tolerance)

async function open(page: Page, file: string): Promise<void> {
  const bytes = readFileSync(join(CORPUS, file))
  await page.route(url => url.pathname === `/__fixtures/${file}`, route => route.fulfill({ body: bytes, contentType: 'application/octet-stream' }))
  await page.goto(`/src/debug/viewer-e2e.html?${new URLSearchParams({ src: `/__fixtures/${file}`, name: '1' })}`)
  await page.locator('[data-officeview-root] canvas[aria-label^="page "]').first().waitFor({ timeout: 60_000 })
}

/** RGB at slide fractions of the first page canvas, once its first point has painted something other than white. */
async function probe(page: Page, points: Array<[number, number]>): Promise<number[][]> {
  const canvas = page.locator('[data-officeview-root] canvas[aria-label^="page "]').first()
  const read = () => canvas.evaluate((el, pts) => {
    const c = el as HTMLCanvasElement, ctx = c.getContext('2d')!
    return pts.map(([fx, fy]) => [...ctx.getImageData(Math.floor(fx * (c.width - 1)), Math.floor(fy * (c.height - 1)), 1, 1).data])
  }, points)
  await expect.poll(async () => (await read())[0].slice(0, 3).join(','), { timeout: 30_000 }).not.toBe('255,255,255')
  return read()
}

test.describe('slide appearance (spec 8, D2)', () => {
  test('a corpus deck shows its designed background and master art (browser)', async ({ page }) => {
    test.skip(!existsSync(join(CORPUS, 'lo-tdf105739.pptx')) || !existsSync(join(CORPUS, 'lo-tdf123684.pptx')), 'slideshow corpus not fetched: bun scripts/fetch-corpus.ts')
    // lo-tdf105739: the slide's own linear gradient, red (FF0000) at the top left to green (00B050) at the bottom right.
    await open(page, 'lo-tdf105739.pptx')
    const [topLeft, bottomRight] = await probe(page, [[0.01, 0.01], [0.99, 0.99]])
    expect(near(topLeft, [255, 0, 0], 40), `top left ${topLeft}`).toBe(true)
    expect(near(bottomRight, [0, 176, 80], 40), `bottom right ${bottomRight}`).toBe(true)
    await page.unrouteAll({ behavior: 'ignoreErrors' })
    // lo-tdf123684: a theme picture background (bgRef 1003) under the master's artwork, whose solid red
    // rectangle sits at the top right of the slide; without enrichment the slide is blank white.
    await open(page, 'lo-tdf123684.pptx')
    const [backgroundCorner, masterRect] = await probe(page, [[0.01, 0.99], [0.885, 0.08]])
    expect(near(masterRect, [178, 22, 22], 40), `master rectangle ${masterRect}`).toBe(true)
    expect(near(backgroundCorner, [0, 0, 0], 40), `picture background corner ${backgroundCorner}`).toBe(true)
  })

  test('print output includes the background (browser)', async ({ page }) => {
    test.skip(!existsSync(join(CORPUS, 'lo-tdf105739.pptx')), 'slideshow corpus not fetched: bun scripts/fetch-corpus.ts')
    // No system print dialog: the print frame's print() is observed, never called through.
    await page.addInitScript(() => {
      const calls: Window[] = []
      Object.assign(window, { __printCalls: calls })
      window.print = () => { calls.push(window) }
      const append = Node.prototype.appendChild
      Node.prototype.appendChild = function <T extends Node>(node: T): T {
        const result = append.call(this, node) as T
        if (node instanceof HTMLIFrameElement && node.hasAttribute('data-officeview-print-frame')) node.contentWindow!.print = () => { calls.push(node.contentWindow!) }
        return result
      }
    })
    await open(page, 'lo-tdf105739.pptx')
    await page.getByRole('toolbar', { name: 'Document toolbar' }).getByRole('button', { name: 'Print', exact: true }).click()
    const image = page.frameLocator('[data-officeview-print-frame]').locator('img')
    await expect(image).toHaveCount(1)
    await expect.poll(() => page.evaluate(() => (window as unknown as { __printCalls: Window[] }).__printCalls.length)).toBe(1)
    const [topLeft, bottomRight] = await image.evaluate(async el => {
      const img = el as HTMLImageElement
      if (!img.complete) await new Promise(resolve => { img.onload = resolve })
      const c = document.createElement('canvas')
      c.width = img.naturalWidth
      c.height = img.naturalHeight
      const ctx = c.getContext('2d')!
      ctx.drawImage(img, 0, 0)
      return [[2, 2], [c.width - 3, c.height - 3]].map(([x, y]) => [...ctx.getImageData(x, y, 1, 1).data])
    })
    expect(near(topLeft, [255, 0, 0], 40), `print top left ${topLeft}`).toBe(true)
    expect(near(bottomRight, [0, 176, 80], 40), `print bottom right ${bottomRight}`).toBe(true)
  })
})
