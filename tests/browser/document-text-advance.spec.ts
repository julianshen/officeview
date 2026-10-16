/**
 * Visual review #16 / NEW-3: word gaps in static ODT/DOCX pages looked doubled
 * in Chrome ("Writer  document", "python-docx  was here!").
 *
 * Cause: layout measured text on a detached <canvas>, whose context has
 * lang="inherit" with nothing to inherit from, while pages paint on canvases
 * attached under <html lang="en">. Chrome resolves the generic sans-serif
 * family differently for the two (Helvetica vs Arial metrics on macOS), so
 * every laid-out advance was ~7% wider than the painted glyphs and the excess
 * piled up in front of each space.
 *
 * The check is engine-agnostic: every painted segment must start where the
 * previous segment on that line ends, measured by the painting context itself.
 */
import { expect, test, type Page } from '@playwright/test'
import { buildOdt, odfDecimalListStyle } from '../../src/testdata/odf-builders'
import { existsSync, readFileSync } from 'node:fs'

// Playwright's bundled Chromium resolves generic families identically for
// detached and attached canvases, so the defect only reproduces in branded
// Chrome. Use it when installed (as the visual review did).
const CHROME = ['/Applications/Google Chrome.app', '/opt/google/chrome/chrome'].some(path => existsSync(path))
test.use({ channel: async ({ browserName }, use) => { await use(browserName === 'chromium' && CHROME ? 'chrome' : undefined) } })

interface PaintedText { text: string; x: number; y: number; width: number }
declare global {
  interface Window { __paintedText?: PaintedText[] }
}

async function openDocument(page: Page, name: string, bytes: Uint8Array): Promise<PaintedText[]> {
  await page.addInitScript(() => {
    const painted: PaintedText[] = []
    window.__paintedText = painted
    const fillText = CanvasRenderingContext2D.prototype.fillText
    CanvasRenderingContext2D.prototype.fillText = function (this: CanvasRenderingContext2D, text: string, x: number, y: number, maxWidth?: number) {
      // Page canvases only: their aria-label is "page N".
      if (this.canvas.getAttribute('aria-label')?.startsWith('page ')) {
        const t = this.getTransform()
        painted.push({ text, x: t.a * x + t.e, y: t.d * y + t.f, width: t.a * this.measureText(text).width })
      }
      return maxWidth === undefined ? fillText.call(this, text, x, y) : fillText.call(this, text, x, y, maxWidth)
    }
  })
  await page.route(url => url.pathname === `/__fixtures/${name}`, route => route.fulfill({ body: Buffer.from(bytes), contentType: 'application/octet-stream' }))
  await page.goto(`/src/debug/viewer-e2e.html?${new URLSearchParams({ src: `/__fixtures/${name}` })}`)
  await page.locator('canvas[aria-label^="page "]').first().waitFor({ timeout: 60_000 })
  await expect.poll(() => page.evaluate(() => window.__paintedText?.length ?? 0)).toBeGreaterThan(5)
  return page.evaluate(() => window.__paintedText ?? [])
}

/** Largest gap/overlap between consecutive segments painted on one baseline. */
function worstAdvanceError(painted: PaintedText[]): { error: number; at: string } {
  let worst = { error: 0, at: '' }
  for (let i = 1; i < painted.length; i++) {
    const prev = painted[i - 1], cur = painted[i]
    // List markers sit in their own gutter, not in the run flow.
    if (Math.abs(prev.y - cur.y) > 0.01 || cur.x < prev.x || /^(\d+[.)]|\u2022)$/.test(prev.text)) continue
    const error = Math.abs(cur.x - (prev.x + prev.width))
    if (error > worst.error) worst = { error, at: `${JSON.stringify(prev.text)} -> ${JSON.stringify(cur.text)}` }
  }
  return worst
}

test('ODT body text advances match the painted glyph widths', async ({ page }) => {
  const odt = await buildOdt({
    extraAutoStyles: odfDecimalListStyle('DemoList'),
    paras: [
      { text: 'Writer document (ODF)' },
      { runs: [{ text: 'This .odt renders through the same layout engine as Word files: ' }, { text: 'styles cascade', bold: true }, { text: ' from ODF automatic styles.' }] },
    ],
    lists: [{ styleName: 'DemoList', items: [{ text: 'first item' }, { text: 'second item' }] }],
  })
  const painted = await openDocument(page, 'advance.odt', odt)
  expect(painted.map(p => p.text)).toContain('document')
  const worst = worstAdvanceError(painted)
  expect(worst.error, `worst advance error at ${worst.at}`).toBeLessThan(0.75)
})

// pydocx-test.docx is the review's NEW-3 file (theme heading font). The
// corpus is fetched on demand (bun scripts/fetch-corpus.ts), so skip without it.
const PYDOCX = 'corpus/pydocx-test.docx'
test('DOCX heading text advances match the painted glyph widths', async ({ page }) => {
  test.skip(!existsSync(PYDOCX), 'corpus/ not fetched')
  const painted = await openDocument(page, 'pydocx-test.docx', readFileSync(PYDOCX))
  expect(painted.map(p => p.text)).toContain('python-docx')
  const worst = worstAdvanceError(painted)
  expect(worst.error, `worst advance error at ${worst.at}`).toBeLessThan(0.75)
})
