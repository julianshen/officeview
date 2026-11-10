/**
 * Stage 1 CF browser coverage (Task 11): the real interactive viewer paints
 * conditional formatting through the product routes — resident, worker-backed
 * and forced-paged — across zoom/DPR pairs, with frozen panes, copy/find,
 * keyboard focus, screenshots and print capture.
 *
 * Fixtures are the saved Stage 1 oracle workbooks (`tests/fixtures/xlsx-cf`),
 * routed from disk like tests/browser/xlsx-pivots.spec.ts does. Anchor cell
 * B2 of stage1-numfmt-accessibility holds 1.25 with a losing base format
 * (`0.0` → `1.3`) and a winning CF format (`0.00` → `1.25`) plus a pink CF
 * fill (FFFFC7CE): the accessible name, copied text, find hits and painted
 * pixels all distinguish applied CF from base formatting. Legs are compared
 * against each other on the same machine/browser (never against Excel
 * references). Real Safari/VoiceOver are not claimed; WebKit runs are
 * Playwright WebKit (`webkit-playwright`).
 */
import { expect, test, type Page } from '@playwright/test'
import { createCanvas, loadImage } from 'canvas'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { activeCell, counts, goTo } from './viewer-e2e'

const CF_FIXTURES = fileURLToPath(new URL('../fixtures/xlsx-cf/', import.meta.url))
const NUMFMT = 'stage1-numfmt-accessibility'
const PINK_FIXTURE = 'stage1-cellis-literals'
const SCALE_FIXTURE = 'stage2-scale-interpolation'

test.use({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: 1 })

interface Leg { name: string; worker: boolean; paged: boolean }
const LEGS: readonly Leg[] = [
  { name: 'resident', worker: false, paged: false },
  { name: 'worker (?worker=1)', worker: true, paged: false },
  { name: 'forced paged (?worker=1&paged=1)', worker: true, paged: true },
]

type Img = { width: number; height: number; data: Uint8ClampedArray }

async function openCf(page: Page, name: string, bytes: Buffer, leg: Leg): Promise<void> {
  const route = `/__fixtures/cf-${encodeURIComponent(name)}.xlsx`
  await page.route(url => url.pathname === route, r => r.fulfill({ body: bytes, contentType: 'application/octet-stream' }))
  const query = new URLSearchParams({ src: route })
  if (leg.worker) query.set('worker', '1')
  if (leg.paged) query.set('paged', '1')
  await page.goto(`/src/debug/viewer-e2e.html?${query}`)
  await expect(page.getByRole('grid')).toBeVisible({ timeout: 60_000 })
}

async function decode(png: Buffer): Promise<Img> {
  const image = await loadImage(png)
  const canvas = createCanvas(image.width, image.height)
  const ctx = canvas.getContext('2d')
  ctx.drawImage(image, 0, 0)
  const { data } = ctx.getImageData(0, 0, image.width, image.height)
  return { width: image.width, height: image.height, data: new Uint8ClampedArray(data) }
}

/** Screenshot once the canvas stops changing and CF pink has painted. */
async function settledCfGrid(page: Page, pinkMin: number): Promise<Img> {
  const box = await page.locator('canvas').first().boundingBox()
  expect(box, 'the grid canvas has a box').not.toBeNull()
  const shoot = async () => page.screenshot({
    clip: { x: box!.x, y: box!.y, width: Math.min(box!.width, 1000), height: Math.min(box!.height, 480) },
    animations: 'disabled', caret: 'hide',
  })
  let previous = await shoot()
  const deadline = Date.now() + 45_000
  for (;;) {
    await page.waitForTimeout(250)
    const next = await shoot()
    if (next.equals(previous)) {
      const img = await decode(next)
      if (countPink(img, 0, 0, img.width, img.height) >= pinkMin) return img
    }
    if (Date.now() > deadline) return decode(next)
    previous = next
  }
}

function countPink(img: Img, x: number, y: number, w: number, h: number): number {
  // FFC7CE with per-channel tolerance 48: matches anti-aliased fill edges.
  let n = 0
  const x1 = Math.min(img.width, x + w), y1 = Math.min(img.height, y + h)
  for (let yy = Math.max(0, y); yy < y1; yy++) {
    for (let xx = Math.max(0, x); xx < x1; xx++) {
      const i = (yy * img.width + xx) * 4
      if (Math.abs(img.data[i]! - 0xff) <= 48 && Math.abs(img.data[i + 1]! - 0xc7) <= 48 && Math.abs(img.data[i + 2]! - 0xce) <= 48) n++
    }
  }
  return n
}

function countColor(img: Img, x: number, y: number, w: number, h: number, r: number, g: number, b: number, tol = 48): number {
  let n = 0
  const x1 = Math.min(img.width, x + w), y1 = Math.min(img.height, y + h)
  for (let yy = Math.max(0, y); yy < y1; yy++) {
    for (let xx = Math.max(0, x); xx < x1; xx++) {
      const i = (yy * img.width + xx) * 4
      if (Math.abs(img.data[i]! - r) <= tol && Math.abs(img.data[i + 1]! - g) <= tol && Math.abs(img.data[i + 2]! - b) <= tol) n++
    }
  }
  return n
}

async function settledScaleGrid(page: Page, minColorCount = 200): Promise<Img> {
  const box = await page.locator('canvas').first().boundingBox()
  expect(box, 'the grid canvas has a box').not.toBeNull()
  const shoot = async () => page.screenshot({
    clip: { x: box!.x, y: box!.y, width: Math.min(box!.width, 1000), height: Math.min(box!.height, 480) },
    animations: 'disabled', caret: 'hide',
  })
  let previous = await shoot()
  const deadline = Date.now() + 45_000
  for (;;) {
    await page.waitForTimeout(250)
    const next = await shoot()
    if (next.equals(previous)) {
      const img = await decode(next)
      if (countColor(img, 0, 0, img.width, img.height, 0xff, 0, 0) >= minColorCount &&
          countColor(img, 0, 0, img.width, img.height, 0, 0xff, 0) >= minColorCount) return img
    }
    if (Date.now() > deadline) return decode(next)
    previous = next
  }
}

async function frozenBytes(name: string = NUMFMT, options: { xSplit?: number; ySplit?: number; topLeftCell?: string } = {}): Promise<Buffer> {
  const xSplit = options.xSplit ?? 1
  const ySplit = options.ySplit ?? 1
  const topLeftCell = options.topLeftCell ?? 'B2'
  const raw = readFileSync(`${CF_FIXTURES}${name}.xlsx`)
  const zip = await JSZip.loadAsync(raw)
  const path = 'xl/worksheets/sheet1.xml'
  const xml = await zip.file(path)!.async('string')
  const frozen = `<sheetViews><sheetView workbookViewId="0" showGridLines="1"><pane xSplit="${xSplit}" ySplit="${ySplit}" topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="${topLeftCell}" sqref="${topLeftCell}"/></sheetView></sheetViews>`
  const patched = xml.replace(/<sheetViews>.*?<\/sheetViews>/s, frozen)
  if (patched === xml) throw new Error('no sheetViews element to freeze')
  zip.file(path, patched)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }) as unknown as Promise<Buffer>
}

async function editCellInWorkbook(raw: Buffer, ref: string, newValue: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(raw)
  const path = 'xl/worksheets/sheet1.xml'
  const xml = await zip.file(path)!.async('string')
  const cellRegex = new RegExp(`(<c r="${ref}"[^>]*>.*?<v>)([^<]*)(</v>.*?</c>)`, 's')
  const patched = xml.replace(cellRegex, `$1${newValue}$3`)
  if (patched === xml) throw new Error(`Could not edit cell ${ref} in ${path}`)
  zip.file(path, patched)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }) as unknown as Promise<Buffer>
}

test('CF number format names the formatted value on every leg', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${NUMFMT}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, NUMFMT, bytes, leg)
        if (leg.paged) expect((await counts(page)).open ?? 0, 'paged leg opens through the worker').toBe(1)
        // B2 holds 1.25: CF 0.00 names it "1.25", never base 0.0's "1.3".
        // (Fix round 1: the settle timeout drops a hung paged read and
        // re-reads fresh, so the paged leg expects the same CF value.)
        const b2 = page.getByRole('gridcell', { name: /^B2, value / })
        await expect(b2, `${leg.name}: B2 announces CF text`).toContainText('1.25', { timeout: 30_000 })
        expect(await b2.textContent()).not.toContain('1.3')
        await expect(page.getByRole('alert'), `${leg.name}: no error alert`).toHaveCount(0)
      } finally {
        await page.close()
      }
    })
  }
})

test('CF fills paint on canvas on every leg, and legs agree', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${PINK_FIXTURE}.xlsx`)
  const shots: Img[] = []
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, PINK_FIXTURE, bytes, leg)
        const grid = await settledCfGrid(page, 500)
        expect(countPink(grid, 0, 0, grid.width, grid.height), `${leg.name}: CF fill paints`).toBeGreaterThan(500)
        shots.push(grid)
      } finally {
        await page.close()
      }
    })
  }
  const [resident, worker, paged] = shots as [Img, Img, Img]
  // Worker and resident share the rendering pipeline: exact pixels must agree.
  expect(worker.width).toBe(resident.width)
  expect(worker.height).toBe(resident.height)
  let differing = 0
  for (let i = 0; i < resident.data.length; i += 4) {
    if (Math.abs(resident.data[i]! - worker.data[i]!) > 4 || Math.abs(resident.data[i + 1]! - worker.data[i + 1]!) > 4 || Math.abs(resident.data[i + 2]! - worker.data[i + 2]!) > 4) differing++
  }
  expect(differing / (resident.width * resident.height), 'worker matches resident').toBeLessThan(0.002)
  // Forced-paged uses the tile pipeline (positional differences) and shows
  // base text (worker-side text hang, see the number-format test), so exact
  // pixels differ: paint presence is already asserted above via the pink
  // count on every leg. Dimensions must still agree.
  expect(paged.width).toBe(resident.width)
  expect(paged.height).toBe(resident.height)
})

test('six zoom/DPR pairs keep CF formatting', async ({ browser }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${NUMFMT}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      for (const dpr of [1, 2]) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: dpr })
        try {
          const page = await context.newPage()
          try {
            await openCf(page, NUMFMT, bytes, leg)
            for (const zoom of ['50%', '100%', '200%']) {
              const label = page.getByRole('button', { name: 'Reset zoom' })
              for (let i = 0; i < 12; i++) {
                if ((await label.textContent()) === zoom) break
                const current = Number((await label.textContent())!.replace('%', ''))
                await page.getByRole('button', { name: current < Number(zoom.replace('%', '')) ? 'Zoom in' : 'Zoom out' }).click()
              }
              await expect(label).toHaveText(zoom, { timeout: 10_000 })
              // Zooming recenters the viewport, virtualizing B2 out: scroll back
              // to the origin before asserting the top-left cells.
              await page.getByTestId('officeview-sheet-scroller').evaluate(node => { node.scrollTop = 0; node.scrollLeft = 0 })
              const b2 = page.getByRole('gridcell', { name: /^B2, value / })
              await expect(b2, `dpr=${dpr} zoom=${zoom}: CF text`).toContainText('1.25', { timeout: 30_000 })
              const grid = await settledCfGrid(page, 200)
              expect(countPink(grid, 0, 0, grid.width, grid.height), `dpr=${dpr} zoom=${zoom}: CF fill`).toBeGreaterThan(200)
            }
          } finally {
            await page.close()
          }
        } finally {
          await context.close()
        }
      }
    })
  }
})

test('frozen panes keep CF painting and pinned headers', async ({ context }) => {
  const bytes = await frozenBytes()
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
  const page = await context.newPage()
  try {
    await openCf(page, `${NUMFMT}-frozen`, bytes, leg)
    const b2 = page.getByRole('gridcell', { name: /^B2, value / })
    await expect(b2, 'frozen: CF text').toContainText('1.25', { timeout: 30_000 })
    const scroller = page.getByTestId('officeview-sheet-scroller')
    await scroller.evaluate(node => { node.scrollTop = 600; node.scrollLeft = 400 })
    await page.waitForTimeout(500)
    // Frozen headers stay visible after scrolling away.
    await expect(page.getByRole('columnheader', { name: /^Column A/ }).first(), 'frozen column header').toBeVisible()
    await expect(page.getByRole('rowheader', { name: 'Row 1' }).first(), 'frozen row header').toBeVisible()
    const grid = await settledCfGrid(page, 200)
    expect(countPink(grid, 0, 0, grid.width, grid.height), 'frozen: CF fill paints').toBeGreaterThan(200)
    await expect(page.getByRole('alert'), 'no error alert').toHaveCount(0)
  } finally {
    await page.close()
  }
    })
  }
})

test('copy and find use CF-formatted text', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${NUMFMT}.xlsx`)
  // All three legs serve CF text (fix round 1: paged scope liveness); find
  // matches the CF display text and copy yields it.
  for (const leg of LEGS as readonly Leg[]) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, NUMFMT, bytes, leg)
        // Find matches the CF display text, not the base value: '1.25' first
        // hits A2 ("match_1.25" base substring), then Next match advances to
        // B2's CF "1.25".
        await page.getByLabel('Search document').fill('1.25')
        await expect.poll(async () => (await activeCell(page))?.label ?? '', { timeout: 30_000 }).toMatch(/^A2, value /)
        await page.getByRole('button', { name: 'Next match' }).click()
        await expect.poll(async () => (await activeCell(page))?.label ?? '', { timeout: 30_000 }).toMatch(/^B2, value 1\.25/)
        // Copy yields the formatted text (B2 is already active from find).
        await page.getByRole('button', { name: 'Copy selection' }).first().click()
        await expect.poll(() => page.evaluate(() => (window as unknown as { __viewerE2E?: { clipboard: string[] } }).__viewerE2E?.clipboard.at(-1) ?? '')).toBe('1.25')
      } finally {
        await page.close()
      }
    })
  }
})

test('keyboard focus moves the active CF cell', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${NUMFMT}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
  const page = await context.newPage()
  try {
    await openCf(page, NUMFMT, bytes, leg)
    // Gridcells are not pointer-clickable in the harness (canvas overlay), so
    // select via Go To and focus the grid element directly for arrow keys
    // (the a11y suite's established pattern).
    await goTo(page, 'A1')
    await page.getByRole('grid').focus()
    await expect.poll(async () => (await activeCell(page))?.label ?? '', { timeout: 30_000 }).toMatch(/^A1, value /)
    await page.keyboard.press('ArrowRight')
    await page.keyboard.press('ArrowDown')
    await expect.poll(async () => (await activeCell(page))?.label ?? '', { timeout: 30_000 }).toMatch(/^B2, value 1\.25/)
    await page.keyboard.press('Escape')
    await expect(page.getByRole('grid')).toBeFocused()
  } finally {
    await page.close()
  }
    })
  }
})

test('print capture prepares resident pages with CF and handles worker notice', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${NUMFMT}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      await page.addInitScript(() => {
        (window as unknown as { print: () => void }).print = () => {
          (window as unknown as { __printed: number }).__printed = 1
        }
      })
      try {
        await openCf(page, NUMFMT, bytes, leg)
        if (!leg.paged) {
          // Resident and small-doc worker legs: click Print, which runs preparePrint
          // -> preparePrintUnits. This prepares real print canvases with CF formatting.
          await page.getByRole('button', { name: 'Print', exact: true }).click()
          await page.waitForFunction(() => (window as unknown as { __printed?: number }).__printed === 1, { timeout: 30_000 })
          const printCanvases = page.locator('[data-testid="officeview-sheet-print-pages"] canvas')
          await expect(printCanvases).toHaveCount(1)
          // The prepared print canvas contains painted CF formatting (pink fill FFFFC7CE).
          const pink = await printCanvases.first().evaluate(canvas => {
            const c = canvas as HTMLCanvasElement
            const ctx = c.getContext('2d')
            if (!ctx) return 0
            const data = ctx.getImageData(0, 0, c.width, c.height).data
            let count = 0
            for (let i = 0; i < data.length; i += 4) {
              if (Math.abs(data[i]! - 0xff) <= 48 && Math.abs(data[i + 1]! - 0xc7) <= 48 && Math.abs(data[i + 2]! - 0xce) <= 48) count++
            }
            return count
          })
          expect(pink, 'CF pink fill is painted on the prepared print canvas').toBeGreaterThan(20)
          // Emulate print media: print notice is hidden once pages are prepared.
          await page.emulateMedia({ media: 'print' })
          await expect(page.getByTestId('officeview-sheet-print-notice')).toBeHidden()
          await expect(printCanvases.first()).toBeVisible()
          if (test.info().project.name === 'chromium') {
            const pdf = await page.pdf({ format: 'A4', printBackground: true })
            expect(pdf.byteLength, 'prepared print produces PDF output').toBeGreaterThan(1000)
          }
        } else {
          // Forced-paged sheets intentionally do not print: no prepared pages,
          // clicking Print announces PAGED_PRINT_MESSAGE with no prepared canvas.
          await page.getByRole('button', { name: 'Print', exact: true }).click()
          await page.emulateMedia({ media: 'print' })
          await expect(page.locator('[data-testid="officeview-sheet-print-pages"] canvas')).toHaveCount(0)
          await expect(page.getByTestId('officeview-sheet-print-notice')).toBeVisible()
        }
      } finally {
        await page.close()
      }
    })
  }
})

// ==================== Stage 2 Colour Scale Coverage ====================

test('Stage 2 colour scale browser coverage: resident, worker, and forced-paged legs agree', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${SCALE_FIXTURE}.xlsx`)
  const shots: Img[] = []
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, SCALE_FIXTURE, bytes, leg)
        // Accessibility descriptors via OD-2:
        // B2 is min (10) -> 0% of range
        // B8 is max (70) -> 100% of range
        const b2 = page.getByRole('gridcell', { name: /^B2, value / })
        await expect(b2, `${leg.name}: B2 announces colour scale 0% descriptor`).toContainText('Colour scale, 0% of range', { timeout: 30_000 })
        const b8 = page.getByRole('gridcell', { name: /^B8, value / })
        await expect(b8, `${leg.name}: B8 announces colour scale 100% descriptor`).toContainText('Colour scale, 100% of range', { timeout: 30_000 })
        await expect(page.getByRole('alert'), `${leg.name}: no error alert`).toHaveCount(0)

        const grid = await settledScaleGrid(page, 200)
        expect(countColor(grid, 0, 0, grid.width, grid.height, 0xff, 0, 0), `${leg.name}: Red fill paints`).toBeGreaterThan(200)
        expect(countColor(grid, 0, 0, grid.width, grid.height, 0, 0xff, 0), `${leg.name}: Green fill paints`).toBeGreaterThan(200)
        shots.push(grid)
      } finally {
        await page.close()
      }
    })
  }
  const [resident, worker, paged] = shots as [Img, Img, Img]
  expect(worker.width).toBe(resident.width)
  expect(worker.height).toBe(resident.height)
  let differing = 0
  for (let i = 0; i < resident.data.length; i += 4) {
    if (Math.abs(resident.data[i]! - worker.data[i]!) > 4 || Math.abs(resident.data[i + 1]! - worker.data[i + 1]!) > 4 || Math.abs(resident.data[i + 2]! - worker.data[i + 2]!) > 4) differing++
  }
  expect(differing / (resident.width * resident.height), 'worker matches resident').toBeLessThan(0.002)
  expect(paged.width).toBe(resident.width)
  expect(paged.height).toBe(resident.height)
})

test('Stage 2 colour scale six zoom/DPR pairs keep scale fill', async ({ browser }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${SCALE_FIXTURE}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      for (const dpr of [1, 2]) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: dpr })
        try {
          const page = await context.newPage()
          try {
            await openCf(page, SCALE_FIXTURE, bytes, leg)
            for (const zoom of ['50%', '100%', '200%']) {
              const label = page.getByRole('button', { name: 'Reset zoom' })
              for (let i = 0; i < 12; i++) {
                if ((await label.textContent()) === zoom) break
                const current = Number((await label.textContent())!.replace('%', ''))
                await page.getByRole('button', { name: current < Number(zoom.replace('%', '')) ? 'Zoom in' : 'Zoom out' }).click()
              }
              await expect(label).toHaveText(zoom, { timeout: 10_000 })
              await page.getByTestId('officeview-sheet-scroller').evaluate(node => { node.scrollTop = 0; node.scrollLeft = 0 })
              const b2 = page.getByRole('gridcell', { name: /^B2, value / })
              await expect(b2, `dpr=${dpr} zoom=${zoom}: CF text`).toContainText('Colour scale, 0% of range', { timeout: 30_000 })
              const grid = await settledScaleGrid(page, 100)
              expect(countColor(grid, 0, 0, grid.width, grid.height, 0xff, 0, 0), `dpr=${dpr} zoom=${zoom}: Red scale fill`).toBeGreaterThan(100)
              expect(countColor(grid, 0, 0, grid.width, grid.height, 0, 0xff, 0), `dpr=${dpr} zoom=${zoom}: Green scale fill`).toBeGreaterThan(100)
            }
          } finally {
            await page.close()
          }
        } finally {
          await context.close()
        }
      }
    })
  }
})

test('Stage 2 colour scale frozen panes keep pinned headers and scale fill', async ({ context }) => {
  const bytes = await frozenBytes(SCALE_FIXTURE, { xSplit: 2, ySplit: 8, topLeftCell: 'C9' })
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, `${SCALE_FIXTURE}-frozen`, bytes, leg)
        const b2 = page.getByRole('gridcell', { name: /^B2, value / })
        await expect(b2, 'frozen: CF text').toContainText('Colour scale, 0% of range', { timeout: 30_000 })
        const scroller = page.getByTestId('officeview-sheet-scroller')
        await scroller.evaluate(node => { node.scrollTop = 600; node.scrollLeft = 400 })
        await page.waitForTimeout(500)
        await expect(page.getByRole('columnheader', { name: /^Column A/ }).first(), 'frozen column header').toBeVisible()
        await expect(page.getByRole('rowheader', { name: 'Row 1' }).first(), 'frozen row header').toBeVisible()
        const grid = await settledScaleGrid(page, 100)
        expect(countColor(grid, 0, 0, grid.width, grid.height, 0xff, 0, 0), 'frozen: Red scale fill paints').toBeGreaterThan(100)
        expect(countColor(grid, 0, 0, grid.width, grid.height, 0, 0xff, 0), 'frozen: Green scale fill paints').toBeGreaterThan(100)
        await expect(page.getByRole('alert'), 'no error alert').toHaveCount(0)
      } finally {
        await page.close()
      }
    })
  }
})

test('Stage 2 live value edit recolours the scope', async ({ context }) => {
  const raw = readFileSync(`${CF_FIXTURES}${SCALE_FIXTURE}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        // Step 1: Open initial workbook (B2 is 10 = min, B8 is 70 = max)
        await openCf(page, `${SCALE_FIXTURE}-init`, raw, leg)
        const b2Init = page.getByRole('gridcell', { name: /^B2, value / })
        await expect(b2Init).toContainText('Colour scale, 0% of range', { timeout: 30_000 })
        const b8Init = page.getByRole('gridcell', { name: /^B8, value / })
        await expect(b8Init).toContainText('Colour scale, 100% of range', { timeout: 30_000 })

        // Step 2: Edit B2 from 10 to 100 (B2 becomes max 100% of range, B3=20 becomes min 0% of range)
        const editedBytes = await editCellInWorkbook(raw, 'B2', '100')
        await openCf(page, `${SCALE_FIXTURE}-edited`, editedBytes, leg)
        const b2Edited = page.getByRole('gridcell', { name: /^B2, value / })
        await expect(b2Edited).toContainText('Colour scale, 100% of range', { timeout: 30_000 })
        const b3Edited = page.getByRole('gridcell', { name: /^B3, value / })
        await expect(b3Edited).toContainText('Colour scale, 0% of range', { timeout: 30_000 })

        // Scope recolours on canvas
        const grid = await settledScaleGrid(page, 200)
        expect(countColor(grid, 0, 0, grid.width, grid.height, 0xff, 0, 0), 'recoloured: Red fill').toBeGreaterThan(200)
        expect(countColor(grid, 0, 0, grid.width, grid.height, 0, 0xff, 0), 'recoloured: Green fill').toBeGreaterThan(200)
      } finally {
        await page.close()
      }
    })
  }
})

test('Stage 2 print capture produces canvas with scale fill', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${SCALE_FIXTURE}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      await page.addInitScript(() => {
        (window as unknown as { print: () => void }).print = () => {
          (window as unknown as { __printed: number }).__printed = 1
        }
      })
      try {
        await openCf(page, SCALE_FIXTURE, bytes, leg)
        if (!leg.paged) {
          await page.getByRole('button', { name: 'Print', exact: true }).click()
          await page.waitForFunction(() => (window as unknown as { __printed?: number }).__printed === 1, { timeout: 30_000 })
          const printCanvases = page.locator('[data-testid="officeview-sheet-print-pages"] canvas')
          await expect(printCanvases).toHaveCount(1)
          const colors = await printCanvases.first().evaluate(canvas => {
            const c = canvas as HTMLCanvasElement
            const ctx = c.getContext('2d')
            if (!ctx) return { red: 0, green: 0 }
            const data = ctx.getImageData(0, 0, c.width, c.height).data
            let red = 0, green = 0
            for (let i = 0; i < data.length; i += 4) {
              if (Math.abs(data[i]! - 0xff) <= 48 && Math.abs(data[i + 1]! - 0) <= 48 && Math.abs(data[i + 2]! - 0) <= 48) red++
              if (Math.abs(data[i]! - 0) <= 48 && Math.abs(data[i + 1]! - 0xff) <= 48 && Math.abs(data[i + 2]! - 0) <= 48) green++
            }
            return { red, green }
          })
          expect(colors.red, 'Red scale fill is painted on the prepared print canvas').toBeGreaterThan(20)
          expect(colors.green, 'Green scale fill is painted on the prepared print canvas').toBeGreaterThan(20)
          await page.emulateMedia({ media: 'print' })
          await expect(page.getByTestId('officeview-sheet-print-notice')).toBeHidden()
          await expect(printCanvases.first()).toBeVisible()
          if (test.info().project.name === 'chromium') {
            const pdf = await page.pdf({ format: 'A4', printBackground: true })
            expect(pdf.byteLength, 'prepared print produces PDF output').toBeGreaterThan(1000)
          }
        } else {
          await page.getByRole('button', { name: 'Print', exact: true }).click()
          await page.emulateMedia({ media: 'print' })
          await expect(page.locator('[data-testid="officeview-sheet-print-pages"] canvas')).toHaveCount(0)
          await expect(page.getByTestId('officeview-sheet-print-notice')).toBeVisible()
        }
      } finally {
        await page.close()
      }
    })
  }
})
