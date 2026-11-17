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

async function openCf(page: Page, name: string, bytes: Buffer, leg: Leg, extra: { cfev?: boolean } = {}): Promise<void> {
  const route = `/__fixtures/cf-${encodeURIComponent(name)}.xlsx`
  await page.route(url => url.pathname === route, r => r.fulfill({ body: bytes, contentType: 'application/octet-stream' }))
  const query = new URLSearchParams({ src: route })
  if (leg.worker) query.set('worker', '1')
  if (leg.paged) query.set('paged', '1')
  // Stage 5 evidence seam (src/debug/cfEvidence.ts, debug-only): the page and its viewer worker hold the readings production implements.
  if (extra.cfev) query.set('cfev', '1')
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
  const names: string[][] = []
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, PINK_FIXTURE, bytes, leg)
        const grid = await settledCfGrid(page, 500)
        expect(countPink(grid, 0, 0, grid.width, grid.height), `${leg.name}: CF fill paints`).toBeGreaterThan(500)
        shots.push(grid)
        names.push(await describedNames(page))
      } finally {
        await page.close()
      }
    })
  }
  const [resident, worker, paged] = shots as [Img, Img, Img]
  // The three legs announce the same descriptors for the same cells (painted bars and the withheld one).
  expect(names[1], 'worker announces what resident announces').toEqual(names[0])
  expect(names[2], 'forced paged announces what resident announces').toEqual(names[0])
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

// Stage 3 (Task 23): the published evidence paints data bars. The Excel-resaved lengths case draws its bars in
// FF0000FF; column C is the explicit 0/100 min/max bar (values 0, 25, 50, 75, 100), column E the num 20/80 bar, whose
// capability is non-discriminating and so stays withheld (announced as unavailable, painting nothing).
const BARS = 'stage3-bar-lengths'
const BLUE_TOL = 32
const countBlue = (img: Img): number => countColor(img, 0, 0, img.width, img.height, 0, 0, 0xff, BLUE_TOL)

/** Screenshot once the canvas stops changing and at least `minBlue` bar pixels have painted. */
async function settledBarGrid(page: Page, minBlue: number): Promise<Img> {
  const box = await page.locator('canvas').first().boundingBox()
  expect(box, 'the grid canvas has a box').not.toBeNull()
  const shoot = async () => page.screenshot({ clip: { x: box!.x, y: box!.y, width: Math.min(box!.width, 1000), height: Math.min(box!.height, 480) }, animations: 'disabled', caret: 'hide' })
  let previous = await shoot()
  const deadline = Date.now() + 45_000
  for (;;) {
    await page.waitForTimeout(250)
    const next = await shoot()
    if (next.equals(previous)) {
      const img = await decode(next)
      if (countBlue(img) >= minBlue) return img
    }
    if (Date.now() > deadline) return decode(next)
    previous = next
  }
}

const cell = (page: Page, ref: string) => page.getByRole('gridcell', { name: new RegExp(`^${ref}, value `) })
/** Ruling S3-A11Y-PAGED: every leg, the forced-paged one included, announces the bar descriptor of each painted bar. */
const DESCRIBED = ['C3', 'C6', 'E3'] as const
const describedNames = async (page: Page): Promise<string[]> => Promise.all(DESCRIBED.map(async ref => (await cell(page, ref).getAttribute('aria-label')) ?? (await cell(page, ref).textContent()) ?? ''))

test('Stage 3 data bars paint and announce on every leg, and legs agree (Task 23)', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${BARS}.xlsx`)
  const shots: Img[] = []
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, BARS, bytes, leg)
        // The bar descriptor (Ruling S3-P): a fixed label and the value's position in the range.
        await expect(cell(page, 'C6'), `${leg.name}: C6 announces its bar`).toContainText('Data bar, 100% of range', { timeout: 30_000 })
        await expect(cell(page, 'C3'), `${leg.name}: C3 announces its bar`).toContainText('Data bar, 25% of range', { timeout: 30_000 })
        // A bar whose capability is not verified is announced as unavailable, never as a bar.
        await expect(cell(page, 'E3'), `${leg.name}: E3 withholds`).toContainText('Conditional formatting unavailable', { timeout: 30_000 })
        await expect(cell(page, 'E3')).not.toContainText('Data bar')
        await expect(page.getByRole('alert'), `${leg.name}: no error alert`).toHaveCount(0)
        const grid = await settledBarGrid(page, 2000)
        expect(countBlue(grid), `${leg.name}: bar pixels paint`).toBeGreaterThan(2000)
        shots.push(grid)
      } finally {
        await page.close()
      }
    })
  }
  const [resident, worker, paged] = shots as [Img, Img, Img]
  for (const [name, other] of [['worker', worker], ['forced paged', paged]] as const) {
    expect(other.width, name).toBe(resident.width)
    expect(other.height, name).toBe(resident.height)
    let differing = 0
    for (let i = 0; i < resident.data.length; i += 4) {
      if ([0, 1, 2].some(k => Math.abs(resident.data[i + k]! - other.data[i + k]!) > 4)) differing++
    }
    expect(differing / (resident.width * resident.height), `${name} matches resident`).toBeLessThan(0.002)
  }
})

test('Stage 3 data bars keep painting at the six zoom/DPR pairs', async ({ browser }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${BARS}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      for (const dpr of [1, 2]) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: dpr })
        try {
          const page = await context.newPage()
          try {
            await openCf(page, BARS, bytes, leg)
            for (const zoom of ['50%', '100%', '200%']) {
              const label = page.getByRole('button', { name: 'Reset zoom' })
              for (let i = 0; i < 12; i++) {
                if ((await label.textContent()) === zoom) break
                const current = Number((await label.textContent())!.replace('%', ''))
                await page.getByRole('button', { name: current < Number(zoom.replace('%', '')) ? 'Zoom in' : 'Zoom out' }).click()
              }
              await expect(label).toHaveText(zoom, { timeout: 10_000 })
              await page.getByTestId('officeview-sheet-scroller').evaluate(node => { node.scrollTop = 0; node.scrollLeft = 0 })
              await expect(cell(page, 'C6'), `dpr=${dpr} zoom=${zoom}: bar descriptor`).toContainText('Data bar, 100% of range', { timeout: 30_000 })
              const grid = await settledBarGrid(page, 300)
              expect(countBlue(grid), `dpr=${dpr} zoom=${zoom}: bar pixels`).toBeGreaterThan(300)
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

test('Stage 3 data bars paint under frozen panes with pinned headers', async ({ context }) => {
  const bytes = await frozenBytes(BARS, { xSplit: 1, ySplit: 1, topLeftCell: 'B2' })
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, `${BARS}-frozen`, bytes, leg)
        await expect(cell(page, 'C6'), 'frozen: bar descriptor').toContainText('Data bar, 100% of range', { timeout: 30_000 })
        const scroller = page.getByTestId('officeview-sheet-scroller')
        await scroller.evaluate(node => { node.scrollTop = 40; node.scrollLeft = 60 })
        await page.waitForTimeout(500)
        await expect(page.getByRole('columnheader', { name: /^Column A/ }).first(), 'frozen column header').toBeVisible()
        await expect(page.getByRole('rowheader', { name: 'Row 1' }).first(), 'frozen row header').toBeVisible()
        const grid = await settledBarGrid(page, 1000)
        expect(countBlue(grid), 'frozen: bar pixels paint').toBeGreaterThan(1000)
        await expect(page.getByRole('alert'), 'no error alert').toHaveCount(0)
      } finally {
        await page.close()
      }
    })
  }
})

test('Stage 3 live max edit rescales the bars', async ({ context }) => {
  const raw = readFileSync(`${CF_FIXTURES}${BARS}.xlsx`)
  // C6 is the maximum (100). Raising it to 200 moves every other bar of the min/max rule: 75 -> 38%, 25 -> 13%.
  const edited = await editCellInWorkbook(raw, 'C6', '200')
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, `${BARS}-init`, raw, leg)
        await expect(cell(page, 'C5')).toContainText('Data bar, 75% of range', { timeout: 30_000 })
        const before = countBlue(await settledBarGrid(page, 2000))
        await openCf(page, `${BARS}-edited`, edited, leg)
        await expect(cell(page, 'C6')).toContainText('Data bar, 100% of range', { timeout: 30_000 })
        await expect(cell(page, 'C5')).toContainText('Data bar, 38% of range', { timeout: 30_000 })
        await expect(cell(page, 'C3')).toContainText('Data bar, 13% of range', { timeout: 30_000 })
        const after = countBlue(await settledBarGrid(page, 1000))
        // The C bars shrink, so fewer bar pixels paint after the edit.
        expect(after, `${leg.name}: rescaled bars paint fewer pixels (${before} -> ${after})`).toBeLessThan(before)
      } finally {
        await page.close()
      }
    })
  }
})

test('Stage 3 print capture paints data bars on the prepared page', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${BARS}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      await page.addInitScript(() => {
        (window as unknown as { print: () => void }).print = () => { (window as unknown as { __printed: number }).__printed = 1 }
      })
      try {
        await openCf(page, BARS, bytes, leg)
        await expect(cell(page, 'C6')).toContainText('Data bar, 100% of range', { timeout: 30_000 })
        await page.getByRole('button', { name: 'Print', exact: true }).click()
        if (!leg.paged) {
          await page.waitForFunction(() => (window as unknown as { __printed?: number }).__printed === 1, { timeout: 30_000 })
          const printCanvases = page.locator('[data-testid="officeview-sheet-print-pages"] canvas')
          await expect(printCanvases).toHaveCount(1)
          const blue = await printCanvases.first().evaluate(canvas => {
            const c = canvas as HTMLCanvasElement
            const data = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data
            let n = 0
            for (let i = 0; i < data.length; i += 4) if (data[i]! <= 32 && data[i + 1]! <= 32 && data[i + 2]! >= 223) n++
            return n
          })
          expect(blue, 'bar pixels on the prepared print canvas').toBeGreaterThan(200)
          await page.emulateMedia({ media: 'print' })
          await expect(printCanvases.first()).toBeVisible()
        } else {
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

test('the forced-paged view announces its painted bars (Ruling S3-A11Y-PAGED)', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${BARS}.xlsx`)
  const page = await context.newPage()
  try {
    await openCf(page, BARS, bytes, LEGS[2]!)
    expect(countBlue(await settledBarGrid(page, 2000)), 'the bars paint').toBeGreaterThan(2000)
    await expect(cell(page, 'C6'), 'C6 announces its painted bar').toContainText('Data bar, 100% of range', { timeout: 15_000 })
    await expect(cell(page, 'C3'), 'C3 announces its value position').toContainText('Data bar, 25% of range', { timeout: 15_000 })
    // The withheld (unverified num 20/80) bar still announces unavailable, never a bar.
    await expect(cell(page, 'E3')).toContainText('Conditional formatting unavailable', { timeout: 15_000 })
    await expect(cell(page, 'E3')).not.toContainText('Data bar')
  } finally {
    await page.close()
  }
})

// Stage 4 (Task 29; Task 30 published the evidence and the glyphs). Icons paint only on the admitted geometry
// (Calibri 11 regular, 15/18/30 pt rows, centred: S4-ROW-EXACT and S4-ROW-BASELINE). stage4-icon-inset uses 20 pt rows,
// so on every leg its icons stay withheld: each cell keeps its value text un-inset and announces no icon (the tests on
// ICONS below pin that fail-closed outcome). The tests on ROWS further down pin the painted outcome. Column B of
// stage4-icon-inset is a 3Arrows num 0/5/8 rule over left-aligned B2:B4 (10, 50, 90), whose values are visible (the
// thresholds fixture hides its values behind a ';;;' format).
const ICONS = 'stage4-icon-inset'
const ICON_CELLS = ['B2', 'B3', 'B4'] as const
const ICON_VALUES: Record<(typeof ICON_CELLS)[number], string> = { B2: '10', B3: '50', B4: '90' }
const named = (page: Page, ref: string, value: string) => page.getByRole('gridcell', { name: new RegExp(`^${ref}, value ${value}\\b`) })
async function expectNoIcon(page: Page, label: string): Promise<void> {
  for (const ref of ICON_CELLS) {
    await expect(named(page, ref, ICON_VALUES[ref]), `${label}: ${ref} keeps its value in its accessible name`).toHaveCount(1, { timeout: 30_000 })
    await expect(cell(page, ref), `${label}: ${ref} announces no icon`).not.toContainText(/Icon \d of \d|圖示/)
  }
}
/** Screenshot once the canvas stops changing. */
async function settledGrid(page: Page): Promise<Img> {
  const box = await page.locator('canvas').first().boundingBox()
  expect(box, 'the grid canvas has a box').not.toBeNull()
  const shoot = async () => page.screenshot({ clip: { x: box!.x, y: box!.y, width: Math.min(box!.width, 1000), height: Math.min(box!.height, 480) }, animations: 'disabled', caret: 'hide' })
  let previous = await shoot()
  const deadline = Date.now() + 30_000
  for (;;) {
    await page.waitForTimeout(250)
    const next = await shoot()
    if (next.equals(previous) || Date.now() > deadline) return decode(next)
    previous = next
  }
}

test('Stage 4 icons off the admitted geometry (20 pt rows) stay withheld and announce nothing on every leg, and legs agree', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${ICONS}.xlsx`)
  const shots: Img[] = []
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, ICONS, bytes, leg)
        await expectNoIcon(page, leg.name)
        await expect(page.getByRole('alert'), `${leg.name}: no error alert`).toHaveCount(0)
        shots.push(await settledGrid(page))
      } finally {
        await page.close()
      }
    })
  }
  // Pixels: the worker leg matches the resident one. The forced-paged frame lays this fixture's rows out about 1 px
  // apart from the resident one (observed on header row 1, which no CF range covers), so its agreement is the announced
  // labels above plus the un-inset text on every leg; a pixel comparison there would measure the row metrics, not CF
  // (a known limit of this fixture only, docs/validation/xlsx-cf.md §11; the Stage 3 test keeps all three legs).
  const [resident, worker] = shots as [Img, Img, Img]
  for (const [name, other] of [['worker', worker]] as const) {
    expect(other.width, name).toBe(resident.width)
    expect(other.height, name).toBe(resident.height)
    let differing = 0
    for (let i = 0; i < resident.data.length; i += 4) {
      if ([0, 1, 2].some(k => Math.abs(resident.data[i + k]! - other.data[i + k]!) > 4)) differing++
    }
    expect(differing / (resident.width * resident.height), `${name} matches resident`).toBeLessThan(0.002)
  }
})

test('Stage 4 withheld icons stay withheld at the six zoom/DPR pairs', async ({ browser }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${ICONS}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      for (const dpr of [1, 2]) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: dpr })
        try {
          const page = await context.newPage()
          try {
            await openCf(page, ICONS, bytes, leg)
            for (const zoom of ['50%', '100%', '200%']) {
              const label = page.getByRole('button', { name: 'Reset zoom' })
              for (let i = 0; i < 12; i++) {
                if ((await label.textContent()) === zoom) break
                const current = Number((await label.textContent())!.replace('%', ''))
                await page.getByRole('button', { name: current < Number(zoom.replace('%', '')) ? 'Zoom in' : 'Zoom out' }).click()
              }
              await expect(label).toHaveText(zoom, { timeout: 10_000 })
              await page.getByTestId('officeview-sheet-scroller').evaluate(node => { node.scrollTop = 0; node.scrollLeft = 0 })
              await expect(named(page, 'B2', '10'), `dpr=${dpr} zoom=${zoom}`).toHaveCount(1, { timeout: 30_000 })
              await expect(cell(page, 'B2')).not.toContainText(/Icon \d of \d/)
              await expect(page.getByRole('alert'), `dpr=${dpr} zoom=${zoom}: no error alert`).toHaveCount(0)
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

test('Stage 4 withheld icons under frozen panes', async ({ context }) => {
  const bytes = await frozenBytes(ICONS, { xSplit: 1, ySplit: 1, topLeftCell: 'B2' })
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, `${ICONS}-frozen`, bytes, leg)
        await expectNoIcon(page, `${leg.name} frozen`)
        const scroller = page.getByTestId('officeview-sheet-scroller')
        await scroller.evaluate(node => { node.scrollTop = 40; node.scrollLeft = 60 })
        await page.waitForTimeout(500)
        await expect(page.getByRole('columnheader', { name: /^Column A/ }).first(), 'frozen column header').toBeVisible()
        await expect(page.getByRole('rowheader', { name: 'Row 1' }).first(), 'frozen row header').toBeVisible()
        await expect(page.getByRole('alert'), 'no error alert').toHaveCount(0)
      } finally {
        await page.close()
      }
    })
  }
})

test('Stage 4 live max edit on withheld icons keeps the value and no icon', async ({ context }) => {
  const raw = readFileSync(`${CF_FIXTURES}${ICONS}.xlsx`)
  // B4 holds the rule's maximum (90); editing it to 1 moves it from the top bucket to the bottom one. With no icon
  // painted in production, the edit must reach the value text and still announce no icon.
  const edited = await editCellInWorkbook(raw, 'B4', '1')
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, `${ICONS}-init`, raw, leg)
        await expectNoIcon(page, `${leg.name} before`)
        await openCf(page, `${ICONS}-edited`, edited, leg)
        await expect(named(page, 'B4', '1'), `${leg.name}: the edited value`).toHaveCount(1, { timeout: 30_000 })
        await expect(cell(page, 'B4')).not.toContainText(/Icon \d of \d/)
      } finally {
        await page.close()
      }
    })
  }
})

test('Stage 4 print capture paints no withheld icon', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${ICONS}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      await page.addInitScript(() => {
        (window as unknown as { print: () => void }).print = () => { (window as unknown as { __printed: number }).__printed = 1 }
      })
      try {
        await openCf(page, ICONS, bytes, leg)
        await expectNoIcon(page, leg.name)
        await page.getByRole('button', { name: 'Print', exact: true }).click()
        if (!leg.paged) {
          await page.waitForFunction(() => (window as unknown as { __printed?: number }).__printed === 1, { timeout: 30_000 })
          const printCanvases = page.locator('[data-testid="officeview-sheet-print-pages"] canvas')
          await expect(printCanvases).toHaveCount(1)
          // The prepared page paints the sheet's value text (dark ink), and no icon (20 pt rows: withheld).
          const ink = await printCanvases.first().evaluate(canvas => {
            const c = canvas as HTMLCanvasElement
            const data = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data
            let n = 0
            for (let i = 0; i < data.length; i += 4) if (data[i]! < 96 && data[i + 1]! < 96 && data[i + 2]! < 96) n++
            return n
          })
          expect(ink, 'value text on the prepared print canvas').toBeGreaterThan(200)
          await page.emulateMedia({ media: 'print' })
          await expect(printCanvases.first()).toBeVisible()
        } else {
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

// Task 30: painted icons. stage4-icon-rows is Excel's own resave (Excel 16.115): 11 pt Calibri, every cell 10 under num
// 0/5/8 (bucket 2: "Icon 3 of 3"), rows 2..9 at 15, 12.75, 16.5, 18, 20, 22.5, 25 and 30 pt. Its Normal style is centred,
// so B..D and F..I are centred; column E keeps its own xf, bottom-aligned. Every cell is Calibri 11 regular, so bottom
// placement is admitted too (S4-ROW-UNIFORM). Admitted: rows 2, 5 and 9 (15, 18, 30 pt). Withheld: every other row.
const ROWS = 'stage4-icon-rows'
const PAINTED = ['B2', 'F2', 'I2', 'E2', 'B5', 'F5', 'E5', 'B9', 'H9', 'E9'] as const
const WITHHELD = ['B3', 'F4', 'F6', 'B7', 'F8', 'E3', 'E7'] as const
/** Pixels of the 3Arrows:2 / 3TrafficLights1:2 / 3Symbols:2 light green (0x5bbb76 and kin, +-40): only an icon draws it. */
function greenPixels(img: Img): number {
  let n = 0
  for (let i = 0; i < img.data.length; i += 4) {
    const r = img.data[i]!, g = img.data[i + 1]!, b = img.data[i + 2]!
    if (g > 120 && g - r > 40 && g - b > 20) n++
  }
  return n
}
async function expectRowsIcons(page: Page, label: string): Promise<void> {
  for (const ref of PAINTED) await expect(cell(page, ref), `${label}: ${ref} announces its icon`).toContainText('Icon 3 of 3', { timeout: 30_000 })
  for (const ref of WITHHELD) await expect(cell(page, ref), `${label}: ${ref} announces no icon`).not.toContainText(/Icon \d of \d/)
}

test('Stage 4 icons paint on the admitted geometry, announce on every leg, and the three legs agree', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${ROWS}.xlsx`)
  const shots: Img[] = []
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, ROWS, bytes, leg)
        await expectRowsIcons(page, leg.name)
        await expect(page.getByRole('alert'), `${leg.name}: no error alert`).toHaveCount(0)
        const shot = await settledGrid(page)
        expect(greenPixels(shot), `${leg.name}: the icons' green is on the canvas`).toBeGreaterThan(100)
        shots.push(shot)
      } finally {
        await page.close()
      }
    })
  }
  // The worker leg matches the resident one pixel for pixel. The forced-paged frame differs from the resident one from
  // the header row down (text and gridlines, about 1 px; the forced-paged frame offset that exists at base,
  // docs/validation/xlsx-cf.md §11), so its agreement on CF is the icon pixels: every pixel only an icon draws (the
  // glyphs' green) is the same pixel on every leg, its colour within 4 (anti-aliased edges). The left-text inset is not
  // pixel-compared on the forced-paged leg (uncovered there; its announced icon and the unit tests cover it).
  const [resident, worker, forced] = shots as [Img, Img, Img]
  const greenMask = (img: Img): string[] => {
    const out: string[] = []
    for (let i = 0; i < img.data.length; i += 4) {
      const r = img.data[i]!, g = img.data[i + 1]!, b = img.data[i + 2]!
      if (g > 120 && g - r > 40 && g - b > 20) out.push(`${(i / 4) % img.width},${Math.floor(i / 4 / img.width)}:${r},${g},${b}`)
    }
    return out
  }
  expect(worker.width).toBe(resident.width)
  expect(worker.height).toBe(resident.height)
  let differing = 0
  for (let i = 0; i < resident.data.length; i += 4) if ([0, 1, 2].some(k => resident.data[i + k] !== worker.data[i + k])) differing++
  expect(differing, 'worker matches resident pixel for pixel').toBe(0)
  // Positions exact; colours within the Stage 3 three-leg tolerance (4 per channel: anti-aliased glyph edges move by 1).
  const icons = greenMask(resident), forcedIcons = greenMask(forced)
  expect(icons.length).toBeGreaterThan(100)
  expect(forcedIcons.map(p => p.split(':')[0]), 'forced paged paints the icons at the same pixels as resident').toEqual(icons.map(p => p.split(':')[0]))
  for (let i = 0; i < icons.length; i++) {
    const a = icons[i]!.split(':')[1]!.split(',').map(Number), b = forcedIcons[i]!.split(':')[1]!.split(',').map(Number)
    expect(Math.max(...a.map((v, k) => Math.abs(v - b[k]!))), icons[i]).toBeLessThanOrEqual(4)
  }
})

test('Stage 4 painted icons at the six zoom/DPR pairs: painted and announced at each (fidelity *unverified* off 100 %)', async ({ browser }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${ROWS}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      for (const dpr of [1, 2]) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: dpr })
        try {
          const page = await context.newPage()
          try {
            await openCf(page, ROWS, bytes, leg)
            for (const zoom of ['50%', '100%', '200%']) {
              const label = page.getByRole('button', { name: 'Reset zoom' })
              for (let i = 0; i < 12; i++) {
                if ((await label.textContent()) === zoom) break
                const current = Number((await label.textContent())!.replace('%', ''))
                await page.getByRole('button', { name: current < Number(zoom.replace('%', '')) ? 'Zoom in' : 'Zoom out' }).click()
              }
              await expect(label).toHaveText(zoom, { timeout: 10_000 })
              await page.getByTestId('officeview-sheet-scroller').evaluate(node => { node.scrollTop = 0; node.scrollLeft = 0 })
              // The description is decided at the capture scale: the same at every zoom and DPR.
              await expect(cell(page, 'F2'), `dpr=${dpr} zoom=${zoom}`).toContainText('Icon 3 of 3', { timeout: 30_000 })
              await expect(cell(page, 'E3'), `dpr=${dpr} zoom=${zoom}`).not.toContainText(/Icon \d of \d/)
              expect(greenPixels(await settledGrid(page)), `dpr=${dpr} zoom=${zoom}: icons painted`).toBeGreaterThan(20)
              await expect(page.getByRole('alert'), `dpr=${dpr} zoom=${zoom}: no error alert`).toHaveCount(0)
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

test('Stage 4 painted icons under frozen panes', async ({ context }) => {
  const bytes = await frozenBytes(ROWS, { xSplit: 1, ySplit: 1, topLeftCell: 'B2' })
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, `${ROWS}-frozen`, bytes, leg)
        await expectRowsIcons(page, `${leg.name} frozen`)
        expect(greenPixels(await settledGrid(page)), `${leg.name} frozen: icons painted`).toBeGreaterThan(100)
        await expect(page.getByRole('alert'), 'no error alert').toHaveCount(0)
      } finally {
        await page.close()
      }
    })
  }
})

test('Stage 4 live edit re-buckets a painted icon', async ({ context }) => {
  const raw = readFileSync(`${CF_FIXTURES}${ROWS}.xlsx`)
  // F2 holds 10 (bucket 2 under num 0/5/8); 1 moves it to bucket 0 ("Icon 1 of 3", 3Arrows:0).
  const edited = await editCellInWorkbook(raw, 'F2', '1')
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, `${ROWS}-init`, raw, leg)
        await expect(cell(page, 'F2'), `${leg.name} before`).toContainText('Icon 3 of 3', { timeout: 30_000 })
        await openCf(page, `${ROWS}-edited`, edited, leg)
        await expect(named(page, 'F2', '1'), `${leg.name}: the edited value`).toHaveCount(1, { timeout: 30_000 })
        await expect(cell(page, 'F2'), `${leg.name}: re-bucketed`).toContainText('Icon 1 of 3')
        await expect(cell(page, 'B2'), `${leg.name}: its neighbour keeps its icon`).toContainText('Icon 3 of 3')
      } finally {
        await page.close()
      }
    })
  }
})

test('Stage 4 print capture paints the admitted icons', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${ROWS}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      await page.addInitScript(() => {
        (window as unknown as { print: () => void }).print = () => { (window as unknown as { __printed: number }).__printed = 1 }
      })
      try {
        await openCf(page, ROWS, bytes, leg)
        await expect(cell(page, 'F2')).toContainText('Icon 3 of 3', { timeout: 30_000 })
        await page.getByRole('button', { name: 'Print', exact: true }).click()
        if (!leg.paged) {
          await page.waitForFunction(() => (window as unknown as { __printed?: number }).__printed === 1, { timeout: 30_000 })
          const printCanvases = page.locator('[data-testid="officeview-sheet-print-pages"] canvas')
          await expect(printCanvases).toHaveCount(1)
          const green = await printCanvases.first().evaluate(canvas => {
            const c = canvas as HTMLCanvasElement
            const data = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data
            let n = 0
            for (let i = 0; i < data.length; i += 4) if (data[i + 1]! > 120 && data[i + 1]! - data[i]! > 40 && data[i + 1]! - data[i + 2]! > 20) n++
            return n
          })
          expect(green, 'the icons on the prepared print canvas').toBeGreaterThan(100)
        } else {
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

// ==================== Stage 5 formula rules (Task 38) ====================
//
// Task 39 published Stage 5 (42 capabilities, among them the formula dialects stage5-formula-authored needs). Two kinds of run,
// each stating what it is: `cfev` runs open the page with the debug-only evidence seam (`?cfev=1`, src/debug/cfEvidence.ts,
// excluded from the build) so the page AND its viewer worker hold every reading production implements; production runs (no
// `cfev`) use the published set. On these fixtures both paint the formula rules, identically across the legs, forced-paged
// included. Neither run compares with Excel here (tests/xlsx-cf-oracle-conformance.test.ts does).

const S5_AUTHORED = 'stage5-formula-authored'
const S5_CROSS = 'stage5-cross-sheet'
const S5_CELLS = ['B2', 'B3', 'B4', 'B5', 'E2', 'E3', 'E4', 'E5', 'G2', 'G3', 'G4', 'G5'] as const
const S5_CROSS_CELLS = ['B2', 'B3', 'B4', 'B5', 'C2', 'C3', 'C4', 'C5', 'D2', 'D3', 'D4', 'D5', 'E2', 'E3', 'E4', 'E5'] as const
const FORMULA_TEXT = /[=$!]|[A-Z]{2,}\(|TODAY|COUNTIF|ISBLANK|Sheet2/
const namesOf = async (page: Page, refs: readonly string[]): Promise<string[]> => Promise.all(refs.map(async ref => (await cell(page, ref).getAttribute('aria-label')) ?? (await cell(page, ref).textContent()) ?? ''))
/** The names, each stripped of the cell's own value text: what is left is the CF announcement. */
const announcements = (names: readonly string[]): string[] => names.map(n => n.replace(/^[A-Z]+\d+, value [^;]*;?\s*/, ''))
const appliedCount = (names: readonly string[]): number => names.filter(n => /Conditional formatting applied/.test(n)).length
const announced = (refs: readonly string[], min: number) => async (page: Page): Promise<boolean> => (await namesOf(page, refs)).filter(n => /Conditional formatting/.test(n)).length >= min

async function editSheetCell(raw: Buffer, sheetPath: string, ref: string, value: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(raw)
  const xml = await zip.file(sheetPath)!.async('string')
  const patched = xml.replace(new RegExp(`(<c r="${ref}"[^>]*>.*?<v>)([^<]*)(</v>)`, 's'), `$1${value}$3`)
  if (patched === xml) throw new Error(`Could not edit ${ref} in ${sheetPath}`)
  zip.file(sheetPath, patched)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }) as unknown as Promise<Buffer>
}

test('Stage 5 formula rules paint and announce on every leg, and legs agree', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${S5_AUTHORED}.xlsx`)
  const shotsByRun: Record<string, number> = {}
  for (const cfev of [true, false]) {
    const shots: Img[] = [], names: string[][] = []
    for (const leg of LEGS) {
      await test.step(`${leg.name}${cfev ? ' (evidence seam)' : ' (production evidence)'}`, async () => {
        const page = await context.newPage()
        try {
          await openCf(page, S5_AUTHORED, bytes, leg, { cfev })
          await expect.poll(async () => announced(S5_CELLS, 2)(page), { timeout: 30_000, message: `${leg.name}: the formula cells announce their CF state` }).toBe(true)
          const grid = await settledCfGrid(page, 200)
          shots.push(grid)
          names.push(announcements(await namesOf(page, S5_CELLS)))
          for (const n of names.at(-1)!) expect(n).not.toMatch(FORMULA_TEXT)
          // Published (Task 39): production applies the formula rules as the evidence seam does.
          expect(appliedCount(await namesOf(page, S5_CELLS)), `${leg.name}: formula rules apply`).toBeGreaterThan(2)
          expect(countPink(grid, 0, 0, grid.width, grid.height), `${leg.name}: the formula fill paints`).toBeGreaterThan(200)
          await expect(page.getByRole('alert'), `${leg.name}: no error alert`).toHaveCount(0)
        } finally { await page.close() }
      })
    }
    expect(names[1], 'worker announces what resident announces').toEqual(names[0])
    expect(names[2], 'forced paged announces what resident announces').toEqual(names[0])
    const [resident, worker, paged] = shots as [Img, Img, Img]
    expect(worker.width).toBe(resident.width); expect(paged.width).toBe(resident.width)
    let differing = 0
    for (let i = 0; i < resident.data.length; i += 4) if (Math.abs(resident.data[i]! - worker.data[i]!) > 4 || Math.abs(resident.data[i + 1]! - worker.data[i + 1]!) > 4 || Math.abs(resident.data[i + 2]! - worker.data[i + 2]!) > 4) differing++
    expect(differing / (resident.width * resident.height), 'worker paints what resident paints').toBeLessThan(0.002)
    // The forced-paged tile pipeline lays pixels out differently (its fill coverage is not pixel-equal), but it paints the same cells:
    // the same number of filled row bands and column bands as resident, and the fill is present on it.
    const bands = (img: Img): { rows: number; cols: number; pink: number } => {
      const isPink = (x: number, y: number): boolean => { const i = (y * img.width + x) * 4; return Math.abs(img.data[i]! - 0xff) <= 48 && Math.abs(img.data[i + 1]! - 0xc7) <= 48 && Math.abs(img.data[i + 2]! - 0xce) <= 48 }
      const rowHit = Array.from({ length: img.height }, (_, y) => { let n = 0; for (let x = 0; x < img.width; x++) if (isPink(x, y)) n++; return n >= 10 })
      const colHit = Array.from({ length: img.width }, (_, x) => { let n = 0; for (let y = 0; y < img.height; y++) if (isPink(x, y)) n++; return n >= 10 })
      const runs = (hit: boolean[]): number => hit.filter((v, i) => v && !hit[i - 1]).length
      return { rows: runs(rowHit), cols: runs(colHit), pink: countPink(img, 0, 0, img.width, img.height) }
    }
    const [br, bw, bp] = [bands(resident), bands(worker), bands(paged)]
    expect(bw, 'worker bands').toEqual(br)
    expect(bp.pink, 'forced paged paints the fill').toBeGreaterThan(200)
    expect({ rows: bp.rows, cols: bp.cols }, 'forced paged paints the same cells as resident').toEqual({ rows: br.rows, cols: br.cols })
    shotsByRun[cfev ? 'evidence' : 'production'] = br.pink
  }
  // Published: production paints the formula fill the evidence seam paints (within anti-aliasing).
  expect(Math.abs(shotsByRun.evidence! - shotsByRun.production!), 'production paints what the evidence seam paints').toBeLessThanOrEqual(200)
  expect(shotsByRun.production!).toBeGreaterThan(200)
})

test('Stage 5 cross-sheet edit repaints the dependent sheet', async ({ context }) => {
  // The viewer is view-only (no cell-commit API on this page, and a worker-backed workbook has no value commit), so the edit is
  // made in the file and the page is opened on it: a new document, not a same-session commit. The same-session invalidation
  // (a referenced-sheet commit staling the dependent sheet's tiles and subscriptions) is covered by tests/xlsx-cf-live.test.tsx,
  // tests/xlsx-cf-formula-cross-sheet.test.ts and the remote-cache dependency epoch tests.
  const raw = readFileSync(`${CF_FIXTURES}${S5_CROSS}.xlsx`)
  const edited = await editSheetCell(raw, 'xl/worksheets/sheet2.xml', 'A3', '9')
  const seen: Record<string, { names: string[]; pink: number }[]> = {}
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      for (const [phase, bytes] of [['before', raw], ['after', edited]] as const) {
        const page = await context.newPage()
        try {
          await openCf(page, `${S5_CROSS}-${phase}`, bytes, leg, { cfev: true })
          await expect.poll(async () => announced(S5_CROSS_CELLS, 4)(page), { timeout: 30_000, message: `${leg.name} ${phase}: announces` }).toBe(true)
          const grid = await settledCfGrid(page, 0)
          const names = announcements(await namesOf(page, S5_CROSS_CELLS))
          for (const n of names) expect(n).not.toMatch(FORMULA_TEXT)
          ;(seen[phase] ??= []).push({ names, pink: countPink(grid, 0, 0, grid.width, grid.height) })
        } finally { await page.close() }
      }
    })
  }
  for (const phase of ['before', 'after']) { expect(seen[phase]![1]!.names, `${phase}: worker`).toEqual(seen[phase]![0]!.names); expect(seen[phase]![2]!.names, `${phase}: forced paged`).toEqual(seen[phase]![0]!.names) }
  expect(appliedCount(seen.before![0]!.names), 'the cross-sheet rules apply from Sheet2').toBeGreaterThan(1)
  for (const leg of [0, 1, 2]) {
    expect(seen.after![leg]!.names, `leg ${leg}: Sheet2 moved what Sheet1 announces`).not.toEqual(seen.before![leg]!.names)
    expect(seen.after![leg]!.pink, `leg ${leg}: and what it paints`).not.toBe(seen.before![leg]!.pink)
  }
})

test('Stage 5 formula rules under frozen panes', async ({ context }) => {
  const bytes = await frozenBytes(S5_AUTHORED, { xSplit: 1, ySplit: 1, topLeftCell: 'B2' })
  const names: string[][] = []
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, `${S5_AUTHORED}-frozen`, bytes, leg, { cfev: true })
        await expect.poll(async () => announced(S5_CELLS, 2)(page), { timeout: 30_000, message: `${leg.name}: announces` }).toBe(true)
        const scroller = page.getByTestId('officeview-sheet-scroller')
        await scroller.evaluate(node => { node.scrollTop = 600; node.scrollLeft = 400 })
        await page.waitForTimeout(500)
        await expect(page.getByRole('columnheader', { name: /^Column A/ }).first(), 'frozen column header').toBeVisible()
        await expect(page.getByRole('rowheader', { name: 'Row 1' }).first(), 'frozen row header').toBeVisible()
        await scroller.evaluate(node => { node.scrollTop = 0; node.scrollLeft = 0 })
        await page.waitForTimeout(500)
        await expect.poll(async () => announced(S5_CELLS, 2)(page), { timeout: 30_000 }).toBe(true)
        names.push(announcements(await namesOf(page, S5_CELLS)))
        expect(appliedCount(await namesOf(page, S5_CELLS)), `${leg.name}: formula rules apply under frozen panes`).toBeGreaterThan(2)
        const grid = await settledCfGrid(page, 200)
        expect(countPink(grid, 0, 0, grid.width, grid.height), `${leg.name}: the formula fill paints`).toBeGreaterThan(200)
        await expect(page.getByRole('alert')).toHaveCount(0)
      } finally { await page.close() }
    })
  }
  expect(names[1]).toEqual(names[0]); expect(names[2]).toEqual(names[0])
})

test('Stage 5 print capture paints formula rules', async ({ context }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${S5_AUTHORED}.xlsx`)
  const pinkOf = async (page: Page): Promise<number> => page.locator('[data-testid="officeview-sheet-print-pages"] canvas').first().evaluate(canvas => {
    const c = canvas as HTMLCanvasElement, ctx = c.getContext('2d')
    if (!ctx) return -1
    const data = ctx.getImageData(0, 0, c.width, c.height).data
    let count = 0
    for (let i = 0; i < data.length; i += 4) if (Math.abs(data[i]! - 0xff) <= 48 && Math.abs(data[i + 1]! - 0xc7) <= 48 && Math.abs(data[i + 2]! - 0xce) <= 48) count++
    return count
  })
  const totals: Record<string, number> = {}
  for (const cfev of [true, false]) {
    const pink: number[] = []
    for (const leg of LEGS.filter(l => !l.paged)) {
      await test.step(`${leg.name}${cfev ? ' (evidence seam)' : ' (production evidence)'}`, async () => {
        const page = await context.newPage()
        await page.addInitScript(() => { (window as unknown as { print: () => void }).print = () => { (window as unknown as { __printed: number }).__printed = 1 } })
        try {
          await openCf(page, S5_AUTHORED, bytes, leg, { cfev })
          await page.getByRole('button', { name: 'Print', exact: true }).click()
          await page.waitForFunction(() => (window as unknown as { __printed?: number }).__printed === 1, { timeout: 30_000 })
          await expect.poll(async () => page.locator('[data-testid="officeview-sheet-print-pages"] canvas').count(), { message: 'the workbook\'s pages are prepared' }).toBeGreaterThan(0)
          pink.push(await pinkOf(page))
        } finally { await page.close() }
      })
    }
    // Print prepares the same CF the screen does: resident and worker pages carry the same formula fill.
    expect(pink[1], `resident and worker print pages agree (${cfev ? 'evidence seam' : 'production'})`).toBe(pink[0])
    totals[cfev ? 'evidence' : 'production'] = pink[0]!
  }
  // Published (Task 39): production prints the formula fill the evidence seam prints.
  expect(Math.abs(totals.evidence! - totals.production!), 'production prints what the evidence seam prints').toBeLessThanOrEqual(200)
})

// ---- Stage 5: a referenced-sheet commit in the SAME live session (fix round 2, finding 6) ----
//
// The viewer is view-only, but a pivot click is a real in-session commit on all three legs: it replaces the pivot's displayed cells
// (an overlay on the pivot's sheet) without reopening anything. The Excel-saved pivot workbook
// tests/fixtures/xlsx-pivots/engine/baseline.xlsx has the pivot on "Pivot" and its source data on "Data"; this spec adds one
// cross-sheet rule to "Data"!A2, `Pivot!$B$7>5000`. East's first child row is B7 = 1234.5 (false); collapsing East moves the next
// region's 6497.84 into B7 (true). The test visits Data, goes back to Pivot, clicks Collapse East (the commit), returns to Data and
// expects A2 painted and announced: the dependent rule must read the COMMITTED overlay value of the other sheet, on the resident,
// worker and forced-paged legs alike (no reopen, no forced rerender; the evidence seam is the debug one above).
//
// This test goes through a sheet switch, which re-prepares the sheet it lands on, so it does not exercise the dependency subscriptions
// (reverting `cfResidentDependencyKey` or `cfDependencyEpoch` leaves it green, checked). The test after it does: it keeps the
// dependent sheet on screen while a slicer on that sheet commits another sheet's pivot. The revert-checks that fail THIS one are
// dropping the resident CF workbook (sheetAssets) and the worker's cross-sheet seam (pagedSource cfWorkbook).
const PIVOTS = fileURLToPath(new URL('../fixtures/xlsx-pivots/', import.meta.url))
async function crossSheetPivotBytes(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(readFileSync(`${PIVOTS}engine/baseline.xlsx`))
  const sheet = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
  const rule = '<conditionalFormatting sqref="A2"><cfRule type="expression" dxfId="0" priority="1"><formula>Pivot!$B$7&gt;5000</formula></cfRule></conditionalFormatting>'
  if (!sheet.includes('<pageMargins')) throw new Error('Data sheet has no pageMargins anchor')
  zip.file('xl/worksheets/sheet2.xml', sheet.replace('<pageMargins', `${rule}<pageMargins`))
  const styles = await zip.file('xl/styles.xml')!.async('string')
  if (!styles.includes('<dxfs count="0"/>')) throw new Error('unexpected dxfs')
  zip.file('xl/styles.xml', styles.replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }) as unknown as Promise<Buffer>
}

test('Stage 5 cross-sheet commit in the same live session repaints the dependent sheet', async ({ context }) => {
  const bytes = await crossSheetPivotBytes()
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, 'stage5-pivot-cross', bytes, leg, { cfev: true })
        const sheets = page.getByRole('combobox', { name: 'Worksheet' })
        const collapse = page.getByRole('button', { name: /^Collapse East/ })
        await expect(collapse).toBeAttached({ timeout: 30_000 })
        // Visit the dependent sheet first: its tiles and CF scope are prepared and cached under the pre-commit pivot.
        await sheets.selectOption({ label: 'Data' })
        await expect(cell(page, 'A2')).toContainText('East', { timeout: 30_000 })
        await expect.poll(async () => (await namesOf(page, ['A2']))[0], { timeout: 15_000, message: 'A2 settles before the commit' }).not.toMatch(/Conditional formatting applied/)
        const before = await settledCfGrid(page, 0)
        // The Data sheet's own styling is pinkish at the default tolerance, so the rule's fill (FFC7CE) is counted exactly.
        const exact = (img: Img): number => countColor(img, 0, 0, img.width, img.height, 0xff, 0xc7, 0xce, 2)
        const fillBefore = exact(before)
        expect(fillBefore, `${leg.name}: the rule's fill is absent before the commit`).toBeLessThan(100)
        // The commit: collapse East on the pivot's own sheet, in this session.
        await sheets.selectOption({ label: 'Pivot' })
        await expect(collapse).toBeAttached({ timeout: 30_000 })
        await collapse.click()
        await expect(page.getByRole('button', { name: 'Expand East (Region)' })).toBeAttached({ timeout: 30_000 })
        // Back on the dependent sheet: it must show the committed state, announced and painted.
        await sheets.selectOption({ label: 'Data' })
        await expect.poll(async () => (await namesOf(page, ['A2']))[0], { timeout: 30_000, message: `${leg.name}: A2 announces the rule after the pivot commit` }).toMatch(/Conditional formatting applied/)
        const after = await settledCfGrid(page, 0)
        await expect.poll(async () => exact(await settledCfGrid(page, 0)), { timeout: 30_000, message: `${leg.name}: A2's fill paints after the commit` }).toBeGreaterThan(fillBefore + 300)
        void after
        await expect(page.getByRole('alert')).toHaveCount(0)
      } finally { await page.close() }
    })
  }
})


// ---- the dependent sheet stays on screen while the referenced sheet commits (fix round 3) ----
//
// controls/controls-seed.xlsx (Excel): the pivot on "工作表1" is filtered by a "Product" slicer that starts at All. The slicer's
// drawing is moved from that sheet to "Data", so the slicer lives on the dependent sheet and its click commits the pivot on another
// sheet (InteractiveSheetDoc applyLinked) while Data stays displayed. Data!A2 carries `'工作表1'!$C$8>50`: C8 is Bananas' column
// (20) at All and the Total column (90) once Apples alone is selected. The measured single-item click from All and Clear Filter are
// the actions the slicer controller allows (src/xlsx/pivots/slicers/controller.ts:177-184 and :166 for clear); nothing is injected.
// No sheet switch, no reopen, no forced rerender: the sheet's own subscriptions must repaint and re-announce A2 after each commit.
//
// Legs (fix round 4). Every leg must COMMIT (the slicer republishes its selection only for a committed result). Forced paged holds
// the pivot in the worker, so its commit is a remote controlApply then controlCommit (src/worker/viewerRemotePivots.ts applyLinked),
// counted at the worker port; `?worker=1` alone admits this small workbook resident, so its commit is the main-thread one (zero
// control commands, pinned). The worker's transaction chunk is served by the harness warmup (vite.harness.config.ts).
// Revert-checks, each fails this test: the resident dependency key in the CF live key (resident leg); the remote dependency epoch
// in that key, or RemoteCfCache.invalidateDependents (forced-paged leg); the harness warmup of slicers/transaction.ts (forced paged:
// the commit is refused). What it does NOT prove: that the CF-key tile invalidation (the `paintedCf` check in the canvas effect) is
// needed here. The slicer publication also advances Data's drawing generation, whose tile invalidation repaints Data on its own:
// removing the CF-key invalidation alone leaves this test green, removing both fails it (checked). The announcements follow the CF
// evaluation, not the tiles.
async function slicerOnDataBytes(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(readFileSync(`${PIVOTS}controls/controls-seed.xlsx`))
  const s1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const moved = /<drawing r:id="rId2"\/><extLst>.*<\/extLst>/s.exec(s1)
  if (!moved) throw new Error('controls-seed drawing/extLst anchor not found')
  zip.file('xl/worksheets/sheet1.xml', s1.replace(moved[0], ''))
  const rels1 = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string')
  const keep = (id: string) => new RegExp(`<Relationship Id="${id}"[^>]*/>`).exec(rels1)![0]
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', rels1.replace(keep('rId2'), '').replace(keep('rId3'), '').replace(keep('rId4'), ''))
  zip.file('xl/worksheets/_rels/sheet2.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${keep('rId2')}${keep('rId3')}${keep('rId4')}</Relationships>`)
  const s2 = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
  const rule = `<conditionalFormatting sqref="A2"><cfRule type="expression" dxfId="0" priority="1"><formula>'工作表1'!$C$8&gt;50</formula></cfRule></conditionalFormatting>`
  zip.file('xl/worksheets/sheet2.xml', s2.replace('<pageMargins', `${rule}<pageMargins`).replace('</worksheet>', `${moved[0]}</worksheet>`))
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file('xl/styles.xml', styles.replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }) as unknown as Promise<Buffer>
}

test('Stage 5 slicer commit on another sheet repaints the visible dependent sheet without a switch', async ({ context }) => {
  const bytes = await slicerOnDataBytes()
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      const page = await context.newPage()
      try {
        await openCf(page, 'stage5-slicer-controls', bytes, leg, { cfev: true })
        const sheets = page.getByRole('combobox', { name: 'Worksheet' })
        // The slicer's target is resolved against the pivot definitions of the sheets the viewer holds loaded (paged shells carry
        // none, InteractiveSheetDoc.tsx slicer apply). Wait for 工作表1's pivot display before leaving it: a pivot field button
        // (named after the field on the resident legs, after the saved caption when paged).
        await expect(sheets.locator('option:checked')).toHaveText('工作表1')
        await expect(page.getByRole('button', { name: /^Filter and sort / }).first()).toBeAttached({ timeout: 30_000 })
        await sheets.selectOption({ label: 'Data' })
        await expect(cell(page, 'A2')).toContainText('2023', { timeout: 30_000 })
        const exact = (img: Img): number => countColor(img, 0, 0, img.width, img.height, 0xff, 0xc7, 0xce, 2)
        const a2 = async (): Promise<string> => (await namesOf(page, ['A2']))[0]!
        await expect.poll(a2, { timeout: 15_000, message: `${leg.name}: A2 settles at All` }).not.toMatch(/Conditional formatting applied/)
        const fillAll = exact(await settledCfGrid(page, 0))
        expect(fillAll, `${leg.name}: the rule's fill is absent at All`).toBeLessThan(100)
        const linked = async (): Promise<[number, number]> => { const c = await counts(page); return [c.controlApply ?? 0, c.controlCommit ?? 0] }
        const [applied0, committed0] = await linked()
        // Commit 1, Data on screen throughout: All -> Apples alone.
        await page.getByRole('button', { name: 'Apples, selected' }).click()
        // The commit succeeded: the controller publishes the new selection only for a committed result (controller.ts enqueue).
        await expect(page.getByRole('button', { name: 'Bananas, not selected' })).toBeAttached({ timeout: 30_000 })
        await expect(page.getByRole('button', { name: 'Apples, selected' })).toBeAttached()
        // Forced paged holds the pivot in the worker: the click is one remote controlApply then controlCommit
        // (src/worker/viewerRemotePivots.ts applyLinked). `?worker=1` alone admits this small workbook resident, so its commit is the
        // main-thread one and no control command reaches the worker (pinned, so the leg cannot silently change kind).
        const remote = leg.paged ? 1 : 0
        await expect.poll(linked, { timeout: 30_000, message: `${leg.name}: ${remote} remote apply and commit` }).toEqual([applied0 + remote, committed0 + remote])
        // Data is still the selected worksheet: nothing below switches sheets.
        await expect(sheets.locator('option:checked')).toHaveText('Data')
        await expect.poll(a2, { timeout: 30_000, message: `${leg.name}: A2 announces the rule after the slicer commit` }).toMatch(/Conditional formatting applied/)
        await expect.poll(async () => exact(await settledCfGrid(page, 0)), { timeout: 30_000, message: `${leg.name}: A2's fill paints after the commit` }).toBeGreaterThan(fillAll + 300)
        // Commit 2, still on Data: Clear Filter restores All, and the rule goes away again.
        await page.getByRole('button', { name: 'Clear Filter' }).click()
        await expect(page.getByRole('button', { name: 'Bananas, selected' })).toBeAttached({ timeout: 30_000 })
        await expect.poll(linked, { timeout: 30_000, message: `${leg.name}: the clear's remote apply and commit` }).toEqual([applied0 + 2 * remote, committed0 + 2 * remote])
        await expect.poll(a2, { timeout: 30_000, message: `${leg.name}: A2 withdraws the rule after Clear Filter` }).not.toMatch(/Conditional formatting applied/)
        await expect.poll(async () => exact(await settledCfGrid(page, 0)), { timeout: 30_000, message: `${leg.name}: A2's fill is gone after Clear Filter` }).toBeLessThan(100)
        await expect(page.getByRole('alert')).toHaveCount(0)
      } finally { await page.close() }
    })
  }
})

// ---- Stage 5 at the six zoom/DPR pairs and across a midnight (Task 39 fix rounds 1-2, F6/R3) ----

/** The CF fill pixels (red FFC7CE or green C6EFCE, +-24) inside one cell's on-screen box, its 2 px border excluded. */
async function cellFillPixels(page: Page, ref: string): Promise<number> {
  const box = await cell(page, ref).boundingBox()
  if (!box || box.width < 6 || box.height < 6) throw new Error(`${ref}: no on-screen cell box`)
  const img = await decode(await page.screenshot({ clip: { x: box.x + 2, y: box.y + 2, width: box.width - 4, height: box.height - 4 }, animations: 'disabled', caret: 'hide' }))
  let n = 0
  for (let i = 0; i < img.data.length; i += 4) {
    const [r, g, b] = [img.data[i]!, img.data[i + 1]!, img.data[i + 2]!]
    if ((Math.abs(r - 0xff) <= 24 && Math.abs(g - 0xc7) <= 24 && Math.abs(b - 0xce) <= 24) || (Math.abs(r - 0xc6) <= 24 && Math.abs(g - 0xef) <= 24 && Math.abs(b - 0xce) <= 24)) n++
  }
  return n
}
/** Whether a cell is mostly CF-filled (at least a third of its inner box). */
async function cellFilled(page: Page, ref: string): Promise<boolean> {
  const box = (await cell(page, ref).boundingBox())!
  return (await cellFillPixels(page, ref)) * 3 >= Math.max(1, (box.width - 4) * (box.height - 4)) * (await page.evaluate(() => window.devicePixelRatio)) ** 2
}

test('Stage 5 formula rules paint and announce at the six zoom/DPR pairs on every leg (production evidence)', async ({ browser }) => {
  const bytes = readFileSync(`${CF_FIXTURES}${S5_AUTHORED}.xlsx`)
  for (const leg of LEGS) {
    await test.step(leg.name, async () => {
      // R3: one reference per leg, so announcements are compared across both DPRs as well as across the zooms.
      const reference: string[] = []
      for (const dpr of [1, 2]) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: dpr })
        try {
          const page = await context.newPage()
          try {
            await openCf(page, S5_AUTHORED, bytes, leg)
            for (const zoom of ['50%', '100%', '200%']) {
              const label = page.getByRole('button', { name: 'Reset zoom' })
              for (let i = 0; i < 12; i++) {
                if ((await label.textContent()) === zoom) break
                const current = Number((await label.textContent())!.replace('%', ''))
                await page.getByRole('button', { name: current < Number(zoom.replace('%', '')) ? 'Zoom in' : 'Zoom out' }).click()
              }
              await expect(label).toHaveText(zoom, { timeout: 10_000 })
              await page.getByTestId('officeview-sheet-scroller').evaluate(node => { node.scrollTop = 0; node.scrollLeft = 0 })
              const where = `${leg.name} dpr=${dpr} zoom=${zoom}`
              await expect.poll(async () => appliedCount(await namesOf(page, S5_CELLS)), { timeout: 30_000, message: `${where}: formula rules apply` }).toBeGreaterThan(2)
              // The announcement is decided by the rule, not the scale: identical at every zoom and DPR.
              const names = announcements(await namesOf(page, S5_CELLS))
              if (reference.length === 0) reference.push(...names)
              else expect(names, where).toEqual(reference)
              for (const n of names) expect(n, where).not.toMatch(FORMULA_TEXT)
              await settledCfGrid(page, 50)
              // R3: the pixels of each designated formula cell agree with its announcement: an applied cell is filled, a
              // cell without applied formatting is not.
              const all = await namesOf(page, S5_CELLS)
              let filled = 0
              for (const [i, ref] of S5_CELLS.entries()) {
                const appliedHere = /Conditional formatting applied/.test(all[i] ?? '')
                expect(await cellFilled(page, ref), `${where} ${ref}: pixels follow the announcement`).toBe(appliedHere)
                if (appliedHere) filled++
              }
              expect(filled, `${where}: formula cells painted`).toBeGreaterThan(2)
              await expect(page.getByRole('alert'), `${where}: no error alert`).toHaveCount(0)
            }
          } finally { await page.close() }
        } finally { await context.close() }
      }
    })
  }
})

/**
 * The midnight tick in a browser (resident leg). Playwright's clock (`page.clock`) fakes Date and timers on the page; it
 * never reaches a dedicated worker (init scripts and clock installation do not run in a worker realm), so the worker and
 * forced-paged legs cannot be put on a faked midnight here: their tick is covered by the mounted tests of Task 38
 * (tests/xlsx-cf-live.test.tsx, the remote-cache clock-fact tests). stage5-clock-today holds `=A2=TODAY()` over B2:B10,
 * with A4 and A9 on the anchor day: before the anchor day's midnight they do not match, after it they do.
 */
test('Stage 5 midnight tick repaints TODAY() rules on the resident leg (injected browser clock)', async ({ browser }) => {
  const id = 'stage5-clock-today'
  const bytes = readFileSync(`${CF_FIXTURES}${id}.xlsx`)
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, timezoneId: 'Asia/Taipei' })
  try {
    const page = await context.newPage()
    try {
      // Before the anchor day (2026-11-18, the capture's local day) begins in Asia/Taipei; the clock is paused at
      // 23:59:30 so only runFor crosses midnight (a run without it stays on the 17th; revert-checked).
      // The page loads on a flowing clock (its scheduling needs timers), then the clock is paused before midnight.
      await page.clock.install({ time: new Date('2026-11-17T23:58:00+08:00') })
      await openCf(page, `${id}-midnight`, bytes, LEGS[0]!)
      await page.clock.pauseAt(new Date('2026-11-17T23:59:30+08:00'))
      const applied = async (ref: string): Promise<boolean> => /Conditional formatting applied/.test((await namesOf(page, [ref]))[0] ?? '')
      await expect.poll(async () => (await namesOf(page, ['B2']))[0] ?? '', { timeout: 30_000, message: 'the grid announces' }).toMatch(/B2, value/)
      expect(await applied('B4'), 'before midnight: the anchor-day cell is not today').toBe(false)
      expect(await applied('B9'), 'before midnight: the anchor-day cell is not today').toBe(false)
      // R3: the canvas, not only the announcement: B4/B9 unfilled before midnight, filled after it.
      expect(await cellFilled(page, 'B4'), 'before midnight: B4 unfilled on the canvas').toBe(false)
      expect(await cellFilled(page, 'B9'), 'before midnight: B9 unfilled on the canvas').toBe(false)
      await midnightTick(page)
      await expect.poll(async () => (await applied('B4')) && (await applied('B9')), { timeout: 30_000, message: 'after midnight: the anchor-day cells match TODAY()' }).toBe(true)
      await expect.poll(async () => (await cellFilled(page, 'B4')) && (await cellFilled(page, 'B9')), { timeout: 30_000, message: 'after midnight: the canvas repaints B4/B9' }).toBe(true)
      await expect(page.getByRole('alert')).toHaveCount(0)
    } finally { await page.close() }
  } finally { await context.close() }
})

/** Crosses the anchor day's midnight on the injected clock (the page is paused at 23:59:30). */
async function midnightTick(page: Page): Promise<void> { await page.clock.runFor(60_000) }

/**
 * R3 revert-check, kept as a test: a canvas that is not repainted after midnight (every 2D draw call disabled just
 * before the tick, while the accessibility layer still updates) is caught by the pixel check that the midnight test uses.
 */
test('Stage 5 midnight pixel check catches a stale canvas (revert-check)', async ({ browser }) => {
  const id = 'stage5-clock-today'
  const bytes = readFileSync(`${CF_FIXTURES}${id}.xlsx`)
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, timezoneId: 'Asia/Taipei' })
  try {
    const page = await context.newPage()
    try {
      await page.clock.install({ time: new Date('2026-11-17T23:58:00+08:00') })
      await openCf(page, `${id}-stale`, bytes, LEGS[0]!)
      await page.clock.pauseAt(new Date('2026-11-17T23:59:30+08:00'))
      await expect.poll(async () => (await namesOf(page, ['B2']))[0] ?? '', { timeout: 30_000 }).toMatch(/B2, value/)
      expect(await cellFilled(page, 'B4')).toBe(false)
      // Freeze the canvas: every 2D drawing and clearing call becomes a no-op from here on.
      await page.evaluate(() => {
        const proto = CanvasRenderingContext2D.prototype as unknown as Record<string, unknown>
        for (const name of ['fillRect', 'clearRect', 'strokeRect', 'fillText', 'drawImage', 'fill', 'stroke', 'putImageData']) proto[name] = () => {}
      })
      await midnightTick(page)
      // The announcement moves on, the frozen canvas does not: the pixel check reports the stale frame.
      await expect.poll(async () => /Conditional formatting applied/.test((await namesOf(page, ['B4']))[0] ?? ''), { timeout: 30_000 }).toBe(true)
      await page.waitForTimeout(1000)
      expect(await cellFilled(page, 'B4'), 'a stale canvas is detected (B4 still unfilled)').toBe(false)
    } finally { await page.close() }
  } finally { await context.close() }
})
