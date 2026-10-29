/**
 * Real-browser smoke for the Stage 1 XLSX pivot, slicer and timeline display (view-only), on all three paths:
 *   - resident (no worker);
 *   - `?worker=1`, admitted resident (a worker is supplied but the small file never needs it);
 *   - worker-backed paged (`?worker=1&paged=1`, the debug admission override of src/debug/forcePaged.ts).
 *
 * The node tests paint with node-canvas; this spec looks at what Chromium really paints in the interactive viewer.
 * Each Excel-saved fixture (tests/fixtures/xlsx-pivots) is opened once per leg and the grid is captured as a
 * screenshot, then cropped to the object under test:
 *   - painted: the crop is not blank and carries the style's distinctive colours (PivotStyleLight16 header fill,
 *     SlicerStyleLight1 frame and selected item, TimeSlicerStyleLight1 frame and scrollbar track; the colours are the
 *     values pinned in docs/validation/xlsx-pivots.md and measured from the Excel references);
 *   - parity: the paged and worker crops equal the resident crop within a small tolerance. Legs are compared with
 *     each other, in the same browser on the same machine, never against Excel references, so font hinting and
 *     anti-aliasing differences between machines (Linux CI) cannot matter;
 *   - the pivot's field-dropdown buttons ("Row Labels" and "Column Labels") are painted: the button's bluish #738A93
 *     border and its dark arrow are counted in the header cell's right-edge zone, against the same zone of an empty
 *     cell of the same column (no button).
 *
 * The fixtures are routed from disk like tests/browser/paged-charts.spec.ts routes its documents. The capture is a
 * 1x-DPR 1280x800 viewport so the sheet origin and the cell geometry are the same on every machine; every crop is
 * placed relative to the grid canvas, never to the page chrome.
 */
import { expect, test, type BrowserContext, type Page } from '@playwright/test'
import { createCanvas, loadImage } from 'canvas'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { counts } from './viewer-e2e'

const FIXTURES = fileURLToPath(new URL('../fixtures/xlsx-pivots/', import.meta.url))

test.use({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: 1 })

interface Leg { name: string; worker: boolean; paged: boolean }
const LEGS: readonly Leg[] = [
  { name: 'resident', worker: false, paged: false },
  { name: '?worker=1, admitted resident', worker: true, paged: false },
  { name: 'worker-backed paged (?worker=1&paged=1)', worker: true, paged: true },
]

type Rgb = readonly [number, number, number]
interface Rect { x: number; y: number; width: number; height: number }
/** A grid-canvas region (CSS px, canvas origin) holding one object, and the colours that prove it is styled. */
interface Region {
  name: string
  /** Padded around the object so its frame and anti-aliased edges are inside. */
  rect: Rect
  /** Each colour must appear at least `min` times (per channel tolerance 2). */
  colours: ReadonlyArray<{ what: string; rgb: Rgb; min: number }>
}
interface Img { width: number; height: number; data: Uint8ClampedArray }

// Geometry of the canvas at 100% zoom: the origin is the top-left of cell A1 (row 1 is 20 px tall, so row n starts at
// 20 * (n - 1)); the slicer/timeline workbooks use default columns plus A:B and C resized (see the fixture sheet XML).
const PIVOT_COMPACT: Region = {
  // PivotStyleLight16 on this workbook's own Office theme: the header band is accent1 (#156082) at tint 0.8.
  name: 'pivot A3:I22 (PivotStyleLight16, compact layout)',
  rect: { x: 0, y: 36, width: 860, height: 410 },
  colours: [{ what: 'header band, accent1 tint 0.8', rgb: [0xc1, 0xe5, 0xf5], min: 20_000 }],
}
const PIVOT_BESIDE_CONTROLS: Region = {
  // G1:H6 of the "Report" sheet, PivotStyleLight16 with row stripes on the Office 2007-2010 palette.
  name: 'pivot G1:H6 (PivotStyleLight16)',
  rect: { x: 403, y: 0, width: 138, height: 126 },
  colours: [
    { what: 'header band, accent1 tint 0.8', rgb: [0xdc, 0xe6, 0xf2], min: 500 },
    { what: 'first row stripe', rgb: [0xd9, 0xd9, 0xd9], min: 500 },
  ],
}
const SLICER: Region = {
  // D1:F13 of the "Report" sheet.
  name: 'slicer "Product" (SlicerStyleLight1)',
  rect: { x: 212, y: 6, width: 184, height: 262 },
  colours: [
    { what: 'frame border', rgb: [0x4f, 0x81, 0xbd], min: 300 },
    { what: 'selected item with data', rgb: [0xb8, 0xcc, 0xe4], min: 1500 },
  ],
}
const TIMELINE: Region = {
  // H8:L15 of the "Report" sheet.
  name: 'timeline "Date" (TimeSlicerStyleLight1, months)',
  rect: { x: 551, y: 135, width: 236, height: 170 },
  colours: [
    // Excel's 1 px frame is a partial-coverage accent (#5887BF); the viewer draws the accent itself, #4F81BD (docs "Timeline frame shade").
    { what: 'frame border, accent1', rgb: [0x4f, 0x81, 0xbd], min: 300 },
    // Excel's track is #DAE4F1 (accent tint 0.8 of its capture palette); the viewer's tint of #4F81BD is #DCE6F2.
    { what: 'scrollbar track, accent tint 0.8', rgb: [0xdc, 0xe6, 0xf2], min: 400 },
  ],
}

/** Field-dropdown buttons of pivot-layout-compact: header cell, and a control cell of the same column without one. */
/** The button fills the 20 px row (a 15 pt row at 96 dpi) and is flush with the cell's right edge; the zone is its right 22 px. */
const BUTTON_ZONE_WIDTH = 22
const BUTTONS: ReadonlyArray<{ name: string; cell: Rect; control: Rect }> = [
  // A5 "Row Labels" (col A is 0..90 px, row 5 is 80..100 px); the control is the empty A2 above the pivot.
  { name: 'Row Labels (A5)', cell: { x: 0, y: 80, width: 90, height: 20 }, control: { x: 0, y: 20, width: 90, height: 20 } },
  // B3 "Column Labels" (col B is 90..197 px, row 3 is 40..60 px); the control is the empty B2.
  { name: 'Column Labels (B3)', cell: { x: 90, y: 40, width: 107, height: 20 }, control: { x: 90, y: 20, width: 107, height: 20 } },
]
/** The button border colour (docs/validation/xlsx-pivots.md, "Compact indent and buttons"): bluish, so text never matches. */
const BUTTON_BORDER: Rgb = [0x73, 0x8a, 0x93]

const near = (a: number, b: number, tol: number) => Math.abs(a - b) <= tol
function countColour(img: Img, rect: Rect, rgb: Rgb, tol = 2): number {
  let n = 0
  for (let y = Math.max(0, rect.y); y < Math.min(img.height, rect.y + rect.height); y++) {
    for (let x = Math.max(0, rect.x); x < Math.min(img.width, rect.x + rect.width); x++) {
      const i = (y * img.width + x) * 4
      if (near(img.data[i], rgb[0], tol) && near(img.data[i + 1], rgb[1], tol) && near(img.data[i + 2], rgb[2], tol)) n++
    }
  }
  return n
}
function crop(img: Img, rect: Rect): Img {
  const out = new Uint8ClampedArray(rect.width * rect.height * 4)
  for (let y = 0; y < rect.height; y++) {
    const sy = rect.y + y
    if (sy < 0 || sy >= img.height) continue
    for (let x = 0; x < rect.width; x++) {
      const sx = rect.x + x
      if (sx < 0 || sx >= img.width) continue
      out.set(img.data.subarray((sy * img.width + sx) * 4, (sy * img.width + sx) * 4 + 4), (y * rect.width + x) * 4)
    }
  }
  return { width: rect.width, height: rect.height, data: out }
}
/** Pixels that are neither paper white nor a faint grid line (every channel >= 200 and near-grey). */
function inkPixels(img: Img): number {
  let n = 0
  for (let i = 0; i < img.data.length; i += 4) {
    const [r, g, b] = [img.data[i], img.data[i + 1], img.data[i + 2]]
    if (!(r > 235 && g > 235 && b > 235)) n++
  }
  return n
}
function distinctColours(img: Img): number {
  const seen = new Set<number>()
  for (let i = 0; i < img.data.length; i += 4) seen.add((img.data[i] << 16) | (img.data[i + 1] << 8) | img.data[i + 2])
  return seen.size
}
/** Largest per-channel difference and the share of pixels differing by more than `threshold` between two crops. */
function diff(a: Img, b: Img, threshold = 8): { max: number; differing: number; share: number } {
  expect(`${a.width}x${a.height}`).toBe(`${b.width}x${b.height}`)
  let max = 0, differing = 0
  for (let i = 0; i < a.data.length; i += 4) {
    const d = Math.max(Math.abs(a.data[i] - b.data[i]), Math.abs(a.data[i + 1] - b.data[i + 1]), Math.abs(a.data[i + 2] - b.data[i + 2]))
    if (d > max) max = d
    if (d > threshold) differing++
  }
  return { max, differing, share: differing / (a.width * a.height) }
}

async function decode(png: Buffer): Promise<Img> {
  const image = await loadImage(png)
  const canvas = createCanvas(image.width, image.height)
  const ctx = canvas.getContext('2d')
  ctx.drawImage(image, 0, 0)
  const { data } = ctx.getImageData(0, 0, image.width, image.height)
  return { width: image.width, height: image.height, data: new Uint8ClampedArray(data) }
}

/** Opens a fixture of tests/fixtures/xlsx-pivots on one leg, in its own page. */
async function open(context: BrowserContext, fixture: string, leg: Leg): Promise<Page> {
  const page = await context.newPage()
  const bytes = readFileSync(`${FIXTURES}${fixture}.xlsx`)
  await page.route(url => url.pathname === `/__fixtures/${fixture}.xlsx`, route => route.fulfill({ body: bytes, contentType: 'application/octet-stream' }))
  const query = new URLSearchParams({ src: `/__fixtures/${fixture}.xlsx` })
  if (leg.worker) query.set('worker', '1')
  if (leg.paged) query.set('paged', '1')
  await page.goto(`/src/debug/viewer-e2e.html?${query}`)
  await expect(page.getByRole('grid')).toBeVisible({ timeout: 60_000 })
  await expect(page.getByRole('status').filter({ hasText: 'Visible body' })).toContainText('Visible body', { timeout: 60_000 })
  return page
}

/** The grid canvas, as a PNG at the box the viewer gave it. Waits for the paint to settle (two equal captures) and for `ready`. */
async function settledGrid(page: Page, ready: (img: Img) => boolean): Promise<Img> {
  const canvas = page.locator('canvas').first()
  await expect(canvas).toBeVisible()
  const box = await canvas.boundingBox()
  expect(box, 'the grid canvas has a box').not.toBeNull()
  const shoot = async () => page.screenshot({ clip: { x: box!.x, y: box!.y, width: Math.min(box!.width, 1000), height: Math.min(box!.height, 480) }, animations: 'disabled', caret: 'hide' })
  let previous = await shoot()
  const deadline = Date.now() + 30_000
  for (;;) {
    await page.waitForTimeout(200)
    const next = await shoot()
    if (next.equals(previous)) {
      const img = await decode(next)
      if (ready(img)) return img
    }
    if (Date.now() > deadline) return decode(next)
    previous = next
  }
}

const sheetBox = (page: Page) => page.locator('canvas').first().boundingBox()

interface Case { fixture: string; regions: readonly Region[]; buttons: boolean }
const CASES: readonly Case[] = [
  { fixture: 'pivot-layout-compact', regions: [PIVOT_COMPACT], buttons: true },
  { fixture: 'slicer-light1', regions: [SLICER, PIVOT_BESIDE_CONTROLS], buttons: false },
  { fixture: 'timeline-light1-months', regions: [TIMELINE, PIVOT_BESIDE_CONTROLS], buttons: false },
]

/** Parity budget between legs: the same browser paints both, so anything above a stray anti-aliased pixel is a real difference. */
const PARITY_MAX_SHARE = 0.002

for (const { fixture, regions, buttons } of CASES) {
  test(`${fixture}: painted on every leg, and the legs paint the same pixels`, async ({ context }) => {
    const resident: Record<string, Img> = {}
    const grids: Record<string, Img> = {}
    {
      for (const leg of LEGS) {
        await test.step(leg.name, async () => {
          const page = await open(context, fixture, leg)
          try {
            // The paged leg really went through the worker: one `open`, every read after it a worker read.
            expect((await counts(page)).open ?? 0, `${leg.name}: worker opens`).toBe(leg.paged ? 1 : 0)
            if (leg.paged) await expect(page.getByRole('status').filter({ hasText: 'Showing saved values' }).first()).toBeVisible()

            const wanted = (img: Img) => regions.every(region => region.colours.every(c => countColour(img, region.rect, c.rgb) >= c.min))
              && (!buttons || BUTTONS.every(b => countColour(img, boxAtRightEdge(b.cell), BUTTON_BORDER, 6) >= 50))
            const grid = await settledGrid(page, wanted)
            grids[leg.name] = grid
            await expect(page.getByRole('alert'), `${leg.name}: no error alert`).toHaveCount(0)
            expect(await sheetBox(page), 'canvas box is stable').toMatchObject({ x: 44 })

            for (const region of regions) {
              const img = crop(grid, region.rect)
              expect(inkPixels(img), `${leg.name}: ${region.name} is not blank`).toBeGreaterThan(500)
              expect(distinctColours(img), `${leg.name}: ${region.name} paints more than a flat fill`).toBeGreaterThan(8)
              for (const c of region.colours) {
                const n = countColour(img, { x: 0, y: 0, width: img.width, height: img.height }, c.rgb)
                expect(n, `${leg.name}: ${region.name} carries ${c.what} #${c.rgb.map(v => v.toString(16).padStart(2, '0')).join('')}`).toBeGreaterThanOrEqual(c.min)
              }
              if (leg.name === 'resident') resident[region.name] = img
            }

            if (buttons) {
              for (const b of BUTTONS) {
                const on = countColour(grid, boxAtRightEdge(b.cell), BUTTON_BORDER, 6)
                const off = countColour(grid, boxAtRightEdge(b.control), BUTTON_BORDER, 6)
                // The button box has a perimeter of about 70 px; its bluish border is painted, and the same zone of
                // a button-less cell of that column never holds that colour.
                expect(on, `${leg.name}: ${b.name} field-dropdown button border painted`).toBeGreaterThanOrEqual(50)
                expect(off, `${leg.name}: ${b.name} control box (no button) holds no button border`).toBe(0)
                // The arrow: a dark triangle in the box centre, absent from the control box.
                const arrowOn = darkPixels(grid, shrink(boxAtRightEdge(b.cell), 5))
                const arrowOff = darkPixels(grid, shrink(boxAtRightEdge(b.control), 5))
                expect(arrowOn, `${leg.name}: ${b.name} arrow painted`).toBeGreaterThanOrEqual(10)
                expect(arrowOff, `${leg.name}: ${b.name} control box has no arrow`).toBe(0)
              }
            }
          } finally { await page.close() }
        })
      }
      // Real parity: compare leg against leg (same browser, same machine).
      const report: string[] = []
      for (const leg of LEGS.slice(1)) {
        for (const region of regions) {
          const result = diff(resident[region.name], crop(grids[leg.name], region.rect))
          report.push(`${fixture} | ${leg.name} | ${region.name}: max channel diff ${result.max}, ${result.differing} px differ (${(result.share * 100).toFixed(4)}%)`)
          expect(result.share, `${leg.name} vs resident: ${region.name} (${report[report.length - 1]})`).toBeLessThanOrEqual(PARITY_MAX_SHARE)
        }
        // The whole captured grid too: nothing outside the object may differ either.
        const whole = diff(grids.resident, grids[leg.name])
        report.push(`${fixture} | ${leg.name} | whole grid capture: max channel diff ${whole.max}, ${whole.differing} px differ (${(whole.share * 100).toFixed(4)}%)`)
        expect(whole.share, `${leg.name} vs resident: whole grid`).toBeLessThanOrEqual(PARITY_MAX_SHARE)
      }
      for (const line of report) console.log(`[parity] ${line}`)
    }
  })
}

/** Where a field-dropdown button sits: the right edge of the cell (docs/validation/xlsx-pivots.md, "Compact indent and buttons"). */
function boxAtRightEdge(cell: Rect): Rect {
  return { x: cell.x + cell.width - BUTTON_ZONE_WIDTH, y: cell.y, width: BUTTON_ZONE_WIDTH, height: cell.height }
}
function shrink(rect: Rect, by: number): Rect {
  return { x: rect.x + by, y: rect.y + by, width: rect.width - 2 * by, height: rect.height - 2 * by }
}
/** Dark, neutral-or-bluish ink (the arrow is about #828282; paper and the button fill #FDFDFD are far lighter). */
function darkPixels(img: Img, rect: Rect): number {
  let n = 0
  for (let y = rect.y; y < rect.y + rect.height; y++) {
    for (let x = rect.x; x < rect.x + rect.width; x++) {
      const i = (y * img.width + x) * 4
      if (img.data[i] < 170 && img.data[i + 1] < 170 && img.data[i + 2] < 170) n++
    }
  }
  return n
}
