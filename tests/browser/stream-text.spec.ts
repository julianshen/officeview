/**
 * Stage 3a Task 7: search, selection and copy in the remote view, in a real
 * browser against server/test/app.ts (`?copy=full|cap3|off` sets the session's
 * copy policy). Documents: pypptx-test.pptx ("Presentation Title Text" /
 * "Subtitle Text") and lo-tdf123684.pptx (six one-word "Test" paragraphs
 * stacked on slide 1, for the paragraph cap).
 */
import { expect, test, type Locator, type Page } from '@playwright/test'

const P = Number(process.env.OV_STREAM_PORT ?? 4181)
const open = (page: Page, q: string) => page.goto(`/src/debug/stream-viewer.html?port=${P}&${q}`)
const viewer = (page: Page) => page.locator('[role="document"]')
const painted = (page: Page) => page.locator('canvas[data-painted="1"]')
const status = (page: Page) => page.locator('[role="status"][data-officeview-notice]')
const DOC_WORDS = ['Presentation', 'Title', 'Subtitle']

test.beforeEach(async ({ context, baseURL }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: baseURL })
})

async function search(page: Page, query: string) {
  await viewer(page).focus()
  await page.keyboard.press('Control+f')
  const input = page.getByLabel('Search document')
  await expect(input).toBeFocused()
  await input.fill(query)
  return input
}

/** Drags from just inside the left edge of `from` to just inside the right edge of `to`. */
async function drag(page: Page, from: Locator, to: Locator) {
  const a = (await from.boundingBox())!, b = (await to.boundingBox())!
  await page.mouse.move(a.x + 1, a.y + a.height / 2)
  await page.mouse.down()
  await page.mouse.move((a.x + b.x + b.width) / 2, (a.y + b.y + b.height) / 2, { steps: 5 })
  await page.mouse.move(b.x + b.width - 1, b.y + b.height / 2, { steps: 5 })
  await page.mouse.up()
}
const clipboard = (page: Page) => page.evaluate(() => navigator.clipboard.readText())
const setClipboard = (page: Page, t: string) => page.evaluate((t) => navigator.clipboard.writeText(t), t)

test('Ctrl+F then a query draws hit boxes and a count', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx&copy=full')
  await expect(painted(page)).toHaveCount(1)
  await search(page, 'Text')
  await expect(page.locator('[data-hit]')).toHaveCount(2)
  await expect(page.getByTestId('officeview-search-status')).toHaveText('1 of 2')
  await expect(page.locator('[data-hit-active]')).toHaveCount(1)
  const first = await page.locator('[data-hit-active]').boundingBox()
  await page.keyboard.press('Enter')
  await expect(page.getByTestId('officeview-search-status')).toHaveText('2 of 2')
  expect((await page.locator('[data-hit-active]').boundingBox())!.y).not.toBe(first!.y)
  // The boxes sit inside the unit wrapper over the painted canvas.
  const wrap = (await page.locator('[data-unit-wrap="0"]').boundingBox())!
  for (const b of await page.locator('[data-hit]').all()) {
    const r = (await b.boundingBox())!
    expect(r.x).toBeGreaterThanOrEqual(wrap.x); expect(r.x + r.width).toBeLessThanOrEqual(wrap.x + wrap.width + 0.5)
    expect(r.y).toBeGreaterThanOrEqual(wrap.y); expect(r.y + r.height).toBeLessThanOrEqual(wrap.y + wrap.height + 0.5)
  }
  await page.keyboard.press('Escape')
  await expect(page.getByLabel('Search document')).toHaveCount(0)
  await expect(page.locator('[data-hit]')).toHaveCount(0)
})

test('drag-select then Ctrl+C puts the selected text on the clipboard (copy=full)', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx&copy=full')
  await expect(painted(page)).toHaveCount(1)
  await expect(page.getByText('text unavailable in this viewing mode')).toHaveCount(0)
  await search(page, 'Presentation Title Text')
  const hit = page.locator('[data-hit]')
  await expect(hit).toHaveCount(1)
  await setClipboard(page, 'SENTINEL')
  await drag(page, hit, hit)
  await expect(page.locator('[data-sel]').first()).toBeVisible()
  await expect(viewer(page)).toBeFocused()
  await page.keyboard.press('Control+c')
  await expect.poll(() => clipboard(page)).toBe('Presentation Title Text')
  await expect(status(page)).toHaveText('')
})

test('with cap3, selecting four paragraphs shows the limit message and the clipboard is unchanged', async ({ page }) => {
  await open(page, 'doc=lo-tdf123684.pptx&copy=cap3')
  await expect(painted(page)).toHaveCount(1)
  await search(page, 'Test')
  const hits = page.locator('[data-hit]')
  await expect(hits).toHaveCount(6)
  await setClipboard(page, 'SENTINEL')
  await drag(page, hits.nth(0), hits.nth(3))
  await expect(page.locator('[data-sel]').first()).toBeVisible()
  const before = (await page.locator('[data-unit-wrap="0"]').boundingBox())!.y
  await page.keyboard.press('Control+c')
  await expect(status(page)).toHaveText('Selection too large to copy (limit: 3 paragraphs).')
  expect(await clipboard(page)).toBe('SENTINEL')
  // The message takes the reserved status line: the units do not move.
  expect((await page.locator('[data-unit-wrap="0"]').boundingBox())!.y).toBe(before)
  // Three paragraphs are within the cap; the new selection clears the message.
  await drag(page, hits.nth(0), hits.nth(2))
  await expect(status(page)).toHaveText('')
  await page.keyboard.press('Control+c')
  await expect.poll(() => clipboard(page)).toBe('Test\nTest\nTest')
})

test('with copy off, dragging draws no selection', async ({ page }) => {
  const selects: string[] = []
  page.on('websocket', (ws) => ws.on('framesent', (f) => { if (typeof f.payload === 'string' && /"type":"(select|copy)"/.test(f.payload)) selects.push(f.payload) }))
  await open(page, 'doc=pypptx-test.pptx&copy=off')
  await expect(painted(page)).toHaveCount(1)
  await expect(page.getByText('text unavailable in this viewing mode')).toHaveCount(1)
  await search(page, 'Presentation Title Text')
  const hit = page.locator('[data-hit]')
  await expect(hit).toHaveCount(1)
  await drag(page, hit, hit)
  await viewer(page).focus()
  await page.keyboard.press('Control+a')
  await page.keyboard.press('Control+c')
  await page.waitForTimeout(500)
  await expect(page.locator('[data-sel]')).toHaveCount(0)
  expect(selects).toEqual([])
})

test('the page DOM contains no document text when a11y is off', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx&copy=off')
  await expect(painted(page)).toHaveCount(1)
  await search(page, 'Subtitle')
  await expect(page.locator('[data-hit]')).toHaveCount(1)
  const text = await page.evaluate(() => `${document.body.innerText}\n${document.body.textContent}`)
  for (const w of DOC_WORDS) expect(text).not.toContain(w)
})

test('the unit wrapper keeps the canvas sizing, also at a narrow viewport', async ({ page }) => {
  for (const width of [1280, 400]) {
    await page.setViewportSize({ width, height: 720 })
    await open(page, 'doc=pypptx-test.pptx&copy=full')
    await expect(painted(page)).toHaveCount(1)
    const [w, c] = await Promise.all([page.locator('[data-unit-wrap="0"]').boundingBox(), page.locator('canvas[data-unit="0"]').boundingBox()])
    expect(w).toEqual(c)
    expect(c!.width).toBeLessThanOrEqual(width)
  }
})
