/**
 * Stage 1d (§16 T2): the stream viewer in a real browser against
 * server/test/app.ts — happy / error / expiry / reconnect restores the
 * viewport and visible units / oversized frame / revoke.
 */
import { expect, test, type Page } from '@playwright/test'

const P = Number(process.env.OV_STREAM_PORT ?? 4181)
const CONTROL = `http://127.0.0.1:${P + 1}`
const open = (page: Page, q: string) => page.goto(`/src/debug/stream-viewer.html?port=${P}&${q}`)
const viewer = (page: Page) => page.locator('[role="document"]')
const painted = (page: Page) => page.locator('canvas[data-painted="1"]')
async function sids(page: Page, srv = 'normal'): Promise<string[]> {
  // Live sessions only: earlier pages leave detached ones for the resume grace.
  return (await (await page.request.get(`${CONTROL}/sessions?srv=${srv}`)).json() as Array<{ sid: string; status: string }>).filter((s) => s.status === 'stream').map((s) => s.sid)
}
/** Non-background pixels in a painted canvas (the unit really has content). */
const inked = (page: Page, i: number) => page.evaluate((i) => {
  const c = document.querySelectorAll('canvas')[i] as HTMLCanvasElement
  const d = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data
  let n = 0
  for (let k = 0; k < d.length; k += 4) if (d[k + 3]! > 0 && (d[k]! < 250 || d[k + 1]! < 250 || d[k + 2]! < 250)) n++
  return n
}, i)

// Copy and text a11y are off in the demo default policy, so the notice shows once.
test('happy: a protected deck streams rasters with no text layer and the host watermark ignored', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx')
  await expect(viewer(page)).toBeVisible()
  await expect(painted(page)).toHaveCount(1)
  expect(await inked(page, 0)).toBeGreaterThan(100)
  await expect(page.getByText('text unavailable in this viewing mode')).toHaveCount(1)
  expect(await page.locator('[role="document"] input, [role="document"] [contenteditable]').count()).toBe(0)
  await expect(page.getByText('HOSTMARK')).toHaveCount(0)
  // No selectable text inside the viewer (the Stage 3b Print button's label is chrome, not document text).
  expect(await viewer(page).evaluate((el) => {
    const c = el.cloneNode(true) as Element
    c.querySelectorAll('[data-ov-print-button]').forEach((b) => b.remove())
    return (c.textContent ?? '').replace('text unavailable in this viewing mode', '').trim()
  })).toBe('')
})

test('relaxed + font pack: command units replay in Chromium with every face verified and no unit-failed', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx&relaxed=1&fonts=1')
  await expect(painted(page)).toHaveCount(1)
  const notices = await page.evaluate(() => (window as unknown as { __ovStream: { notices: Array<{ code: string }> } }).__ovStream.notices.map((n) => n.code))
  expect(notices).toContain('relaxed-session')
  expect(notices).not.toContain('unit-failed')
  expect(notices).not.toContain('face-missing')
  expect(await inked(page, 0)).toBeGreaterThan(100)
})

test('error: an unknown document ends in the typed auth-failed error', async ({ page }) => {
  await open(page, 'doc=no-such-file.pptx')
  await expect(page.locator('[data-stream-error="auth-failed"]')).toBeVisible()
})

test('expiry: an 8 s token keeps renewing for 20 s with no error', async ({ page }) => {
  let renews = 0
  page.on('websocket', (ws) => ws.on('framereceived', (f) => { if (typeof f.payload === 'string' && f.payload.includes('"type":"renew"')) renews++ }))
  await open(page, 'doc=pypptx-test.pptx&srv=short')
  await expect(painted(page)).toHaveCount(1)
  await page.waitForTimeout(20_000)
  expect(renews).toBeGreaterThanOrEqual(2)
  expect(await page.locator('[role="alert"]').count()).toBe(0)
  expect((await sids(page, 'short')).length).toBe(1)
})

test('reconnect: a dropped socket resumes the same session; scroll position kept; further units paint', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 720 })
  await open(page, 'doc=lo-tdf104788.pptx')
  await expect(page.locator('canvas')).toHaveCount(6)
  const scroller = viewer(page)
  // A rect delta, not offsetTop: the canvas's offsetParent is its positioned unit wrapper (Stage 3a).
  await scroller.evaluate((el) => { const c = el.querySelectorAll('canvas')[3] as HTMLElement; el.scrollTop += c.getBoundingClientRect().top - el.getBoundingClientRect().top })
  await expect(page.locator('canvas[data-unit="3"][data-painted="1"]')).toHaveCount(1)
  const before = await scroller.evaluate((el) => el.scrollTop)
  const [sid] = await sids(page)
  expect((await page.request.post(`${CONTROL}/drop?sid=${sid}`)).ok()).toBe(true)
  await expect.poll(() => page.evaluate(() => (window as unknown as { __ovStream: { notices: Array<{ code: string }> } }).__ovStream.notices.some((n) => n.code === 'resumed'))).toBe(true)
  expect(Math.abs(await scroller.evaluate((el) => el.scrollTop) - before)).toBeLessThanOrEqual(1)
  expect(await sids(page)).toEqual([sid])
  // A unit not yet painted arrives over the resumed session.
  await scroller.evaluate((el) => { el.scrollTop = el.scrollHeight })
  await expect(page.locator('canvas[data-unit="5"][data-painted="1"]')).toHaveCount(1)
  expect(await page.locator('[role="alert"]').count()).toBe(0)
})

test('oversized frame: the client closes 4007 and shows the typed error', async ({ page }) => {
  await open(page, 'doc=pypptx-test.pptx')
  await expect(painted(page)).toHaveCount(1)
  const [sid] = await sids(page)
  await page.request.post(`${CONTROL}/oversize?sid=${sid}`)
  await expect(page.locator('[data-stream-error="oversized-frame"]')).toBeVisible()
  await expect.poll(async () => (await sids(page)).length).toBe(0)
})

test('revoke: the session closes 4005', async ({ page }) => {
  const closes: number[] = []
  await page.exposeFunction('__close', (c: number) => closes.push(c))
  await page.addInitScript(() => {
    const W = window.WebSocket
    window.WebSocket = class extends W { constructor(u: string | URL, p?: string | string[]) { super(u, p); this.addEventListener('close', (e) => (window as unknown as { __close(c: number): void }).__close(e.code)) } } as typeof WebSocket
  })
  await open(page, 'doc=pypptx-test.pptx')
  await expect(painted(page)).toHaveCount(1)
  const [sid] = await sids(page)
  await page.request.post(`${CONTROL}/revoke?sid=${sid}`)
  await expect.poll(() => closes).toContain(4005)
  await expect(page.locator('[data-stream-error="revoked"]')).toBeVisible()
})
