/**
 * Task 3 (spec §6.5, §6.9, §6.10, §6.11; Review Focus 4, 5, 7): the paged
 * viewer's toolbar and status line, fit-to-width, page tracking and the lazy
 * chrome's layout stability, in a real browser with the real stylesheet.
 *
 * Page: src/debug/viewer-e2e.html (the shipped <OfficeFile>; `name=1` hands
 * the bytes over as a named File, `toolbar=0` starts with showToolbar off,
 * `theme=` sets the theme). Documents are built here and routed in.
 *
 * The Layout Instability API is Chromium-only. On WebKit the same delayed-
 * chunk cases assert the engine-agnostic part: the placeholder and real chrome
 * heights match and the scroll container's box does not move when the chunk
 * lands (no WebKit CLS number is claimed).
 */
import { expect, test, type Page, type Route } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx, buildPptx } from '../../src/testdata/ooxml-builders'

const lines = (n: number, prefix = 'Line') => Array.from({ length: n }, (_, i) => ({ runs: [{ text: `${prefix} ${i + 1} of the paged toolbar test document, with enough words to wrap.` }] }))

/** buildPptx writes one slide; copy it so the deck has `count` slides. */
async function deck(count: number): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildPptx([{ off: ['914400', '914400'], ext: ['7315200', '1828800'], paragraphs: [{ runs: [{ text: 'Slide text for the toolbar test' }] }] }]))
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
  let pres = await zip.file('ppt/presentation.xml')!.async('string')
  let presRels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string')
  let types = await zip.file('[Content_Types].xml')!.async('string')
  for (let i = 2; i <= count; i++) {
    zip.file(`ppt/slides/slide${i}.xml`, slide)
    zip.file(`ppt/slides/_rels/slide${i}.xml.rels`, rels)
    pres = pres.replace('</p:sldIdLst>', `<p:sldId id="${255 + i}" r:id="rIdS${i}"/></p:sldIdLst>`)
    presRels = presRels.replace('</Relationships>', `<Relationship Id="rIdS${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i}.xml"/></Relationships>`)
    types = types.replace('</Types>', `<Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`)
  }
  zip.file('ppt/presentation.xml', pres)
  zip.file('ppt/_rels/presentation.xml.rels', presRels)
  zip.file('[Content_Types].xml', types)
  return zip.generateAsync({ type: 'uint8array' })
}

const fixtures = new Map<string, Promise<Uint8Array>>()
function fixture(name: string): Promise<Uint8Array> {
  if (!fixtures.has(name)) {
    fixtures.set(name, name === 'Report.docx' ? buildDocx(lines(220))
      : name === 'Tall.docx' ? buildDocx(lines(60))
        : name === 'Hundred.docx' ? buildDocx(lines(4400, 'Row'))
          : name === 'Deck.pptx' ? deck(3)
            : Promise.reject(new Error(`no fixture ${name}`)))
  }
  return fixtures.get(name)!
}

const CHUNK = (url: URL) => url.pathname.endsWith('/src/components/paged/PagedChrome.tsx')

interface OpenOptions { query?: Record<string, string>; hold?: boolean }
/** Opens a document; with `hold`, the PagedChrome chunk waits for the returned release(). */
async function open(page: Page, name: string, options: OpenOptions = {}): Promise<{ release(): void }> {
  await page.route((url) => url.pathname.startsWith('/__fixtures/'), async (route) => {
    const requested = decodeURIComponent(new URL(route.request().url()).pathname.slice('/__fixtures/'.length))
    await route.fulfill({ body: Buffer.from(await fixture(requested)), contentType: 'application/octet-stream' })
  })
  let release = () => {}
  if (options.hold) {
    const gate = new Promise<void>((r) => { release = r })
    await page.route(CHUNK, async (route: Route) => { await gate; await route.continue() })
  }
  const query = new URLSearchParams({ src: `/__fixtures/${name}`, name: '1', ...options.query })
  await page.goto(`/src/debug/viewer-e2e.html?${query}`)
  await page.locator('[data-officeview-root] canvas').first().waitFor({ timeout: 60_000 })
  if (!options.hold && options.query?.toolbar !== '0') await toolbar(page).waitFor()
  return { release }
}

/** Task 6: chrome motion has finished (WAAPI or JS-driven): every motion element is opaque and untransformed. */
const motionSettled = (page: Page) => page.waitForFunction(() => document.getAnimations().every((a) => a.playState !== 'running') && [...document.querySelectorAll('[data-ov-motion]')].every((el) => { const cs = getComputedStyle(el); return cs.opacity === '1' && (cs.transform === 'none' || cs.transform === 'matrix(1, 0, 0, 1, 0, 0)') }))
const toolbar = (page: Page) => page.getByRole('toolbar', { name: 'Document toolbar' })
const statusLine = (page: Page) => page.getByTestId('officeview-status')
const region = (page: Page) => page.getByRole('region', { name: 'Document viewer' })
const coarse = (page: Page) => page.evaluate(() => matchMedia('(pointer: coarse)').matches)
/** Two animation frames: the rAF-throttled tracker has run and React has committed. */
const frames = (page: Page) => page.evaluate(() => new Promise<void>((r) => requestAnimationFrame(() => requestAnimationFrame(() => r()))))

type Rect = { left: number; top: number; right: number; bottom: number; width: number; height: number }
const overlap = (a: Rect, b: Rect) => Math.max(0, Math.min(a.right, b.right) - Math.max(a.left, b.left)) * Math.max(0, Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top))

for (const vp of [{ label: 'desktop 1280x800', viewport: { width: 1280, height: 800 }, touch: false }, { label: 'compact 375x667 coarse', viewport: { width: 375, height: 667 }, touch: true }]) {
  test.describe(vp.label, () => {
    test.use({ viewport: vp.viewport, hasTouch: vp.touch, isMobile: vp.touch })

    test('toolbar and status never overlap the visible part of any page', async ({ page }) => {
      for (const name of ['Report.docx', 'Deck.pptx']) {
        await open(page, name)
        for (const scroll of [0, 600, 1e6]) {
          await region(page).evaluate((el, top) => { el.scrollTop = top }, scroll)
          await frames(page)
          const r = await page.evaluate(() => {
            const rect = (el: Element) => el.getBoundingClientRect().toJSON() as Rect
            const container = rect(document.querySelector('[data-officeview-root]')!)
            const pages = [...document.querySelectorAll('[data-officeview-root] canvas')].map(rect)
            return { container, pages, bar: rect(document.querySelector('[data-testid=officeview-toolbar]')!), status: rect(document.querySelector('[data-testid=officeview-status]')!) }
          })
          for (const [i, p] of r.pages.entries()) {
            // the visible part of the page: inside the scroll container
            const visible = { left: Math.max(p.left, r.container.left), right: Math.min(p.right, r.container.right), top: Math.max(p.top, r.container.top), bottom: Math.min(p.bottom, r.container.bottom), width: 0, height: 0 }
            if (visible.right <= visible.left || visible.bottom <= visible.top) continue
            expect(overlap(visible, r.bar), `${name} page ${i} vs toolbar at ${scroll}`).toBeLessThanOrEqual(0.5)
            expect(overlap(visible, r.status), `${name} page ${i} vs status at ${scroll}`).toBeLessThanOrEqual(0.5)
          }
        }
        await page.unrouteAll({ behavior: 'ignoreErrors' })
      }
    })

    test('placeholder and real chrome heights match; CLS ≤ 0.001 with a delayed chunk', async ({ page, browserName }) => {
      await page.addInitScript(() => {
        const w = window as unknown as { __shifts: { value: number; recent: boolean; t: number }[] }
        w.__shifts = []
        try {
          new PerformanceObserver((list) => {
            for (const e of list.getEntries() as unknown as { value: number; hadRecentInput: boolean; startTime: number }[]) w.__shifts.push({ value: e.value, recent: e.hadRecentInput, t: e.startTime })
          }).observe({ type: 'layout-shift', buffered: true })
        } catch { /* no Layout Instability API (WebKit) */ }
      })
      const expected = vp.touch || vp.viewport.width < 600 ? 48 : 40
      const geometry = () => page.evaluate(() => {
        const box = (sel: string) => { const el = document.querySelector(sel) as HTMLElement | null; return el ? el.getBoundingClientRect().toJSON() as Rect : null }
        return { placeholder: box('[data-testid=officeview-toolbar-placeholder]'), bar: box('[data-testid=officeview-toolbar]'), slot: box('[data-testid=officeview-status-slot]'), status: box('[data-testid=officeview-status]'), container: box('[data-officeview-root]')! }
      })
      const cls = (since: number) => page.evaluate((t) => (window as unknown as { __shifts: { value: number; recent: boolean; t: number }[] }).__shifts.filter((s) => !s.recent && s.t >= t).reduce((a, s) => a + s.value, 0), since)
      const now = () => page.evaluate(() => performance.now())

      // document load: placeholders from the first paint, the chunk held for 1 s
      const { release } = await open(page, 'Report.docx', { hold: true })
      if (vp.touch && !(await coarse(page))) test.skip(true, 'this engine does not emulate a coarse pointer')
      const before = await geometry()
      expect(before.placeholder?.height).toBe(expected)
      expect(before.slot?.height).toBe(24)
      expect(before.bar).toBeNull()
      await page.waitForTimeout(1000)
      const t0 = await now()
      release()
      await toolbar(page).waitFor()
      await frames(page)
      const after = await geometry()
      expect(after.bar?.height).toBe(expected)
      expect(after.status?.height).toBe(24)
      expect(after.placeholder).toBeNull()
      for (const k of ['top', 'height', 'left', 'width'] as const) expect(after.container[k], `container ${k}`).toBeCloseTo(before.container[k], 1)
      if (browserName === 'chromium') expect(await cls(t0), 'CLS when the chunk landed').toBeLessThanOrEqual(0.001)

      // showToolbar toggle: from off to on with the chunk held again (a fresh page)
      await page.unrouteAll({ behavior: 'ignoreErrors' })
      const toggled = await open(page, 'Report.docx', { hold: true, query: { toolbar: '0' } })
      expect((await geometry()).placeholder).toBeNull()
      await page.evaluate(() => (window as unknown as { __viewerE2E: { setToolbar(on: boolean): void } }).__viewerE2E.setToolbar(true))
      await frames(page)
      const held = await geometry()
      expect(held.placeholder?.height).toBe(expected)
      expect(held.slot?.height).toBe(24)
      await page.waitForTimeout(1000)
      const t1 = await now()
      toggled.release()
      await toolbar(page).waitFor()
      await frames(page)
      const landed = await geometry()
      expect(landed.bar?.height).toBe(expected)
      for (const k of ['top', 'height'] as const) expect(landed.container[k], `container ${k}`).toBeCloseTo(held.container[k], 1)
      if (browserName === 'chromium') expect(await cls(t1), 'CLS when the chunk landed after the toggle').toBeLessThanOrEqual(0.001)
    })
  })
}

test.describe('compact 375x667 coarse', () => {
  test.use({ viewport: { width: 375, height: 667 }, hasTouch: true, isMobile: true })
  test('coarse: every toolbar control is at least 44×44', async ({ page }) => {
    for (const name of ['Report.docx', 'Deck.pptx']) {
      await open(page, name)
      test.skip(!(await coarse(page)), 'this engine does not emulate a coarse pointer')
      const sizes = async () => toolbar(page).evaluate((bar) => [...bar.querySelectorAll('button, input')].map((el) => ({ label: el.getAttribute('aria-label') ?? el.textContent, ...el.getBoundingClientRect().toJSON() as Rect })))
      for (const s of await sizes()) {
        expect(s.height, `${name} ${s.label}`).toBeGreaterThanOrEqual(44)
        expect(s.width, `${name} ${s.label}`).toBeGreaterThanOrEqual(44)
      }
      // every control fits inside the pane (nothing pushed off-screen)
      for (const s of await sizes()) expect(s.right, `${name} ${s.label} right edge`).toBeLessThanOrEqual(375 + 0.5)
      await toolbar(page).getByLabel('More options').click()
      await motionSettled(page)
      for (const item of await page.getByRole('menu').evaluate((m) => [...m.querySelectorAll('[role^=menuitem]')].map((el) => ({ label: el.textContent, ...el.getBoundingClientRect().toJSON() as Rect })))) {
        expect(item.height, `${name} menu ${item.label}`).toBeGreaterThanOrEqual(44)
      }
      await page.keyboard.press('Escape')
      await toolbar(page).getByLabel('Find').click()
      await motionSettled(page) // the closed menu has finished its exit
      for (const s of await sizes()) {
        expect(s.height, `${name} expanded ${s.label}`).toBeGreaterThanOrEqual(44)
        expect(s.right, `${name} expanded ${s.label} right edge`).toBeLessThanOrEqual(375 + 0.5)
      }
      await page.unrouteAll({ behavior: 'ignoreErrors' })
    }
  })
})

test.describe('desktop 1280x800', () => {
  test.use({ viewport: { width: 1280, height: 800 } })

  for (const theme of ['light', 'dark'] as const) {
    test(`keyboard: Tab order through toolbar groups with a visible focus ring in light and dark (${theme})`, async ({ page, browserName }) => {
      // WebKit follows Safari's default: Tab reaches only text fields, Option-Tab every control
      const tab = browserName === 'webkit' ? 'Alt+Tab' : 'Tab'
      await open(page, 'Report.docx', { query: { theme } })
      // Tab from the start of the page walks the toolbar in DOM order, then the pages.
      await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur())
      const order: string[] = []
      for (let i = 0; i < 20; i++) {
        await page.keyboard.press(tab)
        const info = await page.evaluate(() => {
          const el = document.activeElement as HTMLElement
          const bar = document.querySelector('[data-testid=officeview-toolbar]')!
          const inBar = bar.contains(el)
          const cs = getComputedStyle(el)
          return { label: el.getAttribute('aria-label') ?? el.tagName, inBar, outlineStyle: cs.outlineStyle, outlineWidth: parseFloat(cs.outlineWidth), outlineColor: cs.outlineColor, background: getComputedStyle(bar).backgroundColor }
        })
        if (!info.inBar) { order.push(info.label); if (order.length > 1) break; continue }
        order.push(info.label)
        expect(info.outlineStyle, info.label).toBe('solid')
        expect(info.outlineWidth, info.label).toBeGreaterThanOrEqual(2)
        expect(contrast(info.outlineColor, info.background), `${theme} ${info.label} ring contrast`).toBeGreaterThanOrEqual(3)
      }
      // disabled controls (previous page, zoom out at 100%, match arrows, copy without a selection) are skipped
      expect(order).toEqual(['Next page', 'Reset zoom', 'Zoom in', 'Fit to width', 'Search document', 'Print', 'Document viewer'])
    })
  }

  test('fit-to-width keeps native wheel scrolling and drag-to-select', async ({ page }) => {
    await open(page, 'Report.docx')
    await toolbar(page).getByLabel('Fit to width').click()
    await expect(toolbar(page).getByLabel('Fit to width')).toHaveAttribute('aria-pressed', 'true')
    const state = () => region(page).evaluate((el) => ({ overflow: el.style.overflow, scrollTop: el.scrollTop, transform: (el.querySelector('[data-testid=officeview-content]') as HTMLElement).style.transform, width: el.querySelector('canvas')!.getBoundingClientRect().width, inner: el.clientWidth - 32 }))
    const fit = await state()
    expect(fit.overflow).toBe('auto')
    expect(fit.transform).toBe('none')
    expect(fit.width).toBeCloseTo(fit.inner, 0)
    const canvas = page.locator('[data-officeview-root] canvas').first()
    const box = (await canvas.boundingBox())!
    // drag across the first line selects text (not a pan)
    const s = box.width / 816
    await page.mouse.move(box.x + 100 * s, box.y + 104 * s)
    await page.mouse.down()
    await page.mouse.move(box.x + 300 * s, box.y + 104 * s, { steps: 8 })
    await page.mouse.up()
    await expect(toolbar(page).getByTestId('officeview-copy-selection')).toBeEnabled()
    expect((await state()).transform).toBe('none')
    // a plain wheel scrolls natively
    await page.mouse.move(box.x + box.width / 2, box.y + 300)
    await page.mouse.wheel(0, 600)
    await expect.poll(async () => (await state()).scrollTop).toBeGreaterThan(fit.scrollTop + 100)
    expect((await state()).transform).toBe('none')
  })

  test('fit-to-width follows window resize', async ({ page }) => {
    await open(page, 'Report.docx')
    await toolbar(page).getByLabel('Fit to width').click()
    for (const width of [1000, 1400, 900]) {
      await page.setViewportSize({ width, height: 800 })
      await expect.poll(() => region(page).evaluate((el) => Math.round(el.querySelector('canvas')!.getBoundingClientRect().width - (el.clientWidth - 32)))).toBe(0)
    }
  })

  test('two tall pages: status flips at the dominance crossover', async ({ page }) => {
    await open(page, 'Tall.docx')
    const pages = await page.locator('[data-officeview-root] canvas').count()
    expect(pages).toBeGreaterThanOrEqual(2)
    const first = await region(page).evaluate((el) => { const c = el.querySelectorAll('canvas'); return { top1: (c[1] as HTMLElement).offsetTop, h: el.clientHeight } })
    // walk across the boundary in 40px steps
    for (let y = first.top1 - first.h; y <= first.top1 + 40; y += 40) {
      await region(page).evaluate((el, top) => { el.scrollTop = top }, Math.max(0, y))
      await frames(page)
      const expected = await region(page).evaluate((el) => {
        const view = el.getBoundingClientRect()
        let best = 0, bestPx = -1
        el.querySelectorAll('canvas').forEach((c, i) => {
          const r = c.getBoundingClientRect()
          // the tracker's viewport is the container's client box
          const px = Math.min(view.top + el.clientHeight, r.bottom) - Math.max(view.top, r.top)
          if (px > bestPx) { best = i; bestPx = px }
        })
        return best + 1
      })
      await expect(statusLine(page)).toContainText(`Page ${expected} of ${pages}`)
    }
  })

  test('rapid scrolling a 100-page document: status updates at most once per frame', async ({ page }) => {
    test.setTimeout(240_000)
    await open(page, 'Hundred.docx')
    const total = await page.locator('[data-officeview-root] canvas').count()
    expect(total).toBeGreaterThanOrEqual(90)
    // the word count shares the status line and lands on its own schedule;
    // let it settle and watch only the page position
    await expect(page.getByTestId('officeview-word-count')).not.toHaveText(/^—/, { timeout: 60_000 })
    const r = await page.evaluate(async () => {
      const el = document.querySelector('[data-officeview-root]') as HTMLElement
      const status = document.querySelector('[data-testid=officeview-status-position]')!
      let frame = 0
      const changes: number[] = []
      const mo = new MutationObserver(() => changes.push(frame))
      mo.observe(status, { subtree: true, characterData: true, childList: true })
      await new Promise<void>((done) => {
        const tick = () => {
          frame++
          // several scroll steps inside one frame, as momentum scrolling delivers
          for (let i = 0; i < 4; i++) el.scrollTop += 120
          if (frame < 120 && el.scrollTop < el.scrollHeight - el.clientHeight) requestAnimationFrame(tick)
          else setTimeout(done, 100)
        }
        requestAnimationFrame(tick)
      })
      mo.disconnect()
      return { changes, frames: frame, text: status.textContent }
    })
    expect(r.changes.length).toBeGreaterThan(5)
    const perFrame = new Map<number, number>()
    for (const f of r.changes) perFrame.set(f, (perFrame.get(f) ?? 0) + 1)
    expect(Math.max(...perFrame.values())).toBeLessThanOrEqual(1)
  })

  test('native scroll to page 3, zoom in, goTo(5), zoom out: the tracked page and position stay correct', async ({ page }) => {
    await open(page, 'Report.docx')
    const total = await page.locator('[data-officeview-root] canvas').count()
    expect(total).toBeGreaterThanOrEqual(5)
    await region(page).evaluate((el) => { el.scrollTop = (el.querySelectorAll('canvas')[2] as HTMLElement).offsetTop })
    await expect(statusLine(page)).toContainText(`Page 3 of ${total}`)
    await toolbar(page).getByLabel('Zoom in').click()
    expect(await region(page).evaluate((el) => [el.scrollTop, el.style.overflow])).toEqual([0, 'hidden'])
    await frames(page)
    await expect(statusLine(page)).toContainText(`Page 3 of ${total}`)
    await toolbar(page).getByLabel('Next page').click()
    await expect(statusLine(page)).toContainText(`Page 4 of ${total}`)
    await toolbar(page).getByLabel('Next page').click()
    await expect(statusLine(page)).toContainText(`Page 5 of ${total}`)
    const top = await region(page).evaluate((el) => el.querySelectorAll('canvas')[4].getBoundingClientRect().top - el.getBoundingClientRect().top)
    expect(Math.abs(top)).toBeLessThanOrEqual(1)
    await toolbar(page).getByLabel('Zoom out').click()
    expect(await region(page).evaluate((el) => el.style.overflow)).toBe('auto')
    await frames(page)
    await expect(statusLine(page)).toContainText(`Page 5 of ${total}`)
    // native scrolling resumes
    const before = await region(page).evaluate((el) => el.scrollTop)
    await page.mouse.move(640, 400)
    await page.mouse.wheel(0, 300)
    await expect.poll(() => region(page).evaluate((el) => el.scrollTop)).toBeGreaterThan(before)
  })
})

test('hostile host sheet leaves paged toolbar and status geometry and colours unchanged', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 900 })
  const snap = async (query: string) => {
    await page.goto(`/src/debug/ui-host-isolation.html${query}`)
    await page.locator('[data-viewer=paged] canvas').first().waitFor()
    await page.locator('[data-viewer=paged] [data-testid=officeview-toolbar]').waitFor()
    await page.locator('[data-viewer=paged] [data-testid=officeview-status]').waitFor()
    await page.evaluate(() => document.fonts.ready)
    return page.evaluate(() => {
      const out: Record<string, unknown> = {}
      const els = document.querySelectorAll('[data-viewer=paged] :is([data-testid=officeview-toolbar], [data-testid=officeview-toolbar] *, [data-testid=officeview-status], [data-testid=officeview-status] *)')
      els.forEach((el, i) => {
        const cs = getComputedStyle(el)
        const b = el.getBoundingClientRect()
        out[`${i}:${el.tagName.toLowerCase()}:${el.getAttribute('aria-label') ?? el.getAttribute('data-testid') ?? el.getAttribute('data-icon') ?? ''}`] = {
          box: [b.x, b.y, b.width, b.height].map((v) => Math.round(v * 2) / 2),
          color: cs.color, backgroundColor: cs.backgroundColor, borderColor: cs.borderColor, padding: cs.padding, fontSize: cs.fontSize, boxSizing: cs.boxSizing,
        }
      })
      return out
    })
  }
  const clean = await snap('')
  expect(Object.keys(clean).length).toBeGreaterThan(20)
  for (const order of ['first', 'last']) expect(await snap(`?host=${order}`), `host=${order}`).toEqual(clean)
})

/** WCAG contrast ratio of two CSS rgb()/rgba() colours (alpha ignored). */
function contrast(a: string, b: string): number {
  const lum = (c: string) => {
    const [r, g, bl] = (c.match(/[\d.]+/g) ?? ['0', '0', '0']).slice(0, 3).map((v) => { const x = Number(v) / 255; return x <= 0.03928 ? x / 12.92 : ((x + 0.055) / 1.055) ** 2.4 })
    return 0.2126 * r + 0.7152 * g + 0.0722 * bl
  }
  const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x)
  return (hi + 0.05) / (lo + 0.05)
}

test.describe('backdrop token', () => {
  test.use({ viewport: { width: 1280, height: 800 } })
  test('the paged backdrop follows the theme and host token overrides; an explicit background prop wins; pages stay white', async ({ page }) => {
    const colours = () => region(page).evaluate((el) => ({ backdrop: getComputedStyle(el).backgroundColor, page: getComputedStyle(el.querySelector('canvas')!).backgroundColor }))
    await open(page, 'Report.docx', { query: { theme: 'light' } })
    expect(await colours()).toEqual({ backdrop: 'rgb(240, 239, 238)', page: 'rgb(255, 255, 255)' })
    await page.unrouteAll({ behavior: 'ignoreErrors' })
    await open(page, 'Report.docx', { query: { theme: 'dark' } })
    expect(await colours()).toEqual({ backdrop: 'rgb(22, 21, 20)', page: 'rgb(255, 255, 255)' })
    await page.addStyleTag({ content: '[data-ov-theme-root] { --ov-color-backdrop: #123456 }' })
    expect(await colours()).toEqual({ backdrop: 'rgb(18, 52, 86)', page: 'rgb(255, 255, 255)' })
    await page.unrouteAll({ behavior: 'ignoreErrors' })
    await open(page, 'Report.docx', { query: { theme: 'dark', background: '#654321' } })
    expect(await colours()).toEqual({ backdrop: 'rgb(101, 67, 33)', page: 'rgb(255, 255, 255)' })
  })
})

test.describe('narrow fine pointer 375x667', () => {
  test.use({ viewport: { width: 375, height: 667 } })
  test('compact on a fine pointer: toolbar controls and menu items are still 44px', async ({ page }) => {
    await open(page, 'Report.docx')
    expect(await coarse(page)).toBe(false)
    for (const s of await toolbar(page).evaluate((bar) => [...bar.querySelectorAll('button')].map((el) => ({ label: el.getAttribute('aria-label'), ...el.getBoundingClientRect().toJSON() as Rect })))) {
      expect(s.height, `${s.label}`).toBeGreaterThanOrEqual(44)
      expect(s.right, `${s.label} right edge`).toBeLessThanOrEqual(375 + 0.5)
    }
    await toolbar(page).getByLabel('More options').click()
    // Task 6: the menu enters with a short drop; measure it once settled.
    await motionSettled(page)
    for (const h of await page.getByRole('menu').evaluate((m) => [...m.querySelectorAll('[role^=menuitem]')].map((el) => el.getBoundingClientRect().height))) expect(h).toBeGreaterThanOrEqual(44)
  })
})

// Task 4: print is isolated from the host document. No system print dialog is
// opened by these tests; each new frame gets its own observed print function.
async function observePrint(page: Page) {
  await page.addInitScript(() => {
    const calls: Window[] = []
    Object.assign(window, {__printCalls: calls})
    window.print = () => {calls.push(window)}
    const append = Node.prototype.appendChild
    Node.prototype.appendChild = function<T extends Node>(node: T): T {
      const result = append.call(this, node) as T
      if (node instanceof HTMLIFrameElement && node.hasAttribute('data-officeview-print-frame')) {
        node.contentWindow!.print = () => {calls.push(node.contentWindow!)}
      }
      return result
    }
  })
}

test('paged print frame contains only page images at page size', async ({page}, info) => {
  await observePrint(page)
  await open(page, 'Deck.pptx')
  await page.evaluate(() => { const p = document.createElement('p'); p.textContent = 'HOST-SECRET'; document.body.append(p) })
  await toolbar(page).getByRole('button', {name:'Print', exact:true}).click()
  const frame = page.frameLocator('[data-officeview-print-frame]')
  await expect(frame.locator('img')).toHaveCount(3)
  await expect.poll(() => page.evaluate(() => (window as unknown as {__printCalls:Window[]}).__printCalls.length)).toBe(1)
  expect(await frame.locator('body').textContent()).not.toContain('HOST-SECRET')
  expect(await page.evaluate(() => (window as unknown as {__printCalls:Window[]}).__printCalls.includes(window))).toBe(false)
  const dimensions = await frame.locator('img').first().evaluate(img => ({w:img.getBoundingClientRect().width,h:img.getBoundingClientRect().height,nw:(img as HTMLImageElement).naturalWidth}))
  expect(dimensions.nw).toBe(dimensions.w * 2)
  await page.emulateMedia({media:'print'})
  // Reveal only for the evidence screenshot: production keeps the frame hidden.
  await page.locator('[data-officeview-print-frame]').evaluate((f, d) => {
    const e = f as HTMLIFrameElement
    e.style.cssText = `position:fixed;inset:0;width:${d.w}px;height:${d.h}px;opacity:1;border:0;z-index:99999`
  }, dimensions)
  const shot = await page.locator('[data-officeview-print-frame]').screenshot({path:info.outputPath('white-print-page.png')})
  const corner = await page.evaluate(async png => {
    const image = new Image()
    const loaded = new Promise<void>((resolve, reject) => { image.onload = () => resolve(); image.onerror = () => reject(new Error('screenshot decode failed')) })
    image.src = `data:image/png;base64,${png}`
    await loaded
    const canvas = document.createElement('canvas')
    canvas.width = canvas.height = 1
    const ctx = canvas.getContext('2d')!
    // The fixture's unpainted page corner must be white in the screenshot,
    // not merely declared white by the iframe body's computed style.
    ctx.drawImage(image, 2, 2, 1, 1, 0, 0, 1, 1)
    return [...ctx.getImageData(0, 0, 1, 1).data]
  }, shot.toString('base64'))
  expect(corner).toEqual([255, 255, 255, 255])
  expect(await frame.locator('body').evaluate(el => getComputedStyle(el).backgroundColor)).toBe('rgb(255, 255, 255)')
})

test('print cleanup without afterprint: next print removes the old frame', async ({page}) => {
  await observePrint(page)
  await open(page, 'Deck.pptx')
  await toolbar(page).getByRole('button', {name:'Print', exact:true}).click()
  await expect.poll(() => page.evaluate(() => (window as unknown as {__printCalls:Window[]}).__printCalls.length)).toBe(1)
  await page.locator('[data-officeview-print-frame]').evaluate(f => f.setAttribute('data-old-frame',''))
  await toolbar(page).getByRole('button', {name:'Print', exact:true}).click()
  await expect.poll(() => page.evaluate(() => (window as unknown as {__printCalls:Window[]}).__printCalls.length)).toBe(2)
  await expect(page.locator('[data-old-frame]')).toHaveCount(0)
  await expect(page.locator('[data-officeview-print-frame]')).toHaveCount(1)
})

test('page box size: named pages or the supported single-size fallback', async ({page, context, browserName}, info) => {
  await observePrint(page)
  await open(page, 'Deck.pptx')
  await page.emulateMedia({media:'print'})
  const result = await page.evaluate(async () => {
    const path = '/src/components/paged/printDocument.ts'
    const {printDocument} = await import(/* @vite-ignore */ path)
    const pages = [{widthPx:400,heightPx:600},{widthPx:600,heightPx:400},{widthPx:400,heightPx:600}].map(spec => ({spec,paint(ctx:CanvasRenderingContext2D){ctx.fillStyle='#111';ctx.fillRect(20,20,40,40)}}))
    const job = await printDocument(pages, document.querySelector('[data-ov-theme-root]'), {signal:new AbortController().signal,onNotice(text:string){throw new Error(text)}})
    Object.assign(window, {__disposePageBoxJob: () => job.dispose()})
    const doc = job.frame.contentDocument as Document
    return {html:doc.documentElement.outerHTML,named:CSS.supports('page','p0'),css:doc.querySelector('style')!.textContent,boxes:[...doc.images].map(i => ({w:i.getBoundingClientRect().width,h:i.getBoundingClientRect().height}))}
  })
  if (result.named) {
    expect(result.css).toContain('@page p0 { size: 400px 600px; margin: 0 }')
    expect(result.boxes).toEqual([{w:400,h:600},{w:600,h:400},{w:400,h:600}])
  } else {
    expect(result.css).toContain('@page { size: 600px 600px; margin: 0 }')
    expect(result.boxes).toEqual(Array(3).fill({w:600,h:600}))
  }
  if (browserName === 'chromium') {
    // Playwright exposes PDF only on Page, not Frame. Reproduce the frame's
    // exact HTML in an isolated same-origin page (blob URLs remain live).
    const pdfPage = await context.newPage()
    await pdfPage.goto('/src/debug/viewer-e2e.html')
    await pdfPage.setContent(result.html)
    await pdfPage.waitForFunction(() => [...document.images].every(i => i.complete && i.naturalWidth > 0))
    const pdf = await pdfPage.pdf({preferCSSPageSize:true,printBackground:true,path:info.outputPath('mixed-print.pdf')})
    const boxes = [...pdf.toString('latin1').matchAll(/\/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]/g)].map(m => [Number(m[1]),Number(m[2])])
    expect(boxes).toHaveLength(3)
    for (const [i, box] of boxes.entries()) {
      expect(Math.abs(box[0] - result.boxes[i].w * .75)).toBeLessThanOrEqual(1)
      expect(Math.abs(box[1] - result.boxes[i].h * .75)).toBeLessThanOrEqual(1)
    }
    await pdfPage.close()
  }
  await page.evaluate(() => (window as unknown as {__disposePageBoxJob():void}).__disposePageBoxJob())
  await expect(page.locator('[data-officeview-print-frame]')).toHaveCount(0)
})

for (const stylesheet of [true, false]) test(`denied print guard isolates peers (stylesheet=${stylesheet})`, async ({page}) => {
  await open(page, 'Deck.pptx')
  await page.evaluate(async ({stylesheet}) => {
    if (!stylesheet) {
      document.querySelectorAll('style,link[rel="stylesheet"]').forEach(el => el.remove())
    }
    const path = '/src/debug/paged-print-probe.tsx'
    const {mountPrintPeers} = await import(/* @vite-ignore */ path)
    const dispose = await mountPrintPeers(await (await fetch('/__fixtures/Deck.pptx')).blob())
    Object.assign(window, {__disposePrintPeers: dispose})
  }, {stylesheet})
  const denied = page.locator('[data-peer="denied"] [data-ov-instance]')
  await expect(denied).toHaveCount(1)
  // Wait for the lazy service, then test a competing host !important rule.
  await page.locator('[data-peer="allowed"]').getByRole('button', {name:'Print', exact:true}).waitFor()
  if (!stylesheet) await expect(page.locator('style[data-ov-print-guard]')).toHaveCount(1)
  await page.addStyleTag({content:'@media print { div { display:block !important } }'})
  await page.emulateMedia({media:'print'})
  await expect(denied).toHaveCSS('display', 'none')
  for (const peer of ['allowed', 'sheet']) {
    const root = page.locator(`[data-peer="${peer}"] [data-ov-instance]`)
    await expect(root).toHaveCount(1)
    expect(await root.evaluate(el => getComputedStyle(el).display)).not.toBe('none')
  }
  await page.evaluate(() => (window as unknown as {__disposePrintPeers():void}).__disposePrintPeers())
  await expect(page.locator('style[data-ov-print-guard]')).toHaveCount(0)
})

test('print cleanup without afterprint: React unmount removes the frame', async ({page}) => {
  await observePrint(page)
  await open(page, 'Deck.pptx')
  await page.evaluate(async () => {
    const path = '/src/debug/paged-print-probe.tsx'
    const {mountPrintPeers} = await import(/* @vite-ignore */ path)
    Object.assign(window, {__disposePrintPeers: await mountPrintPeers(await (await fetch('/__fixtures/Deck.pptx')).blob())})
  })
  await page.locator('[data-peer="allowed"]').getByRole('button', {name:'Print', exact:true}).click()
  await expect.poll(() => page.evaluate(() => (window as unknown as {__printCalls:Window[]}).__printCalls.length)).toBe(1)
  await expect(page.locator('[data-officeview-print-frame]')).toHaveCount(1)
  await page.evaluate(() => (window as unknown as {__disposePrintPeers():void}).__disposePrintPeers())
  await expect(page.locator('[data-officeview-print-frame]')).toHaveCount(0)
})

test('dark: paged page pixels are white and the backdrop is dark', async ({page}) => {
  await open(page,'Report.docx',{query:{theme:'dark'}})
  await expect(region(page)).toHaveCSS('background-color','rgb(22, 21, 20)')
  const shot=await region(page).locator('canvas').first().screenshot()
  const corner=await page.evaluate(async png=>{
    const img=new Image(), loaded=new Promise<void>((resolve,reject)=>{img.onload=()=>resolve();img.onerror=()=>reject(new Error('PNG load'))})
    img.src=`data:image/png;base64,${png}`; await loaded
    const c=document.createElement('canvas');c.width=c.height=1
    const ctx=c.getContext('2d')!;ctx.drawImage(img,2,2,1,1,0,0,1,1)
    return [...ctx.getImageData(0,0,1,1).data]
  },shot.toString('base64'))
  expect(corner).toEqual([255,255,255,255])
})

test('host --ov-color-find override reaches the paged highlights without replacing the canvas', async ({page}) => {
  await open(page,'Report.docx')
  const canvas=region(page).locator('canvas').first()
  await canvas.evaluate(c=>c.setAttribute('data-palette-original',''))
  await page.getByRole('searchbox',{name:'Search document'}).fill('Line')
  // C1 flake hardening: wait for the search to have matches (a DOM signal) before overriding the colour, so the
  // poll below only waits for the palette repaint, not for search indexing as well. Under load the two together
  // overran the 15 s poll. The repaint itself is driven by a MutationObserver on the theme root's style, so once
  // matches exist the pixel appears on the next frame; the poll still allows the suite's slowest observed frame.
  await expect(page.getByRole('button',{name:'Next match'})).toBeEnabled()
  await page.locator('[data-ov-theme-root]').evaluate(el=>(el as HTMLElement).style.setProperty('--ov-color-find','#1234ff'))
  await expect.poll(()=>canvas.evaluate(node=>{
    const c=node as HTMLCanvasElement,d=c.getContext('2d')!.getImageData(0,0,c.width,c.height).data
    for(let i=0;i<d.length;i+=4) if(d[i]===18&&d[i+1]===52&&d[i+2]===255&&d[i+3]===255) return true
    return false
  }),{timeout:30_000}).toBe(true)
  await expect(canvas).toHaveAttribute('data-palette-original','')
})
