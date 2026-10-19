/**
 * PR #11 visual review, stream A: real-browser layout evidence for the
 * interactive viewer (host fill #6, popup/toolbar #13, headers #9/#21, status
 * line #18, chrome font #19, print #12). Runs on the viewer-e2e page, whose
 * host (#root) is a 100%-height flex column.
 */
import { expect, test, type Page } from '@playwright/test'
import { ensureFixtures, goTo, openViewer } from './viewer-e2e'

type Box = { left: number; right: number; top: number; bottom: number; width: number; height: number }
const rect = (page: Page, selector: string) => page.evaluate(s => document.querySelector(s)!.getBoundingClientRect().toJSON() as Box, selector)
const overlaps = (a: Box, b: Box) => a.left < b.right - 0.5 && b.left < a.right - 0.5 && a.top < b.bottom - 0.5 && b.top < a.bottom - 0.5

test.describe('visual review: interactive viewer layout', () => {
  test.beforeAll(async () => { await ensureFixtures() })

  test('#6/#18 fills a sized flex host and shows the status line inside it', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    const wrapper = await rect(page, '[data-officeview-sheet-instance]'), pane = await rect(page, '[data-testid=officeview-sheet-pane]')
    expect(wrapper.bottom).toBeGreaterThan(800 - 1)
    expect(pane.bottom).toBeGreaterThan(800 - 1)
    const status = await rect(page, '[data-officeview-ui=viewer-status]')
    expect(status.height).toBeGreaterThan(10)
    expect(status.top).toBeGreaterThanOrEqual(pane.top)
    expect(status.bottom).toBeLessThanOrEqual(pane.bottom + 0.5)
    await expect(page.locator('[data-officeview-ui=viewer-status]')).toBeVisible()
  })

  test('#19 chrome controls and labels use the viewer font, never the host serif', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const fonts = await page.evaluate(() => ['[data-officeview-ui=sheet-navigator] label', '[data-officeview-ui=sheet-navigator] button', '[data-officeview-ui=sheet-navigator] select', '[data-testid=officeview-search] input']
      .map(s => getComputedStyle(document.querySelector(s)!).fontFamily))
    for (const font of fonts) expect(font).toMatch(/^system-ui/)
  })

  test('#13 the filter popup opens under its column, inside the cell area, with OK and Cancel visible', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    await page.getByRole('button', { name: /^Filter column D/ }).click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await expect(dialog.getByRole('checkbox', { name: 'alpha', exact: true })).toBeVisible()
    const popup = await rect(page, '[data-officeview-ui=grid-popup]'), header = await page.getByRole('columnheader', { name: /^Column D/ }).boundingBox()
    const headers = await rect(page, '[data-officeview-ui=grid-headers]'), status = await rect(page, '[data-officeview-ui=viewer-status]')
    expect(Math.abs(popup.left - header!.x)).toBeLessThan(1)
    expect(popup.top).toBeGreaterThanOrEqual(headers.bottom - 0.5)
    expect(popup.bottom).toBeLessThanOrEqual(status.top + 0.5)
    for (const name of ['OK', 'Cancel']) {
      const button = await dialog.getByRole('button', { name, exact: true }).boundingBox()
      expect(button!.y + button!.height).toBeLessThanOrEqual(popup.bottom + 0.5)
      expect(button!.y).toBeGreaterThanOrEqual(popup.top)
    }
    await expect(page.getByTestId('officeview-zoom-controls')).toHaveCount(0)
  })

  test.describe('375px touch viewport', () => {
    test.use({ viewport: { width: 375, height: 667 }, hasTouch: true, isMobile: true, deviceScaleFactor: 2 })
    test('#13 toolbar controls never overlap; the filter sheet is not covered by the zoom buttons', async ({ page }) => {
      await openViewer(page, 'f2-1k.xlsx')
      const boxes = await page.evaluate(() => ['[aria-label=Worksheet]', '[aria-label="Address space"]', '[aria-label="Go To cell"]', '[data-officeview-ui=sheet-navigator] button', '[data-officeview-ui=sheet-actions]', '[data-testid=officeview-search]']
        .map(s => ({ s, ...document.querySelector(s)!.getBoundingClientRect().toJSON() })))
      for (const [i, a] of boxes.entries()) for (const b of boxes.slice(i + 1)) expect(overlaps(a, b), `${a.s} overlaps ${b.s}`).toBe(false)
      const coarse = await page.evaluate(() => matchMedia('(pointer: coarse)').matches)
      if (coarse) for (const s of ['[aria-label="Search document"]', '[aria-label="Next match"]', '[aria-label=Worksheet]']) {
        const box = await rect(page, s)
        expect(box.height, s).toBeGreaterThanOrEqual(44)
      }
      await goTo(page, 'D2')
      await page.getByRole('grid').focus()
      await page.keyboard.press('Alt+ArrowDown')
      // The bottom sheet is fixed to the viewport (its popup container has no box of its own).
      const sheet = page.locator('[data-officeview-ui=filter-dropdown]')
      await expect(sheet).toBeVisible()
      await expect(sheet.getByRole('button', { name: 'OK', exact: true })).toBeInViewport()
      await expect(page.getByTestId('officeview-zoom-controls')).toHaveCount(0)
    })
  })

  test('#9 header labels never truncate and action glyphs never crowd them (100%, 50%, far corner)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    const check = async () => {
      const result = await page.evaluate(() => {
        const labels = [...document.querySelectorAll<HTMLElement>('[data-header-column] > span')].filter(s => s.getBoundingClientRect().width > 0)
        const truncated = labels.filter(s => s.scrollWidth > s.clientWidth + 0.5).map(s => s.textContent)
        // A visible action glyph never overlaps its column's label text.
        const crowded = labels.filter(s => {
          const header = s.closest('[role=columnheader]')!, label = s.getBoundingClientRect()
          const text = document.createRange(); text.selectNodeContents(s); const ink = text.getBoundingClientRect()
          return [...header.querySelectorAll<HTMLElement>('[data-action-column]')].some(b => { const r = b.getBoundingClientRect(); return r.width > 1 && r.left < ink.right - 0.5 && ink.left < r.right && label.width > 0 })
        }).map(s => s.textContent)
        const rows = [...document.querySelectorAll<HTMLElement>('[data-header-row] > span')].filter(s => s.scrollWidth > s.clientWidth + 0.5).map(s => s.textContent)
        return { count: labels.length, truncated, crowded, rows }
      })
      expect(result.count).toBeGreaterThan(0)
      expect(result.truncated).toEqual([])
      expect(result.crowded).toEqual([])
      expect(result.rows).toEqual([])
    }
    await check()
    for (let i = 0; i < 3; i++) await page.getByRole('button', { name: 'Zoom out' }).click()
    await expect(page.getByRole('button', { name: 'Reset zoom' })).toHaveText(/50%|25%/)
    await check()
    // Keyboard reach: F2 control mode still reaches a compact column's actions.
    await page.getByRole('grid').focus()
    await page.keyboard.press('F2')
    const sort = page.getByRole('button', { name: 'Cycle sort column A', exact: true })
    await expect(sort).toHaveCount(1)
    await sort.focus()
    expect((await sort.boundingBox())!.width).toBeGreaterThan(4) // revealed on focus
    await page.getByRole('button', { name: 'Reset zoom' }).click()
    await goTo(page, 'XFD1048576')
    await expect(page.getByRole('rowheader', { name: 'Row 1048576' })).toBeVisible()
    await check()
    const farCell = await rect(page, '[role=gridcell][aria-selected=true]'), zoom = await rect(page, '[data-testid=officeview-zoom-controls]')
    expect(overlaps(farCell, zoom)).toBe(false)
    expect(farCell.right).toBeLessThanOrEqual(zoom.left + 0.5)
    const status = await rect(page, '[data-officeview-ui=viewer-status]')
    expect(overlaps(status, zoom)).toBe(false)
    expect(zoom.bottom).toBeLessThanOrEqual(status.top + 0.5)
    await page.getByRole('grid').focus()
    await page.keyboard.press('ArrowLeft')
    await page.keyboard.press('ArrowRight')
    const navCell = await rect(page, '[role=gridcell][aria-selected=true]')
    expect(overlaps(navCell, zoom)).toBe(false)
    expect(overlaps(status, zoom)).toBe(false)
  })

  test('header buttons carry the explicit label colour instead of the UA button colour (browser)', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const color = await page.evaluate(() => {
      const btn = document.querySelector<HTMLElement>('[data-header-column]')
      return btn ? getComputedStyle(btn).color : null
    })
    expect(color).toBe('rgb(50, 49, 48)')
  })

  test('desktop header action icons are fully visible inside their buttons and not clipped (browser)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')

    const checkActionIcons = async (description: string) => {
      const icons = await page.evaluate(() => {
        const buttons = [...document.querySelectorAll<HTMLElement>('[data-action-column]')].filter(b => b.getBoundingClientRect().width >= 14)
        return buttons.map(btn => {
          const btnRect = btn.getBoundingClientRect()
          const span = btn.querySelector('span')
          const spanRect = span?.getBoundingClientRect()
          const svg = btn.querySelector('svg')
          const svgRect = svg?.getBoundingClientRect()
          return {
            label: btn.getAttribute('aria-label') ?? '',
            btnWidth: btnRect.width,
            spanWidth: spanRect?.width ?? 0,
            svgWidth: svgRect?.width ?? 0,
            svgFill: svg?.getAttribute('fill'),
          }
        })
      })
      expect(icons.length, `${description} icon count`).toBeGreaterThan(0)
      for (const item of icons) {
        expect(item.svgWidth, `${description} ${item.label} svg width`).toBeGreaterThanOrEqual(14)
        expect(item.spanWidth, `${description} ${item.label} span width`).toBeGreaterThanOrEqual(14)
        expect(item.btnWidth, `${description} ${item.label} button width`).toBeGreaterThanOrEqual(14)
      }
      return icons
    }

    await checkActionIcons('initial')

    const filterBtn = page.getByRole('button', { name: /^Filter column A/ })
    await filterBtn.click()
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: 'Deselect All', exact: true }).click()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)

    const activeIcons = await checkActionIcons('filtered')
    const activeFilter = activeIcons.find(i => /Filter column A/.test(i.label))
    expect(activeFilter).toBeDefined()
    expect(activeFilter!.svgFill).toBe('currentColor')
    expect(activeFilter!.svgWidth).toBeGreaterThanOrEqual(14)
    expect(activeFilter!.spanWidth).toBeGreaterThanOrEqual(14)
  })

  test('find scroll-into-view keeps row 1 visible for visible hits and aligns below-the-fold hits to whole rows', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await openViewer(page, 'f2-1k.xlsx')
    const headers = await rect(page, '[data-officeview-ui=grid-headers]')
    const search = page.getByRole('searchbox', { name: 'Search document' })
    await search.fill('alpha')
    await expect(page.getByRole('rowheader', { name: 'Row 1', exact: true })).toBeVisible()
    const row1 = await rect(page, '[role=row][aria-rowindex="2"]')
    expect(Math.abs(row1.top - headers.bottom)).toBeLessThan(1)
    await search.fill('total')
    await expect(page.getByRole('rowheader', { name: 'Row 1000', exact: true })).toBeVisible()
    const topBodyRow = await page.evaluate(() => {
      const rows = [...document.querySelectorAll<HTMLElement>('[role=row][aria-rowindex]:not([data-officeview-ui=grid-headers]):not([data-transient-active=true])')].map(r => r.getBoundingClientRect()).filter(r => r.height > 0).sort((a, b) => a.top - b.top)
      return rows[0]?.toJSON() as Box | undefined
    })
    expect(topBodyRow).toBeDefined()
    expect(Math.abs(topBodyRow!.top - headers.bottom)).toBeLessThan(1)
  })

  test('#21 header labels meet WCAG AA against the header strip', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const ratios = await page.evaluate(() => {
      const parse = (c: string) => (c.match(/[\d.]+/g) ?? []).slice(0, 3).map(Number)
      const lum = (c: number[]) => { const [r, g, b] = c.map(v => v / 255).map(v => v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4); return 0.2126 * r + 0.7152 * g + 0.0722 * b }
      const ratio = (a: number[], b: number[]) => { const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x); return (hi + 0.05) / (lo + 0.05) }
      const strip = parse(getComputedStyle(document.querySelector('[data-officeview-ui=grid-headers]')!).backgroundColor)
      return [...document.querySelectorAll('[data-header-column], [data-header-row]')].slice(0, 20).map(el => ratio(parse(getComputedStyle(el).color), strip))
    })
    expect(ratios.length).toBeGreaterThan(0)
    for (const r of ratios) expect(r).toBeGreaterThanOrEqual(4.5)
  })

  test('header and status computed colours meet contrast in light and dark (browser)', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    const checkContrast = async (themeName: string) => {
      const res = await page.evaluate(() => {
        const parse = (c: string) => (c.match(/[\d.]+/g) ?? []).slice(0, 3).map(Number)
        const lum = (c: number[]) => { const [r, g, b] = c.map(v => v / 255).map(v => v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4); return 0.2126 * r + 0.7152 * g + 0.0722 * b }
        const ratio = (a: number[], b: number[]) => { const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x); return (hi + 0.05) / (lo + 0.05) }

        const headerStripRaw = getComputedStyle(document.querySelector('[data-officeview-ui=grid-headers]')!).backgroundColor
        const headerStrip = parse(headerStripRaw)
        const headerRatios = [...document.querySelectorAll('[data-header-column], [data-header-row]')].slice(0, 20).map(el => ratio(parse(getComputedStyle(el).color), headerStrip))

        const statusEl = document.querySelector('[data-officeview-ui=viewer-status]') as HTMLElement
        const statusText = parse(getComputedStyle(statusEl).color)
        const statusBg = parse(getComputedStyle(statusEl).backgroundColor)
        const statusRatio = ratio(statusText, statusBg)

        return { headerRatios, statusRatio, headerStrip: headerStripRaw }
      })
      expect(res.headerRatios.length, `${themeName} headers`).toBeGreaterThan(0)
      for (const r of res.headerRatios) expect(r, `${themeName} header`).toBeGreaterThanOrEqual(4.5)
      expect(res.statusRatio, `${themeName} status`).toBeGreaterThanOrEqual(4.5)
      return res
    }

    const light = await checkContrast('light')
    await page.emulateMedia({ colorScheme: 'dark' })
    const dark = await checkContrast('dark')
    expect(light.headerStrip).not.toBe(dark.headerStrip)
  })

  test.describe('dark scheme (browser)', () => {
    test('#9 header labels never truncate and action glyphs never crowd them in dark scheme', async ({ page }) => {
      await page.setViewportSize({ width: 1280, height: 800 })
      await openViewer(page, 'f2-1k.xlsx')
      await page.emulateMedia({ colorScheme: 'dark' })
      const result = await page.evaluate(() => {
        const labels = [...document.querySelectorAll<HTMLElement>('[data-header-column] > span')].filter(s => s.getBoundingClientRect().width > 0)
        const truncated = labels.filter(s => s.scrollWidth > s.clientWidth + 0.5).map(s => s.textContent)
        const crowded = labels.filter(s => {
          const header = s.closest('[role=columnheader]')!, label = s.getBoundingClientRect()
          const text = document.createRange(); text.selectNodeContents(s); const ink = text.getBoundingClientRect()
          return [...header.querySelectorAll<HTMLElement>('[data-action-column]')].some(b => { const r = b.getBoundingClientRect(); return r.width > 1 && r.left < ink.right - 0.5 && ink.left < r.right && label.width > 0 })
        }).map(s => s.textContent)
        return { count: labels.length, truncated, crowded }
      })
      expect(result.count).toBeGreaterThan(0)
      expect(result.truncated).toEqual([])
      expect(result.crowded).toEqual([])
    })

    test('#13 toolbar controls never overlap in dark scheme', async ({ page }) => {
      await openViewer(page, 'f2-1k.xlsx')
      await page.emulateMedia({ colorScheme: 'dark' })
      const boxes = await page.evaluate(() => ['[aria-label=Worksheet]', '[aria-label="Address space"]', '[aria-label="Go To cell"]', '[data-officeview-ui=sheet-navigator] button', '[data-officeview-ui=sheet-actions]', '[data-testid=officeview-search]']
        .map(s => ({ s, ...document.querySelector(s)!.getBoundingClientRect().toJSON() })))
      for (const [i, a] of boxes.entries()) for (const b of boxes.slice(i + 1)) expect(overlaps(a, b), `${a.s} overlaps ${b.s}`).toBe(false)
    })
  })

  test('#12 print: white page, 2x canvases, a titled page per sheet, controls hidden', async ({ page, browserName }) => {
    await page.addInitScript(() => { (window as unknown as { print: () => void }).print = () => { (window as unknown as { __printed: number }).__printed = 1 } })
    await openViewer(page, 'f7-two-sheets.xlsx')
    await page.getByRole('button', { name: 'Print', exact: true }).click()
    await page.waitForFunction(() => (window as unknown as { __printed?: number }).__printed === 1)
    await page.emulateMedia({ media: 'print' })
    await expect(page.getByTestId('officeview-search')).toBeHidden()
    await expect(page.getByTestId('officeview-zoom-controls')).toBeHidden()
    const facts = await page.evaluate(() => ({
      background: getComputedStyle(document.querySelector('[data-officeview-root]')!).backgroundColor,
      canvases: [...document.querySelectorAll<HTMLCanvasElement>('[data-testid=officeview-sheet-print-pages] canvas')].map(c => ({ ratio: c.width / c.getBoundingClientRect().width, css: c.getBoundingClientRect().width })),
      titles: [...document.querySelectorAll('[data-officeview-print-sheet-title]')].map(h => h.textContent),
    }))
    expect(facts.background).toMatch(/rgb\(255, 255, 255\)|rgba\(0, 0, 0, 0\)/)
    expect(facts.titles).toEqual(['A', 'B'])
    for (const c of facts.canvases) expect(c.ratio).toBeGreaterThanOrEqual(2 - 1e-6)
    if (browserName === 'chromium') {
      const pdf = await page.pdf({ format: 'A4' })
      const pages = (pdf.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
      expect(pages).toBeGreaterThanOrEqual(2)
    }
    // Browser-menu print (nothing prepared): the notice, with no floating control over it.
    await page.emulateMedia({ media: 'screen' })
    await page.reload()
    await expect(page.getByRole('grid')).toBeVisible({ timeout: 60_000 })
    await page.emulateMedia({ media: 'print' })
    await expect(page.getByTestId('officeview-sheet-print-notice')).toBeVisible()
    await expect(page.getByTestId('officeview-search')).toBeHidden()
  })
})

test.describe('canvas paper and palettes', () => {
  test.beforeAll(async () => { await ensureFixtures() })
  const settle = (page: Page) => page.evaluate(async () => { await document.fonts.ready; await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))) })
  const base = '[data-testid="officeview-sheet-base"]'
  const overlay = '[data-testid="officeview-sheet-overlay"]'
  const hasPixel = (page: Page, rgb: number[]) => page.locator(overlay).evaluate((node, rgb) => {
    const c=node as HTMLCanvasElement, data=c.getContext('2d')!.getImageData(0,0,c.width,c.height).data
    for(let i=0;i<data.length;i+=4) if(data[i]===rgb[0] && data[i+1]===rgb[1] && data[i+2]===rgb[2] && data[i+3]>0) return true
    return false
  },rgb)
  // Find the active outline, then require visible dark glyph pixels inside it.
  // Merely finding the accent colour also passes an opaque fill hiding the text.
  const hasActivePixel = (page: Page) => page.locator(overlay).evaluate(node => {
    const canvas = node as HTMLCanvasElement, root = canvas.closest('[data-ov-theme-root]')!
    const style = getComputedStyle(root), sample = document.createElement('canvas').getContext('2d')!
    sample.fillStyle = style.getPropertyValue('--ov-color-find-active'); sample.fillRect(0,0,1,1)
    sample.fillStyle = style.getPropertyValue('--ov-color-selection-fill'); sample.fillRect(0,0,1,1)
    const accent = sample.getImageData(0,0,1,1).data
    const pixels = canvas.getContext('2d')!.getImageData(0,0,canvas.width,canvas.height).data
    const base = root.querySelector<HTMLCanvasElement>('[data-testid="officeview-sheet-base"]')!
    const paper = base.getContext('2d')!.getImageData(0,0,base.width,base.height).data
    let left=canvas.width, top=canvas.height, right=-1, bottom=-1
    for(let y=0;y<canvas.height;y++) for(let x=0;x<canvas.width;x++) {
      const i=(y*canvas.width+x)*4
      if(pixels[i]===accent[0] && pixels[i+1]===accent[1] && pixels[i+2]===accent[2]) {
        left=Math.min(left,x);right=Math.max(right,x);top=Math.min(top,y);bottom=Math.max(bottom,y)
      }
    }
    const lum = (rgb:number[]) => rgb.map(v=>v/255).map(v=>v<=.04045?v/12.92:((v+.055)/1.055)**2.4).reduce((n,v,i)=>n+v*[.2126,.7152,.0722][i],0)
    for(let y=top+2;y<bottom-1;y++) for(let x=left+2;x<right-1;x++) {
      const i=(y*canvas.width+x)*4, alpha=pixels[i+3]/255
      if(alpha<=0 || alpha>=.63 || paper[i]>32 || paper[i+1]>32 || paper[i+2]>32) continue
      const glyph=[0,1,2].map(k=>pixels[i+k]*alpha+paper[i+k]*(1-alpha))
      const background=[0,1,2].map(k=>pixels[i+k]*alpha+255*(1-alpha))
      if((lum(background)+.05)/(lum(glyph)+.05)>=4.5) return true
    }
    return false
  })
  test('host background prop overrides the sheet backdrop in dark theme', async ({page}) => {
    await openViewer(page,'f2-1k.xlsx')
    const url = new URL(page.url())
    url.searchParams.set('theme','dark'); url.searchParams.set('background','#123456')
    await page.goto(url.toString())
    await expect(page.getByRole('grid')).toBeVisible()
    await expect(page.locator('[data-officeview-root]')).toHaveCSS('background-color','rgb(18, 52, 86)')
  })
  test('dark: sheet paper and gridlines are pixel-identical to light; headers and backdrop are dark', async ({page}) => {
    await page.emulateMedia({colorScheme:'light'})
    await openViewer(page,'f2-1k.xlsx'); await settle(page)
    const pixels=await page.locator(base).evaluate(c=>(c as HTMLCanvasElement).toDataURL())
    const header=page.locator('[data-officeview-ui="grid-headers"]')
    await expect(header).toHaveCSS('background-color','rgb(243, 242, 241)')
    await page.emulateMedia({colorScheme:'dark'})
    await expect(header).toHaveCSS('background-color','rgb(32, 31, 30)'); await settle(page)
    expect(await page.locator(base).evaluate(c=>(c as HTMLCanvasElement).toDataURL())).toBe(pixels)
    await expect(page.locator('[data-officeview-root]')).toHaveCSS('background-color','rgb(22, 21, 20)')
    const tones=await page.locator(base).evaluate(c=>{
      const canvas=c as HTMLCanvasElement,d=canvas.getContext('2d')!.getImageData(0,0,canvas.width,canvas.height).data
      let white=false,grid=false
      for(let i=0;i<d.length;i+=4) { white ||= d[i]===255&&d[i+1]===255&&d[i+2]===255&&d[i+3]===255; grid ||= d[i]===208&&d[i+1]===208&&d[i+2]===208 }
      return {white,grid}
    })
    expect(tones).toEqual({white:true,grid:true})
  })
  test('selection and find highlights are visible on paper in both themes', async ({page}) => {
    await openViewer(page,'f2-1k.xlsx')
    for(const colorScheme of ['light','dark'] as const) {
      await page.emulateMedia({colorScheme})
      await goTo(page,'A2')
      await expect.poll(()=>hasPixel(page,[16,124,65])).toBe(true)
      await page.getByRole('searchbox',{name:'Search document'}).fill('alpha')
      await expect.poll(()=>hasActivePixel(page)).toBe(true)
      await page.getByRole('searchbox',{name:'Search document'}).fill('')
    }
  })
  test('host --ov-color-find override reaches the highlights', async ({page}) => {
    await openViewer(page,'f2-1k.xlsx')
    await page.getByRole('searchbox',{name:'Search document'}).fill('alpha')
    await expect.poll(()=>hasActivePixel(page)).toBe(true)
    await page.locator('[data-officeview-sheet-instance]').evaluate(el=>(el as HTMLElement).style.setProperty('--ov-color-find','#1234ff'))
    const rgb = await page.locator('[data-ov-theme-root]').evaluate(root => {
      const ctx = document.createElement('canvas').getContext('2d')!
      ctx.globalAlpha = .75
      ctx.fillStyle = '#1234ff'; ctx.fillRect(0,0,1,1)
      ctx.globalAlpha = 1
      ctx.fillStyle = getComputedStyle(root).getPropertyValue('--ov-color-selection-fill'); ctx.fillRect(0,0,1,1)
      return [...ctx.getImageData(0,0,1,1).data].slice(0,3)
    })
    await expect.poll(()=>hasPixel(page,rgb)).toBe(true)
  })
  test('CLS 0 on an emulateMedia colour-scheme flip', async ({page,browserName}) => {
    await page.emulateMedia({colorScheme:'light'})
    await openViewer(page,'f2-1k.xlsx'); await settle(page)
    const before=await page.locator('[data-officeview-sheet-instance]').boundingBox()
    await page.evaluate(() => {
      Object.assign(window,{__paletteCLS:0})
      if(PerformanceObserver.supportedEntryTypes.includes('layout-shift')) new PerformanceObserver(list=>{
        for(const e of list.getEntries()) (window as unknown as {__paletteCLS:number}).__paletteCLS+=(e as PerformanceEntry & {value:number}).value
      }).observe({type:'layout-shift'})
    })
    await page.emulateMedia({colorScheme:'dark'}); await settle(page)
    await expect(page.locator('[data-officeview-ui="grid-headers"]')).toHaveCSS('background-color','rgb(32, 31, 30)')
    expect(await page.locator('[data-officeview-sheet-instance]').boundingBox()).toEqual(before)
    if(browserName==='chromium') expect(await page.evaluate(()=>(window as unknown as {__paletteCLS:number}).__paletteCLS)).toBe(0)
  })
})
