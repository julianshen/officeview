/**
 * Task 2 (spec §6.3, §6.8, R4 test split): computed geometry and colours of
 * the shared viewer controls (SearchBar, ZoomControls, copy pill) with the
 * real stylesheet. The jsdom suite (tests/viewer-chrome.test.tsx) asserts the
 * classes; these are the browser twins of its former inline-size checks.
 * Task 3 moved the paged viewer's controls into its toolbar (inline Find,
 * toolbar-sized zoom buttons, compact layout on coarse pointers); the paged
 * toolbar's own geometry lives in paged-toolbar.spec.ts, and the floating
 * sizes here are the sheet viewer's.
 *
 * Page: src/debug/ui-host-isolation.html (sheet + paged viewers inside a host
 * `<div class="app">`; `host=first|last` adds the hostile host sheet).
 */
import { expect, test, type Page } from '@playwright/test'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const MAIN_SIZES = JSON.parse(readFileSync(fileURLToPath(new URL('./fixtures/chrome-sizes.main.json', import.meta.url)), 'utf8')) as Record<string, { width: number; height: number }>
const LABELS = ['Search document', 'Previous match', 'Next match', 'Zoom in', 'Zoom out', 'Reset zoom']
/** The paged toolbar's controls in its compact (coarse pointer) layout. */
const PAGED_COMPACT = ['Previous page', 'Next page', 'Zoom out', 'Reset zoom', 'Zoom in', 'Find', 'More options']

async function open(page: Page, query = '', searchboxes = 2): Promise<void> {
  await page.goto(`/src/debug/ui-host-isolation.html${query}`)
  await page.locator('[data-viewer=sheet] [data-officeview-sheet-instance] canvas').first().waitFor()
  await page.locator('[data-viewer=paged] canvas').first().waitFor()
  await expect(page.getByRole('searchbox')).toHaveCount(searchboxes)
  await expect(page.getByRole('button', { name: 'Zoom in' })).toHaveCount(2)
  await page.evaluate(() => document.fonts.ready)
}

/** Double-click a word of the paged document's first line so Copy is enabled. */
async function showCopyPill(page: Page): Promise<void> {
  const canvas = page.locator('[data-viewer=paged] canvas').first()
  const box = (await canvas.boundingBox())!
  const scale = box.width / 816
  await canvas.dblclick({ position: { x: 130 * scale, y: 104 * scale } })
  // the paged toolbar's Copy button is enabled by the selection; let its
  // 150ms colour transition settle before anything reads its colours
  await expect(page.getByTestId('officeview-copy-selection')).toBeEnabled()
  await page.waitForTimeout(300)
}

type Box = { width: number; height: number; left: number; right: number; top: number; bottom: number }
async function boxes(page: Page, viewer: string, labels: readonly string[] = LABELS): Promise<Record<string, Box>> {
  return page.evaluate(({ labels, v }) => Object.fromEntries(labels.map((name) => {
    const el = document.querySelector(`[data-viewer=${v}] [aria-label="${name}"]`)
    if (!el) throw new Error(`${v}: ${name} missing`)
    return [name, el.getBoundingClientRect().toJSON()]
  })), { labels: [...labels], v: viewer })
}

test.describe('shared viewer controls: geometry (spec §513, AC-A11Y-02)', () => {
  test.describe('coarse pointer', () => {
    test.use({ viewport: { width: 1280, height: 900 }, hasTouch: true, isMobile: true })
    test('coarse pointer: every control is at least 44px high and wide (browser)', async ({ page }) => {
      // coarse: the paged toolbar is compact, its Find collapsed to a button
      await open(page, '', 1)
      test.skip(!(await page.evaluate(() => matchMedia('(pointer: coarse)').matches)), 'this engine does not emulate a coarse pointer')
      const sheet = await boxes(page, 'sheet')
      for (const name of LABELS) {
        expect(sheet[name].height, `sheet ${name}`).toBeGreaterThanOrEqual(44)
        expect(sheet[name].width, `sheet ${name}`).toBeGreaterThanOrEqual(44)
      }
      // the input keeps its 148px width
      expect(sheet['Search document'].width, 'sheet input').toBeCloseTo(148, 0)
      const paged = await boxes(page, 'paged', PAGED_COMPACT)
      for (const name of PAGED_COMPACT) {
        expect(paged[name].height, `paged ${name}`).toBeGreaterThanOrEqual(44)
        expect(paged[name].width, `paged ${name}`).toBeGreaterThanOrEqual(44)
      }
    })
  })

  test('fine pointer: desktop sizes are unchanged (browser)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 900 })
    await open(page)
    expect(await page.evaluate(() => matchMedia('(pointer: fine)').matches)).toBe(true)
    const sheet = await boxes(page, 'sheet')
    const diffs: string[] = []
    for (const name of LABELS) {
      for (const axis of ['width', 'height'] as const) {
        const want = MAIN_SIZES[name][axis], got = sheet[name][axis]
        if (Math.abs(got - want) > 0.5) diffs.push(`sheet ${name} ${axis}: main ${want}, now ${got}`)
      }
    }
    expect(diffs).toEqual([])
    // the paged viewer's Find row is the same SearchBar, inline in its toolbar
    const paged = await boxes(page, 'paged', ['Search document', 'Previous match', 'Next match'])
    for (const name of ['Search document', 'Previous match', 'Next match']) expect(paged[name].height, `paged ${name}`).toBeCloseTo(MAIN_SIZES[name].height, 0)
  })

  // The floating variant (the sheet viewer's; the paged viewer's Find is inline in its toolbar).
  test('is bounded by the viewer width, lets the input shrink and wraps status/buttons (browser)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 900 })
    await open(page)
    await page.evaluate(() => { (document.querySelector('[data-viewer=sheet]') as HTMLElement).style.width = '240px' })
    const search = page.locator('[data-viewer=sheet]').getByRole('searchbox', { name: 'Search document' })
    await search.fill('1')
    await expect(page.locator('[data-viewer=sheet] [data-testid=officeview-search-status]')).toContainText('of')
    const r = await page.evaluate(() => {
      const bar = document.querySelector('[data-viewer=sheet] [data-testid=officeview-search]') as HTMLElement
      // bounded by the box it floats in (the sheet viewer's pane)
      const viewer = (bar.offsetParent as HTMLElement).getBoundingClientRect().toJSON()
      const input = bar.querySelector('input')!.getBoundingClientRect().toJSON()
      const parts = [...bar.children].map((el) => el.getBoundingClientRect().toJSON())
      return { viewer, bar: bar.getBoundingClientRect().toJSON(), barScroll: bar.scrollWidth, barClient: bar.clientWidth, input, parts }
    })
    expect(r.viewer.width).toBeLessThanOrEqual(240)
    // (inside its pane; the sheet pane's own inset differs from the old paged 12px margin)
    expect(r.bar.right).toBeLessThanOrEqual(r.viewer.right + 0.5)
    expect(r.bar.left).toBeGreaterThanOrEqual(r.viewer.left - 0.5)
    expect(r.barScroll).toBeLessThanOrEqual(r.barClient)
    for (const part of r.parts) {
      expect(part.left).toBeGreaterThanOrEqual(r.bar.left - 0.5)
      expect(part.right).toBeLessThanOrEqual(r.bar.right + 0.5)
    }
    expect(r.input.width).toBeLessThan(148)
  })
})

interface Snapshot {
  key: string
  box: { x: number; y: number; width: number; height: number }
  color: string
  backgroundColor: string
  borderColor: string
  borderRadius: string
  padding: string
  icon: { width: number; height: number; stroke: string } | null
}

/** The SearchBar, ZoomControls and copy pill elements of both viewers. */
function snapshot(page: Page): Promise<Snapshot[]> {
  return page.evaluate(() => {
    const out: Snapshot[] = []
    const sel = '[data-testid=officeview-search], [data-testid=officeview-search] *, [data-testid=officeview-zoom-controls], [data-testid=officeview-zoom-controls] button, [data-testid=officeview-copy-selection]'
    for (const v of ['sheet', 'paged']) {
      document.querySelectorAll(`[data-viewer=${v}] :is(${sel})`).forEach((el, i) => {
        if (el.closest('svg') && el.tagName.toLowerCase() !== 'svg') return
        const cs = getComputedStyle(el)
        const b = el.getBoundingClientRect()
        const svg = el.tagName.toLowerCase() === 'svg' ? null : el.querySelector(':scope > svg')
        const ib = svg?.getBoundingClientRect()
        out.push({
          key: `${v}:${i}:${el.tagName.toLowerCase()}:${el.getAttribute('data-testid') ?? ''}:${el.getAttribute('aria-label') ?? el.getAttribute('data-icon') ?? ''}`,
          box: { x: b.x, y: b.y, width: b.width, height: b.height },
          color: cs.color,
          backgroundColor: cs.backgroundColor,
          borderColor: cs.borderColor,
          borderRadius: cs.borderRadius,
          padding: cs.padding,
          icon: svg && ib ? { width: ib.width, height: ib.height, stroke: getComputedStyle(svg).stroke } : null,
        })
      })
    }
    return out
  })
}

test.describe('shared viewer controls: host isolation (spec §6.8)', () => {
  test('hostile host sheet (.app button, all:unset) leaves SearchBar/ZoomControls/copy pill geometry and colours unchanged (browser)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 900 })
    await open(page)
    await showCopyPill(page)
    const clean = await snapshot(page)
    // all three surfaces are covered, icons included
    for (const id of ['officeview-search', 'officeview-zoom-controls', 'officeview-copy-selection']) expect(clean.some((s) => s.key.includes(id)), id).toBe(true)
    expect(clean.filter((s) => s.icon).length).toBeGreaterThanOrEqual(9)
    for (const order of ['first', 'last'] as const) {
      await open(page, `?host=${order}`)
      await showCopyPill(page)
      const hostile = await snapshot(page)
      const byKey = new Map(hostile.map((s) => [s.key, s]))
      const diffs: string[] = []
      for (const want of clean) {
        const got = byKey.get(want.key)
        if (!got) { diffs.push(`${order} ${want.key}: missing`); continue }
        for (const axis of ['x', 'y', 'width', 'height'] as const) {
          if (Math.abs(got.box[axis] - want.box[axis]) > 0.5) diffs.push(`${order} ${want.key} box.${axis}: ${want.box[axis]} -> ${got.box[axis]}`)
        }
        for (const prop of ['color', 'backgroundColor', 'borderColor', 'borderRadius', 'padding'] as const) {
          if (got[prop] !== want[prop]) diffs.push(`${order} ${want.key} ${prop}: ${want[prop]} -> ${got[prop]}`)
        }
        if (JSON.stringify(got.icon) !== JSON.stringify(want.icon)) diffs.push(`${order} ${want.key} icon: ${JSON.stringify(want.icon)} -> ${JSON.stringify(got.icon)}`)
      }
      if (hostile.length !== clean.length) diffs.push(`${order}: ${clean.length} elements -> ${hostile.length}`)
      expect(diffs).toEqual([])
    }
  })
})
