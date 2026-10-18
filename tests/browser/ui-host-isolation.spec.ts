/**
 * Task 1 (spec §6.8, Review Focus 1 and 2): the viewer's chrome is isolated
 * from hostile host CSS, the host is isolated from the viewer's CSS, and every
 * shipped or served officeview selector is scoped to [data-ov-theme-root].
 *
 * Page: src/debug/ui-host-isolation.html (sheet + paged viewers inside a host
 * `<div class="app">`; `host=first|last` adds the hostile sheet before or
 * after officeview.css; `mode=pair` mounts two differently themed viewers).
 */
import { expect, test, type Page } from '@playwright/test'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import postcss from 'postcss'
import { unscopedSelectors } from '../../scripts/postcss-ov-scope'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))

interface Control {
  key: string
  box: { x: number; y: number; width: number; height: number }
  color: string
  backgroundColor: string
  borderColor: string
  padding: string
  boxSizing: string
}

async function open(page: Page, query: string): Promise<void> {
  await page.setViewportSize({ width: 1280, height: 900 })
  await page.goto(`/src/debug/ui-host-isolation.html${query}`)
}

async function ready(page: Page): Promise<void> {
  // Sheet viewer painted, paged viewer painted, and both Find bars mounted
  // (the paged one comes from the lazy chunk).
  await page.locator('[data-viewer=sheet] [data-officeview-sheet-instance] canvas').first().waitFor()
  await page.locator('[data-viewer=paged] canvas').first().waitFor()
  await expect(page.getByRole('searchbox')).toHaveCount(2)
  await expect(page.getByRole('button', { name: 'Zoom in' })).toHaveCount(2)
  await page.evaluate(() => document.fonts.ready)
}

/** Every toolbar and floating control inside a theme root, in DOM order. */
function controls(page: Page): Promise<Control[]> {
  return page.evaluate(() => {
    const out: Control[] = []
    const roots = [...document.querySelectorAll('[data-ov-theme-root]')]
    roots.forEach((root, r) => {
      const els = root.querySelectorAll('button, input, select, textarea, [data-officeview-chrome], [data-testid=officeview-zoom-controls]')
      els.forEach((el, i) => {
        const cs = getComputedStyle(el)
        const b = el.getBoundingClientRect()
        const label = el.getAttribute('aria-label') ?? el.getAttribute('data-testid') ?? el.tagName
        out.push({
          key: `${r}:${i}:${el.tagName.toLowerCase()}:${label}`,
          box: { x: b.x, y: b.y, width: b.width, height: b.height },
          color: cs.color,
          backgroundColor: cs.backgroundColor,
          borderColor: cs.borderColor,
          padding: cs.padding,
          boxSizing: cs.boxSizing,
        })
      })
    })
    return out
  })
}

function expectSameControls(clean: Control[], hostile: Control[]): void {
  expect(clean.length).toBeGreaterThan(6)
  const byKey = new Map(hostile.map((c) => [c.key, c]))
  const diffs: string[] = []
  for (const want of clean) {
    const got = byKey.get(want.key)
    if (!got) { diffs.push(`${want.key}: missing`); continue }
    for (const axis of ['x', 'y', 'width', 'height'] as const) {
      if (Math.abs(got.box[axis] - want.box[axis]) > 0.5) diffs.push(`${want.key} box.${axis}: ${want.box[axis]} -> ${got.box[axis]}`)
    }
    for (const prop of ['color', 'backgroundColor', 'borderColor', 'padding', 'boxSizing'] as const) {
      if (got[prop] !== want[prop]) diffs.push(`${want.key} ${prop}: ${want[prop]} -> ${got[prop]}`)
    }
  }
  for (const got of hostile) if (!clean.some((c) => c.key === got.key)) diffs.push(`${got.key}: extra`)
  expect(diffs).toEqual([])
}

async function hostileOrder(page: Page): Promise<'first' | 'last'> {
  return page.evaluate(() => {
    const styles = [...document.head.querySelectorAll('style')]
    const hostile = styles.findIndex((s) => s.id === 'hostile')
    const ov = styles.findIndex((s) => (s.getAttribute('data-vite-dev-id') ?? '').endsWith('/src/styles/officeview.css'))
    if (hostile < 0 || ov < 0) throw new Error(`style elements missing (hostile ${hostile}, officeview ${ov})`)
    return hostile < ov ? 'first' : 'last'
  })
}

test.describe('host isolation (spec §6.8)', () => {
  for (const order of ['first', 'last'] as const) {
    test(`toolbar geometry and colours are identical with and without a hostile host reset (host sheet ${order})`, async ({ page }) => {
      await open(page, '')
      await ready(page)
      const clean = await controls(page)
      await open(page, `?host=${order}`)
      await ready(page)
      expect(await hostileOrder(page)).toBe(order)
      expectSameControls(clean, await controls(page))
    })
  }

  test('host elements are untouched', async ({ page }) => {
    await open(page, '?host=last')
    await ready(page)
    const host = await page.evaluate(() => ({
      button: getComputedStyle(document.getElementById('host-button')!).color,
      flex: getComputedStyle(document.getElementById('host-flex')!).display,
    }))
    expect(host).toEqual({ button: 'rgb(1, 2, 3)', flex: 'block' })
    // And on a clean page the viewer's stylesheet leaves host elements alone.
    await open(page, '')
    await ready(page)
    const clean = await page.evaluate(() => {
      const button = document.getElementById('host-button')!
      const cs = getComputedStyle(button)
      return { display: getComputedStyle(document.getElementById('host-flex')!).display, appearance: cs.appearance, boxSizing: cs.boxSizing }
    })
    expect(clean.display).toBe('block')
    expect(clean.appearance).not.toBe('none')
    expect(clean.boxSizing).toBe('border-box') // the UA default for buttons, not the viewer's base
  })

  test('dist/style.css: every selector is scoped; no @property, :root or :host', async () => {
    const file = join(ROOT, 'dist', 'style.css')
    expect(existsSync(file), 'dist/style.css missing: run bun run build').toBe(true)
    const css = readFileSync(file, 'utf8')
    const root = postcss.parse(css)
    let rules = 0
    root.walkRules(() => { rules++ })
    expect(rules).toBeGreaterThan(0)
    expect(unscopedSelectors(css)).toEqual([])
    const banned: string[] = []
    root.walkAtRules('property', (r) => { banned.push(`@property ${r.params}`) })
    root.walkRules((r) => { for (const s of r.selectors) if (/:root\b|:host\b/.test(s)) banned.push(s) })
    expect(banned).toEqual([])
  })

  test('the harness-served stylesheet is scoped (dev pipeline)', async ({ page }) => {
    await open(page, '')
    await ready(page)
    const css = await page.evaluate(() => {
      const sheets = [...document.styleSheets].filter((s) => {
        const owner = s.ownerNode as Element | null
        return (owner?.getAttribute('data-vite-dev-id') ?? s.href ?? '').replace(/[?#].*$/, '').endsWith('/src/styles/officeview.css')
      })
      return { count: sheets.length, text: sheets.map((s) => [...s.cssRules].map((r) => r.cssText).join('\n')).join('\n') }
    })
    expect(css.count).toBe(1)
    expect(css.text).toContain('--ov-sentinel')
    expect(css.text).toMatch(/\.ov\\:flex:is\(\[data-ov-theme-root\]/)
    expect(css.text).not.toContain('@property')
    expect(unscopedSelectors(css.text)).toEqual([])
  })

  test('two adjacent viewers resolve their own theme and host overrides, including floating controls', async ({ page }) => {
    await open(page, '?mode=pair')
    await page.locator('[data-viewer="1"] canvas').first().waitFor()
    await page.locator('[data-viewer="2"] canvas').first().waitFor()
    await expect(page.getByRole('searchbox')).toHaveCount(2)
    const resolved = await page.evaluate(() => [1, 2].map((n) => {
      const root = document.querySelector(`[data-viewer="${n}"] [data-ov-theme-root]`)!
      const search = root.querySelector('[data-testid=officeview-search]')!
      const input = root.querySelector('input[type=search]')!
      const zoom = root.querySelector('[data-testid=officeview-zoom-controls]')!
      const v = (el: Element, p: string) => getComputedStyle(el).getPropertyValue(p).trim()
      return {
        theme: root.getAttribute('data-ov-theme'),
        rootColor: getComputedStyle(root).color,
        accent: v(root, '--ov-color-accent'),
        searchAccent: v(search, '--ov-color-accent'),
        zoomSurface: v(zoom, '--ov-color-surface'),
        inputColor: getComputedStyle(input).color,
      }
    }))
    expect(resolved).toEqual([
      { theme: 'light', rootColor: 'rgb(32, 31, 30)', accent: '#c50f1f', searchAccent: '#c50f1f', zoomSurface: '#ffffff', inputColor: 'rgb(32, 31, 30)' },
      { theme: 'dark', rootColor: 'rgb(243, 242, 241)', accent: '#479ef5', searchAccent: '#479ef5', zoomSurface: '#292827', inputColor: 'rgb(243, 242, 241)' },
    ])
  })
})
