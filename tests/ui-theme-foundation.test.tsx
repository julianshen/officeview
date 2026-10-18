/**
 * Task 1 (spec §6.2, §6.8): the theme root, the `theme` prop, the token blocks
 * and the missing-stylesheet warning. jsdom asserts attributes and token
 * values; computed chrome geometry and colours live in the Playwright spec
 * tests/browser/ui-host-isolation.spec.ts.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import postcss from 'postcss'
import css from '../src/styles/officeview.css?inline'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficeFile } from '../src/components/OfficeFile'
import { resetStyleWarningForTests, themeRootProps } from '../src/components/theme'
import { COLOR_TOKENS, CANVAS_TOKENS, THEME_DEFAULTS } from '../src/styles/tokens'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { buildDocx, buildXlsx } from '../src/testdata/ooxml-builders'
import { FIRST_PAINT_TEST } from './helpers/viewer-timing'
import { STYLE_ELEMENT_ID } from './setup/styles'

afterEach(() => { cleanup(); vi.restoreAllMocks(); resetStyleWarningForTests() })

const docxBytes = () => buildDocx([{ runs: [{ text: 'theme root' }] }])
const docx = async () => parseDocx(await OfficePackage.load(await docxBytes()))
const xlsx = async () => parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])))

const themeRootOf = (c: HTMLElement) => c.querySelector('[data-ov-theme-root]') as HTMLElement | null

describe('theme root (spec §6.8)', () => {
  it('themeRootProps defaults to system and carries the instance id', () => {
    expect(themeRootProps(undefined, 'i1')).toEqual({ 'data-ov-theme-root': '', 'data-ov-theme': 'system', 'data-ov-instance': 'i1' })
    expect(themeRootProps('dark', 'i2')['data-ov-theme']).toBe('dark')
  })

  it('marks the outer wrapper as theme root, data-ov-theme=system by default, honouring the prop', FIRST_PAINT_TEST, async () => {
    // Paged (DOCX) viewer: the LegacyOfficeDoc wrapper.
    const doc = await docx()
    const ui = render(<OfficeDoc document={doc} />)
    const root = themeRootOf(ui.container)!
    expect(root).toBe(ui.container.firstElementChild)
    expect(root.getAttribute('data-ov-theme')).toBe('system')
    expect(root.getAttribute('data-ov-instance')).toBeTruthy()
    expect(root.querySelector('[data-officeview-root]')).toBeTruthy()
    // The floating controls come from the lazy chunk and render inside it.
    const search = await ui.findByRole('searchbox', { name: 'Search document' })
    expect(root.contains(search)).toBe(true)
    ui.rerender(<OfficeDoc document={doc} theme="dark" />)
    expect(themeRootOf(ui.container)!.getAttribute('data-ov-theme')).toBe('dark')
    ui.unmount()

    // Interactive sheet viewer: its outer wrapper.
    const book = await xlsx()
    const sheet = render(<OfficeDoc document={book} interactiveSheets />)
    await waitFor(() => expect(sheet.container.querySelector('[data-officeview-sheet-instance]')).toBeTruthy())
    const sheetRoot = sheet.container.querySelector('[data-officeview-sheet-instance]') as HTMLElement
    expect(sheetRoot.hasAttribute('data-ov-theme-root')).toBe(true)
    expect(sheetRoot.getAttribute('data-ov-theme')).toBe('system')
    expect(sheetRoot.getAttribute('data-ov-instance')).toBe(sheetRoot.getAttribute('data-officeview-sheet-instance'))
    expect(sheetRoot.querySelector('[data-officeview-root]')).toBeTruthy()
    expect(sheetRoot.contains(sheet.getByRole('searchbox'))).toBe(true)
    sheet.rerender(<OfficeDoc document={book} interactiveSheets theme="light" />)
    expect(sheetRoot.getAttribute('data-ov-theme')).toBe('light')
    sheet.unmount()

    // OfficeFile passes the prop through.
    const file = render(<OfficeFile data={await docxBytes()} theme="dark" />)
    await waitFor(() => expect(themeRootOf(file.container)).toBeTruthy())
    expect(themeRootOf(file.container)!.getAttribute('data-ov-theme')).toBe('dark')
  })

  it('the lazy-sheet Suspense fallback is a theme root too', async () => {
    // A fresh module graph, so OfficeDoc's lazy() has not resolved the sheet chunk.
    vi.resetModules()
    const fresh = await import('../src/components/OfficeDoc')
    const ui = render(<fresh.OfficeDoc document={await xlsx()} interactiveSheets theme="dark" />)
    const status = ui.getByRole('status')
    expect(status.textContent).toBe('Loading spreadsheet viewer…')
    const root = themeRootOf(ui.container)!
    expect(root.getAttribute('data-ov-theme')).toBe('dark')
    expect(root.getAttribute('data-ov-instance')).toBeTruthy()
    expect(root.contains(status)).toBe(true)
    expect(root.querySelector('[data-officeview-root]')).toBe(status)
    await waitFor(() => expect(ui.container.querySelector('[data-officeview-sheet-instance]')).toBeTruthy())
  })

  it('data-officeview-root keeps host className/style and carries no ov: classes', FIRST_PAINT_TEST, async () => {
    const ui = render(<OfficeDoc document={await docx()} className="host-cls" style={{ outline: '3px solid red' }} />)
    const inner = ui.container.querySelector('[data-officeview-root]') as HTMLElement
    expect(inner.className).toBe('host-cls')
    expect(inner.style.outline).toBe('3px solid red')
    expect(inner.hasAttribute('data-ov-theme-root')).toBe(false)
    ui.unmount()

    const sheet = render(<OfficeDoc document={await xlsx()} interactiveSheets className="host-cls" style={{ outline: '3px solid red' }} />)
    await waitFor(() => expect(sheet.container.querySelector('[data-officeview-sheet-instance] [data-officeview-root]')).toBeTruthy())
    const sheetInner = sheet.container.querySelector('[data-officeview-sheet-instance] [data-officeview-root]') as HTMLElement
    expect(sheetInner.className).toBe('host-cls')
    expect(sheetInner.style.outline).toBe('3px solid red')
    for (const el of sheet.container.querySelectorAll('[data-officeview-root]')) {
      expect([...el.classList].filter((c) => c.startsWith('ov:'))).toEqual([])
    }
  })
})

describe('theme tokens (spec §6.8)', () => {
  function tokenBlock(selector: string, media?: string): Record<string, string> {
    const out: Record<string, string> = {}
    postcss.parse(css).walkRules((rule) => {
      if (rule.selector !== selector) return
      const parent = rule.parent as postcss.AtRule | postcss.Root
      const inMedia = parent.type === 'atrule' ? (parent as postcss.AtRule).params : undefined
      if (inMedia !== media) return
      rule.walkDecls((d) => { out[d.prop] = d.value })
    })
    return out
  }

  it('defines light tokens and the sentinel on the theme root', () => {
    const el = document.createElement('div')
    Object.entries(themeRootProps(undefined, 't')).forEach(([k, v]) => el.setAttribute(k, v))
    document.body.appendChild(el)
    const cs = getComputedStyle(el)
    expect(cs.getPropertyValue('--ov-color-accent').trim()).toBe('#0f6cbd')
    expect(cs.getPropertyValue('--ov-sentinel').trim()).toBe('1')
    el.remove()
  })

  it('dark tokens apply for data-ov-theme=dark', () => {
    const el = document.createElement('div')
    Object.entries(themeRootProps('dark', 't')).forEach(([k, v]) => el.setAttribute(k, v))
    document.body.appendChild(el)
    expect(getComputedStyle(el).getPropertyValue('--ov-color-surface').trim()).toBe('#292827')
    el.remove()
  })

  it('THEME_DEFAULTS mirror the stylesheet token blocks', () => {
    const light = tokenBlock(':where([data-ov-theme-root])')
    const dark = tokenBlock(':where([data-ov-theme-root][data-ov-theme=dark])')
    const system = tokenBlock(':where([data-ov-theme-root][data-ov-theme=system])', '(prefers-color-scheme: dark)')
    expect(light['--ov-sentinel']).toBe('1')
    for (const name of CANVAS_TOKENS) expect(COLOR_TOKENS).toContain(name)
    for (const name of COLOR_TOKENS) {
      expect(light[`--ov-color-${name}`], `light ${name}`).toBe(THEME_DEFAULTS.light[name])
      expect(dark[`--ov-color-${name}`], `dark ${name}`).toBe(THEME_DEFAULTS.dark[name])
      expect(system[`--ov-color-${name}`], `system ${name}`).toBe(THEME_DEFAULTS.dark[name])
    }
    // Radii and shadows (plan Global Constraints, Default values).
    expect(light['--ov-radius-control']).toBe('4px')
    expect(light['--ov-radius-panel']).toBe('6px')
    expect(light['--ov-shadow-rest']).toBe('0 1.6px 3.6px rgba(0,0,0,.13)')
    expect(light['--ov-shadow-popover']).toBe('0 6.4px 14.4px rgba(0,0,0,.18)')
    expect(dark['--ov-shadow-rest']).toBe('0 1.6px 3.6px rgba(0,0,0,.5)')
    expect(system['--ov-shadow-popover']).toBe('0 6.4px 14.4px rgba(0,0,0,.6)')
  })
})

describe('missing stylesheet warning (spec §5, R10)', () => {
  const MESSAGE = '[officeview] style.css not loaded: import "officeview/style.css"'

  it('warns once per page when the sentinel is missing, regardless of build mode', FIRST_PAINT_TEST, async () => {
    const style = document.getElementById(STYLE_ELEMENT_ID)!
    expect(style).toBeTruthy()
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const doc = await docx()
    style.remove()
    try {
      resetStyleWarningForTests()
      const a = render(<OfficeDoc document={doc} />)
      const b = render(<OfficeDoc document={doc} />)
      await waitFor(() => expect(warn).toHaveBeenCalled())
      await a.findAllByRole('searchbox')
      await b.findAllByRole('searchbox')
      expect(warn.mock.calls.filter((c) => String(c[0]).includes('style.css'))).toEqual([[MESSAGE]])
      cleanup()
    } finally {
      document.head.appendChild(style)
    }
    resetStyleWarningForTests()
    warn.mockClear()
    const c = render(<OfficeDoc document={doc} />)
    await c.findByRole('searchbox')
    const sheet = render(<OfficeDoc document={await xlsx()} interactiveSheets />)
    await waitFor(() => expect(sheet.container.querySelector('[data-officeview-sheet-instance]')).toBeTruthy())
    expect(warn.mock.calls.filter((call) => String(call[0]).includes('style.css'))).toEqual([])
  })

  it('the interactive sheet viewer warns too when the sentinel is missing', FIRST_PAINT_TEST, async () => {
    const style = document.getElementById(STYLE_ELEMENT_ID)!
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const book = await xlsx()
    style.remove()
    try {
      resetStyleWarningForTests()
      const ui = render(<OfficeDoc document={book} interactiveSheets />)
      await waitFor(() => expect(ui.container.querySelector('[data-officeview-sheet-instance]')).toBeTruthy())
      await waitFor(() => expect(warn.mock.calls.filter((c) => String(c[0]).includes('style.css'))).toEqual([[MESSAGE]]))
    } finally {
      document.head.appendChild(style)
    }
  })
})
