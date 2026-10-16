/**
 * Viewer chrome (search bar, zoom controls) around both viewers: Ctrl/Cmd+wheel
 * page-zoom suppression over the whole viewer, spreadsheet print rules, and
 * the search row's fit in narrow viewers.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { SearchBar, ZoomControls } from '../src/components/viewerControls'
import type { PaintableArray } from '../src/render/paint'
import { workbookOf } from './helpers/xlsx-find'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { buildDocx, buildXlsx } from '../src/testdata/ooxml-builders'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

afterEach(() => { cleanup(); vi.restoreAllMocks() })

const docx = async () => parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'chrome' }] }])))
const xlsx = async () => parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])))

function zoomWheelPrevented(target: Element): boolean[] {
  return ['ctrlKey', 'metaKey'].map(modifier => {
    const event = new WheelEvent('wheel', { deltaY: -100, [modifier]: true, bubbles: true, cancelable: true })
    act(() => { target.dispatchEvent(event) })
    return event.defaultPrevented
  })
}

describe('Ctrl/Cmd+wheel never zooms the browser page over any part of the viewer', () => {
  it('page viewer: over the search bar and the zoom controls', async () => {
    const ui = render(<OfficeDoc document={await docx()} />)
    await waitFor(() => expect(ui.container.querySelector('canvas')).toBeTruthy())
    expect(zoomWheelPrevented(ui.getByRole('searchbox', { name: 'Search document' }))).toEqual([true, true])
    expect(zoomWheelPrevented(ui.getByRole('button', { name: 'Zoom in' }))).toEqual([true, true])
  })

  it('spreadsheet viewer: over the search bar and the zoom controls', FIRST_PAINT_TEST, async () => {
    const ui = render(<OfficeDoc document={await xlsx()} interactiveSheets />)
    await waitFor(() => expect(ui.getAllByRole('gridcell', { name: /^A1,/ })[0]).toBeTruthy(), FIRST_PAINT_WAIT)
    expect(zoomWheelPrevented(ui.getByRole('searchbox', { name: 'Search document' }))).toEqual([true, true])
    expect(zoomWheelPrevented(ui.getByRole('button', { name: 'Zoom in' }))).toEqual([true, true])
  })
})

/** jsdom never applies `@media print`: evaluate the installed print rules
 * (and plain rules) against an element, !important winning, last rule wins. */
function printDisplay(el: Element): string | undefined {
  let best: { important: boolean; value: string } | undefined
  for (const style of Array.from(document.querySelectorAll('style'))) {
    for (const rule of Array.from(style.sheet?.cssRules ?? [])) {
      const media = (rule as CSSMediaRule).media
      const rules = media ? (media.mediaText.includes('print') ? Array.from((rule as CSSMediaRule).cssRules) : []) : [rule]
      for (const r of rules as CSSStyleRule[]) {
        if (!r.style?.display || !el.matches(r.selectorText)) continue
        const important = r.style.getPropertyPriority('display') === 'important'
        if (!best || important || !best.important) best = { important, value: r.style.display }
      }
    }
  }
  return best?.value
}
function printed(el: Element | null): boolean {
  for (let node = el; node; node = node.parentElement) if (printDisplay(node) === 'none') return false
  return true
}

const sheetDoc = () => workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'Name' }, { row: 1, col: 0, value: 'beta' }] }])

describe('spreadsheet print mode', () => {
  it('hides the floating search and zoom controls in print', async () => {
    const print = vi.spyOn(window, 'print').mockImplementation(() => {})
    const ui = render(<InteractiveSheetDoc doc={sheetDoc()} />)
    const search = ui.getByRole('searchbox', { name: 'Search document' }), zoom = ui.getByRole('button', { name: 'Zoom in' })
    // a browser-menu print shows the notice, not the chrome
    expect(printed(ui.getByTestId('officeview-sheet-print-notice'))).toBe(true)
    expect(printed(search)).toBe(false)
    expect(printed(zoom)).toBe(false)
    let during: boolean[] = []
    print.mockImplementation(() => { during = [printed(search), printed(zoom), printed(ui.getByTestId('officeview-sheet-print-pages'))] })
    fireEvent.keyDown(ui.getByRole('region', { name: 'Spreadsheet viewer' }), { key: 'p', ctrlKey: true })
    await waitFor(() => expect(print).toHaveBeenCalled())
    expect(during).toEqual([false, false, true])
  })

  it('print rules are scoped per viewer: printing one never prints the other', async () => {
    const print = vi.spyOn(window, 'print').mockImplementation(() => {})
    const ui = render(<><InteractiveSheetDoc doc={sheetDoc()} /><InteractiveSheetDoc doc={sheetDoc()} /></>)
    const [a, b] = ui.getAllByRole('region', { name: 'Spreadsheet viewer' })
    const notice = (root: Element) => root.querySelector('[data-testid="officeview-sheet-print-notice"]')
    const screen = (root: Element) => root.querySelector('[data-officeview-sheet-screen]')
    // a browser-menu print of the page labels both viewers
    expect([printed(notice(a)), printed(notice(b))]).toEqual([true, true])
    expect([printed(screen(a)), printed(screen(b))]).toEqual([false, false])
    let during: boolean[] = []
    print.mockImplementation(() => { during = [printed(a.querySelector('[data-testid="officeview-sheet-print-pages"]')), printed(screen(a)), printed(notice(b)), printed(screen(b)), printed(b)] })
    fireEvent.keyDown(a, { key: 'p', ctrlKey: true })
    await waitFor(() => expect(print).toHaveBeenCalled())
    // A prints its pages; B contributes neither its notice nor its screen
    expect(during).toEqual([true, false, false, false, false])
    act(() => { window.dispatchEvent(new Event('afterprint')) })
    expect([printed(notice(a)), printed(notice(b))]).toEqual([true, true])
  })

  it('one document-wide print owner: overlapping preparations without afterprint never blank the printout', async () => {
    // Codex review repro: A and B both prepare, A prints and no afterprint
    // arrives, then B prints. The latest starter owns the printout.
    const lease = () => Object.assign([{ spec: { widthPx: 10, heightPx: 10 }, paint: vi.fn() }], { dispose: vi.fn(), fontDiagnostics: [] }) as unknown as PaintableArray
    let resolveA!: () => void, resolveB!: () => void
    const waitA = new Promise<void>(resolve => { resolveA = resolve }), waitB = new Promise<void>(resolve => { resolveB = resolve })
    const pages = (root: Element) => root.querySelector('[data-testid="officeview-sheet-print-pages"]')
    const during: boolean[][] = []
    const print = vi.spyOn(window, 'print').mockImplementation(() => { during.push([a, b].map(root => printed(pages(root)))) })
    const ui = render(<><InteractiveSheetDoc doc={sheetDoc()} printLoader={async () => { await waitA; return lease() }} /><InteractiveSheetDoc doc={sheetDoc()} printLoader={async () => { await waitB; return lease() }} /></>)
    const [a, b] = ui.getAllByRole('region', { name: 'Spreadsheet viewer' })
    fireEvent.keyDown(a, { key: 'p', ctrlKey: true }); fireEvent.keyDown(b, { key: 'p', ctrlKey: true })
    await act(async () => { resolveA(); await new Promise(resolve => setTimeout(resolve, 0)) })
    await act(async () => { resolveB(); await new Promise(resolve => setTimeout(resolve, 0)) })
    expect(print).toHaveBeenCalledTimes(2)
    expect(during).toEqual([[true, false], [false, true]])
    expect(document.querySelectorAll('[data-officeview-sheet-print-owner]')).toHaveLength(1)
    // afterprint (or unmount) releases the owner: browser-menu print labels both again
    act(() => { window.dispatchEvent(new Event('afterprint')) })
    expect(document.querySelectorAll('[data-officeview-sheet-print-owner]')).toHaveLength(0)
    expect([a, b].map(root => printed(root.querySelector('[data-testid="officeview-sheet-print-notice"]')))).toEqual([true, true])
  })

  it('a print-denied viewer hides only itself (with its chrome), never an allowed viewer', () => {
    const ui = render(<><InteractiveSheetDoc doc={sheetDoc()} allowPrint={false} /><InteractiveSheetDoc doc={sheetDoc()} /></>)
    const [denied, allowed] = ui.getAllByRole('region', { name: 'Spreadsheet viewer' })
    const search = (root: Element) => root.closest('[data-officeview-sheet-instance]')!.querySelector('[data-testid="officeview-search"]')
    expect([printed(denied), printed(search(denied))]).toEqual([false, false])
    expect(printed(allowed.querySelector('[data-testid="officeview-sheet-print-notice"]'))).toBe(true)
  })

  it('unmounting the print owner releases the document-wide exclusion', async () => {
    vi.spyOn(window, 'print').mockImplementation(() => {})
    const ui = render(<InteractiveSheetDoc doc={sheetDoc()} />)
    fireEvent.keyDown(ui.getByRole('region', { name: 'Spreadsheet viewer' }), { key: 'p', ctrlKey: true })
    await waitFor(() => expect(document.querySelectorAll('[data-officeview-sheet-print-owner]')).toHaveLength(1))
    ui.unmount()
    expect(document.querySelectorAll('[data-officeview-sheet-print-owner]')).toHaveLength(0)
  })
})

describe('search bar in a narrow viewer', () => {
  // jsdom has no layout: the real-browser check is in tests/browser/interactive-viewer.spec.ts.
  it('is bounded by the viewer width, lets the input shrink and wraps status/buttons', () => {
    const ui = render(<div style={{ position: 'relative', width: 240 }}><SearchBar inputRef={{ current: null }} query="x" onQuery={() => {}} count={3} active={0} indexing={false} onPrev={() => {}} onNext={() => {}} /></div>)
    const bar = ui.getByTestId('officeview-search')
    expect(bar.style.maxWidth).toBe('calc(100% - 24px)')
    expect(bar.style.boxSizing).toBe('border-box')
    expect(bar.style.flexWrap).toBe('wrap')
    expect(ui.getByRole('searchbox').style.minWidth).toBe('0px')
  })
})

describe('touch targets (spec §513, AC-A11Y-02)', () => {
  const chrome = () => render(<div><SearchBar inputRef={{ current: null }} query="x" onQuery={() => {}} count={3} active={0} indexing={false} onPrev={() => {}} onNext={() => {}} /><ZoomControls zoom={1} onIn={() => {}} onOut={() => {}} onReset={() => {}} atFit={false} /></div>)
  const sizes = (ui: ReturnType<typeof render>) => ['Search document', 'Previous match', 'Next match', 'Zoom in', 'Zoom out', 'Reset zoom']
    .map(name => ui.getByLabelText(name) as HTMLElement).map(el => [parseFloat(el.style.height), Math.max(parseFloat(el.style.minWidth) || 0, parseFloat(el.style.width) || 0)])
  it('coarse pointer: every control is at least 44px high and wide (the input is 148px wide)', () => {
    vi.stubGlobal('matchMedia', (query: string) => ({ matches: query.includes('coarse') }))
    try {
      for (const [height, width] of sizes(chrome())) { expect(height).toBeGreaterThanOrEqual(44); expect(width).toBeGreaterThanOrEqual(44) }
    } finally { vi.unstubAllGlobals() }
  })
  it('fine pointer: desktop sizes are unchanged', () => {
    vi.stubGlobal('matchMedia', () => ({ matches: false }))
    try {
      expect(sizes(chrome()).map(([height]) => height)).toEqual([28, 28, 28, 36, 36, 28])
    } finally { vi.unstubAllGlobals() }
  })
})
