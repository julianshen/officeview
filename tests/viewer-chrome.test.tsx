/**
 * Viewer chrome (search bar, zoom controls) around both viewers: Ctrl/Cmd+wheel
 * page-zoom suppression over the whole viewer, spreadsheet print rules, and
 * the search row's fit in narrow viewers.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
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
})
