/**
 * Viewer chrome (search bar, zoom controls) around both viewers: Ctrl/Cmd+wheel
 * page-zoom suppression over the whole viewer, spreadsheet print rules, and
 * the search row's fit in narrow viewers.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { act, cleanup, render, waitFor } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { buildDocx, buildXlsx } from '../src/testdata/ooxml-builders'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

afterEach(() => cleanup())

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
