import { readFileSync } from 'node:fs'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { afterEach, expect, test, vi } from 'vitest'
import { OfficeDoc } from '../src/components/OfficeDoc'
import type { PptxDocument } from '../src/pptx/types'
const bytes = new Uint8Array(readFileSync('tests/fixtures/fonts/Liter-Regular.ttf'))
function source(family: string): PptxDocument { return { slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], slides: [{ index: 0, widthEmu: 9144000, heightEmu: 6858000, shapes: [{ xEmu: 0, yEmu: 0, widthEmu: 9144000, heightEmu: 6858000, geometry: 'rect', textBody: { anchor: 't', wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, paragraphs: [{ align: 'left', level: 0, runs: [{ text: family, fontFamily: family }] }] } }] }], embeddedFonts: [{ family, variant: 'regular', bytes, partPath: 'font', relationshipId: 'font' }] } }
function harness() {
 const pending: Array<() => void> = [], add = vi.fn(), remove = vi.fn()
 class MockFace { family: string; constructor(family: string) { this.family = family } load() { return new Promise<MockFace>(resolve => pending.push(() => resolve(this))) } }
 vi.stubGlobal('FontFace', MockFace); Object.defineProperty(document, 'fonts', { configurable: true, value: { add, delete: remove } })
 return { pending, add, remove }
}
afterEach(() => { cleanup(); vi.unstubAllGlobals() })
test('replacement during pending font load and unmount release late results', async () => {
 const h = harness(), a = source('A'), b = source('B')
 const view = render(<OfficeDoc document={a} />)
 await waitFor(() => expect(h.pending).toHaveLength(1))
 view.rerender(<OfficeDoc document={b} />)
 await waitFor(() => expect(h.pending).toHaveLength(2))
 await act(async () => h.pending[0]())
 await waitFor(() => expect(h.remove).toHaveBeenCalledTimes(1))
 view.unmount(); await act(async () => h.pending[1]())
 await waitFor(() => expect(h.remove).toHaveBeenCalledTimes(2))
})
test('concurrent viewers retain loaded fonts until their final unmount', async () => {
 const h = harness(), a = source('A')
 const first = render(<OfficeDoc document={a} />), second = render(<OfficeDoc document={a} />)
 await waitFor(() => expect(h.pending).toHaveLength(1)); await act(async () => h.pending[0]())
 await waitFor(() => expect(h.add).toHaveBeenCalledTimes(1))
 first.unmount(); expect(h.remove).not.toHaveBeenCalled()
 second.unmount(); expect(h.remove).toHaveBeenCalledTimes(1)
})
test('lazy indexing fallback owns and releases its own pending extraction on replacement', async () => {
 const h = harness(), a = source('A'), b = source('B')
 const view = render(<OfficeDoc document={a} />)
 await waitFor(() => expect(h.pending).toHaveLength(1))
 fireEvent.change(view.getByRole('searchbox'), { target: { value: 'B' } })
 await act(async () => { await new Promise(r => setTimeout(r, 0)) })
 view.rerender(<OfficeDoc document={b} />)
 await waitFor(() => expect(h.pending).toHaveLength(2))
 await act(async () => h.pending[0]())
 await waitFor(() => expect(h.remove).toHaveBeenCalledTimes(1))
 await act(async () => h.pending[1]())
 await waitFor(() => expect(view.getByTestId('officeview-search-status').textContent).toBe('1 of 1'))
 view.unmount(); await waitFor(() => expect(h.remove).toHaveBeenCalledTimes(2))
})

test.each(['document', 'watermark'] as const)('%s owner change clears stale indexing after the query was emptied', async change => {
 const h = harness(), a = source('A'), b = source('B')
 const view = render(<OfficeDoc document={a} />)
 await waitFor(() => expect(h.pending).toHaveLength(1))
 fireEvent.change(view.getByRole('searchbox'), { target: { value: 'x' } })
 await waitFor(() => expect(view.getByTestId('officeview-search-status').textContent).toBe('indexing…'))
 fireEvent.change(view.getByRole('searchbox'), { target: { value: '' } })
 if (change === 'document') view.rerender(<OfficeDoc document={b} />)
 else view.rerender(<OfficeDoc document={a} watermark={{ text: 'New owner' }} />)
 await waitFor(() => expect(h.pending).toHaveLength(change === 'document' ? 2 : 1))
 await act(async () => { h.pending.forEach(resolve => resolve()) })
 await waitFor(() => expect(view.container.querySelectorAll('canvas')).toHaveLength(1))
 await waitFor(() => expect(view.getByTestId('officeview-search-status').textContent).toBe(''))
 // The watermark replacement shares A's pending registration; old indexing
 // releases only its own reference, leaving the replacement viewer alive.
 expect(h.remove).toHaveBeenCalledTimes(change === 'document' ? 1 : 0)
 view.unmount()
 await waitFor(() => expect(h.remove).toHaveBeenCalledTimes(change === 'document' ? 2 : 1))
})
