import { act, fireEvent, render, waitFor } from '@testing-library/react'
import { afterEach, expect, test, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import type { PptxDocument } from '../src/pptx/types'
const gate = vi.hoisted(() => { let release!: () => void; const promise = new Promise<void>(r => { release = r }); return { promise, release, entered: false } })
vi.mock('../src/core/search', async importOriginal => {
 const original = await importOriginal<typeof import('../src/core/search')>()
 return { ...original, buildTextIndex: async (...args: Parameters<typeof original.buildTextIndex>) => { gate.entered = true; await gate.promise; return original.buildTextIndex(...args) } }
})
import { OfficeDoc } from '../src/components/OfficeDoc'
afterEach(() => vi.unstubAllGlobals())
test('indexing retains a font consumer while async recording outlives main viewer', async () => {
 const add = vi.fn(), remove = vi.fn()
 class Face { async load() { return this } }
 vi.stubGlobal('FontFace', Face); Object.defineProperty(document, 'fonts', { configurable: true, value: { add, delete: remove } })
 const doc: PptxDocument = { slideWidthEmu: 100000, slideHeightEmu: 100000, images: [], slides: [{ index: 0, widthEmu: 100000, heightEmu: 100000, shapes: [] }], embeddedFonts: [{ family: 'Liter', variant: 'regular', relationshipId: 'font', partPath: 'font', bytes: new Uint8Array(readFileSync('tests/fixtures/fonts/Liter-Regular.ttf')) }] }
 const view = render(<OfficeDoc document={doc} />)
 await waitFor(() => expect(view.container.querySelectorAll('canvas')).toHaveLength(1))
 fireEvent.change(view.getByRole('searchbox'), { target: { value: 'x' } })
 await waitFor(() => expect(gate.entered).toBe(true))
 view.unmount(); expect(remove).not.toHaveBeenCalled()
 await act(async () => gate.release())
 await waitFor(() => expect(remove).toHaveBeenCalledTimes(1))
})
