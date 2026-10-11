/**
 * S5.2 review round 1: sheet-mode image/font assets (spec §§10.1, 12.4). Only the
 * active sheet's images decode, each admitted by its header dimensions (S1.1b
 * limits) and charged with its decode transient against a bounded aggregate;
 * a font change swaps leases atomically.
 */
import { describe, expect, it, vi } from 'vitest'
import { act, render, waitFor } from '@testing-library/react'
import { prepareSheetAssets } from '../src/components/xlsx/sheetAssets'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'

/** PNG signature + IHDR declaring width x height (enough for header admission). */
function pngHeader(width: number, height: number, salt: number): Uint8Array {
  const bytes = new Uint8Array(64)
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52])
  const view = new DataView(bytes.buffer)
  view.setUint32(16, width)
  view.setUint32(20, height)
  bytes.set([8, 6, 0, 0, 0], 24)
  bytes[40] = salt // distinct bytes: no decode sharing between assets
  return bytes
}
const source = { partPath: 'd', treePath: 't', element: 'pic', representation: 'native' as const }
const picture = (imageIndex: number) => ({ xEmu: 0, yEmu: 0, widthEmu: 952500, heightEmu: 952500, geometry: 'rect', imageIndex, source })
function workbookWithImages(): XlsxDocument {
  const sheet = (name: string, indexes: number[]) => ({ name, rows: [], cols: [], merges: [], mergeRanges: [], drawings: indexes.map(picture) }) as unknown as XlsxSheet
  return {
    sheets: [sheet('A', [0, 1, 2, 3, 4]), sheet('B', [5, 6, 7, 8, 9])],
    images: Array.from({ length: 10 }, (_, i) => ({ data: pngHeader(4096, 4096, i), mime: 'image/png' })),
  } as unknown as XlsxDocument
}

describe('sheet-mode image admission', () => {
  it('decodes only the requested sheet, admitting by header dimensions within a bounded aggregate', async () => {
    const decode = vi.fn(async () => ({ width: 4096, height: 4096 }) as unknown as CanvasImageSource)
    const lease = await prepareSheetAssets(workbookWithImages(), { decodeImage: decode, sheetIndex: 0, profile: 'D' })
    try {
      // 4096² RGBA = 64 MiB each; with its decode transient one fits the D budget.
      expect(decode).toHaveBeenCalledTimes(1)
      expect(lease.decodedBytes).toBeLessThanOrEqual(lease.imageBudgetBytes)
      expect(lease.refusedImages).toBe(4)
      expect(lease.assets.images?.slice(5).every(image => image === undefined)).toBe(true)
    } finally { lease.dispose() }
  })

  it('refuses an image whose header exceeds the per-image pixel limit before decoding', async () => {
    const decode = vi.fn(async () => ({}) as CanvasImageSource)
    const doc = workbookWithImages()
    doc.images![0] = { data: pngHeader(20_000, 20_000, 0), mime: 'image/png' }
    const lease = await prepareSheetAssets(doc, { decodeImage: decode, sheetIndex: 0, profile: 'M' })
    try {
      expect(decode).not.toHaveBeenCalledWith(doc.images![0].data, expect.anything())
      expect(lease.assets.images?.[0]).toBeUndefined()
    } finally { lease.dispose() }
  })
})

describe('font lease swap', () => {
  it('keeps the current lease until its replacement has published', async () => {
    const disposals: string[] = []
    let release: (() => void) | undefined
    const calls: string[] = []
    const loader: typeof prepareSheetAssets = async (doc, options) => {
      const tag = options?.fallbackFonts ? 'next' : 'first'
      calls.push(tag)
      if (tag === 'next') await new Promise<void>(resolve => { release = resolve })
      const lease = await prepareSheetAssets(doc, options)
      return { ...lease, dispose: () => { disposals.push(tag); lease.dispose() } }
    }
    const { default: InteractiveSheetDoc } = await import('../src/components/xlsx/InteractiveSheetDoc')
    const doc = workbookWithImages()
    const ui = render(<InteractiveSheetDoc doc={doc} assetLoader={loader} />)
    await waitFor(() => expect(calls).toEqual(['first']))
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) })
    ui.rerender(<InteractiveSheetDoc doc={doc} assetLoader={loader} fallbackFonts={{ fallbackChain: ['Arial'] }} />)
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) })
    // The replacement is still pending: the current lease must still be alive.
    expect(disposals).toEqual([])
    await act(async () => { release?.(); await new Promise(resolve => setTimeout(resolve, 5)) })
    expect(disposals).toEqual(['first'])
    ui.unmount()
    expect(disposals).toEqual(['first', 'next'])
  })
})
