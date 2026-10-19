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
    const calls: string[] = [], loaded: string[] = []
    const loader: typeof prepareSheetAssets = async (doc, options) => {
      const tag = options?.fallbackFonts ? 'next' : 'first'
      calls.push(tag)
      if (tag === 'next') await new Promise<void>(resolve => { release = resolve })
      const lease = await prepareSheetAssets(doc, options)
      loaded.push(tag)
      return { ...lease, dispose: () => { disposals.push(tag); lease.dispose() } }
    }
    const { default: InteractiveSheetDoc } = await import('../src/components/xlsx/InteractiveSheetDoc')
    const doc = workbookWithImages()
    const ui = render(<InteractiveSheetDoc doc={doc} assetLoader={loader} />)
    // The first lease must be current before the swap starts: wait for its
    // load to settle (a fixed sleep raced a slow prepareSheetAssets, and a
    // lease that lands after it is superseded is rightly disposed at once).
    await waitFor(() => expect(loaded).toEqual(['first']))
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

describe('round 2: SVG bounds, shared image budget, stale sheet pixels', () => {
  const svg = (w: number, h: number) => new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><rect width="10" height="10"/></svg>`)
  const svgDoc = (bytes: Uint8Array) => ({
    sheets: [{ name: 'A', rows: [], cols: [], merges: [], mergeRanges: [], drawings: [picture(0)] }],
    images: [{ data: bytes, mime: 'image/svg+xml' }],
  }) as unknown as XlsxDocument

  it('refuses an SVG whose rasterized size exceeds the per-image limits; admits a small one', async () => {
    const decode = vi.fn(async () => ({}) as CanvasImageSource)
    const big = await prepareSheetAssets(svgDoc(svg(20_000, 20_000)), { decodeImage: decode, profile: 'D' })
    expect(decode).not.toHaveBeenCalled()
    expect(big.refusedImages).toBe(1)
    big.dispose()
    const small = await prepareSheetAssets(svgDoc(svg(100, 100)), { decodeImage: decode, profile: 'D' })
    expect(decode).toHaveBeenCalledTimes(1)
    expect(small.decodedBytes).toBeGreaterThan(0)
    small.dispose()
  })

  it('charges overlapping leases against one shared image budget and shares identical pictures', async () => {
    const { SheetImageCache } = await import('../src/components/xlsx/sheetAssets')
    const decode = vi.fn(async () => ({ close() {} }) as unknown as CanvasImageSource)
    const doc = workbookWithImages()
    const cache = new SheetImageCache(128 * 1024 * 1024)
    const a = await prepareSheetAssets(doc, { decodeImage: decode, sheetIndex: 0, profile: 'D', imageCache: cache })
    expect(decode).toHaveBeenCalledTimes(1)
    // Same sheet (a font change): the retained picture is shared, not re-decoded.
    const again = await prepareSheetAssets(doc, { decodeImage: decode, sheetIndex: 0, profile: 'D', imageCache: cache })
    expect(decode).toHaveBeenCalledTimes(1)
    // Another sheet while A is retained: its decode must fit what A leaves.
    const b = await prepareSheetAssets(doc, { decodeImage: decode, sheetIndex: 1, profile: 'D', imageCache: cache })
    expect(cache.peakBytes).toBeLessThanOrEqual(cache.limitBytes)
    a.dispose(); again.dispose(); b.dispose()
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(cache.chargedBytes).toBe(0)
  })

  it('a pending sheet switch never keeps presenting the previous sheet\'s pixels', async () => {
    const { OfficePackage } = await import('../src/core/zip')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const doc = await parseXlsx(await OfficePackage.load(await buildXlsx([
      { name: 'A', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'AAAAAAAA' }] }] },
      { name: 'B', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'BBBBBBBB' }] }] },
    ])))
    let release: (() => void) | undefined
    const loader: typeof prepareSheetAssets = async (d, options) => {
      if (options?.sheetIndex === 1) await new Promise<void>(resolve => { release = resolve })
      return prepareSheetAssets(d, options)
    }
    const { default: InteractiveSheetDoc } = await import('../src/components/xlsx/InteractiveSheetDoc')
    const { fireEvent, waitFor: wait } = await import('@testing-library/react')
    const ui = render(<InteractiveSheetDoc doc={doc} assetLoader={loader} />)
    const base = () => ui.getByTestId('officeview-sheet-base') as HTMLCanvasElement
    const pixels = () => [...base().getContext('2d')!.getImageData(0, 0, 120, 24).data].join()
    await wait(() => expect(pixels()).not.toBe(new Array(120 * 24 * 4).fill(0).join()))
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) })
    const sheetA = pixels()
    act(() => { fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: '2' } }) })
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) })
    expect(pixels()).not.toBe(sheetA)
    expect(ui.getByRole('status').textContent).toContain('Loading sheet')
    await act(async () => { release?.(); await new Promise(resolve => setTimeout(resolve, 5)) })
    expect(ui.getByRole('status').textContent).not.toContain('Loading sheet')
    ui.unmount()
  })
})

describe('round 3: SVG raster bounds are proven from the real XML root', () => {
  const enc = (text: string) => new TextEncoder().encode(text)
  const NS = 'xmlns="http://www.w3.org/2000/svg"'
  it('skips comments, PIs and doctype; withdraws CSS- or context-dependent sizes', async () => {
    const { svgIntrinsicSize } = await import('../src/components/xlsx/sheetAssets')
    // A commented-out root must not decide the size.
    expect(svgIntrinsicSize(enc(`<!-- <svg width="1" height="1"> --><svg ${NS} viewBox="0 0 640 360"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<!-- <svg width="1" height="1"> --><svg ${NS} width="20000" height="20000"/>`))).toEqual({ width: 20_000, height: 20_000 })
    // CSS can resize the root: style attribute, <style> element, xml-stylesheet PI.
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="1" height="1" style="width:640px;height:360px"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="1" height="1"><style>svg{width:640px;height:360px}</style></svg>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<?xml-stylesheet href="s.css"?><svg ${NS} width="1" height="1"/>`))).toBeNull()
    // Relative or missing dimensions depend on the decoder's context.
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="100%" height="50%" viewBox="0 0 10 10"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="10em" height="10em"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS}/>`))).toBeNull()
    // Declarations and a doctype (with an internal subset) are skipped.
    expect(svgIntrinsicSize(enc(`<?xml version="1.0"?><!DOCTYPE svg [ <!ENTITY e "<svg width='1' height='1'>"> ]><!-- c --><svg ${NS} width="100" height="80"/>`))).toEqual({ width: 100, height: 80 })
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="1in" height="72pt"/>`))).toEqual({ width: 96, height: 96 })
  })

  it('a disguised oversized root never reaches the decoder', async () => {
    const decode = vi.fn(async () => ({}) as CanvasImageSource)
    const doc = {
      sheets: [{ name: 'A', rows: [], cols: [], merges: [], mergeRanges: [], drawings: [picture(0)] }],
      images: [{ data: enc(`<!-- <svg width="1" height="1"> --><svg ${NS} width="20000" height="20000"/>`), mime: 'image/svg+xml' }],
    } as unknown as XlsxDocument
    const lease = await prepareSheetAssets(doc, { decodeImage: decode, profile: 'D' })
    expect(decode).not.toHaveBeenCalled()
    expect(lease.refusedImages).toBe(1)
    lease.dispose()
  })
})

describe('round 4: exact SVG attribute names; print refusal diagnostics', () => {
  const enc = (text: string) => new TextEncoder().encode(text)
  const NS = 'xmlns="http://www.w3.org/2000/svg"'
  it('withdraws case variants, duplicates and malformed sizing attributes', async () => {
    const { svgIntrinsicSize } = await import('../src/components/xlsx/sheetAssets')
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="640" height="360" WIDTH="1" HEIGHT="1"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS} WIDTH="1" HEIGHT="1"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="1" height="1" ViewBox="0 0 640 360"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="1" width="20000" height="1"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS} width=640 height="360"/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="640" height="360" junk/>`))).toBeNull()
    expect(svgIntrinsicSize(enc(`<svg ${NS} width="640px" height="360"/>`))).toEqual({ width: 640, height: 360 })
  })

  it('a disguised case-variant 20000² root never reaches the decoder', async () => {
    const decode = vi.fn(async () => ({}) as CanvasImageSource)
    const doc = {
      sheets: [{ name: 'A', rows: [], cols: [], merges: [], mergeRanges: [], drawings: [picture(0)] }],
      images: [{ data: enc(`<svg ${NS} width="20000" height="20000" WIDTH="1" HEIGHT="1"/>`), mime: 'image/svg+xml' }],
    } as unknown as XlsxDocument
    const lease = await prepareSheetAssets(doc, { decodeImage: decode, profile: 'D' })
    expect(decode).not.toHaveBeenCalled()
    expect(lease.refusedImages).toBe(1)
    lease.dispose()
  })

  it('print units report pictures refused by the shared image budget', async () => {
    const { preparePrintUnits, SheetImageCache } = await import('../src/components/xlsx/sheetAssets')
    const decode = vi.fn(async () => ({ close() {} }) as unknown as CanvasImageSource)
    const doc = {
      sheets: [
        { name: 'A', rows: [], cols: [], merges: [], mergeRanges: [], drawings: [picture(0)] },
        { name: 'B', rows: [], cols: [], merges: [], mergeRanges: [], drawings: [picture(1)] },
      ],
      images: [{ data: pngHeader(4096, 4096, 1), mime: 'image/png' }, { data: pngHeader(4096, 4096, 2), mime: 'image/png' }],
    } as unknown as XlsxDocument
    const cache = new SheetImageCache(128 * 1024 * 1024)
    const units = await preparePrintUnits(doc, { decodeImage: decode, profile: 'D', imageCache: cache })
    try {
      // Whatever is refused is reported per unit; nothing is silently blank.
      const refused = units.refusedImages
      expect(refused).toHaveLength(2)
      expect(refused[0] + refused[1] + decode.mock.calls.length).toBe(2)
    } finally { units.dispose() }
  })
})
