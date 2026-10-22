/**
 * Task S2-6: media reads and leases (slideshow spec 6.1, 7). One media part is
 * read with the bounded uncached reader, counted inflated chunk by chunk, so a
 * forged declared size or a tight aggregate budget is caught mid-inflate; the
 * Blob URL is revoked exactly once on release.
 */
import { it, describe, expect, vi, type Mock } from 'vitest'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { openShowPackage } from '../src/slideshow/read/deck'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import { MEDIA_PART_LIMIT, openMediaLease } from '../src/slideshow/media/lease'
import type { MediaRecord } from '../src/slideshow/types'
import type { PptxPackageHandle } from '../src/pptx/types'

const MiB = 1024 * 1024

interface FakeWindow extends Window {
  URL: { createObjectURL: Mock<(blob: Blob) => string>; revokeObjectURL: Mock<(url: string) => void> }
  revokes: string[]
}

/** A realm-safe window stub: only URL's Blob-URL statics are used. */
function openWindow(): FakeWindow {
  const revokes: string[] = []
  let n = 0
  return {
    URL: {
      createObjectURL: vi.fn((blob: Blob) => { void blob; return `blob:media-lease-${++n}` }),
      revokeObjectURL: vi.fn((url: string) => { revokes.push(url) }),
    },
    revokes,
  } as unknown as FakeWindow
}

/** The record the parser always hands over: the declared size comes from the package. */
async function mediaRec(handle: PptxPackageHandle, path = 'ppt/media/showMedia1.mp4'): Promise<MediaRecord> {
  const entry = handle.pkg.entry(path) as unknown as { dir?: boolean; _data?: { uncompressedSize?: number } }
  if (!entry || entry.dir) throw new Error(`${path} is not in the package`)
  const size = entry._data?.uncompressedSize
  if (!Number.isSafeInteger(size) || size! < 0) throw new Error('the part has no declared size')
  return { rId: 'rIdShow1', partPath: path, contentType: 'video/mp4', declaredBytes: size! }
}

/** A deck with one slide holding one media part of `data`. */
async function deck(data: Uint8Array, ext = 'mp4'): Promise<PptxPackageHandle> {
  const raw = await buildPptxDeck({ slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['1000', '1000'] }], media: [{ shapeId: 2, kind: 'video', ext, data }] }] })
  return openShowPackage(raw)
}

const random = (n: number) => {
  const b = new Uint8Array(n)
  for (let i = 0; i < n; i += 65536) crypto.getRandomValues(b.subarray(i, Math.min(i + 65536, n)))
  return b
}

/** Rewrites one entry's declared uncompressed size (local header and central directory). */
function forgeDeclaredSize(zip: Uint8Array, name: string, size: number): Uint8Array {
  const out = zip.slice(), view = new DataView(out.buffer), bytes = new TextEncoder().encode(name)
  const matches = (at: number) => bytes.every((b, i) => out[at + i] === b)
  for (let i = 0; i + 46 < out.length; i++) {
    const sig = view.getUint32(i, true)
    if (sig === 0x04034b50 && view.getUint16(i + 26, true) === bytes.length && matches(i + 30)) view.setUint32(i + 22, size, true)
    if (sig === 0x02014b50 && view.getUint16(i + 28, true) === bytes.length && matches(i + 46)) view.setUint32(i + 24, size, true)
  }
  return out
}

describe('bounded media reads and leases', () => {
  it('reads with the inflated counter; a forged declared size is caught at the cap', async () => {
    // The charge is the inflated length, not the declaration.
    const win = openWindow()
    const data = random(256 * 1024)
    const handle = await deck(data)
    const rec = await mediaRec(handle)
    const stream = vi.spyOn(handle.pkg.entry(rec.partPath) as unknown as { internalStream: () => unknown }, 'internalStream')
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    const lease = await openMediaLease(handle, rec, w, win)
    expect(lease).not.toHaveProperty('refused')
    expect((lease as { bytes: number }).bytes).toBe(data.length)
    expect(w.used()).toBe(data.length)
    expect(win.URL.createObjectURL.mock.calls[0]![0]!.size).toBe(data.length)
    expect((lease as { url: string }).url).toMatch(/^blob:/)
    stream.mockRestore()
    ;(lease as { release(): void }).release()

    // A declaration over the 512 MiB cap is refused unread.
    const win2 = openWindow()
    const over = await openMediaLease(handle, { ...rec, declaredBytes: MEDIA_PART_LIMIT + 1 }, w, win2)
    expect(over).toEqual({ refused: 'media-limit' })
    expect(win2.URL.createObjectURL).not.toHaveBeenCalled()
    expect(stream).not.toHaveBeenCalled()
    expect(w.used()).toBe(0)

    // A forged (smaller) declaration cannot lower the charge: with a tiny
    // workspace, the inflated counter itself refuses where the real length
    // crosses it, not where the declaration claims to.
    const win3 = openWindow()
    const forged = await openShowPackage(forgeDeclaredSize(await buildPptxDeck({ slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['1000', '1000'] }], media: [{ shapeId: 2, kind: 'video', ext: 'mp4', data: random(512 * 1024) }] }] }), 'ppt/media/showMedia1.mp4', 100))
    const rec3 = await mediaRec(forged)
    expect(rec3.declaredBytes).toBe(100)
    const w3 = createSlideShowWorkspace('desktop', { limit: 16 * 1024 })
    const refused = await openMediaLease(forged, rec3, w3, win3)
    expect(refused).toEqual({ refused: 'resource-limit' })
    expect(win3.URL.createObjectURL).not.toHaveBeenCalled()
    expect(w3.used()).toBe(0)
  })

  it('charges the inflated length; refuses one byte below the remaining budget', async () => {
    const limit = 1024 * 1024
    const hold = random(4096)
    const remaining = limit - hold.length
    // A part that fits the workspace charges exactly its inflated length,
    // nothing about the declaration.
    const win = openWindow()
    const handle = await deck(random(256 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit })
    w.lease(hold.length, 'hold')
    const lease = await openMediaLease(handle, rec, w, win)
    expect(lease).not.toHaveProperty('refused')
    expect(w.used()).toBe(hold.length + 256 * 1024)
    ;(lease as { release(): void }).release()
    expect(w.used()).toBe(hold.length)

    // One byte below the remaining budget is still refused: the peak (chunks
    // plus Blob) needs twice what the part inflates to, so an implementation
    // that only charged the finished Blob could not pass this.
    const other = await deck(random(remaining - 1))
    const win2 = openWindow()
    const refused = await openMediaLease(other, await mediaRec(other), w, win2)
    expect(refused).toEqual({ refused: 'resource-limit' })
    expect(win2.URL.createObjectURL).not.toHaveBeenCalled()
    expect(w.used()).toBe(hold.length)
  })

  it('on the 256 MiB profile a media part larger than the remaining budget is refused at the crossing chunk, not after inflation', async () => {
    const handle = await deck(random(5 * MiB))
    const rec = await mediaRec(handle)
    // The mobile profile holds everything but a small remaining budget, all of
    // it claimed by earlier work: the part must be refused by the inflated
    // counter, mid-stream.
    const w = createSlideShowWorkspace('mobile')
    w.lease(MiB * 253, 'earlier work')
    const remaining = w.limit - 253 * MiB

    // Observe the chunks the reader is handed (the bounded-chart test pattern).
    const entry = handle.pkg.entry(rec.partPath) as unknown as { internalStream: (type: string) => unknown }
    const internal = entry.internalStream.bind(entry)
    const accepted: number[] = [], refusedChunks: number[] = []
    entry.internalStream = (type: string) => {
      const stream = internal(type) as { on: (event: string, fn: (chunk: Uint8Array) => void) => unknown }
      const on = stream.on.bind(stream)
      stream.on = (event: string, fn: (chunk: Uint8Array) => void) => on(event, event !== 'data' ? fn : (chunk: Uint8Array) => {
        if (refusedChunks.length) { refusedChunks.push(chunk.length); return }
        try { fn(chunk); accepted.push(chunk.length) } catch (error) { refusedChunks.push(chunk.length); throw error }
      })
      return stream
    }

    const win = openWindow()
    const result = await openMediaLease(handle, rec, w, win)
    expect(result).toEqual({ refused: 'resource-limit' })
    // Some chunks were admitted (the read was in flight), and the refusal came
    // exactly at the chunk that crosses the remaining budget.
    expect(accepted.length).toBeGreaterThan(0)
    const fit = accepted.reduce((a, b) => a + b, 0)
    expect(fit).toBeLessThanOrEqual(remaining)
    expect(fit + refusedChunks[0]!).toBeGreaterThan(remaining)
    // No Blob was ever constructed, and the earlier charges are untouched.
    expect(win.URL.createObjectURL).not.toHaveBeenCalled()
    expect(w.used()).toBe(253 * MiB)
  })

  it('the Blob construction peak (chunks plus Blob) is charged, then the chunk lease is released', async () => {
    const data = random(256 * 1024)
    // Chunks plus Blob fit at the peak; afterwards only the Blob stays charged.
    const win = openWindow()
    const handle = await deck(data)
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 2 * data.length + 1024 })
    const lease = await openMediaLease(handle, rec, w, win)
    expect(lease).not.toHaveProperty('refused')
    expect(w.used()).toBe(data.length)
    ;(lease as { release(): void }).release()
    expect(w.used()).toBe(0)

    // One byte short of the peak is refused before any Blob URL exists.
    const other = createSlideShowWorkspace('desktop', { limit: 2 * data.length - 1 })
    const win2 = openWindow()
    const peak = await openMediaLease(handle, rec, other, win2)
    expect(peak).toEqual({ refused: 'resource-limit' })
    expect(win2.URL.createObjectURL).not.toHaveBeenCalled()
    expect(other.used()).toBe(0)
  })

  it('revokes the URL when the slide is left unless retained across slides', async () => {
    const win = openWindow()
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    const lease = await openMediaLease(handle, rec, w, win) as {
      url: string; bytes: number; spans: number; retain(slides: number): void; release(): void
    }
    // Retention across slides is recorded on the lease; nothing is revoked by it.
    lease.retain(3)
    expect(lease.spans).toBe(3)
    lease.retain(2)
    expect(lease.spans).toBe(3)
    expect(win.URL.revokeObjectURL).not.toHaveBeenCalled()
    expect(w.used()).toBe(rec.declaredBytes)

    // Release revokes the URL exactly once and gives the charge back.
    lease.release()
    expect(win.URL.revokeObjectURL).toHaveBeenCalledTimes(1)
    expect(win.URL.revokeObjectURL).toHaveBeenCalledWith(lease.url)
    expect(w.used()).toBe(0)
    lease.release()
    expect(win.URL.revokeObjectURL).toHaveBeenCalledTimes(1)
  })

  it('never uses OfficePackage.bytes', async () => {
    const win = openWindow()
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const bytes = vi.spyOn(handle.pkg, 'bytes')
    const text = vi.spyOn(handle.pkg, 'text')
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    const lease = await openMediaLease(handle, rec, w, win)
    expect(lease).not.toHaveProperty('refused')
    expect(bytes).not.toHaveBeenCalled()
    expect(text).not.toHaveBeenCalled()
    ;(lease as { release(): void }).release()
  })
})
