/**
 * Task S2-6: media reads and leases (slideshow spec 6.1, 7). One media part is
 * read with the bounded uncached reader, counted inflated chunk by chunk, so a
 * forged declared size or a tight aggregate budget is caught mid-inflate; the
 * Blob URL is revoked exactly once on release, and refusals are values, never
 * raw rejections the show coordinator would have to guard against.
 */
import { it, describe, expect, vi, type Mock } from 'vitest'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { openShowPackage } from '../src/slideshow/read/deck'
import { createSlideShowWorkspace } from '../src/slideshow/workspace'
import { MEDIA_PART_LIMIT, openMediaLease } from '../src/slideshow/media/lease'
import type { MediaRecord } from '../src/slideshow/types'
import type { OfficePackage } from '../src/core/zip'
import type { PptxPackageHandle } from '../src/pptx/types'

const MiB = 1024 * 1024

interface FakeWindow extends Window {
  URL: { createObjectURL: Mock<(blob: Blob) => string>; revokeObjectURL: Mock<(url: string) => void> }
  revokes: string[]
}

/** A realm-safe window stub: only URL's Blob-URL statics are used. */
function openWindow(revokeThrows = false): FakeWindow {
  const revokes: string[] = []
  let n = 0
  return {
    URL: {
      createObjectURL: vi.fn((blob: Blob) => { void blob; return `blob:media-lease-${++n}` }),
      revokeObjectURL: vi.fn((url: string) => {
        if (revokeThrows) throw new TypeError('the realm is gone')
        revokes.push(url)
      }),
    },
    revokes,
  } as unknown as FakeWindow
}

/** The record the parser always hands over: the declared size comes from the package. */
async function mediaRec(handle: PptxPackageHandle, path?: string): Promise<MediaRecord> {
  const part = path ?? (handle.pkg.paths(p => p.startsWith('ppt/media/'))[0] ?? 'ppt/media/showMedia1.mp4')
  const entry = handle.pkg.entry(part) as unknown as { dir?: boolean; _data?: { uncompressedSize?: number } }
  if (!entry || entry.dir) throw new Error(`${part} is not in the package`)
  const size = entry._data?.uncompressedSize
  if (!Number.isSafeInteger(size) || size! < 0) throw new Error('the part has no declared size')
  return { rId: 'rIdShow1', partPath: part, contentType: 'video/mp4', declaredBytes: size! }
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

/** A JSZip-shaped package whose stream feeds `sizes` synthetic chunks (chunk memory is not allocated), so the 512 MiB cap and mid-read aborts are drivable. */
function fakeHandle(sizes: number[], between?: () => void): PptxPackageHandle {
  const stream: {
    on(event: string, fn: (arg: never) => void): unknown
    resume(): void
    onData?: (chunk: { length: number }) => void
    onError?: (e: Error) => void
    onEnd?: () => void
    _worker: { previous?: unknown; push(): void; end(): void; error(): void }
  } = {
    on(event, fn) {
      if (event === 'data') stream.onData = fn as (chunk: { length: number }) => void
      else if (event === 'error') stream.onError = fn as (e: Error) => void
      else if (event === 'end') stream.onEnd = fn as () => void
      return stream
    },
    onData: undefined,
    onError: undefined,
    onEnd: undefined,
    _worker: { push() {}, end() {}, error() {} },
    resume() {
      for (const n of sizes) {
        if (stream.onData === undefined) return
        try { stream.onData({ length: n }) } catch (e) { stream.onError?.(e as Error); return }
        try { between?.() } catch (e) { stream.onError?.(e as Error); return }
      }
      stream.onEnd?.()
    },
  }
  return { pkg: { entry: () => ({ internalStream: () => stream }) } as unknown as OfficePackage, slideParts: [] } as unknown as PptxPackageHandle
}

describe('bounded media reads and leases', () => {
  it('reads with the inflated counter; an over-cap declaration is refused unread', async () => {
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

  it('the Blob carries only audio or video content types', async () => {
    // A hostile declaration cannot make the Blob a renderable document type.
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const win = openWindow()
    const hostile = await openMediaLease(handle, { ...rec, contentType: 'text/html' }, createSlideShowWorkspace('desktop', { limit: 1024 * 1024 }), win)
    expect(hostile).not.toHaveProperty('refused')
    expect(win.URL.createObjectURL.mock.calls[0]![0]!.type).toBe('')
    ;(hostile as { release(): void }).release()
    // The honest declarations pass through unchanged.
    for (const type of ['video/mp4', 'audio/mpeg']) {
      const w2 = openWindow()
      const ok = await openMediaLease(handle, { ...rec, contentType: type }, createSlideShowWorkspace('desktop', { limit: 1024 * 1024 }), w2)
      expect(ok).not.toHaveProperty('refused')
      expect(w2.URL.createObjectURL.mock.calls[0]![0]!.type).toBe(type)
      ;(ok as { release(): void }).release()
    }
  })

  it('charges exactly the inflated length; a part one byte below the remaining budget still cannot make the peak', async () => {
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
    // Inflation actually stops at the crossing chunk: exactly one refusal.
    expect(refusedChunks.length).toBe(1)
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
    const handle = await deck(data)
    const rec = await mediaRec(handle)
    // Chunks plus Blob fit at the peak; afterwards only the Blob stays charged.
    const win = openWindow()
    const w = createSlideShowWorkspace('desktop', { limit: 2 * data.length + 1024 })
    // Observe the ledger inside the Blob constructor (the peak the spec charges).
    const RealBlob = Blob
    let peak = 0
    class SpyBlob extends RealBlob {
      constructor(parts: BlobPart[], init: { type?: string }) {
        peak = w.used()
        super(parts, init)
      }
    }
    vi.stubGlobal('Blob', SpyBlob)
    let lease: { release(): void } | undefined
    try {
      lease = (await openMediaLease(handle, rec, w, win)) as { release(): void }
    } finally {
      vi.unstubAllGlobals()
    }
    expect(lease).not.toHaveProperty('refused')
    expect(peak).toBe(2 * data.length)
    expect(w.used()).toBe(data.length)
    lease!.release()
    expect(w.used()).toBe(0)

    // One byte short of the peak is refused before any Blob URL exists.
    const other = createSlideShowWorkspace('desktop', { limit: 2 * data.length - 1 })
    const win2 = openWindow()
    const result = await openMediaLease(handle, rec, other, win2)
    expect(result).toEqual({ refused: 'resource-limit' })
    expect(win2.URL.createObjectURL).not.toHaveBeenCalled()
    expect(other.used()).toBe(0)
  })

  it('a retained lease survives slide leaves and revokes when the span ends', async () => {
    const win = openWindow()
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    const lease = await openMediaLease(handle, rec, w, win) as {
      url: string; bytes: number; spans: number; retain(slides: number): void; leave(): void; release(): void
    }
    lease.retain(3)
    expect(lease.spans).toBe(3)
    // Two slide leaves with retention left: the URL stays, the charge stays.
    lease.leave()
    expect(lease.spans).toBe(2)
    expect(win.URL.revokeObjectURL).not.toHaveBeenCalled()
    expect(w.used()).toBe(rec.declaredBytes)
    lease.leave()
    expect(lease.spans).toBe(1)
    expect(win.URL.revokeObjectURL).not.toHaveBeenCalled()
    // The last leave ends the span: the URL is revoked exactly once and the
    // charge is back.
    lease.leave()
    expect(win.URL.revokeObjectURL).toHaveBeenCalledTimes(1)
    expect(win.URL.revokeObjectURL).toHaveBeenCalledWith(lease.url)
    expect(w.used()).toBe(0)
  })

  it('retain validates its span, the methods survive destructuring, and spans is not writable', async () => {
    const win = openWindow()
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    const lease = await openMediaLease(handle, rec, w, win) as unknown as Record<string, unknown> & { spans: number; release(): void }
    // Destructured methods must not depend on `this`; reads of spans work.
    const { retain, leave, release } = lease as never as { retain(n: number): void; leave(): void; release(): void }
    retain(1)
    leave()
    expect(lease.spans).toBe(0)
    // Strict-mode modules make a getter-only write a loud error.
    expect(() => { lease.spans = 99 }).toThrow(TypeError)
    expect(lease.spans).toBe(0)
    release()
    expect(w.used()).toBe(0)

    // A fractional or negative span is a programming error, not silence.
    const again = await openMediaLease(handle, rec, w, win) as { retain(n: number): void; spans: number }
    for (const bad of [2.5, -1, NaN])
      expect(() => again.retain(bad)).toThrow(RangeError)
    expect(again.spans).toBe(0)
    ;(again as unknown as { release(): void }).release()
  })

  it('indefinite retention never self-revokes; release still does', async () => {
    const win = openWindow()
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    const lease = await openMediaLease(handle, rec, w, win) as { spans: number; retain(n: number): void; leave(): void; release(): void }
    // "Plays until the next transition sound" keeps it alive until it is stopped.
    lease.retain(Infinity)
    expect(lease.spans).toBe(Infinity)
    for (let i = 0; i < 5; i++) lease.leave()
    expect(win.URL.revokeObjectURL).not.toHaveBeenCalled()
    expect(w.used()).toBe(rec.declaredBytes)
    lease.release()
    expect(win.URL.revokeObjectURL).toHaveBeenCalledTimes(1)
    expect(w.used()).toBe(0)
  })

  it('release() refunds the charge even when revokeObjectURL throws', async () => {
    const win = openWindow(true)
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    const lease = await openMediaLease(handle, rec, w, win) as { release(): void }
    expect(() => lease.release()).not.toThrow()
    expect(w.used()).toBe(0)
    lease.release()
    expect(win.URL.revokeObjectURL).toHaveBeenCalledTimes(1)
  })

  it('a window without Blob-URL support is refused before anything is read; a throwing createObjectURL likewise', async () => {
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const entry = handle.pkg.entry(rec.partPath) as unknown as { internalStream: () => unknown }
    const stream = vi.spyOn(entry, 'internalStream')
    const bare = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    expect(await openMediaLease(handle, rec, bare, {} as Window)).toEqual({ refused: 'media-unavailable' })
    expect(stream).not.toHaveBeenCalled()
    expect(bare.used()).toBe(0)
    // A half-torn realm whose createObjectURL throws is a refusal, not a rejection.
    const throwing: Window = { URL: { createObjectURL: () => { throw new TypeError('dead realm') }, revokeObjectURL: () => {} } } as unknown as Window
    const degraded = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    expect(await openMediaLease(handle, rec, degraded, throwing)).toEqual({ refused: 'media-unavailable' })
    expect(degraded.used()).toBe(0)
  })

  it('a Blob construction failure is a resource refusal, not a rejection', async () => {
    const win = openWindow()
    const handle = await deck(random(256 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    vi.stubGlobal('Blob', function failing() { throw new RangeError('quota') })
    try {
      expect(await openMediaLease(handle, rec, w, win)).toEqual({ refused: 'resource-limit' })
    } finally {
      vi.unstubAllGlobals()
    }
    expect(win.URL.createObjectURL).not.toHaveBeenCalled()
    expect(w.used()).toBe(0)
  })

  it('a zero-byte part is refused, never an empty live URL', async () => {
    const win = openWindow()
    const handle = await deck(new Uint8Array(0))
    // A zero-byte entry has no declared size in the zip directory at all; the
    // caller's record still reaches the lease open.
    const rec = { rId: 'rIdShow1', partPath: 'ppt/media/showMedia1.mp4', contentType: 'video/mp4', declaredBytes: 0 }
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    expect(await openMediaLease(handle, rec, w, win)).toEqual({ refused: 'media-empty' })
    expect(win.URL.createObjectURL).not.toHaveBeenCalled()
    expect(w.used()).toBe(0)
  })

  it('a stream that genuinely inflates past the 512 MiB cap gets the media-limit label even under a lying declaration', async () => {
    // Synthetic chunks of 300 MiB each: the second one crosses the cap while
    // the declaration stays at 100 bytes.
    const handle = fakeHandle([300 * MiB, 300 * MiB])
    const rec = await mediaRec(handle, 'ppt/media/fake.mp4').catch(() => ({ rId: 'rIdShow1', partPath: 'ppt/media/fake.mp4', contentType: 'video/mp4', declaredBytes: 100 }))
    const w = createSlideShowWorkspace('desktop', { limit: 600 * MiB })
    const win = openWindow()
    expect(await openMediaLease(handle, rec as MediaRecord, w, win)).toEqual({ refused: 'media-limit' })
    expect(win.URL.createObjectURL).not.toHaveBeenCalled()
    expect(w.used()).toBe(0)
  })

  it('an abort before or during the read rejects with nothing charged', async () => {
    const before = new AbortController()
    before.abort()
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    await expect(openMediaLease(handle, rec, w, openWindow(), before.signal)).rejects.toMatchObject({ code: 'cancelled' })
    expect(w.used()).toBe(0)

    // Mid-read: the abort lands between chunks.
    const mid = new AbortController()
    const fake = fakeHandle([64 * 1024, 64 * 1024], () => mid.abort())
    const win = openWindow()
    await expect(openMediaLease(fake, await mediaRec(fake, 'ppt/media/fake.mp4').catch(() => ({ rId: 'rIdShow1', partPath: 'ppt/media/fake.mp4', contentType: 'video/mp4', declaredBytes: 128 * 1024 })), w, win, mid.signal)).rejects.toMatchObject({ code: 'cancelled' })
    expect(w.used()).toBe(0)
    expect(win.URL.createObjectURL).not.toHaveBeenCalled()
  })

  it('a part missing from the package is refused', async () => {
    const win = openWindow()
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    expect(await openMediaLease(handle, { ...rec, partPath: 'ppt/media/nope.mp4' }, w, win)).toEqual({ refused: 'media-missing' })
    expect(w.used()).toBe(0)
  })

  it('never uses OfficePackage.bytes; the host realm URL statics are untouched', async () => {
    const win = openWindow()
    const handle = await deck(random(64 * 1024))
    const rec = await mediaRec(handle)
    const bytes = vi.spyOn(handle.pkg, 'bytes')
    const text = vi.spyOn(handle.pkg, 'text')
    const w = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
    const leased = await openMediaLease(handle, rec, w, win)
    expect(leased).not.toHaveProperty('refused')
    expect(bytes).not.toHaveBeenCalled()
    expect(text).not.toHaveBeenCalled()
    ;(leased as { release(): void }).release()
    // The global URL's Blob-URL statics never saw a call.
    const global = URL as unknown as { createObjectURL?: Mock<(b: Blob) => string>; revokeObjectURL?: Mock<(u: string) => void> }
    if (typeof global.createObjectURL === 'function' && typeof global.revokeObjectURL === 'function') {
      const created = vi.spyOn(global, 'createObjectURL'), revoked = vi.spyOn(global, 'revokeObjectURL')
      const w2 = createSlideShowWorkspace('desktop', { limit: 1024 * 1024 })
      const again = await openMediaLease(handle, rec, w2, openWindow())
      ;(again as { release(): void }).release()
      expect(created).not.toHaveBeenCalled()
      expect(revoked).not.toHaveBeenCalled()
    }
  })
})
