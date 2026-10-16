/**
 * PR #11 review findings: byte sources are bounded before the legacy reader
 * inflates package metadata, and the per-load abort signal reaches fetch and
 * an already-aborted stream.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { loadOfficeSource } from '../src/components/loadOfficeSource'
import { OfficePackage } from '../src/core/zip'
import { ImportRefusal } from '../src/xlsx/storage/importSession'
import { openByteStream } from '../src/core/stream'
import { buildZip } from './helpers/zip-builder'
import { xlsxEntries } from './helpers/xlsx-package'

const MiB = 1024 * 1024
const CT = '[Content_Types].xml'

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

/** An honest spreadsheet whose content-types part inflates to ~8 MiB from a few KiB. */
function contentTypesBomb(): Uint8Array {
  const entries = xlsxEntries({})
  const ct = entries.find((e) => e.name === CT)!
  const text = typeof ct.data === 'string' ? ct.data : new TextDecoder().decode(ct.data)
  const padded = text.replace('</Types>', `${' '.repeat(8 * MiB)}</Types>`)
  return buildZip(entries.map((e) => (e === ct ? { ...e, data: padded, method: 8 } : e))).bytes
}

async function rejection(p: Promise<unknown>): Promise<unknown> {
  try { await p } catch (e) { return e }
  throw new Error('expected a rejection')
}

describe('byte sources are bounded before the legacy reader inflates metadata (PR #11 P1)', () => {
  for (const [label, wrap] of [
    ['Uint8Array', (b: Uint8Array) => b],
    ['ArrayBuffer', (b: Uint8Array) => b.slice().buffer],
    ['Blob', (b: Uint8Array) => new Blob([b as BlobPart])],
  ] as const) {
    it(`${label}: an oversized [Content_Types].xml is refused as a resource limit without being inflated whole`, async () => {
      const bytes = contentTypesBomb()
      expect(bytes.length).toBeLessThan(256 * 1024)
      const text = vi.spyOn(OfficePackage.prototype, 'text')
      const error = await rejection(loadOfficeSource(wrap(bytes)))
      expect(text.mock.calls.map((c) => c[0])).not.toContain(CT)
      // the typed refusal <OfficeFile> surfaces with its code and reason
      expect(error).toBeInstanceOf(ImportRefusal)
      expect(error).toMatchObject({ code: 'resource-limit', reason: 'limit' })
    })
  }
})

/** A fetch that, like the real one, rejects when its init signal aborts and never sees headers otherwise. */
function pendingFetch() {
  const calls: Array<{ input: unknown; init: RequestInit | undefined }> = []
  const fetch = vi.fn((input: unknown, init?: RequestInit) => {
    calls.push({ input, init })
    return new Promise<Response>((_, reject) => {
      const signal = init?.signal
      if (signal?.aborted) reject(new DOMException('aborted', 'AbortError'))
      signal?.addEventListener('abort', () => reject(new DOMException('aborted', 'AbortError')))
    })
  })
  vi.stubGlobal('fetch', fetch)
  return calls
}

describe('the per-load signal reaches fetch (PR #11 P2, stream.ts openFetch)', () => {
  it('URL descriptor: aborting before headers arrive aborts the fetch, keeping headers and credentials', async () => {
    const calls = pendingFetch()
    const controller = new AbortController()
    const load = loadOfficeSource({ url: 'https://example.test/a.docx', headers: { Authorization: 'Bearer t' }, credentials: 'include' }, { signal: controller.signal })
    await vi.waitFor(() => expect(calls).toHaveLength(1))
    expect(calls[0].init?.headers).toEqual({ Authorization: 'Bearer t' })
    expect(calls[0].init?.credentials).toBe('include')
    controller.abort()
    expect(calls[0].init?.signal?.aborted).toBe(true)
    await expect(load).rejects.toThrow()
  })

  it('Request: the load signal and the Request\'s own signal both abort the fetch, with or without AbortSignal.any', async () => {
    const statics = AbortSignal as { any?: unknown }
    const native = statics.any
    try {
    for (const [which, withAny] of [['load', true], ['request', true], ['load', false], ['request', false]] as const) {
      statics.any = withAny ? native : undefined
      const calls = pendingFetch()
      const loadCtl = new AbortController()
      const ownCtl = new AbortController()
      const load = loadOfficeSource(new Request('https://example.test/a.docx', { signal: ownCtl.signal }), { signal: loadCtl.signal })
      await vi.waitFor(() => expect(calls).toHaveLength(1))
      expect(calls[0].init?.signal?.aborted).toBe(false)
      ;(which === 'load' ? loadCtl : ownCtl).abort()
      expect(calls[0].init?.signal?.aborted).toBe(true)
      await expect(load).rejects.toThrow()
    }
    } finally { statics.any = native }
  })
})

/** A body that records pulls and cancellation of its underlying source. */
function trackedBody() {
  const log = { pulls: 0, cancelled: false }
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) { log.pulls++; controller.enqueue(new Uint8Array(1024)) },
    cancel() { log.cancelled = true },
  }, { highWaterMark: 0 })
  return { stream, log }
}

describe('an already-aborted signal cancels the stream it opens (PR #11 P2, stream.ts pullStream)', () => {
  it('openByteStream with a pre-aborted signal cancels the reader and leaves the body unlocked', async () => {
    const { stream, log } = trackedBody()
    const controller = new AbortController()
    controller.abort()
    const opened = await openByteStream(new Response(stream), {}, controller.signal)
    expect(opened.kind).toBe('stream')
    await vi.waitFor(() => expect(log.cancelled).toBe(true))
    expect(stream.locked).toBe(false)
    expect(log.pulls).toBe(0)
  })

  it('a promised Response that resolves after its load was cancelled is not left downloading', async () => {
    const { stream, log } = trackedBody()
    const controller = new AbortController()
    let resolve!: (r: Response) => void
    const load = loadOfficeSource(new Promise<Response>((r) => { resolve = r }), { signal: controller.signal })
    controller.abort()
    resolve(new Response(stream))
    await expect(load).rejects.toThrow()
    await vi.waitFor(() => expect(log.cancelled).toBe(true))
    expect(stream.locked).toBe(false)
  })
})
