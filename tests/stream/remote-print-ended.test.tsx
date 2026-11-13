/** Stage 3b final fix 2: a terminal session end stops every part of an
 * unfinished print — the network job, the lazy surface import, and the real
 * print surface while its images load — and the frame never prints. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import RemoteDocView from '../../src/components/remote/RemoteDocView'
import { StreamSessionError } from '../../src/stream/remote/errors'
import type { PrintedPage, RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'

// The real surface, behind a gate on its first (dynamic) import, with its calls recorded.
const h = vi.hoisted(() => {
  let open!: () => void
  const gate = new Promise<void>((r) => { open = r })
  return { gate, open, importing: false, calls: [] as Array<{ signal: AbortSignal }> }
})
vi.mock('../../src/components/paged/printDocument', async () => {
  h.importing = true
  await h.gate
  const real = await vi.importActual<typeof import('../../src/components/paged/printDocument')>('../../src/components/paged/printDocument')
  return {
    ...real,
    printRasters: (pages: Parameters<typeof real.printRasters>[0], root: HTMLElement, opts: Parameters<typeof real.printRasters>[2]) => {
      h.calls.push({ signal: opts.signal })
      return real.printRasters(pages, root, opts)
    },
  }
})

type PrintOpts = { signal?: AbortSignal; onProgress?: (done: number, total: number) => void }

function fakeDoc(o: { maxPages?: number; pending?: boolean } = {}) {
  const count = 3, list = Array.from({ length: count }, () => ({ w: 200, h: 100 }))
  let end!: (e: StreamSessionError | undefined) => void
  const ended = new Promise<StreamSessionError | undefined>((r) => { end = r })
  const signals: AbortSignal[] = []
  const doc = {
    kind: 'remote',
    manifest: {
      name: 'deck.pptx', kind: 'pptx', units: { count, list },
      can: { search: true, copy: false, print: true, textA11y: false, present: false, notes: false, links: false },
      caps: { copy: null, a11y: 'off', print: { maxPages: o.maxPages ?? null, maxSessionPages: null } },
      recorder: { scalesBuckets: [1, 2] }, notices: [],
    },
    onManifestChange: () => () => {},
    units: async () => list.map((u) => ({ widthPx: u.w, heightPx: u.h })),
    paint: vi.fn(async () => {}),
    search: vi.fn(async () => ({ total: 0, truncated: false })),
    select: vi.fn(async () => ({ units: [], paragraphs: 0 })),
    copy: vi.fn(async () => ''),
    a11y: vi.fn(),
    print: vi.fn((r: { from: number; to: number }, po?: PrintOpts) => {
      if (po?.signal) signals.push(po.signal)
      const pages = Array.from({ length: r.to - r.from + 1 }, (_, i): PrintedPage => ({ unitIndex: r.from + i, widthPx: 200, heightPx: 100, scale: 2, png: new Blob(['png'], { type: 'image/png' }) }))
      if (!o.pending) return Promise.resolve(pages)
      return new Promise<PrintedPage[]>((_, reject) => po?.signal?.addEventListener('abort', () => reject(po.signal!.reason), { once: true }))
    }),
    close: vi.fn(),
    ended,
  }
  return { doc, end, signals }
}
type Fake = ReturnType<typeof fakeDoc>['doc']
const asDoc = (d: Fake) => d as unknown as RemoteOfficeDoc
const printBtn = (v: ReturnType<typeof render>) => v.container.querySelector<HTMLButtonElement>('button[data-ov-print-button]')
const frames = () => [...document.querySelectorAll<HTMLIFrameElement>('iframe[data-officeview-print-frame]')]
const revoked = () => new StreamSessionError('revoked')

let framePrint: ReturnType<typeof vi.fn>
beforeEach(() => {
  framePrint = vi.fn()
  let n = 0
  vi.stubGlobal('URL', Object.assign(class extends URL {}, { createObjectURL: () => `blob:raster-${++n}`, revokeObjectURL: () => {} }))
  vi.stubGlobal('CSS', { supports: () => true })
  vi.spyOn(window, 'print').mockImplementation(() => { throw new Error('the screen must never print') })
  const append = Node.prototype.appendChild
  vi.spyOn(Node.prototype, 'appendChild').mockImplementation(function <T extends Node>(this: Node, node: T): T {
    const result = append.call(this, node) as T
    if (node instanceof HTMLIFrameElement) node.contentWindow!.print = framePrint as () => void
    return result
  })
})
afterEach(() => { cleanup(); vi.restoreAllMocks(); vi.unstubAllGlobals() })

describe('remote print: a terminal session end', () => {
  // First in the file: the surface module has not been imported yet.
  it('during the lazy surface import: the surface is never started and nothing prints', async () => {
    const { doc, end, signals } = fakeDoc()
    const v = render(<RemoteDocView document={asDoc(doc)} />)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(h.importing).toBe(true))
    await act(async () => { end(revoked()) })
    expect(v.container.textContent).toContain('Viewing session ended (revoked).')
    expect(signals[0]!.aborted).toBe(true)
    await act(async () => { h.open(); await new Promise((r) => setTimeout(r, 20)) })
    expect(h.calls).toEqual([])
    expect(frames()).toEqual([])
    expect(framePrint).not.toHaveBeenCalled()
    expect(v.container.textContent).not.toContain('Preparing pages')
    expect(v.container.querySelector('[data-ov-print-notice]')).toBeNull()
  })

  it('while the surface loads its images: the surface signal aborts, the frame goes, and a late image load never prints', async () => {
    h.open()
    const { doc, end } = fakeDoc()
    const v = render(<RemoteDocView document={asDoc(doc)} />)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(frames()[0]?.contentDocument?.querySelector('img')).toBeTruthy())
    const frame = frames()[0]!, img = frame.contentDocument!.querySelector('img')!
    const { signal } = h.calls.at(-1)!
    expect(signal.aborted).toBe(false)
    await act(async () => { end(revoked()) })
    expect(signal.aborted).toBe(true)
    expect(frame.isConnected).toBe(false)
    // The image finishing late changes nothing.
    await act(async () => { img.dispatchEvent(new Event('load')); await new Promise((r) => setTimeout(r, 20)) })
    expect(framePrint).not.toHaveBeenCalled()
    expect(v.container.textContent).not.toContain('Preparing pages')
    expect(v.container.querySelector('[data-ov-print-notice]')).toBeNull()
    expect(printBtn(v)).toBeNull()
  })

  it('during the network job: the job aborts silently, progress clears and Print is gone', async () => {
    h.open()
    const { doc, end, signals } = fakeDoc({ pending: true })
    const v = render(<RemoteDocView document={asDoc(doc)} />)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(v.container.textContent).toContain('Preparing pages 0 of 3'))
    await act(async () => { end(revoked()) })
    expect(signals[0]!.aborted).toBe(true)
    await act(async () => { await new Promise((r) => setTimeout(r, 20)) })
    expect(v.container.textContent).not.toContain('Preparing pages')
    expect(v.container.querySelector('[data-ov-print-notice]')).toBeNull()
    expect(printBtn(v)).toBeNull()
    expect(frames()).toEqual([])
    expect(framePrint).not.toHaveBeenCalled()
  })

  it('the surface aborts on the end itself, before any React render or effect (microtasks only)', async () => {
    h.open()
    const { doc, end } = fakeDoc()
    const v = render(<RemoteDocView document={asDoc(doc)} />)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(frames()[0]?.contentDocument?.querySelector('img')).toBeTruthy())
    const { signal } = h.calls.at(-1)!
    // No act(), no timers: a late image load in the same task must already find the surface aborted.
    end(revoked())
    for (let i = 0; i < 5; i++) await Promise.resolve()
    expect(signal.aborted).toBe(true)
    await act(async () => { await new Promise((r) => setTimeout(r, 20)) })
    expect(framePrint).not.toHaveBeenCalled()
  })

  for (const first of ['print', 'ended'] as const) {
    it(`doc.print rejecting with the terminal error (${first} settling first) shows no failure notice`, async () => {
      h.open()
      const { doc, end } = fakeDoc()
      let fail!: (e: unknown) => void
      doc.print.mockImplementation(() => new Promise<PrintedPage[]>((_, reject) => { fail = reject }))
      const v = render(<RemoteDocView document={asDoc(doc)} />)
      fireEvent.click(printBtn(v)!)
      await waitFor(() => expect(doc.print).toHaveBeenCalledTimes(1))
      const e = revoked()
      await act(async () => { if (first === 'print') { fail(e); end(e) } else { end(e); fail(e) } })
      await act(async () => { await new Promise((r) => setTimeout(r, 20)) })
      expect(v.container.textContent).toContain('Viewing session ended (revoked).')
      expect(v.container.querySelector('[data-ov-print-notice]')).toBeNull()
      expect(v.container.textContent).not.toContain('could not be printed')
      expect(v.container.textContent).not.toContain('Preparing pages')
    })
  }

  it('closes an open range dialog', async () => {
    const { doc, end } = fakeDoc({ maxPages: 2 })
    const v = render(<RemoteDocView document={asDoc(doc)} />)
    fireEvent.click(printBtn(v)!)
    expect(v.queryByRole('dialog', { name: 'Print pages' })).not.toBeNull()
    await act(async () => { end(revoked()) })
    expect(v.queryByRole('dialog', { name: 'Print pages' })).toBeNull()
    expect(doc.print).not.toHaveBeenCalled()
  })

  it("an old document's late end leaves the replacement's print alone", async () => {
    h.open()
    const a = fakeDoc(), b = fakeDoc()
    const v = render(<RemoteDocView document={asDoc(a.doc)} />)
    v.rerender(<RemoteDocView document={asDoc(b.doc)} />)
    await act(async () => { a.end(revoked()) })
    expect(printBtn(v)).not.toBeNull()
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(frames()[0]?.contentDocument?.querySelector('img')).toBeTruthy())
    const frame = frames()[0]!
    for (const img of frame.contentDocument!.querySelectorAll('img')) img.dispatchEvent(new Event('load'))
    // Images load one at a time: release each as it appears.
    await waitFor(() => {
      for (const img of frame.contentDocument!.querySelectorAll('img')) img.dispatchEvent(new Event('load'))
      expect(framePrint).toHaveBeenCalledTimes(1)
    })
    expect(b.signals[0]!.aborted).toBe(false)
  })
})
