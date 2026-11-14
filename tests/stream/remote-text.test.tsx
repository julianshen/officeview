/** Stage 3a Task 7: search bar, hit/selection overlay and copy in the remote view. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { StrictMode } from 'react'
import RemoteDocView, { TEXT_UNAVAILABLE } from '../../src/components/remote/RemoteDocView'
import { A11Y_RETRY_MS } from '../../src/components/remote/RemoteText'
import { StreamSessionError } from '../../src/stream/remote/errors'
import type { RemoteOfficeDoc, RemoteRange, RemoteSearchHit } from '../../src/stream/remote/remote-doc'

afterEach(() => { cleanup(); vi.restoreAllMocks(); vi.useRealTimers() })

type Can = { search?: boolean; copy?: boolean; textA11y?: boolean }
type TextCaps = Omit<RemoteOfficeDoc['manifest']['caps'], 'print'>
function fakeDoc(can: Can = {}, textCaps: TextCaps = { copy: null, a11y: 'off' }) {
  const caps: RemoteOfficeDoc['manifest']['caps'] = { ...textCaps, print: { maxPages: null, maxSessionPages: null } }
  const manifest = {
    name: 'deck.pptx', kind: 'pptx', units: { count: 2, list: [{ w: 200, h: 100 }, { w: 200, h: 100 }] },
    can: { search: true, copy: false, print: true, textA11y: false, present: true, notes: false, links: false, ...can },
    caps, recorder: { scalesBuckets: [1, 2] }, notices: [],
  }
  const listeners = new Set<() => void>()
  const doc = {
    kind: 'remote', manifest,
    onManifestChange: vi.fn((cb: () => void) => { listeners.add(cb); return () => { listeners.delete(cb) } }),
    /** A resume replacing the manifest (R3): new `can`/`caps`, then the listeners. */
    resume(can: Can, textCaps: TextCaps) {
      const caps: RemoteOfficeDoc['manifest']['caps'] = { ...textCaps, print: { maxPages: null, maxSessionPages: null } }
      doc.manifest = { ...doc.manifest, can: { ...doc.manifest.can, ...can }, caps }
      act(() => { for (const l of [...listeners]) l() })
    },
    listeners,
    units: async () => manifest.units.list.map((u) => ({ widthPx: u.w, heightPx: u.h })),
    paint: vi.fn(async () => {}),
    search: vi.fn(async (_q: string, _o?: { signal?: AbortSignal; onHits?: (h: RemoteSearchHit[]) => void }) => ({ total: 0, truncated: false })),
    select: vi.fn(async (_r: RemoteRange, _s?: AbortSignal) => ({ units: [] as RemoteSearchHit[], paragraphs: 0 })),
    copy: vi.fn(async (_r: RemoteRange) => ''),
    a11y: vi.fn(),
    close: vi.fn(),
    ended: new Promise(() => {}),
  }
  return doc
}
const view = (doc: ReturnType<typeof fakeDoc>) => render(<RemoteDocView document={doc as unknown as RemoteOfficeDoc} />)
const scroller = (v: ReturnType<typeof render>) => v.container.querySelector('[data-officeview-root]') as HTMLElement
const key = (el: Element, k: string, o: KeyboardEventInit = {}) => fireEvent.keyDown(el, { key: k, ctrlKey: true, ...o })
const status = (v: ReturnType<typeof render>) => v.container.querySelector('[role="status"][data-officeview-notice]')!
const rect = (x: number, y: number, width = 20, height = 10) => ({ x, y, width, height })

describe('remote view: TEXT_UNAVAILABLE notice', () => {
  it('is absent when copy is on', () => {
    const v = view(fakeDoc({ copy: true }))
    expect(v.container.textContent).not.toContain(TEXT_UNAVAILABLE)
  })
  it('is absent when only text a11y is on', () => {
    const v = view(fakeDoc({ textA11y: true }))
    expect(v.container.textContent).not.toContain(TEXT_UNAVAILABLE)
  })
  it('is shown once when copy and text a11y are both off', () => {
    const v = view(fakeDoc())
    expect(v.getAllByText(TEXT_UNAVAILABLE).length).toBe(1)
  })
})

describe('remote view: wrapper', () => {
  it('each canvas sits in a data-unit-wrap wrapper with its sizing unchanged', () => {
    const v = view(fakeDoc())
    const wraps = v.container.querySelectorAll('[data-unit-wrap]')
    expect([...wraps].map((w) => w.getAttribute('data-unit-wrap'))).toEqual(['0', '1'])
    const c = wraps[0]!.querySelector('canvas[data-unit="0"]') as HTMLCanvasElement
    expect(c.style.width).toBe('200px')
    expect(c.style.maxWidth).toBe('100%')
    expect(c.style.height).toBe('auto')
  })
})

describe('remote view: search', () => {
  it('Ctrl+F opens the bar; a query searches after the debounce and draws percent-positioned hit boxes', async () => {
    const doc = fakeDoc()
    doc.search.mockImplementation(async (_q, o) => {
      o?.onHits?.([{ unit: 1, rects: [rect(100, 50)] }])
      o?.onHits?.([{ unit: 1, rects: [rect(100, 50)] }, { unit: 0, rects: [rect(20, 10), rect(60, 10)] }])
      return { total: 2, truncated: false }
    })
    const v = view(doc)
    expect(v.container.querySelector('input')).toBeNull()
    key(scroller(v), 'f')
    const input = v.getByLabelText('Search document') as HTMLInputElement
    expect(document.activeElement).toBe(input)
    fireEvent.change(input, { target: { value: 'tex' } })
    expect(doc.search).not.toHaveBeenCalled()
    await waitFor(() => expect(doc.search).toHaveBeenCalledTimes(1))
    expect(doc.search.mock.calls[0]![0]).toBe('tex')
    expect(doc.search.mock.calls[0]![1]!.signal).toBeInstanceOf(AbortSignal)
    await waitFor(() => expect(v.container.querySelectorAll('[data-hit]').length).toBe(3))
    // Document order: unit 0's hit is first and active.
    const active = v.container.querySelectorAll('[data-hit-active]')
    expect(active.length).toBe(2)
    expect(active[0]!.closest('[data-unit-wrap]')!.getAttribute('data-unit-wrap')).toBe('0')
    const box = active[0] as HTMLElement
    expect([box.style.left, box.style.top, box.style.width, box.style.height]).toEqual(['10%', '10%', '10%', '10%'])
    expect(box.getAttribute('aria-hidden')).toBe('true')
    expect(box.style.pointerEvents).toBe('none')
    expect(box.textContent).toBe('')
    expect(v.getByTestId('officeview-search-status').textContent).toBe('1 of 2')
    fireEvent.keyDown(input, { key: 'Enter' })
    expect(v.getByTestId('officeview-search-status').textContent).toBe('2 of 2')
    expect(v.container.querySelector('[data-hit-active]')!.closest('[data-unit-wrap]')!.getAttribute('data-unit-wrap')).toBe('1')
    fireEvent.keyDown(input, { key: 'Enter', shiftKey: true })
    expect(v.getByTestId('officeview-search-status').textContent).toBe('1 of 2')
    // Escape closes the bar, clears the boxes and aborts the search.
    fireEvent.keyDown(input, { key: 'Escape' })
    expect(v.container.querySelector('input')).toBeNull()
    expect(v.container.querySelectorAll('[data-hit]').length).toBe(0)
    expect(doc.search.mock.calls[0]![1]!.signal!.aborted).toBe(true)
  })

  it('a truncated search shows "+"; a new query aborts the previous search; an empty query sends nothing', async () => {
    const doc = fakeDoc()
    let pending: ((v: { total: number; truncated: boolean }) => void) | undefined
    doc.search.mockImplementationOnce((_q, o) => new Promise((ok, no) => {
      o?.onHits?.([{ unit: 0, rects: [rect(0, 0)] }])
      pending = ok
      o?.signal?.addEventListener('abort', () => no(new DOMException('aborted', 'AbortError')))
    }))
    doc.search.mockImplementationOnce(async (_q, o) => { o?.onHits?.([{ unit: 0, rects: [rect(0, 0)] }]); return { total: 5000, truncated: true } })
    const v = view(doc)
    key(scroller(v), 'f')
    const input = v.getByLabelText('Search document')
    fireEvent.change(input, { target: { value: 'a' } })
    await waitFor(() => expect(doc.search).toHaveBeenCalledTimes(1))
    // The running count shows while the search is pending.
    await waitFor(() => expect(v.getByTestId('officeview-search-status').textContent).toBe('1 of 1'))
    fireEvent.change(input, { target: { value: 'ab' } })
    await waitFor(() => expect(doc.search).toHaveBeenCalledTimes(2))
    expect(doc.search.mock.calls[0]![1]!.signal!.aborted).toBe(true)
    await waitFor(() => expect(v.getByTestId('officeview-search-status').textContent).toBe('1 of 5000+'))
    pending?.({ total: 1, truncated: false })
    fireEvent.change(input, { target: { value: '' } })
    await new Promise((r) => setTimeout(r, 300))
    expect(doc.search).toHaveBeenCalledTimes(2)
    expect(v.container.querySelectorAll('[data-hit]').length).toBe(0)
  })

  it('without can.search Ctrl+F does not open a bar', () => {
    const v = view(fakeDoc({ search: false }))
    key(scroller(v), 'f')
    expect(v.container.querySelector('input')).toBeNull()
  })
})

describe('remote view: selection and copy', () => {
  function writer() {
    const writeText = vi.fn(async (_t: string) => {})
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
    return writeText
  }

  it('Ctrl+A selects everything and draws data-sel boxes; Ctrl+C copies it to the clipboard', async () => {
    const writeText = writer()
    const doc = fakeDoc({ copy: true }, { copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'unit' })
    doc.select.mockResolvedValue({ units: [{ unit: 0, rects: [rect(0, 0, 200, 10)] }, { unit: 1, rects: [rect(0, 0, 100, 10)] }], paragraphs: 2 })
    doc.copy.mockResolvedValue('one\ntwo')
    const v = view(doc)
    key(scroller(v), 'a')
    expect(doc.select).toHaveBeenCalledWith({ all: true }, expect.any(AbortSignal))
    await waitFor(() => expect(v.container.querySelectorAll('[data-sel]').length).toBe(2))
    expect((v.container.querySelector('[data-sel]') as HTMLElement).style.width).toBe('100%')
    key(scroller(v), 'c')
    await waitFor(() => expect(writeText).toHaveBeenCalledWith('one\ntwo'))
    expect(doc.copy).toHaveBeenCalledWith({ all: true })
    expect(status(v).textContent).toBe('')
  })

  it('Ctrl+C with no selection copies nothing', async () => {
    const doc = fakeDoc({ copy: true })
    const v = view(doc)
    key(scroller(v), 'c')
    await act(async () => {})
    expect(doc.copy).not.toHaveBeenCalled()
  })

  const refusals: Array<[string, StreamSessionError | Error, RemoteOfficeDoc['manifest']['caps']['copy']]> = [
    ['Selection too large to copy (limit: 3 paragraphs).', new StreamSessionError('copy-too-large', 'x', { limit: 3, requested: 4 }), { maxParagraphs: 3, maxSessionParagraphs: null }],
    ['Selection too large to copy.', new StreamSessionError('copy-too-large', 'x', { limit: 4194304, requested: 5000000 }), { maxParagraphs: 3, maxSessionParagraphs: null }],
    ['Selection too large to copy.', new StreamSessionError('copy-too-large', 'x', { limit: 4194304, requested: 5000000 }), { maxParagraphs: null, maxSessionParagraphs: null }],
    ['Copy limit for this session reached (10 paragraphs).', new StreamSessionError('copy-quota-exhausted', 'x', { limit: 10, requested: 2 }), { maxParagraphs: 3, maxSessionParagraphs: 10 }],
    ['Too many requests. Try again shortly.', new StreamSessionError('rate-limited'), { maxParagraphs: null, maxSessionParagraphs: null }],
    ['Copy failed.', new StreamSessionError('session-rotated'), { maxParagraphs: null, maxSessionParagraphs: null }],
    ['Copy failed.', new StreamSessionError('policy-refused'), { maxParagraphs: null, maxSessionParagraphs: null }],
  ]
  for (const [message, error, copyCaps] of refusals) it(`a ${(error as StreamSessionError).code} refusal says "${message}" and copies nothing`, async () => {
    const writeText = writer()
    const doc = fakeDoc({ copy: true }, { copy: copyCaps, a11y: 'focused' })
    doc.copy.mockRejectedValue(error)
    const v = view(doc)
    key(scroller(v), 'a')
    key(scroller(v), 'c')
    await waitFor(() => expect(status(v).textContent).toBe(message))
    expect(writeText).not.toHaveBeenCalled()
    // A new selection clears the message.
    key(scroller(v), 'a')
    await waitFor(() => expect(status(v).textContent).toBe(''))
  })

  it('a failed clipboard write says "Copy failed."', async () => {
    Object.defineProperty(navigator, 'clipboard', { value: { writeText: async () => { throw new Error('denied') } }, configurable: true })
    // jsdom has no execCommand, so the legacy fallback fails too.
    const doc = fakeDoc({ copy: true })
    doc.copy.mockResolvedValue('text')
    const v = view(doc)
    key(scroller(v), 'a')
    key(scroller(v), 'c')
    await waitFor(() => expect(status(v).textContent).toBe('Copy failed.'))
  })

  it('Ctrl+A and Ctrl+C in the search input are left to the input', async () => {
    const doc = fakeDoc({ copy: true })
    const v = view(doc)
    key(scroller(v), 'f')
    const input = v.getByLabelText('Search document')
    key(input, 'a')
    key(input, 'c')
    await act(async () => {})
    expect(doc.select).not.toHaveBeenCalled()
    expect(doc.copy).not.toHaveBeenCalled()
  })

  it('without can.copy, Ctrl+A selects nothing and pointer drags draw nothing', async () => {
    const doc = fakeDoc()
    const v = view(doc)
    key(scroller(v), 'a')
    const wrap = v.container.querySelector('[data-unit-wrap="0"]')!
    fireEvent.pointerDown(wrap, { button: 0, clientX: 1, clientY: 1 })
    fireEvent.pointerMove(window, { clientX: 50, clientY: 50 })
    fireEvent.pointerUp(window, { clientX: 50, clientY: 50 })
    await act(async () => {})
    expect(doc.select).not.toHaveBeenCalled()
    expect(v.container.querySelectorAll('[data-sel]').length).toBe(0)
  })
})

const deferred = <T,>() => { let ok!: (v: T) => void, no!: (e: unknown) => void; const p = new Promise<T>((a, b) => { ok = a; no = b }); return { p, ok, no } }
const flush = () => act(async () => { for (let i = 0; i < 5; i++) await Promise.resolve() })
function clip() {
  const writeText = vi.fn(async (_t: string) => {})
  Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
  return writeText
}
/** A select that honours its AbortSignal, as the real session does. */
const honest = (doc: ReturnType<typeof fakeDoc>) => doc.select.mockImplementation(async (_r, signal) => {
  if (signal?.aborted) throw new DOMException('aborted', 'AbortError')
  return { units: [{ unit: 0, rects: [rect(0, 0)] }], paragraphs: 1 }
})

describe('fix round 1: lifetimes', () => {
  it('StrictMode: Ctrl+A draws the selection (the select signal is live after the mount replay)', async () => {
    const doc = fakeDoc({ copy: true })
    honest(doc)
    const v = render(<StrictMode><RemoteDocView document={doc as unknown as RemoteOfficeDoc} /></StrictMode>)
    key(scroller(v), 'a')
    expect(doc.select.mock.calls[0]![1]!.aborted).toBe(false)
    await waitFor(() => expect(v.container.querySelectorAll('[data-sel]').length).toBe(1))
  })
  it('StrictMode: a search runs on a live signal and draws its hits', async () => {
    const doc = fakeDoc()
    doc.search.mockImplementation(async (_q, o) => {
      if (o?.signal?.aborted) throw new DOMException('aborted', 'AbortError')
      o?.onHits?.([{ unit: 0, rects: [rect(0, 0)] }]); return { total: 1, truncated: false }
    })
    const v = render(<StrictMode><RemoteDocView document={doc as unknown as RemoteOfficeDoc} /></StrictMode>)
    key(scroller(v), 'f')
    fireEvent.change(v.getByLabelText('Search document'), { target: { value: 'x' } })
    await waitFor(() => expect(v.container.querySelectorAll('[data-hit]').length).toBe(1))
  })
  it('a copy that resolves after unmount writes nothing to the clipboard', async () => {
    const writeText = clip(), doc = fakeDoc({ copy: true }), d = deferred<string>()
    doc.copy.mockReturnValue(d.p)
    const v = view(doc)
    key(scroller(v), 'a'); key(scroller(v), 'c')
    expect(doc.copy).toHaveBeenCalledTimes(1)
    v.unmount()
    d.ok('late text')
    await flush()
    expect(writeText).not.toHaveBeenCalled()
  })
  it('a copy that resolves after the document is replaced writes nothing and says nothing', async () => {
    const writeText = clip(), doc = fakeDoc({ copy: true }), next = fakeDoc({ copy: true }), d = deferred<string>()
    doc.copy.mockReturnValue(d.p)
    const v = view(doc)
    key(scroller(v), 'a'); key(scroller(v), 'c')
    v.rerender(<RemoteDocView document={next as unknown as RemoteOfficeDoc} />)
    d.ok('late text')
    await flush()
    expect(writeText).not.toHaveBeenCalled()
    expect(status(v).textContent).toBe('')
    // The replaced document's selection is gone: Ctrl+C on the new one copies nothing.
    key(scroller(v), 'c')
    expect(next.copy).not.toHaveBeenCalled()
  })
  it('a late refusal does not restore a message cleared by a newer selection', async () => {
    const doc = fakeDoc({ copy: true }, { copy: { maxParagraphs: 3, maxSessionParagraphs: null }, a11y: 'focused' }), d = deferred<string>()
    doc.copy.mockReturnValue(d.p)
    const v = view(doc)
    key(scroller(v), 'a'); key(scroller(v), 'c')
    key(scroller(v), 'a')
    d.no(new StreamSessionError('copy-too-large', 'x', { limit: 3, requested: 4 }))
    await flush()
    expect(status(v).textContent).toBe('')
  })
  it('a late success after a newer selection does not write the clipboard', async () => {
    const writeText = clip(), doc = fakeDoc({ copy: true }), d = deferred<string>()
    doc.copy.mockReturnValue(d.p)
    const v = view(doc)
    key(scroller(v), 'a'); key(scroller(v), 'c')
    key(scroller(v), 'a')
    d.ok('old text')
    await flush()
    expect(writeText).not.toHaveBeenCalled()
  })
  it('a second Ctrl+C while a copy is in flight sends no second request', async () => {
    const writeText = clip(), doc = fakeDoc({ copy: true }), d = deferred<string>()
    doc.copy.mockReturnValueOnce(d.p).mockResolvedValueOnce('again')
    const v = view(doc)
    key(scroller(v), 'a'); key(scroller(v), 'c'); key(scroller(v), 'c')
    expect(doc.copy).toHaveBeenCalledTimes(1)
    d.ok('first')
    await waitFor(() => expect(writeText).toHaveBeenCalledWith('first'))
    key(scroller(v), 'c')
    await waitFor(() => expect(doc.copy).toHaveBeenCalledTimes(2))
  })
})

describe('fix round 1: the select limiter (fake clock)', () => {
  const wrap = (v: ReturnType<typeof render>) => v.container.querySelector('[data-unit-wrap="0"]')!
  const pe = (type: string, x: number, y: number) => { const e = new MouseEvent(type, { clientX: x, clientY: y, button: 0, bubbles: true }); return e }
  const times = (doc: ReturnType<typeof fakeDoc>, t0: number, log: number[]) => doc.select.mockImplementation(async () => { log.push(Date.now() - t0); return { units: [], paragraphs: 0 } })
  const focusOf = (doc: ReturnType<typeof fakeDoc>, k: number) => (doc.select.mock.calls[k]![0] as { focus: { x: number; y: number } }).focus

  it('a move then an immediate pointer-up: the final position goes as the trailing request at the next slot', () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] })
    const doc = fakeDoc({ copy: true }), log: number[] = [], t0 = Date.now()
    times(doc, t0, log)
    const v = view(doc)
    fireEvent.pointerDown(wrap(v), { button: 0, clientX: 10, clientY: 10 })
    act(() => { window.dispatchEvent(pe('pointermove', 50, 50)) })
    vi.advanceTimersByTime(1)
    act(() => { window.dispatchEvent(pe('pointerup', 60, 70)) })
    expect(log).toEqual([0])
    act(() => { vi.advanceTimersByTime(198) })
    expect(log).toEqual([0])
    act(() => { vi.advanceTimersByTime(1) })
    expect(log).toEqual([0, 200])
    expect(focusOf(doc, 1)).toMatchObject({ x: 60, y: 70 })
    act(() => { vi.advanceTimersByTime(1000) })
    expect(log.length).toBe(2)
  })

  it('during a drag an answer still draws while the next request waits for its slot', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] })
    const doc = fakeDoc({ copy: true }), d1 = deferred<{ units: RemoteSearchHit[]; paragraphs: number }>()
    doc.select.mockReturnValueOnce(d1.p)
    const v = view(doc)
    fireEvent.pointerDown(wrap(v), { button: 0, clientX: 10, clientY: 10 })
    act(() => { window.dispatchEvent(pe('pointermove', 50, 50)) })
    act(() => { window.dispatchEvent(pe('pointermove', 60, 60)) })
    expect(doc.select).toHaveBeenCalledTimes(1)
    d1.ok({ units: [{ unit: 0, rects: [rect(0, 0)] }], paragraphs: 1 })
    await flush()
    expect(v.container.querySelectorAll('[data-sel]').length).toBe(1)
  })

  it('consecutive short drags and Ctrl+A share one limiter: requests stay ≥ 200 ms apart and the last position wins', () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] })
    const doc = fakeDoc({ copy: true }), log: number[] = [], t0 = Date.now()
    times(doc, t0, log)
    const v = view(doc)
    for (const [a, b] of [[10, 40], [20, 50], [30, 90]] as const) {
      fireEvent.pointerDown(wrap(v), { button: 0, clientX: a, clientY: a })
      act(() => { window.dispatchEvent(pe('pointermove', b, b)) })
      act(() => { window.dispatchEvent(pe('pointerup', b, b)) })
      act(() => { vi.advanceTimersByTime(20) })
    }
    act(() => { vi.advanceTimersByTime(1000) })
    for (let k = 1; k < log.length; k++) expect(log[k]! - log[k - 1]!).toBeGreaterThanOrEqual(200)
    expect(focusOf(doc, log.length - 1)).toMatchObject({ x: 90, y: 90 })
    // Ctrl+A twice in a row: the second waits for its slot.
    const before = log.length
    key(scroller(v), 'a'); key(scroller(v), 'a')
    expect(log.length).toBe(before + 1)
    act(() => { vi.advanceTimersByTime(200) })
    expect(log.length).toBe(before + 2)
    for (let k = 1; k < log.length; k++) expect(log[k]! - log[k - 1]!).toBeGreaterThanOrEqual(200)
  })
})

describe('remote view: a11y text layer (Task 8)', () => {
  type P = { rect: { x: number; y: number; width: number; height: number }; text?: string }
  const r0 = rect(0, 0)
  const unitDoc = () => fakeDoc({ copy: true, textA11y: true }, { copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'unit' })
  const focusedDoc = () => fakeDoc({ copy: true, textA11y: true }, { copy: { maxParagraphs: 3, maxSessionParagraphs: 5 }, a11y: 'focused' })
  const layer = (v: ReturnType<typeof render>, i: number) => v.container.querySelector(`[data-unit-wrap="${i}"] [data-a11y-layer]`) as HTMLElement | null
  /** A controllable IntersectionObserver: `show(i)` reports unit i's canvas entering view. */
  function fakeIO() {
    const observers: Array<{ cb: IntersectionObserverCallback; els: Element[] }> = []
    class IO { els: Element[] = []; constructor(readonly cb: IntersectionObserverCallback) { observers.push(this) } observe(el: Element) { this.els.push(el) } disconnect() { this.els = [] } unobserve() {} takeRecords() { return [] } }
    vi.stubGlobal('IntersectionObserver', IO)
    return (i: number, isIntersecting = true) => act(() => {
      for (const o of observers) for (const el of o.els) if ((el as HTMLElement).dataset.unit === String(i)) o.cb([{ isIntersecting, target: el } as unknown as IntersectionObserverEntry], o as never)
    })
  }
  afterEach(() => vi.unstubAllGlobals())

  it("'off': no layer and no request", async () => {
    for (const doc of [fakeDoc(), fakeDoc({ copy: true }, { copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'off' }), fakeDoc({ textA11y: true })]) {
      const v = view(doc)
      await act(async () => {})
      expect(v.container.querySelector('[data-a11y-layer]')).toBeNull()
      expect(doc.a11y).not.toHaveBeenCalled()
      v.unmount()
    }
  })

  it("'unit': a visually-hidden group per unit, one <p> per paragraph in order, requested once per unit", async () => {
    const doc = unitDoc()
    doc.a11y.mockImplementation(async (u: number) => [{ rect: r0, text: `first ${u}` }, { rect: r0, text: `second ${u}` }] as P[])
    const v = view(doc)
    await waitFor(() => expect(layer(v, 1)?.querySelectorAll('p').length).toBe(2))
    for (const i of [0, 1]) {
      const l = layer(v, i)!
      expect(l.getAttribute('role')).toBe('group')
      expect(l.getAttribute('aria-label')).toBe(`Page ${i + 1}`)
      expect(l.hasAttribute('aria-hidden')).toBe(false)
      expect(l.hasAttribute('data-ov-remote-chrome')).toBe(true)
      expect(l.style.display).not.toBe('none')
      expect([l.style.position, l.style.width, l.style.height, l.style.overflow]).toEqual(['absolute', '1px', '1px', 'hidden'])
      expect([...l.querySelectorAll('p')].map((p) => p.textContent)).toEqual([`first ${i}`, `second ${i}`])
      expect(v.container.querySelector(`canvas[data-unit="${i}"]`)!.getAttribute('aria-hidden')).toBe('true')
    }
    expect(doc.a11y.mock.calls).toEqual([[0], [1]])
    expect(status(v).textContent).toBe('')
  })

  it("'unit': requested when the unit first becomes visible; a failure says 'Text unavailable.', leaves the layer empty and retries on the next visibility", async () => {
    const show = fakeIO()
    const doc = unitDoc()
    doc.a11y.mockRejectedValueOnce(new StreamSessionError('rate-limited'))
    doc.a11y.mockImplementation(async () => [{ rect: r0, text: 'hello' }] as P[])
    const v = view(doc)
    await act(async () => {})
    expect(doc.a11y).not.toHaveBeenCalled()
    expect(layer(v, 0)!.childElementCount).toBe(0)
    await show(0)
    await act(async () => {})
    // Painting reaches the neighbours; text only the unit in view.
    expect(doc.a11y.mock.calls).toEqual([[0]])
    expect(layer(v, 0)!.childElementCount).toBe(0)
    expect(status(v).textContent).toBe('Text unavailable.')
    await show(0, false)
    await show(0)
    await waitFor(() => expect(layer(v, 0)!.textContent).toBe('hello'))
    await show(0)
    await act(async () => {})
    expect(doc.a11y.mock.calls).toEqual([[0], [0]])
  })

  it('an initial failure on a unit that stays visible is retried once after 5 s (not in a loop); a visibility loss cancels the retry', async () => {
    vi.useFakeTimers()
    const show = fakeIO()
    const doc = unitDoc()
    doc.a11y.mockRejectedValue(new StreamSessionError('policy-refused'))
    const v = view(doc)
    await show(0)
    await act(async () => {})
    expect(doc.a11y).toHaveBeenCalledTimes(1)
    await act(async () => { vi.advanceTimersByTime(A11Y_RETRY_MS - 1) })
    expect(doc.a11y).toHaveBeenCalledTimes(1)
    await act(async () => { vi.advanceTimersByTime(1) })
    expect(doc.a11y).toHaveBeenCalledTimes(2)
    await act(async () => { vi.advanceTimersByTime(4 * A11Y_RETRY_MS) })
    expect(doc.a11y).toHaveBeenCalledTimes(2)
    expect(status(v).textContent).toBe('Text unavailable.')
    // Leaving view before the timer fires cancels it; coming back retries (and re-arms one timer).
    await show(0, false)
    await show(0)
    await act(async () => {})
    expect(doc.a11y).toHaveBeenCalledTimes(3)
    await show(0, false)
    await act(async () => { vi.advanceTimersByTime(2 * A11Y_RETRY_MS) })
    expect(doc.a11y).toHaveBeenCalledTimes(3)
    // Unmount clears a pending timer.
    await show(0)
    await act(async () => {})
    expect(doc.a11y).toHaveBeenCalledTimes(4)
    v.unmount()
    await act(async () => { vi.advanceTimersByTime(2 * A11Y_RETRY_MS) })
    expect(doc.a11y).toHaveBeenCalledTimes(4)
  })

  it("'focused': an initial (rects) failure is reported too; copy-quota-exhausted states the quota", async () => {
    for (const [e, msg] of [[new StreamSessionError('rate-limited'), 'Text unavailable.'], [new StreamSessionError('copy-quota-exhausted', 'x', { limit: 5, requested: 1 }), 'Copy limit for this session reached (5 paragraphs).']] as const) {
      const doc = focusedDoc()
      doc.a11y.mockRejectedValue(e)
      const v = view(doc)
      await waitFor(() => expect(status(v).textContent).toBe(msg))
      expect(layer(v, 0)!.childElementCount).toBe(0)
      v.unmount()
    }
  })

  it("'focused': n unread paragraphs are named buttons; activating one reads it once, replaces it by a focused <p> and caches it", async () => {
    const doc = focusedDoc()
    doc.a11y.mockImplementation(async (_u: number, k?: number) => [0, 1, 2].map((j) => (k === j ? { rect: r0, text: `text ${j}` } : { rect: r0 })) as P[])
    const v = view(doc)
    await waitFor(() => expect(layer(v, 0)?.querySelectorAll('button').length).toBe(3))
    const buttons = () => [...layer(v, 0)!.querySelectorAll('button')]
    expect(buttons().map((b) => [b.getAttribute('type'), b.textContent, b.hasAttribute('aria-label')])).toEqual([1, 2, 3].map((k) => ['button', `Paragraph ${k} of 3. Press Enter to read.`, false]))
    expect(layer(v, 0)!.querySelectorAll('p').length).toBe(0)
    expect(doc.a11y.mock.calls).toEqual([[0], [1]])
    const second = buttons()[1]!
    second.focus()
    fireEvent.click(second)
    fireEvent.click(second) // a second activation before the answer sends nothing
    await waitFor(() => expect(layer(v, 0)!.querySelector('p')?.textContent).toBe('text 1'))
    const p = layer(v, 0)!.querySelector('p')!
    expect(p.getAttribute('tabindex')).toBe('0')
    expect(p.hasAttribute('aria-label')).toBe(false)
    expect(document.activeElement).toBe(p)
    // Order kept: button, p, button.
    expect([...layer(v, 0)!.children].map((c) => c.tagName)).toEqual(['BUTTON', 'P', 'BUTTON'])
    fireEvent.keyDown(p, { key: 'Enter' })
    fireEvent.click(buttons()[1]!)
    await waitFor(() => expect(layer(v, 0)!.querySelectorAll('p').length).toBe(2))
    expect(doc.a11y.mock.calls).toEqual([[0], [1], [0, 1], [0, 2]])
    expect(status(v).textContent).toBe('')
  })

  it("'focused': a read refused for the quota shows the copy-limit message; other failures 'Text unavailable.'; a failed read may be retried", async () => {
    const doc = focusedDoc()
    doc.a11y.mockImplementation(async (_u: number, k?: number) => {
      if (k === 0) throw new StreamSessionError('copy-quota-exhausted', 'x', { limit: 5, requested: 1 })
      if (k === 1) throw new StreamSessionError('policy-refused')
      return [{ rect: r0 }, { rect: r0 }] as P[]
    })
    const v = view(doc)
    await waitFor(() => expect(layer(v, 0)?.querySelectorAll('button').length).toBe(2))
    const buttons = () => [...layer(v, 0)!.querySelectorAll('button')]
    fireEvent.click(buttons()[0]!)
    await waitFor(() => expect(status(v).textContent).toBe('Copy limit for this session reached (5 paragraphs).'))
    fireEvent.click(buttons()[1]!)
    await waitFor(() => expect(status(v).textContent).toBe('Text unavailable.'))
    expect(buttons().length).toBe(2)
    fireEvent.click(buttons()[1]!)
    await waitFor(() => expect(doc.a11y.mock.calls.filter((c) => c[1] === 1).length).toBe(2))
  })

  it('a resume that turns a11y off removes the layer and its text at once; a late answer is dropped', async () => {
    const doc = unitDoc()
    let late!: (p: P[]) => void
    doc.a11y.mockImplementationOnce(async () => [{ rect: r0, text: 'secret 0' }] as P[])
    doc.a11y.mockImplementationOnce(() => new Promise<P[]>((r) => { late = r }))
    const v = view(doc)
    await waitFor(() => expect(layer(v, 0)?.textContent).toBe('secret 0'))
    doc.resume({ textA11y: false, copy: false }, { copy: null, a11y: 'off' })
    expect(v.container.querySelector('[data-a11y-layer]')).toBeNull()
    expect(v.container.textContent).not.toContain('secret')
    await act(async () => { late([{ rect: r0, text: 'secret 1' }]) })
    expect(v.container.textContent).not.toContain('secret')
    expect(v.getAllByText(TEXT_UNAVAILABLE).length).toBe(1)
  })

  it('unit → focused → unit on the same document: the layer is rebuilt in the new mode and cached texts are dropped', async () => {
    const doc = unitDoc()
    doc.a11y.mockImplementation(async (_u: number, k?: number) => {
      const mode = doc.manifest.caps.a11y
      return [0, 1].map((j) => (mode === 'unit' || k === j ? { rect: r0, text: `${mode} ${j}` } : { rect: r0 })) as P[]
    })
    const v = view(doc)
    await waitFor(() => expect(layer(v, 0)?.textContent).toBe('unit 0unit 1'))
    doc.resume({}, { copy: { maxParagraphs: 3, maxSessionParagraphs: 5 }, a11y: 'focused' })
    expect(layer(v, 0)!.textContent).not.toContain('unit')
    await waitFor(() => expect(layer(v, 0)?.querySelectorAll('button').length).toBe(2))
    fireEvent.click(layer(v, 0)!.querySelector('button')!)
    await waitFor(() => expect(layer(v, 0)!.querySelector('p')?.textContent).toBe('focused 0'))
    doc.resume({}, { copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'unit' })
    expect(layer(v, 0)!.textContent).not.toContain('focused')
    await waitFor(() => expect(layer(v, 0)?.textContent).toBe('unit 0unit 1'))
    doc.resume({}, { copy: { maxParagraphs: 3, maxSessionParagraphs: 5 }, a11y: 'focused' })
    await waitFor(() => expect(layer(v, 0)?.querySelectorAll('button').length).toBe(2))
    // The paragraph read in the earlier focused period is not served from a stale cache.
    expect(layer(v, 0)!.querySelector('p')).toBeNull()
  })

  it('a resume that turns copy off clears the selection and detaches selection and copy; turning it on re-attaches', async () => {
    const writeText = vi.fn(async (_t: string) => {})
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
    const doc = fakeDoc({ copy: true }, { copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'off' })
    doc.select.mockResolvedValue({ units: [{ unit: 0, rects: [rect(0, 0, 200, 10)] }], paragraphs: 1 })
    doc.copy.mockResolvedValue('x')
    const v = view(doc)
    key(scroller(v), 'a')
    await waitFor(() => expect(v.container.querySelectorAll('[data-sel]').length).toBe(1))
    doc.resume({ copy: false }, { copy: null, a11y: 'off' })
    expect(v.container.querySelectorAll('[data-sel]').length).toBe(0)
    key(scroller(v), 'c'); key(scroller(v), 'a')
    fireEvent.pointerDown(v.container.querySelector('[data-unit-wrap="0"]')!, { button: 0, clientX: 1, clientY: 1 })
    await act(async () => {})
    expect(doc.copy).not.toHaveBeenCalled()
    expect(doc.select).toHaveBeenCalledTimes(1)
    expect(v.getAllByText(TEXT_UNAVAILABLE).length).toBe(1)
    doc.resume({ copy: true }, { copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'off' })
    key(scroller(v), 'a')
    await waitFor(() => expect(v.container.querySelectorAll('[data-sel]').length).toBe(1))
    key(scroller(v), 'c')
    await waitFor(() => expect(writeText).toHaveBeenCalledWith('x'))
  })

  it('stale answers are dropped: unmount, and a document change', async () => {
    let resolve!: (p: P[]) => void
    const doc = unitDoc()
    doc.a11y.mockImplementation(() => new Promise<P[]>((r) => { resolve = r }))
    const err = vi.spyOn(console, 'error').mockImplementation(() => {})
    const v = view(doc)
    await act(async () => {})
    v.unmount()
    await act(async () => { resolve([{ rect: r0, text: 'late' }]) })
    expect(err).not.toHaveBeenCalled()
    // A document change: the old document's answer never fills the new one's layer.
    const a = unitDoc(), b = unitDoc()
    const pending: Array<(p: P[]) => void> = []
    a.a11y.mockImplementation(() => new Promise<P[]>((r) => { pending.push(r) }))
    b.a11y.mockImplementation(() => new Promise<P[]>(() => {}))
    const w = render(<RemoteDocView document={a as unknown as RemoteOfficeDoc} />)
    await act(async () => {})
    w.rerender(<RemoteDocView document={b as unknown as RemoteOfficeDoc} />)
    await act(async () => { for (const r of pending) r([{ rect: r0, text: 'old document' }]) })
    expect(w.container.textContent).not.toContain('old document')
  })

  it('StrictMode: the layer still fills', async () => {
    const doc = unitDoc()
    doc.a11y.mockImplementation(async () => [{ rect: r0, text: 'strict' }] as P[])
    const v = render(<StrictMode><RemoteDocView document={doc as unknown as RemoteOfficeDoc} /></StrictMode>)
    await waitFor(() => expect(layer(v, 0)?.textContent).toBe('strict'))
  })
})
