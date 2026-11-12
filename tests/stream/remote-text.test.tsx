/** Stage 3a Task 7: search bar, hit/selection overlay and copy in the remote view. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import RemoteDocView, { TEXT_UNAVAILABLE } from '../../src/components/remote/RemoteDocView'
import { StreamSessionError } from '../../src/stream/remote/errors'
import type { RemoteOfficeDoc, RemoteRange, RemoteSearchHit } from '../../src/stream/remote/remote-doc'

afterEach(() => { cleanup(); vi.restoreAllMocks() })

type Can = { search?: boolean; copy?: boolean; textA11y?: boolean }
function fakeDoc(can: Can = {}, caps: RemoteOfficeDoc['manifest']['caps'] = { copy: null, a11y: 'off' }) {
  const manifest = {
    name: 'deck.pptx', kind: 'pptx', units: { count: 2, list: [{ w: 200, h: 100 }, { w: 200, h: 100 }] },
    can: { search: true, copy: false, print: true, textA11y: false, present: false, notes: false, links: false, ...can },
    caps, recorder: { scalesBuckets: [1, 2] }, notices: [],
  }
  const doc = {
    kind: 'remote', manifest,
    units: async () => manifest.units.list.map((u) => ({ widthPx: u.w, heightPx: u.h })),
    paint: vi.fn(async () => {}),
    search: vi.fn(async (_q: string, _o?: { signal?: AbortSignal; onHits?: (h: RemoteSearchHit[]) => void }) => ({ total: 0, truncated: false })),
    select: vi.fn(async (_r: RemoteRange) => ({ units: [] as RemoteSearchHit[], paragraphs: 0 })),
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
