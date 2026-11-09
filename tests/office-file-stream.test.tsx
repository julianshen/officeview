/** 1d Task 4: a StreamSource routes to the lazy remote viewer (§6.3, §8.2). */
import { afterEach, describe, expect, it, vi, type Mock } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import { createRef } from 'react'
import { OfficeFile, useOfficeFile, type OfficeFileState } from '../src/components/OfficeFile'
import type { OfficeViewerHandle } from '../src/components/present'
import { StreamSessionError } from '../src/stream/remote/errors'
import type { RemoteOfficeDoc } from '../src/stream/remote/remote-doc'

const h = vi.hoisted(() => ({ open: vi.fn() }))
vi.mock('../src/stream/remote/open', () => ({ openRemoteDocument: (...a: unknown[]) => h.open(...a) }))

afterEach(() => { cleanup(); vi.restoreAllMocks() })

function fakeDoc(name = 'deck.pptx'): RemoteOfficeDoc & { close: Mock; paint: Mock } {
  const manifest = {
    name, kind: 'pptx', units: { count: 2, list: [{ w: 100, h: 50 }, { w: 100, h: 50 }] },
    can: { search: false, copy: false, print: true, textA11y: false, present: false, notes: false, links: false },
    recorder: { scalesBuckets: [1, 2] }, notices: [],
  }
  return {
    kind: 'remote', manifest: manifest as never,
    units: async () => manifest.units.list.map((u) => ({ widthPx: u.w, heightPx: u.h })),
    paint: vi.fn(async () => {}),
    search: () => Promise.reject(new StreamSessionError('policy-refused')),
    copy: () => Promise.reject(new StreamSessionError('policy-refused')),
    close: vi.fn(),
    ended: new Promise(() => {}),
  } as unknown as RemoteOfficeDoc & { close: Mock; paint: Mock }
}
const stream = (url = 'wss://x/stream') => ({ kind: 'stream' as const, url })
const deferred = <T,>() => { let ok!: (v: T) => void; const p = new Promise<T>((r) => { ok = r }); return { p, ok } }

describe('<OfficeFile data={StreamSource}>', () => {
  it('opens a remote document once, never fetches, and renders the remote viewer', async () => {
    const doc = fakeDoc(), fetchSpy = vi.spyOn(globalThis, 'fetch')
    const src = stream()
    // Set before the first import; the mock factory assigns h.open on import.
    h.open.mockReset(); h.open.mockResolvedValueOnce(doc)
    const view = render(<OfficeFile data={src} watermark={{ text: 'HOSTMARK' }} />)
    await waitFor(() => expect(view.container.querySelectorAll('canvas').length).toBe(2))
    expect(h.open).toHaveBeenCalledTimes(1)
    expect(h.open.mock.calls[0]![0]).toBe(src)
    expect(fetchSpy).not.toHaveBeenCalled()
    expect(view.getByRole('document')).toBeTruthy()
    const notice = view.container.querySelectorAll('[data-officeview-notice]')
    expect([...notice].filter((n) => n.textContent?.includes('text unavailable in this viewing mode')).length).toBe(1)
    expect(view.container.querySelector('input')).toBeNull()
    expect(view.container.textContent).not.toContain('HOSTMARK')
    await waitFor(() => expect(doc.paint).toHaveBeenCalled())
  })
  it('replacing the source mid-open closes the first document exactly once; unmount closes the second', async () => {
    const first = fakeDoc('a'), second = fakeDoc('b'), d1 = deferred<RemoteOfficeDoc>()
    h.open.mockReset()
    h.open.mockReturnValueOnce(d1.p).mockResolvedValueOnce(second)
    const view = render(<OfficeFile data={stream('wss://a')} />)
    await waitFor(() => expect(h.open).toHaveBeenCalledTimes(1))
    view.rerender(<OfficeFile data={stream('wss://b')} />)
    await waitFor(() => expect(h.open).toHaveBeenCalledTimes(2))
    d1.ok(first)
    await waitFor(() => expect(first.close).toHaveBeenCalledTimes(1))
    await waitFor(() => expect(view.getByRole('document').getAttribute('aria-label')).toBe('b'))
    view.unmount()
    expect(second.close).toHaveBeenCalledTimes(1)
    expect(first.close).toHaveBeenCalledTimes(1)
  })
  it('present() on a remote document rejects not-presentable', async () => {
    h.open.mockReset(); h.open.mockResolvedValueOnce(fakeDoc())
    const ref = createRef<OfficeViewerHandle>()
    const view = render(<OfficeFile ref={ref} data={stream()} />)
    await waitFor(() => expect(view.container.querySelectorAll('canvas').length).toBe(2))
    await expect(ref.current!.present()).rejects.toMatchObject({ code: 'not-presentable' })
  })
  it('a typed session error becomes an error state carrying its code', async () => {
    h.open.mockReset(); h.open.mockRejectedValueOnce(new StreamSessionError('auth-failed'))
    let last: OfficeFileState | undefined
    const src = stream()
    const Probe = () => { last = useOfficeFile(src); return null }
    render(<Probe />)
    await waitFor(() => expect(last?.status).toBe('error'))
    expect(last).toMatchObject({ status: 'error', code: 'auth-failed' })
  })
  it('a session that ends after the view is shown surfaces its typed code once', async () => {
    const doc = fakeDoc()
    let end!: (e: StreamSessionError) => void
    ;(doc as { ended: Promise<StreamSessionError> }).ended = new Promise((r) => { end = r })
    h.open.mockReset(); h.open.mockResolvedValueOnce(doc)
    const view = render(<OfficeFile data={stream()} />)
    await waitFor(() => expect(view.container.querySelectorAll('canvas').length).toBe(2))
    end(new StreamSessionError('revoked'))
    await waitFor(() => expect(view.container.querySelector('[data-stream-error="revoked"]')).toBeTruthy())
    expect(view.getAllByRole('alert').length).toBe(1)
  })
  it('review 6: denied print marks the themed instance root the print guard matches', async () => {
    const doc = fakeDoc()
    ;(doc.manifest.can as { print: boolean }).print = false
    h.open.mockReset(); h.open.mockResolvedValueOnce(doc)
    const view = render(<OfficeFile data={stream()} />)
    await waitFor(() => expect(view.container.querySelectorAll('canvas').length).toBe(2))
    expect(view.container.querySelector('[data-ov-theme-root][data-ov-instance][data-ov-print="false"]')).toBeTruthy()
  })
  it('PR #64: printing stays guarded until every unit is painted; beforeprint paints the off-screen ones', async () => {
    // An IntersectionObserver that never fires: no unit nears the viewport.
    vi.stubGlobal('IntersectionObserver', class { observe() {} disconnect() {} })
    try {
      const doc = fakeDoc()
      h.open.mockReset(); h.open.mockResolvedValueOnce(doc)
      const view = render(<OfficeFile data={stream()} />)
      await waitFor(() => expect(view.container.querySelectorAll('canvas').length).toBe(2))
      const root = () => view.container.querySelector('[data-ov-theme-root][data-ov-instance]')!
      expect(doc.paint).not.toHaveBeenCalled()
      expect(root().getAttribute('data-ov-print')).toBe('false')
      window.dispatchEvent(new Event('beforeprint'))
      await waitFor(() => expect(doc.paint).toHaveBeenCalledTimes(2))
      await waitFor(() => expect(root().getAttribute('data-ov-print')).toBe('true'))
    } finally { vi.unstubAllGlobals() }
  })
  it('PR #64: Ctrl/Cmd+P paints every unit, then prints', async () => {
    vi.stubGlobal('IntersectionObserver', class { observe() {} disconnect() {} })
    const print = vi.spyOn(window, 'print').mockImplementation(() => {})
    try {
      const doc = fakeDoc()
      h.open.mockReset(); h.open.mockResolvedValueOnce(doc)
      const view = render(<OfficeFile data={stream()} />)
      await waitFor(() => expect(view.container.querySelectorAll('canvas').length).toBe(2))
      const root = view.container.querySelector('[data-ov-theme-root][data-ov-instance]')!
      root.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true }))
      await waitFor(() => expect(print).toHaveBeenCalledTimes(1))
      expect(doc.paint).toHaveBeenCalledTimes(2)
      expect(root.getAttribute('data-ov-print')).toBe('true')
    } finally { vi.unstubAllGlobals() }
  })
})
