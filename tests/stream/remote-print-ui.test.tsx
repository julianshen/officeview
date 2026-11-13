/** Stage 3b Task 7: the remote view prints only server-drawn print rasters. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import { StrictMode } from 'react'
import RemoteDocView, { remotePrintGuard } from '../../src/components/remote/RemoteDocView'
import { StreamSessionError } from '../../src/stream/remote/errors'
import type { PrintedPage, RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'

const pr = vi.hoisted(() => ({ printRasters: vi.fn(async (..._a: unknown[]) => null) }))
vi.mock('../../src/components/paged/printDocument', () => ({ printRasters: pr.printRasters }))

type PrintCaps = RemoteOfficeDoc['manifest']['caps']['print']
type PrintOpts = { signal?: AbortSignal; onProgress?: (done: number, total: number) => void }

function fakeDoc(o: { print?: boolean; caps?: PrintCaps; count?: number } = {}) {
  const count = o.count ?? 3
  const list = Array.from({ length: count }, () => ({ w: 200, h: 100 }))
  const manifest = {
    name: 'deck.pptx', kind: 'pptx', units: { count, list },
    can: { search: true, copy: false, print: o.print ?? true, textA11y: false, present: false, notes: false, links: false },
    caps: { copy: null, a11y: 'off', print: o.caps === undefined ? { maxPages: null, maxSessionPages: null } : o.caps },
    recorder: { scalesBuckets: [1, 2] }, notices: [],
  }
  const listeners = new Set<() => void>()
  const pages = (from: number, to: number): PrintedPage[] =>
    Array.from({ length: to - from + 1 }, (_, i) => ({ unitIndex: from + i, widthPx: 200, heightPx: 100, scale: 2 as const, png: new Blob(['png']) }))
  const doc = {
    kind: 'remote', manifest,
    onManifestChange: vi.fn((cb: () => void) => { listeners.add(cb); return () => { listeners.delete(cb) } }),
    resume(m: { print?: boolean; caps?: PrintCaps }) {
      doc.manifest = {
        ...doc.manifest,
        can: { ...doc.manifest.can, ...(m.print === undefined ? {} : { print: m.print }) },
        caps: { ...doc.manifest.caps, ...(m.caps === undefined ? {} : { print: m.caps }) },
      }
      act(() => { for (const l of [...listeners]) l() })
    },
    units: async () => list.map((u) => ({ widthPx: u.w, heightPx: u.h })),
    paint: vi.fn(async (_i: number, _s: number, _d: number, c: HTMLCanvasElement) => { void c }),
    search: vi.fn(async () => ({ total: 0, truncated: false })),
    select: vi.fn(async () => ({ units: [], paragraphs: 0 })),
    copy: vi.fn(async () => ''),
    a11y: vi.fn(),
    print: vi.fn(async (r: { from: number; to: number }, po?: PrintOpts) => {
      po?.onProgress?.(r.to - r.from + 1, r.to - r.from + 1)
      return pages(r.from, r.to)
    }),
    close: vi.fn(),
    ended: new Promise(() => {}),
  }
  return doc
}
type Fake = ReturnType<typeof fakeDoc>
const view = (doc: Fake, props: { allowPrint?: boolean; title?: string } = {}) =>
  render(<RemoteDocView document={doc as unknown as RemoteOfficeDoc} {...props} />)
const themeRoot = (v: ReturnType<typeof render>) => v.container.querySelector('[data-ov-theme-root][data-ov-instance]') as HTMLElement
const printBtn = (v: ReturnType<typeof render>) => v.container.querySelector<HTMLButtonElement>('button[data-ov-print-button]')
const dialog = (v: ReturnType<typeof render>) => v.queryByRole('dialog', { name: 'Print pages' })

let winPrint: ReturnType<typeof vi.fn>
beforeEach(() => {
  pr.printRasters.mockReset(); pr.printRasters.mockResolvedValue(null)
  winPrint = vi.fn()
  vi.spyOn(window, 'print').mockImplementation(winPrint as () => void)
})
afterEach(() => {
  // The screen canvases are never printed, by any path.
  expect(winPrint).not.toHaveBeenCalled()
  cleanup(); vi.restoreAllMocks()
})

describe('remote print: button', () => {
  it('uncapped: Print requests the whole document with no dialog, then prints the rasters', async () => {
    const doc = fakeDoc({ caps: { maxPages: null, maxSessionPages: null } })
    const v = view(doc, { title: 'Quarterly' })
    const btn = printBtn(v)!
    expect(btn.closest('[data-ov-remote-chrome]')).toBeTruthy()
    fireEvent.click(btn)
    expect(dialog(v)).toBeNull()
    await waitFor(() => expect(pr.printRasters).toHaveBeenCalledTimes(1))
    expect(doc.print).toHaveBeenCalledTimes(1)
    expect(doc.print.mock.calls[0]![0]).toEqual({ from: 0, to: 2 })
    const [pages, root, opts] = pr.printRasters.mock.calls[0]! as [PrintedPage[], HTMLElement, { signal: AbortSignal; title?: string }]
    expect(pages.map((p) => p.unitIndex)).toEqual([0, 1, 2])
    // Hosted outside the view's root, which is hidden in print media.
    expect(root).toBe(document.body)
    expect(themeRoot(v).contains(root)).toBe(false)
    expect(opts.title).toBe('Quarterly')
    expect(opts.signal).toBe(doc.print.mock.calls[0]![1]!.signal)
    // The progress line clears once the job is handed to the print surface.
    await waitFor(() => expect(v.container.textContent).not.toContain('Preparing pages'))
  })
  it('is absent when the host disallows print or the server denies it', () => {
    const a = view(fakeDoc(), { allowPrint: false })
    expect(printBtn(a)).toBeNull()
    cleanup()
    const b = view(fakeDoc({ print: false }))
    expect(printBtn(b)).toBeNull()
  })
  it('a second click while a job runs is ignored', async () => {
    const doc = fakeDoc()
    doc.print.mockImplementation(() => new Promise(() => {}))
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(doc.print).toHaveBeenCalledTimes(1))
    expect((printBtn(v) as HTMLButtonElement).disabled).toBe(true)
    fireEvent.click(printBtn(v)!)
    fireEvent.keyDown(themeRoot(v), { key: 'p', ctrlKey: true })
    expect(doc.print).toHaveBeenCalledTimes(1)
  })
  it('works under StrictMode', async () => {
    const doc = fakeDoc()
    const v = render(<StrictMode><RemoteDocView document={doc as unknown as RemoteOfficeDoc} /></StrictMode>)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(pr.printRasters).toHaveBeenCalledTimes(1))
    expect(doc.print).toHaveBeenCalledTimes(1)
  })
})

describe('remote print: range dialog under a cap', () => {
  it('opens with defaults 1 and min(count, maxPages), shows both cap lines, and submits 0-based', async () => {
    const doc = fakeDoc({ caps: { maxPages: 2, maxSessionPages: 7 } })
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    const d = dialog(v)!
    expect(d.getAttribute('aria-modal')).toBe('true')
    expect(d.closest('[data-ov-remote-chrome]')).toBeTruthy()
    const from = v.getByLabelText('From') as HTMLInputElement, to = v.getByLabelText('To') as HTMLInputElement
    expect([from.value, to.value]).toEqual(['1', '2'])
    expect([from.min, from.max, to.min, to.max]).toEqual(['1', '3', '1', '3'])
    expect(d.textContent).toContain('Up to 2 pages per print.')
    expect(d.textContent).toContain('Up to 7 pages in this session.')
    expect(doc.print).not.toHaveBeenCalled()
    fireEvent.click(within(d).getByRole('button', { name: 'Print' }))
    await waitFor(() => expect(doc.print).toHaveBeenCalledTimes(1))
    expect(doc.print.mock.calls[0]![0]).toEqual({ from: 0, to: 1 })
    expect(dialog(v)).toBeNull()
  })
  it('defaults To to the count when maxPages is larger; only the session line shows when maxPages is null', () => {
    const a = view(fakeDoc({ caps: { maxPages: 5, maxSessionPages: null } }))
    fireEvent.click(printBtn(a)!)
    expect((a.getByLabelText('To') as HTMLInputElement).value).toBe('3')
    expect(dialog(a)!.textContent).toContain('Up to 5 pages per print.')
    expect(dialog(a)!.textContent).not.toContain('in this session')
    cleanup()
    const b = view(fakeDoc({ caps: { maxPages: null, maxSessionPages: 4 } }))
    fireEvent.click(printBtn(b)!)
    expect((b.getByLabelText('To') as HTMLInputElement).value).toBe('3')
    expect(dialog(b)!.textContent).toContain('Up to 4 pages in this session.')
    expect(dialog(b)!.textContent).not.toContain('per print')
  })
  it('From > To, empty or out-of-range input disables Print', () => {
    const v = view(fakeDoc({ caps: { maxPages: 2, maxSessionPages: null } }))
    fireEvent.click(printBtn(v)!)
    const submit = within(dialog(v)!).getByRole('button', { name: 'Print' }) as HTMLButtonElement
    const from = v.getByLabelText('From'), to = v.getByLabelText('To')
    expect(submit.disabled).toBe(false)
    fireEvent.change(from, { target: { value: '3' } }); fireEvent.change(to, { target: { value: '2' } })
    expect(submit.disabled).toBe(true)
    fireEvent.change(from, { target: { value: '' } })
    expect(submit.disabled).toBe(true)
    fireEvent.change(from, { target: { value: '1' } }); fireEvent.change(to, { target: { value: '4' } })
    expect(submit.disabled).toBe(true)
    fireEvent.change(to, { target: { value: '1.5' } })
    expect(submit.disabled).toBe(true)
    fireEvent.change(to, { target: { value: '3' } })
    expect(submit.disabled).toBe(false)
  })
  it('Cancel and Escape close the dialog without printing', () => {
    const doc = fakeDoc({ caps: { maxPages: 2, maxSessionPages: null } })
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    fireEvent.click(v.getByRole('button', { name: 'Cancel' }))
    expect(dialog(v)).toBeNull()
    fireEvent.click(printBtn(v)!)
    fireEvent.keyDown(v.getByLabelText('From'), { key: 'Escape' })
    expect(dialog(v)).toBeNull()
    expect(doc.print).not.toHaveBeenCalled()
  })
  it('a manifest change updates an open dialog; can.print turning false closes it', () => {
    const doc = fakeDoc({ caps: { maxPages: 5, maxSessionPages: null } })
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    expect(dialog(v)!.textContent).toContain('Up to 5 pages per print.')
    doc.resume({ caps: { maxPages: 2, maxSessionPages: 3 } })
    expect(dialog(v)!.textContent).toContain('Up to 2 pages per print.')
    expect(dialog(v)!.textContent).toContain('Up to 3 pages in this session.')
    doc.resume({ print: false, caps: null })
    expect(dialog(v)).toBeNull()
    expect(printBtn(v)).toBeNull()
  })
})

describe('remote print: notices', () => {
  const cases: Array<[string, StreamSessionError, string]> = [
    ['print-too-large with a limit', new StreamSessionError('print-too-large', 'x', { limit: 4, requested: 9 }), 'Too many pages to print at once (limit: 4 pages).'],
    ['print-too-large without a limit', new StreamSessionError('print-too-large'), 'This document is too large to print from the viewer.'],
    ['print-quota-exhausted', new StreamSessionError('print-quota-exhausted', 'x', { limit: 6, requested: 3 }), 'Print limit for this session reached (6 pages).'],
    ['rate-limited', new StreamSessionError('rate-limited'), 'Too many requests. Try again shortly.'],
    ['session-rotated', new StreamSessionError('session-rotated'), 'The document could not be printed. Please try again.'],
    ['policy-refused', new StreamSessionError('policy-refused'), 'The document could not be printed. Please try again.'],
    ['a plain error', new Error('boom') as StreamSessionError, 'The document could not be printed. Please try again.'],
  ]
  for (const [name, err, text] of cases) it(`${name}: "${text}"`, async () => {
    const doc = fakeDoc()
    doc.print.mockRejectedValueOnce(err)
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(v.getByText(text).closest('[data-ov-remote-chrome]')).toBeTruthy())
    expect(pr.printRasters).not.toHaveBeenCalled()
    expect(v.container.textContent).not.toContain('Preparing pages')
  })
  it('a notice from the print surface is shown', async () => {
    const doc = fakeDoc()
    pr.printRasters.mockImplementationOnce(async (...a: unknown[]) => {
      (a[2] as { onNotice(t: string): void }).onNotice('This document is too large to print from the viewer.'); return null
    })
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(v.getByText('This document is too large to print from the viewer.')).toBeTruthy())
  })
})

describe('remote print: progress and cancel', () => {
  it('shows progress; Cancel aborts the signal passed to doc.print and shows no notice', async () => {
    const doc = fakeDoc()
    let progress!: (d: number, t: number) => void
    doc.print.mockImplementation((_r, po) => new Promise((_ok, fail) => {
      progress = po!.onProgress!
      po!.signal!.addEventListener('abort', () => fail(po!.signal!.reason))
    }))
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(doc.print).toHaveBeenCalledTimes(1))
    expect(v.container.textContent).toContain('Preparing pages 0 of 3…')
    act(() => progress(2, 3))
    expect(v.container.textContent).toContain('Preparing pages 2 of 3…')
    const signal = doc.print.mock.calls[0]![1]!.signal!
    expect(signal.aborted).toBe(false)
    fireEvent.click(v.getByRole('button', { name: 'Cancel' }))
    expect(signal.aborted).toBe(true)
    await waitFor(() => expect(v.container.textContent).not.toContain('Preparing pages'))
    expect(v.container.textContent).not.toContain('could not be printed')
    expect(pr.printRasters).not.toHaveBeenCalled()
    expect((printBtn(v) as HTMLButtonElement).disabled).toBe(false)
  })
  it('unmount after the hand-off still aborts the print surface (its frame is disposed)', async () => {
    const doc = fakeDoc()
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(pr.printRasters).toHaveBeenCalledTimes(1))
    await waitFor(() => expect(v.container.textContent).not.toContain('Preparing pages'))
    const { signal } = pr.printRasters.mock.calls[0]![2] as { signal: AbortSignal }
    expect(signal.aborted).toBe(false)
    v.unmount()
    expect(signal.aborted).toBe(true)
  })
  it('unmount aborts a running job', async () => {
    const doc = fakeDoc()
    doc.print.mockImplementation(() => new Promise(() => {}))
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(doc.print).toHaveBeenCalledTimes(1))
    v.unmount()
    expect(doc.print.mock.calls[0]![1]!.signal!.aborted).toBe(true)
  })
  it('losing print permission mid-job aborts it silently', async () => {
    const doc = fakeDoc()
    doc.print.mockImplementation(() => new Promise(() => {}))
    const v = view(doc)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(doc.print).toHaveBeenCalledTimes(1))
    doc.resume({ print: false, caps: null })
    expect(doc.print.mock.calls[0]![1]!.signal!.aborted).toBe(true)
    expect(v.container.textContent).not.toContain('Preparing pages')
  })
})

describe('remote print: document replacement', () => {
  const asDoc = (d: Fake) => d as unknown as RemoteOfficeDoc
  it('a pending job on A leaves nothing behind on B; A settling late changes nothing', async () => {
    const a = fakeDoc(), b = fakeDoc({ count: 2 })
    let settle!: (p: PrintedPage[]) => void, progress!: (d: number, t: number) => void
    a.print.mockImplementation((_r, po) => new Promise((ok) => { settle = ok; progress = po!.onProgress! }))
    const v = render(<RemoteDocView document={asDoc(a)} />)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(a.print).toHaveBeenCalledTimes(1))
    expect(v.container.textContent).toContain('Preparing pages 0 of 3…')
    v.rerender(<RemoteDocView document={asDoc(b)} />)
    expect(a.print.mock.calls[0]![1]!.signal!.aborted).toBe(true)
    expect(v.container.textContent).not.toContain('Preparing pages')
    expect(printBtn(v)!.disabled).toBe(false)
    await act(async () => { progress(2, 3); settle([]) })
    expect(v.container.textContent).not.toContain('Preparing pages')
    expect(pr.printRasters).not.toHaveBeenCalled()
    expect(printBtn(v)!.disabled).toBe(false)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(b.print).toHaveBeenCalledTimes(1))
    expect(b.print.mock.calls[0]![0]).toEqual({ from: 0, to: 1 })
  })
  it("an open dialog on A is closed on B (A's range never carries over)", () => {
    const a = fakeDoc({ caps: { maxPages: 2, maxSessionPages: null } }), b = fakeDoc({ count: 5, caps: { maxPages: 4, maxSessionPages: null } })
    const v = render(<RemoteDocView document={asDoc(a)} />)
    fireEvent.click(printBtn(v)!)
    fireEvent.change(v.getByLabelText('From'), { target: { value: '2' } })
    v.rerender(<RemoteDocView document={asDoc(b)} />)
    expect(dialog(v)).toBeNull()
    fireEvent.click(printBtn(v)!)
    expect([(v.getByLabelText('From') as HTMLInputElement).value, (v.getByLabelText('To') as HTMLInputElement).value]).toEqual(['1', '4'])
  })
  it("A's notice is cleared on B", async () => {
    const a = fakeDoc(), b = fakeDoc()
    a.print.mockRejectedValueOnce(new StreamSessionError('rate-limited'))
    const v = render(<RemoteDocView document={asDoc(a)} />)
    fireEvent.click(printBtn(v)!)
    await waitFor(() => expect(v.getByText('Too many requests. Try again shortly.')).toBeTruthy())
    v.rerender(<RemoteDocView document={asDoc(b)} />)
    expect(v.queryByText('Too many requests. Try again shortly.')).toBeNull()
  })
})

describe('remote print: dialog focus', () => {
  it('is modal: focus starts in From and Tab / Shift+Tab wrap inside the dialog', () => {
    const v = view(fakeDoc({ caps: { maxPages: 2, maxSessionPages: null } }))
    fireEvent.click(printBtn(v)!)
    const d = dialog(v)!
    expect(d.getAttribute('aria-modal')).toBe('true')
    const from = v.getByLabelText('From'), submit = within(d).getByRole('button', { name: 'Print' })
    expect(document.activeElement).toBe(from)
    submit.focus()
    expect(fireEvent.keyDown(submit, { key: 'Tab' })).toBe(false)
    expect(document.activeElement).toBe(from)
    expect(fireEvent.keyDown(from, { key: 'Tab', shiftKey: true })).toBe(false)
    expect(document.activeElement).toBe(submit)
  })
  it('Escape returns focus to the Print button', () => {
    const v = view(fakeDoc({ caps: { maxPages: 2, maxSessionPages: null } }))
    printBtn(v)!.focus()
    fireEvent.click(printBtn(v)!)
    fireEvent.keyDown(v.getByLabelText('From'), { key: 'Escape' })
    expect(document.activeElement).toBe(printBtn(v))
  })
})

describe('remote print: Ctrl/Cmd+P', () => {
  it('with print allowed opens the same flow', async () => {
    const doc = fakeDoc()
    const v = view(doc)
    expect(fireEvent.keyDown(themeRoot(v), { key: 'p', ctrlKey: true })).toBe(false)
    await waitFor(() => expect(doc.print).toHaveBeenCalledTimes(1))
    expect(doc.print.mock.calls[0]![0]).toEqual({ from: 0, to: 2 })
  })
  it('Cmd+P under a cap opens the dialog', () => {
    const v = view(fakeDoc({ caps: { maxPages: 2, maxSessionPages: null } }))
    expect(fireEvent.keyDown(themeRoot(v), { key: 'P', metaKey: true })).toBe(false)
    expect(dialog(v)).toBeTruthy()
  })
  it('can.print=true with caps.print=null is contradictory: print unavailable, Ctrl+P swallowed', () => {
    const doc = fakeDoc({ print: true, caps: null })
    const v = view(doc)
    expect(printBtn(v)).toBeNull()
    expect(fireEvent.keyDown(themeRoot(v), { key: 'p', ctrlKey: true })).toBe(false)
    expect(dialog(v)).toBeNull()
    expect(doc.print).not.toHaveBeenCalled()
  })
  it('with can.print=false is preventDefaulted and does nothing else; no button', () => {
    const doc = fakeDoc({ print: false, caps: null })
    const v = view(doc)
    expect(printBtn(v)).toBeNull()
    expect(fireEvent.keyDown(themeRoot(v), { key: 'p', ctrlKey: true })).toBe(false)
    expect(dialog(v)).toBeNull()
    expect(doc.print).not.toHaveBeenCalled()
  })
})

describe('remote print: the screen never prints', () => {
  it('the root carries data-ov-print="false" at all times, including after every unit has painted', async () => {
    vi.stubGlobal('IntersectionObserver', undefined)
    try {
      const doc = fakeDoc()
      doc.paint.mockImplementation(async (_i, _s, _d, c) => { void c })
      const v = view(doc)
      expect(themeRoot(v).getAttribute('data-ov-print')).toBe('false')
      await waitFor(() => expect(v.container.querySelectorAll('canvas[data-painted="1"]').length).toBe(3))
      expect(themeRoot(v).getAttribute('data-ov-print')).toBe('false')
      window.dispatchEvent(new Event('beforeprint'))
      await act(async () => {})
      expect(themeRoot(v).getAttribute('data-ov-print')).toBe('false')
    } finally { vi.unstubAllGlobals() }
  })
  it('carries its own print guard (no officeview.css loaded) that hides this root in print media', () => {
    const v = view(fakeDoc())
    const root = themeRoot(v)
    const guard = root.querySelector('style[data-ov-print-guard]') as HTMLStyleElement
    expect(guard).toBeTruthy()
    const media = guard.sheet!.cssRules[0] as CSSMediaRule
    expect(media.media.mediaText).toBe('print')
    const rule = media.cssRules[0] as CSSStyleRule
    expect(root.matches(rule.selectorText)).toBe(true)
    expect(rule.style.getPropertyValue('display')).toBe('none')
    expect(rule.style.getPropertyPriority('display')).toBe('important')
    // Scoped to this instance: a second view's root is not matched by the first guard.
    const w = view(fakeDoc())
    const other = [...document.querySelectorAll('[data-ov-theme-root][data-ov-instance]')].find((r) => r !== root)!
    expect(other.matches(rule.selectorText)).toBe(false)
    void w
  })
  it('the guard escapes quotes and backslashes in the instance id', () => {
    const css = remotePrintGuard('a"b\\c')
    expect(css).toBe('@media print{[data-ov-theme-root][data-ov-instance="a\\22 b\\5c c"]{display:none!important}}')
  })
  it('the view registers no beforeprint listener', () => {
    const add = vi.spyOn(window, 'addEventListener')
    view(fakeDoc())
    expect(add.mock.calls.filter((c) => c[0] === 'beforeprint')).toEqual([])
  })
})
