/**
 * Slideshow plan S1-17: the eager host API, the per-document show
 * coordinator and the synchronous start sequence (spec 4.1, 4.2, 5.3, 8, 9).
 *
 * jsdom with an injected PresentPlatform: nothing here waits on wall time.
 * The lazy chunk is reached through `showLoader`, which the tests defer,
 * fail or resolve with the real (stub) `runShow` handoff.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { render, act, fireEvent, cleanup, waitFor } from '@testing-library/react'
import { createRef } from 'react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficeFile } from '../src/components/OfficeFile'
import {
  acquireShowToken, showLoader, startPresent, SlideShowError,
  type OfficeViewerHandle, type PresentPlatform, type ShowStart, type SlideShowCallbacks,
} from '../src/components/present'
import * as slideshow from '../src/slideshow'
import { docxModel } from './helpers/paged-viewer'
import { buildPptx } from '../src/testdata/ooxml-builders'
import type { PptxDocument } from '../src/pptx/types'
import css from '../src/styles/officeview.css?inline'

vi.mock('../src/render/paint', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/render/paint')>()
  const { fakePaintables } = await import('./helpers/paged-viewer')
  return { ...m, getPaintables: () => Promise.resolve(fakePaintables([{ w: 400, h: 300 }, { w: 400, h: 300 }, { w: 400, h: 300 }])) }
})
vi.mock('../src/components/paged/PagedChrome', () => ({ default: () => null }))

// A caller-constructed deck: S1-19 validates slideWidthEmu/slideHeightEmu (spec 7), so the sizes are the real fields.
const slideOf = (index: number) => ({ index, widthEmu: 9144000, heightEmu: 6858000, shapes: [] })
const PPTX = { slides: [slideOf(0), slideOf(1), slideOf(2)], slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [] } as unknown as PptxDocument
const DOCX = docxModel(['not a deck'])

interface Deferred<T> { promise: Promise<T>; resolve(v: T): void; reject(e: unknown): void }
function deferred<T>(): Deferred<T> {
  let resolve!: (v: T) => void, reject!: (e: unknown) => void
  const promise = new Promise<T>((a, b) => { resolve = a; reject = b })
  return { promise, resolve, reject }
}

type Log = string[]
interface FakePlatform extends PresentPlatform {
  log: Log
  fs: Element | null
  /** 'now' enters at once; 'reject' refuses at once; 'defer' waits for settle(). */
  settle?: { enter(): void; refuse(): void }
}
function fakePlatform(log: Log, o: { activation?: boolean | undefined; fs?: 'now' | 'reject' | 'defer' } = {}): FakePlatform {
  const p: FakePlatform = {
    log,
    fs: null,
    hasActivation: () => ('activation' in o ? o.activation : true),
    requestFullscreen: (el) => {
      log.push(`requestFullscreen:${(el as HTMLElement).style.visibility}`)
      if (o.fs === 'reject') return Promise.reject(new Error('denied'))
      if (o.fs === 'defer') {
        return new Promise<void>((ok, no) => { p.settle = { enter: () => { p.fs = el; ok() }, refuse: () => no(new Error('denied')) } })
      }
      p.fs = el
      return Promise.resolve()
    },
    exitFullscreen: () => { log.push('exitFullscreen'); p.fs = null; return Promise.resolve() },
    fullscreenElement: () => p.fs,
    open: () => { log.push('open'); return null },
  }
  return p
}

function callbacks(log: Log, over: Partial<SlideShowCallbacks> = {}): SlideShowCallbacks {
  return {
    onStart: () => log.push('onStart'),
    onEnd: () => log.push('onEnd'),
    onDiagnostic: (d) => log.push(`diag:${d.code}`),
    currentSlide: () => 1,
    restoreFocus: () => log.push('restoreFocus'),
    scrollToSlide: (i) => log.push(`scroll:${i}`),
    ...over,
  }
}

/** A host in a fresh document, so each pure test has its own coordinator. */
function freshHost(): HTMLElement {
  const doc = document.implementation.createHTMLDocument('t')
  const host = doc.createElement('div')
  host.tabIndex = -1
  host.setAttribute('data-officeview-show', '')
  host.style.visibility = 'hidden'
  doc.body.append(host)
  return host
}

let loads: Deferred<typeof slideshow>[]
let log: Log
beforeEach(() => {
  log = []
  loads = []
  vi.spyOn(showLoader, 'load').mockImplementation(() => {
    log.push('import')
    const d = deferred<typeof slideshow>()
    loads.push(d)
    return d.promise
  })
  Element.prototype.scrollIntoView = function () {}
})
afterEach(() => {
  cleanup()
  vi.restoreAllMocks()
})

function start(host: HTMLElement, platform: PresentPlatform, over: Partial<Parameters<typeof startPresent>[0]> = {}) {
  return startPresent({ doc: PPTX, host, platform, generation: 1, callbacks: callbacks(log), allow: true, motion: 'auto', ...over })
}
const codeOf = (p: Promise<unknown>) => p.then(() => 'resolved', (e: unknown) => (e instanceof SlideShowError ? e.code : String(e)))
const flush = () => act(async () => { for (let i = 0; i < 5; i++) await Promise.resolve() })

describe('startPresent: the synchronous prologue', () => {
  it('rejects gesture-required when userActivation reports none and opens nothing', async () => {
    const host = freshHost()
    const p = fakePlatform(log, { activation: false })
    expect(await codeOf(start(host, p))).toBe('gesture-required')
    expect(log).toEqual([])
    expect(host.style.visibility).toBe('hidden')
    expect(acquireShowToken(host)).toBeDefined()
  })

  it('attempts the start when userActivation is missing (undefined)', async () => {
    const host = freshHost()
    const p = fakePlatform(log, { activation: undefined })
    const run = start(host, p)
    expect(log).toEqual(['requestFullscreen:visible', 'import'])
    host.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    expect(await codeOf(run)).toBe('cancelled')
  })

  it('PR #18 review 3: the ref reaches an interactive-sheets viewer; its present() rejects not-presentable', async () => {
    const { OfficePackage } = await import('../src/core/zip')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const book = await parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])))
    const ref = createRef<OfficeViewerHandle>()
    const ui = render(<OfficeDoc ref={ref} document={book} interactiveSheets />)
    expect(ref.current).toBeTruthy()
    expect(await codeOf(ref.current!.present())).toBe('not-presentable')
    ui.unmount()
    expect(ref.current).toBeNull()
  })

  it('PR #18 review 4: a full-screen retry that succeeds after the show ended leaves full screen again', async () => {
    const host = freshHost()
    // The first request is refused (the notice offers a retry); the retry is still pending when the show ends.
    const p = fakePlatform(log, { fs: 'defer' })
    let starts = 0
    const real = p.requestFullscreen
    p.requestFullscreen = (el) => (starts++ === 0 ? Promise.reject(new Error('denied')) : real(el))
    let handed!: ShowStart
    const run = start(host, p)
    loads[0].resolve({ ...slideshow, runShow: async (s: ShowStart) => { handed = s; return {} as never } } as typeof slideshow)
    await run
    expect(await handed.fullscreen).toBe('refused')
    const retry = handed.retryFullscreen()
    handed.cleanup()
    expect(log).not.toContain('exitFullscreen')
    p.settle!.enter()
    expect(await retry).toBe('entered')
    await flush()
    expect(log).toContain('exitFullscreen')
    expect(p.fs).toBeNull()
  })

  it('not-presentable for DOCX and for allowPresent=false; invalid-slide for 1.5, -1 and slideCount', async () => {
    const host = freshHost()
    const p = fakePlatform(log)
    expect(await codeOf(start(host, p, { doc: DOCX }))).toBe('not-presentable')
    expect(await codeOf(start(host, p, { allow: false }))).toBe('not-presentable')
    for (const from of [1.5, -1, 3, Number.NaN, Infinity]) expect(await codeOf(start(host, p, { opts: { from } }))).toBe('invalid-slide')
    // not-presentable is checked before the slide and the gesture
    expect(await codeOf(start(host, fakePlatform(log, { activation: false }), { allow: false, opts: { from: -1 } }))).toBe('not-presentable')
    expect(await codeOf(start(host, fakePlatform(log, { activation: false }), { opts: { from: -1 } }))).toBe('invalid-slide')
    expect(log).toEqual([])
    expect(acquireShowToken(host)).toBeDefined()
  })

  it('requestFullscreen is called synchronously before the chunk import', async () => {
    const host = freshHost()
    const p = fakePlatform(log)
    const run = start(host, p)
    // No await yet: the host was shown first, then full screen, then the import.
    expect(log).toEqual(['requestFullscreen:visible', 'import'])
    loads[0].resolve(slideshow)
    const session = await run
    session.exit()
    await session.ended
  })

  it("A's load fails after B entered full screen: A does not exit B's full screen and releases its token", async () => {
    const host = freshHost()
    const p = fakePlatform(log)
    const run = start(host, p)
    const other = host.ownerDocument.createElement('div')
    host.ownerDocument.body.append(other)
    p.fs = other // B now owns the document's full screen
    loads[0].reject(new Error('chunk 404'))
    expect(await codeOf(run)).toBe('load-failed')
    expect(log).not.toContain('exitFullscreen')
    expect(p.fs).toBe(other)
    expect(host.style.visibility).toBe('hidden')
    expect(acquireShowToken(host)).toBeDefined()
  })

  it('a failed start that owns full screen exits it, hides the host and releases the token, in that order', async () => {
    const host = freshHost()
    const p = fakePlatform(log)
    p.exitFullscreen = () => { log.push(`exitFullscreen:${host.style.visibility}`); p.fs = null; return Promise.resolve() }
    const run = start(host, p)
    await Promise.resolve()
    loads[0].reject(new Error('chunk 404'))
    expect(await codeOf(run)).toBe('load-failed')
    expect(log).toEqual(['requestFullscreen:visible', 'import', 'exitFullscreen:visible'])
    expect(host.style.visibility).toBe('hidden')
    expect(acquireShowToken(host)).toBeDefined()
  })

  it('a start that fails before the browser has entered full screen exits it once the request settles', async () => {
    const host = freshHost()
    const p = fakePlatform(log, { fs: 'defer' })
    const run = start(host, p)
    loads[0].reject(new Error('chunk 404'))
    expect(await codeOf(run)).toBe('load-failed')
    expect(host.style.visibility).toBe('hidden')
    expect(log).not.toContain('exitFullscreen')
    // the browser completes the request after the failure: it is left at once
    p.settle!.enter()
    await flush()
    expect(log.filter((l) => l === 'exitFullscreen')).toHaveLength(1)
    expect(p.fs).toBeNull()
    // and a request that is refused after the failure exits nothing
    log.length = 0
    const p2 = fakePlatform(log, { fs: 'defer' })
    const run2 = start(host, p2)
    loads[1].reject(new Error('chunk 404'))
    expect(await codeOf(run2)).toBe('load-failed')
    p2.settle!.refuse()
    await flush()
    expect(log).not.toContain('exitFullscreen')
  })

  it('a fullscreenchange for another element is ignored', async () => {
    const host = freshHost()
    const doc = host.ownerDocument
    const p = fakePlatform(log, { fs: 'defer' })
    const run = start(host, p)
    let settled = false
    void run.then(() => { settled = true }, () => { settled = true })
    const other = doc.createElement('div')
    doc.body.append(other)
    // Another element enters and leaves full screen while A has not entered.
    p.fs = other; doc.dispatchEvent(new Event('fullscreenchange'))
    p.fs = null; doc.dispatchEvent(new Event('fullscreenchange'))
    await Promise.resolve()
    expect(settled).toBe(false)
    // A enters; another element stacking on top is still not A leaving.
    p.settle!.enter(); doc.dispatchEvent(new Event('fullscreenchange'))
    p.fs = other; doc.dispatchEvent(new Event('fullscreenchange'))
    await Promise.resolve()
    expect(settled).toBe(false)
    // Full screen ends for the document (browser Esc): A cancels.
    p.fs = null; doc.dispatchEvent(new Event('fullscreenchange'))
    expect(await codeOf(run)).toBe('cancelled')
    expect(log).not.toContain('exitFullscreen')
  })

  it('a late chunk completion from an older generation is discarded', async () => {
    const host = freshHost()
    const p = fakePlatform(log)
    const epoch = new AbortController()
    const run = start(host, p, { signal: epoch.signal, generation: 1 })
    epoch.abort() // document change: generation 1 is over
    expect(await codeOf(run)).toBe('cancelled')
    const runShow = vi.fn()
    loads[0].resolve({ ...slideshow, runShow } as typeof slideshow)
    await flush()
    expect(runShow).not.toHaveBeenCalled()
    // Generation 2 starts normally.
    const run2 = start(host, p, { generation: 2 })
    loads[1].resolve(slideshow)
    const s = await run2
    s.exit()
    await s.ended
  })

  it('fullscreen resolving after runShow starts sets surface fullscreen; rejecting after it starts switches to in-window with a diagnostic; either outcome for a stale generation is ignored', async () => {
    const host = freshHost()
    // resolves after the handoff
    let p = fakePlatform(log, { fs: 'defer' })
    let run = start(host, p)
    loads[0].resolve(slideshow)
    let s = await run
    expect(s.surface).toBe('in-window')
    await act(async () => { p.settle!.enter() })
    await flush()
    expect(s.surface).toBe('fullscreen')
    s.exit(); await s.ended
    // rejects after the handoff
    log.length = 0
    p = fakePlatform(log, { fs: 'defer' })
    run = start(host, p)
    loads[1].resolve(slideshow)
    s = await run
    await act(async () => { p.settle!.refuse() })
    await flush()
    expect(s.surface).toBe('in-window')
    expect(log).toContain('diag:fullscreen-refused')
    expect(host.querySelector('[role=status]')?.textContent).toMatch(/full screen/i)
    s.exit(); await s.ended
    // stale: the session is over before full screen settles
    for (const outcome of ['enter', 'refuse'] as const) {
      log.length = 0
      p = fakePlatform(log, { fs: 'defer' })
      run = start(host, p)
      loads[loads.length - 1].resolve(slideshow)
      s = await run
      s.exit(); await s.ended
      await act(async () => { p.settle![outcome]() })
      await flush()
      expect(s.surface).toBe('in-window')
      expect(log).not.toContain('diag:fullscreen-refused')
    }
  })

  it('onSlideShowEnd fires exactly once before ended resolves; exit() is idempotent', async () => {
    const host = freshHost()
    const p = fakePlatform(log)
    const run = start(host, p)
    loads[0].resolve(slideshow)
    const s = await run
    expect(log).toContain('onStart')
    void s.ended.then(() => log.push('ended'))
    s.exit()
    s.exit()
    await s.ended
    s.exit()
    await flush()
    const tail = log.slice(log.indexOf('onStart') + 1)
    expect(tail).toEqual(['exitFullscreen', 'scroll:0', 'restoreFocus', 'onEnd', 'ended'])
    expect(host.style.visibility).toBe('hidden')
    expect(host.childElementCount).toBe(0)
    expect(acquireShowToken(host)).toBeDefined()
  })

  it('presenterView true before stage 2 is diagnosed and audience-only', async () => {
    const host = freshHost()
    const p = fakePlatform(log)
    const run = start(host, p, { opts: { presenterView: true } })
    loads[0].resolve(slideshow)
    const s = await run
    expect(log).toContain('diag:presenter-view-unavailable')
    expect(log).not.toContain('open')
    expect(s.presenter).toBe('none')
    s.exit(); await s.ended
  })

  it('the host shows a black letterbox with a busy indicator until the first paint, and hides it after', async () => {
    const host = freshHost()
    const p = fakePlatform(log)
    const run = start(host, p)
    // Loading: visible, empty, so the stylesheet's :empty busy indicator shows.
    expect(host.style.visibility).toBe('visible')
    expect(host.childElementCount).toBe(0)
    expect(css).toMatch(/\[data-officeview-show\]\[style\*="?visible"?\]:empty::after/)
    expect(host.matches('[data-officeview-show][style*="visible"]:empty')).toBe(true)
    expect(css).toMatch(/\[data-officeview-show\][^{]*\{[^}]*background:\s*(#000|black|rgb\(0, 0, 0\))/)
    loads[0].resolve(slideshow)
    const s = await run
    // First paint: the surface fills the host, so the indicator is gone.
    expect(host.childElementCount).toBeGreaterThan(0)
    expect(host.matches(':empty')).toBe(false)
    s.exit(); await s.ended
    expect(host.style.visibility).toBe('hidden')
  })
})

// ---------------------------------------------------------------------------
// Through the components
// ---------------------------------------------------------------------------

function mountDoc(props: Partial<Parameters<typeof OfficeDoc>[0]> = {}, platform = fakePlatform(log)) {
  const ref = createRef<OfficeViewerHandle>()
  const ui = render(<OfficeDoc ref={ref} document={PPTX} __presentPlatform={platform} {...props} />)
  const host = () => ui.container.querySelector<HTMLElement>('[data-officeview-show]')!
  const region = () => ui.getByRole('region', { name: 'Document viewer' })
  return { ref, ui, host, region, platform }
}

describe('OfficeDoc host handle', () => {
  it('renders the pre-mounted host hidden and empty, never display:none', () => {
    const { host } = mountDoc()
    expect(host()).toBeTruthy()
    expect(host().style.visibility).toBe('hidden')
    expect(host().style.display).not.toBe('none')
    expect(host().childElementCount).toBe(0)
  })

  it('a second present() in the same document rejects already-presenting, from the same or another viewer', async () => {
    const platform = fakePlatform(log)
    const a = createRef<OfficeViewerHandle>(), b = createRef<OfficeViewerHandle>()
    render(<><OfficeDoc ref={a} document={PPTX} __presentPlatform={platform} /><OfficeDoc ref={b} document={PPTX} __presentPlatform={platform} /></>)
    const first = a.current!.present()
    expect(await codeOf(a.current!.present())).toBe('already-presenting')
    expect(await codeOf(b.current!.present())).toBe('already-presenting')
    expect(log.filter((l) => l.startsWith('requestFullscreen'))).toHaveLength(1)
    loads[0].resolve(slideshow)
    const s = await first
    // a running show still holds the token
    expect(await codeOf(b.current!.present())).toBe('already-presenting')
    await act(async () => { s.exit(); await s.ended })
    // released: B can start now
    const second = b.current!.present()
    loads[1].resolve(slideshow)
    const s2 = await second
    await act(async () => { s2.exit(); await s2.ended })
  })

  it('load failure rejects load-failed, shows the notice and leaves the viewer usable; cancel by Esc, unmount, document change and allowPresent=false rejects cancelled', async () => {
    const m = mountDoc()
    let run = m.ref.current!.present()
    loads[0].reject(new Error('chunk 404'))
    expect(await codeOf(run)).toBe('load-failed')
    await flush()
    expect(m.ui.getByText('Slideshow could not load. Reload to retry.')).toBeTruthy()
    expect(m.host().style.visibility).toBe('hidden')
    expect(m.region()).toBeTruthy()
    // usable: a new start runs (and clears the notice)
    run = m.ref.current!.present()
    await flush()
    expect(m.ui.queryByText('Slideshow could not load. Reload to retry.')).toBeNull()
    // Esc during the load cancels
    fireEvent.keyDown(m.host(), { key: 'Escape' })
    expect(await codeOf(run)).toBe('cancelled')
    expect(log.filter((l) => l === 'exitFullscreen')).toHaveLength(2)
    expect(m.host().style.visibility).toBe('hidden')
    // document change cancels
    run = m.ref.current!.present()
    m.ui.rerender(<OfficeDoc ref={m.ref} document={{ ...PPTX } as PptxDocument} __presentPlatform={m.platform} />)
    expect(await codeOf(run)).toBe('cancelled')
    // allowPresent=false cancels
    run = m.ref.current!.present()
    m.ui.rerender(<OfficeDoc ref={m.ref} document={PPTX} allowPresent={false} __presentPlatform={m.platform} />)
    expect(await codeOf(run)).toBe('cancelled')
    expect(await codeOf(m.ref.current!.present())).toBe('not-presentable')
    // unmount cancels
    m.ui.rerender(<OfficeDoc ref={m.ref} document={PPTX} __presentPlatform={m.platform} />)
    run = m.ref.current!.present()
    m.ui.unmount()
    expect(await codeOf(run)).toBe('cancelled')
    // the late chunks never start a show
    for (const l of loads) l.resolve(slideshow)
    await flush()
    expect(log).not.toContain('onStart')
  })

  it('unmount, document change and allowPresent=false end a running session with exit semantics', async () => {
    const onEnd = vi.fn()
    for (const change of ['unmount', 'document', 'allow'] as const) {
      onEnd.mockClear()
      const m = mountDoc({ onSlideShowEnd: onEnd })
      const run = m.ref.current!.present()
      loads[loads.length - 1].resolve(slideshow)
      const s = await run
      await act(async () => {
        if (change === 'unmount') m.ui.unmount()
        else if (change === 'document') m.ui.rerender(<OfficeDoc ref={m.ref} document={{ ...PPTX } as PptxDocument} onSlideShowEnd={onEnd} __presentPlatform={m.platform} />)
        else m.ui.rerender(<OfficeDoc ref={m.ref} document={PPTX} allowPresent={false} onSlideShowEnd={onEnd} __presentPlatform={m.platform} />)
      })
      await s.ended
      expect(onEnd).toHaveBeenCalledTimes(1)
      cleanup()
    }
  })

  it('PR #18 review 8: the viewer hands its allowCopy policy to the show, whatever the consumer style says', async () => {
    const starts: ShowStart[] = []
    const runShow = vi.fn(async (st: ShowStart) => { starts.push(st); st.cleanup(); return {} as never })
    const m = mountDoc({ allowCopy: false, style: { userSelect: 'text' } })
    fireEvent.keyDown(m.region(), { key: 'F5' })
    loads[0].resolve({ ...slideshow, runShow } as typeof slideshow)
    await flush()
    expect(starts[0].allowCopy).toBe(false)
  })

  it('F5 and Shift+F5 inside the viewer start the show; presentShortcuts=false disables them; outside focus is ignored', async () => {
    const starts: ShowStart[] = []
    const runShow = vi.fn(async (st: ShowStart) => { starts.push(st); st.cleanup(); return {} as never })
    const m = mountDoc()
    const ev = fireEvent.keyDown(m.region(), { key: 'F5' })
    expect(ev).toBe(false) // default prevented
    loads[0].resolve({ ...slideshow, runShow } as typeof slideshow)
    await flush()
    fireEvent.keyDown(m.region(), { key: 'F5', shiftKey: true })
    loads[1].resolve({ ...slideshow, runShow } as typeof slideshow)
    await flush()
    expect(starts.map((s) => s.from)).toEqual(['beginning', 'current'])
    // Alt+F5 is Presenter View (stage 2): not routed here
    expect(fireEvent.keyDown(m.region(), { key: 'F5', altKey: true })).toBe(true)
    // PR #18 review 8: Ctrl/Cmd+F5 stays the browser's hard reload, with or without Shift
    expect(fireEvent.keyDown(m.region(), { key: 'F5', ctrlKey: true })).toBe(true)
    expect(fireEvent.keyDown(m.region(), { key: 'F5', metaKey: true })).toBe(true)
    expect(fireEvent.keyDown(m.region(), { key: 'F5', ctrlKey: true, shiftKey: true })).toBe(true)
    expect(loads).toHaveLength(2)
    // outside the viewer
    fireEvent.keyDown(document.body, { key: 'F5' })
    expect(loads).toHaveLength(2)
    // disabled
    m.ui.rerender(<OfficeDoc ref={m.ref} document={PPTX} presentShortcuts={false} __presentPlatform={m.platform} />)
    expect(fireEvent.keyDown(m.region(), { key: 'F5' })).toBe(true)
    // a DOCX never routes F5
    m.ui.rerender(<OfficeDoc ref={m.ref} document={DOCX} __presentPlatform={m.platform} />)
    expect(fireEvent.keyDown(m.region(), { key: 'F5' })).toBe(true)
    expect(loads).toHaveLength(2)
  })

  it('restores focus to the start element and scrolls to the last slide on exit', async () => {
    const m = mountDoc()
    m.region().focus()
    const run = m.ref.current!.present({ from: 2 })
    expect(document.activeElement).toBe(m.host())
    loads[0].resolve(slideshow)
    const s = await run
    expect(s.slideIndex).toBe(2)
    await act(async () => { s.exit(); await s.ended })
    expect(document.activeElement).toBe(m.region())
  })
})

describe('viewer navigation (lazy viewerCallbacks)', () => {
  it("'current' starts on the page under the viewer's middle; exit scrolls the viewer to the last slide shown", async () => {
    const m = mountDoc()
    await waitFor(() => expect(m.region().querySelectorAll('canvas')).toHaveLength(3))
    const region = m.region()
    let scrollTop = 0
    Object.defineProperty(region, 'scrollTop', { configurable: true, get: () => scrollTop, set: (v: number) => { scrollTop = v } })
    region.getBoundingClientRect = () => ({ top: 100, height: 800, bottom: 900 }) as DOMRect
    // pages 300px tall from y=0: the middle (500) falls on page index 1
    m.region().querySelectorAll('canvas').forEach((c, i) => { c.getBoundingClientRect = () => ({ top: i * 316 - scrollTop, bottom: i * 316 + 300 - scrollTop }) as DOMRect })
    const run = m.ref.current!.present({ from: 'current' })
    loads[0].resolve(slideshow)
    const s = await run
    expect(s.slideIndex).toBe(1)
    s.goTo(2)
    await act(async () => { s.exit() })
    await s.ended
    // page 2's top (632) lands at the region's top (100) plus the 16px page padding
    expect(scrollTop).toBe(632 - 100 - 16)
  })
})

describe('OfficeFile source', () => {
  it('a re-readable source reaches runShow; a one-shot source does not', async () => {
    const bytes = await buildPptx([])
    const starts: ShowStart[] = []
    const runShow = vi.fn(async (st: ShowStart) => { starts.push(st); st.cleanup(); return {} as never })
    const sources = [
      { data: bytes, keep: true },
      { data: new Blob([bytes as BlobPart]), keep: true },
      { data: bytes.buffer.slice(0) as ArrayBuffer, keep: true },
      { data: new Response(bytes as BodyInit), keep: false },
      { data: Promise.resolve(bytes), keep: false },
    ]
    for (const { data } of sources) {
      const ref = createRef<OfficeViewerHandle>()
      const ui = render(<OfficeFile ref={ref} data={data} __presentPlatform={fakePlatform(log)} />)
      await waitFor(() => expect(ref.current).toBeTruthy())
      const run = ref.current!.present()
      loads[loads.length - 1].resolve({ ...slideshow, runShow } as typeof slideshow)
      await run
      ui.unmount()
    }
    expect(starts.map((s) => s.source)).toEqual(sources.map((s) => (s.keep ? s.data : undefined)))
    // the parsed document carries its package handle
    expect(starts.every((s) => s.handle && s.handle.slideParts.length === 1)).toBe(true)
  })

  it('the viewer watermark reaches runShow (S1-20: every slideshow surface paints it)', async () => {
    const bytes = await buildPptx([])
    const starts: ShowStart[] = []
    const runShow = vi.fn(async (st: ShowStart) => { starts.push(st); st.cleanup(); return {} as never })
    const ref = createRef<OfficeViewerHandle>()
    const ui = render(<OfficeFile ref={ref} data={bytes} watermark={{ text: 'DRAFT' }} __presentPlatform={fakePlatform(log)} />)
    await waitFor(() => expect(ref.current).toBeTruthy())
    const run = ref.current!.present()
    loads[loads.length - 1].resolve({ ...slideshow, runShow } as typeof slideshow)
    await run
    ui.unmount()
    expect(starts[0].watermark).toMatchObject({ text: 'DRAFT' })
  })
})
