/**
 * Streaming Stage 3 piece 3, Task 7: `present()` on a streamed document
 * (design §4.3; parent §8.2, §12a.7). The eager start reads the manifest
 * synchronously, keeps the gesture rule (host shown and full screen requested
 * before the lazy remote show module resolves) and hands the lazy
 * `runRemoteShow` the start; `from` is a streamed index (hidden excluded, R19).
 *
 * jsdom with an injected PresentPlatform. Both chunk loaders are spied: the
 * remote show is deferred and never the local slideshow chunk.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { createRef } from 'react'
import { OfficeDoc } from '../../src/components/OfficeDoc'
import RemoteDocView from '../../src/components/remote/RemoteDocView'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import {
  acquireShowToken, preparedPlacement, showLoader, SlideShowError, startPresent,
  type OfficeViewerHandle, type PresentPlatform, type SlideShowSession,
} from '../../src/components/present'
import { presentRemote, remoteShowLoader } from '../../src/stream/remote/present-start'
import type { RemoteShowStart } from '../../src/stream/remote/show'
import type { RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'

type RemoteShowModule = typeof import('../../src/stream/remote/show')

function fakeDoc(o: { kind?: 'pptx' | 'docx'; present?: boolean; notes?: boolean; count?: number; hidden?: readonly number[] } = {}) {
  const count = o.count ?? 3
  const list = Array.from({ length: count }, (_, i) => (o.hidden?.includes(i) ? { w: 200, h: 100, hidden: true as const } : { w: 200, h: 100 }))
  const listeners = new Set<() => void>()
  const doc = {
    kind: 'remote',
    manifest: {
      name: 'deck.pptx', kind: o.kind ?? 'pptx', units: { count, list },
      can: { search: false, copy: false, print: false, textA11y: false, present: o.present ?? true, notes: o.notes ?? false, links: false },
      caps: { copy: null, a11y: 'off', print: null }, recorder: { scalesBuckets: [1, 2] }, notices: [],
    },
    onManifestChange: vi.fn((cb: () => void) => { listeners.add(cb); return () => { listeners.delete(cb) } }),
    resume(can: Partial<RemoteOfficeDoc['manifest']['can']>) {
      doc.manifest = { ...doc.manifest, can: { ...doc.manifest.can, ...can } }
      act(() => { for (const l of [...listeners]) l() })
    },
    units: async () => list.map((u) => ({ widthPx: u.w, heightPx: u.h })),
    paint: vi.fn(async () => {}),
    paintNotes: vi.fn(async () => {}),
    search: vi.fn(), select: vi.fn(), copy: vi.fn(), a11y: vi.fn(), print: vi.fn(),
    close: vi.fn(),
    ended: new Promise(() => {}),
  }
  return doc
}
type Fake = ReturnType<typeof fakeDoc>
const asRemote = (d: Fake) => d as unknown as RemoteOfficeDoc

function fakePlatform(log: string[], o: { activation?: boolean } = {}): PresentPlatform {
  return {
    hasActivation: () => o.activation ?? true,
    requestFullscreen: (el) => { log.push(`requestFullscreen:${(el as HTMLElement).style.visibility}`); return new Promise(() => {}) },
    exitFullscreen: async () => { log.push('exitFullscreen') },
    fullscreenElement: () => null,
    open: () => { log.push('open'); return null },
  }
}

function freshHost(): HTMLElement {
  const d = document.implementation.createHTMLDocument('t')
  const host = d.createElement('div')
  host.tabIndex = -1
  host.style.visibility = 'hidden'
  d.body.append(host)
  return host
}

const codeOf = (p: Promise<unknown>) => p.then(() => 'resolved', (e: unknown) => (e instanceof SlideShowError ? e.code : String(e)))
const flush = () => act(async () => { for (let i = 0; i < 10; i++) await Promise.resolve() })

let log: string[]
let started: RemoteShowStart[]
let release: Array<(m: RemoteShowModule) => void>
const SESSION = { slideIndex: 0 } as unknown as SlideShowSession
beforeEach(() => {
  log = []
  started = []
  release = []
  vi.spyOn(showLoader, 'load').mockImplementation(() => { log.push('local-import'); return new Promise(() => {}) })
  vi.spyOn(remoteShowLoader, 'load').mockImplementation(() => {
    log.push('remote-import')
    return new Promise<RemoteShowModule>((ok) => { release.push(ok) })
  })
})
afterEach(() => { cleanup(); vi.restoreAllMocks() })

/** The deferred remote module: its runRemoteShow records the start and, as the real one, cleans up when the start's signal aborts. */
const remoteModule = () => ({
  runRemoteShow: async (s: RemoteShowStart) => {
    started.push(s); log.push('runRemoteShow')
    s.signal.addEventListener('abort', () => s.cleanup())
    return SESSION
  },
}) as unknown as RemoteShowModule

function start(doc: Fake, host: HTMLElement, platform: PresentPlatform, opts: Parameters<typeof startPresent>[0]['opts'] = {}) {
  return presentRemote({
    doc: asRemote(doc), host, platform, generation: 1, allow: true, motion: 'auto', opts,
    callbacks: { currentSlide: () => 0, scrollToSlide: () => {}, restoreFocus: () => {} },
  })
}

describe('startPresent on a streamed document: the synchronous prologue', () => {
  it('without user activation rejects gesture-required and opens nothing', async () => {
    const host = freshHost()
    expect(await codeOf(start(fakeDoc(), host, fakePlatform(log, { activation: false })))).toBe('gesture-required')
    expect(log).toEqual([])
    expect(host.style.visibility).toBe('hidden')
    expect(acquireShowToken(host)).toBeDefined()
  })

  it('a streamed docx (can.present false) is not-presentable', async () => {
    const host = freshHost()
    expect(await codeOf(start(fakeDoc({ kind: 'docx', present: false }), host, fakePlatform(log)))).toBe('not-presentable')
    expect(log).toEqual([])
  })

  it('a deck whose every unit is hidden is not-presentable', async () => {
    expect(await codeOf(start(fakeDoc({ count: 2, hidden: [0, 1] }), freshHost(), fakePlatform(log)))).toBe('not-presentable')
    expect(log).toEqual([])
  })

  it('presenterView: true with can.notes false is not-presentable: no window.open, no full screen', async () => {
    const host = freshHost()
    expect(await codeOf(start(fakeDoc({ notes: false }), host, fakePlatform(log), { presenterView: true }))).toBe('not-presentable')
    expect(log).toEqual([])
    expect(host.style.visibility).toBe('hidden')
  })

  it('presenterView: "auto" with can.notes false opens no presenter window even on a prepared screen', async () => {
    const host = freshHost()
    preparedPlacement.set(host.ownerDocument, { popupFeatures: 'popup', audienceScreen: { availWidth: 1, availHeight: 1 } })
    void start(fakeDoc({ notes: false }), host, fakePlatform(log)).catch(() => {})
    expect(log).toEqual(['requestFullscreen:visible', 'remote-import'])
  })

  it('presenterView: true with can.notes opens the presenter window first, then full screen', async () => {
    const host = freshHost()
    void start(fakeDoc({ notes: true }), host, fakePlatform(log), { presenterView: true }).catch(() => {})
    expect(log).toEqual(['open', 'requestFullscreen:visible', 'remote-import'])
  })

  it('from is a streamed index: hidden units are not counted (R19)', async () => {
    // Units 0..3, unit 1 hidden: three streamed slides 0..2.
    const doc = fakeDoc({ count: 4, hidden: [1] })
    expect(await codeOf(start(doc, freshHost(), fakePlatform(log), { from: 3 }))).toBe('invalid-slide')
    expect(await codeOf(start(doc, freshHost(), fakePlatform(log), { from: -1 }))).toBe('invalid-slide')
    expect(await codeOf(start(doc, freshHost(), fakePlatform(log), { from: 1.5 }))).toBe('invalid-slide')
    expect(log).toEqual([])
    const run = start(doc, freshHost(), fakePlatform(log), { from: 2 })
    release[0]!(remoteModule())
    await expect(run).resolves.toBe(SESSION)
    expect(started[0]!.from).toBe(2)
  })

  it('with activation: host shown and full screen requested before the remote module resolves; never the local chunk', async () => {
    const host = freshHost()
    const doc = fakeDoc()
    const run = start(doc, host, fakePlatform(log))
    // Synchronous: the surface is up inside the gesture, the import only started.
    expect(log).toEqual(['requestFullscreen:visible', 'remote-import'])
    expect(host.style.visibility).toBe('visible')
    expect(started).toEqual([])
    release[0]!(remoteModule())
    await expect(run).resolves.toBe(SESSION)
    expect(log).toEqual(['requestFullscreen:visible', 'remote-import', 'runRemoteShow'])
    expect(showLoader.load).not.toHaveBeenCalled()
    const s = started[0]!
    expect(s.doc).toBe(doc)
    expect(s.from).toBe('beginning')
    expect(s.host).toBe(host)
    expect(s.popup).toBeUndefined()
    expect(typeof s.cleanup).toBe('function')
    expect(s.token.owns(host)).toBe(true)
    // The token is held: a second start is refused while the show runs.
    expect(await codeOf(start(doc, host, fakePlatform([])))).toBe('already-presenting')
  })

  it('a remote module that fails to load rejects load-failed and hides the host', async () => {
    const host = freshHost()
    vi.mocked(remoteShowLoader.load).mockImplementationOnce(() => Promise.reject(new Error('chunk')))
    expect(await codeOf(start(fakeDoc(), host, fakePlatform(log)))).toBe('load-failed')
    expect(host.style.visibility).toBe('hidden')
    expect(acquireShowToken(host)).toBeDefined()
  })

  it('the eager starter knows nothing remote: a stream handed to startPresent without its launch is not-presentable', async () => {
    const host = freshHost()
    const run = startPresent({ doc: asRemote(fakeDoc()), host, platform: fakePlatform(log), generation: 1, allow: true, motion: 'auto', callbacks: {} })
    expect(await codeOf(run)).toBe('not-presentable')
    expect(log).toEqual([])
    // Lazy boundary: no manifest reading, hidden filtering or remote show loader in the eager module.
    const src = readFileSync(resolve(__dirname, '../../src/components/present.ts'), 'utf8').replace(/\/\*[\s\S]*?\*\/|\/\/.*$/gm, '')
    expect(src).not.toMatch(/stream\/remote\/show|manifest|\.hidden\b|can\.notes|runRemoteShow/)
  })

  it('a local deck never loads the remote show module', async () => {
    const host = freshHost()
    const deck = { slides: [{ index: 0, widthEmu: 9144000, heightEmu: 6858000, shapes: [] }], slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [] }
    void startPresent({ doc: deck as never, host, platform: fakePlatform(log), generation: 1, allow: true, motion: 'auto', callbacks: {} }).catch(() => {})
    expect(log).toEqual(['requestFullscreen:visible', 'local-import'])
    expect(remoteShowLoader.load).not.toHaveBeenCalled()
  })
})

const view = (doc: Fake, props: Record<string, unknown> = {}) =>
  render(<RemoteDocView document={asRemote(doc)} __presentPlatform={fakePlatform(log)} {...props} />)
const presentBtn = (v: ReturnType<typeof render>) => v.container.querySelector<HTMLButtonElement>('button[data-ov-present-button]')
const showHost = (v: ReturnType<typeof render>) => v.container.querySelector<HTMLElement>('[data-officeview-show]')

describe('RemoteDocView: Present button, F5 and the show host', () => {
  it('no Present button when can.present is false; it appears when a resume turns it on', () => {
    const doc = fakeDoc({ kind: 'docx', present: false })
    const v = view(doc)
    expect(presentBtn(v)).toBeNull()
    doc.resume({ present: true })
    expect(presentBtn(v)).not.toBeNull()
    doc.resume({ present: false })
    expect(presentBtn(v)).toBeNull()
  })

  it('a session that ends drops the Present button and refuses present() (not-presentable)', async () => {
    let end!: () => void
    const doc = fakeDoc()
    doc.ended = new Promise<void>((r) => { end = r })
    const presentRef = { current: null as null | ((o?: object) => Promise<SlideShowSession>) }
    const v = view(doc, { presentRef })
    expect(presentBtn(v)).not.toBeNull()
    end()
    await flush()
    expect(presentBtn(v)).toBeNull()
    expect(await codeOf(presentRef.current!())).toBe('not-presentable')
    expect(log).toEqual([])
  })

  it('a session that ends while the show module is still loading cancels the start at once', async () => {
    let end!: () => void
    const doc = fakeDoc()
    doc.ended = new Promise<void>((r) => { end = r })
    const v = view(doc)
    fireEvent.click(presentBtn(v)!)
    expect(showHost(v)!.style.visibility).toBe('visible')
    end()
    await flush()
    // The eager cleanup ran: host hidden, token released, without waiting for the import.
    expect(showHost(v)!.style.visibility).toBe('hidden')
    expect(acquireShowToken(showHost(v)!)?.release()).toBeUndefined()
    // The import resolving later starts nothing.
    release[0]!(remoteModule())
    await flush()
    expect(started).toEqual([])
    expect(log).not.toContain('runRemoteShow')
  })

  it('a session end aborts only its own document epoch: a new document can still present', async () => {
    let end!: () => void
    const first = fakeDoc()
    first.ended = new Promise<void>((r) => { end = r })
    const v = view(first)
    const second = fakeDoc()
    v.rerender(<RemoteDocView document={asRemote(second)} __presentPlatform={fakePlatform(log)} />)
    end()
    await flush()
    fireEvent.click(presentBtn(v)!)
    expect(showHost(v)!.style.visibility).toBe('visible')
    release[0]!(remoteModule())
    await flush()
    expect(started.length).toBe(1)
    expect(started[0]!.signal.aborted).toBe(false)
  })

  it('a chunk that fails to load says so in the view', async () => {
    vi.mocked(remoteShowLoader.load).mockImplementationOnce(() => Promise.reject(new Error('chunk')))
    const v = view(fakeDoc())
    fireEvent.click(presentBtn(v)!)
    await flush()
    expect(v.container.querySelector('[data-ov-present-notice]')?.textContent).toBe('Slideshow could not load. Reload to retry.')
    expect(showHost(v)!.style.visibility).toBe('hidden')
  })

  it('no Present button when the host withholds presenting (allowPresent false)', () => {
    expect(presentBtn(view(fakeDoc(), { allowPresent: false }))).toBeNull()
  })

  it('the show host sits inside the view theme root, beside the page region, hidden', () => {
    const v = view(fakeDoc())
    const host = showHost(v)!
    expect(host).not.toBeNull()
    expect(host.style.visibility).toBe('hidden')
    expect(host.closest('[data-ov-theme-root]')).not.toBeNull()
    expect(host.parentElement!.querySelector(':scope > [data-officeview-root]')).not.toBeNull()
  })

  it('the Present button starts the show synchronously in the click: host shown, full screen, then the import', async () => {
    const v = view(fakeDoc())
    fireEvent.click(presentBtn(v)!)
    expect(log).toEqual(['requestFullscreen:visible', 'remote-import'])
    expect(showHost(v)!.style.visibility).toBe('visible')
    release[0]!(remoteModule())
    await flush()
    expect(started[0]!.from).toBe('beginning')
    expect(started[0]!.host).toBe(showHost(v))
  })

  it('F5 starts from the beginning, Shift+F5 from the current slide; Ctrl/Cmd+F5 is left to the browser', async () => {
    const v = view(fakeDoc())
    const root = v.container.querySelector('[data-officeview-root]')!
    fireEvent.keyDown(root, { key: 'F5', ctrlKey: true })
    fireEvent.keyDown(root, { key: 'F5', metaKey: true })
    expect(log).toEqual([])
    fireEvent.keyDown(root, { key: 'F5', shiftKey: true })
    expect(log).toEqual(['requestFullscreen:visible', 'remote-import'])
    release[0]!(remoteModule())
    await flush()
    expect(started[0]!.from).toBe('current')
    // The viewer callbacks speak unit (page) indices: units are 100 px tall, the region 300 px, so unit 1 is under its middle.
    vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
      const u = (this as HTMLElement).dataset?.unit
      return (u === undefined ? { top: 0, bottom: 300, height: 300 } : { top: Number(u) * 100, bottom: Number(u) * 100 + 100, height: 100 }) as DOMRect
    })
    expect(started[0]!.callbacks.currentSlide()).toBe(1)
    const region = root as HTMLElement
    region.scrollTop = 0
    started[0]!.callbacks.scrollToSlide(2)
    expect(region.scrollTop).toBe(200 - 16)
  })

  it('F5 does nothing with presentShortcuts off or on a non-presentable stream', () => {
    const a = view(fakeDoc(), { presentShortcuts: false })
    fireEvent.keyDown(a.container.querySelector('[data-officeview-root]')!, { key: 'F5' })
    cleanup()
    const b = view(fakeDoc({ present: false }))
    fireEvent.keyDown(b.container.querySelector('[data-officeview-root]')!, { key: 'F5' })
    expect(log).toEqual([])
  })
})

describe('<OfficeDoc> ref on a streamed document', () => {
  it('present() before the lazy view has mounted rejects not-presentable (typed, never a TypeError)', async () => {
    const ref = createRef<OfficeViewerHandle>()
    const v = render(<OfficeDoc ref={ref} document={asRemote(fakeDoc())} __presentPlatform={fakePlatform(log)} />)
    if (!showHost(v)) expect(await codeOf(ref.current!.present())).toBe('not-presentable')
    expect(log).toEqual([])
  })

  it('present() reaches the remote start once the view is mounted', async () => {
    const ref = createRef<OfficeViewerHandle>()
    const doc = fakeDoc()
    const v = render(<OfficeDoc ref={ref} document={asRemote(doc)} __presentPlatform={fakePlatform(log)} />)
    await waitFor(() => expect(showHost(v)).not.toBeNull())
    const run = ref.current!.present({ from: 1 })
    expect(log).toEqual(['requestFullscreen:visible', 'remote-import'])
    release[0]!(remoteModule())
    await expect(run).resolves.toBe(SESSION)
    expect(started[0]!.from).toBe(1)
  })

  it('present() rejects not-presentable on a stream without can.present, and with allowPresent false', async () => {
    const ref = createRef<OfficeViewerHandle>()
    const v = render(<OfficeDoc ref={ref} document={asRemote(fakeDoc({ kind: 'docx', present: false }))} __presentPlatform={fakePlatform(log)} />)
    await waitFor(() => expect(showHost(v)).not.toBeNull())
    expect(await codeOf(ref.current!.present())).toBe('not-presentable')
    v.rerender(<OfficeDoc ref={ref} document={asRemote(fakeDoc())} allowPresent={false} __presentPlatform={fakePlatform(log)} />)
    expect(await codeOf(ref.current!.present())).toBe('not-presentable')
    expect(log).toEqual([])
  })

  it('unmounting the viewer ends a running show (the document epoch)', async () => {
    const ref = createRef<OfficeViewerHandle>()
    const v = render(<OfficeDoc ref={ref} document={asRemote(fakeDoc())} __presentPlatform={fakePlatform(log)} />)
    await waitFor(() => expect(showHost(v)).not.toBeNull())
    const run = ref.current!.present()
    v.unmount()
    expect(await codeOf(run)).toBe('cancelled')
  })
})
