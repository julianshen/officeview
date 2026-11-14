/**
 * Streaming Stage 3 piece 3, Task 6: the remote show session (`runRemoteShow`)
 * over a fake RemoteOfficeDoc and a manual fetcher clock.
 *
 * jsdom with node-canvas. Every unit paint and notes paint is held until the
 * test settles it; the fetcher's pacing and the notes settle delay run on the
 * injected clock, so nothing waits on wall time. Presenter View is the real
 * one, mounted by the session (single-screen swap: `popup: null`).
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { remoteShowDeps, runRemoteShow, type RemoteShowStart } from '../../src/stream/remote/show'
import { StreamSessionError, type StreamErrorCode } from '../../src/stream/remote/errors'
import type { Manifest } from '../../src/stream/manifest'
import type { RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'

const UNIT = { w: 960, h: 540 }

interface Held {
  unit: number; scale: number | 'fit'; dpr: number; canvas: HTMLCanvasElement; signal?: AbortSignal; onSend?: () => void
  settled: boolean
  ok(color?: string): Promise<void>
  fail(code: StreamErrorCode): Promise<void>
}

/** A streamed deck: `count` units, `hidden` flagged; paints and notes are held until settled. */
function fakeDoc(o: { count?: number; hidden?: readonly number[]; notes?: boolean; buckets?: number[] } = {}) {
  const count = o.count ?? 4
  const list = Array.from({ length: count }, (_, i) => (o.hidden?.includes(i) ? { ...UNIT, hidden: true as const } : { ...UNIT }))
  let manifest = {
    units: { count, list },
    can: { search: false, copy: false, print: false, textA11y: false, present: true, notes: o.notes ?? false, links: false },
    recorder: { scalesBuckets: o.buckets ?? [1, 2] },
  } as unknown as Manifest
  const paints: Held[] = []
  const notes: Held[] = []
  const listeners = new Set<() => void>()
  let endIt!: (e: StreamSessionError | undefined) => void
  const ended = new Promise<StreamSessionError | undefined>((r) => { endIt = r })
  const hold = (into: Held[], unit: number, scale: number | 'fit', dpr: number, canvas: HTMLCanvasElement, signal?: AbortSignal, onSend?: () => void, w = UNIT.w, h = UNIT.h) =>
    new Promise<void>((resolve, reject) => {
      into.push({
        unit, scale, dpr, canvas, signal, onSend, settled: false,
        async ok(color = '#f00') {
          this.settled = true
          if (!signal?.aborted) {
            canvas.width = w; canvas.height = h
            const ctx = canvas.getContext('2d')!
            ctx.fillStyle = color
            ctx.fillRect(0, 0, w, h)
          }
          resolve()
          await flush()
        },
        async fail(code) { this.settled = true; reject(new StreamSessionError(code)); await flush() },
      })
    })
  const doc = {
    kind: 'remote',
    get manifest() { return manifest },
    paint: vi.fn((unit: number, scale: number | 'fit', dpr: number, canvas: HTMLCanvasElement, signal?: AbortSignal, po?: { onSend?: () => void }) => {
      po?.onSend?.()
      const k = typeof scale === 'number' ? scale * dpr : dpr
      return hold(paints, unit, scale, dpr, canvas, signal, po?.onSend, UNIT.w * k, UNIT.h * k)
    }),
    paintNotes: vi.fn((unit: number, dpr: number, canvas: HTMLCanvasElement, signal?: AbortSignal) => {
      if (!manifest.can.notes) return Promise.reject(new StreamSessionError('policy-refused'))
      return hold(notes, unit, 1, dpr, canvas, signal, undefined, 720, 100)
    }),
    onManifestChange(cb: () => void) { listeners.add(cb); return () => { listeners.delete(cb) } },
    ended,
    close: vi.fn(),
  } as unknown as RemoteOfficeDoc & { paint: ReturnType<typeof vi.fn>; paintNotes: ReturnType<typeof vi.fn>; close: ReturnType<typeof vi.fn> }
  return {
    doc, paints, notes,
    end: async (e?: StreamSessionError) => { endIt(e); await flush() },
    setManifest(m: Partial<Manifest['can']>) { manifest = { ...manifest, can: { ...manifest.can, ...m } }; for (const l of [...listeners]) l() },
    listeners,
  }
}

/** The fetcher's (and the notes settle delay's) clock: time moves only through `advance`. */
function manualClock() {
  let t = 0
  let timers: Array<{ at: number; cb: () => void }> = []
  return {
    now: () => t,
    wake(ms: number, cb: () => void) {
      const e = { at: t + ms, cb }
      timers.push(e)
      return () => { timers = timers.filter((x) => x !== e) }
    },
    async advance(ms: number) {
      const end = t + ms
      for (;;) {
        const due = timers.filter((x) => x.at <= end).sort((a, b) => a.at - b.at)[0]
        if (!due) break
        timers = timers.filter((x) => x !== due)
        t = Math.max(t, due.at)
        due.cb()
        await flush()
      }
      t = end
      await flush()
    },
  }
}

const flush = async () => { for (let i = 0; i < 10; i++) await new Promise((r) => setTimeout(r, 0)) }
const px = (c: HTMLCanvasElement, x: number, y: number) => [...c.getContext('2d')!.getImageData(x, y, 1, 1).data]
const all = (c: HTMLCanvasElement) => [...c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data]

let clock: ReturnType<typeof manualClock>
beforeEach(() => {
  clock = manualClock()
  vi.spyOn(remoteShowDeps, 'clock').mockImplementation(() => clock)
  // The presenter's 1 Hz refresh is driven by tests through `tick()`.
  vi.spyOn(remoteShowDeps, 'every').mockImplementation(() => () => {})
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: 960 })
  Object.defineProperty(window, 'innerHeight', { configurable: true, value: 540 })
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
})

async function show(fd: ReturnType<typeof fakeDoc>, o: Partial<RemoteShowStart> = {}) {
  const host = document.createElement('div')
  host.tabIndex = -1
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const cleanup = vi.fn(() => { ac.abort() })
  const start: RemoteShowStart = {
    doc: fd.doc, host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'),
    retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'auto',
    callbacks: {
      onStart: () => log.push('onStart'),
      onEnd: () => log.push('onEnd'),
      onSlideChange: (i, b) => log.push(`slide:${i}:${b}`),
      onDiagnostic: (d) => log.push(`diag:${d.code}`),
      currentSlide: () => 0, restoreFocus() {}, scrollToSlide: (i) => log.push(`scroll:${i}`),
    },
    signal: ac.signal, generation: 1, cleanup,
    ...o,
  }
  const s = await runRemoteShow(start)
  await flush()
  const canvas = host.querySelector<HTMLCanvasElement>('canvas[data-officeview-show-canvas]')!
  const root = host.querySelector<HTMLElement>('[data-officeview-show-surface]')!
  return { s, host, log, cleanup, canvas, root }
}

const frame = (root: HTMLElement) => root.getAttribute('data-officeview-show-frame')
const status = (host: HTMLElement) => host.querySelector('[data-officeview-presenter-status]')?.textContent
const units = (fd: ReturnType<typeof fakeDoc>) => fd.paints.map((p) => p.unit)

describe('remote show: order and navigation', () => {
  it('the order skips hidden units and the counter excludes them', async () => {
    const fd = fakeDoc({ count: 4, hidden: [1] })
    const { s, host } = await show(fd, { popup: null })
    expect(s.state.order).toEqual([0, 1, 2])
    expect(s.presenterSource.slideCount).toBe(3)
    expect(status(host)).toBe('Slide 1 of 3')
    expect(units(fd)).toEqual([0])
    await fd.paints[0]!.ok()
    s.next()
    await clock.advance(400)
    expect(s.state.slide).toBe(1)
    expect(s.slideIndex).toBe(1)
    expect(status(host)).toBe('Slide 2 of 3')
    // The streamed slide 2 is unit 2: the hidden unit 1 is never requested.
    expect(units(fd).at(-1)).toBe(2)
    await clock.advance(5_000)
    expect(units(fd)).not.toContain(1)
    s.exit()
  })

  it('goTo(2) with unit 1 hidden shows unit 3; out-of-range goTo is ignored', async () => {
    const fd = fakeDoc({ count: 4, hidden: [1] })
    const { s, log, canvas } = await show(fd)
    await fd.paints[0]!.ok('#00f')
    s.goTo(2)
    await clock.advance(400)
    expect(s.slideIndex).toBe(2)
    const p = fd.paints.find((x) => x.unit === 3)!
    expect(p).toBeDefined()
    await p.ok('#0f0')
    expect(px(canvas, canvas.width >> 1, canvas.height >> 1)).toEqual([0, 255, 0, 255])
    expect(log).toContain('slide:2:0')
    s.goTo(3)
    expect(s.slideIndex).toBe(2)
    expect(log).toContain('diag:invalid-slide')
    s.exit()
  })

  it('next past the end shows the end slide; next again exits', async () => {
    const fd = fakeDoc({ count: 2 })
    const { s, root, log } = await show(fd)
    s.next(); s.next()
    expect(s.state.slide).toBe('end')
    expect(root.getAttribute('data-slide')).toBe('end')
    s.previous()
    expect(s.state.slide).toBe(1)
    s.next()
    s.next()
    await s.ended
    expect(log).toContain('onEnd')
    expect(log).toContain('scroll:1')
  })

  it('a failed unit shows a typed error frame; next() still advances', async () => {
    const fd = fakeDoc({ count: 3 })
    const { s, root } = await show(fd)
    expect(frame(root)).toBe('loading')
    expect(root.querySelector('[data-officeview-show-frame-text]')?.textContent).toBe('Loading slide 1')
    await fd.paints[0]!.fail('unit-failed')
    expect(frame(root)).toBe('error')
    expect(root.querySelector('[data-officeview-show-frame-text]')?.textContent).toBe('Slide 1 could not be shown (unit-failed)')
    s.next()
    expect(s.state.slide).toBe(1)
    expect(frame(root)).toBe('loading')
    await clock.advance(400)
    await fd.paints.find((p) => p.unit === 1)!.ok()
    expect(frame(root)).toBe('image')
    s.exit()
  })

  it('a rate-limited refusal is no error frame (the fetcher re-issues it)', async () => {
    const fd = fakeDoc({ count: 2 })
    const { s, root } = await show(fd)
    await fd.paints[0]!.fail('rate-limited')
    expect(frame(root)).toBe('loading')
    await clock.advance(2_100)
    expect(units(fd)).toEqual([0, 0])
    s.exit()
  })

  it('io.paint threads a per-send charge hook into RemoteOfficeDoc.paint (R16)', async () => {
    const fd = fakeDoc({ count: 2 })
    const { s } = await show(fd)
    expect(typeof fd.paints[0]!.onSend).toBe('function')
    s.exit()
  })

  it('a looping single-slide show names that slide as Next, not a build (Task 5 minor)', async () => {
    const fd = fakeDoc({ count: 2, hidden: [1] })
    const { s, host } = await show(fd, { popup: null, loop: true })
    expect(host.querySelector('[data-officeview-presenter-next] figcaption')?.textContent).toBe('Next slide (1)')
    expect(host.querySelector('[data-officeview-presenter-next]')?.getAttribute('aria-label')).toBe('Next: slide 1')
    s.exit()
  })
})

describe('remote show: Presenter View in streamed indices', () => {
  it('the go-to field and the grid jump use streamed slide numbers', async () => {
    const fd = fakeDoc({ count: 4, hidden: [1] })
    const { s, host } = await show(fd, { popup: null })
    const field = host.querySelector<HTMLInputElement>('input[aria-label="Go to slide"]')!
    field.value = '3'
    field.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    await flush()
    expect(s.slideIndex).toBe(2)
    expect(status(host)).toBe('Slide 3 of 3')
    await clock.advance(400)
    expect(units(fd)).toContain(3)
    s.exit()
  })

  it('slideImage is cache-only and paints the streamed slide letterboxed', async () => {
    const fd = fakeDoc({ count: 3, hidden: [1] })
    const { s } = await show(fd, { popup: null })
    const c = document.createElement('canvas')
    c.width = 192; c.height = 108
    const before = fd.doc.paint.mock.calls.length
    expect(s.presenterSource.slideImage!(1, c)).toBe(false)
    expect(fd.doc.paint.mock.calls.length).toBe(before)
    await fd.paints[0]!.ok()
    await clock.advance(1_500)
    await fd.paints.find((p) => p.unit === 2)!.ok('#0f0')
    expect(s.presenterSource.slideImage!(1, c)).toBe(true)
    expect(px(c, 96, 54)).toEqual([0, 255, 0, 255])
    expect(s.presenterSource.hidden!(1)).toBe(false)
    s.exit()
  })
})

describe('remote show: blank screens reach Presenter View (PR #75)', () => {
  it('entering and leaving black/white notifies frame subscribers, so the Current preview matches the audience', async () => {
    const fd = fakeDoc({ count: 3 })
    const { s } = await show(fd)
    await fd.paints[0]!.ok()
    const seen: string[] = []
    s.presenterSource.onFrame(() => seen.push(s.state.screen))
    s.presenterSource.input({ kind: 'black' })
    expect(seen).toEqual(['black'])
    s.presenterSource.input({ kind: 'next' }) // any key restores the slide first
    expect(seen).toEqual(['black', 'slide'])
    expect(s.state.slide).toBe(0)
    s.presenterSource.input({ kind: 'white' })
    expect(seen).toEqual(['black', 'slide', 'white'])
    s.exit()
  })

  it.each([[-1], [3], [0.5]])('an invalid numbered goto (%p) from a blank screen restores the slide and notifies too', async (index) => {
    const fd = fakeDoc({ count: 3 })
    const { s } = await show(fd)
    await fd.paints[0]!.ok()
    const seen: string[] = []
    s.presenterSource.onFrame(() => seen.push(s.state.screen))
    s.presenterSource.input({ kind: 'black' })
    s.presenterSource.input({ kind: 'goto', index })
    expect(s.state.screen).toBe('slide')
    expect(seen.at(-1)).toBe('slide')
    s.exit()
  })
})

describe('remote show: end and revocation', () => {
  it('doc.ended exits the show; a later ready paints nothing', async () => {
    const fd = fakeDoc({ count: 3 })
    const { s, host, log, cleanup } = await show(fd)
    const p = fd.paints[0]!
    const frames = vi.fn()
    s.presenterSource.onFrame(frames)
    await fd.end(new StreamSessionError('revoked'))
    await s.ended
    expect(log).toContain('onEnd')
    expect(cleanup).toHaveBeenCalledTimes(1)
    expect(host.children.length).toBe(0)
    expect(p.signal?.aborted).toBe(true)
    await p.ok()
    // The late bitmap is closed (zeroed), never drawn.
    expect(p.canvas.width).toBe(0)
    expect(frames).not.toHaveBeenCalled()
    expect(fd.listeners.size).toBe(0)
    // Navigation after the end does nothing.
    s.next()
    expect(fd.paints.length).toBe(1)
  })

  it('exit() does not close the viewer-owned document', async () => {
    const fd = fakeDoc()
    const { s } = await show(fd)
    s.exit()
    await s.ended
    expect(fd.doc.close).not.toHaveBeenCalled()
  })
})

describe('remote show: notes', () => {
  it('an audience session (can.notes false) sends no notes request across a full run', async () => {
    const fd = fakeDoc({ count: 3, notes: false })
    const { s } = await show(fd, { popup: null })
    const empty = s.presenterSource.notes()
    expect(empty).toEqual([])
    for (let i = 0; i < 3; i++) {
      await clock.advance(1_500)
      for (const p of fd.paints.filter((x) => !x.signal?.aborted)) await p.ok()
      s.next()
    }
    await clock.advance(5_000)
    s.goTo(0)
    await clock.advance(5_000)
    expect(fd.doc.paintNotes).not.toHaveBeenCalled()
    expect(s.presenterSource.notes()).toBe(empty)
    s.exit()
    await s.ended
    expect(fd.doc.paintNotes).not.toHaveBeenCalled()
  })

  it('a presenter: notes only for the settled current slide, never on the audience canvas', async () => {
    const fd = fakeDoc({ count: 4, notes: true })
    const { s, canvas, host } = await show(fd, { popup: null })
    await fd.paints[0]!.ok('#00f')
    const n0 = s.presenterSource.notes()
    expect('kind' in n0 && n0.state).toBe('loading')
    expect(s.presenterSource.notes()).toBe(n0)
    // Passing slides quickly: only the slide the presenter stops on asks for notes.
    s.next(); s.next()
    await clock.advance(1_000)
    expect(fd.notes.map((n) => n.unit)).toEqual([2])
    const audience = all(canvas)
    await fd.notes[0]!.ok('#ff0')
    expect(all(canvas)).toEqual(audience)
    const ready = s.presenterSource.notes()
    expect('kind' in ready && ready.state).toBe('ready')
    expect(s.presenterSource.notes()).toBe(ready)
    const pane = host.querySelector<HTMLCanvasElement>('[data-officeview-presenter-notes] canvas')!
    expect(pane).not.toBeNull()
    expect(px(pane, 1, 1)).toEqual([255, 255, 0, 255])
    s.exit()
  })

  it('R11: a rate-limited notes refusal is retried once when the fetcher is idle; a second failure is an error', async () => {
    const fd = fakeDoc({ count: 2, notes: true })
    const { s, host } = await show(fd, { popup: null })
    await clock.advance(1_000)
    expect(fd.notes.length).toBe(1)
    await fd.notes[0]!.fail('rate-limited')
    // A unit is still in flight: the retry waits for it.
    expect(fd.notes.length).toBe(1)
    await fd.paints[0]!.ok()
    await clock.advance(10)
    // The prefetch that went out when slide 1 landed is in flight too: still waiting.
    expect(fd.notes.length).toBe(1)
    for (const p of fd.paints.filter((x) => !x.signal?.aborted && x.unit !== 0)) await p.ok()
    await clock.advance(10)
    expect(fd.notes.length).toBe(2)
    await fd.notes[1]!.fail('rate-limited')
    await clock.advance(5_000)
    for (const p of fd.paints.filter((x) => !x.signal?.aborted)) await p.ok()
    await clock.advance(5_000)
    expect(fd.notes.length).toBe(2)
    const n = s.presenterSource.notes()
    expect('kind' in n && n.state).toBe('error')
    expect(host.querySelector('[data-officeview-presenter-notes-status]')?.textContent).toBe('Notes could not be loaded')
    s.exit()
  })

  it('notes turned off by a resume: the pane empties and no further request goes out', async () => {
    const fd = fakeDoc({ count: 3, notes: true })
    const { s } = await show(fd, { popup: null })
    await clock.advance(1_000)
    await fd.notes[0]!.ok()
    fd.setManifest({ notes: false })
    expect(s.presenterSource.notes()).toEqual([])
    s.next()
    await clock.advance(5_000)
    expect(fd.notes.length).toBe(1)
    s.exit()
  })
})

describe('remote show: scale bucket', () => {
  it('the bucket follows the audience backing; a resize changes it', async () => {
    const fd = fakeDoc({ count: 2, buckets: [1, 2, 3] })
    const { s, canvas } = await show(fd)
    expect(fd.paints[0]!.scale).toBe(1)
    await fd.paints[0]!.ok('#f00')
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1920 })
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 1080 })
    window.dispatchEvent(new Event('resize'))
    await clock.advance(400)
    expect(canvas.width).toBe(1920)
    const again = fd.paints.filter((p) => p.unit === 0)
    expect(again.length).toBe(2)
    expect(again[1]!.scale).toBe(2)
    // Meanwhile the old bitmap is shown scaled.
    expect(px(canvas, 960, 540)).toEqual([255, 0, 0, 255])
    s.exit()
  })
})

// ---------------------------------------------------------------------------
// Fix round 1: retained surfaces are byte-admitted; re-entrant exit; presenter refresh
// ---------------------------------------------------------------------------

const SLIDE_B1 = 960 * 540 * 4
const SLIDE_B2 = 1920 * 1080 * 4

/** A deterministic workspace (no terminal scratch) the show's every retained surface must be leased from. */
function workspaceOf(limit: number) {
  const ws = createSlideShowWorkspace('desktop', { limit, scratch: 0 })
  vi.spyOn(remoteShowDeps, 'workspace').mockImplementation(() => ws)
  return ws
}

describe('remote show: retained surfaces are byte-admitted', () => {
  it('a full workspace refuses the slide surface before any request: typed error frame, still navigable', async () => {
    // Room for the audience output (960×540) and nothing else.
    workspaceOf(SLIDE_B1 + 1000)
    const fd = fakeDoc({ count: 3 })
    const { s, root } = await show(fd)
    expect(fd.doc.paint).not.toHaveBeenCalled()
    expect(frame(root)).toBe('error')
    expect(root.querySelector('[data-officeview-show-frame-text]')?.textContent).toBe('Slide 1 could not be shown (workspace-full)')
    s.next()
    expect(s.state.slide).toBe(1)
    s.exit()
  })

  it('a retained slide holds its bytes; exit returns the workspace to its pre-show usage', async () => {
    const ws = workspaceOf(512 * 1024 * 1024)
    const before = ws.used()
    const fd = fakeDoc({ count: 3 })
    const { s } = await show(fd)
    // Terminal output + the in-flight destination, reserved before the paint.
    expect(ws.used() - before).toBe(SLIDE_B1 + SLIDE_B1)
    await fd.paints[0]!.ok()
    expect(ws.used() - before).toBe(SLIDE_B1 + SLIDE_B1)
    s.exit()
    await s.ended
    expect(ws.used()).toBe(before)
  })

  it('an aborted request gives its bytes back at once', async () => {
    const ws = workspaceOf(512 * 1024 * 1024)
    const fd = fakeDoc({ count: 8 })
    const { s } = await show(fd)
    const p0 = fd.paints[0]!
    s.goTo(6)
    await clock.advance(400)
    expect(p0.signal?.aborted).toBe(true)
    // Terminal + the new target only: the aborted destination was released before its late result.
    expect(ws.used()).toBe(SLIDE_B1 + SLIDE_B1)
    await p0.ok()
    expect(ws.used()).toBe(SLIDE_B1 + SLIDE_B1)
    s.exit()
    await s.ended
    expect(ws.used()).toBe(0)
  })

  it('a bucket replacement holds old + new while the new one is in flight, then frees the old', async () => {
    const ws = workspaceOf(512 * 1024 * 1024)
    const fd = fakeDoc({ count: 1, buckets: [1, 2] })
    const { s } = await show(fd)
    await fd.paints[0]!.ok()
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1920 })
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 1080 })
    window.dispatchEvent(new Event('resize'))
    await clock.advance(400)
    // Terminal (now 1920×1080) + the old bitmap + the new destination.
    expect(ws.used()).toBe(SLIDE_B2 + SLIDE_B1 + SLIDE_B2)
    await fd.paints[1]!.ok()
    expect(ws.used()).toBe(SLIDE_B2 + SLIDE_B2)
    s.exit()
    await s.ended
    expect(ws.used()).toBe(0)
  })

  it('the cache is bounded by bytes: evicted slides release their leases', async () => {
    // Terminal + room for a small cache only.
    const ws = workspaceOf(SLIDE_B1 + 8 * SLIDE_B1)
    const fd = fakeDoc({ count: 10 })
    const { s } = await show(fd)
    let peak = 0
    for (let i = 0; i < 9; i++) {
      for (let k = 0; k < 4; k++) {
        await clock.advance(1_100)
        for (const p of fd.paints.filter((x) => !x.signal?.aborted && !x.settled)) await p.ok()
        peak = Math.max(peak, ws.used())
      }
      s.next()
    }
    expect(peak).toBeLessThanOrEqual(SLIDE_B1 + 8 * SLIDE_B1)
    // Bytes held equal the bitmaps retained (plus the terminal): nothing evicted stays charged.
    const held = fd.paints.filter((p) => p.canvas.width > 0).length
    const inFlight = fd.paints.filter((p) => !p.settled && !p.signal?.aborted).length
    expect(ws.used()).toBe(SLIDE_B1 + (held + inFlight) * SLIDE_B1)
    expect(held).toBeLessThan(9)
    s.exit()
    await s.ended
    expect(ws.used()).toBe(0)
  })

  it('a notes image and its Presenter View copy are leased; dropping the notes releases both', async () => {
    const ws = workspaceOf(512 * 1024 * 1024)
    const fd = fakeDoc({ count: 2, notes: true })
    const { s } = await show(fd, { popup: null })
    await fd.paints[0]!.ok()
    await clock.advance(1_000)
    const pending = ws.used()
    await fd.notes[0]!.ok()
    const NOTES = 720 * 100 * 4
    const withNotes = ws.used()
    // The worst-case reservation shrank to the delivered image plus its pane copy.
    expect(withNotes).toBeLessThan(pending)
    fd.setManifest({ notes: false })
    expect(withNotes - ws.used()).toBe(2 * NOTES)
    s.exit()
    await s.ended
    expect(ws.used()).toBe(0)
  })
})

describe('remote show: exit inside a notification loop', () => {
  it('exit() from a frame subscriber stops the rest of the frame dispatch', async () => {
    const fd = fakeDoc({ count: 2 })
    const { s } = await show(fd)
    const later = vi.fn()
    s.presenterSource.onFrame(() => s.exit())
    s.presenterSource.onFrame(later)
    await fd.paints[0]!.ok()
    expect(s.state.exited).toBe(true)
    expect(later).not.toHaveBeenCalled()
  })

  it('a frame subscriber removed by an earlier one is not called', async () => {
    const fd = fakeDoc({ count: 2 })
    const { s } = await show(fd)
    const later = vi.fn()
    let off = () => {}
    s.presenterSource.onFrame(() => off())
    off = s.presenterSource.onFrame(later)
    await fd.paints[0]!.ok()
    expect(later).not.toHaveBeenCalled()
    s.exit()
  })

  it('exit() from a view subscriber stops the rest of the change dispatch', async () => {
    const fd = fakeDoc({ count: 2 })
    const { s } = await show(fd)
    const later = vi.fn()
    s.presenterSource.subscribe(() => s.exit())
    s.presenterSource.subscribe(later)
    s.presenterSource.touch()
    expect(s.state.exited).toBe(true)
    expect(later).not.toHaveBeenCalled()
  })
})

describe('remote show: presenter refresh', () => {
  it('Alt+F5 into Presenter View starts the 1 Hz refresh once; exit stops it', async () => {
    const stop = vi.fn()
    const every = vi.spyOn(remoteShowDeps, 'every').mockImplementation(() => stop)
    const fd = fakeDoc({ count: 2 })
    const { s, host } = await show(fd)
    expect(every).not.toHaveBeenCalled()
    const altF5 = () => host.dispatchEvent(new KeyboardEvent('keydown', { key: 'F5', altKey: true, bubbles: true, cancelable: true }))
    altF5()
    expect(every).toHaveBeenCalledTimes(1)
    altF5(); altF5()
    expect(every).toHaveBeenCalledTimes(1)
    s.exit()
    expect(stop).toHaveBeenCalledTimes(1)
  })
})

describe('remote show: the notes pane copy is owned with its bytes', () => {
  const NOTES = 720 * 100 * 4
  const paneOf = (host: HTMLElement) => host.querySelector<HTMLCanvasElement>('[data-officeview-presenter-notes] canvas')!
  async function withNotes(count = 3) {
    const ws = workspaceOf(512 * 1024 * 1024)
    const fd = fakeDoc({ count, notes: true })
    const r = await show(fd, { popup: null })
    await fd.paints[0]!.ok()
    await clock.advance(1_000)
    await fd.notes[0]!.ok()
    const pane = paneOf(r.host)
    expect([pane.width, pane.height]).toEqual([720, 100])
    return { ws, fd, pane, ...r }
  }

  it('permission withdrawal destroys the pane copy as its bytes are refunded', async () => {
    const { ws, fd, pane, s } = await withNotes()
    const before = ws.used()
    fd.setManifest({ notes: false })
    expect(before - ws.used()).toBe(2 * NOTES)
    expect([pane.width, pane.height]).toEqual([0, 0])
    s.exit()
  })

  it('a slide change destroys the pane copy', async () => {
    const { pane, s } = await withNotes()
    s.next()
    expect([pane.width, pane.height]).toEqual([0, 0])
    s.exit()
  })

  it('closing and reopening Presenter View keeps one accounted pane copy', async () => {
    const { ws, pane, s, host } = await withNotes()
    const used = ws.used()
    const altF5 = () => host.dispatchEvent(new KeyboardEvent('keydown', { key: 'F5', altKey: true, bubbles: true, cancelable: true }))
    altF5(); altF5()
    await flush()
    const again = paneOf(host)
    expect(again).not.toBe(pane)
    expect([again.width, again.height]).toEqual([720, 100])
    expect([pane.width, pane.height]).toEqual([0, 0])
    // Presenter View's own surfaces come back to the same size: only the notes copies are compared.
    expect(ws.used()).toBe(used)
    s.exit()
  })

  it('exit destroys the pane copy before refunding it', async () => {
    const { ws, pane, s } = await withNotes()
    s.exit()
    await s.ended
    expect(ws.used()).toBe(0)
    expect([pane.width, pane.height]).toEqual([0, 0])
  })
})

describe('R28: the show paces against the session\'s unit budget', () => {
  it('a doc whose session budget has no room paints nothing; once it has room the target goes out', async () => {
    const fd = fakeDoc({ count: 3 })
    let room = 0
    Object.assign(fd.doc, { unitBudget: { room: (pending = 0) => Math.max(0, room - pending), waitMs: () => (room > 0 ? 0 : 1_000) } })
    const { s } = await show(fd, { popup: null })
    await clock.advance(5_000)
    expect(units(fd)).toEqual([])
    room = 48
    await clock.advance(1_000)
    expect(units(fd)).toEqual([0])
    s.exit()
  })
})
