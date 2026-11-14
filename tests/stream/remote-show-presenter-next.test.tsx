/**
 * Piece 4 Task 13 (spec 5.2): Presenter View "Next" for a streamed animated show. Within a slide it is the next
 * build checkpoint composed on the client from the streamed layers (`PlanStage.composeAt`); at the slide's end it
 * is the next slide's entry frame; with no workspace room it is the entry frame plus a `preview-reduced`
 * diagnostic. Driven over the Task 9 recorded bundles through the Task 11 fake session (jsdom + node-canvas).
 */
import { createCanvas, loadImage } from 'canvas'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { remoteShowDeps, runRemoteShow, type RemoteShowStart } from '../../src/stream/remote/show'
import { StreamSessionError, type StreamErrorCode } from '../../src/stream/remote/errors'
import type { Manifest } from '../../src/stream/manifest'
import type { RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'
import type { ShowReply, ShowRequestOptions } from '../../src/stream/remote/session'
import { decodeShowMessage, encodeShowMessage, SHOW_PLAN_LIMITS, type ShowMessage, type ShowPlanRecord } from '../../src/stream/show-plan'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'
import type { TimeNode } from '../../src/slideshow/types'

interface Rec { op: number; unit: number; bytes: Uint8Array }
function records(name: string): Rec[] {
  const buf = new Uint8Array(readFileSync(resolve(__dirname, '../fixtures/show-plan', `${name}.bin`)))
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  const out: Rec[] = []
  for (let at = 0; at < buf.length;) {
    const len = v.getUint32(at + 13, true)
    out.push({ op: v.getUint8(at), unit: v.getUint32(at + 1, true), bytes: buf.subarray(at + 17, at + 17 + len) })
    at += 17 + len
  }
  return out
}
const decodeAll = (rs: Rec[]) => rs.map((r) => decodeShowMessage(r.bytes, { unit: r.unit, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }))
const FLY = decodeAll(records('fly').filter((r) => r.op === 1))
const MORPH_RECS = records('morph')
/** The Morph target slide's bundle (a transition-only slide: plan + entry frame). */
const MORPH_TO = decodeAll(MORPH_RECS.filter((r) => r.op === 1))
/** The recorded layered pair (0 → 1 at 0): the tag-5 record and its four assets. */
const PAIR = decodeAll(MORPH_RECS.filter((r) => r.op === 3))

type Plan = ShowPlanRecord
/** A recorded bundle re-addressed to `unit` (the plan's index and every asset's unit), the plan optionally edited. */
function bundleOf(src: ShowMessage[], unit: number, edit: (p: Plan) => Plan = (p) => p): Uint8Array[] {
  return src.map((m) => (m.tag === 1
    ? encodeShowMessage({ tag: 1, plan: edit({ ...m.plan, slide: { ...m.plan.slide, index: unit } }) })
    : encodeShowMessage({ ...(m as Extract<ShowMessage, { tag: 2 | 3 | 4 }>), unit })))
}
/** The recorded pair re-addressed to `(from → to @ fromStep)`. */
function pairOf(from: number, to: number, fromStep: number): Uint8Array[] {
  return PAIR.map((m) => (m.tag === 5
    ? encodeShowMessage({ tag: 5, morph: { ...m.morph, from, to, fromStep } })
    : encodeShowMessage({ ...(m as Extract<ShowMessage, { tag: 2 | 3 | 4 }>), unit: to })))
}

/** The fly-in timing with its one click step repeated `n` times (fresh node ids). */
function flyTiming(n: number): Plan['show']['timing'] {
  const plan = (FLY[0] as Extract<ShowMessage, { tag: 1 }>).plan
  const timing = plan.show.timing!
  if ('refused' in timing) throw new Error('refused')
  const root = structuredClone(timing.root) as TimeNode & { children: TimeNode[] }
  const main = root.children[0] as TimeNode & { children: TimeNode[] }
  const click = main.children[0]!
  const renumber = (node: TimeNode, k: number): TimeNode => {
    const c = structuredClone(node) as TimeNode
    const walk = (x: TimeNode) => { (x.ctn as { id: number }).id += 1000 * k; if ('children' in x) for (const y of x.children) walk(y) }
    walk(c)
    return c
  }
  main.children = Array.from({ length: n }, (_, k) => (k === 0 ? click : renumber(click, k)))
  return { ...timing, root }
}

const FRAME_W = 96, FRAME_H = 72
const COLORS = ['#ff0000', '#00ff00', '#0000ff']
const png = (color: string, w = FRAME_W, h = FRAME_H) => {
  const c = createCanvas(w, h)
  const ctx = c.getContext('2d')
  ctx.fillStyle = color
  ctx.fillRect(0, 0, w, h)
  return new Uint8Array(c.toBuffer('image/png'))
}
/** A fly slide whose `steps` click steps are all composited, each with a 3-frame keyframe track. */
const compositedFly = (steps: number) => (p: Plan): Plan => ({
  ...p,
  show: { ...p.show, timing: flyTiming(steps) },
  steps: Array.from({ length: steps }, () => ({ trigger: 'click' as const, eligibility: 'composited' as const, track: { fps: 15 as const, count: 3 } })),
})
const trackOf = (unit: number, step: number) => [0, 1, 2].map((i) => encodeShowMessage({ tag: 4, unit, assetId: i, w: FRAME_W, h: FRAME_H, png: png(COLORS[step % 3]!) }))

// ---------------------------------------------------------------------------
// The fake streamed document and session
// ---------------------------------------------------------------------------

type Kind = 'still' | 'fly' | 'morphTo' | { messages: (unit: number) => Uint8Array[] }
type Rule = 'hold' | StreamErrorCode | (() => Uint8Array[])

const flush = async () => { for (let i = 0; i < 6; i++) await new Promise((r) => setTimeout(r, 0)) }

function animDoc(deck: Kind[], o: { tracks?: (unit: number, step: number) => Uint8Array[]; notes?: boolean } = {}) {
  const list = deck.map((k) => (k === 'still' ? { w: 960, h: 720 } : { w: 960, h: 720, anim: true as const }))
  const manifest = {
    units: { count: deck.length, list },
    can: { search: false, copy: false, print: false, textA11y: false, present: true, notes: !!o.notes, links: false },
    recorder: { scalesBuckets: [1, 2] },
  } as unknown as Manifest
  const log: string[] = []
  const rules = new Map<string, Rule>()
  const held = new Map<string, Array<() => void>>()
  const signals = new Map<string, AbortSignal>()
  let endIt!: (e: StreamSessionError | undefined) => void
  const ended = new Promise<StreamSessionError | undefined>((r) => { endIt = r })
  const messagesFor = (unit: number): Uint8Array[] => {
    const k = deck[unit]!
    if (k === 'fly') return bundleOf(FLY, unit)
    if (k === 'morphTo') return bundleOf(MORPH_TO, unit)
    if (k === 'still') throw new Error(`unit ${unit} is a still`)
    return k.messages(unit)
  }
  /** One show reply: logged, charged once, each message leased at its bytes in the show workspace (as the session). */
  function reply(key: string, sig: AbortSignal, o: ShowRequestOptions, make: () => Uint8Array[]): Promise<ShowReply> {
    log.push(key)
    signals.set(key, sig)
    o.onSend?.()
    const rule = rules.get(key)
    return new Promise<ShowReply>((res, rej) => {
      const answer = () => {
        if (sig.aborted) { rej(sig.reason ?? new Error('aborted')); return }
        const r = rules.get(key)
        if (typeof r === 'string' && r !== 'hold') { rej(new StreamSessionError(r)); return }
        const messages = typeof r === 'function' ? r() : make()
        const leases = messages.map((m) => o.workspace.lease(m.length, 'remote-show:message'))
        if (leases.some((l) => !l)) { for (const l of leases) l?.release(); rej(new StreamSessionError('workspace-full')); return }
        let freed = false
        res({ messages, release: () => { if (freed) return; freed = true; for (const l of leases) l!.release() } })
      }
      sig.addEventListener('abort', () => rej(sig.reason ?? new Error('aborted')), { once: true })
      if (rule === 'hold') { const q = held.get(key) ?? []; q.push(answer); held.set(key, q) } else queueMicrotask(answer)
    })
  }
  const session = {
    showPlan: vi.fn((unit: number, scale: number, motion: string, sig: AbortSignal, so: ShowRequestOptions) => {
      log.push(`motion:${motion}@${scale}`)
      return reply(`plan:${unit}`, sig, so, () => messagesFor(unit))
    }),
    showTrack: vi.fn((unit: number, step: number, _scale: number, _motion: string, sig: AbortSignal, so: ShowRequestOptions) =>
      reply(`track:${unit}:${step}`, sig, so, () => (o.tracks ?? trackOf)(unit, step))),
    showMorph: vi.fn((from: number, to: number, fromStep: number, _scale: number, _motion: string, sig: AbortSignal, so: ShowRequestOptions) =>
      reply(`morph:${from}>${to}@${fromStep}`, sig, so, () => pairOf(from, to, fromStep))),
  }
  const doc = {
    kind: 'remote',
    get manifest() { return manifest },
    session,
    paint: vi.fn(async (unit: number, scale: number | 'fit', dpr: number, canvas: HTMLCanvasElement, signal?: AbortSignal, po?: { onSend?: () => void }) => {
      log.push(`paint:${unit}`)
      po?.onSend?.()
      await Promise.resolve()
      if (signal?.aborted) return
      const k = (typeof scale === 'number' ? scale : 1) * dpr
      canvas.width = 960 * k; canvas.height = 720 * k
      const ctx = canvas.getContext('2d')!
      ctx.fillStyle = '#123456'
      ctx.fillRect(0, 0, canvas.width, canvas.height)
    }),
    paintNotes: vi.fn(() => Promise.reject(new StreamSessionError('policy-refused'))),
    onManifestChange() { return () => {} },
    ended,
    close: vi.fn(),
  } as unknown as RemoteOfficeDoc
  return {
    doc, session, log, rules, signals,
    release(key: string) { const q = held.get(key) ?? []; held.delete(key); for (const f of q) f() },
    end: async (e?: StreamSessionError) => { endIt(e); await flush() },
  }
}

/** The fetcher, notes and show clock: time moves only through `advance`. */
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
    advance(ms: number) {
      const end = t + ms
      for (;;) {
        const due = timers.filter((x) => x.at <= end).sort((a, b) => a.at - b.at)[0]
        if (!due) break
        timers = timers.filter((x) => x !== due)
        t = Math.max(t, due.at)
        due.cb()
      }
      t = end
    },
  }
}

let clock: ReturnType<typeof manualClock>
let frame: (() => void) | undefined
const bitmapOf = async (bytes: Uint8Array) => {
  const img = await loadImage(Buffer.from(bytes))
  const cv = document.createElement('canvas')
  cv.width = img.width; cv.height = img.height
  cv.getContext('2d')!.drawImage(img as unknown as CanvasImageSource, 0, 0)
  return cv
}
beforeEach(() => {
  clock = manualClock()
  frame = undefined
  vi.spyOn(remoteShowDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(remoteShowDeps, 'every').mockImplementation(() => () => {})
  vi.spyOn(remoteShowDeps, 'frames').mockImplementation((_w, f) => { frame = f; return () => { frame = undefined } })
  vi.spyOn(remoteShowDeps, 'decode').mockImplementation(bitmapOf)
  remoteShowDeps.backend = 'canvas2d'
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: 960 })
  Object.defineProperty(window, 'innerHeight', { configurable: true, value: 540 })
})
afterEach(() => {
  remoteShowDeps.backend = undefined
  vi.restoreAllMocks()
  document.body.replaceChildren()
})

/** Runs the show `ms` at ~60 fps: the clock, then one frame, letting decodes settle. */
async function run(ms: number) {
  for (let t = 0; t < ms; t += 16) {
    clock.advance(16)
    frame?.()
    await flush()
  }
}

async function show(fd: ReturnType<typeof animDoc>, o: Partial<RemoteShowStart> = {}) {
  const host = document.createElement('div')
  host.tabIndex = -1
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const start: RemoteShowStart = {
    doc: fd.doc, host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'),
    retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'full',
    callbacks: {
      onStart: () => log.push('onStart'),
      onEnd: () => log.push('onEnd'),
      onSlideChange: (i, b) => log.push(`slide:${i}:${b}`),
      onDiagnostic: (d) => log.push(`diag:${d.code}`),
      currentSlide: () => 0, restoreFocus() {}, scrollToSlide: () => {},
    },
    signal: ac.signal, generation: 1, cleanup: () => { ac.abort() },
    ...o,
  }
  const s = await runRemoteShow(start)
  await flush()
  const canvas = host.querySelector<HTMLCanvasElement>('canvas[data-officeview-show-canvas]')!
  const root = host.querySelector<HTMLElement>('[data-officeview-show-surface]')!
  root.getBoundingClientRect = () => ({ left: 0, top: 0, width: canvas.width, height: canvas.height, right: canvas.width, bottom: canvas.height, x: 0, y: 0, toJSON: () => ({}) })
  return { s, host, log, canvas, root }
}

const px = (c: HTMLCanvasElement, x: number, y: number) => [...c.getContext('2d')!.getImageData(x, y, 1, 1).data].slice(0, 3)
const pixels = (c: HTMLCanvasElement) => c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data

/** A preview canvas the size of the audience backing. */
const previewOf = (like: HTMLCanvasElement) => {
  const c = document.createElement('canvas')
  c.width = like.width; c.height = like.height
  return c
}
/** Whether two canvases hold the same pixels (alpha included). */
const same = (a: HTMLCanvasElement, b: HTMLCanvasElement) => {
  if (a.width !== b.width || a.height !== b.height) return false
  const x = pixels(a), y = pixels(b)
  for (let i = 0; i < x.length; i++) if (x[i] !== y[i]) return false
  return true
}
const mid = (c: HTMLCanvasElement) => px(c, c.width >> 1, c.height >> 1)

describe('Presenter View Next on a streamed animated show (piece 4 Task 13)', () => {
  it('Next within a slide shows the next checkpoint composed from streamed layers', async () => {
    const fd = animDoc(['fly', 'fly'])
    const { s, canvas } = await show(fd, { popup: null })
    await run(1_500)
    const src = s.presenterSource
    expect(typeof src.nextImage).toBe('function')
    const next = previewOf(canvas)
    expect(src.nextImage!(next)).toBe(true)
    // The entry frame is what the slide shows before the click; the composed checkpoint is not it.
    const entry = previewOf(canvas)
    expect(src.slideImage!(0, entry)).toBe(true)
    expect(same(next, entry)).toBe(false)
    // The click plays the step; once it settles the audience shows what Next showed.
    s.next()
    await run(2_000)
    expect(s.state.slide).toBe(0)
    expect(same(next, canvas)).toBe(true)
    s.exit()
  })

  it('Next within a composited step shows the step\'s last keyframe', async () => {
    const three = { messages: (u: number) => bundleOf(FLY, u, compositedFly(3)) }
    const fd = animDoc([three, three])
    const { s, canvas } = await show(fd, { popup: null })
    await run(1_500)
    const next = previewOf(canvas)
    expect(s.presenterSource.nextImage!(next)).toBe(true)
    expect(mid(next)).toEqual([255, 0, 0])
    s.next()
    await run(600)
    const after = previewOf(canvas)
    expect(s.presenterSource.nextImage!(after)).toBe(true)
    expect(mid(after)).toEqual([0, 255, 0])
    s.exit()
  })

  it('Next at slide end shows the next slide\'s entry frame', async () => {
    const fd = animDoc(['fly', 'fly'])
    const { s, canvas } = await show(fd, { popup: null })
    await run(1_500)
    s.next()
    await run(2_000)
    expect(s.state.slide).toBe(0)
    const src = s.presenterSource
    const next = previewOf(canvas)
    expect(src.nextImage!(next)).toBe(true)
    const entry = previewOf(canvas)
    expect(src.slideImage!(1, entry)).toBe(true)
    expect(same(next, entry)).toBe(true)
    // And it is not the settled checkpoint of the slide being shown.
    expect(same(next, canvas)).toBe(false)
    expect(fd.doc.paintNotes).not.toHaveBeenCalled()
    s.exit()
  })

  it('no workspace room: the entry frame plus a preview-reduced diagnostic', async () => {
    const ws = createSlideShowWorkspace('desktop')
    vi.spyOn(remoteShowDeps, 'workspace').mockReturnValue(ws)
    let refuse = false
    const lease = ws.lease.bind(ws)
    vi.spyOn(ws, 'lease').mockImplementation((n, tag) => (refuse && tag === 'plan:compose' ? undefined : lease(n, tag)))
    const fd = animDoc(['fly', 'fly'])
    const { s, canvas } = await show(fd, { popup: null })
    await run(1_500)
    const src = s.presenterSource
    expect(src.diagnostics().map((d) => d.code)).not.toContain('preview-reduced')
    refuse = true
    const next = previewOf(canvas)
    expect(src.nextImage!(next)).toBe(true)
    const entry = previewOf(canvas)
    expect(src.slideImage!(1, entry)).toBe(true)
    expect(same(next, entry)).toBe(true)
    expect(src.diagnostics().filter((d) => d.code === 'preview-reduced')).toHaveLength(1)
    // Asked again, the diagnostic is the same one (the sink deduplicates).
    expect(src.nextImage!(previewOf(canvas))).toBe(true)
    expect(src.diagnostics().filter((d) => d.code === 'preview-reduced')).toHaveLength(1)
    s.exit()
  })

  it('a click repaints the Presenter View Next canvas (the view hears the build change)', async () => {
    const three = { messages: (u: number) => bundleOf(FLY, u, compositedFly(3)) }
    const fd = animDoc([three, three])
    const { s, canvas, host } = await show(fd, { popup: null })
    await run(1_500)
    const pane = host.querySelector<HTMLCanvasElement>('[data-officeview-presenter-next] canvas')!
    expect(pane).not.toBeNull()
    expect(mid(pane)).toEqual([255, 0, 0])
    s.next()
    await run(600)
    expect(mid(pane)).toEqual([0, 255, 0])
    expect(canvas).toBeTruthy()
    s.exit()
  })

  it('audience role: nextImage never paints notes', async () => {
    const fd = animDoc(['fly', 'fly'], { notes: true })
    // No popup: no presenter, so the notes are never requested; Next is slide pixels only.
    const { s, canvas } = await show(fd)
    await run(1_500)
    const src = s.presenterSource
    expect(src.notes()).toEqual([])
    const next = previewOf(canvas)
    expect(src.nextImage!(next)).toBe(true)
    expect(src.notes()).toEqual([])
    expect(fd.doc.paintNotes).not.toHaveBeenCalled()
    s.exit()
  })

  it('an unanimated deck has no nextImage: Next stays the piece-3 still path', async () => {
    const fd = animDoc(['still', 'still'])
    const { s } = await show(fd, { popup: null })
    await run(200)
    expect(s.presenterSource.nextImage).toBeUndefined()
    s.exit()
  })
})
