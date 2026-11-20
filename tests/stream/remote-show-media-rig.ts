/**
 * Piece 5 Task 11: the shared rig of the remote-show media tests. A fake streamed document whose session serves the
 * recorded media-deck bundles (`tests/fixtures/show-plan/media-<name>.bin`, `scripts/stream-p5/record-media-show-
 * fixtures.ts`: real server plans with media on, their session media ids made deterministic) and answers the piece-5
 * media transport (`mediaShow`, `mediaCursor`, `mediaRequest`, `mediaEnd`) by rule; jsdom media elements whose
 * `play`/`pause`/`currentTime` are recorded; Blob URLs counted; the show clock, fetcher pacing and frame loop on a
 * manual clock (as `remote-show-anim.test.tsx`).
 */
import { loadImage } from 'canvas'
import { AsyncLocalStorage } from 'node:async_hooks'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { vi, it as vitestIt } from 'vitest'
import { remoteShowDeps, runRemoteShow, type RemoteShowStart } from '../../src/stream/remote/show'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { MEDIA_AUDIO_CAPS, type Manifest, type ManifestUnit } from '../../src/stream/manifest'
import type { RemoteOfficeDoc } from '../../src/stream/remote/remote-doc'
import type { MediaDelivery, MediaRoom, ShowReply, ShowRequestOptions } from '../../src/stream/remote/session'
import { decodeShowMessage, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../src/stream/show-plan'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'
import type { SlideShowWorkspace } from '../../src/slideshow/types'

// ---------------------------------------------------------------------------
// The recorded media decks
// ---------------------------------------------------------------------------

export type MediaFixtureName = 'narration' | 'effect-sound' | 'transition-sound' | 'cross-slide' | 'video' | 'unrelated' | 'eight-parts' | 'refused' | 'still-sound' | 'still-transition'
export interface MediaFixture { list: ManifestUnit[]; plans: Map<number, Uint8Array[]>; plan(unit: number): ShowPlanRecord }
export function mediaFixture(name: MediaFixtureName): MediaFixture {
  const buf = new Uint8Array(readFileSync(resolve(__dirname, '../fixtures/show-plan', `media-${name}.bin`)))
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  let list: ManifestUnit[] = []
  const plans = new Map<number, Uint8Array[]>()
  for (let at = 0; at < buf.length;) {
    const op = v.getUint8(at), unit = v.getUint32(at + 1, true), len = v.getUint32(at + 13, true)
    const bytes = buf.subarray(at + 17, at + 17 + len)
    if (op === 0) list = (JSON.parse(new TextDecoder().decode(bytes)) as { list: ManifestUnit[] }).list
    else if (op === 1) { const q = plans.get(unit) ?? []; q.push(bytes); plans.set(unit, q) }
    at += 17 + len
  }
  return {
    list, plans,
    plan: (unit) => (decodeShowMessage(plans.get(unit)![0]!, { unit, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan,
  }
}

/** Six macrotask turns (see `remote-show-anim.test.tsx`). */
export const flush = async () => { for (let i = 0; i < 6; i++) await new Promise((r) => setImmediate(r)) }

// ---------------------------------------------------------------------------
// Media elements and Blob URLs (jsdom has neither playback nor createObjectURL)
// ---------------------------------------------------------------------------

export interface ElementLog { el: HTMLMediaElement; at: number }
export interface MediaElements {
  plays: ElementLog[]
  /** Elements whose last call was play() (and that were not paused, unloaded or removed since). */
  playing(): HTMLMediaElement[]
  created: string[]
  revoked: string[]
  restore(): void
}
export function installMedia(o: { canPlay?: string } = {}): MediaElements {
  const P = HTMLMediaElement.prototype
  const times = new WeakMap<HTMLMediaElement, number>()
  const live = new Set<HTMLMediaElement>()
  const out: MediaElements = { plays: [], playing: () => [...live].filter((e) => e.isConnected), created: [], revoked: [], restore() {} }
  const desc = Object.getOwnPropertyDescriptor(P, 'currentTime')
  Object.defineProperty(P, 'currentTime', { configurable: true, get(this: HTMLMediaElement) { return times.get(this) ?? 0 }, set(this: HTMLMediaElement, x: number) { times.set(this, x) } })
  vi.spyOn(P, 'canPlayType').mockImplementation(() => (o.canPlay ?? 'probably') as CanPlayTypeResult)
  vi.spyOn(P, 'play').mockImplementation(function (this: HTMLMediaElement) { out.plays.push({ el: this, at: times.get(this) ?? 0 }); live.add(this); return Promise.resolve() })
  vi.spyOn(P, 'pause').mockImplementation(function (this: HTMLMediaElement) { live.delete(this) })
  vi.spyOn(P, 'load').mockImplementation(function (this: HTMLMediaElement) { live.delete(this) })
  const U = window.URL as unknown as { createObjectURL?: (b: Blob) => string; revokeObjectURL?: (u: string) => void }
  const had = { c: U.createObjectURL, r: U.revokeObjectURL }
  let n = 0
  U.createObjectURL = () => { const u = `blob:media-${n++}`; out.created.push(u); return u }
  U.revokeObjectURL = (u: string) => { out.revoked.push(u) }
  out.restore = () => {
    if (desc) Object.defineProperty(P, 'currentTime', desc)
    else delete (P as unknown as Record<string, unknown>).currentTime
    U.createObjectURL = had.c; U.revokeObjectURL = had.r
  }
  return out
}

// ---------------------------------------------------------------------------
// The fake streamed document
// ---------------------------------------------------------------------------

/** How a media-request is answered: delivered (`ok`, the default), held until `deliver(id)`, or refused. */
export type MediaRule = 'ok' | 'hold' | { refused: string }
/** The transport a show uses besides show plans (a real StreamSession's, for the pacing test). */
export type MediaTransport = Pick<Required<import('../../src/stream/remote/remote-doc').SessionLike>, 'mediaShow' | 'mediaCursor' | 'mediaRequest' | 'mediaEnd'>

export function mediaDoc(fx: MediaFixture, o: { media?: boolean; notes?: boolean; transport?: MediaTransport; planRule?: (unit: number) => string | undefined } = {}) {
  const media = o.media ?? true
  const manifest = {
    units: { count: fx.list.length, list: fx.list },
    can: { search: false, copy: false, print: false, textA11y: false, present: true, notes: !!o.notes, links: false, media },
    caps: { copy: null, a11y: 'off', print: null, ...(media ? { media: { audio: { ...MEDIA_AUDIO_CAPS } } } : {}) },
    recorder: { scalesBuckets: [1, 2] },
  } as unknown as Manifest
  const log: string[] = []
  const rules = new Map<string, MediaRule>()
  const holds = new Map<string, Array<() => void>>()
  const live = new Set<(d: MediaDelivery) => void>()
  let endIt!: (e: StreamSessionError | undefined) => void
  const ended = new Promise<StreamSessionError | undefined>((r) => { endIt = r })
  const FILE = 4096
  function reply(key: string, sig: AbortSignal, so: ShowRequestOptions, make: () => Uint8Array[]): Promise<ShowReply> {
    log.push(key)
    so.onSend?.()
    return new Promise<ShowReply>((res, rej) => {
      queueMicrotask(() => {
        if (sig.aborted) { rej(sig.reason ?? new Error('aborted')); return }
        let messages: Uint8Array[]
        try { messages = make() } catch (e) { rej(e); return }
        const leases = messages.map((m) => so.workspace.lease(m.length, 'remote-show:message'))
        if (leases.some((l) => !l)) { for (const l of leases) l?.release(); rej(new StreamSessionError('workspace-full')); return }
        res({ messages, release: () => { for (const l of leases.splice(0)) l!.release() } })
      })
    })
  }
  const fake: MediaTransport = {
    mediaShow: vi.fn(async (unit: number) => { log.push(`media-show:${unit}`); return media ? { queued: 0 } : { refused: 'media-unavailable' } }),
    mediaCursor: vi.fn((unit: number) => { log.push(`media-cursor:${unit}`) }),
    mediaEnd: vi.fn(() => { log.push('media-end'); for (const r of [...live]) r({ refused: 'media-unavailable' }) }),
    mediaRequest: vi.fn((mediaId: string, sig: AbortSignal, room: MediaRoom, unit?: number) => {
      log.push(`media-request:${mediaId}@${unit}`)
      if (!media) return Promise.resolve({ refused: 'media-unavailable' })
      if (!room.reserve(MEDIA_AUDIO_CAPS.maxBytes)) return Promise.resolve({ refused: 'resource-limit' })
      return new Promise<MediaDelivery>((res, rej) => {
        let done = false
        const settle = (d: MediaDelivery) => { if (done) return; done = true; live.delete(settle); if ('refused' in d) room.release(); res(d) }
        live.add(settle)
        sig.addEventListener('abort', () => { if (done) return; done = true; live.delete(settle); room.release(); rej(sig.reason ?? new Error('aborted')) }, { once: true })
        const answer = () => {
          const r = rules.get(mediaId) ?? 'ok'
          if (typeof r === 'object') { settle(r); return }
          const start = { reqId: 1, type: 'media-start' as const, mediaId, contentType: 'audio/mp4' as const, durationMs: 1000, bytes: FILE, chunks: 1 }
          room.fit(start)
          settle({ chunks: [new Uint8Array(FILE)], start })
        }
        if ((rules.get(mediaId) ?? 'ok') === 'hold') { const q = holds.get(mediaId) ?? []; q.push(answer); holds.set(mediaId, q) } else queueMicrotask(answer)
      })
    }),
  }
  const transport = o.transport ?? fake
  const session = {
    showPlan: vi.fn((unit: number, _scale: number, _motion: string, sig: AbortSignal, so: ShowRequestOptions) => reply(`plan:${unit}`, sig, so, () => {
      const code = o.planRule?.(unit)
      if (code) throw new StreamSessionError(code as never)
      const m = fx.plans.get(unit)
      if (!m) throw new StreamSessionError('still' as never)
      return m
    })),
    showTrack: vi.fn(() => Promise.reject(new StreamSessionError('policy-refused'))),
    showMorph: vi.fn(() => Promise.reject(new StreamSessionError('policy-refused'))),
    mediaShow: (u: number) => transport.mediaShow(u),
    mediaCursor: (u: number) => transport.mediaCursor(u),
    mediaRequest: (id: string, sig: AbortSignal, room: MediaRoom, unit?: number) => transport.mediaRequest(id, sig, room, unit),
    mediaEnd: () => transport.mediaEnd(),
  }
  const manifestListeners = new Set<() => void>()
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
    onManifestChange(cb: () => void) { manifestListeners.add(cb); return () => { manifestListeners.delete(cb) } },
    ended,
    close: vi.fn(),
  } as unknown as RemoteOfficeDoc
  return {
    doc, session, fake, log, rules, manifest,
    deliver(id: string) { const q = holds.get(id) ?? []; holds.delete(id); for (const f of q) f() },
    requests: () => log.filter((l) => l.startsWith('media-request:')),
    resumed() { for (const l of [...manifestListeners]) l() },
    end: async (e?: StreamSessionError) => { endIt(e); await flush() },
  }
}
export type MediaDoc = ReturnType<typeof mediaDoc>

// ---------------------------------------------------------------------------
// The manual clock and the per-test owner (as remote-show-anim.test.tsx)
// ---------------------------------------------------------------------------

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
interface Owner { clock: ReturnType<typeof manualClock>; frame?: () => void; live: boolean; media?: MediaElements; workspace?: SlideShowWorkspace }
const owners = new AsyncLocalStorage<Owner>()
let current: Owner | undefined
function own(): Owner {
  const o = owners.getStore()
  if (!o || !o.live || o !== current) throw new Error('run: the test that started this run has ended')
  return o
}
export const it = Object.assign(
  (name: string, fn: () => unknown, timeout?: number) => vitestIt(name, () => owners.run(current!, fn), timeout),
  { each: vitestIt.each.bind(vitestIt) as typeof vitestIt.each },
)
const bitmapOf = async (bytes: Uint8Array) => {
  const img = await loadImage(Buffer.from(bytes))
  const cv = document.createElement('canvas')
  cv.width = img.width; cv.height = img.height
  cv.getContext('2d')!.drawImage(img as unknown as CanvasImageSource, 0, 0)
  return cv
}

/** beforeEach: a fresh owner, its clock and frame loop behind the show's seams, media elements stubbed. */
export function setUp(o: { canPlay?: string } = {}): void {
  if (current) current.live = false
  const own: Owner = current = { clock: manualClock(), live: true }
  vi.spyOn(remoteShowDeps, 'clock').mockImplementation(() => own.clock)
  vi.spyOn(remoteShowDeps, 'every').mockImplementation(() => () => {})
  vi.spyOn(remoteShowDeps, 'frames').mockImplementation((_w, f) => { own.frame = f; return () => { if (own.frame === f) own.frame = undefined } })
  vi.spyOn(remoteShowDeps, 'decode').mockImplementation(bitmapOf)
  vi.spyOn(remoteShowDeps, 'workspace').mockImplementation(() => (own.workspace = createSlideShowWorkspace('desktop')))
  remoteShowDeps.backend = 'canvas2d'
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: 960 })
  Object.defineProperty(window, 'innerHeight', { configurable: true, value: 540 })
  own.media = installMedia(o)
}
export function tearDown(): void {
  current!.live = false
  current!.media?.restore()
  remoteShowDeps.backend = undefined
  vi.restoreAllMocks()
  document.body.replaceChildren()
}
/** The calling test's media elements and show workspace. */
export const elements = (): MediaElements => own().media!
export const showWorkspace = (): SlideShowWorkspace => own().workspace!
export const clockNow = (): number => own().clock.now()

/** Runs the show `ms` at ~60 fps. */
export async function run(ms: number) {
  for (let t = 0; t < ms; t += 16) {
    const o = own()
    o.clock.advance(16)
    o.frame?.()
    await flush()
  }
  own()
}

export async function show(fd: MediaDoc, o: Partial<RemoteShowStart> = {}) {
  const host = document.createElement('div')
  host.tabIndex = -1
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const diags: Array<{ code: string; message: string }> = []
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
      onDiagnostic: (d) => { log.push(`diag:${d.code}`); diags.push({ code: d.code, message: d.message }) },
      currentSlide: () => 0, restoreFocus() {}, scrollToSlide: () => {},
    },
    signal: ac.signal, generation: 1, cleanup: () => { ac.abort() },
    ...o,
  }
  const s = await runRemoteShow(start)
  await flush()
  const root = host.querySelector<HTMLElement>('[data-officeview-show-surface]')!
  return { s, host, log, diags, root }
}

/** The driver's recorded media ends on the slide on screen: `[nodeId, t]`. */
export const mediaEnds = (s: Awaited<ReturnType<typeof show>>['s']) =>
  (s.cursor?.events ?? []).flatMap((e) => (e.kind === 'mediaEnd' ? [[e.nodeId, e.t] as [number, number]] : []))
