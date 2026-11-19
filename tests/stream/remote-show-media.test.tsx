/**
 * Piece 5 Task 11 (spec §4.4, §5; plan D4, D12, D16, D18, D19, D22): the remote show plays audio. The media player
 * is mounted on the audience surface with the remote opener (`openRemoteMediaLease` over the session), the
 * transition sound plays through the same opener, `media-show` / `media-cursor` / `media-end` follow the show, a
 * slide's media is requested when it becomes the next or the current slide (paced by the session's `media` log, 6 per
 * 60 s), and audio ends on the show clock (`begin + durationMs`) whatever the delivery does (D4).
 *
 * Driven over the recorded media-deck bundles (`remote-show-media-rig.ts`) with the piece-4 T5 throw-on-import mocks,
 * plus the package's media lease and part reader (`slideshow/media/lease`, `slideshow/media/read`).
 *
 * Task 5 ruling (binding): a slide whose drawing waits on a sound end is a still, with media on as off, so no
 * animated slide has a build after a narration that the narration's end feeds. The D4 case is therefore asserted on
 * the history itself: a runtime refusal leaves the driver's media-end events exactly those of the delivered run,
 * at `begin + durationMs`, the shared computation (`slideMediaEnds`) the classifier and driver evaluate.
 */
import { afterEach, beforeEach, describe, expect, vi } from 'vitest'

const { boom, calls } = vi.hoisted(() => ({
  boom: (m: string) => () => { throw new Error(`T5: package module reached: ${m}`) },
  calls: (m: string) => () => new Proxy({}, {
    get: (_, k) => (k === 'then' || typeof k === 'symbol' || k === '__esModule' ? undefined : () => { throw new Error(`T5: package code called: ${m}.${String(k)}`) }),
    has: () => true,
  }),
}))
vi.mock('../../src/pptx/types', boom('pptx/types'))
vi.mock('../../src/core/zip', boom('core/zip'))
vi.mock('../../src/core/package-sniff', boom('core/package-sniff'))
vi.mock('../../src/slideshow/read/store', boom('read/store'))
vi.mock('../../src/slideshow/read/deck', boom('read/deck'))
vi.mock('../../src/slideshow/prepare/index', boom('prepare/index'))
vi.mock('../../src/slideshow/render/raster', boom('render/raster'))
vi.mock('../../src/slideshow/render/plan', boom('render/plan'))
vi.mock('../../src/slideshow/render/backend', boom('render/backend'))
vi.mock('../../src/slideshow/morph/match', boom('morph/match'))
vi.mock('../../src/slideshow/player/stage', boom('player/stage'))
vi.mock('../../src/slideshow/player/session', boom('player/session'))
vi.mock('../../src/slideshow/player/text-builds', boom('player/text-builds'))
vi.mock('../../src/slideshow/player/hit-test', boom('player/hit-test'))
vi.mock('../../src/slideshow/media/lease', boom('media/lease'))
vi.mock('../../src/slideshow/media/read', boom('media/read'))
vi.mock('../../src/slideshow/render/direct', calls('render/direct'))
vi.mock('../../src/slideshow/render/scene', calls('render/scene'))
vi.mock('../../src/pptx/parse', calls('pptx/parse'))
vi.mock('../../src/core/package-error', calls('core/package-error'))

import { gzipSync } from 'node:zlib'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { encodeMessage } from '../../src/stream/frame'
import { encodeManifest, MEDIA_AUDIO_CAPS, type Manifest } from '../../src/stream/manifest'
import { StreamSession } from '../../src/stream/remote/session'
import { clockNow, elements, flush, it, mediaDoc, mediaEnds, mediaFixture, run, setUp, show, showWorkspace, tearDown, type MediaDoc } from './remote-show-media-rig'

const THROW_ON_IMPORT = [
  '../../src/pptx/types', '../../src/core/zip', '../../src/core/package-sniff',
  '../../src/slideshow/read/store', '../../src/slideshow/read/deck', '../../src/slideshow/prepare/index',
  '../../src/slideshow/render/raster', '../../src/slideshow/render/plan', '../../src/slideshow/render/backend',
  '../../src/slideshow/morph/match', '../../src/slideshow/player/stage', '../../src/slideshow/player/session', '../../src/slideshow/player/text-builds', '../../src/slideshow/player/hit-test',
  '../../src/slideshow/media/lease', '../../src/slideshow/media/read',
]
const THROW_ON_CALL = ['../../src/slideshow/render/direct', '../../src/slideshow/render/scene', '../../src/pptx/parse', '../../src/core/package-error']

const NARRATION = mediaFixture('narration')
const NARRATION_ID = NARRATION.plan(0).media![0]!.mediaId
const NARRATION_MS = NARRATION.plan(0).media![0]!.durationMs
/** The narration's timing node (its `cTn` id, `mediaDeck` shape 3). */
const NARRATION_NODE = 3
const DRIVES = 15_000

beforeEach(() => setUp())
afterEach(() => tearDown())

const audios = (root: ParentNode) => [...root.querySelectorAll('audio, video')] as HTMLMediaElement[]
const diagCodes = (log: string[]) => log.filter((l) => l.startsWith('diag:'))

describe('remote show plays audio (Task 11)', () => {
  it('narrationDeck: media-show at start, request on entering next/current, audio plays, node ends at begin+durationMs', async () => {
    const fd = mediaDoc(NARRATION)
    const { s, root } = await show(fd)
    await run(200)
    expect(fd.log).toContain('media-show:0')
    expect(fd.log.indexOf('media-show:0')).toBeLessThan(fd.log.findIndex((l) => l.startsWith('media-request:')))
    expect(fd.requests()).toEqual([`media-request:${NARRATION_ID}@0`])
    const els = audios(root)
    expect(els).toHaveLength(1)
    expect(els[0]!.getAttribute('data-officeview-show-media-spid')).toBe('3')
    expect(elements().playing()).toEqual(els)
    expect(mediaEnds(s)).toEqual([])
    await run(NARRATION_MS + 200)
    // On the show clock: the node's end is begin (0) + durationMs, whatever the element does.
    expect(mediaEnds(s)).toEqual([[NARRATION_NODE, NARRATION_MS]])
    expect(elements().playing()).toEqual([])
    expect(fd.requests()).toHaveLength(1)
    s.exit()
  }, DRIVES)

  it("bytes after begin start at the node's offset; after its end never play", async () => {
    const fd = mediaDoc(NARRATION)
    fd.rules.set(NARRATION_ID, 'hold')
    const { s } = await show(fd)
    await run(800)
    expect(elements().plays).toEqual([])
    const t = clockNow()
    fd.deliver(NARRATION_ID)
    await run(32)
    expect(elements().plays).toHaveLength(1)
    // Started at the node's current offset (seconds), within a frame or two of the delivery.
    expect(elements().plays[0]!.at).toBeGreaterThanOrEqual(0.7)
    expect(elements().plays[0]!.at).toBeLessThanOrEqual((clockNow() - 0) / 1000)
    expect(t).toBeGreaterThan(700)
    s.exit()

    // A second show: the bytes arrive after the node ended. Nothing plays.
    const fd2 = mediaDoc(NARRATION)
    fd2.rules.set(NARRATION_ID, 'hold')
    const before = elements().plays.length
    const b = await show(fd2)
    await run(NARRATION_MS + 300)
    expect(mediaEnds(b.s)).toEqual([[NARRATION_NODE, NARRATION_MS]])
    fd2.deliver(NARRATION_ID)
    await run(200)
    expect(elements().plays.length).toBe(before)
    expect(elements().playing()).toEqual([])
    b.s.exit()
  }, DRIVES)

  it('a runtime media-timeout keeps the clock end: builds after the narration fire at begin+durationMs', async () => {
    const ok = mediaDoc(NARRATION)
    const a = await show(ok)
    await run(NARRATION_MS + 300)
    const delivered = a.s.cursor!.events
    a.s.exit()

    const fd = mediaDoc(NARRATION)
    fd.rules.set(NARRATION_ID, { refused: 'media-timeout' })
    const b = await show(fd)
    await run(300)
    // The typed notice, the poster (the slide's picture) stays, and nothing plays.
    expect(b.diags.find((d) => d.code === 'media-refused')?.message).toMatch(/media-timeout/)
    expect(b.root.getAttribute('data-officeview-show-frame')).toBe('image')
    expect(elements().playing()).toEqual([])
    expect(mediaEnds(b.s)).toEqual([])
    await run(NARRATION_MS)
    // The history the classifier evaluated: the end at begin + durationMs, exactly the delivered run's.
    expect(mediaEnds(b.s)).toEqual([[NARRATION_NODE, NARRATION_MS]])
    expect(b.s.cursor!.events).toEqual(delivered)
    b.s.exit()
  }, DRIVES)

  it('eight media parts across current+next stay within 6 requests/min and never close 4008', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'Date'] })
    try {
      const fx = mediaFixture('eight-parts')
      const { session, sock } = await realSession(fx.list)
      const fd = mediaDoc(fx, { transport: session })
      const { s } = await show(fd)
      const times = () => sock.of('media-request').map((m) => m.at as number)
      const answer = () => { for (const m of sock.of('media-request')) if (!m.answered) { m.answered = true; sock.text({ reqId: m.reqId, type: 'media-refused', reason: 'media-busy' }) } }
      for (let k = 0; k < 20; k++) { await run(100); await vi.advanceTimersByTimeAsync(100); answer() }
      const show0 = sock.of('media-show')
      expect(show0.map((m) => m.unit)).toEqual([0])
      // The current slide's four first, then look-ahead for slide 1, every one naming its slide; 6 in the window.
      expect(sock.of('media-request')).toHaveLength(6)
      expect(sock.of('media-request').slice(0, 4).map((m) => m.unit)).toEqual([0, 0, 0, 0])
      expect(sock.of('media-request').slice(4).map((m) => m.unit)).toEqual([1, 1])
      for (let k = 0; k < 70; k++) { await vi.advanceTimersByTimeAsync(1000); answer() }
      const all = times()
      expect(all).toHaveLength(8)
      // Never more than 6 in any 60 s window (the server's `media` class), so no refusal is ever needed.
      for (let i = 6; i < all.length; i++) expect(all[i]! - all[i - 6]!).toBeGreaterThan(60_000)
      expect(new Set(sock.of('media-request').map((m) => m.mediaId)).size).toBe(8)
      expect(sock.closed).toBeUndefined()
      s.exit()
      expect(sock.of('media-end')).toHaveLength(1)
    } finally { vi.useRealTimers() }
  }, DRIVES)

  it('cross-slide sound keeps its lease across numSld slides and is not re-requested', async () => {
    const fx = mediaFixture('cross-slide')
    const id = fx.plan(0).media![0]!.mediaId
    const fd = mediaDoc(fx)
    const { s, root } = await show(fd)
    await run(300)
    const [el] = audios(root)
    expect(el).toBeDefined()
    expect(elements().playing()).toEqual([el])
    s.next(); await run(400)
    expect(s.state.slide).toBe(1)
    expect(audios(root)).toEqual([el])
    expect(elements().playing()).toEqual([el])
    s.next(); await run(400)
    expect(s.state.slide).toBe(2)
    expect(audios(root)).toEqual([el])
    expect(fd.requests()).toEqual([`media-request:${id}@0`])
    expect(fd.log.filter((l) => l.startsWith('media-cursor:'))).toEqual(['media-cursor:1', 'media-cursor:2'])
    expect(elements().revoked).toEqual([])
    s.exit()
    expect(elements().revoked).toEqual(elements().created)
  }, DRIVES)

  it('transition sound plays through TransitionSoundDeps.open remote', async () => {
    const fx = mediaFixture('transition-sound')
    const id = fx.plan(0).media![0]!.mediaId
    const fd = mediaDoc(fx)
    const { s, root } = await show(fd)
    await run(400)
    expect(fd.requests()).toEqual([`media-request:${id}@0`])
    const sound = root.querySelector<HTMLMediaElement>('audio[data-officeview-show-sound]')
    expect(sound).not.toBeNull()
    expect(elements().playing()).toEqual([sound])
    s.exit()
  }, DRIVES)

  it("effect sound plays through the snd target's mediaId", async () => {
    const fx = mediaFixture('effect-sound')
    const id = fx.plan(0).media![0]!.mediaId
    const fd = mediaDoc(fx)
    const { s, root } = await show(fd)
    await run(300)
    // Look-ahead on entry: the slide is current, its sound is requested before the click that plays it.
    expect(fd.requests()).toEqual([`media-request:${id}@0`])
    expect(audios(root)).toEqual([])
    s.next(); await run(300)
    const el = root.querySelector<HTMLMediaElement>('audio[data-officeview-show-effect-sound="90"]')
    expect(el).not.toBeNull()
    expect(elements().playing()).toEqual([el])
    // The look-ahead delivery is the one played: not asked for twice.
    expect(fd.requests()).toHaveLength(1)
    s.exit()
  }, DRIVES)

  it('plan-refused record shows poster + notice and ends per local fallback', async () => {
    const fx = mediaFixture('refused')
    expect(fx.plan(0).media![0]!.refused).toBe('media-unreadable')
    const fd = mediaDoc(fx)
    const { s, root, diags } = await show(fd)
    await run(300)
    expect(fd.requests()).toEqual([])
    expect(diags.find((d) => d.code === 'media-refused')?.message).toMatch(/media-unreadable/)
    expect(root.getAttribute('data-officeview-show-frame')).toBe('image')
    expect(elements().plays).toEqual([])
    // No finite dur: it ends at once (§5.3), at its begin.
    expect(mediaEnds(s)).toEqual([[NARRATION_NODE, 0]])
    s.exit()
  }, DRIVES)

  it('goTo sends media-cursor; show end sends media-end and releases every lease', async () => {
    const fx = mediaFixture('cross-slide')
    const fd = mediaDoc(fx)
    const ws = () => showWorkspace()
    const { s, log } = await show(fd)
    await run(300)
    const base = ws().used()
    expect(elements().created).toHaveLength(1)
    s.goTo(2); await run(300)
    expect(fd.log.filter((l) => l.startsWith('media-cursor:'))).toEqual(['media-cursor:2'])
    s.next(); await run(100)
    expect(s.state.slide).toBe('end')
    s.next(); await run(50)
    expect(log).toContain('onEnd')
    expect(fd.log.filter((l) => l === 'media-end')).toHaveLength(1)
    expect(elements().revoked).toEqual(elements().created)
    expect(ws().used()).toBe(0)
    expect(base).toBeGreaterThan(0)
  }, DRIVES)

  it('revocation mid-narration exits; no media element plays after ended', async () => {
    const fd = mediaDoc(NARRATION)
    const { log, host } = await show(fd)
    await run(500)
    expect(elements().playing()).toHaveLength(1)
    await fd.end(new StreamSessionError('revoked' as never))
    expect(log).toContain('onEnd')
    expect(elements().playing()).toEqual([])
    expect(audios(host)).toEqual([])
    const plays = elements().plays.length
    await run(500)
    expect(elements().plays.length).toBe(plays)
    expect(elements().revoked).toEqual(elements().created)
  }, DRIVES)

  it('videoDeck: the still plays, no media-request is sent, and the fetches equal those with media off', async () => {
    const fx = mediaFixture('video')
    const fetches = async (media: boolean) => {
      const fd: MediaDoc = mediaDoc(fx, { media, planRule: () => 'still' })
      const { s, root, log } = await show(fd)
      await run(600)
      expect(root.getAttribute('data-officeview-show-frame')).toBe('image')
      expect(fd.requests()).toEqual([])
      expect(audios(root)).toEqual([])
      expect(diagCodes(log).filter((d) => d.startsWith('diag:media'))).toEqual([])
      s.exit()
      return fd.log.filter((l) => l.startsWith('plan:') || l.startsWith('paint:'))
    }
    const on = await fetches(true)
    expect(on).toEqual(await fetches(false))
    expect(on).toEqual(['plan:0', 'paint:0'])
  }, DRIVES)

  it("mediaEndRevealDeck('unrelated') with media on: the player opens nothing for the video, raises no notice and creates no element", async () => {
    const fx = mediaFixture('unrelated')
    const fd = mediaDoc(fx)
    const { s, root, log } = await show(fd)
    await run(300)
    s.next(); await run(800)
    expect(s.state.build).toBe(1)
    expect(fd.requests()).toEqual([])
    expect(audios(root)).toEqual([])
    expect(elements().plays).toEqual([])
    expect(diagCodes(log).filter((d) => d.startsWith('diag:media'))).toEqual([])
    // The video ends as piece 4 ends it: at its first live frame.
    expect(mediaEnds(s).map(([id]) => id)).toEqual([3])
    s.exit()
  }, DRIVES)

  it('media off (D18): no player, no request, and media nodes end at once as in piece 4', async () => {
    const fd = mediaDoc(NARRATION, { media: false })
    const { s, root } = await show(fd)
    await run(300)
    expect(fd.requests()).toEqual([])
    expect(audios(root)).toEqual([])
    expect(mediaEnds(s)).toEqual([[NARRATION_NODE, expect.any(Number)]])
    expect(mediaEnds(s)[0]![1]).toBeLessThan(100)
    s.exit()
  }, DRIVES)

  it('a drop: after the resume media-show and the cursor are sent again, and a rotated request is retried', async () => {
    const fx = mediaFixture('cross-slide')
    const id = fx.plan(0).media![0]!.mediaId
    const fd = mediaDoc(fx)
    fd.rules.set(id, { refused: 'session-rotated' })
    const { s, diags } = await show(fd)
    await run(200)
    expect(fd.requests()).toEqual([`media-request:${id}@0`])
    fd.rules.set(id, 'ok')
    fd.resumed()
    await run(200)
    expect(fd.log.filter((l) => l.startsWith('media-show:'))).toEqual(['media-show:0', 'media-show:0'])
    expect(fd.requests()).toEqual([`media-request:${id}@0`, `media-request:${id}@0`])
    expect(diags.filter((d) => d.code === 'media-refused')).toEqual([])
    expect(elements().playing()).toHaveLength(1)
    s.exit()
  }, DRIVES)

  it('no package module is imported', async () => {
    for (const m of THROW_ON_IMPORT) await expect(import(/* @vite-ignore */ m), m).rejects.toThrow()
    for (const m of THROW_ON_CALL) {
      const mod = await import(/* @vite-ignore */ m) as Record<string, () => unknown>
      expect(() => mod.anything!(), m).toThrow(/package code called/)
    }
    const fd = mediaDoc(mediaFixture('effect-sound'))
    const { s, log } = await show(fd)
    await run(300)
    s.next(); await run(300)
    expect(elements().playing()).toHaveLength(1)
    expect(log.filter((l) => /T5/.test(l))).toEqual([])
    s.exit()
  }, DRIVES)
})

// ---------------------------------------------------------------------------
// A real StreamSession over a fake socket, for the pacing test (the session holds the `media` log)
// ---------------------------------------------------------------------------

const b64url = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
class FakeSocket extends EventTarget {
  binaryType: BinaryType = 'blob'
  readyState = 0
  closed: number | undefined
  sent: Array<Record<string, any>> = []
  send(d: string) { this.sent.push({ ...JSON.parse(d), at: Date.now() }) }
  close(code?: number) { if (this.readyState === 3) return; this.readyState = 3; this.closed = code ?? 1005 }
  open() { this.readyState = 1; this.dispatchEvent(new Event('open')) }
  deliver(data: ArrayBuffer | string) { this.dispatchEvent(Object.assign(new Event('message'), { data })) }
  text(m: object) { this.deliver(JSON.stringify(m)) }
  frame(kind: number, reqId: number, payload: Uint8Array) {
    for (const f of encodeMessage({ ver: 1, kind: kind as never, reqId, payload, compress: (b) => new Uint8Array(gzipSync(b)) })) this.deliver(f.slice().buffer)
  }
  of(type: string) { return this.sent.filter((m) => m.type === type) }
}
async function realSession(list: Manifest['units']['list']) {
  const manifest: Manifest = {
    v: 1, docRef: 'a'.repeat(24), kind: 'pptx', name: 'deck.pptx',
    units: { count: list.length, list },
    layoutFingerprint: 'f'.repeat(64),
    session: { sid: '0123456789ab', token: `h.${b64url({ jti: 'j1' })}.s`, expiresSec: 300, resume: 'session' },
    policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
    can: { search: false, copy: false, print: false, textA11y: false, present: true, notes: false, links: false, media: true },
    caps: { copy: null, a11y: 'off', print: null, media: { audio: { ...MEDIA_AUDIO_CAPS } } },
    recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [1, 2], requiredFaces: [], projection: 'final-raster' },
    limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: 1 << 20, tilePx: 512 },
    notices: [],
  }
  const socks: FakeSocket[] = []
  const p = StreamSession.open({ kind: 'stream', url: 'wss://x/stream' }, { client: () => { const s = new FakeSocket(); socks.push(s); return s as never }, faces: async () => [], now: () => Date.now() })
  const sock = socks[0]!
  sock.open(); await vi.advanceTimersByTimeAsync(10)
  sock.frame(1, sock.of('hello')[0]!.reqId, encodeManifest(manifest))
  const session = await p
  // media-show is answered as the server does.
  const answerShows = () => { for (const m of sock.of('media-show')) if (!m.answered) { m.answered = true; sock.text({ reqId: m.reqId, type: 'media-show-ok', queued: 8 }) } }
  const origSend = sock.send.bind(sock)
  sock.send = (d: string) => { origSend(d); queueMicrotask(answerShows) }
  await flush()
  return { session, sock }
}
