/**
 * Piece 5 Task 8: the per-session show-start encode queue over the shared cache and the admission class (plan D1,
 * D12, D22, D23; Review Focus 3, 4, 8). A fake transcode with controllable gates, sizes and a call counter, and a
 * real cache over the media temp dir; the real transcode where a test says so.
 */
import { afterAll, describe, expect, it } from 'bun:test'
import { execFileSync } from 'node:child_process'
import { randomBytes } from 'node:crypto'
import { existsSync, mkdtempSync, readFileSync, rmSync, unlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { mediaDeck } from '../../../src/testdata/show-decks'
import type { MediaRefusal } from '../../../src/stream/media-types'
import { createMediaAdmission, MEDIA_QUEUE, type MediaAdmission } from '../../src/media/admission'
import { bindMedia, mediaIdentity, sessionMediaId } from '../../src/media/bind'
import { createEncodeCache, MEDIA_IDLE_TTL_MS, MEDIA_SERVER_CACHE, MEDIA_SESSION_LOOKAHEAD, type EncodeCache } from '../../src/media/cache'
import type { FfmpegInfo } from '../../src/media/ffmpeg'
import { mediaJobsFor, type MediaItem } from '../../src/media/jobs'
import { encodeKeyId, encodeKeyOf } from '../../src/media/params'
import type { ProbedMedia } from '../../src/media/probe'
import { createSessionMediaQueue, type SessionMediaQueue, type SessionMediaQueueOptions } from '../../src/media/queue'
import { createMediaServices } from '../../src/media/services'
import { createMediaTemp } from '../../src/media/temp'
import { transcode, type EncodeResult } from '../../src/media/transcode'
import { DocumentHost } from '../../src/render'
import type { BuiltShowPlan, ShowPlanJob, ShowRefusal } from '../../src/show/build'
import { ffmpegForTests } from './ffmpeg-gate'

const info = await ffmpegForTests()
const root = mkdtempSync(join(tmpdir(), 'ov-queue-'))
const temp = await createMediaTemp(root)
afterAll(async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) })

const MiB = 1 << 20
const FAKE_INFO: FfmpegInfo = info ?? { ffmpeg: '/nonexistent/ffmpeg', ffprobe: '/nonexistent/ffprobe', version: '8.1.1', platform: 'darwin' }
const until = async (cond: () => boolean, what = 'condition', ms = 20_000) => {
  const t0 = Date.now()
  while (!cond()) { if (Date.now() - t0 > ms) throw new Error(`timed out waiting for ${what}`); await Bun.sleep(1) }
}

// --- fakes ------------------------------------------------------------------------------------------------------

interface Call { name: string; out: string; signal: AbortSignal; finish(): void; fail(r: 'media-too-large' | 'media-timeout' | 'media-unreadable'): void; aborted: boolean }
/** A transcode that writes a 1-byte output reported as `size(name)`; `manual` holds each call until `finish`/`fail`. */
function fakeTranscode(o: { manual?: boolean; size?: (name: string) => number; admission?: () => MediaAdmission } = {}) {
  const calls: Call[] = []
  const s = { count: 0, active: 0, maxActive: 0, maxRunning: 0 }
  const fn: typeof transcode = async (_info, job, signal) => {
    const name = job.entry.partPath.split('/').pop()!
    s.count++; s.active++; s.maxActive = Math.max(s.maxActive, s.active)
    if (o.admission) s.maxRunning = Math.max(s.maxRunning, o.admission().running)
    try {
      writeFileSync(job.out, 'x')
      let settle!: (r: 'ok' | MediaRefusal | 'abort') => void
      const gate = new Promise<'ok' | MediaRefusal | 'abort'>((ok) => { settle = ok })
      const call: Call = { name, out: job.out, signal, finish: () => settle('ok'), fail: (r) => settle(r), aborted: false }
      calls.push(call)
      const onAbort = () => settle('abort')
      if (signal.aborted) settle('abort'); else signal.addEventListener('abort', onAbort, { once: true })
      if (!o.manual) queueMicrotask(() => settle('ok'))
      const r = await gate
      signal.removeEventListener('abort', onAbort)
      if (r === 'abort' || signal.aborted) { call.aborted = true; try { unlinkSync(job.out) } catch { /* gone */ }; throw signal.reason }
      if (r !== 'ok') { try { unlinkSync(job.out) } catch { /* gone */ }; return { ok: false, reason: r as 'media-too-large' | 'media-timeout' | 'media-unreadable' } }
      return { ok: true, bytes: o.size?.(name) ?? 1, durationMs: job.entry.durationMs, contentType: 'audio/mp4', paramsSha: 'p', ffmpegVersion: FAKE_INFO.version } satisfies EncodeResult
    } finally { s.active-- }
  }
  return { fn, calls, s }
}

const entryOf = (name: string, durationMs = 2000): ProbedMedia => ({ k: 0, partPath: `/ppt/media/${name}`, kind: 'audio', hasAudio: true, durationMs, sourceHash: `hash-${name}`, demuxer: 'mov' })
/** A synthetic item: session `sid`, document `docId`, slide `unit`, part `name` (its bytes are `name`'s). `a#1` and
 * `a#2` are two uses (two mediaIds) of part `a`: one key. */
function itemOf(sid: string, docId: string, unit: number, name: string): MediaItem {
  const e = entryOf(name.split('#')[0]!)
  return { mediaId: `${sid}~${unit}~${name}`, unit, ref: { kind: 'audio', k: 0 }, key: encodeKeyOf(docId, e, FAKE_INFO), partPath: e.partPath, job: (out, src) => ({ entry: e, src, out }) }
}
interface World { cache: EncodeCache; admission: MediaAdmission; advance(ms: number): void }
function world(o: { slots?: number; queue?: number; sessionBytes?: number; serverBytes?: number } = {}): World {
  let t = 5_000_000
  return {
    cache: createEncodeCache({ temp, sessionBytes: o.sessionBytes ?? MEDIA_SESSION_LOOKAHEAD, serverBytes: o.serverBytes ?? MEDIA_SERVER_CACHE, now: () => t }),
    admission: createMediaAdmission({ slots: o.slots ?? 4, queue: o.queue ?? MEDIA_QUEUE }),
    advance: (ms) => { t += ms },
  }
}
const fakePart = async () => temp.create(new Uint8Array([1, 2, 3]))
/** A queue over `deck[unit]` = part names on that slide (synthetic items). */
function synthQueue(w: World, sid: string, docId: string, deck: string[][], tc: ReturnType<typeof fakeTranscode>, part: SessionMediaQueueOptions['part'] = fakePart, extra: Partial<SessionMediaQueueOptions> = {}): SessionMediaQueue & { id(unit: number, name: string): string } {
  const q = createSessionMediaQueue({
    sid, units: deck.length, info: FAKE_INFO, temp, admission: w.admission, cache: w.cache, transcode: tc.fn, part,
    items: async (u) => deck[u]!.map((n) => itemOf(sid, docId, u, n)),
    ...extra,
  })
  return Object.assign(q, { id: (unit: number, name: string) => `${sid}~${unit}~${name}` })
}
const ok = (r: Awaited<ReturnType<SessionMediaQueue['request']>>) => { if (!r.ok) throw new Error(`request refused: ${r.reason}`); return r }

// --- real decks (the probe needs ffmpeg) -----------------------------------------------------------------------

const box = { id: 2, name: 'Box', prst: 'rect', off: ['3619500', '2476500'] as [string, string], ext: ['1905000', '1905000'] as [string, string], fill: '336699' }
const tgt = (spid: number) => `<p:spTgt spid="${spid}"/>`
const audioNode = (id: number, target: string) => `<p:audio><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl>${target}</p:tgtEl></p:cMediaNode></p:audio>`
const videoNode = (id: number, spid: number) => `<p:video><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl>${tgt(spid)}</p:tgtEl></p:cMediaNode></p:video>`
const SND = '<p:sndTgt r:embed="rIdSnd1" name="a.wav"/>'
/** Slide 0: a shape audio (a.m4a) and an effect sound (a.wav). Slide 1: the same a.m4a bytes on its own shape, a
 * hidden shape audio (a.mp3) and a video (av.mp4). */
const twoSlideDeck = () => mediaDeck({ slides: [
  { boxes: [box], media: [{ spid: 3, file: 'a.m4a' }], sounds: [{ rId: 'rIdSnd1', file: 'a.wav' }], timing: audioNode(3, tgt(3)) + audioNode(4, SND) },
  { boxes: [box], media: [{ spid: 3, file: 'a.m4a' }, { spid: 4, file: 'a.mp3', hidden: true }, { spid: 5, file: 'av.mp4', kind: 'video' }], timing: audioNode(3, tgt(3)) + audioNode(6, tgt(4)) + videoNode(7, 5) },
] })
const oneAudioDeck = (file: 'a.m4a' | 'a.mp3' | 'a.wav' | 'a.wma' = 'a.m4a', extra = 0) => mediaDeck({ slides: [{ boxes: [box, ...Array.from({ length: extra }, (_, i) => ({ ...box, id: 10 + i, name: `X${i}` }))], media: [{ spid: 3, file }], timing: audioNode(3, tgt(3)) }] })

function realHost() { const h = new DocumentHost(); h.configureMedia({ info: info!, dir: temp.dir }); return h }
async function openDoc(bytes: Uint8Array, policy: 'visible-only' | 'full' = 'visible-only') {
  const h = realHost()
  const { docKey, units } = await h.open(`doc-${Math.random()}`, bytes, policy)
  return { h, docKey, units: units.length }
}
type Doc = Awaited<ReturnType<typeof openDoc>>
const session = (doc: Doc, docId: string) => ({ mediaKey: new Uint8Array(randomBytes(32)), docId, doc })
function realQueue(w: World, sid: string, s: ReturnType<typeof session>, tc: typeof transcode) {
  const { h, docKey, units } = s.doc
  return createSessionMediaQueue({
    sid, units, info: info!, temp, admission: w.admission, cache: w.cache, transcode: tc,
    items: async (u) => mediaJobsFor(await h.mediaRefs(docKey, u), u, s, info!),
    part: async (pp) => { const r = await h.mediaPart(docKey, pp, temp.dir); return temp.adopt(r.path, r.bytes) },
  })
}
const job = (unitIndex: number): ShowPlanJob => ({ unitIndex, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false, media: true })
const built = (r: BuiltShowPlan | ShowRefusal): BuiltShowPlan => { if ('refused' in r) throw new Error(`refused: ${r.refused}`); return r }

describe.skipIf(!info)('mediaJobsFor (D2, D21, D22, D23)', () => {
  it("mediaJobsFor ids equal bindMedia's ids for the same slide and session", async () => {
    const doc = await openDoc(await twoSlideDeck())
    const s = session(doc, 'docA')
    for (const unit of [0, 1]) {
      const b = built(await doc.h.showPlan(doc.docKey, job(unit)))
      const plan = bindMedia(b.plan, (ref) => sessionMediaId(s.mediaKey, mediaIdentity(ref)), b.mediaRefs)
      const shipped = (plan.media ?? []).filter((r) => !r.refused).map((r) => r.mediaId).sort()
      const items = mediaJobsFor(await doc.h.mediaRefs(doc.docKey, unit), unit, s, info!)
      expect(shipped.length).toBeGreaterThan(0)
      expect(items.map((i) => i.mediaId).sort()).toEqual(shipped)
    }
  })

  it('mediaJobsFor keys are equal for two sessions of one document while their mediaIds differ; the same bytes in a second document give a different key', async () => {
    const doc = await openDoc(await oneAudioDeck())
    const other = await openDoc(await oneAudioDeck('a.m4a', 1)) // another document, byte-identical audio
    const refs = await doc.h.mediaRefs(doc.docKey, 0)
    const [a] = mediaJobsFor(refs, 0, session(doc, 'docA'), info!)
    const [b] = mediaJobsFor(refs, 0, session(doc, 'docA'), info!)
    const [c] = mediaJobsFor(await other.h.mediaRefs(other.docKey, 0), 0, session(other, 'docB'), info!)
    expect(encodeKeyId(a!.key)).toBe(encodeKeyId(b!.key))
    expect(a!.mediaId).not.toBe(b!.mediaId)
    expect(c!.key.sourceHash).toBe(a!.key.sourceHash)
    expect(encodeKeyId(c!.key)).not.toBe(encodeKeyId(a!.key))
  })

  it('mediaJobsFor skips hidden slides and hidden shapes, and yields no item for a video node', async () => {
    const full = await openDoc(await twoSlideDeck(), 'full')
    const refs = await full.h.mediaRefs(full.docKey, 1)
    expect(refs.refs.some((r) => r.hiddenShape)).toBe(true) // the hidden a.mp3 shape is listed under full
    const items = mediaJobsFor(refs, 1, session(full, 'docA'), info!)
    expect(items.map((i) => i.ref)).toEqual([{ kind: 'audio', k: items[0]!.ref.k, unit: 1, spid: '3' }])
    expect(items.every((i) => !i.partPath.endsWith('.mp4'))).toBe(true) // no video
    expect(mediaJobsFor({ ...refs, hidden: true }, 1, session(full, 'docA'), info!)).toEqual([])
  })
})

describe('the session queue (D1, D12)', () => {
  it('start encodes in slide order from the start unit, then wraps', async () => {
    const w = world(), tc = fakeTranscode()
    const q = synthQueue(w, 's', 'd-order', [['u0'], ['u1'], ['u2'], ['u3']], tc)
    expect(await q.start(2)).toBe(4)
    await until(() => tc.s.count === 4 && q.running === undefined, 'four encodes')
    expect(tc.calls.map((c) => c.name)).toEqual(['u2', 'u3', 'u0', 'u1'])
    await q.end(); await w.cache.close()
  })

  it('Task 11b fix round: encode-ahead skips an item `ahead` declines; a request still encodes it', async () => {
    const w = world(), tc = fakeTranscode()
    const q = synthQueue(w, 's', 'd-ahead', [['u0'], ['u1'], ['u2']], tc, fakePart, { ahead: (it) => it.unit !== 1 })
    expect(await q.start(0)).toBe(2)
    await until(() => tc.s.count === 2 && q.running === undefined, 'two encodes')
    expect(tc.calls.map((c) => c.name)).toEqual(['u0', 'u2'])
    ok(await q.request(q.id(1, 'u1')))
    expect(tc.calls.map((c) => c.name)).toEqual(['u0', 'u2', 'u1'])
    await q.end(); await w.cache.close()
  })

  it('one key awaited at a time per session', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const q = synthQueue(w, 's', 'd-one', [['a', 'b'], ['c'], ['d']], tc)
    await q.start(0)
    for (let i = 0; i < 4; i++) {
      await until(() => tc.calls.length === i + 1, `call ${i}`)
      await Bun.sleep(5) // give a second encode every chance to start
      expect(tc.s.active).toBe(1)
      tc.calls[i]!.finish()
    }
    await until(() => q.running === undefined && tc.s.active === 0, 'drain')
    expect(tc.s.maxActive).toBe(1)
    await q.end(); await w.cache.close()
  })

  it('request moves an item to the front; a cached item resolves at once with origin hit', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const deck = [['u0'], ['u1'], ['u2'], ['u3']]
    // Another session of the document encoded u3 earlier.
    const other = synthQueue(w, 'o', 'd-front', deck, tc)
    const early = other.request(other.id(3, 'u3'))
    await until(() => tc.calls.length === 1, 'the earlier encode')
    tc.calls[0]!.finish()
    expect(ok(await early).origin).toBe('miss')
    const q = synthQueue(w, 's', 'd-front', deck, tc)
    expect(await q.start(0)).toBe(3)
    await until(() => tc.calls.length === 2, 'u0 running')
    expect(tc.calls[1]!.name).toBe('u0')
    const front = q.request(q.id(2, 'u2'))
    tc.calls[1]!.finish()
    await until(() => tc.calls.length === 3, 'the next encode')
    expect(tc.calls[2]!.name).toBe('u2') // before u1
    tc.calls[2]!.finish()
    expect(ok(await front).origin).toBe('miss')
    const n = tc.s.count
    const hit = ok(await q.request(q.id(3, 'u3')))
    expect(hit.origin).toBe('hit')
    expect(tc.s.count).toBe(n)
    await until(() => tc.calls.length === 4, 'u1'); tc.calls[3]!.finish()
    await q.end(); await other.end(); await w.cache.close()
  })

  it('jump re-queues from N', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const q = synthQueue(w, 's', 'd-jump', [['u0'], ['u1'], ['u2'], ['u3'], ['u4'], ['u5']], tc)
    await q.start(0)
    await until(() => tc.calls.length === 1, 'u0')
    q.cursor(4)
    tc.calls[0]!.finish()
    for (let i = 1; i < 6; i++) { await until(() => tc.calls.length === i + 1, `call ${i}`); tc.calls[i]!.finish() }
    expect(tc.calls.map((c) => c.name)).toEqual(['u0', 'u4', 'u5', 'u1', 'u2', 'u3'])
    await q.end(); await w.cache.close()
  })

  it('a session starting after the encode finished starts none: origin hit, zero transcode calls', async () => {
    const w = world(), tc = fakeTranscode()
    const deck = [['a'], ['b']]
    const s1 = synthQueue(w, 's1', 'd-late', deck, tc)
    await s1.start(0)
    await until(() => tc.s.count === 2 && s1.running === undefined, 's1 encodes')
    const s2 = synthQueue(w, 's2', 'd-late', deck, tc)
    expect(await s2.start(0)).toBe(0)
    const r = ok(await s2.request(s2.id(1, 'b')))
    expect(r.origin).toBe('hit')
    expect(tc.s.count).toBe(2)
    await s1.end(); await s2.end(); await w.cache.close()
  })

  it('two documents with byte-identical audio produce two encodes: two transcode calls, two entries, no joins', async () => {
    const w = world(), tc = fakeTranscode()
    const a = synthQueue(w, 'sa', 'doc-A', [['same']], tc), b = synthQueue(w, 'sb', 'doc-B', [['same']], tc)
    const [ra, rb] = await Promise.all([a.request(a.id(0, 'same')), b.request(b.id(0, 'same'))])
    expect([ok(ra).origin, ok(rb).origin]).toEqual(['miss', 'miss'])
    expect(tc.s.count).toBe(2)
    expect(w.cache.stats().entries).toBe(2)
    await a.end(); await b.end(); await w.cache.close()
  })

  it('revoking a document mid-delivery: its entries are deleted, the delivery fails media-unavailable, its in-flight encode is aborted, and another document\'s entries are untouched', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const b = synthQueue(w, 'sb', 'doc-B', [['b']], tc)
    const rb = b.request(b.id(0, 'b'))
    await until(() => tc.calls.length === 1, 'b'); tc.calls[0]!.finish()
    const kept = ok(await rb)
    const a = synthQueue(w, 'sa', 'doc-A', [['a1'], ['a2']], tc)
    const done1 = a.request(a.id(0, 'a1'))
    await until(() => tc.calls.length === 2, 'a1'); tc.calls[1]!.finish()
    const a1 = ok(await done1)
    const inflight = a.request(a.id(1, 'a2'))
    await until(() => tc.calls.length === 3, 'a2')
    await w.cache.dropDocument('doc-A')
    expect(await inflight).toEqual({ ok: false, reason: 'media-unavailable' })
    expect(tc.calls[2]!.aborted).toBe(true)
    expect(existsSync(a1.path)).toBe(false)
    expect(w.cache.peek(itemOf('x', 'doc-A', 0, 'a1').key)).toBeUndefined()
    expect(w.cache.peek(itemOf('x', 'doc-B', 0, 'b').key)).toBeDefined()
    expect(existsSync(kept.path)).toBe(true)
    // The queue sees media-unavailable for the rest of the show.
    expect(await a.request(a.id(0, 'a1'))).toEqual({ ok: false, reason: 'media-unavailable' })
    expect(tc.s.count).toBe(3)
    await a.end(); await b.end(); await w.cache.close()
  })

  it('start enqueues only keys not already cached: the returned count is 0 for a deck whose keys are all cached or encoding', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const deck = [['a'], ['b']]
    const s1 = synthQueue(w, 's1', 'd-count', deck, tc)
    expect(await s1.start(0)).toBe(2)
    await until(() => tc.calls.length === 1, 'a'); tc.calls[0]!.finish()
    await until(() => tc.calls.length === 2, 'b running') // a cached, b encoding
    const s2 = synthQueue(w, 's2', 'd-count', deck, tc)
    expect(await s2.start(0)).toBe(0)
    // The cached key took its reference at start; the encoding one is joined at its turn (already, as nothing precedes it).
    expect(w.cache.referencesOf('s2')).toBe(2)
    tc.calls[1]!.finish()
    expect(ok(await s2.request(s2.id(1, 'b'))).origin).toBe('joined')
    expect(tc.s.count).toBe(2)
    await s1.end(); await s2.end(); await w.cache.close()
  })

  it('the encode survives the starter\'s end() while another session waits; it is killed when the last waiter ends', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const live0 = temp.live()
    const s1 = synthQueue(w, 's1', 'd-survive', [['a']], tc), s2 = synthQueue(w, 's2', 'd-survive', [['a']], tc)
    await s1.start(0)
    await until(() => tc.calls.length === 1, 'a')
    void s2.request(s2.id(0, 'a')) // joins
    await until(() => s2.running !== undefined, 's2 awaits the key')
    await s1.end()
    expect(tc.calls[0]!.aborted).toBe(false)
    expect(w.cache.inflight(itemOf('x', 'd-survive', 0, 'a').key)).toBe(true)
    await s2.end()
    expect(tc.calls[0]!.aborted).toBe(true)
    expect(w.admission.running).toBe(0)
    expect(temp.live()).toBe(live0) // part and output disposed
    await w.cache.close()
  })

  it('a media-timeout reaches every joiner and is not cached; the next request encodes again', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const s1 = synthQueue(w, 's1', 'd-to', [['a']], tc), s2 = synthQueue(w, 's2', 'd-to', [['a']], tc)
    const r1 = s1.request(s1.id(0, 'a'))
    await until(() => tc.calls.length === 1, 'a')
    const r2 = s2.request(s2.id(0, 'a'))
    await until(() => s2.running !== undefined, 'joined')
    tc.calls[0]!.fail('media-timeout')
    expect(await r1).toEqual({ ok: false, reason: 'media-timeout' })
    expect(await r2).toEqual({ ok: false, reason: 'media-timeout' })
    expect(w.cache.peek(itemOf('x', 'd-to', 0, 'a').key)).toBeUndefined()
    const again = s2.request(s2.id(0, 'a'))
    await until(() => tc.calls.length === 2, 'the re-encode'); tc.calls[1]!.finish()
    expect((await again).ok).toBe(true)
    await s1.end(); await s2.end(); await w.cache.close()
  })

  it('a media-too-large result is cached as a refusal for the key: a request from a second session answers it without re-encoding', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const s1 = synthQueue(w, 's1', 'd-big', [['a']], tc), s2 = synthQueue(w, 's2', 'd-big', [['a']], tc)
    const r1 = s1.request(s1.id(0, 'a'))
    await until(() => tc.calls.length === 1, 'a'); tc.calls[0]!.fail('media-too-large')
    expect(await r1).toEqual({ ok: false, reason: 'media-too-large' })
    expect(await s2.request(s2.id(0, 'a'))).toEqual({ ok: false, reason: 'media-too-large' })
    expect(tc.s.count).toBe(1)
    await s1.end(); await s2.end(); await w.cache.close()
  })

  it('twelve 25 MiB encodes on one slide all deliver under the 256 MiB look-ahead without deadlock', async () => {
    const w = world(), tc = fakeTranscode({ size: () => 25 * MiB })
    const names = Array.from({ length: 12 }, (_, i) => `n${i}`)
    const q = synthQueue(w, 's', 'd-twelve', [names], tc)
    await q.start(0)
    // Encode-ahead stops at 9: 9 × 25 MiB + the next one's 32 MiB reservation is over 256 MiB.
    await until(() => q.running === undefined && tc.s.count === 9, 'the look-ahead fills')
    await Bun.sleep(5)
    expect(tc.s.count).toBe(9)
    expect(w.cache.stats().bytes).toBeLessThanOrEqual(MEDIA_SESSION_LOOKAHEAD)
    for (const n of names) {
      const r = ok(await q.request(q.id(0, n)))
      expect(r.result.bytes).toBe(25 * MiB)
      w.cache.delivered('s', itemOf('s', 'd-twelve', 0, n).key)
    }
    expect(tc.s.count).toBe(12)
    expect(w.cache.referencesOf('s')).toBe(0)
    await q.end()
    expect(w.cache.referencesOf('s')).toBe(0)
    await w.cache.close()
  })

  it('backward jump with the look-ahead full of undelivered ahead entries drops the farthest own references and delivers', async () => {
    const w = world(), tc = fakeTranscode({ size: () => 25 * MiB })
    const deck = Array.from({ length: 13 }, (_, u) => [`n${u}`])
    // Another session of the document has encoded slides 1–12 (delivered: unreferenced, cached).
    const other = synthQueue(w, 'o', 'd-back', deck, tc)
    for (let u = 1; u < 13; u++) { ok(await other.request(other.id(u, `n${u}`))); w.cache.delivered('o', itemOf('o', 'd-back', u, `n${u}`).key) }
    expect(tc.s.count).toBe(12)
    const q = synthQueue(w, 's', 'd-back', deck, tc)
    expect(await q.start(1)).toBe(1) // slide 0 only
    // Slides 1–10 are referenced ahead of the cursor (10 × 25 MiB); slide 11 waits (250 + 25 > 256).
    expect(w.cache.referencesOf('s')).toBe(10)
    await Bun.sleep(5)
    expect(tc.s.count).toBe(12)
    q.cursor(0) // jump back
    const r = ok(await q.request(q.id(0, 'n0')))
    expect(r.origin).toBe('miss')
    expect(tc.s.count).toBe(13)
    // The farthest-ahead own reference (slide 10, from cursor 0) was dropped; its entry stays cached.
    expect(w.cache.referencesOf('s')).toBe(10) // slides 1–9 and slide 0
    expect(w.cache.peek(itemOf('s', 'd-back', 10, 'n10').key)).toBeDefined()
    // A delivery makes room, and the waiting encode-ahead resumes (slide 11, a hit): nothing waits forever.
    w.cache.delivered('s', itemOf('s', 'd-back', 9, 'n9').key)
    await until(() => w.cache.referencesOf('s') === 10, 'slide 11 referenced')
    expect(tc.s.count).toBe(13)
    await q.end(); await other.end(); await w.cache.close()
  })

  it('server queue full: encode-ahead waits silently; a request gets media-busy; a joiner takes no place', async () => {
    const w = world({ slots: 1, queue: 1 })
    const tc = fakeTranscode({ manual: true })
    const a = synthQueue(w, 'sa', 'd-busy', [['a']], tc)
    await a.start(0)
    await until(() => tc.calls.length === 1, 'a runs')
    expect(w.admission.running).toBe(1)
    const b = synthQueue(w, 'sb', 'd-busy', [['b']], tc)
    await b.start(0)
    await until(() => w.admission.queued === 1, 'b takes the place')
    const c = synthQueue(w, 'sc', 'd-busy', [['c']], tc)
    await c.start(0)
    await Bun.sleep(5) // c's encode-ahead parks on admission
    expect(w.admission.queued).toBe(1) // parked silently, no place
    const d = synthQueue(w, 'sd', 'd-busy', [['d']], tc)
    expect(await d.request(d.id(0, 'd'))).toEqual({ ok: false, reason: 'media-busy' })
    const e = synthQueue(w, 'se', 'd-busy', [['a']], tc)
    const joined = e.request(e.id(0, 'a'))
    await until(() => e.running !== undefined, 'e joins a')
    expect(w.admission.queued).toBe(1)
    expect(w.admission.running).toBe(1)
    tc.calls[0]!.finish()
    expect(ok(await joined).origin).toBe('joined')
    for (let i = 1; i < 3; i++) { await until(() => tc.calls.length === i + 1, `call ${i}`); tc.calls[i]!.finish() }
    expect(tc.calls.map((x) => x.name)).toEqual(['a', 'b', 'c'])
    await until(() => w.admission.running === 0 && w.admission.queued === 0, 'drain')
    for (const q of [a, b, c, d, e]) await q.end()
    await w.cache.close()
  })

  it('end(): queued items cancelled, references dropped, entries remain cached until the TTL', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const live0 = temp.live()
    const q = synthQueue(w, 's', 'd-end', [['a'], ['b'], ['c']], tc)
    await q.start(0)
    await until(() => tc.calls.length === 1, 'a'); tc.calls[0]!.finish()
    await until(() => tc.calls.length === 2, 'b'); tc.calls[1]!.finish()
    await until(() => tc.calls.length === 3, 'c')
    await q.end()
    expect(tc.calls[2]!.aborted).toBe(true)
    await Bun.sleep(5)
    expect(tc.s.count).toBe(3) // nothing more runs
    expect(w.cache.referencesOf('s')).toBe(0)
    expect(w.cache.stats()).toMatchObject({ entries: 2, inflight: 0 })
    expect(temp.live()).toBe(live0 + 2) // the two cached outputs only
    w.advance(MEDIA_IDLE_TTL_MS)
    await w.cache.sweep()
    expect(w.cache.stats().entries).toBe(0)
    expect(temp.live()).toBe(live0)
    expect(await q.request(q.id(0, 'a'))).toEqual({ ok: false, reason: 'media-unavailable' })
    await w.cache.close()
  })

  it("end() is bounded: one session's end never hangs on another session's stuck, abandoned encode", async () => {
    const w = world()
    let unstick!: () => void
    const stuck = new Promise<void>((ok) => { unstick = ok })
    // A transcode that ignores its abort until released (a wedged child).
    let started = 0
    const wedged: typeof transcode = async (_i, job) => { started++; writeFileSync(job.out, 'x'); await stuck; unlinkSync(job.out); throw new Error('aborted late') }
    const mk = (sid: string, name: string) => createSessionMediaQueue({ sid, units: 1, info: FAKE_INFO, temp, admission: w.admission, cache: w.cache, transcode: wedged, part: fakePart, items: async (u) => [itemOf(sid, 'd-stuck', u, name)], quiesceMs: 50 })
    const s1 = mk('s1', 'a'), s2 = mk('s2', 'b')
    await s1.start(0)
    await until(() => started === 1, 's1 encodes')
    const t0 = Date.now()
    await s1.end() // its own flight is abandoned and stuck: bounded
    await s2.start(0)
    await until(() => started === 2, 's2 encodes')
    await s2.end() // waits on the server's abandoned flights, s1's stuck one included: bounded
    expect(Date.now() - t0).toBeLessThan(5000)
    unstick()
    await w.cache.quiesce()
    await w.cache.close()
  })

  it('a request never waits behind a parked encode-ahead: with the admission queue full it answers media-busy, and no one joins the parked flight', async () => {
    const w = world({ slots: 1, queue: 1 })
    const tc = fakeTranscode({ manual: true })
    const a = synthQueue(w, 'sa', 'd-park', [['a']], tc), b = synthQueue(w, 'sb', 'd-park', [['b']], tc)
    await a.start(0)
    await until(() => tc.calls.length === 1, 'a runs')
    await b.start(0)
    await until(() => w.admission.queued === 1, 'b takes the place')
    const c = synthQueue(w, 'sc', 'd-park', [['c']], tc)
    await c.start(0) // its encode-ahead parks
    await Bun.sleep(5)
    expect(await c.request(c.id(0, 'c'))).toEqual({ ok: false, reason: 'media-busy' })
    const d = synthQueue(w, 'sd', 'd-park', [['c']], tc) // another session of the document, the same key
    expect(await d.request(d.id(0, 'c'))).toEqual({ ok: false, reason: 'media-busy' })
    expect(w.cache.inflight(itemOf('x', 'd-park', 0, 'c').key)).toBe(false) // a parked encode-ahead holds no flight
    tc.calls[0]!.finish()
    for (let i = 1; i < 3; i++) { await until(() => tc.calls.length === i + 1, `call ${i}`); tc.calls[i]!.finish() }
    expect(tc.calls.map((x) => x.name)).toEqual(['a', 'b', 'c']) // c's encode-ahead still ran
    await until(() => w.admission.running === 0, 'drain')
    for (const q of [a, b, c, d]) await q.end()
    await w.cache.close()
  })

  it('revocation stops a queue parked on room and an idle one: no encode for the revoked document, and requests answer media-unavailable', async () => {
    const w = world(), tc = fakeTranscode({ size: () => 25 * MiB })
    const names = Array.from({ length: 12 }, (_, i) => `n${i}`)
    const parked = synthQueue(w, 's', 'doc-R', [names], tc)
    await parked.start(0)
    await until(() => parked.running === undefined && tc.s.count === 9, 'the look-ahead fills')
    const idle = synthQueue(w, 'i', 'doc-R2', [['x']], tc)
    ok(await idle.request(idle.id(0, 'x')))
    const n = tc.s.count
    await w.cache.dropDocument('doc-R')
    await w.cache.dropDocument('doc-R2')
    await Bun.sleep(10)
    expect(tc.s.count).toBe(n)
    expect(w.cache.stats().entries).toBe(0)
    expect(await parked.request(parked.id(0, 'n10'))).toEqual({ ok: false, reason: 'media-unavailable' })
    expect(await idle.request(idle.id(0, 'x'))).toEqual({ ok: false, reason: 'media-unavailable' })
    expect(tc.s.count).toBe(n)
    await parked.end(); await idle.end(); await w.cache.close()
  })

  it('an answered request is pinned until delivered: pipelined requests never evict it, and two mediaIds of one key hold a pin each', async () => {
    const w = world({ serverBytes: 256 * MiB }), tc = fakeTranscode({ size: () => 25 * MiB })
    const names = Array.from({ length: 12 }, (_, i) => `n${i}`)
    const q = synthQueue(w, 's', 'd-pin', [names], tc)
    await q.start(0)
    await until(() => q.running === undefined && tc.s.count === 9, 'the look-ahead fills')
    const r9 = ok(await q.request(q.id(0, 'n9')))
    const r10 = ok(await q.request(q.id(0, 'n10'))) // before n9 is delivered
    expect(existsSync(r9.path)).toBe(true)
    expect(w.cache.peek(itemOf('s', 'd-pin', 0, 'n9').key)).toBeDefined()
    expect(existsSync(r10.path)).toBe(true)
    await q.end(); await w.cache.close()
    const w2 = world(), tc2 = fakeTranscode()
    const two = synthQueue(w2, 't', 'd-pin2', [['a#1', 'a#2']], tc2)
    const [x, y] = [ok(await two.request(two.id(0, 'a#1'))), ok(await two.request(two.id(0, 'a#2')))]
    expect(x.path).toBe(y.path)
    expect(tc2.s.count).toBe(1)
    const k = itemOf('t', 'd-pin2', 0, 'a').key
    w2.cache.delivered('t', k)
    expect(w2.cache.referencesOf('t')).toBe(1) // the second delivery still holds it
    w2.cache.delivered('t', k)
    expect(w2.cache.referencesOf('t')).toBe(0)
    await two.end(); await w2.cache.close()
  })

  it('media-busy and media-timeout leave the key unreadied and record no origin', async () => {
    const w = world({ serverBytes: 64 * MiB }), tc = fakeTranscode({ manual: true, size: (n) => (n === 'x' || n === 'y' ? 32 * MiB : 1) })
    // The server cache wholly referenced by another session (2 × 32 MiB, pinned by undelivered requests).
    const p = synthQueue(w, 'p', 'd-poison', [['x'], ['y']], tc)
    const px = p.request(p.id(0, 'x')); await until(() => tc.calls.length === 1, 'x'); tc.calls[0]!.finish(); ok(await px)
    const py = p.request(p.id(1, 'y')); await until(() => tc.calls.length === 2, 'y'); tc.calls[1]!.finish(); ok(await py)
    const s1 = synthQueue(w, 's1', 'd-poison', [['k']], tc)
    expect(await s1.request(s1.id(0, 'k'))).toEqual({ ok: false, reason: 'media-busy' })
    w.cache.delivered('p', itemOf('p', 'd-poison', 0, 'x').key) // x is now evictable
    // Another session encodes k; s1 then reports a hit, not a stale miss.
    const s2 = synthQueue(w, 's2', 'd-poison', [['k']], tc)
    const r2 = s2.request(s2.id(0, 'k')); await until(() => tc.calls.length === 3, 'k'); tc.calls[2]!.finish(); ok(await r2)
    expect(await s1.start(0)).toBe(0)
    expect(ok(await s1.request(s1.id(0, 'k'))).origin).toBe('hit')
    for (const [sid, u, n] of [['p', 1, 'y'], ['s1', 0, 'k'], ['s2', 0, 'k']] as const) w.cache.delivered(sid, itemOf(sid, 'd-poison', u, n).key)
    // A timed-out encode-ahead is not readied (a request encodes again) and does not loop.
    const t = synthQueue(w, 't', 'd-poison', [['z']], tc)
    await t.start(0)
    await until(() => tc.calls.length === 4, 'z'); tc.calls[3]!.fail('media-timeout')
    await Bun.sleep(10)
    expect(tc.calls.length).toBe(4)
    const again = t.request(t.id(0, 'z'))
    await until(() => tc.calls.length === 5, 'z again'); tc.calls[4]!.finish()
    expect(ok(await again).origin).toBe('miss')
    for (const q of [p, s1, s2, t]) await q.end()
    await w.cache.close()
  })

  // --- fix round 2: the re-review's scenarios -----------------------------------------------------------------

  it('a request arriving while its session holds a queue place keeps that place: no abort, no media-busy; an unrelated revocation disturbs nothing', async () => {
    const w = world({ slots: 1, queue: 1 })
    const tc = fakeTranscode({ manual: true })
    const a = synthQueue(w, 'sa', 'd-keep', [['a']], tc), b = synthQueue(w, 'sb', 'd-keep', [['b']], tc), c = synthQueue(w, 'sc', 'd-keep', [['c']], tc)
    await a.start(0)
    await until(() => tc.calls.length === 1, 'a runs')
    await b.start(0)
    await until(() => w.admission.queued === 1, 'b holds the place')
    await c.start(0) // parked
    await Bun.sleep(5)
    await w.cache.dropDocument('some-other-doc')
    await Bun.sleep(5)
    expect(w.admission.queued).toBe(1)
    const rb = b.request(b.id(0, 'b'))
    await Bun.sleep(5)
    expect(w.admission.queued).toBe(1) // still b's
    tc.calls[0]!.finish()
    await until(() => tc.calls.length === 2, 'the next encode')
    expect(tc.calls[1]!.name).toBe('b') // b's place, used for its request
    tc.calls[1]!.finish()
    expect((await rb).ok).toBe(true)
    await until(() => tc.calls.length === 3, 'c'); tc.calls[2]!.finish()
    await until(() => w.admission.running === 0 && w.admission.queued === 0, 'drain')
    for (const q of [a, b, c]) await q.end()
    await w.cache.close()
  })

  it('pipelined requests pin atomically: a cached answer is never dropped or evicted by the next request, and pins never exceed the look-ahead', async () => {
    // A: 'a' (50 MiB) cached; requesting a and the uncached b together.
    const w = world({ serverBytes: 60 * MiB, sessionBytes: 60 * MiB }), tc = fakeTranscode({ size: () => 50 * MiB })
    const o = synthQueue(w, 'o', 'd-atomic', [['a', 'b']], tc)
    ok(await o.request(o.id(0, 'a'))); w.cache.delivered('o', itemOf('o', 'd-atomic', 0, 'a').key)
    const q = synthQueue(w, 's', 'd-atomic', [['a', 'b']], tc)
    const [ra, rb] = await Promise.all([q.request(q.id(0, 'a')), q.request(q.id(0, 'b'))])
    expect(ra.ok).toBe(true)
    expect(ra.ok && existsSync(ra.path)).toBe(true)
    expect(rb).toEqual({ ok: false, reason: 'media-busy' }) // 50 pinned + 32 > 60
    await q.end(); await o.end(); await w.cache.close()
    // B: a, b, c (50 MiB each) cached; three requests together never hold more than the look-ahead.
    const w2 = world({ serverBytes: 200 * MiB, sessionBytes: 60 * MiB }), tc2 = fakeTranscode({ size: () => 50 * MiB })
    const o2 = synthQueue(w2, 'o', 'd-atomic2', [['a', 'b', 'c']], tc2)
    for (const n of ['a', 'b', 'c']) { ok(await o2.request(o2.id(0, n))); w2.cache.delivered('o', itemOf('o', 'd-atomic2', 0, n).key) }
    const q2 = synthQueue(w2, 's', 'd-atomic2', [['a', 'b', 'c']], tc2)
    const rs = await Promise.all(['a', 'b', 'c'].map((n) => q2.request(q2.id(0, n))))
    expect(rs.filter((r) => r.ok)).toHaveLength(1)
    expect(w2.cache.referencesOf('s')).toBe(1)
    for (const r of rs) if (r.ok) expect(existsSync(r.path)).toBe(true)
    await q2.end(); await o2.end(); await w2.cache.close()
  })

  it('passed() never releases an in-flight reference: a request on a slide the cursor leaves still answers and holds its reference', async () => {
    const w = world(), tc = fakeTranscode({ manual: true })
    const q = synthQueue(w, 's', 'd-pass', [['a'], ['b']], tc)
    const ra = q.request(q.id(0, 'a'))
    await until(() => tc.calls.length === 1, 'a')
    q.cursor(1)
    expect(w.cache.referencesOf('s')).toBe(1)
    tc.calls[0]!.finish()
    expect((await ra).ok).toBe(true)
    expect(w.cache.referencesOf('s')).toBeGreaterThanOrEqual(1)
    await q.end(); await w.cache.close()
  })

  it('a request miss takes its admission slot before it enters the cache: refused, it evicts and creates nothing, and a joining encode-ahead is not disabled', async () => {
    const w = world({ slots: 1, queue: 1, serverBytes: 64 * MiB }), tc = fakeTranscode({ manual: true, size: (n) => (n === 'old' ? 30 * MiB : 1) })
    const o = synthQueue(w, 'o', 'd-adm', [['old']], tc)
    const ro = o.request(o.id(0, 'old')); await until(() => tc.calls.length === 1, 'old'); tc.calls[0]!.finish(); ok(await ro)
    w.cache.delivered('o', itemOf('o', 'd-adm', 0, 'old').key) // unreferenced, evictable
    const a = synthQueue(w, 'sa', 'd-adm', [['a']], tc), b = synthQueue(w, 'sb', 'd-adm', [['b']], tc)
    await a.start(0); await until(() => tc.calls.length === 2, 'a runs')
    await b.start(0); await until(() => w.admission.queued === 1, 'b queued')
    const s1 = synthQueue(w, 's1', 'd-adm', [['c']], tc), s2 = synthQueue(w, 's2', 'd-adm', [['c']], tc)
    const [r1] = await Promise.all([s1.request(s1.id(0, 'c')), s2.start(0)])
    expect(r1).toEqual({ ok: false, reason: 'media-busy' })
    expect(w.cache.peek(itemOf('x', 'd-adm', 0, 'old').key)).toBeDefined() // nothing evicted for it
    expect(w.cache.inflight(itemOf('x', 'd-adm', 0, 'c').key)).toBe(false)
    tc.calls[1]!.finish()
    await until(() => tc.calls.length === 3, 'b'); tc.calls[2]!.finish()
    await until(() => tc.calls.some((x) => x.name === 'c'), "s2's encode-ahead of c")
    tc.calls.find((x) => x.name === 'c')!.finish()
    await until(() => w.admission.running === 0, 'drain')
    for (const q of [o, a, b, s1, s2]) await q.end()
    await w.cache.close()
  })

  it('start() after the document was revoked enqueues nothing', async () => {
    const w = world(), tc = fakeTranscode()
    await w.cache.dropDocument('d-gone')
    const q = synthQueue(w, 's', 'd-gone', [['a'], ['b']], tc)
    expect(await q.start(0)).toBe(0)
    await Bun.sleep(5)
    expect(tc.s.count).toBe(0)
    await q.end(); await w.cache.close()
  })

  // --- fix round 3: answers are confirmed in the step that resolves them --------------------------------------

  const ticks = async (n: number) => { for (let i = 0; i < n; i++) await Promise.resolve() }

  it('a cached answer is confirmed when it resolves: a revocation or end() in the window answers media-unavailable, never ok', async () => {
    for (const which of ['drop', 'end'] as const) for (let k = 0; k <= 6; k++) {
      const w = world(), tc = fakeTranscode()
      const o = synthQueue(w, 'o', 'd-win', [['a']], tc)
      ok(await o.request(o.id(0, 'a'))); w.cache.delivered('o', itemOf('o', 'd-win', 0, 'a').key)
      // Checked in the step that resolves the answer (the observe seam), so the caller's own await adds no window.
      const bad: unknown[] = []
      let ended = false
      const q = synthQueue(w, 's', 'd-win', [['a']], tc, fakePart, {
        observe: (e) => {
          if (e.type !== 'answer' || !e.result.ok) return
          const k2 = itemOf('s', 'd-win', 0, 'a').key
          if (ended || w.cache.revoked('d-win') || !w.cache.peek(k2) || w.cache.pinsOf('s', k2) < 1 || !existsSync(e.result.path)) bad.push({ which, k })
        },
      })
      await q.request('not-an-id') // loads the items
      await Promise.all([q.request(q.id(0, 'a')), ticks(k).then(() => { if (which === 'drop') return w.cache.dropDocument('d-win'); ended = true; return q.end() })])
      expect(bad).toEqual([])
      await q.end(); await o.end(); await w.cache.close()
    }
  })

  it('a request joining an in-flight encode-ahead is pinned at once: a request in the settle window never drops it', async () => {
    for (let k = 0; k <= 12; k++) {
      const w = world({ sessionBytes: 50 * MiB, serverBytes: 128 * MiB }), tc = fakeTranscode({ manual: true, size: () => 30 * MiB })
      const memPart = async () => ({ path: '/mem/part', dispose: async () => {} }) // microtask-only, so the window is addressable
      const s1 = synthQueue(w, 's1', 'd-join', [['a'], ['b']], tc, memPart)
      await s1.start(0) // encode-ahead runs a
      await until(() => tc.calls.length === 1, 'a runs')
      // Another session caches b (delivered: unreferenced).
      const o = synthQueue(w, 'o', 'd-join', [['a'], ['b']], tc, memPart)
      const rbo = o.request(o.id(1, 'b')); await until(() => tc.calls.length === 2, 'b'); tc.calls[1]!.finish(); ok(await rbo)
      w.cache.delivered('o', itemOf('o', 'd-join', 1, 'b').key)
      const ra = s1.request(s1.id(0, 'a')) // joins the running encode-ahead
      await Bun.sleep(1)
      tc.calls[0]!.finish()
      const rb = ticks(k).then(() => s1.request(s1.id(1, 'b'))) // in the window after a settles
      const [a, b] = await Promise.all([ra, rb])
      void b
      expect({ k, a: a.ok ? 'ok' : a.reason }).toEqual({ k, a: 'ok' })
      await s1.end(); await o.end(); await w.cache.close()
    }
  })

  it('start() counts nothing when the document is revoked or the queue ends while it takes cached references', async () => {
    for (const which of ['drop', 'end'] as const) for (let k = 0; k <= 6; k++) {
      const w = world(), tc = fakeTranscode()
      const o = synthQueue(w, 'o', 'd-cnt', [['a', 'b']], tc)
      ok(await o.request(o.id(0, 'a'))); w.cache.delivered('o', itemOf('o', 'd-cnt', 0, 'a').key)
      const bad: unknown[] = []
      let ended = false
      const q = synthQueue(w, 's', 'd-cnt', [['a', 'b']], tc, fakePart, {
        observe: (e) => { if (e.type === 'start' && e.count > 0 && (ended || w.cache.revoked('d-cnt'))) bad.push({ which, k, n: e.count }) },
      })
      await q.request('not-an-id')
      const [n] = await Promise.all([q.start(0), ticks(k).then(() => { if (which === 'drop') return w.cache.dropDocument('d-cnt'); ended = true; return q.end() })])
      expect(bad).toEqual([])
      expect([0, 1]).toContain(n)
      await q.end(); await o.end(); await w.cache.close()
    }
  })

  it('MEDIA_SERVER_CACHE is 2 GiB and positive', () => {
    expect(MEDIA_SERVER_CACHE).toBe(2147483648)
    expect(MEDIA_SERVER_CACHE).toBeGreaterThan(0)
    expect(MEDIA_SESSION_LOOKAHEAD).toBe(256 * MiB)
  })
})

describe.skipIf(!info)('the queue with real parts', () => {
  it('two sessions requesting the same part start exactly one encode', async () => {
    let w!: World
    w = world()
    const tc = fakeTranscode({ manual: true, admission: () => w.admission })
    const doc = await openDoc(await oneAudioDeck())
    const s1 = session(doc, 'docA'), s2 = session(doc, 'docA')
    const q1 = realQueue(w, 'q1', s1, tc.fn), q2 = realQueue(w, 'q2', s2, tc.fn)
    const [n1, n2] = [await q1.start(0), await q2.start(0)]
    const [i1] = mediaJobsFor(await doc.h.mediaRefs(doc.docKey, 0), 0, s1, info!)
    const [i2] = mediaJobsFor(await doc.h.mediaRefs(doc.docKey, 0), 0, s2, info!)
    expect(i1!.mediaId).not.toBe(i2!.mediaId)
    const r1 = q1.request(i1!.mediaId), r2 = q2.request(i2!.mediaId)
    await until(() => tc.calls.length === 1, 'the encode')
    await until(() => q2.running !== undefined, 'q2 joins')
    tc.calls[0]!.finish()
    const [a, b] = [ok(await r1), ok(await r2)]
    expect(n1 + n2).toBe(1)
    expect(tc.s.count).toBe(1)
    expect([a.origin, b.origin]).toEqual(['miss', 'joined'])
    expect(readFileSync(a.path)).toEqual(readFileSync(b.path))
    expect(tc.s.maxRunning).toBeLessThanOrEqual(1)
    expect(w.admission.running).toBe(0)
    await q1.end(); await q2.end(); await w.cache.close()
  })

  it('rapid jumps re-queue with one running group and no orphan pid or temp file', async () => {
    const w = world()
    const live0 = temp.live()
    const files = ['a.m4a', 'a.mp3', 'a.wav', 'a.wma'] as const
    const bytes = await mediaDeck({ slides: files.map((file) => ({ boxes: [box], media: [{ spid: 3, file }], timing: audioNode(3, tgt(3)) })) })
    const doc = await openDoc(bytes)
    const s = { active: 0, max: 0, calls: 0 }
    const counted: typeof transcode = async (...a) => { s.calls++; s.active++; s.max = Math.max(s.max, s.active); try { return await transcode(...a) } finally { s.active-- } }
    const q = realQueue(w, 'rj', session(doc, 'docJ'), counted)
    await q.start(0)
    await until(() => s.active === 1, 'an encode runs')
    q.cursor(3); q.cursor(1); q.cursor(3) // goTo 9, 2, 9 within one second
    await until(() => s.calls >= 2, 'the next encode')
    expect(s.max).toBe(1)
    await q.end()
    expect(s.active).toBe(0)
    const pgrep = () => { try { return execFileSync('/usr/bin/pgrep', ['-fl', temp.dir], { encoding: 'utf8' }).trim() } catch { return '' } }
    expect(pgrep()).toBe('')
    expect(w.admission.running).toBe(0)
    expect(temp.live()).toBe(live0 + w.cache.stats().entries) // cached outputs only
    await w.cache.close()
    expect(temp.live()).toBe(live0)
  })

  it('createMediaServices: close() ends queues, closes the cache and removes the temp dir', async () => {
    const r = mkdtempSync(join(tmpdir(), 'ov-services-'))
    const services = await createMediaServices({ env: process.env, enabled: true, testOnly: { root: r, slots: 1 } })
    try {
      expect(services.info).toBeDefined()
      expect(existsSync(services.temp.dir)).toBe(true)
      const h = new DocumentHost(); h.configureMedia({ info: services.info!, dir: services.temp.dir })
      const { docKey, units } = await h.open('svc', await oneAudioDeck(), 'visible-only')
      const s = { mediaKey: new Uint8Array(randomBytes(32)), docId: 'svc' }
      const q = services.createQueue({
        sid: 'svc', units: units.length,
        items: async (u) => mediaJobsFor(await h.mediaRefs(docKey, u), u, s, services.info!),
        part: async (pp) => { const x = await h.mediaPart(docKey, pp, services.temp.dir); return services.temp.adopt(x.path, x.bytes) },
      })
      const [item] = mediaJobsFor(await h.mediaRefs(docKey, 0), 0, s, services.info!)
      const got = ok(await q.request(item!.mediaId))
      expect(existsSync(got.path)).toBe(true)
      expect(services.cache.stats().entries).toBe(1)
      await services.close()
      expect(existsSync(services.temp.dir)).toBe(false)
      expect(services.temp.live()).toBe(0)
      expect(await q.request(item!.mediaId)).toEqual({ ok: false, reason: 'media-unavailable' })
      expect(() => services.createQueue({ sid: 'late', units: 1, items: async () => [], part: fakePart })).toThrow()
    } finally { await services.close(); rmSync(r, { recursive: true, force: true }) }
  })

  it('createMediaServices with media disabled probes nothing and touches no disk', async () => {
    const services = await createMediaServices({ env: process.env, enabled: false })
    expect(services.info).toBeUndefined()
    expect(services.temp.dir).toBe('')
    expect(() => services.createQueue({ sid: 'x', units: 1, items: async () => [], part: fakePart })).toThrow()
    await services.close()
  })
})
