/**
 * Piece 5 Task 9 (spec §4.3, §7; plan D1, D12, D13, D14, D23): the media wire through the harness and real ffmpeg.
 * `media-show` starts the session's queue; `media-request` answers `media-start` then exactly `chunks` stored
 * MEDIA_CHUNK messages, or `media-refused`; the `media` rate class; the `media_deliveries` audit table.
 */
import { afterAll, describe, expect, it } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { decodeShowMessage, MEDIA_ID, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../../src/stream/show-plan'
import { MEDIA_WIRE_LIMITS, parseMediaCursor, parseMediaRefused, parseMediaRequest, parseMediaShow, parseMediaShowOk, parseMediaStart } from '../../../src/stream/media-wire'
import { mediaDeck, mixedMediaDeck, narrationDeck } from '../../../src/testdata/show-decks'
import { encodeKeyId } from '../../src/media/params'
import { Client, disposePool, start, type Harness, type HarnessMedia } from '../harness'
import { ffmpegForTests } from './ffmpeg-gate'

const info = await ffmpegForTests()
const scratch = mkdtempSync(join(tmpdir(), 'ov-t9-src-'))
afterAll(() => { disposePool(); rmSync(scratch, { recursive: true, force: true }) })

/** An m4a source made by ffmpeg from a lavfi graph (test input only, never product code). */
async function source(name: string, graph: string, seconds: number, bitrate = '64k'): Promise<Uint8Array> {
  const out = join(scratch, name)
  const p = Bun.spawn([info!.ffmpeg, '-nostdin', '-hide_banner', '-loglevel', 'error', '-y', '-f', 'lavfi', '-i', graph, '-t', String(seconds), '-c:a', 'aac', '-b:a', bitrate, '-f', 'ipod', out], { stdout: 'ignore', stderr: 'pipe' })
  if (await p.exited !== 0) throw new Error(`ffmpeg: ${await new Response(p.stderr).text()}`)
  return new Uint8Array(readFileSync(out))
}
/** `deck` with its one shape-audio part replaced by `bytes`. */
async function withPart(deck: Uint8Array, bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(deck)
  const parts = Object.keys(zip.files).filter((f) => /^ppt\/media\/showMedia\d+\.m4a$/.test(f))
  if (parts.length !== 1) throw new Error(`expected one m4a part, got ${parts.join(',')}`)
  zip.file(parts[0]!, bytes)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}
let noisy: Promise<Uint8Array> | undefined
/** ~150 s of noise: its 128 kb/s encode is about 2.4 MB, three chunks. */
const noiseDeck = () => (noisy ??= (async () => withPart(await narrationDeck(), await source('noise.m4a', 'anoisesrc=c=pink:r=48000:a=0.5', 150)))())

const box = { id: 2, name: 'Box', prst: 'rect', off: ['3619500', '2476500'] as [string, string], ext: ['1905000', '1905000'] as [string, string], fill: '336699' }
const sndNode = (id: number, start = '<p:cond delay="0"/>') => `<p:audio><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst>${start}</p:stCondLst></p:cTn><p:tgtEl><p:sndTgt r:embed="rIdSnd1" name="a.wav"/></p:tgtEl></p:cMediaNode></p:audio>`
const videoNode = (id: number, spid: number) => `<p:video><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:video>`
/**
 * One sound part (spid-less, identity `audio|k`: one mediaId on both slides). Slide 0: the sound begins on a video's
 * end (D22: refused media-unavailable there). Slide 1: the same part begins with the slide (playable).
 */
async function sharedSoundDeck(): Promise<Uint8Array> {
  const bytes = await mediaDeck({ slides: [
    { boxes: [box], media: [{ spid: 3, file: 'av.mp4', kind: 'video' }], sounds: [{ rId: 'rIdSnd1', file: 'a.wav' }], timing: videoNode(3, 3) + sndNode(4, '<p:cond evt="onEnd" delay="0"><p:tn val="3"/></p:cond>') },
    { boxes: [box], sounds: [{ rId: 'rIdSnd1', file: 'a.wav' }], timing: sndNode(3) },
  ] })
  const zip = await JSZip.loadAsync(bytes)
  const rels = 'ppt/slides/_rels/slide2.xml.rels'
  zip.file(rels, (await zip.file(rels)!.async('string')).replace('../media/sound2.wav', '../media/sound1.wav'))
  zip.remove('ppt/media/sound2.wav')
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace(/<Override PartName="\/ppt\/media\/sound2\.wav"[^>]*\/>/, ''))
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}
/** A transport that holds every kind-9 send after the first under backpressure while `held()`. */
function heldTransport(held: () => boolean, onFirst: () => void = () => {}) {
  let sent = 0
  return {
    sent: () => sent,
    transport: {
      sendBinary: (ws: unknown, f: Uint8Array) => { const n = (ws as { sendBinary(f: Uint8Array): number }).sendBinary(f); if (f[5] === 9 && ++sent === 1) onFirst(); return n },
      bufferedAmount: (ws: unknown) => (sent >= 1 && held() ? 64 << 20 : (ws as { getBufferedAmount(): number }).getBufferedAmount()),
    },
  }
}

async function harness(doc: Uint8Array, media: HarnessMedia = {}, extra: Parameters<typeof start>[0] = {}): Promise<Harness> {
  return await start({ doc, media, ...extra })
}
async function session(h: Harness): Promise<Client> {
  const c = new Client(h.url)
  await c.hello(); await c.ready()
  return c
}
async function planMedia(c: Client, unit = 0) {
  const r = await c.showPlan(unit, 1, 'full')
  if (r.end) throw new Error(`show-plan refused ${r.end.code}`)
  return (decodeShowMessage(r.messages[0]!, { unit, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan.media ?? []
}
async function idOf(c: Client, unit = 0) { return (await planMedia(c, unit))[0]!.mediaId }
const sidOf = (c: Client) => c.manifest!.session.sid
const concat = (cs: Uint8Array[]) => { const out = new Uint8Array(cs.reduce((n, c) => n + c.length, 0)); let at = 0; for (const c of cs) { out.set(c, at); at += c.length } return out }
/** Waits until `f()` holds (polling), or throws. */
async function until(f: () => boolean, ms = 10_000) { const end = Date.now() + ms; while (!f()) { if (Date.now() > end) throw new Error('timed out'); await Bun.sleep(20) } }
/** The refusal of a request that never started a delivery. */
function refusal(r: Awaited<ReturnType<Client['mediaRequest']>>) {
  expect(r.start).toBeUndefined()
  expect(r.chunks).toHaveLength(0)
  return parseMediaRefused(r.end).reason
}

describe('media wire validators (Task 9)', () => {
  const ok = { reqId: 3, type: 'media-start', mediaId: 'A'.repeat(22), contentType: 'audio/mp4', durationMs: 2000, bytes: 1048513, chunks: 2 }
  it('parseMediaStart: the exact key set, audio/mp4, bytes ≤ audioBytes, chunks = ceil(bytes / sendChunk)', () => {
    expect(parseMediaStart(ok)).toEqual(ok as never)
    expect(MEDIA_WIRE_LIMITS).toEqual({ chunkBytes: 4194304, sendChunk: 1048512, audioBytes: 33554432, textBytes: 65536 })
    for (const bad of [{ ...ok, extra: 1 }, { ...ok, contentType: 'video/mp4' }, { ...ok, bytes: MEDIA_WIRE_LIMITS.audioBytes + 1, chunks: 33 }, { ...ok, chunks: 1 }, { ...ok, chunks: 3 }, { ...ok, bytes: 0, chunks: 0 }, { ...ok, mediaId: 'short' }, { ...ok, type: 'media-show-ok' }, { ...ok, durationMs: -1 }]) expect(() => parseMediaStart(bad)).toThrow()
    expect(parseMediaStart({ ...ok, bytes: MEDIA_WIRE_LIMITS.audioBytes, chunks: 33 }).chunks).toBe(33)
  })
  it('client and server parsers refuse unknown keys and bad values', () => {
    expect(parseMediaShowOk({ reqId: 1, type: 'media-show-ok', queued: 0 })).toEqual({ reqId: 1, type: 'media-show-ok', queued: 0 })
    expect(() => parseMediaShowOk({ reqId: 1, type: 'media-show-ok', queued: -1 })).toThrow()
    expect(parseMediaRefused({ reqId: 1, type: 'media-refused', reason: 'invalid-media' })).toEqual({ reqId: 1, type: 'media-refused', reason: 'invalid-media' })
    expect(() => parseMediaRefused({ reqId: 1, type: 'media-refused', reason: 'nope' })).toThrow()
    expect(() => parseMediaRefused({ reqId: 1, type: 'media-refused', reason: 'media-busy', code: 1 })).toThrow()
    expect(parseMediaShow({ reqId: 1, type: 'media-show', token: 't', unit: 2 })).toEqual({ reqId: 1, token: 't', unit: 2 })
    expect(() => parseMediaShow({ reqId: 1, type: 'media-show', token: 't', unit: 2, scale: 1 })).toThrow()
    expect(parseMediaCursor({ type: 'media-cursor', unit: 0 })).toEqual({ unit: 0 })
    expect(() => parseMediaCursor({ type: 'media-cursor', unit: 1.5 })).toThrow()
    expect(parseMediaRequest({ reqId: 1, type: 'media-request', token: 't', mediaId: 'x' }).mediaId).toBe('x')
    expect(() => parseMediaRequest({ reqId: 1, type: 'media-request', token: 't', mediaId: 7 })).toThrow()
  })
})

describe.skipIf(!info)('media wire through the harness (Task 9)', () => {
  it('media-show then media-request: media-start then exactly chunks MEDIA_CHUNK, bytes sum, each ≤ sendChunk', async () => {
    const h = await harness(await noiseDeck())
    const c = await session(h)
    try {
      expect(c.manifest!.can.media).toBe(true)
      const id = await idOf(c)
      expect(id).toMatch(MEDIA_ID)
      const show = await c.mediaShow(0)
      expect(parseMediaShowOk(show).type).toBe('media-show-ok')
      const r = await c.mediaRequest(id, { ms: 60_000 })
      expect(r.end).toBeUndefined()
      const st = parseMediaStart(r.start)
      expect(st.mediaId).toBe(id)
      expect(st.chunks).toBeGreaterThanOrEqual(2)
      expect(r.chunks).toHaveLength(st.chunks)
      expect(r.chunks.reduce((n, x) => n + x.length, 0)).toBe(st.bytes)
      for (const x of r.chunks) expect(x.length).toBeLessThanOrEqual(MEDIA_WIRE_LIMITS.sendChunk)
      // The bytes are the cached encode's, and no source byte (an M4A of its own, from ftyp).
      const entry = h.media!.cache.inspect().entries.find((e) => e.path)!
      expect(concat(r.chunks)).toEqual(new Uint8Array(readFileSync(entry.path!)))
      // Nothing more on that reqId.
      await Bun.sleep(100)
      expect(c.msgs.filter((m) => ('text' in m ? m.text.reqId : m.binary.reqId) === r.reqId)).toHaveLength(0)
      // Delivered: the pin, and with it the session's reference, is released (D12: a delivery releases it).
      expect(h.media!.cache.pinsOf(sidOf(c), { docId: c.manifest!.docRef, sourceHash: h.stream.audit.mediaDeliveries(sidOf(c))[0]!.sourceHash!, paramsSha: h.stream.audit.mediaDeliveries(sidOf(c))[0]!.encodeParamsSha!, ffmpegVersion: info!.version })).toBe(0)
      expect(h.media!.cache.referencesOf(sidOf(c))).toBe(0)
      expect(h.media!.cache.stats().entries).toBe(1)
    } finally { c.close(); await h.stop() }
  }, 120_000)

  it('media-request before media-show → media-refused invalid-media', async () => {
    const h = await harness(await narrationDeck())
    const c = await session(h)
    try {
      const id = await idOf(c)
      expect(refusal(await c.mediaRequest(id))).toBe('invalid-media')
      expect(h.stream.audit.mediaDeliveries(sidOf(c))).toEqual([])
    } finally { c.close(); await h.stop() }
  }, 60_000)

  it('foreign mediaId → invalid-media', async () => {
    const h = await harness(await narrationDeck())
    const c = await session(h)
    try {
      await idOf(c)
      await c.mediaShow(0)
      for (const bogus of ['A'.repeat(22), 'not-an-id', '@a0.3']) {
        const r = await c.mediaRequest(bogus)
        expect(refusal(r)).toBe('invalid-media')
        // The id is echoed only when well-formed.
        expect('mediaId' in r.end!).toBe(MEDIA_ID.test(bogus))
      }
      expect(h.stream.audit.mediaDeliveries(sidOf(c))).toEqual([])
    } finally { c.close(); await h.stop() }
  }, 60_000)

  it("7th media-request in 60 s → rate-limited; never closes 4008; unit budget untouched", async () => {
    const h = await harness(await narrationDeck())
    const c = await session(h)
    try {
      const id = await idOf(c)
      await c.mediaShow(0)
      for (let i = 0; i < 6; i++) expect((await c.mediaRequest(id, { ms: 60_000 })).end).toBeUndefined()
      for (let i = 0; i < 6; i++) expect(refusal(await c.mediaRequest(id))).toBe('rate-limited')
      expect(c.ws.readyState).toBe(WebSocket.OPEN)
      // The unit budget is untouched: a unit and a show-plan still answer.
      c.unit(77, 0, 1)
      expect((await c.unitMessages(77)).messages.length).toBeGreaterThan(0)
      expect((await c.showPlan(0, 1, 'full')).end).toBeUndefined()
      const rows = h.stream.audit.mediaDeliveries(sidOf(c))
      expect(rows.map((r) => r.reason ?? 'ok')).toEqual([...Array(6).fill('ok'), ...Array(6).fill('rate-limited')])
      expect(rows.slice(6).every((r) => r.cacheState === undefined)).toBe(true)
    } finally { c.close(); await h.stop() }
  }, 120_000)

  it('can.media false → media-unavailable', async () => {
    const h = await harness(await narrationDeck(), { enabled: false })
    const c = await session(h)
    try {
      expect(c.manifest!.can.media).toBe(false)
      expect(parseMediaRefused(await c.mediaShow(0)).reason).toBe('media-unavailable')
      expect(refusal(await c.mediaRequest('A'.repeat(22)))).toBe('media-unavailable')
      // Off, nothing touched the disk and no queue exists.
      expect(h.media!.info).toBeUndefined()
      expect(h.media!.cache.stats().entries).toBe(0)
    } finally { c.close(); await h.stop() }
  }, 60_000)

  it('revoked mid-delivery stops sending', async () => {
    let sent = 0, revoke: (() => void) | undefined
    const h = await harness(await noiseDeck(), {}, {
      transport: {
        // The first chunk goes out, then the session is revoked before the next.
        sendBinary: (ws, f) => { const n = (ws as { sendBinary(f: Uint8Array): number }).sendBinary(f); if (f[5] === 9 && ++sent === 1) revoke?.(); return n },
        bufferedAmount: (ws) => (ws as { getBufferedAmount(): number }).getBufferedAmount(),
      },
    })
    const c = await session(h)
    try {
      const id = await idOf(c)
      await c.mediaShow(0)
      revoke = () => h.stream.revoke(sidOf(c))
      c.send({ reqId: 4000, type: 'media-request', token: c.manifest!.session.token, mediaId: id })
      expect((await c.closed).code).toBe(4005)
      await Bun.sleep(100)
      expect(sent).toBe(1)
      const start = c.msgs.find((m) => 'text' in m && m.text.type === 'media-start') as { text: Record<string, any> }
      expect(start.text.chunks).toBeGreaterThan(1)
      expect(c.msgs.filter((m) => 'binary' in m && m.binary.kind === 9)).toHaveLength(1)
      await until(() => h.media!.cache.referencesOf(sidOf(c)) === 0)
    } finally { c.close(); await h.stop() }
  }, 120_000)

  it('media_deliveries row has exactly the twelve columns; ffmpeg_version and encode_params_sha match the result; cache_state is miss', async () => {
    const h = await harness(await narrationDeck())
    const c = await session(h)
    try {
      expect(h.stream.audit.mediaColumns()).toEqual(['sid', 'unit_index', 'media_id', 'kind', 'source_hash', 'out_bytes', 'duration_ms', 'reason', 'at', 'ffmpeg_version', 'encode_params_sha', 'cache_state'])
      const rec = (await planMedia(c))[0]!
      await c.mediaShow(0)
      const r = await c.mediaRequest(rec.mediaId, { ms: 60_000 })
      const st = parseMediaStart(r.start)
      const [row] = h.stream.audit.mediaDeliveries(sidOf(c))
      const key = h.media!.cache.inspect().entries[0]!
      expect(row).toEqual({
        sid: sidOf(c), unitIndex: 0, mediaId: rec.mediaId, kind: 'audio', sourceHash: expect.stringMatching(/^[0-9a-f]{64}$/), outBytes: st.bytes, durationMs: st.durationMs,
        at: expect.any(Number), ffmpegVersion: info!.version, encodeParamsSha: expect.stringMatching(/^[0-9a-f]{64}$/), cacheState: 'miss',
      })
      expect(st.durationMs).toBe(rec.durationMs)
      // The key the row names is the cache entry's.
      expect(encodeKeyId({ docId: c.manifest!.docRef, sourceHash: row!.sourceHash!, paramsSha: row!.encodeParamsSha!, ffmpegVersion: row!.ffmpegVersion! })).toBe(key.id)
      expect(h.stream.audit.mediaDeliveryById(rec.mediaId)).toEqual(row)
    } finally { c.close(); await h.stop() }
  }, 60_000)

  it('two sessions of one document: the second media-request starts no encode; its row reads joined or hit, the first reads miss; the bytes are identical and the mediaIds differ', async () => {
    const h = await harness(await narrationDeck())
    const a = await session(h), b = await session(h)
    try {
      const ia = await idOf(a), ib = await idOf(b)
      expect(ia).not.toBe(ib)
      await a.mediaShow(0)
      const ra = await a.mediaRequest(ia, { ms: 60_000 })
      const before = h.media!.cache.inspect().entries
      expect(before).toHaveLength(1)
      await b.mediaShow(0)
      const rb = await b.mediaRequest(ib, { ms: 60_000 })
      const after = h.media!.cache.inspect().entries
      expect(after).toEqual([{ ...before[0]!, refs: after[0]!.refs }]) // the same entry, the same file: no encode
      expect(concat(rb.chunks)).toEqual(concat(ra.chunks))
      expect(h.stream.audit.mediaDeliveries(sidOf(a)).map((r) => r.cacheState)).toEqual(['miss'])
      expect(['joined', 'hit']).toContain(h.stream.audit.mediaDeliveries(sidOf(b))[0]!.cacheState!)
      expect(h.stream.audit.mediaDeliveries(sidOf(b))[0]!.mediaId).toBe(ib)
    } finally { a.close(); b.close(); await h.stop() }
  }, 60_000)

  it("session B requesting session A's mediaId, or a cache key, → invalid-media; no row reads from the cache", async () => {
    const h = await harness(await narrationDeck())
    const a = await session(h), b = await session(h)
    try {
      const ia = await idOf(a)
      await idOf(b)
      await a.mediaShow(0); await b.mediaShow(0)
      await a.mediaRequest(ia, { ms: 60_000 })
      const keyId = h.media!.cache.inspect().entries[0]!.id
      expect(refusal(await b.mediaRequest(ia))).toBe('invalid-media')
      expect(refusal(await b.mediaRequest(keyId))).toBe('invalid-media')
      expect(refusal(await b.mediaRequest(keyId.slice(0, 22)))).toBe('invalid-media')
      expect(h.stream.audit.mediaDeliveries(sidOf(b))).toEqual([])
      expect(h.stream.audit.mediaDeliveryById(ia)!.sid).toBe(sidOf(a))
    } finally { a.close(); b.close(); await h.stop() }
  }, 60_000)

  it("document revoked mid-delivery: the send stops, the client gets media-refused media-unavailable, and the document's entries are gone", async () => {
    let sent = 0, revoke: (() => void) | undefined
    const h = await harness(await noiseDeck(), {}, {
      transport: {
        sendBinary: (ws, f) => { if (f[5] === 9 && ++sent === 1) revoke?.(); return (ws as { sendBinary(f: Uint8Array): number }).sendBinary(f) },
        bufferedAmount: (ws) => (ws as { getBufferedAmount(): number }).getBufferedAmount(),
      },
    })
    const c = await session(h)
    try {
      const id = await idOf(c)
      await c.mediaShow(0)
      const docRef = c.manifest!.docRef
      revoke = () => { void h.media!.cache.dropDocument(docRef) }
      const r = await c.mediaRequest(id, { ms: 60_000 })
      expect(r.start!.chunks).toBeGreaterThan(1)
      expect(r.chunks).toHaveLength(1)
      expect(parseMediaRefused(r.end).reason).toBe('media-unavailable')
      expect(r.end!.mediaId).toBe(id)
      expect(sent).toBe(1)
      expect(h.media!.cache.inspect().entries.filter((e) => e.docId === docRef)).toEqual([])
      // The rest of the show sees media-unavailable.
      expect(refusal(await c.mediaRequest(id))).toBe('media-unavailable')
    } finally { c.close(); await h.stop() }
  }, 120_000)

  it('services row records media on/off', async () => {
    for (const enabled of [true, false]) {
      const h = await harness(await narrationDeck(), { enabled })
      const c = await session(h)
      try {
        expect(h.stream.audit.get(sidOf(c))!.services.media).toBe(enabled)
      } finally { c.close(); await h.stop() }
    }
  }, 60_000)

  it('media-cursor flood is message-rate-limited only', async () => {
    const h = await harness(await narrationDeck())
    const c = await session(h)
    try {
      const id = await idOf(c)
      await c.mediaShow(0)
      // Under the message burst (20/s): no reply at all, and the media class is not charged.
      for (let i = 0; i < 15; i++) c.mediaCursor(i % 2 ? 0 : 99)
      await Bun.sleep(1100)
      for (let i = 0; i < 6; i++) expect((await c.mediaRequest(id, { ms: 60_000 })).end).toBeUndefined()
      expect(c.msgs.filter((m) => 'text' in m && m.text.reqId === 0)).toEqual([])
      await Bun.sleep(1100)
      // Over it: the message class refuses, then closes 4008 (the media class plays no part).
      for (let i = 0; i < 40; i++) c.mediaCursor(0)
      expect((await c.closed).code).toBe(4008)
      expect(c.msgs.some((m) => 'text' in m && m.text.type === 'refused' && m.text.code === 'rate-limited')).toBe(true)
    } finally { c.close(); await h.stop() }
  }, 60_000)

  describe('refusals', () => {
    async function refusedFor(doc: Uint8Array, media: HarnessMedia = {}, before?: (h: Harness) => Promise<() => void>) {
      const h = await harness(doc, media)
      const c = await session(h)
      const after = before ? await before(h) : () => {}
      try {
        const rec = (await planMedia(c))[0]!
        await c.mediaShow(0)
        const reason = refusal(await c.mediaRequest(rec.mediaId, { ms: 60_000 }))
        const rows = h.stream.audit.mediaDeliveries(sidOf(c))
        expect(rows).toHaveLength(1)
        expect(rows[0]!.reason).toBe(reason)
        expect(rows[0]!.cacheState).toBeUndefined()
        return { reason, rec }
      } finally { after(); c.close(); await h.stop() }
    }
    it('refusal media-too-long', async () => {
      const long = await withPart(await narrationDeck(), await source('long.m4a', 'anullsrc=r=8000:cl=mono', 31 * 60, '16k'))
      const { reason, rec } = await refusedFor(long)
      expect(rec.refused).toBe('media-too-long')
      expect(reason).toBe('media-too-long')
    }, 120_000)
    it('refusal media-too-large', async () => {
      expect((await refusedFor(await narrationDeck(), { testOnly: { maxBytes: 1000 } })).reason).toBe('media-too-large')
    }, 60_000)
    it('refusal media-timeout', async () => {
      expect((await refusedFor(await narrationDeck(), { testOnly: { wallCapMs: 1 } })).reason).toBe('media-timeout')
    }, 60_000)
    it('refusal media-busy', async () => {
      const { reason } = await refusedFor(await narrationDeck(), { testOnly: { slots: 1, queue: 0 } }, async (h) => {
        const held = h.media!.admission.slot('holder', 'request', new AbortController().signal)
        if ('refused' in held) throw new Error('no slot to hold')
        const lease = await held
        return () => lease.release()
      })
      expect(reason).toBe('media-busy')
    }, 60_000)
    it('refusal media-unreadable', async () => {
      const junk = new Uint8Array(4096).map((_, i) => (i * 131 + 7) & 0xff)
      const { reason, rec } = await refusedFor(await withPart(await narrationDeck(), junk))
      expect(rec.refused).toBe('media-unreadable')
      expect(reason).toBe('media-unreadable')
    }, 60_000)
    it('refusal media-linked', async () => {
      const zip = await JSZip.loadAsync(await narrationDeck())
      const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
      zip.file(relsPath, (await zip.file(relsPath)!.async('string')).replace(/Target="\.\.\/media\/[^"]+"/g, 'Target="https://example.com/x.m4a" TargetMode="External"'))
      const { reason, rec } = await refusedFor(new Uint8Array(await zip.generateAsync({ type: 'uint8array' })))
      expect(rec.refused).toBe('media-linked')
      expect(reason).toBe('media-linked')
    }, 60_000)
    it('refusal media-unavailable', async () => {
      // A narration chained on a video's end: the build ships its record refused media-unavailable (its part probes fine).
      const h = await harness(await mixedMediaDeck('chain'))
      const c = await session(h)
      try {
        const rec = (await planMedia(c)).find((r) => r.refused)!
        expect(rec.refused).toBe('media-unavailable')
        await c.mediaShow(0)
        expect(refusal(await c.mediaRequest(rec.mediaId))).toBe('media-unavailable')
        expect(h.stream.audit.mediaDeliveries(sidOf(c)).map((r) => [r.reason, r.cacheState])).toEqual([['media-unavailable', undefined]])
      } finally { c.close(); await h.stop() }
    }, 60_000)
  })

  it("session expiry ends the queue and drops the session's references; the cache entries remain", async () => {
    // Expiry: a token whose renewal is never acknowledged dies at its exp (4004). B acknowledges, so it holds the
    // document open (closing the document would delete its entries, D12).
    const h = await harness(await narrationDeck(), {}, { tokenTtlSec: 3, renewLeadSec: 2 })
    const a = await session(h), b = await session(h)
    try {
      await idOf(a)
      // Encode-ahead readies the slide's audio and A references it (undelivered).
      expect(parseMediaShowOk(await a.mediaShow(0)).queued).toBe(1)
      await until(() => h.media!.cache.stats().entries === 1 && h.media!.cache.referencesOf(sidOf(a)) > 0)
      const renew = await b.next((m) => 'text' in m && m.text.type === 'renew', 10_000) as { text: Record<string, any> }
      b.send({ reqId: 0, type: 'renew-ack', jti: renew.text.jti })
      expect((await a.closed).code).toBe(4004)
      await until(() => h.media!.cache.referencesOf(sidOf(a)) === 0)
      expect(h.media!.cache.stats().entries).toBe(1)
      expect(h.media!.cache.stats().inflight).toBe(0)
      expect(h.media!.admission.snapshot().running).toBe(0)
    } finally { a.close(); b.close(); await h.stop() }
  }, 60_000)

  it('a use the build refuses on every slide is never encoded ahead, and is refused whatever the order of show-plan and media-request', async () => {
    const h = await harness(await mixedMediaDeck('chain'))
    const a = await session(h)
    let resumed: Client | undefined
    try {
      const rec = (await planMedia(a)).find((r) => r.refused)!
      // A second session that never fetched the plan learns nothing new: use A's resume to request before show-plan.
      expect(parseMediaShowOk(await a.mediaShow(0)).queued).toBe(0)
      await Bun.sleep(300)
      expect(h.media!.cache.stats()).toMatchObject({ entries: 0, inflight: 0 })
      const m0 = a.manifest!
      const ws = a.ws as WebSocket & { terminate?: () => void }
      if (ws.terminate) ws.terminate(); else ws.close(1001)
      await a.closed; await Bun.sleep(50)
      resumed = new Client(h.url)
      await resumed.hello({ resume: { sid: m0.session.sid, resumeProof: m0.session.token } }); await resumed.ready()
      await resumed.mediaShow(0)
      // No show-plan on this socket: still refused, no bytes.
      expect(refusal(await resumed.mediaRequest(rec.mediaId))).toBe('media-unavailable')
      expect(h.media!.cache.stats()).toMatchObject({ entries: 0, inflight: 0 })
    } finally { resumed?.close(); a.close(); await h.stop() }
  }, 60_000)

  it('a refusal is per (mediaId, slide): a sound refused on slide 0 plays on slide 1; the row records the slide requested', async () => {
    const h = await harness(await sharedSoundDeck())
    const c = await session(h)
    try {
      const [r0] = await planMedia(c, 0), [r1] = await planMedia(c, 1)
      expect(r0!.mediaId).toBe(r1!.mediaId)
      expect(r0!.refused).toBe('media-unavailable')
      expect(r1!.refused).toBeUndefined()
      const id = r0!.mediaId
      expect(parseMediaShowOk(await c.mediaShow(1)).queued).toBe(1)
      // On the cursor's slide (1, no unit given): delivered.
      const ok = await c.mediaRequest(id, { ms: 60_000 })
      expect(ok.end).toBeUndefined()
      // Named slide 0: refused; named slide 1: delivered.
      expect(refusal(await c.mediaRequest(id, { unit: 0 }))).toBe('media-unavailable')
      expect((await c.mediaRequest(id, { unit: 1, ms: 60_000 })).end).toBeUndefined()
      // A slide that does not use it: invalid-media.
      expect(refusal(await c.mediaRequest(id, { unit: 5 }))).toBe('invalid-media')
      // The cursor moves to slide 0: refused there.
      c.mediaCursor(0); await Bun.sleep(100)
      expect(refusal(await c.mediaRequest(id))).toBe('media-unavailable')
      expect(h.stream.audit.mediaDeliveries(sidOf(c)).map((r) => [r.unitIndex, r.reason ?? 'ok'])).toEqual([[1, 'ok'], [0, 'media-unavailable'], [1, 'ok'], [0, 'media-unavailable']])
    } finally { c.close(); await h.stop() }
  }, 60_000)

  it('a revocation under backpressure answers media-refused media-unavailable at once and releases the pin', async () => {
    let held = true
    const t = heldTransport(() => held)
    const h = await harness(await noiseDeck(), {}, { transport: t.transport })
    const c = await session(h)
    try {
      const id = await idOf(c)
      await c.mediaShow(0)
      c.send({ reqId: 4100, type: 'media-request', token: c.manifest!.session.token, mediaId: id })
      await until(() => t.sent() === 1, 60_000)
      await Bun.sleep(200) // parked in the high-water wait
      expect(h.stream.stats().media.deliveries).toBe(1)
      const at = Date.now()
      await h.media!.cache.dropDocument(c.manifest!.docRef)
      const r = await c.mediaRequest(id, { reqId: 4100, send: false, ms: 2000 })
      expect(Date.now() - at).toBeLessThan(1000)
      expect(parseMediaRefused(r.end).reason).toBe('media-unavailable')
      expect(r.chunks).toHaveLength(1)
      await until(() => h.stream.stats().media.deliveries === 0, 1000)
      expect(t.sent()).toBe(1)
    } finally { held = false; c.close(); await h.stop() }
  }, 120_000)

  it('media-end under backpressure ends the delivery at once', async () => {
    const t = heldTransport(() => true)
    const h = await harness(await noiseDeck(), {}, { transport: t.transport })
    const c = await session(h)
    try {
      const id = await idOf(c)
      await c.mediaShow(0)
      c.send({ reqId: 4200, type: 'media-request', token: c.manifest!.session.token, mediaId: id })
      await until(() => t.sent() === 1, 60_000)
      await Bun.sleep(200)
      expect(h.stream.stats().media.deliveries).toBe(1)
      c.mediaEnd()
      await until(() => h.stream.stats().media.deliveries === 0, 1000)
      expect(h.media!.cache.referencesOf(sidOf(c))).toBe(0)
      expect(t.sent()).toBe(1)
    } finally { c.close(); await h.stop() }
  }, 120_000)

  it('a server-initiated close (revoke) ends the media queue at once, before the peer answers the close', async () => {
    const t = heldTransport(() => true)
    const h = await harness(await noiseDeck(), {}, { transport: t.transport })
    const c = await session(h)
    try {
      const id = await idOf(c)
      await c.mediaShow(0)
      c.send({ reqId: 4300, type: 'media-request', token: c.manifest!.session.token, mediaId: id })
      await until(() => t.sent() === 1, 60_000)
      expect(h.media!.cache.referencesOf(sidOf(c))).toBeGreaterThan(0)
      h.stream.revoke(sidOf(c))
      // Synchronously: the queue has ended (dropSession), before any close handshake.
      expect(h.media!.cache.referencesOf(sidOf(c))).toBe(0)
      await until(() => h.stream.stats().media.deliveries === 0, 1000)
    } finally { c.close(); await h.stop() }
  }, 120_000)

  it('the only session closes and the same deck reopens: a media-request is served again', async () => {
    const h = await harness(await narrationDeck())
    try {
      for (let round = 0; round < 2; round++) {
        const c = await session(h)
        try {
          const id = await idOf(c)
          await c.mediaShow(0)
          const r = await c.mediaRequest(id, { ms: 60_000 })
          expect(r.end).toBeUndefined()
          expect(parseMediaStart(r.start).mediaId).toBe(id)
        } finally { c.close(); await c.closed }
        // Closing the document (its last session) deletes its entries at once.
        await until(() => h.media!.cache.stats().entries === 0)
      }
    } finally { await h.stop() }
  }, 120_000)
})
