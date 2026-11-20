/**
 * Piece 5 Task 13 (spec §4.2–§4.4, §7; plan D3, D8, D22, D23; pre-flight X4, X29 R8; Review Focus 7, 8): what the
 * server puts on the wire, audited over every media deck variant and every fixture. Each deck is played through the
 * harness once (`beforeAll`: the plan of every slide, `media-show`, then a `media-request` for every playable use on
 * its slide), and every test below reads what was received.
 */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { parseMediaRefused, parseMediaStart } from '../../../src/stream/media-wire'
import { MEDIA_DECKS, mediaDeck, mixedMediaDeck, narrationDeck, stillSoundDeck } from '../../../src/testdata/show-decks'
import type { MediaDeliveryRecord } from '../../src/audit-store'
import { MUXER_TAG_ALLOWLIST, PROBE_REPORTED_OPTIONAL, encodeKeyId } from '../../src/media/params'
import { MEMORY_BYTES, runSandboxed } from '../../src/media/sandbox'
import { disposePool, start } from '../harness'
import { AuditClient, concat, jsonOf, packageMedia, planOf, session, sidOf, spawnLog, until } from './battery'
import { ffmpegForTests } from './ffmpeg-gate'

const info = await ffmpegForTests()
const scratch = mkdtempSync(join(tmpdir(), 'ov-t13-wire-'))
afterAll(() => { disposePool(); rmSync(scratch, { recursive: true, force: true }) })

const box = { id: 2, name: 'Box', prst: 'rect', off: ['3619500', '2476500'] as [string, string], ext: ['1905000', '1905000'] as [string, string], fill: '336699' }
const audioNode = (id: number, spid: number) => `<p:audio><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:audio>`
const FIXTURES = ['a.m4a', 'a.mp3', 'a.wav', 'a.wma', 'av.mp4'] as const
/** One slide per fixture, each an audio shape (an audio relationship, a `p:audio` node) begun with the slide: `av.mp4`
 * is consumed as audio (D22). */
const fixturesDeck = () => mediaDeck({ slides: FIXTURES.map((file) => ({ boxes: [box], media: [{ spid: 3, file, kind: 'audio' as const }], timing: audioNode(3, 3) })) })

const DECKS: Array<[string, () => Promise<Uint8Array>]> = [
  ...Object.entries(MEDIA_DECKS),
  ...(['video', 'none', 'chain', 'chain-build'] as const).map((g): [string, () => Promise<Uint8Array>] => [`mixedMediaDeck(${g})`, () => mixedMediaDeck(g)]),
  ...(['audio-end', 'transition', 'shared'] as const).map((k): [string, () => Promise<Uint8Array>] => [`stillSoundDeck(${k})`, () => stillSoundDeck(k)]),
  ['fixturesDeck', fixturesDeck],
]

interface Delivery { mediaId: string; unit: number; bytes: Uint8Array; durationMs: number }
interface Played {
  name: string
  pkg: Awaited<ReturnType<typeof packageMedia>>
  texts: string[]
  plans: string[]
  deliveries: Delivery[]
  rows: MediaDeliveryRecord[]
  /** The cache's key ids for this deck's entries, and the server-only key fields its rows hold. */
  secrets: string[]
}

/** stdout of a sandboxed ffprobe over `path`, after BOTH exit and the stream's end. */
async function ffprobe(path: string, entries: string[]): Promise<any> {
  const run = runSandboxed(info!, 'ffprobe', ['-v', 'error', '-hide_banner', ...entries, '-of', 'json', `file:${path}`], { cpuSeconds: 30, fileBytes: 1 << 20, memoryBytes: MEMORY_BYTES[info!.platform], wallMs: 30_000, stderrBytes: 65536 }, new AbortController().signal)
  const parts: Buffer[] = []
  run.stdout.on('data', (b: Buffer) => parts.push(b))
  const drained = new Promise<void>((ok) => { run.stdout.once('end', ok); run.stdout.once('close', ok); run.stdout.once('error', () => ok()) })
  const [res] = await Promise.all([run.done, drained])
  if (res.code !== 0) throw new Error(`ffprobe failed: ${res.stderr}`)
  return JSON.parse(Buffer.concat(parts).toString('utf8'))
}

/** Plays `deck`: every slide's plan, `media-show`, a `media-request` per playable use (the media class allows six a
 * session per minute: no deck here has more). */
async function playDeck(name: string, deck: Uint8Array): Promise<Played> {
  const h = await start({ doc: deck, media: {} })
  const c = await session(h)
  try {
    const plans: string[] = []
    const uses: Array<{ mediaId: string; unit: number }> = []
    for (let u = 0; u < c.manifest!.units.count; u++) {
      const p = await planOf(c, u, { still: true })
      plans.push(...p.decoded.map(jsonOf))
      for (const r of p.media) if (!r.refused && !uses.some((x) => x.mediaId === r.mediaId && x.unit === u)) uses.push({ mediaId: r.mediaId, unit: u })
    }
    if (uses.length > 6) throw new Error(`${name}: ${uses.length} playable uses, over one session's media rate`)
    await c.mediaShow(0)
    const deliveries: Delivery[] = []
    for (const u of uses) {
      const r = await c.mediaRequest(u.mediaId, { unit: u.unit, ms: 60_000 })
      if (r.end) throw new Error(`${name}: ${u.mediaId} on slide ${u.unit} refused ${parseMediaRefused(r.end).reason}`)
      const st = parseMediaStart(r.start)
      expect(r.chunks).toHaveLength(st.chunks)
      deliveries.push({ mediaId: u.mediaId, unit: u.unit, bytes: concat(r.chunks), durationMs: st.durationMs })
    }
    const rows = h.stream.audit.mediaDeliveries(sidOf(c))
    const secrets = [...h.media!.cache.inspect().entries.map((e) => e.id), ...rows.flatMap((r) => [r.sourceHash, r.encodeParamsSha].filter((x): x is string => !!x))]
    await Bun.sleep(50)
    return { name, pkg: await packageMedia(deck), texts: [...c.texts], plans, deliveries, rows, secrets }
  } finally { c.close(); await h.stop() }
}

/** udta { meta { hdlr 'mdir' 'appl', ilst {} } }: 61 bytes the mp4 muxer writes in every output (Task 7 ruling: a
 * muxer constant, not source metadata; transcode.test.ts pins it). */
const EMPTY_UDTA = Buffer.from('0000003d75647461000000356d657461000000000000002168646c7200000000000000006d6469726170706c00000000000000000000000008696c7374', 'hex')

describe.skipIf(!info)('the media wire audit (Task 13)', () => {
  const played: Played[] = []
  beforeAll(async () => {
    for (const [name, make] of DECKS) played.push(await playDeck(name, await make()))
  }, 600_000)

  it('every deck variant and fixture is played: each fixture is delivered as an audio use', () => {
    expect(played.map((p) => p.name)).toEqual(DECKS.map(([n]) => n))
    // A deck whose uses were all refused would pass every "nothing on the wire" check below: count the deliveries.
    const fx = played.find((p) => p.name === 'fixturesDeck')!
    expect(fx.deliveries.map((d) => d.unit)).toEqual([0, 1, 2, 3, 4])
    expect(played.reduce((n, p) => n + p.deliveries.length, 0)).toBeGreaterThanOrEqual(20)
    // Only these have no playable use: a video (D22: no use) and a narration chained on a video's end (refused).
    const silent = ['videoDeck', 'mixedMediaDeck(chain)', 'mixedMediaDeck(chain-build)']
    for (const p of played) expect(p.deliveries.length > 0, p.name).toBe(!silent.includes(p.name))
  })

  /** The first 64-byte window of a source part (one with at least 32 distinct byte values) found in `out`. */
  function leak(out: Buffer, parts: Map<string, Uint8Array>): { path: string; at: number } | undefined {
    const hit = leakAny([out], parts)
    return hit && { path: hit.path, at: hit.at }
  }

  /** `leak` over several outputs at once: every output window is indexed once, and each source part is walked once
   *  (not once per output), so a deck's audit costs one pass over its parts. `out` is the output that holds the hit. */
  function leakAny(outs: readonly Buffer[], parts: Map<string, Uint8Array>): { path: string; at: number; out: number } | undefined {
    // Task 7 ruling: the empty moov/udta box (61 bytes) is a muxer constant, allowed in every output. The allowance
    // exempts only windows lying wholly inside it, and since a window (64 bytes) is longer than the box there are
    // none (fix round 1): every output window is indexed, so source bytes against either edge are still caught.
    const windows = new Map<string, number>()
    outs.forEach((out, k) => {
      for (let i = 0; i + 64 <= out.length; i++) {
        const w = out.toString('latin1', i, i + 64)
        if (!windows.has(w)) windows.set(w, k)
      }
    })
    const seen = new Uint8Array(256)
    for (const [path, bytes] of parts) {
      const src = Buffer.from(bytes)
      for (let i = 0; i + 64 <= src.length; i++) {
        const k = windows.get(src.toString('latin1', i, i + 64))
        if (k === undefined) continue
        // X29 R8: zero padding and silence legitimately recur in a re-encode. Checked only on a match (it decides
        // nothing otherwise): a window with fewer than 32 distinct byte values is not a leak.
        seen.fill(0)
        let distinct = 0
        for (let j = i; j < i + 64; j++) if (!seen[src[j]!]++) distinct++
        if (distinct >= 32) return { path, at: i, out: k }
      }
    }
    return undefined
  }

  it('the window audit catches source bytes against either edge of the muxer\'s empty udta (positive controls)', () => {
    // Fix round 1: the 61-byte box is shorter than a window, so every window over it holds bytes from outside it.
    // An output with the constant between two runs of high-entropy bytes; a source holding 64 bytes that cross
    // its left edge (63 before + the first box byte), or its right edge (the last box byte + 63 after), is a leak.
    let x = 0x2545f491
    const noise = (n: number) => Buffer.from(Array.from({ length: n }, () => { x ^= x << 13; x ^= x >>> 17; x ^= x << 5; return (x >>> 0) & 0xff }))
    const out = Buffer.concat([noise(200), EMPTY_UDTA, noise(200)])
    const at = out.indexOf(EMPTY_UDTA)
    const left = out.subarray(at - 63, at + 1), right = out.subarray(at + EMPTY_UDTA.length - 1, at + EMPTY_UDTA.length + 63)
    for (const [edge, w] of [['left', left], ['right', right]] as const) {
      expect(new Set(w).size, edge).toBeGreaterThanOrEqual(32)
      expect(leak(out, new Map([[`ppt/media/${edge}.m4a`, new Uint8Array(w)]])), edge).toEqual({ path: `ppt/media/${edge}.m4a`, at: 0 })
    }
  })

  it('no MEDIA_CHUNK contains any 64-byte window of any source part', () => {
    let outputs = 0
    for (const p of played) {
      const outs = p.deliveries.map((d) => Buffer.from(d.bytes))
      for (const out of outs) {
        // Every output holds the muxer's empty udta exactly once.
        const at = out.indexOf(EMPTY_UDTA)
        expect(at, `${p.name}: the muxer's empty udta`).toBeGreaterThan(0)
        expect(out.indexOf(EMPTY_UDTA, at + 1)).toBe(-1)
        outputs++
      }
      const hit = leakAny(outs, p.pkg.parts)
      if (hit) throw new Error(`${p.name}: ${hit.path} window at ${hit.at} appears in ${p.deliveries[hit.out]!.mediaId}'s output`)
    }
    expect(outputs).toBeGreaterThanOrEqual(20)
    // Positive control: a part delivered verbatim is caught.
    const fx = played.find((p) => p.name === 'fixturesDeck')!
    const [path, bytes] = [...fx.pkg.parts].find(([k]) => k.endsWith('.wav'))!
    expect(leak(Buffer.from(bytes), fx.pkg.parts)?.path).toBe(path)
    // ...and, among a deck's outputs indexed together, it is attributed to the output that holds it.
    const clean = Buffer.from(fx.deliveries[0]!.bytes)
    expect(leakAny([clean, Buffer.from(bytes)], fx.pkg.parts)).toMatchObject({ path, out: 1 })
  })

  it('reassembled outputs show no metadata under ffprobe -show_format -show_streams', async () => {
    // X4: exactly the muxer's tags at their pinned values; a probe-reported optional tag (ffprobe 6.1.1's vendor_id)
    // only at its pinned value; nothing else (transcode.test.ts `check`).
    const check = (tags: Record<string, string> | undefined, allow: Record<string, readonly string[]>, optional: Record<string, readonly string[]> = {}) => {
      const seen = Object.keys(tags ?? {}).filter((k) => !(k in optional)).sort()
      expect(seen).toEqual(Object.keys(allow).sort())
      for (const [k, v] of Object.entries(tags ?? {})) expect(allow[k] ?? optional[k]).toContain(v)
    }
    let n = 0
    for (const p of played) for (const d of p.deliveries) {
      const path = join(scratch, `out-${n++}.m4a`)
      writeFileSync(path, d.bytes)
      const j = await ffprobe(path, ['-show_format', '-show_streams', '-show_chapters'])
      expect(j.chapters ?? []).toEqual([])
      expect(j.streams).toHaveLength(1)
      expect(j.format.nb_streams).toBe(1)
      check(j.format.tags, MUXER_TAG_ALLOWLIST.format)
      check(j.streams[0].tags, MUXER_TAG_ALLOWLIST.stream, PROBE_REPORTED_OPTIONAL.stream)
      for (const tags of [j.format.tags, j.streams[0].tags]) for (const k of ['encoder', 'creation_time']) expect(tags ?? {}).not.toHaveProperty(k)
      const raw = Buffer.from(d.bytes)
      for (const s of ['Lavf', 'Lavc', 'encoder', 'creation_time']) expect(raw.includes(s), `${p.name}: ${s}`).toBe(false)
    }
    expect(n).toBeGreaterThanOrEqual(20)
  }, 120_000)

  it('no shipped plan record or text message contains an rId or part path from the package', () => {
    let rIds = 0
    for (const p of played) {
      const banned = new Set<string>()
      for (const id of p.pkg.rIds) { banned.add(id); rIds++ }
      for (const t of p.pkg.targets) { banned.add(t); banned.add(t.split('/').pop()!) }
      for (const path of p.pkg.parts.keys()) { banned.add(path); banned.add(`/${path}`); banned.add(path.split('/').pop()!) }
      expect(p.texts.length).toBeGreaterThan(0)
      for (const s of [...p.texts, ...p.plans]) for (const b of banned) if (s.includes(b)) throw new Error(`${p.name}: ${JSON.stringify(b)} in ${s.slice(0, 200)}`)
    }
    expect(rIds).toBeGreaterThan(20)
    expect(played.reduce((n, p) => n + p.plans.length, 0)).toBeGreaterThan(20)
  })

  it('no reassembled output has a video stream', async () => {
    // av.mp4 used as audio (D22, Review Focus 7) included: the fixtures deck plays it on slide 4.
    const fx = played.find((p) => p.name === 'fixturesDeck')!
    expect(fx.deliveries.some((d) => d.unit === FIXTURES.indexOf('av.mp4'))).toBe(true)
    let n = 0
    for (const p of played) for (const d of p.deliveries) {
      const path = join(scratch, `video-${n++}.m4a`)
      writeFileSync(path, d.bytes)
      const j = await ffprobe(path, ['-show_streams'])
      expect(j.streams.map((s: any) => [s.codec_type, s.codec_name]), p.name).toEqual([['audio', 'aac']])
      const raw = Buffer.from(d.bytes)
      for (const s of ['vide', 'avc1', 'avcC', 'x264']) expect(raw.includes(s), `${p.name}: ${s}`).toBe(false)
    }
  }, 120_000)

  it('every delivery has a media_deliveries row with a cache_state', () => {
    for (const p of played) {
      const ok = p.rows.filter((r) => r.reason === undefined)
      expect(ok.map((r) => [r.mediaId, r.unitIndex]), p.name).toEqual(p.deliveries.map((d) => [d.mediaId, d.unit]))
      for (const r of ok) {
        expect(['miss', 'joined', 'hit'], p.name).toContain(r.cacheState!)
        expect(r.outBytes).toBe(p.deliveries.find((d) => d.mediaId === r.mediaId && d.unit === r.unitIndex)!.bytes.length)
      }
      expect(p.rows.length).toBe(ok.length) // no refusal: every request was for a playable use
    }
  })

  // --- D23: the shared cache across sessions and documents ---------------------------------------------------------

  async function deliver(c: AuditClient) {
    const { media } = await planOf(c)
    await c.mediaShow(0)
    const r = await c.mediaRequest(media[0]!.mediaId, { ms: 60_000 })
    if (r.end) throw new Error(`refused ${parseMediaRefused(r.end).reason}`)
    return { mediaId: media[0]!.mediaId, bytes: concat(r.chunks) }
  }

  it('two sessions of one document: one ffmpeg spawn, identical bytes, different mediaIds', async () => {
    const h = await start({ doc: await narrationDeck(), media: {} })
    const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
    const a = await session(h), b = await session(h)
    try {
      const da = await deliver(a)
      const db = await deliver(b)
      expect(log.ffmpeg()).toHaveLength(1)
      expect(db.bytes).toEqual(da.bytes)
      expect(da.mediaId).not.toBe(db.mediaId)
      expect(h.media!.cache.stats().entries).toBe(1)
      expect(h.stream.audit.mediaDeliveries(sidOf(a)).map((r) => r.cacheState)).toEqual(['miss'])
      expect(['joined', 'hit']).toContain(h.stream.audit.mediaDeliveries(sidOf(b))[0]!.cacheState!)
    } finally { log.stop(); a.close(); b.close(); await h.stop() }
  }, 60_000)

  it("two documents with byte-identical audio: two ffmpeg spawns, and neither session's request is answered from the other's entry", async () => {
    // docId is an HMAC of the document's bytes (Task 7): the decks differ outside the audio (the box's fill), so they
    // are two documents whose one audio part is the same bytes.
    const deckA = await narrationDeck()
    const zip = await JSZip.loadAsync(deckA)
    const slide = 'ppt/slides/slide1.xml'
    const xml = await zip.file(slide)!.async('string')
    expect(xml).toContain('336699')
    zip.file(slide, xml.replace('336699', '993366'))
    const deckB = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    const [pa, pb] = [await packageMedia(deckA), await packageMedia(deckB)]
    expect([...pb.parts.values()]).toEqual([...pa.parts.values()])
    const h = await start({ media: {}, resolveDocument: async (req) => ({ bytes: req.url.searchParams.get('doc') === 'b' ? deckB : deckA, name: 'doc' }) })
    const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
    const a = new AuditClient(`${h.url}?doc=a`), b = new AuditClient(`${h.url}?doc=b`)
    try {
      await a.hello(); await a.ready(); await b.hello(); await b.ready()
      expect(a.manifest!.docRef).not.toBe(b.manifest!.docRef)
      const da = await deliver(a)
      const db = await deliver(b)
      expect(log.ffmpeg()).toHaveLength(2)
      const [ra] = h.stream.audit.mediaDeliveries(sidOf(a)), [rb] = h.stream.audit.mediaDeliveries(sidOf(b))
      expect(ra!.sourceHash).toBe(rb!.sourceHash)
      expect([ra!.cacheState, rb!.cacheState]).toEqual(['miss', 'miss'])
      // Two entries, one per document; each session's key names its own document.
      const entries = h.media!.cache.inspect().entries
      expect(entries.map((e) => e.docId).sort()).toEqual([a.manifest!.docRef, b.manifest!.docRef].sort())
      for (const [c, r] of [[a, ra!], [b, rb!]] as const) expect(entries.find((e) => e.docId === c.manifest!.docRef)!.id).toBe(encodeKeyId({ docId: c.manifest!.docRef, sourceHash: r.sourceHash!, paramsSha: r.encodeParamsSha!, ffmpegVersion: r.ffmpegVersion! }))
      expect(new Set(entries.map((e) => e.path)).size).toBe(2)
      // The same encode settings over the same bytes: the same output, from two encodes.
      expect(db.bytes).toEqual(da.bytes)
    } finally { log.stop(); a.close(); b.close(); await h.stop() }
  }, 60_000)

  it("a session cannot fetch another session's media: a foreign mediaId → invalid-media, and no cache key or key id appears in any message or plan", async () => {
    const h = await start({ doc: await narrationDeck(), media: {} })
    const a = await session(h), b = await session(h)
    try {
      const plansA = (await planOf(a)).decoded.map(jsonOf), pb = await planOf(b)
      const plansB = pb.decoded.map(jsonOf)
      await a.mediaShow(0); await b.mediaShow(0)
      const ia = (await planOf(a)).media[0]!.mediaId
      expect((await a.mediaRequest(ia, { ms: 60_000 })).end).toBeUndefined()
      const [row] = h.stream.audit.mediaDeliveries(sidOf(a))
      const key = h.media!.cache.inspect().entries[0]!.id
      expect(key).toBe(encodeKeyId({ docId: a.manifest!.docRef, sourceHash: row!.sourceHash!, paramsSha: row!.encodeParamsSha!, ffmpegVersion: row!.ffmpegVersion! }))
      const asked = new Set<number>()
      for (const foreign of [ia, key, key.slice(0, 22)]) {
        const r = await b.mediaRequest(foreign)
        expect(r.start).toBeUndefined()
        expect(parseMediaRefused(r.end).reason).toBe('invalid-media')
        asked.add(r.reqId)
      }
      expect(h.stream.audit.mediaDeliveries(sidOf(b))).toEqual([])
      expect(pb.media[0]!.mediaId).not.toBe(ia)
      await until(() => a.texts.length > 0 && b.texts.length > 0)
      // The key and the server-only fields it is made of (the docId is the manifest's docRef, on the wire by design).
      const secrets = [key, key.slice(0, 22), row!.sourceHash!, row!.encodeParamsSha!]
      // B's refusals echo the ids B itself sent (a well-formed one is echoed): those are not the server's to leak.
      const echoes = (t: string) => { const m = JSON.parse(t); return m.type === 'media-refused' && asked.has(m.reqId) }
      for (const s of [...a.texts, ...b.texts.filter((t) => !echoes(t)), ...plansA, ...plansB]) for (const x of secrets) if (s.includes(x)) throw new Error(`${x} in ${s.slice(0, 300)}`)
    } finally { a.close(); b.close(); await h.stop() }
  }, 60_000)

  it('the played decks carried no cache key or key field either', () => {
    for (const p of played) for (const s of [...p.texts, ...p.plans]) for (const x of p.secrets) if (s.includes(x)) throw new Error(`${p.name}: a key field in ${s.slice(0, 200)}`)
  })
})
