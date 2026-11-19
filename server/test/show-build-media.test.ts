/**
 * Piece 5 Task 4: show plans carry audio records, whitelisted media commands and transition sounds (spec §4.1; plan
 * D2, D3, D15, D18, D19, D21, D22, X13). The build (`DocumentHost.showPlan`), its binding (`bindMedia`) and the wire.
 */
import { afterAll, describe, expect, it } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { randomBytes } from 'node:crypto'
import { decodeShowMessage, encodeShowMessage, MEDIA_ID, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../src/stream/show-plan'
import { validateManifest } from '../../src/stream/manifest'
import type { BehaviourNode, TimeNode } from '../../src/slideshow/types'
import {
  crossSlideSoundDeck, effectSoundDeck, MEDIA_DECKS, mediaDeck, mediaEndRevealDeck, narrationDeck, playFromDeck, repeatDeck, togglePauseDeck, transitionSoundDeck, twoShapeAudioDeck, videoDeck,
} from '../../src/testdata/show-decks'
import { DocumentHost } from '../src/render'
import { Pool } from '../src/pool'
import { createMediaTemp } from '../src/media/temp'
import { bindMedia, mediaIdentity, sessionMediaId } from '../src/media/bind'
import { createDocMediaTable } from '../src/media/probe'
import { prepareShowSlide, type BuiltShowPlan, type ShowPlanJob, type ShowRefusal } from '../src/show/build'
import { Client, disposePool, start } from './harness'
import { ffmpegForTests } from './media/ffmpeg-gate'

const info = await ffmpegForTests()
const root = mkdtempSync(join(tmpdir(), 'ov-t4-'))
const temp = await createMediaTemp(root)
let mediaPool: Pool | undefined
const poolOn = () => (mediaPool ??= new Pool({ workers: 1, media: { info: info!, dir: temp.dir } }))
afterAll(async () => { mediaPool?.dispose(); disposePool(); await temp.remove(); rmSync(root, { recursive: true, force: true }) })

const job = (over: Partial<ShowPlanJob> = {}): ShowPlanJob => ({ unitIndex: 0, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false, media: true, ...over })
const built = (r: BuiltShowPlan | ShowRefusal): BuiltShowPlan => { if ('refused' in r) throw new Error(`refused: ${r.refused}`); return r }
function host() { const h = new DocumentHost(); h.configureMedia({ info: info!, dir: temp.dir }); return h }
async function plan(bytes: Uint8Array, o: Partial<ShowPlanJob> = {}, policy: 'visible-only' | 'full' = 'visible-only') {
  const h = host()
  const { docKey } = await h.open(`d${Math.random()}`, bytes, policy)
  return { h, docKey, r: await h.showPlan(docKey, job(o)) }
}
const key = () => new Uint8Array(randomBytes(32))
const bind = (p: ShowPlanRecord, k = key()) => bindMedia(p, (ref) => sessionMediaId(k, mediaIdentity(ref)))
/** The bound plan as a client decodes it. */
const viaClient = (p: ShowPlanRecord) => (decodeShowMessage(encodeShowMessage({ tag: 1, plan: p }), { unit: p.slide.index, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
function leaves(p: ShowPlanRecord): BehaviourNode[] {
  const out: BehaviourNode[] = []
  const walk = (n: TimeNode) => { if ('children' in n) n.children.forEach(walk); else out.push(n) }
  if (p.show.timing && !('refused' in p.show.timing)) walk(p.show.timing.root)
  return out
}
const cmds = (p: ShowPlanRecord) => leaves(p).flatMap(n => (n.kind === 'cmd' ? [n.cmd] : []))
const snds = (p: ShowPlanRecord) => leaves(p).flatMap(n => ('snd' in n.target ? [n.target.snd] : []))

// Raw timing for decks the shared fixtures do not cover.
const tgt = (spid: number) => `<p:spTgt spid="${spid}"/>`
const audioNode = (id: number, target: string, start = '<p:cond delay="0"/>') => `<p:audio><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst>${start}</p:stCondLst></p:cTn><p:tgtEl>${target}</p:tgtEl></p:cMediaNode></p:audio>`
const videoNode = (id: number, spid: number) => `<p:video><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl>${tgt(spid)}</p:tgtEl></p:cMediaNode></p:video>`
const cmd = (id: number, type: string, text: string, target = '<p:sldTgt/>', delay = 0) => `<p:cmd type="${type}" cmd="${text}"><p:cBhvr><p:cTn id="${id}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn><p:tgtEl>${target}</p:tgtEl></p:cBhvr></p:cmd>`
const box = { id: 2, name: 'Box', prst: 'rect', off: ['3619500', '2476500'] as [string, string], ext: ['1905000', '1905000'] as [string, string], fill: '336699' }

describe.skipIf(!info)('show plans carry audio media (Task 4)', () => {
  it('narrationDeck plan ships one audio record with placeholder, durationMs from the probe', async () => {
    const { h, docKey, r } = await plan(await narrationDeck())
    const p = built(r).plan
    const part = (await h.mediaTable(docKey).parts())[0]!
    const e = h.mediaTable(docKey).peek(part)
    expect(e && !('refused' in e) ? e.durationMs : -1).toBeGreaterThan(0)
    expect(p.media).toEqual([{ mediaId: '@a0.3', kind: 'audio', spid: '3', durationMs: (e as { durationMs: number }).durationMs }])
    // The cached plan holds the placeholder; only a bound plan frames.
    expect(() => encodeShowMessage({ tag: 1, plan: p })).toThrow()
    expect(viaClient(bind(p)).media![0]!.mediaId).toMatch(MEDIA_ID)
  })

  it('videoDeck plan ships no record, no placeholder and no units.list[i].media (D22)', async () => {
    const { r } = await plan(await videoDeck())
    // Piece 4: a video end that reveals a box is unprovable: the still, with media on as off.
    expect('refused' in r && r.refused).toBe('still')
    const unrelated = await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'av.mp4', kind: 'video' }], timing: videoNode(3, 3) }] })
    const p = built((await plan(unrelated)).r).plan
    expect(p.media).toEqual([])
    expect(JSON.stringify(p)).not.toContain('@a')
    const h = await start({ doc: await videoDeck(), pool: poolOn(), media: { enabled: true, info } }), c = new Client(h.url)
    try {
      const m = await c.hello()
      expect(m.can.media).toBe(true)
      expect(m.units.list[0]).not.toHaveProperty('media')
    } finally { c.close(); h.stop() }
  })

  it("media cmd text survives in canonical form; onstopaudio ships; any other cmd stays ''", async () => {
    const timing = audioNode(3, tgt(3)) + cmd(4, 'call', 'playFrom(1.50)', tgt(3)) + cmd(5, 'evt', 'OnStopAudio') + cmd(6, 'call', 'stop', tgt(3)) +
      cmd(7, 'call', 'play') + cmd(8, 'verb', '0') + cmd(9, 'call', 'rewind', tgt(3)) + cmd(10, 'evt', 'play', tgt(3))
    const p = built((await plan(await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'a.m4a' }], timing }] }))).r).plan
    expect(cmds(p)).toEqual(['playFrom(1.5)', 'onstopaudio', 'stop', '', '', '', ''])
    expect(p.diagnostics.map(d => d.code)).toContain('cmd-dropped')
    // A click-to-toggle on the audio shape survives streaming (piece 4 still classifies its trigger slide a still).
    const t = await plan(await togglePauseDeck())
    expect('refused' in t.r && t.r.refused).toBe('still')
    const show = (t.h as unknown as { docs: Map<string, { show: Parameters<typeof prepareShowSlide>[0] }> }).docs.get(t.docKey)!.show
    const streamed = prepareShowSlide(show, 0, 'full', { on: true })
    expect(cmds({ show: streamed.ps.meta } as unknown as ShowPlanRecord)).toEqual(['togglePause'])
    expect(cmds({ show: prepareShowSlide(show, 0, 'full', { on: false }).ps.meta } as unknown as ShowPlanRecord)).toEqual([''])
  })

  it("a togglePause or playFrom targeting a video shape stays ''", async () => {
    const timing = videoNode(3, 3) + cmd(4, 'call', 'togglePause', tgt(3)) + cmd(5, 'call', 'playFrom(1)', tgt(3)) + audioNode(6, tgt(4)) + cmd(7, 'call', 'pause', tgt(4))
    const p = built((await plan(await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'av.mp4', kind: 'video' }, { spid: 4, file: 'a.wav' }], timing }] }))).r).plan
    expect(cmds(p)).toEqual(['', '', 'pause'])
    expect(p.media!.map(r => r.spid)).toEqual(['4'])
  })

  it('playFrom over 1800 s is blanked', async () => {
    expect(cmds(built((await plan(await playFromDeck(1800))).r).plan)).toEqual(['playFrom(1800)'])
    expect(cmds(built((await plan(await playFromDeck(1800.5))).r).plan)).toEqual([''])
    expect(cmds(built((await plan(await playFromDeck('0.0000001'))).r).plan)).toEqual(['playFrom(0.0000001)'])
  })

  it('one audio part used as two effect sounds and the transition sound ships one record (X13)', async () => {
    const sound = (id: number) => audioNode(id, '<p:sndTgt r:embed="rIdSnd1" name="a.wav"/>')
    const timing = sound(3) + sound(4) + cmd(5, 'evt', 'onstopaudio')
    const transition = '<p:transition spd="med"><p:fade/><p:sndAc><p:stSnd loop="1"><p:snd r:embed="rIdSnd1" name="a.wav"/></p:stSnd></p:sndAc></p:transition>'
    const p = built((await plan(await mediaDeck({ slides: [{ boxes: [box], sounds: [{ rId: 'rIdSnd1', file: 'a.wav' }], timing, transition }] }))).r).plan
    expect(p.media).toHaveLength(1)
    expect(p.media![0]).toMatchObject({ mediaId: '@a0', kind: 'audio' })
    expect(p.media![0]).not.toHaveProperty('spid')
    expect(snds(p)).toEqual(['@a0', '@a0'])
    expect(p.show.transition?.sound).toEqual({ start: { mediaId: '@a0', loop: true } })
    const b = viaClient(bind(p))
    expect(new Set([...snds(b), b.show.transition!.sound!.start!.mediaId])).toEqual(new Set([b.media![0]!.mediaId]))
  })

  it('transition sound ships {start:{mediaId,loop}}; effect sound snd target is the mediaId after bind', async () => {
    const t = viaClient(bind(built((await plan(await transitionSoundDeck())).r).plan))
    expect(t.media).toHaveLength(1)
    expect(t.show.transition!.sound).toEqual({ start: { mediaId: t.media![0]!.mediaId, loop: false } })
    expect(t.media![0]!.mediaId).toMatch(MEDIA_ID)
    const e = viaClient(bind(built((await plan(await effectSoundDeck())).r).plan))
    expect(snds(e)).toEqual([e.media![0]!.mediaId])
    expect(e.media![0]!.durationMs).toBeGreaterThan(0)
  })

  it('no rId or part path in the bound plan JSON', async () => {
    const decks: Array<[string, Uint8Array]> = []
    for (const [name, f] of Object.entries(MEDIA_DECKS)) decks.push([name, await f()])
    decks.push(['builtIn', await mediaEndRevealDeck('effectSound')])
    let shipped = 0
    for (const [name, bytes] of decks) {
      const zip = await JSZip.loadAsync(bytes)
      const relIds = new Set<string>(), parts = new Set<string>()
      for (const f of Object.keys(zip.files)) {
        if (f.startsWith('ppt/media/')) { parts.add(f); parts.add(`/${f}`); parts.add(f.slice('ppt/'.length)) }
        if (f.endsWith('.rels')) for (const m of (await zip.file(f)!.async('string')).matchAll(/Id="([^"]+)"/g)) relIds.add(m[1]!)
      }
      const { h, docKey } = await plan(bytes)
      const units = (await h.open('again', bytes, 'visible-only')).units.length
      for (let u = 0; u < units; u++) {
        const r = await h.showPlan(docKey, job({ unitIndex: u }))
        if ('refused' in r) continue
        shipped++
        const json = JSON.stringify(viaClient(bind(r.plan)))
        for (const id of relIds) expect(json.includes(`"${id}"`), `${name}: ${id}`).toBe(false)
        for (const p of parts) expect(json.includes(p), `${name}: ${p}`).toBe(false)
        expect(json).not.toContain('builtIn')
        expect(json).not.toContain('/media/')
        expect(json).not.toContain('@a')
      }
    }
    // Every deck but videoDeck (a video end reveals its box: the still, D22) ships slide 1's plan.
    expect(shipped).toBeGreaterThanOrEqual(decks.length - 1)
  }, 120_000)

  it("hidden shape's media is not shipped", async () => {
    const timing = audioNode(3, tgt(3)) + audioNode(4, tgt(4)) + cmd(5, 'call', 'pause', tgt(3))
    const bytes = await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'a.m4a', hidden: true }, { spid: 4, file: 'a.wav' }], timing }] })
    const p = built((await plan(bytes, {}, 'full')).r).plan
    expect(p.media!.map(r => r.spid)).toEqual(['4'])
    // Its command targets no shipped audio: blanked.
    expect(cmds(p)).toEqual([''])
  })

  it('builtIn sound → refused media-unreadable', async () => {
    const p = built((await plan(await mediaEndRevealDeck('effectSound'))).r).plan
    expect(p.media).toHaveLength(1)
    expect(p.media![0]).toMatchObject({ kind: 'audio', durationMs: 0, refused: 'media-unreadable' })
    expect(snds(p)).toEqual([p.media![0]!.mediaId])
    expect(JSON.stringify(p)).not.toContain('builtIn')
  })

  it('can.media false → every record refused media-unavailable, and L1 key differs', async () => {
    const h = host()
    const { docKey } = await h.open('off', await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'a.m4a' }], sounds: [{ rId: 'rIdSnd1', file: 'a.wav' }], timing: audioNode(3, tgt(3)) + audioNode(4, '<p:sndTgt r:embed="rIdSnd1" name="a.wav"/>') + cmd(5, 'call', 'pause', tgt(3)), transition: '<p:transition><p:fade/><p:sndAc><p:stSnd><p:snd r:embed="rIdSnd1" name="a.wav"/></p:stSnd></p:sndAc></p:transition>' }] }), 'visible-only')
    const off = built(await h.showPlan(docKey, job({ media: false }))).plan
    expect(off.media).toEqual([
      { mediaId: '@a0.3', kind: 'audio', spid: '3', durationMs: 0, refused: 'media-unavailable' },
      { mediaId: '@a1', kind: 'audio', durationMs: 0, refused: 'media-unavailable' },
    ])
    // Off: the transition sound is stripped and every command blanked, as in piece 4 (D18).
    expect(off.show.transition?.sound).toBeUndefined()
    expect(cmds(off)).toEqual([''])
    const misses = () => h.showCacheStats().l1.misses
    const m0 = misses()
    const on = built(await h.showPlan(docKey, job({ media: true }))).plan
    expect(misses()).toBe(m0 + 1)
    expect(on.media!.every(r => r.refused === undefined && r.durationMs > 0)).toBe(true)
    // Absent is off: the same L1 entry as `media: false`.
    built(await h.showPlan(docKey, job({ media: undefined, mark: { kind: 'off' } })))
    expect(misses()).toBe(m0 + 1)
  })

  it('two sessions bind different mediaIds from one cached plan', async () => {
    const h = await start({ doc: await narrationDeck(), pool: poolOn(), media: { enabled: true, info } })
    const a = new Client(h.url), b = new Client(h.url)
    try {
      await a.hello(); await a.ready(); await b.hello(); await b.ready()
      const idOf = async (c: Client) => {
        const r = await c.showPlan(0, 1, 'full')
        expect(r.end).toBeUndefined()
        const p = (decodeShowMessage(r.messages[0]!, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
        expect(new TextEncoder().encode(JSON.stringify(p)).length).toBe(r.messages[0]!.length - 1)
        return p.media![0]!.mediaId
      }
      const ia = await idOf(a), ib = await idOf(b)
      expect(ia).toMatch(MEDIA_ID); expect(ib).toMatch(MEDIA_ID)
      expect(ia).not.toBe(ib)
      // Stable within a session.
      expect(await idOf(a)).toBe(ia)
    } finally { a.close(); b.close(); h.stop() }
  }, 60_000)

  it('one audio part on two audio shapes of one slide binds two distinct mediaIds', async () => {
    const p = built((await plan(await twoShapeAudioDeck())).r).plan
    expect(p.media!.map(r => [r.mediaId, r.spid])).toEqual([['@a0.3', '3'], ['@a0.4', '4']])
    const b = viaClient(bind(p))
    expect(new Set(b.media!.map(r => r.mediaId)).size).toBe(2)
  })

  it('an unanimated slide with an audio shape gets units.list[i].media and a show-plan is served for it (D19)', async () => {
    // Its only timing is the narration (no visual effect); slides 2 and 3 of the cross-slide deck have none at all.
    const h = await start({ doc: await crossSlideSoundDeck(), pool: poolOn(), media: { enabled: true, info } }), c = new Client(h.url)
    try {
      const m = await c.hello(); await c.ready()
      expect(validateManifest(JSON.parse(JSON.stringify(m))).ok).toBe(true)
      expect(m.caps.media).toEqual({ audio: { maxSeconds: 1800, maxBytes: 33554432 } })
      expect(m.units.list.map((u: { media?: true }) => u.media === true)).toEqual([true, false, false])
      const r = await c.showPlan(0, 1, 'full')
      expect(r.end).toBeUndefined()
      const p = (decodeShowMessage(r.messages[0]!, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
      expect(p.media).toHaveLength(1)
      expect(p.steps).toEqual([])
      expect((await c.showPlan(1, 1, 'full')).end?.code).toBe('invalid-unit')
    } finally { c.close(); h.stop() }
  }, 60_000)

  it('media.enabled false: records refused media-unavailable and piece-4 still classification unchanged (D18)', async () => {
    for (const media of [undefined, { enabled: false, info }, { enabled: true, info: undefined }]) {
      const h = await start({ doc: await narrationDeck(), pool: poolOn(), ...(media ? { media } : {}) }), c = new Client(h.url)
      try {
        const m = await c.hello(); await c.ready()
        expect(m.can.media).toBe(false)
        expect(m.caps.media ?? null).toBeNull()
        expect(m.caps).not.toHaveProperty('media')
        expect(m.units.list[0]).toEqual({ w: 960, h: 720, anim: true })
        const r = await c.showPlan(0, 1, 'full')
        expect(r.end).toBeUndefined()
        const p = (decodeShowMessage(r.messages[0]!, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
        expect(p.media).toEqual([{ mediaId: expect.stringMatching(MEDIA_ID), kind: 'audio', spid: '3', durationMs: 0, refused: 'media-unavailable' }])
      } finally { c.close(); h.stop() }
    }
    // Piece 4's stills stay stills: a video end reveals a box (D22), an audio repeats; with media on and off.
    for (const doc of [await mediaEndRevealDeck('onEnd'), await videoDeck(), await repeatDeck(2)]) {
      for (const media of [undefined, { enabled: true, info }]) {
        const h = await start({ doc, pool: poolOn(), ...(media ? { media } : {}) }), c = new Client(h.url)
        try {
          await c.hello(); await c.ready()
          const r = await c.showPlan(0, 1, 'full')
          expect(r.end).toEqual({ reqId: r.reqId, type: 'error', code: 'still' })
        } finally { c.close(); h.stop() }
      }
    }
  }, 120_000)

  it("a unit whose show meta isn't loaded fails loudly in load, digest and the build (Task 3 carried finding)", async () => {
    const h = host()
    const { docKey } = await h.open('loud', await narrationDeck(), 'visible-only')
    const doc = (h as unknown as { docs: Map<string, { show: Parameters<typeof createDocMediaTable>[0]['show'] }> }).docs.get(docKey)!
    const table = createDocMediaTable({ show: doc.show }, info, temp)
    // Slide 0's metadata is not in `show.metas` yet (open releases it).
    expect(doc.show!.metas.has(0)).toBe(false)
    await expect(table.load(0)).rejects.toThrow(/metadata not loaded/)
    expect(() => table.digest(0)).toThrow()
    // The host's own path loads it first, so a plan builds.
    expect('refused' in await h.showPlan(docKey, job())).toBe(false)
    expect(() => h.mediaTable(docKey).digest(0)).not.toThrow()
  })
})
