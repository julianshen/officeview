/**
 * Piece 5 Task 4: show plans carry audio records, whitelisted media commands and transition sounds (spec §4.1; plan
 * D2, D3, D15, D18, D19, D21, D22, X13). The build (`DocumentHost.showPlan`), its binding (`bindMedia`) and the wire.
 */
import { afterAll, describe, expect, it } from 'bun:test'
import { rejectionMessage } from './fixtures/rejection'
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
/** Each built plan's server-only media refs, which `bindMedia` needs (`BuiltShowPlan.mediaRefs`). */
const refsOf = new WeakMap<ShowPlanRecord, BuiltShowPlan['mediaRefs']>()
const built = (r: BuiltShowPlan | ShowRefusal): BuiltShowPlan => { if ('refused' in r) throw new Error(`refused: ${r.refused}`); refsOf.set(r.plan, r.mediaRefs); return r }
function host() { const h = new DocumentHost(); h.configureMedia({ info: info!, dir: temp.dir }); return h }
async function plan(bytes: Uint8Array, o: Partial<ShowPlanJob> = {}, policy: 'visible-only' | 'full' = 'visible-only') {
  const h = host()
  const { docKey } = await h.open(`d${Math.random()}`, bytes, policy)
  return { h, docKey, r: await h.showPlan(docKey, job(o)) }
}
const key = () => new Uint8Array(randomBytes(32))
const bind = (p: ShowPlanRecord, k = key(), refs = refsOf.get(p)!) => bindMedia(p, (ref) => sessionMediaId(k, mediaIdentity(ref)), refs)
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
    const shipped: string[] = [], streamedOnly: string[] = []
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
        const clean = (json: string) => {
          for (const id of relIds) expect(json.includes(`"${id}"`), `${name}: ${id}`).toBe(false)
          for (const p of parts) expect(json.includes(p), `${name}: ${p}`).toBe(false)
          expect(json).not.toContain('builtIn')
          expect(json).not.toContain('/media/')
        }
        if ('refused' in r) {
          // A slide refused to its still ships no plan; what it would stream (its prepared meta and records) is clean too.
          const show = (h as unknown as { docs: Map<string, { show: Parameters<typeof prepareShowSlide>[0] }> }).docs.get(docKey)!.show
          const st = prepareShowSlide(show, u, 'full', { on: true })
          clean(JSON.stringify({ timing: st.ps.meta.timing, transition: st.ps.meta.transition, records: st.records }))
          streamedOnly.push(`${name}#${u}`)
          continue
        }
        shipped.push(`${name}#${u}`)
        const json = JSON.stringify(viaClient(bind(r.plan, key(), r.mediaRefs)))
        clean(json)
        expect(json).not.toContain('@a')
      }
    }
    // Piece 4 stills (Task 5 owns them): a video end that reveals a box (D22), a trigger cut, a repeating audio.
    expect(streamedOnly.sort()).toEqual(['repeatDeck#0', 'togglePauseDeck#0', 'videoDeck#0'])
    expect(shipped.sort()).toEqual(['builtIn#0', 'crossSlideSoundDeck#0', 'crossSlideSoundDeck#1', 'crossSlideSoundDeck#2', 'effectSoundDeck#0', 'narrationDeck#0', 'playFromDeck#0', 'transitionSoundDeck#0', 'twoShapeAudioDeck#0'])
  }, 120_000)

  it("hidden shape's media is not shipped", async () => {
    const timing = audioNode(3, tgt(3)) + audioNode(4, tgt(4)) + cmd(5, 'call', 'pause', tgt(3))
    const bytes = await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'a.m4a', hidden: true }, { spid: 4, file: 'a.wav' }], timing }] })
    const p = built((await plan(bytes, {}, 'full')).r).plan
    expect(p.media!.map(r => r.spid)).toEqual(['4'])
    // Its command targets no shipped audio: blanked.
    expect(cmds(p)).toEqual([''])
  })

  it('an external media link ships refused media-linked (shape audio and effect sound)', async () => {
    const timing = audioNode(3, tgt(3)) + audioNode(4, '<p:sndTgt r:embed="rIdSnd1" name="a.wav"/>')
    const zip = await JSZip.loadAsync(await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'a.m4a' }], sounds: [{ rId: 'rIdSnd1', file: 'a.wav' }], timing }] }))
    const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
    const rels = (await zip.file(relsPath)!.async('string')).replace(/Target="\.\.\/media\/[^"]+"/g, 'Target="https://example.com/x.m4a" TargetMode="External"')
    zip.file(relsPath, rels)
    const p = built((await plan(new Uint8Array(await zip.generateAsync({ type: 'uint8array' })))).r).plan
    expect(p.media!.map(r => [r.spid, r.refused, r.durationMs])).toEqual([['3', 'media-linked', 0], [undefined, 'media-linked', 0]])
    const json = JSON.stringify(viaClient(bind(p)))
    expect(json).not.toContain('example.com')
    expect(json).not.toContain('rIdSnd1')
  })

  // Fix round 1 (review Important): one definition of a unit's audio uses, and part-less identities scoped to their slide.
  it('A: part-less sounds on two slides bind different mediaIds in one session', async () => {
    const builtIn = (id: number, name: string) => audioNode(id, `<p:sndTgt name="${name}"/>`)
    const bytes = await mediaDeck({ slides: [{ boxes: [box], timing: builtIn(3, 'chime.wav') }, { boxes: [box], timing: builtIn(3, 'wind.wav') }] })
    const h = host()
    const { docKey } = await h.open('A', bytes, 'visible-only')
    const k = key()
    const ids = [] as string[]
    for (const u of [0, 1]) {
      const b = built(await h.showPlan(docKey, job({ unitIndex: u })))
      expect(b.plan.media).toHaveLength(1)
      ids.push(viaClient(bindMedia(b.plan, (ref) => sessionMediaId(k, mediaIdentity(ref)), b.mediaRefs)).media![0]!.mediaId)
    }
    expect(ids[0]).not.toBe(ids[1])
  })

  it('B: mediaRefs and the bound plan agree on k, identity and refusal for an external link', async () => {
    const timing = audioNode(3, tgt(3)) + audioNode(4, '<p:sndTgt r:embed="rIdSnd1" name="a.wav"/>') + audioNode(5, '<p:sndTgt name="chime.wav"/>')
    const zip = await JSZip.loadAsync(await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'a.m4a' }], sounds: [{ rId: 'rIdSnd1', file: 'a.wav' }], timing }] }))
    const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
    zip.file(relsPath, (await zip.file(relsPath)!.async('string')).replace(/(Id="rIdSnd1"[^>]*?)Target="\.\.\/media\/[^"]+"/, '$1Target="https://example.com/x.wav" TargetMode="External"'))
    const h = host()
    const { docKey } = await h.open('B', new Uint8Array(await zip.generateAsync({ type: 'uint8array' })), 'visible-only')
    const b = built(await h.showPlan(docKey, job()))
    const refs = await h.mediaRefs(docKey, 0)
    const fromPlan = b.plan.media!.map((r) => { const m = b.mediaRefs.find((x) => x.placeholder === r.mediaId)!; return [mediaIdentity(m.ref), m.ref.k, r.refused ?? 'ok'] })
    const fromRefs = refs.refs.map((r) => [r.identity, r.k, 'refused' in r.entry ? r.entry.refused : 'ok'])
    expect(fromRefs).toEqual(fromPlan)
    expect(fromPlan.map((x) => x[2])).toEqual(['ok', 'media-linked', 'media-unreadable'])
    expect(refs.refs.every((r) => r.k >= 0)).toBe(true)
  })

  it('C: a condition-only snd on an external link is refused media-linked', async () => {
    const cond = '<p:cond evt="onBegin" delay="0"><p:tgtEl><p:sndTgt r:embed="rIdSnd1" name="a.wav"/></p:tgtEl></p:cond>'
    const timing = `<p:set><p:cBhvr><p:cTn id="3" dur="1" fill="hold"><p:stCondLst>${cond}</p:stCondLst></p:cTn><p:tgtEl>${tgt(2)}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
    const zip = await JSZip.loadAsync(await mediaDeck({ slides: [{ boxes: [box], sounds: [{ rId: 'rIdSnd1', file: 'a.wav' }], timing }] }))
    const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
    zip.file(relsPath, (await zip.file(relsPath)!.async('string')).replace(/(Id="rIdSnd1"[^>]*?)Target="\.\.\/media\/[^"]+"/, '$1Target="https://example.com/x.wav" TargetMode="External"'))
    const h = host()
    const { docKey } = await h.open('C', new Uint8Array(await zip.generateAsync({ type: 'uint8array' })), 'visible-only')
    const r = await h.showPlan(docKey, job())
    const p = 'refused' in r ? undefined : r.plan
    const records = p?.media ?? prepareShowSlide((h as unknown as { docs: Map<string, { show: Parameters<typeof prepareShowSlide>[0] }> }).docs.get(docKey)!.show, 0, 'full', { on: true }).records
    expect(records!.map((x) => x.refused)).toEqual(['media-linked'])
    expect((await h.mediaRefs(docKey, 0)).refs.map((x) => 'refused' in x.entry && x.entry.refused)).toEqual(['media-linked'])
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
        // `decodedBytes` (measured on the unbound plan with 22-character stand-ins) is exact for the bound plan.
        const assets = r.messages.slice(1).map((m) => decodeShowMessage(m, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded }) as { w: number; h: number })
        expect(r.start!.decodedBytes).toBe(assets.reduce((n, x) => n + x.w * x.h * 4, 0) + r.messages[0]!.length - 1)
        return p.media![0]!.mediaId
      }
      const ia = await idOf(a), ib = await idOf(b)
      expect(ia).toMatch(MEDIA_ID); expect(ib).toMatch(MEDIA_ID)
      expect(ia).not.toBe(ib)
      // Stable within a session.
      expect(await idOf(a)).toBe(ia)
    } finally { a.close(); b.close(); h.stop() }
  }, 60_000)

  it('a resume keeps the session media key: the same mediaId before and after', async () => {
    const h = await start({ doc: await narrationDeck(), pool: poolOn(), media: { enabled: true, info } })
    const idOf = async (c: Client) => {
      const r = await c.showPlan(0, 1, 'full')
      expect(r.end).toBeUndefined()
      return (decodeShowMessage(r.messages[0]!, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan.media![0]!.mediaId
    }
    const c0 = new Client(h.url)
    try {
      const m0 = await c0.hello(); await c0.ready()
      const before = await idOf(c0)
      const ws = c0.ws as WebSocket & { terminate?: () => void }
      if (ws.terminate) ws.terminate(); else ws.close(1001)
      await c0.closed; await Bun.sleep(50)
      const c1 = new Client(h.url)
      try {
        const m1 = await c1.hello({ resume: { sid: m0.session.sid, resumeProof: m0.session.token } })
        expect(m1.session.sid).toBe(m0.session.sid)
        await c1.ready()
        expect(await idOf(c1)).toBe(before)
      } finally { c1.close() }
    } finally { h.stop() }
  }, 60_000)

  it('D19: a unit the manifest marks media but not anim is served a show-plan only with can.media', async () => {
    // No deck yields media without anim (an audio node is timing), so the pool's open result is edited: unit 0 keeps
    // `media` and loses `anim`. Without `|| (canMedia && d.media.has(u))` the request is invalid-unit.
    const real = poolOn()
    const pool = new Proxy(real, { get(t, k) {
      if (k === 'open') return async (...a: Parameters<Pool['open']>) => ({ ...(await t.open(...a)), anim: [] })
      const v = Reflect.get(t, k, t); return typeof v === 'function' ? v.bind(t) : v
    } }) as Pool
    for (const enabled of [true, false]) {
      const h = await start({ doc: await narrationDeck(), pool, media: { enabled, info } }), c = new Client(h.url)
      try {
        const m = await c.hello(); await c.ready()
        expect(m.units.list[0]!.anim).toBeUndefined()
        expect(m.units.list[0]!.media).toBe(enabled ? true : undefined)
        const r = await c.showPlan(0, 1, 'full')
        if (enabled) expect(r.end).toBeUndefined()
        else expect(r.end?.code).toBe('invalid-unit')
      } finally { c.close(); h.stop() }
    }
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
    // Task 11b: with media on, the audio still (repeatDeck) answers its sound-only plan (`fallback: 'still'`, nothing
    // drawn) in place of the error; it still plays its still.
    for (const [doc, sounds] of [[await mediaEndRevealDeck('onEnd'), false], [await videoDeck(), false], [await repeatDeck(2), true]] as const) {
      for (const media of [undefined, { enabled: true, info }]) {
        const h = await start({ doc, pool: poolOn(), ...(media ? { media } : {}) }), c = new Client(h.url)
        try {
          await c.hello(); await c.ready()
          const r = await c.showPlan(0, 1, 'full')
          if (sounds && media) {
            expect(r.end).toBeUndefined()
            expect((decodeShowMessage(r.messages[0]!, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan.fallback).toBe('still')
          } else expect(r.end).toEqual({ reqId: r.reqId, type: 'error', code: 'still' })
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
    expect(await rejectionMessage(table.load(0))).toMatch(/metadata not loaded/)
    expect(() => table.digest(0)).toThrow()
    // The host's own path loads it first, so a plan builds.
    expect('refused' in await h.showPlan(docKey, job())).toBe(false)
    expect(() => h.mediaTable(docKey).digest(0)).not.toThrow()
  })
})
