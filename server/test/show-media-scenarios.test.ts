/**
 * Piece 5 Task 5, fix round 1: the reviewer's soundness scenarios (C1–C4), executed with media on and off.
 *
 * Controller ruling: with media on a slide classifies exactly as with media off; a media end (audio or video) that
 * can reach drawing, or an analysis that cannot tell, plays the still. Each scenario states that property.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { installCanvasBackend, type CanvasBackend } from '../../src/core/canvas-backend'
import { HIDDEN_INK, mediaDeck, type MediaShapeSpec } from '../../src/testdata/show-decks'
import { serverBackend } from '../src/backend'
import { createMediaTemp } from '../src/media/temp'
import { DocumentHost } from '../src/render'
import type { ShowPlanJob } from '../src/show/build'
import { ffmpegForTests } from './media/ffmpeg-gate'
import { classify, reconstruct } from './show-classify-helper'

let backend: CanvasBackend
let previous: CanvasBackend | undefined
beforeAll(() => { backend = serverBackend(); previous = installCanvasBackend(backend) })
afterAll(() => { installCanvasBackend(previous) })

const px = (n: number) => String(Math.round(n * 9525))
const box = (id: number, x: number, y: number, fill: string) => ({ id, name: `B${id}`, prst: 'rect', off: [px(x), px(y)] as [string, string], ext: [px(160), px(120)] as [string, string], fill })
const INK = box(2, 400, 300, HIDDEN_INK)
const tgt = (spid: number) => `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`
const show = (id: number, spid: number, o: { delay?: number; dur?: number | 'indefinite'; fill?: string } = {}) =>
  `<p:set><p:cBhvr><p:cTn id="${id}" dur="${o.dur ?? 1}" fill="${o.fill ?? 'hold'}"><p:stCondLst><p:cond delay="${o.delay ?? 0}"/></p:stCondLst></p:cTn>${tgt(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
const media = (kind: 'audio' | 'video', id: number, spid: number, o: { start?: string; repeat?: 'indefinite' } = {}) =>
  `<p:${kind}><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"${o.repeat ? ` repeatCount="${o.repeat}"` : ''}><p:stCondLst>${o.start ?? '<p:cond delay="0"/>'}</p:stCondLst></p:cTn>${tgt(spid)}</p:cMediaNode></p:${kind}>`
const onEnd = (tn: number, delay = 0) => `<p:cond evt="onEnd" delay="${delay}"><p:tn val="${tn}"/></p:cond>`
/** An entrance container (`entr`) with raw attributes, conditions and children. */
const entr = (id: number, attrs: string, conds: string, kids: string) =>
  `<p:par><p:cTn id="${id}" presetID="1" presetClass="entr" presetSubtype="0"${attrs} nodeType="afterEffect">${conds}<p:childTnLst>${kids}</p:childTnLst></p:cTn></p:par>`

const A3: MediaShapeSpec = { spid: 3, file: 'a.m4a' } // 2000 ms
const A4: MediaShapeSpec = { spid: 4, file: 'a.mp3' } // 3000 ms
const V3: MediaShapeSpec = { spid: 3, file: 'av.mp4', kind: 'video' }

/** The scenarios whose property is "plays its still" (build level, media on and off). */
const STILL_SCENARIOS: Record<string, () => Promise<Uint8Array>> = {
  // C1: audio begins at 100 ms; its end reveals the ink (a 0-step slide).
  C1: () => mediaDeck({ slides: [{ boxes: [INK], media: [A3], timing: media('audio', 3, 3, { start: '<p:cond delay="100"/>' }) + entr(5, ' fill="hold"', `<p:stCondLst>${onEnd(3)}</p:stCondLst>`, show(6, 2)) }] }),
  // C2: the ink shows at 20 ms inside a fill-remove entrance that the audio's end cuts short.
  C2: () => mediaDeck({ slides: [{ boxes: [INK], media: [A3], timing: media('audio', 3, 3) + entr(5, ' dur="10000" fill="remove"', `<p:stCondLst><p:cond delay="20"/></p:stCondLst><p:endCondLst>${onEnd(3)}</p:endCondLst>`, show(6, 2, { dur: 10000, fill: 'remove' })) }] }),
  // C3b: the audio sits in a reversing container with a finite dur; the ink enters 300 ms after its end.
  C3b: () => mediaDeck({ slides: [{ boxes: [INK], media: [A3], timing: `<p:par><p:cTn id="20" dur="1200" autoRev="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${media('audio', 3, 3, { start: '<p:cond delay="1000"/>' })}</p:childTnLst></p:cTn></p:par>` + entr(5, ' fill="hold"', `<p:stCondLst>${onEnd(3, 300)}</p:stCondLst>`, show(6, 2)) }] }),
  // C4.1: a seq whose second child (audio 4) begins on audio 3's end, cutting short the first child, which shows the
  // ink at 100 ms (fill remove: the cut hides it).
  'C4.1': () => mediaDeck({ slides: [{ boxes: [INK], media: [A3, A4], timing: media('audio', 3, 3) +
    `<p:seq><p:cTn id="30" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="31" presetID="1" presetClass="entr" presetSubtype="0" dur="indefinite" fill="remove" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${show(32, 2, { delay: 100, dur: 'indefinite' })}</p:childTnLst></p:cTn></p:par>` +
    `${media('audio', 4, 4, { start: onEnd(3) })}</p:childTnLst></p:cTn></p:seq>` }] }),
  // C4.2: the video's end fires a `stop` on audio 4 (indefinite), whose end reveals the ink.
  'C4.2': () => mediaDeck({ slides: [{ boxes: [INK], media: [V3, A4], timing: media('video', 3, 3) +
    `<p:cmd type="call" cmd="stop"><p:cBhvr><p:cTn id="10" dur="1" fill="hold"><p:stCondLst>${onEnd(3)}</p:stCondLst></p:cTn>${tgt(4)}</p:cBhvr></p:cmd>` +
    media('audio', 4, 4, { repeat: 'indefinite' }) + entr(5, ' fill="hold"', `<p:stCondLst>${onEnd(4)}</p:stCondLst>`, show(6, 2)) }] }),
}

const info = await ffmpegForTests()

describe.skipIf(!info)('reviewer scenarios at the build (media on and off)', () => {
  const root = mkdtempSync(join(tmpdir(), 'ov-t5s-'))
  let temp: Awaited<ReturnType<typeof createMediaTemp>>
  beforeAll(async () => { temp = await createMediaTemp(root) })
  afterAll(async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) })
  const job = (media: boolean): ShowPlanJob => ({ unitIndex: 0, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false, media })
  const planOf = async (bytes: Uint8Array, media: boolean) => {
    const h = new DocumentHost()
    h.configureMedia({ info: info!, dir: temp.dir })
    const { docKey } = await h.open(`s${Math.random()}`, bytes, 'visible-only')
    return h.showPlan(docKey, job(media))
  }

  for (const [name, deck] of Object.entries(STILL_SCENARIOS)) {
    for (const media of [false, true]) {
      it(`${name} media ${media ? 'on' : 'off'}: plays its still`, async () => {
        const r = await planOf(await deck(), media)
        expect('refused' in r && r.refused).toBe('still')
      }, 60_000)
    }
  }

  // C3a: an audio node under a reversing (or repeating) ancestor never ships a playable record, so the sound's clock
  // never depends on the activation query under a mirrored local time: the slide is a still (the analysis does not
  // model the ancestor's timing), or the record ships refused media-unavailable.
  for (const anc of ['autoRev="1"', 'repeatCount="2000"'] as const) {
    for (const media of [false, true]) {
      it(`C3a (${anc}) media ${media ? 'on' : 'off'}: no playable record`, async () => {
        const bytes = await mediaDeck({ slides: [{ boxes: [box(2, 400, 300, '336699')], media: [A3], timing: `<p:par><p:cTn id="20" dur="1200" ${anc} fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${media_('audio', 3, 3, '<p:cond delay="1000"/>')}</p:childTnLst></p:cTn></p:par>` }] })
        const r = await planOf(bytes, media)
        if ('refused' in r) { expect(r.refused).toBe('still'); return }
        expect(r.plan.media?.map(m => m.refused)).toEqual(['media-unavailable'])
      }, 60_000)
    }
  }
})
const media_ = (kind: 'audio', id: number, spid: number, start: string) => media(kind, id, spid, { start })

describe('C2 without media: a non-media end condition cuts a fill-remove entrance short (piece 4)', () => {
  it('the slide plays its still, or the ink is kept and reconstructs while shown', async () => {
    // Shape 9's set (500 ms) ends at 500; the entrance (shown from 20 ms, fill remove) ends with it.
    const bytes = await mediaDeck({ slides: [{ boxes: [INK, box(9, 40, 40, '336699')], timing:
      `<p:par><p:cTn id="40" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${show(41, 9, { dur: 500 })}</p:childTnLst></p:cTn></p:par>` +
      entr(5, ' dur="10000" fill="remove"', `<p:stCondLst><p:cond delay="20"/></p:stCondLst><p:endCondLst>${onEnd(41)}</p:endCondLst>`, show(6, 2, { dur: 10000, fill: 'remove' })) }] })
    const run = await classify(bytes)
    if (run.c.fallback === 'still') return
    const ink = run.ps.layers.find(l => l.kind === 'target' && l.key === 'sp:2')!
    expect(run.c.drop.has(ink)).toBe(false)
    expect(reconstruct(run, run.timeline.initial(), 250).diff).toBe(0)
  })
})
