/**
 * Slideshow plan S1-22: per-frame work. One frame computes the shown slide's
 * render state once: `frame()` computes it and the accessibility update
 * reuses it (S1-20 left a second `stateAt` per tick in `updateA11y`).
 *
 * jsdom with node-canvas; every clock is driven. The deck is "Office-shaped
 * synthetic".
 */
import { afterEach, beforeEach, expect, it } from 'vitest'
import { sessionDeps } from '../src/slideshow/player/session'
import type { StageModel } from '../src/slideshow/player/stage'
import { buildPptxDeck, mainSeqXml, timingXmlOf } from '../src/testdata/ooxml-builders'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

/** A click step fading shape `spid` in over 1000 ms. */
const fadeStep = (id: number, spid: number) =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 2}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
  `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${id + 4}" dur="1000"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

let seams: ShowSeams
let stateCalls = 0
let realStage: typeof sessionDeps.stage

beforeEach(() => {
  seams = installShowSeams()
  stateCalls = 0
  realStage = sessionDeps.stage
  // Count the scheduler snapshots the shown slide's timeline computes.
  sessionDeps.stage = (deps) => {
    const stage = realStage(deps)
    const enter = stage.enter.bind(stage)
    stage.enter = (model: StageModel, t, neighbours) => {
      const tl = model.timeline
      const stateAt = tl.stateAt.bind(tl)
      tl.stateAt = (cursor, at) => { stateCalls++; return stateAt(cursor, at) }
      return enter(model, t, neighbours)
    }
    return stage
  }
})
afterEach(() => {
  sessionDeps.stage = realStage
  seams.restore()
  document.body.replaceChildren()
})

it('one frame computes the render state once: the accessibility update reuses the frame\'s state', async () => {
  const doc = await loadDeck(await buildPptxDeck({
    slides: [{
      shapes: [{ id: 2, name: 'Box', off: ['914400', '914400'], ext: ['3657600', '914400'], fill: '336699', paragraphs: [{ runs: [{ text: 'Fades in' }] }] }],
      timingXml: timingXmlOf(mainSeqXml([fadeStep(10, 2)])),
    }],
  }))
  const r = await startShow(doc, { seams })
  r.s.next()
  await r.tick(0)
  await r.tick(500)
  expect(r.state().get('sp:2')?.opacity).toBeCloseTo(0.5, 6)
  expect(r.text()).toEqual(['Fades in'])
  // Mid-animation: every frame changes the state, so the accessibility key is recomputed each tick.
  for (const step of [16, 17, 16]) {
    seams.clock.advance(step)
    stateCalls = 0
    r.s.tick()
    expect(stateCalls).toBe(1)
  }
  r.exit()
})
