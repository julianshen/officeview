/**
 * Slideshow PR #18 review 8: the flattened rung renders at checkpoint granularity (the spec's checkpoint-cut
 * fallback). While an effect runs, the one full-slide frame shows the checkpoint it is heading to; it is not
 * re-rasterised and re-uploaded every animation frame.
 *
 * The residency is wrapped so every slide settles on the flattened rung (its own file: `vi.mock` is per
 * module). jsdom with node-canvas; every clock is driven. Decks are "Office-shaped synthetic".
 */
import { afterEach, expect, it, vi } from 'vitest'
import { buildPptxDeck, mainSeqXml, timingXmlOf } from '../src/testdata/ooxml-builders'
import { installShowSeams, loadDeck, startShow, type ShowSeams } from './helpers/slideshow-show-rig'

vi.mock('../src/slideshow/residency', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/residency')>()
  return {
    ...real,
    createResidency: (...args: Parameters<typeof real.createResidency>) => {
      const r = real.createResidency(...args)
      const ensure = r.ensure.bind(r)
      r.ensure = (slide, level) => ensure(slide, level === 'terminal' ? level : 'flattened')
      return r
    },
  }
})

let seams: ShowSeams | undefined
afterEach(() => { seams?.restore(); seams = undefined; vi.restoreAllMocks(); document.body.replaceChildren() })

const W = 12192000, H = 6858000
/** Shape 3 fades in over 1000 ms on the first click. */
const fadeIn = '<p:par><p:cTn id="10" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
  '<p:set><p:cBhvr><p:cTn id="11" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="3"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>' +
  '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="12" dur="1000"/><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:animEffect>' +
  '</p:childTnLst></p:cTn></p:par>'
const step = `<p:par><p:cTn id="20" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="21" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${fadeIn}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

it('a running effect at the flattened rung shows its target checkpoint, rasterised once, not every frame', async () => {
  seams = installShowSeams({ width: 1280, height: 720 })
  const doc = await loadDeck(await buildPptxDeck({
    slideSize: [W, H],
    slides: [{
      shapes: [
        { id: 2, name: 'Back', off: ['0', '0'], ext: [String(W), String(H)], fill: 'FFFFFF' },
        { id: 3, name: 'Box', off: [String(W / 4), String(H / 4)], ext: [String(W / 2), String(H / 2)], fill: '0000CC' },
      ],
      timingXml: timingXmlOf(mainSeqXml([step])),
    }],
  }))
  const r = await startShow(doc, { seams, size: { width: 1280, height: 720 } })
  try {
    await r.tick(100)
    expect(r.s.stage!.level).toBe('flattened')
    const created = vi.spyOn(document, 'createElement')
    const canvases = () => created.mock.calls.filter(c => c[0] === 'canvas').length
    r.s.next()
    await r.tick(0)
    const afterCut = canvases()
    // Mid-fade: the frame already shows the box at full opacity (the checkpoint it is heading to).
    const c = r.frameCanvas()!
    const mid = () => [...c.getContext('2d')!.getImageData(c.width >> 1, c.height >> 1, 1, 1).data.slice(0, 3)]
    for (let i = 0; i < 10; i++) await r.tick(50)
    expect(mid()).toEqual([0, 0, 0xcc])
    // No re-rasterisation while the effect runs.
    expect(canvases()).toBe(afterCut)
    await r.tick(1000)
    expect(canvases()).toBe(afterCut)
    expect(mid()).toEqual([0, 0, 0xcc])
  } finally { r.exit() }
})
