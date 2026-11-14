/**
 * Server-test decks for the show-plan bundle build (piece 4 Task 6): the shapes the shared fixtures
 * (`src/testdata/show-decks.ts`) do not cover — a transition-only slide, a `cmd` node, links into hidden
 * slides, growth, an appear-and-move, and a tiny box. Built the same way, with `buildPptxDeck`.
 */
import JSZip from 'jszip'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxShapeSpec } from '../../src/testdata/ooxml-builders'
import { DocumentHost } from '../src/render'

const px = (n: number) => String(Math.round(n * 9525))
export const rect = (id: number, name: string, x: number, y: number, w: number, h: number, fill: string, extra: Partial<PptxShapeSpec> = {}): PptxShapeSpec => ({
  id, name, prst: 'rect', off: [px(x), px(y)], ext: [px(w), px(h)], fill, ...extra,
})
const tgt = (spid: number) => `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`
const ctn = (id: number, dur: number) => `<p:cTn id="${id}" dur="${dur}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
const appear = (id: () => number, spid: number) =>
  `<p:set><p:cBhvr>${ctn(id(), 1)}${tgt(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`

/** One main-sequence click step holding `behaviours` in one click effect. */
function oneStep(behaviours: (id: () => number) => string[]): string {
  let n = 3
  const id = () => n++
  const outer = id(), inner = id(), own = id()
  const body = behaviours(id).join('')
  const step = `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${own}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
    `<p:childTnLst>${body}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
  return timingXmlOf(mainSeqXml([step]))
}

/** One slide, no timing, entered with a fade. */
export const transitionOnlyDeck = () => buildPptxDeck({
  slides: [{ bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Box', 380, 260, 200, 200, '336699')], transition: { xml: '<p:transition spd="med"><p:fade/></p:transition>' } }],
})

/** One slide: click 1 makes the box (id 2) appear and plays an OLE verb on it (`p:cmd`). */
export const cmdDeck = () => buildPptxDeck({
  slides: [{
    bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Box', 380, 260, 200, 200, '336699')],
    timingXml: oneStep(id => [appear(id, 2), `<p:cmd type="verb" cmd="0"><p:cBhvr>${ctn(id(), 1)}${tgt(2)}</p:cBhvr></p:cmd>`]),
  }],
})

/**
 * Three slides, the second hidden. Slide 0 links shape 2 to slide 3 (unit 2, streamed index 1), shape 3 to
 * the hidden slide 2 (unit 1), and shape 4 to an external URL; click 1 makes shape 5 appear.
 */
export const hiddenLinkDeck = () => buildPptxDeck({
  slides: [
    {
      bg: { solid: 'FFFFFF' },
      shapes: [rect(2, 'ToThree', 80, 80, 240, 120, '336699'), rect(3, 'ToHidden', 400, 80, 240, 120, 'F4B183'), rect(4, 'Web', 80, 300, 240, 120, 'DDEBF7'), rect(5, 'Late', 400, 300, 240, 120, '70AD47')],
      links: [
        { shapeId: 2, target: 'slide3.xml', action: 'ppaction://hlinksldjump' },
        { shapeId: 3, target: 'slide2.xml', action: 'ppaction://hlinksldjump' },
        { shapeId: 4, target: 'https://example.com/', external: true },
      ],
      timingXml: oneStep(id => [appear(id, 5)]),
    },
    { bg: { solid: 'DDEBF7' }, shapes: [rect(2, 'Hidden', 300, 300, 200, 120, '336699')], hidden: true },
    { bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Three', 300, 300, 200, 120, 'C00000')] },
  ],
})

/** One slide: click 1 makes the shape (id 2, `w` × `h` at x, y) appear and grows it by `by` (`p:animScale`). */
export const growthDeck = (by: number, box: { x: number; y: number; w: number; h: number } = { x: 380, y: 260, w: 200, h: 200 }) => buildPptxDeck({
  slides: [{
    bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Grow', box.x, box.y, box.w, box.h, '336699')],
    timingXml: oneStep(id => [appear(id, 2), `<p:animScale><p:cBhvr>${ctn(id(), 500)}${tgt(2)}</p:cBhvr><p:by x="${Math.round(by * 100000)}" y="${Math.round(by * 100000)}"/></p:animScale>`]),
  }],
})

/** One slide: click 1 makes the box (id 2) appear and moves it right by `dx` slide fractions (its one visible rest is moved). */
export const appearMoveDeck = (dx = 0.25) => buildPptxDeck({
  slides: [{
    bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Title', 80, 40, 800, 100, 'DDEBF7'), rect(3, 'Mover', 80, 260, 200, 200, '336699')],
    timingXml: oneStep(id => [appear(id, 3), `<p:animMotion path="M 0 0 L ${dx} 0 E"><p:cBhvr>${ctn(id(), 500)}${tgt(3)}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`]),
  }],
})

/** One slide: click 1 makes a 3 × 3 px box (id 2) appear: too small for one forensic carrier cell. */
export const tinyBoxDeck = () => buildPptxDeck({
  slides: [{ bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Dot', 401, 301, 3, 3, '336699')], timingXml: oneStep(id => [appear(id, 2)]) }],
})

/** Opens `bytes` on a fresh host under `'full'` (hidden slides stay units). */
export async function openDeck(bytes: Uint8Array, docRef = 'deck', host = new DocumentHost()) {
  const open = await host.open(docRef, bytes, 'full')
  return { host, open, docKey: open.docKey }
}

/** One slide whose click 1 targets shape 50, which is on the master only (and shape 2, on the slide). */
export const masterTargetDeck = () => buildPptxDeck({
  masters: [{ shapes: [rect(50, 'Logo', 860, 20, 80, 40, 'C00000')], layouts: [{}] }],
  slides: [{ bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Box', 380, 260, 200, 200, '336699')], timingXml: oneStep(id => [appear(id, 2), appear(id, 50)]) }],
})

/**
 * One slide holding a group (id 10) of two boxes (ids 11 and 12): click 1 makes the group appear, click 2 makes
 * box 12 appear inside it. The group's state composes over its sub-layers (a container record on the wire).
 */
export async function groupDeck(): Promise<Uint8Array> {
  const two = (id: () => number) => [appear(id, 10)]
  let n = 3
  const id = () => n++
  const step = (body: string) => {
    const outer = id(), inner = id(), own = id()
    return `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${own}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
      `<p:childTnLst>${body}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
  }
  const s1 = step(two(id).join('')), s2 = step(appear(id, 12))
  const bytes = await buildPptxDeck({
    slides: [{
      bg: { solid: 'FFFFFF' },
      shapes: [rect(2, 'Title', 80, 40, 800, 100, 'DDEBF7'), rect(11, 'A', 200, 260, 200, 200, '336699'), rect(12, 'B', 500, 260, 200, 200, 'F4B183')],
      timingXml: timingXmlOf(mainSeqXml([s1, s2])),
    }],
  })
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/slide1.xml'
  let xml = await zip.file(path)!.async('string')
  const open = xml.search(/<p:sp>(?:(?!<p:sp>)[\s\S])*?<p:cNvPr id="11"/)
  const tail = xml.indexOf('</p:sp>', xml.search(/<p:cNvPr id="12"/)) + '</p:sp>'.length
  if (open < 0 || tail < open) throw new Error('groupDeck: shapes not found')
  const e = (v: number) => String(Math.round(v * 9525))
  const xfrm = `<a:xfrm><a:off x="${e(200)}" y="${e(260)}"/><a:ext cx="${e(500)}" cy="${e(200)}"/><a:chOff x="${e(200)}" y="${e(260)}"/><a:chExt cx="${e(500)}" cy="${e(200)}"/></a:xfrm>`
  xml = `${xml.slice(0, open)}<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrm}</p:grpSpPr>${xml.slice(open, tail)}</p:grpSp>${xml.slice(tail)}`
  zip.file(path, xml)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}
