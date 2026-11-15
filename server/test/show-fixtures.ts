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
function oneStep(behaviours: (id: () => number) => string[], start: 'indefinite' | '0' = 'indefinite'): string {
  let n = 3
  const id = () => n++
  const outer = id(), inner = id(), own = id()
  const body = behaviours(id).join('')
  const step = `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="${start}"/></p:stCondLst><p:childTnLst>` +
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

/**
 * Two slides. Slide 0 holds one text box (id 2, 80,80 600×200) of three runs on one line: "Jump" links to slide 2,
 * " plain text " has no link, "Web" links to an external URL. Click 1 makes shape 5 appear (so a plan is built).
 */
export const runLinkDeck = () => buildPptxDeck({
  slides: [
    {
      bg: { solid: 'FFFFFF' },
      shapes: [
        rect(2, 'Runs', 80, 80, 600, 200, 'FFFFFF', { paragraphs: [{ runs: [{ text: 'Jump', sz: '2800' }, { text: ' plain text ', sz: '2800' }, { text: 'Web', sz: '2800' }] }] }),
        rect(5, 'Late', 400, 360, 240, 120, '70AD47'),
      ],
      links: [
        { shapeId: 2, target: 'slide2.xml', action: 'ppaction://hlinksldjump', onRun: 0 },
        { shapeId: 2, target: 'https://example.com/', external: true, onRun: 2 },
      ],
      timingXml: oneStep(id => [appear(id, 5)]),
    },
    { bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Two', 300, 300, 200, 120, 'C00000')] },
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

/** One slide, no click steps: the box (id 2) fades in with the slide (an autoplayed main-sequence group, delay 0). */
export const autoplayFadeDeck = () => buildPptxDeck({
  slides: [{
    bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Box', 380, 260, 200, 200, '336699')],
    timingXml: oneStep(id => [appear(id, 2), `<p:animEffect transition="in" filter="fade"><p:cBhvr>${ctn(id(), 500)}${tgt(2)}</p:cBhvr></p:animEffect>`], '0'),
  }],
})

/** `bytes` with every `filter="from"` in slide 1's timing replaced by `filter="to"`. */
export async function withFilter(bytes: Uint8Array, from: string, to: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/slide1.xml'
  const xml = await zip.file(path)!.async('string')
  if (!xml.includes(`filter="${from}"`)) throw new Error(`withFilter: no ${from}`)
  zip.file(path, xml.split(`filter="${from}"`).join(`filter="${to}"`))
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}

/** `triggerFadeDeck` with its fade replaced by a motion path 240 px right (same `cBhvr`): a triggered motion entrance. */
export async function withTriggeredMove(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/slide1.xml'
  const xml = await zip.file(path)!.async('string')
  const out = xml.replace(/<p:animEffect transition="in" filter="fade">(<p:cBhvr>[\s\S]*?<\/p:cBhvr>)<\/p:animEffect>/, '<p:animMotion path="M 0 0 L 0.25 0 E">$1</p:animMotion>')
  if (out === xml) throw new Error('withTriggeredMove: no fade')
  zip.file(path, out)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}

/** `bytes` with its `style.visibility` sets removed: the triggered target is shown from the start. */
export async function withoutEntrance(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/slide1.xml'
  const xml = await zip.file(path)!.async('string')
  const out = xml.replace(/<p:set>[\s\S]*?<\/p:set>/g, '').replace(/presetClass="entr"/g, 'presetClass="path"')
  if (out === xml) throw new Error('withoutEntrance: no set')
  zip.file(path, out)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}

/** An interactive sequence started by a click on `trigger`, holding `body` in one effect. */
function triggerSeq(id: () => number, trigger: number, body: string, cls: 'entr' | 'path' = 'entr'): string {
  const seq = id(), par = id(), own = id()
  const cond = `<p:cond evt="onClick" delay="0">${tgt(trigger)}</p:cond>`
  return `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${seq}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
    `<p:stCondLst>${cond}</p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>` +
    `<p:par><p:cTn id="${par}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${own}" presetID="1" presetClass="${cls}" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:nextCondLst>${cond}</p:nextCondLst></p:seq>`
}

/**
 * The reviewer's reveal-then-move probe (fix round 3), on `triggerFadeDeck`'s layout: trigger A, a click on the
 * button (id 3), fades the initially hidden box (id 2) in; trigger B, a click on the box itself, moves it 240 px
 * right (a held motion path, no entrance). Only "A, then B" shows the box at its moved rest.
 */
export function revealThenMoveDeck(): Promise<Uint8Array> {
  let n = 3
  const id = () => n++
  const reveal = triggerSeq(id, 3, appear(id, 2) + `<p:animEffect transition="in" filter="fade"><p:cBhvr>${ctn(id(), 500)}${tgt(2)}</p:cBhvr></p:animEffect>`)
  const move = triggerSeq(id, 2, `<p:animMotion path="M 0 0 L 0.25 0 E"><p:cBhvr>${ctn(id(), 500)}${tgt(2)}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`, 'path')
  return buildPptxDeck({
    slides: [{
      bg: { solid: 'FFFFFF' },
      shapes: [rect(3, 'Button', 80, 560, 160, 80, 'F4B183'), rect(2, 'Box', 380, 200, 200, 200, '336699')],
      timingXml: timingXmlOf(reveal + move),
    }],
  })
}

// ---------------------------------------------------------------------------
// Morph pairs (piece 4 Task 7, D11)
// ---------------------------------------------------------------------------

const morphOf = (durMs: number) => ({ xml: `<p:transition spd="slow" p14:dur="${durMs}"><p159:morph xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" option="byObject"/></p:transition>` })
const MORPH = morphOf(1200)

/** Two slides; slide 0's box (id 2) is shown and its `Late` box (id 3) appears on click 1; slide 1 moves both, entered by Morph. */
export const morphLateDeck = () => buildPptxDeck({
  slides: [
    { bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Box', 80, 80, 200, 140, 'D06040'), rect(3, 'Late', 80, 400, 200, 140, '336699')], timingXml: oneStep(id => [appear(id, 3)]) },
    { bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Box', 600, 300, 200, 140, 'D06040'), rect(3, 'Late', 600, 80, 200, 140, '336699')], transition: MORPH },
  ],
})

/**
 * Two slides, entered by Morph: `Under` (id 4) and `Cover` (id 5) stay put on both, `Cover` hiding part of `Under`
 * (removed pixels); `Mover` (id 2), between them in z, crosses that part: its swept hull covers a masked layer.
 */
export const morphSweepDeck = (durMs = 1200) => buildPptxDeck({
  slides: [
    { bg: { solid: 'FFFFFF' }, shapes: [rect(4, 'Under', 300, 200, 300, 300, '70AD47'), rect(2, 'Mover', 60, 300, 100, 100, 'C00000'), rect(5, 'Cover', 380, 280, 140, 140, '1F4E79')] },
    { bg: { solid: 'FFFFFF' }, shapes: [rect(4, 'Under', 300, 200, 300, 300, '70AD47'), rect(2, 'Mover', 760, 300, 100, 100, 'C00000'), rect(5, 'Cover', 380, 280, 140, 140, '1F4E79')], transition: morphOf(durMs) },
  ],
})

/** Two slides of `n` small boxes each (ids 2 …), the second entered by Morph: past 256 objects the Morph cannot run. */
export const manyObjectsMorphDeck = (n = 257) => {
  const boxes = (dx: number) => Array.from({ length: n }, (_, i) => rect(2 + i, `B${i}`, 10 + (i % 32) * 28 + dx, 10 + Math.floor(i / 32) * 60, 20, 20, '336699'))
  return buildPptxDeck({ slides: [{ bg: { solid: 'FFFFFF' }, shapes: boxes(0) }, { bg: { solid: 'FFFFFF' }, shapes: boxes(4), transition: MORPH }] })
}
