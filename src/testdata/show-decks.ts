/**
 * Shared slideshow fixture decks for streamed animation playback (piece 4 plan D18).
 *
 * Every fixture is small, deterministic and "Office-shaped synthetic": built
 * with `buildPptxDeck`, its timing with `mainSeqXml` / `timingXmlOf`. Each
 * returns the package bytes; `loadShowDeck` parses them DOM-free (bun or
 * vitest) into the document and each slide's `ShowSlide` metadata.
 *
 * The slide is 960 × 720 px (the builder's default 9144000 × 6858000 EMU).
 * Every adversary deck hides its never-shown ink in `HIDDEN_INK` (#13F00D),
 * the hidden-ink signature the wire audit looks for.
 */
import { OfficePackage } from '../core/zip'
import { enrichSlideAppearance } from '../pptx/appearance'
import { parsePptx } from '../pptx/parse'
import type { PptxDocument } from '../pptx/types'
import { createDiagnosticSink } from '../slideshow/diagnostics'
import { openShowPackage, readShowIndex } from '../slideshow/read/deck'
import { shapeResolver } from '../slideshow/read/resolve'
import { createSlideMetadataStore } from '../slideshow/read/store'
import { readSlideTiming } from '../slideshow/read/timing'
import type { DiagnosticSink, Lease, ShowIndex, ShowSlide, SlideShowWorkspace } from '../slideshow/types'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxDeckSpec, type PptxShapeSpec, type PptxSlideSpec } from './ooxml-builders'

/** The hidden-ink signature: no transmitted asset may carry this colour (ΔRGB ≤ 6, α > 0). */
export const HIDDEN_INK = '13F00D'

/** Slide px → EMU. */
const px = (n: number) => String(Math.round(n * 9525))

/** A rectangle at slide px (x, y), size (w, h). */
const rect = (id: number, name: string, x: number, y: number, w: number, h: number, fill: string, extra: Partial<PptxShapeSpec> = {}): PptxShapeSpec => ({
  id, name, prst: 'rect', off: [px(x), px(y)], ext: [px(w), px(h)], fill, ...extra,
})

// ---------------------------------------------------------------------------
// Timing XML
// ---------------------------------------------------------------------------

/** Sequential `p:cTn` ids for one slide's timing (the root takes the next free id in `timingXmlOf`). */
function ids(start = 3): () => number {
  let n = start
  return () => n++
}

const tgt = (spid: number) => `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`
const attrs = (...names: string[]) => `<p:attrNameLst>${names.map(n => `<p:attrName>${n}</p:attrName>`).join('')}</p:attrNameLst>`
const ctn = (id: number, dur: number | 'indefinite', delay = 0) =>
  `<p:cTn id="${id}" dur="${dur}" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>`

/** `p:set` of `style.visibility`. */
const setVisibility = (id: () => number, spid: number, to: 'visible' | 'hidden', delay = 0) =>
  `<p:set><p:cBhvr>${ctn(id(), 1, delay)}${tgt(spid)}${attrs('style.visibility')}</p:cBhvr><p:to><p:strVal val="${to}"/></p:to></p:set>`

/** `p:anim` of one attribute over string keyframes `[tm, value]` (tm in 1/1000 %). */
const animStr = (id: () => number, spid: number, attr: string, dur: number, keys: Array<[number, string]>) =>
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">${ctn(id(), dur)}${tgt(spid)}${attrs(attr)}</p:cBhvr>` +
  `<p:tavLst>${keys.map(([tm, v]) => `<p:tav tm="${tm}"><p:val><p:strVal val="${v}"/></p:val></p:tav>`).join('')}</p:tavLst></p:anim>`

/** `p:anim` of one attribute over number keyframes. */
const animNum = (id: () => number, spid: number, attr: string, dur: number, keys: Array<[number, number]>) =>
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr>${ctn(id(), dur)}${tgt(spid)}${attrs(attr)}</p:cBhvr>` +
  `<p:tavLst>${keys.map(([tm, v]) => `<p:tav tm="${tm}"><p:val><p:fltVal val="${v}"/></p:val></p:tav>`).join('')}</p:tavLst></p:anim>`

/** `p:animEffect` (fade, wipe(down), …), optionally pinned at `progress`. */
const animEffect = (id: () => number, spid: number, transition: 'in' | 'out', filter: string, dur: number, progress?: number) =>
  `<p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr>${ctn(id(), dur)}${tgt(spid)}</p:cBhvr>` +
  `${progress === undefined ? '' : `<p:progress><p:fltVal val="${progress}"/></p:progress>`}</p:animEffect>`

/** One effect: a preset-classed `clickEffect` (or `withEffect`) par holding its behaviours. */
interface Effect { cls: 'entr' | 'exit' | 'emph' | 'path'; preset: number; behaviours: (id: () => number) => string[]; node?: 'clickEffect' | 'withEffect' }

const effectXml = (id: () => number, e: Effect) => {
  const own = id()
  return `<p:par><p:cTn id="${own}" presetID="${e.preset}" presetClass="${e.cls}" presetSubtype="0" fill="hold" nodeType="${e.node ?? 'clickEffect'}">` +
    `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${e.behaviours(id).join('')}</p:childTnLst></p:cTn></p:par>`
}

/** One click step of the main sequence holding `effects` (concurrent); `start` "0" makes it autoplay with the slide. */
const clickStep = (id: () => number, effects: Effect[], start: 'indefinite' | '0' = 'indefinite') => {
  const outer = id(), inner = id()
  return `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="${start}"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `${effects.map((e, i) => effectXml(id, i === 0 ? e : { ...e, node: e.node ?? 'withEffect' })).join('')}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
}

/** A main sequence of click steps, wrapped in the timing root. */
const mainSequence = (steps: Effect[][]) => {
  const id = ids()
  const xml = steps.map(s => clickStep(id, s))
  return timingXmlOf(mainSeqXml(xml))
}

/** Entrance: appear (a `set` of visibility). */
const appear = (spid: number): Effect => ({ cls: 'entr', preset: 1, behaviours: id => [setVisibility(id, spid, 'visible')] })
/** Entrance: fade in over `dur` ms. */
const fadeIn = (spid: number, dur = 500): Effect => ({ cls: 'entr', preset: 10, behaviours: id => [setVisibility(id, spid, 'visible'), animEffect(id, spid, 'in', 'fade', dur)] })
/** Entrance: fly in from the left over `dur` ms. */
const flyIn = (spid: number, dur = 500): Effect => ({
  cls: 'entr', preset: 2, behaviours: id => [
    setVisibility(id, spid, 'visible'),
    animStr(id, spid, 'ppt_x', dur, [[0, '0-#ppt_w/2'], [100000, '#ppt_x']]),
    animStr(id, spid, 'ppt_y', dur, [[0, '#ppt_y'], [100000, '#ppt_y']]),
  ],
})
/** Exit: disappear. */
const disappear = (spid: number): Effect => ({ cls: 'exit', preset: 1, behaviours: id => [setVisibility(id, spid, 'hidden')] })
/** Emphasis: spin one full turn over `dur` ms. */
const spin = (spid: number, dur: number): Effect => ({
  cls: 'emph', preset: 8, behaviours: id => [`<p:animRot by="21600000"><p:cBhvr>${ctn(id(), dur)}${tgt(spid)}${attrs('r')}</p:cBhvr></p:animRot>`],
})

/** An interactive sequence started by a click on `trigger`, holding `effects`. */
const interactiveSeq = (id: () => number, trigger: number, effects: Effect[]) => {
  const seq = id(), par = id()
  const cond = `<p:cond evt="onClick" delay="0">${tgt(trigger)}</p:cond>`
  return `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${seq}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
    `<p:stCondLst>${cond}</p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>` +
    `<p:par><p:cTn id="${par}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `${effects.map(e => effectXml(id, e)).join('')}</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn><p:nextCondLst>${cond}</p:nextCondLst></p:seq>`
}

const MORPH_NS = 'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"'
const morphTransition = (option: 'byObject' | 'byWord' | 'byChar') => ({
  xml: `<p:transition spd="slow" p14:dur="1200"><p159:morph ${MORPH_NS} option="${option}"/></p:transition>`,
})

const deck = (slides: PptxSlideSpec[], extra: Partial<PptxDeckSpec> = {}) => buildPptxDeck({ slides, ...extra })

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

/**
 * One slide: a static title (id 2), a box (id 3) that flies in from the left
 * on click 1, and a static footer (id 4) above it. Plans as static
 * (background), target `sp:3`, static.
 */
export function flyInDeck(): Promise<Uint8Array> {
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [
      rect(2, 'Title', 80, 40, 800, 100, 'DDEBF7', { paragraphs: [{ runs: [{ text: 'Fly in' }] }] }),
      rect(3, 'Box', 380, 260, 200, 200, '336699'),
      rect(4, 'Footer', 80, 600, 800, 60, 'F4B183'),
    ],
    timingXml: mainSequence([[flyIn(3)]]),
  }])
}

export interface MorphDeckOptions {
  option?: 'byObject' | 'byWord' | 'byChar'
  /** Words in the box's text (default 2); `byWord` / `byChar` split it into units. */
  words?: number
}

/**
 * One slide (as `lo-bnc591147`): the main sequence's first group autoplays a media call (`p:cmd playFrom`) on the
 * box (id 3) held for `holdMs`, so a → during it seeks the group; there is no click step. Task 14 ruling: a
 * streamed `cmd` stays in the tree, inert, so the clicks equal local.
 */
export function cmdHoldDeck(holdMs = 3000): Promise<Uint8Array> {
  const id = ids()
  const outer = id(), inner = id(), effect = id(), behaviour = id()
  // The group starts with the main sequence (`onBegin` of its cTn), as PowerPoint writes an autoplayed media call.
  const group = `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="__SEQ__"/></p:cond></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${effect}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>${ctn(behaviour, holdMs)}${tgt(3)}</p:cBhvr></p:cmd>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
  const steps = [group]
  const seq = mainSeqXml(steps)
  const seqId = /<p:cTn id="(\d+)" dur="indefinite" nodeType="mainSeq"/.exec(seq)![1]!
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Box', 380, 260, 200, 200, '336699'), rect(3, 'Media', 80, 560, 160, 80, 'F4B183')],
    timingXml: timingXmlOf(seq.replace('__SEQ__', seqId)),
  }])
}

/**
 * Three white slides for slide-change checks (PR #78 round 2): a fly-in (as `flyInDeck`), a static slide entered
 * by a cut, and a second fly-in. Nothing on them is black.
 */
export function navDeck(): Promise<Uint8Array> {
  const fly = (title: string) => ({
    bg: { solid: 'FFFFFF' },
    shapes: [
      rect(2, 'Title', 80, 40, 800, 100, 'DDEBF7', { paragraphs: [{ runs: [{ text: title }] }] }),
      rect(3, 'Box', 380, 260, 200, 200, '336699'),
    ],
    timingXml: mainSequence([[flyIn(3)]]),
  })
  return deck([
    fly('One'),
    { bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Two', 120, 120, 300, 200, 'F4B183')] },
    fly('Three'),
  ])
}

/**
 * Two slides; the second enters with a 1 s fade transition, and its box (id 3) fades in with the slide (the main
 * sequence's first group starts at delay 0): its arrival frame has the box hidden (PR #78 round 7).
 */
export function autoplayTransitionDeck(): Promise<Uint8Array> {
  const id = ids()
  return deck([
    { bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'One', 120, 120, 300, 200, 'F4B183')] },
    {
      bg: { solid: 'FFFFFF' },
      shapes: [rect(2, 'Title', 80, 40, 800, 100, 'DDEBF7'), rect(3, 'Box', 380, 260, 200, 200, '336699')],
      transition: { xml: '<p:transition spd="slow" p14:dur="1000"><p:fade/></p:transition>' },
      timingXml: timingXmlOf(mainSeqXml([clickStep(id, [fadeIn(3, 500)], '0')])),
    },
  ])
}

/**
 * Two slides, each holding shape id 2 named "Box" (the same kind, a text
 * rectangle), moved and resized; the second enters with a Morph
 * (`option`, default `byObject`).
 */
export function morphDeck(options: MorphDeckOptions = {}): Promise<Uint8Array> {
  const words = Math.max(0, Math.floor(options.words ?? 2))
  const text = Array.from({ length: words }, (_, i) => `w${i}`).join(' ')
  const box = (x: number, y: number, w: number, h: number, fill: string) =>
    rect(2, 'Box', x, y, w, h, fill, text ? { paragraphs: [{ runs: [{ text, sz: '1200' }] }] } : {})
  return deck([
    { bg: { solid: 'FFFFFF' }, shapes: [box(80, 80, 240, 160, 'D06040')] },
    { bg: { solid: 'FFFFFF' }, shapes: [box(520, 360, 360, 240, 'D06040')], transition: morphTransition(options.option ?? 'byObject') },
  ])
}

/**
 * Two slides; the second enters with the Cube transition, a 3-D row on the
 * shader path (`p14:prism`, PowerPoint's Cube, with a fade fallback).
 */
export function cubeDeck(): Promise<Uint8Array> {
  return deck([
    { bg: { solid: '1F4E79' }, shapes: [rect(2, 'One', 120, 120, 300, 200, 'F4B183')] },
    {
      bg: { solid: 'FFFFFF' }, shapes: [rect(2, 'Two', 520, 380, 300, 200, '336699')],
      transition: { xml: '<p:transition spd="slow" p14:dur="1000"><p14:prism/></p:transition>', wrapInAlternateContent: { requires: 'p14', fallbackXml: '<p:transition spd="slow"><p:fade/></p:transition>' } },
    },
  ])
}

/**
 * One slide: a trigger button (id 3) whose click starts an interactive
 * sequence fading in the box (id 2). No main sequence.
 */
export function triggerDeck(): Promise<Uint8Array> {
  const id = ids()
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [
      rect(2, 'Box', 380, 200, 200, 200, '336699'),
      rect(3, 'Button', 80, 560, 160, 80, 'F4B183', { paragraphs: [{ runs: [{ text: 'Go' }] }] }),
    ],
    timingXml: timingXmlOf(interactiveSeq(id, 3, [fadeIn(2)])),
  }])
}

/**
 * Two slides. Slide 0: a shape (id 2) linking to slide 2, a shape (id 3)
 * linking to an external URL, and a text run link on shape id 4.
 */
export function linkDeck(): Promise<Uint8Array> {
  return deck([
    {
      bg: { solid: 'FFFFFF' },
      shapes: [
        rect(2, 'Next', 80, 80, 240, 120, '336699'),
        rect(3, 'Web', 400, 80, 240, 120, 'F4B183'),
        rect(4, 'Text', 80, 300, 560, 120, 'DDEBF7', { paragraphs: [{ runs: [{ text: 'See the web' }] }] }),
      ],
      links: [
        { shapeId: 2, target: 'slide2.xml', action: 'ppaction://hlinksldjump' },
        { shapeId: 3, target: 'https://example.com/', external: true },
        { shapeId: 4, target: 'https://example.com/run', external: true, onRun: true },
      ],
    },
    { bg: { solid: 'DDEBF7' }, shapes: [rect(2, 'Two', 300, 300, 200, 120, '336699')] },
  ])
}

/** One slide: click 1 spins the box (id 2) for 9 s (an emphasis longer than the 6 s keyframe cap). */
export function longStepDeck(): Promise<Uint8Array> {
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Box', 380, 260, 200, 200, '336699')],
    timingXml: mainSequence([[spin(2, 9000)]]),
  }])
}

/** A motion path moving `spid` by (dx, dy) slide fractions over `dur` ms and holding there. */
const moveBy = (spid: number, dx: number, dy: number, dur = 500): Effect => ({
  cls: 'path', preset: 0, behaviours: id => [`<p:animMotion path="M 0 0 L ${dx} ${dy} E"><p:cBhvr>${ctn(id(), dur)}${tgt(spid)}${attrs('ppt_x', 'ppt_y')}</p:cBhvr></p:animMotion>`],
})

/** One slide: click 1 moves the box (id 2) right 240 px and holds; click 2 moves it down 120 px (two distinct rests, D8). */
export function moveTwiceDeck(): Promise<Uint8Array> {
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Box', 80, 80, 160, 120, '336699')],
    timingXml: mainSequence([[moveBy(2, 0.25, 0)], [moveBy(2, 0, 1 / 6)]]),
  }])
}

/** One slide: `n` click steps, each spinning the box (id 2) one turn over `durMs` ms. */
export function spinStepsDeck(n: number, durMs: number): Promise<Uint8Array> {
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Box', 380, 260, 200, 200, '336699')],
    timingXml: mainSequence(Array.from({ length: Math.max(1, Math.floor(n)) }, () => [spin(2, durMs)])),
  }])
}

/** One slide: a trigger button (id 3) whose click recolours the box's fill (id 2) to #C00000 over 500 ms. */
export function triggerColourDeck(): Promise<Uint8Array> {
  const id = ids()
  const recolour: Effect = {
    cls: 'emph', preset: 1, behaviours: id => [
      `<p:animClr clrSpc="rgb"><p:cBhvr>${ctn(id(), 500)}${tgt(2)}${attrs('fillcolor')}</p:cBhvr><p:to><a:srgbClr val="C00000"/></p:to></p:animClr>`,
    ],
  }
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [
      rect(2, 'Box', 380, 200, 200, 200, '336699'),
      rect(3, 'Button', 80, 560, 160, 80, 'F4B183', { paragraphs: [{ runs: [{ text: 'Go' }] }] }),
    ],
    timingXml: timingXmlOf(interactiveSeq(id, 3, [recolour])),
  }])
}

/** `moverReveal`'s 1 s hold and 20 ms dart on the cover (id 3). */
const dartEffect = (): Effect => ({
  cls: 'path', preset: 0, behaviours: id => [
    animNum(id, 3, 'style.opacity', MOVER_REVEAL.stepMs, [[0, 1], [100000, 1]]),
    `<p:animMotion path="M 0 0 L 0.5 0 L 0 0 E"><p:cBhvr>${ctn(id(), MOVER_REVEAL.dartMs, MOVER_REVEAL.dartAt)}${tgt(3)}${attrs('ppt_x', 'ppt_y')}</p:cBhvr></p:animMotion>`,
  ],
})

/** `moverReveal`'s ink, cover and dart, but autoplayed with the slide (the main sequence's group starts at delay 0). */
export function autoplayRevealDeck(): Promise<Uint8Array> {
  const id = ids()
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Ink', 400, 300, 160, 120, HIDDEN_INK), rect(3, 'Cover', 360, 260, 240, 200, '336699')],
    timingXml: timingXmlOf(mainSeqXml([clickStep(id, [dartEffect()], '0')])),
  }])
}

/**
 * One slide: the ink box (id 2) appears on click 1 under an opaque cover (id 3); clicking the cover (an
 * interactive sequence) hides it. Only "click 1, then the trigger" shows the ink.
 */
export function triggerRevealDeck(): Promise<Uint8Array> {
  const id = ids()
  const main = clickStep(id, [appear(2)])
  id() // The main sequence takes the next id (`mainSeqXml`).
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Ink', 400, 300, 160, 120, HIDDEN_INK), rect(3, 'Cover', 360, 260, 240, 200, '336699')],
    timingXml: timingXmlOf(mainSeqXml([main]) + interactiveSeq(id, 3, [disappear(3)])),
  }])
}

/** One slide: two identical opaque boxes (ids 2 and 3, same place and fill) that both appear on click 1. */
export function twinDeck(): Promise<Uint8Array> {
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Twin A', 400, 300, 160, 120, HIDDEN_INK), rect(3, 'Twin B', 400, 300, 160, 120, HIDDEN_INK)],
    timingXml: mainSequence([[appear(2), appear(3)]]),
  }])
}

/**
 * One slide: the ink box (id 2) under an opaque cover (id 3). Click 1 enters the ink and pulses its opacity
 * 0 → 1 → 1 → 0 over 1 s (keys at 0, 25, 75, 100 %); clicking the cover hides it. The ink shows only when the
 * cover is hidden while the pulse is up: a trigger mid-step, or the step played after the trigger.
 */
export function triggerPulseMainDeck(): Promise<Uint8Array> {
  const id = ids()
  const pulse: Effect = {
    cls: 'entr', preset: 10, behaviours: id => [
      setVisibility(id, 2, 'visible'), animNum(id, 2, 'style.opacity', 1000, [[0, 0], [25000, 1], [75000, 1], [100000, 0]]),
    ],
  }
  const main = clickStep(id, [pulse])
  id() // The main sequence takes the next id (`mainSeqXml`).
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Ink', 400, 300, 160, 120, HIDDEN_INK), rect(3, 'Cover', 360, 260, 240, 200, '336699')],
    timingXml: timingXmlOf(mainSeqXml([main]) + interactiveSeq(id, 3, [disappear(3)])),
  }])
}

/**
 * One slide: the ink box (id 2, static) under an opaque cover (id 3) whose click pulses the cover's opacity
 * 1 → 0 → 0 → 1 over 1 s: the ink shows mid-pulse only, never at the trigger's end state.
 */
export function triggerPulseDeck(): Promise<Uint8Array> {
  const id = ids()
  const pulse: Effect = { cls: 'emph', preset: 9, behaviours: id => [animNum(id, 3, 'style.opacity', 1000, [[0, 1], [25000, 0], [75000, 0], [100000, 1]])] }
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Ink', 400, 300, 160, 120, HIDDEN_INK), rect(3, 'Cover', 360, 260, 240, 200, '336699')],
    timingXml: timingXmlOf(interactiveSeq(id, 3, [pulse])),
  }])
}

/**
 * One slide: a trigger button (id 3) below a box (id 2) in z-order, so they share no layer boundary over the
 * background: the button's click fades the box in. Nothing is ever covered, so no layer needs a drop or mask.
 */
export function triggerFadeDeck(): Promise<Uint8Array> {
  const id = ids()
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [
      rect(3, 'Button', 80, 560, 160, 80, 'F4B183'),
      rect(2, 'Box', 380, 200, 200, 200, '336699'),
    ],
    timingXml: timingXmlOf(interactiveSeq(id, 3, [fadeIn(2)])),
  }])
}

/** `moverReveal`'s cover dart, started by a click on a trigger button (id 4) instead of the main sequence. */
export function triggerDartDeck(): Promise<Uint8Array> {
  const id = ids()
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [
      rect(2, 'Ink', 400, 300, 160, 120, HIDDEN_INK), rect(3, 'Cover', 360, 260, 240, 200, '336699'),
      rect(4, 'Button', 80, 560, 160, 80, 'F4B183'),
    ],
    timingXml: timingXmlOf(interactiveSeq(id, 4, [dartEffect()])),
  }])
}

/**
 * One slide: click 1 is a 1 s step (the box's opacity held at 1) holding a 20 ms recolour of the box (id 2)
 * from 300 ms that is not held (`fill="remove"`): no sample of the step shows it tinted.
 */
export function flashColourDeck(): Promise<Uint8Array> {
  return deck([{
    bg: { solid: 'FFFFFF' },
    shapes: [rect(2, 'Box', 380, 260, 200, 200, '336699')],
    timingXml: mainSequence([[{
      cls: 'emph', preset: 1, behaviours: id => [
        animNum(id, 2, 'style.opacity', 1000, [[0, 1], [100000, 1]]),
        `<p:animClr clrSpc="rgb"><p:cBhvr><p:cTn id="${id()}" dur="20" fill="remove"><p:stCondLst><p:cond delay="300"/></p:stCondLst></p:cTn>${tgt(2)}${attrs('fillcolor')}</p:cBhvr>` +
        `<p:to><a:srgbClr val="${HIDDEN_INK}"/></p:to></p:animClr>`,
      ],
    }]]),
  }])
}

export interface ManyLayersOptions {
  /** Click steps the entrances are spread over, round-robin (default 5). */
  steps?: number
}

/**
 * Two slides of `n` small boxes (ids 2 … n + 1, a grid). Slide 0 fades every
 * box in, spread round-robin over `steps` click steps, so it plans about `n`
 * target layers. Slide 1 holds the same ids shifted and enters with a Morph
 * (both sides of the Morph hold `n` objects).
 */
export function manyLayersDeck(n: number, options: ManyLayersOptions = {}): Promise<Uint8Array> {
  const count = Math.max(1, Math.floor(n))
  const steps = Math.max(1, Math.min(count, Math.floor(options.steps ?? 5)))
  const cols = Math.ceil(Math.sqrt(count))
  const cell = Math.floor(Math.min(960, 720) / cols)
  const size = Math.max(2, cell - 2)
  const boxes = (dx: number) => Array.from({ length: count }, (_, i) =>
    rect(2 + i, `Box ${i}`, (i % cols) * cell + dx, Math.floor(i / cols) * cell, size, size, i % 2 ? '336699' : 'F4B183'))
  const plan: Effect[][] = Array.from({ length: steps }, () => [])
  for (let i = 0; i < count; i++) plan[i % steps].push(fadeIn(2 + i, 300))
  return deck([
    { bg: { solid: 'FFFFFF' }, shapes: boxes(0), timingXml: mainSequence(plan) },
    { bg: { solid: 'FFFFFF' }, shapes: boxes(Math.floor(cell / 2)), transition: morphTransition('byObject') },
  ])
}

/**
 * `moverReveal`'s timing (ms from the click): the step lasts `stepMs`; the dart begins at `dartAt` and lasts
 * `dartMs`. The cover (x 360–600) uncovers the ink (x 400–560) once it is ≥ 200 px out, so the ink shows for
 * t in `reveal` only.
 */
export const MOVER_REVEAL = Object.freeze({ stepMs: 1000, dartAt: 300, dartMs: 20, reveal: Object.freeze([300 + 20 * (200 / 960), 320 - 20 * (200 / 960)] as const) })

export type AdversaryKind ='cover' | 'alpha0' | 'recolourBg' | 'fullMask' | 'crop' | 'exitBeforeEntrance' | 'moverReveal'
export const ADVERSARY_KINDS: readonly AdversaryKind[] = ['cover', 'alpha0', 'recolourBg', 'fullMask', 'crop', 'exitBeforeEntrance', 'moverReveal']

/**
 * One slide whose ink in `HIDDEN_INK` sits in its own layer (or under a
 * mover) but is never shown by the composed show, hidden by `kind`:
 *
 * - `cover`: the ink box (id 2) fades in on click 1 under an opaque static
 *   cover (id 3) that contains it.
 * - `alpha0`: the ink box enters on click 1 at opacity 0 and holds 0.
 * - `recolourBg`: click 1 recolours the ink box's fill to the white
 *   background (an emphasis `animClr`, 1 ms) and, 2 ms later, the box
 *   enters: it only ever shows as background.
 * - `fullMask`: the ink box enters on click 1 by a wipe pinned at progress 0
 *   (fully masked).
 * - `crop`: a picture (id 2) whose left half (cropped away by `srcRect`) is
 *   the ink fades in on click 1; the shown right half is #3366CC.
 * - `exitBeforeEntrance`: click 1 runs the ink box's entrance and, listed
 *   after it at the same instant, its exit: the exit wins, so it never shows.
 * - `moverReveal`: the ink box (id 2, static) sits under an opaque mover
 *   (id 3). Click 1 is a 1 s step (the cover's opacity held at 1) holding a
 *   20 ms dart from 300 ms, 480 px right and back (`MOVER_REVEAL`). The ink
 *   shows only while the dart is out (`MOVER_REVEAL.reveal`, about
 *   304–316 ms), strictly between the classifier's samples (8 even times
 *   plus keyframe and behaviour begin/end times) and between 15 fps frames.
 */
export function adversaryDeck(kind: AdversaryKind): Promise<Uint8Array> {
  const ink = (extra: Partial<PptxShapeSpec> = {}) => rect(2, 'Ink', 400, 300, 160, 120, HIDDEN_INK, extra)
  const cover = rect(3, 'Cover', 360, 260, 240, 200, '336699')
  const slide = (shapes: PptxShapeSpec[], steps: Effect[][]): PptxSlideSpec => ({ bg: { solid: 'FFFFFF' }, shapes, timingXml: mainSequence(steps) })
  switch (kind) {
    case 'cover':
      return deck([slide([ink(), cover], [[fadeIn(2)]])])
    case 'alpha0':
      return deck([slide([ink()], [[{ cls: 'entr', preset: 10, behaviours: id => [setVisibility(id, 2, 'visible'), animNum(id, 2, 'style.opacity', 500, [[0, 0], [100000, 0]])] }]])])
    case 'recolourBg':
      return deck([slide([ink()], [[
        {
          cls: 'emph', preset: 1, behaviours: id => [
            `<p:animClr clrSpc="rgb"><p:cBhvr>${ctn(id(), 1)}${tgt(2)}${attrs('fillcolor')}</p:cBhvr>` +
            `<p:from><a:srgbClr val="${HIDDEN_INK}"/></p:from><p:to><a:srgbClr val="FFFFFF"/></p:to></p:animClr>`,
          ],
        },
        { cls: 'entr', preset: 1, behaviours: id => [setVisibility(id, 2, 'visible', 2)] },
      ]])])
    case 'fullMask':
      return deck([slide([ink()], [[{ cls: 'entr', preset: 22, behaviours: id => [setVisibility(id, 2, 'visible'), animEffect(id, 2, 'in', 'wipe(down)', 500, 0)] }]])])
    case 'crop':
      return deck([slide([{ id: 2, name: 'Picture', off: [px(400), px(300)], ext: [px(160), px(120)], image: { data: halfInkPng(), name: 'ink.png', srcRect: { l: 50000 } } }], [[fadeIn(2)]])])
    case 'exitBeforeEntrance':
      return deck([slide([ink()], [[appear(2), disappear(2)]])])
    case 'moverReveal':
      return deck([slide([ink(), cover], [[{
        cls: 'path', preset: 0, behaviours: id => [
          // The step's 1 s interval: the cover's opacity held at 1 (keyframes at 0 and 1000 ms).
          animNum(id, 3, 'style.opacity', MOVER_REVEAL.stepMs, [[0, 1], [100000, 1]]),
          // The dart: 20 ms from 300 ms, out 480 px and back (it begins and ends covering the ink).
          `<p:animMotion path="M 0 0 L 0.5 0 L 0 0 E"><p:cBhvr>${ctn(id(), MOVER_REVEAL.dartMs, MOVER_REVEAL.dartAt)}${tgt(3)}${attrs('ppt_x', 'ppt_y')}</p:cBhvr></p:animMotion>`,
        ],
      }]])])
  }
}

export type MorphAdversaryKind = 'hiddenFrom' | 'sweptCover' | 'coveredPair' | 'movingCover'
export const MORPH_ADVERSARY_KINDS: readonly MorphAdversaryKind[] = ['hiddenFrom', 'sweptCover', 'coveredPair', 'movingCover']

/**
 * Two slides, the second entered by Morph (`byObject`), whose ink in
 * `HIDDEN_INK` is never shown by the composed show (T5 wire audit, fix round
 * 1), hidden by `kind`:
 *
 * - `hiddenFrom`: on slide 1 the ink box `Ink` (id 3) appears and, listed
 *   after it at the same instant, disappears on click 1, so it is hidden at
 *   the Morph's `fromStep` (1). Slide 2's `Ink` is #336699: the pair matches
 *   by name, its from side is hidden. `Box` (id 2) moves.
 * - `sweptCover`: `morphSweepDeck`'s layout (`Under` partly covered by
 *   `Cover`, `Mover` crossing it, so the pair is unproven and composited),
 *   plus an ink box (id 6) wholly under `Cover`, below `Mover`, on both
 *   slides.
 * - `coveredPair`: the ink box `Ink` (id 2) under its opaque `Cover` (id 3),
 *   which contains it, both in place on both slides, while `Box` (id 4)
 *   moves clear of them: a layered Morph whose matched ink pair is under the
 *   cover throughout.
 * - `movingCover`: the ink box `Box` (id 2) under its containing `Cover`
 *   (id 3), both moved by the same (+320, +160) on slide 2: the ink stays
 *   under the cover throughout (the server composites this pair).
 */
export function morphAdversaryDeck(kind: MorphAdversaryKind): Promise<Uint8Array> {
  const t = morphTransition('byObject')
  const W = 'FFFFFF'
  switch (kind) {
    case 'hiddenFrom':
      return deck([
        { bg: { solid: W }, shapes: [rect(2, 'Box', 80, 80, 200, 140, 'D06040'), rect(3, 'Ink', 80, 400, 200, 140, HIDDEN_INK)], timingXml: mainSequence([[appear(3), disappear(3)]]) },
        { bg: { solid: W }, shapes: [rect(2, 'Box', 600, 300, 200, 140, 'D06040'), rect(3, 'Ink', 600, 80, 200, 140, '336699')], transition: t },
      ])
    case 'sweptCover': {
      const shapes = (moverX: number) => [
        rect(4, 'Under', 300, 200, 300, 300, '70AD47'),
        rect(6, 'Ink', 400, 300, 100, 100, HIDDEN_INK),
        rect(2, 'Mover', moverX, 300, 100, 100, 'C00000'),
        rect(5, 'Cover', 380, 280, 140, 140, '1F4E79'),
      ]
      return deck([{ bg: { solid: W }, shapes: shapes(60) }, { bg: { solid: W }, shapes: shapes(760), transition: t }])
    }
    case 'coveredPair': {
      const shapes = (x: number) => [rect(2, 'Ink', 100, 100, 160, 120, HIDDEN_INK), rect(3, 'Cover', 60, 60, 240, 200, '336699'), rect(4, 'Box', x, 500, 160, 120, 'D06040')]
      return deck([{ bg: { solid: W }, shapes: shapes(100) }, { bg: { solid: W }, shapes: shapes(700), transition: t }])
    }
    case 'movingCover': {
      const shapes = (dx: number, dy: number) => [rect(2, 'Box', 100 + dx, 100 + dy, 160, 120, HIDDEN_INK), rect(3, 'Cover', 60 + dx, 60 + dy, 240, 200, '336699')]
      return deck([{ bg: { solid: W }, shapes: shapes(0, 0) }, { bg: { solid: W }, shapes: shapes(320, 160), transition: t }])
    }
  }
}

// ---------------------------------------------------------------------------
// A tiny PNG encoder (stored deflate): the crop adversary's picture
// ---------------------------------------------------------------------------

const CRC = (() => {
  const t = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[n] = c >>> 0
  }
  return t
})()
function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff
  for (const b of bytes) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}
const be32 = (n: number) => [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255]

/** An RGB PNG of `w` × `h` from `pixel(x, y)` → [r, g, b] (one stored deflate block; tiny images only). */
function png(w: number, h: number, pixel: (x: number, y: number) => [number, number, number]): Uint8Array {
  const raw: number[] = []
  for (let y = 0; y < h; y++) {
    raw.push(0)
    for (let x = 0; x < w; x++) raw.push(...pixel(x, y))
  }
  let a = 1, b = 0
  for (const v of raw) { a = (a + v) % 65521; b = (b + a) % 65521 }
  const zlib = [0x78, 0x01, 1, raw.length & 255, raw.length >>> 8, ~raw.length & 255, (~raw.length >>> 8) & 255, ...raw, ...be32(((b << 16) | a) >>> 0)]
  const chunk = (type: string, data: number[]) => {
    const body = new Uint8Array([...[...type].map(c => c.charCodeAt(0)), ...data])
    return [...be32(data.length), ...body, ...be32(crc32(body))]
  }
  return new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    ...chunk('IHDR', [...be32(w), ...be32(h), 8, 2, 0, 0, 0]),
    ...chunk('IDAT', zlib),
    ...chunk('IEND', []),
  ])
}

const hex = (c: string): [number, number, number] => [parseInt(c.slice(0, 2), 16), parseInt(c.slice(2, 4), 16), parseInt(c.slice(4, 6), 16)]
let halfInk: Uint8Array | undefined
/** 8 × 4 px: the left half `HIDDEN_INK`, the right half #3366CC. */
function halfInkPng(): Uint8Array {
  return halfInk ??= png(8, 4, x => (x < 4 ? hex(HIDDEN_INK) : hex('3366CC')))
}

// ---------------------------------------------------------------------------
// Loading (DOM-free)
// ---------------------------------------------------------------------------

/** An unbounded test workspace ledger. */
function openWorkspace(): SlideShowWorkspace {
  let used = 0
  return {
    limit: Number.POSITIVE_INFINITY,
    used: () => used,
    reserveTerminal: () => undefined,
    lease: (bytes, tag) => {
      let held = bytes, live = true
      used += bytes
      const lease: Lease = {
        tag,
        get bytes() { return live ? held : 0 },
        grow(more) { if (!live) return false; held += more; used += more; return true },
        shrink(less) { const n = Math.min(less, held); held -= n; used -= n },
        transfer: () => lease,
        release() { if (!live) return; live = false; used -= held; held = 0 },
      }
      return lease
    },
  }
}

export interface LoadedShowDeck {
  doc: PptxDocument
  index: ShowIndex
  sink: DiagnosticSink
  /** Slide `i`'s show metadata (timing, transition, links, builds), read through the real metadata store. */
  meta(i: number): Promise<ShowSlide>
  /** The charge callback of slide `i`'s metadata handle. */
  charge(i: number): Promise<(bytes: number, tag: string) => boolean>
}

/** Parses `bytes` with the slide appearance enrichment and opens the metadata store over it, as a show does. */
export async function loadShowDeck(bytes: Uint8Array): Promise<LoadedShowDeck> {
  const doc = await parsePptx(await OfficePackage.load(bytes))
  await enrichSlideAppearance(doc)
  const sink = createDiagnosticSink()
  const handle = await openShowPackage(bytes)
  const index = await readShowIndex(doc, handle, sink)
  const store = createSlideMetadataStore({ doc, handle, index, workspace: openWorkspace(), sink, timing: readSlideTiming, resolver: (slide) => shapeResolver(slide, []) })
  return {
    doc, index, sink,
    meta: async (i) => (await store.acquire(i)).slide,
    charge: async (i) => { const h = await store.acquire(i); return (b, t) => h.attach(b, t) },
  }
}
