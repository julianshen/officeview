/**
 * S1-8: transition and timing-tree reader, the requirements table and the
 * unknown-node policy (slideshow spec 6.1, 4.4).
 *
 * Builder decks are "Office-shaped synthetic" and are never reported as
 * Office evidence; the corpus tests (LibreOffice sd/qa/unit/data/pptx) are.
 */
import { existsSync, readFileSync } from 'node:fs'
import { resolve as resolvePath } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseXmlOrdered, type XmlNode } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parseArtworkTree, parsePptx } from '../src/pptx/parse'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxDeckSpec, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { rowForPreset } from '../src/slideshow/catalogue'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { targetKey } from '../src/slideshow/target'
import { SLIDESHOW_REQUIREMENTS, slideshowChoiceSupported } from '../src/slideshow/read/requirements'
import { SPD_MS, diagnoseInheritedTransition, readTransition } from '../src/slideshow/read/transition'
import { createSlideTimingReader, expandBuilds, readSlideTiming } from '../src/slideshow/read/timing'
import { shapeResolver } from '../src/slideshow/read/resolve'
import type {
  AnimBehaviour, AnimClrBehaviour, AnimEffectBehaviour, AnimMotionBehaviour, AnimRotBehaviour, AnimScaleBehaviour,
  AudioBehaviour, BehaviourNode, BuildTemplate, CmdBehaviour, ContainerNode, DiagnosticSink, SetBehaviour, TargetResolver,
  TimeNode, TimingTree, VideoBehaviour,
} from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
const sld = (inner: string, ns = '') => parseXmlOrdered(`<p:sld ${P}${ns ? ' ' + ns : ''}><p:cSld><p:spTree/></p:cSld>${inner}</p:sld>`)

/** Every spid resolves on the slide with three level-0 paragraphs. */
const ANY: TargetResolver = {
  shape: () => 'slide',
  paragraphs: () => [{ level: 0, runs: 1, empty: false }, { level: 0, runs: 1, empty: false }, { level: 0, runs: 1, empty: false }],
}
const codes = (sink: DiagnosticSink) => sink.list().map(d => d.code)
const has = (sink: DiagnosticSink, code: string, effect?: string) =>
  sink.list().some(d => d.code === code && (effect === undefined || d.effect === effect))

function tree(result: ReturnType<typeof readSlideTiming>): TimingTree {
  const t = result.timing
  if (!t || 'refused' in t) throw new Error(`timing not read: ${JSON.stringify(t)}`)
  return t
}
function walk(node: TimeNode, visit: (n: TimeNode, depth: number) => void, depth = 1): void {
  visit(node, depth)
  if (node.kind === 'par' || node.kind === 'seq' || node.kind === 'excl') for (const c of node.children) walk(c, visit, depth + 1)
}
function behaviours(t: TimingTree): BehaviourNode[] {
  const out: BehaviourNode[] = []
  walk(t.root, n => { if (n.kind !== 'par' && n.kind !== 'seq' && n.kind !== 'excl') out.push(n as BehaviourNode) })
  return out
}
function byId(t: TimingTree, id: number): TimeNode | undefined {
  let found: TimeNode | undefined
  walk(t.root, n => { if (n.ctn.id === id) found = n })
  return found
}
function mainSeq(t: TimingTree): ContainerNode {
  let found: ContainerNode | undefined
  walk(t.root, n => { if (n.kind === 'seq' && n.ctn.nodeType === 'mainSeq') found = n })
  if (!found) throw new Error('no main sequence')
  return found
}

/** One click step holding `inner` behaviours (ids from 3; the builder numbers the wrappers above them). */
const click = (inner: string, preset = 'presetID="10" presetClass="entr" presetSubtype="0"') =>
  `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="5" ${preset} fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
const tgt = (spid = '2', extra = '') => `<p:tgtEl><p:spTgt spid="${spid}">${extra}</p:spTgt></p:tgtEl>`
const bhvr = (id: number, attrs = '', names: string[] = [], spid = '2', cbhvr = '') =>
  `<p:cBhvr${cbhvr}><p:cTn id="${id}"${/\bdur=/.test(attrs) ? '' : ' dur="500"'} ${attrs}/>${tgt(spid)}${names.length ? `<p:attrNameLst>${names.map(n => `<p:attrName>${n}</p:attrName>`).join('')}</p:attrNameLst>` : ''}</p:cBhvr>`
const timingOf = (steps: string[], build = '') => timingXmlOf(mainSeqXml(steps)).replace('</p:tnLst>', `</p:tnLst>${build}`)

const shape = (id: number, paragraphs: number[] = [], text = 'x'): PptxShapeSpec => ({
  id, prst: 'rect', off: ['0', '0'], ext: ['914400', '914400'],
  ...(paragraphs.length ? { paragraphs: paragraphs.map(() => ({ runs: [{ text }] })) } : {}),
})

async function patchSlide(bytes: Uint8Array, i: number, fn: (xml: string) => string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const path = `ppt/slides/slide${i + 1}.xml`
  zip.file(path, fn(await zip.file(path)!.async('string')))
  return zip.generateAsync({ type: 'uint8array' })
}
async function loadDeck(bytes: Uint8Array, i = 0) {
  const zip = await JSZip.loadAsync(bytes)
  const xml = parseXmlOrdered(await zip.file(`ppt/slides/slide${i + 1}.xml`)!.async('string'))
  const pkg = await OfficePackage.load(bytes)
  const doc = await parsePptx(pkg)
  return { xml, slide: doc.slides[i], doc, pkg, zip }
}
async function deck(spec: PptxDeckSpec, i = 0) { return loadDeck(await buildPptxDeck(spec), i) }

/** Read one behaviour in a builder deck (shape id 2 with three paragraphs). */
async function readOne(inner: string) {
  const d = await deck({ slides: [{ shapes: [shape(2, [0, 0, 0])], timingXml: timingOf([click(inner)]) }] })
  const sink = createDiagnosticSink()
  const result = readSlideTiming(d.xml, 0, shapeResolver(d.slide), sink)
  return { t: tree(result), sink, result }
}

const corpus = resolvePath('corpus') + '/'
const corpusAbsent = !existsSync(corpus + 'lo-tdf104788.pptx')
async function corpusSlide(file: string, i: number) {
  const bytes = new Uint8Array(readFileSync(corpus + file))
  return loadDeck(bytes, i)
}
async function corpusPart(file: string, part: string): Promise<XmlNode> {
  const zip = await JSZip.loadAsync(readFileSync(corpus + file))
  return parseXmlOrdered(await zip.file(part)!.async('string'))
}

// ---------------------------------------------------------------------------
// Requirements and transitions
// ---------------------------------------------------------------------------

describe('requirements table', () => {
  it('names p14, p15 and p159 and every namespace the catalogue uses', () => {
    for (const ns of ['p', 'p14', 'p15', 'p159']) expect(SLIDESHOW_REQUIREMENTS).toContain(ns)
  })

  it('mc:Choice p14 transition is chosen by namespace, not prefix; an aliased prefix works; unknown Requires falls back', () => {
    const sink = createDiagnosticSink()
    const ac = (requires: string, decl: string, choice: string) =>
      sld(`<mc:AlternateContent><mc:Choice ${decl} Requires="${requires}"><p:transition spd="slow" ${choice}/></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`)
    // The prefix "p14" bound to the real URI: chosen.
    const canonical = readTransition(sld(`<mc:AlternateContent><mc:Choice xmlns:p14="${P14}" Requires="p14"><p:transition spd="slow" p14:dur="1200"><p14:vortex dir="r"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`), 0, sink)
    expect(canonical).toMatchObject({ element: 'vortex', ns: 'p14', durationMs: 1200, attrs: { dir: 'r' } })
    // An aliased prefix bound to the p14 URI: chosen just the same.
    const aliased = readTransition(sld(`<mc:AlternateContent><mc:Choice xmlns:q="${P14}" Requires="q"><p:transition q:dur="300"><q:ripple/></p:transition></mc:Choice><mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`), 0, sink)
    expect(aliased).toMatchObject({ element: 'ripple', ns: 'p14', durationMs: 300 })
    // The literal prefix "p14" bound to an unknown URI: falls back, diagnosed.
    const fake = readTransition(ac('p14', 'xmlns:p14="urn:not-powerpoint"', ''), 1, sink)
    expect(fake).toMatchObject({ element: 'fade', ns: 'p', durationMs: SPD_MS.fast })
    expect(has(sink, 'transition-requires', 'p14')).toBe(true)
    // An undeclared Requires prefix: falls back.
    const undeclared = readTransition(ac('zz', '', ''), 2, sink)
    expect(undeclared).toMatchObject({ element: 'fade', ns: 'p' })
    expect(has(sink, 'transition-requires', 'zz')).toBe(true)
    // Choice helper directly.
    const choice = parseXmlOrdered(`<mc:Choice xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="${P14}" Requires="x"/>`)
    expect(slideshowChoiceSupported(choice)).toEqual({ supported: true })
  })

  it('p:extLst transition children are read; an unknown extension URI is diagnosed', () => {
    const sink = createDiagnosticSink()
    const spec = readTransition(sld(`<p:transition spd="med"><p:fade/><p:extLst><p:ext uri="{SOME-GUID}"><p14:prism xmlns:p14="${P14}" isInverted="1"/></p:ext><p:ext uri="{UNKNOWN-URI}"><x:thing xmlns:x="urn:x"/></p:ext></p:extLst></p:transition>`), 0, sink)
    expect(spec).toMatchObject({ element: 'prism', ns: 'p14', attrs: { isInverted: '1' }, durationMs: SPD_MS.med })
    expect(has(sink, 'transition-extension', '{UNKNOWN-URI}')).toBe(true)
    expect(has(sink, 'transition-extension', '{SOME-GUID}')).toBe(false)
  })

  it('spd without dur uses SPD_MS; p14:dur wins; advClick=0 and advTm parse; missing advTm is undefined', () => {
    expect(SPD_MS).toEqual({ fast: 500, med: 750, slow: 1000 })
    const sink = createDiagnosticSink()
    const read = (t: string) => readTransition(sld(t, `xmlns:p14="${P14}"`), 0, sink)
    expect(read('<p:transition spd="slow"><p:wipe dir="d"/></p:transition>')).toMatchObject({ element: 'wipe', ns: 'p', durationMs: 1000, advClick: true, attrs: { dir: 'd' } })
    expect(read('<p:transition spd="med"><p:fade/></p:transition>')!.durationMs).toBe(750)
    expect(read('<p:transition><p:fade/></p:transition>')!.durationMs).toBe(500)
    expect(read('<p:transition spd="slow" p14:dur="700"><p:fade/></p:transition>')!.durationMs).toBe(700)
    const adv = read('<p:transition advClick="0" advTm="2500"><p:cut/></p:transition>')!
    expect(adv.advClick).toBe(false)
    expect(adv.advTmMs).toBe(2500)
    expect(read('<p:transition><p:cut/></p:transition>')!.advTmMs).toBeUndefined()
    // Sound: start with loop, then end.
    const snd = read('<p:transition><p:fade/><p:sndAc><p:stSnd loop="1"><p:snd r:embed="rId9" name="a.wav"/></p:stSnd></p:sndAc></p:transition>')!
    expect(snd.sound).toEqual({ start: { rId: 'rId9', loop: true } })
    expect(read('<p:transition><p:fade/><p:sndAc><p:endSnd/></p:sndAc></p:transition>')!.sound).toEqual({ end: true })
    // p15 preset transitions keep the preset; no transition element is the documented "none".
    expect(read(`<p:transition><p15:prstTrans xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" prst="curtains"/></p:transition>`)).toMatchObject({ element: 'prstTrans', ns: 'p15', preset: 'curtains' })
    expect(read('<p:transition advTm="10"/>')).toMatchObject({ element: 'none', ns: 'p', advTmMs: 10 })
    // No transition at all.
    expect(read('')).toBeUndefined()
  })

  it('an unknown transition element is diagnosed and the fallback branch is used', () => {
    const sink = createDiagnosticSink()
    const spec = readTransition(sld(`<mc:AlternateContent><mc:Choice xmlns:p14="${P14}" Requires="p14"><p:transition><p14:notAThing/></p:transition></mc:Choice><mc:Fallback><p:transition><p:dissolve/></p:transition></mc:Fallback></mc:AlternateContent>`), 0, sink)
    expect(spec).toMatchObject({ element: 'dissolve', ns: 'p' })
    expect(has(sink, 'transition-unknown', 'notAThing')).toBe(true)
  })

  it('a transition on a layout is diagnosed and not inherited', async () => {
    const d = await deck({
      masters: [{ layouts: [{}] }],
      slides: [{ shapes: [shape(2)] }],
    })
    const layoutXml = parseXmlOrdered(`<p:sldLayout ${P}><p:cSld><p:spTree/></p:cSld><p:transition spd="slow"><p:fade/></p:transition></p:sldLayout>`)
    const sink = createDiagnosticSink()
    expect(readTransition(d.xml, 0, sink)).toBeUndefined()
    expect(diagnoseInheritedTransition(layoutXml, 'layout', 0, sink)).toBe(true)
    expect(has(sink, 'transition-inherited', 'layout')).toBe(true)
    const plainMaster = parseXmlOrdered(`<p:sldMaster ${P}><p:cSld><p:spTree/></p:cSld></p:sldMaster>`)
    expect(diagnoseInheritedTransition(plainMaster, 'master', 0, sink)).toBe(false)
    // The slide reader never consults the layout: the slide's own result is still undefined.
    expect(readSlideTiming(d.xml, 0, shapeResolver(d.slide), sink).transition).toBeUndefined()
  })
})

describe.skipIf(corpusAbsent)(`corpus transitions${corpusAbsent ? ' — run bun scripts/fetch-corpus.ts' : ''}`, () => {
  it('tdf115394 transition durations match the file', async () => {
    const sink = createDiagnosticSink()
    // slide1 spd=slow; slide2 spd=med; slide3 no spd; slide4 p14:dur=250; slide5 spd=slow p14:dur=4250.
    const want = [1000, 750, 500, 250, 4250]
    for (let i = 0; i < 5; i++) {
      const xml = await corpusPart('lo-tdf115394.pptx', `ppt/slides/slide${i + 1}.xml`)
      expect(readTransition(xml, i, sink)).toMatchObject({ element: 'fade', ns: 'p', durationMs: want[i], attrs: { thruBlk: '1' } })
    }
  })

  it('tdf142915 reads advClick=0 advTm=1000 with dur 0 and 10', async () => {
    const sink = createDiagnosticSink()
    const one = readTransition(await corpusPart('lo-tdf142915.pptx', 'ppt/slides/slide1.xml'), 0, sink)!
    const two = readTransition(await corpusPart('lo-tdf142915.pptx', 'ppt/slides/slide2.xml'), 1, sink)!
    expect(one).toMatchObject({ advClick: false, advTmMs: 1000, durationMs: 10 })
    expect(two).toMatchObject({ advClick: false, advTmMs: 1000, durationMs: 0 })
  })

  it('p15 airplane, p14 flythrough and p159 morph are chosen over their fallbacks', async () => {
    const sink = createDiagnosticSink()
    expect(readTransition(await corpusPart('lo-airplane-transition.pptx', 'ppt/slides/slide2.xml'), 1, sink))
      .toMatchObject({ element: 'prstTrans', ns: 'p15', preset: 'airplane', durationMs: 1250 })
    expect(readTransition(await corpusPart('lo-tdf156808.pptx', 'ppt/slides/slide1.xml'), 0, sink))
      .toMatchObject({ element: 'flythrough', ns: 'p14', durationMs: 800, attrs: { dir: 'out' } })
    expect(readTransition(await corpusPart('lo-tdf149588_transparentSolidFill.pptx', 'ppt/slides/slide1.xml'), 0, sink))
      .toMatchObject({ element: 'morph', ns: 'p159', durationMs: 1500, attrs: { option: 'byObject' } })
    expect(codes(sink)).not.toContain('transition-requires')
  })

  it('tdf44223 split transition carries its start sound', async () => {
    const sink = createDiagnosticSink()
    const t = readTransition(await corpusPart('lo-tdf44223.pptx', 'ppt/slides/slide2.xml'), 1, sink)
    expect(t).toMatchObject({ element: 'split', ns: 'p', durationMs: 1500, attrs: { orient: 'vert' }, sound: { start: { rId: 'rId2', loop: false } } })
  })
})

// ---------------------------------------------------------------------------
// Timing units and behaviours
// ---------------------------------------------------------------------------

describe('timing tree', () => {
  it('repeatCount 2500 is 2.5 iterations; accel 50000 is 0.5; animRot by 5400000 is 90 degrees; animScale 150000 is 1.5', async () => {
    const { t } = await readOne(
      `<p:animRot by="5400000">${bhvr(6, 'repeatCount="2500" accel="50000" decel="25000"', ['r'])}</p:animRot>` +
      `<p:animScale>${bhvr(7, 'repeatCount="indefinite" repeatDur="3000"')}<p:by x="150000" y="150000"/></p:animScale>`)
    const rot = byId(t, 6) as AnimRotBehaviour
    expect(rot.kind).toBe('animRot')
    expect(rot.ctn.repeatCount).toBe(2.5)
    expect(rot.ctn.accel).toBe(0.5)
    expect(rot.ctn.decel).toBe(0.25)
    expect(rot.by).toBe(90)
    const scale = byId(t, 7) as AnimScaleBehaviour
    expect(scale.by).toEqual({ x: 1.5, y: 1.5 })
    expect(scale.ctn.repeatCount).toBe('indefinite')
    expect(scale.ctn.repeatDurMs).toBe(3000)
  })

  it('every behaviour element of the semantics table parses its fields', async () => {
    // p:set
    let r = await readOne(`<p:set>${bhvr(6, 'fill="hold"', ['style.visibility'], '2', ' additive="sum" accumulate="always" override="childStyle" xfrmType="img" rctx="PPT"')}<p:to><p:strVal val="visible"/></p:to></p:set>`)
    const set = byId(r.t, 6) as SetBehaviour
    expect(set).toMatchObject({ kind: 'set', to: 'visible', attrNames: ['style.visibility'], additive: 'sum', accumulate: 'always', override: 'childStyle', xfrmType: 'img', rctx: 'PPT', target: { spid: '2' } })
    expect(set.ctn).toMatchObject({ id: 6, durMs: 500, fill: 'hold' })
    // p:anim with keyframes, formula and value types
    r = await readOne(`<p:anim calcmode="fmla" valueType="num" from="0" to="1" by="0.5">${bhvr(6, '', ['ppt_x'])}<p:tavLst><p:tav tm="0" fmla="#ppt_x+$"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav><p:tav tm="50000"><p:val><p:fltVal val="0.25"/></p:val></p:tav><p:tav tm="100000"><p:val><p:intVal val="3"/></p:val></p:tav></p:tavLst></p:anim>`)
    const anim = byId(r.t, 6) as AnimBehaviour
    expect(anim).toMatchObject({ kind: 'anim', calcmode: 'fmla', valueType: 'num', from: '0', to: '1', by: '0.5' })
    expect(anim.keys).toEqual([{ tm: 0, value: '#ppt_x', fmla: '#ppt_x+$' }, { tm: 0.5, value: 0.25 }, { tm: 1, value: 3 }])
    // p:animEffect
    r = await readOne(`<p:animEffect transition="out" filter="wipe(down)" prLst="gradientSize: 0.1">${bhvr(6)}<p:progress><p:fltVal val="0.5"/></p:progress></p:animEffect>`)
    expect(byId(r.t, 6)).toMatchObject({ kind: 'animEffect', transition: 'out', filter: 'wipe(down)', prLst: 'gradientSize: 0.1', progress: 0.5 } satisfies Partial<AnimEffectBehaviour>)
    // p:animMotion (no path parser injected: the path is diagnosed, the rest parses)
    r = await readOne(`<p:animMotion origin="layout" path="M 0 0 L 0.25 0.5 E" pathEditMode="relative" rAng="2700000" ptsTypes="AA">${bhvr(6, '', ['ppt_x', 'ppt_y'])}<p:by x="10000" y="20000"/><p:from x="0" y="0"/><p:to x="25000" y="50000"/><p:rCtr x="50000" y="50000"/></p:animMotion>`)
    expect(byId(r.t, 6)).toMatchObject({ kind: 'animMotion', origin: 'layout', rAngDeg: 45, ptsTypes: 'AA', by: { x: 0.1, y: 0.2 }, from: { x: 0, y: 0 }, to: { x: 0.25, y: 0.5 }, rCtr: { x: 0.5, y: 0.5 } } satisfies Partial<AnimMotionBehaviour>)
    expect(has(r.sink, 'motion-path-deferred')).toBe(true)
    // with an injected path parser the path is set
    const withPath = createSlideTimingReader({ parseMotionPath: s => ({ ok: true, path: { segments: [{ cmd: 'M', pts: [0, 0] }], length: s.length } }) })
    const d = await deck({ slides: [{ shapes: [shape(2)], timingXml: timingOf([click(`<p:animMotion path="M 0 0 E">${bhvr(6)}</p:animMotion>`)]) }] })
    const motion = byId(tree(withPath(d.xml, 0, shapeResolver(d.slide), createDiagnosticSink())), 6) as AnimMotionBehaviour
    expect(motion.path?.segments[0]).toEqual({ cmd: 'M', pts: [0, 0] })
    // p:animRot from/to
    r = await readOne(`<p:animRot from="-5400000" to="10800000">${bhvr(6, '', ['r'])}</p:animRot>`)
    expect(byId(r.t, 6)).toMatchObject({ kind: 'animRot', from: -90, to: 180 } satisfies Partial<AnimRotBehaviour>)
    // p:animScale from/to/zoomContents
    r = await readOne(`<p:animScale zoomContents="1">${bhvr(6)}<p:from x="100000" y="100000"/><p:to x="80000" y="100000"/></p:animScale>`)
    expect(byId(r.t, 6)).toMatchObject({ kind: 'animScale', zoomContents: true, from: { x: 1, y: 1 }, to: { x: 0.8, y: 1 } } satisfies Partial<AnimScaleBehaviour>)
    // p:animClr
    r = await readOne(`<p:animClr clrSpc="hsl" dir="ccw">${bhvr(6, '', ['fillcolor'])}<p:by><p:hsl h="10800000" s="0" l="-20000"/></p:by><p:from><a:srgbClr val="FF0000"/></p:from><p:to><a:srgbClr val="00ff00"/></p:to></p:animClr>`)
    expect(byId(r.t, 6)).toMatchObject({ kind: 'animClr', clrSpc: 'hsl', dir: 'ccw', by: 'hsl(180,0,-0.2)', from: '#ff0000', to: '#00ff00' } satisfies Partial<AnimClrBehaviour>)
    // p:cmd
    r = await readOne(`<p:cmd type="call" cmd="playFrom(0.0)">${bhvr(6)}</p:cmd>`)
    expect(byId(r.t, 6)).toMatchObject({ kind: 'cmd', cmdType: 'call', cmd: 'playFrom(0.0)' } satisfies Partial<CmdBehaviour>)
    // p:audio / p:video through p:cMediaNode
    r = await readOne(`<p:audio isNarration="1"><p:cMediaNode vol="50000" numSld="3"><p:cTn id="6" fill="hold"/>${tgt()}</p:cMediaNode></p:audio><p:video fullScrn="1"><p:cMediaNode><p:cTn id="7"/>${tgt()}</p:cMediaNode></p:video>`)
    expect(byId(r.t, 6)).toMatchObject({ kind: 'audio', isNarration: true, attrNames: [], target: { spid: '2' } } satisfies Partial<AudioBehaviour>)
    expect(byId(r.t, 7)).toMatchObject({ kind: 'video', fullScrn: true } satisfies Partial<VideoBehaviour>)
    expect(has(r.sink, 'timing-ignored', 'numSld')).toBe(true)
    // excl container and iterate
    const exclXml = `<p:excl><p:cTn id="6" fill="hold"><p:iterate type="wd" backwards="1"><p:tmPct val="20000"/></p:iterate><p:childTnLst><p:set>${bhvr(7)}<p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:excl>`
    r = await readOne(exclXml)
    const excl = byId(r.t, 6) as ContainerNode
    expect(excl.kind).toBe('excl')
    expect(excl.ctn.iterate).toEqual({ type: 'wd', backwards: true, tmPct: 0.2 })
    // targets
    r = await readOne(
      `<p:set>${bhvr(6).replace(tgt(), tgt('2', '<p:txEl><p:pRg st="1" end="2"/></p:txEl>'))}</p:set>` +
      `<p:set>${bhvr(7).replace(tgt(), tgt('2', '<p:txEl><p:charRg st="0" end="4"/></p:txEl>'))}</p:set>` +
      `<p:set>${bhvr(8).replace(tgt(), tgt('2', '<p:bg/>'))}</p:set>` +
      `<p:set>${bhvr(9).replace(tgt(), '<p:tgtEl><p:sldTgt/></p:tgtEl>')}</p:set>` +
      `<p:set>${bhvr(10).replace(tgt(), '<p:tgtEl><p:sndTgt r:embed="rId4" name="x.wav"/></p:tgtEl>')}</p:set>` +
      `<p:set>${bhvr(11).replace(tgt(), '<p:tgtEl><p:inkTgt spid="2"/></p:tgtEl>')}</p:set>`)
    expect((byId(r.t, 6) as SetBehaviour).target).toEqual({ spid: '2', text: { kind: 'p', st: 1, end: 2 } })
    expect((byId(r.t, 7) as SetBehaviour).target).toEqual({ spid: '2', text: { kind: 'char', st: 0, end: 4 } })
    expect((byId(r.t, 8) as SetBehaviour).target).toEqual({ spid: '2', sub: 'bg' })
    expect((byId(r.t, 9) as SetBehaviour).target).toEqual({ sld: true })
    expect((byId(r.t, 10) as SetBehaviour).target).toEqual({ snd: 'rId4' })
    expect((byId(r.t, 11) as SetBehaviour).target).toEqual({ spid: '2', sub: 'ink' })
  })

  it('container and condition fields follow the semantics table', async () => {
    const d = await deck({ slides: [{ shapes: [shape(2), shape(3)], timingXml: timingOf([click(`<p:set>${bhvr(6)}<p:to><p:strVal val="visible"/></p:to></p:set>`)]) }] })
    const sink = createDiagnosticSink()
    const t = tree(readSlideTiming(d.xml, 0, shapeResolver(d.slide), sink))
    const root = t.root as ContainerNode
    expect(root.kind).toBe('par')
    expect(root.ctn).toMatchObject({ nodeType: 'tmRoot', durMs: 'indefinite', restart: 'never', accel: 0, decel: 0, autoRev: false })
    const seq = mainSeq(t)
    expect(seq.seq).toEqual({ concurrent: true, nextAc: 'seek', prevAc: 'none' })
    const step = seq.children[0] as ContainerNode
    expect(step.ctn.start).toEqual([{ delayMs: 'indefinite' }])
    expect(step.ctn.durMs).toBeUndefined()
    const effect = (step.children[0] as ContainerNode).children[0] as ContainerNode
    expect(effect.ctn).toMatchObject({ id: 5, presetClass: 'entr', presetId: 10, presetSubtype: 0, grpId: 0, nodeType: 'clickEffect', fill: 'hold' })
    expect(effect.ctn.start).toEqual([{ delayMs: 0 }])
  })

  it('unknown element is dropped with siblings kept; unknown container drops its subtree', async () => {
    const { t, sink } = await readOne(
      `<p:set>${bhvr(6)}<p:to><p:strVal val="visible"/></p:to></p:set>` +
      `<p:mystery foo="1"/>` +
      `<p:wrapper><p:cTn id="20"><p:childTnLst><p:set>${bhvr(21)}</p:set></p:childTnLst></p:cTn></p:wrapper>` +
      `<p:animEffect transition="in" filter="fade">${bhvr(7)}<p:bogus/></p:animEffect>`)
    expect(byId(t, 6)).toBeDefined()
    expect(byId(t, 7)).toBeDefined()
    expect(byId(t, 20)).toBeUndefined()
    expect(byId(t, 21)).toBeUndefined()
    expect(has(sink, 'timing-unknown-element', 'mystery')).toBe(true)
    expect(has(sink, 'timing-unknown-element', 'bogus')).toBe(true)
    expect(has(sink, 'timing-unknown-container', 'wrapper')).toBe(true)
  })

  it('an unknown attribute value is dropped with a diagnostic and the node kept', async () => {
    const { t, sink } = await readOne(`<p:set>${bhvr(6, 'fill="sideways" restart="sometimes" dur="soon"')}<p:to><p:strVal val="visible"/></p:to></p:set>`)
    const set = byId(t, 6) as SetBehaviour
    expect(set.ctn.fill).toBeUndefined()
    expect(set.ctn.restart).toBeUndefined()
    expect(has(sink, 'timing-invalid-value', 'cTn@fill')).toBe(true)
    expect(has(sink, 'timing-invalid-value', 'cTn@restart')).toBe(true)
  })

  it('syncBehavior, masterRel, xfrmType and rctx are carried on the normalised nodes', async () => {
    const { t, sink } = await readOne(`<p:set>${bhvr(6, 'syncBehavior="locked" masterRel="nextClick"', [], '2', ' xfrmType="pt" rctx="IE"')}<p:to><p:strVal val="visible"/></p:to></p:set>`)
    const set = byId(t, 6) as SetBehaviour
    expect(set.ctn.syncBehavior).toBe('locked')
    expect(set.ctn.masterRel).toBe('nextClick')
    expect(set.xfrmType).toBe('pt')
    expect(set.rctx).toBe('IE')
    for (const a of ['syncBehavior', 'masterRel', 'rctx']) expect(has(sink, 'timing-carried', a)).toBe(true)
    // Once per slide: a second carrier adds nothing.
    const before = sink.list().length
    sink.add({ code: 'timing-carried', slideIndex: 0, effect: 'rctx', message: 'again' })
    expect(sink.list().length).toBe(before)
  })

  it('a tn cycle is broken with a diagnostic', async () => {
    const sink = createDiagnosticSink()
    const xml = sld(timingXmlOf(
      `<p:par><p:cTn id="3"><p:stCondLst><p:cond evt="onBegin" delay="0"><p:tn val="4"/></p:cond></p:stCondLst></p:cTn></p:par>` +
      `<p:par><p:cTn id="4"><p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="3"/></p:cond></p:stCondLst></p:cTn></p:par>` +
      // Legal: 5 ends when 6 ends, 6 begins when 5 begins.
      `<p:par><p:cTn id="5" dur="indefinite"><p:endCondLst><p:cond evt="onEnd" delay="0"><p:tn val="6"/></p:cond></p:endCondLst></p:cTn></p:par>` +
      `<p:par><p:cTn id="6" dur="100"><p:stCondLst><p:cond evt="onBegin" delay="0"><p:tn val="5"/></p:cond></p:stCondLst></p:cTn></p:par>`))
    const t = tree(readSlideTiming(xml, 0, ANY, sink))
    expect(sink.list().filter(d => d.code === 'timing-cycle')).toHaveLength(1)
    const refs = [3, 4].map(id => byId(t, id)!.ctn.start.filter(c => c.tn !== undefined).length)
    expect(refs.reduce((a, b) => a + b)).toBe(1)
    expect(byId(t, 6)!.ctn.start).toEqual([{ delayMs: 0, evt: 'onBegin', tn: 5 }])
    expect(byId(t, 5)!.ctn.end).toEqual([{ delayMs: 0, evt: 'onEnd', tn: 6 }])
  })

  it('a syncbase condition without evt never fires, as in the S1-9 scheduler, and is diagnosed (no cycle edge)', () => {
    // S1-9 (ss/s1-9 instances(): case undefined only fires without tn) ignores
    // a tn reference with no evt, so the two conditions below cannot form a
    // cycle; the reader agrees and diagnoses each one instead.
    const sink = createDiagnosticSink()
    const xml = sld(timingXmlOf(
      `<p:par><p:cTn id="3"><p:stCondLst><p:cond delay="0"><p:tn val="4"/></p:cond></p:stCondLst></p:cTn></p:par>` +
      `<p:par><p:cTn id="4"><p:stCondLst><p:cond delay="0"><p:tn val="3"/></p:cond></p:stCondLst></p:cTn></p:par>`))
    const t = tree(readSlideTiming(xml, 0, ANY, sink))
    expect(has(sink, 'timing-syncbase-no-event', 'tn:3')).toBe(true)
    expect(has(sink, 'timing-syncbase-no-event', 'tn:4')).toBe(true)
    expect(codes(sink)).not.toContain('timing-cycle')
    expect(byId(t, 3)!.ctn.start).toEqual([{ delayMs: 0, tn: 4 }])
    expect(byId(t, 4)!.ctn.start).toEqual([{ delayMs: 0, tn: 3 }])
  })
})

// ---------------------------------------------------------------------------
// Limits
// ---------------------------------------------------------------------------

describe('reader limits', () => {
  const root = (inner: string, rootStart = '') => sld(`<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot">${rootStart}<p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`)
  const pars = (n: number) => Array.from({ length: n }, (_, i) => `<p:par><p:cTn id="${i + 2}"/></p:par>`).join('')
  // fast-xml-parser caps XML nesting at 100 tags (about 32 time nodes), so deep
  // trees are built as XmlNode objects directly (the reader falls back to
  // elementChildren order for nodes without parse order).
  const nested = (depth: number): XmlNode => {
    let ctn: XmlNode = { '@attrs': { id: String(depth) } }
    for (let i = depth - 1; i >= 1; i--) ctn = { '@attrs': i === 1 ? { id: '1', nodeType: 'tmRoot' } : { id: String(i) }, childTnLst: { '@attrs': {}, par: { '@attrs': {}, cTn: ctn } } }
    return { '@attrs': {}, cSld: { '@attrs': {} }, timing: { '@attrs': {}, tnLst: { '@attrs': {}, par: { '@attrs': {}, cTn: ctn } } } }
  }
  const conds = (n: number) => `<p:par><p:cTn id="2"><p:stCondLst>${'<p:cond delay="0"/>'.repeat(n)}</p:stCondLst></p:cTn></p:par>`
  const anim = (id: number, tavs: number) => `<p:anim>${bhvr(id, '', ['ppt_x'])}<p:tavLst>${Array.from({ length: tavs }, (_, i) => `<p:tav tm="${Math.min(100000, i)}"><p:val><p:fltVal val="0"/></p:val></p:tav>`).join('')}</p:tavLst></p:anim>`
  const anims = (tavsPer: number, count: number, extra = 0) => Array.from({ length: count }, (_, i) => anim(i + 2, tavsPer)).join('') + (extra ? anim(count + 2, extra) : '')
  const read = (xml: XmlNode) => { const sink = createDiagnosticSink(); return { r: readSlideTiming(xml, 0, ANY, sink), sink } }

  it('2000 time nodes, depth 64, 10000 conditions, 1000 tav and 20000 tav per slide are accepted', () => {
    for (const xml of [root(pars(1999)), nested(64), root(conds(10000)), root(anims(1000, 1)), root(anims(1000, 20))]) {
      const { r, sink } = read(xml)
      expect(r.timing && 'refused' in r.timing ? r.timing.refused : 'ok').toBe('ok')
      expect(codes(sink)).not.toContain('timing-limit')
    }
    let depth = 0
    walk(tree(read(nested(64)).r).root, (_, d) => { depth = Math.max(depth, d) })
    expect(depth).toBe(64)
  })

  it('2001 time nodes, depth 65, 10001 conditions, 1001 tav and 20001 tav per slide each refuse to static', () => {
    const cases: Array<[XmlNode, string]> = [
      [root(pars(2000)), 'time-nodes'],
      [nested(65), 'nesting'],
      [root(conds(10001)), 'conditions'],
      [root(anims(1001, 1)), 'keyframes-per-behaviour'],
      [root(anims(1000, 20, 1)), 'keyframes'],
    ]
    for (const [xml, why] of cases) {
      const { r, sink } = read(xml)
      expect(r.timing).toEqual({ refused: `limit:${why}` })
      expect(r.builds).toEqual([])
      expect(has(sink, 'timing-limit', why)).toBe(true)
    }
  })

  it('a motion path over 64 KiB refuses to static', () => {
    const big = 'M 0 0 ' + 'L 0.1 0.1 '.repeat(7000)
    expect(big.length).toBeGreaterThan(65536)
    const { r } = read(root(`<p:animMotion path="${big}">${bhvr(2)}</p:animMotion>`))
    expect(r.timing).toEqual({ refused: 'limit:motion-path' })
  })

  it('the transition is still read when the timing is refused', () => {
    const xml = sld(`<p:transition spd="slow"><p:fade/></p:transition><p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>${pars(2000)}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`)
    const r = readSlideTiming(xml, 0, ANY, createDiagnosticSink())
    expect(r.transition).toMatchObject({ element: 'fade', durationMs: 1000 })
    expect(r.timing).toEqual({ refused: 'limit:time-nodes' })
  })
})

// ---------------------------------------------------------------------------
// Target resolution
// ---------------------------------------------------------------------------

describe('targets', () => {
  it('unresolved spid drops the behaviour, keeps the step', async () => {
    const d = await deck({ slides: [{ shapes: [shape(2)], timingXml: timingOf([click(`<p:set>${bhvr(6, '', [], '99')}<p:to><p:strVal val="visible"/></p:to></p:set><p:set>${bhvr(7)}<p:to><p:strVal val="visible"/></p:to></p:set>`), click(`<p:set>${bhvr(8, '', [], '98')}</p:set>`).replace(/id="([345])"/g, (_, n) => `id="${Number(n) + 10}"`)]) }] })
    const sink = createDiagnosticSink()
    const t = tree(readSlideTiming(d.xml, 0, shapeResolver(d.slide), sink))
    expect(byId(t, 6)).toBeUndefined()
    expect(byId(t, 7)).toBeDefined()
    expect(byId(t, 8)).toBeUndefined()
    // Both steps survive, the second now empty.
    expect(mainSeq(t).children).toHaveLength(2)
    expect(byId(t, 15)).toBeDefined()
    expect(has(sink, 'unresolved-target', '99')).toBe(true)
    expect(has(sink, 'unresolved-target', '98')).toBe(true)
  })

  it('an spid present only in a discarded mc:Choice branch is unresolved; a master-only id is master-target', async () => {
    const base = await buildPptxDeck({
      masters: [{ shapes: [{ ...shape(40), name: 'Logo' }], layouts: [{}] }],
      slides: [{ shapes: [shape(2)], timingXml: timingOf([click(
        `<p:set>${bhvr(6, '', [], '50')}</p:set><p:set>${bhvr(7, '', [], '51')}</p:set><p:set>${bhvr(8, '', [], '40')}</p:set><p:set>${bhvr(9)}</p:set>`)]) }],
    })
    const choice = (id: number) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="S${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`
    const bytes = await patchSlide(base, 0, xml => xml.replace('</p:spTree>',
      `<mc:AlternateContent><mc:Choice xmlns:v="urn:unknown-vendor" Requires="v">${choice(50)}</mc:Choice><mc:Fallback>${choice(51)}</mc:Fallback></mc:AlternateContent></p:spTree>`))
    const d = await loadDeck(bytes)
    const underlay = [
      ...await parseArtworkTree(d.pkg, 'ppt/slideMasters/slideMaster1.xml', d.slide.theme!, []),
      ...await parseArtworkTree(d.pkg, 'ppt/slideLayouts/slideLayout1.xml', d.slide.theme!, []),
    ]
    const resolver = shapeResolver(d.slide, underlay)
    expect(resolver.shape('50')).toBe('missing')
    expect(resolver.shape('51')).toBe('slide')
    expect(resolver.shape('40')).toBe('master-only')
    const sink = createDiagnosticSink()
    const t = tree(readSlideTiming(d.xml, 0, resolver, sink))
    expect(byId(t, 6)).toBeUndefined()
    expect(byId(t, 7)).toBeDefined()
    expect(byId(t, 8)).toBeUndefined()
    expect(byId(t, 9)).toBeDefined()
    expect(has(sink, 'unresolved-target', '50')).toBe(true)
    expect(has(sink, 'master-target', '40')).toBe(true)
  })

  it('shapeResolver reports paragraph levels, run counts and emptiness, recursing through groups', async () => {
    const d = await deck({ slides: [{ shapes: [{ ...shape(2), paragraphs: [{ runs: [{ text: 'a' }, { text: 'b' }] }, { runs: [] }] }] }] })
    const r = shapeResolver(d.slide)
    expect(r.shape('2')).toBe('slide')
    expect(r.paragraphs('2')).toEqual([{ level: 0, runs: 2, empty: false }, { level: 0, runs: 0, empty: true }])
    expect(r.paragraphs('nope')).toEqual([])
    const grouped = shapeResolver({ ...d.slide, shapes: [{ ...d.slide.shapes[0], source: { ...d.slide.shapes[0].source!, id: '9' }, children: d.slide.shapes }] })
    expect(grouped.shape('9')).toBe('slide')
    expect(grouped.shape('2')).toBe('slide')
  })
})

// ---------------------------------------------------------------------------
// Builds
// ---------------------------------------------------------------------------

describe('builds', () => {
  const resolverOf = (levels: number[], empty: boolean[] = []): TargetResolver => ({
    shape: spid => spid === '5' ? 'slide' : 'missing',
    paragraphs: spid => spid === '5' ? levels.map((level, i) => ({ level, runs: 1, empty: empty[i] ?? false })) : [],
  })
  const bldP = (o: Partial<Extract<BuildTemplate, { kind: 'bldP' }>> = {}): BuildTemplate => ({ kind: 'bldP', spid: '5', grpId: 0, build: 'p', bldLvl: 1, animBg: false, rev: false, ...o })
  const range = (st: number, end: number) => ({ key: targetKey({ spid: '5', text: { kind: 'p', st, end } }), st, end })

  it('identical shapes with bldLvl 1 vs 2 and animBg 0 vs 1 expand to different BuildOwnership', () => {
    const r = resolverOf([0, 1, 2, 0])
    const lvl1 = expandBuilds([bldP()], r)
    const lvl2 = expandBuilds([bldP({ bldLvl: 2 })], r)
    expect(lvl1).toEqual([{ spid: '5', mode: 'paragraph', ranges: [range(0, 2), range(3, 3)], animBg: false, body: targetKey({ spid: '5' }) }])
    expect(lvl2[0].ranges).toEqual([range(0, 0), range(1, 2), range(3, 3)])
    const bg = expandBuilds([bldP({ animBg: true })], r)
    expect(bg[0].animBg).toBe(true)
    expect(bg[0]).not.toEqual(lvl1[0])
    expect(bg[0].body).toBe('sp:5')
  })

  it('two shapes with equal paragraph and run counts but levels [0,1,0] vs [0,0,0] expand to different ranges for a top-level build', () => {
    expect(expandBuilds([bldP()], resolverOf([0, 1, 0]))[0].ranges).toEqual([range(0, 1), range(2, 2)])
    expect(expandBuilds([bldP()], resolverOf([0, 0, 0]))[0].ranges).toEqual([range(0, 0), range(1, 1), range(2, 2)])
  })

  it('empty paragraphs join the preceding range; whole builds have no ranges; unresolved builds are diagnosed', () => {
    expect(expandBuilds([bldP()], resolverOf([0, 0, 0], [false, true, false]))[0].ranges).toEqual([range(0, 1), range(2, 2)])
    expect(expandBuilds([bldP({ build: 'whole' })], resolverOf([0, 0]))).toEqual([{ spid: '5', mode: 'whole', ranges: [], animBg: false, body: 'sp:5' }])
    const sink = createDiagnosticSink()
    expect(expandBuilds([bldP({ spid: '77' })], resolverOf([0]), { sink, slideIndex: 0 })).toEqual([])
    expect(has(sink, 'unresolved-target', '77')).toBe(true)
  })

  it('bldLst templates are read with schema defaults', async () => {
    const d = await deck({ slides: [{ shapes: [shape(5, [0, 0]), shape(6)], timingXml: timingOf([click(`<p:set>${bhvr(6, '', [], '5')}</p:set>`)],
      '<p:bldLst><p:bldP spid="5" grpId="0" build="p" animBg="1"/><p:bldP spid="6" grpId="1"/><p:bldGraphic spid="6" grpId="2"><p:bldAsOne/></p:bldGraphic></p:bldLst>') }] })
    const r = readSlideTiming(d.xml, 0, shapeResolver(d.slide), createDiagnosticSink())
    expect(tree(r).build).toEqual([
      { kind: 'bldP', spid: '5', grpId: 0, build: 'p', bldLvl: 1, animBg: true, rev: false },
      { kind: 'bldP', spid: '6', grpId: 1, build: 'whole', bldLvl: 1, animBg: false, rev: false },
      { kind: 'bldGraphic', spid: '6', grpId: 2, bld: 'asOne' },
    ])
    expect(r.builds[0]).toEqual({ spid: '5', mode: 'paragraph', ranges: [range(0, 0), range(1, 1)], animBg: true, body: 'sp:5' })
  })

  // S1-7's store and S1-12's buildRenderPlan are not on this branch; the S1-10
  // integration adds the end-to-end run. The builds above already pin the S1-12
  // contract (inclusive st..end, body = targetKey({spid}), keys via targetKey).
  it.todo('reader to store to planner: builds survive in ShowSlide.builds and drive buildRenderPlan')
})

// ---------------------------------------------------------------------------
// Corpus timing
// ---------------------------------------------------------------------------

describe.skipIf(corpusAbsent)(`corpus timing${corpusAbsent ? ' — run bun scripts/fetch-corpus.ts' : ''}`, () => {
  it('tdf104788 timing tree has six entrance steps with the file presets', async () => {
    const d = await corpusSlide('lo-tdf104788.pptx', 5)
    const sink = createDiagnosticSink()
    const r = readSlideTiming(d.xml, 5, shapeResolver(d.slide), sink)
    const t = tree(r)
    const steps = mainSeq(t).children as ContainerNode[]
    expect(steps).toHaveLength(6)
    const presets = steps.map(s => ((s.children[0] as ContainerNode).children[0] as ContainerNode).ctn)
    expect(presets.map(c => c.presetId)).toEqual([37, 34, 58, 47, 42, 10])
    for (const c of presets) {
      expect(c.presetClass).toBe('entr')
      expect(c.nodeType).toBe('clickEffect')
      expect(rowForPreset('entr', c.presetId!, c.presetSubtype)).toBeDefined()
    }
    for (const s of steps) expect(s.ctn.start).toEqual([{ delayMs: 'indefinite' }])
    expect(behaviours(t).length).toBeGreaterThanOrEqual(19)
    expect(t.build).toEqual([{ kind: 'bldP', spid: '2', grpId: 0, build: 'whole', bldLvl: 1, animBg: false, rev: false }])
    expect(codes(sink)).not.toContain('unresolved-target')
    expect(codes(sink)).not.toContain('timing-unknown-element')
  })

  it('tdf44223 keeps the step sound from p:subTnLst and the animBg paragraph build', async () => {
    const d = await corpusSlide('lo-tdf44223.pptx', 0)
    const sink = createDiagnosticSink()
    const r = readSlideTiming(d.xml, 0, shapeResolver(d.slide), sink)
    const t = tree(r)
    const audio = behaviours(t).find(b => b.kind === 'audio') as AudioBehaviour
    expect(audio.target).toEqual({ snd: 'rId2' })
    expect(audio.ctn.masterRel).toBe('sameClick')
    expect(audio.ctn.start).toEqual([{ delayMs: 0, evt: 'begin', tn: 5 }])
    expect(audio.ctn.end).toEqual([{ delayMs: 0, evt: 'onStopAudio', target: { sld: true } }])
    expect(t.build[0]).toMatchObject({ kind: 'bldP', spid: '4', animBg: true })
    expect(r.builds[0]).toMatchObject({ spid: '4', animBg: true, body: 'sp:4' })
    expect(has(sink, 'timing-ignored', 'display')).toBe(true)
  })

  it('bnc591147 reads the interactive sequence trigger and the media call', async () => {
    const d = await corpusSlide('lo-bnc591147.pptx', 0)
    const sink = createDiagnosticSink()
    const t = tree(readSlideTiming(d.xml, 0, shapeResolver(d.slide), sink))
    let inter: ContainerNode | undefined
    walk(t.root, n => { if (n.kind === 'seq' && n.ctn.nodeType === 'interactiveSeq') inter = n })
    expect(inter!.ctn.start).toEqual([{ delayMs: 0, evt: 'onClick', target: { spid: '4' } }])
    expect(inter!.ctn.endSync).toEqual({ delayMs: 0, evt: 'end', rtn: 'all' })
    expect(inter!.ctn.restart).toBe('whenNotActive')
    const cmds = behaviours(t).filter(b => b.kind === 'cmd') as CmdBehaviour[]
    expect(cmds.map(c => c.cmd)).toEqual(['playFrom(0.0)', 'togglePause'])
    expect(byId(t, 5)!.ctn).toMatchObject({ presetClass: 'mediacall', nodeType: 'afterEffect' })
    expect(has(sink, 'timing-ignored', 'evtFilter')).toBe(true)
  })

  it('every LibreOffice corpus slide reads without unknown nodes', async () => {
    for (const file of ['lo-tdf104788.pptx', 'lo-tdf115394.pptx', 'lo-tdf142915.pptx', 'lo-tdf44223.pptx', 'lo-bnc591147.pptx', 'lo-airplane-transition.pptx', 'lo-tdf156808.pptx', 'lo-tdf149588_transparentSolidFill.pptx', 'lo-tdf104792-smart-art-animation.pptx', 'lo-connector-shape-animations.pptx']) {
      if (!existsSync(corpus + file)) continue
      const zip = await JSZip.loadAsync(readFileSync(corpus + file))
      const parts = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p))
      for (const part of parts) {
        const sink = createDiagnosticSink()
        const r = readSlideTiming(parseXmlOrdered(await zip.file(part)!.async('string')), 0, ANY, sink)
        expect(r.timing && 'refused' in r.timing ? `${file} ${part} ${r.timing.refused}` : 'ok').toBe('ok')
        expect(codes(sink).filter(c => /unknown/.test(c)).map(c => `${file} ${part} ${c} ${sink.list().find(d => d.code === c)?.effect}`)).toEqual([])
      }
    }
  })
})

// ---------------------------------------------------------------------------
// Reader -> scheduler contract (pinned snapshot of ss/s1-9's scheduler)
// ---------------------------------------------------------------------------

describe('reader to scheduler contract', () => {
  it('builder deck timing drives the S1-9 scheduler: next begins the click step, a trigger begins the interactive sequence', async () => {
    const { createScheduler } = await import('./fixtures/slideshow-s1-9/scheduler')
    const interactive = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="20" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>` +
      `<p:par><p:cTn id="21" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="22" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="23" dur="1" fill="hold"/>${tgt('2')}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
      `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`
    const timingXml = timingXmlOf(mainSeqXml([click(`<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="6" dur="1000"/>${tgt('2')}</p:cBhvr></p:animEffect>`)]) + interactive)
    const d = await deck({ slides: [{ shapes: [shape(2), shape(3)], timingXml }] })
    const t = tree(readSlideTiming(d.xml, 0, shapeResolver(d.slide), createDiagnosticSink()))
    const s = createScheduler(t)
    const ids = (snap: { active: ReadonlyArray<{ node: BehaviourNode }> }) => snap.active.map(e => e.node.ctn.id)
    expect(ids(s.evaluate([], 100))).toEqual([])
    const next = [{ kind: 'next' as const, t: 0 }]
    expect(ids(s.evaluate(next, 500))).toEqual([6])
    const trigger = [{ kind: 'trigger' as const, spid: '3', t: 0 }]
    const snap = s.evaluate(trigger, 0.5)
    expect([...ids(snap), ...snap.frozen.map(e => e.node.ctn.id)]).toContain(23)
    expect(ids(s.evaluate(trigger, 500))).not.toContain(6)
  })
})
