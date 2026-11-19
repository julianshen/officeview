import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { gzipSync } from 'node:zlib'
import { encodeMessage, Kind, Reassembler } from '../../src/stream/frame'
import { CAPS } from '../../src/stream/remote/session'
import {
  decodeShowMessage, encodeShowMessage, PlanRejected, SHOW_PLAN_LIMITS,
  type MorphPlanRecord, type PlanLayerRecord, type ShowPlanRecord,
} from '../../src/stream/show-plan'
import type { BehaviourNode, CTn, ContainerNode, TimeNode } from '../../src/slideshow/types'
import { buildPptxDeck, mainSeqXml, timingXmlOf } from '../../src/testdata/ooxml-builders'
import { ADVERSARY_KINDS, adversaryDeck, cubeDeck, flyInDeck, linkDeck, loadShowDeck, longStepDeck, morphDeck, triggerDeck } from '../../src/testdata/show-decks'

const MiB = 1 << 20
const E = { unit: 0, maxDecoded: 128 * MiB }
const enc = new TextEncoder()

const ctn = (id: number, extra: Partial<CTn> = {}): CTn => ({ id, durMs: 500, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...extra })
const setVisible = (id: number, spid: string): BehaviourNode => ({ kind: 'set', ctn: ctn(id, { durMs: 1 }), target: { spid }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' })
const flyY = (id: number, spid: string): BehaviourNode => ({
  kind: 'anim', ctn: ctn(id), target: { spid }, attrNames: ['ppt_y'], additive: 'base', accumulate: 'none',
  keys: [{ tm: 0, value: '#ppt_y+1' }, { tm: 100000, value: '#ppt_y' }], calcmode: 'lin', valueType: 'num',
})
const par = (id: number, children: TimeNode[], extra: Partial<CTn> = {}): ContainerNode => ({ kind: 'par', ctn: ctn(id, extra), children })

/** A fly-in click build: the shape of a real `flyInDeck` slide (entrance fly-in on shape 2 over a background layer). */
function flyIn(): ShowPlanRecord {
  const effect = par(5, [setVisible(6, '2'), flyY(7, '2')], { presetClass: 'entr', presetId: 2, presetSubtype: 4, nodeType: 'clickEffect', grpId: 0 })
  const click = par(3, [par(4, [effect])], { durMs: 'indefinite', start: [{ delayMs: 'indefinite', evt: 'onClick' }] })
  const main = { kind: 'seq', ctn: ctn(2, { durMs: 'indefinite', nodeType: 'mainSeq' }), children: [click], seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' } } as ContainerNode
  return {
    v: 1,
    slide: { index: 0, w: 960, h: 540 },
    options: { showAnimation: true, useTimings: true },
    show: {
      transition: { element: 'fade', ns: 'p', attrs: { thruBlk: '1' }, durationMs: 700, advClick: true },
      timing: { root: par(1, [main], { durMs: 'indefinite', nodeType: 'tmRoot' }), build: [{ kind: 'bldP', spid: '2', grpId: 0, build: 'whole', bldLvl: 1, animBg: false, rev: false }] },
      builds: [{ spid: '2', mode: 'whole', ranges: [{ key: '2', st: 0, end: 0 }], animBg: false, body: '2' }],
    },
    layers: [
      { id: 0, part: 'body', bounds: { x: 0, y: 0, w: 960, h: 540 }, z: 0, assetId: 1, decodedBytes: 960 * 540 * 4, scale: 1, background: true },
      { id: 1, key: '2', part: 'fill', bounds: { x: 100, y: 100, w: 300, h: 120 }, z: 1, assetId: 2, decodedBytes: 300 * 120 * 4, scale: 2, maskBox: { x: 100, y: 100, w: 300, h: 120 } },
    ],
    base: [['2', { x: 100, y: 100, w: 300, h: 120, rot: 0 }]],
    steps: [{ trigger: 'click', eligibility: 'layers' }, { trigger: 'after', eligibility: 'composited', reason: 'mark-rest', track: { fps: 15, count: 8 } }],
    triggers: [{ spid: '2', rects: [{ x: 100, y: 100, w: 300, h: 120 }] }],
    linkRegions: [
      { rect: { x: 10, y: 10, w: 50, h: 20 }, action: { kind: 'slide', index: 3 } },
      { rect: { x: 10, y: 40, w: 50, h: 20 }, action: { kind: 'jump', verb: 'nextslide' } },
      { rect: { x: 10, y: 70, w: 50, h: 20 }, action: { kind: 'withheld' } },
    ],
    iterate: { keys: [['3', ['3#0', '3#1']]], groups: [['3', { wd: [['3#0', '3#1']] }]] },
    entry: 3,
    diagnostics: [{ code: 'step-composited', step: 1 }, { code: 'cmd-dropped' }],
  }
}

const morph = (): MorphPlanRecord => ({
  v: 1, from: 0, to: 1, fromStep: 2,
  items: [
    { background: true, from: { assetId: 1, geometry: { x: 0, y: 0, w: 960, h: 540, rot: 0 }, bounds: { x: 0, y: 0, w: 960, h: 540 } }, to: { assetId: 2, geometry: { x: 0, y: 0, w: 960, h: 540, rot: 0 }, bounds: { x: 0, y: 0, w: 960, h: 540 } } },
    {
      under: true,
      from: { assetId: 3, geometry: { x: 10, y: 10, w: 100, h: 50, rot: 0 }, bounds: { x: 0, y: 0, w: 120, h: 70 }, state: { visible: true, opacity: 1, transform: [1, 0, 0, 1, 0, 0] }, part: 'fill', key: 'a' },
      to: { assetId: 4, geometry: { x: 50, y: 60, w: 100, h: 50, rot: 15 }, bounds: { x: 38.5, y: 41.25, w: 124, h: 88 }, key: 'b' },
      path: { from: { x: 10, y: 10, w: 100, h: 50, rot: 0 }, to: { x: 50, y: 60, w: 100, h: 50, rot: 15 } },
      pathStates: { to: { visible: true, opacity: 0.5, transform: [1, 0, 0, 1, 4, 4], mask: { kind: 'wipe', progress: 0.5, params: { dir: 'l', n: 2 } }, color: { fill: Array(20).fill(0) } } },
    },
    { to: { assetId: 5, geometry: { x: 0, y: 0, w: 5, h: 5, rot: 0 }, bounds: { x: -10, y: -10, w: 25, h: 25 } } },
  ],
  over: [6], pairs: [{ fromKey: 'a', toKey: 'b', rule: 'name-kind' }],
})

/** Payload bytes built without the encoder's own validation, so hostile shapes reach the decoder. */
const raw = (tag: number, v: unknown) => rawText(tag, JSON.stringify(v))
const rawText = (tag: number, s: string) => { const b = enc.encode(s), out = new Uint8Array(1 + b.length); out[0] = tag; out.set(b, 1); return out }
// biome-ignore lint: test mutation helper
const mutated = (f: (p: any) => void, base: unknown = flyIn()) => { const p = structuredClone(base); f(p); return raw(1, p) }

function png(w: number, h: number): Uint8Array {
  const b = new Uint8Array(33), v = new DataView(b.buffer)
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]); v.setUint32(8, 13); b.set(enc.encode('IHDR'), 12); v.setUint32(16, w); v.setUint32(20, h); b[24] = 8; b[25] = 6
  return b
}
const assetBytes = (tag: number, unit: number, assetId: number, w: number, h: number, p = png(w, h)) => {
  const out = new Uint8Array(1 + 8 + 12 + p.length), v = new DataView(out.buffer)
  out[0] = tag; v.setUint32(1, unit, true); v.setUint32(5, assetId, true); v.setUint32(9, w, true); v.setUint32(13, h, true); v.setUint32(17, p.length, true); out.set(p, 21)
  return out
}

function pathOf(fn: () => unknown): string {
  expect(fn).toThrow(PlanRejected)
  try { fn() } catch (e) { return (e as PlanRejected).path }
  throw new Error('unreachable')
}

describe('SHOW_PLAN_LIMITS', () => {
  it('carries the D2 caps verbatim', () => {
    expect(SHOW_PLAN_LIMITS).toEqual({ layers: 128, steps: 64, trackFrames: 90, slideKeyframes: 512, rectsPerStep: 32, triggers: 64, links: 64, morphItems: 512, planRecordBytes: 4194304, assetDecoded: 67108864, showDecoded: 134217728, str: 256 })
    expect(Object.isFrozen(SHOW_PLAN_LIMITS)).toBe(true)
  })
})

describe('round trips', () => {
  it('round-trips a flyInDeck-shaped record', () => {
    const plan = flyIn()
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan }), E)).toEqual({ tag: 1, plan })
  })
  it('round-trips a Morph record carrying under, background, over, path and pathStates', () => {
    const m = morph()
    expect(decodeShowMessage(encodeShowMessage({ tag: 5, morph: m }), E)).toEqual({ tag: 5, morph: m })
  })
  it('round-trips a composited Morph record', () => {
    const m: MorphPlanRecord = { v: 1, from: 4, to: 5, fromStep: 0, items: [], over: [], pairs: [], composited: { fps: 15, count: 90 } }
    expect(decodeShowMessage(encodeShowMessage({ tag: 5, morph: m }), E)).toEqual({ tag: 5, morph: m })
  })
  it.each([2, 3, 4] as const)('round-trips an asset of tag %i with its unit', (tag) => {
    const p = png(30, 20)
    const msg = { tag, unit: 7, assetId: 9, w: 30, h: 20, png: p }
    expect(decodeShowMessage(encodeShowMessage(msg), { unit: 7, maxDecoded: 30 * 20 * 4 })).toEqual(msg)
  })
  it('a record at exactly 64 nesting levels and 2000 time nodes is accepted', () => {
    const plan = flyIn()
    let node: ContainerNode = par(2, [])
    const root = node
    for (let i = 3; i <= 64; i++) { const c = par(i, []); node.children.push(c); node = c }
    plan.show.timing = { root, build: [] }
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan }), E)).toEqual({ tag: 1, plan })
    plan.show.timing = { root: par(1, Array.from({ length: 1999 }, (_, i) => setVisible(i + 2, '2'))), build: [] }
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan }), E)).toEqual({ tag: 1, plan })
  })
  it('a transition-only slide (no layers, no timing) is valid', () => {
    const plan = flyIn()
    plan.layers = []; plan.steps = []; plan.base = []; plan.triggers = []; plan.show = { transition: plan.show.transition, timing: undefined, builds: [] }
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan }), E)).toEqual({ tag: 1, plan })
  })
  it('a refused timing round-trips', () => {
    const plan = flyIn()
    plan.show = { timing: { refused: 'too-many-nodes' }, builds: [] }
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan }), E)).toEqual({ tag: 1, plan })
  })
})

/** What the real reader produces (transition, timing tree, builds) is what the codec must carry, byte for byte. */
describe('real reader output', () => {
  const kinds = new Set<string>()
  const walk = (n: TimeNode) => { kinds.add(n.kind); if ('children' in n) n.children.forEach(walk) }
  async function roundTrip(bytes: Uint8Array) {
    const d = await loadShowDeck(bytes)
    let slides = 0
    for (let i = 0; i < d.doc.slides.length; i++) {
      const m = await d.meta(i)
      // Piece 5 (D3): a package rId never ships. Each sound (an effect's `snd` target, the transition's start) names a
      // media record by an opaque id, as the server's build and `bindMedia` produce.
      const ids = new Map<string, string>()
      const idOf = (rId: string) => { let id = ids.get(rId); if (!id) ids.set(rId, id = `m${String(ids.size).padStart(21, '0')}`); return id }
      const bind = (n: TimeNode): TimeNode => {
        const ctn = { ...n.ctn, start: n.ctn.start.map((c) => (c.target && 'snd' in c.target ? { ...c, target: { snd: idOf(c.target.snd) } } : c)) }
        if ('children' in n) return { ...n, ctn, children: n.children.map(bind) } as TimeNode
        return { ...n, ctn, ...('snd' in n.target ? { target: { snd: idOf(n.target.snd) } } : {}) } as TimeNode
      }
      const timing = m.timing && !('refused' in m.timing) ? { ...m.timing, root: bind(m.timing.root) } : m.timing
      const start = m.transition?.sound?.start
      const transition = m.transition && { ...m.transition, ...(m.transition.sound ? { sound: { ...m.transition.sound, ...(start ? { start: { mediaId: idOf(start.rId), loop: start.loop } } : {}) } } : {}) }
      const show = { ...(transition ? { transition } : {}), timing, builds: m.builds } as ShowPlanRecord['show']
      const plan = { ...flyIn(), show, media: [...ids.values()].map((mediaId) => ({ mediaId, kind: 'audio' as const, durationMs: 0, refused: 'media-unavailable' as const })) }
      const back = decodeShowMessage(encodeShowMessage({ tag: 1, plan }), E)
      expect(back).toEqual({ tag: 1, plan: JSON.parse(JSON.stringify(plan)) })
      if (m.timing && !('refused' in m.timing)) walk(m.timing.root)
      slides++
    }
    return slides
  }
  it.each([
    ['flyInDeck', flyInDeck], ['cubeDeck', cubeDeck], ['triggerDeck', triggerDeck], ['linkDeck', linkDeck], ['longStepDeck', longStepDeck],
    ['morphDeck byObject', () => morphDeck({ option: 'byObject' })], ['morphDeck byWord', () => morphDeck({ option: 'byWord', words: 3 })], ['morphDeck byChar', () => morphDeck({ option: 'byChar', words: 3 })],
  ])('%s: every slide round-trips', async (_n, make) => { expect(await roundTrip(await make())).toBeGreaterThan(0) })
  it.each(ADVERSARY_KINDS.map((k) => [k]))('adversaryDeck(%s): every slide round-trips', async (k) => { expect(await roundTrip(await adversaryDeck(k))).toBeGreaterThan(0) })
  it('a motion path, colour, scale, rotation, per-word iterate and text-range target round-trip', async () => {
    const tg = (spid: number, txt = '') => `<p:tgtEl><p:spTgt spid="${spid}">${txt}</p:spTgt></p:tgtEl>`
    const bh = (id: number, dur: number, spid: number, attrs: string[], txt = '') =>
      `<p:cBhvr><p:cTn id="${id}" dur="${dur}" fill="hold"/>${tg(spid, txt)}<p:attrNameLst>${attrs.map((a) => `<p:attrName>${a}</p:attrName>`).join('')}</p:attrNameLst></p:cBhvr>`
    const effect = (id: number, cls: string, inner: string, extra = '') =>
      `<p:par><p:cTn id="${id}" presetID="1" presetClass="${cls}" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>${extra}<p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`
    const step = (id: number, inner: string) => `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
    const words = '<p:txEl><p:pRg st="0" end="0"/></p:txEl>'
    const xml = timingXmlOf(mainSeqXml([
      step(10, [
        effect(12, 'path', `<p:animMotion origin="layout" path="M 0 0 L 0.25 0.1 C 0.3 0.2 0.4 0.2 0.5 0.1 E" pathEditMode="relative" ptsTypes="AAAA">${bh(13, 2000, 3, ['ppt_x', 'ppt_y'])}<p:rCtr x="0.1" y="0"/></p:animMotion>`),
        effect(14, 'emph', `<p:animClr clrSpc="hsl" dir="cw">${bh(15, 500, 3, ['fillcolor'])}<p:by><p:hsl h="0" s="0" l="10000"/></p:by></p:animClr><p:animScale>${bh(16, 500, 3, ['ppt_w'])}<p:by x="150000" y="150000"/></p:animScale><p:animRot by="5400000">${bh(17, 500, 3, ['r'])}</p:animRot>`),
      ].join('')),
      step(20, effect(22, 'entr', `<p:set>${bh(23, 1, 2, ['style.visibility'], words)}<p:to><p:strVal val="visible"/></p:to></p:set>`, '<p:iterate type="wd"><p:tmPct val="10000"/></p:iterate>')),
    ]))
    const bytes = await buildPptxDeck({ slides: [{ bg: { solid: 'FFFFFF' }, shapes: [
      { id: 2, name: 'Text', prst: 'rect', off: ['952500', '952500'], ext: ['4762500', '952500'], fill: 'DDEBF7', paragraphs: [{ runs: [{ text: 'one two three' }] }] },
      { id: 3, name: 'Box', prst: 'rect', off: ['3810000', '2540000'], ext: ['1905000', '1905000'], fill: '336699' },
    ], timingXml: xml }] })
    const d = await loadShowDeck(bytes)
    const m = await d.meta(0)
    expect(m.timing && 'refused' in m.timing ? m.timing.refused : undefined).toBeUndefined()
    const seen = new Set<string>()
    let iterate = false
    const walk2 = (n: TimeNode) => { seen.add(n.kind); iterate ||= !!n.ctn.iterate; if ('children' in n) n.children.forEach(walk2) }
    if (m.timing && !('refused' in m.timing)) walk2(m.timing.root)
    expect([...seen]).toEqual(expect.arrayContaining(['animMotion', 'animClr', 'animScale', 'animRot', 'set']))
    expect(iterate).toBe(true)
    expect(await roundTrip(bytes)).toBe(1)
  })
  it('unsigned timing ids above 2^31, a condition referencing one, and a high grpId round-trip', async () => {
    const hi = 2147483648
    const tg = '<p:tgtEl><p:spTgt spid="3"/></p:tgtEl>'
    const set = (id: number) => `<p:set><p:cBhvr><p:cTn id="${id}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tg}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
    const effect = (id: number, grp: number, cond: string) => `<p:par><p:cTn id="${id}" presetID="1" presetClass="entr" presetSubtype="0" grpId="${grp}" fill="hold" nodeType="clickEffect"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${set(id + 1)}</p:childTnLst></p:cTn></p:par>`
    const step = `<p:par><p:cTn id="${hi}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${hi + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effect(hi + 2, 4294967295, '<p:cond delay="0"/>')}${effect(hi + 4, hi + 5, `<p:cond evt="onEnd" delay="0"><p:tn val="${hi + 2}"/></p:cond>`)}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
    const bytes = await buildPptxDeck({ slides: [{ bg: { solid: 'FFFFFF' }, shapes: [{ id: 3, name: 'Box', prst: 'rect', off: ['952500', '952500'], ext: ['1905000', '1905000'], fill: '336699' }], timingXml: timingXmlOf(mainSeqXml([step])) }] })
    const m = await (await loadShowDeck(bytes)).meta(0)
    expect(m.timing && 'refused' in m.timing ? m.timing.refused : undefined).toBeUndefined()
    const ids: number[] = [], tns: unknown[] = []
    const walk3 = (n: TimeNode) => { ids.push(n.ctn.id); n.ctn.start.forEach((c) => tns.push(c.tn)); if ('children' in n) n.children.forEach(walk3) }
    if (m.timing && !('refused' in m.timing)) walk3(m.timing.root)
    expect(Math.max(...ids)).toBeGreaterThan(2 ** 31)
    expect(tns).toContain(hi + 2)
    expect(await roundTrip(bytes)).toBe(1)
  })
  const corpus = 'corpus'
  const decks = existsSync(corpus) ? readdirSync(corpus).filter((f) => /^lo-.*\.pptx$/.test(f)) : []
  it.skipIf(!decks.length)('every slide of the LibreOffice corpus decks round-trips', async () => {
    let slides = 0
    for (const f of decks) slides += await roundTrip(new Uint8Array(readFileSync(`${corpus}/${f}`)))
    expect(slides).toBeGreaterThan(0)
    // eslint-disable-next-line no-console
    console.info(`OVSP real-reader round trip: ${decks.length} decks, ${slides} slides, node kinds: ${[...kinds].sort().join(',')}`)
  }, 120_000)
})

describe('hostile plan records', () => {
  const rej = (p: Uint8Array, e = E) => pathOf(() => decodeShowMessage(p, e))
  const layerAt = (i: number, over: object = {}) => ({ id: i, bounds: { x: 0, y: 0, w: 10, h: 10 }, z: i, assetId: 100 + i, decodedBytes: 400, scale: 1, ...over })

  it('rejects an empty payload', () => expect(rej(new Uint8Array(0))).toBe('payload'))
  it.each([0, 6, 7, 255])('rejects a tag outside 1-5 (%i)', (t) => expect(rej(new Uint8Array([t, 1, 2, 3]))).toBe('tag'))
  it('rejects 129 layers', () => expect(rej(mutated((p) => { p.layers = Array.from({ length: 129 }, (_, i) => layerAt(i)) }))).toBe('layers'))
  it('accepts exactly 128 layers', () => {
    expect(() => decodeShowMessage(mutated((p) => { p.layers = Array.from({ length: 128 }, (_, i) => layerAt(i)) }), E)).not.toThrow()
  })
  it('rejects 65 steps', () => expect(rej(mutated((p) => { p.steps = Array.from({ length: 65 }, () => ({ trigger: 'click', eligibility: 'layers' })) }))).toBe('steps'))
  it('rejects 91 track frames', () => expect(rej(mutated((p) => { p.steps[1].track.count = 91 }))).toBe('steps[1].track.count'))
  it('rejects 0 track frames', () => expect(rej(mutated((p) => { p.steps[1].track.count = 0 }))).toBe('steps[1].track.count'))
  it('rejects a track at another frame rate', () => expect(rej(mutated((p) => { p.steps[1].track.fps = 30 }))).toBe('steps[1].track.fps'))
  it('rejects more than 512 track frames in a slide', () => {
    expect(rej(mutated((p) => { p.steps = Array.from({ length: 6 }, () => ({ trigger: 'click', eligibility: 'composited', track: { fps: 15, count: 90 } })) }))).toBe('steps')
  })
  it('rejects 33 rects in a step', () => expect(rej(mutated((p) => { p.triggers[0].rects = Array.from({ length: 33 }, () => ({ x: 0, y: 0, w: 1, h: 1 })) }))).toBe('triggers[0].rects'))
  it('accepts a trigger marked cut: true', () => {
    expect(() => decodeShowMessage(mutated((p) => { p.triggers[0].cut = true }), E)).not.toThrow()
  })
  it('rejects a trigger cut other than true', () => {
    expect(rej(mutated((p) => { p.triggers[0].cut = false }))).toBe('triggers[0].cut')
    expect(rej(mutated((p) => { p.triggers[0].cut = 'x' }))).toBe('triggers[0].cut')
  })
  it('rejects an unknown trigger key', () => expect(rej(mutated((p) => { p.triggers[0].extra = 1 }))).toBe('triggers[0].extra'))
  it('rejects 65 triggers', () => expect(rej(mutated((p) => { p.triggers = Array.from({ length: 65 }, (_, i) => ({ spid: String(i), rects: [] })) }))).toBe('triggers'))
  it('rejects 65 link regions', () => {
    expect(rej(mutated((p) => { p.linkRegions = Array.from({ length: 65 }, () => ({ rect: { x: 0, y: 0, w: 1, h: 1 }, action: { kind: 'withheld' } })) }))).toBe('linkRegions')
  })
  it('rejects 129 iterate items', () => {
    expect(rej(mutated((p) => { p.iterate = { keys: [['a', Array.from({ length: 127 }, (_, i) => String(i))], ['b', ['c']]], groups: [] } }))).toBe('iterate')
    expect(rej(mutated((p) => { p.iterate = { keys: [], groups: [['a', { wd: [Array.from({ length: 129 }, (_, i) => String(i))] }]] } }))).toBe('iterate')
  })
  it('accepts exactly 128 iterate items', () => {
    expect(() => decodeShowMessage(mutated((p) => { p.iterate = { keys: [['a', Array.from({ length: 127 }, (_, i) => String(i))]], groups: [] } }), E)).not.toThrow()
  })
  describe('group container records (Task 6 follow-up)', () => {
    const grouped = (p: Record<string, any>) => {
      p.layers.push({ id: 7, key: 'g', z: 1.5, container: true })
      p.layers.push({ id: 8, key: '9', parent: 7, bounds: { x: 0, y: 0, w: 10, h: 10 }, z: 2, assetId: 9, decodedBytes: 400, scale: 1 })
    }
    it('accepts a container and a child naming it', () => {
      const m = decodeShowMessage(mutated(grouped), E)
      expect(m.tag === 1 && m.plan.layers.map(l => l.id)).toEqual([0, 1, 7, 8])
    })
    it('accepts a container nested in a container', () => expect(() => decodeShowMessage(mutated((p) => { grouped(p); p.layers.splice(3, 0, { id: 10, key: 'h', z: 1.6, container: true, parent: 7 }); p.layers[4].parent = 10 }), E)).not.toThrow())
    it('rejects a container carrying an assetId', () => expect(rej(mutated((p) => { grouped(p); p.layers[2].assetId = 20 }))).toBe('layers[2].assetId'))
    it('rejects a container carrying decodedBytes', () => expect(rej(mutated((p) => { grouped(p); p.layers[2].decodedBytes = 4 }))).toBe('layers[2].decodedBytes'))
    it('rejects a container carrying bounds', () => expect(rej(mutated((p) => { grouped(p); p.layers[2].bounds = { x: 0, y: 0, w: 1, h: 1 } }))).toBe('layers[2].bounds'))
    it('rejects a container without a key', () => expect(rej(mutated((p) => { grouped(p); delete p.layers[2].key }))).toBe('layers[2].key'))
    it('rejects container: false', () => expect(rej(mutated((p) => { grouped(p); p.layers[2].container = false }))).toBe('layers[2].container'))
    it('rejects a parent that names no record', () => expect(rej(mutated((p) => { grouped(p); p.layers[3].parent = 99 }))).toBe('layers[3].parent'))
    it('rejects a parent listed later', () => expect(rej(mutated((p) => { grouped(p); p.layers.reverse() }))).toMatch(/^layers\[\d\]\.parent$/))
    it('rejects a parent that is a raster layer', () => expect(rej(mutated((p) => { grouped(p); p.layers[3].parent = 0 }))).toBe('layers[3].parent'))
    it('rejects a self-parent and a cycle', () => {
      expect(rej(mutated((p) => { grouped(p); p.layers[2].parent = 7 }))).toBe('layers[2].parent')
      expect(rej(mutated((p) => { grouped(p); p.layers.splice(3, 0, { id: 10, key: 'h', z: 1.6, container: true, parent: 7 }); p.layers[2].parent = 10 }))).toBe('layers[2].parent')
    })
    it('rejects a duplicate layer id', () => expect(rej(mutated((p) => { grouped(p); p.layers[3].id = 7 }))).toBe('layers[3].id'))
  })
  it('rejects a duplicate assetId', () => expect(rej(mutated((p) => { p.layers[1].assetId = p.layers[0].assetId }))).toBe('layers[1].assetId'))
  it('rejects an entry assetId that collides with a layer', () => expect(rej(mutated((p) => { p.entry = p.layers[0].assetId }))).toBe('entry'))
  it('rejects a layer decodedBytes of 64 MiB + 1', () => expect(rej(mutated((p) => { p.layers[0].decodedBytes = 64 * MiB + 1 }))).toBe('layers[0].decodedBytes'))
  it('rejects a Σ decodedBytes of 128 MiB + 1', () => {
    expect(rej(mutated((p) => { p.layers = [layerAt(0, { decodedBytes: 64 * MiB }), layerAt(1, { decodedBytes: 64 * MiB }), layerAt(2, { decodedBytes: 1 })] }))).toBe('layers.decodedBytes')
  })
  it('rejects a Σ decodedBytes over the grant the caller passes', () => {
    expect(rej(mutated(() => undefined), { unit: 0, maxDecoded: 960 * 540 * 4 })).toBe('layers.decodedBytes')
  })
  it.each([['null', 'null'], ['1e999', 'Infinity'], ['"7"', 'a string']])('rejects a non-finite bound (%s, %s)', (lit) => {
    const text = JSON.stringify(flyIn()).replace('"bounds":{"x":0,', `"bounds":{"x":${lit},`)
    expect(text).toContain(`"x":${lit}`)
    expect(rej(rawText(1, text))).toBe('layers[0].bounds.x')
  })
  it('rejects a bound beyond 5·slide', () => {
    expect(rej(mutated((p) => { p.layers[0].bounds.x = 5 * 960 + 1 }))).toBe('layers[0].bounds.x')
    expect(rej(mutated((p) => { p.layers[0].bounds.y = -4 * 540 - 1 }))).toBe('layers[0].bounds.y')
    expect(rej(mutated((p) => { p.layers[0].bounds.x = 4 * 960; p.layers[0].bounds.w = 960 + 1 }))).toBe('layers[0].bounds.w')
  })
  it('accepts a bound exactly at -4·slide and 5·slide', () => {
    expect(() => decodeShowMessage(mutated((p) => { p.layers[0].bounds = { x: -4 * 960, y: -4 * 540, w: 9 * 960, h: 9 * 540 } }), E)).not.toThrow()
  })
  it('rejects an unknown top-level key', () => expect(rej(mutated((p) => { p.extra = 1 }))).toBe('extra'))
  it('rejects an unknown nested key', () => expect(rej(mutated((p) => { p.layers[0].extra = 1 }))).toBe('layers[0].extra'))
  it('rejects a missing required key', () => expect(rej(mutated((p) => { delete p.entry }))).toBe('entry'))
  it('rejects a __proto__ key', () => {
    expect(rej(rawText(1, JSON.stringify(flyIn()).replace('"v":1,', '"v":1,"__proto__":{"x":1},')))).toBe('__proto__')
  })
  it('rejects an unknown part', () => expect(rej(mutated((p) => { p.layers[0].part = 'shadow' }))).toBe('layers[0].part'))
  it('rejects a 257-char string', () => expect(rej(mutated((p) => { p.layers[0].key = 'k'.repeat(257) }))).toBe('layers[0].key'))
  it('accepts a 256-char string', () => expect(() => decodeShowMessage(mutated((p) => { p.layers[0].key = 'k'.repeat(256) }), E)).not.toThrow())
  it('animation formula strings may reach the reader\'s 1024 bytes, no further', () => {
    const at0 = (p: { show: { timing: { root: { children: Array<{ children: unknown[] }> } } } }, v: string) => { p.show.timing.root.children[0]!.children = [{ ...flyY(9, '2'), keys: [{ tm: 0, value: v }] }] }
    expect(() => decodeShowMessage(mutated((p) => at0(p, 'x'.repeat(1024))), E)).not.toThrow()
    expect(rej(mutated((p) => at0(p, 'x'.repeat(1025))))).toBe('show.timing.root.children[0].children[0].keys[0].value')
  })
  describe('motion-path structure (the reader\'s parseMotionPath invariants)', () => {
    const motion = (segments: unknown[], length = 1) => ({
      kind: 'animMotion', ctn: ctn(2), target: { spid: '2' }, attrNames: ['ppt_x'], additive: 'base', accumulate: 'none', path: { segments, length },
    })
    const withPath = (segments: unknown[]) => mutated((p) => { p.show.timing = { root: par(1, [motion(segments) as unknown as BehaviourNode]), build: [] } })
    const seg = (i: number, leaf: string) => `show.timing.root.children[0].path.segments[${i}].${leaf}`
    const M = { cmd: 'M', pts: [0, 0] }
    it('accepts a well-formed M L C Z path', () => {
      expect(() => decodeShowMessage(withPath([M, { cmd: 'L', pts: [1, 1] }, { cmd: 'C', pts: [1, 2, 3, 4, 5, 6] }, { cmd: 'Z', pts: [] }]), E)).not.toThrow()
    })
    it.each([
      ['M with one operand', [{ cmd: 'M', pts: [0] }], 0],
      ['L with three operands', [M, { cmd: 'L', pts: [1, 1, 1] }], 1],
      ['an incomplete cubic', [M, { cmd: 'C', pts: [1] }], 1],
      ['a cubic with four operands', [M, { cmd: 'C', pts: [1, 2, 3, 4] }], 1],
      ['Z with an operand', [M, { cmd: 'Z', pts: [0, 0] }], 1],
    ])('rejects %s', (_n, segments, i) => expect(rej(withPath(segments))).toBe(seg(i, 'pts')))
    it('rejects a path whose first command is not M', () => {
      expect(rej(withPath([{ cmd: 'L', pts: [1, 1] }, M]))).toBe(seg(0, 'cmd'))
      expect(rej(withPath([{ cmd: 'Z', pts: [] }]))).toBe(seg(0, 'cmd'))
    })
    it('rejects an empty path', () => expect(rej(withPath([]))).toBe('show.timing.root.children[0].path.segments'))
    it('rejects more than 10000 commands', () => {
      expect(rej(withPath([M, ...Array.from({ length: 10000 }, () => ({ cmd: 'L', pts: [1, 1] }))]))).toBe('show.timing.root.children[0].path.segments')
    })
  })
  it('rejects 2001 time nodes', () => {
    expect(rej(mutated((p) => { p.show.timing = { root: par(1, Array.from({ length: 2000 }, (_, i) => setVisible(i + 2, '2'))), build: [] } }))).toBe('show.timing.root.children[1999]')
  })
  it('rejects nesting 65', () => {
    expect(rej(mutated((p) => {
      let node: ContainerNode = par(1, [])
      const root = node
      for (let i = 2; i <= 65; i++) { const c = par(i, []); node.children.push(c); node = c }
      p.show.timing = { root, build: [] }
    }))).toBe(`show.timing.root${'.children[0]'.repeat(64)}`)
  })
  it('rejects more than 20000 behaviour keyframes in a slide', () => {
    const keys = (n: number) => Array.from({ length: n }, (_, i) => ({ tm: i, value: 1 }))
    const beh = (id: number, n: number) => ({ ...flyY(id, '2'), keys: keys(n) })
    expect(rej(mutated((p) => { p.show.timing = { root: par(1, Array.from({ length: 21 }, (_, i) => beh(i + 2, 1000))), build: [] } }))).toBe('show.timing.root.children[20].keys')
    expect(rej(mutated((p) => { p.show.timing = { root: par(1, [beh(2, 1001)]), build: [] } }))).toBe('show.timing.root.children[0].keys')
    expect(() => decodeShowMessage(mutated((p) => { p.show.timing = { root: par(1, Array.from({ length: 20 }, (_, i) => beh(i + 2, 1000))), build: [] } }), E)).not.toThrow()
  })
  it('rejects an unknown timing key and an unknown behaviour kind', () => {
    expect(rej(mutated((p) => { p.show.timing.root.children[0].ctn.extra = 1 }))).toBe('show.timing.root.children[0].ctn.extra')
    expect(rej(mutated((p) => { p.show.timing.root.children[0].kind = 'script' }))).toBe('show.timing.root.children[0].kind')
  })
  it('rejects a plan record of 4 MiB + 1 B', () => {
    const body = JSON.stringify(flyIn()), pad = ' '.repeat(4 * MiB + 1 - body.length)
    const bytes = rawText(1, body + pad)
    expect(bytes.length - 1).toBe(4 * MiB + 1)
    expect(rej(bytes)).toBe('plan')
  })
  it('does not parse JSON before the byte cap (invalid JSON over the cap names the cap, not a syntax error)', () => {
    expect(rej(rawText(1, '{'.repeat(4 * MiB + 1)))).toBe('plan')
  })
  it('rejects malformed JSON and invalid UTF-8', () => {
    expect(rej(rawText(1, '{"v":'))).toBe('plan')
    expect(rej(new Uint8Array([1, 0xff, 0xfe]))).toBe('plan')
    expect(rej(rawText(1, '[]'))).toBe('plan')
  })
  it('rejects v: 2', () => expect(rej(mutated((p) => { p.v = 2 }))).toBe('v'))
  it('rejects a plan for a unit other than expect.unit', () => expect(rej(mutated((p) => { p.slide.index = 4 }))).toBe('slide.index'))
  it('rejects an unknown diagnostic code and an unknown link verb', () => {
    expect(rej(mutated((p) => { p.diagnostics[0].code = 'who-knows' }))).toBe('diagnostics[0].code')
    expect(rej(mutated((p) => { p.linkRegions[1].action.verb = 'rm -rf' }))).toBe('linkRegions[1].action.verb')
  })
  it('rejects a negative or fractional assetId and an assetId over u32', () => {
    expect(rej(mutated((p) => { p.layers[0].assetId = -1 }))).toBe('layers[0].assetId')
    expect(rej(mutated((p) => { p.layers[0].assetId = 1.5 }))).toBe('layers[0].assetId')
    expect(rej(mutated((p) => { p.layers[0].assetId = 2 ** 32 }))).toBe('layers[0].assetId')
  })
  it('encode refuses what decode refuses', () => {
    const plan = flyIn()
    ;(plan.layers[1] as PlanLayerRecord).assetId = (plan.layers[0] as PlanLayerRecord).assetId
    expect(pathOf(() => encodeShowMessage({ tag: 1, plan }))).toBe('layers[1].assetId')
    expect(pathOf(() => encodeShowMessage({ tag: 5, morph: { ...morph(), composited: { fps: 15, count: 3 } } }))).toBe('composited')
  })
})

describe('hostile Morph records', () => {
  const rej = (p: Uint8Array, e = E) => pathOf(() => decodeShowMessage(p, e))
  const mm = (f: (p: any) => void) => { const p = structuredClone(morph()); f(p); return raw(5, p) } // eslint-disable-line @typescript-eslint/no-explicit-any
  const item = (id: number) => ({ from: { assetId: id, geometry: { x: 0, y: 0, w: 1, h: 1, rot: 0 }, bounds: { x: 0, y: 0, w: 1, h: 1 } } })
  it('rejects 513 Morph items', () => expect(rej(mm((p) => { p.items = Array.from({ length: 513 }, (_, i) => item(i + 10)) }))).toBe('items'))
  it('accepts 512 Morph items', () => expect(() => decodeShowMessage(mm((p) => { p.items = Array.from({ length: 512 }, (_, i) => item(i + 10)) }), E)).not.toThrow())
  it.each([0, 91, -1, 1.5])('rejects composited.count %p', (count) => {
    expect(rej(mm((p) => { p.items = []; p.composited = { fps: 15, count } }))).toBe('composited.count')
  })
  it('rejects composited alongside non-empty items', () => expect(rej(mm((p) => { p.composited = { fps: 15, count: 3 } }))).toBe('composited'))
  it('rejects a duplicate assetId across items and over', () => {
    expect(rej(mm((p) => { p.items[2].to.assetId = 1 }))).toBe('items[2].to.assetId')
    expect(rej(mm((p) => { p.over = [6, 6] }))).toBe('over[1]')
    expect(rej(mm((p) => { p.over = [3] }))).toBe('over[0]')
  })
  it('rejects an item with neither side', () => expect(rej(mm((p) => { p.items.push({ under: true }) }))).toBe('items[3]'))
  it('rejects a non-finite or out-of-range geometry', () => {
    expect(rej(mm((p) => { p.items[0].from.geometry.w = 'x' }))).toBe('items[0].from.geometry.w')
    expect(rej(mm((p) => { p.items[0].from.geometry.x = 5 * 8192 + 1 }))).toBe('items[0].from.geometry.x')
  })
  it('rejects a side without its raster bounds, or with non-finite or out-of-range bounds (Task 7 fix round 1)', () => {
    expect(rej(mm((p) => { delete p.items[0].from.bounds }))).toBe('items[0].from.bounds')
    expect(rej(mm((p) => { p.items[1].to.bounds.x = Number.NaN }))).toBe('items[1].to.bounds.x')
    expect(rej(mm((p) => { p.items[1].to.bounds.w = 'x' }))).toBe('items[1].to.bounds.w')
    expect(rej(mm((p) => { p.items[2].to.bounds.x = 5 * 8192 + 1 }))).toBe('items[2].to.bounds.x')
    expect(rej(mm((p) => { p.items[2].to.bounds.w = -1 }))).toBe('items[2].to.bounds.w')
    expect(rej(mm((p) => { p.items[2].to.bounds.x = 4 * 8192; p.items[2].to.bounds.w = 8193 }))).toBe('items[2].to.bounds.w')
    expect(rej(mm((p) => { p.items[0].from.bounds.extra = 1 }))).toBe('items[0].from.bounds.extra')
  })
  it('rejects an unknown rule, key, state key and v: 2', () => {
    expect(rej(mm((p) => { p.pairs[0].rule = 'fuzzy' }))).toBe('pairs[0].rule')
    expect(rej(mm((p) => { p.extra = 1 }))).toBe('extra')
    expect(rej(mm((p) => { p.items[1].from.state.extra = 1 }))).toBe('items[1].from.state.extra')
    expect(rej(mm((p) => { p.v = 2 }))).toBe('v')
  })
  it('rejects a Morph record of 4 MiB + 1 B', () => {
    expect(rej(rawText(5, JSON.stringify(morph()) + ' '.repeat(4 * MiB)))).toBe('morph')
  })
})

describe('hostile assets', () => {
  const rej = (p: Uint8Array, e = E) => pathOf(() => decodeShowMessage(p, e))
  it('rejects an asset with w·h·4 > the decoded bytes the caller admitted', () => {
    expect(rej(assetBytes(2, 0, 1, 100, 100), { unit: 0, maxDecoded: 100 * 100 * 4 - 1 })).toBe('decodedBytes')
    expect(() => decodeShowMessage(assetBytes(2, 0, 1, 100, 100), { unit: 0, maxDecoded: 100 * 100 * 4 })).not.toThrow()
  })
  it('rejects an asset over 64 MiB decoded whatever the caller admits', () => {
    expect(rej(assetBytes(2, 0, 1, 4097, 4096), { unit: 0, maxDecoded: 1 << 30 })).toBe('decodedBytes')
  })
  it('rejects an asset with max(w,h) = 8193', () => {
    expect(rej(assetBytes(2, 0, 1, 8193, 1))).toBe('w')
    expect(rej(assetBytes(2, 0, 1, 1, 8193))).toBe('h')
    expect(() => decodeShowMessage(assetBytes(2, 0, 1, 8192, 1), E)).not.toThrow()
  })
  it('rejects an asset for a unit other than expect.unit', () => {
    expect(rej(assetBytes(3, 5, 1, 10, 10))).toBe('unit')
    expect(rej(assetBytes(4, 1, 1, 10, 10), { unit: 0, maxDecoded: 1000 })).toBe('unit')
  })
  it('rejects zero dimensions, a truncated body, a bad signature and a header that disagrees with the PNG', () => {
    expect(rej(assetBytes(2, 0, 1, 0, 10))).toBe('w')
    expect(rej(assetBytes(2, 0, 1, 10, 0))).toBe('h')
    expect(rej(new Uint8Array([2, 0, 0, 0, 0, 1, 0, 0, 0]))).toBe('body')
    const bad = png(10, 10); bad[1] = 0
    expect(rej(assetBytes(2, 0, 1, 10, 10, bad))).toBe('png')
    expect(rej(assetBytes(2, 0, 1, 10, 10, png(11, 10)))).toBe('png')
  })
  it('rejects an asset whose declared PNG length disagrees with the body', () => {
    const b = assetBytes(2, 0, 1, 10, 10)
    new DataView(b.buffer).setUint32(17, 99, true)
    expect(rej(b)).toBe('body')
  })
})

describe('module boundary', () => {
  it('has no runtime import of prepare/, morph/, pptx/, core/package or render/', () => {
    const src = readFileSync('src/stream/show-plan.ts', 'utf8')
    const runtime = [...src.matchAll(/^import\s+(?!type\b)[^\n]*?from\s+'([^']+)'/gm)].map((m) => m[1])
    expect(runtime.sort()).toEqual(['../slideshow/read/limits', './raster'])
    expect(src).not.toMatch(/import\([^)]*\)/)
  })
})

describe('SHOW_PLAN caps through the reassembler', () => {
  const caps = CAPS[Kind.SHOW_PLAN]!
  const run = async (payload: Uint8Array, decoded = caps[1]) => {
    const fragments = encodeMessage({ ver: 1, kind: Kind.SHOW_PLAN, reqId: 7, payload, compress: (b) => gzipSync(b, { level: 6 }) })
    const r = new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: () => decoded, encodedCap: (k) => CAPS[k]?.[0] ?? 0 })
    let last
    for (const f of fragments) { last = await r.push(f); if (last) break }
    return last
  }
  it('the session caps are 8 MiB encoded and 32 MiB decoded', () => expect(caps).toEqual([8 * MiB, 32 * MiB]))
  it('a message over 8 MiB encoded is refused (encoded-cap)', async () => {
    const noise = new Uint8Array(8 * MiB + 1024)
    for (let i = 0; i < noise.length; i += 65536) crypto.getRandomValues(noise.subarray(i, Math.min(i + 65536, noise.length)))
    expect(await run(noise)).toMatchObject({ ok: false, reason: 'encoded-cap' })
  })
  it('a message over 32 MiB decoded is refused (decoded-cap)', async () => {
    expect(await run(new Uint8Array(32 * MiB + 1))).toMatchObject({ ok: false, reason: 'decoded-cap' })
  })
  it('a message within both caps is delivered', async () => {
    const payload = encodeShowMessage({ tag: 1, plan: flyIn() })
    expect(await run(payload)).toMatchObject({ ok: true, kind: Kind.SHOW_PLAN, reqId: 7 })
  })
})
