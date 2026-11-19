/** Piece 5 Task 4 (spec §4.1, D2, D3, D22, X13): the plan's media records, snd targets and transition sounds. */
import { describe, expect, it } from 'vitest'
import { decodeShowMessage, encodeShowMessage, localTransition, PlanRejected, SHOW_PLAN_MEDIA_LIMITS, type PlanMediaRecord, type ShowPlanRecord } from '../../src/stream/show-plan'
import type { BehaviourNode, CTn, ContainerNode, TimeNode } from '../../src/slideshow/types'

const E = { unit: 0, maxDecoded: 128 << 20 }
const enc = new TextEncoder()
const id = (n: number) => `m${String(n).padStart(21, '0')}`
const ctn = (i: number, extra: Partial<CTn> = {}): CTn => ({ id: i, durMs: 500, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...extra })
const audio = (i: number, target: BehaviourNode['target']): BehaviourNode => ({ kind: 'audio', ctn: ctn(i, { durMs: 'indefinite' }), target, attrNames: [], additive: 'base', accumulate: 'none' } as BehaviourNode)
const par = (i: number, children: TimeNode[], extra: Partial<CTn> = {}): ContainerNode => ({ kind: 'par', ctn: ctn(i, extra), children })

/** A narration on shape 3, an effect sound and a transition sound sharing one record (X13). */
function plan(over: { media?: PlanMediaRecord[]; snd?: string; start?: string; extra?: TimeNode[] } = {}): ShowPlanRecord {
  return {
    v: 1, slide: { index: 0, w: 960, h: 540 }, options: { showAnimation: true, useTimings: true },
    show: {
      transition: { element: 'fade', ns: 'p', attrs: {}, durationMs: 700, advClick: true, sound: { start: { mediaId: over.start ?? id(1), loop: false } } },
      timing: { root: par(1, [audio(2, { spid: '3' }), audio(3, { snd: over.snd ?? id(1) }), ...(over.extra ?? [])], { durMs: 'indefinite', nodeType: 'tmRoot' }), build: [] },
      builds: [],
    },
    media: over.media ?? [
      { mediaId: id(0), kind: 'audio', spid: '3', durationMs: 2000 },
      { mediaId: id(1), kind: 'audio', durationMs: 3000 },
    ],
    layers: [], base: [], steps: [], triggers: [], linkRegions: [], iterate: { keys: [], groups: [] }, entry: 0, diagnostics: [],
  }
}
const raw = (o: unknown) => { const j = enc.encode(JSON.stringify(o)); const out = new Uint8Array(1 + j.length); out[0] = 1; out.set(j, 1); return out }
const pathOf = (f: () => unknown) => { try { f(); return 'accepted' } catch (e) { if (e instanceof PlanRejected) return e.path; throw e } }
const rej = (o: unknown, e: Parameters<typeof decodeShowMessage>[1] = E) => pathOf(() => decodeShowMessage(raw(o), e))
const withMedia = (media: unknown[]) => ({ ...plan(), media })
const rec = (o: Record<string, unknown>) => ({ mediaId: id(9), kind: 'audio', durationMs: 1000, ...o })

describe('show plan media records', () => {
  it('round-trips an audio record set', () => {
    const p = plan()
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan: p }), E)).toEqual({ tag: 1, plan: p })
    const refused = plan({ media: [{ mediaId: id(0), kind: 'audio', spid: '3', durationMs: 0, refused: 'media-unavailable' }, { mediaId: id(1), kind: 'audio', durationMs: 0, refused: 'media-too-large' }] })
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan: refused }), E)).toEqual({ tag: 1, plan: refused })
    // Absent = none (a piece-4 plan, no sound anywhere).
    const none: ShowPlanRecord = { ...plan(), show: { timing: undefined, builds: [] } }
    delete none.media
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan: none }), E)).toEqual({ tag: 1, plan: none })
  })

  it('rejects kind video', () => {
    expect(rej(withMedia([rec({ kind: 'video' })]))).toBe('media[0].kind')
  })

  // Each case appends its records after the two of `plan()` (media[0], media[1]).
  it.each([
    ['media[2].extra', [rec({ extra: 1 })]],
    ['media[2].mediaId', [rec({ mediaId: 'short' })]],
    ['media[2].mediaId', [rec({ mediaId: 'x'.repeat(21) + '!' })]],
    ['media[2].mediaId', [rec({ mediaId: id(0) })]],
    ['media[2].durationMs', [rec({ durationMs: 1.5 })]],
    ['media[2].durationMs', [rec({ durationMs: -1 })]],
    ['media[2].durationMs', [rec({ durationMs: SHOW_PLAN_MEDIA_LIMITS.durationMs + 1 })]],
    ['media[2].refused', [rec({ refused: 'media-timeout' })]],
    ['media[2].refused', [rec({ refused: 'nope' })]],
    ['media[2].spid', [rec({ spid: 3 })]],
    ['media[2].kind', [{ mediaId: id(9), durationMs: 1 }]],
    ['media', Array.from({ length: SHOW_PLAN_MEDIA_LIMITS.records - 1 }, (_, i) => rec({ mediaId: id(100 + i) }))],
  ])('rejects %s', (path, media) => {
    expect(rej(withMedia([...plan().media!, ...media]))).toBe(path)
  })
  it('accepts exactly 64 records', () => {
    expect(rej(withMedia([...plan().media!, ...Array.from({ length: SHOW_PLAN_MEDIA_LIMITS.records - 2 }, (_, i) => rec({ mediaId: id(100 + i) }))]))).toBe('accepted')
  })

  it('rejects an @ placeholder on the client; the server-side decoder takes it', () => {
    const p = plan({ media: [{ mediaId: '@a0.3', kind: 'audio', spid: '3', durationMs: 1 }, { mediaId: '@a1', kind: 'audio', durationMs: 1 }], snd: '@a1', start: '@a1' })
    expect(rej(p)).toBe('media[0].mediaId')
    expect(rej(p, { ...E, placeholders: true })).toBe('accepted')
    expect(() => encodeShowMessage({ tag: 1, plan: p })).toThrow(PlanRejected)
    expect(() => encodeShowMessage({ tag: 1, plan: p }, { placeholders: true })).not.toThrow()
    // A malformed placeholder rejects on the server too.
    expect(rej({ ...p, media: [{ ...p.media![0]!, mediaId: '@a12345' }, p.media![1]] }, { ...E, placeholders: true })).toBe('media[0].mediaId')
  })

  it('rejects a snd target or a transition start naming no record', () => {
    expect(rej(plan({ snd: id(7) }))).toBe('show.timing.root.children[1].target.snd')
    expect(rej(plan({ start: id(7) }))).toBe('show.transition.sound.start.mediaId')
    const cond: TimeNode = par(9, [], { start: [{ delayMs: 0, evt: 'onBegin', target: { snd: id(8) } }] })
    expect(rej(plan({ extra: [cond] }))).toBe('show.timing.root.children[2].ctn.start[0].target.snd')
    // A piece-4 transition sound (a package rId) rejects.
    const p = plan() as unknown as { show: { transition: { sound: unknown } } }
    p.show.transition.sound = { start: { rId: 'rId3', loop: false } }
    expect(rej(p)).toBe('show.transition.sound.start.rId')
    // No records at all: a snd names none.
    const bare = plan(); delete bare.media
    expect(rej(bare)).toBe('show.transition.sound.start.mediaId')
    bare.show = { ...bare.show, transition: { ...bare.show.transition!, sound: { end: true } } }
    expect(rej(bare)).toBe('show.timing.root.children[1].target.snd')
  })

  it('localTransition maps start.mediaId to the local rId slot (X12)', () => {
    const t = plan().show.transition!
    expect(localTransition(t).sound).toEqual({ start: { rId: id(1), loop: false } })
    const plain = { element: 'fade', ns: 'p' as const, attrs: {}, durationMs: 1, advClick: true }
    expect(localTransition(plain)).toBe(plain)
  })
})

/**
 * Task 11b (owner decision 2026-11-19): a still slide's sound-only plan. `fallback: 'still'` (the refusal the slide
 * plays as) carries the records, the transition (its sound) and the no-input audio schedule; nothing drawn.
 */
describe('a still sound-only plan (Task 11b)', () => {
  const still = (over: Partial<ShowPlanRecord> = {}): ShowPlanRecord => ({ ...plan(), fallback: 'still', ...over })
  it('round-trips with fallback "still" and no layers, steps, base, triggers, links or iterations', () => {
    const p = still()
    expect(decodeShowMessage(encodeShowMessage({ tag: 1, plan: p }), E)).toEqual({ tag: 1, plan: p })
  })
  it.each([
    ['fallback', { fallback: 'show-budget' }],
    ['fallback', { fallback: true }],
    ['layers', { layers: [{ id: 0, bounds: { x: 0, y: 0, w: 10, h: 10 }, z: 0, assetId: 1, decodedBytes: 400, scale: 1 }] }],
    ['steps', { steps: [{ trigger: 'click', eligibility: 'layers' }] }],
    ['base', { base: [['3', { x: 0, y: 0, w: 1, h: 1, rot: 0 }]] }],
    ['triggers', { triggers: [{ spid: '3', rects: [] }] }],
    ['linkRegions', { linkRegions: [{ rect: { x: 0, y: 0, w: 1, h: 1 }, action: { kind: 'withheld' } }] }],
    ['iterate', { iterate: { keys: [['3', ['3/p0']]], groups: [] } }],
  ])('rejects a still plan with %s', (path, over) => {
    expect(rej({ ...still(), ...over })).toBe(path)
  })
  it('rejects a still plan with builds', () => {
    const p = still()
    expect(rej({ ...p, show: { ...p.show, builds: [{ spid: '3', mode: 'whole', ranges: [], animBg: false, body: '' }] } })).toBe('show.builds')
  })
  it('a plan without fallback is unchanged (piece 4 / Task 4 plans)', () => {
    expect(rej(plan())).toBe('accepted')
    expect('fallback' in (decodeShowMessage(encodeShowMessage({ tag: 1, plan: plan() }), E) as { plan: ShowPlanRecord }).plan).toBe(false)
  })
})
