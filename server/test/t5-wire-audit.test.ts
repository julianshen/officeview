/**
 * Piece 4 Task 9 — the T5 wire audit (spec §9 item 3, §6; parent §12a.3).
 * Every show asset a client receives must hold no pixel within ΔRGB ≤ 6 of
 * the hidden ink (#13F00D) with α > 0:
 *
 * - **Appearance adversaries** (`adversaryDeck(kind)`, one slide): its
 *   `show-plan` bundle (tag 3 entry frame, tag 2 layers) and the `show-track`
 *   of every `composited` step (tag 4 keyframes).
 * - **Morph adversaries** (`morphAdversaryDeck(kind)`, two slides, the second
 *   entered by Morph; fix round 1): both slides' bundles and tracks, and the
 *   `show-morph` pair (tag 5 record, then tag 2 sides for a layered Morph or
 *   tag 4 frames for a composited one). The path each takes is asserted, so
 *   an empty reply or a refusal cannot pass.
 *
 * Coverage is required: every adversary must actually ship show assets (a
 * show-plan refusal fails the gate), so a run where every slide falls back
 * cannot claim to have exercised animation-asset security.
 *
 * Refusals are recorded separately. A slide the server refuses (`still`,
 * `show-limits`, `show-budget`, `show-plan-timeout`) plays its piece-3
 * still, the existing static view, which the controller's ruling exempts
 * from this invariant: it is fetched and logged, never ink-asserted. (A
 * static still draws shapes at rest, so it can legitimately carry the ink:
 * see the positive control.)
 *
 * The primary pass runs with both marks off, so no mark can tint leaked ink
 * out of the signature; a second pass runs with the default marks (fixed
 * visible mark, Layer-1). A positive control proves the detector live on a
 * real wire PNG where the ink is plainly shown.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { decodeRasterBody } from '../../src/stream/raster'
import { decodeShowMessage, SHOW_PLAN_LIMITS, type MorphPlanRecord, type ShowMessage, type ShowPlanRecord } from '../../src/stream/show-plan'
import { ADVERSARY_KINDS, adversaryDeck, MORPH_ADVERSARY_KINDS, morphAdversaryDeck, type AdversaryKind, type MorphAdversaryKind } from '../../src/testdata/show-decks'
import { decodePng } from '../src/mark/png'
import type { AuthDecision } from '../src/hooks'
import { Client, disposePool, start } from './harness'
import { inkPixels } from './show-classify-helper'

afterAll(() => disposePool())

type Policy = AuthDecision['policy']
const MARKS_OFF: Policy = { contentPolicy: 'full', visibleMark: 'off', forensicMark: false }
const MARKS_ON: Policy = { contentPolicy: 'full' }

interface Audited { what: string; w: number; h: number; ink: number }
type Eligibility = ShowPlanRecord['steps'][number]['eligibility']
interface SlidePlay { unit: number; refused?: string; steps: Eligibility[]; tracks: number; plan?: ShowPlanRecord }
/** `assets`: show assets (ink-asserted). `stills`: piece-3 stills of refused slides (exempt, logged). */
interface Outcome { name: string; slides: SlidePlay[]; assets: Audited[]; stills: Audited[]; morph?: { path: 'layered' | 'composited'; record: MorphPlanRecord; tag2: number; tag4: number } }

async function session(doc: Uint8Array, policy: Policy) {
  const h = await start({ doc, policy }), c = new Client(h.url)
  await c.hello(); await c.ready()
  return { c, stop: () => { c.close(); h.stop() } }
}

const audit = (what: string, png: Uint8Array): Audited => { const d = decodePng(png); return { what, w: d.w, h: d.h, ink: inkPixels(d.rgba) } }

/** The unit's piece-3 still over the wire (a raster unit body). */
async function stillOf(c: Client, reqId: number, unit: number): Promise<Audited> {
  c.unit(reqId, unit, 1)
  const r = await c.unitMessages(reqId)
  expect(r.unit.discriminator).toBe(2)
  return audit(`still:${unit}`, decodeRasterBody(r.unit.body).png)
}

/** Plays slide `unit` as a client would: its bundle and every composited step's track; a refusal fetches the still. */
async function playSlide(c: Client, unit: number, out: Outcome): Promise<SlidePlay> {
  const r = await c.showPlan(unit, 1, 'full')
  if (r.end) {
    out.stills.push(await stillOf(c, 50 + unit, unit))
    return { unit, refused: r.end.code, steps: [], tracks: 0 }
  }
  const ms: ShowMessage[] = r.messages.map((p, i) => decodeShowMessage(p, { unit, maxDecoded: i === 0 ? SHOW_PLAN_LIMITS.showDecoded : SHOW_PLAN_LIMITS.assetDecoded }))
  const plan = (ms[0] as { plan: ShowPlanRecord }).plan
  const sp: SlidePlay = { unit, steps: plan.steps.map((st) => st.eligibility), tracks: 0, plan }
  for (const m of ms.slice(1)) {
    if (m.tag !== 2 && m.tag !== 3) throw new Error(`unexpected tag ${m.tag} in a plan bundle`)
    out.assets.push(audit(`u${unit}:tag${m.tag}:${m.assetId}`, m.png))
  }
  for (const [k, st] of plan.steps.entries()) {
    if (st.eligibility !== 'composited') continue
    const t = await c.showTrack(unit, k, 1, 'full')
    expect(t.end).toBeUndefined()
    expect(t.messages.length).toBe(st.track!.count)
    sp.tracks++
    for (const p of t.messages) {
      const m = decodeShowMessage(p, { unit, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded })
      if (m.tag !== 4) throw new Error(`unexpected tag ${m.tag} in a track`)
      out.assets.push(audit(`u${unit}:tag4:${k}:${m.assetId}`, m.png))
    }
  }
  return sp
}

async function play(kind: AdversaryKind, policy: Policy): Promise<Outcome> {
  const s = await session(await adversaryDeck(kind), policy)
  try {
    const out: Outcome = { name: kind, slides: [], assets: [], stills: [] }
    out.slides.push(await playSlide(s.c, 0, out))
    return out
  } finally { s.stop() }
}

/** Plays both slides of a Morph adversary, then requests the pair (0 → 1 at slide 0's step count) and audits it. */
async function playMorph(kind: MorphAdversaryKind, policy: Policy): Promise<Outcome> {
  const s = await session(await morphAdversaryDeck(kind), policy)
  try {
    const out: Outcome = { name: kind, slides: [], assets: [], stills: [] }
    const units = s.c.manifest!.units.list
    for (const u of [0, 1]) {
      if (units[u]!.anim) out.slides.push(await playSlide(s.c, u, out))
      else out.slides.push({ unit: u, steps: [], tracks: 0 })
    }
    const fromStep = out.slides[0]!.steps.length
    const r = await s.c.showMorph(0, 1, fromStep, 1, 'full')
    if (r.end) throw new Error(`${kind}: show-morph refused ${r.end.code}`)
    const ms = r.messages.map((p) => decodeShowMessage(p, { unit: 1, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded }))
    if (ms[0]?.tag !== 5) throw new Error(`${kind}: the Morph reply does not open with its record`)
    const record = ms[0].morph
    let tag2 = 0, tag4 = 0
    for (const m of ms.slice(1)) {
      if (m.tag === 2) tag2++
      else if (m.tag === 4) tag4++
      else throw new Error(`unexpected tag ${m.tag} in a Morph reply`)
      out.assets.push(audit(`morph:tag${m.tag}:${(m as { assetId: number }).assetId}`, (m as { png: Uint8Array }).png))
    }
    out.morph = { path: record.composited ? 'composited' : 'layered', record, tag2, tag4 }
    return out
  } finally { s.stop() }
}

const summary = (o: Outcome) => ({
  name: o.name, slides: o.slides.map((p) => ({ unit: p.unit, refused: p.refused, steps: p.steps, tracks: p.tracks })),
  assets: o.assets.length, morph: o.morph && { path: o.morph.path, items: o.morph.record.items.length, tag2: o.morph.tag2, tag4: o.morph.tag4 },
  inked: o.assets.filter((a) => a.ink > 0).map((a) => `${a.what}=${a.ink}`),
  stills: o.stills.map((a) => `${a.what}=${a.ink}`),
})
const PASSES = [['marks off', MARKS_OFF], ['marks on', MARKS_ON]] as const

describe('T5 wire audit', () => {
  test('positive control: the detector finds the ink in a wire PNG where it is shown', async () => {
    // The alpha0 slide's piece-3 still draws its shapes at rest, the ink box at full opacity.
    const s = await session(await adversaryDeck('alpha0'), MARKS_OFF)
    try {
      const a = await stillOf(s.c, 40, 0)
      // The box is 160 × 120 slide px at scale 1.
      expect(a.ink).toBeGreaterThan(160 * 120 * 0.9)
    } finally { s.stop() }
  }, 60_000)

  for (const kind of ADVERSARY_KINDS) {
    test(`no transmitted asset carries the hidden ink (${kind})`, async () => {
      for (const [pass, policy] of PASSES) {
        const o = await play(kind, policy)
        console.log(`[t5-wire-audit] ${pass}`, JSON.stringify(summary(o)))
        // Coverage: the slide shipped show assets (a refusal is recorded, and fails the gate here); every composited step's track was fetched.
        expect({ pass, refused: o.slides[0]!.refused }).toEqual({ pass, refused: undefined })
        expect(o.assets.length).toBeGreaterThan(0)
        expect(o.slides[0]!.tracks).toBe(o.slides[0]!.steps.filter((e) => e === 'composited').length)
        expect({ pass, inked: summary(o).inked }).toEqual({ pass, inked: [] })
      }
    }, 120_000)
  }

  /** The path each Morph adversary must take over the wire. */
  const MORPH_PATH: Record<MorphAdversaryKind, 'layered' | 'composited'> = { hiddenFrom: 'layered', sweptCover: 'composited', coveredPair: 'layered', movingCover: 'composited' }
  for (const kind of MORPH_ADVERSARY_KINDS) {
    test(`no transmitted Morph asset carries the hidden ink (${kind})`, async () => {
      for (const [pass, policy] of PASSES) {
        const o = await playMorph(kind, policy)
        console.log(`[t5-wire-audit] ${pass}`, JSON.stringify(summary(o)))
        const m = o.morph!
        expect({ pass, path: m.path }).toEqual({ pass, path: MORPH_PATH[kind] })
        // Coverage: the pair shipped exactly the assets its record names (layered) or its frames (composited).
        if (m.path === 'layered') {
          const named = new Set(m.record.items.flatMap((it) => [it.from?.assetId, it.to?.assetId]).filter((x): x is number => x !== undefined))
          expect(named.size).toBeGreaterThan(0)
          expect({ tag2: m.tag2, tag4: m.tag4 }).toEqual({ tag2: named.size, tag4: 0 })
        } else {
          expect(m.record.items).toEqual([])
          expect(m.tag4).toBe(m.record.composited!.count)
          expect(m.tag4).toBeGreaterThan(0)
          expect(m.tag2).toBe(0)
        }
        if (kind === 'hiddenFrom') {
          // The ink box is hidden at fromStep: no from side for it.
          expect(o.slides[0]!.steps.length).toBe(1)
          expect(m.record.items.flatMap((it) => (it.from?.key ? [it.from.key] : []))).not.toContain('sp:3')
        }
        expect(o.slides.filter((p) => p.refused)).toEqual([])
        expect({ pass, inked: summary(o).inked }).toEqual({ pass, inked: [] })
      }
    }, 180_000)
  }
})
