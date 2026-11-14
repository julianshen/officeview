/**
 * Piece 4 Task 9 — the T5 wire audit (spec §9 item 3, §6; parent §12a.3).
 * For each appearance adversary (`adversaryDeck(kind)`: a layer holds ink,
 * #13F00D, that the composed show never displays), a real harness session
 * plays the slide the way a client would: its `show-plan` bundle, then the
 * `show-track` of every `composited` step. Every transmitted asset (tag 2
 * layers, tag 3 entry frames, tag 4 keyframes) is decoded and must hold no
 * pixel within ΔRGB ≤ 6 of the ink with α > 0.
 *
 * A slide the server refuses (`still`, `show-limits`, `show-budget`,
 * `show-plan-timeout`) plays its piece-3 still: that still is a transmitted
 * asset too and is fetched and audited the same way (controller ruling,
 * 2026-11-15).
 *
 * The primary pass runs with both marks off, so no mark can tint leaked ink
 * out of the signature; a second pass runs with the default marks (fixed
 * visible mark, Layer-1). A positive control proves the detector live on a
 * real wire PNG where the ink is plainly shown.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { decodeRasterBody } from '../../src/stream/raster'
import { decodeShowMessage, SHOW_PLAN_LIMITS, type ShowMessage, type ShowPlanRecord } from '../../src/stream/show-plan'
import { ADVERSARY_KINDS, adversaryDeck, type AdversaryKind } from '../../src/testdata/show-decks'
import { decodePng } from '../src/mark/png'
import type { AuthDecision } from '../src/hooks'
import { Client, disposePool, start } from './harness'
import { inkPixels } from './show-classify-helper'

afterAll(() => disposePool())

type Policy = AuthDecision['policy']
const MARKS_OFF: Policy = { contentPolicy: 'full', visibleMark: 'off', forensicMark: false }
const MARKS_ON: Policy = { contentPolicy: 'full' }

interface Audited { what: string; w: number; h: number; ink: number }
interface Outcome { kind: AdversaryKind; refused?: string; steps: Array<ShowPlanRecord['steps'][number]['eligibility']>; assets: Audited[]; tracks: number }

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

/** Plays slide 0 of `kind` over the wire as a client would and audits every transmitted asset. */
async function play(kind: AdversaryKind, policy: Policy): Promise<Outcome> {
  const s = await session(await adversaryDeck(kind), policy)
  try {
    const out: Outcome = { kind, steps: [], assets: [], tracks: 0 }
    const r = await s.c.showPlan(0, 1, 'full')
    if (r.end) {
      out.refused = r.end.code
      out.assets.push(await stillOf(s.c, 50, 0))
      return out
    }
    const ms: ShowMessage[] = r.messages.map((p, i) => decodeShowMessage(p, { unit: 0, maxDecoded: i === 0 ? SHOW_PLAN_LIMITS.showDecoded : SHOW_PLAN_LIMITS.assetDecoded }))
    const plan = (ms[0] as { plan: ShowPlanRecord }).plan
    out.steps = plan.steps.map((st) => st.eligibility)
    for (const m of ms.slice(1)) {
      if (m.tag !== 2 && m.tag !== 3) throw new Error(`unexpected tag ${m.tag} in a plan bundle`)
      out.assets.push(audit(`tag${m.tag}:${m.assetId}`, m.png))
    }
    for (const [k, st] of plan.steps.entries()) {
      if (st.eligibility !== 'composited') continue
      const t = await s.c.showTrack(0, k, 1, 'full')
      expect(t.end).toBeUndefined()
      expect(t.messages.length).toBe(st.track!.count)
      out.tracks++
      for (const p of t.messages) {
        const m = decodeShowMessage(p, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded })
        if (m.tag !== 4) throw new Error(`unexpected tag ${m.tag} in a track`)
        out.assets.push(audit(`tag4:${k}:${m.assetId}`, m.png))
      }
    }
    return out
  } finally { s.stop() }
}

const summary = (o: Outcome) => ({ kind: o.kind, refused: o.refused, steps: o.steps, assets: o.assets.length, tracks: o.tracks, inked: o.assets.filter((a) => a.ink > 0).map((a) => `${a.what}=${a.ink}`) })

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
      for (const [pass, policy] of [['marks off', MARKS_OFF], ['marks on', MARKS_ON]] as const) {
        const o = await play(kind, policy)
        console.log(`[t5-wire-audit] ${pass}`, JSON.stringify(summary(o)))
        // Coverage: something was transmitted and audited; every composited step's track was fetched.
        expect(o.assets.length).toBeGreaterThan(0)
        if (!o.refused) expect(o.tracks).toBe(o.steps.filter((e) => e === 'composited').length)
        expect({ pass, inked: summary(o).inked }).toEqual({ pass, inked: [] })
      }
    }, 120_000)
  }
})
