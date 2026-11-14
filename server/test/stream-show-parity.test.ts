/**
 * Piece 4 Task 14 — the streamed-show parity gate on the synthetic decks (spec §8, plan D11, D20, D24).
 *
 * `scripts/stream-show-compare.ts` plays each deck locally (`runShow` under jsdom + node-canvas, the visible mark
 * as the plan's watermark layer) and streamed (this live test server with `testOnly.forensicOff`, the client
 * codecs, `createPlanSource`, the shared driver over `createPlanStage`), and compares the frames at the capture
 * plan's `(slide, step, pct)` samples. `installDom` runs before anything imports `startShow` or `PlanStage`: the
 * harness imports both lazily, after it. The corpus run (and its published widening) is the CLI's
 * (`bun scripts/stream-show-compare.ts --deck all`); this file holds the synthetic decks only.
 */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { compareStreamDeck, deckBytes, installDom, measureWidening, uninstallDom, wideningByFamily, type StreamDeckResult } from '../../scripts/stream-show-compare'
import { disposePool } from './harness'
import { TOLERANCE_APPROX } from '../../src/slideshow/catalogue'

const DECKS = ['flyInDeck', 'morphDeck', 'cubeDeck'] as const
const runs = new Map<string, StreamDeckResult>()
/** Each deck's own measured renderer difference (spec §8: per deck), frozen before any parity sample. */
const widening = new Map<string, { meanAbs: number; rowMax: number }>()
let family = { meanAbs: 0, rowMax: 0 }

beforeAll(async () => {
  await installDom()
  // D20: the renderer widening, measured on the decks' own layers and frozen before any parity sample.
  const all = []
  for (const d of [...DECKS, 'morphCompositedDeck']) {
    const rows = await measureWidening(d, await deckBytes(d))
    expect(rows.length).toBeGreaterThan(0)
    widening.set(d, { meanAbs: Math.max(...rows.map((r) => r.meanAbs)), rowMax: Math.max(...rows.map((r) => r.rowMax)) })
    all.push(...rows)
  }
  const latin = wideningByFamily(all).get('latin')!
  family = { meanAbs: latin.meanAbs, rowMax: latin.rowMax }
}, 60_000)

afterAll(() => { disposePool(); uninstallDom() })

async function run(deck: string, marks: 'off' | 'fixed'): Promise<StreamDeckResult> {
  const key = `${deck}|${marks}`
  let r = runs.get(key)
  if (!r) { r = await compareStreamDeck(deck, { marks: { kind: marks }, widening: widening.get(deck)!, family: 'latin' }); runs.set(key, r) }
  return r
}

describe('streamed show parity (synthetic decks)', () => {
  test('the measured renderer widening is within TOLERANCE_APPROX', () => {
    expect(family.meanAbs).toBeLessThanOrEqual(TOLERANCE_APPROX.meanAbs)
    expect(family.rowMax).toBeLessThanOrEqual(TOLERANCE_APPROX.rowMax)
  })

  test("flyInDeck, morphDeck, cubeDeck: structure equal; every sampled frame within the deck's Tolerance", async () => {
    for (const deck of DECKS) {
      for (const marks of ['off', 'fixed'] as const) {
        const r = await run(deck, marks)
        const at = `${deck} [marks ${marks}]`
        expect(r.structure.equal, `${at}: ${JSON.stringify(r.structure)}`).toBe(true)
        expect(r.structure.stills, at).toEqual([])
        const fails = r.outcomes.filter((o) => o.status === 'fail')
        expect(fails.map((o) => `${o.slide}-${o.step}-${o.pct} ${o.kind}: ${o.detail}`), at).toEqual([])
        const compared = r.outcomes.filter((o) => o.status === 'pass')
        expect(compared.length, at).toBeGreaterThanOrEqual(marks === 'off' ? 6 : 3)
        if (marks === 'off') expect(r.outcomes.filter((o) => o.status === 'skipped-mark'), at).toEqual([])
        // D24: under the fixed mark only mid-motion samples (never 0 % or 100 %) are left out.
        for (const o of r.outcomes.filter((x) => x.status === 'skipped-mark')) expect([0, 100], at).not.toContain(o.pct)
        // Not vacuous: within a step that animates, the frames change across its samples on both sides.
        const byStep = new Map<string, Set<string>>()
        for (const o of r.outcomes) { const k = `${o.slide}-${o.step}`; if (!byStep.has(k)) byStep.set(k, new Set()); byStep.get(k)!.add(`${o.hash?.local}|${o.hash?.streamed}`) }
        expect(Math.max(...[...byStep.values()].map((s) => s.size)), at).toBeGreaterThanOrEqual(4)
      }
    }
    const fly = await run('flyInDeck', 'off')
    expect(fly.structure.local).toEqual({ steps: [1], transitions: ['none'], morphPairs: [], clicks: 2 })
    const cube = await run('cubeDeck', 'off')
    expect(cube.structure.streamed.transitions).toEqual(['none', 'p14:prism'])
    expect(cube.outcomes.filter((o) => o.kind === 'transition' && o.status === 'pass').length).toBe(5)
  }, 240_000)

  test('Morph pairs map to the local keys', async () => {
    const r = await run('morphDeck', 'off')
    expect(r.structure.local.morphPairs).toEqual([['sp:2', 'sp:2', 'name-kind']])
    expect(r.structure.streamed.morphPairs).toEqual(r.structure.local.morphPairs)
    const c = await run('morphCompositedDeck', 'off')
    expect(c.structure.streamed.morphPairs.length).toBeGreaterThan(1)
    expect(c.structure.streamed.morphPairs).toEqual(c.structure.local.morphPairs)
  }, 120_000)

  test('composited Morph sampled at its keyframe times within Tolerance', async () => {
    for (const marks of ['off', 'fixed'] as const) {
      const r = await run('morphCompositedDeck', marks)
      expect(r.diagnostics).toContain('1:morph-composited/swept')
      const kf = r.outcomes.filter((o) => o.kind === 'morph-composited')
      expect(kf.length).toBeGreaterThanOrEqual(4)
      expect(kf.filter((o) => o.status === 'fail').map((o) => o.detail)).toEqual([])
      // D11: track frame i (shown mid-interval, (i + ½)/15 s) is compared with local at the time it depicts.
      for (const o of kf.filter((x) => x.pct !== 100)) expect(o.ms - o.streamedMs).toBeCloseTo(1000 / 30, 1)
      // The keyframes are distinct pictures (the Morph moves between them).
      expect(new Set(kf.map((o) => o.hash?.streamed)).size).toBe(kf.length)
    }
  }, 120_000)

  test('both backends measure Calibri Light text identically (the pack alias, no host fallback)', async () => {
    const { nodeCanvas } = await import('../../scripts/stream-show-compare')
    const { serverBackend, showBackend } = await import('../src/backend')
    const b = showBackend(serverBackend())
    const text = 'Standard transition – medium'
    for (const font of ['58.67px "Calibri Light"', '58.67px "Nonexistent Font"', '58.67px sans-serif', 'bold 20px Calibri']) {
      const local = nodeCanvas().createCanvas(1, 1).getContext('2d'); local.font = font
      const server = b.createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D; server.font = font
      expect(Math.abs(local.measureText(text).width - server.measureText(text).width), font).toBeLessThan(0.5)
    }
  })

  test('cmdHoldDeck: an inert cmd keeps its interval, so the clicks equal local', async () => {
    const r = await compareStreamDeck('cmdHoldDeck', { marks: { kind: 'off' }, widening: { meanAbs: 0.01, rowMax: 0.05 }, family: 'latin' })
    // Local: one → seeks the held media call, one leaves the slide.
    expect(r.structure.local.clicks).toBe(2)
    expect(r.structure.streamed).toEqual(r.structure.local)
    expect(r.structure.equal).toBe(true)
    expect(r.outcomes.filter((o) => o.status === 'fail' || o.status === 'unmatched').map((o) => o.detail)).toEqual([])
  }, 120_000)

  test('fallback rate reported', async () => {
    for (const deck of DECKS) {
      const r = await run(deck, 'off')
      expect(r.fallback.steps).toBeGreaterThanOrEqual(0)
      expect(r.fallback.rate).toBe(r.fallback.steps ? r.fallback.composited / r.fallback.steps : 0)
    }
    expect((await run('flyInDeck', 'off')).fallback).toEqual({ steps: 1, composited: 0, rate: 0 })
  }, 120_000)
})
