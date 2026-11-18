/**
 * PR #78 round 13b: every canvas a show job creates is reserved against the job's grant first. The raw backend is
 * instrumented (as codex's r13-accounting probe), so the checks see real canvas creation, not the ledger alone:
 * - every byte created during a job was leased (created = leased);
 * - the ledger never passes its grant;
 * - the ledger ends empty after every outcome, refusals at any phase included (a grant sweep).
 * One host per test: the host installs its backend process-wide.
 */
import { describe, expect, it } from 'bun:test'
import { serverBackend } from '../src/backend'
import { DocumentHost } from '../src/render'
import { showBuildGrant, showMorphGrant } from '../src/show/grant'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'
import type { SlideShowWorkspace } from '../../src/slideshow/types'
import { flyInDeck, morphAdversaryDeck, morphDeck } from '../../src/testdata/show-decks'
import { fullSlideLayersDeck, growthDeck, openDeck } from './show-fixtures'

function rig() {
  const base = serverBackend()
  const s = { created: 0, leased: 0, peak: 0, grant: -1, ws: undefined as SlideShowWorkspace | undefined, counting: false }
  const backend = { ...base, createCanvas(w: number, h: number) { if (s.counting) s.created += Math.max(1, Math.ceil(w)) * Math.max(1, Math.ceil(h)) * 4; return base.createCanvas(w, h) } }
  const host = new DocumentHost(backend)
  host.showOptions = {
    ledger: (grant) => {
      const ws = createSlideShowWorkspace('desktop', { limit: grant, scratch: 0 })
      s.grant = grant; s.ws = ws
      return {
        limit: ws.limit, used: ws.used, reserveTerminal: ws.reserveTerminal,
        lease(bytes, tag) { const l = ws.lease(bytes, tag); if (l && tag === 'show:canvas') s.leased += bytes; s.peak = Math.max(s.peak, ws.used()); return l },
      }
    },
  }
  /** Runs one job with creation counted; resets the counters first. */
  const job = async <T>(f: () => Promise<T>) => { s.created = 0; s.leased = 0; s.peak = 0; s.ws = undefined; s.counting = true; try { return await f() } finally { s.counting = false } }
  return { s, host, job }
}
const plan = (scale: 1 | 2) => ({ unitIndex: 0, scale, motion: 'full' as const, mark: { kind: 'fixed' as const }, strictMarks: false })
const outcome = (r: object) => ('refused' in r ? (r as { refused: string }).refused : 'built')

describe('show job canvas accounting (PR #78 round 13b)', () => {
  it('codex probe: two full-slide boxes at scale 2 create only what they lease, within the grant, and end empty', async () => {
    const { s, host, job } = rig()
    const { docKey } = await openDeck(await fullSlideLayersDeck(2), 'probe', host)
    await host.showPlan(docKey, plan(1)) // loads the slide and its images outside the measured job
    await job(() => host.showPlan(docKey, plan(2)))
    expect(s.grant).toBe(showBuildGrant({ w: 960, h: 720 }, 2))
    expect(s.created).toBeGreaterThan(0)
    expect(s.created).toBe(s.leased)
    expect(s.peak).toBeLessThanOrEqual(s.grant)
    expect(s.ws!.used()).toBe(0)
  }, 60_000)

  it('a grant sweep: every outcome, refused at any phase or built, creates only leased bytes and ends empty', async () => {
    const page = 960 * 720 * 4
    const seen = new Set<string>()
    for (const pages of [0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8]) {
      const { s, host, job } = rig()
      const { docKey } = await openDeck(await flyInDeck(), 'sweep', host)
      const grant = Math.floor(pages * page)
      await host.showPlan(docKey, plan(1), 64 * 1024 * 1024)
      const r = await job(() => host.showPlan(docKey, { ...plan(1), mark: { kind: 'off' } }, grant))
      seen.add(outcome(r))
      expect({ pages, created: s.created }).toEqual({ pages, created: s.leased })
      expect(s.peak).toBeLessThanOrEqual(grant)
      expect(s.ws!.used()).toBe(0)
    }
    // The sweep reaches both a refusal and a build.
    expect(seen).toEqual(new Set(['show-budget', 'built']))
  }, 120_000)

  it('a slide that fits its rasters but not its working set refuses before it over-allocates', async () => {
    // The grown strip: its rasters fit one grant of 3 pages; masks, marking and validation do not.
    const { s, host, job } = rig()
    const { docKey } = await openDeck(await growthDeck(10, { x: 0, y: 340, w: 960, h: 24 }), 'strip', host)
    await host.showPlan(docKey, plan(1), 64 * 1024 * 1024)
    const page = 960 * 720 * 4 * 4
    const r = await job(() => host.showPlan(docKey, { ...plan(2), mark: { kind: 'off' } }, 3 * page))
    expect(outcome(r)).toBe('show-budget')
    expect(s.created).toBe(s.leased)
    expect(s.peak).toBeLessThanOrEqual(3 * page)
    expect(s.ws!.used()).toBe(0)
  }, 60_000)

  it('a track job (its plan and every frame) and a Morph pair (proof, fallback, validation) are held to one grant and end empty', async () => {
    for (const grantPages of [1, 3, 64]) {
      const { s, host, job } = rig()
      const { docKey } = await openDeck(await flyInDeck(), 'track', host)
      await host.showPlan(docKey, plan(1), 64 * 1024 * 1024)
      const grant = Math.min(64 * 1024 * 1024, grantPages * 960 * 720 * 4 * 4)
      await job(() => host.showTrack(docKey, { ...plan(2), step: 0 }, grant))
      expect(s.created).toBe(s.leased)
      expect(s.peak).toBeLessThanOrEqual(grant)
      expect(s.ws!.used()).toBe(0)
    }
    for (const grant of [960 * 720 * 4, 3 * 960 * 720 * 4, showMorphGrant({ w: 960, h: 720 }, { w: 960, h: 720 }, 1)]) {
      const { s, host, job } = rig()
      const { docKey } = await openDeck(await morphDeck(), 'morph', host)
      await host.showPlan(docKey, { ...plan(1), unitIndex: 1 }, 64 * 1024 * 1024)
      await host.showPlan(docKey, plan(1), 64 * 1024 * 1024).catch(() => undefined)
      await job(() => host.showMorph(docKey, { from: 0, to: 1, fromStep: 0, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false }, grant))
      expect(s.created).toBe(s.leased)
      expect(s.peak).toBeLessThanOrEqual(grant)
      expect(s.ws!.used()).toBe(0)
    }
  }, 180_000)

  it('a layered Morph pair lets each masked original go once its copy exists (T5 coveredPair peaks under 6 pages, not 7.4)', async () => {
    const { s, host, job } = rig()
    const { docKey } = await openDeck(await morphAdversaryDeck('coveredPair'), 'covered', host)
    await host.showPlan(docKey, { ...plan(1), unitIndex: 1 }, 64 * 1024 * 1024)
    const r = await job(() => host.showMorph(docKey, { from: 0, to: 1, fromStep: 0, scale: 1, motion: 'full', mark: { kind: 'off' }, strictMarks: false }))
    expect('refused' in r || 'frames' in r).toBe(false)
    expect(s.created).toBe(s.leased)
    expect(s.peak).toBeLessThanOrEqual(6 * 960 * 720 * 4)
    expect(s.ws!.used()).toBe(0)
  }, 60_000)
})
