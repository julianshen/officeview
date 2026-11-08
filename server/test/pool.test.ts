import { afterAll, describe, expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { DEFAULT_CAPS } from '../src/admission'
import { AdmissionRefused, Pool, RenderTimeout, type RenderJob } from '../src/pool'

const root = resolve(import.meta.dir, '../..')
const deck = new Uint8Array(readFileSync(resolve(root, 'corpus/pypptx-test.pptx')))
const sheet = new Uint8Array(readFileSync(resolve(root, 'corpus/poi-56295.xlsx')))
const job = (docRef: string, extra: Partial<RenderJob> = {}): RenderJob => ({ docRef, contentPolicy: 'visible-only', unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'fixed' }, ...extra })
const pools: Pool[] = []
const mk = (o: ConstructorParameters<typeof Pool>[0]) => { const p = new Pool(o); pools.push(p); return p }
afterAll(() => { for (const p of pools) p.dispose() })

describe('render pool (§9.2/§14)', () => {
  it('per-session caps: 8 running, then a queue of 32, then refusal', async () => {
    const pool = mk({ workers: 1 })
    await pool.open('deck', deck, 'visible-only')
    const runs = Array.from({ length: 40 }, () => pool.render('s1', job('deck'), 1024))
    expect(pool.stats()).toMatchObject({ running: 8, queued: 32 })
    await expect(pool.render('s1', job('deck'), 1024)).rejects.toEqual(new AdmissionRefused('session-queue'))
    const done = await Promise.all(runs)
    expect(done.every((u) => u.kind === 'raster')).toBe(true)
    expect(pool.stats()).toMatchObject({ running: 0, queued: 0, reserved: 0 })
  }, 60_000)

  it('a request larger than the session transient budget is refused', async () => {
    const pool = mk({ workers: 1 })
    await expect(pool.render('s1', job('deck'), DEFAULT_CAPS.sessionTransientBytes + 1)).rejects.toEqual(new AdmissionRefused('too-large'))
  })

  it('a wedged paint is retired at the deadline; its charge is refunded only after the worker is gone', async () => {
    const pool = mk({ workers: 1, caps: { ...DEFAULT_CAPS, renderDeadlineMs: 400 }, testHooks: true })
    await pool.open('deck', deck, 'visible-only')
    const run = pool.render('s1', job('deck', { hangMs: 1500 }), 4096)
    await Bun.sleep(300)
    expect(pool.stats().reserved).toBe(4096)
    await expect(run).rejects.toBeInstanceOf(RenderTimeout)
    const s = pool.stats()
    expect(s).toMatchObject({ reserved: 0, running: 0, retired: 1 })
    expect(s.workers[0]!.generation).toBe(1)
    // The replacement re-opens the document on demand.
    expect((await pool.render('s1', job('deck'), 4096)).kind).toBe('raster')
  }, 30_000)

  it('recycles a worker after the page budget and keeps serving', async () => {
    const pool = mk({ workers: 1, recyclePages: 3 })
    await pool.open('deck', deck, 'visible-only')
    for (let i = 0; i < 4; i++) await pool.render('s1', job('deck'), 1024)
    expect(pool.stats().recycled).toBe(1)
    expect(pool.stats().workers[0]!.generation).toBe(1)
  }, 30_000)

  it('pins documents to workers and parses once per (docRef, policy)', async () => {
    const pool = mk({ workers: 2 })
    await pool.open('deck', deck, 'visible-only')
    await pool.open('sheet', sheet, 'visible-only')
    const { affinity } = pool.stats()
    expect(affinity.deck).not.toBe(affinity.sheet)
    for (let i = 0; i < 3; i++) await pool.render('s1', job('deck'), 1024)
    expect(pool.stats().workers[affinity.deck!]!.parses).toBe(1)
    await pool.render('s1', job('deck', { contentPolicy: 'full' }), 1024)
    expect(pool.stats().workers[affinity.deck!]!.parses).toBe(2)
  }, 30_000)

  it('worker refusals surface with their code', async () => {
    const pool = mk({ workers: 1 })
    await pool.open('deck', deck, 'visible-only')
    await expect(pool.render('s1', job('deck', { unitIndex: 99 }), 1024)).rejects.toMatchObject({ refusal: 'unit-range' })
    expect(pool.stats().reserved).toBe(0)
  })
})
