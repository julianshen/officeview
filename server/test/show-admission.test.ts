/**
 * PR #78 round 13b owner ruling (reserve per worker): a show job is admitted at request time for a small, bounded
 * request charge (`showRequestBytes`) against the session aggregate; its preparation grant is reserved only when
 * it holds its worker's slot, so server preparation memory is bounded by workers × `SHOW_BUILD_BYTES`.
 */
import { describe, expect, it } from 'bun:test'
import { DEFAULT_CAPS } from '../src/admission'
import { Pool } from '../src/pool'
import { SHOW_BUILD_BYTES, showBuildGrant, showRequestBytes } from '../src/show/grant'
import { flyInDeck } from '../../src/testdata/show-decks'

const job = { contentPolicy: 'full' as const, unitIndex: 0, scale: 1 as const, motion: 'full' as const, mark: { kind: 'off' as const }, strictMarks: false }
const grant = showBuildGrant({ w: 960, h: 720 }, 1)
const until = async (f: () => boolean) => { for (let i = 0; i < 400 && !f(); i++) await new Promise(r => setTimeout(r, 10)) }

describe('show admission: request charge at request time, the grant at slot time', () => {
  it('queued show jobs reserve only their request charge; the running one holds its grant on its worker', async () => {
    const pool = new Pool({ workers: 1, testHooks: true })
    try {
      await pool.open('fly', await flyInDeck(), 'full')
      const a = pool.showPlan('s1', { docRef: 'fly', ...job, hangMs: 1500 }, grant)
      const b = pool.showPlan('s1', { docRef: 'fly', ...job, mark: { kind: 'fixed' } }, grant)
      const c = pool.showPlan('s1', { docRef: 'fly', ...job, motion: 'reduced' }, grant)
      await until(() => pool.stats().prep.reserved > 0)
      const st = pool.stats()
      expect(st.reserved).toBe(3 * showRequestBytes(grant))
      expect(st.prep).toMatchObject({ reserved: grant, bound: SHOW_BUILD_BYTES })
      await Promise.all([a, b, c])
      expect(pool.stats().prep.reserved).toBe(0)
      expect(pool.stats().reserved).toBe(0)
    } finally { pool.dispose() }
  }, 60_000)

  it('preparation reserved never passes workers × SHOW_BUILD_BYTES, and each worker holds one grant at a time', async () => {
    const pool = new Pool({ workers: 2 })
    try {
      await pool.open('a', await flyInDeck(), 'full')
      await pool.open('b', await flyInDeck(), 'full')
      await Promise.all(Array.from({ length: 8 }, (_, i) => pool.showPlan(`s${i}`, { docRef: i % 2 ? 'a' : 'b', ...job, scale: (1 + (i % 2)) as 1 | 2, mark: { kind: i < 4 ? 'off' : 'fixed' } }, showBuildGrant({ w: 960, h: 720 }, 1 + (i % 2)))))
      const st = pool.stats()
      expect(st.prep.bound).toBe(2 * SHOW_BUILD_BYTES)
      expect(st.prep.peak).toBeGreaterThan(0)
      expect(st.prep.peak).toBeLessThanOrEqual(st.prep.bound)
      expect(st.prep.peakPerWorker).toBeLessThanOrEqual(SHOW_BUILD_BYTES)
      expect(st.prep.reserved).toBe(0)
    } finally { pool.dispose() }
  }, 120_000)

  it("a holding-style flood (40 plan requests, one session, scale-2 16:9 grants) queues within the session aggregate: none refused", async () => {
    const pool = new Pool({ workers: 1, testHooks: true })
    try {
      await pool.open('fly', await flyInDeck(), 'full')
      const big = showBuildGrant({ w: 1280, h: 720 }, 2)
      // More than the whole session aggregate: as a request charge it would be refused too-large.
      expect(big).toBeGreaterThan(DEFAULT_CAPS.sessionTransientBytes)
      const all = Array.from({ length: 40 }, (_, i) => pool.showPlan('s1', { docRef: 'fly', ...job, unitIndex: 0, strictMarks: i % 2 === 1, mark: { kind: 'viewer' as const, text: `v${i}` }, ...(i === 0 ? { hangMs: 500 } : {}) }, big).then(() => 'ok', (e: Error) => e.constructor.name))
      const out = await Promise.all(all)
      expect(out.filter(x => x !== 'ok')).toEqual([])
      expect(pool.admission.peakSessionReserved).toBeLessThanOrEqual(DEFAULT_CAPS.sessionTransientBytes)
      expect(pool.admission.peakSessionQueued).toBeLessThanOrEqual(DEFAULT_CAPS.perSessionQueue)
    } finally { pool.dispose() }
  }, 120_000)
})
