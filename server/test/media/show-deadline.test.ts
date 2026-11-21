/**
 * Piece 5 Task 8: the deadline risk carried from Task 3. A show job's `loadShow` probes the slide's parts before
 * building; with the pool's media on, the probes run first as a media read job under their own deadline, so a cold,
 * media-heavy slide is not refused `show-plan-timeout` (nor is that refusal remembered).
 */
import { afterAll, describe, expect, it } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { mediaDeck } from '../../../src/testdata/show-decks'
import { createMediaTemp } from '../../src/media/temp'
import { Pool } from '../../src/pool'
import { SHOW_BUILD_BYTES } from '../../src/show/build'
import { ffmpegForTests } from './ffmpeg-gate'
import { rejectionMessage } from '../fixtures/rejection'

const info = await ffmpegForTests()
const root = mkdtempSync(join(tmpdir(), 'ov-deadline-'))
const temp = await createMediaTemp(root)
afterAll(async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) })

const box = { id: 2, name: 'Box', prst: 'rect', off: ['3619500', '2476500'] as [string, string], ext: ['1905000', '1905000'] as [string, string], fill: '336699' }
const audioNode = (id: number, spid: number) => `<p:audio><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:audio>`
const FILES = ['a.m4a', 'a.mp3', 'a.wav', 'a.wma'] as const

describe.skipIf(!info)('show jobs and the media probe deadline (Task 3 carry)', () => {
  it('a slide whose probes outlast the show deadline still builds: the probes run first under the media deadline', async () => {
    // Four parts held 900 ms each: 3.6 s of probing against a 2.5 s show deadline.
    const pool = new Pool({ workers: 1, media: { info: info!, dir: temp.dir }, testHooks: true, mediaHoldMs: 900, showDeadlineMs: 2500 })
    try {
      const bytes = await mediaDeck({ slides: [{ boxes: [box], media: FILES.map((file, i) => ({ spid: 3 + i, file })), timing: FILES.map((_, i) => audioNode(3 + i, 3 + i)).join('') }] })
      await pool.open('heavy', bytes, 'visible-only')
      const job = { docRef: 'heavy', contentPolicy: 'visible-only' as const, unitIndex: 0, scale: 1 as const, motion: 'full' as const, mark: { kind: 'fixed' as const }, strictMarks: false, media: true }
      const r = await pool.showPlan('s1', job, SHOW_BUILD_BYTES)
      expect('refused' in r ? r.refused : 'built').not.toBe('show-plan-timeout')
      expect(pool.stats().workers[0]!.probes).toBe(4)
      const again = await pool.showPlan('s2', job, SHOW_BUILD_BYTES)
      expect('refused' in again ? again.refused : 'built').toBe('refused' in r ? r.refused : 'built')
      expect(pool.stats().workers[0]!.probes).toBe(4) // one table per document: probed once
    } finally { pool.dispose() }
  }, 30_000)

  it('a warm-up past the media deadline answers that request show-plan-timeout and is never remembered', async () => {
    // Probes held 900 ms against a 300 ms media deadline (test-only): the warm-up times out every time.
    const pool = new Pool({ workers: 1, media: { info: info!, dir: temp.dir }, testHooks: true, mediaHoldMs: 900, mediaDeadlineMs: 300 })
    try {
      const bytes = await mediaDeck({ slides: [{ boxes: [box], media: [{ spid: 3, file: 'a.m4a' }], timing: audioNode(3, 3) }] })
      await pool.open('slow', bytes, 'visible-only')
      const job = { docRef: 'slow', contentPolicy: 'visible-only' as const, unitIndex: 0, scale: 1 as const, motion: 'full' as const, mark: { kind: 'fixed' as const }, strictMarks: false, media: true }
      const r1 = await pool.showPlan('s1', job, SHOW_BUILD_BYTES)
      expect('refused' in r1 && r1.refused).toBe('show-plan-timeout')
      const retired = pool.stats().retired
      const r2 = await pool.showPlan('s2', job, SHOW_BUILD_BYTES) // runs again: not answered from a remembered refusal
      expect('refused' in r2 && r2.refused).toBe('show-plan-timeout')
      expect(pool.stats().retired).toBe(retired + 1)
      // The track and morph jobs time out the same way, and are not remembered either.
      expect(await rejectionMessage(pool.showTrack('s3', { ...job, step: 0 }, SHOW_BUILD_BYTES))).toMatch(/deadline/)
      expect(await rejectionMessage(pool.showTrack('s3', { ...job, step: 0 }, SHOW_BUILD_BYTES))).toMatch(/deadline/)
      expect(pool.stats().retired).toBe(retired + 3)
    } finally { pool.dispose() }
  }, 30_000)
})
