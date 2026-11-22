/**
 * Final review finding 1 (plan D1, D23): the production `StreamSession` against the real server admits keyframe
 * tracks and composited Morph pairs whose summed frames exceed the 128 MiB show budget but whose live D23 window
 * (at most 8 decoded frames) fits. The starts declare `decodedBytes` = the window bound (plus a Morph's record)
 * and `encodedBytes` = Σ encoded payloads; the reply then plays through the window to its terminal frame.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { StreamSession } from '../../src/stream/remote/session'
import { decodeMorph, decodeTrack, type ShowTrack } from '../../src/stream/remote/show-bundle'
import type { PlanTrack } from '../../src/stream/remote/plan-stage'
import { decodeShowMessage, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../src/stream/show-plan'
import { createSlideShowWorkspace } from '../../src/slideshow/workspace'
import { longStepDeck } from '../../src/testdata/show-decks'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { morphSweepDeck } from './show-fixtures'
import { ORIGIN, disposePool, start } from './harness'

afterAll(async () => { await disposePool() })

const MiB = 1024 * 1024
/** A decode yielding a closable bitmap of the PNG's IHDR size (the session path never needs the pixels here). */
const decode = async (png: Uint8Array) => {
  const v = new DataView(png.buffer, png.byteOffset, png.byteLength)
  return { width: v.getUint32(16), height: v.getUint32(20), close() {} }
}
const frameSize = (m: Uint8Array) => { const v = new DataView(m.buffer, m.byteOffset, m.byteLength); return v.getUint32(9, true) * v.getUint32(13, true) * 4 }

async function opened(doc: Uint8Array) {
  const h = await start({ doc, audit: new AuditTable({ store: AuditStore.open(':memory:') }), policy: { contentPolicy: 'full' } })
  const s = await StreamSession.open({ kind: 'stream', url: h.url }, { client: (o: { url: string }) => new WebSocket(o.url, { headers: { Origin: ORIGIN } } as never) })
  return { s, stop: () => { s.close(); h.stop() } }
}
/** Rolls the window frame by frame (as playback asks for them) and returns the last frame it shows. */
async function playToEnd(t: PlanTrack) {
  for (let i = 0; i < t.count; i++) {
    t.want(i)
    for (let k = 0; k < 20 && !t.frame(i); k++) await Bun.sleep(0)
    expect({ i, shown: !!t.frame(i) }).toEqual({ i, shown: true })
  }
  return t.frame(t.count - 1)
}

describe('production StreamSession admits D23-windowed replies', () => {
  test('longStepDeck: the 90-frame 960×720 track (Σ ≈ 249 MB decoded) is admitted and plays to its terminal frame', async () => {
    const { s, stop } = await opened(await longStepDeck())
    const ws = createSlideShowWorkspace('desktop', { limit: 512 * MiB, scratch: 0 })
    const sig = new AbortController().signal
    const bundle = await s.showPlan(0, 1, 'full', sig, { workspace: ws })
    const plan = (decodeShowMessage(bundle.messages[0]!, { unit: 0, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: ShowPlanRecord }).plan
    bundle.release()
    const step = plan.steps[0]!
    expect(step.eligibility).toBe('composited')
    expect(step.track?.count).toBe(90)
    const reply = await s.showTrack(0, 0, 1, 'full', sig, { workspace: ws })
    expect(reply.messages.length).toBe(90)
    expect(reply.messages.reduce((n, m) => n + frameSize(m), 0)).toBeGreaterThan(SHOW_PLAN_LIMITS.showDecoded)
    expect(frameSize(reply.messages[0]!)).toBe(960 * 720 * 4)
    const t: ShowTrack = decodeTrack(0, 90, reply, ws, decode)
    const last = await playToEnd(t)
    expect(last).toMatchObject({ width: 960, height: 720 })
    t.dispose()
    expect(ws.used()).toBe(0)
    stop()
  }, 120_000)

  test('a composited Morph whose Σ frames exceed 128 MiB is admitted on its window and plays to its terminal frame', async () => {
    // 4 s at 15 fps: 60 frames of 960×720 (≈ 166 MB decoded in all; the window holds 8).
    const { s, stop } = await opened(await morphSweepDeck(4000))
    const ws = createSlideShowWorkspace('desktop', { limit: 512 * MiB, scratch: 0 })
    const reply = await s.showMorph(0, 1, 0, 1, 'full', new AbortController().signal, { workspace: ws })
    const frames = reply.messages.slice(1)
    expect(frames.reduce((n, m) => n + frameSize(m), 0)).toBeGreaterThan(SHOW_PLAN_LIMITS.showDecoded)
    const m = await decodeMorph({ from: 0, to: 1, fromStep: 0 }, reply, ws, decode)
    expect(m.record.composited?.count).toBe(frames.length)
    const last = await playToEnd(m.track!)
    expect(last).toMatchObject({ width: 960, height: 720 })
    m.release()
    expect(ws.used()).toBe(0)
    stop()
  }, 120_000)
})
