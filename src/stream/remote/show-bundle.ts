/**
 * A streamed slide bundle (piece 4 Task 10, D1, D2; spec §5): the kind-6 messages of one
 * `show-plan` reply, decoded into the plan record and its bitmaps.
 *
 * The codec (`decodeShowMessage`) checks each message on its own; this module holds the reply
 * together, failing closed (`policy-refused`) on anything else:
 * - message 0 is the plan record, for the unit asked;
 * - message 1 is the entry frame, its assetId the plan's `entry`;
 * - then exactly one tag-2 asset per raster layer, none missing, none twice, none unnamed, each
 *   no larger than its layer's `decodedBytes` (`w·h·4`); no keyframe or Morph record;
 * - every asset for the unit asked, and the decoded bitmap the asset's size;
 * - Σ decoded (the plan record's bytes and every asset's `w·h·4`) within the 128 MiB show budget.
 * The plan record is charged at its message bytes for the bundle's life (fix round 1), and each
 * bitmap at `w·h·4` before it is decoded, so the workspace holds exactly `decodedBytes`; no room rejects
 * `workspace-full`. On any failure every bitmap decoded so far is closed and every lease released.
 * `release()` (idempotent) does the same for a delivered bundle. Given the session's `ShowReply`,
 * it releases the reply's retained-payload leases once it is done with the messages, on success or
 * failure (controller ruling 5: the encoded bytes are charged until decoded).
 *
 * A bundle over the 128 MiB show budget on its own is refused `show-budget` (the slide plays its
 * piece-3 still, spec §5).
 *
 * Piece 4 Task 11 adds the other two replies: `decodeTrack` (a `show-track` reply, D4, D23) and
 * `decodeMorph` (a `show-morph` reply, D11), with the same fail-closed checks.
 *
 * Package-free (D25): its only value imports are the codec, the error type and the D23 track window.
 */
import { decodeShowMessage, isRasterLayer, SHOW_PLAN_LIMITS, type MorphPlanRecord, type ShowMessage, type ShowPlanRecord } from '../show-plan'
import { StreamSessionError } from './errors'
import type { Lease, SlideShowWorkspace } from '../../slideshow/types'
import { createTrackWindow, type PlanAssets, type PlanTrack } from './plan-stage'
import type { ShowReply } from './session'
import type { ImageBitmapLike } from './show-fetch'

export interface ShowBundle { unit: number; plan: ShowPlanRecord; assets: PlanAssets; decodedBytes: number; release(): void }

type Asset = Extract<ShowMessage, { tag: 2 | 3 | 4 }>
interface Held { lease: Lease; bytes: number; bmp?: ImageBitmapLike }

const closeBitmap = (b: ImageBitmapLike | undefined) => { try { b?.close?.() } catch { /* already released */ } }

export async function decodeBundle(unit: number, reply: readonly Uint8Array[] | ShowReply, workspace: SlideShowWorkspace, decode: (png: Uint8Array) => Promise<ImageBitmapLike>): Promise<ShowBundle> {
  const messages: readonly Uint8Array[] = Array.isArray(reply) ? reply : (reply as ShowReply).messages
  const done = () => { if (!Array.isArray(reply)) (reply as ShowReply).release() }
  const held = new Map<number, Held>()
  let released = false
  const release = () => {
    if (released) return
    released = true
    for (const h of held.values()) { closeBitmap(h.bmp); h.lease.release() }
    held.clear()
  }
  const reject = (why: string): never => { throw new StreamSessionError('policy-refused', `show bundle rejected: ${why}`) }
  const expect = { unit, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }
  const message = (i: number): ShowMessage => {
    try { return decodeShowMessage(messages[i]!, expect) } catch (e) { return reject(`message ${i}: ${(e as Error).message}`) }
  }
  try {
    if (!messages.length) reject('no messages')
    const head = message(0)
    if (head.tag !== 1) return reject('message 0 is not the plan record')
    const plan = head.plan
    // The record's accounting size (its message bytes, as `decodedBytes` counts it), held while the bundle lives.
    const record = workspace.lease(messages[0]!.length, 'remote-show:plan')
    if (!record) throw new StreamSessionError('workspace-full', 'no workspace room for a show plan record')
    held.set(-1, { lease: record, bytes: 0 })
    // assetId → the decoded bytes its layer allows; the entry frame is bounded by the codec's per-asset cap.
    const want = new Map<number, number>([[plan.entry, SHOW_PLAN_LIMITS.assetDecoded]])
    for (const l of plan.layers) if (isRasterLayer(l)) want.set(l.assetId, l.decodedBytes)
    if (messages.length !== want.size + 1) reject(`${messages.length} messages for ${want.size} assets`)
    let total = messages[0]!.length
    for (let i = 1; i < messages.length; i++) {
      const m = message(i)
      if (m.tag !== (i === 1 ? 3 : 2)) reject(`message ${i} has tag ${m.tag}`)
      const a = m as Asset
      if (i === 1 && a.assetId !== plan.entry) reject('the entry frame is not the plan\'s')
      // The entry frame is always held by now, so a second copy of it is a duplicate too.
      const allow = want.get(a.assetId)
      if (allow === undefined || held.has(a.assetId)) reject(`asset ${a.assetId} is not one the plan names once`)
      const bytes = a.w * a.h * 4
      if (bytes > allow!) reject(`asset ${a.assetId} is larger than its layer`)
      total += bytes
      if (total > SHOW_PLAN_LIMITS.showDecoded) throw new StreamSessionError('show-budget', 'show bundle rejected: the bundle is over the show budget')
      const lease = workspace.lease(bytes, `remote-show:asset:${a.assetId}`)
      if (!lease) throw new StreamSessionError('workspace-full', 'no workspace room for a show asset')
      const h: Held = { lease, bytes }
      held.set(a.assetId, h)
      let bmp: ImageBitmapLike
      try { bmp = await decode(a.png) } catch { return reject(`asset ${a.assetId} does not decode`) }
      h.bmp = bmp
      if (bmp.width !== a.w || bmp.height !== a.h) reject(`asset ${a.assetId} decoded to another size`)
    }
    const assets: PlanAssets = {
      bitmap: (id) => (released ? undefined : held.get(id)?.bmp),
      bytes: (id) => (released || id < 0 ? 0 : held.get(id)?.bytes ?? 0),
    }
    return { unit, plan, assets, decodedBytes: total, release }
  } catch (e) {
    release()
    throw e
  } finally {
    done()
  }
}

// ---------------------------------------------------------------------------
// Keyframe tracks and Morph pairs (Task 11)
// ---------------------------------------------------------------------------

/** A decoded keyframe track: the D23 window over the reply's frames. `bytes`: the encoded frames held (charged until `dispose`). */
export interface ShowTrack extends PlanTrack { readonly bytes: number; dispose(): void }

type Frame = Extract<ShowMessage, { tag: 2 | 3 | 4 }>
const messagesOf = (reply: readonly Uint8Array[] | ShowReply): readonly Uint8Array[] => (Array.isArray(reply) ? reply : (reply as ShowReply).messages)
const releaseReply = (reply: readonly Uint8Array[] | ShowReply) => { if (!Array.isArray(reply)) (reply as ShowReply).release() }
const refuse = (what: string, why: string): never => { throw new StreamSessionError('policy-refused', `${what} rejected: ${why}`) }

/**
 * The frames `messages[from..]` as one track for `unit`: every message a tag-4 keyframe of the unit, all one size,
 * exactly `count` of them (1..90). The PNGs stay encoded (the reply's leases keep them charged); the window decodes
 * at most 8 at a time, each charged `w·h·4` before it is decoded, and checks each decoded size.
 */
function trackOf(what: string, unit: number, messages: readonly Uint8Array[], from: number, count: number, workspace: SlideShowWorkspace, decode: (png: Uint8Array) => Promise<ImageBitmapLike>): { track: PlanTrack & { dispose(): void }; bytes: number } {
  if (!Number.isInteger(count) || count < 1 || count > SHOW_PLAN_LIMITS.trackFrames) refuse(what, `a track of ${count} frames`)
  if (messages.length - from !== count) refuse(what, `${messages.length - from} frames for a track of ${count}`)
  const frames: Frame[] = []
  let bytes = 0
  for (let i = from; i < messages.length; i++) {
    let m: ShowMessage
    try { m = decodeShowMessage(messages[i]!, { unit, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded }) } catch (e) { return refuse(what, `message ${i}: ${(e as Error).message}`) }
    if (m.tag !== 4) refuse(what, `message ${i} has tag ${m.tag}`)
    const f = m as Frame
    if (frames.length && (f.w !== frames[0]!.w || f.h !== frames[0]!.h)) refuse(what, `frame ${frames.length} is another size`)
    frames.push(f)
    bytes += messages[i]!.length
  }
  const { w, h } = frames[0]!
  const track = createTrackWindow({
    count, frameBytes: w * h * 4, workspace,
    async decode(i) {
      const f = frames[i]
      if (!f) return undefined
      const bmp = await decode(f.png)
      // A frame that decodes to another size is dropped (the window holds the previous picture).
      if (bmp.width !== w || bmp.height !== h) { closeBitmap(bmp); throw new Error('frame size') }
      return bmp
    },
  })
  return { track, bytes }
}

/**
 * One `show-track` reply of `unit` (D4): `count` tag-4 frames (the plan step's `track.count`). Given a `ShowReply`,
 * its encoded leases are held until `dispose()` (D23: the encoded PNGs stay charged); on a rejection they are
 * released at once.
 */
export function decodeTrack(unit: number, count: number, reply: readonly Uint8Array[] | ShowReply, workspace: SlideShowWorkspace, decode: (png: Uint8Array) => Promise<ImageBitmapLike>): ShowTrack {
  let t: ReturnType<typeof trackOf>
  try { t = trackOf('show track', unit, messagesOf(reply), 0, count, workspace, decode) } catch (e) { releaseReply(reply); throw e }
  let disposed = false
  return {
    fps: 15, count: t.track.count, bytes: t.bytes,
    frame: (i) => t.track.frame(i),
    want: (i) => t.track.want(i),
    dispose() { if (disposed) return; disposed = true; t.track.dispose(); releaseReply(reply) },
  }
}

/** A decoded Morph pair (D11): its record and either its layer assets or its composited track. `bytes`: what it holds. */
export interface ShowMorph { record: MorphPlanRecord; assets: PlanAssets; track?: PlanTrack; bytes: number; release(): void }

/**
 * One `show-morph` reply for `(from, to, fromStep)` (unit indices). Message 0 is the tag-5 record, for that pair.
 * A layered record is followed by exactly one tag-2 asset per assetId its sides and `over` name (every one present
 * once, none other; each for `to`'s unit), each leased at `w·h·4` before it is decoded and checked against its size;
 * the reply's encoded leases are released once decoded. A composited record is followed by exactly its `count`
 * tag-4 frames, decoded through the D23 window; its encoded leases are held until `release()`. Σ within the show
 * budget. Anything else rejects `policy-refused` with every lease released.
 */
export async function decodeMorph(pair: { from: number; to: number; fromStep: number }, reply: readonly Uint8Array[] | ShowReply, workspace: SlideShowWorkspace, decode: (png: Uint8Array) => Promise<ImageBitmapLike>): Promise<ShowMorph> {
  const what = 'Morph pair'
  const messages = messagesOf(reply)
  const held = new Map<number, Held>()
  let released = false
  let window: (PlanTrack & { dispose(): void }) | undefined
  let keepReply = false
  const release = () => {
    if (released) return
    released = true
    for (const h of held.values()) { closeBitmap(h.bmp); h.lease.release() }
    held.clear()
    window?.dispose()
    if (keepReply) releaseReply(reply)
  }
  try {
    if (!messages.length) refuse(what, 'no messages')
    let head: ShowMessage
    try { head = decodeShowMessage(messages[0]!, { unit: pair.to, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) } catch (e) { return refuse(what, `message 0: ${(e as Error).message}`) }
    if (head.tag !== 5) return refuse(what, 'message 0 is not the Morph record')
    const record = head.morph
    if (record.from !== pair.from || record.to !== pair.to || record.fromStep !== pair.fromStep) refuse(what, 'the record is for another pair')
    const record0 = workspace.lease(messages[0]!.length, 'remote-show:morph')
    if (!record0) throw new StreamSessionError('workspace-full', 'no workspace room for a Morph record')
    held.set(-1, { lease: record0, bytes: 0 })
    let total = messages[0]!.length
    if (record.composited) {
      if (record.items.length) refuse(what, 'a composited record names items')
      const t = trackOf(what, pair.to, messages, 1, record.composited.count, workspace, decode)
      window = t.track
      keepReply = true
      total += t.bytes
      if (total > SHOW_PLAN_LIMITS.showDecoded) throw new StreamSessionError('show-budget', 'Morph pair rejected: over the show budget')
      return { record, assets: { bitmap: () => undefined, bytes: () => 0 }, track: t.track, bytes: total, release }
    }
    // Every assetId a side or `over` names ships once (sides may share one asset).
    const want = new Set<number>()
    for (const it of record.items) for (const side of [it.from, it.to]) if (side) want.add(side.assetId)
    for (const id of record.over) want.add(id)
    if (messages.length !== want.size + 1) refuse(what, `${messages.length} messages for ${want.size} assets`)
    for (let i = 1; i < messages.length; i++) {
      let m: ShowMessage
      try { m = decodeShowMessage(messages[i]!, { unit: pair.to, maxDecoded: SHOW_PLAN_LIMITS.assetDecoded }) } catch (e) { return refuse(what, `message ${i}: ${(e as Error).message}`) }
      if (m.tag !== 2) refuse(what, `message ${i} has tag ${m.tag}`)
      const a = m as Frame
      if (!want.has(a.assetId) || held.has(a.assetId)) refuse(what, `asset ${a.assetId} is not one the record names once`)
      const bytes = a.w * a.h * 4
      total += bytes
      if (total > SHOW_PLAN_LIMITS.showDecoded) throw new StreamSessionError('show-budget', 'Morph pair rejected: over the show budget')
      const lease = workspace.lease(bytes, `remote-show:morph-asset:${a.assetId}`)
      if (!lease) throw new StreamSessionError('workspace-full', 'no workspace room for a Morph asset')
      const h: Held = { lease, bytes }
      held.set(a.assetId, h)
      let bmp: ImageBitmapLike
      try { bmp = await decode(a.png) } catch { return refuse(what, `asset ${a.assetId} does not decode`) }
      h.bmp = bmp
      if (bmp.width !== a.w || bmp.height !== a.h) refuse(what, `asset ${a.assetId} decoded to another size`)
    }
    releaseReply(reply)
    const assets: PlanAssets = {
      bitmap: (id) => (released ? undefined : held.get(id)?.bmp),
      bytes: (id) => (released || id < 0 ? 0 : held.get(id)?.bytes ?? 0),
    }
    return { record, assets, bytes: total, release }
  } catch (e) {
    keepReply = true
    release()
    throw e
  }
}
