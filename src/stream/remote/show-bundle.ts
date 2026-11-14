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
 * Each bitmap is charged to the show workspace at `w·h·4` before it is decoded; no room rejects
 * `workspace-full`. On any failure every bitmap decoded so far is closed and every lease released.
 * `release()` (idempotent) does the same for a delivered bundle. Given the session's `ShowReply`,
 * it releases the reply's retained-payload leases once it is done with the messages, on success or
 * failure (controller ruling 5: the encoded bytes are charged until decoded).
 *
 * Package-free (D25): its only value imports are the codec and the error type.
 */
import { decodeShowMessage, isRasterLayer, SHOW_PLAN_LIMITS, type ShowMessage, type ShowPlanRecord } from '../show-plan'
import { StreamSessionError } from './errors'
import type { Lease, SlideShowWorkspace } from '../../slideshow/types'
import type { PlanAssets } from './plan-stage'
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
      if (total > SHOW_PLAN_LIMITS.showDecoded) reject('the bundle is over the show budget')
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
      bytes: (id) => (released ? 0 : held.get(id)?.bytes ?? 0),
    }
    return { unit, plan, assets, decodedBytes: total, release }
  } catch (e) {
    release()
    throw e
  } finally {
    done()
  }
}
