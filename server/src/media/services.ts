/**
 * Piece 5 (plan X8, D12, D18, D23): the one media object a process holds. `main.ts` builds it at startup and hands
 * its `info` and `temp.dir` to the pool; `ws.ts` (Task 9) holds it for the session queues.
 *
 * Fails closed. With `enabled` false (D18: until Task 11) or a failed ffmpeg probe (`probeFfmpeg`: `OV_MEDIA=off`,
 * a missing binary, a sandbox limit the platform cannot apply), `info` is undefined, nothing touches the disk, and
 * `temp` is an inert stub; every media record is then `media-unavailable`. Otherwise it claims the process's one
 * media temp directory (`createMediaTemp`), the admission class (⌈cores / 4⌉ slots, queue 16) and the shared cache
 * (2 GiB, 256 MiB per session). A temp directory that cannot be claimed is media unavailable too.
 *
 * Queues are created through `createQueue` (an addition to the brief's interface: `close()` must end every queue, so
 * the services track them); a queue leaves the registry when it ends.
 */
import { availableParallelism } from 'node:os'
import { createMediaAdmission, MEDIA_QUEUE, type MediaAdmission } from './admission'
import { createEncodeCache, MEDIA_SERVER_CACHE, MEDIA_SESSION_LOOKAHEAD, type EncodeCache } from './cache'
import { probeFfmpeg, type FfmpegInfo } from './ffmpeg'
import { createSessionMediaQueue, type SessionMediaQueue, type SessionMediaQueueOptions } from './queue'
import { createMediaTemp, type MediaTemp } from './temp'
import { transcode } from './transcode'

/** D18: media in streamed shows stays off until the client can play it (Task 11 turns it on). */
export const MEDIA_ENABLED_DEFAULT = false

/** TEST-ONLY. `root`: the parent of the media temp directory (default the OS temp dir). */
export interface MediaTestOnly { maxBytes?: number; wallCapMs?: number; slots?: number; queue?: number; root?: string }
export type MediaQueueOptions = Pick<SessionMediaQueueOptions, 'sid' | 'items' | 'part' | 'units'>
export interface MediaServices {
  readonly info: FfmpegInfo | undefined
  readonly temp: MediaTemp
  readonly admission: MediaAdmission
  readonly cache: EncodeCache
  readonly testOnly?: MediaTestOnly
  /** A session's queue over the shared admission and cache. Throws when media is unavailable. */
  createQueue(o: MediaQueueOptions): SessionMediaQueue
  /** Ends every queue, closes the cache, removes the temp dir. */
  close(): Promise<void>
}

export const mediaSlots = (cores = availableParallelism()) => Math.max(1, Math.ceil(cores / 4))

/** Media off: no directory; any use throws. */
const INERT_TEMP: MediaTemp = {
  dir: '', create: () => Promise.reject(new Error('media unavailable')), adopt: () => { throw new Error('media unavailable') }, reserve: () => { throw new Error('media unavailable') },
  live: () => 0, bytes: () => 0, remove: () => Promise.resolve(),
}

export async function createMediaServices(o: { env: Record<string, string | undefined>; enabled: boolean; testOnly?: MediaTestOnly }): Promise<MediaServices> {
  const t = o.testOnly
  let info: FfmpegInfo | undefined
  let temp = INERT_TEMP
  if (o.enabled) {
    const probe = await probeFfmpeg({ OV_FFMPEG: o.env.OV_FFMPEG, OV_FFPROBE: o.env.OV_FFPROBE, OV_MEDIA: o.env.OV_MEDIA })
    if (probe.ok) {
      try { temp = await createMediaTemp(t?.root); info = probe.info } catch (e) { console.warn(`[media] unavailable: ${(e as Error).message}`) }
    } else console.warn(`[media] unavailable: ffmpeg probe ${probe.reason}`)
  }
  const admission = createMediaAdmission({ slots: t?.slots ?? mediaSlots(), queue: t?.queue ?? MEDIA_QUEUE })
  const cache = createEncodeCache({ temp, sessionBytes: MEDIA_SESSION_LOOKAHEAD, serverBytes: MEDIA_SERVER_CACHE })
  const queues = new Set<SessionMediaQueue>()
  const seams = t && (t.maxBytes !== undefined || t.wallCapMs !== undefined) ? { ...(t.maxBytes !== undefined ? { maxBytes: t.maxBytes } : {}), ...(t.wallCapMs !== undefined ? { wallCapMs: t.wallCapMs } : {}) } : undefined
  let closing: Promise<void> | undefined
  return {
    info, temp, admission, cache, ...(t ? { testOnly: t } : {}),
    createQueue(q) {
      if (!info || closing) throw new Error('media unavailable')
      const queue = createSessionMediaQueue({ ...q, info, temp, admission, cache, transcode, ...(seams ? { seams } : {}) })
      const end = queue.end.bind(queue)
      const tracked: SessionMediaQueue = {
        start: (u) => queue.start(u), cursor: (u) => queue.cursor(u), request: (id) => queue.request(id),
        async end() { try { await end() } finally { queues.delete(tracked) } },
        get running() { return queue.running },
      }
      queues.add(tracked)
      return tracked
    },
    close() {
      return (closing ??= (async () => {
        await Promise.all([...queues].map((q) => q.end()))
        await cache.close()
        await temp.remove()
      })())
    },
  }
}
