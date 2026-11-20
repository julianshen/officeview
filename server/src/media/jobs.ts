/**
 * Piece 5 (plan D2, D21, D22, D23): a session's encode items for one slide, from `pool.mediaRefs`.
 *
 * Each `MediaUnitRefs` entry is one audio use (`uses.ts`, the single definition the plan build binds from), so an
 * item's `mediaId` is the one `bindMedia` gives the plan record of the same use: both hash the same `identity` under
 * the session's key. Items join plan records by that identity, never by `k`.
 *
 * Skipped: a hidden slide, a use on a hidden shape, and an entry the probe refused (its record ships `refused`).
 * Video nodes are not uses (D22), so no item is ever a video. The key is the document's (D23), so every session of
 * one document computes the same key for a part while its `mediaId`s differ.
 */
import type { MediaUnitRefs } from '../render'
import { mediaIdentity, sessionMediaId, type MediaRef } from './bind'
import type { FfmpegInfo } from './ffmpeg'
import { encodeKeyOf, type EncodeKey } from './params'
import type { ProbedMedia } from './probe'
import type { EncodeJob } from './transcode'

/** `partPath`: the part `pool.mediaPart` stages for the encode (server-only, never on the wire). */
export interface MediaItem { mediaId: string; unit: number; ref: MediaRef; key: EncodeKey; partPath: string; job: (out: string, src: string) => EncodeJob }

export function mediaJobsFor(refs: MediaUnitRefs /* pool.mediaRefs */, unit: number, s: { mediaKey: Uint8Array; docId: string }, info: FfmpegInfo): MediaItem[] {
  if (refs.unit !== unit) throw new Error('media: refs of another unit')
  if (refs.hidden) return []
  const items: MediaItem[] = []
  for (const r of refs.refs) {
    if (r.hiddenShape || 'refused' in r.entry) continue
    const entry: ProbedMedia = r.entry
    const ref: MediaRef = r.spid !== undefined ? { kind: 'audio', k: r.k, unit, spid: r.spid } : { kind: 'audio', k: r.k }
    // The identity is the join: a ref rebuilt here must hash to what the plan build bound.
    if (mediaIdentity(ref) !== r.identity) throw new Error('media: a use whose identity does not match its ref')
    items.push({
      mediaId: sessionMediaId(s.mediaKey, r.identity), unit, ref, key: encodeKeyOf(s.docId, entry, info), partPath: entry.partPath,
      job: (out, src) => ({ entry, src, out }),
    })
  }
  return items
}
