/**
 * Piece 5 (spec §4.1; plan D2, D3, D21, D22, X13; Task 4 fix round 1): THE definition of a slide's audio uses — their
 * identity, `k` and refusal. The host's open (the D19 unit flag), `loadShow` (whether to read the slide's media
 * context), `mediaRefs` (the main thread's view, D21) and the plan build (`streamedMeta`) all derive from it, so a
 * plan's `mediaId` and a `mediaRefs` entry always name the same use.
 *
 * - A shape audio node on a part: `@a<k>.<spid>`, identity `audio|k|unit|spid`; on an external link: refused
 *   `media-linked`; with neither, nothing to play (no use). A use on a hidden shape is listed `hiddenShape` (the plan
 *   ships nothing for it).
 * - A sound (`target.snd` of any behaviour or condition, or the transition's start): its part's `@a<k>`, identity
 *   `audio|k` (one record per part however often it is used, X13). A `builtIn:` sound is refused `media-unreadable`
 *   (no part), an external one `media-linked`, one naming no relationship `media-unreadable`.
 * - A part-less use gets a synthetic `k` at or past `partCount` (distinct per use within the slide) for its
 *   placeholder, and an identity scoped to its slide and source key (`audio|-|unit|<source>`), so two part-less uses
 *   never share a `mediaId` (a real part, `0 ≤ k < partCount`, keeps its document-global identity: one encode).
 * - A video node is no use (D22): its spid is listed in `videoSpids`, which keeps its commands blanked.
 *
 * The walk is bounded at the reader's nesting limit (64).
 */
import type { Condition, ShowSlide, TargetRef, TimeNode } from '../../../src/slideshow/types'
import type { PlanMediaRecord } from '../../../src/stream/show-plan'
import { placeholderOf, type MediaRef } from './bind'

/** The reader's own timing nesting limit (`READER_LIMITS.nesting`). */
export const MAX_TIMING_DEPTH = 64

/**
 * What the synchronous build needs to know of a slide's package beyond its metadata. `hiddenSpids`: its hidden shapes
 * (under `full`; projected away otherwise); `linkedRIds`: its external media relationships; `linkedSpids`: the shapes
 * they play on.
 */
export interface SlideMediaContext { hiddenSpids: ReadonlySet<string>; linkedRIds: ReadonlySet<string>; linkedSpids: ReadonlySet<string> }
export const NO_MEDIA_CONTEXT: SlideMediaContext = Object.freeze({ hiddenSpids: new Set<string>(), linkedRIds: new Set<string>(), linkedSpids: new Set<string>() })

/** One audio use, before binding. `partPath`: its part (none: refused unread). `ref`: what `mediaIdentity` hashes. */
export interface PlannedMedia { placeholder: string; ref: MediaRef; spid?: string; partPath?: string; refused?: PlanMediaRecord['refused']; hiddenShape: boolean }
export interface SlideAudioUses { records: PlannedMedia[]; snd: Map<string, string>; shapeAudio: Set<string>; videoSpids: Set<string>; transitionStart?: string }

type Visit = { sound(rId: string): void; shape(spid: string): void; video(spid: string): void }
function walkUses(meta: ShowSlide, v: Visit): void {
  const onTarget = (t: TargetRef | undefined) => { if (t && 'snd' in t) v.sound(t.snd) }
  const conds = (cs: readonly Condition[] | undefined) => { for (const c of cs ?? []) onTarget(c.target) }
  const timing = meta.timing
  if (timing && !('refused' in timing)) {
    const walk = (n: TimeNode, depth: number) => {
      if (depth > MAX_TIMING_DEPTH) return
      conds(n.ctn.start); conds(n.ctn.end); if (n.ctn.endSync) conds([n.ctn.endSync])
      if ('children' in n) { for (const c of n.children) walk(c, depth + 1); return }
      if (n.kind === 'video' && 'spid' in n.target) v.video(n.target.spid)
      if (n.kind !== 'audio') { onTarget(n.target); return }
      if ('snd' in n.target) v.sound(n.target.snd)
      else if ('spid' in n.target) v.shape(n.target.spid)
    }
    walk(timing.root, 0)
  }
  const start = meta.transition?.sound?.start
  if (start) v.sound(start.rId)
}

/** Whether the slide has any audio use (sound or audio node): then its media context is worth reading. */
export function hasAudioUse(meta: ShowSlide): boolean {
  let any = false
  walkUses(meta, { sound: () => { any = true }, shape: () => { any = true }, video: () => {} })
  return any
}

/** Slide `unit`'s audio uses. `kOf(partPath)` is the part's `k` (the build reads the probe table; it throws for an unprobed part). */
export function slideAudioUses(meta: ShowSlide, unit: number, ctx: SlideMediaContext, kOf: (partPath: string) => number, partCount: number): SlideAudioUses {
  const records = new Map<string, PlannedMedia>()
  const snd = new Map<string, string>()
  const shapeAudio = new Set<string>(), videoSpids = new Set<string>()
  const synthetic = new Map<string, number>()
  const synth = (source: string) => { let k = synthetic.get(source); if (k === undefined) synthetic.set(source, k = partCount + synthetic.size); return k }
  const add = (r: PlannedMedia) => { if (!records.has(r.placeholder)) records.set(r.placeholder, r); return r.placeholder }
  const partLess = (source: string, refused: NonNullable<PlannedMedia['refused']>, spid?: string): PlannedMedia => {
    const k = synth(source)
    return { placeholder: placeholderOf({ k, ...(spid !== undefined ? { spid } : {}) }), ref: { kind: 'audio', k, unit, source, ...(spid !== undefined ? { spid } : {}) }, ...(spid !== undefined ? { spid } : {}), refused, hiddenShape: false }
  }
  const sound = (rId: string): string => {
    const have = snd.get(rId)
    if (have) return have
    let ph: string
    if (rId.startsWith('builtIn:')) ph = add(partLess(rId, 'media-unreadable'))
    else {
      const rec = meta.media.find(m => m.rId === rId)
      if (rec) { const k = kOf(rec.partPath); ph = add({ placeholder: placeholderOf({ k }), ref: { kind: 'audio', k }, partPath: rec.partPath, hiddenShape: false }) }
      else ph = add(partLess(`snd:${rId}`, ctx.linkedRIds.has(rId) ? 'media-linked' : 'media-unreadable'))
    }
    snd.set(rId, ph)
    return ph
  }
  let transitionStart: string | undefined
  const start = meta.transition?.sound?.start
  walkUses({ ...meta, transition: undefined }, {
    sound,
    video: (spid) => { videoSpids.add(spid) },
    shape: (spid) => {
      if (!/^\d{1,10}$/.test(spid)) return
      const hiddenShape = ctx.hiddenSpids.has(spid)
      const rec = meta.media.find(m => m.spid === spid)
      if (rec) { const k = kOf(rec.partPath); add({ placeholder: placeholderOf({ k, spid }), ref: { kind: 'audio', k, unit, spid }, spid, partPath: rec.partPath, hiddenShape }) }
      else if (ctx.linkedSpids.has(spid)) add({ ...partLess(`shape:${spid}`, 'media-linked', spid), hiddenShape })
      else return
      if (!hiddenShape) shapeAudio.add(spid)
    },
  })
  if (start) transitionStart = sound(start.rId)
  return { records: [...records.values()], snd, shapeAudio, videoSpids, ...(transitionStart ? { transitionStart } : {}) }
}
