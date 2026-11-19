/**
 * The plan-backed slide source of a streamed show (piece 4 Task 11; spec §5, plan D5, D13, D16, D25):
 * what the shared driver reads instead of the package.
 *
 * - The source speaks streamed indices (hidden units are not in `unitOf`), so nothing is hidden.
 * - A slide whose unit is neither `anim` nor `media` in the manifest (D5; piece 5 D19) is a static `ShowSlide`: no
 *   timing, no transition; it plays as its piece-3 still.
 * - A slide whose bundle is held (`bundles(i)`) plays the plan's `show` (its timing tree, transition and builds)
 *   with `notes: []` (the piece-3 image service) and `links: []` (link regions are the stage's, §5.1). Without a
 *   bundle (refused to a still, failed) it is static as well.
 * - Piece 5 (D16, D18, D22, X12), with `media` on (`can.media`): `ShowSlide.media` holds one record per plan audio
 *   record, `{rId: mediaId, partPath: 'remote:' + mediaId, contentType: 'audio/mp4', declaredBytes: 0, spid}`; the
 *   transition's start sound names its record by `rId = mediaId` (`localTransition`), so the unchanged
 *   `TransitionSoundDeps.open(rec)` finds it; `mediaClock(i)` is the plan's records on the show clock (no record is
 *   ever a video's, D22). Built from any bundle's plan, so a still's sound-only plan would play unchanged. Media off:
 *   `media: []` and no clock (the driver ends media nodes as piece 4 does, D18).
 * - The deck's show settings (`showAnimation`, `useTimings`) are the plan's `options` (D13), read
 *   through getters: the driver reads them after a slide is acquired, so the first plan sets them.
 *
 * The show's byte budget (128 MiB, eviction of non-current bundles, `show-budget` for a current
 * bundle over it on its own) lives in the fetcher that holds the bundles (show-fetch.ts); waiting
 * for a bundle, pinning it while the slide is shown and charging the text fan are the show's
 * (show.ts), which wraps `acquire`.
 *
 * Package-free (D25): type imports, plus `localTransition` from the DOM-free wire codec (`show-plan.ts`) and the
 * DOM-free media clock (`slideshow/engine/media-ends.ts`).
 */
import { mediaClockOf, type MediaClock } from '../../slideshow/engine/media-ends'
import type { SlideSource } from '../../slideshow/player/driver'
import type { MediaRecord, ShowSlide } from '../../slideshow/types'
import type { Manifest } from '../manifest'
import { localTransition, type PlanMediaRecord } from '../show-plan'
import type { ShowBundle } from './show-bundle'

/** What a remote media record stands for (`remoteRecordOf`): its server unit and the plan's record. */
export interface RemoteRecordInfo { unit: number; record: PlanMediaRecord }
const remoteInfo = new WeakMap<MediaRecord, RemoteRecordInfo>()
/** The plan record and server unit behind a `ShowSlide.media` record this source built (undefined for any other). */
export const remoteRecordOf = (rec: MediaRecord): RemoteRecordInfo | undefined => remoteInfo.get(rec)

/** D16: the local record a plan audio record plays as. */
export function remoteMediaRecord(r: PlanMediaRecord, unit: number): MediaRecord {
  const rec: MediaRecord = { rId: r.mediaId, partPath: `remote:${r.mediaId}`, contentType: 'audio/mp4', declaredBytes: 0, ...(r.spid !== undefined ? { spid: r.spid } : {}) }
  remoteInfo.set(rec, { unit, record: r })
  return rec
}

/** D19: a unit the show asks for a plan bundle (an animated slide, or one that ships audio or a transition sound). */
export const planUnit = (u: { anim?: true; media?: true } | undefined): boolean => !!u?.anim || !!u?.media

export function createPlanSource(o: { manifest: Manifest; unitOf: readonly number[]; bundles: (streamed: number) => ShowBundle | undefined; loop: boolean; kiosk: boolean; media?: boolean }): SlideSource {
  /** The deck's show settings, from the last plan acquired (every plan carries the same deck settings). */
  let options: { showAnimation: boolean; useTimings: boolean } | undefined
  const settings: SlideSource['settings'] = {
    get showAnimation() { return options?.showAnimation ?? true },
    get useTimings() { return options?.useTimings ?? true },
    loop: o.loop || o.kiosk,
    type: o.kiosk ? 'kiosk' : 'present',
  }
  const bundleOf = (i: number) => {
    const unit = o.unitOf[i]
    return unit !== undefined && planUnit(o.manifest.units.list[unit]) ? o.bundles(i) : undefined
  }
  /** The records of the plan last read for each streamed slide (one object per record, as the player keys them). */
  const records = new Map<number, { plan: ShowBundle['plan']; media: MediaRecord[] }>()
  const mediaOf = (i: number, plan: ShowBundle['plan']): MediaRecord[] => {
    if (!o.media) return []
    const had = records.get(i)
    if (had?.plan === plan) return had.media
    const media = (plan.media ?? []).map((r) => remoteMediaRecord(r, o.unitOf[i]!))
    records.set(i, { plan, media })
    return media
  }
  const slideOf = (i: number): ShowSlide => {
    const bundle = bundleOf(i)
    if (!bundle) return { index: i, hidden: false, timing: undefined, notes: [], links: [], media: [], builds: [] }
    const { plan } = bundle
    options = plan.options
    const { timing, transition, builds } = plan.show
    return { index: i, hidden: false, timing, ...(transition ? { transition: localTransition(transition) } : {}), builds: builds ?? [], notes: [], links: [], media: mediaOf(i, plan) }
  }
  return {
    count: o.unitOf.length,
    settings,
    hidden: () => false,
    async acquire(i, signal) {
      if (signal.aborted) throw signal.reason ?? new Error('aborted')
      if (!Number.isInteger(i) || i < 0 || i >= o.unitOf.length) throw new RangeError(`no streamed slide ${i}`)
      return { slide: slideOf(i), release() {} }
    },
    ...(o.media ? {
      mediaClock(i: number): MediaClock | undefined {
        const plan = bundleOf(i)?.plan
        // A slide with no plan (a still) has no media node to end; one with a plan always gets a clock, even empty.
        if (!plan) return undefined
        const timing = plan.show.timing
        return mediaClockOf([timing && !('refused' in timing) ? timing.root : undefined], plan.media ?? [])
      },
    } : {}),
  }
}
