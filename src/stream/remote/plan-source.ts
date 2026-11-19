/**
 * The plan-backed slide source of a streamed show (piece 4 Task 11; spec §5, plan D5, D13, D16, D25):
 * what the shared driver reads instead of the package.
 *
 * - The source speaks streamed indices (hidden units are not in `unitOf`), so nothing is hidden.
 * - A slide whose unit is not `anim` in the manifest (D5) is a static `ShowSlide`: no timing, no
 *   transition; it plays as its piece-3 still.
 * - An animated slide whose bundle is held (`bundles(i)`) plays the plan's `show` (its timing tree,
 *   transition and builds) with `notes: []` (the piece-3 image service), `links: []` (link regions
 *   are the stage's, §5.1) and `media: []` (piece 5; the driver ends media nodes at once, D16).
 *   Without a bundle (refused to a still, failed) it is static as well.
 * - The deck's show settings (`showAnimation`, `useTimings`) are the plan's `options` (D13), read
 *   through getters: the driver reads them after a slide is acquired, so the first plan sets them.
 *
 * The show's byte budget (128 MiB, eviction of non-current bundles, `show-budget` for a current
 * bundle over it on its own) lives in the fetcher that holds the bundles (show-fetch.ts); waiting
 * for a bundle, pinning it while the slide is shown and charging the text fan are the show's
 * (show.ts), which wraps `acquire`.
 *
 * Package-free (D25): type imports only.
 */
import type { SlideSource } from '../../slideshow/player/driver'
import type { ShowSlide } from '../../slideshow/types'
import type { Manifest } from '../manifest'
import { localTransition } from '../show-plan'
import type { ShowBundle } from './show-bundle'

export function createPlanSource(o: { manifest: Manifest; unitOf: readonly number[]; bundles: (streamed: number) => ShowBundle | undefined; loop: boolean; kiosk: boolean }): SlideSource {
  /** The deck's show settings, from the last plan acquired (every plan carries the same deck settings). */
  let options: { showAnimation: boolean; useTimings: boolean } | undefined
  const settings: SlideSource['settings'] = {
    get showAnimation() { return options?.showAnimation ?? true },
    get useTimings() { return options?.useTimings ?? true },
    loop: o.loop || o.kiosk,
    type: o.kiosk ? 'kiosk' : 'present',
  }
  const slideOf = (i: number): ShowSlide => {
    const unit = o.unitOf[i]
    const bundle = unit !== undefined && o.manifest.units.list[unit]?.anim ? o.bundles(i) : undefined
    if (!bundle) return { index: i, hidden: false, timing: undefined, notes: [], links: [], media: [], builds: [] }
    const { plan } = bundle
    options = plan.options
    const { timing, transition, builds } = plan.show
    return { index: i, hidden: false, timing, ...(transition ? { transition: localTransition(transition) } : {}), builds: builds ?? [], notes: [], links: [], media: [] }
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
  }
}
