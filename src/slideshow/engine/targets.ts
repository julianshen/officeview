/**
 * Timing targets of a slide's timing tree (moved from `player/stage.ts`, piece 4 plan D22): which targets
 * animate, which animate colour, and which start hidden with their entrance behaviours. Reads only the
 * `ShowSlide`, so a package-free player derives `hidden` and `entrances` itself.
 */
import { targetKey } from '../target'
import type { ShowSlide, TargetKey, TimeNode } from '../types'

function isContainer(n: TimeNode): n is Extract<TimeNode, { children: TimeNode[] }> {
  return n.kind === 'par' || n.kind === 'seq' || n.kind === 'excl'
}

const COLOUR_ATTR = new Set(['fillcolor', 'stroke.color', 'style.color'])

/**
 * Animated targets, colour-animated targets, and the initially hidden ones with their entrance behaviours.
 *
 * Initial hiding (spec 6.2 "initial"; S1-20 review I2): a target with any behaviour under an entrance preset is
 * hidden until one of those entrance behaviours begins, unless a non-entrance behaviour on it runs before its
 * first entrance in time. "Before in time" is read at click-step granularity: a non-entrance behaviour in a
 * strictly earlier main-sequence step (or before the first click) reveals the target from the start, the way
 * PowerPoint shows a shape whose first effect is an emphasis. Behaviours in the same step run concurrently, and
 * interactive sequences run only when triggered, so neither reveals the target before its entrance.
 */
export function timingTargets(slide: ShowSlide): {
  animated: Set<TargetKey>; colour: Set<TargetKey>; hidden: Set<TargetKey>; entrances: Map<TargetKey, Set<TimeNode>>
} {
  const animated = new Set<TargetKey>(), colour = new Set<TargetKey>(), hidden = new Set<TargetKey>()
  const entrances = new Map<TargetKey, Set<TimeNode>>()
  const timing = slide.timing
  if (!timing || 'refused' in timing) return { animated, colour, hidden, entrances }
  const firstEntrance = new Map<TargetKey, number>()
  const firstOther = new Map<TargetKey, number>()
  const walk = (n: TimeNode, cls: string | undefined, step: number, depth: number): void => {
    if (depth > 128) return
    const here = n.ctn.presetClass ?? cls
    if (isContainer(n)) {
      // The main sequence's children are its steps, in order; everything off it (interactive) is "never before".
      const main = n.kind === 'seq' && n.ctn.nodeType === 'mainSeq'
      n.children.forEach((c, i) => walk(c, here, main ? i : n.ctn.nodeType === 'interactiveSeq' ? Infinity : step, depth + 1))
      return
    }
    if (!('spid' in n.target)) return
    const key = targetKey(n.target)
    animated.add(key)
    if (n.kind === 'animClr' || n.attrNames.some(a => COLOUR_ATTR.has(a))) colour.add(key)
    const first = here === 'entr' ? firstEntrance : firstOther
    first.set(key, Math.min(first.get(key) ?? Infinity, step))
    if (here === 'entr') {
      let set = entrances.get(key)
      if (!set) entrances.set(key, set = new Set())
      set.add(n)
    }
  }
  walk(timing.root, undefined, Infinity, 0)
  for (const [key, at] of firstEntrance) if (!((firstOther.get(key) ?? Infinity) < at)) hidden.add(key)
  for (const key of [...entrances.keys()]) if (!hidden.has(key)) entrances.delete(key)
  return { animated, colour, hidden, entrances }
}
