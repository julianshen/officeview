/**
 * Show order (slideshow spec 4.2 starting slide, 4.3 hidden slides, 4.4
 * `p:sldRg` and loop). Pure: deck indices in, deck indices out.
 *
 * The visible order is the deck order without hidden slides and without
 * slides outside the range. Next and previous move through it; a slide off
 * the order (a hidden slide shown by goTo, number entry or a 'current' start)
 * moves to the nearest visible slide after or before its own index.
 */
import type { ShowIndex } from '../types'

/** Deck indices that play, in deck order: not hidden, inside `p:sldRg` (already 0-based and validated by the reader). */
export function visibleOrder(index: ShowIndex): number[] {
  const range = index.settings.range
  const out: number[] = []
  for (const s of index.slides) {
    if (s.hidden) continue
    if (range && (s.index < range.st || s.index > range.end)) continue
    out.push(s.index)
  }
  return out
}

/** The first visible slide after deck index `from`, or undefined at the end of the order. */
export function nextVisible(order: readonly number[], from: number): number | undefined {
  return order.find((i) => i > from)
}

/** The last visible slide before deck index `from`, or undefined at the start of the order. */
export function previousVisible(order: readonly number[], from: number): number | undefined {
  for (let k = order.length - 1; k >= 0; k--) if (order[k] < from) return order[k]
  return undefined
}
