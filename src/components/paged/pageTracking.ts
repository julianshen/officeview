/**
 * The current page: the one with the largest visible pixel height in the
 * viewport, ties to the lower index (spec §6.11, R7). Pure.
 */
export { pageOffsets } from '../scrollPan'

export function mostVisiblePage(o: { top: number[]; height: number[] }, viewTop: number, viewHeight: number): number {
  const viewBottom = viewTop + viewHeight
  let best = 0
  let bestPx = -1
  for (let i = 0; i < o.top.length; i++) {
    const top = o.top[i]
    if (top >= viewBottom) break
    const px = Math.min(viewBottom, top + o.height[i]) - Math.max(viewTop, top)
    if (px > bestPx) { best = i; bestPx = px }
  }
  return best
}
