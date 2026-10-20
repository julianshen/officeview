/** One category-band rule for marks and their labels, including date/reversed axes. */
import { categoryFraction, type ResolvedAxis } from './axes'
import type { ChartGroup } from './model'

export function barGeometry(group: Pick<ChartGroup, 'grouping' | 'overlap' | 'gapWidth'> & { series: { length: number } }, category: ResolvedAxis, between: boolean, plot: { x: number; y: number; w: number; h: number }) {
  const vertical = category.orientation === 'v'
  const n = group.grouping === 'clustered' ? Math.max(1, group.series.length) : 1
  const stride = 1 - group.overlap / 100
  const span = 1 + (n - 1) * stride
  const len = vertical ? plot.h : plot.w
  const band = between ? len / category.slots : len / Math.max(1, category.slots - 1)
  const width = band / (span + group.gapWidth / 100)
  const slotStart = (idx: number): number => {
    if (!between) {
      const fraction = categoryFraction(category, idx, false)
      return (vertical ? plot.y + plot.h - fraction * plot.h : plot.x + fraction * plot.w) - band / 2
    }
    if (vertical) return category.reversed ? plot.y + band * idx : plot.y + plot.h - band * (idx + 1)
    return category.reversed ? plot.x + plot.w - band * (idx + 1) : plot.x + band * idx
  }
  return {
    width,
    start(idx: number, series: number): number {
      const slot = group.grouping === 'clustered' ? series : 0
      return slotStart(idx) + (band - width * span) / 2 + width * slot * stride
    },
  }
}
