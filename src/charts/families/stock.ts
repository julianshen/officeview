/**
 * Stock charts (spec 4.1; lazy chart chunk).
 *
 * Supported types:
 * - HLC (High-Low-Close): 3 series
 * - OHLC (Open-High-Low-Close): 4 series
 * - Volume-HLC: 1 bar group + 1 stock group with 3 series
 * - Volume-OHLC: 1 bar group + 1 stock group with 4 series
 *
 * Rules:
 * - Volume is the series of the bar group in the same plot.
 * - Stock series are, in order: Open (only when 4 series), High, Low, Close.
 * - Axes come from authored axId bindings, never assumed.
 * - Draw hiLowLines and upDownBars (up/down fills from c:upBars/c:downBars).
 * - A series-count/role mismatch becomes a card with a reason.
 */
import { attrs, getChildren, type XmlNode } from '../../core/xml'
import type { ThemeContext } from '../../drawing/style'
import type { ChartFillStyle } from '../fills'
import { parseFillStyle } from '../fills'
import type { ChartGroup, ChartUpDownBars, ChartUpDownBar } from '../model'
import { parseLineStyle, applyLineStyle, type ChartLineStyle } from '../style'
import { barGeometry } from '../barGeometry'
import type { GroupAxes, ResolvedAxis } from '../axes'
import { chargeChartLine, type ChartCharge } from '../own'

const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const val = (node: XmlNode | undefined) => attrs(node).val

interface Box { x: number; y: number; w: number; h: number }

export interface StockSeriesMapping<T> {
  open?: T
  high: T
  low: T
  close: T
}

/** Map series in order: Open (only when 4 series), High, Low, Close. */
export function mapStockSeries<T>(series: readonly T[]): StockSeriesMapping<T> | undefined {
  if (series.length === 3) {
    return { high: series[0], low: series[1], close: series[2] }
  }
  if (series.length === 4) {
    return { open: series[0], high: series[1], low: series[2], close: series[3] }
  }
  return undefined
}

/** Reads stock group options (`hiLowLines`, `upDownBars`, `gapWidth`). */
export function readStockOptions(
  node: XmlNode,
  theme?: ThemeContext,
  charge?: ChartCharge
): Pick<ChartGroup, 'hiLowLines' | 'upDownBars'> & { gapWidth?: number } {
  const lineOf = (spPr: XmlNode | undefined) => parseLineStyle(child(spPr, 'ln'), theme, charge)

  // 1. hiLowLines
  const hiLowNode = child(node, 'hiLowLines')
  const hiLowLines: ChartLineStyle | undefined = hiLowNode
    ? lineOf(child(hiLowNode, 'spPr')) ?? (charge ? (chargeChartLine(charge), { noFill: false, color: '000000', width: 1 }) : { noFill: false, color: '000000', width: 1 })
    : undefined

  // 2. upDownBars
  const upDownNode = child(node, 'upDownBars')
  let upDownBars: ChartUpDownBars | undefined
  let gapWidth: number | undefined

  if (upDownNode) {
    const gapRaw = val(child(upDownNode, 'gapWidth'))
    gapWidth = gapRaw !== undefined && Number.isFinite(Number(gapRaw))
      ? Math.max(0, Math.min(500, Number(gapRaw)))
      : 150

    const parseBar = (barNode: XmlNode | undefined): ChartUpDownBar => {
      if (!barNode) return {}
      const spPr = child(barNode, 'spPr')
      const fill = parseFillStyle(spPr, theme, charge)
      const line = lineOf(spPr)
      return {
        ...(fill ? { fill } : {}),
        ...(line ? { line } : {}),
      }
    }

    const upBars = parseBar(child(upDownNode, 'upBars'))
    const downBars = parseBar(child(upDownNode, 'downBars'))
    if (charge) charge(64)
    upDownBars = { upBars, downBars }
  }

  return {
    ...(hiLowLines ? { hiLowLines } : {}),
    ...(upDownBars ? { upDownBars } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
  }
}

export interface StockSeriesLike {
  source: import('../model').ChartSeries
  color: string
  markers?: boolean
  plot(slot: number, x?: boolean): number | 'gap' | 'span'
}

export interface StockPaintKit {
  plot: Box
  slots: number
  source(slot: number): number
  admitPath(cost: number): boolean
  fillWith(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, color: string, box: Box): boolean
  strokeBorder(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, color: string, box: Box, rect?: Box): void
  drawMarker(ctx: CanvasRenderingContext2D, series: StockSeriesLike, pointIndex: number, x: number, y: number, r: number): void
  plotted(v: number | undefined): number | undefined
  at(v: number): number
  categoryIndex(axis: ResolvedAxis, slot: number): number
}

/** Paints a stock group on its axes. */
export function paintStock(
  ctx: CanvasRenderingContext2D,
  group: Pick<ChartGroup, 'hiLowLines' | 'upDownBars' | 'gapWidth'>,
  series: readonly StockSeriesLike[],
  axes: GroupAxes,
  kit: StockPaintKit
): void {
  const mapping = mapStockSeries(series)
  if (!mapping) return

  const catAxis = axes.category ?? axes.x!
  const gap = group.gapWidth ?? 150
  const geom = barGeometry({ grouping: 'standard', overlap: 0, gapWidth: gap, series: { length: 1 } }, catAxis, axes.between, kit.plot)

  for (let slot = 0; slot < kit.slots; slot++) {
    const idx = kit.categoryIndex(catAxis, slot)
    if (!Number.isFinite(idx)) continue
    const xCenter = geom.start(idx, 0) + geom.width / 2

    const H = mapping.high.plot(slot)
    const L = mapping.low.plot(slot)
    const C = mapping.close.plot(slot)
    const O = mapping.open?.plot(slot)

    const hVal = typeof H === 'number' && kit.plotted(H) !== undefined ? H : undefined
    const lVal = typeof L === 'number' && kit.plotted(L) !== undefined ? L : undefined
    const cVal = typeof C === 'number' && kit.plotted(C) !== undefined ? C : undefined
    const oVal = typeof O === 'number' && kit.plotted(O) !== undefined ? O : undefined

    // 1. hiLowLines: connects Low to High
    if (group.hiLowLines && !group.hiLowLines.noFill && hVal !== undefined && lVal !== undefined) {
      if (kit.admitPath(2)) {
        const yH = kit.at(hVal), yL = kit.at(lVal)
        ctx.save()
        try {
          ctx.strokeStyle = group.hiLowLines.color ? `#${group.hiLowLines.color}` : '#000000'
          ctx.lineWidth = group.hiLowLines.width ?? 1
          if (applyLineStyle(ctx, group.hiLowLines, kit.plot)) {
            ctx.beginPath()
            ctx.moveTo(xCenter, yL)
            ctx.lineTo(xCenter, yH)
            ctx.stroke()
          }
        } finally {
          ctx.restore()
        }
      }
    }

    // 2. upDownBars (OHLC)
    if (mapping.open && group.upDownBars && oVal !== undefined && cVal !== undefined) {
      if (kit.admitPath(5)) {
        const yO = kit.at(oVal), yC = kit.at(cVal)
        const yTop = Math.min(yO, yC)
        const yBot = Math.max(yO, yC)
        const h = Math.max(1, yBot - yTop)
        const barBox: Box = { x: geom.start(idx, 0), y: yTop, w: geom.width, h }

        const isUp = cVal >= oVal
        const bar = isUp ? group.upDownBars.upBars : group.upDownBars.downBars
        const defaultFill = isUp ? 'FFFFFF' : '000000'
        const defaultLineColor = '000000'

        if (kit.fillWith(ctx, bar?.fill, defaultFill, barBox)) {
          ctx.fillRect(barBox.x, barBox.y, barBox.w, barBox.h)
        }
        const lineStyle = bar?.line ?? { color: defaultLineColor, width: 1, noFill: false }
        kit.strokeBorder(ctx, lineStyle, lineStyle.color ?? defaultLineColor, barBox, barBox)
      }
    }

    // 3. HLC Close tick (when no Open series)
    if (!mapping.open && cVal !== undefined) {
      const yC = kit.at(cVal)
      if (mapping.close.markers || mapping.close.source.markerStyle) {
        kit.drawMarker(ctx, mapping.close, kit.source(slot), xCenter, yC, 3)
      } else if (kit.admitPath(2)) {
        ctx.save()
        try {
          const line = group.hiLowLines ?? { color: '000000', width: 1, noFill: false }
          ctx.strokeStyle = line.color ? `#${line.color}` : '#000000'
          ctx.lineWidth = line.width ?? 1
          if (applyLineStyle(ctx, line, kit.plot)) {
            ctx.beginPath()
            ctx.moveTo(xCenter, yC)
            ctx.lineTo(xCenter + geom.width / 2, yC)
            ctx.stroke()
          }
        } finally {
          ctx.restore()
        }
      }
    }

    // 4. Authored markers on individual series
    for (const s of series) {
      if (s.markers && s.source.markerStyle && s !== (!mapping.open ? mapping.close : undefined)) {
        const v = s.plot(slot)
        if (typeof v === 'number' && kit.plotted(v) !== undefined) {
          kit.drawMarker(ctx, s, kit.source(slot), xCenter, kit.at(v), 3)
        }
      }
    }
  }
}
