/**
 * Chart model ownership (spec 6.4/6.5; lazy chunk). A model's point storage
 * is charged to its document's ChartWorkspace under one tag, *before* the
 * storage is built, and stays charged for the model's lifetime:
 *
 * - a fill that returns a new model object takes over the tag (the charge is
 *   kept, never charged twice); storage it adds is admitted under the same tag;
 * - `releaseChartModel` (dispose, a discarded parse, a released pool entry)
 *   refunds it; a model collected without an explicit release is refunded by
 *   a FinalizationRegistry.
 *
 * The record is realm-local: a structured clone (worker -> main reply) is not
 * owned here; the receiving side charges its copy (spec 6.4).
 */
import type { ChartModel } from './model'
import type { ChartWorkspace } from './workspace'

/** A refused storage reservation: the build stops before the allocation (spec 6.4/6.5). */
export class ChartRefusal extends Error {}
interface Owner { workspace: ChartWorkspace; tag: string }
const releaseListeners = new Set<(model: ChartModel, next?: ChartModel) => void>()
/** Lazy sidecars register their disposal hook without making parse depend on layout. */
export function onChartModelRelease(listener: (model: ChartModel, next?: ChartModel) => void): void { releaseListeners.add(listener) }
const owners = new WeakMap<ChartModel, Owner>()
const collected = typeof FinalizationRegistry === 'function' ? new FinalizationRegistry<Owner>(owner => owner.workspace.release(owner.tag)) : undefined
let tags = 0

/** A fresh model tag in `workspace`. */
export const modelOwner = (workspace: ChartWorkspace): Owner => ({ workspace, tag: `chart-model#${++tags}` })
export function adoptChartModel(model: ChartModel, owner: Owner): void {
  owners.set(model, owner)
  collected?.register(model, owner, model)
}
export const chartModelOwner = (model: ChartModel): Owner | undefined => owners.get(model)
/** `next` (a fill's new model object) takes over `previous`'s charge. */
export function succeedChartModel(previous: ChartModel, next: ChartModel): void {
  const owner = owners.get(previous)
  if (!owner || previous === next) return
  for (const release of releaseListeners) release(previous, next)
  owners.delete(previous)
  collected?.unregister(previous)
  adoptChartModel(next, owner)
}
/**
 * Bytes a chart operation published into a longer-lived source structure (a formatter diagnostic on the sheet):
 * transferred from the operation's tag to one owner per target, which lives as long as the target (the sheet,
 * hence the document), not the chart model. Refunded once: `releaseChartRetained` (the document's release) or
 * the target's collection.
 */
const retained = new WeakMap<object, Owner>()
export function retainWithTarget(target: object, workspace: ChartWorkspace, fromTag: string, bytes: number): void {
  let owner = retained.get(target)
  if (!owner || owner.workspace !== workspace) {
    if (owner) releaseChartRetained(target)
    owner = modelOwner(workspace)
    retained.set(target, owner)
    collected?.register(target, owner, target)
  }
  workspace.transfer(fromTag, owner.tag, bytes)
}
export function releaseChartRetained(target: object): void {
  const owner = retained.get(target)
  if (!owner) return
  retained.delete(target)
  collected?.unregister(target)
  owner.workspace.release(owner.tag)
}
/** Refunds a model's storage (idempotent). */
export function releaseChartModel(model: ChartModel): void {
  for (const release of releaseListeners) release(model)
  const owner = owners.get(model)
  if (!owner) return
  owners.delete(model)
  collected?.unregister(model)
  owner.workspace.release(owner.tag)
}
/**
 * Retained chart formatting (A4a, spec 5.4/5.11): fills, gradient stops, lines, markers and per-point overrides are
 * priced here and charged to the model's tag *before* each object is built (the parser passes its `charge`). A
 * gradient keeps at most CHART_GRADIENT_STOP_LIMIT stops, each charged, so a long `gsLst` neither grows unbounded
 * nor costs a flat estimate.
 */
export const CHART_GRADIENT_STOP_LIMIT = 32
export const CHART_STYLE_BYTES = Object.freeze({ fill: 96, stop: 64, line: 112, dash: 64, marker: 96, point: 160, color: 32 })
export type ChartCharge = (bytes: number) => void
export const chargeChartFill = (charge: ChartCharge | undefined, stops = 0): void => charge?.(CHART_STYLE_BYTES.fill + Math.min(stops, CHART_GRADIENT_STOP_LIMIT) * CHART_STYLE_BYTES.stop)
export const chargeChartLine = (charge: ChartCharge | undefined): void => charge?.(CHART_STYLE_BYTES.line)
export const chargeChartMarker = (charge: ChartCharge | undefined): void => charge?.(CHART_STYLE_BYTES.marker)
export const chargeChartPoint = (charge: ChartCharge | undefined): void => charge?.(CHART_STYLE_BYTES.point)
export const chargeChartColor = (charge: ChartCharge | undefined): void => charge?.(CHART_STYLE_BYTES.color)
/**
 * Trendlines (A6a, spec 5.6): a parsed record (plus its label format and name text) is charged at parse; a prepared
 * fit retains its coefficients, at most 512 samples (x, y) and its label, priced by `preparedTrendlineBytes`.
 */
export const CHART_TRENDLINE_BYTES = Object.freeze({ spec: 160, fit: 192 })
export const chargeChartTrendline = (charge: ChartCharge | undefined, textUnits = 0): void => charge?.(CHART_TRENDLINE_BYTES.spec + (textUnits ? textUnits * 2 + 64 : 0))
