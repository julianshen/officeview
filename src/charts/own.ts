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
  const owner = owners.get(model)
  if (!owner) return
  owners.delete(model)
  collected?.unregister(model)
  owner.workspace.release(owner.tag)
}
