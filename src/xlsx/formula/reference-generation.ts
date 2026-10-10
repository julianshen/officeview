/** Internal generation witnesses and branded invalidations; no public barrel. */
import type { AstNode, EvaluationContext, EvaluationValue, RefArea, ResolvedRef, ReferenceServices } from './types'

export const MAX_REFERENCE_GENERATION_RESTARTS = 8

class ReferenceGenerationChanged extends Error {
  constructor(bound: number, current: number, kind: 'reference' | 'cursor') {
    super(`stale ${kind} generation (bound ${bound}, current ${current})`)
    this.name = 'ReferenceGenerationChanged'
  }
}
const ownedInvalidations = new WeakSet<object>()
const serviceGenerations = new WeakMap<ReferenceServices, () => number>()

export function referenceGenerationChanged(bound: number, current: number, kind: 'reference' | 'cursor' = 'reference'): Error {
  const error = new ReferenceGenerationChanged(bound, current, kind)
  ownedInvalidations.add(error)
  return error
}
/** Identity branding cannot be forged with a user's error name/message/prototype. */
export function isReferenceGenerationChanged(error: unknown): error is Error {
  return typeof error === 'object' && error !== null && ownedInvalidations.has(error)
}
export function registerReferenceGeneration(services: ReferenceServices, generation: () => number): void {
  serviceGenerations.set(services, generation)
}
export function currentReferenceGeneration(services: ReferenceServices): number | undefined {
  return serviceGenerations.get(services)?.()
}


/** Only evaluator-created scoped memos are registered, with exact context
 * identity. Deliberately supplied maps have no registration. */
interface OwnedMemoScope {
  ctx: EvaluationContext
  owner: object
  epoch: () => number
  restart: () => never
}
const ownedMemoScopes = new WeakMap<WeakMap<AstNode, EvaluationValue>, OwnedMemoScope>()

export function installOwnedMemoScope(
  ctx: EvaluationContext,
  memo: WeakMap<AstNode, EvaluationValue>,
  owner: object,
  epoch: () => number,
  restart: () => never,
): () => void {
  const previous = ownedMemoScopes.get(memo)
  ownedMemoScopes.set(memo, { ctx, owner, epoch, restart })
  return () => {
    if (previous) ownedMemoScopes.set(memo, previous)
    else ownedMemoScopes.delete(memo)
  }
}
function ownedMemoScope(ctx: EvaluationContext | undefined): OwnedMemoScope | undefined {
  const scope = ctx?.nodeValues ? ownedMemoScopes.get(ctx.nodeValues) : undefined
  return scope?.ctx === ctx ? scope : undefined
}
/** Inherit the owned scope registered for `sourceMemo` onto a privately created
 * `targetMemo` so a mode-switched child view participates in the SAME root
 * owner/epoch/matching restart. A caller-provided map has no scope and is
 * therefore never registered as owned. */
export function inheritOwnedMemoScope(
  ctx: EvaluationContext,
  sourceMemo: WeakMap<AstNode, EvaluationValue>,
  targetMemo: WeakMap<AstNode, EvaluationValue>,
): () => void {
  const scope = ownedMemoScopes.get(sourceMemo)
  if (!scope || scope.ctx !== ctx) return () => {}
  const previous = ownedMemoScopes.get(targetMemo)
  ownedMemoScopes.set(targetMemo, scope)
  return () => {
    if (previous) ownedMemoScopes.set(targetMemo, previous)
    else ownedMemoScopes.delete(targetMemo)
  }
}
export function ownedMemoEpoch(ctx: EvaluationContext | undefined): { owner: object; epoch: number } | undefined {
  const scope = ownedMemoScope(ctx)
  return scope ? { owner: scope.owner, epoch: scope.epoch() } : undefined
}
/** A genuine invalidation requests a cooperative unwind to the matching owner.
 * The owner callback throws an invocation-private token; no user error matching. */
export function requestOwnedMemoRestart(ctx: EvaluationContext | undefined): void {
  ownedMemoScope(ctx)?.restart()
}


/** Generation-bound region creators stay private. Conditional optional-target
 * resizing can preserve the exact store/sheet identity without guessing names. */
const referenceCreators = new WeakMap<ResolvedRef, (areas: readonly RefArea[]) => ResolvedRef>()
export function registerReferenceCreator(ref: ResolvedRef, create: (areas: readonly RefArea[]) => ResolvedRef): void {
  referenceCreators.set(ref, create)
}
export function createReferenceRegion(ref: ResolvedRef, areas: readonly RefArea[]): ResolvedRef | undefined {
  return referenceCreators.get(ref)?.(areas)
}
