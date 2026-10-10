/**
 * GROUP 3 UNION BRIDGE — module-private active-name identity state
 * (name-frame-design-review/UNION-CYCLE.md verbatim-adapted to existing names).
 *
 * NO public export: never re-exported from ./refs barrels, index.ts, or any
 * public surface; no EvaluationContext/NameBinding/ReferenceServices fields
 * are added. The evaluator driver installs its invocation-local active set
 * under the SAME live EvaluationContext identity; refs' resolver queries it.
 */
import type { AstNode, EvaluationContext } from './types'

type ActiveNameScope = {
  names: ReadonlySet<AstNode>
  parent?: ActiveNameScope
}

const activeNameScopes = new WeakMap<EvaluationContext, ActiveNameScope>()

/** Scoped installer: saves/restores the previous registration in the returned
 * undo closure (root try/finally; every throw/NodeMemoSuspension escape and
 * PendingDependency restores it). Exported ONLY for the evaluator/refs pair. */
export function installActiveNameScope(
  ctx: EvaluationContext,
  names: ReadonlySet<AstNode>,
): () => void {
  const previous = activeNameScopes.get(ctx)
  activeNameScopes.set(ctx, { names, parent: previous })
  return () => {
    if (previous) activeNameScopes.set(ctx, previous)
    else activeNameScopes.delete(ctx)
  }
}

/** Query: is this binding SOURCE AST active on the executing driver path? */
export function isBindingActive(ctx: EvaluationContext, sourceAst: AstNode): boolean {
  for (let scope = activeNameScopes.get(ctx); scope; scope = scope.parent) {
    if (scope.names.has(sourceAst)) return true
  }
  return false
}
