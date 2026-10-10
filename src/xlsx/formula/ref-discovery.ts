/** Internal reference-topology discovery; never exported by a public barrel. */
import type { AstNode, EvaluationContext, EvaluationError, NameBinding, NameSyntax, ReferenceNode } from './types'

/** Find 3D syntax in reference topology with one fixed use context. Bound source
 * identity distinguishes local scopes and terminates repeated alias branches.
 * Value expressions, including calls and their lazy branches, are not visited. */
export function contains3d(
  root: ReferenceNode,
  bindName: (name: NameSyntax, ctx: EvaluationContext) => NameBinding | EvaluationError,
  ctx: EvaluationContext,
): boolean {
  const pending: AstNode[] = [root]
  const seenBindings = new Set<AstNode>()
  while (pending.length > 0) {
    const node = pending.pop()!
    switch (node.type) {
      case 'ref3d':
        return true
      case 'union':
        for (let i = node.refs.length - 1; i >= 0; i--) pending.push(node.refs[i])
        break
      case 'intersect':
        pending.push(node.right)
        pending.push(node.left)
        break
      case 'name': {
        const binding = bindName(node.ref, ctx)
        if ('kind' in binding) break
        const source = binding.ast
        if (seenBindings.has(source)) break
        seenBindings.add(source)
        pending.push(source)
        break
      }
    }
  }
  return false
}

/** Private metadata on cloned value-name reference nodes. The original source
 * AST and public AST/context shapes stay unchanged. Consumers query only when
 * a node is actually reached, so lazy branches and cycle precedence survive. */
interface RelativeNodeGate { feature: 'relative-name' | 'relative-name-context-missing'; message: string }
const relativeNodeGates = new WeakMap<AstNode, RelativeNodeGate>()
export function registerRelativeNodeGate(node: AstNode, gate: RelativeNodeGate): void {
  relativeNodeGates.set(node, gate)
}
export function relativeNodeGate(node: AstNode): RelativeNodeGate | undefined {
  return relativeNodeGates.get(node)
}
