/**
 * Relative-reference translation of a compiled CF formula (Stage 5 Task 34, Ruling S5-C). Its own module so the
 * per-cell evaluator crosses a module boundary to translate (tests replace the translation to prove the view-engine
 * tree guard is wired at the call site).
 */
import type { AstNode } from '../formula/types'
import { CF_FORMULA_NODE_BYTES, type CfMemory } from './limits'
import type { CfFormula, CfLease, CfOutcome } from './model'

const MAX_ROW = 1048576, MAX_COL = 16384

/** How a relative reference shifted past the sheet edge reads (Ruling S5-C): wrap modulo the grid, or `#REF!`. Absent
 * (no `formula:relative-wrap` evidence) a translation that leaves the sheet is unresolved. */
export type CfEdgeMode = 'wrap' | 'ref-error'
export const EDGE_UNMEASURED: unique symbol = Symbol('cf-edge-unmeasured')
type Shifted = number | 'ref'
function shiftIndex(index: number, delta: number, pinned: boolean, max: number, edge: CfEdgeMode | undefined): Shifted {
  if (pinned || delta === 0) return index
  const moved = index + delta
  if (moved >= 0 && moved < max) return moved
  if (edge === 'wrap') return ((moved % max) + max) % max
  if (edge === 'ref-error') return 'ref'
  throw EDGE_UNMEASURED
}
/** A span's two endpoints shift together. Under `wrap` an endpoint pair that lands on two sides of the edge has no single
 * rectangle (Excel's reading of such a range is unmeasured), so it is unresolved rather than guessed. */
export function shiftPair(a: number, b: number, delta: number, pinned: boolean, max: number, edge: CfEdgeMode | undefined): [Shifted, Shifted] {
  const x = shiftIndex(a, delta, pinned, max, edge), y = shiftIndex(b, delta, pinned, max, edge)
  if (x === 'ref' || y === 'ref') return ['ref', 'ref']
  if (edge === 'wrap' && !pinned && ((a + delta >= 0 && a + delta < max) !== (b + delta >= 0 && b + delta < max))) throw EDGE_UNMEASURED
  return [x, y]
}
const refErrorNode = (): AstNode => ({ type: 'error', error: '#REF!' })

/**
 * Translate a compiled formula by `(dRow, dCol)` from its anchor (Ruling S5-C). `$`-pinned axes never move. Subtrees with
 * nothing to move are shared, not copied, so hoisted nodes keep their identity. `substitute` replaces a node (a hoisted
 * aggregate's value) before it is walked. The output tree is reserved on `memory` BEFORE any node is built, at
 * `f.nodes x CF_FORMULA_NODE_BYTES` (a translated tree never has more nodes than the compiled one, and each node is at
 * most CF_FORMULA_NODE_BYTES); the lease keeps what was actually rebuilt.
 */
export function translateCfFormula(f: CfFormula, dRow: number, dCol: number, edge: CfEdgeMode | undefined, memory: CfMemory,
  substitute?: (node: AstNode) => AstNode | undefined): CfOutcome<CfLease<AstNode>> {
  if (!Number.isSafeInteger(dRow) || !Number.isSafeInteger(dCol)) return { kind: 'unresolved', reason: 'cf-outside-preparation' }
  const hold = memory.reserve(f.nodes * CF_FORMULA_NODE_BYTES)
  if (hold.kind !== 'reserved') return { kind: 'unresolved', reason: 'cf-resource-limit' }
  let built = 0
  const walk = (node: AstNode): AstNode => {
    const replaced = substitute?.(node)
    if (replaced) { built++; return replaced }
    switch (node.type) {
      case 'cell': {
        const r = node.ref
        const row = shiftIndex(r.row, dRow, r.absRow, MAX_ROW, edge), col = shiftIndex(r.col, dCol, r.absCol, MAX_COL, edge)
        if (row === 'ref' || col === 'ref') { built++; return refErrorNode() }
        if (row === r.row && col === r.col) return node
        built++
        return { type: 'cell', ref: { ...r, row, col } }
      }
      case 'range': {
        const { from, to } = node.ref
        const [r0, r1] = shiftPair(from.row, to.row, dRow, from.absRow, MAX_ROW, edge)
        const [c0, c1] = shiftPair(from.col, to.col, dCol, from.absCol, MAX_COL, edge)
        if (r0 === 'ref' || c0 === 'ref' || r1 === 'ref' || c1 === 'ref') { built++; return refErrorNode() }
        if (r0 === from.row && r1 === to.row && c0 === from.col && c1 === to.col) return node
        built++
        return { type: 'range', ref: { ...node.ref, from: { ...from, row: r0, col: c0 }, to: { ...to, row: r1, col: c1 } } }
      }
      case 'wholeCol': case 'wholeRow': {
        const { from, to } = node.ref
        const delta = node.type === 'wholeCol' ? dCol : dRow, max = node.type === 'wholeCol' ? MAX_COL : MAX_ROW
        const [a, b] = shiftPair(from.index, to.index, delta, from.absolute, max, edge)
        if (a === 'ref' || b === 'ref') { built++; return refErrorNode() }
        if (a === from.index && b === to.index) return node
        built++
        return { ...node, ref: { ...node.ref, from: { ...from, index: a }, to: { ...to, index: b } } } as AstNode
      }
      case 'unary': { const e = walk(node.expr); if (e === node.expr) return node; built++; return { ...node, expr: e } }
      case 'binary': {
        const l = walk(node.left), r = walk(node.right)
        if (l === node.left && r === node.right) return node
        built++
        return { ...node, left: l, right: r }
      }
      case 'call': {
        let changed = false
        const args = node.args.map(a => { const t = walk(a); if (t !== a) changed = true; return t })
        if (!changed) return node
        built++
        return { type: 'call', name: node.name, args }
      }
      default: return node
    }
  }
  try {
    const tree = walk(f.ast)
    return { kind: 'ready', value: hold.adopt(tree, Math.min(hold.bytes, built * CF_FORMULA_NODE_BYTES)) }
  } catch (error) {
    hold.release()
    if (error === EDGE_UNMEASURED) return { kind: 'unresolved', reason: 'cf-formula-edge-unmeasured' }
    throw error
  }
}
