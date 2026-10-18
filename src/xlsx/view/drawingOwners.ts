/** Drawing projections keep the existing reply allocation, not a second data
 * charge. Weak keys do not root a document: a root's self-edge and its aliases
 * are an ephemeron cycle which disappears when no consumer reaches it. */
import { DOCUMENT_DRAWING_NODE_LIMIT, DRAWING_GROUP_DEPTH } from '../../drawing/parts'
import type { SheetDrawings } from './contracts'
import type { MainMemoryLedger } from './mainLedger'

export const DRAWING_OWNER_LINK_BYTES = 64
export const DRAWING_OWNER_ROOT_BYTES = 128
/** One bounded Set (64 B per structural node) plus the two nested bounded DFS
 * stacks. Reserved before constructing the Set, retained through binding. */
export const DRAWING_OWNER_WORKSPACE_BYTES = 64 * DOCUMENT_DRAWING_NODE_LIMIT + 16 * 1024
const owners = new WeakMap<object, object>()
const metadata = new WeakMap<object, number>()
let serial = 0

interface ImageAlias { data?: Uint8Array; svg?: { bytes: Uint8Array } }
interface OwnerNode {
  source?: object; group?: object; image?: ImageAlias; children?: readonly OwnerNode[]
  content?: { kind: string; chart?: object; shapes?: readonly OwnerNode[] }
}
interface Plan { nodes: Set<OwnerNode>; drawingNodes: number; diagramShapes: number; links: number }

/** Only known structural projections: no traversal through chart points,
 * arbitrary model properties or binary bytes. Drawing and diagram nodes share
 * the parser's document node bound; chart models/images are counted as aliases,
 * never as extra drawing nodes. */
function planOwners(payload: SheetDrawings): Plan | undefined {
  const plan: Plan = { nodes: new Set(), drawingNodes: 0, diagramShapes: 0, links: 0 }
  const walk = (list: readonly OwnerNode[], depth: number, shapes: boolean): boolean => {
    if (depth > DRAWING_GROUP_DEPTH) return false
    for (let i = 0; i < list.length; i++) {
      const node = list[i]
      if (plan.nodes.has(node)) continue
      if (plan.nodes.size >= DOCUMENT_DRAWING_NODE_LIMIT) return false
      plan.nodes.add(node)
      if (shapes) plan.diagramShapes++; else plan.drawingNodes++
      if (node.children && !walk(node.children, depth + 1, shapes)) return false
      if (!shapes && node.content?.kind === 'diagram' && node.content.shapes && !walk(node.content.shapes, 0, true)) return false
    }
    return true
  }
  if (!walk(payload.drawings, 0, false)) return undefined
  visitAliases(payload, plan, () => { plan.links++ }, false)
  // Indexed images are not drawing nodes. Price their maximum six known alias
  // links each in O(1), before reading/binding any image element; the existing
  // byte budget bounds that later loop, without an invented image-count cap.
  plan.links += 6 * payload.images.length
  return plan
}

function visitAliases(payload: SheetDrawings, plan: Plan, visit: (alias: object) => void, indexed = true): void {
  const image = (asset: ImageAlias | undefined) => {
    if (!asset) return
    visit(asset)
    if (asset.data) { visit(asset.data); visit(asset.data.buffer) }
    if (asset.svg) { visit(asset.svg); visit(asset.svg.bytes); visit(asset.svg.bytes.buffer) }
  }
  visit(payload.drawings); visit(payload.images)
  for (const node of plan.nodes) {
    visit(node)
    if (node.source) visit(node.source)
    if (node.group) visit(node.group)
    if (node.children) visit(node.children)
    if (node.content) {
      visit(node.content)
      if (node.content.kind === 'xlsxChart' && node.content.chart) visit(node.content.chart)
      if (node.content.kind === 'diagram' && node.content.shapes) visit(node.content.shapes)
    }
    image(node.image)
  }
  if (indexed) for (let i = 0; i < payload.images.length; i++) image(payload.images[i])
}

/** The immutable source root, useful when making another shallow projection. */
export function drawingOwnerOf(value: object): object | undefined {
  const owner = owners.get(value)
  return owner && (owners.get(owner) ?? owner)
}
/** Explicit metadata component of a drawing reply's ledger charge. */
export function drawingOwnerMetadataBytes(value: object): number {
  const root = drawingOwnerOf(value)
  return root ? metadata.get(root) ?? 0 : 0
}

/** Copy callers reserve this weak-link metadata in SIZED_DRAWING_BYTES first.
 * Flatten repeated remaps so one live copy never retains a chain of old copies. */
export function retainDrawingCopyOwners(copies: readonly OwnerNode[], source: readonly OwnerNode[]): void {
  const root = drawingOwnerOf(source) ?? source
  if (!copies.length) return
  owners.set(copies, root)
  for (let i = 0; i < copies.length; i++) {
    const node = copies[i]
    if (node !== source[i]) {
      owners.set(node, copies)
      if (node.group && node.group !== source[i]?.group) owners.set(node.group, copies)
    }
  }
}

/** Admits only the new metadata. The reply's existing reachability charge
 * already prices the chart/image data and stays alive through the source root.
 * A failure constructs no ownership links and leaves that data charge intact. */
export function admitDrawingOwners(payload: SheetDrawings, ledger: MainMemoryLedger, budget: number): boolean {
  if (!payload.drawings.length && !payload.images.length) return true
  if (owners.has(payload.drawings)) return true
  if (ledger.bytes + DRAWING_OWNER_WORKSPACE_BYTES > budget) return false
  const workspace = `drawing-owner-work:${++serial}`, retained = `drawing-owner-links:${serial}`
  ledger.set('drawings', workspace, DRAWING_OWNER_WORKSPACE_BYTES)
  let plan: Plan | undefined
  try {
    plan = planOwners(payload)
    if (!plan) return false
    // Duplicate aliases deliberately over-count; no unreserved dedup Set.
    const bytes = DRAWING_OWNER_ROOT_BYTES + plan.links * DRAWING_OWNER_LINK_BYTES
    if (ledger.bytes + bytes > budget) return false
    ledger.set('drawings', retained, bytes)
    const root = payload.drawings
    ledger.adoptWhileReachable('drawings', retained, 'drawings', [root], bytes, () => {
      owners.set(root, root)
      metadata.set(root, bytes)
      visitAliases(payload, plan!, alias => { if (!owners.has(alias)) owners.set(alias, root) })
    })
    return true
  } finally {
    plan = undefined
    ledger.release('drawings', retained)
    ledger.release('drawings', workspace)
  }
}
