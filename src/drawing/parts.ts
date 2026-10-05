/** Package-scoped drawing traversal state. Cached source XML is theme-independent. */
import { attrs, getChildren, orderedChildren, type XmlNode } from '../core/xml'
import type { OfficePackage } from '../core/zip'

export const CONTENT_REFERENCE_DEPTH = 32
export const DRAWING_GROUP_DEPTH = 64
export const DOCUMENT_DRAWING_NODE_LIMIT = 10000
export interface ContentDiagnostic {
  kind: 'missing-part' | 'malformed-part' | 'external-reference' | 'content-cycle' | 'content-depth' | 'group-depth' | 'node-budget' | 'unsupported-content'
  message: string
  partPath: string
  feature?: string
  /** Relationship ID or resolved part which led to this branch. */
  identity?: string
  /** Source part whose relationship selected this target, when partPath names the target. */
  ownerPartPath?: string
  /** Relationship on the original drawing owner that selected this content chain. */
  sourceReferenceId?: string
  reason?: string
  limit?: number
}
export interface DrawingPartContext {
  nodes: number
  diagnostics: ContentDiagnostic[]
  coverage: { loaded: number; repeated: number; fallbacks: number }
  relationships: Map<string, Promise<Map<string, { path?: string; external: boolean; type?: string }>>>
  relationshipNodes: Map<string, Promise<XmlNode[]>>
  placements: Set<string>
}
const contexts = new WeakMap<OfficePackage, DrawingPartContext>()
export function drawingPartContext(pkg: OfficePackage): DrawingPartContext {
  let context = contexts.get(pkg)
  if (!context) {
    context = { nodes: 0, diagnostics: [], coverage: { loaded: 0, repeated: 0, fallbacks: 0 }, relationships: new Map(), relationshipNodes: new Map(), placements: new Set() }
    contexts.set(pkg, context)
  }
  return context
}
export function contentDiagnostic(context: DrawingPartContext, kind: ContentDiagnostic['kind'], partPath: string, feature?: string, details: Pick<ContentDiagnostic, 'identity' | 'reason' | 'limit' | 'ownerPartPath' | 'sourceReferenceId'> = {}): void {
  context.diagnostics.push({ kind, partPath, feature, ...details, message: `Drawing content ${kind}${feature ? `: ${feature}` : ''}` })
}
export function reserveDrawingNode(context: DrawingPartContext, partPath: string): boolean {
  if (context.nodes >= DOCUMENT_DRAWING_NODE_LIMIT) {
    if (!context.diagnostics.some(issue => issue.kind === 'node-budget' && issue.reason === 'document-budget')) contentDiagnostic(context, 'node-budget', partPath, undefined, { reason: 'document-budget', limit: DOCUMENT_DRAWING_NODE_LIMIT })
    return false
  }
  context.nodes++
  return true
}
export function resolvePartTarget(owner: string, target: string): string {
  const bits = target.startsWith('/') ? [] : owner.split('/').slice(0, -1)
  for (const bit of target.split('/')) {
    if (bit === '..') bits.pop()
    else if (bit && bit !== '.') bits.push(bit)
  }
  return bits.join('/')
}
/** Optional owner relationships are read once and fail as a diagnosed branch. */
export function partRelationshipNodes(pkg: OfficePackage, owner: string): Promise<XmlNode[]> {
  const context = drawingPartContext(pkg)
  let pending = context.relationshipNodes.get(owner)
  if (!pending) {
    pending = (async () => {
      const slash = owner.lastIndexOf('/')
      const path = `${slash < 0 ? '' : owner.slice(0, slash) + '/'}_rels/${owner.slice(slash + 1)}.rels`
      try { return getChildren(await pkg.xmlOrdered(path), 'Relationship') }
      catch { contentDiagnostic(context, 'malformed-part', path, undefined, { identity: owner, reason: 'invalid-relationship-xml' }); return [] }
    })()
    context.relationshipNodes.set(owner, pending)
  }
  return pending
}
export function partRelationships(pkg: OfficePackage, owner: string): Promise<Map<string, { path?: string; external: boolean; type?: string }>> {
  const context = drawingPartContext(pkg)
  let pending = context.relationships.get(owner)
  if (!pending) {
    pending = (async () => {
      const out = new Map<string, { path?: string; external: boolean; type?: string }>()
      for (const node of await partRelationshipNodes(pkg, owner)) {
        const a = attrs(node), external = a.TargetMode === 'External'
        if (a.Id && a.Target) out.set(a.Id, { path: external ? undefined : resolvePartTarget(owner, a.Target), external, type: a.Type })
      }
      return out
    })()
    context.relationships.set(owner, pending)
  }
  return pending
}
export async function referencedPart(pkg: OfficePackage, owner: string, id: string): Promise<string | undefined> {
  const rel = (await partRelationships(pkg, owner)).get(id), context = drawingPartContext(pkg)
  if (rel?.external) { contentDiagnostic(context, 'external-reference', owner, id, { identity: id, reason: 'external-relationship' }); return undefined }
  if (!rel?.path || !pkg.has(rel.path)) { contentDiagnostic(context, 'missing-part', owner, id, { identity: id, reason: rel?.path ? 'part-not-found' : 'relationship-not-found' }); return undefined }
  return rel.path
}
/** Iterative DFS retains source order and avoids recursive XML-stack exhaustion. */
export function contentDescendants(node: XmlNode | undefined, tag: string): XmlNode[] {
  const out: XmlNode[] = [], pending = [...orderedChildren(node)].reverse()
  while (pending.length) {
    const [name, child] = pending.pop()!
    if (name === tag) out.push(child)
    const children = orderedChildren(child)
    for (let i = children.length - 1; i >= 0; i--) if (children[i][0] !== '#text') pending.push(children[i])
  }
  return out
}
