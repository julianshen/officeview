/**
 * Master and layout artwork (slideshow spec 3 and 11 D2): the non-placeholder shapes
 * of a slide's layout and master (logos, bands, decorations), honouring showMasterSp.
 * Lazy: loaded only through the appearance chunk.
 */
import type { OfficePackage } from '../../core/zip'
import { attrs, getChildren, orderedChildren, type XmlNode } from '../../core/xml'
import { partRelationshipNodes, resolvePartTarget } from '../../drawing/parts'
import { emuToPx } from '../../core/geometry'
import { resolveGeometry } from '../../drawing/geometry'
import type { ThemeContext } from '../../drawing/style'
import { parseArtworkTree } from '../parse'
import type { PptxDiagnostic, PptxShape } from '../types'

export interface MasterArtwork { master: PptxShape[]; layout: PptxShape[]; diagnostics: PptxDiagnostic[] }

const MAX_DEPTH = 128
const NONVISUAL = ['nvSpPr', 'nvCxnSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvGrpSpPr', 'nvContentPartPr']

/** The first local relationship of a kind; external or absent parts have no artwork. */
async function related(pkg: OfficePackage, partPath: string, kind: 'slideLayout' | 'slideMaster'): Promise<string | undefined> {
  for (const rel of await partRelationshipNodes(pkg, partPath)) {
    const a = attrs(rel)
    if (a.TargetMode !== 'External' && a.Target && a.Type?.endsWith(`/${kind}`)) return resolvePartTarget(partPath, a.Target)
  }
  return undefined
}

/** showMasterSp is an xsd:boolean defaulting to true. */
const hidesMaster = (root: XmlNode | undefined): boolean => ['0', 'false'].includes(attrs(root).showMasterSp ?? '')

/** `cNvPr@id` of every drawing element whose non-visual properties carry `p:nvPr/p:ph`, at any depth (groups, both AlternateContent branches). */
function placeholderIds(root: XmlNode | undefined): Set<string> {
  const ids = new Set<string>()
  const pending: Array<[XmlNode, number]> = root ? [[root, 0]] : []
  while (pending.length) {
    const [node, depth] = pending.pop()!
    if (depth >= MAX_DEPTH) continue
    for (const [name, child] of orderedChildren(node)) {
      if (name === '#text') continue
      if (NONVISUAL.includes(name)) {
        const id = attrs(getChildren(child, 'cNvPr')[0]).id
        if (id !== undefined && getChildren(getChildren(child, 'nvPr')[0], 'ph').length) ids.add(id)
        continue
      }
      pending.push([child, depth + 1])
    }
  }
  return ids
}

/** True when a drawing element's non-visual properties carry `p:nvPr/p:ph`. */
const isPlaceholder = (node: XmlNode): boolean => NONVISUAL.some(name => getChildren(getChildren(getChildren(node, name)[0], 'nvPr')[0], 'ph').length > 0)

/**
 * The image rIds (`r:embed`, `r:link`) the part's non-placeholder elements reference, at any depth and in every
 * AlternateContent branch, so placeholder-only and unreferenced image relationships are never read.
 */
function usedImageRids(root: XmlNode | undefined): Set<string> {
  const rids = new Set<string>()
  const pending: Array<[XmlNode, number]> = root ? [[root, 0]] : []
  while (pending.length) {
    const [node, depth] = pending.pop()!
    if (depth >= MAX_DEPTH) continue
    for (const [name, child] of orderedChildren(node)) {
      if (name === '#text' || isPlaceholder(child)) continue
      const a = attrs(child)
      if (a.embed) rids.add(a.embed)
      if (a.link) rids.add(a.link)
      pending.push([child, depth + 1])
    }
  }
  return rids
}

/** `cNvPr@id` of every `p:sp` with `useBgFill` (an xsd:boolean) in a part's shape tree, at any depth. */
function useBgFillIds(root: XmlNode | undefined): Set<string> {
  const ids = new Set<string>()
  const pending: Array<[XmlNode, number]> = root ? [[root, 0]] : []
  while (pending.length) {
    const [node, depth] = pending.pop()!
    if (depth >= MAX_DEPTH) continue
    for (const [name, child] of orderedChildren(node)) {
      if (name === '#text') continue
      if (name === 'sp' && ['1', 'true'].includes(attrs(child).useBgFill ?? '')) {
        const id = attrs(getChildren(getChildren(child, 'nvSpPr')[0], 'cNvPr')[0]).id
        if (id !== undefined) ids.add(id)
      }
      pending.push([child, depth + 1])
    }
  }
  return ids
}

/**
 * Flag a part's `useBgFill` shapes so the painter shows the slide background through them (in slide
 * coordinates, whatever the background: solid, gradient, pattern or picture). The flag is a property of the
 * part, so cached artwork shared between slides stays slide-independent; each slide supplies its own background.
 */
export function flagUseBgFill(shapes: PptxShape[], spTree: XmlNode | undefined, partPath: string): void {
  const ids = useBgFillIds(spTree)
  if (!ids.size) return
  const walk = (list: PptxShape[]) => {
    for (const shape of list) {
      if (shape.source?.partPath === partPath && shape.source.element === 'sp' && shape.source.id !== undefined && ids.has(shape.source.id)) shape.useBgFill = true
      if (shape.children) walk(shape.children)
    }
  }
  walk(shapes)
}

function withoutPlaceholders(shapes: PptxShape[], ids: Set<string>): PptxShape[] {
  const kept: PptxShape[] = []
  for (const shape of shapes) {
    if (shape.placeholder || (shape.source?.id !== undefined && ids.has(shape.source.id))) continue
    if (shape.children) shape.children = withoutPlaceholders(shape.children, ids)
    kept.push(shape)
  }
  return kept
}

/** One part's artwork: parsed whole (reading only the images kept shapes use), then placeholders and their diagnostics are filtered out. */
async function partArtwork(pkg: OfficePackage, partPath: string, theme: ThemeContext, diagnostics: PptxDiagnostic[]): Promise<PptxShape[]> {
  const issues: PptxDiagnostic[] = []
  const spTree = getChildren(getChildren(await pkg.xmlOrdered(partPath), 'cSld')[0], 'spTree')[0]
  const parsed = await parseArtworkTree(pkg, partPath, theme, issues, usedImageRids(spTree))
  const ids = placeholderIds(spTree)
  const shapes = withoutPlaceholders(parsed, ids)
  flagUseBgFill(shapes, spTree, partPath)
  diagnostics.push(...issues.filter(issue => !(issue.source?.partPath === partPath && issue.source.id !== undefined && ids.has(issue.source.id))))
  // As the slide loop does: geometry and style issues of each kept shape.
  const walk = (list: PptxShape[]) => {
    for (const shape of list) {
      const geometry = shape.drawingGeometry ? resolveGeometry(shape.drawingGeometry, emuToPx(shape.widthEmu), emuToPx(shape.heightEmu)).issues : []
      shape.diagnostics = [...(shape.diagnostics ?? []), ...[...geometry, ...(shape.drawingStyle?.issues ?? [])].map(issue => ({ ...issue, source: shape.source }))]
      diagnostics.push(...shape.diagnostics)
      if (shape.children) walk(shape.children)
    }
  }
  walk(shapes)
  return shapes
}

/** One enrichment run's parsed parts, keyed by part and effective colour map, so a master shared by many slides
 * is parsed (and charged to the run's node budget) once. Slides sharing an entry share its shape objects. */
export type ArtworkCache = Map<string, Promise<{ shapes: PptxShape[]; diagnostics: PptxDiagnostic[] }>>

async function cachedArtwork(pkg: OfficePackage, partPath: string, theme: ThemeContext, diagnostics: PptxDiagnostic[], cache?: ArtworkCache): Promise<PptxShape[]> {
  if (!cache) return partArtwork(pkg, partPath, theme, diagnostics)
  const key = `${partPath}\u0000${JSON.stringify(theme.colorMap)}`
  let entry = cache.get(key)
  if (!entry) {
    const own: PptxDiagnostic[] = []
    entry = partArtwork(pkg, partPath, theme, own).then(shapes => ({ shapes, diagnostics: own }))
    cache.set(key, entry)
  }
  const { shapes, diagnostics: own } = await entry
  diagnostics.push(...own)
  return shapes
}

/**
 * The slide's master and layout non-placeholder shapes, in paint order (master below layout).
 * Slide `showMasterSp="0"` drops both; layout `showMasterSp="0"` drops the master's only.
 * Shapes keep their layout/master `source.partPath`, so their ids never collide with slide ids.
 */
export async function masterArtwork(pkg: OfficePackage, slidePath: string, theme: ThemeContext, cache?: ArtworkCache): Promise<MasterArtwork> {
  const diagnostics: PptxDiagnostic[] = []
  const art: MasterArtwork = { master: [], layout: [], diagnostics }
  if (hidesMaster(await pkg.xmlOrdered(slidePath))) return art
  const layoutPath = await related(pkg, slidePath, 'slideLayout')
  // A slide without a layout may still link its master directly.
  const masterPath = await related(pkg, layoutPath ?? slidePath, 'slideMaster')
  if (masterPath && !(layoutPath && hidesMaster(await pkg.xmlOrdered(layoutPath)))) art.master = await cachedArtwork(pkg, masterPath, theme, diagnostics, cache)
  if (layoutPath) art.layout = await cachedArtwork(pkg, layoutPath, theme, diagnostics, cache)
  return art
}
