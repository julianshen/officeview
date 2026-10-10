/** Serializable facts about source drawing placements and selected representations. */
export interface DrawingCoverageEntry {
  partPath: string
  treePath: string
  element: string
  id?: string
  name?: string
  /** Relationship identifier when a placement depends on a separate part. */
  referenceId?: string
  feature: string
  status: 'native' | 'fallback' | 'unsupported' | 'malformed'
  /** Actual payload consumed by this static renderer, independent of MC branch. */
  selectedRepresentation: 'native-shape' | 'picture' | 'raster-fallback' | 'cached-diagram' | 'column-chart' | 'inkml' | 'textbox' | 'table' | 'group' | 'text-only' | 'blank' | 'none' | 'unverified'
  representation: 'native' | 'choice' | 'fallback'
  reason?: string
  /** Why the retained paint assessment is inconclusive; source failure remains in reason. */
  assessmentReason?: 'winding-analysis-limit' | 'curve-paint-inconclusive' | 'dash-paint-inconclusive'
  /** Selected SVG/raster representation and phase for a picture placement. */
  imageSelection?: ImageSelection
  /** Maximum distinct straight edges analyzed for completed nonzero winding. */
  analysisLimit?: number
  /** Original placements/containers, selected nested nodes, and failures without a placement. */
  scope: 'original' | 'descendant' | 'diagnostic'
  /** Zero-based slide/sheet index. Omitted for Word: flow placement has no fixed page. */
  unit?: number
  limit?: number
  requestedExtent?: { width: number; height: number }
  retainedExtent?: { width: number; height: number }
}

export interface DrawingCoverageDocument { drawingCoverage?: DrawingCoverageEntry[] }
/**
 * Map prepared content to its coverage representation. Text-only diagrams
 * (cacheless SmartArt synthesized from dataModel text) reuse the existing
 * text-only status instead of claiming a cached diagram rendering.
 */
export function contentRepresentation<Paragraph, Text>(content: DrawingContent<Paragraph, Text>): DrawingCoverageEntry['selectedRepresentation'] {
  if (content.kind === 'diagram' && content.textOnly) return 'text-only'
  const kind = content.kind
  return kind === 'diagram' ? 'cached-diagram' : kind === 'chart' ? 'column-chart' : kind === 'ink' ? 'inkml' : kind === 'textbox' ? 'textbox' : 'none'
}

import { attrs, getChildren, namespaceUri, orderedChildren, type XmlNode } from '../core/xml'
import type { ImageSelection } from '../core/svg'
import type { PptxShape, PptxSlide } from '../pptx/types'
import type { XlsxDrawing } from '../xlsx/types'
import { resolveGeometry, type ResolvedPath } from './geometry'
import { emuToPx } from '../core/geometry'
import type { DrawingFill, DrawingLine } from './style'
import type { ContentDiagnostic } from './parts'
import type { DrawingContent } from './content'

/** Match a diagnosed relationship to its selected source interpretation. */
export function coverageIssueMatchesEntry(issue: ContentDiagnostic, entry: DrawingCoverageEntry): boolean {
  const sourceReferenceId = issue.sourceReferenceId ?? issue.identity
  if (!sourceReferenceId || !entry.referenceId || (issue.ownerPartPath ?? issue.partPath) !== entry.partPath || sourceReferenceId !== entry.referenceId) return false
  if ((issue.sourceReferenceId && issue.feature !== 'image' || ['chart', 'diagram', 'contentPart'].includes(issue.feature ?? '')) &&
    (entry.element === 'pic' || entry.feature === 'picture' || entry.selectedRepresentation === 'picture' || entry.selectedRepresentation === 'raster-fallback')) return false
  // A target may be a valid image while parsing the same bytes as chart XML
  // fails. The XML failure belongs to the chart/content attempt only.
  if (issue.kind === 'malformed-part' && issue.reason === 'invalid-content-xml')
    return entry.element !== 'pic' && entry.feature !== 'picture' && entry.selectedRepresentation !== 'picture' && entry.selectedRepresentation !== 'raster-fallback'
  return true
}

const knownNamespaceUris: Record<string, string> = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  dsp: 'http://schemas.microsoft.com/office/drawing/2008/diagram',
  ink: 'http://www.w3.org/2003/InkML',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
}

/** MC Requires identifies namespace URIs, not the lexical prefix spelling. */
export function supportedChoiceRequirements(choice: XmlNode, supported: readonly string[]): { supported: boolean; unknown?: string } {
  const allowedUris = new Set(supported.map(prefix => knownNamespaceUris[prefix]).filter(Boolean))
  for (const prefix of (attrs(choice).Requires ?? '').trim().split(/\s+/).filter(Boolean)) {
    if (!allowedUris.has(namespaceUri(choice, prefix) ?? '')) return { supported: false, unknown: prefix }
  }
  return { supported: true }
}

const drawingNames = new Set(['sp', 'cxnSp', 'pic', 'graphicFrame', 'grpSp', 'contentPart', 'AlternateContent'])
const identity = (node: XmlNode | undefined) => {
  const nv = ['nvSpPr', 'nvCxnSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvGrpSpPr'].map(name => getChildren(node, name)[0]).find(Boolean)
  const a = attrs(getChildren(nv, 'cNvPr')[0])
  return { id: a.id, name: a.name }
}
function firstIdentity(node: XmlNode | undefined, depth = 0): { id?: string; name?: string } {
  if (!node || depth > 32) return {}
  const own = identity(node)
  if (own.id || own.name) return own
  for (const [, child] of orderedChildren(node)) {
    const nested = firstIdentity(child, depth + 1)
    if (nested.id || nested.name) return nested
  }
  return {}
}
function sourceFeature(node: XmlNode | undefined, name: string, depth = 0): string {
  if (depth > 32) return name
  if (name === 'pic') return 'picture'
  if (name === 'sp' || name === 'cxnSp') return 'shape'
  if (name === 'grpSp') return 'group'
  const graphicData = getChildren(getChildren(node, 'graphic')[0], 'graphicData')[0] ?? getChildren(node, 'graphicData')[0]
  const uri = attrs(graphicData).uri?.toLowerCase() ?? ''
  if (uri.includes('chartex')) return 'ChartEx'
  if (uri.includes('model3d')) return 'model3D'
  if (getChildren(graphicData, 'chart').length) return 'chart'
  if (getChildren(graphicData, 'relIds').length) return 'diagram'
  if (getChildren(graphicData, 'contentPart').length || getChildren(node, 'contentPart').length) return 'ink'
  if (graphicData) return 'graphicData'
  for (const [childName, child] of orderedChildren(node)) {
    if (childName === 'Fallback') continue
    if (childName === 'oMath' || childName === 'oMathPara' || childName === 'm') return 'OMML'
    if (childName.toLowerCase() === 'model3d') return 'model3D'
    if (childName === 'graphicData' && (attrs(child).uri ?? '').toLowerCase().includes('chartex')) return 'ChartEx'
    const found = sourceFeature(child, childName, depth + 1)
    if (['OMML', 'model3D', 'ChartEx'].includes(found)) return found
  }
  return name === 'AlternateContent' ? 'AlternateContent' : name
}

/** A round cap paints a point path; butt and square caps need a real segment. */
function strokePathRetainsPaint(path: ResolvedPath, cap: CanvasLineCap): boolean {
  let current: [number, number] | undefined, start: [number, number] | undefined
  let stroked = false
  const differs = (point: [number, number]): boolean => current !== undefined && (current[0] !== point[0] || current[1] !== point[1])
  for (const command of path.commands) {
    switch (command[0]) {
      case 'moveTo':
        current = start = [command[1], command[2]]
        break
      case 'lnTo': {
        const end: [number, number] = [command[1], command[2]]
        if (differs(end)) return true
        current = end; stroked = true
        break
      }
      case 'quadBezTo': {
        const control: [number, number] = [command[1], command[2]], end: [number, number] = [command[3], command[4]]
        if (differs(control) || differs(end)) return true
        current = end; stroked = true
        break
      }
      case 'cubicBezTo': {
        const first: [number, number] = [command[1], command[2]], second: [number, number] = [command[3], command[4]], end: [number, number] = [command[5], command[6]]
        if (differs(first) || differs(second) || differs(end)) return true
        current = end; stroked = true
        break
      }
      case 'arcTo': {
        const end: [number, number] = [command[1] + command[3] * Math.cos(command[6]), command[2] + command[4] * Math.sin(command[6])]
        if (command[3] !== 0 && command[4] !== 0) {
          const arcStart: [number, number] = [command[1] + command[3] * Math.cos(command[5]), command[2] + command[4] * Math.sin(command[5])]
          if (differs(arcStart) || command[5] !== command[6]) return true
        } else if (differs(end)) return true
        current = end; stroked = true
        break
      }
      case 'close':
        if (start && differs(start)) return true
        current = start; stroked = true
        break
    }
  }
  return cap === 'round' && stroked
}

type PaintAssessment = { kind: 'paint' | 'none' | 'unverified'; reason?: DrawingCoverageEntry['assessmentReason']; limit?: number }
const paint: PaintAssessment = { kind: 'paint' }, noPaint: PaintAssessment = { kind: 'none' }

/** Remove only curves that cannot change the painter's path. Other curves stay inconclusive. */
function straightCommands(path: ResolvedPath): ResolvedPath['commands'] | undefined {
  const commands: ResolvedPath['commands'] = []
  let current: [number, number] | undefined, start: [number, number] | undefined
  const same = (x: number, y: number): boolean => current?.[0] === x && current[1] === y
  for (const command of path.commands) {
    switch (command[0]) {
      case 'moveTo': current = start = [command[1], command[2]]; commands.push(command); break
      case 'lnTo': current = [command[1], command[2]]; commands.push(command); break
      case 'close': current = start; commands.push(command); break
      case 'quadBezTo':
        if (!same(command[1], command[2]) || !same(command[3], command[4])) return undefined
        break
      case 'cubicBezTo':
        if (!same(command[1], command[2]) || !same(command[3], command[4]) || !same(command[5], command[6])) return undefined
        break
      case 'arcTo': {
        const end: [number, number] = [command[1] + command[3] * Math.cos(command[6]), command[2] + command[4] * Math.sin(command[6])]
        if (command[3] !== 0 && command[4] !== 0 && command[5] !== command[6]) return undefined
        if (!same(...end)) {
          if (command[3] !== 0 && command[4] !== 0) return undefined
          commands.push(['lnTo', ...end])
        }
        current = end
        break
      }
    }
  }
  return commands
}

/** A completed straight path fills only where its nonzero winding has area.
 *  Exact inverse segments cancel in a linear prepass, even for long contours.
 *  At most 64 distinct remaining edges enter the crossing search (under 2,016
 *  pairs); more complex residuals have an explicit inconclusive assessment.
 */
function fillPathAssessment(path: ResolvedPath): PaintAssessment {
  const maxWindingEdges = 64
  const commands = straightCommands(path)
  if (!commands) return { kind: 'unverified', reason: 'curve-paint-inconclusive' }
  type Point = [number, number]
  type Edge = [Point, Point]
  const edges: Edge[] = []
  let current: Point | undefined, start: Point | undefined
  const addEdge = (end: Point): void => {
    if (current && (current[0] !== end[0] || current[1] !== end[1])) edges.push([current, end])
    current = end
  }
  for (const command of commands) {
    if (command[0] === 'moveTo') {
      if (start) addEdge(start) // Canvas implicitly closes a filled subpath.
      current = start = [command[1], command[2]]
    } else if (command[0] === 'lnTo') addEdge([command[1], command[2]])
    else if (command[0] === 'close' && start) addEdge(start)
  }
  if (start) addEdge(start)
  const key = (a: Point, b: Point): string => `${a[0]},${a[1]}|${b[0]},${b[1]}`
  const uncancelled = new Map<string, { edge: Edge; count: number }>()
  for (const edge of edges) {
    const [a, b] = edge, reverseKey = key(b, a), opposite = uncancelled.get(reverseKey)
    if (opposite) {
      if (--opposite.count === 0) uncancelled.delete(reverseKey)
    } else {
      const forwardKey = key(a, b), currentCount = uncancelled.get(forwardKey)
      if (currentCount) currentCount.count++
      else uncancelled.set(forwardKey, { edge, count: 1 })
    }
  }
  const remaining = [...uncancelled.values()]
  if (remaining.length > maxWindingEdges) return { kind: 'unverified', reason: 'winding-analysis-limit', limit: maxWindingEdges }
  if (remaining.length < 3) return noPaint
  const cross = (ax: number, ay: number, bx: number, by: number): number => ax * by - ay * bx
  const levels = remaining.flatMap(({ edge: [a, b] }) => [a[1], b[1]])
  for (let i = 0; i < remaining.length; i++) for (let j = i + 1; j < remaining.length; j++) {
    const [[ax, ay], [bx, by]] = remaining[i].edge, [[cx, cy], [dx, dy]] = remaining[j].edge
    const rx = bx - ax, ry = by - ay, sx = dx - cx, sy = dy - cy
    const denominator = cross(rx, ry, sx, sy)
    if (denominator === 0) continue
    const t = cross(cx - ax, cy - ay, sx, sy) / denominator
    const u = cross(cx - ax, cy - ay, rx, ry) / denominator
    if (t > 0 && t < 1 && u > 0 && u < 1) levels.push(ay + t * ry)
  }
  levels.sort((a, b) => a - b)
  for (let i = 1; i < levels.length; i++) {
    if (levels[i] - levels[i - 1] <= 1e-9) continue
    const y = (levels[i] + levels[i - 1]) / 2
    const events: Array<{ x: number; delta: number }> = []
    for (const { edge: [[ax, ay], [bx, by]], count } of remaining) {
      if (y <= Math.min(ay, by) || y >= Math.max(ay, by)) continue
      events.push({ x: ax + (y - ay) * (bx - ax) / (by - ay), delta: (by > ay ? 1 : -1) * count })
    }
    events.sort((a, b) => a.x - b.x)
    let winding = 0
    for (let j = 0; j + 1 < events.length; j++) {
      winding += events[j].delta
      if (winding !== 0 && events[j + 1].x - events[j].x > 1e-9) return paint
    }
  }
  return noPaint
}

/** A delayed dash paints only when a known first segment reaches its first ink slot. */
function strokeDashAssessment(path: ResolvedPath, line: DrawingLine): PaintAssessment {
  if (!Array.isArray(line.dash) || line.dash.length === 0) return paint
  const dash = line.dash
  // setupLine treats invalid/all-zero patterns as solid; odd patterns repeat.
  const width = line.width ?? 1
  if (dash.some(n => !Number.isFinite(n) || n < 0 || !Number.isFinite(n * width)) || !dash.some(n => n > 0)) return paint
  const effective = dash.length % 2 ? [...dash, ...dash] : dash
  const firstInk = effective.findIndex((length, index) => index % 2 === 0 && length > 0)
  if (firstInk < 0) return (line.cap ?? 'butt') === 'butt' ? noPaint : paint
  // Square and round caps visibly paint the zero-length starting dash.
  if ((line.cap ?? 'butt') !== 'butt' || firstInk === 0) return paint
  const gap = effective.slice(0, firstInk).reduce((sum, length) => sum + length * width, 0)
  if (gap === 0) return paint
  // A first straight segment longer than the gap proves that the following
  // positive slot is reached. Curves and shorter paths stay inconclusive.
  let current: [number, number] | undefined
  for (const command of path.commands) {
    if (command[0] === 'moveTo') {
      if (current) break
      current = [command[1], command[2]]
    } else if (command[0] === 'lnTo' && current) {
      const length = Math.hypot(command[1] - current[0], command[2] - current[1])
      if (length > gap) return paint
      if (length > 0) break
      current = [command[1], command[2]]
    } else if (command[0] !== 'close') break
  }
  return { kind: 'unverified', reason: 'dash-paint-inconclusive' }
}

/** Inventory the final selected PPTX scene against its original placement tree. */
export function pptxCoverage(slide: PptxSlide, spTree: XmlNode | undefined, partPath: string, selectedBranch?: (alternate: XmlNode) => XmlNode | undefined): DrawingCoverageEntry[] {
  const entries: DrawingCoverageEntry[] = []
  const visibleFill = (fill: DrawingFill | undefined, width: number, height: number): boolean => {
    if (!fill || fill.kind === 'none') return false
    if (fill.kind === 'solid') return [fill.color.r, fill.color.g, fill.color.b, fill.color.a].every(Number.isFinite) && fill.color.a > 0
    if (fill.gradient === 'circle' && (width <= 0 || height <= 0)) return false
    const dx = Math.cos(fill.angle) * (fill.scaled ? height : 1)
    const dy = Math.sin(fill.angle) * (fill.scaled ? width : 1)
    return Number.isFinite(fill.angle) && (fill.gradient === 'circle' || Number.isFinite(Math.hypot(dx, dy)) && Math.hypot(dx, dy) > 0) && fill.stops.length > 0 &&
      fill.stops.every(stop => Number.isFinite(stop.position) && [stop.color.r, stop.color.g, stop.color.b, stop.color.a].every(Number.isFinite)) &&
      fill.stops.some(stop => stop.color.a > 0)
  }
  const retainedGeometry = (shape: PptxShape | undefined): PaintAssessment => {
    if (!shape || (shape.source?.element !== 'sp' && shape.source?.element !== 'cxnSp') ||
      shape.transformValid === false || ![shape.widthEmu, shape.heightEmu].every(n => Number.isFinite(n) && n >= 0) ||
      shape.diagnostics?.some(issue => ['unknown-preset', 'invalid-guide', 'invalid-path', 'invalid-text-rect', 'invalid-extents', 'invalid-transform'].includes(issue.kind))) return noPaint
    if (shape.drawingGeometry) {
      const width = emuToPx(shape.widthEmu), height = emuToPx(shape.heightEmu)
      const geometry = resolveGeometry(shape.drawingGeometry, width, height)
      const style = shape.drawingStyle
      if (geometry.issues.length) return noPaint
      let unknown: PaintAssessment | undefined
      for (const path of geometry.paths) {
        if (path.fill !== 'none' && width > 0 && height > 0 && visibleFill(style?.fill, width, height)) {
          const result = fillPathAssessment(path)
          if (result.kind === 'paint') return paint
          if (result.kind === 'unverified') unknown ??= result
        }
        if (path.stroke && style?.line && Number.isFinite(style.line.width ?? 1) && (style.line.width ?? 1) > 0 &&
          visibleFill(style.line.fill, width, height) && strokePathRetainsPaint(path, style.line.cap ?? 'butt')) {
          if ((style.line.headEnd?.type && style.line.headEnd.type !== 'none' || style.line.tailEnd?.type && style.line.tailEnd.type !== 'none') &&
            strokePathRetainsPaint(path, 'butt')) return paint
          const dash = strokeDashAssessment(path, style.line)
          if (dash.kind === 'paint') return paint
          if (dash.kind === 'unverified') unknown ??= dash
        }
      }
      return unknown ?? noPaint
    }
    // Hand-built legacy models use the same simple fill and line fields as scene-paint.
    return shape.fill && shape.widthEmu > 0 && shape.heightEmu > 0 || shape.line && (shape.line.widthEmu ?? 9525) > 0 ? paint : noPaint
  }
  const retainedGeometryAt = (entry: DrawingCoverageEntry): boolean => allShapes.some(shape =>
    (shape.source?.treePath === entry.treePath || entry.element === 'AlternateContent' && shape.source?.treePath.startsWith(`${entry.treePath}/`)) &&
    retainedGeometry(shape).kind !== 'none')
  const unselectedReferences = new Set<string>(), selectedReferences = new Set<string>()
  const referenceIds = (node: XmlNode, sink: Set<string>, depth = 0, rootName?: string): void => {
    if (depth >= 128) return
    if (rootName && ['blip', 'chart', 'relIds', 'contentPart', 'imagedata'].includes(rootName)) {
      const a = attrs(node)
      for (const key of ['id', 'embed', 'dm', 'lo', 'qs', 'cs']) if (a[key]) sink.add(a[key])
    }
    for (const [name, child] of orderedChildren(node)) {
      if (['blip', 'chart', 'relIds', 'contentPart', 'imagedata'].includes(name)) {
        const a = attrs(child)
        for (const key of ['id', 'embed', 'dm', 'lo', 'qs', 'cs']) if (a[key]) sink.add(a[key])
      }
      if (name !== '#text') referenceIds(child, sink, depth + 1)
    }
  }
  const scanSelection = (node: XmlNode | undefined, depth = 0): void => {
    if (!node || depth >= 128) return
    for (const [name, child] of orderedChildren(node)) {
      if (name === 'AlternateContent') {
        const chosen = selectedBranch?.(child)
        if (chosen) for (const branch of [...getChildren(child, 'Choice'), ...getChildren(child, 'Fallback')]) {
          if (branch !== chosen) referenceIds(branch, unselectedReferences)
        }
        scanSelection(chosen, depth + 1)
      } else if (name !== '#text') {
        if (['blip', 'chart', 'relIds', 'contentPart'].includes(name)) {
          const a = attrs(child)
          for (const key of ['id', 'embed', 'dm', 'lo', 'qs', 'cs']) if (a[key]) selectedReferences.add(a[key])
        }
        scanSelection(child, depth + 1)
      }
    }
  }
  scanSelection(spTree)
  const allShapes: PptxShape[] = []
  const visit = (shapes: PptxShape[]) => { for (const shape of shapes) { allShapes.push(shape); visit(shape.children ?? []) } }
  visit(slide.shapes)
  const add = (node: XmlNode, name: string, path: string, scope: DrawingCoverageEntry['scope'], inheritedRepresentation?: 'choice' | 'fallback') => {
    const selected = allShapes.filter(shape => shape.source?.treePath === path || (name === 'AlternateContent' && shape.source?.treePath.startsWith(`${path}/`)))
    const shape = selected[0]
    const source = shape?.source
    const chosen = name === 'AlternateContent' ? selectedBranch?.(node) : undefined
    const attempted = name === 'AlternateContent' && !chosen
      ? getChildren(node, 'Choice').find(choice => supportedChoiceRequirements(choice, ['a', 'p', 'c', 'dgm', 'dsp', 'ink']).supported) ?? getChildren(node, 'Choice')[0]
      : undefined
    const own = name === 'AlternateContent' ? firstIdentity(chosen ?? attempted) : identity(node)
    let nested: { branch: XmlNode; representation: 'choice' | 'fallback' } | undefined
    if (name !== 'AlternateContent') {
      const scan = (parent: XmlNode | undefined, depth = 0): void => {
        if (!parent || depth >= 128) return
        for (const [childName, child] of orderedChildren(parent)) {
          if (childName === 'AlternateContent') {
            const branch = selectedBranch?.(child)
            if (branch) { nested = { branch, representation: getChildren(child, 'Fallback').includes(branch) ? 'fallback' : 'choice' }; scan(branch, depth + 1) }
          } else if (childName !== '#text') scan(child, depth + 1)
        }
      }
      scan(node)
    }
    const originalFeature = sourceFeature(node, name)
    const selectedPayload = (chosen ?? attempted) && orderedChildren(chosen ?? attempted).find(([childName]) => drawingNames.has(childName))
    const selectedFeature = selectedPayload ? sourceFeature(selectedPayload[1], selectedPayload[0]) : originalFeature
    let feature = source?.representation === 'fallback'
      ? (['ChartEx', 'OMML', 'model3D'].includes(originalFeature) ? originalFeature : source.feature ?? originalFeature)
      : shape ? shape.content?.kind ?? sourceFeature(undefined, source?.element ?? name) : selectedFeature
    let status: DrawingCoverageEntry['status'] = source?.representation === 'fallback' ? 'fallback' : shape ? 'native' : 'unsupported'
    let reason = source?.reason
    if (name === 'AlternateContent' && !shape) {
      const supportedEmpty = chosen && getChildren(node, 'Choice').includes(chosen) &&
        orderedChildren(chosen).every(([childName]) => childName === '#text')
      if (supportedEmpty) { status = 'native'; feature = 'empty-choice'; reason = 'Supported empty Choice selected' }
    }
    if (name === 'graphicFrame' && !shape && nested?.representation === 'choice' && orderedChildren(nested.branch).every(([childName]) => childName === '#text')) {
      status = 'native'; feature = 'empty-choice'; reason = 'Supported empty Choice selected'
    }
    const issues = shape?.diagnostics ?? []
    const geometryIssue = issues.find(issue => ['unknown-preset', 'invalid-guide', 'invalid-path', 'invalid-text-rect', 'invalid-extents'].includes(issue.kind))
    if (geometryIssue) { feature = 'geometry'; status = geometryIssue.kind === 'unknown-preset' ? 'unsupported' : 'malformed'; reason = `${geometryIssue.kind}: ${geometryIssue.message}` }
    else if (issues.some(issue => issue.kind === 'missing-image') || name === 'pic' && shape && !shape.image) { feature = 'image'; status = 'malformed'; reason = 'missing image part or relationship' }
    else if (shape?.transformValid === false) { feature = 'transform'; status = 'malformed'; reason = 'invalid-transform' }
    else if (name === 'graphicFrame' && !shape?.content && !shape?.table && feature !== 'empty-choice') { feature = 'graphicData'; status = 'unsupported'; reason = 'unsupported graphicData payload' }
    else if (!shape && status !== 'native') reason = 'No selected static representation'
    const geometryAssessment = retainedGeometry(shape)
    const selectedRepresentation: DrawingCoverageEntry['selectedRepresentation'] = status === 'fallback' && shape?.image ? 'raster-fallback'
      : feature === 'empty-choice' ? 'blank'
      : shape?.content ? contentRepresentation(shape.content)
      : shape?.table ? 'table'
      : status === 'unsupported' || status === 'malformed' ? geometryAssessment.kind === 'paint' ? 'native-shape' : shape?.textBody ? 'text-only' : geometryAssessment.kind === 'unverified' ? 'unverified' : 'none'
      : shape?.image ? 'picture' : shape?.group ? 'group' : shape ? 'native-shape' : 'none'
    const referenced = new Set<string>()
    referenceIds(selectedPayload?.[1] ?? node, referenced, 0, selectedPayload?.[0] ?? name)
    const multipleSelectedObjects = name === 'AlternateContent' && (chosen ?? attempted) &&
      orderedChildren(chosen ?? attempted).filter(([childName]) => drawingNames.has(childName)).length > 1
    entries.push({ partPath, treePath: path, element: name, id: multipleSelectedObjects ? undefined : name === 'AlternateContent' ? own.id ?? source?.id : source?.id ?? own.id, name: multipleSelectedObjects ? undefined : name === 'AlternateContent' ? own.name ?? source?.name : source?.name ?? own.name,
      referenceId: referenced.values().next().value,
      feature, status, selectedRepresentation, representation: source?.representation !== 'native' && source?.representation ? source.representation : nested?.representation ?? inheritedRepresentation ?? (chosen ? getChildren(node, 'Fallback').includes(chosen) ? 'fallback' : 'choice' : 'native'), reason, scope, unit: slide.index,
      ...(shape?.image?.imageSelection ? { imageSelection: shape.image.imageSelection } : {}),
      ...(geometryAssessment.kind === 'unverified' && (status === 'unsupported' || status === 'malformed') ? { assessmentReason: geometryAssessment.reason, analysisLimit: geometryAssessment.limit } : {}) })
  }
  const walk = (parent: XmlNode | undefined, prefix: string, scope: DrawingCoverageEntry['scope'], depth = 0, inheritedRepresentation?: 'choice' | 'fallback') => {
    if (depth >= 128) return
    for (const [index, [name, node]] of orderedChildren(parent).filter(([name]) => name !== '#text').entries()) {
      if (!drawingNames.has(name)) continue
      const path = `${prefix}/${name}[${index}]`
      add(node, name, path, scope, inheritedRepresentation)
      if (name === 'grpSp') walk(node, path, scope === 'diagnostic' ? 'diagnostic' : 'descendant', depth + 1, inheritedRepresentation)
      if (name === 'AlternateContent') {
        const chosen = selectedBranch?.(node)
        const attempted = !chosen ? getChildren(node, 'Choice').find(choice => supportedChoiceRequirements(choice, ['a', 'p', 'c', 'dgm', 'dsp', 'ink']).supported) ?? getChildren(node, 'Choice')[0] : undefined
        const branch = chosen ?? attempted
        if (branch) walk(branch, `${path}/${chosen ? getChildren(node, 'Fallback').includes(chosen) ? 'fallback' : 'choice' : 'attempted'}`, chosen ? 'descendant' : 'diagnostic', depth + 1,
          chosen ? getChildren(node, 'Fallback').includes(chosen) ? 'fallback' : 'choice' : inheritedRepresentation)
      }
    }
  }
  walk(spTree, 'spTree', 'original')
  for (const issue of slide.diagnostics ?? []) {
    const source = issue.source
    const content = issue as Partial<ContentDiagnostic>
    const sourceReferenceId = content.sourceReferenceId ?? content.identity
    if (sourceReferenceId && unselectedReferences.has(sourceReferenceId) && !selectedReferences.has(sourceReferenceId)) continue
    if (issue.kind === 'fallback-representation' && source) {
      const owning = entries.find(entry => entry.partPath === source.partPath && entry.treePath === source.treePath)
      if (owning) {
        owning.representation = 'fallback'
        // The branch selected a fallback, but a failed payload is still a
        // failure. Keep its more precise geometry/image/transform reason.
        if (owning.status !== 'unsupported' && owning.status !== 'malformed') {
          owning.status = 'fallback'
          owning.reason = issue.message
          if (owning.selectedRepresentation === 'picture') owning.selectedRepresentation = 'raster-fallback'
        }
        continue
      }
    }
    const relationshipIssue = issue.kind === 'missing-part' || issue.kind === 'malformed-part' || issue.kind === 'external-reference'
    const associated = content.identity || content.sourceReferenceId ? entries.filter(entry => relationshipIssue || content.sourceReferenceId
      ? coverageIssueMatchesEntry(issue as ContentDiagnostic, entry)
      : entry.partPath === content.partPath && (entry.referenceId === content.identity || entry.id === content.identity || entry.treePath === content.identity)) : []
    if (associated.length && ['missing-part', 'malformed-part', 'external-reference', 'content-cycle', 'group-depth', 'content-depth', 'node-budget', 'unsupported-content'].includes(issue.kind)) {
      for (const affected of associated) {
        affected.status = issue.kind === 'missing-part' || issue.kind === 'malformed-part' ? 'malformed' : 'unsupported'
        affected.reason = content.reason ?? issue.message
        affected.limit = content.limit
        if (!allShapes.some(shape => shape.source?.treePath === affected.treePath && (shape.content || shape.image || shape.table || shape.textBody)) &&
          !retainedGeometryAt(affected)) affected.selectedRepresentation = 'none'
        if (affected.scope === 'original' && (issue.kind === 'missing-part' || issue.kind === 'malformed-part')) {
          for (const entry of entries) if (entry.scope === 'diagnostic' && entry.treePath.startsWith(`${affected.treePath}/`) && entry.referenceId === content.identity) {
            entry.status = 'malformed'; entry.reason = content.reason ?? issue.message
          }
        }
        if (affected.scope !== 'original' && (issue.kind === 'missing-part' || issue.kind === 'malformed-part')) {
          const owner = entries.find(entry => entry.scope === 'original' && entry.element === 'AlternateContent' &&
            affected.treePath.startsWith(`${entry.treePath}/`) && entry.status === 'unsupported' && entry.selectedRepresentation === 'none' && entry.referenceId === content.identity)
          if (owner) { owner.status = 'malformed'; owner.reason = content.reason ?? issue.message }
        }
      }
      continue
    }
    if (source && entries.some(entry => entry.treePath === source.treePath || source.treePath.startsWith(`${entry.treePath}/`))) continue
    if (issue.kind === 'unsupported-content' && entries.some(entry => entry.status === 'fallback' && entry.partPath === partPath &&
      ((issue.feature ?? '').toLowerCase().includes(entry.feature.toLowerCase()) || entry.feature === 'model3D' && (issue.feature ?? '').toLowerCase().includes('model3d')))) continue
    if (!['missing-part', 'malformed-part', 'external-reference', 'content-cycle', 'content-depth', 'group-depth', 'node-budget', 'unsupported-content', 'missing-representation', 'unsupported-object'].includes(issue.kind)) continue
    entries.push({ partPath: content.partPath ?? source?.partPath ?? partPath, treePath: source?.treePath ?? content.identity ?? 'diagnostic',
      element: source?.element ?? issue.kind, id: source?.id ?? content.identity, name: source?.name, feature: 'feature' in issue ? issue.feature ?? issue.kind : issue.kind,
      status: issue.kind === 'malformed-part' || issue.kind === 'missing-part' ? 'malformed' : 'unsupported', selectedRepresentation: 'none', representation: source?.representation ?? 'native',
      reason: content.reason ?? issue.message, scope: 'diagnostic', unit: slide.index, limit: content.limit })
  }
  return entries
}

export function xlsxNodeCoverage(node: XlsxDrawing, unit: number, scope: DrawingCoverageEntry['scope']): DrawingCoverageEntry[] {
  const source = node.source
  const issue = node.drawingGeometry ? resolveGeometry(node.drawingGeometry, node.widthEmu / 9525, node.heightEmu / 9525).issues[0] : undefined
  const feature = issue ? 'geometry' : node.content?.kind ?? (source.emptySelection ? 'empty-choice' : source.element === 'pic' ? 'picture' : source.element === 'grpSp' ? 'group' : source.element === 'contentPart' ? 'ink' : source.element === 'graphicFrame' ? 'graphicData' : 'shape')
  const status: DrawingCoverageEntry['status'] = issue ? issue.kind === 'unknown-preset' ? 'unsupported' : 'malformed' : node.transformValid === false ? 'malformed' : (source.element === 'graphicFrame' || source.element === 'contentPart') && !node.content && !node.children?.length && !source.emptySelection ? 'unsupported' : source.element === 'pic' && !node.image ? 'malformed' : source.representation === 'fallback' ? 'fallback' : 'native'
  const reason = issue ? `${issue.kind}: ${issue.message}` : node.transformValid === false ? 'invalid-transform' : source.emptySelection ? 'Supported empty Choice selected' : status === 'malformed' && source.element === 'pic' ? 'missing image part or relationship' : status === 'unsupported' ? 'unsupported graphicData payload' : source.reason
  const selectedRepresentation: DrawingCoverageEntry['selectedRepresentation'] = node.content ? contentRepresentation(node.content)
    : source.emptySelection ? 'blank' : status === 'fallback' && node.image ? 'raster-fallback' : status === 'unsupported' || status === 'malformed' ? node.textBody ? 'text-only' : 'none'
    : node.image ? 'picture' : node.group ? 'group' : 'native-shape'
  return [{ partPath: source.partPath, treePath: source.treePath, element: source.element, id: source.id, name: source.name, referenceId: source.referenceId,
    feature, status, selectedRepresentation, representation: source.representation, reason, scope, unit,
    ...(node.image?.imageSelection ? { imageSelection: node.image.imageSelection } : {}) }, ...(node.children ?? []).flatMap(child => xlsxNodeCoverage(child, unit, 'descendant'))]
}

export function drawingReport(document: DrawingCoverageDocument) {
  const entries = document.drawingCoverage ?? []
  return {
    entries,
    counts: {
      originalObjects: entries.filter(entry => entry.scope === 'original').length,
      selectedDescendants: entries.filter(entry => entry.scope === 'descendant').length,
      diagnosticEntries: entries.filter(entry => entry.scope === 'diagnostic').length,
    },
  }
}
