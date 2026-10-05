/** Worksheet DrawingML adapter. Placement belongs to the worksheet anchor. */
import type { OfficePackage } from '../core/zip'
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { sniffImageMime } from '../core/images'
import { parseGeometry } from '../drawing/geometry'
import { coverageIssueMatchesEntry, supportedChoiceRequirements, xlsxNodeCoverage } from '../drawing/coverage'
import { attemptedMalformedRelationshipIssue, malformedRelationshipAttempt, prepareCompatibleDrawingContent, prepareDrawingContent, reserveDrawingContent } from '../drawing/content'
import { contentDiagnostic, drawingPartContext, DOCUMENT_DRAWING_NODE_LIMIT, DRAWING_GROUP_DEPTH, partRelationships, referencedPart, reserveDrawingNode } from '../drawing/parts'
import { resolveDrawingStyle, type ThemeContext } from '../drawing/style'
import { parseTextBody, textFontDefaults } from '../drawing/text-parse'
import type { XlsxDrawing, XlsxImage, XlsxSheet } from './types'

const EMU = 9525
const MAX_COLS = 4096
const MAX_ROWS = 16384
const MAX_COMPATIBILITY_DEPTH = 128
const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const finite = (raw: string | undefined): number | undefined => raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
const bool = (raw: string | undefined) => raw === '1' || raw === 'true'
const contentTheme = (theme: ThemeContext) => ({ colors: new Map(Object.entries(theme.palette).map(([key, value]) => [key, value.replace(/^#/, '')])), fonts: new Map([['minorHAnsi', theme.fonts.minor.latin ?? 'Calibri']]) })
const contentAdapters = (theme: ThemeContext) => ({ parseDiagramText: (body: XmlNode, shape: XmlNode) => parseTextBody(body, theme, undefined, textFontDefaults(getChildren(shape, 'style')[0], theme)), parseParagraph: (p: XmlNode) => parseTextBody({ p }, theme).paragraphs[0] })

/** Decide MC branches from their selected descendants only. Empty supported Choices are valid. */
async function selectedAlternative(pkg: OfficePackage, owner: string, alternate: XmlNode, theme: ThemeContext, depth: number, carrier?: XmlNode): Promise<{ branch: XmlNode; mode: 'choice' | 'fallback' } | undefined> {
  if (depth >= MAX_COMPATIBILITY_DEPTH) {
    contentDiagnostic(drawingPartContext(pkg), 'unsupported-content', owner, 'AlternateContent', { identity: owner, reason: 'compatibility-depth', limit: MAX_COMPATIBILITY_DEPTH })
    return undefined
  }
  const diagnostics = drawingPartContext(pkg).diagnostics
  const failed = new Set<(typeof diagnostics)[number]>()
  const discardFailed = () => { for (let i = diagnostics.length - 1; i >= 0; i--) if (failed.has(diagnostics[i])) diagnostics.splice(i, 1) }
  for (const choice of getChildren(alternate, 'Choice')) {
    if (!supportedChoiceRequirements(choice, ['a', 'xdr', 'c', 'dgm', 'dsp', 'ink']).supported) continue
    const before = diagnostics.length
    if (await branchAvailable(pkg, owner, choice, theme, depth + 1, carrier)) { discardFailed(); return { branch: choice, mode: 'choice' } }
    for (const diagnostic of diagnostics.slice(before)) failed.add(diagnostic)
  }
  const fallback = child(alternate, 'Fallback')
  if (fallback && await branchAvailable(pkg, owner, fallback, theme, depth + 1, carrier)) { discardFailed(); return { branch: fallback, mode: 'fallback' } }
  return undefined
}

async function branchAvailable(pkg: OfficePackage, owner: string, node: XmlNode, theme: ThemeContext, depth: number, carrier?: XmlNode): Promise<boolean> {
  if (depth >= MAX_COMPATIBILITY_DEPTH) return false
  let selected = false, failed = false
  for (const [name, descendant] of orderedChildren(node)) {
    if (name === '#text') continue
    if (name === 'sp' || name === 'pic' || name === 'grpSp' || name === 'cxnSp') { selected = true; continue }
    if (name === 'AlternateContent') {
      if (await selectedAlternative(pkg, owner, descendant, theme, depth + 1, carrier)) selected = true
      else failed = true
      continue
    }
    if (name === 'chart' || name === 'relIds' || name === 'contentPart' || name === 'wsp') {
      if (await prepareDrawingContent(pkg, { ...(carrier ? { '@attrs': attrs(carrier) } : {}), [name]: descendant }, owner, contentTheme(theme), theme, contentAdapters(theme))) selected = true
      else failed = true
      continue
    }
    if (name === 'graphicFrame' || name === 'graphic' || name === 'graphicData') {
      if (await branchAvailable(pkg, owner, descendant, theme, depth + 1, name === 'graphicData' ? descendant : carrier)) selected = true
      else failed = true
    }
  }
  return selected || !failed
}

/** Physical prefixes may extend to a meaningful anchor, regardless of the output viewport. */
function prefixes(sheet: XlsxSheet, axis: 'col' | 'row', through: number): number[] {
  const limit = axis === 'col' ? MAX_COLS : MAX_ROWS
  const n = Math.min(limit, Math.max(1, through + 1))
  const values = new Array<number>(n).fill(axis === 'col' ? 64 : 20)
  if (axis === 'col') {
    for (const spec of sheet.cols) {
      if (!Number.isSafeInteger(spec.min) || !Number.isSafeInteger(spec.max)) continue
      for (let i = Math.max(0, spec.min); i <= spec.max && i < n; i++) values[i] = spec.hidden ? 0 : spec.widthChars !== undefined && Number.isFinite(spec.widthChars) && spec.widthChars >= 0 ? Math.round(spec.widthChars * 7 + 5) : 64
    }
  } else {
    for (const row of sheet.rows) if (Number.isSafeInteger(row.index) && row.index >= 0 && row.index < n) {
      values[row.index] = row.hidden ? 0 : row.customHeight && row.heightPt !== undefined && Number.isFinite(row.heightPt) && row.heightPt >= 0 ? Math.round(row.heightPt * 96 / 72) : 20
    }
  }
  const out = [0]
  for (const value of values) out.push(out[out.length - 1] + value)
  return out
}

function marker(sheet: XlsxSheet, node: XmlNode | undefined): { x: number; y: number; col: number; row: number } | undefined {
  const col = finite(textOf(child(node, 'col'))), row = finite(textOf(child(node, 'row')))
  const xo = finite(textOf(child(node, 'colOff'))), yo = finite(textOf(child(node, 'rowOff')))
  if (col === undefined || row === undefined || xo === undefined || yo === undefined || !Number.isSafeInteger(col) || !Number.isSafeInteger(row) || col < 0 || row < 0 || col >= MAX_COLS || row >= MAX_ROWS) return undefined
  sheet.drawingMarkers ??= { maxCol: 0, maxRow: 0 }
  sheet.drawingMarkers.maxCol = Math.max(sheet.drawingMarkers.maxCol, col)
  sheet.drawingMarkers.maxRow = Math.max(sheet.drawingMarkers.maxRow, row)
  return { x: (prefixes(sheet, 'col', col)[col] * EMU) + xo, y: (prefixes(sheet, 'row', row)[row] * EMU) + yo, col, row }
}

function transform(node: XmlNode | undefined): Pick<XlsxDrawing, 'xEmu' | 'yEmu' | 'widthEmu' | 'heightEmu' | 'rotationDeg' | 'flipH' | 'flipV' | 'transformValid'> {
  const oa = attrs(child(node, 'off')), ea = attrs(child(node, 'ext')), xa = attrs(node)
  const x = finite(oa.x), y = finite(oa.y), w = finite(ea.cx), h = finite(ea.cy), rot = xa.rot === undefined ? 0 : finite(xa.rot)
  const valid = x !== undefined && y !== undefined && w !== undefined && h !== undefined && rot !== undefined && w >= 0 && h >= 0
  return { xEmu: x ?? 0, yEmu: y ?? 0, widthEmu: w ?? 0, heightEmu: h ?? 0, rotationDeg: (rot ?? 0) / 60000, flipH: bool(xa.flipH), flipV: bool(xa.flipV), transformValid: valid }
}

function anchorPlacement(sheet: XlsxSheet, name: string, node: XmlNode): Pick<XlsxDrawing, 'xEmu' | 'yEmu' | 'widthEmu' | 'heightEmu'> | undefined {
  if (name === 'absoluteAnchor') {
    const pa = attrs(child(node, 'pos')), ea = attrs(child(node, 'ext'))
    const x = finite(pa.x), y = finite(pa.y), w = finite(ea.cx), h = finite(ea.cy)
    return x !== undefined && y !== undefined && w !== undefined && h !== undefined && w >= 0 && h >= 0 ? { xEmu: x, yEmu: y, widthEmu: w, heightEmu: h } : undefined
  }
  const from = marker(sheet, child(node, 'from'))
  if (!from) return undefined
  if (name === 'oneCellAnchor') {
    const ea = attrs(child(node, 'ext')), w = finite(ea.cx), h = finite(ea.cy)
    return w !== undefined && h !== undefined && w >= 0 && h >= 0 ? { xEmu: from.x, yEmu: from.y, widthEmu: w, heightEmu: h } : undefined
  }
  const to = marker(sheet, child(node, 'to'))
  return to && to.x >= from.x && to.y >= from.y ? { xEmu: from.x, yEmu: from.y, widthEmu: to.x - from.x, heightEmu: to.y - from.y } : undefined
}

function rectFractions(node: XmlNode | undefined): { l: number; t: number; r: number; b: number } | undefined {
  if (!node) return undefined
  const a = attrs(node), values = [a.l, a.t, a.r, a.b].map(value => value === undefined ? 0 : finite(value))
  if (values.some(value => value === undefined)) return undefined
  return { l: values[0]! / 100000, t: values[1]! / 100000, r: values[2]! / 100000, b: values[3]! / 100000 }
}

async function imageFor(pkg: OfficePackage, owner: string, node: XmlNode, shape: XlsxDrawing): Promise<void> {
  const fill = child(node, 'blipFill') ?? child(child(node, 'spPr'), 'blipFill')
  const blip = child(fill, 'blip'), rid = attrs(blip).embed
  if (!rid) return
  shape.source.referenceId = rid
  const diagnosticStart = drawingPartContext(pkg).diagnostics.length
  const path = await referencedPart(pkg, owner, rid)
  if (!path) malformedRelationshipAttempt(pkg, owner, rid, 'image', diagnosticStart)
  if (!path) return
  let data: Uint8Array | undefined
  try { data = await pkg.bytes(path) }
  catch { contentDiagnostic(drawingPartContext(pkg), 'malformed-part', path, 'image', { identity: path, reason: 'unreadable-media' }); return }
  if (!data) return
  let opacity = 1
  for (const effect of getChildren(blip, 'alphaModFix')) {
    const amount = finite(attrs(effect).amt)
    if (amount !== undefined && amount >= 0 && amount <= 100000) opacity *= amount / 100000
  }
  shape.image = { data, mime: sniffImageMime(data), partPath: path, opacity,
    srcRect: rectFractions(child(fill, 'srcRect')),
    fillRect: rectFractions(child(child(fill, 'stretch'), 'fillRect')) }
}

function source(owner: string, path: string, name: string, node: XmlNode, representation: 'native' | 'choice' | 'fallback'): XlsxDrawing['source'] {
  const nv = child(node, name === 'grpSp' ? 'nvGrpSpPr' : name === 'pic' ? 'nvPicPr' : name === 'graphicFrame' ? 'nvGraphicFramePr' : 'nvSpPr')
  const a = attrs(child(nv, 'cNvPr'))
  return { partPath: owner, treePath: path, element: name, id: a.id ?? (name === 'contentPart' ? attrs(node).id : undefined), name: a.name, representation }
}

async function parseObject(pkg: OfficePackage, owner: string, name: string, node: XmlNode, theme: ThemeContext, path: string, depth: number, representation: 'native' | 'choice' | 'fallback' = 'native', parentCarrier?: XmlNode, anchorTopLevel = false): Promise<XlsxDrawing | undefined> {
  const context = drawingPartContext(pkg)
  if (name === 'grpSp' && depth >= DRAWING_GROUP_DEPTH) {
    contentDiagnostic(context, 'group-depth', owner, path, { identity: attrs(child(child(node, 'nvGrpSpPr'), 'cNvPr')).id, reason: 'group-depth', limit: DRAWING_GROUP_DEPTH })
    return undefined
  }
  if (!reserveDrawingNode(context, owner)) return undefined
  const spPr = child(node, name === 'grpSp' ? 'grpSpPr' : 'spPr')
  const localTransform = child(spPr, 'xfrm')
  const t = transform(localTransform ?? (name === 'graphicFrame' ? child(node, 'xfrm') : undefined))
  // A worksheet anchor owns the top-level box. Pictures and native shapes may
  // omit their optional local transform, but an explicitly malformed one is invalid.
  if (anchorTopLevel && !localTransform && (name === 'sp' || name === 'cxnSp' || name === 'pic')) t.transformValid = true
  const result: XlsxDrawing = { ...t, source: source(owner, path, name, node, representation) }
  if (name === 'grpSp') {
    const xfrm = child(spPr, 'xfrm'), ca = attrs(child(xfrm, 'chOff')), ce = attrs(child(xfrm, 'chExt'))
    const cx = finite(ca.x), cy = finite(ca.y), cw = finite(ce.cx), ch = finite(ce.cy)
    if (cx === undefined || cy === undefined || cw === undefined || ch === undefined || cw <= 0 || ch <= 0) { result.transformValid = false; return result }
    result.group = { off: { x: t.xEmu, y: t.yEmu }, ext: { width: t.widthEmu, height: t.heightEmu }, chOff: { x: cx, y: cy }, chExt: { width: cw, height: ch } }
    result.children = await parseObjects(pkg, owner, node, theme, path, depth + 1, representation, parentCarrier)
  } else if (name === 'sp' || name === 'cxnSp') {
    result.drawingGeometry = parseGeometry(spPr)
    result.drawingStyle = resolveDrawingStyle(spPr, child(node, 'style'), theme, name === 'cxnSp' ? { fill: { kind: 'none' } } : {})
    const body = child(node, 'txBody')
    if (body) result.textBody = parseTextBody(body, theme, undefined, textFontDefaults(child(node, 'style'), theme))
    await imageFor(pkg, owner, node, result)
  } else if (name === 'pic') await imageFor(pkg, owner, node, result)
  else if (name === 'graphicFrame' || name === 'contentPart') {
    const palette = contentTheme(theme), adapters = contentAdapters(theme)
    const graphic = child(child(node, 'graphic'), 'graphicData')
    const carrier = name === 'contentPart' ? { ...(parentCarrier ? { '@attrs': attrs(parentCarrier) } : {}), contentPart: node } : graphic
    if (carrier || name === 'graphicFrame') {
      const direct = carrier ? await prepareDrawingContent(pkg, carrier, owner, palette, theme, adapters) : undefined
      const compatible = !direct && name === 'graphicFrame' ? await prepareCompatibleDrawingContent(pkg, node, owner, palette, theme, adapters) : undefined
      const prepared = direct ?? compatible?.content
      const reference = compatible?.reference ?? child(carrier, 'chart') ?? child(carrier, 'relIds') ?? child(carrier, 'contentPart')
      result.source.referenceId = attrs(reference).id ?? attrs(reference).dm
      const selected = compatible?.selections.at(-1)
      if (selected) {
        result.source.representation = selected.representation
        result.source.reason = selected.reason
        result.source.emptySelection = !prepared && orderedChildren(selected.node).every(([childName]) => childName === '#text')
      }
      if (prepared && reserveDrawingContent(pkg, prepared, owner, false)) {
        result.content = prepared
      }
      else if (t.widthEmu > 0 && t.heightEmu > 0) {
        const native = await parseObjects(pkg, owner, node, theme, `${path}/native`, depth)
        if (native.length) {
          result.group = { off: { x: t.xEmu, y: t.yEmu }, ext: { width: t.widthEmu, height: t.heightEmu }, chOff: { x: 0, y: 0 }, chExt: { width: t.widthEmu, height: t.heightEmu } }
          result.children = native
        }
      }
    }
  }
  return result
}

async function parseObjects(pkg: OfficePackage, owner: string, container: XmlNode, theme: ThemeContext, path: string, depth = 0, representation: 'native' | 'choice' | 'fallback' = 'native', carrier?: XmlNode, anchorTopLevel = false): Promise<XlsxDrawing[]> {
  const out: XlsxDrawing[] = []
  let index = 0
  for (const [name, node] of orderedChildren(container)) {
    if (name === '#text' || name === 'clientData') continue
    if (!['AlternateContent', 'sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame', 'contentPart', 'graphic', 'graphicData'].includes(name)) continue
    const treePath = `${path}/${name}[${index++}]`
    if (name === 'AlternateContent') {
      const selected = await selectedAlternative(pkg, owner, node, theme, 0, carrier)
      if (selected) out.push(...await parseObjects(pkg, owner, selected.branch, theme, treePath, depth, selected.mode, carrier, anchorTopLevel))
    } else if (name === 'graphic' || name === 'graphicData') {
      out.push(...await parseObjects(pkg, owner, node, theme, treePath, depth, representation, name === 'graphicData' ? node : carrier, anchorTopLevel))
    } else if (['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame', 'contentPart'].includes(name)) {
      const drawing = await parseObject(pkg, owner, name, node, theme, treePath, depth, representation, carrier, anchorTopLevel)
      if (drawing) out.push(drawing)
    }
  }
  return out
}

/** Resolve each worksheet-owned drawing relationship and retain anchor order. */
export async function parseWorksheetDrawings(pkg: OfficePackage, sheet: XlsxSheet, worksheetPath: string, root: XmlNode, theme: ThemeContext): Promise<void> {
  const context = drawingPartContext(pkg), start = context.diagnostics.length
  sheet.drawingCoverage = []
  for (const drawingRef of getChildren(root, 'drawing')) {
    const id = attrs(drawingRef).id
    if (!id) continue
    const owner = await referencedPart(pkg, worksheetPath, id)
    if (!owner) continue
    let drawingRoot: XmlNode | undefined
    try { drawingRoot = await pkg.xmlOrdered(owner) }
    catch { contentDiagnostic(context, 'malformed-part', owner, undefined, { identity: id, reason: 'invalid-drawing-xml' }); continue }
    if (!drawingRoot) continue
    // Resolve relationships through the shared cached optional-part path.
    await partRelationships(pkg, owner)
    let anchorIndex = 0
    for (const [name, anchor] of orderedChildren(drawingRoot)) {
      if (!['oneCellAnchor', 'twoCellAnchor', 'absoluteAnchor'].includes(name)) continue
      const anchorPath = `${name}[${anchorIndex++}]`
      const placement = anchorPlacement(sheet, name, anchor)
      if (!placement) {
        contentDiagnostic(context, 'unsupported-content', owner, name, { identity: anchorPath, reason: 'invalid-anchor' })
        sheet.drawingCoverage.push({ partPath: owner, treePath: anchorPath, element: name, feature: 'anchor', status: 'malformed', selectedRepresentation: 'none', representation: 'native', reason: 'invalid-anchor', scope: 'original', unit: 0 })
        continue
      }
      const objects = await parseObjects(pkg, owner, anchor, theme, anchorPath, 0, 'native', undefined, true)
      if (!objects.length) {
        const exhausted = context.nodes >= DOCUMENT_DRAWING_NODE_LIMIT && context.diagnostics.some(issue => issue.kind === 'node-budget' && issue.reason === 'document-budget')
        const selected = await Promise.all(getChildren(anchor, 'AlternateContent').map(alternate => selectedAlternative(pkg, owner, alternate, theme, 0)))
        const blank = selected.find(choice => choice && orderedChildren(choice.branch).every(([childName]) => childName === '#text'))
        sheet.drawingCoverage.push({ partPath: owner, treePath: anchorPath, element: name,
          feature: blank ? 'empty-choice' : 'anchor', status: blank ? blank.mode === 'fallback' ? 'fallback' : 'native' : 'unsupported', selectedRepresentation: blank ? 'blank' : 'none',
          representation: blank?.mode ?? 'native', reason: blank ? 'Selected empty compatibility branch' : exhausted ? 'document-budget' : 'No selected drawing object',
          limit: exhausted ? DOCUMENT_DRAWING_NODE_LIMIT : undefined, scope: 'original', unit: 0 })
      }
      for (const [objectIndex, object] of objects.entries()) {
        Object.assign(object, placement)
        if (object.source.element === 'contentPart') object.transformValid = true
        if (object.group) { object.group.off = { x: placement.xEmu, y: placement.yEmu }; object.group.ext = { width: placement.widthEmu, height: placement.heightEmu } }
        ;(sheet.drawings ??= []).push(object)
        sheet.drawingCoverage.push(...xlsxNodeCoverage(object, 0, objectIndex === 0 ? 'original' : 'descendant'))
      }
    }
  }
  sheet.drawingDiagnostics = context.diagnostics.slice(start)
  // A cached relationship failure may have been recorded on an earlier sheet.
  // Match by both owner part and selected relationship, so an identical rId in
  // another drawing part cannot inherit this failure.
  for (const issue of context.diagnostics) {
    if (attemptedMalformedRelationshipIssue(pkg, issue)) continue
    const affected = sheet.drawingCoverage.filter(entry => coverageIssueMatchesEntry(issue, entry))
    if (issue.kind === 'malformed-part' && issue.feature === 'image') for (const entry of sheet.drawingCoverage) {
      if (!entry.referenceId || affected.includes(entry) || entry.element === 'graphicFrame' || entry.element === 'contentPart') continue
      const relation = (await partRelationships(pkg, entry.partPath)).get(entry.referenceId)
      if (relation?.path === issue.partPath) affected.push(entry)
    }
    if (affected.length && ['missing-part', 'malformed-part', 'external-reference', 'unsupported-content', 'content-cycle', 'content-depth', 'group-depth'].includes(issue.kind)) {
      for (const entry of affected) {
        if (issue.kind === 'missing-part' || issue.kind === 'malformed-part') entry.status = 'malformed'
        else if (issue.kind === 'external-reference') entry.status = 'unsupported'
        entry.reason = issue.reason ?? issue.message
        entry.limit = issue.limit
        if ((entry.element === 'contentPart' || entry.element === 'graphicFrame') && entry.selectedRepresentation !== 'raster-fallback' && !sheet.drawings?.some(drawing => drawing.source.treePath === entry.treePath && drawing.content)) entry.selectedRepresentation = 'none'
      }
      continue
    }
    if (!sheet.drawingDiagnostics.includes(issue)) continue
    if (issue.kind === 'node-budget' && issue.reason === 'document-budget' && sheet.drawingCoverage.some(entry => entry.partPath === issue.partPath && entry.reason === 'document-budget' && entry.limit === issue.limit)) continue
    if (sheet.drawingCoverage.some(entry => entry.treePath === issue.identity && entry.reason === issue.reason)) continue
    sheet.drawingCoverage.push({ partPath: issue.partPath, treePath: issue.identity ?? issue.partPath, element: issue.kind,
      id: issue.identity, feature: issue.feature ?? issue.kind, status: issue.kind === 'missing-part' || issue.kind === 'malformed-part' ? 'malformed' : 'unsupported', selectedRepresentation: 'none',
      representation: 'native', reason: issue.reason ?? issue.message, scope: 'diagnostic', unit: 0, limit: issue.limit })
  }
  sheet.drawingTheme = theme
}

/** Stable first-use media indexing across all worksheet scenes. */
export function collectXlsxImages(sheets: XlsxSheet[]): XlsxImage[] {
  const images: XlsxImage[] = [], indices = new Map<string, number>()
  const walk = (nodes: XlsxDrawing[]) => {
    for (const node of nodes) {
      if (node.image?.data && node.image.partPath) {
        let index = indices.get(node.image.partPath)
        if (index === undefined) { index = images.length; indices.set(node.image.partPath, index); images.push({ data: node.image.data, mime: node.image.mime, partPath: node.image.partPath }) }
        node.imageIndex = index
      }
      if (node.children) walk(node.children)
    }
  }
  for (const sheet of sheets) walk(sheet.drawings ?? [])
  return images
}
