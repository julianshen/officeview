import { attemptedMalformedRelationshipIssue, malformedRelationshipAttempt, prepareCompatibleDrawingContent, prepareDrawingContent, reserveDrawingContent, type DrawingContent } from '../drawing/content'
import { drawingPartContext, partRelationshipNodes, reserveDrawingNode, DRAWING_GROUP_DEPTH, contentDiagnostic, type ContentDiagnostic } from '../drawing/parts'
import { parseEmbeddedFonts } from '../core/fonts/parts'
/** Parse PPTX parts (presentation.xml, slides/slideN.xml) into PptxDocument. */
import type { OfficePackage } from '../core/zip'
import { attrs, elementChildren, getChildren, orderedChildren, parseXmlOrdered, textOf, type XmlNode } from '../core/xml'
import { hexRgbToCss } from '../core/color'
import type { PptxDiagnostic, PptxDocument, PptxImageRef, PptxShape, PptxSlide, PptxSource, PptxTable, PptxTableBorders, PptxTableCell, PptxTableRow, PptxTextBody, PptxParagraph } from './types'
import { pptxCoverage, supportedChoiceRequirements } from '../drawing/coverage'
import { sniffImageMime } from '../core/images'
import { emuToPx } from '../core/geometry'
import { parseGeometry, resolveGeometry } from '../drawing/geometry'
import { parseDrawingColor, parseFillDefinition, parseThemeContext, resolveDrawingColor, resolveDrawingStyle, resolveFill, type DrawingColor, type DrawingIssue, type ThemeContext } from '../drawing/style'

import { parseTextBody, textFontDefaults, type InheritedTextLayer } from './text-parse'

function num(v: string | undefined, dflt = 0): number {
  const n = v === undefined ? NaN : parseFloat(v)
  return Number.isFinite(n) ? n : dflt
}

function colorOf(srgbClr: XmlNode | undefined): string | undefined {
  if (!srgbClr) return undefined
  return hexRgbToCss(attrs(srgbClr).val as string)
}

/**
 * Placeholder geometry from a slide layout (or its master). Real decks rely on
 * this heavily: a title/body placeholder on the slide carries only `p:ph`, and
 * its position lives in the layout. Without it those shapes render at 0x0.
 */
async function layoutPlaceholderGeometry(
  pkg: OfficePackage,
  slidePath: string,
): Promise<Map<string, { x: number; y: number; w: number; h: number }>> {
  const out = new Map<string, { x: number; y: number; w: number; h: number }>()
  const rels = await partRelationshipNodes(pkg, slidePath)
  let layoutPath: string | undefined
  for (const rel of rels) {
    const a = attrs(rel)
    if (a.TargetMode !== 'External' && (a.Type as string | undefined)?.endsWith('/slideLayout') && a.Target) {
      layoutPath = resolveTarget(slidePath, a.Target)
      break
    }
  }
  if (!layoutPath) return out

  const collect = (root: XmlNode | undefined): void => {
    if (!root) return
    const cSld = getChildren(root, 'cSld')[0]
    const spTree = cSld ? getChildren(cSld, 'spTree')[0] : undefined
    if (!spTree) return
    for (const [, node] of elementChildren(spTree)) {
      const spPr = getChildren(node, 'spPr')[0]
      const xfrm = spPr ? getChildren(spPr, 'xfrm')[0] : undefined
      if (!xfrm) continue
      const oa = attrs(getChildren(xfrm, 'off')[0])
      const ea = attrs(getChildren(xfrm, 'ext')[0])
      const nvSpPr = getChildren(node, 'nvSpPr')[0]
      const ph = nvSpPr ? getChildren(getChildren(nvSpPr, 'nvPr')[0], 'ph')[0] : undefined
      if (!ph) continue
      const pa = attrs(ph)
      const type = (pa.type as string | undefined) ?? 'body'
      const idx = pa.idx !== undefined ? parseInt(pa.idx as string, 10) || 0 : 0
      // Collect the closest layout first; the master only fills missing slots.
      if (out.has(`${type}|${idx}`)) continue
      out.set(`${type}|${idx}`, {
        x: num(oa.x as string),
        y: num(oa.y as string),
        w: num(ea.cx as string),
        h: num(ea.cy as string),
      })
    }
  }

  collect(await pkg.xml(layoutPath))
  // a layout may itself defer geometry to the master
  const layoutRels = await partRelationshipNodes(pkg, layoutPath)
  if (layoutRels.length) {
    for (const rel of layoutRels) {
      const a = attrs(rel)
      if (a.TargetMode !== 'External' && (a.Type as string | undefined)?.endsWith('/slideMaster') && a.Target) {
        collect(await pkg.xml(resolveTarget(layoutPath, a.Target)))
        break
      }
    }
  }
  return out
}

/** Reject malformed transforms instead of coercing their numbers into a painted shape. */
function parseTransform(xfrm: XmlNode | undefined): Pick<PptxShape, 'xEmu' | 'yEmu' | 'widthEmu' | 'heightEmu' | 'rotationDeg' | 'flipH' | 'flipV' | 'transformValid'> {
  const oa = attrs(getChildren(xfrm, 'off')[0]), ea = attrs(getChildren(xfrm, 'ext')[0]), xa = attrs(xfrm)
  const values = [oa.x, oa.y, ea.cx, ea.cy, xa.rot]
  const valid = values.every(value => value === undefined || value.trim() !== '' && Number.isFinite(Number(value))) &&
    num(ea.cx) >= 0 && num(ea.cy) >= 0
  return {
    xEmu: num(oa.x), yEmu: num(oa.y), widthEmu: num(ea.cx), heightEmu: num(ea.cy),
    rotationDeg: num(xa.rot) / 60000,
    flipH: xa.flipH === '1' || xa.flipH === 'true', flipV: xa.flipV === '1' || xa.flipV === 'true',
    transformValid: valid,
  }
}

function cssColor(color: DrawingColor): string {
  return `#${[color.r, color.g, color.b].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()}`
}

interface SlideTextInheritance {
  layers(type: string, idx: number, placeholder: boolean): InheritedTextLayer[]
}

function parseShape(sp: XmlNode, slideImages: Map<string, PptxImageRef>, theme: ThemeContext, source: PptxSource, textDefaults?: XmlNode, inheritance?: SlideTextInheritance): PptxShape | undefined {
  const spPr = getChildren(sp, 'spPr')[0]
  if (!spPr) return undefined
  const xfrm = getChildren(spPr, 'xfrm')[0]
  const drawingGeometry = parseGeometry(spPr)
  const prst = drawingGeometry.preset
  const drawingStyle = resolveDrawingStyle(spPr, getChildren(sp, 'style')[0], theme,
    source.element === 'cxnSp' ? { fill: { kind: 'none' } } : {})
  const shape: PptxShape = {
    ...parseTransform(xfrm),
    geometry: prst === 'ellipse' ? 'ellipse' : prst === 'roundRect' ? 'roundRect' : prst === 'rect' ? 'rect' : 'other',
    drawingGeometry, drawingStyle, presetName: prst, source, diagnostics: [],
    // Retain the simple color fields consumed by existing clients.
    fill: drawingStyle.fill?.kind === 'solid' ? cssColor(drawingStyle.fill.color) : undefined,
    line: drawingStyle.line?.fill?.kind === 'solid' ? { color: cssColor(drawingStyle.line.fill.color), widthEmu: (drawingStyle.line.width ?? 1) * 9525 } : undefined,
  }
  const txBody = getChildren(sp, 'txBody')[0]
  if (txBody) {
    const ph = getChildren(getChildren(getChildren(sp, 'nvSpPr')[0], 'nvPr')[0], 'ph')[0]
    const pa = attrs(ph)
    const inherited = inheritance?.layers(pa.type ?? 'body', pa.idx === undefined ? 0 : num(pa.idx), !!ph)
    shape.textBody = parseTextBody(txBody, theme, textDefaults, textFontDefaults(getChildren(sp, 'style')[0], theme), inherited)
    shape.diagnostics!.push(...(shape.textBody.diagnostics ?? []))
  }
  const nvSpPr = getChildren(sp, 'nvSpPr')[0]
  const ph = getChildren(getChildren(nvSpPr, 'nvPr')[0], 'ph')[0]
  if (ph) {
    const pa = attrs(ph)
    shape.placeholder = { type: pa.type ?? 'body', idx: pa.idx !== undefined ? parseInt(pa.idx, 10) || 0 : 0 }
  }
  // Native fallbacks may be p:sp with a:blipFill inside spPr, even alongside noFill.
  const blipFill = getChildren(sp, 'blipFill')[0] ?? getChildren(spPr, 'blipFill')[0]
  if (blipFill) {
    const blip = getChildren(blipFill, 'blip')[0]
    const rid = attrs(blip).embed
    let opacity: number | undefined
    for (const [name, effect] of orderedChildren(blip)) {
      if (name === '#text') continue
      if (name === 'alphaModFix') {
        const amount = attrs(effect).amt
        const value = amount === undefined ? 100000 : /^[+-]?\d+$/.test(amount.trim()) ? Number(amount) : NaN
        if (!Number.isInteger(value) || value < 0 || value > 2147483647) {
          shape.diagnostics!.push({ kind: 'invalid-paint', message: `Invalid fixed alpha amount ${amount ?? ''}`, feature: name, source })
        } else if (value > 100000) {
          // Canvas globalAlpha can attenuate, but cannot amplify source pixel alpha.
          shape.diagnostics!.push({ kind: 'unsupported-effect', message: 'Image alpha amplification is deferred', feature: name, source })
        } else opacity = (opacity ?? 1) * value / 100000
      } else shape.diagnostics!.push({ kind: 'unsupported-effect', message: `Image effect ${name} is deferred`, feature: name, source })
    }
    const image = rid ? slideImages.get(rid) : undefined
    if (image) {
      const relativeRect = (node: XmlNode): { l: number; t: number; r: number; b: number } => {
        const a = attrs(node)
        return { l: num(a.l) / 100000, t: num(a.t) / 100000, r: num(a.r) / 100000, b: num(a.b) / 100000 }
      }
      const srcRect = getChildren(blipFill, 'srcRect')[0]
      const fillRect = getChildren(getChildren(blipFill, 'stretch')[0], 'fillRect')[0]
      shape.image = { ...image,
        ...(opacity !== undefined ? { opacity } : {}),
        ...(srcRect ? { srcRect: relativeRect(srcRect) } : {}),
        ...(fillRect ? { fillRect: relativeRect(fillRect) } : {}),
      }
      // The adapter paints blipFill natively; shared paint's deferred-fill issue does not apply.
      drawingStyle.issues = drawingStyle.issues.filter(issue => !(issue.kind === 'unsupported-fill' && issue.feature === 'blipFill'))
    } else {
      shape.diagnostics!.push({ kind: 'missing-image', message: `No usable embedded image for ${rid ?? 'blipFill'}`, feature: 'blipFill', source })
    }
  }
  return shape
}

/** Resolve a relationship Target against the part's directory. */
function resolveTarget(partPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const base = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/')) : ''
  const segs = base ? base.split('/') : []
  for (const seg of target.split('/')) {
    if (seg === '.' || seg === '') continue
    if (seg === '..') segs.pop()
    else segs.push(seg)
  }
  return segs.join('/')
}

/** Load image parts referenced by a part's .rels file, keyed by rId. */
async function loadSlideImages(pkg: OfficePackage, partPath: string): Promise<Map<string, PptxImageRef>> {
  const out = new Map<string, PptxImageRef>()
  const rels = await partRelationshipNodes(pkg, partPath)
  // one PptxImageRef per media part, so multiple rIds for the same part dedupe
  const byPath = new Map<string, PptxImageRef>()
  for (const rel of rels) {
    const a = attrs(rel)
    const type = a.Type as string | undefined
    if (!type || !type.includes('/image') || a.TargetMode === 'External' || !a.Target) continue
    const path = resolveTarget(partPath, (a.Target as string) ?? '')
    let ref = byPath.get(path)
    if (!ref) {
      const data = await pkg.bytes(path)
      if (!data) continue
      ref = { data, mime: sniffImageMime(data), partPath: path }
      byPath.set(path, ref)
    }
    if (a.Id) out.set(a.Id, ref)
  }
  return out
}

/** Follow a local slide/layout relationship. An absent or external part cannot inherit a background. */
async function relatedPath(pkg: OfficePackage, partPath: string, kind: 'slideLayout' | 'slideMaster'): Promise<string | undefined> {
  for (const rel of await partRelationshipNodes(pkg, partPath)) {
    const a = attrs(rel)
    if (a.TargetMode !== 'External' && a.Target && a.Type?.endsWith(`/${kind}`)) return resolveTarget(partPath, a.Target)
  }
  return undefined
}

/** Only placeholder metadata is inherited; content and source shape identity stay on the slide. */
async function slideTextInheritance(pkg: OfficePackage, slidePath: string): Promise<SlideTextInheritance> {
  const layoutPath = await relatedPath(pkg, slidePath, 'slideLayout')
  const layout = layoutPath ? await pkg.xmlOrdered(layoutPath) : undefined
  const masterPath = layoutPath ? await relatedPath(pkg, layoutPath, 'slideMaster') : await relatedPath(pkg, slidePath, 'slideMaster')
  const master = masterPath ? await pkg.xmlOrdered(masterPath) : undefined
  interface PlaceholderBody { type: string; idx: number; body: XmlNode }
  const placeholderBodies = (root: XmlNode | undefined): PlaceholderBody[] => {
    const result: PlaceholderBody[] = []
    const tree = getChildren(getChildren(root, 'cSld')[0], 'spTree')[0]
    for (const sp of getChildren(tree, 'sp')) {
      const ph = getChildren(getChildren(getChildren(sp, 'nvSpPr')[0], 'nvPr')[0], 'ph')[0]
      const body = getChildren(sp, 'txBody')[0]
      if (!ph || !body) continue
      const pa = attrs(ph)
      result.push({ type: pa.type ?? 'body', idx: pa.idx === undefined ? 0 : num(pa.idx), body })
    }
    return result
  }
  const masterBodies = placeholderBodies(master), layoutBodies = placeholderBodies(layout)
  const txStyles = getChildren(master, 'txStyles')[0]
  const contentRoles = new Set(['body', 'subTitle', 'obj', 'pic', 'chart', 'clipArt', 'dgm', 'media', 'tbl'])
  const category = (type: string) => type === 'title' || type === 'ctrTitle' ? 'title'
    : contentRoles.has(type) ? 'body' : type
  return { layers(type, idx, placeholder) {
    const result: InheritedTextLayer[] = []
    // Slide placeholders bind to the layout primarily by index. The master
    // has its own indices: bind that hop by placeholder role, then use its text style.
    const layoutBody = placeholder ? layoutBodies.find(item => item.idx === idx && item.type === type)
      ?? layoutBodies.find(item => item.idx === idx)
      ?? layoutBodies.find(item => category(item.type) === category(type)) : undefined
    const role = category(layoutBody?.type ?? type)
    const styleName = placeholder && role === 'title' ? 'titleStyle' : placeholder && role === 'body' ? 'bodyStyle' : 'otherStyle'
    const style = getChildren(txStyles, styleName)[0]
    if (style) result.push({ style, origin: 'master' })
    const masterBody = placeholder ? masterBodies.find(item => category(item.type) === role && item.idx === (layoutBody?.idx ?? idx))
      ?? masterBodies.find(item => category(item.type) === role) : undefined
    if (masterBody) result.push({ body: masterBody.body, origin: 'master' })
    if (layoutBody) result.push({ body: layoutBody.body, origin: 'layout' })
    return result
  } }
}

function backgroundColor(root: XmlNode | undefined, theme: ThemeContext): string | undefined {
  const bg = getChildren(getChildren(root, 'cSld')[0], 'bg')[0]
  const bgPr = getChildren(bg, 'bgPr')[0]
  if (bgPr) {
    const definition = parseFillDefinition(bgPr)
    if (definition?.kind === 'solid') {
      const fill = resolveFill(definition, theme)
      if (fill?.kind === 'solid') return cssColor(fill.color)
    }
  }
  const bgRef = getChildren(bg, 'bgRef')[0]
  if (bgRef) {
    const idx = Number(attrs(bgRef).idx)
    const definition = Number.isInteger(idx) && idx >= 1001 ? theme.bgFillStyles[idx - 1001] : undefined
    if (definition?.kind === 'solid') {
      const placeholder = resolveDrawingColor(parseDrawingColor(bgRef), theme)
      const fill = resolveFill(definition, theme, placeholder)
      if (fill?.kind === 'solid') return cssColor(fill.color)
    }
  }
  return undefined
}

/** Nearest resolvable solid definition wins. Unsupported fills leave the prior solid available. */
async function slideBackground(pkg: OfficePackage, slidePath: string, slideRoot: XmlNode | undefined, theme: ThemeContext): Promise<string | undefined> {
  const direct = backgroundColor(slideRoot, theme)
  if (direct) return direct
  const layoutPath = await relatedPath(pkg, slidePath, 'slideLayout')
  if (!layoutPath) return undefined
  const layoutRoot = await pkg.xmlOrdered(layoutPath)
  if (!layoutRoot) return undefined
  const layout = backgroundColor(layoutRoot, theme)
  if (layout) return layout
  const masterPath = await relatedPath(pkg, layoutPath, 'slideMaster')
  if (!masterPath) return undefined
  return backgroundColor(await pkg.xmlOrdered(masterPath), theme)
}

/** Depth-first search for the first a:srgbClr under a node. */
function firstSrgb(node: XmlNode | undefined): string | undefined {
  if (!node) return undefined
  for (const [name, child] of elementChildren(node)) {
    if (name === 'srgbClr') return colorOf(child)
    const found = firstSrgb(child)
    if (found) return found
  }
  return undefined
}

interface TableStyleEntry {
  fills: { firstRow?: string; band1?: string; band2?: string; wholeTable?: string; lastRow?: string; firstCol?: string }
  firstRowTextColor?: string
  firstRowBold?: boolean
  borders?: PptxTableBorders
  firstRowBorders?: PptxTableBorders
}

/** Table scheme colors come from this slide's master theme, not theme1 by name. */
async function slideTheme(pkg: OfficePackage, slidePath: string): Promise<{ context: ThemeContext; palette: Map<string, string> }> {
  const mappings: Record<string, string> = {}
  let mappingSelected = false
  let path: string | undefined = slidePath
  let relationshipKind = 'slide'
  let themeRoot: XmlNode | undefined
  const visited = new Set<string>()
  while (path && !visited.has(path)) {
    visited.add(path)
    const root = await pkg.xmlOrdered(path)
    const override = getChildren(root, 'clrMapOvr')[0]
    if (!mappingSelected && override) {
      const customMap = getChildren(override, 'overrideClrMapping')[0]
      if (getChildren(override, 'masterClrMapping').length || customMap) {
        // The closest declaration chooses either its own map or the master's map.
        // An explicit masterClrMapping must bypass intervening layout overrides.
        mappingSelected = true
        Object.assign(mappings, attrs(customMap))
      }
    }
    for (const [key, value] of Object.entries(attrs(getChildren(root, 'clrMap')[0]))) if (!(key in mappings)) mappings[key] = value
    if (getChildren(root, 'themeElements')[0]) { themeRoot = root; break }
    const relationships: XmlNode[] = (await partRelationshipNodes(pkg, path)).filter(rel => attrs(rel).TargetMode !== 'External' && attrs(rel).Target)
    // Follow slide -> layout -> master -> theme; tolerate direct theme links on any part.
    const types = relationshipKind === 'slide' ? ['slideLayout', 'theme'] : relationshipKind === 'slideLayout' ? ['slideMaster', 'theme'] : ['theme']
    const next: { type: string; rel: XmlNode | undefined } | undefined = types.map(type => ({ type, rel: relationships.find(rel => String(attrs(rel).Type).endsWith(`/${type}`)) })).find(entry => entry.rel)
    relationshipKind = next?.type ?? ''
    path = next?.rel ? resolveTarget(path, attrs(next.rel).Target) : undefined
  }
  const context = parseThemeContext(themeRoot, mappings)
  const palette = new Map(Object.entries(context.palette).map(([key, value]) => [key, value.replace(/^#/, '')]))
  for (const [alias, key] of Object.entries(context.colorMap)) {
    const color = context.palette[key]
    if (color) palette.set(alias, color.replace(/^#/, ''))
  }
  return { context, palette }
}

function tableColor(node: XmlNode | undefined, theme: Map<string, string>): string | undefined {
  if (!node) return undefined
  const rgb = getChildren(node, 'srgbClr')[0]
  const scheme = getChildren(node, 'schemeClr')[0]
  const color = rgb ?? scheme
  const value = rgb ? attrs(rgb).val as string : theme.get(String(attrs(scheme).val))
  if (!color || !value || !/^[\da-f]{6}$/i.test(value)) return undefined
  let channels = [0, 2, 4].map(index => parseInt(value.slice(index, index + 2), 16) / 255)
  // DrawingML tints/shades operate in linear RGB. A tint's
  // percentage is the retained input color; the remainder is white.
  for (const [name, transform] of elementChildren(color)) {
    const factor = Math.max(0, Math.min(1, num(attrs(transform).val as string) / 100000))
    if (name === 'tint' || name === 'shade') {
      channels = channels.map(channel => {
        const linear = channel <= 0.04045 ? channel / 12.92 : Math.pow((channel + 0.055) / 1.055, 2.4)
        const mixed = linear * factor + (name === 'tint' ? 1 - factor : 0)
        return mixed <= 0.0031308 ? mixed * 12.92 : 1.055 * Math.pow(mixed, 1 / 2.4) - 0.055
      })
    }
  }
  return `#${channels.map(channel => Math.round(channel * 255).toString(16).padStart(2, '0')).join('').toUpperCase()}`
}

function tableBorders(tcStyle: XmlNode | undefined, theme: Map<string, string>): PptxTableBorders {
  const out: PptxTableBorders = {}
  const borders = getChildren(tcStyle, 'tcBdr')[0]
  for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'] as const) {
    const line = getChildren(getChildren(borders, side)[0], 'ln')[0]
    const color = tableColor(getChildren(line, 'solidFill')[0], theme)
    if (color) out[side] = { color, widthEmu: num(attrs(line).w as string, 12700) }
  }
  return out
}

async function readTableStyles(pkg: OfficePackage, theme: Map<string, string>): Promise<Map<string, TableStyleEntry>> {
  const out = new Map<string, TableStyleEntry>()
  const root = await pkg.xml('ppt/tableStyles.xml')
  if (!root) return out
  // the part root is a:tblStyleLst itself, but tolerate a wrapper
  const list = getChildren(root, 'tblStyleLst')[0] ?? root
  for (const style of getChildren(list, 'tblStyle')) {
    const id = attrs(style).styleId as string | undefined
    if (!id) continue
    const entry: TableStyleEntry = { fills: {} }
    // Native DrawingML styles use direct regions, with tcStyle/fill nesting.
    for (const [regionName, fillName] of [
      ['wholeTbl', 'wholeTable'], ['band1H', 'band1'], ['band2H', 'band2'],
      ['firstRow', 'firstRow'], ['lastRow', 'lastRow'], ['firstCol', 'firstCol'],
    ] as const) {
      const region = getChildren(style, regionName)[0]
      if (!region) continue
      const tcStyle = getChildren(region, 'tcStyle')[0]
      const fill = tableColor(getChildren(getChildren(tcStyle, 'fill')[0], 'solidFill')[0], theme)
      if (fill) entry.fills[fillName] = fill
      if (regionName === 'wholeTbl') entry.borders = tableBorders(tcStyle, theme)
      if (regionName === 'firstRow') {
        const tx = getChildren(region, 'tcTxStyle')[0]
        entry.firstRowTextColor = tableColor(tx, theme)
        const bold = attrs(tx).b
        if (bold !== undefined) entry.firstRowBold = bold === 'on' || bold === '1' || bold === 'true'
        entry.firstRowBorders = tableBorders(tcStyle, theme)
      }
    }
    // whole-table fill lives directly on the style
    const tableFill = getChildren(getChildren(style, 'tblPr')[0], 'solidFill')[0]
    const whole = tableFill ? colorOf(getChildren(tableFill, 'srgbClr')[0]) : undefined
    if (whole) entry.fills.wholeTable = whole
    for (const region of getChildren(style, 'tblStylePr')) {
      const kind = attrs(region).type as string | undefined
      if (!kind) continue
      const tcPr = getChildren(region, 'tcPr')[0]
      const solid = tcPr ? getChildren(tcPr, 'solidFill')[0] : undefined
      const fill = solid ? colorOf(getChildren(solid, 'srgbClr')[0]) : undefined
      switch (kind) {
        case 'firstRow':
          if (fill) entry.fills.firstRow = fill
          // a:txStyles is a sibling of a:tcPr under a:tblStylePr
          entry.firstRowTextColor = firstSrgb(getChildren(region, 'txStyles')[0])
          break
        case 'band1Horz': if (fill) entry.fills.band1 = fill; break
        case 'band2Horz': if (fill) entry.fills.band2 = fill; break
        case 'lastRow': if (fill) entry.fills.lastRow = fill; break
        case 'firstCol': if (fill) entry.fills.firstCol = fill; break
        default: break
      }
    }
    out.set(id, entry)
  }
  return out
}

/** p:graphicFrame -> a:graphic/a:graphicData/a:tbl */
function parseGraphicFrame(frame: XmlNode, tableStyles: Map<string, TableStyleEntry>, theme: ThemeContext, textDefaults?: XmlNode): PptxShape | undefined {
  const xfrm = getChildren(frame, 'xfrm')[0]
  const graphic = getChildren(frame, 'graphic')[0]
  const graphicData = graphic ? getChildren(graphic, 'graphicData')[0] : undefined
  const tbl = graphicData ? getChildren(graphicData, 'tbl')[0] : undefined
  if (!tbl) return undefined

  const textIssues: PptxDiagnostic[] = []
  const table: PptxTable = { colWidthsEmu: [], rows: [] }
  const issues: DrawingIssue[] = []
  const tblPr = getChildren(tbl, 'tblPr')[0]
  if (tblPr) {
    const ta = attrs(tblPr)
    table.firstRow = ta.firstRow === '1' || ta.firstRow === 'true'
    table.bandRow = ta.bandRow === '1' || ta.bandRow === 'true'
    const styleId = getChildren(tblPr, 'tableStyleId')[0]
      ? textOf(getChildren(tblPr, 'tableStyleId')[0]).trim() || undefined
      : undefined
    if (styleId) {
      table.styleId = styleId
      const entry = tableStyles?.get(styleId)
      if (entry) {
        table.styleFills = entry.fills
        table.firstRowTextColor = entry.firstRowTextColor
        table.firstRowBold = entry.firstRowBold
        table.styleBorders = entry.borders
        table.firstRowBorders = entry.firstRowBorders
      }
    }
  }
  const grid = getChildren(tbl, 'tblGrid')[0]
  if (grid) {
    for (const col of getChildren(grid, 'gridCol')) {
      table.colWidthsEmu.push(num(attrs(col).w as string))
    }
  }
  for (const tr of getChildren(tbl, 'tr')) {
    const ta = attrs(tr)
    const row: PptxTableRow = { cells: [], heightEmu: ta.h !== undefined ? num(ta.h as string) : undefined }
    for (const tc of getChildren(tr, 'tc')) {
      const ca = attrs(tc)
      const cell: PptxTableCell = {
        paragraphs: [],
        gridSpan: ca.gridSpan !== undefined ? num(ca.gridSpan as string) : 1,
        rowSpan: ca.rowSpan !== undefined ? num(ca.rowSpan as string) : 1,
        // hMerge/vMerge without a value marks a cell absorbed by a merge
        merged: ca.hMerge !== undefined || ca.vMerge !== undefined,
      }
      const tcPr = getChildren(tc, 'tcPr')[0]
      if (tcPr) {
        cell.drawingFill = resolveFill(parseFillDefinition(tcPr, issues), theme, undefined, issues)
        if (cell.drawingFill?.kind === 'solid') cell.fill = cssColor(cell.drawingFill.color)
        for (const [name, side] of [['lnL', 'left'], ['lnR', 'right'], ['lnT', 'top'], ['lnB', 'bottom']] as const) {
          const borderNode = getChildren(tcPr, name)[0]
          if (!borderNode) continue
          const border = resolveDrawingStyle({ ln: borderNode }, undefined, theme)
          issues.push(...border.issues)
          if (border.line) {
            // Omitted width remains inheritable from the selected table style.
            if (attrs(borderNode).w === undefined) delete border.line.width
            cell.drawingBorders ??= {}
            cell.drawingBorders[side] = border.line
          }
        }
        const ca = attrs(tcPr)
        for (const [name, key] of [['marL', 'leftEmu'], ['marR', 'rightEmu'], ['marT', 'topEmu'], ['marB', 'bottomEmu']] as const) {
          if (ca[name] === undefined) continue
          const value = ca[name].trim() === '' ? NaN : Number(ca[name])
          if (Number.isFinite(value) && value >= 0) { cell.margins ??= {}; cell.margins[key] = value }
          else issues.push({ kind: 'invalid-paint', message: `Invalid table margin ${name}`, feature: name })
        }
        if (['t', 'ctr', 'b'].includes(ca.anchor)) cell.anchor = ca.anchor as PptxTextBody['anchor']
      }
      const txBody = getChildren(tc, 'txBody')[0]
      if (txBody) {
        const body = parseTextBody(txBody, theme, textDefaults)
        cell.paragraphs = body.paragraphs
        textIssues.push(...(body.diagnostics ?? []))
      }
      row.cells.push(cell)
    }
    table.rows.push(row)
  }

  return {
    ...parseTransform(xfrm),
    geometry: 'rect',
    table,
    diagnostics: [...issues, ...textIssues],
  }
}

interface Representation {
  node: XmlNode
  representation: 'choice' | 'fallback'
  reason?: string
  feature?: string
}

const MAX_DRAWING_DEPTH = 128
const SUPPORTED_DRAWING_REQUIREMENTS = ['a', 'p', 'c', 'dgm', 'dsp', 'ink']

/** Per-slide decisions avoid rescanning a selected subtree at every enclosing Choice. */
class DrawingCompatibility {
  constructor(private contents: Map<XmlNode, DrawingContent<PptxParagraph, PptxTextBody>>) {}
  private features = new WeakMap<XmlNode, string | null>()
  private selections = new WeakMap<XmlNode, Representation | undefined>()
  private failures = new WeakMap<XmlNode, string>()

  failure(node: XmlNode): string { return this.failures.get(node) ?? 'AlternateContent' }
  useRepresentation(node: XmlNode, selected: Representation): void { this.selections.set(node, selected); this.features = new WeakMap() }

  /** Only the selected branch can veto an enclosing native Choice. */
  unsupportedFeature(node: XmlNode, depth = 0): string | undefined {
    if (depth >= MAX_DRAWING_DEPTH) return 'drawing-depth'
    if (this.features.has(node)) return this.features.get(node) ?? undefined
    const feature = this.inspect(node, depth)
    this.features.set(node, feature ?? null)
    return feature
  }

  private inspect(node: XmlNode, depth: number): string | undefined {
    for (const [name, child] of orderedChildren(node)) {
      if (name === '#text') continue
      if (name === 'AlternateContent') {
        const selected = this.selectRepresentation(child, depth + 1)
        if (!selected) return this.failure(child)
        const feature = this.unsupportedFeature(selected.node, depth + 2)
        if (feature) return feature
        continue
      }
      if (name === 'm' || name === 'oMath' || name === 'oMathPara') return 'OMML'
      if (name.toLowerCase() === 'model3d') return 'model3D'
      if (name === 'graphicData') {
        const uri = attrs(child).uri?.toLowerCase() ?? ''
        if (uri.includes('model3d')) return 'model3D'
        if (uri.includes('chartex')) return 'ChartEx'
        if (['chart', 'relIds', 'contentPart', 'wsp'].some(tag => getChildren(child, tag).length) && !this.contents.has(child)) return 'drawing-content'
      }
      if (name === 'contentPart') { if (!this.contents.has(child) && !this.contents.has(node)) return 'drawing-content'; continue }
      const feature = this.unsupportedFeature(child, depth + 1)
      if (feature) return feature
    }
    return undefined
  }

  selectRepresentation(node: XmlNode, depth = 0): Representation | undefined {
    if (depth >= MAX_DRAWING_DEPTH) { this.failures.set(node, 'drawing-depth'); return undefined }
    if (this.selections.has(node)) return this.selections.get(node)
    let feature: string | undefined
    for (const choice of getChildren(node, 'Choice')) {
      const unsupported = this.unsupportedFeature(choice, depth + 1)
      // Requires is resolved through the Choice's in-scope namespace bindings.
      const unknownRequirement = supportedChoiceRequirements(choice, SUPPORTED_DRAWING_REQUIREMENTS).unknown
      if (!unsupported && !unknownRequirement) {
        const selected: Representation = { node: choice, representation: 'choice' }
        this.selections.set(node, selected)
        return selected
      }
      feature ??= unsupported ?? unknownRequirement
    }
    this.failures.set(node, feature ?? 'AlternateContent')
    const fallback = getChildren(node, 'Fallback')[0]
    const selected: Representation | undefined = fallback ? { node: fallback, representation: 'fallback', feature: feature ?? 'AlternateContent', reason: `Native fallback selected for unsupported ${feature ?? 'AlternateContent'}` } : undefined
    this.selections.set(node, selected)
    return selected
  }
}

function sourceOf(node: XmlNode, element: string, partPath: string, treePath: string, representation?: Omit<Representation, 'node'>): PptxSource {
  const nonvisual = getChildren(node, 'nvSpPr')[0] ?? getChildren(node, 'nvCxnSpPr')[0] ?? getChildren(node, 'nvPicPr')[0] ?? getChildren(node, 'nvGraphicFramePr')[0] ?? getChildren(node, 'nvGrpSpPr')[0]
  const a = attrs(getChildren(nonvisual, 'cNvPr')[0])
  return { partPath, treePath, element, ...(a.id !== undefined ? { id: a.id } : {}), ...(a.name !== undefined ? { name: a.name } : {}), representation: 'native', ...representation }
}

const xmlEscape = (value: string): string => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/** Reparse only when nested compatibility nodes need selection, preserving shared XML ordering. */
function selectNestedAlternates(node: XmlNode, source: PptxSource, diagnostics: PptxDiagnostic[], compatibility: DrawingCompatibility): XmlNode {
  const pending = [node]
  let containsAlternate = false
  while (pending.length && !containsAlternate) {
    for (const [name, child] of orderedChildren(pending.pop())) {
      if (name === 'AlternateContent') { containsAlternate = true; break }
      if (name !== '#text') pending.push(child)
    }
  }
  if (!containsAlternate) return node
  const serialize = (parent: XmlNode, depth = 0): string => {
    if (depth >= MAX_DRAWING_DEPTH) {
      diagnostics.push({ kind: 'missing-representation', message: 'Drawing nesting exceeds the static renderer limit', feature: 'drawing-depth', source })
      return ''
    }
    return orderedChildren(parent).map(([name, child]) => {
      if (name === '#text') return xmlEscape(textOf(child))
      if (name === 'AlternateContent') {
        const selected = compatibility.selectRepresentation(child, depth + 1)
        if (!selected) {
          diagnostics.push({ kind: 'missing-representation', message: 'AlternateContent has no supported Choice or native fallback', feature: compatibility.failure(child), source })
          return ''
        }
        if (selected.representation === 'fallback') diagnostics.push({ kind: 'fallback-representation', message: selected.reason!, feature: selected.feature, source })
        return serialize(selected.node, depth + 2)
      }
      const attributes = Object.entries(attrs(child)).map(([key, value]) => ` ${key}="${xmlEscape(value)}"`).join('')
      return `<${name}${attributes}>${serialize(child, depth + 1)}</${name}>`
    }).join('')
  }
  return parseXmlOrdered(`<root>${serialize(node)}</root>`)
}

/** Find the selected payload on the original XML tree; reparsed shapes have different node identities. */
function selectedGraphicData(frame: XmlNode, compatibility: DrawingCompatibility): XmlNode | undefined {
  const find = (parent: XmlNode | undefined, depth: number): XmlNode | undefined => {
    if (!parent || depth >= MAX_DRAWING_DEPTH) return undefined
    for (const [name, node] of orderedChildren(parent)) {
      if (name === 'graphicData') return node
      if (name === '#text') continue
      const branch = name === 'AlternateContent' ? compatibility.selectRepresentation(node, depth + 1)?.node : node
      const found = find(branch, depth + 1)
      if (found) return found
    }
    return undefined
  }
  return find(frame, 0)
}

/** Retain ordinary source text without painting the unselected mathematical representation. */
function retainChoiceText(choice: XmlNode | undefined, selectedShapes: PptxShape[]): void {
  const bodies = new Map<string, PptxTextBody>()
  const pending = [{ name: 'Choice', parent: choice, depth: 0 }]
  while (pending.length) {
    const { name, parent, depth } = pending.pop()!
    if (depth >= MAX_DRAWING_DEPTH) continue
    if (name === 'sp') {
      const id = attrs(getChildren(getChildren(parent, 'nvSpPr')[0], 'cNvPr')[0]).id
      const text = getChildren(parent, 'txBody')[0]
      if (id && text && !bodies.has(id)) bodies.set(id, parseTextBody(text))
    }
    // Process nodes on pop, pushing later siblings first to retain source-order DFS.
    const children = orderedChildren(parent)
    for (let i = children.length - 1; i >= 0; i--) {
      const [childName, child] = children[i]
      if (childName !== '#text') pending.push({ name: childName, parent: child, depth: depth + 1 })
    }
  }
  walkShapes(selectedShapes, shape => {
    if (shape.source?.id) shape.sourceTextBody = bodies.get(shape.source.id)
  })
}

function parseShapeTree(
  pkg: OfficePackage, contents: Map<XmlNode, DrawingContent<PptxParagraph, PptxTextBody>>, parent: XmlNode | undefined, partPath: string, slideImages: Map<string, PptxImageRef>, theme: ThemeContext,
  tableStyles: Map<string, TableStyleEntry>, diagnostics: PptxDiagnostic[], compatibility: DrawingCompatibility, treePath = 'spTree',
  representation?: Omit<Representation, 'node'>, depth = 0, textDefaults?: XmlNode, groupDepth = 0, inheritance?: SlideTextInheritance,
): PptxShape[] {
  const shapes: PptxShape[] = []
  if (depth >= MAX_DRAWING_DEPTH) {
    diagnostics.push({ kind: 'missing-representation', message: 'Drawing nesting exceeds the static renderer limit', feature: 'drawing-depth', source: sourceOf(parent ?? {}, 'spTree', partPath, treePath, representation) })
    return shapes
  }
  for (const [index, [name, original]] of orderedChildren(parent).filter(([name]) => name !== '#text').entries()) {
    const path = `${treePath}/${name}[${index}]`
    if (name === 'AlternateContent') {
      const selected = compatibility.selectRepresentation(original)
      if (selected) {
        const { node, ...selection } = selected
        const selectedShapes = parseShapeTree(pkg, contents, node, partPath, slideImages, theme, tableStyles, diagnostics, compatibility, `${path}/${selected.representation}`, selection, depth + 1, textDefaults, groupDepth, inheritance)
        if (selected.representation === 'fallback') retainChoiceText(getChildren(original, 'Choice')[0], selectedShapes)
        if (!selectedShapes.length) diagnostics.push({ kind: 'missing-representation', message: 'Selected AlternateContent representation contains no usable drawing', feature: selected.feature, source: sourceOf(original, name, partPath, path, selection) })
        shapes.push(...selectedShapes)
      } else diagnostics.push({ kind: 'missing-representation', message: 'AlternateContent has no supported Choice or native fallback', feature: compatibility.failure(original), source: sourceOf(original, name, partPath, path) })
      continue
    }
    if (!['sp', 'cxnSp', 'pic', 'graphicFrame', 'grpSp', 'contentPart'].includes(name)) continue
    if (name === 'grpSp' && groupDepth >= DRAWING_GROUP_DEPTH) {
      const source = sourceOf(original, name, partPath, path, representation)
      contentDiagnostic(drawingPartContext(pkg), 'group-depth', partPath, undefined, { identity: source.id ?? path, reason: 'group-depth', limit: DRAWING_GROUP_DEPTH })
      continue
    }
    if (!reserveDrawingNode(drawingPartContext(pkg), partPath)) break
    const source = sourceOf(original, name, partPath, path, representation)
    const shapeIssues: PptxDiagnostic[] = []
    if (representation?.representation === 'fallback') shapeIssues.push({ kind: 'fallback-representation', message: representation.reason!, feature: representation.feature, source })
    let shape: PptxShape | undefined
    if (name === 'grpSp') {
      const xfrm = getChildren(getChildren(original, 'grpSpPr')[0], 'xfrm')[0]
      const transform = parseTransform(xfrm)
      const ca = attrs(getChildren(xfrm, 'chOff')[0]), ce = attrs(getChildren(xfrm, 'chExt')[0])
      const validChildren = [ca.x, ca.y, ce.cx, ce.cy].every(value => value !== undefined && value.trim() !== '' && Number.isFinite(Number(value))) && num(ce.cx) > 0 && num(ce.cy) > 0
      shape = { ...transform, geometry: 'other', source, transformValid: transform.transformValid && validChildren,
        group: { off: { x: transform.xEmu, y: transform.yEmu }, ext: { width: transform.widthEmu, height: transform.heightEmu }, chOff: { x: num(ca.x), y: num(ca.y) }, chExt: { width: num(ce.cx), height: num(ce.cy) } },
        children: parseShapeTree(pkg, contents, original, partPath, slideImages, theme, tableStyles, diagnostics, compatibility, path, representation, depth + 1, textDefaults, groupDepth + 1, inheritance), diagnostics: [],
      }
    } else {
      const node = selectNestedAlternates(original, source, shapeIssues, compatibility)
      const graphic = name === 'graphicFrame' ? selectedGraphicData(original, compatibility) : undefined
      const content = name === 'contentPart' ? contents.get(original) : graphic ? contents.get(graphic) : undefined
      if (content && (name === 'graphicFrame' || name === 'contentPart')) {
        reserveDrawingContent(pkg, content, partPath, false)
        shape = { ...parseTransform(getChildren(node, 'xfrm')[0]), geometry: 'other', content }
      } else shape = name === 'graphicFrame' ? parseGraphicFrame(node, tableStyles, theme, textDefaults) : parseShape(node, slideImages, theme, source, textDefaults, inheritance)
    }
    if (!shape) {
      diagnostics.push({ kind: 'unsupported-object', message: `No static renderer for ${name}`, feature: compatibility.unsupportedFeature(original) ?? name, source })
      continue
    }
    shape.source = source
    shape.diagnostics = [...shapeIssues, ...(shape.diagnostics ?? [])].map(issue => ({ ...issue, source: issue.source ?? source }))
    if (shape.transformValid === false) shape.diagnostics.push({ kind: 'invalid-transform', message: 'Malformed shape/group transform; painting skipped', source })
    shapes.push(shape)
  }
  return shapes
}

function walkShapes(shapes: PptxShape[], visit: (shape: PptxShape) => void): void {
  for (const shape of shapes) { visit(shape); if (shape.children) walkShapes(shape.children, visit) }
}

export async function parsePptx(pkg: OfficePackage): Promise<PptxDocument> {
  const presentation = await pkg.xml('ppt/presentation.xml')
  if (!presentation) throw new Error('ppt/presentation.xml missing — not a valid pptx?')
  const sldSz = getChildren(presentation, 'sldSz')[0]
  const sa = attrs(sldSz)
  const doc: PptxDocument = {
    drawingCoverage: [],
    slideWidthEmu: num(sa.cx as string, 9144000),
    slideHeightEmu: num(sa.cy as string, 6858000),
    slides: [],
    images: [],
  }
  // slide order from presentation rels
  const rels = await pkg.xml('ppt/_rels/presentation.xml.rels')
  Object.assign(doc, await parseEmbeddedFonts(pkg, presentation, rels))
  const relMap = new Map<string, string>()
  if (rels) {
    for (const rel of getChildren(rels, 'Relationship')) {
      const a = attrs(rel)
      if (a.Id && a.Target && a.TargetMode !== 'External') relMap.set(a.Id, a.Target)
    }
  }
  const sldIdLst = getChildren(presentation, 'sldIdLst')[0]
  const slideIds = sldIdLst ? getChildren(sldIdLst, 'sldId') : []
  for (let i = 0; i < slideIds.length; i++) {
    const a = attrs(slideIds[i])
    const rid = (a['r:id'] ?? a.id) as string
    const target = relMap.get(rid)
    if (!target) continue
    const path = resolveTarget('ppt/presentation.xml', target)
    const slideRoot = await pkg.xmlOrdered(path)
    const slideImages = await loadSlideImages(pkg, path)
    const preloadedRelationships = drawingPartContext(pkg).diagnostics.filter(issue =>
      issue.kind === 'malformed-part' && issue.reason === 'invalid-relationship-xml' && issue.identity === path && !issue.sourceReferenceId)
    const theme = await slideTheme(pkg, path)
    const tableStyles = await readTableStyles(pkg, theme.palette)
    const slide: PptxSlide = { index: i, widthEmu: doc.slideWidthEmu, heightEmu: doc.slideHeightEmu, shapes: [], theme: theme.context, diagnostics: [] }
    slide.background = await slideBackground(pkg, path, slideRoot, theme.context)
    const spTree = getChildren(getChildren(slideRoot, 'cSld')[0], 'spTree')[0]
    const contents = new Map<XmlNode, DrawingContent<PptxParagraph, PptxTextBody>>()
    const compatibility = new DrawingCompatibility(contents)
    const defaults = getChildren(presentation, 'defaultTextStyle')[0]
    const contentTheme = { colors: theme.palette, fonts: new Map([['minorHAnsi', theme.context.fonts.minor.latin ?? 'Calibri']]) }
    const pending = [slideRoot]
    const diagnosticStart = drawingPartContext(pkg).diagnostics.length
    while (pending.length) {
      const parent = pending.pop()
      for (const [name, node] of orderedChildren(parent)) {
        if (name === '#text') continue
        if (name === 'graphicData' || name === 'contentPart') {
          const adapters = {
            parseDiagramText: (body: XmlNode, shape: XmlNode) => parseTextBody(body, theme.context, defaults, textFontDefaults(getChildren(shape, 'style')[0], theme.context)),
            parseParagraph: (p: XmlNode) => parseTextBody({ p }, theme.context, defaults).paragraphs[0],
          }
          let payload = await prepareDrawingContent(pkg, name === 'contentPart' ? { contentPart: node } : node, path, contentTheme, theme.context, adapters)
          if (!payload && name === 'graphicData' && getChildren(node, 'AlternateContent').length) {
            const selected = await prepareCompatibleDrawingContent(pkg, node, path, contentTheme, theme.context, adapters)
            if (selected) {
              payload = selected.content
              if (payload && selected.reference) contents.set(selected.reference, payload)
              for (const choice of selected.selections) compatibility.useRepresentation(choice.alternate, { node: choice.node, representation: choice.representation, feature: choice.feature, reason: choice.reason })
            }
          }
          if (payload) contents.set(node, payload)
        }
        if (name !== 'graphicData') pending.push(node)
      }
    }
    const inheritance = await slideTextInheritance(pkg, path)
    slide.shapes = parseShapeTree(pkg, contents, spTree, path, slideImages, theme.context, tableStyles, slide.diagnostics!, compatibility, 'spTree', undefined, 0, defaults, 0, inheritance)
    slide.diagnostics!.push(...preloadedRelationships, ...drawingPartContext(pkg).diagnostics.slice(diagnosticStart))
    if (getChildren(slideRoot, 'timing').length) slide.diagnostics!.push({ kind: 'deferred-animation', message: 'Slide animation/timing is outside static drawing rendering', feature: 'timing' })
    // fill in placeholder geometry from the layout/master
    let needsGeometry = false
    walkShapes(slide.shapes, shape => { if (shape.placeholder && (shape.widthEmu === 0 || shape.heightEmu === 0)) needsGeometry = true })
    if (needsGeometry) {
      const inherited = await layoutPlaceholderGeometry(pkg, path)
      walkShapes(slide.shapes, shape => {
        if (!shape.placeholder || (shape.widthEmu !== 0 && shape.heightEmu !== 0)) return
        const geom = inherited.get(`${shape.placeholder.type}|${shape.placeholder.idx}`)
        if (geom) {
          shape.xEmu = geom.x
          shape.yEmu = geom.y
          shape.widthEmu = geom.w
          shape.heightEmu = geom.h
        }
      })
    }
    walkShapes(slide.shapes, shape => {
      const issues = shape.drawingGeometry ? resolveGeometry(shape.drawingGeometry, emuToPx(shape.widthEmu), emuToPx(shape.heightEmu)).issues : []
      shape.diagnostics = [...(shape.diagnostics ?? []), ...issues.map(issue => ({ ...issue, source: shape.source })), ...(shape.drawingStyle?.issues ?? []).map(issue => ({ ...issue, source: shape.source }))]
      slide.diagnostics!.push(...shape.diagnostics)
    })
    const selectedBranch = (alternate: XmlNode) => compatibility.selectRepresentation(alternate)?.node
    if (preloadedRelationships.length) {
      // Inventory selected source pictures before attributing the cached .rels
      // error; a discarded Choice must not become a painted source object.
      const inventory = pptxCoverage(slide, spTree, path, selectedBranch)
      const pictureRefs = new Set(inventory.filter(entry => (entry.element === 'pic' || entry.feature === 'image') && entry.referenceId).map(entry => entry.referenceId!))
      const selectedRefs = new Set(inventory.filter(entry => entry.referenceId).map(entry => entry.referenceId!))
      for (const referenceId of pictureRefs) malformedRelationshipAttempt(pkg, path, referenceId, 'image')
      // A slide part may occur in multiple presentation units. Cache-level
      // failures are reused, but each unit needs its own selected-source audit.
      for (const issue of drawingPartContext(pkg).diagnostics) if (issue.reason === 'invalid-relationship-xml' &&
        issue.ownerPartPath === path && issue.sourceReferenceId && selectedRefs.has(issue.sourceReferenceId) &&
        !slide.diagnostics!.includes(issue)) slide.diagnostics!.push(issue)
      slide.diagnostics = slide.diagnostics!.filter(issue => !attemptedMalformedRelationshipIssue(pkg, issue as ContentDiagnostic))
    }
    doc.drawingCoverage!.push(...pptxCoverage(slide, spTree, path, selectedBranch))
    doc.slides.push(slide)
  }
  // Recursive first-use source order, deduplicated by media part while preserving use-specific crop.
  const imageIndices = new Map<string | PptxImageRef, number>()
  for (const slide of doc.slides) walkShapes(slide.shapes, shape => {
    if (!shape.image) return
    const key = shape.image.partPath ?? shape.image
    let index = imageIndices.get(key)
    if (index === undefined) {
      index = doc.images.length
      const { opacity: _opacity, srcRect: _srcRect, fillRect: _fillRect, ...part } = shape.image
      doc.images.push(part)
      imageIndices.set(key, index)
    }
    shape.imageIndex = index
  })
  return doc
}
