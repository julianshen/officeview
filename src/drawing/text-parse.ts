/** DrawingML text metadata shared by slides, cached diagrams, and sheet drawings. */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { parseDrawingColor, resolveDrawingColor, type DrawingColor, type ThemeContext } from './style'
import type { DrawingTextParagraph as PptxParagraph, DrawingTextBody as PptxTextBody, DrawingTextRun as PptxTextRun, DrawingTextSpacing as PptxTextSpacing, DrawingTextStyle as PptxTextStyle, DrawingTabStop as PptxTabStop, TextDirection } from './text'

const number = (v: string | undefined, fallback = 0): number => v !== undefined && Number.isFinite(Number(v)) ? Number(v) : fallback
function align(v: string | undefined): PptxParagraph['align'] {
  return v === 'ctr' ? 'center' : v === 'r' ? 'right' : v === 'just' ? 'justify' : 'left'
}
function textCssColor(color: DrawingColor): string {
  const rgb = [color.r, color.g, color.b]
  return color.a < 1 ? `rgba(${rgb.join(',')},${color.a})` : `#${rgb.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()}`
}
function style(node: XmlNode | undefined, theme?: ThemeContext): PptxTextStyle {
  const a = attrs(node), out: PptxTextStyle = {}
  if (a.b !== undefined) out.bold = a.b === '1' || a.b === 'true'
  if (a.i !== undefined) out.italic = a.i === '1' || a.i === 'true'
  if (a.sz !== undefined) out.fontSizePt = number(a.sz) / 100
  if (a.spc !== undefined) out.characterSpacingPt = number(a.spc) / 100
  if (a.lang !== undefined) out.language = a.lang
  for (const [tag, field] of [['latin', 'fontFamily'], ['ea', 'fontFamilyEastAsia'], ['cs', 'fontFamilyComplexScript']] as const) {
    const face = attrs(getChildren(node, tag)[0]).typeface
    if (face !== undefined) out[field] = face
  }
  if (out.fontFamily === undefined && a.typeface !== undefined) out.fontFamily = a.typeface
  if (getChildren(node, 'noFill').length) out.noFill = true
  else if (getChildren(node, 'solidFill').length) out.noFill = false
  const color = resolveDrawingColor(parseDrawingColor(getChildren(node, 'solidFill')[0]), theme)
  if (color) out.color = textCssColor(color)
  return out
}
function spacing(node: XmlNode | undefined): PptxTextSpacing | undefined {
  const points = attrs(getChildren(node, 'spcPts')[0]).val
  const percent = attrs(getChildren(node, 'spcPct')[0]).val
  return points !== undefined ? { kind: 'points', value: number(points) / 100 }
    : percent !== undefined ? { kind: 'percent', value: number(percent) / 100000 } : undefined
}
function paragraphProperties(node: XmlNode | undefined): Partial<PptxParagraph> {
  if (!node) return {}
  const a = attrs(node), out: Partial<PptxParagraph> = {}
  if (a.algn !== undefined) { out.align = align(a.algn); out.sourceAlign = a.algn }
  for (const [name, key] of [['marL', 'marginLeftEmu'], ['marR', 'marginRightEmu'], ['indent', 'indentEmu'], ['defTabSz', 'defaultTabSizeEmu']] as const) {
    if (a[name] !== undefined) out[key] = number(a[name])
  }
  for (const [name, key] of [['lnSpc', 'lineSpacing'], ['spcBef', 'spaceBefore'], ['spcAft', 'spaceAfter']] as const) {
    const s = spacing(getChildren(node, name)[0]); if (s) out[key] = s
  }
  const tabList = getChildren(node, 'tabLst')[0]
  if (tabList) out.tabStops = getChildren(tabList, 'tab').map((tab): PptxTabStop => {
    const a = attrs(tab)
    return { positionEmu: number(a.pos), align: a.algn === 'ctr' ? 'center' : a.algn === 'r' ? 'right' : a.algn === 'dec' ? 'decimal' : 'left' }
  }).sort((a, b) => a.positionEmu - b.positionEmu)
  const bullet = getChildren(node, 'buChar')[0]
  if (bullet) { out.bullet = true; out.bulletCharacter = attrs(bullet).char ?? '•' }
  if (getChildren(node, 'buNone').length) { out.bullet = false; out.bulletCharacter = undefined }
  return out
}

export interface InheritedTextLayer {
  /** Master text style or another level-based DrawingML style. */
  style?: XmlNode
  /** Placeholder txBody contributes bodyPr, list defaults and paragraph metadata, never its content. */
  body?: XmlNode
  origin: 'master' | 'layout' | 'placeholder'
}

const directions = new Set<TextDirection>(['horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl'])

export interface ParsedDrawingTextBody extends PptxTextBody {
  diagnostics?: Array<{ kind: 'unsupported-text-alignment'; feature: string; message: string }>
}

export function parseTextBody(txBody: XmlNode, theme?: ThemeContext, defaults?: XmlNode, fontDefaults: PptxTextStyle = {}, inheritedLayers: readonly InheritedTextLayer[] = []): ParsedDrawingTextBody {
  const a = Object.assign({}, ...inheritedLayers.map(layer => attrs(getChildren(layer.body, 'bodyPr')[0])), attrs(getChildren(txBody, 'bodyPr')[0]))
  const body: ParsedDrawingTextBody = { paragraphs: [], anchor: a.anchor === 'ctr' ? 'ctr' : a.anchor === 'b' ? 'b' : 't',
    insetLeftEmu: number(a.lIns, 91440), insetRightEmu: number(a.rIns, 91440), insetTopEmu: number(a.tIns, 45720), insetBottomEmu: number(a.bIns, 45720), wrap: a.wrap !== 'none',
    ...(directions.has(a.vert as TextDirection) ? { direction: a.vert as TextDirection } : {}) }
  const list = getChildren(txBody, 'lstStyle')[0]
  for (const p of getChildren(txBody, 'p')) {
    const pPr = getChildren(p, 'pPr')[0], level = number(attrs(pPr).lvl)
    const listNodes = (node: XmlNode | undefined) => [getChildren(node, 'defPPr')[0], getChildren(node, `lvl${level + 1}pPr`)[0]]
    const propertyLayers: Array<[XmlNode | undefined, 'default' | 'master' | 'layout' | 'placeholder' | 'list' | 'paragraph']> =
      listNodes(defaults).map(node => [node, 'default'])
    for (const layer of inheritedLayers) {
      for (const node of listNodes(layer.style)) propertyLayers.push([node, layer.origin])
      for (const node of listNodes(getChildren(layer.body, 'lstStyle')[0])) propertyLayers.push([node, layer.origin])
      // A placeholder paragraph can supply pPr even when its text is intentionally empty.
      const placeholderParagraph = getChildren(layer.body, 'p').find(candidate => number(attrs(getChildren(candidate, 'pPr')[0]).lvl) === level)
      if (placeholderParagraph) propertyLayers.push([getChildren(placeholderParagraph, 'pPr')[0], layer.origin])
    }
    propertyLayers.push(...listNodes(list).map(node => [node, 'list'] as [XmlNode | undefined, 'list']), [pPr, 'paragraph'])
    let inherited = { ...fontDefaults }
    const propertySources: NonNullable<PptxTextRun['propertySources']> = {}
    for (const key of Object.keys(fontDefaults) as Array<keyof PptxTextStyle>) propertySources[key] = 'default'
    const para: PptxParagraph = { runs: [], align: 'left', level }
    for (const [node, origin] of propertyLayers) {
      Object.assign(para, paragraphProperties(node))
      const values = style(getChildren(node, 'defRPr')[0], theme)
      inherited = { ...inherited, ...values }
      for (const key of Object.keys(values) as Array<keyof PptxTextStyle>) propertySources[key] = origin
    }
    if (para.sourceAlign && !['l', 'ctr', 'r', 'just'].includes(para.sourceAlign)) {
      body.diagnostics ??= []
      body.diagnostics.push({ kind: 'unsupported-text-alignment', feature: para.sourceAlign, message: `DrawingML alignment ${para.sourceAlign} is deferred; using left alignment` })
    }
    para.defaultProperties = inherited
    para.endProperties = style(getChildren(p, 'endParaRPr')[0], theme)
    for (const [name, node] of orderedChildren(p)) {
      if (name !== 'r' && name !== 'br' && name !== 'fld') continue
      const directProperties = style(getChildren(node, 'rPr')[0], theme)
      const end = name === 'br' ? para.endProperties : {}
      const origins = { ...propertySources }
      for (const key of Object.keys(end) as Array<keyof PptxTextStyle>) origins[key] = 'end'
      for (const key of Object.keys(directProperties) as Array<keyof PptxTextStyle>) origins[key] = 'run'
      para.runs.push({ ...inherited, ...end, ...directProperties, text: name === 'br' ? '\n' : textOf(getChildren(node, 't')[0]), directProperties, propertySources: origins })
    }
    body.paragraphs.push(para)
  }
  return body
}

/** Shape fontRef supplies theme face/color defaults; it never overrides explicit rPr. */
export function textFontDefaults(shapeStyle: XmlNode | undefined, theme?: ThemeContext): PptxTextStyle {
  const ref = getChildren(shapeStyle, 'fontRef')[0], index = attrs(ref).idx
  const out: PptxTextStyle = {}
  if (index === 'major' || index === 'minor') {
    const prefix = index === 'major' ? '+mj' : '+mn'
    Object.assign(out, { fontFamily: `${prefix}-lt`, fontFamilyEastAsia: `${prefix}-ea`, fontFamilyComplexScript: `${prefix}-cs` })
  }
  const color = resolveDrawingColor(parseDrawingColor(ref), theme)
  if (color) out.color = textCssColor(color)
  return out
}
