/** Theme-resolved payloads; source parts are cached separately from these models. */
import { attrs, getChildren, parseXmlOrdered, textOf, type XmlNode } from '../core/xml'
import { emuToPx } from '../core/geometry'
import type { GeometryDefinition, GeometryIssue } from './geometry'
import type { DrawingColor, DrawingStyle, ThemeContext } from './style'
import type { SceneGroupTransform } from './scene'
import type { OfficePackage } from '../core/zip'
import { parseGeometry, resolveGeometry } from './geometry'
import { supportedChoiceRequirements } from './coverage'
import { parseGroupShapeProperties, resolveDrawingStyle } from './style'
import { parseTextBody, textFontDefaults, parseAutofitNode } from './text-parse'
import { isVmlTrue } from './vml'
import { resolveTextFamily } from './text-layout'
import { CONTENT_REFERENCE_DEPTH, DRAWING_GROUP_DEPTH, DOCUMENT_DRAWING_NODE_LIMIT, contentDescendants as descendants, contentDiagnostic, drawingPartContext, partRelationships, referencedPart, reserveDrawingNode, type ContentDiagnostic } from './parts'
import { orderedChildren } from '../core/xml'
import type { TextDirection, TextWarp, TextWarpPreset, DrawingTextAutofit } from './text'
import { SUPPORTED_TEXT_WARP_PRESETS, DEFAULT_WARP_ADJUSTMENTS, parseAdjustGuides, parseTextPlainAdjustment } from './text'
export interface ContentTextRun { text: string; fontFamily?: string; fontSizePt?: number; color?: string; bold?: boolean; italic?: boolean }
export type ContentParagraphAlign = 'left' | 'center' | 'right' | 'justify'
export interface ContentParagraph { runs: ContentTextRun[]; align: ContentParagraphAlign }
export interface ContentTheme { colors: Map<string, string>; fonts: Map<string, string> }
/** Minimal raster asset identity for nested paint lookup (structural: any
 * image model carrying bytes qualifies, keeping shared code adapter-free). */
export interface ContentImageAsset {
  data: Uint8Array
  mime?: string
}
/** Optional decoded-asset handoff for nested picture paint, added compatibly. */
export interface ContentPaintAssets {
  imageFor?: (image: ContentImageAsset) => CanvasImageSource | undefined
  /** Optional explicit fallback chain/region, threaded into adapter resolvers. */
  fallbackFonts?: import('../core/fonts/fallback').FallbackFontsOptions
  /** Optional embedded font resolver for mapping families to registered aliases. */
  resolveFont?: import('../core/fonts/register').FontResolver
}
const num = (v: string | undefined, fallback = 0): number => v !== undefined && Number.isFinite(Number(v)) ? Number(v) : fallback
const child = (n: XmlNode | undefined, name: string) => getChildren(n, name)[0]
/** Diagnostic-only memory survives a discarded Choice when its relationship XML was cached. */
const malformedRelationshipParts = new WeakMap<OfficePackage, Map<string, string>>()
/** Attribute a cached malformed owner .rels file to an actual source reference. */
export function malformedRelationshipAttempt(
  pkg: OfficePackage, owner: string, referenceId: string, feature?: string,
  diagnosticStart?: number, origin: { ownerPartPath: string; sourceReferenceId: string } = { ownerPartPath: owner, sourceReferenceId: referenceId },
  reuseFreshRaw = false,
): boolean {
  const context = drawingPartContext(pkg)
  const malformed = context.diagnostics.find(issue => issue.kind === 'malformed-part' && issue.reason === 'invalid-relationship-xml' && issue.identity === owner)
  const relsPath = malformed?.partPath ?? malformedRelationshipParts.get(pkg)?.get(owner)
  if (!relsPath) return false
  let known = malformedRelationshipParts.get(pkg)
  if (!known) { known = new Map(); malformedRelationshipParts.set(pkg, known) }
  known.set(owner, relsPath)
  if (diagnosticStart !== undefined) {
    for (let index = context.diagnostics.length - 1; index >= diagnosticStart; index--) {
      const issue = context.diagnostics[index]
      if (issue.kind === 'missing-part' && issue.partPath === owner && issue.identity === referenceId && issue.reason === 'relationship-not-found') context.diagnostics.splice(index, 1)
    }
  }
  // follow() annotates diagnostics created inside its own attempt in finally.
  // A picture path has no such annotation, so it needs an explicit source issue.
  if (reuseFreshRaw && malformed && diagnosticStart !== undefined && context.diagnostics.indexOf(malformed) >= diagnosticStart) return true
  if (context.diagnostics.some(issue => issue.kind === 'malformed-part' && issue.reason === 'invalid-relationship-xml' &&
    issue.partPath === relsPath && issue.identity === owner && issue.ownerPartPath === origin.ownerPartPath &&
    issue.sourceReferenceId === origin.sourceReferenceId && issue.feature === feature)) return true
  contentDiagnostic(context, 'malformed-part', relsPath, feature, {
    identity: owner, ownerPartPath: origin.ownerPartPath, sourceReferenceId: origin.sourceReferenceId, reason: 'invalid-relationship-xml',
  })
  return true
}
/** The cache-level .rels diagnostic has a real source attempt, even if its Choice was discarded. */
export function attemptedMalformedRelationshipIssue(pkg: OfficePackage, issue: ContentDiagnostic): boolean {
  return issue.kind === 'malformed-part' && issue.reason === 'invalid-relationship-xml' && !issue.sourceReferenceId && !!issue.identity &&
    malformedRelationshipParts.get(pkg)?.get(issue.identity) === issue.partPath
}
export interface DrawingContentShape<Text = never> {
  xEmu: number
  yEmu: number
  widthEmu: number
  heightEmu: number
  geometry: string
  image?: ContentImageAsset
  /** Cached DrawingML outline; absent on handwritten legacy drawing models. */
  drawingGeometry?: GeometryDefinition
  drawingStyle?: DrawingStyle
  geometryIssues?: GeometryIssue[]
  textBody?: Text
  group?: SceneGroupTransform
  children?: DrawingContentShape<Text>[]
  flipH?: boolean
  flipV?: boolean
  rotationDeg?: number
  fill?: string
  line?: { color: string; widthEmu: number }
  paragraphs: Array<{ runs: ContentTextRun[]; align: ContentParagraphAlign }>
  fontFamily: string
  textColor?: string
}

export type DrawingContent<Paragraph = ContentParagraph, Text = never> =
  | { kind: 'diagram'; shapes: DrawingContentShape<Text>[]; textOnly?: boolean }
  | {
      kind: 'ink'
      strokes: Array<{ points: Array<[number, number]>; widthEmu: number; color: string }>
      widthEmu: number
      heightEmu: number
    }
  | {
      kind: 'chart'
      title?: string
      categories: string[]
      series: Array<{ name: string; values: Array<number | undefined>; color: string }>
      min?: number
      max?: number
      majorUnit?: number
      gapWidth: number
      overlap: number
      legend: boolean
      fontFamily: string
      fontSizePt: number
    }
  | {
      kind: 'textbox'
      paragraphs: Paragraph[]
      /** DrawingML bodyPr text direction. Absent means horizontal. */
      direction?: TextDirection
      /** Legacy Word textbox flag; explicit direction takes precedence. */
      vertical?: boolean
      fontFamily: string
      fontSizePt: number
      insets: { left: number; top: number; right: number; bottom: number }
      anchor?: 't' | 'ctr' | 'b'
      wrap?: boolean
      fill?: string
      line?: { color: string; widthEmu: number }
      textWarp?: TextWarp
      autofit?: DrawingTextAutofit
      diagnostics?: Array<{ kind: 'unsupported-text-alignment' | 'unsupported-text-appearance' | 'unsupported-text-warp'; feature: string; message: string }>
    }

function color(fill: XmlNode | undefined, theme: ContentTheme): string | undefined {
  const c = child(fill, 'srgbClr') ?? child(fill, 'schemeClr') ?? child(fill, 'sysClr') ?? child(fill, 'prstClr')
  if (!c) return undefined
  const a = attrs(c)
  const preset: Record<string, string> = {
    black: '000000',
    white: 'FFFFFF',
    red: 'FF0000',
    blue: '0000FF',
    green: '008000'
  }
  const value = child(fill, 'schemeClr')
    ? theme.colors.get(a.val)
    : child(fill, 'prstClr')
      ? preset[a.val]
      : child(fill, 'sysClr')
        ? a.lastClr
        : a.val
  if (!value || !/^[\da-f]{6}$/i.test(value)) return undefined
  let channels = [0, 2, 4].map((i) => parseInt(value.slice(i, i + 2), 16) / 255)
  for (const [name, transform] of orderedChildren(c)) {
    const amount = Math.min(1, Math.max(0, num(attrs(transform).val) / 100000))
    if (name === 'tint' || name === 'shade')
      channels = channels.map((channel) => {
        const linear = channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4
        const adjusted = name === 'tint' ? linear * amount + 1 - amount : linear * amount
        return adjusted <= 0.0031308 ? adjusted * 12.92 : 1.055 * adjusted ** (1 / 2.4) - 0.055
      })
    else if (name === 'lumMod') channels = channels.map((channel) => channel * amount)
    else if (name === 'lumOff') channels = channels.map((channel) => Math.min(1, channel + amount))
  }
  return channels
    .map((c) =>
      Math.round(c * 255)
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase()
}

function hexDrawingColor(color: DrawingColor): string {
  const hex = (n: number) => Math.round(n * 255).toString(16).padStart(2, '0').toUpperCase()
  return `${hex(color.r)}${hex(color.g)}${hex(color.b)}`
}

function shape(node: XmlNode, theme: ContentTheme, drawingTheme: ThemeContext, groupStyle?: Partial<DrawingStyle>, name?: string): DrawingContentShape {
  const pr = child(node, 'spPr'),
    transform = child(pr, 'xfrm')
  const offset = attrs(child(transform, 'off')),
    extent = attrs(child(transform, 'ext'))
  const ln = child(pr, 'ln'),
    lineColor = color(child(ln, 'solidFill'), theme)
  const txAttrs = attrs(transform)
  const textColor = color(child(child(node, 'style'), 'fontRef'), theme)
  const fontFamily = theme.fonts.get('minorHAnsi') ?? 'Calibri'
  const drawingGeometry = parseGeometry(pr)
  const drawingStyle = resolveDrawingStyle(pr, child(node, 'style'), drawingTheme, name === 'cxnSp' ? { ...groupStyle, fill: { kind: 'none' } } : groupStyle)
  const geometryIssues = resolveGeometry(drawingGeometry, emuToPx(num(extent.cx)), emuToPx(num(extent.cy))).issues
  const textBody = child(node, 'txBody')
  const parsedText = textBody ? parseTextBody(textBody, drawingTheme, undefined, textFontDefaults(child(node, 'style'), drawingTheme)) : undefined
  const paragraphs = parsedText?.paragraphs.map((p) => ({
    align: p.align,
    runs: p.runs.map((run) => ({
      text: run.text,
      fontFamily: theme.fonts.get(run.fontFamily ?? '') ?? resolveTextFamily(run, drawingTheme) ?? fontFamily,
      fontSizePt: run.fontSizePt ?? 24,
      color: run.color?.replace(/^#/, '') ?? textColor,
      bold: run.bold,
      italic: run.italic,
    }))
  })) ?? []
  return {
    xEmu: num(offset.x),
    yEmu: num(offset.y),
    widthEmu: num(extent.cx),
    heightEmu: num(extent.cy),
    geometry: drawingGeometry.preset ?? 'custom',
    drawingGeometry,
    drawingStyle,
    geometryIssues,
    rotationDeg: num(txAttrs.rot) / 60000,
    flipH: isVmlTrue(txAttrs.flipH),
    flipV: isVmlTrue(txAttrs.flipV),
    fill: color(child(pr, 'solidFill'), theme) ?? (drawingStyle.fill?.kind === 'solid' ? hexDrawingColor(drawingStyle.fill.color) : undefined),
    line: lineColor
      ? { color: lineColor, widthEmu: num(attrs(ln).w, 12700) }
      : (drawingStyle.line?.fill?.kind === 'solid'
          ? { color: hexDrawingColor(drawingStyle.line.fill.color), widthEmu: (drawingStyle.line.width ?? 1) * 9525 }
          : undefined),
    textBody: parsedText as unknown as any,
    paragraphs,
    fontFamily,
    textColor
  }
}

/** Cache points are indexed; missing points must retain their category slot. */
function cacheValues(node: XmlNode | undefined): Array<string | undefined> {
  const cache =
    child(child(node, 'strRef'), 'strCache') ??
    child(child(node, 'numRef'), 'numCache') ??
    child(node, 'strLit') ??
    child(node, 'numLit')
  const out: Array<string | undefined> = []
  for (const p of getChildren(cache, 'pt')) out[num(attrs(p).idx)] = textOf(child(p, 'v'))
  return Array.from({ length: out.length }, (_, index) => out[index])
}

function chart(root: XmlNode, theme: ContentTheme): Extract<DrawingContent, { kind: 'chart' }> | undefined {
  const chart = child(root, 'chart'),
    plot = child(chart, 'plotArea'),
    bars = child(plot, 'barChart')
  if (
    !bars ||
    attrs(child(bars, 'barDir')).val === 'bar' ||
    !['clustered', undefined].includes(attrs(child(bars, 'grouping')).val)
  )
    return undefined
  const series = getChildren(bars, 'ser').map((s, index) => {
    const tx = child(s, 'tx')
    return {
      name: textOf(child(tx, 'v')) || cacheValues(tx)[0] || `Series ${index + 1}`,
      values: cacheValues(child(s, 'val')).map((v) =>
        v === undefined || !Number.isFinite(Number(v)) ? undefined : Number(v)
      ),
      color:
        color(child(child(s, 'spPr'), 'solidFill'), theme) ?? theme.colors.get(`accent${(index % 6) + 1}`) ?? '4F81BD'
    }
  })
  const categories = cacheValues(child(getChildren(bars, 'ser')[0], 'cat')).map((v) => v ?? '')
  const title = child(chart, 'title')
  const titleText = descendants(child(title, 'tx'), 't').map(textOf).join('') || cacheValues(child(title, 'tx'))[0]
  const locale = attrs(child(root, 'lang')).val ?? ''
  const axis = child(plot, 'valAx'),
    scaling = child(axis, 'scaling')
  const optional = (n: XmlNode | undefined) => (n ? num(attrs(n).val) : undefined)
  return {
    kind: 'chart',
    categories,
    series,
    title: title ? titleText || (locale.startsWith('zh') ? '圖表標題' : 'Chart Title') : undefined,
    min: optional(child(scaling, 'min')),
    max: optional(child(scaling, 'max')),
    majorUnit: optional(child(axis, 'majorUnit')),
    gapWidth: Math.max(0, num(attrs(child(bars, 'gapWidth')).val, 150)),
    overlap: Math.max(-100, Math.min(100, num(attrs(child(bars, 'overlap')).val))),
    legend: !!child(chart, 'legend'),
    fontFamily: theme.fonts.get('minorHAnsi') ?? 'Calibri',
    fontSizePt: num(attrs(descendants(child(axis, 'txPr'), 'defRPr')[0]).sz, 900) / 100
  }
}

const LENGTH_EMU: Record<string, number> = { cm: 360000, mm: 36000, in: 914400, pt: 12700, px: 9525 }

/** Cartesian InkML traces. Unsupported encodings keep the compatibility picture. */
export function parseInk(root: XmlNode): Extract<DrawingContent, { kind: 'ink' }> | undefined {
  const contexts = descendants(root, 'context')
  const brushes = descendants(root, 'brush')
  const strokes: Extract<DrawingContent, { kind: 'ink' }>['strokes'] = []
  for (const trace of descendants(root, 'trace')) {
    if (attrs(trace).type === 'penUp') continue
    const context = contexts.find((node) => `#${attrs(node).id}` === attrs(trace).contextRef)
    const source = getChildren(context, 'inkSource')[0]
    const format = getChildren(source, 'traceFormat')[0] ?? getChildren(root, 'traceFormat')[0]
    const channels = getChildren(format, 'channel')
    if (channels.length < 2 || attrs(channels[0]).name !== 'X' || attrs(channels[1]).name !== 'Y') return undefined
    const scale = channels.slice(0, 2).map((channel) => {
      const a = attrs(channel)
      const resolution = descendants(source, 'channelProperty').find(
        (property) => attrs(property).channel === a.name && attrs(property).name === 'resolution'
      )
      const divisor = Number(attrs(resolution).value ?? 1)
      return (LENGTH_EMU[a.units ?? 'mm'] ?? 0) / divisor
    })
    if (scale.some((value) => !Number.isFinite(value) || value <= 0)) return undefined
    const values = channels.map(() => 0)
    const differences = channels.map(() => 0)
    const modes = channels.map(() => '!')
    const points: Array<[number, number]> = []
    // InkML prefixes persist per channel: ! absolute, ' first difference,
    // " second difference. See https://www.w3.org/TR/InkML/#trace.
    for (const sample of textOf(trace).split(',')) {
      const tokens = [...sample.matchAll(/([!'\"]?)([+-]?(?:\d*\.\d+|\d+)(?:[eE][+-]?\d+)?)/g)]
      if (tokens.length !== channels.length) return undefined
      for (let index = 0; index < tokens.length; index++) {
        const [, prefix, raw] = tokens[index]
        if (prefix) modes[index] = prefix
        const value = Number(raw)
        if (modes[index] === '!') {
          differences[index] = value - values[index]
          values[index] = value
        } else if (modes[index] === "'") {
          differences[index] = value
          values[index] += value
        } else {
          differences[index] += value
          values[index] += differences[index]
        }
      }
      if (values.some(value => !Number.isFinite(value))) return undefined
      points.push([values[0] * scale[0], values[1] * scale[1]])
    }
    const brush = brushes.find((node) => `#${attrs(node).id}` === attrs(trace).brushRef)
    const properties = getChildren(brush, 'brushProperty')
    const width = attrs(properties.find((node) => attrs(node).name === 'width'))
    const widthEmu = Number(width.value ?? 1) * (LENGTH_EMU[width.units ?? 'mm'] ?? 0)
    if (!Number.isFinite(widthEmu) || widthEmu <= 0) return undefined
    const color = attrs(properties.find((node) => attrs(node).name === 'color')).value?.replace(/^#/, '')
    strokes.push({ points, widthEmu, color: color && /^[\da-f]{6}$/i.test(color) ? color : '000000' })
  }
  if (!strokes.length) return undefined
  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity
  for (const stroke of strokes)
    for (const [x, y] of stroke.points) {
      minX = Math.min(minX, x)
      minY = Math.min(minY, y)
      maxX = Math.max(maxX, x)
      maxY = Math.max(maxY, y)
    }
  for (const stroke of strokes) stroke.points = stroke.points.map(([x, y]) => [x - minX, y - minY])
  return { kind: 'ink', strokes, widthEmu: maxX - minX, heightEmu: maxY - minY }
}


export interface ContentLoadAdapters<Paragraph, Text> {
  parseParagraph?: (p: XmlNode) => Paragraph
  parseDiagramText?: (node: XmlNode, shape: XmlNode) => Text
}
const deferredDiagrams = new WeakMap<object, { materializeSelected: () => boolean }>()
const SOURCE_PROBE_LIMIT = DOCUMENT_DRAWING_NODE_LIMIT * 2
const cachedGroupIdentity = (group: XmlNode): string | undefined => attrs(child(child(group, 'nvGrpSpPr'), 'cNvPr')).id || attrs(group).id
/** Raw source capability check: no styles, text models, or cached shapes built. */
function inspectCachedTree(tree: XmlNode): { supported: boolean; sawPicture: boolean; pictureIds: string[]; invalidGroup: boolean; limitReached: boolean; overDepthGroup?: XmlNode } {
  const result: { supported: boolean; sawPicture: boolean; pictureIds: string[]; invalidGroup: boolean; limitReached: boolean; overDepthGroup?: XmlNode } = { supported: false, sawPicture: false, pictureIds: [], invalidGroup: false, limitReached: false }
  const pending: Array<{ name: string; node: XmlNode; groupDepth: number }> = [{ name: 'spTree', node: tree, groupDepth: 0 }]
  let inspected = 0
  while (pending.length && inspected < SOURCE_PROBE_LIMIT) {
    const { name, node, groupDepth } = pending.pop()!
    if (name !== 'spTree') {
      inspected++
      if (name === 'sp' || name === 'cxnSp') { result.supported = true; continue }
      if (name === 'graphicFrame') {
        const tbl = child(child(child(node, 'graphic'), 'graphicData'), 'tbl')
        if (tbl) result.supported = true
        continue
      }
      if (name === 'pic') {
        result.sawPicture = true
        const blip = child(child(node, 'blipFill'), 'blip')
        const picId = attrs(blip).embed ?? attrs(blip).link
        if (picId) result.pictureIds.push(picId)
        continue
      }
      if (groupDepth >= DRAWING_GROUP_DEPTH) {
        result.overDepthGroup ??= node
        continue
      }
      if (name === 'grpSp') {
        const transform = child(child(node, 'grpSpPr'), 'xfrm')
        const off = attrs(child(transform, 'off')), ext = attrs(child(transform, 'ext'))
        const chOff = attrs(child(transform, 'chOff')), chExt = attrs(child(transform, 'chExt'))
        const valid = [off.x, off.y, ext.cx, ext.cy, chOff.x, chOff.y, chExt.cx, chExt.cy].every(v => v !== undefined && v.trim() !== '' && Number.isFinite(Number(v))) && num(chExt.cx) > 0 && num(chExt.cy) > 0
        if (!valid) { result.invalidGroup = true; continue }
      }
    }
    const children = orderedChildren(node)
    for (let i = children.length - 1; i >= 0; i--) {
      const [childName, sourceNode] = children[i]
      if (['sp', 'cxnSp', 'grpSp', 'pic', 'graphicFrame'].includes(childName)) pending.push({ name: childName, node: sourceNode, groupDepth: name === 'grpSp' ? groupDepth + 1 : groupDepth })
    }
  }
  result.limitReached = pending.length > 0
  return result
}
/** Bound dataModel point collection so a hostile part cannot stage unbounded paragraphs. */
const SMARTART_TEXT_PARAGRAPH_LIMIT = 1000
/**
 * Collect SmartArt dataModel point texts in document order, one entry per
 * point with non-blank dgm:t content. Empty points contribute nothing.
 * Returns the texts plus whether points were dropped at the limit.
 */
/**
 * Paragraph texts for one dataModel point. Plain string points yield a
 * single paragraph. Nested DrawingML (non-conformant but observed when
 * pretty-printed) yields one paragraph per nested p, with br line breaks
 * preserved as newline characters; pretty whitespace between elements is
 * skipped while intentional spacing inside runs is preserved.
 */
function pointParagraphs(pt: XmlNode): string[] {
  const tops = getChildren(pt, 't')
  if (!tops.some((t) => orderedChildren(t).some(([name]) => name !== '#text'))) {
    const text = tops.map((t) => textOf(t)).join('').trim()
    return text ? [text] : []
  }
  const paras: string[] = []
  let cur = ''
  const flush = (): void => { if (cur !== '') { paras.push(cur); cur = '' } }
  const walk = (node: XmlNode): void => {
    for (const [name, child] of orderedChildren(node)) {
      if (name === '#text') {
        const s = textOf(child)
        if (s.trim() !== '') cur += s
      } else if (name === 'p') {
        flush()
        walk(child)
        flush()
      } else if (name === 'br') {
        cur += '\n'
      } else if (name === 't' && !orderedChildren(child).some(([n]) => n !== '#text')) {
        cur += textOf(child)
      } else {
        walk(child)
      }
    }
  }
  for (const t of tops) walk(t)
  flush()
  return paras
}
function collectDataModelTexts(root: XmlNode | undefined): { texts: string[]; truncated: boolean } {
  const dataModel = root && child(root, 'ptLst') ? root : child(root, 'dataModel')
  const pts = dataModel ? getChildren(child(dataModel, 'ptLst'), 'pt') : []
  const texts: string[] = []
  let truncated = false
  for (const pt of pts) {
    for (const para of pointParagraphs(pt)) {
      const text = para.trim()
      if (!text) continue
      if (texts.length >= SMARTART_TEXT_PARAGRAPH_LIMIT) { truncated = true; break }
      texts.push(text)
    }
    if (truncated) break
  }
  return { texts, truncated }
}
/**
 * Build a minimal DrawingML text body carrying the fallback paragraphs, so
 * each caller parses them through its own parseDiagramText pipeline
 * (DOCX and PPTX adapters both consume a:p runs; the node is produced by
 * the real XML parser, never hand-shaped). Plain &/<> escaping only;
 * unparseable text degrades at the call site, never here.
 */
function smartArtTextBodyXml(texts: string[]): string {
  const esc = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  const paras = texts.map((text) => `<a:p><a:r><a:t>${esc(text)}</a:t></a:r></a:p>`).join('')
  return `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/>${paras}</a:txBody>`
}
/** Resolve a single supported cached payload. Never cache theme-resolved models. */
export async function prepareDrawingContent<Paragraph = ContentParagraph, Text = never>(
  pkg: OfficePackage, graphic: XmlNode, owner: string, theme: ContentTheme, drawingTheme: ThemeContext,
  adapters: ContentLoadAdapters<Paragraph, Text> = {}, ancestry: readonly string[] = [owner], depth = 0,
  source?: { ownerPartPath: string; sourceReferenceId: string; interpretation: string },
): Promise<DrawingContent<Paragraph, Text> | undefined> {
  const context = drawingPartContext(pkg)
  const follow = async (id: string, from = owner, interpretation = 'content'): Promise<DrawingContent<Paragraph, Text> | undefined> => {
    const origin = source ?? { ownerPartPath: from, sourceReferenceId: id, interpretation }
    const diagnosticStart = context.diagnostics.length
    try {
    const path = await referencedPart(pkg, from, id)
    if (!path) {
      malformedRelationshipAttempt(pkg, from, id, undefined, diagnosticStart, origin, true)
      return undefined
    }
    if (ancestry.includes(path)) { contentDiagnostic(context, 'content-cycle', from, path, { identity: path, reason: 'ancestry' }); return undefined }
    if (depth >= CONTENT_REFERENCE_DEPTH) { contentDiagnostic(context, 'content-depth', from, path, { identity: path, reason: 'reference-depth', limit: CONTENT_REFERENCE_DEPTH }); return undefined }
    let root: XmlNode | undefined
    try { root = await pkg.xmlOrdered(path) }
    catch { contentDiagnostic(context, 'malformed-part', path, interpretation, { identity: id, ownerPartPath: from, reason: 'invalid-content-xml' }); return undefined }
    if (!root) return undefined
    if (context.placements.has(path)) context.coverage.repeated++
    context.placements.add(path)
    context.coverage.loaded++
    const next = [...ancestry, path]
    const payload = parseInk(root) ?? chart(root, theme)
    if (payload) return payload
    const tree = child(root, 'spTree')
    if (tree) {
      const capability = inspectCachedTree(tree)
      if (!capability.supported) {
        let hasValidPicture = false
        if (capability.sawPicture) {
          const rels = await partRelationships(pkg, path)
          for (const picId of capability.pictureIds) {
            const rel = rels.get(picId)
            if (rel?.external) {
              contentDiagnostic(context, 'external-reference', path, picId, { identity: picId, reason: 'external-relationship' })
            } else if (!rel?.path || !pkg.has(rel.path)) {
              contentDiagnostic(context, 'missing-part', path, picId, { identity: picId, reason: rel?.path ? 'part-not-found' : 'relationship-not-found' })
            } else {
              hasValidPicture = true
            }
          }
          if (!hasValidPicture) {
            contentDiagnostic(context, 'unsupported-content', path, 'pic', { reason: 'cached-picture-unsupported' })
          }
        }
        if (!hasValidPicture) {
          if (capability.invalidGroup) contentDiagnostic(context, 'unsupported-content', path, 'invalid-group-transform')
          if (capability.overDepthGroup) contentDiagnostic(context, 'group-depth', path, undefined, { identity: cachedGroupIdentity(capability.overDepthGroup), reason: 'group-depth', limit: DRAWING_GROUP_DEPTH })
          if (capability.limitReached) contentDiagnostic(context, 'unsupported-content', path, 'source-probe-limit', { reason: 'source-probe-limit', limit: SOURCE_PROBE_LIMIT })
          contentDiagnostic(context, 'unsupported-content', path, 'empty-diagram', { reason: 'no-supported-shapes' })
          return undefined
        }
      }
      const pictureRels = capability.sawPicture ? await partRelationships(pkg, path) : new Map()
      const pictureBytes = new Map<string, Uint8Array>()
      if (capability.sawPicture) {
        // Preload only picture ids actually referenced by the cached tree,
        // not every relationship of the part (which may point at large
        // non-image payloads).
        for (const picId of capability.pictureIds) {
          if (pictureBytes.has(picId)) continue
          const rel = pictureRels.get(picId)
          if (!rel || rel.external || !rel.path || !pkg.has(rel.path)) continue
          try {
            const bytes = await pkg.bytes(rel.path)
            if (bytes) pictureBytes.set(picId, bytes)
          } catch {}
        }
      }
      let shapes: DrawingContentShape<Text>[] | undefined
      const content: Extract<DrawingContent<Paragraph, Text>, { kind: 'diagram' }> = { kind: 'diagram', shapes: [] }
      Object.defineProperty(content, 'shapes', {
        enumerable: true,
        get: () => shapes ?? (shapes = diagramShapes(tree, path, pictureRels, pictureBytes, false)),
        set: (value: DrawingContentShape<Text>[]) => { shapes = value },
      })
      deferredDiagrams.set(content, { materializeSelected: () => {
        shapes = diagramShapes(tree, path, pictureRels, pictureBytes, true)
        return true
      } })
      return content
    }
    const cached = descendants(root, 'dataModelExt')[0]
    if (cached) {
      const cachedId = attrs(cached).relId
      // Existing Word producers may attach the cached drawing rel to either owner.
      const local = (await partRelationships(pkg, path)).get(cachedId)
      const cachedOwner = local ? path : from
      const cachedContent = await prepareDrawingContent(pkg, { contentPart: { '@attrs': { id: cachedId } } }, cachedOwner, theme, drawingTheme, adapters, next, depth + 1, origin)
      if (cachedContent) return cachedContent
      // A broken cache keeps its diagnostics and falls through to the
      // dataModel text fallback below instead of losing all text.
    }
    // A chart document contains c:chart with plotArea; it is not a new
    // relationship carrier. Unsupported chart types end at this target.
    const chartDocument = child(root, 'chart')
    if (chartDocument && !attrs(chartDocument).id) {
      contentDiagnostic(context, 'unsupported-content', path, 'chart', { reason: 'unsupported-chart' })
      return undefined
    }
    // Cacheless (or cache-broken) SmartArt: the dataModel carries author
    // text but no usable cached dsp:drawing. Synthesize a searchable
    // text-only box from dgm:pt/dgm:t instead of losing all text. Cached
    // rendering stays preferred; this only runs when nothing usable was
    // found above. Generic paragraphs carry the text adapter-free while
    // textBody flows through the caller's own parseDiagramText pipeline,
    // so DOCX and PPTX adapters each parse format-correct paragraphs; the
    // unsupported-content diagnostic lets the existing coverage mapping
    // report the long-standing text-only representation status.
    const { texts: smartArtTexts, truncated: smartArtTruncated } = collectDataModelTexts(root)
    if (smartArtTexts.length > 0) {
      const paragraphs = smartArtTexts.map((text) => ({ runs: [{ text }], align: 'left' as const }))
      let textBody: Text | undefined
      if (adapters.parseDiagramText) {
        try {
          textBody = adapters.parseDiagramText(parseXmlOrdered(smartArtTextBodyXml(smartArtTexts)), {})
        } catch {
          textBody = undefined
        }
      }
      contentDiagnostic(context, 'unsupported-content', path, 'smartart', { reason: 'cacheless-smartart-text-fallback' })
      if (smartArtTruncated) {
        contentDiagnostic(context, 'unsupported-content', path, 'smartart', { reason: 'smartart-text-truncated', limit: SMARTART_TEXT_PARAGRAPH_LIMIT })
      }
      const fallbackShape: DrawingContentShape<Text> = {
        xEmu: 0,
        yEmu: 0,
        widthEmu: 0,
        heightEmu: 0,
        geometry: 'rect',
        ...(textBody !== undefined ? { textBody } : {}),
        paragraphs,
        fontFamily: theme.fonts.get('minorHAnsi') ?? 'Calibri',
      }
      return { kind: 'diagram', shapes: [fallbackShape], textOnly: true }
    }
    const data = child(root, 'graphicData') ?? child(child(root, 'graphic'), 'graphicData') ?? root
    const beforeNested = context.diagnostics.length
    const loaded = await prepareDrawingContent(pkg, data, path, theme, drawingTheme, adapters, next, depth + 1, origin)
    if (!loaded && context.diagnostics.length === beforeNested) contentDiagnostic(context, 'unsupported-content', path, origin.interpretation, { reason: 'unsupported-content' })
    return loaded
    } finally {
      // Preserve the immediate failure identity and resolved target; carry
      // the selected source relationship through nested reference chains.
      for (const issue of context.diagnostics.slice(diagnosticStart)) {
        issue.ownerPartPath = origin.ownerPartPath
        issue.sourceReferenceId ??= origin.sourceReferenceId
        if (!issue.feature && ['unsupported-content', 'content-cycle', 'content-depth'].includes(issue.kind)) issue.feature = origin.interpretation
      }
    }
  }
  const diagramShapes = (
    tree: XmlNode, part: string,
    pictureRels: Map<string, { path?: string; external: boolean; type?: string }>,
    pictureBytes: Map<string, Uint8Array>,
    selected: boolean,
    groupDepth = 0,
    state = { nodes: 0, exhausted: false, pictureReported: false, missingPictures: new Set<string>() },
    groupStyle?: Partial<DrawingStyle>,
  ): DrawingContentShape<Text>[] => {
    const out: DrawingContentShape<Text>[] = []
    for (const [name, node] of orderedChildren(tree)) {
      if (state.exhausted) break
      if (!['sp', 'cxnSp', 'grpSp', 'pic', 'graphicFrame'].includes(name)) continue
      if (selected) {
        if (!reserveDrawingNode(context, part)) { state.exhausted = true; break }
      } else if (++state.nodes > DOCUMENT_DRAWING_NODE_LIMIT) {
        contentDiagnostic(context, 'node-budget', part, undefined, { reason: 'source-node-limit', limit: DOCUMENT_DRAWING_NODE_LIMIT })
        state.exhausted = true
        break
      }
      if (name === 'pic') {
        const image = child(child(node, 'blipFill'), 'blip')
        const id = attrs(image).embed ?? attrs(image).link
        let imgData: Uint8Array | undefined
        if (id) {
          const rel = pictureRels.get(id)
          // Dedupe repeated ids and bound the tracking set so a hostile tree
          // with many distinct dangling references cannot grow it without limit.
          if (rel?.external) {
            if (!state.missingPictures.has(id) && state.missingPictures.size < 64) {
              contentDiagnostic(context, 'external-reference', part, id, { identity: id, reason: 'external-relationship' })
              state.missingPictures.add(id)
            }
          } else if (!rel?.path || !pkg.has(rel.path)) {
            if (!state.missingPictures.has(id) && state.missingPictures.size < 64) {
              contentDiagnostic(context, 'missing-part', part, id, { identity: id, reason: rel?.path ? 'part-not-found' : 'relationship-not-found' })
              state.missingPictures.add(id)
            }
          } else {
            imgData = pictureBytes.get(id)
          }
        }
        if (!imgData) {
          if (!state.pictureReported) {
            contentDiagnostic(context, 'unsupported-content', part, 'pic', { reason: 'cached-picture-unsupported' })
            state.pictureReported = true
          }
          continue
        }
        const pr = child(node, 'spPr')
        const transform = child(pr, 'xfrm')
        const txAttrs = attrs(transform)
        const offset = attrs(child(transform, 'off'))
        const extent = attrs(child(transform, 'ext'))
        const drawingGeometry = parseGeometry(pr)
        out.push({
          xEmu: num(offset.x),
          yEmu: num(offset.y),
          widthEmu: num(extent.cx),
          heightEmu: num(extent.cy),
          geometry: drawingGeometry.preset ?? 'rect',
          rotationDeg: num(txAttrs.rot) / 60000,
          flipH: isVmlTrue(txAttrs.flipH),
          flipV: isVmlTrue(txAttrs.flipV),
          image: { data: imgData },
          paragraphs: [],
          fontFamily: '',
        })
        continue
      }
      if (name === 'graphicFrame') {
        const transform = child(node, 'xfrm')
        const frameOffset = attrs(child(transform, 'off'))
        const frameExtent = attrs(child(transform, 'ext'))
        const frameX = num(frameOffset.x)
        const frameY = num(frameOffset.y)
        const frameW = num(frameExtent.cx)
        const frameH = num(frameExtent.cy)

        const graphic = child(node, 'graphic')
        const graphicData = child(graphic, 'graphicData')
        const tbl = child(graphicData, 'tbl')
        if (tbl) {
          const tblGrid = child(tbl, 'tblGrid')
          const cols = getChildren(tblGrid, 'gridCol').map(c => num(attrs(c).w))
          const trs = getChildren(tbl, 'tr')
          const rowHeights = trs.map(tr => num(attrs(tr).h))
          let currentY = frameY
          for (let r = 0; r < trs.length; r++) {
            if (state.exhausted) break
            const tr = trs[r]
            const defaultRowH = (rowHeights[r] && rowHeights[r] > 0) ? rowHeights[r] : (trs.length > 0 ? frameH / trs.length : 0)
            let currentX = frameX
            let colIdx = 0
            const tcs = getChildren(tr, 'tc')
            // Standard DrawingML rows retain merged-away physical cells;
            // compact producer rows omit those placeholders.
            const physicalColumns = tcs.length === cols.length
            for (const [cellIndex, tc] of tcs.entries()) {
              if (selected) {
                if (!reserveDrawingNode(context, part)) { state.exhausted = true; break }
              } else if (++state.nodes > DOCUMENT_DRAWING_NODE_LIMIT) {
                contentDiagnostic(context, 'node-budget', part, undefined, { reason: 'source-node-limit', limit: DOCUMENT_DRAWING_NODE_LIMIT })
                state.exhausted = true
                break
              }
              if (isVmlTrue(attrs(tc).hMerge) || isVmlTrue(attrs(tc).vMerge)) {
                currentX += cols[colIdx] ?? (tcs.length > 0 ? frameW / tcs.length : 0)
                colIdx++
                continue
              }
              const boundedSpan = (feature: 'gridSpan' | 'rowSpan', remaining: number): number => {
                const raw = attrs(tc)[feature]
                const span = raw === undefined ? 1 : Number(raw)
                if (!Number.isInteger(span) || span < 1 || span > remaining) {
                  contentDiagnostic(context, 'malformed-part', part, feature, {
                    identity: raw, reason: 'invalid-table-span', limit: remaining,
                  })
                  // Preserve malformed cell text as one slot, without
                  // expanding beyond the actual table grid or row list.
                  return Math.min(1, remaining)
                }
                return span
              }
              let gridSpan = boundedSpan('gridSpan', Math.max(0, (cols.length || tcs.length) - colIdx))
              if (physicalColumns && gridSpan > 1 && !tcs.slice(cellIndex + 1, cellIndex + gridSpan).every(covered => isVmlTrue(attrs(covered).hMerge))) {
                contentDiagnostic(context, 'malformed-part', part, 'gridSpan', { reason: 'missing-merge-placeholder' })
                // Contradictory physical rows retain each cell in its own
                // column rather than painting a span over later text.
                gridSpan = 1
              }
              const rowSpan = boundedSpan('rowSpan', trs.length - r)
              if (gridSpan === 0) continue
              let cellW = 0
              for (let s = 0; s < gridSpan; s++) {
                cellW += (cols.length > 0 ? cols[colIdx + s] : undefined) ?? (tcs.length > 0 ? frameW / tcs.length : 0)
              }
              let cellH = 0
              for (let s = 0; s < rowSpan; s++) {
                const rIdx = r + s
                cellH += (rowHeights[rIdx] && rowHeights[rIdx] > 0) ? rowHeights[rIdx] : defaultRowH
              }
              const tcPr = child(tc, 'tcPr')
              const cellFill = color(child(tcPr, 'solidFill'), theme)
              const ln = child(tcPr, 'lnL') ?? child(tcPr, 'lnT') ?? child(tcPr, 'lnB') ?? child(tcPr, 'lnR') ?? child(tcPr, 'ln')
              const cellLineColor = ln ? color(child(ln, 'solidFill'), theme) : undefined
              const txBody = child(tc, 'txBody')
              const parsedText = txBody ? parseTextBody(txBody, drawingTheme, undefined, textFontDefaults(undefined, drawingTheme)) : undefined
              const fontFamily = theme.fonts.get('minorHAnsi') ?? 'Calibri'
              const paragraphs = parsedText?.paragraphs.map((p) => ({
                align: p.align,
                runs: p.runs.map((run) => ({
                  text: run.text,
                  fontFamily: theme.fonts.get(run.fontFamily ?? '') ?? resolveTextFamily(run, drawingTheme) ?? fontFamily,
                  fontSizePt: run.fontSizePt ?? 12,
                  color: run.color?.replace(/^#/, ''),
                  bold: run.bold,
                  italic: run.italic,
                }))
              })) ?? []
              const cellShape: DrawingContentShape<Text> = {
                xEmu: currentX,
                yEmu: currentY,
                widthEmu: cellW,
                heightEmu: cellH,
                geometry: 'rect',
                fill: cellFill,
                drawingStyle: tcPr ? resolveDrawingStyle(tcPr, undefined, drawingTheme) : undefined,
                line: cellLineColor ? { color: cellLineColor, widthEmu: num(attrs(ln).w, 12700) } : undefined,
                paragraphs,
                fontFamily,
                textBody: parsedText as unknown as any,
              }
              if (txBody && adapters.parseDiagramText) cellShape.textBody = adapters.parseDiagramText(txBody, tc)
              out.push(cellShape)
              currentX += physicalColumns ? cols[colIdx] : cellW
              colIdx += physicalColumns ? 1 : gridSpan
            }
            currentY += defaultRowH
          }
          continue
        }
        contentDiagnostic(context, 'unsupported-content', part, attrs(graphicData).uri || 'graphicFrame')
        continue
      }
      if (name === 'grpSp') {
        if (groupDepth >= DRAWING_GROUP_DEPTH) { contentDiagnostic(context, 'group-depth', part, undefined, { identity: cachedGroupIdentity(node), reason: 'group-depth', limit: DRAWING_GROUP_DEPTH }); continue }
        const grpSpPr = child(node, 'grpSpPr')
        const currentGroupStyle = parseGroupShapeProperties(grpSpPr, drawingTheme, groupStyle)
        const transform = child(grpSpPr, 'xfrm'), a = attrs(transform)
        const off = attrs(child(transform, 'off')), ext = attrs(child(transform, 'ext'))
        const chOff = attrs(child(transform, 'chOff')), chExt = attrs(child(transform, 'chExt'))
        if (![off.x, off.y, ext.cx, ext.cy, chOff.x, chOff.y, chExt.cx, chExt.cy].every(v => v !== undefined && v.trim() !== '' && Number.isFinite(Number(v))) || num(chExt.cx) <= 0 || num(chExt.cy) <= 0) {
          contentDiagnostic(context, 'unsupported-content', part, 'invalid-group-transform'); continue
        }
        out.push({ xEmu: num(off.x), yEmu: num(off.y), widthEmu: num(ext.cx), heightEmu: num(ext.cy), geometry: 'group', paragraphs: [], fontFamily: '',
          rotationDeg: num(a.rot) / 60000, flipH: isVmlTrue(a.flipH), flipV: isVmlTrue(a.flipV),
          drawingStyle: currentGroupStyle,
          group: { off: { x: num(off.x), y: num(off.y) }, ext: { width: num(ext.cx), height: num(ext.cy) }, chOff: { x: num(chOff.x), y: num(chOff.y) }, chExt: { width: num(chExt.cx), height: num(chExt.cy) } },
          children: diagramShapes(node, part, pictureRels, pictureBytes, selected, groupDepth + 1, state, currentGroupStyle),
        })
      } else {
        const model = shape(node, theme, drawingTheme, groupStyle, name) as DrawingContentShape<Text>
        const txBody = child(node, 'txBody')
        if (txBody && adapters.parseDiagramText) model.textBody = adapters.parseDiagramText(txBody, node)
        out.push(model)
      }
    }
    return out
  }
  const uri = attrs(graphic).uri?.toLowerCase() ?? ''
  if (uri.includes('chartex') || uri.includes('model3d') || child(graphic, 'oMath') || child(graphic, 'oMathPara')) {
    contentDiagnostic(context, 'unsupported-content', owner, uri || 'OMML'); return undefined
  }
  const ids = child(graphic, 'relIds')
  if (ids) return follow(attrs(ids).dm, owner, 'diagram')
  const chartRef = child(graphic, 'chart')
  if (chartRef) return follow(attrs(chartRef).id, owner, 'chart')
  const content = child(graphic, 'contentPart')
  if (content) return follow(attrs(content).id, owner, 'contentPart')
  const wsp = child(graphic, 'wsp')
  if (wsp) {
    if (adapters.parseParagraph) {
      const bodyNode = child(wsp, 'bodyPr')
      const body = attrs(bodyNode), pr = child(wsp, 'spPr'), ln = child(pr, 'ln')
      const lineColor = color(child(ln, 'solidFill'), theme)
      const direction = (body.vert ?? undefined) as TextDirection | undefined

      let textWarp: TextWarp | undefined
      const diagnostics: Array<{ kind: 'unsupported-text-alignment' | 'unsupported-text-appearance' | 'unsupported-text-warp'; feature: string; message: string }> = []
      const prstWarpNode = child(bodyNode, 'prstTxWarp')
      if (prstWarpNode) {
        const warpPrst = attrs(prstWarpNode).prst as TextWarpPreset | undefined
        if (warpPrst) {
          if (warpPrst === 'textNoShape') {
            // Handled as unwarped standard text
          } else if (warpPrst === 'textPlain') {
            const nonDefaultAdj = parseTextPlainAdjustment(prstWarpNode)
            if (nonDefaultAdj !== undefined) {
              diagnostics.push({
                kind: 'unsupported-text-warp',
                feature: 'textPlain',
                message: `WordArt warp preset textPlain with nondefault adjustment ${nonDefaultAdj} is unsupported; using unwarped text`,
              })
            }
          } else if (SUPPORTED_TEXT_WARP_PRESETS.has(warpPrst)) {
            const avLst = child(prstWarpNode, 'avLst')
            const adjustments: Record<string, number> = {
              ...(DEFAULT_WARP_ADJUSTMENTS[warpPrst] ?? {}),
              ...parseAdjustGuides(avLst),
            }
            textWarp = {
              preset: warpPrst,
              adjustments: Object.keys(adjustments).length > 0 ? adjustments : undefined,
            }
          } else {
            diagnostics.push({
              kind: 'unsupported-text-warp',
              feature: warpPrst,
              message: `WordArt warp preset ${warpPrst} is unsupported; using unwarped text`,
            })
          }
        }
      }

      const paragraphs = getChildren(child(child(wsp, 'txbx'), 'txbxContent'), 'p').map(adapters.parseParagraph)
      for (const p of paragraphs) {
        const pDiag = (p as any).diagnostics
        if (Array.isArray(pDiag)) {
          for (const d of pDiag) {
            if (!diagnostics.some(existing => existing.kind === d.kind && existing.feature === d.feature)) {
              diagnostics.push(d)
            }
          }
        }
      }

      const rawAnchor = body.anchor?.trim().toLowerCase()
      let anchor: 't' | 'ctr' | 'b' | undefined
      if (rawAnchor === 'ctr' || rawAnchor === 'b' || rawAnchor === 't') {
        anchor = rawAnchor
      } else if (rawAnchor === 'just' || rawAnchor === 'dist') {
        diagnostics.push({
          kind: 'unsupported-text-alignment',
          feature: `anchor-${rawAnchor}`,
          message: `Vertical text anchoring ${rawAnchor} is deferred; using top anchoring`,
        })
      }
      const wrap = body.wrap?.trim().toLowerCase() === 'none' ? false : true
      const autofit = parseAutofitNode(bodyNode)

      return {
        kind: 'textbox',
        paragraphs,
        vertical: direction !== undefined && direction !== 'horz' && ['vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl'].includes(direction),
        ...(direction === 'horz' || direction === 'vert' || direction === 'vert270' || direction === 'wordArtVert' || direction === 'eaVert' || direction === 'mongolianVert' || direction === 'wordArtVertRtl' ? { direction } : {}),
        fontFamily: theme.fonts.get('minorHAnsi') ?? 'Calibri',
        fontSizePt: 12,
        insets: { left: num(body.lIns, 91440), top: num(body.tIns, 45720), right: num(body.rIns, 91440), bottom: num(body.bIns, 45720) },
        ...(anchor ? { anchor } : {}),
        wrap,
        fill: color(child(pr, 'solidFill'), theme),
        line: lineColor ? { color: lineColor, widthEmu: num(attrs(ln).w, 6350) } : undefined,
        ...(textWarp ? { textWarp } : {}),
        ...(autofit ? { autofit } : {}),
        ...(diagnostics.length ? { diagnostics } : {}),
      }
    }
  }
  return undefined
}
/** Public direct loads retain an eagerly usable model; adapters prepare first. */
export async function loadDrawingContent<Paragraph = ContentParagraph, Text = never>(
  pkg: OfficePackage, graphic: XmlNode, owner: string, theme: ContentTheme, drawingTheme: ThemeContext,
  adapters: ContentLoadAdapters<Paragraph, Text> = {}, ancestry: readonly string[] = [owner], depth = 0,
): Promise<DrawingContent<Paragraph, Text> | undefined> {
  const content = await prepareDrawingContent(pkg, graphic, owner, theme, drawingTheme, adapters, ancestry, depth)
  if (content?.kind === 'diagram') void content.shapes
  return content
}

export interface ContentSelection {
  alternate: XmlNode
  node: XmlNode
  representation: 'choice' | 'fallback'
  feature?: string
  reason?: string
}
export interface PreparedCompatibleContent<Paragraph, Text> {
  content?: DrawingContent<Paragraph, Text>
  carrier?: XmlNode
  reference?: XmlNode
  selections: ContentSelection[]
}

/** Resolve vector compatibility on original XML nodes before an adapter reserves a placement. */
export async function prepareCompatibleDrawingContent<Paragraph = ContentParagraph, Text = never>(
  pkg: OfficePackage, root: XmlNode, owner: string, theme: ContentTheme, drawingTheme: ThemeContext,
  adapters: ContentLoadAdapters<Paragraph, Text> = {},
): Promise<PreparedCompatibleContent<Paragraph, Text> | undefined> {
  const maxDepth = 128
  const requirements = ['a', 'p', 'c', 'dgm', 'dsp', 'ink']
  const payloadNames = ['chart', 'relIds', 'contentPart', 'wsp']
  type Prepared = Omit<PreparedCompatibleContent<Paragraph, Text>, 'selections'>
  type Result = { prepared?: Prepared; selections: ContentSelection[]; failure?: string }
  const structuralFeature = (node: XmlNode, depth: number): string | undefined => {
    if (depth >= maxDepth) return 'drawing-depth'
    const uri = attrs(node).uri?.toLowerCase() ?? ''
    if (uri.includes('chartex')) return 'ChartEx'
    if (uri.includes('model3d')) return 'model3D'
    for (const [name, childNode] of orderedChildren(node)) {
      if (name === 'AlternateContent' || name === '#text') continue
      if (name === 'm' || name === 'oMath' || name === 'oMathPara') return 'OMML'
      if (name.toLowerCase() === 'model3d') return 'model3D'
      if (name.toLowerCase() === 'chartex') return 'ChartEx'
      const feature = structuralFeature(childNode, depth + 1)
      if (feature) return feature
    }
    return undefined
  }
  const resolveAlternate = async (alternate: XmlNode, carrier: XmlNode | undefined, depth: number): Promise<Result> => {
    let failure = 'drawing-content'
    const diagnostics = drawingPartContext(pkg).diagnostics
    const failedDiagnostics = new Set<ContentDiagnostic>()
    const discardFailed = () => {
      if (!failedDiagnostics.size) return
      for (let i = diagnostics.length - 1; i >= 0; i--) if (failedDiagnostics.has(diagnostics[i])) diagnostics.splice(i, 1)
    }
    for (const choice of getChildren(alternate, 'Choice')) {
      const unknown = supportedChoiceRequirements(choice, requirements).unknown
      if (unknown) { failure = unknown; continue }
      const diagnosticStart = diagnostics.length
      const result = await resolveBranch(choice, carrier, depth + 1)
      if (result.prepared || !result.failure) { discardFailed(); return { prepared: result.prepared, selections: [{ alternate, node: choice, representation: 'choice' }, ...result.selections] } }
      for (const diagnostic of diagnostics.slice(diagnosticStart)) failedDiagnostics.add(diagnostic)
      failure = result.failure ?? 'drawing-content'
    }
    const fallback = getChildren(alternate, 'Fallback')[0]
    if (fallback) {
      const result = await resolveBranch(fallback, carrier, depth + 1)
      if (result.prepared || !result.failure) { discardFailed(); return { prepared: result.prepared, selections: [{ alternate, node: fallback, representation: 'fallback', feature: failure, reason: `Native fallback selected for unsupported ${failure}` }, ...result.selections] } }
      failure = result.failure ?? 'drawing-content'
    }
    return { selections: [], failure }
  }
  const resolveBranch = async (node: XmlNode, carrier: XmlNode | undefined, depth: number): Promise<Result> => {
    const unsupported = structuralFeature(node, depth)
    if (unsupported) return { selections: [], failure: unsupported }
    let prepared: Prepared | undefined
    const selections: ContentSelection[] = []
    let failure: string | undefined
    for (const [name, childNode] of orderedChildren(node)) {
      if (name === '#text') continue
      let result: Result
      if (name === 'AlternateContent') result = await resolveAlternate(childNode, carrier, depth + 1)
      else if (payloadNames.includes(name)) {
        if (prepared) continue
        const source: XmlNode = { [name]: childNode, ...(carrier ? { '@attrs': attrs(carrier) } : {}) }
        const content = await prepareDrawingContent(pkg, source, owner, theme, drawingTheme, adapters)
        result = content ? { prepared: { content, carrier: carrier ?? childNode, reference: childNode }, selections: [] } : { selections: [], failure: 'drawing-content' }
      } else result = await resolveBranch(childNode, name === 'graphicData' ? childNode : carrier, depth + 1)
      if (result.failure && name === 'AlternateContent') return { selections: [], failure: result.failure }
      failure ??= result.failure
      selections.push(...result.selections)
      prepared ??= result.prepared
    }
    return { prepared, selections, failure: prepared ? undefined : failure }
  }
  const result = await resolveBranch(root, attrs(root).uri !== undefined ? root : undefined, 0)
  return result.prepared || (result.selections.length > 0 && !result.failure)
    ? { ...result.prepared, selections: result.selections } : undefined
}
/** Charge selected placements only; cached nodes may appear at several placements. */
export function reserveDrawingContent<Paragraph, Text>(pkg: OfficePackage, content: DrawingContent<Paragraph, Text>, owner: string, includeContainer = true): boolean {
  const context = drawingPartContext(pkg)
  if (includeContainer && !reserveDrawingNode(context, owner)) return false
  if (content.kind !== 'diagram') return true
  if (deferredDiagrams.get(content)?.materializeSelected()) return true
  const prune = (nodes: DrawingContentShape<Text>[]): DrawingContentShape<Text>[] => {
    const out: DrawingContentShape<Text>[] = []
    for (const node of nodes) {
      if (!reserveDrawingNode(context, owner)) break
      if (node.children) node.children = prune(node.children)
      out.push(node)
    }
    return out
  }
  content.shapes = prune(content.shapes)
  return true
}
