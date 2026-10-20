/**
 * A chart's style and colour parts (spec 5.11; lazy chunk): `cs:chartStyle` and `cs:colorStyle`, reached from
 * the chart part's `.rels` (types `.../chartStyle` and `.../chartColorStyle`). Read before the chart is parsed
 * so the parser can bake series colours from them (charts/style.ts `resolveSeriesColor`):
 *
 * - each part goes through the bounded, admitted chart-part reader (charts/read.ts); its tree is released as
 *   soon as the few values below are extracted;
 * - what is kept (at most CHART_STYLE_COLORS colours and CHART_STYLE_VARIATIONS variations of at most
 *   CHART_STYLE_TRANSFORMS transforms) is charged to the document workspace under its own tag before it is
 *   built, and refunded by `release()` once the parse has used it;
 * - a missing, malformed, unsupported or refused part is skipped: precedence falls to the next level.
 */
import { attrs, getChildren, orderedChildren, type XmlNode } from '../core/xml'
import { parseDrawingColor, type ColorDefinition } from '../drawing/style'
import type { ChartParseContext } from './context'
import { resolveRelatedParts } from './embeddedZip'
import { readChartPart } from './read'
import type { ChartColorPart, ChartStyleEntry, ChartStyleParts, ChartStylePart } from './style'

export const CHART_STYLE_COLORS = 64, CHART_STYLE_VARIATIONS = 64, CHART_STYLE_TRANSFORMS = 16
const COLOR_BYTES = 96, TRANSFORM_BYTES = 64, ENTRY_BYTES = 96
const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
let reads = 0

export interface ChartStylePartsRead { parts: ChartStyleParts; release(): void }

/** Reads the chart's style and colour parts; undefined when it has neither (or neither could be read). */
export async function readChartStyleParts(context: ChartParseContext): Promise<ChartStylePartsRead | undefined> {
  if (context.signal?.aborted) return undefined
  const { workspace } = context, tag = `chart-style:${context.partPath}#${++reads}`
  const release = () => workspace.release(tag)
  const charge = (bytes: number) => workspace.reserve(bytes, tag)
  try {
    const targets = await resolveRelatedParts(context, ['chartStyle', 'chartColorStyle'])
    if (!targets.size) return undefined
    const parts: ChartStyleParts = {}
    const colorPath = targets.get('chartColorStyle'), stylePath = targets.get('chartStyle')
    if (colorPath) { const color = await extract(context, colorPath, root => colorPart(root, charge)); if (color) parts.color = color }
    if (stylePath) { const style = await extract(context, stylePath, root => stylePart(root, charge)); if (style) parts.style = style }
    if (!parts.color && !parts.style) { release(); return undefined }
    return { parts, release }
  } catch {
    release()
    return undefined
  }
}

async function extract<T>(context: ChartParseContext, path: string, read: (root: XmlNode) => T | undefined): Promise<T | undefined> {
  const tree = await readChartPart(context.pkg, path, context)
  if (tree?.kind !== 'tree') return undefined
  try { return read(tree.root) } finally { tree.release() }
}

/** A colour, its transforms capped, charged before it is kept. */
function keepColor(color: ColorDefinition | undefined, charge: (bytes: number) => boolean): ColorDefinition | undefined {
  if (!color) return undefined
  const transforms = color.transforms.slice(0, CHART_STYLE_TRANSFORMS)
  if (!charge(COLOR_BYTES + transforms.length * TRANSFORM_BYTES + 2 * (color.value?.length ?? 0))) return undefined
  return { ...color, transforms }
}

/** `cs:colorStyle`: `meth`, its colour list (any DrawingML colour) and its `cs:variation` transform lists. */
function colorPart(root: XmlNode, charge: (bytes: number) => boolean): ChartColorPart | undefined {
  const colors: ColorDefinition[] = [], variations: ChartColorPart['variations'] = []
  for (const [name, node] of orderedChildren(root)) {
    if (name === 'variation') {
      if (variations.length >= CHART_STYLE_VARIATIONS) continue
      const transforms = orderedChildren(node).filter(([n]) => n !== '#text').slice(0, CHART_STYLE_TRANSFORMS)
        .map(([n, t]) => ({ name: n, ...(attrs(t).val !== undefined ? { value: attrs(t).val } : {}) }))
      if (!charge(COLOR_BYTES + transforms.length * TRANSFORM_BYTES)) break
      variations.push(transforms)
    } else if (name !== '#text' && name !== 'extLst' && colors.length < CHART_STYLE_COLORS) {
      const color = keepColor(colorOf(name, node), charge)
      if (color) colors.push(color)
    }
  }
  if (!colors.length) return undefined
  return { method: attrs(root).meth ?? 'cycle', colors, variations }
}
/** One colour element (`a:schemeClr` and the like) read on its own. */
function colorOf(name: string, node: XmlNode): ColorDefinition | undefined {
  const holder = { '@attrs': {} } as XmlNode
  ;(holder as Record<string, unknown>)[name] = node
  return parseDrawingColor(holder)
}

/** `cs:chartStyle`: the data point entries' spPr colour and fillRef/lnRef `cs:styleClr`. */
function stylePart(root: XmlNode, charge: (bytes: number) => boolean): ChartStylePart | undefined {
  const entry = (name: string, line: boolean): ChartStyleEntry | undefined => {
    const node = child(root, name)
    if (!node) return undefined
    const ref = child(node, line ? 'lnRef' : 'fillRef')
    const raw = attrs(child(ref, 'styleClr')).val
    const styleColor = raw === 'auto' ? 'auto' as const : raw !== undefined && /^\d+$/.test(raw) ? Number(raw) : undefined
    const spPr = child(node, 'spPr')
    const holder = line ? child(child(spPr, 'ln'), 'solidFill') : child(spPr, 'solidFill')
    // The reference's own colour (a phClr stand-in when the spPr has none) is the styleClr; an explicit spPr colour wins.
    const color = holder ? keepColor(parseDrawingColor(holder), charge) : undefined
    if (!color && styleColor === undefined) return undefined
    if (!charge(ENTRY_BYTES)) return undefined
    return { ...(color ? { color } : {}), ...(styleColor !== undefined ? { styleColor } : {}) }
  }
  const dataPoint = entry('dataPoint', false), dataPointLine = entry('dataPointLine', true)
  if (!dataPoint && !dataPointLine) return undefined
  return { ...(dataPoint ? { dataPoint } : {}), ...(dataPointLine ? { dataPointLine } : {}) }
}
