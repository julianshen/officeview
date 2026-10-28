import { attrs, elementNamespaceUri, getChildren, textOf, type XmlNode } from '../../core/xml'
import { parseA1Area } from '../ranges'
import { resolveSheetColor, type SheetColorContext } from '../sheetColors'
import type { SparklineDiagnostic, XlsxSparkline, XlsxSparklineGroup } from './model'

export type Diagnostic = SparklineDiagnostic
export const X14_SPARKLINE_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
export const XM_SPARKLINE_NS = 'http://schemas.microsoft.com/office/excel/2006/main'
export const SPREADSHEETML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
export const SPARKLINE_EXTENSION_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
export const MAX_SPARKLINE_GROUPS_PER_SHEET = 1024
export const MAX_SPARKLINES_PER_GROUP = 16384
export const MAX_POINTS_PER_SPARKLINE = 16384
export const MAX_POINTS_PER_SHEET = 1_048_576

const x14Children = (node: XmlNode | undefined, name: string) => getChildren(node, name).filter(child => elementNamespaceUri(child) === X14_SPARKLINE_NS)
const xmChild = (node: XmlNode | undefined, name: string) => getChildren(node, name).find(child => elementNamespaceUri(child) === XM_SPARKLINE_NS)
const oneOf = <T extends string>(value: string | undefined, choices: readonly T[]): T | undefined => choices.includes(value as T) ? value as T : undefined
const bool = (value: string | undefined): boolean | undefined => value === '1' || value === 'true' ? true : value === '0' || value === 'false' ? false : undefined
const finite = (value: string | undefined): number | undefined => value !== undefined && value.trim() !== '' && Number.isFinite(Number(value)) ? Number(value) : undefined
const areaPoints = (area: NonNullable<ReturnType<typeof parseA1Area>>) => (area.r1 - area.r0 + 1) * (area.c1 - area.c0 + 1)

function diagnostic(out: Diagnostic[], kind: Diagnostic['kind'], feature: string, message: string): void {
  if (!out.some(item => item.feature === feature && item.message === message)) out.push({ kind, feature, message })
}

function refused(group: XlsxSparklineGroup, out: Diagnostic[], feature: string, message: string): void {
  group.refused ??= feature
  diagnostic(out, 'unsupported', feature, message)
}

function areaOf(raw: string | undefined): ReturnType<typeof parseA1Area> {
  if (!raw) return undefined
  let quoted = false
  for (let i = 0; i < raw.length; i++) {
    if (raw[i] === "'") {
      if (quoted && raw[i + 1] === "'") { i++; continue }
      quoted = !quoted
    } else if (!quoted && /\s/.test(raw[i])) return undefined
  }
  if (quoted) return undefined
  const area = parseA1Area(raw)
  if (!area || area.definedName || area.sheet?.startsWith('[')) return undefined
  return area
}

function readSparkline(node: XmlNode, diagnostics: Diagnostic[]): XlsxSparkline {
  const fNode = xmChild(node, 'f'), sqrefNode = xmChild(node, 'sqref')
  const f = fNode ? textOf(fNode) : undefined
  const sqref = sqrefNode ? textOf(sqrefNode) : undefined
  const data = areaOf(f)
  if (f === undefined) diagnostic(diagnostics, 'parse', 'sparkline-data-reference', 'Sparkline is missing its xm:f data reference')
  else if (!data) diagnostic(diagnostics, 'unsupported', 'sparkline-data-reference', `Unsupported sparkline data reference: ${f.slice(0, 80)}`)
  const hostRef = sqref?.trim() ?? ''
  const hostMatch = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(hostRef)
  let host: XlsxSparkline['host']
  if (hostMatch) {
    const col = hostMatch[1].toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    const row = Number(hostMatch[2]) - 1
    if (Number.isSafeInteger(col) && Number.isSafeInteger(row) && row >= 0 && row < 1_048_576 && col >= 0 && col < 16_384) host = { row, col }
  }
  if (sqref === undefined) diagnostic(diagnostics, 'parse', 'sparkline-host-reference', 'Sparkline is missing its xm:sqref host cell')
  else if (!host) diagnostic(diagnostics, 'parse', 'sparkline-host-reference', `Sparkline sqref must be one cell: ${sqref.slice(0, 80)}`)
  return { host, data, raw: { ...(f !== undefined ? { f } : {}), ...(sqref !== undefined ? { sqref } : {}) } }
}

function readGroup(node: XmlNode, color: SheetColorContext, diagnostics: Diagnostic[]): XlsxSparklineGroup {
  const a = attrs(node)
  const typeRaw = a.type
  const type = oneOf(typeRaw, ['line', 'column', 'stacked'] as const) ?? 'line'
  const lineWeightRaw = finite(a.lineWeight)
  const lineWeightPt = lineWeightRaw !== undefined && lineWeightRaw > 0 ? lineWeightRaw : 0.75
  const displayRaw = a.displayEmptyCellsAs
  const displayEmptyCellsAs = oneOf(displayRaw, ['gap', 'zero', 'span'] as const) ?? 'zero'
  const minRaw = a.minAxisType, maxRaw = a.maxAxisType
  const minAxisType = oneOf(minRaw, ['individual', 'group', 'custom'] as const) ?? 'individual'
  const maxAxisType = oneOf(maxRaw, ['individual', 'group', 'custom'] as const) ?? 'individual'
  const group: XlsxSparklineGroup = {
    type, lineWeightPt, displayEmptyCellsAs,
    markers: false, high: false, low: false, first: false, last: false, negative: false,
    displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false,
    minAxisType, maxAxisType, colors: {}, sparklines: [],
  }
  for (const key of ['markers', 'high', 'low', 'first', 'last', 'negative', 'displayXAxis', 'displayHidden', 'dateAxis', 'rightToLeft'] as const) {
    const parsed = bool(a[key])
    if (parsed !== undefined) group[key] = parsed
    else if (a[key] !== undefined) diagnostic(diagnostics, 'parse', 'sparkline-boolean', `Invalid sparkline ${key} value ${a[key]}; using false`)
  }
  if (typeRaw !== undefined && !oneOf(typeRaw, ['line', 'column', 'stacked'] as const)) refused(group, diagnostics, 'sparkline-type', `Unsupported sparkline type: ${typeRaw}`)
  if (displayRaw !== undefined && !oneOf(displayRaw, ['gap', 'zero', 'span'] as const)) refused(group, diagnostics, 'sparkline-empty-cells', `Unsupported displayEmptyCellsAs: ${displayRaw}`)
  if (minRaw !== undefined && !oneOf(minRaw, ['individual', 'group', 'custom'] as const)) refused(group, diagnostics, 'sparkline-min-axis-type', `Unsupported minAxisType: ${minRaw}`)
  if (maxRaw !== undefined && !oneOf(maxRaw, ['individual', 'group', 'custom'] as const)) refused(group, diagnostics, 'sparkline-max-axis-type', `Unsupported maxAxisType: ${maxRaw}`)
  if (a.lineWeight !== undefined && (lineWeightRaw === undefined || lineWeightRaw <= 0)) diagnostic(diagnostics, 'parse', 'sparkline-line-weight', `Invalid lineWeight ${a.lineWeight}; using 0.75pt`)
  for (const [field, axisType] of [['manualMin', minAxisType], ['manualMax', maxAxisType]] as const) {
    const raw = a[field]
    if (raw === undefined) continue
    if (axisType !== 'custom') continue
    const value = finite(raw)
    if (value === undefined) refused(group, diagnostics, 'sparkline-axis-nonconforming', `Invalid ${field} value: ${raw}`)
    else if (field === 'manualMin') group.manualMin = value
    else group.manualMax = value
  }
  if ((a.manualMin !== undefined && minAxisType !== 'custom') || (a.manualMax !== undefined && maxAxisType !== 'custom')) {
    diagnostic(diagnostics, 'parse', 'sparkline-axis-manual-ignored', 'manualMin/manualMax ignored unless their axis type is custom')
  }
  if (minAxisType === 'custom' && a.manualMin === undefined) refused(group, diagnostics, 'sparkline-axis-nonconforming', 'custom minAxisType requires manualMin')
  if (maxAxisType === 'custom' && a.manualMax === undefined) refused(group, diagnostics, 'sparkline-axis-nonconforming', 'custom maxAxisType requires manualMax')

  const colorNames = { colorSeries: 'series', colorNegative: 'negative', colorAxis: 'axis', colorMarkers: 'markers', colorFirst: 'first', colorLast: 'last', colorHigh: 'high', colorLow: 'low' } as const
  for (const [element, key] of Object.entries(colorNames) as Array<[keyof typeof colorNames, typeof colorNames[keyof typeof colorNames]]>) {
    const colorNode = x14Children(node, element)[0]
    if (!colorNode) continue
    const resolved = resolveSheetColor(attrs(colorNode), color)
    if (resolved !== undefined) group.colors[key] = resolved
  }

  // Date-axis rendering is staged out in Stage 1. Preserve a valid optional
  // date range for later stages, but leave malformed and unused ranges silent.
  const dateFormulaNode = xmChild(node, 'f')
  const dateFormula = dateFormulaNode ? textOf(dateFormulaNode) : undefined
  if (dateFormula) group.dateRange = areaOf(dateFormula)
  const sparklineContainer = x14Children(node, 'sparklines')[0]
  const sparkNodes = x14Children(sparklineContainer, 'sparkline')
  if (sparkNodes.length > MAX_SPARKLINES_PER_GROUP) {
    group.incomplete = true
    diagnostic(diagnostics, 'unsupported', 'sparkline-count-limit', `Sparkline group exceeds ${MAX_SPARKLINES_PER_GROUP} entries`)
  }
  for (const sparkline of sparkNodes.slice(0, MAX_SPARKLINES_PER_GROUP)) {
    const parsed = readSparkline(sparkline, diagnostics)
    if (!parsed.host || !parsed.data) group.incomplete = true
    if (parsed.data) {
      const points = areaPoints(parsed.data)
      if (points > MAX_POINTS_PER_SPARKLINE) {
        parsed.data = undefined
        group.incomplete = true
        diagnostic(diagnostics, 'unsupported', 'sparkline-point-limit', `Sparkline exceeds ${MAX_POINTS_PER_SPARKLINE} source points`)
      } else if (parsed.data.r0 !== parsed.data.r1 && parsed.data.c0 !== parsed.data.c1) {
        parsed.data = undefined
        group.incomplete = true
        diagnostic(diagnostics, 'unsupported', 'sparkline-data-reference', 'Sparkline data reference must be one-dimensional')
      }
    }
    group.sparklines.push(parsed)
  }
  return group
}

/** Map one x14:sparklineGroups extension node into bounded renderer-neutral records. */
export function parseSparklineGroups(groupsNode: XmlNode, color: SheetColorContext, diagnostics: Diagnostic[]): XlsxSparklineGroup[] {
  const nodes = x14Children(groupsNode, 'sparklineGroup')
  if (nodes.length > MAX_SPARKLINE_GROUPS_PER_SHEET) diagnostic(diagnostics, 'unsupported', 'sparkline-group-limit', `Worksheet exceeds ${MAX_SPARKLINE_GROUPS_PER_SHEET} sparkline groups`)
  const groups = nodes.slice(0, MAX_SPARKLINE_GROUPS_PER_SHEET).map(node => readGroup(node, color, diagnostics))
  let points = 0
  for (const group of groups) for (const sparkline of group.sparklines) {
    if (!sparkline.data) continue
    const count = areaPoints(sparkline.data)
    if (points + count > MAX_POINTS_PER_SHEET) {
      sparkline.data = undefined
      group.incomplete = true
      diagnostic(diagnostics, 'unsupported', 'sparkline-sheet-point-limit', `Worksheet exceeds ${MAX_POINTS_PER_SHEET} sparkline source points`)
    } else points += count
  }
  return groups
}
