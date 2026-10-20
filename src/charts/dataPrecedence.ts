/**
 * Per-role data precedence for DOCX/PPTX charts (spec 6.4; lazy chunk).
 *
 * Roles: series name, categories, values, x values (bubble sizes and custom
 * error amounts join when their families are read). For each role:
 *
 * 1. **Saved data** (a cache or a literal, even partial) is used as read by
 *    the chart part parser: sparse indices kept, missing ones blank,
 *    `ptCount=0` a present, empty role. Saved data is never overridden.
 * 2. **No saved data, a reference**: the role is read from the chart's
 *    embedded workbook (`c:externalData`, charts/embeddedWorkbook.ts).
 * 3. **Neither**, or a reference that cannot be read (no embedding, an
 *    external target or name, a missing sheet or name, a corrupt workbook):
 *    the role is unavailable, with a diagnostic.
 *
 * Values are required for every family, x values for scatter (when the
 * series references them): one unavailable required role makes the whole
 * chart unsupported ("This chart's data is not available"); a series is
 * never drawn silently missing. Categories and names are optional and fall
 * back to their defaults (1..n, "Series n"). Mixing cached and workbook
 * roles is allowed. A budget breach refuses the chart.
 *
 * A chart whose roles all have saved data never opens the embedding. XLSX
 * charts keep their own fill path (the containing workbook's cells).
 */
import { attrs, getChildren, type XmlNode } from '../core/xml'
import type { ThemeContext } from '../drawing/style'
import type { ChartParseContext } from './context'
import type { ChartDiagnostic, ChartModel, ChartRange, ChartResult, ChartSeries } from './model'
import { CHART_POINT_LIMIT, parseXlsxChart } from './parse'
import { chartHasPending, fillChartRanges, seriesLimit, type ChartCellLookup } from './fill'
import { chartModelOwner, releaseChartModel, succeedChartModel } from './own'
import { stringBytes } from './points'
import { readChartStyleParts } from './styleParts'
import { requestEmbeddedCells, type EmbeddedCellsResult, type EmbeddedRequest, type EmbeddedResolution } from './embeddedWorkbook'

export const CHART_DATA_UNAVAILABLE = "This chart's data is not available"
const XLSX_PART = /^\/?xl\//i

/**
 * The chart part adapter for every format (spec 6.2, 6.3): reads the model, applies DOCX/PPTX data
 * precedence, and shapes the representation. `parse` is the module's parser (passed by the caller so a
 * test's module-level substitute is honoured). A discarded model returns its storage to the workspace.
 */
export async function parseChartPart(root: XmlNode, theme: ThemeContext | undefined, context: ChartParseContext, parse: typeof parseXlsxChart = parseXlsxChart): Promise<ChartResult> {
  let chart: ChartModel | undefined
  // The style and colour parts (spec 5.11) are read first: the parse bakes the series colours from them.
  const style = await readChartStyleParts(context)
  try {
    chart = parse(root, theme, style ? { ...context, styleParts: style.parts } : context)
    style?.release()
    if (chart && !XLSX_PART.test(context.partPath)) chart = await applyDataPrecedence(chart, root, context)
  } catch { if (chart) releaseChartModel(chart); chart = undefined } finally { style?.release() }
  if (chart && !chart.unsupported) return { kind: 'native', model: chart }
  if (chart) releaseChartModel(chart)
  return { kind: chart?.refused ? 'refused' : 'unsupported', reason: chart?.unsupported ?? 'This chart could not be read', ...(chart?.title ? { title: chart.title } : {}) }
}

type Role = 'values' | 'xValues' | 'categories' | 'name'
interface Slot { group: number; series: number; role: Role }
const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]

/** Successor model: takes over the old one's charge; its new diagnostics are admitted (dropped if refused). */
function successor(model: ChartModel, patch: Partial<ChartModel>, diagnostics: ChartDiagnostic[] = []): ChartModel {
  const owner = chartModelOwner(model)
  const admitted = diagnostics.filter(d => !owner || owner.workspace.reserve(stringBytes(d.code) + stringBytes(d.reason), owner.tag))
  const next: ChartModel = { ...model, ...patch, ...(admitted.length ? { diagnostics: [...model.diagnostics ?? [], ...admitted] } : {}) }
  succeedChartModel(model, next)
  return next
}
const unsupported = (model: ChartModel, diagnostics: ChartDiagnostic[]) => successor(model, { unsupported: CHART_DATA_UNAVAILABLE, preparedAxes: undefined }, diagnostics)

/**
 * Applies the precedence to a parsed DOCX/PPTX chart model (returns the successor model; never throws).
 * The embedded workbook is opened only when some role has a reference and no saved data.
 */
export async function applyDataPrecedence(model: ChartModel, root: XmlNode, context: ChartParseContext): Promise<ChartModel> {
  if (model.unsupported) return model
  // A required role (values; x values for scatter, and bubble sizes once bubble is read) with neither saved
  // data nor a reference, cached-only charts included.
  const absent = model.groups.flatMap(group => group.series.flatMap(series => series.absent ?? []))
  if (absent.length) return unsupported(model, [{ code: 'chart-data-unavailable', reason: `A series has no ${absent.includes('values') ? 'values' : 'x values'}: neither saved data nor a reference` }])
  if (!chartHasPending(model)) return model
  // Every role still to read, in a fixed order: the shared categories first, then each series' roles.
  const slots: Array<Slot | 'categories'> = [], requests: EmbeddedRequest[] = []
  if (model.categoryPending) { slots.push('categories'); requests.push({ range: model.categoryPending, text: true, limit: CHART_POINT_LIMIT }) }
  model.groups.forEach((group, g) => group.series.forEach((series, s) => {
    const pending = series.pending
    if (!pending) return
    const limit = seriesLimit(series)
    if (pending.values) { slots.push({ group: g, series: s, role: 'values' }); requests.push({ range: pending.values, text: false, limit }) }
    if (pending.xValues) { slots.push({ group: g, series: s, role: 'xValues' }); requests.push({ range: pending.xValues, text: false, limit }) }
    if (pending.categories) { slots.push({ group: g, series: s, role: 'categories' }); requests.push({ range: pending.categories, text: true, limit: CHART_POINT_LIMIT }) }
    if (pending.name) { slots.push({ group: g, series: s, role: 'name' }); requests.push({ range: pending.name, text: true, limit: 1, single: true }) }
  }))
  const rId = attrs(child(root, 'externalData')).id
  let result: EmbeddedCellsResult
  try {
    result = rId ? await requestEmbeddedCells(context, rId, requests) : { kind: 'unavailable', diagnostic: { code: 'chart-data-missing', reason: 'The chart has no embedded workbook' } }
  } catch { result = { kind: 'unavailable', diagnostic: { code: 'chart-data-corrupt', reason: "The chart's embedded workbook could not be read" } } }
  if (result.kind === 'refused') return successor(model, { unsupported: 'This chart needs more memory than is available for charts', refused: true, preparedAxes: undefined }, [result.diagnostic])
  if (result.kind === 'cancelled') return successor(model, { unsupported: 'Reading the chart data was cancelled', preparedAxes: undefined }, [result.diagnostic])
  const cells = result.kind === 'cells' ? result : undefined
  try {
    const resolved: EmbeddedResolution[] = cells ? cells.resolved : requests.map(() => ({ unavailable: (result as Exclude<EmbeddedCellsResult, { kind: 'cells' }>).diagnostic }))
    const diagnostics: ChartDiagnostic[] = []
    const seen = new Set<string>()
    const note = (d: ChartDiagnostic) => { const id = `${d.code}\n${d.reason}`; if (!seen.has(id)) { seen.add(id); diagnostics.push(d) } }
    let required = false
    // Concrete ranges replace the pending references; an unavailable role is dropped (its default stays).
    let categoryPending: ChartRange | undefined
    const pendings = model.groups.map(group => group.series.map(() => ({} as NonNullable<ChartSeries['pending']>)))
    slots.forEach((slot, i) => {
      const resolution = resolved[i]
      if ('unavailable' in resolution) {
        note(resolution.unavailable)
        if (slot !== 'categories' && (slot.role === 'values' || slot.role === 'xValues')) required = true
        return
      }
      if (slot === 'categories') categoryPending = resolution.range
      else pendings[slot.group][slot.series][slot.role] = resolution.range
    })
    if (required) return unsupported(model, diagnostics)
    const groups = model.groups.map((group, g) => ({ ...group, series: group.series.map((series, s) => {
      if (!series.pending) return series
      const { pending: _old, ...rest } = series
      void _old
      const pending = pendings[g][s]
      return Object.keys(pending).length ? { ...rest, pending } : rest
    }) }))
    const epoch = cells?.epoch
    const next = successor(model, { groups, categoryPending, ...(epoch === '1904' ? { workbookEpoch: '1904' as const } : {}) }, diagnostics)
    if (!cells || !chartHasPending(next)) return next
    // Formula cells without a saved value are unavailable points, diagnosed once per chart.
    let unsaved = 0
    const lookup: ChartCellLookup = (sheet, row, col, text, allocation) => {
      const cell = cells.lookup(sheet, row, col, text, allocation)
      if (cell?.unavailable) unsaved++
      return cell
    }
    const filled = fillChartRanges(next, cells.ownerSheet, lookup, { hidden: cells.hidden, workspace: context.workspace, epoch: cells.epoch })
    return unsaved && !filled.unsupported ? successor(filled, {}, [{ code: 'chart-formula-unsaved', reason: `${unsaved} formula ${unsaved === 1 ? 'cell has' : 'cells have'} no saved value` }]) : filled
  } finally { cells?.done() }
}
