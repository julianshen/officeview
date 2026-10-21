/**
 * Chart parse contexts (spec 6.4, 6.5).
 *
 * `documentChartContext` is created once per parse and owns the document's
 * chart workspace. `parseXlsx` (later `parseDocx`/`parsePptx`) take it as the
 * `chartContext` option and thread it to `ContentLoadAdapters`; without one,
 * the lazy chart chunk creates it on demand, still once per parse (the parse
 * resets the package slot). A standalone or transferred render makes its own.
 *
 * This module is lazy: eager code imports its types only.
 */
import type { OfficePackage } from '../core/zip'
import type { ThemeContext } from '../drawing/style'
import { drawingPartContext } from '../drawing/parts'
import { createChartWorkspace, scopeChartWorkspace, type ChartParentLedger, type ChartProfile, type ChartWorkspace } from './workspace'

/** Document-level: one per parse, shared by every chart of the document. */
export interface ChartDocumentContext {
  readonly profile: ChartProfile
  readonly signal?: AbortSignal
  readonly workspace: ChartWorkspace
}

/** Per chart part. */
export interface ChartParseContext {
  theme?: ThemeContext
  pkg: OfficePackage
  partPath: string
  signal?: AbortSignal
  profile: ChartProfile
  workspace: ChartWorkspace
  epoch?: '1900' | '1904'
  /** The chart's style and colour parts, read before the parse (charts/styleParts.ts; spec 5.11). */
  styleParts?: import('./style').ChartStyleParts
}

/** The document's chart context; `parent` is the paged worker's document-pool share, when there is one. */
export function documentChartContext(profile: ChartProfile, signal?: AbortSignal, parent?: ChartParentLedger): ChartDocumentContext {
  return { profile, ...(signal ? { signal } : {}), workspace: createChartWorkspace(profile, parent) }
}

/** The context for one chart part, derived from the document's. */
export function chartParseContext(doc: ChartDocumentContext, pkg: OfficePackage, partPath: string, theme?: ThemeContext, epoch?: '1900' | '1904'): ChartParseContext {
  return { ...(theme ? { theme } : {}), pkg, partPath, ...(doc.signal ? { signal: doc.signal } : {}), profile: doc.profile, workspace: doc.workspace, ...(epoch ? { epoch } : {}) }
}

/**
 * A package that carries its own memory ledger (the paged worker's
 * `RangePackage`): the chart workspace it gets on demand uses that ledger as its
 * parent, so chart bytes are counted once in the document pool share (D0).
 */
interface LedgeredPackage { chartLedger?: ChartParentLedger; chartProfile?: ChartProfile; chartDocument?: { workspace?: ChartWorkspace } }
/**
 * The production hook: the package's document context for one chart part,
 * created on demand (desktop unless the package says). A paged read's package
 * also carries its document's chart facet holder: the read then works in a
 * scoped view of that one document facet (the chart sublimit across every read
 * and retained chart), charged to the read's share until the read is answered.
 */
export function chartParseContextFor(pkg: OfficePackage, partPath: string, theme?: ThemeContext): ChartParseContext {
  const slot = drawingPartContext(pkg), own = pkg as LedgeredPackage
  if (!slot.chart && own.chartDocument && own.chartLedger) {
    const profile = own.chartProfile ?? 'desktop'
    const document = own.chartDocument.workspace ??= createChartWorkspace(profile)
    slot.chart = { profile, workspace: scopeChartWorkspace(document, own.chartLedger) }
  }
  return chartParseContext(slot.chart ??= documentChartContext(own.chartProfile ?? 'desktop', undefined, own.chartLedger), pkg, partPath, theme)
}
