/**
 * Paintables metadata for charts (spec 6.6, 6.5): before a document's pages are handed to a viewer, each chart's
 * user-shape pictures are decoded in this realm (so the first paint of a model transferred from a worker realm
 * already shows them), and each page gets its charts' accessible descriptions with their placement. The
 * descriptions are charged to the document's chart workspace and refunded when the pages are released.
 */
import type { ChartDocumentContext } from './context'
import type { ChartModel } from './model'
import { chartDescription } from './describe'
import { chartUserShapesReady } from './userShapes'

/** A chart as placed on one page or slide, in page px. */
export interface PlacedChart { chart: ChartModel; x: number; y: number; width: number; height: number }
/** A chart's accessible description with its placement on the page (page px). */
export interface PaintableChart { description: string; x: number; y: number; width: number; height: number }

/** At most this many described charts per page; the rest are summarised by the count in the list's name. */
export const CHART_DESCRIPTIONS_PER_PAGE = 64
/** Bytes charged per entry beside its string (UTF-16, charged once per distinct model). */
const ENTRY_BYTES = 64
let serial = 0

/**
 * Decodes the user-shape pictures of every chart on the pages (charged to the workspace like any paint-time
 * decode), then describes each page's charts in reading order. A refused reservation omits the remaining
 * descriptions rather than failing the pages.
 */
export async function prepareChartPaintables(context: ChartDocumentContext, pages: ReadonlyArray<readonly PlacedChart[]>): Promise<{ charts: PaintableChart[][]; release(): void }> {
  const workspace = context.workspace
  const models = new Set<ChartModel>()
  for (const page of pages) for (const placed of page) models.add(placed.chart)
  await Promise.all([...models].map(model => model.userShapes?.images.length ? chartUserShapesReady(model, workspace).catch(() => {}) : undefined))
  const tag = `chart-descriptions#${++serial}`
  const described = new Map<ChartModel, string>()
  let refused = false
  const charts = pages.map(page => {
    const out: PaintableChart[] = []
    for (const placed of page.slice(0, CHART_DESCRIPTIONS_PER_PAGE)) {
      if (refused) break
      let description = described.get(placed.chart)
      const bytes = ENTRY_BYTES + (description === undefined ? 2 * (description = chartDescription(placed.chart)).length : 0)
      if (!workspace.reserve(bytes, tag)) { refused = true; break }
      described.set(placed.chart, description)
      out.push({ description, x: placed.x, y: placed.y, width: placed.width, height: placed.height })
    }
    return out
  })
  return { charts, release: () => workspace.release(tag) }
}
