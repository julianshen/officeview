/**
 * Slicer and timeline display model (spec §3, Task 12). Types only: eager code
 * may `import type` from here; the parsers, style resolution and painters live
 * in the lazy pivots chunk beside it.
 *
 * A control's appearance is resolved when its model is built (theme and the
 * workbook's custom styles are at hand then), so painting is a pure function of
 * the model, its box, the scale and the device transform. Colours are `#RRGGBB`.
 */

/** A box in the painter's canvas units (worksheet CSS px). */
export interface Box { x: number; y: number; width: number; height: number }

/** The x14 slicer style element an item paints with. Hover states exist in the
 * style but a view-only render never paints them (Ruling R). */
export type SlicerItemState = 'selectedItemWithData' | 'unselectedItemWithData' | 'selectedItemWithNoData' | 'unselectedItemWithNoData'

export interface SlicerItemPaint { fill: string; border: string; text: string; bold: boolean }

export interface SlicerAppearance {
  /** wholeTable: the 1 px frame and the slicer's background. */
  border: string
  background: string
  /** headerRow: caption text and the 1 px separator under the header. */
  header: { fill: string; text: string; bold: boolean; separator: string }
  items: Record<SlicerItemState, SlicerItemPaint>
  /** The theme's minor (body) font. */
  fontFamily: string
}

export interface SlicerItem { label: string; selected: boolean; noData: boolean }

export interface SlicerModel {
  name: string
  caption?: string
  showCaption: boolean
  columnCount: number
  /** Item height in EMU (`slicer@rowHeight`). */
  rowHeightEmu: number
  /** The style painted (an unknown name is already replaced by the default). */
  style: string
  /** `slicer@startItem` as written; the painter clamps it so the last row fills the view. */
  startItem: number
  /** Display order (the slicer cache's item order, no-data items after the others when cross-filtering). */
  items: SlicerItem[]
  /** The pivot tables (or, for table slicers, tables) the slicer filters. */
  drives: { pivotTable?: string; table?: string }[]
  appearance?: SlicerAppearance
}

export type TimelineLevel = 'years' | 'quarters' | 'months' | 'days'

export interface TimelineAppearance {
  border: string
  background: string
  header: { text: string; bold: boolean; separator: string }
  /** Selection label ("All Periods") and the time level picker. */
  selectionLabel: string
  timeLevel: string
  /** Period labels (months, quarters, days) and the year row above them. */
  periodLabel: string
  yearLabel: string
  /** Time bar: selected periods (flat [approx] for Excel's vertical gradient), unselected periods, and the divider
   * after a selected period (after an unselected one it is the background). Static timelines draw no selection
   * handles (Ruling AK). */
  selected: string
  unselected: { fill: string; border: string }
  divider: string
  /** Horizontal scrollbar: pale accent track, grey thumb with a lighter edge. */
  track: string
  thumb: string
  thumbEdge: string
  fontFamily: string
}

export interface TimelineModel {
  name: string
  caption?: string
  style: string
  level: TimelineLevel
  /** UTC milliseconds; `end` inclusive (the last selected day), absent for All Periods. */
  selection?: { start: number; end: number }
  /** UTC milliseconds, the cache state's bounds; `end` exclusive. */
  range: { start: number; end: number }
  /** UTC milliseconds: the first date in view (`timeline@scrollPosition`), else `range.start`. */
  scrollPosition?: number
  show?: { header: boolean; selectionLabel: boolean; timeLevel: boolean; scrollbar: boolean }
  appearance?: TimelineAppearance
}

/** A drawing's control content: the `sle:slicer` / `tsle:timeslicer` graphic frame, resolved. */
export type XlsxControlContent = { kind: 'slicer'; slicer: SlicerModel } | { kind: 'timeline'; timeline: TimelineModel }

export interface ControlDiagnostic { kind: 'unsupported' | 'parse'; feature: string; message: string }

/** One worksheet's resolved controls, by the name its drawing frames carry. */
export interface SheetControls {
  slicers: Map<string, SlicerModel>
  timelines: Map<string, TimelineModel>
  diagnostics: ControlDiagnostic[]
}

/** What a sheet asks the lazy chunk for: its worksheet part, the extension ids
 * (relationship ids) of its slicer and timeline parts, the theme, and the
 * styles source for custom styles. */
export interface SheetControlsRequest {
  sheetPath: string
  slicerIds: readonly string[]
  timelineIds: readonly string[]
  theme: import('../../../drawing/style').ThemeContext
  styles?: import('../../tableStyles/model').TableStyleSource
}
