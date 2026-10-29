/**
 * Table, pivot and slicer style model (spec §6.5). Types only: eager code may
 * `import type` from here; every runtime value of the style engine lives in the
 * lazy modules beside it (parseStylesExtras.ts, resolve.ts).
 */
import type { XlsxCellStyle } from '../types'

/** ECMA-376 Part 1 §18.18.83 ST_TableStyleType. */
export type TableStyleElementType =
  | 'wholeTable' | 'headerRow' | 'totalRow' | 'firstColumn' | 'lastColumn'
  | 'firstRowStripe' | 'secondRowStripe' | 'firstColumnStripe' | 'secondColumnStripe'
  | 'firstHeaderCell' | 'lastHeaderCell' | 'firstTotalCell' | 'lastTotalCell'
  | 'firstSubtotalColumn' | 'secondSubtotalColumn' | 'thirdSubtotalColumn'
  | 'firstSubtotalRow' | 'secondSubtotalRow' | 'thirdSubtotalRow'
  | 'blankRow'
  | 'firstColumnSubheading' | 'secondColumnSubheading' | 'thirdColumnSubheading'
  | 'firstRowSubheading' | 'secondRowSubheading' | 'thirdRowSubheading'
  | 'pageFieldLabels' | 'pageFieldValues'

/** An unresolved SpreadsheetML colour reference. Indexed references are
 * resolved against the workbook palette at parse time (it lives in styles.xml);
 * theme references stay symbolic until a theme is applied. */
export type ColorSpec = { rgb: string } | { theme: number; tint?: number } | { indexed: number }

export type DxfBorderSide = 'left' | 'right' | 'top' | 'bottom' | 'horizontal' | 'vertical'

/** A differential format (`<dxf>`): only what it sets is present. */
export interface DxfStyle {
  /** `underline: 'none'` is an explicit clear of a lower element's underline. */
  font?: { name?: string; sizePt?: number; bold?: boolean; italic?: boolean; underline?: XlsxCellStyle['underline'] | 'none'; strike?: boolean; color?: ColorSpec }
  /** In a dxf a solid fill's colour is `bg` (fgColor is the pattern colour). */
  fill?: { pattern?: string; fg?: ColorSpec; bg?: ColorSpec }
  /** `horizontal`/`vertical` are the element band's inner lines (§18.8.40). */
  border?: Partial<Record<DxfBorderSide, { style: string; color?: ColorSpec }>>
  numFmt?: { id: number; code?: string }
}

export interface TableStyleDef {
  name: string
  /** `pivot`/`table` attributes (both default true). */
  pivot: boolean
  table: boolean
  elements: Partial<Record<TableStyleElementType, { dxf: DxfStyle; size?: number }>>
}

/** x14 slicer style element types (MS-XLSX ST_SlicerStyleType). */
export type SlicerStyleElementType =
  | 'unselectedItemWithData' | 'selectedItemWithData'
  | 'unselectedItemWithNoData' | 'selectedItemWithNoData'
  | 'hoveredUnselectedItemWithData' | 'hoveredSelectedItemWithData'
  | 'hoveredUnselectedItemWithNoData' | 'hoveredSelectedItemWithNoData'

/** A custom slicer style: its whole-slicer and header parts live in the
 * same-named `tableStyle`; its item states in x14 dxfs. */
export interface SlicerStyleDef {
  name: string
  elements: Partial<Record<SlicerStyleElementType, { dxf: DxfStyle }>>
}

export interface StylesExtras {
  dxfs: DxfStyle[]
  tableStyles: Map<string, TableStyleDef>
  defaultTableStyle?: string
  defaultPivotStyle?: string
  slicerStyles: Map<string, SlicerStyleDef>
  defaultSlicerStyle?: string
}

/** Which outer edges of an element's band a cell lies on. A bare role (no
 * position) is a band of one cell: every edge is outer. */
export interface BandEdges { top: boolean; bottom: boolean; left: boolean; right: boolean }
export type RoleRef = TableStyleElementType | { element: TableStyleElementType; edges: BandEdges }

export interface TableStyleDiagnostic { kind: 'approximate' | 'unsupported'; feature: string; message: string }
