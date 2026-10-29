/**
 * Built-in Excel table, pivot, slicer and timeline style provider.
 * Lazy chunk: eager code may only import type from here.
 */
import { getTableStyle } from './table'
import { getPivotStyle } from './pivot'
import { getSlicerTableStyle, getSlicerStyle } from './slicer'
import type { SlicerStyleDef, TableStyleDef } from '../model'

/**
 * Returns the TableStyleDef for any built-in table style, pivot style,
 * or slicer/timeline frame style. Returns undefined for unknown names.
 */
export function builtinTableStyle(name: string): TableStyleDef | undefined {
  if (name.startsWith('TableStyle')) return getTableStyle(name)
  if (name.startsWith('PivotStyle')) return getPivotStyle(name)
  if (name.startsWith('SlicerStyle') || name.startsWith('TimeSlicerStyle')) return getSlicerTableStyle(name)
  return undefined
}

/**
 * Returns the SlicerStyleDef (item states) for any built-in slicer or timeline style.
 * Returns undefined for table styles, pivot styles, or unknown names.
 */
export function builtinSlicerStyle(name: string): SlicerStyleDef | undefined {
  if (name.startsWith('SlicerStyle') || name.startsWith('TimeSlicerStyle')) return getSlicerStyle(name)
  return undefined
}
