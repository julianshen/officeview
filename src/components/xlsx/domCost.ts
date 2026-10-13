/**
 * S10.2 main ledger: the viewer's DOM (with each element's React fiber, props
 * and Blink layout objects) charged by element class. Classes, by the nearest
 * classifying ancestor (the element itself included):
 * - cell: inside a `[role=gridcell]` (the semantic cell and its value span);
 * - header: inside a `[role=columnheader]` or `[role=rowheader]` (header
 *   select/sort/filter/more buttons with their handlers and glyphs);
 * - other: everything else the viewer renders (wrappers, toolbar, sheet
 *   navigation, status, live regions, the pane and its canvases' elements).
 * Per-class costs come from a calibration that renders the real GridSemantics
 * tree (see DOM_CLASS_BYTES).
 */
import { DOM_CLASS_BYTES, type DomClass } from '../../xlsx/view/mainLedger'

const CELL = '[role=gridcell], [role=gridcell] *'
const HEADER = '[role=columnheader], [role=columnheader] *, [role=rowheader], [role=rowheader] *'

export type DomClassCounts = Record<DomClass, number>
/** Elements under `root` (excluded) by class. */
export function domClassCounts(root: Element | null | undefined): DomClassCounts {
  if (!root) return { cell: 0, header: 0, other: 0 }
  const total = root.getElementsByTagName('*').length
  const cell = root.querySelectorAll(CELL).length
  const header = root.querySelectorAll(HEADER).length
  return { cell, header, other: Math.max(0, total - cell - header) }
}
/** Charged bytes for the counts. */
export function domClassBytes(counts: DomClassCounts): number {
  return counts.cell * DOM_CLASS_BYTES.cell + counts.header * DOM_CLASS_BYTES.header + counts.other * DOM_CLASS_BYTES.other
}
