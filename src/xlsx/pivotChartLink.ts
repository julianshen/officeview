/**
 * Which pivot a chart's `c:pivotSource` names (Task 27 r2 R2-4): `[book.xlsx]Sheet!Name`, with the sheet quoted when it
 * needs it (`'[book.xlsx]My Sheet'!Name`, `''` for a quote). Pivot names repeat across sheets (Excel's default name),
 * so a chart links a pivot by sheet and name; a reference without a sheet links by name. Shared by the resident client,
 * the viewer's chart install and the worker, so all three agree. Tiny and dependency-free.
 */
export interface PivotSourceRef { sheet?: string; pivot: string }
export function pivotSourceRef(name: string): PivotSourceRef {
  const bang = name.lastIndexOf('!')
  if (bang < 0) return { pivot: name }
  let sheet = name.slice(0, bang)
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'")
  sheet = sheet.replace(/^\[[^\]]*\]/, '')
  return { sheet, pivot: name.slice(bang + 1) }
}
/** True when `chart` is a pivot chart of the pivot `pivotName` on the sheet `sheetName` (sheet names compare without case). */
export function chartLinksPivot(chart: { pivot?: { name: string } }, sheetName: string, pivotName: string): boolean {
  if (!chart.pivot) return false
  const ref = pivotSourceRef(chart.pivot.name)
  return ref.pivot === pivotName && (ref.sheet === undefined || ref.sheet.toLowerCase() === sheetName.toLowerCase())
}
/** Two pivot charts name the same pivot (sheet and name). */
export function samePivotSource(a: { pivot?: { name: string } }, b: { pivot?: { name: string } }): boolean {
  const x = pivotSourceRef(a.pivot?.name ?? ''), y = pivotSourceRef(b.pivot?.name ?? '')
  return x.pivot === y.pivot && (x.sheet ?? '').toLowerCase() === (y.sheet ?? '').toLowerCase()
}
