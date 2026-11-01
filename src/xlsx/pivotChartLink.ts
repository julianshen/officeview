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

export interface PivotKey { sheet: string; name?: string }
/**
 * The charts linked to `pivot` among `charts` (Ruling S3-Q). A sheet-qualified source links its sheet's pivot; a source
 * without a sheet links the pivot only when no other sheet has a pivot of that name (an unnamed pivot counts as
 * possibly that name). Otherwise it is `ambiguous`: it links to none, and every pivot it could name refuses
 * `pivot-chart-mapping` (fail closed; a chart rebuilt from the wrong pivot would be a wrong picture).
 */
export function linkPivotCharts<C extends { pivot?: { name: string } }>(charts: readonly C[], pivot: { sheet: string; name: string }, pivots: readonly PivotKey[]): { charts: C[]; ambiguous?: C } {
  const linked: C[] = []
  for (const chart of charts) {
    if (!chart.pivot) continue
    const ref = pivotSourceRef(chart.pivot.name)
    if (ref.pivot !== pivot.name) continue
    if (ref.sheet !== undefined) { if (ref.sheet.toLowerCase() === pivot.sheet.toLowerCase()) linked.push(chart); continue }
    const candidates = pivots.filter(entry => entry.name === undefined || entry.name === ref.pivot)
    if (candidates.length > 1) return { charts: [], ambiguous: chart }
    linked.push(chart)
  }
  return { charts: linked }
}
