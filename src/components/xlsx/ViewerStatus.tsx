import { useEffect, useRef, useState, type ReactElement } from 'react'
import type { CellRead, Completeness, SourceStore } from '../../xlsx/view/contracts'
import { cellId, formatA1, MAX_ROWS, type SourceCell } from '../../xlsx/view/coordinates'
import { boundedSelectionText } from '../../xlsx/view/selection'

export interface HiddenCellInspection {
  target: SourceCell
  hiddenReason: string
  source: Pick<SourceStore, 'stamp' | 'readCell'>
}
export interface ViewerStatusProps {
  records: { visible: number; total: number; completeness?: Completeness }
  announcement?: string
  inspection?: HiddenCellInspection
  countMode?: 'records' | 'body'
  loadReport?: { status: 'ready'; notices?: readonly string[] } | { status: 'error'; message: string }
}
function inspectionText(read: CellRead | null): string {
  if (!read || read.kind === 'unknown') return 'Loading cell…'
  if (read.kind === 'blank') return 'Blank cell'
  if (read.kind === 'error') return `Cell unavailable (${read.code}); ${read.retryable ? 'retry available' : 'cannot retry'}`
  const formula = read.cell.formula ? `; formula ${boundedSelectionText(read.cell.formula)}` : ''
  return read.kind === 'unavailable' ? `Formula value unavailable${formula}` : `Value ${boundedSelectionText(String(read.cell.value ?? ''))}${formula}`
}
export function ViewerStatus({ records, announcement, inspection, countMode, loadReport }: ViewerStatusProps): ReactElement {
  const validCounts = Number.isSafeInteger(records?.total) && records.total >= 0 && records.total <= MAX_ROWS && Number.isSafeInteger(records.visible) && records.visible >= 0 && records.visible <= records.total
  const source = inspection?.source, row = inspection?.target.row, col = inspection?.target.col
  const owner = source?.stamp
  const sourceRef = useRef(source)
  sourceRef.current = source
  const key = inspection ? JSON.stringify([owner?.docId, owner?.sheetId, owner?.valueGen, row, col]) : ''
  const [snapshot, setSnapshot] = useState<{ key: string; read: CellRead } | null>(null)
  useEffect(() => {
    const port = sourceRef.current
    if (!port || row === undefined || col === undefined) return
    const controller = new AbortController(), stamp = { ...port.stamp }, target = { row, col }
    let alive = true
    void Promise.resolve().then(() => alive && !controller.signal.aborted ? port.readCell(target, { stamp, signal: controller.signal }) : null).then(read => {
      if (!read || !alive || controller.signal.aborted) return
      const current = sourceRef.current?.stamp
      if (!current || current.docId !== stamp.docId || current.sheetId !== stamp.sheetId || current.valueGen !== stamp.valueGen) {
        read = { kind: 'error', code: 'stale-source', retryable: true }
      } else if ((read.kind === 'cell' || read.kind === 'unavailable') && (read.cell.row !== row || read.cell.col !== col || read.cell.id !== cellId(row, col))) {
        read = { kind: 'error', code: 'page-error', retryable: false }
      }
      setSnapshot({ key, read })
    }).catch(() => { if (alive && !controller.signal.aborted) setSnapshot({ key, read: { kind: 'error', code: 'page-error', retryable: true } }) })
    return () => { alive = false; controller.abort() }
  }, [key])
  const counts = !validCounts ? 'Record counts unavailable' : countMode === 'body' ? `Visible body ${records.visible} / ${records.completeness === 'complete' ? 'complete body' : 'body'} ${records.total}` : records.total === 0 ? '0 records' : `${records.visible} of ${records.total} records`
  return <div role="status" aria-live="polite" aria-atomic="true" data-officeview-ui="viewer-status">
    <span>{counts}{validCounts && records.completeness && records.completeness !== 'complete' ? ` (${records.completeness})` : ''}</span>
    {(loadReport?.status === 'error' ? [loadReport.message] : loadReport?.notices ?? []).map((message, index) => <span key={index}> — {message}</span>)}
    {announcement && <span> — {announcement}</span>}
    {inspection && <span> — Original cell {formatA1(inspection.target)}; {inspection.hiddenReason}; read-only. {inspectionText(snapshot?.key === key ? snapshot.read : null)}</span>}
  </div>
}
