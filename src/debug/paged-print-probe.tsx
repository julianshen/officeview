/** Dev-only mounted peers for the paged print browser regressions. */
import { createRoot } from 'react-dom/client'
import { OfficeDoc } from '../components/OfficeDoc'
import { loadOfficeSource } from '../components/loadOfficeSource'
import type { XlsxDocument } from '../xlsx/types'

export async function mountPrintPeers(source: Blob): Promise<() => void> {
  const document = await loadOfficeSource(source)
  const host = window.document.createElement('div')
  host.dataset.printPeers = ''
  window.document.body.append(host)
  const root = createRoot(host)
  const sheet: XlsxDocument = {sheets: [{name: 'Sheet', rows: [], cols: [], merges: [], mergeRanges: []}]}
  root.render(<>
    <div data-peer="denied"><OfficeDoc document={document} allowPrint={false} /></div>
    <div data-peer="allowed"><OfficeDoc document={document} /></div>
    <div data-peer="sheet"><OfficeDoc document={sheet} interactiveSheets /></div>
  </>)
  return () => { root.unmount(); host.remove() }
}
