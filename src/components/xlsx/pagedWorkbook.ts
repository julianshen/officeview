/**
 * Opens a paged spreadsheet admission in the resident viewer worker (plan
 * S5.4), loaded on demand by <OfficeFile interactiveSheets viewerWorker>. The
 * main thread never parses the package: the worker re-runs admission over the
 * Blob, imports pages into its storage, and this side receives the sheet list
 * and compact layouts. The returned shell XlsxDocument carries sheet names only
 * and marks its worker-backed source (pagedDocument.ts).
 */
import type { ProfileId } from '../../xlsx/view/contracts'
import type { XlsxDocument } from '../../xlsx/types'
import { ViewerClient, type ViewerPort } from '../../worker/viewerClient'
import { RemoteWorkbook, openRemoteWorkbook } from '../../worker/viewerRemoteSource'
import { ImportRefusal } from '../../xlsx/storage/importSession'
import type { PackageErrorCode } from '../../core/package-error'
import { PAGED_WORKBOOK, type PagedWorkbookHandle } from '../pagedDocument'
import { semanticsFor } from '../../xlsx/view/semanticsFor'

/** Open refusals that name a real limit; any other failure keeps compatibility mode. */
const REFUSALS: ReadonlySet<string> = new Set(['storage-error', 'quota', 'resource-limit'])
let serial = 0

export interface PagedOpenOptions {
  port: ViewerPort & { terminate?(): void }
  profile: ProfileId
  signal: AbortSignal
  /** Releases the caller's spooled copy of the source, once the worker is done with it. */
  release?: () => Promise<void>
}

/** The shell document, or undefined when the worker cannot serve this file
 * (the caller then keeps today's compatibility parse). */
export async function openPagedWorkbook(blob: Blob, options: PagedOpenOptions): Promise<XlsxDocument | undefined> {
  const client = new ViewerClient(options.port, 'paged')
  const docId = `paged-${++serial}-${Date.now().toString(36)}`
  let ended = false
  // A refused open leaves the source with the caller (it falls back to the
  // compatibility parse); an opened document releases it when it closes.
  const end = async (releaseSource: boolean) => {
    if (ended) return
    ended = true
    client.dispose()
    options.port.terminate?.()
    if (releaseSource) await options.release?.()
  }
  // Cancellation (unmount, replacement) and a dead worker end the open at once:
  // close the half-open document, terminate the worker, never wait for a reply
  // that may not come (§10.3).
  const stop = new Promise<'aborted' | 'failed'>(resolve => {
    if (options.signal.aborted) { resolve('aborted'); return }
    options.signal.addEventListener('abort', () => resolve('aborted'), { once: true })
    ;(options.port as { addEventListener?(type: 'error', listener: () => void): void }).addEventListener?.('error', () => resolve('failed'))
  })
  const outcome = await Promise.race([
    openRemoteWorkbook(client, docId, blob, options.profile).then(value => ({ value }), (error: unknown) => ({ error })),
    stop,
  ])
  if (outcome === 'aborted' || outcome === 'failed') {
    try { options.port.postMessage({ kind: 'close', requestId: client.nextRequestId(), docId }, []) } catch { /* worker gone */ }
    await end(false)
    return undefined
  }
  if ('error' in outcome) {
    await end(false)
    throw outcome.error
  }
  const opened = outcome.value
  if (!('layouts' in opened)) {
    await end(false)
    const code = opened.kind === 'error' ? opened.code : ''
    if (REFUSALS.has(code)) throw new ImportRefusal(code as PackageErrorCode, 'limit', opened.diagnostics[0]?.message ?? `This spreadsheet could not be opened (${code}).`)
    return undefined
  }
  const remote = new RemoteWorkbook(client, opened, () => end(true))
  if (options.signal.aborted) { await remote.close(); return undefined }
  const doc: XlsxDocument = {
    sheets: opened.sheets.map((sheet, index) => ({ name: sheet.name, sheetId: sheet.sheetId, workbookIndex: index, rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(opened.formatting?.dateSystem ?? '1900'),
      // The indent unit of the main-thread layout host (render.ts layoutCellBody).
      ...(opened.formatting?.maxDigitWidthPx !== undefined ? { maxDigitWidthPx: opened.formatting.maxDigitWidthPx } : {}) })),
    ...(opened.valuePolicy === 'cached' ? { valuePolicy: 'cached' as const } : {}),
  }
  const handle: PagedWorkbookHandle = { source: remote, release: () => remote.close() }
  Object.defineProperty(doc, PAGED_WORKBOOK, { value: handle })
  return doc
}
