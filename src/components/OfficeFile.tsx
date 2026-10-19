/**
 * <OfficeFile> — convenience wrapper: loads an office file from bytes,
 * detects the format from [Content_Types].xml, parses it, and hands the
 * document model to <OfficeDoc>. Accepts a promise for async usage.
 */
import { useEffect, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { OfficeDoc, type OfficeDocProps } from './OfficeDoc'
import type { ByteSource, Progress, ProtectionListener, ProtectionPolicy, WatermarkListener } from '../core/stream'
import { mergeWatermark, type WatermarkOptions } from '../core/watermark'
import type { PackageErrorCode } from '../core/package-error'
import { documentNotices, loadOfficeSource, type AnyDoc } from './loadOfficeSource'
import { ImportRefusal, type RefusalReason } from '../xlsx/storage/importSession'
import { pagedWorkbookOf } from './pagedDocument'

export type { AnyDoc } from './loadOfficeSource'

export interface LoadProgress extends Progress {
  status: 'loading' | 'done' | 'error'
  message?: string
}

export type OfficeFileState =
  | { status: 'loading' }
  /** `code`/`reason` are set when a spreadsheet was refused by admission
   * (`reason: 'limit'` names the applied limit; `'paged-not-yet'` is no longer produced). */
  | { status: 'error'; message: string; code?: PackageErrorCode; reason?: RefusalReason }
  /** `notices`: user-visible load notices (e.g. saved values shown, data not shown). */
  | { status: 'ready'; document: AnyDoc; notices?: readonly string[] }

/**
 * Load an office file from any byte source — bytes in memory, a Blob, a
 * fetch() Response, a ReadableStream, a protected URL ({ url, headers }), a
 * Request, or a promise of any of these (e.g. fetch() itself). `onProgress`
 * reports incremental byte counts while the source is still arriving.
 */
export async function loadOfficeFile(
  source: ByteSource,
  options: {
    onProgress?: (p: Progress) => void
    onProtection?: ProtectionListener
    onWatermark?: WatermarkListener
  } = {},
): Promise<AnyDoc> {
  // Spreadsheets go through admission (object / paged / refuse); every
  // other format takes today's path unchanged. A spreadsheet that cannot load
  // as an object document rejects with a typed ImportRefusal naming the limit.
  return loadOfficeSource(source, options)
}

/** A dedicated worker (or MessagePort) running officeview's dist/worker.js. */
export interface ViewerWorkerPort {
  postMessage(message: unknown, transfer: Transferable[]): void
  addEventListener(type: 'message', listener: (event: MessageEvent) => void): void
  removeEventListener(type: 'message', listener: (event: MessageEvent) => void): void
  terminate?(): void
}

export function useOfficeFile(
  source: ByteSource | null | undefined,
  options: {
    onProgress?: (p: Progress) => void
    onProtection?: ProtectionListener
    onWatermark?: WatermarkListener
    /** Response headers of the download (dropped once the load is replaced). */
    onHeaders?: (headers: Headers, url: string) => void
    /** With `viewerWorker`, a large (paged) spreadsheet opens in that worker for the interactive viewer. */
    interactiveSheets?: boolean
    viewerWorker?: () => ViewerWorkerPort
  } = {},
): OfficeFileState {
  const [state, setState] = useState<OfficeFileState>({ status: 'loading' })
  const progressRef = useRef(options.onProgress)
  progressRef.current = options.onProgress
  const protectionRef = useRef(options.onProtection)
  protectionRef.current = options.onProtection
  const watermarkRef = useRef(options.onWatermark)
  watermarkRef.current = options.onWatermark
  const headersRef = useRef(options.onHeaders)
  headersRef.current = options.onHeaders
  const workerRef = useRef(options.viewerWorker)
  workerRef.current = options.viewerWorker
  // Paged admissions route to the worker only with the flag AND a worker;
  // otherwise the load is exactly today's (compatibility mode for large files).
  const pagedRoute = Boolean(options.interactiveSheets && options.viewerWorker)
  const routeRef = useRef(pagedRoute)
  routeRef.current = pagedRoute
  // A one-shot source (Response, stream, Request, promise) is consumed by its
  // first load, so a later mode change keeps the loaded document (a worker-backed
  // one stays in the interactive viewer) instead of re-reading a drained body;
  // re-readable sources (bytes, Blob, URL) re-open under the new route.
  const routeKey = source && oneShot(source) ? null : pagedRoute
  useEffect(() => {
    let cancelled = false
    // No source: the previous cleanup released its document, so stop rendering it.
    if (!source) { setState({ status: 'loading' }); return }
    let loaded: AnyDoc | undefined
    // One controller per source: cleanup aborts the download, the spool,
    // the verification pass and the parse, and the loader releases the spool.
    const controller = new AbortController()
    setState({ status: 'loading' })
    loadOfficeSource(source, {
      signal: controller.signal,
      onProgress: (p) => progressRef.current?.(p),
      // protection events from a cancelled load are dropped: the ref below
      // always points at the latest render's callback, which would otherwise
      // tag the old policy with the current data (see the source-tagged state
      // in <OfficeFile>) and let a stale download overwrite the live policy.
      onProtection: (p) => { if (!cancelled) protectionRef.current?.(p) },
      onWatermark: (m) => { if (!cancelled) watermarkRef.current?.(m) },
      onHeaders: (h, url) => { if (!cancelled) headersRef.current?.(h, url) },
      ...(routeRef.current ? {
        paged: async (blob, ctx, release) => {
          const factory = workerRef.current
          if (!factory) return undefined
          // A factory that cannot construct a worker (no Worker global, CSP)
          // is the same outcome as a dead worker: keep compatibility mode.
          let port: ViewerWorkerPort
          try { port = factory() } catch { return undefined }
          const { openPagedWorkbook } = await import('./xlsx/pagedWorkbook')
          return openPagedWorkbook(blob, { port, profile: ctx.profile, signal: ctx.signal, ...(release ? { release } : {}) })
        },
      } : {}),
    })
      .then((document) => {
        // A worker-backed document belongs to this load: released with it.
        if (cancelled) { void pagedWorkbookOf(document)?.release(); return }
        loaded = document
        const notices = documentNotices(document)
        setState(notices.length > 0 ? { status: 'ready', document, notices } : { status: 'ready', document })
      })
      .catch((e: unknown) => {
        if (cancelled) return
        const message = e instanceof Error ? e.message : String(e)
        setState(e instanceof ImportRefusal ? { status: 'error', message, code: e.code, reason: e.reason } : { status: 'error', message })
      })
    return () => {
      cancelled = true
      controller.abort()
      void pagedWorkbookOf(loaded)?.release()
    }
  }, [source, routeKey])
  return state
}

const oneShot = (s: ByteSource): boolean =>
  typeof (s as { then?: unknown }).then === 'function'
  || (typeof Request !== 'undefined' && s instanceof Request)
  || (typeof Response !== 'undefined' && s instanceof Response)
  || (typeof ReadableStream !== 'undefined' && s instanceof ReadableStream)

export interface OfficeFileProps extends Omit<OfficeDocProps, 'document'> {
  /** The office file: bytes, a Blob, a fetch() Response, a ReadableStream, a URL ({ url, headers }), a Request, or a promise of any of these. */
  data: ByteSource | null | undefined
  /**
   * Rendered while loading. A function form receives live progress so the
   * host can show a real bar: loading={({ loaded, total }) => <Bar pct={...} />}.
   */
  loading?: ReactElement | ((progress: Progress) => ReactElement) | null
  /** Rendered on parse error. */
  error?: (message: string) => ReactElement
  style?: CSSProperties
  className?: string
  /**
   * Resident viewer worker for large spreadsheets with `interactiveSheets`
   * (experimental), e.g. `() => new Worker(new URL('officeview/dist/worker.js',
   * import.meta.url), { type: 'module' })`. A file admitted as large then
   * opens in that worker: pages stay there and the interactive viewer reads
   * bounded windows. Without it, large files keep compatibility mode. Each
   * large document gets its own worker, terminated when the document closes.
   */
  viewerWorker?: () => ViewerWorkerPort
}

export function OfficeFile({
  data,
  loading,
  error,
  allowCopy = true,
  allowPrint = true,
  watermark,
  viewerWorker,
  ...rest
}: OfficeFileProps): ReactElement {
  const [progress, setProgress] = useState<Progress | undefined>(undefined)
  // Server policy, sealed at download time. Stale loads are dropped at the
  // source (useOfficeFile ignores protection events after cleanup); the tag
  // below is a second line of defence so a policy can only ever apply to the
  // source it arrived with. Policy lives in state — not props — so there is
  // no boolean for devtools to flip. The merge below is deliberately
  // most-restrictive-wins: a server denial cannot be re-enabled from props.
  const [headerPolicy, setHeaderPolicy] = useState<{ source: ByteSource | null | undefined; policy: ProtectionPolicy }>()
  // Same sealing rule for the watermark: it belongs to the source it arrived
  // with, so a stale download can never stamp the next document.
  const [headerMark, setHeaderMark] = useState<{
    source: ByteSource | null | undefined
    mark: WatermarkOptions | undefined
  }>()
  // File name (spec §6.5), sealed to its source like the policy above: the
  // response's Content-Disposition or URL, or the File a promised source
  // resolves to (which reports no headers). Headers win whichever lands
  // first. The parser is lazy (out of the eager entry); a name that resolves
  // after its source was replaced is dropped so it never overwrites the live one.
  const [named, setNamed] = useState<{ source: ByteSource | null | undefined; name: string | undefined }>()
  const current = useRef(data)
  current.current = data
  const name = (from: unknown, disposition?: string | null, keep?: boolean) => {
    const source = data
    void import('./fileName').then(({ fileNameFrom }) => current.current === source && setNamed((n) => keep && n?.source === source ? n : { source, name: fileNameFrom(from, disposition) }), () => {})
  }
  const state = useOfficeFile(data, {
    onProgress: (p) => setProgress(p),
    onProtection: (policy) => setHeaderPolicy({ source: data, policy }),
    onWatermark: (mark) => setHeaderMark({ source: data, mark }),
    onHeaders: (h, url) => name({ url }, h.get('content-disposition')),
    interactiveSheets: rest.interactiveSheets,
    viewerWorker,
  })
  useEffect(() => { Promise.resolve(data).then((s) => name(s, null, true), () => {}) }, [data])
  const serverPolicy = headerPolicy && headerPolicy.source === data ? headerPolicy.policy : undefined
  const effectiveAllowCopy = allowCopy && (serverPolicy?.allowCopy ?? true)
  const effectiveAllowPrint = allowPrint && (serverPolicy?.allowPrint ?? true)
  // A watermark sent by the server wins over the prop: the server is the
  // authority on what a document is stamped with.
  const serverMark = headerMark && headerMark.source === data ? headerMark.mark : undefined
  const effectiveWatermark = mergeWatermark(watermark, serverMark)
  const title = rest.title ?? (named?.source === data ? named?.name : undefined)
  if (state.status === 'loading') {
    if (typeof loading === 'function') return loading(progress ?? { loaded: 0 })
    return <>{loading}</>
  }
  if (state.status === 'error') return error ? error(state.message) : <div role="alert">{state.message}</div>
  const doc = (
    <OfficeDoc
      document={state.document}
      allowCopy={effectiveAllowCopy}
      allowPrint={effectiveAllowPrint}
      {...(effectiveWatermark ? { watermark: effectiveWatermark } : {})}
      {...rest}
      {...(title ? { title } : {})}
    />
  )
  if (!state.notices || ((rest.interactiveSheets || pagedWorkbookOf(state.document)) && 'sheets' in state.document)) return doc
  return (
    <>
      <div role="status" data-officeview-notice="">{state.notices.join(' · ')}</div>
      {doc}
    </>
  )
}

export default OfficeFile
