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

export function useOfficeFile(
  source: ByteSource | null | undefined,
  options: {
    onProgress?: (p: Progress) => void
    onProtection?: ProtectionListener
    onWatermark?: WatermarkListener
  } = {},
): OfficeFileState {
  const [state, setState] = useState<OfficeFileState>({ status: 'loading' })
  const progressRef = useRef(options.onProgress)
  progressRef.current = options.onProgress
  const protectionRef = useRef(options.onProtection)
  protectionRef.current = options.onProtection
  const watermarkRef = useRef(options.onWatermark)
  watermarkRef.current = options.onWatermark
  useEffect(() => {
    let cancelled = false
    if (!source) return
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
    })
      .then((document) => {
        if (cancelled) return
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
    }
  }, [source])
  return state
}

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
}

export function OfficeFile({
  data,
  loading,
  error,
  allowCopy = true,
  allowPrint = true,
  watermark,
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
  const state = useOfficeFile(data, {
    onProgress: (p) => setProgress(p),
    onProtection: (policy) => setHeaderPolicy({ source: data, policy }),
    onWatermark: (mark) => setHeaderMark({ source: data, mark }),
  })
  const serverPolicy = headerPolicy && headerPolicy.source === data ? headerPolicy.policy : undefined
  const effectiveAllowCopy = allowCopy && (serverPolicy?.allowCopy ?? true)
  const effectiveAllowPrint = allowPrint && (serverPolicy?.allowPrint ?? true)
  // A watermark sent by the server wins over the prop: the server is the
  // authority on what a document is stamped with.
  const serverMark = headerMark && headerMark.source === data ? headerMark.mark : undefined
  const effectiveWatermark = mergeWatermark(watermark, serverMark)
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
    />
  )
  if (!state.notices || (rest.interactiveSheets && 'sheets' in state.document)) return doc
  return (
    <>
      <div role="status" data-officeview-notice="">{state.notices.join(' · ')}</div>
      {doc}
    </>
  )
}

export default OfficeFile
