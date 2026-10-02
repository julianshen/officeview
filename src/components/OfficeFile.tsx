/**
 * <OfficeFile> — convenience wrapper: loads an office file from bytes,
 * detects the format from [Content_Types].xml, parses it, and hands the
 * document model to <OfficeDoc>. Accepts a promise for async usage.
 */
import { useEffect, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { OfficePackage } from '../core/zip'
import { parseDocx } from '../docx/parse'
import type { DocxDocument } from '../docx/types'
import { parseXlsx } from '../xlsx/parse'
import type { XlsxDocument } from '../xlsx/types'
import { parsePptx } from '../pptx/parse'
import type { PptxDocument } from '../pptx/types'
import { OfficeDoc, type OfficeDocProps } from './OfficeDoc'
import { odfKind } from '../odf/container'
import { parseOdt } from '../odt/parse'
import { readSource, type ByteSource, type Progress, type ProtectionListener, type ProtectionPolicy } from '../core/stream'
import { parseRtf, isRtf } from '../rtf/parse'

export type AnyDoc = DocxDocument | XlsxDocument | PptxDocument

export interface LoadProgress extends Progress {
  status: 'loading' | 'done' | 'error'
  message?: string
}

export type OfficeFileState =
  | { status: 'loading' }
  | { status: 'error'; message: string }
  | { status: 'ready'; document: AnyDoc }

async function detectAndParse(pkg: OfficePackage): Promise<AnyDoc> {
  const ct = await pkg.xml('[Content_Types].xml')
  const ctText = JSON.stringify(ct ?? {})
  if (ctText.includes('wordprocessingml.document.main')) return parseDocx(pkg)
  if (ctText.includes('spreadsheetml.sheet.main')) return parseXlsx(pkg)
  if (ctText.includes('presentationml.presentation.main')) return parsePptx(pkg)
  // fall back to probing known parts
  if (pkg.has('word/document.xml')) return parseDocx(pkg)
  if (pkg.has('xl/workbook.xml')) return parseXlsx(pkg)
  if (pkg.has('ppt/presentation.xml')) return parsePptx(pkg)
  // OpenDocument (Writer first — Calc/Impress route here when they land)
  const kind = await odfKind(pkg)
  if (kind === 'text') return parseOdt(pkg)
  if (kind) throw new Error(`Unsupported OpenDocument type '${kind}' — only word-processing (.odt) documents render for now`)
  throw new Error('Unrecognized office file: no word/xl/ppt content detected')
}

/**
 * Load an office file from any byte source — bytes in memory, a Blob, a
 * fetch() Response, a ReadableStream, a protected URL ({ url, headers }), a
 * Request, or a promise of any of these (e.g. fetch() itself). `onProgress`
 * reports incremental byte counts while the source is still arriving.
 */
export async function loadOfficeFile(
  source: ByteSource,
  options: { onProgress?: (p: Progress) => void; onProtection?: ProtectionListener } = {},
): Promise<AnyDoc> {
const data = await readSource(source, options.onProgress, options.onProtection)
  // RTF is plain text, not a zip, so it must be recognised before the bytes
  // reach OfficePackage.load, which would throw on them.
  if (isRtf(data)) return parseRtf(data)
  return OfficePackage.load(data).then(detectAndParse)
}

export function useOfficeFile(
  source: ByteSource | null | undefined,
  options: { onProgress?: (p: Progress) => void; onProtection?: ProtectionListener } = {},
): OfficeFileState {
  const [state, setState] = useState<OfficeFileState>({ status: 'loading' })
  const progressRef = useRef(options.onProgress)
  progressRef.current = options.onProgress
  const protectionRef = useRef(options.onProtection)
  protectionRef.current = options.onProtection
  useEffect(() => {
    let cancelled = false
    if (!source) return
    setState({ status: 'loading' })
    loadOfficeFile(source, {
      onProgress: (p) => progressRef.current?.(p),
      // protection events from a cancelled load are dropped: the ref below
      // always points at the latest render's callback, which would otherwise
      // tag the old policy with the current data (see the source-tagged state
      // in <OfficeFile>) and let a stale download overwrite the live policy.
      onProtection: (p) => { if (!cancelled) protectionRef.current?.(p) },
    })
      .then((document) => {
        if (!cancelled) setState({ status: 'ready', document })
      })
      .catch((e: unknown) => {
        if (!cancelled) setState({ status: 'error', message: e instanceof Error ? e.message : String(e) })
      })
    return () => {
      cancelled = true
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
  const state = useOfficeFile(data, {
    onProgress: (p) => setProgress(p),
    onProtection: (policy) => setHeaderPolicy({ source: data, policy }),
  })
  const serverPolicy = headerPolicy && headerPolicy.source === data ? headerPolicy.policy : undefined
  const effectiveAllowCopy = allowCopy && (serverPolicy?.allowCopy ?? true)
  const effectiveAllowPrint = allowPrint && (serverPolicy?.allowPrint ?? true)
  if (state.status === 'loading') {
    if (typeof loading === 'function') return loading(progress ?? { loaded: 0 })
    return <>{loading}</>
  }
  if (state.status === 'error') return error ? error(state.message) : <div role="alert">{state.message}</div>
  return <OfficeDoc document={state.document} allowCopy={effectiveAllowCopy} allowPrint={effectiveAllowPrint} {...rest} />
}

export default OfficeFile
