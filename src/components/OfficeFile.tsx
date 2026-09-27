/**
 * <OfficeFile> — convenience wrapper: loads an office file from bytes,
 * detects the format from [Content_Types].xml, parses it, and hands the
 * document model to <OfficeDoc>. Accepts a promise for async usage.
 */
import { useEffect, useState, type CSSProperties, type ReactElement } from 'react'
import { OfficePackage } from '../core/zip'
import { parseDocx } from '../docx/parse'
import type { DocxDocument } from '../docx/types'
import { parseXlsx } from '../xlsx/parse'
import type { XlsxDocument } from '../xlsx/types'
import { parsePptx } from '../pptx/parse'
import type { PptxDocument } from '../pptx/types'
import { OfficeDoc, type OfficeDocProps } from './OfficeDoc'

export type AnyDoc = DocxDocument | XlsxDocument | PptxDocument

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
  throw new Error('Unrecognized office file: no word/xl/ppt content detected')
}

export function loadOfficeFile(data: ArrayBuffer | Uint8Array): Promise<AnyDoc> {
  return OfficePackage.load(data).then(detectAndParse)
}

export function useOfficeFile(data: ArrayBuffer | Uint8Array | null | undefined): OfficeFileState {
  const [state, setState] = useState<OfficeFileState>({ status: 'loading' })
  useEffect(() => {
    let cancelled = false
    if (!data) return
    setState({ status: 'loading' })
    loadOfficeFile(data)
      .then((document) => {
        if (!cancelled) setState({ status: 'ready', document })
      })
      .catch((e: unknown) => {
        if (!cancelled) setState({ status: 'error', message: e instanceof Error ? e.message : String(e) })
      })
    return () => {
      cancelled = true
    }
  }, [data])
  return state
}

export interface OfficeFileProps extends Omit<OfficeDocProps, 'document'> {
  /** Raw bytes of the office file. */
  data: ArrayBuffer | Uint8Array | null | undefined
  /** Rendered while loading. */
  loading?: ReactElement | null
  /** Rendered on parse error. */
  error?: (message: string) => ReactElement
  style?: CSSProperties
  className?: string
}

export function OfficeFile({ data, loading, error, ...rest }: OfficeFileProps): ReactElement {
  const state = useOfficeFile(data)
  if (state.status === 'loading') return <>{loading}</>
  if (state.status === 'error') return error ? error(state.message) : <div role="alert">{state.message}</div>
  return <OfficeDoc document={state.document} {...rest} />
}

export default OfficeFile
