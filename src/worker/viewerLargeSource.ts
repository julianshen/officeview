/**
 * Blob opens in the resident viewer worker (plan S1.2): admission decides
 * object or paged before anything large is read. A paged decision opens a
 * progressive import into SourcePageStore pages and returns the paged
 * WorkbookSource; the object path (and any file the paged path cannot serve
 * yet) keeps today's parse. Loaded on demand by the kernel.
 */
import type { ProfileId, WorkerDiagnostic } from '../xlsx/view/contracts'
import { blobRangeSource } from '../core/zip-stream'
import { admitSpreadsheet } from '../xlsx/storage/admission'
import { importLimitsFor, type ImportLimitOverrides } from '../xlsx/storage/importSession'
import { compatibilityFallbackAllowed, firstPageOrOutcome, openImportSession } from '../xlsx/storage/pagedImport'
import type { PageStorageEnv } from '../xlsx/storage/sourcePages'
import { createPagedWorkbookSource } from '../xlsx/view/pagedSource'
import type { WorkbookSource } from './viewerSession'

export interface LargeOpenOptions {
  profile: ProfileId
  signal: AbortSignal
  limits?: ImportLimitOverrides
  storage?: PageStorageEnv
}
export type LargeOpen =
  | { kind: 'paged'; source: WorkbookSource }
  /** The object path decides (admitted object, or a file the paged path cannot serve yet). */
  | { kind: 'object'; diagnostics: WorkerDiagnostic[] }
  | { kind: 'refuse'; code: string; message: string }

export const COMPATIBILITY_DIAGNOSTIC = 'large-file-compatibility'
export const PAGED_UNSUPPORTED_DIAGNOSTIC = 'paged-unsupported'

export async function openLargeBlob(blob: Blob, docId: string, options: LargeOpenOptions): Promise<LargeOpen> {
  const limits = importLimitsFor(options.profile, options.limits)
  const admission = await admitSpreadsheet(blobRangeSource(blob), options.profile, limits, options.signal)
  if (admission.mode === 'refuse') return { kind: 'refuse', code: admission.failure.code, message: admission.failure.message }
  if (admission.mode === 'object') return { kind: 'object', diagnostics: [] }
  if (admission.mode === 'legacy-fallback') return { kind: 'object', diagnostics: [{ code: COMPATIBILITY_DIAGNOSTIC, message: `Opened through the compatibility path (${admission.reason})` }] }
  const fallback = compatibilityFallbackAllowed(blob.size, limits)
  let session
  try {
    session = await openImportSession({ kind: 'blob', blob }, { docId, profile: options.profile, signal: options.signal, limits: options.limits, storage: options.storage })
  } catch (error) {
    const code = (error as { code?: unknown } | null)?.code
    // Storage/quota refusals are actionable and never fall back to unbounded memory.
    if (code === 'storage-error' || code === 'quota' || code === 'cancelled') throw error
    if (!fallback) return { kind: 'refuse', code: typeof code === 'string' ? code : 'parse-error', message: (error as Error).message }
    return { kind: 'object', diagnostics: [compat(error)] }
  }
  // A file whose first sheet the paged reader refuses before any page landed
  // (e.g. rows without numbers) opens through the compatibility path, but only
  // within the object caps; beyond them it stays a typed incomplete source.
  const early = await firstPageOrOutcome(session)
  const refusedEarly = early !== undefined && early.kind === 'incomplete' && early.failure.code === 'unsupported' && session.workbook.store.stats.committed === 0
  if (refusedEarly && fallback) {
    await session.release()
    return { kind: 'object', diagnostics: [compat(early.failure)] }
  }
  const source = createPagedWorkbookSource(session)
  if (refusedEarly) {
    return { kind: 'paged', source: { ...source, openDiagnostics: [{ code: PAGED_UNSUPPORTED_DIAGNOSTIC, message: `This large file uses a layout the large-file reader does not support yet (${early.failure.message})` }] } }
  }
  return { kind: 'paged', source }
}

function compat(error: unknown): WorkerDiagnostic {
  const message = error instanceof Error || (typeof error === 'object' && error !== null && 'message' in error) ? String((error as { message: unknown }).message) : String(error)
  return { code: COMPATIBILITY_DIAGNOSTIC, message: `Large file — opened in compatibility mode (${message})` }
}
