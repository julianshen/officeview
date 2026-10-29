/**
 * Large-source import: limits, the page port and the admission outcome used
 * by the entry points (design note §4.3-§4.5).
 *
 * S1.1b owns admission (object / paged / refuse) and the typed refusal; the
 * paged importer (S1.2) lives in ./pagedImport and commits pages to a SourcePageStore.
 */
import type { Completeness, ProfileId, SourceFailure, SourceStamp, WorkerDiagnostic } from '../view/contracts'
import type { PackageErrorCode } from '../../core/package-error'

const KiB = 1024
const MiB = 1024 * KiB
const GiB = 1024 * MiB

/** All §4.3 numbers in one place; defaults per ProfileId (proposed, frozen in S10). */
export interface ImportLimits {
  // input and queues (spec §4.3.4)
  rangeReadBytes: number
  inflateOutputBytes: number
  maxQueuedExpandedBytes: number
  targetPageBytes: number
  inlineStringSoftBytes: number
  maxTokenChars: number
  // gate 1 (declared) and gate 2 (actual): object path
  object: {
    budgetBytes: number
    compressedBytes: number
    nonDrawingExpandedBytes: number
    nonDrawingPartBytes: number
    drawingPartBytes: number
    rows: number
    cells: number
    uniqueStrings: number
    merges: number
    styleRecords: number
    maxAttributesPerElement: number
    entries: number
    maxRatio: number
    envelope: { constantBytes: number; perElement: number; perAttribute: number; perTextChar: number }
    evaluation: { perFormulaBytes: number; perMaterializedCellBytes: number; maxArrayCells: number }
  }
  // paged path (otherwise refuse)
  paged: {
    compressedBytes: number
    expandedBytes: number
    cells: number
    uniqueStrings: number
    sharedStringsBytes: number
    entries: number
    centralDirectoryBytes: number
    maxRatio: number
    drawingPartEnvelopeBytes: number
    minFreeStorageBytes: number
    /** Per sheet: merges, table parts, drawing references and array formulas each (S1.2). */
    structuralRecords: number
    /** Per pivot table, pivot cache, slicer or timeline part the worker reads (Ruling AP: the drawing-part cap). */
    pivotPartBytes: number
  }
  spool: { memoryLimitBytes: number; hardMemoryCapBytes: number; sniffPrefixBytes: number }
}

const DESKTOP: ImportLimits = {
  rangeReadBytes: 262_144,
  inflateOutputBytes: 65_536,
  maxQueuedExpandedBytes: 8 * MiB,
  targetPageBytes: 4 * MiB,
  inlineStringSoftBytes: 64 * KiB,
  maxTokenChars: 1 * MiB,
  object: {
    budgetBytes: 256 * MiB,
    compressedBytes: 32 * MiB,
    nonDrawingExpandedBytes: 6 * MiB,
    nonDrawingPartBytes: 4 * MiB,
    drawingPartBytes: 8 * MiB,
    rows: 100_000,
    cells: 100_000,
    uniqueStrings: 50_000,
    merges: 20_000,
    styleRecords: 20_000,
    maxAttributesPerElement: 128,
    entries: 1024,
    maxRatio: 200,
    envelope: { constantBytes: 64 * MiB, perElement: 1750, perAttribute: 128, perTextChar: 64 },
    evaluation: { perFormulaBytes: 2048, perMaterializedCellBytes: 2000, maxArrayCells: 16_384 },
  },
  paged: {
    compressedBytes: 256 * MiB,
    expandedBytes: 2 * GiB,
    cells: 2 ** 25,
    uniqueStrings: 4_194_304,
    sharedStringsBytes: 512 * MiB,
    entries: 4096,
    centralDirectoryBytes: 4 * MiB,
    maxRatio: 200,
    drawingPartEnvelopeBytes: 128 * MiB,
    minFreeStorageBytes: 64 * MiB,
    structuralRecords: 1_048_576,
    pivotPartBytes: 8 * MiB,
  },
  spool: { memoryLimitBytes: 32 * MiB, hardMemoryCapBytes: 64 * MiB, sniffPrefixBytes: 64 * KiB },
}

const MOBILE: ImportLimits = {
  ...DESKTOP,
  object: {
    ...DESKTOP.object,
    budgetBytes: 128 * MiB,
    compressedBytes: 16 * MiB,
    nonDrawingExpandedBytes: 3 * MiB,
    nonDrawingPartBytes: 2 * MiB,
    drawingPartBytes: 4 * MiB,
    rows: 50_000,
    cells: 50_000,
    uniqueStrings: 25_000,
    merges: 10_000,
    styleRecords: 10_000,
    entries: 512,
    evaluation: { ...DESKTOP.object.evaluation, maxArrayCells: 4096 },
  },
  paged: {
    ...DESKTOP.paged,
    compressedBytes: 128 * MiB,
    expandedBytes: 1 * GiB,
    cells: 2 ** 24,
    uniqueStrings: 2_097_152,
    sharedStringsBytes: 256 * MiB,
    entries: 2048,
    drawingPartEnvelopeBytes: 64 * MiB,
    pivotPartBytes: 4 * MiB,
  },
  spool: { memoryLimitBytes: 16 * MiB, hardMemoryCapBytes: 32 * MiB, sniffPrefixBytes: 64 * KiB },
}

export interface ImportLimitOverrides extends Partial<Omit<ImportLimits, 'object' | 'paged' | 'spool'>> {
  object?: Partial<Omit<ImportLimits['object'], 'envelope' | 'evaluation'>> & {
    envelope?: Partial<ImportLimits['object']['envelope']>
    evaluation?: Partial<ImportLimits['object']['evaluation']>
  }
  paged?: Partial<ImportLimits['paged']>
  spool?: Partial<ImportLimits['spool']>
}

export function importLimitsFor(profile: ProfileId, overrides: ImportLimitOverrides = {}): ImportLimits {
  const base = profile === 'M' ? MOBILE : DESKTOP
  const { object, paged, spool, ...top } = overrides
  return {
    ...base,
    ...top,
    object: {
      ...base.object,
      ...object,
      envelope: { ...base.object.envelope, ...object?.envelope },
      evaluation: { ...base.object.evaluation, ...object?.evaluation },
    },
    paged: { ...base.paged, ...paged },
    spool: { ...base.spool, ...spool },
  }
}

/**
 * Profile at the entry point: 'M' for phones and tablets (including iPadOS,
 * which reports a desktop UA with touch), 'D' otherwise. Internal callers and
 * tests pass a profile explicitly.
 */
export function detectProfile(): ProfileId {
  if (typeof navigator === 'undefined') return 'D'
  const nav = navigator as Navigator & { userAgentData?: { mobile?: boolean } }
  if (nav.userAgentData?.mobile === true) return 'M'
  const ua = nav.userAgent ?? ''
  if (/Mobi|Android|iPhone|iPad|iPod/i.test(ua)) return 'M'
  if (/Macintosh/.test(ua) && (nav.maxTouchPoints ?? 0) > 1) return 'M'
  return 'D'
}

/** Why a load did not produce a document. `paged-not-yet` is retained for type
 * compatibility only: since S1.2 nothing produces it (paged admissions open a
 * paged session internally, or the compatibility document on exported entries). */
export type RefusalReason = 'paged-not-yet' | 'undetermined-format' | 'limit'

/** Typed rejection of loadOfficeFile/<OfficeFile> for a spreadsheet that cannot load as an object document. */
export class ImportRefusal extends Error {
  readonly code: PackageErrorCode
  readonly reason: RefusalReason
  readonly failure: SourceFailure
  readonly partId?: string

  constructor(code: PackageErrorCode, reason: RefusalReason, message: string, partId?: string) {
    super(message)
    this.name = 'ImportRefusal'
    this.code = code
    this.reason = reason
    this.failure = { code, retryable: code === 'cancelled' || code === 'storage-error' || code === 'quota', message }
    if (partId !== undefined) this.partId = partId
  }
}

// ---------------------------------------------------------------------------
// Paged import port (implemented by S1.2's storage; defined here so the
// admission result and the future importer share one contract).

export interface ImportedPage { sheetId: string; pageId: string; r0: number; r1: number; cellCount: number; bytes: number }
export interface PageSink {
  /** Awaited: the downstream backpressure point; a page is never half-published. */
  commit(page: ImportedPage): Promise<void>
  /** Cancel/failure: remove everything committed for this owner. */
  abandon(docId: string): Promise<void>
}
export interface ImportProgress {
  phase: 'inventory' | 'styles' | 'sharedStrings' | 'sheet' | 'done'
  compressedRead: number
  compressedTotal: number | undefined
  expandedBytes: number
  expandedTotal: number | 'unknown'
  populatedCells: number
}
export type ImportOutcome =
  | { kind: 'complete'; stamp: SourceStamp }
  | { kind: 'incomplete'; failure: SourceFailure; stamp: SourceStamp }
  | { kind: 'cancelled' }
export interface ImportSession {
  readonly stamp: SourceStamp
  readonly completeness: Completeness
  readonly diagnostics: readonly WorkerDiagnostic[]
  readonly done: Promise<ImportOutcome>
  cancel(): Promise<void>
  release(): Promise<void>
}
