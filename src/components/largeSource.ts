/**
 * The large-source half of the entry wiring, loaded on demand (dynamic
 * import) so DOCX/PPTX-only consumers do not pay for the ZIP/XML readers,
 * pako, the spool and admission: Blob and network inputs, and any byte input
 * that turns out to be a spreadsheet, load this chunk (design note §4.5).
 */
import type { OfficePackage } from '../core/zip'
import { parseXlsx } from '../xlsx/parse'
import { countFormulaCells, savedValuesStatus } from '../xlsx/view/values'
import type { XlsxDocument } from '../xlsx/types'
import { isRtf } from '../rtf/parse'
import type { ByteStreamReader } from '../core/stream'
import { isPackageError, throwIfAborted, type PackageErrorCode } from '../core/package-error'
import { blobRangeSource, bytesRangeSource, readZipInventory, type RangeSource } from '../core/zip-stream'
import { classifyPackage, sniffPackage, type OfficeKind } from '../core/package-sniff'
import { browserStorageEnv, openSpool, sweepLeftoverSpools, type SpooledSource } from '../core/spool'
import { admitSpreadsheet, evaluationHitCap, inventoryLimits, type CompatibilityKind, type SpreadsheetAdmission } from '../xlsx/storage/admission'
import { ImportRefusal } from '../xlsx/storage/importSession'
import { compatibilityFallbackAllowed, firstPageOrOutcome, openImportSession, type PagedImportSession } from '../xlsx/storage/pagedImport'
import type { PageStorageEnv } from '../xlsx/storage/sourcePages'
import { NOTICE_ALTERNATE_CONTENT, NOTICE_CACHED_VALUES, NOTICE_COMPATIBILITY, NOTICE_DAMAGED, NOTICE_STRICT_FALLBACK, type AnyDoc, type LoadCtx } from './loadOfficeSource'

type Ctx = LoadCtx
type LoadBytes = (bytes: Uint8Array, ctx: Ctx, decided?: SpreadsheetAdmission) => Promise<AnyDoc>

let swept = false

function addDiagnostic(doc: XlsxDocument, feature: string, message: string): void {
  doc.diagnostics ??= []
  doc.diagnostics.push({ kind: 'large-source', feature, message })
}

function refusalFrom(admission: Extract<SpreadsheetAdmission, { mode: 'refuse' }>): ImportRefusal {
  return new ImportRefusal(admission.failure.code as PackageErrorCode, 'limit', admission.failure.message)
}

/** Only a refusal is settled before the legacy package exists; every other route loads the bytes. */
function settledEarly(admission: SpreadsheetAdmission): admission is Extract<SpreadsheetAdmission, { mode: 'refuse' }> {
  return admission.mode === 'refuse'
}

/** Engine evaluation with caps; an evaluator resource hit reopens with saved values (decision 2). */
async function parseWithCaps(pkg: OfficePackage, evaluation: { maxArrayCells: number; maxArraySide: number; maxMaterializedCells: number }, policy: 'engine' | 'cached', formulasNotRecalculated: number | undefined, ctx: Ctx): Promise<XlsxDocument> {
  let doc: XlsxDocument | undefined
  if (policy === 'engine') {
    doc = await parseXlsx(pkg, { evaluation })
    // Rule 5: an evaluator resource hit is never published; reopen with saved values.
    if (evaluationHitCap(doc)) doc = undefined
    throwIfAborted(ctx.signal)
  }
  if (doc === undefined) {
    // Saved values only: the allocation is the element envelope already admitted.
    doc = await parseXlsx(pkg, { valuePolicy: 'cached' })
    addDiagnostic(doc, NOTICE_CACHED_VALUES, savedValuesStatus('cached', formulasNotRecalculated ?? countFormulaCells(doc))!)
  }
  return doc
}

const COMPAT_NOTICE = 'Large file — opened in compatibility mode'
const DAMAGED_NOTICE = 'This file may be damaged; some content may be wrong'

/**
 * Today's legacy object path, used whenever admission does not choose the
 * object path and the file is not refused (orchestrator decisions A-C). The
 * evaluator stays capped, so a resource hit still opens with saved values.
 */
async function compatibilityParse(pkg: OfficePackage, kind: CompatibilityKind, reason: string, ctx: Ctx): Promise<XlsxDocument> {
  const o = ctx.limits.object
  const ev = o.evaluation
  const evaluation = { maxArrayCells: ev.maxArrayCells, maxArraySide: ev.maxArrayCells, maxMaterializedCells: Math.floor((o.budgetBytes - o.envelope.constantBytes) / ev.perMaterializedCellBytes) }
  const doc = await parseWithCaps(pkg, evaluation, 'engine', undefined, ctx)
  if (kind === 'large') {
    addDiagnostic(doc, NOTICE_COMPATIBILITY, COMPAT_NOTICE)
    addDiagnostic(doc, 'compatibility-reason', reason)
  } else if (kind === 'damaged') {
    addDiagnostic(doc, NOTICE_DAMAGED, DAMAGED_NOTICE)
    addDiagnostic(doc, 'damage-reason', reason)
  } else {
    addDiagnostic(doc, NOTICE_STRICT_FALLBACK, `strict reader refused (${reason}); parsed by the legacy reader`)
  }
  return doc
}


export async function parseSpreadsheet(pkg: OfficePackage, bytes: Uint8Array, ctx: Ctx, decided?: SpreadsheetAdmission): Promise<XlsxDocument> {
  const admission = decided ?? await admitSpreadsheet(bytesRangeSource(bytes), ctx.profile, ctx.limits, ctx.signal)
  throwIfAborted(ctx.signal)
  switch (admission.mode) {
    case 'refuse':
      throw refusalFrom(admission)
    case 'paged':
      // The exported entry returns an XlsxDocument, which no paged session can
      // be; paged sessions open through openSpreadsheetSource (internal entry,
      // used by the resident viewer) and here keep the compatibility parse with
      // its notice (decision A) until a paged consumer renders them.
      return compatibilityParse(pkg, 'large', admission.reason, ctx)
    case 'legacy-fallback':
      return compatibilityParse(pkg, admission.kind, admission.reason, ctx)
    case 'object':
      break
  }
  const { decision } = admission
  const doc = await parseWithCaps(pkg, decision.evaluation, decision.valuePolicy, decision.counters.loadEvaluatedFormulas, ctx)
  if (decision.legacyFallback) {
    addDiagnostic(doc, NOTICE_STRICT_FALLBACK, `strict reader refused ${decision.legacyFallback.part} (${decision.legacyFallback.message}); parsed by the legacy reader`)
  }
  for (const part of decision.counters.alternateContentData) {
    addDiagnostic(doc, NOTICE_ALTERNATE_CONTENT, `Some cell data in ${part} is wrapped in mc:AlternateContent and is not shown`)
  }
  for (const part of decision.counters.highRatioParts) {
    addDiagnostic(doc, 'high-compression-ratio', `${part} declares a compression ratio above ${ctx.limits.object.maxRatio}:1 (sizes verified exactly)`)
  }
  return doc
}

/** Strict-reader classification; a refusal means "not ours to judge": the legacy path decides. */
async function classifySeekable(src: RangeSource, ctx: Ctx): Promise<OfficeKind | undefined> {
  try {
    const inventory = await readZipInventory(src, inventoryLimits(ctx.limits), ctx.signal)
    return await classifyPackage(src, inventory, ctx.signal)
  } catch (error) {
    if (isPackageError(error) && error.code === 'cancelled') throw error
    return undefined
  }
}

async function readWhole(blob: Blob, ctx: Ctx): Promise<Uint8Array> {
  const bytes = new Uint8Array(await blob.arrayBuffer())
  throwIfAborted(ctx.signal)
  return bytes
}

/** Seekable input: only the tail, the central directory and small parts are read before admission. */
export async function loadSeekable(blob: Blob, ctx: Ctx, loadBytes: LoadBytes): Promise<AnyDoc> {
  const progress = ctx.options.onProgress
  progress?.({ loaded: 0, total: blob.size })
  const head = new Uint8Array(await blob.slice(0, Math.min(blob.size, 4096)).arrayBuffer())
  let kind: OfficeKind | undefined
  if (!isRtf(head)) kind = await classifySeekable(blobRangeSource(blob), ctx)
  let decided: SpreadsheetAdmission | undefined
  if (kind === 'xlsx') {
    decided = await admitSpreadsheet(blobRangeSource(blob), ctx.profile, ctx.limits, ctx.signal)
    if (settledEarly(decided)) throw refusalFrom(decided)
  }
  const bytes = await readWhole(blob, ctx)
  progress?.({ loaded: bytes.length, total: blob.size })
  return loadBytes(bytes, ctx, decided)
}

/** Sequential input past the 64 KiB prefix: accumulate (DOCX/PPTX/other hint) or spool, then classify. */
export async function continueSequential(reader: ByteStreamReader, chunks: Uint8Array[], held: number, ctx: Ctx, loadBytes: LoadBytes): Promise<AnyDoc> {
  const prefixBytes = ctx.limits.spool.sniffPrefixBytes
  const prefix = concat(chunks, held).subarray(0, prefixBytes)
  const hint = isRtf(prefix) ? 'rtf' : sniffPackage(prefix)
  if (hint !== 'xlsx' && hint !== 'undetermined') {
    // DOCX/PPTX/other/RTF/not-zip: today's accumulate-everything path.
    try {
      for (;;) {
        const chunk = await reader.read()
        throwIfAborted(ctx.signal)
        if (chunk === undefined) break
        chunks.push(chunk)
        held += chunk.length
      }
    } catch (error) {
      await reader.cancel()
      throw error
    }
    return loadBytes(concat(chunks, held), ctx)
  }
  const spooled = await spool(reader, chunks, hint, ctx)
  try {
    ctx.options.onSpooled?.(spooled)
    const kind = await classifySeekable(blobRangeSource(spooled.blob), ctx)
    let decided: SpreadsheetAdmission | undefined
    if (kind === 'xlsx') {
      decided = await admitSpreadsheet(blobRangeSource(spooled.blob), ctx.profile, ctx.limits, ctx.signal)
      if (settledEarly(decided)) throw refusalFrom(decided)
    }
    // Object path or a non-spreadsheet: today's in-memory legacy load. The
    // OfficePackage keeps these bytes, so the spool is released right away.
    const bytes = await readWhole(spooled.blob, ctx)
    await spooled.release()
    return await loadBytes(bytes, ctx, decided)
  } finally {
    await spooled.release()
  }
}

async function spool(reader: ByteStreamReader, held: Uint8Array[], hint: 'xlsx' | 'undetermined', ctx: Ctx): Promise<SpooledSource> {
  const env = ctx.options.storage ?? {
    ...browserStorageEnv(ctx.limits.spool.memoryLimitBytes),
    hardMemoryCapBytes: ctx.limits.spool.hardMemoryCapBytes,
  }
  if (!swept) {
    swept = true
    void sweepLeftoverSpools(env).catch(() => {})
  }
  let sink
  try {
    sink = await openSpool(env, { maxBytes: ctx.limits.paged.compressedBytes, advertisedBytes: reader.total }, ctx.signal)
    for (const chunk of held.splice(0)) await sink.append(chunk, ctx.signal)
    for (;;) {
      const chunk = await reader.read()
      if (chunk === undefined) break
      // awaited: a slow sink pauses the download (at most one chunk read ahead)
      await sink.append(chunk, ctx.signal)
    }
    throwIfAborted(ctx.signal)
    return await sink.finish()
  } catch (error) {
    await sink?.abort()
    await reader.cancel()
    if (isPackageError(error) && error.code !== 'cancelled') {
      if (hint === 'undetermined') {
        throw new ImportRefusal(error.code, 'undetermined-format',
          `The file type could not be determined within the first ${ctx.limits.spool.sniffPrefixBytes} bytes, and the download could not be stored to check it: ${error.message}`)
      }
      throw new ImportRefusal(error.code, 'limit', `This spreadsheet download could not be stored for loading: ${error.message}`)
    }
    throw error
  }
}

function concat(chunks: Uint8Array[], total: number): Uint8Array {
  if (chunks.length === 1 && chunks[0].length === total) return chunks[0]
  const out = new Uint8Array(total)
  let at = 0
  for (const c of chunks) { out.set(c, at); at += c.length }
  return out
}

/** Internal entry result (design note §4.5 item 4). */
export type SpreadsheetSource =
  | { kind: 'paged'; session: PagedImportSession }
  | { kind: 'document'; document: XlsxDocument }

/**
 * Internal seekable-spreadsheet entry: admission, then a paged import session
 * for a paged decision (no whole-file copy). The object path, and a file the
 * paged reader refuses before any page landed, load as today's document; the
 * latter carries the compatibility notice. Storage/quota refusals are typed
 * ImportRefusals: no unbounded in-memory fallback.
 */
export async function openSpreadsheetSource(blob: Blob, ctx: Ctx, storage?: PageStorageEnv): Promise<SpreadsheetSource> {
  const admission = await admitSpreadsheet(blobRangeSource(blob), ctx.profile, ctx.limits, ctx.signal)
  throwIfAborted(ctx.signal)
  const legacy = async (decided: SpreadsheetAdmission): Promise<SpreadsheetSource> => {
    const bytes = await readWhole(blob, ctx)
    const { OfficePackage } = await import('../core/zip')
    return { kind: 'document', document: await parseSpreadsheet(await OfficePackage.load(bytes), bytes, ctx, decided) }
  }
  if (admission.mode !== 'paged') return legacy(admission)
  const fallback = await compatibilityFallbackAllowed(blob, ctx.limits, ctx.signal)
  let session: PagedImportSession
  try {
    session = await openImportSession({ kind: 'blob', blob }, { docId: `doc-${Date.now().toString(36)}`, profile: ctx.profile, signal: ctx.signal, limits: ctx.options.limits, storage })
  } catch (error) {
    if (isPackageError(error) && (error.code === 'storage-error' || error.code === 'quota' || error.code === 'cancelled')) {
      if (error.code === 'cancelled') throw error
      throw new ImportRefusal(error.code, 'limit', `This spreadsheet needs browser storage to open: ${error.message}`)
    }
    if (!fallback) {
      const code = isPackageError(error) ? error.code : 'parse-error'
      throw new ImportRefusal(code, 'limit', `This large spreadsheet could not be read: ${(error as Error).message}`)
    }
    return legacy(admission)
  }
  // Within the object caps a file the paged reader refuses before any page
  // landed loads through the compatibility parse (with its notice); beyond
  // them the incomplete session is returned with its typed failure.
  const early = await firstPageOrOutcome(session)
  if (fallback && early && early.kind === 'incomplete' && early.failure.code === 'unsupported' && session.workbook.store.stats.committed === 0) {
    await session.release()
    // compatibilityParse adds the 'opened in compatibility mode' notice.
    return legacy({ mode: 'paged', reason: early.failure.message })
  }
  return { kind: 'paged', session }
}
