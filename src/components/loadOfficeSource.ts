/**
 * Entry-point routing behind loadOfficeFile / <OfficeFile> (design note §4.5).
 *
 * DOCX, PPTX, ODT, RTF and anything that is not a spreadsheet keep today's
 * path: the bytes are accumulated and handed to OfficePackage plus the legacy
 * detectAndParse, unchanged. A spreadsheet never reaches parseXlsx without
 * admission: seekable inputs are classified from the central directory
 * without reading the whole file, sequential inputs whose 64 KiB prefix hints
 * at a spreadsheet (or cannot be classified) are spooled to a seekable source
 * first, and every remaining route to parseXlsx goes through the same guard.
 */
import { OfficePackage } from '../core/zip'
import { parseDocx } from '../docx/parse'
import type { DocxDocument } from '../docx/types'
import { parseXlsx } from '../xlsx/parse'
import type { XlsxDocument } from '../xlsx/types'
import { parsePptx } from '../pptx/parse'
import type { PptxDocument } from '../pptx/types'
import { odfKind } from '../odf/container'
import { parseOdt } from '../odt/parse'
import { parseRtf, isRtf } from '../rtf/parse'
import { openByteStream, type ByteSource, type ByteStreamReader, type Progress, type ProtectionListener, type WatermarkListener } from '../core/stream'
import { isPackageError, throwIfAborted, type PackageErrorCode } from '../core/package-error'
import { blobRangeSource, bytesRangeSource, readZipInventory, type RangeSource } from '../core/zip-stream'
import { classifyPackage, sniffPackage, type OfficeKind } from '../core/package-sniff'
import { browserStorageEnv, openSpool, sweepLeftoverSpools, type SpooledSource, type StorageEnv } from '../core/spool'
import { admitSpreadsheet, evaluationHitCap, inventoryLimits, type SpreadsheetAdmission } from '../xlsx/storage/admission'
import { detectProfile, importLimitsFor, ImportRefusal, type ImportLimitOverrides, type ImportLimits } from '../xlsx/storage/importSession'
import type { ProfileId } from '../xlsx/view/contracts'

export type AnyDoc = DocxDocument | XlsxDocument | PptxDocument

type XlsxParser = (pkg: OfficePackage) => Promise<XlsxDocument>

/** The legacy format decision, unchanged except that both xlsx branches go through `xlsx`. */
export async function detectAndParse(pkg: OfficePackage, xlsx: XlsxParser = parseXlsx): Promise<AnyDoc> {
  const ct = await pkg.xml('[Content_Types].xml')
  const ctText = JSON.stringify(ct ?? {})
  if (ctText.includes('wordprocessingml.document.main')) return parseDocx(pkg)
  if (ctText.includes('spreadsheetml.sheet.main')) return xlsx(pkg)
  if (ctText.includes('presentationml.presentation.main')) return parsePptx(pkg)
  // fall back to probing known parts
  if (pkg.has('word/document.xml')) return parseDocx(pkg)
  if (pkg.has('xl/workbook.xml')) return xlsx(pkg)
  if (pkg.has('ppt/presentation.xml')) return parsePptx(pkg)
  // OpenDocument (Writer first — Calc/Impress route here when they land)
  const kind = await odfKind(pkg)
  if (kind === 'text') return parseOdt(pkg)
  if (kind) throw new Error(`Unsupported OpenDocument type '${kind}' — only word-processing (.odt) documents render for now`)
  throw new Error('Unrecognized office file: no word/xl/ppt content detected')
}

export interface LoadSourceOptions {
  onProgress?: (p: Progress) => void
  onProtection?: ProtectionListener
  onWatermark?: WatermarkListener
  /** Aborts the download, the spool, the verification pass and the parse. */
  signal?: AbortSignal
  /** Default: run-time probe of OPFS/IndexedDB (memory only where neither exists). */
  storage?: StorageEnv
  profile?: ProfileId
  limits?: ImportLimitOverrides
  /** Observes every spool handed out, for ownership checks. */
  onSpooled?: (source: SpooledSource) => void
}

interface Ctx {
  signal: AbortSignal
  profile: ProfileId
  limits: ImportLimits
  options: LoadSourceOptions
}

let swept = false

function refusalFrom(admission: Exclude<SpreadsheetAdmission, { mode: 'object' }>): ImportRefusal {
  if (admission.mode === 'paged') {
    return new ImportRefusal('unsupported', 'paged-not-yet',
      `This spreadsheet is too large to open in the current viewer (${admission.reason}); the large-workbook viewer is not available yet`)
  }
  return new ImportRefusal(admission.failure.code as PackageErrorCode, 'limit', admission.failure.message)
}

/** The only way an admitted workbook reaches parseXlsx: with the returned evaluator caps. */
function admittedParser(bytes: Uint8Array, ctx: Ctx, decided?: SpreadsheetAdmission): XlsxParser {
  return async (pkg) => {
    const admission = decided ?? await admitSpreadsheet(bytesRangeSource(bytes), ctx.profile, ctx.limits, ctx.signal)
    if (admission.mode !== 'object') throw refusalFrom(admission)
    throwIfAborted(ctx.signal)
    const doc = await parseXlsx(pkg, { evaluation: admission.decision.evaluation })
    const hit = evaluationHitCap(doc)
    // Rule 5: an evaluator resource hit reroutes to paged/`cached` before publication.
    if (hit) {
      throw new ImportRefusal('unsupported', 'paged-not-yet',
        `Formula evaluation in this spreadsheet exceeds the load-time budget (${hit}); saved values need the large-workbook viewer, which is not available yet`)
    }
    return doc
  }
}

async function loadBytes(bytes: Uint8Array, ctx: Ctx, decided?: SpreadsheetAdmission): Promise<AnyDoc> {
  // RTF is plain text, not a zip, so it must be recognised before the bytes
  // reach OfficePackage.load, which would throw on them.
  if (isRtf(bytes)) return parseRtf(bytes)
  const pkg = await OfficePackage.load(bytes)
  throwIfAborted(ctx.signal)
  return detectAndParse(pkg, admittedParser(bytes, ctx, decided))
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
async function loadSeekable(blob: Blob, ctx: Ctx): Promise<AnyDoc> {
  const progress = ctx.options.onProgress
  progress?.({ loaded: 0, total: blob.size })
  const head = new Uint8Array(await blob.slice(0, Math.min(blob.size, 4096)).arrayBuffer())
  let kind: OfficeKind | undefined
  if (!isRtf(head)) kind = await classifySeekable(blobRangeSource(blob), ctx)
  let decided: SpreadsheetAdmission | undefined
  if (kind === 'xlsx') {
    decided = await admitSpreadsheet(blobRangeSource(blob), ctx.profile, ctx.limits, ctx.signal)
    if (decided.mode !== 'object') throw refusalFrom(decided)
  }
  const bytes = await readWhole(blob, ctx)
  progress?.({ loaded: bytes.length, total: blob.size })
  return loadBytes(bytes, ctx, decided)
}

function concat(chunks: Uint8Array[], total: number): Uint8Array {
  if (chunks.length === 1 && chunks[0].length === total) return chunks[0]
  const out = new Uint8Array(total)
  let at = 0
  for (const c of chunks) { out.set(c, at); at += c.length }
  return out
}

async function loadSequential(reader: ByteStreamReader, ctx: Ctx): Promise<AnyDoc> {
  const prefixBytes = ctx.limits.spool.sniffPrefixBytes
  const chunks: Uint8Array[] = []
  let held = 0
  let ended = false
  try {
    while (held < prefixBytes) {
      const chunk = await reader.read()
      throwIfAborted(ctx.signal)
      if (chunk === undefined) { ended = true; break }
      chunks.push(chunk)
      held += chunk.length
    }
  } catch (error) {
    await reader.cancel()
    throw error
  }
  // The whole download fits in the prefix: it is already in memory.
  if (ended) return loadBytes(concat(chunks, held), ctx)
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
      if (decided.mode !== 'object') throw refusalFrom(decided)
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

/** Load any accepted source into a document, routing spreadsheets through admission. */
export async function loadOfficeSource(source: ByteSource, options: LoadSourceOptions = {}): Promise<AnyDoc> {
  const profile = options.profile ?? detectProfile()
  const ctx: Ctx = {
    signal: options.signal ?? new AbortController().signal,
    profile,
    limits: importLimitsFor(profile, options.limits),
    options,
  }
  throwIfAborted(ctx.signal)
  const opened = await openByteStream(source, {
    onProgress: options.onProgress,
    onProtection: options.onProtection,
    onWatermark: options.onWatermark,
  }, ctx.signal)
  throwIfAborted(ctx.signal)
  switch (opened.kind) {
    case 'bytes': return loadBytes(opened.bytes, ctx)
    case 'blob': return loadSeekable(opened.blob, ctx)
    case 'stream': return loadSequential(opened.reader, ctx)
  }
}

