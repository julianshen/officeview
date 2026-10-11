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
 * The large-source code lives in ./largeSource and is loaded on demand, so a
 * byte-input DOCX/PPTX never loads it.
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
import { throwIfAborted } from '../core/package-error'
import type { SpooledSource, StorageEnv } from '../core/spool'
import type { SpreadsheetAdmission } from '../xlsx/storage/admission'
import { detectProfile, importLimitsFor, type ImportLimitOverrides, type ImportLimits } from '../xlsx/storage/importSession'
import type { ProfileId } from '../xlsx/view/contracts'

export type AnyDoc = DocxDocument | XlsxDocument | PptxDocument

type XlsxParser = (pkg: OfficePackage) => Promise<XlsxDocument>

const largeSource = () => import('./largeSource')

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

export interface LoadCtx {
  signal: AbortSignal
  profile: ProfileId
  limits: ImportLimits
  options: LoadSourceOptions
}

/** Diagnostic features the viewer shows to the user (OfficeFileState.notices). */
export const NOTICE_ALTERNATE_CONTENT = 'alternate-content-data'
export const NOTICE_CACHED_VALUES = 'cached-value-policy'
export const NOTICE_STRICT_FALLBACK = 'strict-reader-fallback'
export const NOTICE_COMPATIBILITY = 'large-file-compatibility'
export const NOTICE_DAMAGED = 'possibly-damaged'
const NOTICE_FEATURES = new Set([NOTICE_ALTERNATE_CONTENT, NOTICE_CACHED_VALUES, NOTICE_COMPATIBILITY, NOTICE_DAMAGED])

export function documentNotices(doc: AnyDoc): string[] {
  const diagnostics = (doc as { diagnostics?: Array<{ feature: string; message: string }> }).diagnostics ?? []
  return diagnostics.filter((d) => NOTICE_FEATURES.has(d.feature)).map((d) => d.message)
}

async function loadBytes(bytes: Uint8Array, ctx: LoadCtx, decided?: SpreadsheetAdmission): Promise<AnyDoc> {
  // RTF is plain text, not a zip, so it must be recognised before the bytes
  // reach OfficePackage.load, which would throw on them.
  if (isRtf(bytes)) return parseRtf(bytes)
  const pkg = await OfficePackage.load(bytes)
  throwIfAborted(ctx.signal)
  const xlsx: XlsxParser = async (p) => (await largeSource()).parseSpreadsheet(p, bytes, ctx, decided)
  return detectAndParse(pkg, xlsx)
}

function concat(chunks: Uint8Array[], total: number): Uint8Array {
  if (chunks.length === 1 && chunks[0].length === total) return chunks[0]
  const out = new Uint8Array(total)
  let at = 0
  for (const c of chunks) { out.set(c, at); at += c.length }
  return out
}

/** Hold the 64 KiB prefix; a download that ends inside it is already in memory. */
async function loadSequential(reader: ByteStreamReader, ctx: LoadCtx): Promise<AnyDoc> {
  const prefixBytes = ctx.limits.spool.sniffPrefixBytes
  const chunks: Uint8Array[] = []
  let held = 0
  try {
    while (held < prefixBytes) {
      const chunk = await reader.read()
      throwIfAborted(ctx.signal)
      if (chunk === undefined) return loadBytes(concat(chunks, held), ctx)
      chunks.push(chunk)
      held += chunk.length
    }
  } catch (error) {
    await reader.cancel()
    throw error
  }
  return (await largeSource()).continueSequential(reader, chunks, held, ctx, loadBytes)
}

/** Load any accepted source into a document, routing spreadsheets through admission. */
export async function loadOfficeSource(source: ByteSource, options: LoadSourceOptions = {}): Promise<AnyDoc> {
  const profile = options.profile ?? detectProfile()
  const ctx: LoadCtx = {
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
    case 'blob': return (await largeSource()).loadSeekable(opened.blob, ctx, loadBytes)
    case 'stream': return loadSequential(opened.reader, ctx)
  }
}
