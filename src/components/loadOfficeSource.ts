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
  /** S5.4: opens a paged admission in the resident viewer worker and returns its
   * shell document; undefined keeps the compatibility parse. `release` hands over
   * a spooled source, which the route then owns. Absent: today's behaviour. */
  paged?: PagedRoute
}
export type PagedRoute = (blob: Blob, ctx: { profile: ProfileId; signal: AbortSignal }, release?: () => Promise<void>) => Promise<XlsxDocument | undefined>

export interface LoadCtx {
  signal: AbortSignal
  profile: ProfileId
  limits: ImportLimits
  options: LoadSourceOptions
}

/** Diagnostic features the viewer shows to the user (OfficeFileState.notices). */
export const NOTICE_ALTERNATE_CONTENT = 'alternate-content-data'
export const NOTICE_CACHED_VALUES = 'cached-value-policy'
/** Debug diagnostic, not a notice: the legacy output is identical to today's. */
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

type BlobHint = 'legacy' | 'spreadsheet'

/**
 * Eager hint for a local Blob, without the large-source chunk (decision D):
 * the first bytes (RTF, ZIP signature) and the classic central directory's
 * part names. Like the network prefix hint it only picks a path: DOCX/PPTX/
 * other go to today's whole-file load, where the parse-time guard still
 * sends any spreadsheet the legacy view finds through admission. A package
 * naming `xl/workbook.xml` (or that cannot be read cheaply: ZIP64, a
 * central directory above 1 MiB, no end record) gets the precise, lazy path.
 */
async function blobHint(blob: Blob): Promise<BlobHint> {
  const read = async (a: number, b: number) => new Uint8Array(await blob.slice(a, b).arrayBuffer())
  const head = await read(0, Math.min(blob.size, 4096))
  if (isRtf(head) || head.length < 4 || head[0] !== 0x50 || head[1] !== 0x4b || head[2] !== 3 || head[3] !== 4) return 'legacy'
  const tailStart = Math.max(0, blob.size - 65_557)
  const tail = await read(tailStart, blob.size)
  const u16 = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8)
  const u32 = (b: Uint8Array, i: number) => u16(b, i) + u16(b, i + 2) * 0x10000
  let eocd = -1
  for (let i = tail.length - 22; i >= 0; i--) {
    if (u32(tail, i) === 0x06054b50 && u16(tail, i + 20) === tail.length - i - 22) { eocd = i; break }
  }
  if (eocd < 0) return 'spreadsheet'
  const count = u16(tail, eocd + 10)
  const cdSize = u32(tail, eocd + 12)
  const cdOffset = u32(tail, eocd + 16)
  if (count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff || cdSize > 1 << 20 || cdOffset + cdSize > tailStart + eocd) return 'spreadsheet'
  const cd = cdOffset >= tailStart ? tail.subarray(cdOffset - tailStart, cdOffset - tailStart + cdSize) : await read(cdOffset, cdOffset + cdSize)
  const names = new Set<string>()
  const decoder = new TextDecoder()
  for (let p = 0, k = 0; k < count && p + 46 <= cd.length; k++) {
    if (u32(cd, p) !== 0x02014b50) return 'spreadsheet'
    const n = u16(cd, p + 28)
    names.add(decoder.decode(cd.subarray(p + 46, p + 46 + n)))
    p += 46 + n + u16(cd, p + 30) + u16(cd, p + 32)
  }
  return names.has('xl/workbook.xml') ? 'spreadsheet' : 'legacy'
}

/** Today's Blob path: one whole read, the legacy load, the parse-time guard. */
async function loadBlobLegacy(blob: Blob, ctx: LoadCtx): Promise<AnyDoc> {
  ctx.options.onProgress?.({ loaded: 0, total: blob.size })
  const bytes = new Uint8Array(await blob.arrayBuffer())
  throwIfAborted(ctx.signal)
  ctx.options.onProgress?.({ loaded: bytes.length, total: blob.size })
  return loadBytes(bytes, ctx)
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
    case 'blob':
      if (await blobHint(opened.blob) === 'legacy') return loadBlobLegacy(opened.blob, ctx)
      return (await largeSource()).loadSeekable(opened.blob, ctx, loadBytes)
    case 'stream': return loadSequential(opened.reader, ctx)
  }
}
