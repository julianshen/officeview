/**
 * Spreadsheet admission gates (design note §4.3).
 *
 * G1 `admit()` looks only at declared central-directory data, before any part
 * is inflated. G2 `verifyObjectCandidate()` inflates every non-media part once
 * through openPart, counts the token stream exactly and stops at the first
 * counter that crosses its cap, so the legacy parser only ever starts on input
 * whose element/attribute/text envelope, secondary caps and evaluation budget
 * all fit. Nothing large is allocated before that decision, so a switch to
 * paged mode never follows an oversized allocation.
 */
import { PackageError, isPackageError, throwIfAborted } from '../../core/package-error'
import { openPart, readZipInventory, type ExpandedBudget, type RangeSource, type ZipEntryInfo, type ZipInventory, type ZipLimits } from '../../core/zip-stream'
import { tokenizeBytes } from '../../core/xml-stream'
import type { ProfileId, SourceFailure } from '../view/contracts'
import type { XlsxDocument } from '../types'
import { createQuotaLedger, quotaLimitsFor, type QuotaLedger } from './quota'
import { countingEvents, newAdmissionCounters, type AdmissionCounters } from './sheetRecords'
import type { ImportLimits } from './importSession'

const MiB = 1024 * 1024
const RATIO_FLOOR = MiB

export type PartClass = 'xml' | 'binary' | 'media'
export interface PartInfo { cls: PartClass; drawing: boolean }

const MEDIA = /\.(png|jpe?g|jfif|gif|bmp|dib|tiff?|emf|wmf|emz|wmz|webp|svg|ico|wdp|jxr|hdp|pict?)$/i
/** Drawing-family parts have their own per-part cap instead of the non-drawing byte guards. */
const DRAWING = /^xl\/(drawings|charts|diagrams)\//i

/**
 * Media are never tokenized and never count toward the XML guards. Only
 * `.xml`/`.rels` parts are tokenized; other parts (printer settings, VBA,
 * OLE objects, legacy VML, whose HTML-ish markup is not XML) are only
 * byte-counted (inflated, CRC-checked) against the byte guards.
 */
export function classifyPart(name: string): PartInfo {
  const drawing = DRAWING.test(name)
  if (MEDIA.test(name) || /^xl\/media\//i.test(name)) return { cls: 'media', drawing: false }
  return { cls: /\.(xml|rels)$/i.test(name) ? 'xml' : 'binary', drawing }
}

function failure(code: SourceFailure['code'], message: string): SourceFailure {
  return { code, retryable: false, message }
}

/** ZIP limits for the inventory read: the paged-path caps, so anything beyond them is a typed refusal. */
export function inventoryLimits(limits: ImportLimits): ZipLimits {
  return {
    maxEntries: limits.paged.entries,
    maxCentralDirectoryBytes: limits.paged.centralDirectoryBytes,
    maxNameBytes: 1024,
    maxPartExpandedBytes: limits.paged.expandedBytes,
    maxTotalExpandedBytes: Number.MAX_SAFE_INTEGER,
    // Ratio alone never refuses: produced == declared is enforced exactly while
    // inflating and the byte guards bound the declared sizes (decision 4).
    maxRatio: Number.MAX_SAFE_INTEGER,
  }
}

export type Admission =
  | { mode: 'object-candidate' }
  | { mode: 'paged'; reason: string }
  | { mode: 'refuse'; failure: SourceFailure }

/** Gate 1: declared limits only (compressed size, central-directory sums, entry count, declared part sizes, ratio). */
export function admit(profile: ProfileId, compressedBytes: number, inventory: ZipInventory, limits: ImportLimits): Admission {
  const { object, paged } = limits
  const refuse = (cap: string, value: number, limit: number, part?: string): Admission => ({
    mode: 'refuse',
    failure: failure('resource-limit', `${cap} ${value} exceeds the ${limit} limit for profile ${profile}${part ? ` (${part})` : ''}`),
  })
  if (compressedBytes > paged.compressedBytes) return refuse('package size', compressedBytes, paged.compressedBytes)
  if (inventory.entries.length > paged.entries) return refuse('entry count', inventory.entries.length, paged.entries)
  let expanded = 0
  let nonDrawing = 0
  let objectReason: string | undefined
  const notObject = (reason: string) => { objectReason ??= reason }
  for (const entry of inventory.entries) {
    const info = classifyPart(entry.name)
    if (info.cls === 'media') continue
    expanded += entry.uncompressedSize
    if (/(^|\/)sharedStrings\.xml$/i.test(entry.name) && entry.uncompressedSize > paged.sharedStringsBytes) {
      return refuse('declared shared strings size', entry.uncompressedSize, paged.sharedStringsBytes, entry.name)
    }
    if (info.drawing) {
      if (entry.uncompressedSize > object.drawingPartBytes) notObject(`drawing part ${entry.name} declares ${entry.uncompressedSize} bytes (> ${object.drawingPartBytes})`)
    } else {
      nonDrawing += entry.uncompressedSize
      if (entry.uncompressedSize > object.nonDrawingPartBytes) notObject(`part ${entry.name} declares ${entry.uncompressedSize} bytes (> ${object.nonDrawingPartBytes})`)
    }
  }
  if (expanded > paged.expandedBytes) return refuse('declared expanded size', expanded, paged.expandedBytes)
  if (compressedBytes > object.compressedBytes) notObject(`package size ${compressedBytes} > ${object.compressedBytes}`)
  if (inventory.entries.length > object.entries) notObject(`entry count ${inventory.entries.length} > ${object.entries}`)
  if (nonDrawing > object.nonDrawingExpandedBytes) notObject(`declared non-drawing parts total ${nonDrawing} bytes (> ${object.nonDrawingExpandedBytes})`)
  return objectReason === undefined ? { mode: 'object-candidate' } : { mode: 'paged', reason: objectReason }
}

export interface ObjectEvaluationCaps { maxArrayCells: number; maxArraySide: number; maxMaterializedCells: number }
export type ObjectDecision =
  | {
    mode: 'object'
    counters: AdmissionCounters
    /** 'engine' when the evaluation reserve fits; otherwise saved values only (no load-time evaluation). */
    valuePolicy: 'engine' | 'cached'
    evaluation: ObjectEvaluationCaps
    predictedPeak: number
    admissionTotal: number
    /** Set when the strict tokenizer refused a part's XML grammar; the legacy parser decides (decision 3). */
    legacyFallback?: { part: string; message: string }
  }
  | { mode: 'paged'; counters: AdmissionCounters; reason: string }
  | { mode: 'refuse'; failure: SourceFailure }

/** `predictedPeak = constant + perElement x elements + perAttribute x attributes + perTextChar x text`. */
export function predictedPeak(counters: AdmissionCounters, limits: ImportLimits): number {
  const e = limits.object.envelope
  return e.constantBytes + e.perElement * counters.elements + e.perAttribute * counters.attributes + e.perTextChar * counters.textChars
}

/** `predictedPeak + N x Cs + (N > 0 ? A x Ccell : 0)`: the G2 admission inequality's left side. */
export function admissionTotal(counters: AdmissionCounters, limits: ImportLimits): number {
  const ev = limits.object.evaluation
  const n = counters.loadEvaluatedFormulas
  return predictedPeak(counters, limits) + n * ev.perFormulaBytes + (n > 0 ? ev.maxArrayCells * ev.perMaterializedCellBytes : 0)
}

/** Parts above 1 MiB whose declared ratio exceeds maxRatio: a diagnostic only. */
export function highRatioParts(inventory: ZipInventory, maxRatio: number): string[] {
  return inventory.entries
    .filter((e) => e.uncompressedSize > RATIO_FLOOR && e.uncompressedSize > maxRatio * e.compressedSize)
    .map((e) => e.name)
}

/** A counter crossed an object-path cap: not an error, the document goes to the paged path. */
class CapTrip {
  constructor(readonly reason: string) {}
}

function checkCaps(counters: AdmissionCounters, pendingString: number, limits: ImportLimits): void {
  const o = limits.object
  if (counters.rows > o.rows) throw new CapTrip(`rows ${counters.rows} > ${o.rows}`)
  if (counters.cells > o.cells) throw new CapTrip(`populated cells ${counters.cells} > ${o.cells}`)
  if (counters.uniqueStrings > o.uniqueStrings) throw new CapTrip(`unique shared strings ${counters.uniqueStrings} > ${o.uniqueStrings}`)
  if (counters.merges > o.merges) throw new CapTrip(`merges ${counters.merges} > ${o.merges}`)
  if (counters.styleRecords > o.styleRecords) throw new CapTrip(`style records ${counters.styleRecords} > ${o.styleRecords}`)
  if (counters.maxAttributesPerElement > o.maxAttributesPerElement) throw new CapTrip(`attributes per element ${counters.maxAttributesPerElement} > ${o.maxAttributesPerElement}`)
  const longest = Math.max(counters.longestString, pendingString)
  if (longest > limits.inlineStringSoftBytes) throw new CapTrip(`string of ${longest} characters > ${limits.inlineStringSoftBytes}`)
  // The element envelope decides object vs paged; the evaluation reserve only
  // decides engine vs saved values at the end (decision 2).
  const peak = predictedPeak(counters, limits)
  if (peak > o.budgetBytes) throw new CapTrip(`predicted peak ${peak} bytes > ${o.budgetBytes}`)
}

/** Workbook first, so fullCalcOnLoad is known before any worksheet formula is counted. */
function verificationOrder(inventory: ZipInventory): ZipEntryInfo[] {
  const rank = (e: ZipEntryInfo) => (e.name === 'xl/workbook.xml' ? 0 : 1)
  return [...inventory.entries].sort((a, b) => rank(a) - rank(b))
}

export interface VerifyOptions {
  /** Ledger to charge (tests inspect its peaks); default: a fresh one for the profile. */
  ledger?: QuotaLedger
  profile?: ProfileId
}

/**
 * Gate 2: one verification pass over every non-media part. Returns the
 * evaluator caps the legacy parse must use, or a paged/refuse decision.
 * Never builds a tree and never retains a part.
 */
export async function verifyObjectCandidate(src: RangeSource, inventory: ZipInventory, limits: ImportLimits, signal: AbortSignal, options: VerifyOptions = {}): Promise<ObjectDecision> {
  const o = limits.object
  const ledger = options.ledger ?? createQuotaLedger(quotaLimitsFor(options.profile ?? 'D'), { maxTotalExpandedBytes: Number.MAX_SAFE_INTEGER })
  const counters = newAdmissionCounters()
  counters.highRatioParts = highRatioParts(inventory, o.maxRatio)
  let nonDrawing = 0
  let legacyFallback: { part: string; message: string } | undefined
  try {
    for (const entry of verificationOrder(inventory)) {
      throwIfAborted(signal)
      const info = classifyPart(entry.name)
      if (info.cls === 'media') continue
      const partCap = info.drawing ? o.drawingPartBytes : o.nonDrawingPartBytes
      let partBytes = 0
      const open = () => {
        const ledgerBudget = ledger.partBudget(entry.name, partCap)
        const budget: ExpandedBudget = {
          charge(bytes) {
            if (partBytes + bytes > partCap) throw new CapTrip(`part ${entry.name} expands beyond ${partCap} bytes`)
            if (!info.drawing && nonDrawing + bytes > o.nonDrawingExpandedBytes) throw new CapTrip(`non-drawing parts expand beyond ${o.nonDrawingExpandedBytes} bytes`)
            ledgerBudget.charge(bytes)
            partBytes += bytes
            if (!info.drawing) nonDrawing += bytes
            counters.expandedBytes += bytes
            if (partBytes > counters.maxPartBytes) counters.maxPartBytes = partBytes
          },
        }
        return tagged(openPart(src, entry, { signal, budget, rangeReadBytes: limits.rangeReadBytes, outputBufferBytes: limits.inflateOutputBytes }), entry.name)
      }
      const drain = async () => { for await (const chunk of open()) void chunk }
      try {
        if (info.cls === 'xml' && legacyFallback === undefined) {
          const events = countingEvents(counters, entry.name, { check: (c, pending) => checkCaps(c, pending, limits), alternateContentData: 'record' })
          try {
            await tokenizeBytes(open(), events, { signal, limits: { maxTokenChars: limits.maxTokenChars, maxAttributes: o.maxAttributesPerElement } })
          } catch (error) {
            if (error instanceof PartByteError || error instanceof CapTrip || !isPackageError(error) || error.code === 'cancelled') throw error
            // An XML grammar the strict tokenizer refuses (DOCTYPE, long
            // references, attribute/depth caps, malformed markup) is not a
            // safety violation: the legacy parser decides, as today
            // (decision 3). This part and the rest are still byte-verified.
            legacyFallback = { part: entry.name, message: error.message }
            if (!info.drawing) nonDrawing -= partBytes
            counters.expandedBytes -= partBytes
            partBytes = 0
            ledger.drainQueue()
            await drain()
          }
        } else {
          await drain()
        }
      } finally {
        ledger.drainQueue()
      }
    }
  } catch (error) {
    if (error instanceof CapTrip) return { mode: 'paged', counters, reason: error.reason }
    const cause = error instanceof PartByteError ? error.error : error
    if (isPackageError(cause)) {
      if (cause.code === 'cancelled') throw cause
      return { mode: 'refuse', failure: failure(cause.code, cause.message) }
    }
    throw cause
  }
  const peak = predictedPeak(counters, limits)
  const total = admissionTotal(counters, limits)
  const ev = o.evaluation
  const maxMaterializedCells = Math.max(0, Math.floor((o.budgetBytes - peak - counters.loadEvaluatedFormulas * ev.perFormulaBytes) / ev.perMaterializedCellBytes))
  return {
    mode: 'object',
    counters,
    valuePolicy: total <= o.budgetBytes ? 'engine' : 'cached',
    predictedPeak: peak,
    admissionTotal: total,
    evaluation: { maxArrayCells: ev.maxArrayCells, maxArraySide: ev.maxArrayCells, maxMaterializedCells },
    ...(legacyFallback ? { legacyFallback } : {}),
  }
}

/** A failure raised by the part's bytes (size, CRC, termination): always a safety refusal. */
class PartByteError {
  constructor(readonly error: unknown) {}
}

async function* tagged(part: AsyncGenerator<Uint8Array, void, void>, partId: string): AsyncGenerator<Uint8Array, void, void> {
  try {
    for await (const chunk of part) yield chunk
  } catch (error) {
    if (error instanceof CapTrip) throw error
    throw new PartByteError(isPackageError(error) && error.partId === undefined ? new PackageError(error.code, error.message, partId) : error)
  }
}

/** Load-time evaluator resource hits that reroute an object document to paged/`cached` (design note §4.3 rule 5). */
const REROUTE_FEATURES = new Set(['array-resource-limit', 'spill-resource', 'lookup-matrix-resource-limit'])

export function evaluationHitCap(doc: XlsxDocument): string | undefined {
  return doc.diagnostics?.find((d) => REROUTE_FEATURES.has(d.feature))?.feature
}

export type CompatibilityKind = 'strict' | 'damaged' | 'large'
export type SpreadsheetAdmission =
  | { mode: 'object'; decision: Extract<ObjectDecision, { mode: 'object' }> }
  | { mode: 'paged'; reason: string; counters?: AdmissionCounters }
  /**
   * The legacy object path decides, as today: `strict` = a grammar or
   * capability limit of the strict reader, `damaged` = archive damage the
   * legacy JSZip path tolerates, `large` = beyond the paged caps.
   * `duplicateName`: a case-insensitive-only duplicate downgrades to `strict`.
   */
  | { mode: 'legacy-fallback'; reason: string; kind: CompatibilityKind; duplicateName?: string }
  /** Only damage the legacy JSZip path rejects as well (decision C). */
  | { mode: 'refuse'; failure: SourceFailure }

/**
 * Strict-reader refusals that are grammar or capability limits (decision 3):
 * entry count, central directory size, name length/encoding/escaping,
 * encryption, methods, multi-disk.
 */
const GRAMMAR_INVENTORY = /entries exceed the limit|central directory of \d+ bytes exceeds|part name of \d+ bytes exceeds|part name is not valid UTF-8|escaping or invalid part name|empty part name/
/**
 * The damage JSZip also rejects (measured: "Bug : uncompressed data size
 * mismatch" after inflating the whole part): produced != declared size and a
 * deflate stream that ends before its final block. Everything else the strict
 * reader calls damage (CRC, trailing bytes, local != central header, exact
 * duplicate names, overlapping entries, inconsistent EOCD/ZIP64) opens today
 * and falls back with a "may be damaged" notice (decision C).
 */
const REFUSED_DAMAGE = /part inflates beyond its declared size|part produced \d+ of \d+ declared bytes|compressed data ends before the deflate stream does|deflate stream consumed \d+ of \d+ declared bytes/

export function failureRoute(failure: { code: string; message: string }): CompatibilityKind | 'refuse' {
  if (failure.code === 'unsupported' || GRAMMAR_INVENTORY.test(failure.message)) return 'strict'
  if (REFUSED_DAMAGE.test(failure.message) || failure.code === 'storage-error' || failure.code === 'quota' || failure.code === 'cancelled') return 'refuse'
  if (failure.code === 'resource-limit') return 'large'
  return 'damaged'
}

function routed(failure: SourceFailure): SpreadsheetAdmission {
  const kind = failureRoute(failure)
  if (kind === 'refuse') return { mode: 'refuse', failure }
  const duplicate = /^duplicate part name (.*)$/.exec(failure.message)
  return { mode: 'legacy-fallback', reason: failure.message, kind, ...(duplicate ? { duplicateName: duplicate[1] } : {}) }
}

/** Inventory (paged caps), G1, then G2 for an object candidate. */
export async function admitSpreadsheet(src: RangeSource, profile: ProfileId, limits: ImportLimits, signal: AbortSignal, options: VerifyOptions = {}): Promise<SpreadsheetAdmission> {
  let inventory: ZipInventory
  try {
    inventory = await readZipInventory(src, inventoryLimits(limits), signal)
  } catch (error) {
    if (!isPackageError(error) || error.code === 'cancelled') throw error
    return routed(failure(error.code, error.message))
  }
  const g1 = admit(profile, src.size, inventory, limits)
  if (g1.mode === 'refuse') return routed(g1.failure)
  if (g1.mode === 'paged') return g1
  const g2 = await verifyObjectCandidate(src, inventory, limits, signal, { profile, ...options })
  if (g2.mode === 'object') return { mode: 'object', decision: g2 }
  if (g2.mode === 'paged') return { mode: 'paged', reason: g2.reason, counters: g2.counters }
  return routed(g2.failure)
}
