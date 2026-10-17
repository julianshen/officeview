/**
 * Resident XLSX viewer worker wire protocol (spec §10.2, plan S4.1).
 *
 * Main-safe: types, message guards, typed reply builders, the compact
 * BoundedRowBitset adapters and the ArrayBuffer copy/transfer ledger. Nothing
 * here imports the parser or touches worker globals, so the client can bundle
 * it without pulling the resident session in.
 */
import type {
  BoundedRowBitset,
  CacheIdentity,
  WorkbookOwner,
  OperationRefusal,
  WorkerCommand,
  WorkerCommandKind,
  WorkerCommandOwner,
  WorkerDiagnostic,
  WorkerReply,
  WorkerReplyBase,
} from '../xlsx/view/contracts'
import { OPERATION_REFUSED_CODE } from '../xlsx/view/contracts'
import { MAX_ROWS, validateCoordinate } from '../xlsx/view/coordinates'
import { CompactRowVisibility } from '../xlsx/view/visibility'

export type {
  WorkerCommand, WorkerCommandKind, WorkerReply, WorkerResult, WorkerOwner, WorkerDiagnostic,
  ViewCommitDelta, CacheIdentity, WorkbookOwner, SourceHandle, BoundedRowBitset,
} from '../xlsx/view/contracts'

export const VIEWER_COMMAND_KINDS: readonly WorkerCommandKind[] = Object.freeze([
  'open', 'activateSheet', 'prepareColumns', 'sort', 'filter', 'viewport', 'valueQuery', 'search', 'read', 'cancel', 'close',
])
const COMMAND_KINDS: ReadonlySet<string> = new Set(VIEWER_COMMAND_KINDS)
const REPLY_KINDS: ReadonlySet<string> = new Set(['progress', 'complete', 'cancelled', 'incomplete', 'quota', 'error'])
const TERMINAL_KINDS: ReadonlySet<string> = new Set(['complete', 'cancelled', 'incomplete', 'quota', 'error'])

/** Viewer commands carry a string requestId and a viewer kind; legacy render
 * requests use numeric ids and 'render-*' kinds, so one worker serves both. */
export function isViewerCommand(data: unknown): data is WorkerCommand {
  if (typeof data !== 'object' || data === null) return false
  const { kind, requestId } = data as { kind?: unknown; requestId?: unknown }
  return typeof requestId === 'string' && typeof kind === 'string' && COMMAND_KINDS.has(kind)
}
export function isViewerReply(data: unknown): data is WorkerReply {
  if (typeof data !== 'object' || data === null) return false
  const { kind, requestId, commandKind } = data as { kind?: unknown; requestId?: unknown; commandKind?: unknown }
  return typeof requestId === 'string' && typeof kind === 'string' && REPLY_KINDS.has(kind)
    && typeof commandKind === 'string' && COMMAND_KINDS.has(commandKind)
}
/** Every reply kind except progress ends its request. */
export function isTerminalReply(reply: WorkerReply): boolean {
  return TERMINAL_KINDS.has(reply.kind)
}

/** Protocol maximum requestId length (UTF-16 units): it bounds every reply's
 * identity part, and so the payload-free envelope of a drawings grant. */
export const MAX_REQUEST_ID_UNITS = 256
/** Why an in-flight request was aborted; echoed as the cancelled diagnostic. */
export type CancelReason = 'cancel-requested' | 'superseded' | 'range-replaced' | 'document-closed' | 'query-timeout'
export const CANCEL_MESSAGES: Readonly<Record<CancelReason, string>> = Object.freeze({
  'cancel-requested': 'Cancelled by request',
  superseded: 'Superseded by a newer view draft on the same sheet',
  'range-replaced': 'Replaced by a newer query on the same column',
  'document-closed': 'The document was closed',
  'query-timeout': 'Query timed out after 120 seconds. Retry or choose a smaller range.',
})

/** Protocol convention: the values published by open are generation 0 for every
 * sheet, view and font, so a client derives its first owners from the open
 * reply's sheet list. Every later owner arrives in a committed view delta. */
export const INITIAL_GENERATION = 0
export function initialIdentity(docId: string, sheetId: string): CacheIdentity {
  return { docId, sheetId, valueGen: INITIAL_GENERATION, fontGen: INITIAL_GENERATION, viewGen: INITIAL_GENERATION }
}
export function initialWorkbookOwner(docId: string, sheets: readonly { sheetId: string }[]): WorkbookOwner {
  return {
    docId, globalViewGen: INITIAL_GENERATION, workbookValueGen: INITIAL_GENERATION, fontGen: INITIAL_GENERATION,
    sheetVector: sheets.map(({ sheetId }) => ({ sheetId, revision: INITIAL_GENERATION, valueGen: INITIAL_GENERATION })),
  }
}

type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never
export type CommandOf<K extends WorkerCommandKind> = Extract<WorkerCommand, { kind: K }>
/** Reply payload without the base fields every reply shares. */
export type ReplyBody<K extends WorkerCommandKind> = DistributiveOmit<WorkerReply<K>, keyof WorkerReplyBase<K>>

/** The only reply constructor: commandKind always echoes the originating command
 * (S0.1 freeze addition to §10.2) and owner is the command's captured owner. */
export function reply<K extends WorkerCommandKind>(
  command: CommandOf<K>,
  owner: WorkerCommandOwner<K>,
  body: ReplyBody<K>,
  diagnostics: readonly WorkerDiagnostic[] = [],
): WorkerReply<K> {
  // The conditional reply union cannot be narrowed through a generic K.
  return { ...body, commandKind: command.kind, requestId: command.requestId, owner, diagnostics } as unknown as WorkerReply<K>
}
/** Guarded refusal (§6.3): only sort/filter carry operation-refused, never retryable. */
export function refusalReply<K extends 'sort' | 'filter'>(
  command: CommandOf<K>,
  owner: WorkerCommandOwner<K>,
  refusal: OperationRefusal,
  diagnostics: readonly WorkerDiagnostic[] = [],
): WorkerReply<K> {
  const body = { kind: 'error', code: OPERATION_REFUSED_CODE, retryable: false, refusal } as unknown as ReplyBody<K>
  return reply(command, owner, body, diagnostics)
}

/** ArrayBuffer accounting (§10.2): 'typed' never means zero-copy. A copy is a
 * new buffer filled from resident data; a transfer moves a buffer's ownership
 * through postMessage. Resident buffers are never transferred. */
export interface TransferLedger {
  copies: number
  copiedBytes: number
  transfers: number
  transferredBytes: number
  /** Portion of copiedBytes that is an estimate (structured clone of plain
   * object graphs such as an object source document), not an exact byteLength. */
  estimatedCopiedBytes: number
}
export class TransferAccounting {
  private state: TransferLedger = { copies: 0, copiedBytes: 0, transfers: 0, transferredBytes: 0, estimatedCopiedBytes: 0 }
  copy(bytes: number, estimated = false): void {
    this.state.copies++
    this.state.copiedBytes += bytes
    if (estimated) this.state.estimatedCopiedBytes += bytes
  }
  transfer(buffers: readonly ArrayBuffer[]): void {
    for (const buffer of buffers) {
      this.state.transfers++
      this.state.transferredBytes += buffer.byteLength
    }
  }
  snapshot(): TransferLedger { return { ...this.state } }
}
/** Realm-independent tag checks: structured clones arrive from another realm
 * (worker or test harness), where instanceof against local constructors fails. */
export const isArrayBuffer = (value: unknown): value is ArrayBuffer =>
  Object.prototype.toString.call(value) === '[object ArrayBuffer]'
export const isUint32Array = (value: unknown): value is Uint32Array =>
  Object.prototype.toString.call(value) === '[object Uint32Array]'
export const isUint16Array = (value: unknown): value is Uint16Array =>
  Object.prototype.toString.call(value) === '[object Uint16Array]'
/** Distinct ArrayBuffers behind every typed array in a message. Bounded walk over
 * plain arrays/objects only; replies never carry Maps or class instances. */
export function transferablesOf(message: unknown): ArrayBuffer[] {
  const found = new Set<ArrayBuffer>()
  const seen = new Set<object>()
  const visit = (value: unknown): void => {
    if (typeof value !== 'object' || value === null || seen.has(value)) return
    seen.add(value)
    if (ArrayBuffer.isView(value)) {
      if (isArrayBuffer(value.buffer)) found.add(value.buffer)
      return
    }
    if (isArrayBuffer(value)) { found.add(value); return }
    for (const entry of Array.isArray(value) ? value : Object.values(value)) visit(entry)
  }
  visit(message)
  return [...found]
}
/** Every BoundedRowBitset in a message, for space-tag audits. */
export function bitsetsOf(message: unknown): BoundedRowBitset<'source' | 'display'>[] {
  const out: BoundedRowBitset<'source' | 'display'>[] = []
  const seen = new Set<object>()
  const visit = (value: unknown): void => {
    if (typeof value !== 'object' || value === null || seen.has(value) || ArrayBuffer.isView(value)) return
    seen.add(value)
    if ((value as { kind?: unknown }).kind === 'bitset') out.push(value as BoundedRowBitset<'source' | 'display'>)
    for (const entry of Array.isArray(value) ? value : Object.values(value)) visit(entry)
  }
  visit(message)
  return out
}

function popcount(word: number): number {
  word -= (word >>> 1) & 0x55555555
  word = (word & 0x33333333) + ((word >>> 2) & 0x33333333)
  return (((word + (word >>> 4)) & 0x0f0f0f0f) * 0x01010101) >>> 24
}
/** Mask of the low `bits` bits (1..32) without 32-bit shift overflow. */
const lowBits = (bits: number): number => bits >= 32 ? 0xffffffff : 2 ** bits - 1

/** Encodes a display visibility mask as bounded bitsets. Wire convention:
 * display slots outside every bitset are visible, so a fully visible sheet sends
 * none. One bitset spans the first to last hidden slot, starting on a word
 * boundary so its words are one contiguous copy of the resident mask. */
export function displayBitsetsFromVisibility(
  visibility: CompactRowVisibility,
  ledger?: TransferAccounting,
): BoundedRowBitset<'display'>[] {
  const { words, logicalCount } = visibility
  let firstWord = -1, lastWord = -1
  for (let word = 0; word < words.length; word++) {
    const valid = word === words.length - 1 && logicalCount % 32 ? lowBits(logicalCount % 32) : 0xffffffff
    if ((words[word] >>> 0) !== valid) {
      if (firstWord < 0) firstWord = word
      lastWord = word
    }
  }
  if (firstWord < 0) return []
  // Bits past logicalCount are zero padding, not hidden rows.
  let lastHidden = Math.min(lastWord * 32 + 31, logicalCount - 1)
  while (lastHidden >= lastWord * 32 && ((words[lastWord] >>> (lastHidden % 32)) & 1)) lastHidden--
  const r0 = firstWord * 32
  const out = words.slice(firstWord, lastWord + 1)
  ledger?.copy(out.byteLength)
  // Padding bits past r1 are zero by contract.
  out[out.length - 1] &= lowBits(lastHidden % 32 + 1)
  let count = 0
  for (const word of out) count += popcount(word)
  return [Object.freeze({ kind: 'bitset', space: 'display', r0, r1: lastHidden, words: out, count })]
}
/** Source-space membership bitset over [r0,r1]; `rows` are absolute source rows. */
export function sourceBitset(r0: number, r1: number, rows: Iterable<number>): BoundedRowBitset<'source'> {
  validateCoordinate(r0, 0)
  validateCoordinate(r1, 0)
  if (r0 > r1) throw new RangeError('Reversed bitset bounds')
  const words = new Uint32Array(Math.ceil((r1 - r0 + 1) / 32))
  let count = 0
  for (const row of rows) {
    validateCoordinate(row, 0)
    if (row < r0 || row > r1) throw new RangeError('Bitset row outside bounds')
    const offset = row - r0, word = Math.floor(offset / 32), bit = 2 ** (offset % 32)
    if (!(words[word] & bit)) { words[word] |= bit; count++ }
  }
  return Object.freeze({ kind: 'bitset', space: 'source', r0, r1, words, count })
}
/** Validates the shared mask encoding, including the explicit space tag. */
export function validateBitset<Space extends 'source' | 'display'>(
  bitset: BoundedRowBitset<'source' | 'display'>,
  space: Space,
): asserts bitset is BoundedRowBitset<Space> {
  if (bitset.kind !== 'bitset' || bitset.space !== space) throw new RangeError(`Bitset space must be ${space}`)
  validateCoordinate(bitset.r0, 0)
  validateCoordinate(bitset.r1, 0)
  if (bitset.r0 > bitset.r1) throw new RangeError('Reversed bitset bounds')
  const length = bitset.r1 - bitset.r0 + 1
  if (!isUint32Array(bitset.words) || bitset.words.length !== Math.ceil(length / 32)) {
    throw new RangeError('Invalid bitset word count')
  }
  const tail = length % 32
  if (tail && (bitset.words[bitset.words.length - 1] & ~lowBits(tail)) !== 0) throw new RangeError('Nonzero bitset padding')
  let count = 0
  for (const word of bitset.words) count += popcount(word)
  if (count !== bitset.count) throw new RangeError('Bitset count differs from popcount')
}
/** Client-side inverse of displayBitsetsFromVisibility. */
export function visibilityFromDisplayBitsets(
  bitsets: readonly BoundedRowBitset<'display'>[],
  logicalCount = MAX_ROWS,
): CompactRowVisibility {
  const words = new Uint32Array(Math.ceil(logicalCount / 32)).fill(0xffffffff)
  for (const bitset of bitsets) {
    validateBitset(bitset, 'display')
    if (bitset.r1 >= logicalCount) throw new RangeError('Bitset outside the display axis')
    for (let slot = bitset.r0; slot <= bitset.r1; slot++) {
      const offset = slot - bitset.r0
      const visible = (bitset.words[Math.floor(offset / 32)] >>> (offset % 32)) & 1
      const word = Math.floor(slot / 32), bit = 2 ** (slot % 32)
      words[word] = visible ? (words[word] | bit) >>> 0 : (words[word] & ~bit) >>> 0
    }
  }
  return new CompactRowVisibility(logicalCount, words)
}
