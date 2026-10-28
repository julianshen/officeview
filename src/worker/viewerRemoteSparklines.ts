/** Worker-resolved sparklines for paged XLSX sheets (S3; spec §6.3). */
import type { CacheIdentity, SheetSparklines, WorkerReply } from '../xlsx/view/contracts'
import type { SparklineDiagnostic } from '../xlsx/sparklines/model'
import type { XlsxDocument, XlsxSheet } from '../xlsx/types'
import type { ResolvedSparkline, ResolvedSparklines } from '../xlsx/sparklines/resolve'
import type { CommandOf } from './viewerProtocol'

const RETRY_MS = 250
const MAX_ATTEMPTS = 20
const sameOwner = (a: CacheIdentity, b: CacheIdentity) => a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
export interface SparklineHost {
  identity(sheetId: string): CacheIdentity
  loadState(sheetId: string): { completeness: 'complete' | 'loading' | 'incomplete' }
  hasSparklines(sheetId: string): boolean
  /** Reserve receiver bytes synchronously, before a request can leave. */
  grant(requestId: string): number | null
  releaseGrant(requestId: string): void
  /** Reserve the main-thread decoded arrays/maps before rebuilding them. */
  admitDecoded(sheetId: string, requestId: string, bytes: number): (() => void) | null
  request(command: CommandOf<'read'>, signal?: AbortSignal): Promise<WorkerReply<'read'> | symbol>
  nextRequestId(): string
  release(reply: WorkerReply<'read'>): void
  changed(sheetId: string, hosts: readonly { row: number; col: number }[]): void
}
export interface PagedSheetSparklines {
  readonly state: 'waiting' | 'loading' | 'ready' | 'failed'
  readonly gen: number
  readonly resolved?: ResolvedSparklines
  readonly diagnostics: readonly SparklineDiagnostic[]
  /** Referenced pages which may still be importing. */
  readonly pendingSheetIds?: readonly string[]
}
interface Entry {
  state: PagedSheetSparklines['state']; gen: number; owner?: CacheIdentity; seq: number; bytes: number; decodedBytes: number; attempts: number
  resolved?: ResolvedSparklines; diagnostics: SparklineDiagnostic[]; reply?: WorkerReply<'read'>
  decodedRelease?: () => void; controller?: AbortController; timer?: ReturnType<typeof setTimeout>
}
const empty = (): PagedSheetSparklines => ({ state: 'waiting', gen: 0, diagnostics: [] })

/** JS arrays and map/object records materialized alongside the retained wire
 * reply. Reserve before Array.from/Map construction; point buffers themselves
 * remain charged by the reply lease. */
function decodedBytesOf(result: SheetSparklines): number {
  if (!result.sparklines.length) return 0
  let bytes = 128
  for (const spark of result.sparklines) {
    bytes += 512 + spark.values.length * 8 + (spark.axis?.positions.length ?? 0) * 8
  }
  // One row Map per spark is a safe upper bound without allocating a Set
  // before the receiver has admitted the decoded cache.
  return bytes + result.sparklines.length * 128
}

function unpack(reply: WorkerReply<'read'>): SheetSparklines | undefined {
  if (reply.kind !== 'complete' || reply.result.kind !== 'read' || reply.result.result.op !== 'sparklines') return undefined
  return reply.result.result.result
}
export function rebuildResolvedSparklines(result: SheetSparklines): ResolvedSparklines | undefined {
  const rows = new Map<number, Map<number, ResolvedSparkline>>()
  for (const transfer of result.sparklines) {
    const group = result.groups[transfer.groupIndex]
    if (!group || transfer.values.length !== transfer.sourceRows.length || (transfer.axis && transfer.axis.positions.length !== transfer.values.length)) continue
    const values = Array.from(transfer.values, value => Number.isNaN(value) ? null : value)
    let row = rows.get(transfer.host.row)
    if (!row) rows.set(transfer.host.row, row = new Map())
    if (row.has(transfer.host.col)) continue
    row.set(transfer.host.col, {
      host: transfer.host, dataSheet: transfer.dataSheet, sourceRows: transfer.sourceRows, group, values,
      bridged: transfer.bridged, firstIndex: transfer.firstIndex, lastIndex: transfer.lastIndex,
      highIndices: transfer.highIndices, lowIndices: transfer.lowIndices, min: transfer.min, max: transfer.max,
      ...(transfer.axis ? { axis: { positions: Array.from(transfer.axis.positions), slot: transfer.axis.slot } } : {}),
    })
  }
  return rows.size ? rows : undefined
}

export class RemoteSparklineCache {
  private readonly entries = new Map<string, Entry>()
  private closed = false
  constructor(private readonly host: SparklineHost, readonly budgetBytes = Number.MAX_SAFE_INTEGER) {}
  get(sheetId: string): PagedSheetSparklines {
    if (!this.host.hasSparklines(sheetId)) return empty()
    let entry = this.entries.get(sheetId)
    if (!entry) this.entries.set(sheetId, entry = { state: 'waiting', gen: 0, seq: 0, bytes: 0, decodedBytes: 0, attempts: 0, diagnostics: [] })
    this.entries.delete(sheetId); this.entries.set(sheetId, entry)
    const owner = this.host.identity(sheetId)
    if (entry.owner && !sameOwner(entry.owner, owner)) {
      const hosts = entry.resolved ? [...entry.resolved.values()].flatMap(row => [...row.values()].map(spark => spark.host)) : []
      this.reset(entry); entry.gen++
      if (hosts.length) this.host.changed(sheetId, hosts)
    }
    if (entry.state === 'waiting' && !entry.timer && this.host.loadState(sheetId).completeness === 'complete') void this.load(sheetId, entry)
    return entry
  }
  get retainedEntries(): number { return [...this.entries.values()].filter(entry => entry.state === 'ready').length }
  get retainedBytes(): number { return [...this.entries.values()].reduce((sum, entry) => sum + entry.bytes + entry.decodedBytes, 0) }
  /** Invalidates values on the host sheet; caller invokes it on a new value owner. */
  invalidate(sheetId: string): void {
    const entry = this.entries.get(sheetId)
    if (!entry) return
    const hosts = entry.resolved ? [...entry.resolved.values()].flatMap(row => [...row.values()].map(spark => spark.host)) : []
    this.reset(entry); entry.gen++
    if (hosts.length) this.host.changed(sheetId, hosts)
  }
  close(): void {
    if (this.closed) return
    this.closed = true
    for (const entry of this.entries.values()) this.reset(entry)
    this.entries.clear()
  }
  private reset(entry: Entry): void {
    entry.seq++; entry.controller?.abort(); entry.controller = undefined
    if (entry.timer) clearTimeout(entry.timer); entry.timer = undefined
    if (entry.reply) this.host.release(entry.reply)
    entry.decodedRelease?.(); entry.decodedRelease = undefined
    entry.reply = undefined; entry.resolved = undefined; entry.diagnostics = []; entry.owner = undefined; entry.state = 'waiting'; entry.bytes = 0; entry.decodedBytes = 0; entry.attempts = 0
  }
  private async load(sheetId: string, entry: Entry): Promise<void> {
    if (this.closed) return
    const requestId = this.host.nextRequestId(), owner = this.host.identity(sheetId), grantBytes = this.host.grant(requestId), waitingSeq = entry.seq
    if (grantBytes === null) {
      entry.state = 'waiting'
      entry.timer = setTimeout(() => { entry.timer = undefined; if (entry.seq === waitingSeq) void this.load(sheetId, entry) }, RETRY_MS)
      return
    }
    const seq = ++entry.seq, controller = entry.controller = new AbortController()
    entry.owner = owner; entry.state = 'loading'; entry.attempts++
    let pending: Promise<WorkerReply<'read'> | symbol>
    try { pending = this.host.request({ kind: 'read', requestId, owner, query: { op: 'sparklines', grantBytes } }, controller.signal) }
    catch (cause) {
      this.host.releaseGrant(requestId); entry.controller = undefined; entry.state = 'failed'
      entry.diagnostics = [{ kind: 'unsupported', feature: 'sparkline-read-failed', message: cause instanceof Error ? cause.message.slice(0, 256) : 'Sparkline request could not be sent' }]
      return
    }
    const out = await pending
    if (this.closed || entry.seq !== seq || this.entries.get(sheetId) !== entry) {
      if (typeof out === 'object') this.host.release(out)
      return
    }
    entry.controller = undefined
    if (typeof out !== 'object') { entry.state = 'waiting'; return }
    const data = unpack(out)
    const owned = sameOwner(out.owner as CacheIdentity, owner) && sameOwner(owner, this.host.identity(sheetId))
    if (!owned) { this.host.release(out); entry.state = 'waiting'; return }
    if (!data && out.kind !== 'quota') {
      this.host.release(out)
      if (out.kind === 'error' && !out.retryable) {
        entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: 'sparkline-read-failed', message: out.diagnostics[0]?.message ?? `Sparkline read failed (${out.code})` }]
      } else if (entry.attempts < MAX_ATTEMPTS) this.retry(sheetId, entry, seq)
      else { entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: 'sparkline-read-failed', message: 'Sparkline read could not complete' }] }
      return
    }
    if (data?.diagnostics.some(item => item.feature === 'sparkline-data-unavailable')) {
      this.host.release(out)
      const pending = data.pendingSheetIds ?? []
      if (pending.some(id => this.host.loadState(id).completeness === 'incomplete')) {
        entry.state = 'failed'
        entry.diagnostics = [data.diagnostics.find(item => item.feature === 'sparkline-data-unavailable') ?? data.diagnostics[0] ?? { kind: 'unsupported', feature: 'sparkline-data-unavailable', message: 'Sparkline data is incomplete' }]
      } else if (pending.some(id => this.host.loadState(id).completeness === 'loading')) this.retry(sheetId, entry, seq)
      else if (entry.attempts < MAX_ATTEMPTS) this.retry(sheetId, entry, seq)
      else { entry.state = 'failed'; entry.diagnostics = [data.diagnostics.find(item => item.feature === 'sparkline-data-unavailable') ?? data.diagnostics[0] ?? { kind: 'unsupported', feature: 'sparkline-data-unavailable', message: 'Sparkline data is unavailable' }] }
      return
    }
    const result = data ?? { sparklines: [], groups: [], diagnostics: [{ kind: 'unsupported' as const, feature: 'sparkline-resource-limit', message: 'Sparkline reply exceeded its receiver grant' }] }
    // A quota reply is only an admission decision, never retained cache data.
    // Its grant-backed ledger lease must end as soon as the refusal is consumed.
    if (out.kind === 'quota') this.host.release(out)
    const bytes = out.kind !== 'quota' && Number.isSafeInteger(out.replyBytes) && (out.replyBytes ?? 0) >= 0 ? out.replyBytes! : 0
    const decodedBytes = decodedBytesOf(result)
    if (bytes + decodedBytes > this.budgetBytes) {
      this.host.release(out); entry.state = 'failed'
      entry.diagnostics = [{ kind: 'unsupported', feature: 'sparkline-resource-limit', message: 'Sparkline data exceeds the main-thread cache budget' }]
      return
    }
    for (const [id, other] of this.entries) {
      if (this.retainedBytes + bytes + decodedBytes <= this.budgetBytes) break
      if (other === entry || other.state !== 'ready') continue
      const oldHosts = other.resolved ? [...other.resolved.values()].flatMap(row => [...row.values()].map(spark => spark.host)) : []
      this.reset(other); other.gen++
      if (oldHosts.length) this.host.changed(id, oldHosts)
    }
    const decodedRelease = decodedBytes > 0 ? this.host.admitDecoded(sheetId, requestId, decodedBytes) : undefined
    if (decodedBytes > 0 && !decodedRelease) {
      this.host.release(out); entry.state = 'failed'
      entry.diagnostics = [{ kind: 'unsupported', feature: 'sparkline-resource-limit', message: 'Sparkline data exceeds the main-thread memory budget' }]
      return
    }
    entry.resolved = rebuildResolvedSparklines(result)
    entry.diagnostics = result.diagnostics
    entry.reply = out.kind === 'quota' ? undefined : out
    entry.bytes = bytes
    entry.decodedBytes = decodedBytes
    entry.decodedRelease = decodedRelease ?? undefined
    entry.state = out.kind === 'quota' ? 'failed' : 'ready'
    entry.gen++
    const hosts = result.sparklines.map(spark => spark.host)
    if (hosts.length) this.host.changed(sheetId, hosts)
  }
  private retry(sheetId: string, entry: Entry, seq: number): void {
    entry.state = 'waiting'
    entry.timer = setTimeout(() => { entry.timer = undefined; if (entry.seq === seq) void this.load(sheetId, entry) }, RETRY_MS)
  }
}

export function sheetWithSparklines(base: XlsxSheet, content: PagedSheetSparklines): XlsxSheet {
  if (content.state !== 'ready' || !content.resolved) return base
  const diagnostics = [...(base.diagnostics ?? [])]
  for (const item of content.diagnostics) if (!diagnostics.some(found => found.feature === item.feature && found.message === item.message)) diagnostics.push(item)
  return { ...base, resolvedSparklines: content.resolved, diagnostics }
}
export function documentWithSparklines(doc: XlsxDocument, index: number, content: PagedSheetSparklines): XlsxDocument {
  if (content.state !== 'ready' || !doc.sheets[index]) return doc
  return { ...doc, sheets: doc.sheets.map((sheet, i) => i === index ? sheetWithSparklines(sheet, content) : sheet) }
}
export function sparklineNotices(content: PagedSheetSparklines): string[] {
  return [...new Set(content.diagnostics.map(item => item.message))].slice(0, 4)
}
