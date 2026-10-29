import type { CacheIdentity, RemotePivotDisplay, RemoteRegionStyles, WorkerReply } from '../xlsx/view/contracts'
import type { RegionStyleLayer } from '../xlsx/tableStyles/regionLayer'
import type { CommandOf } from './viewerProtocol'
import { REGION_STYLE_LAYER_BASE_BYTES, MAX_REGION_STYLE_MEMO_ENTRIES, REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES, REGION_STYLE_REGION_ESTIMATE_BYTES } from '../xlsx/view/contracts'

const sameOwner = (a: CacheIdentity, b: CacheIdentity) => a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
const RETRY_MS = 250
const MAX_ATTEMPTS = 20
const READ_TIMEOUT_MS = 10_000
const layerBytesFor = (tableCount: number) => REGION_STYLE_LAYER_BASE_BYTES + tableCount * REGION_STYLE_REGION_ESTIMATE_BYTES + MAX_REGION_STYLE_MEMO_ENTRIES * REGION_STYLE_MEMO_ENTRY_ESTIMATE_BYTES
export interface PagedRegionStyleState {
  readonly state: 'waiting' | 'loading' | 'ready' | 'failed'
  readonly gen: number
  readonly payload?: RemoteRegionStyles | RemotePivotDisplay
  readonly layer?: RegionStyleLayer
  readonly diagnostics: (RemoteRegionStyles | RemotePivotDisplay)['diagnostics']
}
export interface RegionStyleHost {
  identity(sheetId: string): CacheIdentity
  hasRegionStyles(sheetId: string): boolean
  hasPivotDisplay?(sheetId: string): boolean
  loadState(sheetId: string): { completeness: 'complete' | 'loading' | 'incomplete' }
  grant(sheetId: string, requestId: string, minimumBytes?: number, op?: 'regionStyles' | 'pivotDisplay'): number | null
  releaseGrant(requestId: string): void
  request(command: CommandOf<'read'>, signal?: AbortSignal): Promise<WorkerReply<'read'> | symbol>
  nextRequestId(): string
  release(reply: WorkerReply<'read'>): void
  admitLayer(sheetId: string, requestId: string, bytes: number): (() => void) | null
  changed(sheetId: string, reason?: 'updated' | 'evicted'): void
}

/** Tables and pivots are separate bounded reads; their non-overlapping layers share one paint entry. */
export function mergeRegionStyleLayers(...layers: Array<RegionStyleLayer | undefined>): RegionStyleLayer | undefined {
  const retained = layers.filter((layer): layer is RegionStyleLayer => Boolean(layer))
  if (!retained.length) return undefined
  if (retained.length === 1) return retained[0]
  return {
    rects: retained.flatMap(layer => layer.rects),
    get retainedBytes() { return retained.reduce((sum, layer) => sum + layer.retainedBytes, 0) },
    styleAt(row, col, own, mode) { return retained.reduce((style, layer) => layer.styleAt(row, col, style, mode), own) },
  }
}
interface Entry { state: PagedRegionStyleState['state']; gen: number; owner?: CacheIdentity; seq: number; attempts: number; grantFloor?: number; sizeRetryAttempted?: boolean; payload?: RemoteRegionStyles | RemotePivotDisplay; layer?: RegionStyleLayer; layerReady?: boolean; reply?: WorkerReply<'read'>; releaseReply?: () => void; replyBytes: number; layerBytes: number; layerRelease?: () => void; controller?: AbortController; timer?: ReturnType<typeof setTimeout>; layerPending?: boolean; diagnostics: PagedRegionStyleState['diagnostics'] }
const empty = (): PagedRegionStyleState => ({ state: 'waiting', gen: 0, diagnostics: [] })

function unpack(reply: WorkerReply<'read'>, op: 'regionStyles' | 'pivotDisplay'): RemoteRegionStyles | RemotePivotDisplay | undefined {
  if (reply.kind !== 'complete' || reply.result.kind !== 'read' || reply.result.result.op !== op) return undefined
  return reply.result.result.result
}

/** Bounded receiver cache for the worker's per-sheet table descriptors. */
export class RemoteRegionStyleCache {
  private readonly entries = new Map<string, Entry>()
  private closed = false
  constructor(private readonly host: RegionStyleHost, readonly budgetBytes: number, private readonly op: 'regionStyles' | 'pivotDisplay' = 'regionStyles') {}
  get(sheetId: string): PagedRegionStyleState {
    if (!(this.op === 'pivotDisplay' ? this.host.hasPivotDisplay?.(sheetId) : this.host.hasRegionStyles(sheetId))) return empty()
    let entry = this.entries.get(sheetId)
    if (!entry) this.entries.set(sheetId, entry = { state: 'waiting', gen: 0, seq: 0, attempts: 0, replyBytes: 0, layerBytes: 0, diagnostics: [] })
    this.entries.delete(sheetId); this.entries.set(sheetId, entry)
    const owner = this.host.identity(sheetId)
    if (entry.owner && !sameOwner(entry.owner, owner)) { this.reset(entry); entry.gen++; this.deferChanged(sheetId) }
    if (entry.state === 'waiting' && !entry.timer && this.host.loadState(sheetId).completeness === 'complete') void this.load(sheetId, entry)
    else if (entry.state === 'ready' && entry.payload && (entry.payload.tables.length || ('pivotTables' in entry.payload && entry.payload.pivotTables.length)) && !entry.layerReady && !entry.layerPending) void this.rebuildLayer(sheetId, entry)
    return entry
  }
  get retainedEntries(): number { return [...this.entries.values()].filter(entry => entry.state === 'ready').length }
  get retainedBytes(): number { return [...this.entries.values()].reduce((n, e) => n + e.replyBytes + e.layerBytes, 0) }
  invalidate(sheetId: string): void { const entry = this.entries.get(sheetId); if (!entry) return; this.reset(entry); entry.gen++; this.deferChanged(sheetId) }
  closeSheet(sheetId: string): void { this.invalidate(sheetId); this.entries.delete(sheetId) }
  close(): void { if (this.closed) return; this.closed = true; for (const entry of this.entries.values()) this.reset(entry); this.entries.clear() }
  private reset(entry: Entry): void {
    entry.seq++; entry.controller?.abort(); entry.controller = undefined
    if (entry.timer) clearTimeout(entry.timer); entry.timer = undefined
    entry.releaseReply?.()
    entry.layerRelease?.(); entry.layerRelease = undefined
    entry.reply = undefined; entry.releaseReply = undefined; entry.layer = undefined; entry.layerReady = false; entry.payload = undefined; entry.owner = undefined; entry.replyBytes = 0; entry.layerBytes = 0; entry.layerPending = false; entry.attempts = 0
    entry.state = 'waiting'; entry.diagnostics = []; entry.grantFloor = undefined; entry.sizeRetryAttempted = false
  }
  private deferChanged(sheetId: string, reason: 'updated' | 'evicted' = 'updated'): void {
    queueMicrotask(() => { if (!this.closed) { try { this.host.changed(sheetId, reason) } catch { /* observer failures must not change cache ownership */ } } })
  }
  private retry(sheetId: string, entry: Entry, seq: number): void {
    if (entry.attempts >= MAX_ATTEMPTS) {
      entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('read-failed'), message: this.diagnosticMessage('read-failed', `${this.op === 'pivotDisplay' ? 'Pivot display' : 'Table style'} metadata could not be loaded after repeated attempts.`) }]
      this.deferChanged(sheetId); return
    }
    const delay = Math.min(4_000, RETRY_MS * 2 ** Math.max(0, entry.attempts - 1))
    entry.state = 'waiting'
    entry.timer = setTimeout(() => { entry.timer = undefined; if (!this.closed && entry.seq === seq) void this.load(sheetId, entry) }, delay)
  }
  private evictDecoded(sheetId: string, entry: Entry): void {
    entry.layerRelease?.(); entry.layerRelease = undefined
    entry.layer = undefined; entry.layerReady = false; entry.layerBytes = 0; entry.gen++
    this.deferChanged(sheetId, 'evicted')
  }
  private makeRoom(current: Entry, needed: number): boolean {
    for (const [id, other] of this.entries) {
      if (this.retainedBytes + needed <= this.budgetBytes) break
      if (other === current || other.state !== 'ready' || !other.layer) continue
      this.evictDecoded(id, other)
    }
    for (const [id, other] of this.entries) {
      if (this.retainedBytes + needed <= this.budgetBytes) break
      if (other === current || other.state !== 'ready') continue
      this.reset(other); other.gen++
      this.deferChanged(id, 'evicted')
    }
    return this.retainedBytes + needed <= this.budgetBytes
  }
  private diagnosticFeature(kind: 'read-failed' | 'reply-limit'): string {
    return this.op === 'pivotDisplay' ? `pivot-display-${kind}` : `table-style-${kind}`
  }
  private diagnosticMessage(kind: 'read-failed' | 'reply-limit', detail?: string): string {
    if (detail) return detail
    if (this.op === 'pivotDisplay') return kind === 'reply-limit' ? 'Pivot display metadata exceeds the receiver cache budget.' : 'Pivot display metadata could not be loaded.'
    return kind === 'reply-limit' ? 'Table style metadata exceeds the receiver cache budget.' : 'Table style metadata could not be loaded.'
  }
  private async buildLayer(sheetId: string, entry: Entry, requestId: string, payload: RemoteRegionStyles | RemotePivotDisplay, seq: number, reply?: WorkerReply<'read'>, releaseReply?: () => void): Promise<void> {
    const layerCount = payload.tables.length + ('pivotTables' in payload ? payload.pivotTables.length + payload.pivotStyles.length : 0)
    const layerBytes = layerCount ? layerBytesFor(layerCount) : 0
    if (layerBytes && !this.makeRoom(entry, layerBytes + (reply ? reply.replyBytes ?? 0 : 0))) {
      if (reply) releaseReply?.()
      entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('reply-limit'), message: this.diagnosticMessage('reply-limit') }]
      this.deferChanged(sheetId); return
    }
    let layerRelease: (() => void) | undefined
    try {
      let layer: RegionStyleLayer | undefined
      if (layerBytes) {
        layerRelease = this.host.admitLayer(sheetId, requestId, layerBytes) ?? undefined
        if (!layerRelease) {
          if (reply) releaseReply?.()
          entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('reply-limit'), message: this.diagnosticMessage('reply-limit', `${this.op === 'pivotDisplay' ? 'Pivot display' : 'Table style'} metadata exceeds the main-thread cache budget.`) }]
          this.deferChanged(sheetId); return
        }
        const { preloadTableStyles } = await import('../xlsx/tableStyleLoader')
        await preloadTableStyles()
        const { buildSheetRegionLayer } = await import('../xlsx/tableStyles/regionLayer')
        const tableStyles = new Map(payload.tables.map(item => [item.def.name, item.def]))
        if ('pivotStyles' in payload) for (const def of payload.pivotStyles) tableStyles.set(def.name, def)
        const slicerStyles = new Map()
        const extras = { dxfs: [], tableStyles, slicerStyles }
        const sheet = { name: sheetId, tables: payload.tables.map(item => item.table), ...('pivotTables' in payload ? { pivotTables: payload.pivotTables } : {}) } as import('../xlsx/types').XlsxSheet
        if (sheet.pivotTables?.length) await (await import('../xlsx/pivotLoader')).ensurePivotEngine([sheet])
        layer = buildSheetRegionLayer(sheet, extras, payload.theme ?? (await import('../drawing/style')).parseThemeContext())
      }
      if (this.closed || entry.seq !== seq || !sameOwner(entry.owner ?? this.host.identity(sheetId), this.host.identity(sheetId))) {
        layerRelease?.(); layerRelease = undefined; if (reply) releaseReply?.()
        if (!this.closed && entry.seq === seq) this.retry(sheetId, entry, seq)
        return
      }
      if (reply) { entry.reply = reply; entry.releaseReply = releaseReply; entry.replyBytes = reply.replyBytes ?? 0; entry.payload = payload; entry.owner = this.host.identity(sheetId) }
      entry.layer = layer; entry.layerBytes = layerBytes; entry.layerRelease = layerRelease; entry.layerReady = true; entry.layerPending = false
      layerRelease = undefined
      entry.diagnostics = payload.diagnostics; entry.state = 'ready'; entry.gen++; entry.attempts = 0; entry.grantFloor = undefined; entry.sizeRetryAttempted = false
      this.deferChanged(sheetId)
    } catch (cause) {
      layerRelease?.()
      if (reply) releaseReply?.()
      if (entry.reply === reply) { entry.reply = undefined; entry.releaseReply = undefined; entry.replyBytes = 0; entry.payload = undefined }
      entry.layer = undefined; entry.layerReady = false; entry.layerBytes = 0; entry.layerRelease = undefined; entry.layerPending = false
      entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('read-failed'), message: this.diagnosticMessage('read-failed', cause instanceof Error ? cause.message.slice(0, 256) : undefined) }]
      this.deferChanged(sheetId)
    }
  }
  private async rebuildLayer(sheetId: string, entry: Entry): Promise<void> {
    if (this.closed || !entry.payload || entry.layerPending) return
    entry.layerPending = true
    const seq = entry.seq, payload = entry.payload, requestId = `region-style-layer:${sheetId}:${seq}:${entry.gen}`
    await this.buildLayer(sheetId, entry, requestId, payload, seq)
  }
  private async load(sheetId: string, entry: Entry): Promise<void> {
    if (this.closed) return
    const requestId = this.host.nextRequestId(), owner = this.host.identity(sheetId), grantBytes = this.host.grant(sheetId, requestId, entry.grantFloor, this.op)
    entry.attempts++
    if (grantBytes === null) { this.retry(sheetId, entry, entry.seq); return }
    entry.grantFloor = undefined
    const seq = ++entry.seq, controller = entry.controller = new AbortController()
    entry.owner = owner; entry.state = 'loading'
    let out: WorkerReply<'read'> | symbol
    const timeout = setTimeout(() => controller.abort(), READ_TIMEOUT_MS)
    try { out = await this.host.request({ kind: 'read', requestId, owner, query: { op: this.op, grantBytes } } as CommandOf<'read'>, controller.signal) }
    catch (cause) { clearTimeout(timeout); this.host.releaseGrant(requestId); if (entry.seq === seq) this.retry(sheetId, entry, seq); return }
    clearTimeout(timeout)
    const releaseOnce = (() => { let released = false; return () => { if (released) return; released = true; if (typeof out === 'object') this.host.release(out) } })()
    if (this.closed || entry.seq !== seq || this.entries.get(sheetId) !== entry) { releaseOnce(); return }
    entry.controller = undefined
    if (typeof out !== 'object') { this.retry(sheetId, entry, seq); return }
    if (out.kind === 'cancelled' || out.kind === 'incomplete') {
      releaseOnce()
      this.retry(sheetId, entry, seq)
      return
    }
    if (out.kind === 'error') {
      releaseOnce()
      if (out.retryable || out.code === 'stale-source' || out.code === 'cancelled') this.retry(sheetId, entry, seq)
      else { entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('read-failed'), message: out.diagnostics[0]?.message ?? `${this.op === 'pivotDisplay' ? 'Pivot display' : 'Table style'} read failed (${out.code})` }]; this.deferChanged(sheetId) }
      return
    }
    if (out.kind === 'quota') {
      releaseOnce(); entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('reply-limit'), message: this.diagnosticMessage('reply-limit', `${this.op === 'pivotDisplay' ? 'Pivot display' : 'Table style'} metadata exceeds the receiver grant.`) }]; entry.gen++; this.deferChanged(sheetId); return
    }
    const payload = unpack(out, this.op)
    if (!payload) {
      releaseOnce()
      entry.state = 'failed'
      entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('read-failed'), message: this.diagnosticMessage('read-failed', `${this.op === 'pivotDisplay' ? 'Pivot display' : 'Table style'} read did not complete.`) }]
      entry.gen++; this.deferChanged(sheetId); return
    }
    if (!sameOwner(out.owner as CacheIdentity, owner) || !sameOwner(owner, this.host.identity(sheetId))
      || payload.sheetId !== sheetId || payload.stamp.docId !== owner.docId || payload.stamp.sheetId !== sheetId || payload.stamp.valueGen !== owner.valueGen) {
      releaseOnce(); this.retry(sheetId, entry, seq); return
    }
    if (payload.diagnostics.some(d => d.feature === 'table-style-reply-limit' || d.feature === 'pivot-display-reply-limit')) {
      const required = payload.requiredBytes
      releaseOnce()
      if (!entry.sizeRetryAttempted && Number.isSafeInteger(required) && required! > grantBytes) {
        entry.sizeRetryAttempted = true; entry.grantFloor = required; entry.state = 'waiting'
        void this.load(sheetId, entry)
        return
      }
      entry.state = 'failed'; entry.diagnostics = payload.diagnostics; entry.gen++; this.deferChanged(sheetId); return
    }
    if (!Number.isSafeInteger(out.replyBytes) || (out.replyBytes ?? -1) < 0) {
      releaseOnce(); entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('read-failed'), message: this.diagnosticMessage('read-failed', `${this.op === 'pivotDisplay' ? 'Pivot display' : 'Table style'} reply had no valid byte estimate.`) }]; this.deferChanged(sheetId); return
    }
    const layerCount = payload.tables.length + ('pivotTables' in payload ? payload.pivotTables.length + payload.pivotStyles.length : 0)
    const layerBytes = layerCount ? layerBytesFor(layerCount) : 0
    if (!this.makeRoom(entry, (out.replyBytes ?? 0) + layerBytes)) {
      releaseOnce(); entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: this.diagnosticFeature('reply-limit'), message: this.diagnosticMessage('reply-limit') }]; this.deferChanged(sheetId); return
    }
    await this.buildLayer(sheetId, entry, requestId, payload, seq, out, releaseOnce)
  }
}

export function sheetWithRegionStyles<T extends { tables?: unknown[]; regionStyles?: RegionStyleLayer; regionStylesResolved?: boolean; diagnostics?: Array<{ kind: string; feature: string; message: string }> }>(sheet: T, state: PagedRegionStyleState): T {
  if (state.state !== 'ready' && state.state !== 'failed') return sheet
  const diagnostics = [...(sheet.diagnostics ?? [])]
  for (const diagnostic of state.diagnostics) if (!diagnostics.some(found => found.feature === diagnostic.feature && found.message === diagnostic.message)) diagnostics.push(diagnostic)
  if (state.state !== 'ready' || !state.payload) return { ...sheet, diagnostics } as T
  const tables = 'pivotTables' in state.payload ? (sheet.tables ?? []) : state.payload.tables.map(item => item.table)
  const priorPivotTables = (sheet as T & { pivotTables?: unknown[] }).pivotTables
  const pivotTables = 'pivotTables' in state.payload ? state.payload.pivotTables : priorPivotTables
  return { ...sheet, tables, ...(pivotTables ? { pivotTables } : {}), regionStyles: mergeRegionStyleLayers(sheet.regionStyles, state.layer), regionStylesResolved: true, diagnostics } as T
}

/** Drop the display entry's reference when its decoded role cache is evicted. */
export function sheetWithoutRegionStyles<T extends { regionStyles?: RegionStyleLayer }>(sheet: T): T {
  if (!sheet.regionStyles) return sheet
  return { ...sheet, regionStyles: undefined } as T
}
