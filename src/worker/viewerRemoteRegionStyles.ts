import type { CacheIdentity, RemoteRegionStyles, WorkerReply } from '../xlsx/view/contracts'
import type { RegionStyleLayer } from '../xlsx/tableStyles/regionLayer'
import type { CommandOf } from './viewerProtocol'

const sameOwner = (a: CacheIdentity, b: CacheIdentity) => a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
export interface PagedRegionStyleState {
  readonly state: 'waiting' | 'loading' | 'ready' | 'failed'
  readonly gen: number
  readonly payload?: RemoteRegionStyles
  readonly layer?: RegionStyleLayer
  readonly diagnostics: RemoteRegionStyles['diagnostics']
}
export interface RegionStyleHost {
  identity(sheetId: string): CacheIdentity
  hasRegionStyles(sheetId: string): boolean
  loadState(sheetId: string): { completeness: 'complete' | 'loading' | 'incomplete' }
  grant(requestId: string): number | null
  releaseGrant(requestId: string): void
  request(command: CommandOf<'read'>, signal?: AbortSignal): Promise<WorkerReply<'read'> | symbol>
  nextRequestId(): string
  release(reply: WorkerReply<'read'>): void
  admitLayer(sheetId: string, requestId: string, bytes: number): (() => void) | null
  changed(sheetId: string): void
}
interface Entry { state: PagedRegionStyleState['state']; gen: number; owner?: CacheIdentity; seq: number; payload?: RemoteRegionStyles; layer?: RegionStyleLayer; reply?: WorkerReply<'read'>; replyBytes: number; layerBytes: number; layerRelease?: () => void; controller?: AbortController; diagnostics: RemoteRegionStyles['diagnostics'] }
const empty = (): PagedRegionStyleState => ({ state: 'waiting', gen: 0, diagnostics: [] })

function unpack(reply: WorkerReply<'read'>): RemoteRegionStyles | undefined {
  if (reply.kind !== 'complete' || reply.result.kind !== 'read' || reply.result.result.op !== 'regionStyles') return undefined
  return reply.result.result.result
}

/** Bounded receiver cache for the worker's per-sheet table descriptors. */
export class RemoteRegionStyleCache {
  private readonly entries = new Map<string, Entry>()
  private closed = false
  constructor(private readonly host: RegionStyleHost, readonly budgetBytes: number) {}
  get(sheetId: string): PagedRegionStyleState {
    if (!this.host.hasRegionStyles(sheetId)) return empty()
    let entry = this.entries.get(sheetId)
    if (!entry) this.entries.set(sheetId, entry = { state: 'waiting', gen: 0, seq: 0, replyBytes: 0, layerBytes: 0, diagnostics: [] })
    this.entries.delete(sheetId); this.entries.set(sheetId, entry)
    const owner = this.host.identity(sheetId)
    if (entry.owner && !sameOwner(entry.owner, owner)) { this.reset(entry); entry.gen++ }
    if (entry.state === 'waiting' && this.host.loadState(sheetId).completeness === 'complete') void this.load(sheetId, entry)
    return entry
  }
  get retainedEntries(): number { return [...this.entries.values()].filter(entry => entry.state === 'ready').length }
  get retainedBytes(): number { return [...this.entries.values()].reduce((n, e) => n + e.replyBytes + e.layerBytes, 0) }
  invalidate(sheetId: string): void { const entry = this.entries.get(sheetId); if (!entry) return; this.reset(entry); entry.gen++; this.host.changed(sheetId) }
  closeSheet(sheetId: string): void { this.invalidate(sheetId); this.entries.delete(sheetId) }
  close(): void { if (this.closed) return; this.closed = true; for (const entry of this.entries.values()) this.reset(entry); this.entries.clear() }
  private reset(entry: Entry): void {
    entry.seq++; entry.controller?.abort(); entry.controller = undefined
    if (entry.reply) this.host.release(entry.reply)
    entry.layerRelease?.(); entry.layerRelease = undefined
    entry.reply = undefined; entry.layer = undefined; entry.payload = undefined; entry.owner = undefined; entry.replyBytes = 0; entry.layerBytes = 0
    entry.state = 'waiting'; entry.diagnostics = []
  }
  private async load(sheetId: string, entry: Entry): Promise<void> {
    if (this.closed) return
    const requestId = this.host.nextRequestId(), owner = this.host.identity(sheetId), grantBytes = this.host.grant(requestId)
    if (grantBytes === null) { entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: 'table-style-reply-limit', message: 'Table style metadata exceeds available receiver memory.' }]; this.host.changed(sheetId); return }
    const seq = ++entry.seq, controller = entry.controller = new AbortController()
    entry.owner = owner; entry.state = 'loading'
    let out: WorkerReply<'read'> | symbol
    try { out = await this.host.request({ kind: 'read', requestId, owner, query: { op: 'regionStyles', grantBytes } }, controller.signal) }
    catch (cause) { this.host.releaseGrant(requestId); entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: 'table-style-read-failed', message: cause instanceof Error ? cause.message.slice(0, 256) : 'Table style read failed' }]; this.host.changed(sheetId); return }
    if (this.closed || entry.seq !== seq || this.entries.get(sheetId) !== entry) { if (typeof out === 'object') this.host.release(out); return }
    entry.controller = undefined
    if (typeof out !== 'object') { entry.state = 'waiting'; return }
    const payload = unpack(out)
    if (!payload) {
      this.host.release(out)
      entry.state = 'failed'
      entry.diagnostics = [{ kind: 'unsupported', feature: out.kind === 'quota' ? 'table-style-reply-limit' : 'table-style-read-failed', message: out.kind === 'quota' ? 'Table style metadata exceeds available receiver memory.' : out.kind === 'error' ? out.diagnostics[0]?.message ?? `Table style read failed (${out.code})` : 'Table style read did not complete.' }]
      entry.gen++; this.host.changed(sheetId); return
    }
    if (!sameOwner(out.owner as CacheIdentity, owner) || !sameOwner(owner, this.host.identity(sheetId))
      || payload.sheetId !== sheetId || payload.stamp.docId !== owner.docId || payload.stamp.sheetId !== sheetId || payload.stamp.valueGen !== owner.valueGen) {
      this.host.release(out); entry.state = 'waiting'; this.host.changed(sheetId); return
    }
    if (payload.diagnostics.some(d => d.feature === 'table-style-reply-limit')) {
      this.host.release(out); entry.state = 'failed'; entry.diagnostics = payload.diagnostics; entry.gen++; this.host.changed(sheetId); return
    }
    let layerRelease: (() => void) | undefined
    try {
      const decodedBytes = payload.tables.length ? 6_400_000 + payload.tables.length * 2048 : 0
      const replyBytes = out.replyBytes ?? 0
      if (!Number.isSafeInteger(replyBytes) || replyBytes < 0 || this.retainedBytes + replyBytes + decodedBytes > this.budgetBytes) {
        this.host.release(out); entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: 'table-style-reply-limit', message: 'Table style metadata exceeds the receiver cache budget.' }]; this.host.changed(sheetId); return
      }
      let layer: RegionStyleLayer | undefined
      if (decodedBytes > 0) {
        const admittedLayer = this.host.admitLayer(sheetId, requestId, decodedBytes)
        if (!admittedLayer) { this.host.release(out); entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: 'table-style-reply-limit', message: 'Table style metadata exceeds the main-thread cache budget.' }]; this.host.changed(sheetId); return }
        layerRelease = admittedLayer
        const { preloadTableStyles } = await import('../xlsx/tableStyleLoader')
        await preloadTableStyles()
        const { buildTableRegions } = await import('../xlsx/tableStyles/regionLayer')
        const extras = { dxfs: [], tableStyles: new Map(payload.tables.map(item => [item.def.name, item.def])), slicerStyles: new Map() }
        layer = buildTableRegions(payload.tables.map(item => item.table), extras, payload.theme ?? (await import('../drawing/style')).parseThemeContext())
      }
      // The source owner may have changed while the lazy engine was loading.
      if (this.closed || entry.seq !== seq || !sameOwner(owner, this.host.identity(sheetId))) { layerRelease?.(); layerRelease = undefined; this.host.release(out); entry.state = 'waiting'; return }
      entry.payload = payload; entry.layer = layer; entry.reply = out; entry.replyBytes = replyBytes; entry.layerBytes = decodedBytes; entry.layerRelease = layerRelease
      layerRelease = undefined
      entry.diagnostics = payload.diagnostics; entry.state = 'ready'; entry.gen++
      this.host.changed(sheetId)
    } catch (cause) {
      layerRelease?.(); this.host.release(out); entry.state = 'failed'; entry.diagnostics = [{ kind: 'unsupported', feature: 'table-style-read-failed', message: cause instanceof Error ? cause.message.slice(0, 256) : 'Table style engine failed to load' }]; this.host.changed(sheetId)
    }
  }
}

export function sheetWithRegionStyles<T extends { tables?: unknown[]; regionStyles?: RegionStyleLayer; regionStylesResolved?: boolean; diagnostics?: Array<{ kind: string; feature: string; message: string }> }>(sheet: T, state: PagedRegionStyleState): T {
  if (state.state !== 'ready' && state.state !== 'failed') return sheet
  const diagnostics = [...(sheet.diagnostics ?? [])]
  for (const diagnostic of state.diagnostics) if (!diagnostics.some(found => found.feature === diagnostic.feature && found.message === diagnostic.message)) diagnostics.push(diagnostic)
  if (state.state !== 'ready' || !state.payload) return { ...sheet, diagnostics } as T
  return { ...sheet, tables: state.payload.tables.map(item => item.table), regionStyles: state.layer, regionStylesResolved: true, diagnostics } as T
}
