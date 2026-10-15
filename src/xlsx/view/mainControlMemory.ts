import { StructuralIndex } from './structure'
import type { ProfileId } from './contracts'
import { SOURCE_ALLOCATION, adoptAllocation } from './allocationAccount'
import type { AllocationContext, AllocationReservation } from './allocationAccount'
import { MainMemoryLedger } from './mainLedger'
import { MAIN_THREAD_BUDGETS } from './tileCache'
import { ValueIndexError } from './valueRuns'

let serial = 0
const profiles = new WeakMap<MainMemoryLedger, ProfileId>()
/** Every reservation consults the live document aggregate, including unrelated pools. */
export class MainControlMemory implements AllocationContext {
  readonly limitBytes: number
  /** Explicit owners still in use; adopted reachability owners remain in the ledger. */
  pendingBytes = 0
  private isClosed = false
  get usedBytes(): number { return this.ledger.bytes }
  get closed(): boolean { return this.isClosed }
  close(): void { this.isClosed = true }
  constructor(readonly ledger: MainMemoryLedger, readonly profile: ProfileId = 'D') {
    this.limitBytes = MAIN_THREAD_BUDGETS[profile]
    if (!this.limitBytes || profiles.has(ledger) && profiles.get(ledger) !== profile) throw new RangeError('Main ledger profile must stay fixed')
    profiles.set(ledger, profile)
  }
  canReserve(bytes: number): boolean { return !this.isClosed && Number.isSafeInteger(bytes) && bytes >= 0 && this.ledger.bytes + bytes <= this.limitBytes }
  owns(target: object): boolean { return this.ledger.hasLifetime(target) }
  settle(target: object): void { this.ledger.settle(target) }
  hold(target: object): { release(): void } | null { return this.ledger.hold(target) }
  reserve(): AllocationReservation {
    const thisContext = this
    const ledger = this.ledger, limit = this.limitBytes, owner = `control-allocation-${++serial}`
    let held = 0
    const reservation: AllocationReservation = {
      owner,
      get bytes() { return held },
      trySet(bytes) {
        if (!Number.isSafeInteger(bytes) || bytes < 0 || bytes > held && (thisContext.closed || ledger.bytes - held + bytes > limit)) return false
        thisContext.pendingBytes += bytes - held
        ledger.set('controls', owner, bytes)
        held = bytes
        return true
      },
      charge(bytes) {
        if (bytes > 0 && thisContext.closed) throw new ValueIndexError('cancelled', 'Controls allocation scope closed')
        if (!reservation.trySet(held + bytes)) throw new ValueIndexError('quota', 'Main document memory quota exceeded', { resource: 'memory', requiredBytes: ledger.bytes + bytes, limitBytes: limit, retryable: true })
      },
      release(bytes = held) {
        if (!Number.isSafeInteger(bytes) || bytes < 0 || bytes > held) throw new RangeError('Invalid allocation release')
        reservation.trySet(held - bytes)
      },
      adopt(targets, bytes = held, commit, holdable) {
        if (!bytes) return
        ledger.adoptWhileReachable('controls', owner, 'controls', targets, bytes, commit, holdable)
        held -= bytes
        thisContext.pendingBytes -= bytes
      },
    }
    return reservation
  }
}

/** Adds per-call source admission without retargeting the original source owner. */
export function controlSourceStore(store: import('./contracts').SourceStore, memory: AllocationContext): import('./contracts').SourceStore {
  return new Proxy(store, {
    get(target, key) {
      const value = Reflect.get(target, key, target)
      // A page this proxy's admission adopted ends when its consumer settles it.
      if (key === 'settle') return (payload: object) => { if (typeof value === 'function') value.call(target, payload); memory.settle?.(payload) }
      if (typeof value !== 'function') return value
      if (key === 'queryColumn' || key === 'queryStructural' || key === 'scanPopulated') return (request: unknown, options: object) => value.call(target, request, { ...options, [SOURCE_ALLOCATION]: memory })
      if (key === 'readRows') return (range: unknown, cursor: unknown, options: object) => value.call(target, range, cursor, { ...options, [SOURCE_ALLOCATION]: memory })
      return value.bind(target)
    },
  })
}

/** The frozen index clones records and builds four trees; reserve their construction peak. */
export function controlStructure(result: import('./contracts').StructuralResult, memory: AllocationContext): import('./structure').StructuralIndex {
  const account = memory.reserve()
  const count = result.merges.length + result.tables.length + result.spills.length + result.drawings.length
  account.charge(1024 + count * 1024 + result.missing.length * 128)
  try {
    const index = new StructuralIndex(result)
    adoptAllocation(account, index, index.estimatedBytes)
    return index
  } finally { account.release() }
}

/** One imported/returned allocation, with backing-store identity deduplication.
 * Its charge is holdable: the views that show it hold it (§12.4). With `hold`,
 * the caller (a pending operation reading this view) holds it too and releases
 * the result when it ends. */
export function controlView(memory: AllocationContext, order: import('./contracts').RowOrder, visibility: import('./visibility').CompactRowVisibility, hold = false): () => void {
  const holds: { release(): void }[] = []
  const own = (root: object, count: number, visit: (add: (lane: ArrayBufferView) => void) => void) => {
    if (memory.owns(root)) { const held = hold && memory.hold?.(root); if (held) holds.push(held); return }
    const account = memory.reserve()
    try {
      account.charge(256 + count * 128)
      const buffers = new Set<ArrayBufferLike>(), targets: object[] = [root]
      let bytes = 128
      visit(lane => {
        if (buffers.has(lane.buffer) || memory.owns(lane.buffer)) return
        account.charge(lane.buffer.byteLength)
        buffers.add(lane.buffer)
        bytes += lane.buffer.byteLength
        targets.push(lane, lane.buffer)
      })
      account.adopt(targets, bytes, undefined, true)
      const held = hold && memory.hold?.(root)
      if (held) holds.push(held)
    } finally { account.release() }
  }
  const release = () => { for (const held of holds.splice(0)) held.release() }
  try {
    if (order.kind !== 'identity') own(order, order.bands.length * 2, add => {
      for (const band of order.bands) { add(band.sourceAtOffset); add(band.displayOffsetOfSource) }
    })
    own(visibility, 2, add => { add(visibility.words); add(visibility.summaries) })
  } catch (error) { release(); throw error } // a refused reader keeps nothing it acquired
  return release
}
