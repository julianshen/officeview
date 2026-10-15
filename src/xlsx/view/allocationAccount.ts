/** Internal, per-operation admission; default worker callers need no main account. */
export interface AllocationReservation {
  readonly owner: string
  readonly bytes: number
  trySet(bytes: number): boolean
  charge(bytes: number): void
  release(bytes?: number): void
  adopt(targets: readonly object[], bytes?: number, commit?: () => void): void
}
export interface AllocationContext {
  readonly closed?: boolean
  readonly usedBytes?: number
  readonly limitBytes: number
  reserve(): AllocationReservation
  canReserve(bytes: number): boolean
  owns(target: object): boolean
  /** Ends the lifetime charge `target` names: its consumer has merged or dropped it. */
  settle?(target: object): void
}
/** A scoped source call: never stored as mutable state on a shared adapter. */
export const SOURCE_ALLOCATION = Symbol('source allocation')
export interface SourceAllocationOptions { readonly [SOURCE_ALLOCATION]?: AllocationContext }

/** Complete owned page traversal: extracted backing stores keep the owner alive. */
export function allocationTargets(value: object, borrowed: readonly object[] = [], scratch?: AllocationReservation): object[] {
  scratch?.charge(96 + borrowed.length * 48)
  const targets: object[] = [], seen = new Set<object>(borrowed), pending: object[] = [value]
  while (pending.length) {
    const target = pending.pop()!
    if (seen.has(target)) continue
    scratch?.charge(128)
    seen.add(target)
    targets.push(target)
    if (ArrayBuffer.isView(target)) { scratch?.charge(32); pending.push(target.buffer) }
    else if (!(target instanceof ArrayBuffer)) for (const key in target) {
      if (!Object.prototype.hasOwnProperty.call(target, key)) continue
      const child = (target as Record<string, unknown>)[key]
      if (child && typeof child === 'object') { scratch?.charge(32); pending.push(child) }
    }
  }
  return targets
}

/** Traversal/registry staging coexists with the owned output, then drains. */
export function adoptAllocation(account: AllocationReservation, value: object, bytes = account.bytes, borrowed: readonly object[] = []): void {
  const targets = allocationTargets(value, borrowed, account)
  account.adopt(targets, bytes + targets.length * 64 + 128)
}
