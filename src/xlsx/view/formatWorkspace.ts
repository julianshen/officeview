import type { FormatAllocation } from '../formula/format'

/** The calling operation's aggregate budget, never a per-cell allowance. */
export interface FormatBudget {
  reserve(bytes: number): void
  release(bytes: number): void
  retain?(targets: readonly object[], bytes: number, commit: () => void, kind?: 'diagnostic'): void
}

/** Keep builders charged through the synchronous consumer of their output. */
export function withFormatWorkspace<T>(budget: FormatBudget, run: (allocation: FormatAllocation) => T): T {
  let held = 0, pending = 0
  const flush = () => { if (pending) { budget.release(pending); held -= pending; pending = 0 } }
  const allocation: FormatAllocation = {
    allocate: bytes => { flush(); budget.reserve(bytes); held += bytes },
    // A helper's old argument remains aliased until it returns to its caller.
    release: bytes => { pending += bytes },
    ...(budget.retain ? { retain: (targets: readonly object[], bytes: number, commit: () => void, kind?: 'diagnostic') => {
      flush()
      budget.retain!(targets, bytes, commit, kind)
      held -= bytes
    } } : {}),
  }
  try { return run(allocation) }
  finally { if (held) budget.release(held) }
}
