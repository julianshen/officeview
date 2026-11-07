import type { FormatAllocation } from '../formula/format'
import type { XlsxCellStyle } from '../types'

/** The calling operation's aggregate budget, never a per-cell allowance. */
export interface FormatBudget {
  reserve(bytes: number): void
  release(bytes: number): void
  retain?(targets: readonly object[], bytes: number, commit: () => void, kind?: 'diagnostic'): void
  /** Ends a retained target's charge now: its owner dropped it (no GC wait). */
  settle?(target: object): void
}

/** One prepared current-generation effective format for a source batch (Task 9).
 * Implementations live in lazy modules; eager consumers only call these methods.
 * `describe` returns borrowed text owned (and charged) by the scope: callers must
 * not retain it past `dispose`. Absence of a scope, or an undefined per-cell
 * answer, always means base formatting (fail closed). */
export interface EffectiveFormatScope {
  /** Admitted source-space coverage for blank CF cells in a paged painter. */
  readonly regions?: readonly { minRow: number; maxRow: number; minCol: number; maxCol: number }[]
  /** Live content version the scope was prepared under (cfGen). */
  readonly gen: number
  /** Admitted bytes this scope retains, released by `dispose`. */
  readonly bytes: number
  /** Current-generation effective style, or the base when CF does not win. */
  effectiveStyle(row: number, col: number, base: XlsxCellStyle | undefined): XlsxCellStyle | undefined
  /** Localized winning-state descriptor, or undefined when nothing applies. */
  describe(row: number, col: number): { key: string; text: string } | undefined
  dispose(): void
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
