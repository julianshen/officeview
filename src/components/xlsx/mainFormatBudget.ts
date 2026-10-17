import { MainMemoryLedger, type MainPool } from '../../xlsx/view/mainLedger'
import { FindError } from '../../xlsx/view/find'
import type { FormatBudget } from '../../xlsx/view/formatWorkspace'

let serial = 0
/** One viewer/operation's workspace. Retained diagnostics/hits leave by adoption. */
export function mainFormatBudget(ledger: MainMemoryLedger, limit: number, retainedPool: MainPool = 'document'): FormatBudget {
  const owner = `format-${++serial}`
  let held = 0
  return {
    reserve(bytes) {
      if (ledger.bytes + bytes > limit) throw new FindError('resource-limit', 'Formatting exceeds the memory budget. Narrow the search or inspect fewer cells.')
      ledger.set('layout', owner, held += bytes)
    },
    release(bytes) { ledger.set('layout', owner, held -= bytes) },
    retain(targets, bytes, commit) {
      ledger.adoptWhileReachable('layout', owner, retainedPool, targets, bytes, commit)
      held -= bytes
    },
  }
}
