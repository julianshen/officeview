import type { ProfileId } from './contracts'
import type { MainMemoryLedger } from './mainLedger'
import { OverflowCache } from './overflowCache'
import { MAIN_THREAD_BUDGETS, MemoryPool } from './tileCache'

let nextOwner = 0

/** The subcap is part of the existing shared main budget, never extra memory. */
export function createDocumentOverflowCache(ledger: MainMemoryLedger, profile: ProfileId): OverflowCache {
  const limit = MAIN_THREAD_BUDGETS[profile], owner = `overflow:${nextOwner++}`
  let held = 0
  return new OverflowCache(new MemoryPool({ limitBytes: limit }), 2 * 1024 * 1024, {
    reserve(bytes) {
      if (ledger.bytes + bytes > limit) return false
      held += bytes
      ledger.set('layout', owner, held)
      return true
    },
    release(bytes) { held -= bytes; ledger.set('layout', owner, held) },
  })
}
