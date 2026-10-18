import type { CacheIdentity } from './contracts'

export const identityFields = ['docId', 'sheetId', 'valueGen', 'fontGen', 'viewGen'] as const

/** The same complete owner guards both wire evidence and plan publication. */
export function sameIdentity(a: CacheIdentity | null | undefined, b: CacheIdentity): boolean {
  return a != null && identityFields.every(field => a[field] === b[field])
}
