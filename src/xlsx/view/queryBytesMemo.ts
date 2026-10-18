/** queryPayloadBytes already measured (by the worker's payload walker) for an
 * object at a limit, so a later consumer of the same page does not walk it
 * again (spec §10.2). Worker-side; weakly held. */
const known = new WeakMap<object, { limit: number; bytes: number }>()
export function rememberQueryBytes(value: unknown, limit: number, bytes: number): void {
  if (typeof value === 'object' && value !== null) known.set(value, { limit, bytes })
}
export function knownQueryBytes(value: unknown, limit: number): number | undefined {
  const hit = typeof value === 'object' && value !== null ? known.get(value) : undefined
  return hit?.limit === limit ? hit.bytes : undefined
}
