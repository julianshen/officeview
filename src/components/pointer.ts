/**
 * Whether the primary pointer is coarse (touch). Read synchronously on the
 * first render, so the first paint already has the right chrome size, and
 * kept current through the media query's change event. Eager and tiny.
 */
import { useSyncExternalStore } from 'react'

const query = (): MediaQueryList | undefined => globalThis.matchMedia?.('(pointer: coarse)')

const subscribe = (cb: () => void): (() => void) => {
  const mq = query()
  mq?.addEventListener?.('change', cb)
  return () => mq?.removeEventListener?.('change', cb)
}

export function useCoarsePointer(): boolean {
  return useSyncExternalStore(subscribe, () => !!query()?.matches, () => false)
}
