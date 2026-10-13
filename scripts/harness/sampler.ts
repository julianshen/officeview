/**
 * A periodic probe that never overlaps itself, never leaks a rejection and
 * stops for good: failures while running are reported to `onError`; a probe
 * still in flight when stop() is called settles silently (its page may be
 * closing). Used by the S10.2 memory session.
 */
export function startSampler(intervalMs: number, probe: () => Promise<void>, onError: (error: unknown) => void): () => void {
  let running = true, busy = false
  const timer = setInterval(() => {
    if (busy || !running) return
    busy = true
    probe().catch(error => { if (running) onError(error) }).finally(() => { busy = false })
  }, intervalMs)
  return () => { running = false; clearInterval(timer) }
}
