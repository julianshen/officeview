import { useEffect, useState } from 'react'

/**
 * The live window.devicePixelRatio. A window moved between displays changes
 * the ratio without changing any CSS size, so a ResizeObserver never fires
 * (S13 D-DPR). Sources, any of which re-reads the ratio:
 *  - a `(resolution: Xdppx)` media query, re-armed for each new ratio (the
 *    standard change signal);
 *  - window and visualViewport `resize` (engines that pair a ratio change
 *    with a resize, and browser zoom);
 *  - a slow poll while the page is visible (suspended while hidden): Chrome's
 *    CDP device-metrics emulation changes the ratio and dispatches none of
 *    the above.
 * Setting the same value is a no-op, so unchanged reads never re-render.
 */
export function useDevicePixelRatio(pollMs = 1000): number {
  const [dpr, setDpr] = useState(readDpr)
  useEffect(() => {
    if (typeof window === 'undefined') return
    let media: MediaQueryList | undefined
    const arm = () => {
      media?.removeEventListener('change', update)
      media = window.matchMedia?.(`(resolution: ${readDpr()}dppx)`)
      media?.addEventListener('change', update)
    }
    function update() {
      const next = readDpr()
      setDpr(next)
      if (media?.media !== `(resolution: ${next}dppx)`) arm()
    }
    arm()
    update() // the ratio may have changed between render and effect
    const viewport = window.visualViewport
    window.addEventListener('resize', update)
    viewport?.addEventListener('resize', update)
    // The poll runs only while the page is visible: a hidden page keeps no
    // timer alive, and becoming visible re-reads once before polling again.
    let timer: ReturnType<typeof setInterval> | undefined
    const hidden = () => typeof document !== 'undefined' && document.visibilityState === 'hidden'
    const schedule = () => {
      if (hidden()) { clearInterval(timer); timer = undefined; return }
      if (timer === undefined) { update(); timer = setInterval(update, pollMs) }
    }
    schedule()
    document.addEventListener('visibilitychange', schedule)
    return () => {
      media?.removeEventListener('change', update)
      window.removeEventListener('resize', update)
      viewport?.removeEventListener('resize', update)
      document.removeEventListener('visibilitychange', schedule)
      clearInterval(timer)
    }
  }, [pollMs])
  return dpr
}

function readDpr(): number {
  return typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1
}
