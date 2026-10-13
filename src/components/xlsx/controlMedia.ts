import { useEffect, useLayoutEffect, useRef, useState } from 'react'

/**
 * True while the component is mounted; false from the commit that removes it.
 * A layout-effect cleanup runs inside that commit, before the DOM is detached
 * and before any later task, so an observer, media-query or timer callback
 * that fires between the removal and the passive cleanups can check it and
 * never queue a React update for an unmounted tree (React keeps such an update,
 * and the fiber it names, until the root renders again).
 */
export function useLive(): { readonly current: boolean } {
  const live = useRef(false)
  useLayoutEffect(() => {
    live.current = true
    return () => { live.current = false }
  }, [])
  return live
}

export function useControlMedia(query: string): boolean {
  const live = useLive()
  const [matches, setMatches] = useState(() => typeof window !== 'undefined' && Boolean(window.matchMedia?.(query).matches))
  useEffect(() => {
    const media = window.matchMedia?.(query)
    if (!media) return
    const update = () => { if (live.current) setMatches(media.matches) }
    update()
    media.addEventListener('change', update)
    return () => media.removeEventListener('change', update)
  }, [query, live])
  return matches
}
