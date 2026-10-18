/**
 * The page in view, tracked from analytic page offsets (spec §6.11, R7): one
 * computation per animation frame on scroll, transform, resize or layout
 * change, and a state update only when the index changes. The viewport maps
 * into content only through scrollPan.contentYAt (native scrollTop included).
 * goTo scrolls the viewer's own container, or pans through clampPan while
 * zoomed; never scrollIntoView, so host ancestors never scroll.
 */
import { useEffect, useRef, useState, type RefObject } from 'react'
import type { PageSpec } from '../../render/paint'
import { clampPan, type Transform } from '../../core/zoom'
import { contentYAt, pageOffsets, panYForContentTop, type ScrollPanState } from '../scrollPan'
import { PAGE_PADDING_PX } from './metrics'
import { mostVisiblePage } from './pageTracking'

export interface PagedLayout { specs: readonly PageSpec[]; displayWidth: number; gap: number }

export function usePageTracking(
  container: RefObject<HTMLElement | null>,
  layout: PagedLayout,
  transform: Transform,
  setTransform: (t: Transform) => void,
): { current: number; goTo(index: number): void; reveal(contentY: number): void } {
  const { specs, displayWidth, gap } = layout
  const [current, setCurrent] = useState(0)
  const latest = useRef({ specs, displayWidth, gap, transform, setTransform })
  latest.current = { specs, displayWidth, gap, transform, setTransform }
  const offsets = useRef<{ top: number[]; height: number[] }>({ top: [], height: [] })
  const frame = useRef(0)

  const state = (el: HTMLElement): ScrollPanState => {
    const t = latest.current.transform
    return { scrollTop: el.scrollTop, panY: t.panY, zoom: t.zoom, padding: PAGE_PADDING_PX }
  }
  const compute = () => {
    frame.current = 0
    const el = container.current
    if (!el) return
    const s = state(el)
    const top = contentYAt(0, s)
    setCurrent(mostVisiblePage(offsets.current, top, contentYAt(el.clientHeight, s) - top))
  }
  const schedule = () => { if (!frame.current) frame.current = requestAnimationFrame(compute) }

  useEffect(() => {
    offsets.current = pageOffsets(specs, displayWidth, gap)
    schedule()
  }, [specs, displayWidth, gap, transform])

  useEffect(() => {
    const el = container.current
    if (!el) return
    el.addEventListener('scroll', schedule, { passive: true })
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(schedule) : undefined
    ro?.observe(el)
    return () => {
      el.removeEventListener('scroll', schedule)
      ro?.disconnect()
      if (frame.current) cancelAnimationFrame(frame.current)
      frame.current = 0
    }
  }, [container])

  /** Put content y at the viewport top. */
  const reveal = (y: number) => {
    const el = container.current
    if (!el) return
    const { transform: t, setTransform: set } = latest.current
    if (t.zoom > 1) {
      const o = offsets.current
      const n = o.top.length
      const contentHeight = n > 0 ? o.top[n - 1] + o.height[n - 1] : 0
      set(clampPan({ ...t, panY: panYForContentTop(y, state(el)) }, { width: el.clientWidth, height: el.clientHeight, contentWidth: latest.current.displayWidth, contentHeight }))
    } else if (typeof el.scrollTo === 'function') el.scrollTo({ top: PAGE_PADDING_PX + y })
    else el.scrollTop = PAGE_PADDING_PX + y
  }
  const goTo = (index: number) => {
    const top = offsets.current.top[index]
    if (top !== undefined) reveal(top)
  }
  return { current, goTo, reveal }
}
