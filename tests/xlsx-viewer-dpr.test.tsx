/**
 * S13 D-DPR: the interactive viewer must follow runtime devicePixelRatio
 * changes (a window moved between displays) without any other input. The
 * hook re-arms a `(resolution: Xdppx)` media query for each new ratio, also
 * re-reads on window/visualViewport resize, and keeps a slow visible-page
 * poll because some environments (Chrome's CDP device-metrics emulation,
 * observed in tests/browser/interactive-viewer-dpr.spec.ts) change the ratio
 * without dispatching any of those events.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, render } from '@testing-library/react'
import { useDevicePixelRatio } from '../src/components/xlsx/devicePixelRatio'

interface FakeQuery { query: string; listeners: Set<() => void>; removed: number }
let queries: FakeQuery[] = []
let ratio = 1
const originalMatchMedia = window.matchMedia
const originalViewport = Object.getOwnPropertyDescriptor(window, 'visualViewport')

function setRatio(next: number) { ratio = next }
function liveQueries(): string[] { return queries.filter(q => q.listeners.size > 0).map(q => q.query) }

beforeEach(() => {
  queries = []
  ratio = 1
  Object.defineProperty(window, 'devicePixelRatio', { configurable: true, get: () => ratio })
  window.matchMedia = ((query: string) => {
    const entry: FakeQuery = { query, listeners: new Set(), removed: 0 }
    queries.push(entry)
    return {
      media: query,
      matches: query === `(resolution: ${ratio}dppx)`,
      addEventListener: (_type: string, listener: () => void) => { entry.listeners.add(listener) },
      removeEventListener: (_type: string, listener: () => void) => { if (entry.listeners.delete(listener)) entry.removed++ },
    } as unknown as MediaQueryList
  }) as typeof window.matchMedia
  Object.defineProperty(window, 'visualViewport', { configurable: true, value: new EventTarget() })
})
afterEach(() => {
  vi.useRealTimers()
  window.matchMedia = originalMatchMedia
  if (originalViewport) Object.defineProperty(window, 'visualViewport', originalViewport)
  else delete (window as unknown as Record<string, unknown>).visualViewport
})

function Probe({ onValue }: { onValue: (dpr: number) => void }) {
  onValue(useDevicePixelRatio())
  return null
}
function mount() {
  const seen: number[] = []
  const view = render(<Probe onValue={value => seen.push(value)} />)
  return { seen, latest: () => seen[seen.length - 1], view }
}

describe('useDevicePixelRatio (D-DPR)', () => {
  it('a (resolution) media-query change updates the ratio and re-arms the query for the new ratio', () => {
    const probe = mount()
    expect(probe.latest()).toBe(1)
    expect(liveQueries()).toEqual(['(resolution: 1dppx)'])
    setRatio(2)
    act(() => { for (const listener of [...queries[queries.length - 1].listeners]) listener() })
    expect(probe.latest()).toBe(2)
    expect(liveQueries()).toEqual(['(resolution: 2dppx)'])
    setRatio(1.5)
    act(() => { for (const listener of [...queries[queries.length - 1].listeners]) listener() })
    expect(probe.latest()).toBe(1.5)
    expect(liveQueries()).toEqual(['(resolution: 1.5dppx)'])
  })

  it('window resize and visualViewport resize are fallbacks', () => {
    const probe = mount()
    setRatio(2)
    act(() => { window.dispatchEvent(new Event('resize')) })
    expect(probe.latest()).toBe(2)
    setRatio(3)
    act(() => { window.visualViewport!.dispatchEvent(new Event('resize')) })
    expect(probe.latest()).toBe(3)
  })

  it('a slow visible-page poll catches a change that dispatched no event', () => {
    vi.useFakeTimers()
    const probe = mount()
    setRatio(2)
    act(() => { vi.advanceTimersByTime(1000) })
    expect(probe.latest()).toBe(2)
  })

  it('an unchanged ratio does not re-render', () => {
    vi.useFakeTimers()
    const probe = mount()
    const renders = probe.seen.length
    act(() => { window.dispatchEvent(new Event('resize')); vi.advanceTimersByTime(5000) })
    expect(probe.seen.length).toBe(renders)
  })

  it('unmount removes every listener and stops the poll', () => {
    vi.useFakeTimers()
    const probe = mount()
    const removeWindow = vi.spyOn(window, 'removeEventListener')
    const removeViewport = vi.spyOn(window.visualViewport!, 'removeEventListener')
    probe.view.unmount()
    expect(liveQueries()).toEqual([])
    expect(removeWindow).toHaveBeenCalledWith('resize', expect.any(Function))
    expect(removeViewport).toHaveBeenCalledWith('resize', expect.any(Function))
    expect(vi.getTimerCount()).toBe(0)
  })
})
