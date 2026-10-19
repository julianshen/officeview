import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { render, waitFor, fireEvent, act } from '@testing-library/react'
import { docxModel, mockViewerLayout, type FakePage } from './helpers/paged-viewer'

// I1: the lazy zoom routing fails to load. Every later gesture must not
// produce an unhandled rejection; the failure is logged once, never retried.
const h = vi.hoisted(() => ({ pages: null as FakePage[] | null }))
vi.mock('../src/render/paint', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/render/paint')>()
  const { fakePaintables: fake } = await import('./helpers/paged-viewer')
  return { ...m, getPaintables: (doc: never, opts: never) => (h.pages ? Promise.resolve(fake(h.pages)) : m.getPaintables(doc, opts)) }
})
let restore: () => void
beforeEach(() => { h.pages = Array.from({ length: 3 }, () => ({ w: 800, h: 1000 })); restore = mockViewerLayout({ width: 1000, height: 800 }) })
afterEach(() => { restore(); vi.restoreAllMocks() })

describe('a rejected zoom-routing import', () => {
  it('repeated wheel gestures raise no unhandled rejection, one warning, one load attempt', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const unhandled: unknown[] = []
    const onUnhandled = (e: unknown) => { unhandled.push(e) }
    process.on('unhandledRejection', onUnhandled)
    try {
      const { OfficeDoc, routerLoader } = await import('../src/components/OfficeDoc')
      const load = vi.fn(() => Promise.reject(new Error('chunk load failed')))
      routerLoader.load = load
      const ui = render(<OfficeDoc document={docxModel(['a'])} showToolbar={false} />)
      await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBe(3))
      const canvas = ui.container.querySelector('canvas')!
      for (let i = 0; i < 5; i++) {
        fireEvent.wheel(canvas, { ctrlKey: true, deltaY: -100, clientX: 100, clientY: 100 })
        await act(async () => { await new Promise((r) => setTimeout(r, 10)) })
      }
      await act(async () => { await new Promise((r) => setTimeout(r, 30)) })
      expect(unhandled).toEqual([])
      expect(load).toHaveBeenCalledTimes(1)
      expect(warn.mock.calls.filter((c) => /officeview/.test(String(c[0]))).length).toBe(1)
      expect(ui.container.querySelectorAll('canvas').length).toBe(3)
    } finally {
      process.off('unhandledRejection', onUnhandled)
    }
  })
})
