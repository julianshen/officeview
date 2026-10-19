import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { render, waitFor, act } from '@testing-library/react'
import { docxModel, mockViewerLayout, type FakePage } from './helpers/paged-viewer'

// I1: the lazy chrome chunk fails to load (a host redeploy removed the hashed
// chunk, or a CSP blocks it). Pages must keep rendering and the host must not
// unmount. React's lazy caches per module, so this file holds one failure.
const h = vi.hoisted(() => ({ pages: null as FakePage[] | null }))
vi.mock('../src/components/paged/PagedChrome', () => { throw new Error('chunk load failed') })
vi.mock('../src/render/paint', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/render/paint')>()
  const { fakePaintables: fake } = await import('./helpers/paged-viewer')
  return { ...m, getPaintables: (doc: never, opts: never) => (h.pages ? Promise.resolve(fake(h.pages)) : m.getPaintables(doc, opts)) }
})
let restore: () => void
beforeEach(() => { h.pages = Array.from({ length: 3 }, () => ({ w: 800, h: 1000 })); restore = mockViewerLayout({ width: 1000, height: 800 }) })
afterEach(() => { restore(); vi.restoreAllMocks() })

describe('a rejected PagedChrome import', () => {
  it('leaves the pages rendered, keeps the host mounted, throws nothing and warns once', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const error = vi.spyOn(console, 'error').mockImplementation(() => {})
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const ui = render(<div data-testid="host"><OfficeDoc document={docxModel(['a'])} /></div>)
    await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBe(3))
    await act(async () => { await new Promise((r) => setTimeout(r, 30)) })
    ui.rerender(<div data-testid="host"><OfficeDoc document={docxModel(['b'])} /></div>)
    await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBe(3))
    expect(ui.getByTestId('host').querySelector('[data-ov-theme-root]')).not.toBeNull()
    expect(ui.container.querySelector('[role=region]')).not.toBeNull()
    expect(error).not.toHaveBeenCalled()
    expect(warn.mock.calls.filter((c) => /officeview/.test(String(c[0]))).length).toBe(1)
  })
})
