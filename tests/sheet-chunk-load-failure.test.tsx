/**
 * Final re-review residual: the interactive sheet chunk fails to load (a host
 * redeploy removed the hashed chunk, or a CSP blocks it). The failure stays
 * inside the viewer: a visible status, one warning, and the host's sibling UI
 * and error boundary are untouched. React's lazy caches per module, so this
 * file holds one failure.
 */
import { Component, type ReactNode } from 'react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'

vi.mock('../src/components/xlsx/InteractiveSheetDoc', () => { throw new Error('chunk load failed') })
afterEach(() => { cleanup(); vi.restoreAllMocks() })

class Boundary extends Component<{ onError: (e: unknown) => void; children: ReactNode }, { failed: boolean }> {
  state = { failed: false }
  static getDerivedStateFromError() { return { failed: true } }
  componentDidCatch(e: unknown) { this.props.onError(e) }
  render() { return this.state.failed ? <p>host boundary</p> : this.props.children }
}

describe('a rejected InteractiveSheetDoc import', () => {
  it('shows a status in the viewer, warns once, and leaves the host sibling and boundary untouched', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    vi.spyOn(console, 'error').mockImplementation(() => {})
    const caught = vi.fn()
    const book = await parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])))
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const ui = render(<Boundary onError={caught}><p data-testid="sibling">host sibling</p><OfficeDoc document={book} interactiveSheets theme="dark" /></Boundary>)
    const status = await waitFor(() => ui.getByText('Spreadsheet viewer could not load. Reload to retry.'))
    expect(status.getAttribute('role')).toBe('status')
    const root = status.closest('[data-ov-theme-root]')!
    expect(root.getAttribute('data-ov-theme')).toBe('dark')
    expect(ui.getByTestId('sibling').textContent).toBe('host sibling')
    expect(ui.queryByText('host boundary')).toBeNull()
    expect(caught).not.toHaveBeenCalled()
    expect(warn.mock.calls.filter((c) => /officeview/.test(String(c[0]))).length).toBe(1)
  })
})
