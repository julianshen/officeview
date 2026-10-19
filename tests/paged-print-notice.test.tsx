import { afterEach, describe, expect, it, vi } from 'vitest'
import { render, fireEvent, waitFor, cleanup } from '@testing-library/react'
import { docxModel } from './helpers/paged-viewer'

// Final review M1/M2: the print notice sits above the status line by the
// metrics.ts height (not a literal), and its dismiss button is a Lucide icon.
vi.mock('../src/components/paged/metrics', async (importOriginal) => ({
  ...await importOriginal<typeof import('../src/components/paged/metrics')>(),
  PAGED_STATUS_HEIGHT: 30,
}))
afterEach(() => { cleanup(); vi.restoreAllMocks() })

describe('print notice', () => {
  it.each([[true, '38px'], [false, '8px']])('clears the status line from metrics.ts with toolbar=%s and dismisses through an X icon', async (showToolbar, bottom) => {
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const { printLoader } = await import('../src/components/paged/PagedChrome')
    vi.spyOn(printLoader, 'load').mockRejectedValue(new Error('network'))
    const view = render(<OfficeDoc document={docxModel(['one'])} showToolbar={showToolbar} />)
    await waitFor(() => expect(view.container.querySelector('[data-ov-print-notice]')).not.toBeNull())
    const box = view.container.querySelector('[data-ov-print-notice]') as HTMLElement
    expect(box.style.bottom).toBe(bottom)
    fireEvent.keyDown(view.getByRole('region'), { key: 'p', ctrlKey: true })
    const dismiss = await view.findByRole('button', { name: 'Dismiss print notice' })
    expect(dismiss.querySelector('svg[data-icon="x"]')).not.toBeNull()
    expect(dismiss.textContent).toBe('')
  })
})
