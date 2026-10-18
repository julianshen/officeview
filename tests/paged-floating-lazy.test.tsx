import { describe, it, expect, vi, beforeEach } from 'vitest'
import { render, waitFor, fireEvent } from '@testing-library/react'
import type { ComponentType } from 'react'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx } from '../src/testdata/ooxml-builders'

// The floating controls live in a lazy chunk. `gate` holds that dynamic
// import open so a test can act before the chunk resolves.
const chunk = vi.hoisted(() => ({ gate: Promise.resolve() as Promise<void> }))
vi.mock('../src/components/paged/PagedChrome', async (importOriginal) => {
  await chunk.gate
  return importOriginal()
})

const writeText = vi.fn().mockResolvedValue(undefined)

beforeEach(() => {
  chunk.gate = Promise.resolve()
  writeText.mockClear()
  Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
})

async function mount(props: Record<string, unknown> = {}, Doc?: ComponentType<any>) {
  const OfficeDoc = Doc ?? (await import('../src/components/OfficeDoc')).OfficeDoc
  const doc = await parseDocx(await OfficePackage.load(await buildDocx([
    { runs: [{ text: 'alpha beta gamma' }] },
    { runs: [{ text: 'second line of text' }] },
  ])))
  const utils = render(<OfficeDoc document={doc} {...props} />)
  await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
  // jsdom gives canvases no layout; give them a real box so hit testing works
  const canvas = utils.container.querySelector('canvas') as HTMLCanvasElement
  canvas.getBoundingClientRect = () =>
    ({ left: 0, top: 0, width: 816, height: 1056, right: 816, bottom: 1056, x: 0, y: 0, toJSON: () => ({}) }) as DOMRect
  return { ...utils, canvas, doc }
}

const copyButton = (c: HTMLElement) => c.querySelector('[data-testid="officeview-copy-selection"]') as HTMLElement | null

describe('lazy paged floating controls', () => {
  it('Cmd/Ctrl-F before the chunk loads focuses the search input on mount', async () => {
    // a fresh module graph so OfficeDoc's lazy() has not resolved the chunk yet
    vi.resetModules()
    let release!: () => void
    chunk.gate = new Promise<void>((r) => { release = r })
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const ui = await mount({}, OfficeDoc)
    expect(ui.queryByRole('searchbox')).toBeNull()
    fireEvent.keyDown(ui.getByRole('region', { name: 'Document viewer' }), { key: 'f', ctrlKey: true })
    release()
    const search = await ui.findByRole('searchbox')
    await waitFor(() => expect(document.activeElement).toBe(search))
  })

  it('floating search and zoom controls appear after the lazy chunk resolves', async () => {
    const ui = await mount()
    expect(await ui.findByRole('searchbox')).toBeTruthy()
    expect(await ui.findByLabelText('Zoom in')).toBeTruthy()
  })

  it('copy confirmation still announces Copied', async () => {
    const { container, canvas } = await mount()
    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => expect(copyButton(container)).toBeTruthy())
    fireEvent.click(copyButton(container)!)
    await waitFor(() => expect(copyButton(container)!.textContent).toBe('Copied ✓'))
  })

  it('clicking the lazy copy pill with a selection copies it', async () => {
    const { container, canvas } = await mount()
    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => expect(copyButton(container)).toBeTruthy())
    expect(copyButton(container)!.getAttribute('data-state')).toBe('idle')
    fireEvent.click(copyButton(container)!)
    await waitFor(() => expect(copyButton(container)!.getAttribute('data-state')).toBe('copied'))
    expect(writeText).toHaveBeenCalledTimes(1)
    expect(writeText).toHaveBeenCalledWith('beta')
  })

  it('allowCopy=false renders no copy pill even with a selection', async () => {
    const ui = await mount({ allowCopy: false })
    // the chunk is loaded (search is up) and a drag/double-click still makes no pill
    await ui.findByRole('searchbox')
    fireEvent.doubleClick(ui.canvas, { clientX: 160, clientY: 104, bubbles: true })
    fireEvent.pointerDown(ui.canvas, { pointerId: 1, clientX: 100, clientY: 104, button: 0, bubbles: true })
    fireEvent.pointerMove(ui.canvas, { pointerId: 1, clientX: 260, clientY: 104, bubbles: true })
    fireEvent.pointerUp(ui.canvas, { pointerId: 1, clientX: 260, clientY: 104, bubbles: true })
    await new Promise((r) => setTimeout(r, 40))
    expect(copyButton(ui.container)).toBeNull()
  })
})

describe('queued Cmd/Ctrl-F', () => {
  it('is honoured once, not replayed when the controls remount', async () => {
    vi.resetModules()
    let release!: () => void
    chunk.gate = new Promise<void>((r) => { release = r })
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const ui = await mount({}, OfficeDoc)
    fireEvent.keyDown(ui.getByRole('region', { name: 'Document viewer' }), { key: 'f', ctrlKey: true })
    release()
    await waitFor(() => expect(document.activeElement).toBe(ui.getByRole('searchbox')))
    ui.getByRole('searchbox').blur()
    // hide every floating control (the chunk unmounts), then show them again
    ui.rerender(<OfficeDoc document={ui.doc} showSearch={false} showZoomControls={false} />)
    expect(ui.queryByRole('searchbox')).toBeNull()
    ui.rerender(<OfficeDoc document={ui.doc} />)
    const search = await ui.findByRole('searchbox')
    await new Promise((r) => setTimeout(r, 20))
    expect(document.activeElement).not.toBe(search)
  })
})
