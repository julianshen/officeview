import { describe, test, expect, vi, beforeEach } from 'vitest'
import { render, waitFor, fireEvent } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx } from '../src/testdata/ooxml-builders'

const writeText = vi.fn().mockResolvedValue(undefined)

beforeEach(() => {
  writeText.mockClear()
  Object.defineProperty(navigator, 'clipboard', {
    value: { writeText },
    configurable: true,
  })
})

async function mount() {
  const doc = await parseDocx(await OfficePackage.load(await buildDocx([
    { runs: [{ text: 'alpha beta gamma' }] },
    { runs: [{ text: 'second line of text' }] },
  ])))
  const utils = render(<OfficeDoc document={doc} />)
  await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
  // jsdom gives canvases no layout; make hit testing deterministic by giving
  // them a real box and pointing at the page's natural width
  const canvas = utils.container.querySelector('canvas') as HTMLCanvasElement
  canvas.getBoundingClientRect = () =>
    ({ left: 0, top: 0, width: 816, height: 1056, right: 816, bottom: 1056, x: 0, y: 0, toJSON: () => ({}) }) as DOMRect
  return { ...utils, canvas }
}

const copyButton = (c: HTMLElement) => c.querySelector('[data-testid="officeview-copy-selection"]')

describe('<OfficeDoc> text selection', () => {
  test('drag across text shows a selection and a copy button', async () => {
    const { container, canvas } = await mount()
    // press on the first line's text, then drag to the right
    fireEvent.pointerDown(canvas, { pointerId: 1, clientX: 100, clientY: 104, button: 0, bubbles: true })
    fireEvent.pointerMove(canvas, { pointerId: 1, clientX: 260, clientY: 104, bubbles: true })
    fireEvent.pointerUp(canvas, { pointerId: 1, clientX: 260, clientY: 104, bubbles: true })
    await waitFor(() => expect(copyButton(container)).toBeTruthy())

    // a plain click alone yields no selection and therefore no button
    const { container: c2, canvas: canvas2 } = await mount()
    fireEvent.pointerDown(canvas2, { pointerId: 1, clientX: 100, clientY: 104, button: 0, bubbles: true })
    fireEvent.pointerUp(canvas2, { pointerId: 1, clientX: 100, clientY: 104, bubbles: true })
    await new Promise((r) => setTimeout(r, 30))
    expect(copyButton(c2)).toBeNull()
  })

  test('double-click selects a word and copies it', async () => {
    const { container, canvas } = await mount()
    // inside "beta" (chars 6..10 of 'alpha beta gamma')
    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => expect(copyButton(container)).toBeTruthy())
    fireEvent.click(copyButton(container) as HTMLElement)
    await waitFor(() => expect(writeText).toHaveBeenCalled())
    expect(writeText.mock.calls[0][0]).toBe('beta')
  })

  test('Cmd/Ctrl+C copies the selection', async () => {
    const { container, canvas } = await mount()
    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => expect(copyButton(container)).toBeTruthy())
    const region = container.querySelector('[role="region"]') as HTMLElement
    fireEvent.keyDown(region, { key: 'c', metaKey: true })
    await waitFor(() => expect(writeText).toHaveBeenCalled())
    expect(writeText.mock.calls[0][0]).toBe('beta')
  })

  test('the copy button confirms the copy', async () => {
    const { container, canvas } = await mount()
    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => expect(copyButton(container)).toBeTruthy())
    fireEvent.click(copyButton(container) as HTMLElement)
    await waitFor(() => expect((copyButton(container) as HTMLElement).textContent).toBe('Copied ✓'))
  })

  test('Escape clears the selection', async () => {
    const { container, canvas } = await mount()
    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => expect(copyButton(container)).toBeTruthy())
    const region = container.querySelector('[role="region"]') as HTMLElement
    fireEvent.keyDown(region, { key: 'Escape' })
    await waitFor(() => expect(copyButton(container)).toBeNull())
  })

  test('the selection is painted onto the canvas', async () => {
    const { canvas } = await mount()
    const ctx = canvas.getContext('2d')!
    const before = ctx.getImageData(0, 0, canvas.width, canvas.height).data
    let inkBefore = 0
    for (let i = 0; i < before.length; i += 4) if (before[i] < 200) inkBefore++

    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => {
      const after = ctx.getImageData(0, 0, canvas.width, canvas.height).data
      let blue = 0
      // selection wash is rgba(64,128,255,0.35) over white -> ~(188,207,255)
      for (let i = 0; i < after.length; i += 4) {
        if (after[i] > 175 && after[i] < 205 && after[i + 1] > 195 && after[i + 1] < 225 && after[i + 2] > 240) blue++
      }
      expect(blue).toBeGreaterThan(10)
    })
    void inkBefore
  })

  test('double-click on blank space still zooms (no selection)', async () => {
    const { container, canvas } = await mount()
    // far below the last line: no text there
    fireEvent.doubleClick(canvas, { clientX: 400, clientY: 1000, bubbles: true })
    await new Promise((r) => setTimeout(r, 40))
    expect(copyButton(container)).toBeNull()
    const content = container.querySelector('[data-testid="officeview-content"]') as HTMLElement
    expect(content.style.transform).toContain('scale')
  })
})