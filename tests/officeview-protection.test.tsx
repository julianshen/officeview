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

async function mount(props: Record<string, unknown> = {}) {
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
  return { ...utils, canvas }
}

/** Drag across the first line of text, the way a reader would. */
function dragAcrossText(canvas: HTMLCanvasElement): void {
  fireEvent.pointerDown(canvas, { pointerId: 1, clientX: 100, clientY: 104, button: 0, bubbles: true })
  fireEvent.pointerMove(canvas, { pointerId: 1, clientX: 260, clientY: 104, bubbles: true })
  fireEvent.pointerUp(canvas, { pointerId: 1, clientX: 260, clientY: 104, bubbles: true })
}

const copyButton = (c: HTMLElement) => c.querySelector('[data-testid="officeview-copy-selection"]')
const printGuards = () => document.head.querySelectorAll('style[data-officeview-print-guard]')

describe('allowCopy={false} (copy prevention)', () => {
  test('dragging text produces no selection and no copy button', async () => {
    const { container, canvas } = await mount({ allowCopy: false })
    dragAcrossText(canvas)
    // give any async hit-test a chance to land before asserting absence
    await waitFor(() => expect(canvas.getBoundingClientRect().width).toBe(816))
    expect(copyButton(container)).toBeNull()
  })

  test('Cmd/Ctrl+C is swallowed and never reaches the clipboard', async () => {
    const { container } = await mount({ allowCopy: false })
    const region = container.querySelector('[data-officeview-root]') as HTMLElement
    const event = new KeyboardEvent('keydown', { key: 'c', metaKey: true, bubbles: true, cancelable: true })
    fireEvent(region, event)
    expect(event.defaultPrevented).toBe(true)
    expect(writeText).not.toHaveBeenCalled()
  })

  test('a native copy event inside the viewer is cancelled', async () => {
    const { container } = await mount({ allowCopy: false })
    const region = container.querySelector('[data-officeview-root]') as HTMLElement
    const copy = new Event('copy', { bubbles: true, cancelable: true })
    fireEvent(region, copy)
    expect(copy.defaultPrevented).toBe(true)
  })

  test('the viewer opts out of browser text selection', async () => {
    const { container } = await mount({ allowCopy: false })
    const region = container.querySelector('[data-officeview-root]') as HTMLElement
    expect(region.style.userSelect).toBe('none')
  })
})

describe('allowCopy (default, unchanged behaviour)', () => {
  test('selection and the copy button still work when prevention is off', async () => {
    const { container, canvas } = await mount()
    dragAcrossText(canvas)
    await waitFor(() => expect(copyButton(container)).toBeTruthy())
    fireEvent.click(copyButton(container) as HTMLElement)
    await waitFor(() => expect(writeText).toHaveBeenCalled())
    const region = container.querySelector('[data-officeview-root]') as HTMLElement
    expect(region.style.userSelect).toBe('')
  })

  test('no print guard is installed by default', async () => {
    const { unmount } = await mount()
    expect(printGuards().length).toBe(0)
    unmount()
  })
})

describe('allowPrint={false} (print prevention)', () => {
  test('installs a print-only rule that hides the viewer', async () => {
    const { container, unmount } = await mount({ allowPrint: false })
    const guards = printGuards()
    expect(guards.length).toBe(1)
    expect(guards[0].textContent).toContain('@media print')
    expect(guards[0].textContent).toContain('data-officeview-root')
    expect(container.querySelector('[data-officeview-root]')).toBeTruthy()
    unmount()
  })

  test('the rule is removed on unmount so the host page prints normally again', async () => {
    const { unmount } = await mount({ allowPrint: false })
    expect(printGuards().length).toBe(1)
    unmount()
    expect(printGuards().length).toBe(0)
  })

  test('Cmd/Ctrl+P is swallowed', async () => {
    const { container } = await mount({ allowPrint: false })
    const region = container.querySelector('[data-officeview-root]') as HTMLElement
    const event = new KeyboardEvent('keydown', { key: 'p', metaKey: true, bubbles: true, cancelable: true })
    fireEvent(region, event)
    expect(event.defaultPrevented).toBe(true)
  })

  test('Cmd/Ctrl+P still works when print prevention is off', async () => {
    const { container } = await mount()
    const region = container.querySelector('[data-officeview-root]') as HTMLElement
    const event = new KeyboardEvent('keydown', { key: 'p', metaKey: true, bubbles: true, cancelable: true })
    fireEvent(region, event)
    expect(event.defaultPrevented).toBe(false)
  })
})

describe('copy and print prevention are independent', () => {
  test('allowPrint={false} alone leaves text selection intact', async () => {
    const { container, canvas } = await mount({ allowPrint: false })
    dragAcrossText(canvas)
    await waitFor(() => expect(copyButton(container)).toBeTruthy())
  })

  test('allowCopy={false} alone installs no print guard', async () => {
    const { unmount } = await mount({ allowCopy: false })
    expect(printGuards().length).toBe(0)
    unmount()
  })

  test('both off: no selection affordances and no print output', async () => {
    const { container, canvas } = await mount({ allowCopy: false, allowPrint: false })
    dragAcrossText(canvas)
    await waitFor(() => expect(canvas.getBoundingClientRect().width).toBe(816))
    expect(copyButton(container)).toBeNull()
    expect(printGuards().length).toBe(1)
  })
})