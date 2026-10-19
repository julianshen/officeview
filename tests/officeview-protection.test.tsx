import { describe, test, expect, vi, beforeEach } from 'vitest'
import { act, render, waitFor, fireEvent } from '@testing-library/react'
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
  return { ...utils, canvas, doc }
}

/** Drag across the first line of text, the way a reader would. */
function dragAcrossText(canvas: HTMLCanvasElement): void {
  fireEvent.pointerDown(canvas, { pointerId: 1, clientX: 100, clientY: 104, button: 0, bubbles: true })
  fireEvent.pointerMove(canvas, { pointerId: 1, clientX: 260, clientY: 104, bubbles: true })
  fireEvent.pointerUp(canvas, { pointerId: 1, clientX: 260, clientY: 104, bubbles: true })
}

const copyButton = (c: HTMLElement) => c.querySelector('[data-testid="officeview-copy-selection"]')
const printGuards = () => document.querySelectorAll('[data-ov-print="false"]')

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
  test('marks the denied instance for the scoped print rule', async () => {
    const { container, unmount } = await mount({ allowPrint: false })
    const guards = printGuards()
    expect(guards.length).toBe(1)
    expect(guards[0].hasAttribute('data-ov-instance')).toBe(true)
    expect(guards[0].contains(container.querySelector('[data-officeview-root]'))).toBe(true)
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

  test('Cmd/Ctrl+P is routed to isolated print when print prevention is off', async () => {
    const { container } = await mount()
    const region = container.querySelector('[data-officeview-root]') as HTMLElement
    const event = new KeyboardEvent('keydown', { key: 'p', metaKey: true, bubbles: true, cancelable: true })
    fireEvent(region, event)
    expect(event.defaultPrevented).toBe(true)
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

test('withdrawing copy permission while the API is pending blocks the legacy fallback', async () => {
  let reject!: (error: Error) => void
  writeText.mockImplementationOnce(() => new Promise<void>((_resolve, no) => { reject = no }))
  const previous = document.execCommand, exec = vi.fn(() => true)
  document.execCommand = exec as unknown as typeof document.execCommand
  try {
    const ui = await mount(); dragAcrossText(ui.canvas)
    await waitFor(() => expect(copyButton(ui.container)).toBeTruthy())
    fireEvent.click(copyButton(ui.container) as HTMLElement)
    await waitFor(() => expect(writeText).toHaveBeenCalled())
    ui.rerender(<OfficeDoc document={ui.doc} allowCopy={false} />)
    await act(async () => { reject(new Error('clipboard denied')); await Promise.resolve() })
    expect(exec).not.toHaveBeenCalled()
  } finally { document.execCommand = previous }
})
