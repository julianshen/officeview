import { describe, test, expect } from 'vitest'
import { render, waitFor, fireEvent, act } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx } from '../src/testdata/ooxml-builders'

async function docWithPages() {
  const paras = Array.from({ length: 120 }, (_, i) => ({ runs: [{ text: `line ${i}` }] }))
  return parseDocx(await OfficePackage.load(await buildDocx(paras)))
}

async function mount() {
  const doc = await docWithPages()
  const utils = render(<OfficeDoc document={doc} />)
  await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(1))
  return utils
}

function content(utils: { container: HTMLElement }): HTMLElement {
  return utils.container.querySelector('[data-testid="officeview-content"]') as HTMLElement
}

function region(utils: { container: HTMLElement }): HTMLElement {
  return utils.container.querySelector('[role="region"]') as HTMLElement
}

describe('<OfficeDoc> zoom', () => {
  test('starts at fit with a neutral transform', async () => {
    const utils = await mount()
    const el = content(utils)
    expect(el.style.transform).toBe('none')
    expect(region(utils).style.overflow).toBe('auto')
    // native touch scrolling is preserved at fit
    expect(region(utils).style.touchAction).toBe('pan-y')
  })

  test('zoom in button applies a scale transform and disables native scroll', async () => {
    const utils = await mount()
    const btn = utils.container.querySelector('[aria-label="Zoom in"]') as HTMLButtonElement
    fireEvent.click(btn)
    await waitFor(() => expect(content(utils).style.transform).toContain('scale(1.5)'))
    expect(region(utils).style.overflow).toBe('hidden')
    expect(region(utils).style.touchAction).toBe('none')
  })

  test('zoom percentage readout tracks the zoom level and resets', async () => {
    const utils = await mount()
    const reset = utils.container.querySelector('[aria-label="Reset zoom"]') as HTMLButtonElement
    expect(reset.textContent).toBe('100%')
    fireEvent.click(utils.container.querySelector('[aria-label="Zoom in"]') as HTMLButtonElement)
    await waitFor(() => expect(reset.textContent).toBe('150%'))
    fireEvent.click(reset)
    await waitFor(() => expect(reset.textContent).toBe('100%'))
    expect(content(utils).style.transform).toBe('none')
  })

  test('zoom out is disabled at fit', async () => {
    const utils = await mount()
    const out = utils.container.querySelector('[aria-label="Zoom out"]') as HTMLButtonElement
    expect(out.disabled).toBe(true)
  })

  test('double click toggles between fit and 2x', async () => {
    const utils = await mount()
    const el = region(utils)
    fireEvent.doubleClick(el, { clientX: 100, clientY: 100 })
    await waitFor(() => expect(content(utils).style.transform).toContain('scale(2)'))
    fireEvent.doubleClick(el, { clientX: 100, clientY: 100 })
    await waitFor(() => expect(content(utils).style.transform).toBe('none'))
  })

  test('ctrl+wheel zooms, plain wheel does not', async () => {
    const utils = await mount()
    const el = region(utils)
    fireEvent.wheel(el, { deltaY: -100 })
    expect(content(utils).style.transform).toBe('none')
    fireEvent.wheel(el, { deltaY: -100, ctrlKey: true })
    await waitFor(() => expect(content(utils).style.transform).toContain('scale('))
  })

  test('keyboard +,-,0 control zoom', async () => {
    const utils = await mount()
    const el = region(utils)
    fireEvent.keyDown(el, { key: '+' })
    await waitFor(() => expect(content(utils).style.transform).toContain('scale(1.5)'))
    fireEvent.keyDown(el, { key: '0' })
    await waitFor(() => expect(content(utils).style.transform).toBe('none'))
  })

  test('two-pointer pinch zooms in, drag pans when zoomed', async () => {
    const utils = await mount()
    const el = region(utils)
    // jsdom has no layout, so give the region a size for viewport math
    Object.defineProperty(el, 'clientWidth', { value: 400, configurable: true })
    Object.defineProperty(el, 'clientHeight', { value: 800, configurable: true })
    el.getBoundingClientRect = () => ({ left: 0, top: 0, width: 400, height: 800, right: 400, bottom: 800, x: 0, y: 0, toJSON: () => ({}) })

    const pointerOpts = (id: number, x: number, y: number) => ({ pointerId: id, clientX: x, clientY: y, pointerType: 'touch', bubbles: true, isPrimary: id === 1 })
    // fingers 100px apart, moving to 200px apart => 2x
    fireEvent.pointerDown(el, pointerOpts(1, 150, 400))
    fireEvent.pointerDown(el, pointerOpts(2, 250, 400))
    fireEvent.pointerMove(el, pointerOpts(2, 350, 400))
    await waitFor(() => expect(content(utils).style.transform).toContain('scale(2)'))
    const zoomedTransform = content(utils).style.transform
    expect(zoomedTransform).toContain('translate3d')

    // single-finger drag pans
    fireEvent.pointerUp(el, pointerOpts(2, 350, 400))
    const before = content(utils).style.transform
    fireEvent.pointerDown(el, pointerOpts(1, 200, 400))
    fireEvent.pointerMove(el, pointerOpts(1, 140, 340))
    await waitFor(() => expect(content(utils).style.transform).not.toBe(before))
  })

  test('zoom controls can be hidden', async () => {
    const doc = await docWithPages()
    const utils = render(<OfficeDoc document={doc} showZoomControls={false} />)
    await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
    expect(utils.container.querySelector('[data-testid="officeview-zoom-controls"]')).toBeNull()
    void act
  })
})
