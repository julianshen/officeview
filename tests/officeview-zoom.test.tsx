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

/**
 * An IntersectionObserver whose reports the test drives by hand. `visible` is
 * what every observed canvas is reported as; `null` means "never report",
 * which is how an off-screen page looks to the viewer.
 */
class ControllableObserver {
  static visible: boolean | null = null
  static instances: ControllableObserver[] = []
  private targets: Element[] = []
  constructor(private readonly callback: IntersectionObserverCallback) {
    ControllableObserver.instances.push(this)
  }
  observe(el: Element): void {
    this.targets.push(el)
  }
  unobserve(): void {}
  disconnect(): void {
    this.targets = []
  }
  takeRecords(): [] {
    return []
  }
  root = null
  rootMargin = ''
  thresholds: number[] = []
  /** Deliver a report to every registered target, as a real observer would. */
  static report(visible: boolean): void {
    ControllableObserver.visible = visible
    for (const io of ControllableObserver.instances) {
      io.callback(
        io.targets.map((target) => ({ target, isIntersecting: visible } as IntersectionObserverEntry)),
        io as unknown as IntersectionObserver,
      )
    }
  }
  static reset(): void {
    ControllableObserver.instances = []
    ControllableObserver.visible = null
  }
}

/** Install the controllable observer for the duration of `fn`. */
async function withObserver<T>(fn: () => Promise<T>): Promise<T> {
  const original = (globalThis as { IntersectionObserver?: unknown }).IntersectionObserver
  ControllableObserver.reset()
  ;(globalThis as { IntersectionObserver?: unknown }).IntersectionObserver = ControllableObserver
  try {
    return await fn()
  } finally {
    ControllableObserver.reset()
    ;(globalThis as { IntersectionObserver?: unknown }).IntersectionObserver = original
  }
}

async function mount() {
  const doc = await docWithPages()
  const utils = render(<OfficeDoc document={doc} />)
  await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(1))
  // The canvas backing store is assigned in an effect AFTER the element exists.
  // Waiting on existence alone can capture the pre-layout state, which made
  // later width assertions depend on when the effect had run.
  await waitFor(() => {
    for (const c of Array.from(utils.container.querySelectorAll('canvas'))) {
      expect((c as HTMLCanvasElement).width).toBeGreaterThan(0)
      expect((c as HTMLCanvasElement).height).toBeGreaterThan(0)
    }
  })
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

describe('<OfficeDoc> hi-dpi re-render when zoomed', () => {
  test('backing store grows with zoom and returns to base on reset', async () => {
    const utils = await mount()
    const canvasOf = () => utils.container.querySelector('canvas') as HTMLCanvasElement
    const base = canvasOf().width

    fireEvent.click(utils.container.querySelector('[aria-label="Zoom in"]') as HTMLButtonElement)
    await waitFor(() => expect(canvasOf().width).toBeGreaterThan(base))
    const zoomed = canvasOf().width
    // 1.5x zoom with devicePixelRatio 1 => 1.5x the backing width
    expect(zoomed / base).toBeCloseTo(1.5, 2)

    fireEvent.click(utils.container.querySelector('[aria-label="Reset zoom"]') as HTMLButtonElement)
    await waitFor(() => expect(canvasOf().width).toBe(base))
  })

  test('effective scale is capped at 3x to bound memory', async () => {
    const utils = await mount()
    const canvasOf = () => utils.container.querySelector('canvas') as HTMLCanvasElement
    const base = canvasOf().width
    // 1.5 -> 2.25 -> 3.375 (capped to 3)
    const plus = utils.container.querySelector('[aria-label="Zoom in"]') as HTMLButtonElement
    fireEvent.click(plus)
    fireEvent.click(plus)
    fireEvent.click(plus)
    await waitFor(() => expect(canvasOf().width).toBe(base * 3))
  })

  test('off-screen pages stay at base scale while zoomed (memory bound)', async () => {
    await withObserver(async () => {
      const utils = await mount()
      const canvasOf = () => utils.container.querySelector('canvas') as HTMLCanvasElement
      const base = canvasOf().width
      const plus = utils.container.querySelector('[aria-label="Zoom in"]') as HTMLButtonElement

      // Zoom, and wait for the INVARIANT this test claims rather than for an
      // unrelated signal. The previous version waited for the CSS transform and
      // then asserted the backing store immediately, which raced the effect that
      // applies the visibility decision and made this test intermittently fail.
      fireEvent.click(plus)
      await waitFor(() => expect(canvasOf().width).toBe(base))
      expect(content(utils).style.transform).toContain('scale(1.5)')

      // Zooming further must not allocate a bigger backing store either.
      fireEvent.click(plus)
      await waitFor(() => expect(canvasOf().width).toBe(base))
    })
  })

  test('a page reported visible DOES get the zoomed backing store', async () => {
    // The discriminating counterpart to the test above: without it, "stays at
    // base scale" could pass simply because boosting is broken entirely.
    await withObserver(async () => {
      const utils = await mount()
      const canvasOf = () => utils.container.querySelector('canvas') as HTMLCanvasElement
      const base = canvasOf().width

      fireEvent.click(utils.container.querySelector('[aria-label="Zoom in"]') as HTMLButtonElement)
      // still pessimistic: no report has arrived yet
      await waitFor(() => expect(canvasOf().width).toBe(base))

      // now the observer says the page is on screen
      await act(async () => {
        ControllableObserver.report(true)
      })
      await waitFor(() => expect(canvasOf().width).toBeGreaterThan(base))
    })
  })

  test('a page reported off-screen drops back to base scale', async () => {
    // And the reverse transition, so the memory bound is proven in both
    // directions rather than only at the moment of zooming in.
    await withObserver(async () => {
      const utils = await mount()
      const canvasOf = () => utils.container.querySelector('canvas') as HTMLCanvasElement
      const base = canvasOf().width

      fireEvent.click(utils.container.querySelector('[aria-label="Zoom in"]') as HTMLButtonElement)
      await act(async () => {
        ControllableObserver.report(true)
      })
      await waitFor(() => expect(canvasOf().width).toBeGreaterThan(base))
      const boosted = canvasOf().width

      await act(async () => {
        ControllableObserver.report(false)
      })
      await waitFor(() => expect(canvasOf().width).toBe(base))
      expect(boosted).toBeGreaterThan(base)
    })
  })
})
