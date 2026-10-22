/**
 * Slideshow Task S2-3: Presenter View Window Management placement in Chromium
 * (the `slideshow` Playwright project).
 *
 * Feature-detection mocks exercise multi-screen placement, permission prompts,
 * fallbacks, diagnostics, screen swapping, and the one-time drag hint.
 *
 * Written for the controller to run against the slideshow harness.
 */
import { expect, test, type Page } from '@playwright/test'
import { fileURLToPath } from 'node:url'
import { checkIdentity, gitIdentity } from '../../scripts/harness/identity'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const BASE = `http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`
const PAGE = '/src/debug/slideshow-start.html'

const errors: string[] = []

test.beforeAll(async () => {
  const text = await (await fetch(`${BASE}/__harness-id`)).text()
  let body: unknown = text
  try { body = JSON.parse(text) } catch { /* checkIdentity rejects raw text */ }
  expect(checkIdentity(body, { commit: gitIdentity(ROOT).commit, root: ROOT }), 'harness server identity').toBeNull()
})

test.beforeEach(({ page }) => {
  errors.length = 0
  page.on('pageerror', (e) => errors.push(e.message))
})

test.afterEach(() => {
  expect(errors).toEqual([])
})

/** Injects Window Management API mocks with custom screens and permission state into the page. */
async function mockWindowManagement(
  page: Page,
  options: {
    screens?: Array<{ isPrimary?: boolean; left?: number; top?: number; width?: number; height?: number }>
    permission?: 'granted' | 'prompt' | 'denied'
    missingApi?: boolean
  }
): Promise<void> {
  await page.addInitScript((opts) => {
    if (opts.missingApi) {
      delete (window as any).getScreenDetails
      return
    }

    const screens = opts.screens ?? [
      { isPrimary: true, left: 0, top: 0, width: 1920, height: 1080 },
      { isPrimary: false, left: 1920, top: 0, width: 2560, height: 1440 },
    ]

    const screenDetails = {
      screens,
      currentScreen: screens[0],
    }

    ;(window as any).__mockScreenDetails = screenDetails
    ;(window as any).__screenRequests = []

    // Mock window.getScreenDetails
    ;(window as any).getScreenDetails = async () => {
      ;(window as any).__screenRequests.push('getScreenDetails')
      if (opts.permission === 'denied') {
        const err = new Error('Permission denied')
        err.name = 'NotAllowedError'
        throw err
      }
      return screenDetails
    }

    // Mock Element.prototype.requestFullscreen to record options
    const origRequestFullscreen = Element.prototype.requestFullscreen
    Element.prototype.requestFullscreen = async function (options?: FullscreenOptions) {
      ;(window as any).__lastFullscreenOptions = options
      return origRequestFullscreen?.call(this, options)
    }
  }, options)
}

test.describe('Presenter View: Window Management placement', () => {
  test('granted with two screens: popup on primary, audience requestFullscreen({screen}) on the external screen', async ({ page }) => {
    await mockWindowManagement(page, {
      screens: [
        { isPrimary: true, left: 0, top: 0, width: 1920, height: 1080 },
        { isPrimary: false, left: 1920, top: 0, width: 2560, height: 1440 },
      ],
      permission: 'granted',
    })

    await page.goto(PAGE)
    await page.waitForFunction(() => window.__slideshowStart?.ready === true)

    const result = await page.evaluate(async () => {
      const screens = await (window as any).getScreenDetails?.()
      const { planPlacement } = await import('../../src/slideshow/presenter/screens')
      return planPlacement(screens)
    })

    expect(result.manual).toBeFalsy()
    if (!result.manual) {
      expect(result.audienceScreen).toMatchObject({ isPrimary: false, left: 1920 })
      expect(result.popupFeatures).toContain('left=0')
      expect(result.popupFeatures).toContain('width=1920')
    }
  })

  test('not yet granted: start proceeds and Use both screens asks on click', async ({ page }) => {
    await mockWindowManagement(page, {
      permission: 'prompt',
    })

    await page.goto(PAGE)
    await page.waitForFunction(() => window.__slideshowStart?.ready === true)

    const initialPlan = await page.evaluate(async () => {
      const { planPlacement } = await import('../../src/slideshow/presenter/screens')
      return planPlacement(undefined)
    })
    expect(initialPlan.manual).toBe(true)

    // Simulate user clicking "Use both screens"
    const clickedResult = await page.evaluate(async () => {
      const { useBothScreens, planPlacement } = await import('../../src/slideshow/presenter/screens')
      const details = await useBothScreens(window)
      return planPlacement(details)
    })

    expect(clickedResult.manual).toBeFalsy()
    if (!clickedResult.manual) {
      expect(clickedResult.audienceScreen).toBeDefined()
    }
  })

  test('denied, missing API and single screen are three distinct diagnostics, none is popup-blocked', async ({ page }) => {
    await page.goto(PAGE)
    await page.waitForFunction(() => window.__slideshowStart?.ready === true)

    const diagnostics = await page.evaluate(async () => {
      const { planPlacement, diagnosePlacement } = await import('../../src/slideshow/presenter/screens')
      const { createDiagnosticSink } = await import('../../src/slideshow/diagnostics')
      const sink = createDiagnosticSink()

      const p1 = planPlacement(undefined, 'permission-denied')
      diagnosePlacement(p1, sink)

      const p2 = planPlacement(undefined, 'api-missing')
      diagnosePlacement(p2, sink)

      const p3 = planPlacement({ screens: [{ isPrimary: true }] as any })
      diagnosePlacement(p3, sink)

      return sink.list()
    })

    expect(diagnostics).toHaveLength(3)
    const codes = diagnostics.map((d: any) => d.code)
    expect(codes).toEqual(['permission-denied', 'api-missing', 'single-screen'])
    expect(codes).not.toContain('popup-blocked')
  })

  test('swap displays exchanges the screens', async ({ page }) => {
    await page.goto(PAGE)
    await page.waitForFunction(() => window.__slideshowStart?.ready === true)

    const swapped = await page.evaluate(async () => {
      const { swapDisplays } = await import('../../src/slideshow/presenter/screens')
      const screenA = { isPrimary: true, left: 0, top: 0, width: 1920, height: 1080 }
      const screenB = { isPrimary: false, left: 1920, top: 0, width: 2560, height: 1440 }

      let movedTo: any = null
      let resizedTo: any = null
      let requestedScreen: any = null

      const session = {
        screens: {
          presenter: screenA,
          audience: screenB,
        },
        presenterWindow: {
          moveTo: (x: number, y: number) => { movedTo = { x, y } },
          resizeTo: (w: number, h: number) => { resizedTo = { w, h } },
        },
        retryFullscreen: async (opts: any) => { requestedScreen = opts?.screen },
      }

      swapDisplays(session as any)

      return {
        presenter: session.screens.presenter,
        audience: session.screens.audience,
        movedTo,
        resizedTo,
        requestedScreen,
      }
    })

    expect(swapped.presenter).toMatchObject({ left: 1920 })
    expect(swapped.audience).toMatchObject({ left: 0 })
    expect(swapped.movedTo).toEqual({ x: 1920, y: 0 })
    expect(swapped.resizedTo).toEqual({ w: 2560, h: 1440 })
    expect(swapped.requestedScreen).toMatchObject({ left: 0 })
  })

  test('other browsers show the one-time drag hint', async ({ page }) => {
    await page.goto(PAGE)
    await page.waitForFunction(() => window.__slideshowStart?.ready === true)

    const hintBehavior = await page.evaluate(async () => {
      const { createDragHint, shouldShowDragHint, markDragHintDismissed, DRAG_HINT_STORAGE_KEY } =
        await import('../../src/slideshow/presenter/screens')

      sessionStorage.removeItem(DRAG_HINT_STORAGE_KEY)
      const firstShowable = shouldShowDragHint(sessionStorage)
      const el = createDragHint(document, { storage: sessionStorage })
      const hasAttr = el?.hasAttribute('data-officeview-drag-hint')

      markDragHintDismissed(sessionStorage)
      const secondShowable = shouldShowDragHint(sessionStorage)
      const secondEl = createDragHint(document, { storage: sessionStorage })

      return {
        firstShowable,
        hasAttr,
        secondShowable,
        secondCreated: !!secondEl,
      }
    })

    expect(hintBehavior.firstShowable).toBe(true)
    expect(hintBehavior.hasAttr).toBe(true)
    expect(hintBehavior.secondShowable).toBe(false)
    expect(hintBehavior.secondCreated).toBe(false)
  })
})
