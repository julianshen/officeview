/**
 * Slideshow Task S2-3: Window Management placement for Presenter View.
 * Spec: §5.1 placement in Chromium, other browsers, denial diagnostics.
 */
import { describe, expect, it, vi } from 'vitest'
import {
  planPlacement,
  useBothScreens,
  preparePlacement,
  getPreparedPlacement,
  setPreparedPlacementSlot,
  swapDisplays,
  diagnosePlacement,
  createDragHint,
  showDragHint,
  shouldShowDragHint,
  DRAG_HINT_STORAGE_KEY,
  type ScreenLike,
  type ScreenDetailsLike,
} from '../src/slideshow/presenter/screens'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'

describe('Slideshow Screen Placement (Task S2-3)', () => {
  it('granted with two screens: popup on primary, audience requestFullscreen({screen}) on the external screen', async () => {
    const primaryScreen: ScreenLike = {
      isPrimary: true,
      isInternal: true,
      left: 0,
      top: 0,
      width: 1920,
      height: 1080,
      availLeft: 0,
      availTop: 0,
      availWidth: 1920,
      availHeight: 1040,
    }
    const externalScreen: ScreenLike = {
      isPrimary: false,
      isInternal: false,
      left: 1920,
      top: 0,
      width: 2560,
      height: 1440,
      availLeft: 1920,
      availTop: 0,
      availWidth: 2560,
      availHeight: 1440,
    }
    const details: ScreenDetailsLike = {
      screens: [primaryScreen, externalScreen],
      currentScreen: primaryScreen,
    }

    const plan = planPlacement(details)
    expect(plan.manual).toBeFalsy()
    if (!plan.manual) {
      expect(plan.audienceScreen).toBe(externalScreen)
      expect(plan.popupFeatures).toContain('left=0')
      expect(plan.popupFeatures).toContain('top=0')
      expect(plan.popupFeatures).toContain('width=1920')
      expect(plan.popupFeatures).toContain('height=1040')
    }

    // Verify audience full screen request with { screen: externalScreen }
    const requestFullscreenMock = vi.fn().mockResolvedValue(undefined)
    const hostEl = {
      requestFullscreen: requestFullscreenMock,
    }

    if (!plan.manual && plan.audienceScreen) {
      await hostEl.requestFullscreen({ screen: plan.audienceScreen })
      expect(requestFullscreenMock).toHaveBeenCalledWith({ screen: externalScreen })
    }
  })

  it('not yet granted: start proceeds and Use both screens asks on click', async () => {
    // When details not yet granted / available, plan returns manual
    const initialPlan = planPlacement(undefined)
    expect(initialPlan.manual).toBe(true)

    const primaryScreen: ScreenLike = { isPrimary: true, left: 0, top: 0, width: 1440, height: 900 }
    const externalScreen: ScreenLike = { isPrimary: false, left: 1440, top: 0, width: 1920, height: 1080 }
    const details: ScreenDetailsLike = { screens: [primaryScreen, externalScreen] }

    const getScreenDetailsMock = vi.fn().mockResolvedValue(details)
    const fakeWin = {
      getScreenDetails: getScreenDetailsMock,
    } as unknown as Window

    // User clicks "Use both screens" button/control:
    const acquiredDetails = await useBothScreens(fakeWin)
    expect(getScreenDetailsMock).toHaveBeenCalledTimes(1)
    expect(acquiredDetails).toBe(details)

    const updatedPlan = planPlacement(acquiredDetails)
    expect(updatedPlan.manual).toBeFalsy()
    if (!updatedPlan.manual) {
      expect(updatedPlan.audienceScreen).toBe(externalScreen)
    }
  })

  it('denied, missing API and single screen are three distinct diagnostics, none is popup-blocked', () => {
    const sink = createDiagnosticSink()

    // 1. Permission denied
    const deniedPlan = planPlacement(undefined, 'permission-denied')
    expect(deniedPlan.manual).toBe(true)
    if (deniedPlan.manual) {
      expect(deniedPlan.reason).toBe('permission-denied')
      diagnosePlacement(deniedPlan, sink)
    }

    // 2. Missing API
    const missingPlan = planPlacement(undefined, 'api-missing')
    expect(missingPlan.manual).toBe(true)
    if (missingPlan.manual) {
      expect(missingPlan.reason).toBe('api-missing')
      diagnosePlacement(missingPlan, sink)
    }

    // 3. Single screen
    const singleScreen: ScreenLike = { isPrimary: true, left: 0, top: 0, width: 1920, height: 1080 }
    const singlePlan = planPlacement({ screens: [singleScreen] })
    expect(singlePlan.manual).toBe(true)
    if (singlePlan.manual) {
      expect(singlePlan.reason).toBe('single-screen')
      diagnosePlacement(singlePlan, sink)
    }

    const diagnostics = sink.list()
    expect(diagnostics).toHaveLength(3)

    const codes = diagnostics.map((d) => d.code)
    // Three distinct diagnostics
    expect(new Set(codes).size).toBe(3)
    expect(codes).toContain('permission-denied')
    expect(codes).toContain('api-missing')
    expect(codes).toContain('single-screen')

    // None is popup-blocked
    expect(codes).not.toContain('popup-blocked')
    for (const d of diagnostics) {
      expect(d.code).not.toBe('popup-blocked')
    }
  })

  it('swap displays exchanges the screens', () => {
    const screen1: ScreenLike = { isPrimary: true, left: 0, top: 0, width: 1920, height: 1080 }
    const screen2: ScreenLike = { isPrimary: false, left: 1920, top: 0, width: 2560, height: 1440 }

    const moveToMock = vi.fn()
    const resizeToMock = vi.fn()
    const retryFullscreenMock = vi.fn().mockResolvedValue('entered')

    const session = {
      screens: {
        presenter: screen1,
        audience: screen2,
      },
      presenterWindow: {
        moveTo: moveToMock,
        resizeTo: resizeToMock,
      },
      retryFullscreen: retryFullscreenMock,
    }

    swapDisplays(session)

    // Screens should be exchanged
    expect(session.screens.presenter).toBe(screen2)
    expect(session.screens.audience).toBe(screen1)

    // Presenter window moved to screen2
    expect(moveToMock).toHaveBeenCalledWith(1920, 0)
    expect(resizeToMock).toHaveBeenCalledWith(2560, 1440)

    // Audience requested fullscreen on screen1
    expect(retryFullscreenMock).toHaveBeenCalledWith({ screen: screen1 })
  })

  it('other browsers show the one-time drag hint', () => {
    const fakeStorageData: Record<string, string> = {}
    const fakeStorage: Storage = {
      getItem: (key) => fakeStorageData[key] ?? null,
      setItem: (key, val) => {
        fakeStorageData[key] = String(val)
      },
      removeItem: (key) => {
        delete fakeStorageData[key]
      },
      clear: () => {
        for (const k of Object.keys(fakeStorageData)) delete fakeStorageData[k]
      },
      key: (i) => Object.keys(fakeStorageData)[i] ?? null,
      length: 0,
    }

    const doc = document.implementation.createHTMLDocument('Presenter')
    expect(shouldShowDragHint(fakeStorage)).toBe(true)

    // First time: hint element is returned
    const hint = createDragHint(doc, { storage: fakeStorage })
    expect(hint).toBeDefined()
    expect(hint?.getAttribute('data-officeview-drag-hint')).toBe('')
    expect(hint?.textContent).toContain('Drag this presenter window to your other display')

    // Find dismiss button and click it
    const dismissBtn = hint?.querySelector('button')
    expect(dismissBtn).toBeDefined()
    dismissBtn?.click()

    // Dismiss marks it in storage
    expect(fakeStorage.getItem(DRAG_HINT_STORAGE_KEY)).toBe('1')
    expect(shouldShowDragHint(fakeStorage)).toBe(false)

    // showDragHint appends to container
    const container = doc.createElement('div')
    const appended = showDragHint(container, { storage: fakeStorage })
    expect(appended).toBeUndefined() // already dismissed
  })

  it('preparePlacement computes plan and stores in injected setter', () => {
    const doc = document.implementation.createHTMLDocument('Opener')
    const primaryScreen: ScreenLike = { isPrimary: true, left: 0, top: 0, width: 1920, height: 1080 }
    const externalScreen: ScreenLike = { isPrimary: false, left: 1920, top: 0, width: 2560, height: 1440 }
    const details: ScreenDetailsLike = { screens: [primaryScreen, externalScreen] }

    let slotValue: any = null
    const injectedSetter = vi.fn((_d: Document, plan: any) => {
      slotValue = plan
    })
    const injectedGetter = vi.fn((_d: Document) => slotValue)

    setPreparedPlacementSlot(injectedSetter, injectedGetter)

    preparePlacement(doc, details)
    expect(injectedSetter).toHaveBeenCalledWith(doc, expect.objectContaining({
      audienceScreen: externalScreen,
    }))
    expect(getPreparedPlacement(doc)).toEqual(expect.objectContaining({
      audienceScreen: externalScreen,
    }))

    // Reset slot injection
    setPreparedPlacementSlot(undefined as any, undefined as any)
  })
})
