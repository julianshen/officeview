/**
 * A driven slideshow rig (slideshow plan S1-22): the real lazy engine
 * (`runShow`) in a DOM with node-canvas, every clock injected. The show
 * clock, the frame driver, the wall timers and the monotonic clock are
 * replaced through `sessionDeps`; the surface measures `width` × `height`
 * CSS px (jsdom has no layout). Nothing waits on wall time.
 *
 * Used by tests/slideshow-corpus.test.ts (vitest, jsdom) and by
 * scripts/slideshow-compare.ts (bun, after it installs a jsdom window), so it
 * never imports vitest: `installShowSeams` patches and `restore` undoes.
 */
import { runShow } from '../../src/slideshow'
import { sessionDeps, type SessionController } from '../../src/slideshow/player/session'
import { ManualClock } from '../../src/slideshow/clock'
import type { ShowStart, SlideShowCallbacks } from '../../src/components/present'
import { OfficePackage } from '../../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../../src/pptx/parse'
import { enrichSlideAppearance } from '../../src/pptx/appearance'
import type { PptxDocument } from '../../src/pptx/types'
import type { RenderState, SlideShowDiagnostic } from '../../src/slideshow/types'
import { slideSize } from '../../src/slideshow/render/scene'

export interface ShowSeams {
  clock: ManualClock
  /** Moves the wall timers (cursor hide, restore listener) forward. */
  advanceWall(ms: number): void
  restore(): void
}

/** Replaces the session's clocks, frame driver and timers, and sizes every element's box to `width` × `height`. */
export function installShowSeams(size: { width: number; height: number } = { width: 1024, height: 768 }): ShowSeams {
  const clock = new ManualClock()
  const timers: Array<{ at: number; cb: () => void; id: number }> = []
  let wall = 0
  let next = 1
  const saved = { clock: sessionDeps.clock, frames: sessionDeps.frames, timer: sessionDeps.timer, now: sessionDeps.now, backend: sessionDeps.backend }
  const rect = HTMLElement.prototype.getBoundingClientRect
  // The session sizes the surface from the host (clientWidth, 0 in jsdom) or the window.
  const inner = { w: window.innerWidth, h: window.innerHeight }
  Object.defineProperty(window, 'innerWidth', { value: size.width, configurable: true, writable: true })
  Object.defineProperty(window, 'innerHeight', { value: size.height, configurable: true, writable: true })
  sessionDeps.clock = () => clock
  sessionDeps.frames = () => () => {}
  sessionDeps.timer = () => ({
    set: (cb, ms) => { const id = next++; timers.push({ at: wall + ms, cb, id }); return id },
    clear: (id) => { const i = timers.findIndex(t => t.id === id); if (i >= 0) timers.splice(i, 1) },
  })
  sessionDeps.now = () => wall
  HTMLElement.prototype.getBoundingClientRect = function () {
    return { left: 0, top: 0, right: size.width, bottom: size.height, width: size.width, height: size.height, x: 0, y: 0, toJSON() {} } as DOMRect
  }
  return {
    clock,
    advanceWall(ms) {
      wall += ms
      for (const t of [...timers].sort((a, b) => a.at - b.at)) if (t.at <= wall && timers.includes(t)) { timers.splice(timers.indexOf(t), 1); t.cb() }
    },
    restore() {
      Object.assign(sessionDeps, saved)
      HTMLElement.prototype.getBoundingClientRect = rect
      Object.defineProperty(window, 'innerWidth', { value: inner.w, configurable: true, writable: true })
      Object.defineProperty(window, 'innerHeight', { value: inner.h, configurable: true, writable: true })
    },
  }
}

/**
 * Parses a deck and runs the slide appearance enrichment (backgrounds, master and layout artwork), as the viewer's
 * `getPaintables` has done by the time its Present button or `present()` starts a show.
 */
export async function loadDeck(bytes: Uint8Array): Promise<PptxDocument> {
  const doc = await parsePptx(await OfficePackage.load(bytes))
  await enrichSlideAppearance(doc)
  return doc
}

/** Lets pending promise work (metadata reads, image decodes, ensure chains) settle. */
export const flush = async (): Promise<void> => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

export interface ShowRig {
  s: SessionController
  host: HTMLElement
  log: string[]
  diagnostics: SlideShowDiagnostic[]
  /** Advances the show clock by `ms` and runs frames until pending work settles. */
  tick(ms?: number): Promise<void>
  /** Presses a key on the host. */
  key(k: string): Promise<void>
  /** Clicks at slide px (x, y) of the shown slide. */
  click(x: number, y: number): Promise<void>
  /** The shown slide's render state now. */
  state(): RenderState
  /** The offscreen text list (spec 9). */
  text(): string[]
  /** The canvas that holds the composited frame now (the compositor's, or the terminal output). */
  frameCanvas(): HTMLCanvasElement | undefined
  exit(): void
}

export interface ShowOptions {
  seams: ShowSeams
  size?: { width: number; height: number }
  motion?: 'auto' | 'reduced' | 'full'
  callbacks?: Partial<SlideShowCallbacks>
}

/** Starts a show of `doc` from the beginning, as the eager start sequence hands it to the lazy chunk. */
export async function startShow(doc: PptxDocument, opts: ShowOptions): Promise<ShowRig> {
  const { seams } = opts
  const size = opts.size ?? { width: 1024, height: 768 }
  const host = document.createElement('div')
  host.tabIndex = -1
  host.setAttribute('data-officeview-show', '')
  host.style.visibility = 'visible'
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const diagnostics: SlideShowDiagnostic[] = []
  const start = {
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: {
      hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null,
      open: () => null,
    },
    fullscreen: Promise.resolve('entered' as const),
    retryFullscreen: async () => 'entered' as const,
    from: 'beginning' as const, motion: opts.motion ?? 'auto',
    callbacks: {
      onStart: () => log.push('onStart'), onEnd: () => log.push('onEnd'),
      onSlideChange: (i: number, b: number) => log.push(`change:${i}:${b}`),
      onDiagnostic: (d: SlideShowDiagnostic) => { diagnostics.push(d); log.push(`diag:${d.code}`) },
      currentSlide: () => 0, restoreFocus: () => {}, scrollToSlide: () => {},
      ...opts.callbacks,
    },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted,
    cleanup: () => { ac.abort(); host.style.visibility = 'hidden' },
  } as ShowStart
  const s = await runShow(start)
  await flush()
  s.tick()
  const surface = () => host.querySelector<HTMLElement>('[data-officeview-show-surface]')!
  const tick = async (ms = 0) => { seams.clock.advance(ms); s.tick(); await flush(); s.tick() }
  return {
    s, host, log, diagnostics, tick,
    async key(k) { host.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true })); await flush(); s.tick() },
    async click(x, y) {
      const st = s.stage!
      const b = st.backingStore
      const { w, h } = slideSize(doc.slides[typeof s.state.slide === 'number' ? s.state.slide : 0])
      // Slide px → backing px (the viewport) → client px (object-fit: contain of the backing store in the surface).
      const v = Math.min(b.w / w, b.h / h)
      const bx = (b.w - w * v) / 2 + x * v, by = (b.h - h * v) / 2 + y * v
      const k = Math.min(size.width / b.w, size.height / b.h)
      const cx = (size.width - b.w * k) / 2 + bx * k, cy = (size.height - b.h * k) / 2 + by * k
      surface().dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: cx, clientY: cy }))
      await flush()
      s.tick()
    },
    state: () => s.stage!.stateAt(s.cursor, seams.clock.now()),
    text: () => [...host.querySelectorAll('[data-officeview-show-text] li')].map(li => li.textContent || li.getAttribute('aria-label') || ''),
    frameCanvas() {
      if (s.stage?.level === 'terminal') return host.querySelector<HTMLCanvasElement>('[data-officeview-show-canvas]') ?? undefined
      // Canvas 2D composites on the output canvas (PR #18 review P2); WebGL2 on its own canvas in the stage.
      return host.querySelector<HTMLCanvasElement>('canvas[data-slideshow-backend="canvas2d"]')
        ?? host.querySelector<HTMLCanvasElement>('[data-officeview-show-stage] canvas[data-slideshow-backend="webgl2"]') ?? undefined
    },
    exit() { s.exit() },
  }
}
