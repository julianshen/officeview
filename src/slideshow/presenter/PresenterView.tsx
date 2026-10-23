/**
 * Presenter View (slideshow spec 5.2; 5.1 single-screen swap; 9).
 *
 * One state owner: the session controller in the opener. The view is a
 * React tree rendered from the opener's realm into a container of the
 * target document (`mountPresenterView`): the popup's theme root
 * (`host: 'window'`), or a box over the audience surface in the opener for
 * the single-screen swap (`host: 'swap'`). React creates its nodes through
 * the container's own document and listens on the container, so events and
 * focus stay realm-safe; every other listener goes on `doc`/its window.
 *
 * Contents, as PowerPoint, simplified:
 * - Current slide: a leased surface mirroring the rendered audience frame
 *   every frame (S2-4 `mirrorFrame`, from the session's frame hook, so a
 *   WebGL frame is copied before it is presented);
 * - Next: the settled next click checkpoint (S2-4 `nextCheckpoint` on the
 *   show clock, `primePreview` for the shown and next slides), painted with
 *   the stage's own state and images for the slides it holds;
 * - speaker notes with A+/A− (notes.tsx), the elapsed timer with pause and
 *   reset on the injected clock and the wall clock (timer.ts), "Slide n of m";
 * - controls: previous, next, black screen, the slide-number field, "See
 *   all slides" (grid.tsx), screen placement (placement.ts), end show;
 * - the diagnostics panel (diagnostics.tsx).
 *
 * Protection (spec 6.3): every presenter surface paints the watermark;
 * allowCopy=false makes the notes unselectable; allowPrint=false keeps the
 * popup out of print (the session marks its theme root).
 */
import { useEffect, useId, useLayoutEffect, useRef, useState, useSyncExternalStore, type ReactElement } from 'react'
import { createRoot } from 'react-dom/client'
import { flushSync } from 'react-dom'
import type { WatermarkOptions } from '../../core/watermark'
import type { PptxDocument } from '../../pptx/types'
import type { SlideCursor, SlideTimeline } from '../engine/steps'
import type { SessionController, ShowInput } from '../player/session'
import type { StagePreview } from '../player/stage'
import { nextVisible } from '../player/order'
import type { DiagnosticSink, NoteParagraph, ShowModel, SlideShowDiagnostic, SlideShowWorkspace } from '../types'
import { DiagnosticsPanel } from './diagnostics'
import { SlideGrid } from './grid'
import { NotesPane } from './notes'
import type { PresenterPlacement } from './placement'
import { createPresenterSurfaces, END_CURSOR, mirrorFrame, nextCheckpoint, paintPreview, previewTimeline, primePreview, releasePreview, type PresenterSurfaces, type PreviewOptions, type PreviewPosition } from './preview'
import { formatElapsed, formatWallClock, type ElapsedTimer } from './timer'

/** What the session hands Presenter View (additive `SessionController.presenterSource`). */
export interface PresenterSource {
  /** Change notifications: slide, build, screen, presenter mode, diagnostics, the timer's second and the clock's minute. */
  subscribe(cb: () => void): () => void
  readonly version: number
  readonly deck: PptxDocument
  notes(): readonly NoteParagraph[]
  diagnostics(): readonly SlideShowDiagnostic[]
  readonly sink: DiagnosticSink
  /** The show's elapsed timer (one per show, so it survives a swap or a closed popup). */
  readonly timer: ElapsedTimer
  /** The timer's monotonic clock reading. */
  now(): number
  /** Epoch milliseconds for the wall clock. */
  wall(): number
  /** The shown slide's timeline time on the show clock (never negative): what `next()` judges at. */
  slideTime(): number
  readonly model: ShowModel | undefined
  readonly previewOptions: PreviewOptions
  readonly workspace: SlideShowWorkspace
  readonly watermark?: WatermarkOptions
  readonly protection: { copy: boolean; print: boolean }
  /** The canvas holding the audience frame now (compositor or terminal output). */
  audienceFrame(): HTMLCanvasElement | undefined
  /** Runs after every drawn frame; returns the unsubscribe. */
  onFrame(cb: () => void): () => void
  /**
   * Composites the playing videos (DOM overlays on the audience, not in its canvas) into the mirror, where the
   * audience frame landed (`dest`, the mirror's px). Additive (PR #19 review 2).
   */
  drawMedia?(ctx: CanvasRenderingContext2D, dest: { x: number; y: number; w: number; h: number }): void
  /** Paint inputs for slide `index` (the stage's, or a static stand-in). */
  previewOf(index: number, cursor?: SlideCursor, timeline?: SlideTimeline): StagePreview
  /** The presenter key rule (spec 4.3): text controls and the notes keep their keys, Esc exits. */
  key(e: KeyboardEvent): void
  /** Presenter input: the show's input table, never gated by kiosk (controller ruling C6). */
  input(i: ShowInput): void
  /** Single-screen: swaps the full-screen surface between the audience and Presenter View. */
  toggleView(): void
  readonly placement: PresenterPlacement | undefined
  /** Re-renders after a change the session's key does not see (the timer controls). */
  touch(): void
}

export interface PresenterViewProps {
  session: SessionController
  host: 'window' | 'swap'
  doc: Document
}

const sr = 'data-officeview-presenter-sr'

/** Decoded pictures in a slide's (sparse, filled-as-decoded) image list. */
const countReady = (images: ReadonlyArray<unknown>) => { let n = 0; for (const i of images) if (i) n++; return n }

/** The preview's held slides: each with the request that holds it (PR #19 review 4). */
export type PrimeHolds = Map<number, object>

/**
 * Brings the preview's holds to `want`: slides no longer wanted are released, newly wanted ones primed. A prime
 * that fails is forgotten, so the next change retries it. Each request carries a token, so a stale callback (a
 * slide left and primed again before its first prime settled) never touches a newer request's hold. Exported for tests.
 */
export function syncPrimeHolds(held: PrimeHolds, want: ReadonlySet<number>, io: {
  release(i: number): void
  prime(i: number): Promise<boolean>
  primed(): void
  forgotten(): void
}): void {
  for (const i of [...held.keys()]) {
    if (want.has(i)) continue
    io.release(i)
    held.delete(i)
  }
  for (const i of want) {
    if (held.has(i)) continue
    const token = {}
    held.set(i, token)
    // Callbacks are scoped to this request: once the slide is released (or held by a newer request), they do nothing.
    const mine = () => held.get(i) === token
    const forget = () => { if (!mine()) return; held.delete(i); io.forgotten() }
    io.prime(i).then((ok) => { if (!mine()) return; if (ok) io.primed(); else forget() }, forget)
  }
}

export function PresenterView({ session, host, doc }: PresenterViewProps): ReactElement {
  const src = session.presenterSource
  useSyncExternalStore(src.subscribe, () => src.version)
  const st = session.state
  const count = src.deck.slides.length
  const slide = st.slide
  const model = src.model
  const ids = useId()
  const [surfaces, setSurfaces] = useState<PresenterSurfaces>()
  const [primed, setPrimed] = useState(0)
  const [gridOpen, setGridOpen] = useState(false)
  const [timerStatus, setTimerStatus] = useState('')
  const section = useRef<HTMLElement>(null)
  const allButton = useRef<HTMLButtonElement>(null)
  const mirrorSlot = useRef<HTMLDivElement>(null)
  const previewSlot = useRef<HTMLDivElement>(null)
  const hintSlot = useRef<HTMLDivElement>(null)
  const held = useRef<PrimeHolds>(new Map())
  const gridWasOpen = useRef(false)

  // The presenter's leased surfaces live with this view; closing it drops the preview's holds.
  useEffect(() => {
    const s = createPresenterSurfaces(doc, src.workspace, src.sink, () => { if (model) releasePreview(model) })
    setSurfaces(s)
    return () => { held.current.clear(); s.dispose() }
  }, [doc])

  // Focus moves into the view: the popup's own document, or the swapped surface (spec 9).
  useLayoutEffect(() => { section.current?.focus() }, [])

  // Keys on the popup's own document (the swapped view's keys reach the show host's listener).
  useEffect(() => {
    if (host !== 'window') return
    doc.addEventListener('keydown', src.key)
    if (hintSlot.current) src.placement?.hint(hintSlot.current)
    return () => doc.removeEventListener('keydown', src.key)
  }, [doc, host])

  // The panes take the leased canvases (a resize replaces them).
  useLayoutEffect(() => {
    if (!surfaces) return
    if (mirrorSlot.current && mirrorSlot.current.firstChild !== surfaces.mirror) mirrorSlot.current.replaceChildren(surfaces.mirror)
    if (previewSlot.current && previewSlot.current.firstChild !== surfaces.preview) previewSlot.current.replaceChildren(surfaces.preview)
  })

  // The panes' backing follows their CSS size and the window's DPR.
  const [epoch, setEpoch] = useState(0)
  useEffect(() => {
    if (!surfaces) return
    const win = doc.defaultView
    const fit = () => {
      const el = mirrorSlot.current
      const w = el?.clientWidth ?? 0, h = el?.clientHeight ?? 0
      if (w > 0 && h > 0 && surfaces.resize({ w, h }, win?.devicePixelRatio ?? 1)) setEpoch((n) => n + 1)
    }
    fit()
    win?.addEventListener('resize', fit)
    return () => win?.removeEventListener('resize', fit)
  }, [surfaces])

  const paintSlide = (canvas: HTMLCanvasElement, index: number, cursor?: SlideCursor, timeline?: SlideTimeline) => {
    const ctx = canvas.width > 0 ? canvas.getContext('2d') : null
    if (!ctx || !surfaces) return undefined
    const p = src.previewOf(index, cursor, timeline)
    const result = { images: p.images, ready: countReady(p.images) }
    try {
      paintPreview(ctx, { slide: p.slide, state: p.state, backing: { w: canvas.width, h: canvas.height } }, {
        images: p.images, resolveFont: p.resolveFont, builds: p.builds, scale: 1, scratch: surfaces.scratch, sink: src.sink,
        ...(src.watermark ? { watermark: src.watermark } : {}),
      })
    } catch (e) {
      src.sink.add({ code: 'presenter-paint', slideIndex: index, message: `Presenter View could not paint slide ${index + 1} (${e instanceof Error ? e.message : String(e)})` })
    }
    return result
  }

  // Current: the audience frame, copied after every drawn frame.
  useEffect(() => {
    if (!surfaces) return
    const paint = () => {
      const c = surfaces.mirror
      const ctx = c.width > 0 ? c.getContext('2d') : null
      if (!ctx) return
      const screen = session.state.screen
      if (screen !== 'slide' || session.state.slide === 'end') {
        ctx.fillStyle = screen === 'white' ? '#fff' : '#000'
        ctx.fillRect(0, 0, c.width, c.height)
        return
      }
      const source = src.audienceFrame()
      const dest = source ? mirrorFrame(source, ctx) : undefined
      if (dest) src.drawMedia?.(ctx, dest)
    }
    paint()
    return src.onFrame(paint)
  }, [surfaces, epoch])

  // Next: prime the shown and the next slide's timing; hand back what the preview left.
  const after = slide === 'end' ? undefined : nextVisible(st.order, slide) ?? (st.loop ? st.order[0] : undefined)
  // A refused prime is retried on the next change (the per-second tick included), once the workspace may have room.
  const [retry, setRetry] = useState(0)
  useEffect(() => {
    if (!model) return
    // At the end of the show nothing is wanted: every hold goes back.
    const want = new Set(slide === 'end' ? [] : [slide, ...(after === undefined ? [] : [after])])
    syncPrimeHolds(held.current, want, {
      release: (i) => releasePreview(model, i),
      prime: (i) => primePreview(model, i, { ...src.previewOptions, slideIndex: i }),
      primed: () => setPrimed((n) => n + 1),
      forgotten: () => setRetry((n) => n + 1),
    })
  }, [model, slide, after, retry > 0 && src.version])

  let pos: PreviewPosition | 'end' = 'end'
  if (slide !== 'end' && model) {
    try {
      pos = nextCheckpoint(model, st.order, slide, session.cursor ?? END_CURSOR, src.slideTime())
    } catch {
      // A cursor the preview timeline cannot replay (a text build's fanned history): the slide advance.
      pos = after === undefined ? 'end' : { slide: after, cursor: END_CURSOR }
    }
  }
  const posKey = pos === 'end' ? 'end' : `${pos.slide}:${pos.cursor.step}:${pos.cursor.mode}`
  /** What Next was last painted with: the stage decodes a neighbour's pictures after its metadata, so they can arrive late. */
  const painted = useRef<{ images: ReadonlyArray<unknown>; ready: number } | undefined>(undefined)
  const paintNext = () => {
    if (!surfaces || pos === 'end' || !model) return
    painted.current = paintSlide(surfaces.preview, pos.slide, pos.cursor, previewTimeline(model, pos.slide))
  }
  const paintNextRef = useRef(paintNext)
  paintNextRef.current = paintNext
  useLayoutEffect(paintNext, [posKey, surfaces, primed, epoch])
  useEffect(() => src.onFrame(() => {
    const p = painted.current
    if (p && countReady(p.images) !== p.ready) paintNextRef.current()
  }), [])

  // The grid gives focus back to the control that opened it.
  useEffect(() => {
    if (gridWasOpen.current && !gridOpen) allButton.current?.focus()
    gridWasOpen.current = gridOpen
  }, [gridOpen])

  const t = src.now()
  const timer = src.timer
  const status = slide === 'end' ? 'End of slide show' : `Slide ${slide + 1} of ${count}`
  const nextLabel = pos === 'end' ? 'Next: end of slide show' : `Next: slide ${pos.slide + 1}${pos.slide === slide ? ', next build' : ''}`
  const placement = host === 'window' ? src.placement : undefined
  const goField = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key !== 'Enter') return
    e.preventDefault()
    const n = Number(e.currentTarget.value.trim())
    e.currentTarget.value = ''
    if (Number.isInteger(n) && n >= 1 && n <= count) session.goTo(n - 1)
  }

  return (
    <section data-officeview-presenter-view="" data-host={host} aria-label="Presenter View" tabIndex={-1} ref={section}>
      <header data-officeview-presenter-bar="">
        <p data-officeview-presenter-status="" role="status" aria-live="polite">{status}</p>
        <div data-officeview-presenter-timer="" role="group" aria-label="Timer">
          <span id={`${ids}-elapsed`} {...{ [sr]: '' }}>Elapsed time</span>
          <span data-officeview-presenter-elapsed="" role="timer" aria-labelledby={`${ids}-elapsed`}>{formatElapsed(timer.elapsed(t))}</span>
          <button type="button" aria-label={timer.paused ? 'Resume timer' : 'Pause timer'} onClick={() => {
            if (timer.paused) { timer.resume(src.now()); setTimerStatus('Timer running') } else { timer.pause(src.now()); setTimerStatus('Timer paused') }
            src.touch()
          }}>{timer.paused ? 'Resume' : 'Pause'}</button>
          <button type="button" aria-label="Reset timer" onClick={() => { timer.reset(src.now()); setTimerStatus('Timer reset'); src.touch() }}>Reset</button>
          <span data-officeview-presenter-timer-status="" aria-live="polite" {...{ [sr]: '' }}>{timerStatus}</span>
        </div>
        <span id={`${ids}-clock`} {...{ [sr]: '' }}>Current time</span>
        <span data-officeview-presenter-clock="" aria-labelledby={`${ids}-clock`}>{formatWallClock(src.wall())}</span>
      </header>
      <div data-officeview-presenter-hint="" ref={hintSlot} />
      <div data-officeview-presenter-panes="">
        <figure data-officeview-presenter-current="" aria-label="Current slide">
          <div data-officeview-presenter-pane="" ref={mirrorSlot} />
          <figcaption>Current slide</figcaption>
        </figure>
        <figure data-officeview-presenter-next="" aria-label={nextLabel}>
          <div data-officeview-presenter-pane="" ref={previewSlot} hidden={pos === 'end'} />
          {pos === 'end' && <div data-officeview-presenter-end="">End of slide show</div>}
          <figcaption>{pos === 'end' ? 'Next: end of slide show' : pos.slide === slide ? 'Next build' : `Next slide (${pos.slide + 1})`}</figcaption>
        </figure>
        <NotesPane notes={src.notes()} selectable={src.protection.copy} />
      </div>
      <div data-officeview-presenter-controls="" role="toolbar" aria-label="Slide controls">
        <button type="button" aria-label="Previous slide" onClick={() => session.previous()}>Previous</button>
        <button type="button" aria-label="Next slide" onClick={() => session.next()}>Next</button>
        <button type="button" aria-label="Black screen" aria-pressed={st.screen === 'black'} onClick={() => src.input({ kind: 'black' })}>Black screen</button>
        <button type="button" ref={allButton} aria-label="See all slides" aria-expanded={gridOpen} onClick={() => setGridOpen((o) => !o)}>See all slides</button>
        <input type="text" inputMode="numeric" aria-label="Go to slide" placeholder="Slide #" size={5} onKeyDown={goField} />
        {placement?.canSwap && <button type="button" aria-label="Swap displays" aria-busy={placement.busy || undefined} onClick={() => placement.swap()}>Swap displays</button>}
        {placement?.canUseBoth && <button type="button" aria-label="Use both screens" aria-busy={placement.busy || undefined} onClick={() => placement.useBoth()}>Use both screens</button>}
        {host === 'swap' && <button type="button" aria-label="Show audience view" onClick={() => src.toggleView()}>Audience view</button>}
        <button type="button" aria-label="End show" onClick={() => session.exit()}>End show</button>
      </div>
      {gridOpen && surfaces && (
        <SlideGrid
          count={count}
          current={slide === 'end' ? -1 : slide}
          hidden={(i) => !!src.model?.index.slides[i]?.hidden}
          thumb={(i) => surfaces.grid(i)}
          release={(i) => surfaces.releaseGrid(i)}
          paint={(i, c) => { paintSlide(c, i) }}
          onJump={(i) => { setGridOpen(false); session.goTo(i) }}
          onClose={() => setGridOpen(false)}
        />
      )}
      <DiagnosticsPanel list={src.diagnostics()} />
    </section>
  )
}

/** Renders Presenter View into `container` (the popup's theme root, or the swap box) from the opener's realm. */
export function mountPresenterView(container: HTMLElement, props: PresenterViewProps): { dispose(): void } {
  const root = createRoot(container)
  flushSync(() => root.render(<PresenterView {...props} />))
  let live = true
  return {
    dispose() {
      if (!live) return
      live = false
      // A popup mid-pagehide may refuse the teardown; the session's fallback to single-screen still runs.
      try { root.unmount() } catch { /* the window is going away */ }
    },
  }
}
