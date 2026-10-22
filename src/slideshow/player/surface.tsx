/**
 * The audience surface (slideshow spec 4.1 end-of-show, 4.3 black/white
 * screen, 9 live region; checkpoint 1: static slides and cut transitions).
 *
 * Plain DOM built through the host's own document (realm-safe: Stage 2 moves
 * surfaces between windows), updated synchronously, so the session's state
 * and the page never disagree. The host is a child of the viewer's theme
 * root, so the scoped `[data-ov-theme-root] [data-officeview-show-*]` rules
 * in officeview.css apply.
 *
 * The static painter is the terminal path (spec 7, ladder step 3): the
 * scrolling viewer's static slide painted straight into one output canvas
 * whose bytes the permanent reserve already holds, with images admitted
 * against the reserve's scratch allowance. S1-20 replaces it with the
 * engine (`player/stage.ts`) and keeps it as the terminal fallback.
 */
import { acquireFonts, type FontLease } from '../../core/fonts/register'
import { decodeImageAsset, releaseDecodedImage } from '../../core/images'
import type { PptxDocument } from '../../pptx/types'
import { paintDirect } from '../render/direct'
import { imageIndices, slideSize, underlayOf } from '../render/scene'
import type { DiagnosticSink, TerminalReserve } from '../types'

export interface SurfaceView {
  /** The deck index shown, or the end-of-show slide. */
  slide: number | 'end'
  count: number
  screen: 'slide' | 'black' | 'white'
}

export interface AudienceSurface {
  readonly root: HTMLElement
  readonly canvas: HTMLCanvasElement
  show(view: SurfaceView): void
  dispose(): void
}

export const END_OF_SHOW = 'End of slide show, click to exit.'

/** Builds the surface inside `host` (appended; the host's other children, such as the full-screen notice, are left alone). */
export function createAudienceSurface(host: HTMLElement, backing: { w: number; h: number }, onClick: (e: MouseEvent) => void): AudienceSurface {
  const doc = host.ownerDocument
  const root = doc.createElement('div')
  root.setAttribute('data-officeview-show-surface', '')
  const canvas = doc.createElement('canvas')
  canvas.setAttribute('data-officeview-show-canvas', '')
  canvas.setAttribute('aria-hidden', 'true')
  canvas.width = backing.w
  canvas.height = backing.h
  const blank = doc.createElement('div')
  blank.hidden = true
  const end = doc.createElement('div')
  end.setAttribute('data-officeview-show-end', '')
  end.textContent = END_OF_SHOW
  end.hidden = true
  const live = doc.createElement('div')
  live.setAttribute('data-officeview-show-live', '')
  live.setAttribute('aria-live', 'polite')
  root.append(canvas, blank, end, live)
  root.addEventListener('click', onClick)
  host.append(root)
  let announced = ''
  return {
    root,
    canvas,
    show(view) {
      root.setAttribute('data-slide', String(view.slide))
      root.setAttribute('data-screen', view.screen)
      const isEnd = view.slide === 'end'
      canvas.hidden = isEnd
      end.hidden = !isEnd
      blank.hidden = view.screen === 'slide'
      if (view.screen === 'slide') blank.removeAttribute('data-officeview-show-blank')
      else blank.setAttribute('data-officeview-show-blank', view.screen)
      // Announced once per change (spec 9); titles come with S1-20's text list.
      const text = isEnd ? END_OF_SHOW : `Slide ${(view.slide as number) + 1} of ${view.count}`
      if (text !== announced) live.textContent = announced = text
    },
    dispose() {
      root.removeEventListener('click', onClick)
      root.remove()
    },
  }
}

export interface SlidePainter {
  /** Paints the deck slide `index` into `canvas`, letterboxed, as the scrolling viewer paints it. Resolves when painted or abandoned. */
  paint(canvas: HTMLCanvasElement, index: number): Promise<void>
  dispose(): void
}

export interface PainterContext {
  doc: PptxDocument
  terminal: TerminalReserve
  sink: DiagnosticSink
}

/**
 * The static painter. Fonts are acquired once per show (the viewer's own
 * registration is shared and reference-counted). Each paint decodes only the
 * images its slide uses and closes them right after painting: a static slide
 * is painted once per visit. A paint superseded by a later one is abandoned.
 */
export function createStaticPainter({ doc, terminal, sink }: PainterContext): SlidePainter {
  let fonts: Promise<FontLease | undefined> | undefined
  let generation = 0
  let disposed = false
  const lease = () => fonts ??= acquireFonts(doc).catch(() => undefined)
  return {
    async paint(canvas, index) {
      const mine = ++generation
      const slide = doc.slides[index]
      const used = [...new Set(imageIndices([...underlayOf(slide), ...(slide.shapes ?? [])]))]
      const images: Array<CanvasImageSource | undefined> = []
      try {
        const font = await lease()
        await Promise.all(used.map(async (i) => {
          const asset = doc.images?.[i]
          images[i] = asset ? await decodeImageAsset(asset).catch(() => undefined) : undefined
        }))
        if (disposed || mine !== generation) return
        const ctx = canvas.getContext('2d')
        if (!ctx) return
        const { w, h } = slideSize(slide)
        const scale = Math.min(canvas.width / w, canvas.height / h)
        ctx.setTransform(1, 0, 0, 1, 0, 0)
        ctx.fillStyle = '#000'
        ctx.fillRect(0, 0, canvas.width, canvas.height)
        ctx.translate((canvas.width - w * scale) / 2, (canvas.height - h * scale) / 2)
        paintDirect(ctx, slide, new Map(), {
          images, scale, scratch: terminal.scratch, sink,
          resolveFont: font?.resolve ?? ((family: string) => family),
        })
        ctx.setTransform(1, 0, 0, 1, 0, 0)
      } catch (e) {
        sink.add({ code: 'paint-failed', slideIndex: index, message: `The slide could not be painted: ${e instanceof Error ? e.message : String(e)}` })
      } finally {
        for (const img of new Set(images)) releaseDecodedImage(img)
      }
    },
    dispose() {
      disposed = true
      generation++
      void fonts?.then((l) => l?.dispose())
    },
  }
}
