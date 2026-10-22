/**
 * Dev-only harness for the slideshow start sequence (slideshow plan S1-17).
 *
 * Mounts the shipped <OfficeDoc> with a ref, as a host would, on a synthetic
 * one-slide deck ("Office-shaped synthetic", from the test builders). The
 * Present buttons call `ref.present()` synchronously inside their click
 * handler; `window.__slideshowStart.present()` calls it with no gesture.
 * Outcomes and viewer events go to `log`. The slideshow chunk is never
 * imported here: it loads through the viewer's own dynamic import, so the
 * spec can delay or fail it with page.route.
 *
 * Query: viewers=2 (two viewers, buttons #present-a, #present-b and
 * #present-both; the key "b", caught in the capture phase, presents viewer B
 * while A's show covers the page); frame=cross (this page inside a
 * cross-origin iframe without allow="fullscreen"); deck=nav (five slides,
 * the third hidden) or deck=loop (three slides, the second hidden, looping)
 * for the S1-19 navigation spec, which reads `onSlideShowSlideChange` from
 * the log as `a:change:<slide>:<build>`.
 *
 * No-gesture calls: Playwright's evaluate runs with a user gesture, and an
 * activation reaches ancestors and same-origin descendants only. So the
 * spec asks the cross-origin inner frame to present through postMessage
 * (which carries no activation); the result comes back the same way and is
 * kept in the outer page's `window.__slideshowFrame`.
 * Excluded from the library build (tsconfig.build.json excludes src/debug).
 */
import { createRef, useState, type JSX } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeDoc } from '../components/OfficeDoc'
import type { OfficeViewerHandle, PresentOptions, SlideShowSession } from '../components/present'
import { OfficePackage } from '../core/zip'
import { parsePptx } from '../pptx/parse'
import type { PptxDocument } from '../pptx/types'
import { buildPptx, buildPptxDeck } from '../testdata/ooxml-builders'
import '../styles/officeview.css'

export interface SlideshowStart {
  ready: boolean
  log: string[]
  /** The last session per viewer. */
  sessions: Record<string, SlideShowSession | undefined>
  /** present() with no user gesture. */
  present(which?: 'a' | 'b', opts?: PresentOptions): Promise<string>
  unmount(): void
  changeDocument(): void
  fullscreenElement(): string | null
}
export interface SlideshowFrame { ready: boolean; results: Array<{ result: string; active: boolean | undefined }> }
declare global { interface Window { __slideshowStart?: SlideshowStart; __slideshowFrame?: SlideshowFrame } }

const params = new URLSearchParams(location.search)
if (params.get('frame') === 'cross') {
  // Same page on the other loopback name: a different origin, and no allow attribute.
  const inner = new URL(location.href)
  inner.hostname = inner.hostname === 'localhost' ? '127.0.0.1' : 'localhost'
  inner.searchParams.delete('frame')
  const frame = document.createElement('iframe')
  frame.src = inner.href
  frame.title = 'viewer frame'
  frame.style.cssText = 'width: 100%; height: 100%; border: 0'
  document.getElementById('root')!.append(frame)
  const state: SlideshowFrame = { ready: false, results: [] }
  window.__slideshowFrame = state
  window.addEventListener('message', (e) => {
    if (e.source !== frame.contentWindow) return
    if (e.data?.ready) state.ready = true
    if (typeof e.data?.result === 'string') state.results.push(e.data)
  })
} else {
  void boot()
}

async function deck(text: string): Promise<PptxDocument> {
  const kind = params.get('deck')
  if (kind === 'nav' || kind === 'loop') {
    const hidden = kind === 'nav' ? 2 : 1
    const slides = Array.from({ length: kind === 'nav' ? 5 : 3 }, (_, i) => ({
      hidden: i === hidden,
      shapes: [{ off: ['914400', '914400'] as [string, string], ext: ['7315200', '1828800'] as [string, string], fill: '336699', paragraphs: [{ runs: [{ text: `${text} ${i + 1}`, sz: '4400' }] }] }],
    }))
    const presProps = kind === 'loop' ? '<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:showPr loop="1"><p:present/></p:showPr></p:presentationPr>' : undefined
    return parsePptx(await OfficePackage.load(await buildPptxDeck({ slides, presProps })))
  }
  const bytes = await buildPptx([{ off: ['914400', '914400'], ext: ['7315200', '1828800'], fill: '336699', paragraphs: [{ runs: [{ text, sz: '4400' }] }] }])
  return parsePptx(await OfficePackage.load(bytes))
}

async function boot(): Promise<void> {
  const log: string[] = []
  const refs = { a: createRef<OfficeViewerHandle>(), b: createRef<OfficeViewerHandle>() }
  const sessions: SlideshowStart['sessions'] = {}
  const docs = [await deck('First deck'), await deck('Second deck')]
  const two = params.get('viewers') === '2'
  const names = two ? (['a', 'b'] as const) : (['a'] as const)

  const start = (which: 'a' | 'b', opts?: PresentOptions): Promise<string> =>
    refs[which].current!.present(opts).then(
      (s) => { sessions[which] = s; log.push(`${which}:started`); void s.ended.then(() => log.push(`${which}:ended`)); return 'started' },
      (e: { code?: string }) => { log.push(`${which}:rejected:${e.code}`); return `rejected:${e.code}` },
    )

  let setDoc: (i: number) => void = () => {}
  let setMounted: (on: boolean) => void = () => {}
  function App(): JSX.Element | null {
    const [doc, set] = useState(0)
    const [mounted, mount] = useState(true)
    setDoc = set
    setMounted = mount
    if (!mounted) return null
    return (
      <>
        <div>
          {names.map((n) => <button key={n} id={`present-${n}`} type="button" onClick={() => void start(n)}>Present {n}</button>)}
          {two && <button id="present-both" type="button" onClick={() => { void start('a'); void start('b') }}>Present both</button>}
        </div>
        {names.map((n) => (
          <div key={n} className="viewer" data-viewer={n}>
            <OfficeDoc
              ref={refs[n]}
              document={docs[doc]}
              showToolbar={false}
              onSlideShowStart={() => log.push(`${n}:onStart`)}
              onSlideShowEnd={() => log.push(`${n}:onEnd`)}
              onSlideShowDiagnostic={(d) => log.push(`${n}:diag:${d.code}`)}
              onSlideShowSlideChange={params.has('deck') ? (i, b) => log.push(`${n}:change:${i}:${b}`) : undefined}
            />
          </div>
        ))}
      </>
    )
  }
  document.addEventListener('fullscreenchange', () => {
    const el = document.fullscreenElement
    log.push(`fullscreen:${el ? (el.closest('[data-viewer]') as HTMLElement | null)?.dataset.viewer ?? 'other' : 'none'}`)
  })
  if (two) window.addEventListener('keydown', (e) => { if (e.key === 'b') void start('b') }, true)
  if (parent !== window) {
    window.addEventListener('message', (e) => {
      if (e.source !== parent || e.data?.cmd !== 'present') return
      const active = (navigator as { userActivation?: { isActive: boolean } }).userActivation?.isActive
      void start('a').then((result) => parent.postMessage({ result, active }, '*'))
    })
  }
  createRoot(document.getElementById('root')!).render(<App />)
  window.__slideshowStart = {
    ready: true,
    log,
    sessions,
    present: (which = 'a', opts) => start(which, opts),
    unmount: () => setMounted(false),
    changeDocument: () => setDoc(1),
    fullscreenElement: () => {
      const el = document.fullscreenElement
      return el ? (el.closest('[data-viewer]') as HTMLElement | null)?.dataset.viewer ?? 'other' : null
    },
  }
  if (parent !== window) parent.postMessage({ ready: true }, '*')
}
