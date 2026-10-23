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
 * the log as `a:change:<slide>:<build>` S1-20 (slideshow-play.spec.ts):
 * deck=play (two click builds, a trigger, a slide-jump link, an animated
 * link, a push transition), force=canvas2d (the forced Canvas 2D path),
 * budget=tight (a workspace that holds the permanent reserve and the
 * metadata but no layers, so the ladder reaches the terminal rung) and
 * motion=reduced|full (`slideShowMotion`). The force and budget seams are
 * set through a dynamic import of the session module, so other decks never
 * load the chunk early. S1-22 (slideshow-corpus.spec.ts, slideshow-perf.spec.ts):
 * deck=corpus&name=<deck> loads a corpus deck from /__corpus/lo-<deck>.pptx,
 * which the spec serves with page.route (the decks are never committed and
 * the locked-down server cannot read corpus/); bench=1 wraps the frame
 * driver to record each frame's work (`performance.now()` around the
 * session's tick, plus a 1-px readback of the WebGL2 canvas with sync=1, so
 * the GPU's work is included) and the rAF timestamps, in
 * `window.__slideshowStart.bench`; budget=wide gives the show a 1 GiB
 * workspace (profiling a transition at a size the 512 MiB budget refuses).
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

/** S1-22 frame-work samples (bench=1). */
export interface SlideshowBench {
  /** Per frame: ms of work in the session's tick (with sync=1, through a 1-px GPU readback). */
  work: number[]
  /** Per frame: the rAF timestamp. */
  stamps: number[]
  /** Clears both lists. */
  reset(): void
  /** UNMASKED_RENDERER_WEBGL of a probe WebGL2 context (which GPU or software rasteriser ran). */
  renderer: string
  /** The show workspace's bytes in use and its limit, now. */
  workspace(): { used: number; limit: number } | undefined
}

export interface SlideshowStart {
  ready: boolean
  bench?: SlideshowBench
  log: string[]
  /** Every show diagnostic, in full (the log keeps only codes). */
  diagnostics: Array<{ viewer: string; code: string; slideIndex?: number; effect?: string; message: string }>
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
  // After the module body: boot reads the deck helpers declared below.
  void Promise.resolve().then(boot)
}

const px = (v: number) => String(v * 9525)
const box = (id: number, x: number, y: number, w: number, h: number, text: string) => ({
  id, name: text, off: [px(x), px(y)] as [string, string], ext: [px(w), px(h)] as [string, string], fill: '336699',
  paragraphs: [{ runs: [{ text, sz: '2000' }] }],
})
const effect = (id: number, spid: number, extra = '') =>
  `<p:par><p:cTn id="${id}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${id + 1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
  `${extra}</p:childTnLst></p:cTn></p:par>`
const fadeIn = (id: number, spid: number) => `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${id}" dur="300"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>`
const step = (id: number, effects: string) =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
const trigger = (id: number, spid: number, effects: string) =>
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${id}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>` +
  `<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`

/** deck=play: slide 1 builds "First build" (with "Animated link") then "Second build" (with a fade); "Press" reveals "Triggered"; "Jump" goes to slide 3; "Animated link" to the last slide. Slide 2 pushes in. */
async function playDeck(): Promise<PptxDocument> {
  const main = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="90" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${step(10, effect(12, 2) + effect(14, 7))}${step(20, effect(22, 3, fadeIn(24, 3)))}</p:childTnLst></p:cTn></p:seq>`
  const timingXml = `<p:timing><p:tnLst><p:par><p:cTn id="99" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${main}${trigger(30, 4, effect(32, 5))}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
  const slides = [
    {
      shapes: [box(2, 80, 80, 400, 80, 'First build'), box(3, 80, 200, 400, 80, 'Second build'), box(4, 600, 80, 200, 80, 'Press'),
        box(5, 600, 200, 200, 80, 'Triggered'), box(6, 600, 400, 200, 80, 'Jump'), box(7, 80, 400, 400, 80, 'Animated link')],
      timingXml,
      links: [{ shapeId: 6, target: 'slide3.xml', action: 'ppaction://hlinksldjump' }, { shapeId: 7, action: 'ppaction://hlinkshowjump?jump=lastslide' }],
    },
    { shapes: [box(2, 80, 80, 400, 80, 'Pushed in')], transition: { xml: '<p:transition spd="slow"><p:push dir="u"/></p:transition>' } },
    { shapes: [box(2, 80, 80, 400, 80, 'Jump target')] },
    { shapes: [box(2, 80, 80, 400, 80, 'Last')] },
  ]
  return parsePptx(await OfficePackage.load(await buildPptxDeck({ slides })))
}

async function deck(text: string): Promise<PptxDocument> {
  const kind = params.get('deck')
  if (kind === 'play') return playDeck()
  if (kind === 'corpus') {
    const name = params.get('name') ?? ''
    if (!/^[\w.-]+$/.test(name)) throw new Error(`deck=corpus needs a deck name (got ${JSON.stringify(name)})`)
    const res = await fetch(`/__corpus/lo-${name}.pptx`)
    if (!res.ok) throw new Error(`corpus deck ${name}: HTTP ${res.status}`)
    return parsePptx(await OfficePackage.load(new Uint8Array(await res.arrayBuffer())))
  }
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

function rendererName(): string {
  const gl = document.createElement('canvas').getContext('webgl2')
  if (!gl) return 'no webgl2'
  const info = gl.getExtension('WEBGL_debug_renderer_info')
  const name = String(gl.getParameter(info ? info.UNMASKED_RENDERER_WEBGL : gl.RENDERER))
  gl.getExtension('WEBGL_lose_context')?.loseContext()
  return name
}

async function boot(): Promise<void> {
  if (params.has('force') || params.has('budget')) {
    const [{ sessionDeps }, { createSlideShowWorkspace }] = await Promise.all([import('../slideshow/player/session'), import('../slideshow/workspace')])
    if (params.get('force') === 'canvas2d') sessionDeps.backend = 'canvas2d'
    // budget=wide (S1-22 profiling only): a 1 GiB workspace, so a 3840 × 2160 transition is not refused.
    if (params.get('budget') === 'wide') sessionDeps.workspace = (profile) => createSlideShowWorkspace(profile, { limit: 2 ** 30 })
    if (params.get('budget') === 'tight') {
      sessionDeps.workspace = (profile) => {
        const out = Math.ceil(innerWidth * Math.min(2, devicePixelRatio)) * Math.ceil(innerHeight * Math.min(2, devicePixelRatio)) * 4
        return createSlideShowWorkspace(profile, { limit: out + 16 * 2 ** 20 + 2 * 2 ** 20, scratch: 16 * 2 ** 20 })
      }
    }
  }
  let bench: SlideshowBench | undefined
  if (params.get('bench') === '1') {
    const { sessionDeps } = await import('../slideshow/player/session')
    const sync = params.get('sync') === '1'
    let ws: { used(): number; limit: number } | undefined
    const b: SlideshowBench = {
      work: [], stamps: [], reset() { b.work.length = 0; b.stamps.length = 0 }, renderer: rendererName(),
      workspace: () => ws && { used: ws.used(), limit: ws.limit },
    }
    const makeWorkspace = sessionDeps.workspace
    sessionDeps.workspace = (profile) => (ws = makeWorkspace(profile))
    const pixel = new Uint8Array(4)
    // The production driver, with the tick timed: the work a frame does before the browser composites it.
    sessionDeps.frames = (win, frame) => {
      if (!win) return () => {}
      let on = true
      let id = 0
      const loop = (stamp: number) => {
        if (!on) return
        const t0 = win.performance.now()
        frame()
        if (sync) {
          const c = win.document.querySelector<HTMLCanvasElement>('canvas[data-slideshow-backend="webgl2"]')
          const gl = c?.getContext('webgl2')
          gl?.readPixels(0, 0, 1, 1, gl.RGBA, gl.UNSIGNED_BYTE, pixel)
        }
        b.work.push(win.performance.now() - t0)
        b.stamps.push(stamp)
        id = win.requestAnimationFrame(loop)
      }
      id = win.requestAnimationFrame(loop)
      return () => { on = false; win.cancelAnimationFrame(id) }
    }
    bench = b
  }
  const log: string[] = []
  const diagnostics: SlideshowStart['diagnostics'] = []
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
              slideShowMotion={(params.get('motion') as 'reduced' | 'full' | null) ?? undefined}
              onSlideShowStart={() => log.push(`${n}:onStart`)}
              onSlideShowEnd={() => log.push(`${n}:onEnd`)}
              onSlideShowDiagnostic={(d) => { log.push(`${n}:diag:${d.code}`); diagnostics.push({ viewer: n, ...d }) }}
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
    ...(bench ? { bench } : {}),
    log,
    diagnostics,
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
