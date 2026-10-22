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
 * load the chunk early.
 *
 * No-gesture calls: Playwright's evaluate runs with a user gesture, and an
 * activation reaches ancestors and same-origin descendants only. So the
 * spec asks the cross-origin inner frame to present through postMessage
 * (which carries no activation); the result comes back the same way and is
 * kept in the outer page's `window.__slideshowFrame`.
 *
 * S2-7 (slideshow-media.spec.ts): deck=media (slide 1 plays "Spanning", a
 * 3 s WAV with numSld=2, through a media call as it begins; click 1 plays
 * "Clip", a 1.2 s WAV whose media node begins with the click, and "After
 * clip" waits for its end; click 2 does the same for "Second" and "After
 * second"; slides 2 and 3 are plain) and deck=narration
 * (corpus/lo-narration.pptx, fetched; absent until bun scripts/fetch-corpus.ts).
 * Media elements report `media:playing:<spid>` and `media:ended:<spid>` to
 * the log through capture listeners.
 * S2-2 (slideshow-presenter.spec.ts): presenter=1 makes the Present buttons
 * ask for Presenter View; frame=sandbox embeds this page in a cross-origin
 * iframe with `sandbox="allow-scripts allow-same-origin"` and no
 * allow="fullscreen" (both refused); slideshow-csp.html runs this module
 * under a strict Content Security Policy.
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
if (params.get('frame') === 'sandbox') {
  // Cross-origin (the other loopback name) with no allow attribute, so full screen is refused; the sandbox has no allow-popups.
  const inner = new URL(location.href)
  inner.hostname = inner.hostname === 'localhost' ? '127.0.0.1' : 'localhost'
  inner.searchParams.delete('frame')
  const frame = document.createElement('iframe')
  frame.src = inner.href
  frame.title = 'viewer frame'
  frame.setAttribute('sandbox', 'allow-scripts allow-same-origin')
  frame.style.width = '100%'
  frame.style.height = '100%'
  frame.style.border = '0'
  document.getElementById('root')!.append(frame)
} else if (params.get('frame') === 'cross') {
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

/** A silent PCM WAV of `ms` milliseconds (8 kHz, mono, 16-bit): a format every Chromium decodes. */
function wav(ms: number): Uint8Array {
  const samples = Math.round(8 * ms)
  const out = new Uint8Array(44 + samples * 2)
  const v = new DataView(out.buffer)
  const str = (at: number, t: string) => { for (let i = 0; i < t.length; i++) out[at + i] = t.charCodeAt(i) }
  str(0, 'RIFF'); v.setUint32(4, 36 + samples * 2, true); str(8, 'WAVE'); str(12, 'fmt ')
  v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, 1, true); v.setUint32(24, 8000, true)
  v.setUint32(28, 16000, true); v.setUint16(32, 2, true); v.setUint16(34, 16, true); str(36, 'data'); v.setUint32(40, samples * 2, true)
  return out
}

const playCmd = (id: number, spid: number) =>
  `<p:par><p:cTn id="${id}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${id + 1}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`
/** A media node PowerPoint-style: it waits for its play command; `numSld` spans slides. */
const audioNode = (id: number, spid: number, numSld?: number, start = 'indefinite') =>
  `<p:audio><p:cMediaNode vol="80000"${numSld ? ` numSld="${numSld}"` : ''}><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="${start}"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:audio>`
/** A media node that begins with its click group, so its end (real or synthetic) is on the timeline. */
const mediaEffect = (id: number, node: string) =>
  `<p:par><p:cTn id="${id}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${node}</p:childTnLst></p:cTn></p:par>`
/** An entrance on `spid` that waits for node `tn` to end. */
const afterEnd = (id: number, spid: number, tn: number) =>
  `<p:par><p:cTn id="${id}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${id + 1}" dur="1" fill="hold"><p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="${tn}"/></p:cond></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>`

/** deck=media: see the header. Shapes 2 "Clip", 3 "After clip", 4 "Spanning", 5 "Second", 6 "After second". */
async function mediaDeck(): Promise<PptxDocument> {
  const preamble = `<p:par><p:cTn id="10" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="90"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="11" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${playCmd(12, 4).replace('clickEffect', 'afterEffect')}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
  const main = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="90" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${preamble}${step(20, mediaEffect(22, audioNode(80, 2, undefined, '0')) + afterEnd(24, 3, 80))}${step(30, mediaEffect(32, audioNode(82, 5, undefined, '0')) + afterEnd(34, 6, 82))}</p:childTnLst></p:cTn></p:seq>`
  const timingXml = `<p:timing><p:tnLst><p:par><p:cTn id="99" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${main}${audioNode(81, 4, 2)}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
  const slides = [
    {
      shapes: [box(2, 80, 80, 300, 80, 'Clip'), box(3, 80, 200, 300, 80, 'After clip'), box(4, 500, 80, 300, 80, 'Spanning'), box(5, 80, 320, 300, 80, 'Second'), box(6, 80, 440, 300, 80, 'After second')],
      timingXml,
      media: [
        { shapeId: 2, kind: 'audio' as const, ext: 'wav', data: wav(1200), p14: true },
        { shapeId: 4, kind: 'audio' as const, ext: 'wav', data: wav(3000), p14: true },
        { shapeId: 5, kind: 'audio' as const, ext: 'wav', data: wav(1000), p14: true },
      ],
    },
    { shapes: [box(2, 80, 80, 400, 80, 'Slide two')] },
    { shapes: [box(2, 80, 80, 400, 80, 'Slide three')] },
  ]
  return parsePptx(await OfficePackage.load(await buildPptxDeck({ slides })))
}

async function deck(text: string): Promise<PptxDocument> {
  const kind = params.get('deck')
  if (kind === 'play') return playDeck()
  if (kind === 'media') return mediaDeck()
  if (kind === 'narration') {
    const res = await fetch('/corpus/lo-narration.pptx')
    if (!res.ok) throw new Error('corpus/lo-narration.pptx is absent (bun scripts/fetch-corpus.ts)')
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

async function boot(): Promise<void> {
  if (params.has('force') || params.has('budget')) {
    const [{ sessionDeps }, { createSlideShowWorkspace }] = await Promise.all([import('../slideshow/player/session'), import('../slideshow/workspace')])
    if (params.get('force') === 'canvas2d') sessionDeps.backend = 'canvas2d'
    if (params.get('budget') === 'tight') {
      sessionDeps.workspace = (profile) => {
        const out = Math.ceil(innerWidth * Math.min(2, devicePixelRatio)) * Math.ceil(innerHeight * Math.min(2, devicePixelRatio)) * 4
        return createSlideShowWorkspace(profile, { limit: out + 16 * 2 ** 20 + 2 * 2 ** 20, scratch: 16 * 2 ** 20 })
      }
    }
  }
  const log: string[] = []
  const refs = { a: createRef<OfficeViewerHandle>(), b: createRef<OfficeViewerHandle>() }
  const sessions: SlideshowStart['sessions'] = {}
  const docs = [await deck('First deck'), await deck('Second deck')]
  const two = params.get('viewers') === '2'
  const presenter: PresentOptions | undefined = params.get('presenter') === '1' ? { presenterView: true } : undefined
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
          {names.map((n) => <button key={n} id={`present-${n}`} type="button" onClick={() => void start(n, presenter)}>Present {n}</button>)}
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
  // Media events do not bubble; capture listeners still see them.
  for (const type of ['playing', 'ended'] as const) {
    document.addEventListener(type, (e) => {
      const spid = (e.target as Element | null)?.getAttribute?.('data-officeview-show-media-spid')
      if (spid) log.push(`media:${type}:${spid}`)
    }, true)
  }
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
