/**
 * Stage 1d browser harness page: <OfficeFile data={{ kind: 'stream', url }}>
 * against server/test/app.ts. Query: ?doc=<corpus file>&deck=fly|morph|cube&backend=webgl2|canvas2d&relaxed=1&srv=short&fonts=1&copy=full|cap3|off&a11y=unit|focused|off&print=on|off|cap2|session3&role=presenter&notes=off&present=off.
 * deck= opens a synthetic show deck (piece 4); backend=canvas2d makes WebGL2 unavailable to the page, so the show
 * composites on Canvas 2D (webgl2, the default, leaves the browser as it is).
 * present=off sets allowPresent={false} (no Present button in the view; the harness buttons are then refused).
 * Assertions read window.__ovStream (notices, error, the show's diagnostics) and the DOM. The Present and
 * Presenter View buttons call the ref's present() (Stage 3c).
 */
import { useMemo, useRef } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeFile } from '../components/OfficeFile'
import type { OfficeViewerHandle } from '../components/present'
import type { StreamNotice } from '../stream/remote/session'
import '../styles/officeview.css'

const q = new URLSearchParams(location.search)
if (q.get('backend') === 'canvas2d') {
  // The forced Canvas 2D path, as a browser without WebGL2 has it: no 'webgl2' context, ever.
  const getContext = HTMLCanvasElement.prototype.getContext as (this: HTMLCanvasElement, id: string, o?: unknown) => RenderingContext | null
  HTMLCanvasElement.prototype.getContext = function (this: HTMLCanvasElement, id: string, o?: unknown) { return id === 'webgl2' ? null : getContext.call(this, id, o) } as typeof HTMLCanvasElement.prototype.getContext
}
const P = Number(q.get('port') ?? 4181)
const wsPort = q.get('srv') === 'short' ? P + 2 : P
const params = new URLSearchParams({ doc: q.get('doc') ?? 'pypptx-test.pptx', ...(q.get('relaxed') ? { relaxed: '1' } : {}), ...(q.get('copy') ? { copy: q.get('copy')! } : {}), ...(q.get('a11y') ? { a11y: q.get('a11y')! } : {}), ...(q.get('print') ? { print: q.get('print')! } : {}), ...(q.get('role') ? { role: q.get('role')! } : {}), ...(q.get('notes') ? { notes: q.get('notes')! } : {}), ...(q.get('deck') ? { deck: q.get('deck')! } : {}) })
const state = { notices: [] as StreamNotice[], error: undefined as string | undefined, present: undefined as string | undefined, endedAt: undefined as number | undefined, showDiagnostics: [] as Array<{ code: string; message: string }> }
;(window as unknown as { __ovStream: typeof state }).__ovStream = state

function App() {
  const source = useMemo(() => ({ kind: 'stream' as const, url: `ws://127.0.0.1:${wsPort}/stream?${params}` }), [])
  const ref = useRef<OfficeViewerHandle>(null)
  // Called inside the click (the gesture rule); the outcome lands in __ovStream.present.
  const present = (presenterView?: boolean) => {
    state.present = 'starting'
    ref.current?.present(presenterView ? { presenterView } : {}).then((s) => { state.present = 'started'; void s.ended.then(() => { state.present = 'ended'; state.endedAt = performance.now() }) }, (e: { code?: string }) => { state.present = e.code ?? 'error' })
  }
  return (
    <div style={{ height: '100vh' }}>
      {/* Fixed in a corner so the viewer's layout (and the earlier specs' geometry) is unchanged. */}
      <div style={{ position: 'fixed', right: 8, bottom: 8, zIndex: 10 }}>
        <button type="button" data-debug-present="" onClick={() => present()}>Present</button>
        <button type="button" data-debug-presenter="" onClick={() => present(true)}>Presenter View</button>
      </div>
      <OfficeFile
        ref={ref}
        data={source}
        style={{ height: '100%' }}
        watermark={{ text: 'HOSTMARK' }}
        {...(q.get('present') === 'off' ? { allowPresent: false } : {})}
        {...(q.get('fonts') ? { streamFonts: { baseUrl: `http://127.0.0.1:${P + 1}/fontpack` } } : {})}
        onStreamNotice={(n) => state.notices.push(n)}
        onSlideShowDiagnostic={(d) => state.showDiagnostics.push({ code: d.code, message: d.message })}
        error={(m) => { state.error = m; return <div role="alert" data-stream-error={m}>{m}</div> }}
      />
    </div>
  )
}
createRoot(document.getElementById('root')!).render(<App />)
