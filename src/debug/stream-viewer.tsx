/**
 * Stage 1d browser harness page: <OfficeFile data={{ kind: 'stream', url }}>
 * against server/test/app.ts. Query: ?doc=<corpus file>&deck=fly|morph|cube&backend=webgl2|canvas2d&relaxed=1&srv=short&fonts=1&copy=full|cap3|off&a11y=unit|focused|off&print=on|off|cap2|session3&role=presenter&notes=off&present=off.
 * deck= opens a synthetic show deck (piece 4); backend=canvas2d makes WebGL2 unavailable to the page, so the show
 * composites on Canvas 2D (webgl2, the default, leaves the browser as it is).
 * present=off sets allowPresent={false} (no Present button in the view; the harness buttons are then refused).
 * Assertions read window.__ovStream (notices, error, the show's diagnostics) and the DOM. The Present and
 * Presenter View buttons call the ref's present() (Stage 3c).
 * Stage 4a Task 14: `sheets=tiles` sets `streamSheets: 'tiles'` (the client HELLO picks the mode; the server-side
 * `?sheets=` alone does nothing), and `doc=` may name one of the harness's synthetic workbooks (million-rows,
 * grid500x100, frozen). A tiles-mode sheet's frame stats are exposed as window.__ovSheetStats: the last frame's
 * fields, plus `frames` (count) and `basePaintsTotal` (the running sum of `basePaints`).
 */
import { useMemo, useRef } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeFile } from '../components/OfficeFile'
import type { OfficeViewerHandle } from '../components/present'
import type { SheetFrameStats } from '../components/xlsx/InteractiveSheetDoc'
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
const state = { notices: [] as StreamNotice[], error: undefined as string | undefined, present: undefined as string | undefined, endedAt: undefined as number | undefined, showDiagnostics: [] as Array<{ code: string; message: string }>, mediaEnds: [] as Array<{ slide: number; nodeId: number; t: number; wall: number }>, mediaTrace: false }
;(window as unknown as { __ovStream: typeof state }).__ovStream = state
const sheetsTiles = q.get('sheets') === 'tiles'
type SheetStats = SheetFrameStats & { frames: number; basePaintsTotal: number }
let sheetStats: SheetStats | undefined
/** The viewer's @internal `onFrame` (forwarded by OfficeFile → OfficeDoc → RemoteDocView → RemoteSheetView). */
const onFrame = (stats: SheetFrameStats) => {
  sheetStats = { ...stats, frames: (sheetStats?.frames ?? 0) + 1, basePaintsTotal: (sheetStats?.basePaintsTotal ?? 0) + stats.basePaints }
  ;(window as unknown as { __ovSheetStats: SheetStats }).__ovSheetStats = sheetStats
}
// TEST-ONLY `?trace=media` (piece 5 Task 14): each computed media end the remote show's driver records (its node and
// show-clock time), with the wall time it was first seen. Loads the remote show module only when asked.
if (q.get('trace') === 'media') {
  void import('../stream/remote/show').then((m) => {
    m.remoteShowDeps.mediaEnds = (slide, ends) => {
      for (const e of ends) if (!state.mediaEnds.some((x) => x.slide === slide && x.nodeId === e.nodeId && x.t === e.t)) state.mediaEnds.push({ slide, ...e, wall: performance.now() })
    }
    state.mediaTrace = true
  })
}

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
        {...(sheetsTiles ? { streamSheets: 'tiles' as const, ...({ onFrame } as Record<string, unknown>) } : {})}
        {...(q.get('fonts') ? { streamFonts: { baseUrl: `http://127.0.0.1:${P + 1}/fontpack` } } : {})}
        onStreamNotice={(n) => state.notices.push(n)}
        onSlideShowDiagnostic={(d) => state.showDiagnostics.push({ code: d.code, message: d.message })}
        error={(m) => { state.error = m; return <div role="alert" data-stream-error={m}>{m}</div> }}
      />
    </div>
  )
}
createRoot(document.getElementById('root')!).render(<App />)
