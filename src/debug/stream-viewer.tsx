/**
 * Stage 1d browser harness page: <OfficeFile data={{ kind: 'stream', url }}>
 * against server/test/app.ts. Query: ?doc=<corpus file>&relaxed=1&srv=short&fonts=1&copy=full|cap3|off&a11y=unit|focused|off&print=on|off|cap2|session3&role=presenter&notes=off&present=off.
 * present=off sets allowPresent={false} (no Present button in the view; the harness buttons are then refused).
 * Assertions read window.__ovStream (notices, error) and the DOM. The Present and
 * Presenter View buttons call the ref's present() (Stage 3c).
 */
import { useMemo, useRef } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeFile } from '../components/OfficeFile'
import type { OfficeViewerHandle } from '../components/present'
import type { StreamNotice } from '../stream/remote/session'
import '../styles/officeview.css'

const q = new URLSearchParams(location.search)
const P = Number(q.get('port') ?? 4181)
const wsPort = q.get('srv') === 'short' ? P + 2 : P
const params = new URLSearchParams({ doc: q.get('doc') ?? 'pypptx-test.pptx', ...(q.get('relaxed') ? { relaxed: '1' } : {}), ...(q.get('copy') ? { copy: q.get('copy')! } : {}), ...(q.get('a11y') ? { a11y: q.get('a11y')! } : {}), ...(q.get('print') ? { print: q.get('print')! } : {}), ...(q.get('role') ? { role: q.get('role')! } : {}), ...(q.get('notes') ? { notes: q.get('notes')! } : {}) })
const state = { notices: [] as StreamNotice[], error: undefined as string | undefined, present: undefined as string | undefined }
;(window as unknown as { __ovStream: typeof state }).__ovStream = state

function App() {
  const source = useMemo(() => ({ kind: 'stream' as const, url: `ws://127.0.0.1:${wsPort}/stream?${params}` }), [])
  const ref = useRef<OfficeViewerHandle>(null)
  // Called inside the click (the gesture rule); the outcome lands in __ovStream.present.
  const present = (presenterView?: boolean) => {
    state.present = 'starting'
    ref.current?.present(presenterView ? { presenterView } : {}).then((s) => { state.present = 'started'; void s.ended.then(() => { state.present = 'ended' }) }, (e: { code?: string }) => { state.present = e.code ?? 'error' })
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
        error={(m) => { state.error = m; return <div role="alert" data-stream-error={m}>{m}</div> }}
      />
    </div>
  )
}
createRoot(document.getElementById('root')!).render(<App />)
