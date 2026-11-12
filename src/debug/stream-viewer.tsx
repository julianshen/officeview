/**
 * Stage 1d browser harness page: <OfficeFile data={{ kind: 'stream', url }}>
 * against server/test/app.ts. Query: ?doc=<corpus file>&relaxed=1&srv=short&fonts=1&copy=full|cap3|off&a11y=unit|focused|off.
 * Assertions read window.__ovStream (notices, error) and the DOM.
 */
import { useMemo } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeFile } from '../components/OfficeFile'
import type { StreamNotice } from '../stream/remote/session'
import '../styles/officeview.css'

const q = new URLSearchParams(location.search)
const P = Number(q.get('port') ?? 4181)
const wsPort = q.get('srv') === 'short' ? P + 2 : P
const params = new URLSearchParams({ doc: q.get('doc') ?? 'pypptx-test.pptx', ...(q.get('relaxed') ? { relaxed: '1' } : {}), ...(q.get('copy') ? { copy: q.get('copy')! } : {}), ...(q.get('a11y') ? { a11y: q.get('a11y')! } : {}) })
const state = { notices: [] as StreamNotice[], error: undefined as string | undefined }
;(window as unknown as { __ovStream: typeof state }).__ovStream = state

function App() {
  const source = useMemo(() => ({ kind: 'stream' as const, url: `ws://127.0.0.1:${wsPort}/stream?${params}` }), [])
  return (
    <div style={{ height: '100vh' }}>
      <OfficeFile
        data={source}
        style={{ height: '100%' }}
        watermark={{ text: 'HOSTMARK' }}
        {...(q.get('fonts') ? { streamFonts: { baseUrl: `http://127.0.0.1:${P + 1}/fontpack` } } : {})}
        onStreamNotice={(n) => state.notices.push(n)}
        error={(m) => { state.error = m; return <div role="alert" data-stream-error={m}>{m}</div> }}
      />
    </div>
  )
}
createRoot(document.getElementById('root')!).render(<App />)
