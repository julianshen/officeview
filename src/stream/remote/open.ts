/** The lazy entry `useOfficeFile` imports for a `StreamSource` (§6.3). */
import type { CanvasBackend } from '../../core/canvas-backend'
import type { Manifest } from '../manifest'
import type { StreamSource } from '../source'
import { SHEET_TILES_FEATURE, SHEET_VIEW_FEATURE } from '../sheet-wire'
import { resolveFaces } from './faces'
import { createRemoteDoc, type RemoteOfficeDoc } from './remote-doc'
import { StreamSession, type StreamClientFactory, type StreamNotice } from './session'

export interface OpenRemoteOptions {
  client?: StreamClientFactory | undefined
  /** Where the host serves the font pack (`pack.json` + TTFs). */
  fonts?: { baseUrl: string } | undefined
  signal?: AbortSignal
  onNotice?: (n: StreamNotice) => void
  /** Tests: the paint backend and a face resolver override. */
  backend?: CanvasBackend
  faces?: (m: Manifest) => Promise<string[]>
  /**
   * The HELLO `features` for an XLSX document (S4 §6.10; Stage 4b plan P1, F1). Unset (the default) sends
   * `['sheet-view@1']`: tiles with the 4b ops on a 4b server, units on a 4a server (which ignores it). `'tiles'` sends
   * exactly `['sheet-tiles@1']` (4a's opt-in and behaviour). `'units'` sends no `features` (Stage 3's HELLO).
   */
  streamSheets?: 'tiles' | 'units' | undefined
}

export async function openRemoteDocument(source: StreamSource, o: OpenRemoteOptions = {}): Promise<RemoteOfficeDoc> {
  const session = await StreamSession.open(source, {
    ...(o.client ? { client: o.client } : {}),
    ...(o.signal ? { signal: o.signal } : {}),
    ...(o.onNotice ? { onNotice: o.onNotice } : {}),
    ...(o.streamSheets === 'units' ? {} : { features: [o.streamSheets === 'tiles' ? SHEET_TILES_FEATURE : SHEET_VIEW_FEATURE] }),
    faces: o.faces ?? ((m) => resolveFaces(m.recorder.requiredFaces, o.fonts?.baseUrl)),
  })
  return createRemoteDoc(session, { ...(o.backend ? { backend: o.backend } : {}), ...(o.onNotice ? { onNotice: o.onNotice } : {}) })
}

export type { RemoteOfficeDoc } from './remote-doc'
export { StreamSessionError } from './errors'
