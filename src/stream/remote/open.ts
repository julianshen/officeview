/** The lazy entry `useOfficeFile` imports for a `StreamSource` (§6.3). */
import type { CanvasBackend } from '../../core/canvas-backend'
import type { Manifest } from '../manifest'
import type { StreamSource } from '../source'
import { resolveFaces } from './faces'
import { createRemoteDoc, type RemoteOfficeDoc } from './remote-doc'
import { StreamSession, type StreamClientFactory, type StreamNotice } from './session'

export interface OpenRemoteOptions {
  client?: StreamClientFactory
  /** Where the host serves the font pack (`pack.json` + TTFs). */
  fonts?: { baseUrl: string }
  signal?: AbortSignal
  onNotice?: (n: StreamNotice) => void
  /** Tests: the paint backend and a face resolver override. */
  backend?: CanvasBackend
  faces?: (m: Manifest) => Promise<string[]>
}

export async function openRemoteDocument(source: StreamSource, o: OpenRemoteOptions = {}): Promise<RemoteOfficeDoc> {
  const session = await StreamSession.open(source, {
    ...(o.client ? { client: o.client } : {}),
    ...(o.signal ? { signal: o.signal } : {}),
    ...(o.onNotice ? { onNotice: o.onNotice } : {}),
    faces: o.faces ?? ((m) => resolveFaces(m.recorder.requiredFaces, o.fonts?.baseUrl)),
  })
  return createRemoteDoc(session, { ...(o.backend ? { backend: o.backend } : {}), ...(o.onNotice ? { onNotice: o.onNotice } : {}) })
}

export type { RemoteOfficeDoc } from './remote-doc'
export { StreamSessionError } from './errors'
