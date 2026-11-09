/**
 * The stream source (spec §6.1). React-free and engine-free: imported by the
 * `officeview/core` entry and by the byte-source readers, which refuse it.
 */

export interface StreamSource {
  kind: 'stream'
  /** A WebSocket endpoint run by the developer's backend. `wss://` expected;
   * the origin must be in the deployment's allowlist (§9.3). */
  url: string
  /** Optional preauthorization proof for the open decision (§6.2). Never put
   * in a URL path; validated at HELLO and then discarded — it is NOT the
   * session token. */
  token?: string
}

/** True for a `{ kind: 'stream', url: string }` object. */
export function isStreamSource(v: unknown): v is StreamSource {
  return typeof v === 'object' && v !== null
    && (v as { kind?: unknown }).kind === 'stream'
    && typeof (v as { url?: unknown }).url === 'string'
}

/** Thrown when a StreamSource reaches a byte reader: a `wss://` url must never
 * reach `fetch` (§6.1). */
export class StreamSourceMisuseError extends Error {
  override readonly name = 'StreamSourceMisuseError'
  constructor() {
    super('A StreamSource cannot be read as bytes; pass it to <OfficeFile data> to open a stream session.')
  }
}
