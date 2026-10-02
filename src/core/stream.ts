/**
 * Streaming ingestion for office files.
 *
 * A URL (with headers, e.g. Authorization), a Request, a fetch() promise, a
 * Response, a Blob, or any ReadableStream<Uint8Array> can feed the viewer.
 * Chunks accumulate into one buffer with progress callbacks, so a
 * 50 MB deck shows real download progress instead of a blank spinner.
 *
 * Honest limitation: an OOXML file is a zip whose central directory sits at
 * the END of the file, so rendering can only begin once the final byte has
 * arrived. Streaming here buys progress feedback and chunked memory
 * accumulation — not partial rendering (see README).
 */

export interface Progress {
  /** Bytes received so far. */
  loaded: number
  /** Total bytes when the source declares it (Content-Length), else undefined. */
  total?: number
}

type DirectByteSource =
  | ArrayBuffer
  | Uint8Array
  | ReadableStream<Uint8Array>
  | Blob
  | Response
  | Request
  | HttpSource

export type ByteSource = DirectByteSource | Promise<DirectByteSource>

/**
 * An http(s) URL to download with streaming progress, plus the headers that
 * protect it (e.g. `{ Authorization: 'Bearer …' }`). `credentials` forwards
 * cookies/TLS-client-cert behaviour to fetch (`'include'` for cross-site
 * cookies). Prefer a `Request` when you also need method/mode/cache control.
 */
export interface HttpSource {
  url: string
  headers?: HeadersInit
  credentials?: RequestCredentials
}

export type ProgressListener = (progress: Progress) => void

/** Response header carrying the document owner's protection policy. */
export const PROTECTION_HEADER = 'X-OfficeView-Protection'

/**
 * What a viewer may do with a document. Defaults to allow-all; a server
 * denies capabilities by sending `X-OfficeView-Protection: no-copy, no-print`
 * on the download response. Unknown tokens are ignored so the scheme can grow
 * without breaking old readers.
 */
export interface ProtectionPolicy {
  allowCopy: boolean
  allowPrint: boolean
}

/** Read the protection policy off download response headers (pure). */
export function protectionFromHeaders(headers: Headers): ProtectionPolicy {
  const raw = headers.get(PROTECTION_HEADER)
  if (!raw) return { allowCopy: true, allowPrint: true }
  const tokens = new Set(
    raw
      .split(',')
      .map((t) => t.trim().toLowerCase())
      .filter((t) => t.length > 0),
  )
  return { allowCopy: !tokens.has('no-copy'), allowPrint: !tokens.has('no-print') }
}

export type ProtectionListener = (policy: ProtectionPolicy) => void

/** Read a byte stream to completion, reporting progress as chunks arrive. */
export async function readByteStream(
  stream: ReadableStream<Uint8Array>,
  onProgress?: ProgressListener,
  total?: number,
): Promise<Uint8Array> {
  const reader = stream.getReader()
  const chunks: Uint8Array[] = []
  let loaded = 0
  let announced = -1
  const report = () => {
    // throttle to whole-percent steps when the total is known: progress
    // callbacks fire per chunk and big files produce thousands of them
    if (total && total > 0) {
      const pct = Math.floor((loaded / total) * 100)
      if (pct === announced) return
      announced = pct
    }
    onProgress?.({ loaded, total })
  }
  report()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    if (value && value.length > 0) {
      chunks.push(value)
      loaded += value.length
      report()
    }
  }
  const out = new Uint8Array(loaded)
  let at = 0
  for (const chunk of chunks) {
    out.set(chunk, at)
    at += chunk.length
  }
  onProgress?.({ loaded, total })
  return out
}

/** Blob → one buffer (Blob.arrayBuffer gives the whole thing; fine for files already in memory). */
async function readBlob(blob: Blob, onProgress?: ProgressListener): Promise<Uint8Array> {
  onProgress?.({ loaded: 0, total: blob.size })
  const buffer = await blob.arrayBuffer()
  const bytes = new Uint8Array(buffer)
  onProgress?.({ loaded: bytes.length, total: blob.size })
  return bytes
}

/** A fetch() Response → one buffer, with content-length progress. */
async function readResponse(
  response: Response,
  onProgress?: ProgressListener,
  onProtection?: ProtectionListener,
): Promise<Uint8Array> {
  // An auth failure (or any non-2xx) must surface as itself, not as a
  // confusing zip-parse error three layers down.
  if (!response.ok) {
    const status = `HTTP ${response.status}${response.statusText ? ` ${response.statusText}` : ''}`
    throw new Error(`${status} while downloading the document`)
  }
  // Server policy arrives before the first byte: protection engages while the
  // download is still in flight, never after the render.
  onProtection?.(protectionFromHeaders(response.headers))
  const totalHeader = response.headers.get('content-length')
  const total = totalHeader ? parseInt(totalHeader, 10) : undefined
  if (response.body) return readByteStream(response.body, onProgress, Number.isFinite(total) ? total : undefined)
  return new Uint8Array(await response.arrayBuffer())
}

/**
 * Download a URL (or re-fetch a Request) with streaming progress.
 * Headers/cookies travel in `init` (or inside the Request); they are never
 * copied into error messages.
 */
async function fetchInput(
  input: string | Request,
  init: RequestInit | undefined,
  onProgress?: ProgressListener,
  onProtection?: ProtectionListener,
): Promise<Uint8Array> {
  if (typeof fetch !== 'function') {
    const url = typeof input === 'string' ? input : input.url
    throw new Error(`Cannot download ${url}: fetch is not available in this environment`)
  }
  return readResponse(await fetch(input, init), onProgress, onProtection)
}

/** Normalize any accepted byte source into a single Uint8Array buffer. */
export async function readSource(
  source: ByteSource,
  onProgress?: ProgressListener,
  onProtection?: ProtectionListener,
): Promise<Uint8Array> {
  // A promise of a source (e.g. the promise fetch() returns): wait for it,
  // then handle what it resolves to. Progress callbacks only start once the
  // inner source exists — a bare fetch promise reports nothing while the
  // request is still in flight.
  if (typeof (source as Promise<DirectByteSource> | null)?.then === 'function') {
    return readSource(await (source as Promise<DirectByteSource>), onProgress, onProtection)
  }
  // A Request already bundles url + headers + credentials (+ method/mode),
  // so fetch it as-is. Realm note: a cross-realm Request fails instanceof
  // and falls through to the { url } branch below, which still forwards its
  // url, headers and credentials (but not method/body).
  if (typeof Request !== 'undefined' && source instanceof Request) {
    return fetchInput(source, undefined, onProgress, onProtection)
  }
  if (source instanceof Response) return readResponse(source, onProgress, onProtection)
  if (source instanceof Blob) return readBlob(source, onProgress)
  // A { url, headers? } descriptor. Checked after Response/Blob: a Response
  // also carries a (possibly empty) .url, and must keep its own branch.
  if (typeof source === 'object' && source !== null && typeof (source as HttpSource).url === 'string') {
    const { url, headers, credentials } = source as HttpSource
    return fetchInput(url, { headers, credentials }, onProgress, onProtection)
  }
  if (typeof ReadableStream !== 'undefined' && source instanceof ReadableStream) {
    return readByteStream(source as ReadableStream<Uint8Array>, onProgress)
  }
  // Realm-proof byte checks. `instanceof Uint8Array` fails when the source was
  // created in a different VM realm — under vitest's jsdom environment a Node
  // Buffer fails `instanceof` against the jsdom-context Uint8Array, which made
  // every real file "Unsupported byte source". ArrayBuffer.isView checks the
  // internal slots instead and works across realms.
  if (ArrayBuffer.isView(source)) {
    const view = source as Uint8Array
    return view.byteOffset === 0 && view.byteLength === (view.buffer as ArrayBuffer).byteLength
      ? view
      : new Uint8Array(view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength))
  }
  if (source instanceof ArrayBuffer) return new Uint8Array(source)
  throw new Error('Unsupported byte source')
}
