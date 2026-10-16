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

import { watermarkFromHeaders, type WatermarkOptions } from './watermark'
import { throwIfAborted } from './package-error'

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

/** Fires once with the watermark the response header asked for, if any. */
export type WatermarkListener = (mark: WatermarkOptions | undefined) => void

/** Read a byte stream to completion, reporting progress as chunks arrive.
 * An abort (even one that happened earlier) cancels the stream and rejects
 * with PackageError 'cancelled'. */
export async function readByteStream(
  stream: ReadableStream<Uint8Array>,
  onProgress?: ProgressListener,
  total?: number,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  const reader = stream.getReader()
  const stop = () => { reader.cancel().catch(() => {}) }
  if (signal?.aborted) stop()
  else signal?.addEventListener('abort', stop, { once: true })
  try {
    return await drain(reader, onProgress, total, signal)
  } finally {
    signal?.removeEventListener('abort', stop)
    if (signal?.aborted) try { reader.releaseLock() } catch { /* a read is still settling */ }
  }
}

async function drain(reader: ReadableStreamDefaultReader<Uint8Array>, onProgress: ProgressListener | undefined, total: number | undefined, signal: AbortSignal | undefined): Promise<Uint8Array> {
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
    throwIfAborted(signal)
    const { done, value } = await reader.read()
    throwIfAborted(signal)
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
async function readBlob(blob: Blob, onProgress?: ProgressListener, signal?: AbortSignal): Promise<Uint8Array> {
  onProgress?.({ loaded: 0, total: blob.size })
  const buffer = await blob.arrayBuffer()
  throwIfAborted(signal)
  const bytes = new Uint8Array(buffer)
  onProgress?.({ loaded: bytes.length, total: blob.size })
  return bytes
}

/** A fetch() Response → one buffer, with content-length progress. */
async function readResponse(
  response: Response,
  onProgress?: ProgressListener,
  onProtection?: ProtectionListener,
  onWatermark?: WatermarkListener,
  signal?: AbortSignal,
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
  // Same timing as protection: the watermark is known before the first byte, so
  // the page paints already stamped and never flashes unstamped.
  onWatermark?.(watermarkFromHeaders(response.headers))
  const totalHeader = response.headers.get('content-length')
  const total = totalHeader ? parseInt(totalHeader, 10) : undefined
  if (response.body) return readByteStream(response.body, onProgress, Number.isFinite(total) ? total : undefined, signal)
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
  onWatermark?: WatermarkListener,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  if (typeof fetch !== 'function') {
    const url = typeof input === 'string' ? input : input.url
    throw new Error(`Cannot download ${url}: fetch is not available in this environment`)
  }
  return readResponse(await fetch(input, withSignal(input, init, signal)), onProgress, onProtection, onWatermark, signal)
}

/** The caller's init plus a signal that aborts on either the load signal or the
 * caller's own (init.signal, or the Request's, which fetch(request, { signal }) replaces). */
function withSignal(input: string | Request, init: RequestInit | undefined, signal: AbortSignal | undefined): RequestInit | undefined {
  const own = init?.signal ?? (typeof input === 'string' ? undefined : input.signal)
  const combined = eitherSignal(own ?? undefined, signal)
  return combined ? { ...init, signal: combined } : init
}

/** Normalize any accepted byte source into a single Uint8Array buffer. `signal`
 * aborts the fetch and the download (PackageError 'cancelled' once aborted). */
export async function readSource(
  source: ByteSource,
  onProgress?: ProgressListener,
  onProtection?: ProtectionListener,
  onWatermark?: WatermarkListener,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  // A promise of a source (e.g. the promise fetch() returns): wait for it,
  // then handle what it resolves to. Progress callbacks only start once the
  // inner source exists — a bare fetch promise reports nothing while the
  // request is still in flight.
  if (typeof (source as Promise<DirectByteSource> | null)?.then === 'function') {
    const resolved = await (source as Promise<DirectByteSource>)
    return readSource(resolved, onProgress, onProtection, onWatermark, signal)
  }
  // A Request already bundles url + headers + credentials (+ method/mode),
  // so fetch it as-is. Realm note: a cross-realm Request fails instanceof
  // and falls through to the { url } branch below, which still forwards its
  // url, headers and credentials (but not method/body).
  if (typeof Request !== 'undefined' && source instanceof Request) {
    return fetchInput(source, undefined, onProgress, onProtection, onWatermark, signal)
  }
  if (source instanceof Response) return readResponse(source, onProgress, onProtection, onWatermark, signal)
  if (source instanceof Blob) return readBlob(source, onProgress, signal)
  // A { url, headers? } descriptor. Checked after Response/Blob: a Response
  // also carries a (possibly empty) .url, and must keep its own branch.
  if (typeof source === 'object' && source !== null && typeof (source as HttpSource).url === 'string') {
    const { url, headers, credentials } = source as HttpSource
    return fetchInput(url, { headers, credentials }, onProgress, onProtection, onWatermark, signal)
  }
  if (typeof ReadableStream !== 'undefined' && source instanceof ReadableStream) {
    return readByteStream(source as ReadableStream<Uint8Array>, onProgress, undefined, signal)
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

/** A sequential byte source opened for pulling: progress is reported exactly
 * as readByteStream reports it (an initial report, one per chunk, throttled
 * to whole percents when the total is known, and a final unconditional one). */
export interface ByteStreamReader {
  /** Declared length (Content-Length): a hint only. */
  readonly total: number | undefined
  /** Next non-empty chunk, or undefined at the end. Never reads ahead. */
  read(): Promise<Uint8Array | undefined>
  cancel(): Promise<void>
}

export type OpenedSource =
  | { kind: 'bytes'; bytes: Uint8Array }
  | { kind: 'blob'; blob: Blob }
  | { kind: 'stream'; reader: ByteStreamReader }

export interface OpenSourceListeners {
  onProgress?: ProgressListener
  onProtection?: ProtectionListener
  onWatermark?: WatermarkListener
}

function pullStream(stream: ReadableStream<Uint8Array>, onProgress: ProgressListener | undefined, total: number | undefined, signal?: AbortSignal): ByteStreamReader {
  const reader = stream.getReader()
  let loaded = 0
  let announced = -1
  let done = false
  const report = () => {
    if (total && total > 0) {
      const pct = Math.floor((loaded / total) * 100)
      if (pct === announced) return
      announced = pct
    }
    onProgress?.({ loaded, total })
  }
  const cancel = async () => {
    if (done) return
    done = true
    try { await reader.cancel() } catch { /* already errored */ }
    // cancel() settles pending reads; release the lock so the body is not left locked.
    try { reader.releaseLock() } catch { /* a read is still settling */ }
  }
  // 'abort' never fires again for a signal that is already aborted (a promised
  // Response that resolved after its load was replaced): cancel straight away.
  if (signal?.aborted) void cancel()
  else signal?.addEventListener('abort', () => { void cancel() }, { once: true })
  report()
  return {
    total,
    async read() {
      while (!done) {
        const { done: end, value } = await reader.read()
        if (end) {
          if (!done) {
            done = true
            onProgress?.({ loaded, total })
          }
          return undefined
        }
        if (value && value.length > 0) {
          loaded += value.length
          report()
          return value
        }
      }
      return undefined
    },
    cancel,
  }
}

async function openResponse(response: Response, listeners: OpenSourceListeners, signal?: AbortSignal): Promise<OpenedSource> {
  if (!response.ok) {
    const status = `HTTP ${response.status}${response.statusText ? ` ${response.statusText}` : ''}`
    throw new Error(`${status} while downloading the document`)
  }
  listeners.onProtection?.(protectionFromHeaders(response.headers))
  listeners.onWatermark?.(watermarkFromHeaders(response.headers))
  const totalHeader = response.headers.get('content-length')
  const total = totalHeader ? parseInt(totalHeader, 10) : undefined
  if (response.body) return { kind: 'stream', reader: pullStream(response.body, listeners.onProgress, Number.isFinite(total) ? total : undefined, signal) }
  return { kind: 'bytes', bytes: new Uint8Array(await response.arrayBuffer()) }
}

async function openFetch(input: string | Request, init: RequestInit | undefined, listeners: OpenSourceListeners, signal?: AbortSignal): Promise<OpenedSource> {
  if (typeof fetch !== 'function') {
    const url = typeof input === 'string' ? input : input.url
    throw new Error(`Cannot download ${url}: fetch is not available in this environment`)
  }
  return openResponse(await fetch(input, withSignal(input, init, signal)), listeners, signal)
}

/** Aborts when either signal does (AbortSignal.any where it exists and accepts both). */
function eitherSignal(a: AbortSignal | undefined, b: AbortSignal | undefined): AbortSignal | undefined {
  if (!a || a === b) return b
  if (!b) return a
  const any = (AbortSignal as { any?: (signals: AbortSignal[]) => AbortSignal }).any
  try { if (typeof any === 'function') return any.call(AbortSignal, [a, b]) } catch { /* cross-realm signal: fall back */ }
  const controller = new AbortController()
  for (const s of [a, b]) {
    if (s.aborted) { controller.abort(s.reason); break }
    s.addEventListener('abort', () => controller.abort(s.reason), { once: true })
  }
  return controller.signal
}

/**
 * The same source handling as readSource (promise, Request, Response, Blob,
 * `{ url, headers, credentials }`, ReadableStream, cross-realm byte views),
 * but a sequential source is returned as a pull reader instead of being
 * accumulated, so the caller can look at a prefix and choose to accumulate
 * or spool. Protection and watermark listeners fire before the first byte.
 */
export async function openByteStream(source: ByteSource, listeners: OpenSourceListeners = {}, signal?: AbortSignal): Promise<OpenedSource> {
  if (typeof (source as Promise<DirectByteSource> | null)?.then === 'function') {
    return openByteStream(await (source as Promise<DirectByteSource>), listeners, signal)
  }
  if (typeof Request !== 'undefined' && source instanceof Request) return openFetch(source, undefined, listeners, signal)
  if (source instanceof Response) return openResponse(source, listeners, signal)
  if (source instanceof Blob) return { kind: 'blob', blob: source }
  if (typeof source === 'object' && source !== null && typeof (source as HttpSource).url === 'string') {
    const { url, headers, credentials } = source as HttpSource
    return openFetch(url, { headers, credentials }, listeners, signal)
  }
  if (typeof ReadableStream !== 'undefined' && source instanceof ReadableStream) {
    return { kind: 'stream', reader: pullStream(source as ReadableStream<Uint8Array>, listeners.onProgress, undefined, signal) }
  }
  if (ArrayBuffer.isView(source) || source instanceof ArrayBuffer) return { kind: 'bytes', bytes: await readSource(source) }
  throw new Error('Unsupported byte source')
}
