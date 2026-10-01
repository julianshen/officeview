/**
 * Streaming ingestion for office files.
 *
 * A fetch(), a Response, a Blob, or any ReadableStream<Uint8Array> can feed the
 * viewer. Chunks accumulate into one buffer with progress callbacks, so a
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

export type ByteSource =
  | ArrayBuffer
  | Uint8Array
  | ReadableStream<Uint8Array>
  | Blob
  | Response

export type ProgressListener = (progress: Progress) => void

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

/** Normalize any accepted byte source into a single Uint8Array buffer. */
export async function readSource(source: ByteSource, onProgress?: ProgressListener): Promise<Uint8Array> {
  if (source instanceof Response) {
    const totalHeader = source.headers.get('content-length')
    const total = totalHeader ? parseInt(totalHeader, 10) : undefined
    if (source.body) return readByteStream(source.body, onProgress, Number.isFinite(total) ? total : undefined)
    return new Uint8Array(await source.arrayBuffer())
  }
  if (source instanceof Blob) return readBlob(source, onProgress)
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
