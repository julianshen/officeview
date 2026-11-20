/**
 * Container access. An Office file is a zip ("package") of parts with
 * relationships. This module hides JSZip behind a small cached-read API so
 * renderers can lazily grab parts without re-unzipping.
 */
import JSZip from 'jszip'
import { parseXml, parseXmlOrdered, type XmlNode } from './xml'
import { PackageError } from './package-error'

export class OfficePackage {
  private zip: JSZip
  private cache = new Map<string, unknown>()

  /** Protected: a subclass may serve parts from elsewhere (the paged drawing
   * reader reads them on demand through the range reader, S13.2). */
  protected constructor(zip: JSZip) {
    this.zip = zip
  }

  static async load(data: ArrayBuffer | Uint8Array): Promise<OfficePackage> {
    const zip = await JSZip.loadAsync(data)
    return new OfficePackage(zip)
  }

  /** Drop the read cache (the PPTX package handle keeps the archive, not parse-time reads). */
  dropCache(): void { this.cache.clear() }

  /** The underlying zip entry, read uncached by the lazy bounded chart reader. */
  entry(path: string): JSZip.JSZipObject | null { return this.zip.file(path) }

  has(path: string): boolean {
    return this.zip.file(path) !== null
  }

  /** Read a part as text. Cached. */
  async text(path: string): Promise<string | undefined> {
    if (this.cache.has(path)) return this.cache.get(path) as string | undefined
    const f = this.zip.file(path)
    if (!f) return undefined
    const s = await f.async('string')
    this.cache.set(path, s)
    return s
  }

  /** Read a part as raw bytes. Cached. */
  async bytes(path: string, maxBytes?: number): Promise<Uint8Array | undefined> {
    const cacheKey = `bytes:${path}`
    if (maxBytes !== undefined && (!Number.isSafeInteger(maxBytes) || maxBytes < 0)) throw new Error('Invalid part byte limit')
    if (this.cache.has(cacheKey)) {
      const cached = this.cache.get(cacheKey) as Uint8Array
      if (maxBytes !== undefined && cached.length > maxBytes) throw new Error('Part byte limit exceeded')
      return cached
    }
    const f = this.zip.file(path)
    if (!f) return undefined
    const buf = maxBytes === undefined ? await f.async('uint8array') : await boundedBytes(f, maxBytes)
    this.cache.set(cacheKey, buf)
    return buf
  }

  /** Streams a part, uncached, into `onChunk`, refusing past `maxBytes` of actual inflated bytes (declared sizes
   * are untrusted). `onChunk` throwing stops inflation and rejects with its error. Undefined: no such part. */
  async stream(path: string, maxBytes: number, onChunk: (chunk: Uint8Array) => void): Promise<number | undefined> {
    if (!Number.isSafeInteger(maxBytes) || maxBytes < 0) throw new Error('Invalid part byte limit')
    const f = this.zip.file(path)
    if (!f) return undefined
    return boundedStream(f, maxBytes, onChunk)
  }

  /** The part's declared (central directory) uncompressed size: untrusted, for refusing before any read. */
  declaredSize(path: string): number | undefined {
    const size = (this.zip.file(path) as unknown as { _data?: { uncompressedSize?: number } } | null)?._data?.uncompressedSize
    return Number.isSafeInteger(size) && size! >= 0 ? size : undefined
  }

  /** Read a part as UTF-8 text, refusing with PackageError 'resource-limit' once it
   * inflates past maxBytes (counted chunk by chunk; declared sizes are untrusted). */
  async boundedText(path: string, maxBytes: number): Promise<string | undefined> {
    let bytes: Uint8Array | undefined
    try {
      bytes = await this.bytes(path, maxBytes)
    } catch (error) {
      if (error instanceof Error && error.message.startsWith('Part byte limit exceeded')) {
        throw new PackageError('resource-limit', `part exceeds ${maxBytes} bytes`, path)
      }
      throw error
    }
    return bytes && new TextDecoder().decode(bytes)
  }

  /** Read only the first maxBytes of a part. Unlike boundedText this succeeds
   * for larger parts and stops JSZip's inflater as soon as the prefix arrives. */
  async prefixText(path: string, maxBytes: number): Promise<string | undefined> {
    const f = this.zip.file(path)
    if (!f) return undefined
    const bytes = await prefixBytes(f, maxBytes)
    return new TextDecoder().decode(bytes)
  }

  /** Read a part and parse it as XML. Cached. */
  async xml(path: string): Promise<XmlNode | undefined> {
    const cacheKey = `xml:${path}`
    if (this.cache.has(cacheKey)) return this.cache.get(cacheKey) as XmlNode | undefined
    const s = await this.text(path)
    if (s === undefined) return undefined
    const node = parseXml(s)
    this.cache.set(cacheKey, node)
    return node
  }

  /** Read a part and parse it as XML, keeping document order. Cached. */
  async xmlOrdered(path: string): Promise<XmlNode | undefined> {
    const cacheKey = `xml-ordered:${path}`
    if (this.cache.has(cacheKey)) return this.cache.get(cacheKey) as XmlNode | undefined
    const s = await this.text(path)
    if (s === undefined) return undefined
    const node = parseXmlOrdered(s)
    this.cache.set(cacheKey, node)
    return node
  }

  /** List all part paths matching a predicate. */
  paths(predicate?: (p: string) => boolean): string[] {
    const out: string[] = []
    this.zip.forEach((relPath, file) => {
      if (file.dir) return
      if (!predicate || predicate(relPath)) out.push(relPath)
    })
    return out
  }
}

/** Count actual inflate chunks before retaining them. Declared ZIP sizes are untrusted. */
async function boundedBytes(file: JSZip.JSZipObject, maxBytes: number): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let length: number
  try {
    length = await boundedStream(file, maxBytes, (chunk) => { chunks.push(chunk) })
  } catch (error) {
    chunks.length = 0
    throw error
  }
  const bytes = new Uint8Array(length)
  let offset = 0
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length }
  chunks.length = 0
  return bytes
}

/**
 * Inflates `file` chunk by chunk into `onChunk`, counting actual bytes (declared sizes are untrusted). Past
 * `maxBytes`, or when `onChunk` throws, inflation stops and the promise rejects with that error. Resolves with the
 * byte count. Uses JSZip 3.10 internals (as `prefixBytes`); if they change, the patching throws and callers fail
 * closed.
 */
function boundedStream(file: JSZip.JSZipObject, maxBytes: number, onChunk: (chunk: Uint8Array) => void): Promise<number> {
  return new Promise((resolve, reject) => {
    const stream = (file as unknown as { internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array> }).internalStream('uint8array')
    let length = 0
    let settled = false
    // JSZip 3.10 StreamHelper has no cancel API. Catch at the source push/end
    // boundary so throwing from onData aborts the current pako inflate call,
    // then propagate error upstream to stop scheduled ticks and free workers.
    type Worker = { previous?: Worker; push: (...args: unknown[]) => unknown; end: () => unknown; error: (error: Error) => unknown }
    const last = (stream as unknown as { _worker: Worker })._worker
    let source = last
    while (source.previous) source = source.previous
    const push = source.push.bind(source), end = source.end.bind(source)
    const abort = (error: unknown) => {
      if (!settled) { settled = true; reject(error) }
      last.error(error instanceof Error ? error : new Error(String(error)))
    }
    source.push = (...args) => { try { return push(...args) } catch (error) { abort(error) } }
    source.end = () => { try { return end() } catch (error) { abort(error) } }
    stream.on('data', (chunk: Uint8Array) => {
      length += chunk.length
      if (length > maxBytes) throw new Error('Part byte limit exceeded during inflation')
      onChunk(chunk)
    }).on('error', (error: Error) => { if (!settled) { settled = true; reject(error) } }).on('end', () => {
      if (!settled) { settled = true; resolve(length) }
    }).resume()
  })
}

/** Prefix counterpart to boundedBytes; aborts upstream inflation once enough
 * decompressed bytes have been observed, without caching a partial part.
 * Uses the same JSZip 3.10 internals as boundedBytes; if they change, the
 * patching throws and callers fail closed (an unreadable diagnostic). */
function prefixBytes(file: JSZip.JSZipObject, maxBytes: number): Promise<Uint8Array> {
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 1) return Promise.reject(new RangeError('maxBytes must be a positive safe integer'))
  return new Promise((resolve, reject) => {
    const stream = (file as unknown as { internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array> }).internalStream('uint8array')
    const chunks: Uint8Array[] = []
    let length = 0
    let stopping = false
    let settled = false
    const STOP = new Error('part prefix complete')
    type Worker = { previous?: Worker; push: (...args: unknown[]) => unknown; end: () => unknown; error: (error: Error) => unknown }
    const last = (stream as unknown as { _worker: Worker })._worker
    let source = last
    while (source.previous) source = source.previous
    const push = source.push.bind(source), end = source.end.bind(source)
    const join = () => {
      const out = new Uint8Array(length)
      let offset = 0
      for (const chunk of chunks) { out.set(chunk, offset); offset += chunk.length }
      chunks.length = 0
      return out
    }
    const fail = (error: unknown) => {
      if (settled) return
      settled = true
      chunks.length = 0
      reject(error)
      last.error(error instanceof Error ? error : new Error(String(error)))
    }
    source.push = (...args) => {
      try { return push(...args) } catch (error) {
        if (error === STOP) {
          stopping = true
          last.error(STOP)
          if (!settled) { settled = true; resolve(join()) }
          return
        }
        fail(error)
      }
    }
    source.end = () => {
      try { return end() } catch (error) { fail(error) }
    }
    stream.on('data', (chunk: Uint8Array) => {
      const take = Math.min(maxBytes - length, chunk.length)
      if (take > 0) { chunks.push(chunk.subarray(0, take)); length += take }
      if (length === maxBytes) throw STOP
    }).on('error', (error: Error) => {
      if (stopping || settled) return
      settled = true
      chunks.length = 0
      reject(error)
    }).on('end', () => {
      if (settled) return
      settled = true
      resolve(join())
    }).resume()
  })
}
