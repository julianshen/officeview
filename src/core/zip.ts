/**
 * Container access. An Office file is a zip ("package") of parts with
 * relationships. This module hides JSZip behind a small cached-read API so
 * renderers can lazily grab parts without re-unzipping.
 */
import JSZip from 'jszip'
import { parseXml, parseXmlOrdered, type XmlNode } from './xml'

export class OfficePackage {
  private zip: JSZip
  private cache = new Map<string, unknown>()

  private constructor(zip: JSZip) {
    this.zip = zip
  }

  static async load(data: ArrayBuffer | Uint8Array): Promise<OfficePackage> {
    const zip = await JSZip.loadAsync(data)
    return new OfficePackage(zip)
  }

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
function boundedBytes(file: JSZip.JSZipObject, maxBytes: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const stream = (file as unknown as { internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array> }).internalStream('uint8array')
    const chunks: Uint8Array[] = []
    let length = 0
    // JSZip 3.10 StreamHelper has no cancel API. Catch at the source push/end
    // boundary so throwing from onData aborts the current pako inflate call,
    // then propagate error upstream to stop scheduled ticks and free workers.
    type Worker = { previous?: Worker; push: (...args: unknown[]) => unknown; end: () => unknown; error: (error: Error) => unknown }
    const last = (stream as unknown as { _worker: Worker })._worker
    let source = last
    while (source.previous) source = source.previous
    const push = source.push.bind(source), end = source.end.bind(source)
    const abort = (error: unknown) => { chunks.length = 0; reject(error); last.error(error instanceof Error ? error : new Error(String(error))) }
    source.push = (...args) => { try { return push(...args) } catch (error) { abort(error) } }
    source.end = () => { try { return end() } catch (error) { abort(error) } }
    stream.on('data', (chunk: Uint8Array) => {
      length += chunk.length
      if (length > maxBytes) throw new Error('Part byte limit exceeded during inflation')
      chunks.push(chunk)
    }).on('error', (error: Error) => { chunks.length = 0; reject(error) }).on('end', () => {
      const bytes = new Uint8Array(length)
      let offset = 0
      for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length }
      chunks.length = 0
      resolve(bytes)
    }).resume()
  })
}
