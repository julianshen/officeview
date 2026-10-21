/**
 * Bounded, uncached part reads through the PPTX package handle (slideshow spec 6.1, 8). Parts are
 * streamed from `entry()` and counted chunk by chunk (declared ZIP sizes are untrusted); nothing read
 * here lands in the document package's cache. The streamer is local: `charts/read.ts streamEntry`
 * lives in the lazy chart chunk, and importing it would load the chart engine with every enrichment.
 */
import type JSZip from 'jszip'
import type { OfficePackage } from '../../core/zip'
import { parseXml, parseXmlOrdered, type XmlNode } from '../../core/xml'
import { PackageError } from '../../core/package-error'
import type { PptxPackageHandle } from '../types'

export interface ReadPartOptions {
  signal?: AbortSignal
  /** Asked before each chunk is retained; false aborts the inflate with PackageError('resource-limit'). */
  admit?: (chunkBytes: number) => boolean
  /** 'chunks' returns the chunk list without a final concatenation (one copy of a large part). */
  output?: 'bytes' | 'chunks'
}

type Source = Pick<OfficePackage, 'entry'>
type Worker = { previous?: Worker; push: (...args: unknown[]) => unknown; end: () => unknown; error: (error: Error) => unknown }

export function readPartBounded(source: Source, path: string, maxBytes: number, options: ReadPartOptions & { output: 'chunks' }): Promise<Uint8Array[] | undefined>
export function readPartBounded(source: Source, path: string, maxBytes: number, options?: ReadPartOptions & { output?: 'bytes' }): Promise<Uint8Array | undefined>
export function readPartBounded(source: Source, path: string, maxBytes: number, options?: ReadPartOptions): Promise<Uint8Array | Uint8Array[] | undefined>
export function readPartBounded(source: Source, path: string, maxBytes: number, options: ReadPartOptions = {}): Promise<Uint8Array | Uint8Array[] | undefined> {
  const { signal, admit, output = 'bytes' } = options
  return new Promise((resolve, reject) => {
    signal?.throwIfAborted()
    const file = source.entry(path)
    if (!file || file.dir) return resolve(undefined)
    const stream = (file as unknown as { internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array> }).internalStream('uint8array')
    const chunks: Uint8Array[] = []
    let length = 0, failed = false
    // As core/zip boundedBytes: JSZip 3.10 has no cancel API, so a throw from the data handler is
    // caught at the source push/end boundary, which aborts the inflate and errors the worker chain.
    const last = (stream as unknown as { _worker: Worker })._worker
    let first = last
    while (first.previous) first = first.previous
    const push = first.push.bind(first), end = first.end.bind(first)
    const fail = (error: unknown) => {
      chunks.length = 0
      if (failed) return
      failed = true
      reject(error)
      last.error(error instanceof Error ? error : new Error(String(error)))
    }
    first.push = (...args) => { try { return push(...args) } catch (error) { fail(error) } }
    first.end = () => { try { return end() } catch (error) { fail(error) } }
    stream.on('data', (chunk: Uint8Array) => {
      if (failed) return
      signal?.throwIfAborted()
      if (length + chunk.length > maxBytes) throw new PackageError('resource-limit', `part exceeds ${maxBytes} bytes`, path)
      if (admit && !admit(chunk.length)) throw new PackageError('resource-limit', 'part read refused by its owner', path)
      length += chunk.length
      chunks.push(chunk)
    }).on('error', (error: Error) => { chunks.length = 0; if (!failed) { failed = true; reject(error) } }).on('end', () => {
      if (failed) return
      if (output === 'chunks') return resolve(chunks.splice(0))
      const bytes = new Uint8Array(length)
      let offset = 0
      for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length }
      chunks.length = 0
      resolve(bytes)
    }).resume()
  })
}

async function readText(source: Source, path: string, maxBytes: number, signal?: AbortSignal): Promise<string | undefined> {
  const bytes = await readPartBounded(source, path, maxBytes, { signal })
  return bytes && new TextDecoder().decode(bytes)
}

/** Read and parse one XML part under an inflated byte cap; undefined when absent. */
export async function readXmlBounded(source: Source, path: string, maxBytes: number, signal?: AbortSignal): Promise<XmlNode | undefined> {
  const text = await readText(source, path, maxBytes, signal)
  return text === undefined ? undefined : parseXml(text)
}

/** As readXmlBounded, keeping document order. */
export async function readXmlOrderedBounded(source: Source, path: string, maxBytes: number, signal?: AbortSignal): Promise<XmlNode | undefined> {
  const text = await readText(source, path, maxBytes, signal)
  return text === undefined ? undefined : parseXmlOrdered(text)
}

/** The measured cost of the lease (spec 8): the compressed entry bytes the handle keeps alive. */
export function retainedArchiveBytes(handle: PptxPackageHandle): number {
  let total = 0
  for (const path of handle.pkg.paths()) {
    const size = (handle.pkg.entry(path) as unknown as { _data?: { compressedSize?: unknown } } | null)?._data?.compressedSize
    if (typeof size === 'number' && Number.isFinite(size)) total += size
  }
  return total
}
