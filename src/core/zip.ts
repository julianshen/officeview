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
  async bytes(path: string): Promise<Uint8Array | undefined> {
    const cacheKey = `bytes:${path}`
    if (this.cache.has(cacheKey)) return this.cache.get(cacheKey) as Uint8Array | undefined
    const f = this.zip.file(path)
    if (!f) return undefined
    const buf = await f.async('uint8array')
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
