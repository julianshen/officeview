/**
 * An OfficePackage view over the document's package handle (slideshow spec 8), so existing parse
 * helpers can run lazily without touching the document package's cache. Metadata calls delegate to
 * the handle's package; every materialising read is bounded and uncached at the source, with a
 * per-part cap and an instance-wide total. The private cache lives only as long as this instance.
 */
import JSZip from 'jszip'
import { OfficePackage } from '../../core/zip'
import { parseXml, parseXmlOrdered, type XmlNode } from '../../core/xml'
import { PackageError } from '../../core/package-error'
import type { PptxPackageHandle } from '../types'
import { readPartBounded } from './read'

/** Default inflated bytes per part. */
export const HANDLE_PART_BYTES = 4 * 1024 * 1024
/** Inflated bytes across all parts one instance (one enrichment run) reads. */
export const HANDLE_AGGREGATE_BYTES = 32 * 1024 * 1024

export class HandlePackage extends OfficePackage {
  private readonly reads = new Map<string, Promise<Uint8Array | undefined>>()
  private readonly trees = new Map<string, Promise<XmlNode | undefined>>()
  private total = 0
  /** Bumped by dropCache, so a read still in flight never republishes into a dropped cache. */
  private generation = 0

  private constructor(private readonly handle: PptxPackageHandle, private readonly partBytes: number, private readonly textBytes: number) {
    super(new JSZip())
  }

  /** `textBytes` (default `partBytes`) caps text and XML reads below a picture-sized `partBytes`. */
  static over(handle: PptxPackageHandle, caps: { partBytes?: number; textBytes?: number } = {}): HandlePackage {
    const partBytes = caps.partBytes ?? HANDLE_PART_BYTES
    return new HandlePackage(handle, partBytes, Math.min(caps.textBytes ?? partBytes, partBytes))
  }

  /** Drop this instance's reads and trees. The instance total is kept: bytes already inflated stay counted. */
  override dropCache(): void {
    this.generation++
    this.reads.clear()
    this.trees.clear()
    super.dropCache()
  }

  override has(path: string): boolean { return this.handle.pkg.has(path) }
  override entry(path: string): JSZip.JSZipObject | null { return this.handle.pkg.entry(path) }
  override paths(predicate?: (p: string) => boolean): string[] { return this.handle.pkg.paths(predicate) }

  /** One bounded read per part, charged once against the instance total. A read under a caller's
   * tighter cap is cached only when it completes (the part then fits every cap). */
  private read(path: string, cap = this.partBytes): Promise<Uint8Array | undefined> {
    const cached = this.reads.get(path)
    if (cached) return cached
    const read = readPartBounded(this.handle.pkg, path, Math.min(cap, this.partBytes), {
      admit: bytes => {
        if (this.total + bytes > HANDLE_AGGREGATE_BYTES) return false
        this.total += bytes
        return true
      },
    })
    const generation = this.generation
    if (cap >= this.partBytes) this.reads.set(path, read)
    else void read.then(bytes => { if (generation === this.generation && !this.reads.has(path)) this.reads.set(path, Promise.resolve(bytes)) }, () => {})
    return read
  }

  private async limited(path: string, maxBytes: number): Promise<Uint8Array | undefined> {
    if (!Number.isSafeInteger(maxBytes) || maxBytes < 0) throw new Error('Invalid part byte limit')
    const bytes = await this.read(path, maxBytes)
    if (bytes && bytes.length > maxBytes) throw new PackageError('resource-limit', `part exceeds ${maxBytes} bytes`, path)
    return bytes
  }

  override async text(path: string): Promise<string | undefined> {
    const bytes = await this.limited(path, this.textBytes)
    return bytes && new TextDecoder().decode(bytes)
  }

  /** A caller's own limit refuses with the base class's message, which its callers branch on. */
  override async bytes(path: string, maxBytes?: number): Promise<Uint8Array | undefined> {
    if (maxBytes === undefined) return this.read(path)
    try { return await this.limited(path, maxBytes) } catch (error) {
      if (error instanceof PackageError && error.code === 'resource-limit') throw new Error(`Part byte limit exceeded: ${error.message}`)
      throw error
    }
  }

  override async boundedText(path: string, maxBytes: number): Promise<string | undefined> {
    const bytes = await this.limited(path, maxBytes)
    return bytes && new TextDecoder().decode(bytes)
  }

  override xml(path: string): Promise<XmlNode | undefined> { return this.tree(`x:${path}`, path, parseXml) }
  override xmlOrdered(path: string): Promise<XmlNode | undefined> { return this.tree(`o:${path}`, path, parseXmlOrdered) }

  private tree(key: string, path: string, parse: (text: string) => XmlNode): Promise<XmlNode | undefined> {
    let tree = this.trees.get(key)
    if (!tree) {
      tree = this.text(path).then(text => text === undefined ? undefined : parse(text))
      this.trees.set(key, tree)
    }
    return tree
  }
}
