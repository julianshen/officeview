import { parseXmlOrdered, type XmlNode } from '../../core/xml'
import { PackageError, isPackageError, throwIfAborted } from '../../core/package-error'
import { readPartBounded } from '../../pptx/appearance/read'
import { validateSlideSize } from '../../pptx/appearance/size'
import type { PptxDocument, PptxPackageHandle, PptxSlide } from '../../pptx/types'
import type { AllocationOwner, DiagnosticSink, Lease, MetadataHandle, ShowIndex, ShowSlide, SlideMetadataStore, SlideShowWorkspace, SlideTimingReader, TargetResolver } from '../types'
import { relationships, relsPart, type Relationship } from './deck'
import { readNotes } from './notes'
import { readLinks } from './links'
import { readMedia } from './media'
import { READER_LIMITS as L } from './limits'

interface Args {
  doc: PptxDocument
  handle?: PptxPackageHandle
  index: ShowIndex
  workspace: SlideShowWorkspace
  sink: DiagnosticSink
  timing?: SlideTimingReader
  resolver: (slide: PptxSlide) => TargetResolver
  signal?: AbortSignal
  /** Injectable model-cost estimator; the construction reservation remains enforced. */
  estimate?: (slide: ShowSlide) => number
}

interface Pending {
  job: Promise<Entry>
  waiters: number
  cached: boolean
  entry?: Entry
}

interface Entry {
  slide: ShowSlide
  lease: Lease
  refs: number
  cached: boolean
}

/**
 * The store keeps two separate ledgers, and a slide is charged to each for what
 * that ledger is about.
 *
 * 1. The workspace carries TRANSIENT construction state through temporary leases:
 *    the inflated chunks, the joined buffer and decoded text, and the parsed
 *    XML tree. All of it is released when `load` returns, whatever the outcome.
 * 2. The 64 MiB metadata aggregate (spec 6.1: "model objects, estimated") carries
 *    only what a slide RETAINS: notes, links, media records, timing and builds.
 *    It is reserved, from counted source features, before the timing reader runs
 *    and then reconciled to the estimate of the finished model.
 *
 * Calibration (node 22 / V8, fixtures of `<p:sp>` shapes, `<a:r/>` runs, attribute-heavy
 * elements and long text, each padded to 4 MiB; heap after a forced GC and peak RSS
 * rise while parsing with `parseXmlOrdered`; the recalibration is recorded in the S1-7
 * fix-round report):
 *
 *   fixture   tags   '='      retained tree   peak RSS rise
 *   shape     218k   116k     86 MB           187 MB
 *   mixed     149k   79k      57 MB           132 MB
 *   `<a:r/>`  699k   0        193 MB          329 MB
 *   attrs     63k    626k     24 MB           83 MB
 *   text      32k    0        20 MB           65 MB
 *
 * The retained tree costs 276-625 B per opening tag; the rise in resident memory while parsing
 * is 470-2,000 B per tag with the parser's scratch. The charge below is 640 B per opening tag,
 * 16 B per '=' and 3 B per character, on top of 8 B per inflated byte for the buffers. For the
 * fixtures the charge is, in MB against the measured peak: shape 187 vs 187, mixed 143 vs 132,
 * attrs 96 vs 83, text 67 vs 65, `<a:r/>` 493 vs 329. So it is at or above the measured peak
 * for each, and about twice the retained tree; the `<a:r/>` extreme (699k tags) is refused on a
 * 256 MiB workspace. These are accounting estimates for admission, not heap telemetry.
 */
const BUFFER_BYTES_PER_BYTE = 8
const PARSE_BYTES_PER_TAG = 640
const PARSE_BYTES_PER_ATTRIBUTE = 16
const PARSE_BYTES_PER_CHAR = 3

/** Retained-model constants: the shape of one record plus the strings it can copy out of the source. */
const SHELL_BYTES = 2048
const LINK_BYTES = 256
const NOTE_PARAGRAPH_BYTES = 192
const MEDIA_BYTES = 256 + 2 * 255
/** Timing, transition and builds are built from the part of the slide after `p:cSld`; 16 B retained per source character. */
const TIMING_BYTES_PER_SOURCE_CHAR = 16
/** What the reader limits (spec 6.1) allow a slide's timing model to retain at most. */
const TIMING_CEILING = L.timeNodes * 1024 + L.conditions * 128 + L.keyframesPerSlide * 160 + L.motionPathBytes * 2

/** Object estimator for the retained slide model (strings count as UTF-16, objects carry a fixed header). */
export function slideMetadataBytes(value: unknown): number {
  if (typeof value === 'string') return 24 + value.length * 2
  if (!value || typeof value !== 'object') return 8
  if (Array.isArray(value)) return 32 + value.reduce((n, x) => n + 8 + slideMetadataBytes(x), 0)
  let n = 48
  for (const k of Object.keys(value)) n += 16 + k.length * 2 + slideMetadataBytes((value as Record<string, unknown>)[k])
  return n
}

interface Source {
  node: XmlNode
  text: string
}

/** Charge for parsing `text`: opening tags, attributes and characters. Closing tags, comments, declarations and CDATA openers are not nodes. */
function parseCost(text: string): number {
  let tags = 0
  let attributes = 0
  for (let i = text.indexOf('<'); i >= 0; i = text.indexOf('<', i + 1)) {
    const next = text.charCodeAt(i + 1)
    if (next !== 47 && next !== 63 && next !== 33) tags++
  }
  for (let i = text.indexOf('='); i >= 0; i = text.indexOf('=', i + 1)) attributes++
  return tags * PARSE_BYTES_PER_TAG + attributes * PARSE_BYTES_PER_ATTRIBUTE + text.length * PARSE_BYTES_PER_CHAR
}

/**
 * Reads and parses one part under the cap. Every byte and node is charged to a
 * temporary lease BEFORE it is allocated: chunks as they arrive, the parse tree from
 * the decoded text before `parseXmlOrdered` runs. Failures other than an abort are
 * typed PackageErrors (a corrupt entry or malformed XML is a `parse-error`).
 */
async function xml(handle: PptxPackageHandle, path: string, cap: number, owner: AllocationOwner, signal: AbortSignal, leases: Lease[]): Promise<Source | undefined> {
  const raw = owner.lease(0, `xml:${path}`)
  if (!raw) throw new PackageError('resource-limit', 'XML workspace')
  leases.push(raw)
  let bytes = 0
  let chunks: Uint8Array[] | undefined
  try {
    chunks = await readPartBounded(handle.pkg, path, cap, {
      signal,
      output: 'chunks',
      admit: n => {
        if (!raw.grow(n * BUFFER_BYTES_PER_BYTE + 128)) return false
        bytes += n
        return true
      },
    })
  } catch (e) {
    throw isPackageError(e) ? e : new PackageError('parse-error', e instanceof Error ? e.message : 'part read failed', path)
  }
  throwIfAborted(signal)
  if (!chunks) return undefined
  const joined = new Uint8Array(bytes)
  let offset = 0
  for (const c of chunks) {
    joined.set(c, offset)
    offset += c.length
  }
  const text = new TextDecoder().decode(joined)
  if (!raw.grow(parseCost(text) + 1024)) throw new PackageError('resource-limit', 'XML object workspace', path)
  try {
    return { node: parseXmlOrdered(text), text }
  } catch (e) {
    throw new PackageError('parse-error', e instanceof Error ? e.message : 'XML parse failed', path)
  }
}

const HYPERLINK_TAG = /<(?:[\w.-]+:)?hlinkClick\b(?:"[^"]*"|'[^']*'|[^'">])*>/g
/** The first end of `p:cSld`, or a self-closing one. A match inside a comment can only come earlier, which over-reserves. */
const CSLD_END = /<\/(?:[\w.-]+:)?cSld\s*>|<(?:[\w.-]+:)?cSld\b(?:"[^"]*"|'[^']*'|[^'">])*\/>/
const PARAGRAPH_TAG = /<(?:[\w.-]+:)?p[\s>/]/g

function count(text: string, pattern: RegExp): { n: number; chars: number } {
  const re = new RegExp(pattern.source, 'g')
  let n = 0
  let chars = 0
  for (let m = re.exec(text); m; m = re.exec(text)) {
    n++
    chars += m[0].length
  }
  return { n, chars }
}

/**
 * Upper bound on what building the slide's model can retain, from features counted in
 * the sources and from the reader limits; independent of how much plain markup the
 * slide has. It is reserved before any model object is built (spec 6.1: "charged as
 * slides are read"), and the finished model is reconciled to its own estimate.
 *
 * - links: one record per `a:hlinkClick`; the strings they copy (a verb) are no longer than the tag;
 *   hrefs and shape ids are shared references.
 * - notes: at most 64 KiB of text plus one record per paragraph tag.
 * - media: one record per media relationship.
 * - timing, transition and builds: 16 B per character of the slide after `p:cSld`, never more
 *   than the reader limits allow (2,000 time nodes, 10,000 conditions, 20,000 keyframes, 64 KiB paths).
 */
function modelBound(slide: Source, notes: Source | undefined, rels: readonly Relationship[], withTiming: boolean): number {
  const links = count(slide.text, HYPERLINK_TAG)
  const paragraphs = notes ? count(notes.text, PARAGRAPH_TAG).n : 0
  const end = CSLD_END.exec(slide.text)
  const region = end ? slide.text.length - end.index : slide.text.length
  let media = 0
  for (const r of rels) if (!r.external && /\/(audio|video|media)$/.test(r.type)) media += MEDIA_BYTES + 2 * r.target.length
  return SHELL_BYTES
    + links.n * LINK_BYTES + 2 * links.chars
    + (notes ? 2 * Math.min(notes.text.length, L.notesChars) + paragraphs * NOTE_PARAGRAPH_BYTES : 0)
    + media
    + (withTiming ? Math.min(TIMING_CEILING, TIMING_BYTES_PER_SOURCE_CHAR * region) : 0)
}

export function createSlideMetadataStore(args: Args): SlideMetadataStore {
  const { doc, handle, index, workspace, sink } = args
  const cache = new Map<number, Entry>()
  const pending = new Map<number, Pending>()
  const abort = new AbortController()
  let total = 0
  let closed = false
  const stop = () => store.close()
  args.signal?.addEventListener('abort', stop, { once: true })
  const diagnose = (i: number, code: string, message: string) => sink.add({ code, slideIndex: i, message })

  function dispose(e: Entry) {
    if (!e.cached && e.refs === 0) {
      total -= e.lease.bytes
      e.lease.release()
    }
  }

  function evict(i: number) {
    const p = pending.get(i)
    if (p) {
      p.cached = false
      pending.delete(i)
    }
    const e = cache.get(i)
    if (e) {
      cache.delete(i)
      e.cached = false
      dispose(e)
    }
  }

  /** Makes room in the aggregate by evicting unheld cached neighbours; true when `bytes` now fit. */
  function room(bytes: number, except?: Entry) {
    for (const [i, e] of cache) {
      if (total + bytes <= L.metadataBytes) break
      if (e !== except && e.refs === 0 && !pending.has(i)) evict(i)
    }
    return total + bytes <= L.metadataBytes
  }

  const empty = (i: number, why: string): ShowSlide => ({ index: i, hidden: index.slides[i]?.hidden ?? false, timing: { refused: why }, notes: [], links: [], media: [], builds: [] })

  function reserve(bytes: number, i: number): Lease | undefined {
    if (!Number.isSafeInteger(bytes) || bytes < 0 || !room(bytes)) return undefined
    const lease = workspace.lease(bytes, `meta:${i}`)
    if (lease) total += bytes
    return lease
  }

  /** Moves a construction reservation to the finished model's estimate; undefined (released) when growth does not fit. */
  function settle(model: Lease, bytes: number): Lease | undefined {
    if (!Number.isSafeInteger(bytes) || bytes < 0) throw new PackageError('resource-limit', 'Invalid metadata estimate')
    if (bytes > model.bytes) {
      const growth = bytes - model.bytes
      if (!room(growth) || !model.grow(growth)) {
        total -= model.bytes
        model.release()
        return undefined
      }
      total += growth
    } else {
      const refund = model.bytes - bytes
      model.shrink(refund)
      total -= refund
    }
    return model
  }

  async function load(i: number): Promise<Entry> {
    const signal = abort.signal
    const temporary: Lease[] = []
    let slide: ShowSlide | undefined
    let model: Lease | undefined
    throwIfAborted(signal)
    const shell = workspace.lease(2048, `metadata-shell:${i}`)
    if (!shell) throw new PackageError('resource-limit', 'Metadata shell workspace')
    temporary.push(shell)
    try {
      if (!handle || !index.slides[i]?.partPath) {
        sink.add({ code: 'no-package', message: 'No package; slides are static' })
        slide = empty(i, 'no-package')
      } else if (!validateSlideSize(doc.slides[i].widthEmu, doc.slides[i].heightEmu).ok) {
        diagnose(i, 'slide-size', 'Invalid slide size')
        slide = empty(i, 'slide-size')
      } else {
        const part = index.slides[i].partPath!
        const source = await xml(handle, part, L.slideBytes, workspace, signal, temporary)
        if (!source) throw new PackageError('parse-error', 'Missing slide', part)
        const rels = relationships((await xml(handle, relsPart(part), L.relsBytes, workspace, signal, temporary))?.node, part)
        const note = rels.find(r => !r.external && r.type.endsWith('/notesSlide'))
        let notes: Source | undefined
        let notesRefused = false
        if (note) {
          try {
            notes = await xml(handle, note.target, L.notesBytes, workspace, signal, temporary)
          } catch {
            throwIfAborted(signal)
            notesRefused = true
            diagnose(i, 'notes-limit', 'Notes part refused')
          }
        }
        const types = rels.some(r => /\/(audio|video|media)$/.test(r.type))
          ? await xml(handle, '[Content_Types].xml', L.contentTypesBytes, workspace, signal, temporary)
          : undefined
        // Admit the model against BOTH budgets before any of it is built.
        model = reserve(modelBound(source, notes, rels, !!args.timing), i)
        if (model) {
          const paragraphs = readNotes(notes?.node, i, {
            add(d) {
              if (d.code === 'notes-truncated') notesRefused = true
              sink.add(d)
            },
            list: () => sink.list(),
          })
          const timing = args.timing?.(source.node, i, args.resolver(doc.slides[i]), sink)
          slide = {
            index: i,
            hidden: index.slides[i].hidden,
            ...timing,
            timing: notesRefused ? { refused: 'notes-limit' } : timing?.timing,
            notes: paragraphs,
            links: readLinks(source.node, rels, handle.slideParts, i, sink),
            media: readMedia(rels, types?.node, handle.pkg, i, sink, source.node),
            // Builds belong to the timing they build for.
            builds: notesRefused ? [] : timing?.builds ?? [],
          }
          throwIfAborted(signal)
        }
      }
      const bytes = slide ? args.estimate?.(slide) ?? slideMetadataBytes(slide) : 0
      let lease: Lease | undefined
      if (slide) lease = model ? settle(model, bytes) : reserve(bytes, i)
      model = undefined
      if (slide && lease) return { slide, lease, refs: 0, cached: true }
      diagnose(i, 'metadata-limit', 'Aggregate metadata admission refused; slide is static')
      return await staticEntry(i, 'metadata-limit')
    } catch (e) {
      if (model) {
        total -= model.bytes
        model.release()
        model = undefined
      }
      throwIfAborted(signal)
      // A programming error (an injected reader throwing, a bug here) is not a limit refusal.
      if (!isPackageError(e)) throw e
      diagnose(i, 'slide-refused', e.message)
      return await staticEntry(i, e.code === 'resource-limit' ? 'read-limit' : 'read-error')
    } finally {
      for (const lease of temporary) lease.release()
    }
  }

  /** A static entry still needs its own small charge; when even that does not fit the acquire is rejected. */
  async function staticEntry(i: number, why: string): Promise<Entry> {
    const slide = empty(i, why)
    const lease = reserve(slideMetadataBytes(slide), i)
    if (!lease) throw new PackageError('resource-limit', 'Static metadata admission refused')
    return { slide, lease, refs: 0, cached: true }
  }

  function handleFor(e: Entry): MetadataHandle {
    e.refs++
    let released = false
    return {
      slide: e.slide,
      attach(bytes, _tag) {
        if (released || closed || !Number.isSafeInteger(bytes) || bytes < 0 || !room(bytes, e)) return false
        if (!e.lease.grow(bytes)) return false
        total += bytes
        return true
      },
      release() {
        if (released) return
        released = true
        e.refs--
        dispose(e)
      },
    }
  }

  const store: SlideMetadataStore = {
    async acquire(i, signal) {
      throwIfAborted(signal)
      throwIfAborted(abort.signal)
      if (!Number.isInteger(i) || i < 0 || i >= doc.slides.length) throw new RangeError('slide index')
      const cached = cache.get(i)
      if (cached) {
        cache.delete(i)
        cache.set(i, cached)
        return handleFor(cached)
      }
      let p = pending.get(i)
      if (!p) {
        const record: Pending = { job: undefined as never, waiters: 0, cached: true }
        record.job = load(i).then(e => {
          // The job owns one reference until all waiting acquisitions settle.
          e.refs++
          record.entry = e
          e.cached = record.cached && !closed
          if (e.cached) cache.set(i, e)
          return e
        })
        p = record
        pending.set(i, p)
      }
      p.waiters++
      try {
        const entry = await p.job
        throwIfAborted(signal)
        throwIfAborted(abort.signal)
        return handleFor(entry)
      } finally {
        p.waiters--
        if (p.waiters === 0) {
          if (pending.get(i) === p) pending.delete(i)
          if (p.entry) {
            p.entry.refs--
            dispose(p.entry)
          }
        }
      }
    },
    evict,
    resident: () => [...cache.keys()],
    close() {
      if (closed) return
      closed = true
      abort.abort(new PackageError('cancelled', 'Store closed'))
      args.signal?.removeEventListener('abort', stop)
      for (const i of cache.keys()) evict(i)
    },
  }
  if (args.signal?.aborted) store.close()
  return store
}
