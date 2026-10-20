/**
 * The bounded chart-part reader (spec 6.5, lazy chunk).
 *
 * A chart part is streamed (never cached in the package), held to the
 * 8/4 MiB byte cap chunk by chunk, decoded (UTF-8, or UTF-16 with a BOM) and
 * tokenized with the bounded SAX tokenizer (core/xml-stream). The ordered tree
 * the classic parser reads is built from the tokens, and every allocation is
 * charged to the document's ChartWorkspace *before* it is made:
 *
 * - `crossing`: the in-flight inflate chunk (pako emits at most 16 KiB; a
 *   larger chunk is refused) and the tokenizer's carried token, reserved once
 *   before the stream starts;
 * - per chunk, the accepted bytes and their decoded text, released after the
 *   tokenizer consumes them;
 * - every tree node (element, attributes, text), held until `release()`.
 *
 * A refusal (the workspace or its parent ledger) stops the read before the
 * allocation: the result is `refused` and everything charged is refunded.
 * A paged package's pool is the workspace's parent (`chartLedger`, wired by
 * `chartParseContextFor`; D0), so every byte charged here reaches that pool
 * once, through the workspace, and is refunded to it on release.
 *
 * The root element's namespace is resolved from its own declarations (default
 * or prefixed). Only the classic chart namespace (transitional or strict), or
 * an unbound root, continues; ChartEx and any other namespace stop the read at
 * the root, so a non-classic part never reaches the classic parser.
 */
import type JSZip from 'jszip'
import type { OfficePackage } from '../core/zip'
import { registerOrdered, type XmlNode } from '../core/xml'
import { XmlTokenizer, localName } from '../core/xml-stream'
import { PackageError } from '../core/package-error'
import type { ChartParseContext } from './context'
import { CHART_CHUNK_TAG, CHART_CROSSING_TAG } from './workspace'

const MiB = 1024 * 1024
export const CHART_PART_LIMIT = { desktop: 8 * MiB, mobile: 4 * MiB } as const
/** The largest inflate chunk admitted (pako's output chunk is 16 KiB). */
export const CHART_CHUNK_LIMIT = 64 * 1024
export const CHART_TOKEN_CHARS = 64 * 1024
const TOKEN_CHARS = CHART_TOKEN_CHARS
/** Reserved once before the stream: the crossing chunk plus the tokenizer's carried token (UTF-16). */
export const CHART_READ_CROSSING = CHART_CHUNK_LIMIT + 2 * TOKEN_CHARS
const SLICE = 16 * 1024
export { CHART_CROSSING_TAG, CHART_CHUNK_TAG }
const CLASSIC = new Set(['http://schemas.openxmlformats.org/drawingml/2006/chart', 'http://purl.oclc.org/ooxml/drawingml/chart'])

export type ChartRead =
  | { kind: 'tree'; root: XmlNode; release: () => void }
  | { kind: 'unsupported' | 'refused' | 'malformed'; reason: string; code: string; feature?: 'ChartEx'; limit?: number }

class Stop extends Error { constructor(readonly result: Exclude<ChartRead, { kind: 'tree' }>) { super(result.reason) } }

let reads = 0
/** Read one chart part through the bounded, admitted tokenizer path. `undefined` when the part is absent. */
export async function readChartPart(pkg: OfficePackage, path: string, context: Pick<ChartParseContext, 'profile' | 'workspace'>): Promise<ChartRead | undefined> {
  const entry = (pkg as { entry?: (path: string) => JSZip.JSZipObject | null }).entry?.(path) ?? null
  if (!entry && !pkg.has(path)) return undefined
  const workspace = context.workspace, cap = CHART_PART_LIMIT[context.profile]
  // Tags: the tree, the crossing reservation and the per-chunk transient. A parent that already admitted
  // the part's bytes and text (a paged package) is told which charges are transient decode work by the tag.
  const tag = `chart-read:${path}#${++reads}`, chunkTag = `${tag}${CHART_CHUNK_TAG}`, crossingTag = `${tag}${CHART_CROSSING_TAG}`
  const refund = () => { workspace.release(chunkTag); workspace.release(crossingTag); workspace.release(tag) }
  const refused = (): never => { throw new Stop({ kind: 'refused', reason: 'This chart needs more memory than is available for charts', code: 'chart-refused' }) }
  /** Tree bytes, charged to the workspace (and through it to its parent) before the node exists. */
  const charge = (bytes: number) => { if (!workspace.reserve(bytes, tag)) refused() }
  if (!workspace.reserve(CHART_READ_CROSSING, crossingTag)) return { kind: 'refused', reason: 'This chart needs more memory than is available for charts', code: 'chart-refused' }

  const builder = treeBuilder(charge)
  const tokenizer = new XmlTokenizer(builder.events, { maxTokenChars: TOKEN_CHARS })
  // Comments never reach the tree: they are discarded while streaming, so their length is not a token limit.
  const comments = new CommentStripper()
  let decoder: TextDecoder | undefined, head: Uint8Array | undefined, accepted = 0
  /** One inflated chunk: capped, charged, decoded and tokenized, then its transient charge is refunded. */
  const feed = (chunk: Uint8Array, owned: boolean) => {
    if (chunk.length > CHART_CHUNK_LIMIT) refused()
    if (accepted + chunk.length > cap) throw new Stop({ kind: 'unsupported', reason: 'This chart is too large to display', code: 'chart-part-limit', limit: cap })
    accepted += chunk.length
    // Accepted bytes (unless a paged package's ledger already owns them) and their decoded text.
    if (!workspace.reserve((owned ? chunk.length : 0) + 2 * chunk.length, chunkTag)) refused()
    try {
      let bytes = chunk
      if (!decoder) {
        bytes = head ? concat(head, chunk) : chunk
        if (bytes.length < 2) { head = bytes; return }
        head = undefined
        decoder = new TextDecoder(bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le' : bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8', { fatal: true })
      }
      tokenizer.write(comments.strip(decoder.decode(bytes, { stream: true })))
    } finally { workspace.release(chunkTag) }
  }
  const finish = () => {
    if (!decoder) { decoder = new TextDecoder('utf-8', { fatal: true }); if (head) tokenizer.write(comments.strip(decoder.decode(head, { stream: true }))) }
    tokenizer.write(comments.strip(decoder.decode(), true))
    tokenizer.end()
  }
  try {
    if (entry) await stream(entry, chunk => feed(chunk, true))
    else {
      // A paged package reads the part itself, admitted against its own pool (bounded by the same cap).
      let bytes: Uint8Array | undefined
      try { bytes = await pkg.bytes(path, cap) } catch (error) {
        if (String((error as Error | undefined)?.message).startsWith('Part byte limit')) throw new Stop({ kind: 'unsupported', reason: 'This chart is too large to display', code: 'chart-part-limit', limit: cap })
        throw error
      }
      if (!bytes) { refund(); return undefined }
      for (let at = 0; at < bytes.length; at += SLICE) feed(bytes.subarray(at, at + SLICE), false)
    }
    finish()
  } catch (error) {
    refund()
    if (error instanceof Stop) return error.result
    if (builder.stop) return builder.stop
    // A token over the reader's limit is a resource refusal, never malformed XML.
    if (error instanceof PackageError && error.code === 'resource-limit') return { kind: 'unsupported', reason: 'This chart contains markup larger than the viewer reads', code: 'chart-token-limit', limit: TOKEN_CHARS }
    if (error instanceof PackageError && error.code === 'unsupported') return { kind: 'unsupported', reason: 'This chart uses XML the viewer does not read', code: 'unsupported-xml' }
    return { kind: 'malformed', reason: 'Chart data could not be read', code: 'invalid-content-xml' }
  }
  workspace.release(chunkTag)
  return { kind: 'tree', root: builder.root!, release: refund }
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length)
  out.set(a); out.set(b, a.length)
  return out
}

/** Push the entry's inflated chunks, uncached; a throwing consumer aborts the inflate at the source. */
function stream(file: JSZip.JSZipObject, onChunk: (chunk: Uint8Array) => void): Promise<void> {
  return new Promise((resolve, reject) => {
    const helper = (file as unknown as { internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array> }).internalStream('uint8array')
    // As core/zip boundedBytes: JSZip 3.10 has no cancel; catch at the source push/end boundary.
    type Worker = { previous?: Worker; push: (...args: unknown[]) => unknown; end: () => unknown; error: (error: Error) => unknown }
    const last = (helper as unknown as { _worker: Worker })._worker
    let source = last
    while (source.previous) source = source.previous
    const push = source.push.bind(source), end = source.end.bind(source)
    let failed = false
    const abort = (error: unknown) => { if (failed) return; failed = true; reject(error); last.error(error instanceof Error ? error : new Error(String(error))) }
    source.push = (...args) => { try { return push(...args) } catch (error) { abort(error) } }
    source.end = () => { try { return end() } catch (error) { abort(error) } }
    helper.on('data', (chunk: Uint8Array) => { if (!failed) onChunk(chunk) })
      .on('error', (error: Error) => { if (!failed) { failed = true; reject(error) } })
      .on('end', () => { if (!failed) resolve() })
      .resume()
  })
}

/** Streams comments out (outside CDATA and processing instructions), holding at most 8 characters
 * between chunks; a comment's content is never retained. */
class CommentStripper {
  private state: 'text' | 'comment' | 'cdata' | 'pi' = 'text'
  private carry = ''
  strip(chunk: string, last = false): string {
    const input = this.carry + chunk
    this.carry = ''
    let out = '', i = 0
    while (i < input.length) {
      if (this.state === 'text') {
        const at = input.indexOf('<', i)
        if (at < 0) { out += input.slice(i); break }
        out += input.slice(i, at)
        if (!last && input.length - at < 9) { this.carry = input.slice(at); break }
        if (input.startsWith('<!--', at)) { this.state = 'comment'; i = at + 4 }
        else if (input.startsWith('<![CDATA[', at)) { out += '<![CDATA['; this.state = 'cdata'; i = at + 9 }
        else if (input.startsWith('<?', at)) { out += '<?'; this.state = 'pi'; i = at + 2 }
        else { out += '<'; i = at + 1 }
        continue
      }
      const close = this.state === 'comment' ? '-->' : this.state === 'cdata' ? ']]>' : '?>'
      const at = input.indexOf(close, i)
      if (at < 0) {
        const keep = Math.max(i, input.length - (close.length - 1))
        if (this.state !== 'comment') out += input.slice(i, keep)
        this.carry = last ? '' : input.slice(keep)
        if (last && this.state === 'comment') throw new PackageError('parse-error', 'unterminated comment')
        if (last) out += input.slice(keep)
        break
      }
      if (this.state !== 'comment') out += input.slice(i, at + close.length)
      this.state = 'text'
      i = at + close.length
    }
    return out
  }
}

/** Retained bytes priced per tree node (charged before the node exists): the node, its attribute record,
 * its ordered and grouped child lists and its slot in the parent's lists. */
export const ELEMENT_BYTES = 256
const ATTRIBUTE_BYTES = 48, TEXT_BYTES = 160
/** A namespace scope object, and each binding it owns. Inherited bindings are shared through the parent
 * scope, never copied, so each binding is priced once, by the element that declares it. */
export const NAMESPACE_SCOPE_BYTES = 96, NAMESPACE_BINDING_BYTES = 64

/** A persistent namespace scope: an element's own declarations over its parent's scope (shared). */
export class NamespaceScope implements ReadonlyMap<string, string> {
  /** Scopes constructed (tests observe the admission boundary). */
  static built = 0
  constructor(private readonly parent: NamespaceScope | undefined, private readonly own: ReadonlyMap<string, string>) { NamespaceScope.built++ }
  get ownSize(): number { return this.own.size }
  get(prefix: string): string | undefined { return this.own.get(prefix) ?? this.parent?.get(prefix) }
  has(prefix: string): boolean { return this.own.has(prefix) || !!this.parent?.has(prefix) }
  /** The visible bindings, materialized on demand (never retained). */
  private visible(): Map<string, string> {
    const out = new Map(this.parent?.visible())
    for (const [prefix, uri] of this.own) out.set(prefix, uri)
    return out
  }
  get size(): number { return this.visible().size }
  forEach(fn: (value: string, key: string, map: ReadonlyMap<string, string>) => void): void { this.visible().forEach((value, key) => fn(value, key, this)) }
  entries(): MapIterator<[string, string]> { return this.visible().entries() }
  keys(): MapIterator<string> { return this.visible().keys() }
  values(): MapIterator<string> { return this.visible().values() }
  [Symbol.iterator](): MapIterator<[string, string]> { return this.entries() }
}
/** The shared scope of nodes under no declaration at all. */
const NO_NAMESPACES = new NamespaceScope(undefined, new Map())

/** Builds the same ordered tree as `parseXmlOrdered` from tokenizer events, charging each node before it is built. */
function treeBuilder(charge: (bytes: number) => void) {
  interface Open { node: XmlNode; attrs: Record<string, string>; ordered: Array<[string, XmlNode]>; grouped: Map<string, XmlNode[]>; namespaces: NamespaceScope | undefined; text: string }
  const stack: Open[] = []
  const state: { root?: XmlNode; stop?: Exclude<ChartRead, { kind: 'tree' }> } = {}
  const flushText = (open: Open) => {
    if (!open.text) return
    charge(TEXT_BYTES)
    open.node['#text'] = ((open.node['#text'] as string | undefined) ?? '') + open.text
    open.ordered.push(['#text', { '@attrs': {}, '#text': open.text }])
    open.text = ''
  }
  const finish = (open: Open) => {
    flushText(open)
    for (const [name, list] of open.grouped) open.node[name] = list.length === 1 ? list[0] : list
    open.node['@attrs'] = open.attrs
    registerOrdered(open.node, open.ordered, open.namespaces ?? NO_NAMESPACES)
  }
  const declaration = (name: string) => name === 'xmlns' || name.startsWith('xmlns:')
  const events = {
    open(qname: string, names: readonly string[], values: readonly string[], count: number, selfClosing: boolean) {
      const parent = stack[stack.length - 1]
      if (parent) flushText(parent)
      // Price everything this node retains first: nothing below is allocated before the charge succeeds.
      let bytes = ELEMENT_BYTES + 2 * qname.length, declared = 0
      for (let i = 0; i < count; i++) {
        bytes += ATTRIBUTE_BYTES + 2 * (names[i].length + values[i].length)
        if (declaration(names[i])) declared++
      }
      if (declared) bytes += NAMESPACE_SCOPE_BYTES + declared * NAMESPACE_BINDING_BYTES
      if (!parent) {
        // Dispatch on the root's real namespace (its own declarations): only classic (or unbound) continues.
        const colon = qname.indexOf(':'), want = colon < 0 ? 'xmlns' : `xmlns:${qname.slice(0, colon)}`
        let uri: string | undefined
        for (let i = 0; i < count; i++) if (names[i] === want) uri = values[i]
        if (localName(qname) === 'chartSpace' && uri !== undefined && !CLASSIC.has(uri)) {
          state.stop = /chartex/i.test(uri)
            ? { kind: 'unsupported', reason: 'ChartEx charts are not supported yet', code: 'unsupported-chart', feature: 'ChartEx' }
            : { kind: 'unsupported', reason: 'This chart format is not supported', code: 'unsupported-chart' }
          throw new Stop(state.stop)
        }
      }
      charge(bytes)
      let namespaces = parent?.namespaces
      if (declared) {
        const own = new Map<string, string>()
        for (let i = 0; i < count; i++) if (declaration(names[i])) own.set(names[i] === 'xmlns' ? '' : names[i].slice(6), values[i])
        namespaces = new NamespaceScope(namespaces, own)
      }
      const attrs: Record<string, string> = {}
      for (let i = 0; i < count; i++) if (!names[i].startsWith('xmlns')) attrs[localName(names[i])] = values[i]
      const open: Open = { node: {}, attrs, ordered: [], grouped: new Map(), namespaces, text: '' }
      if (parent) {
        const name = localName(qname)
        const list = parent.grouped.get(name) ?? []
        list.push(open.node); parent.grouped.set(name, list); parent.ordered.push([name, open.node])
      } else state.root = open.node
      if (selfClosing) finish(open)
      else stack.push(open)
    },
    close() { finish(stack.pop()!) },
    // The segment and the node's merged text: charged before either is kept.
    text(text: string) { const open = stack[stack.length - 1]; if (open) { charge(4 * text.length); open.text += text } },
  }
  return { events, get root() { return state.root }, get stop() { return state.stop } }
}
