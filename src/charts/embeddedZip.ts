/**
 * Bounded access to a DOCX/PPTX chart's embedded workbook (spec 6.4; lazy
 * chunk). It produces the embedding's bytes and its central-directory
 * inventory; reading cells, and data precedence, are the embedded workbook
 * reader's (charts/embeddedWorkbook.ts).
 *
 * - **Resolve.** Chart part `.rels` -> `c:externalData@r:id`. Only a target
 *   inside the package is followed: `TargetMode="External"`, an absolute URI,
 *   a backslash or NUL, or a `..` that climbs out of the package root (also
 *   percent-encoded) is unavailable with a diagnostic.
 * - **Admit.** The embedding is inflated through the outer package's raw zip
 *   entry (`OfficePackage.entry`), never through the cached `bytes()`, so the
 *   outer package retains nothing. Actual inflated bytes are counted against
 *   8 MiB desktop / 4 MiB mobile; the entry's declared size is never trusted.
 *   The embedding's central directory is then inspected before any load: at
 *   most 256 entries, declared sizes summing to at most 32 / 16 MiB.
 * - **Ownership.** Every allocation is charged to the document's chart
 *   workspace before it is made: the in-flight chunk, each retained chunk, the
 *   contiguous copy and the inventory. A success hands those charges to the
 *   returned workbook until its `release()`; a failure refunds everything.
 * - **One at a time.** One embedded read per document (the workspace's
 *   document facet): the next waits until the current workbook is released, or
 *   until a failed or cancelled producer settles.
 * - **Cancellation.** The document's `AbortSignal` stops publication at once
 *   (`cancelled`). The inflate is not abortable mid-chunk, so its charges stay
 *   held (`ChartWorkspace.hold`) until it actually settles; its next chunk
 *   throws at the source, which ends it.
 */
import type JSZip from 'jszip'
import type { OfficePackage } from '../core/zip'
import { PackageError } from '../core/package-error'
import { XmlTokenizer, localName } from '../core/xml-stream'
import { bytesRangeSource, readZipInventory, type ZipInventory, type ZipLimits } from '../core/zip-stream'
import type { ChartParseContext } from './context'
import type { ChartDiagnostic } from './model'
import { CHART_CHUNK_LIMIT, CommentStripper, streamEntry } from './read'
import { CHART_CHUNK_TAG, CHART_CROSSING_TAG, type ChartProfile, type ChartWorkspace, type ScopedChartWorkspace } from './workspace'

const MiB = 1024 * 1024
/** Inflated embedding cap. Real chart workbooks are 15-100 KB. */
export const EMBEDDED_WORKBOOK_LIMIT: Readonly<Record<ChartProfile, number>> = Object.freeze({ desktop: 8 * MiB, mobile: 4 * MiB })
export const EMBEDDED_ENTRY_LIMIT = 256
/** Declared inflated sizes of the embedding's parts, summed (the same figure as the per-package XML budget). */
export const EMBEDDED_DECLARED_LIMIT: Readonly<Record<ChartProfile, number>> = Object.freeze({ desktop: 32 * MiB, mobile: 16 * MiB })
const NAME_BYTES = 512
/** The inventory's upper bound, charged before it is built: per entry its record and its name twice (entry and folded key) at 2 B per unit. */
export const EMBEDDED_INVENTORY_BYTES = EMBEDDED_ENTRY_LIMIT * (384 + 2 * 2 * NAME_BYTES)

/** A chart part's relationships part is read only up to this many inflated bytes (Office writes well under 1 KB). */
export const EMBEDDED_RELS_LIMIT = 256 * 1024
const RELS_TOKEN_CHARS = 16 * 1024
/** The relationships grammar (OPC Part 2 §9.3): `Relationships` > flat `Relationship`; anything else is malformed. */
const RELS_DEPTH = 2, RELS_NAME_CHARS = 64, RELS_ATTRIBUTES = 8, RELS_VALUE_CHARS = 4096, RELS_TEXT_CHARS = 1024
/**
 * Reserved once before the relationships stream, in UTF-16 bytes: the in-flight chunk; the carried token; the
 * last accepted token's names and values the tokenizer still references (at most one token); the element stack
 * (depth 2, names checked before they are pushed); the pending whitespace segment and its slice.
 */
export const EMBEDDED_RELS_CROSSING = CHART_CHUNK_LIMIT + 2 * (2 * RELS_TOKEN_CHARS + RELS_DEPTH * RELS_NAME_CHARS + 2 * RELS_TEXT_CHARS)

export interface EmbeddedWorkbook {
  kind: 'workbook'
  /** The embedding's part path in the outer package. */
  path: string
  /** The inflated embedding (throws after `release()`). */
  readonly bytes: Uint8Array
  /** Its central directory, admitted. */
  inventory: ZipInventory
  /** Bytes charged to the document workspace until `release()`. */
  charged: number
  /** Refunds the charges and frees the document's embedded-read slot (idempotent). */
  release(): void
}
export type EmbeddedWorkbookResult =
  | EmbeddedWorkbook
  | { kind: 'unavailable' | 'refused' | 'cancelled'; diagnostic: ChartDiagnostic; limit?: number }
export type EmbeddedFailure = Exclude<EmbeddedWorkbookResult, EmbeddedWorkbook>
type Failure = EmbeddedFailure
export type EmbeddedTarget = { kind: 'internal'; path: string } | Failure

export const failure = (kind: Failure['kind'], code: string, reason: string, limit?: number): Failure => ({ kind, diagnostic: { code, reason }, ...(limit !== undefined ? { limit } : {}) })
let reads = 0
const unavailable = (code: string, reason: string) => failure('unavailable', code, reason) as { kind: 'unavailable'; diagnostic: ChartDiagnostic }
const CORRUPT = () => failure('unavailable', 'chart-data-corrupt', "The chart's embedded workbook could not be read")
const CANCELLED = () => failure('cancelled', 'chart-cancelled', 'Reading the chart data was cancelled')
const REFUSED = () => failure('refused', 'chart-refused', 'This chart needs more memory than is available for charts')

/** A package-internal part path, or undefined when the target leaves the package or is not a relative/absolute part reference. */
export function internalPath(owner: string, target: string): string | undefined {
  let decoded: string
  try { decoded = decodeURIComponent(target) } catch { return undefined }
  if (/[\\\0?#]/.test(decoded) || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(decoded) || decoded.startsWith('//')) return undefined
  const bits = decoded.startsWith('/') ? [] : owner.split('/').slice(0, -1)
  for (const bit of decoded.split('/')) {
    if (bit === '..') { if (!bits.length) return undefined; bits.pop() } else if (bit && bit !== '.') bits.push(bit)
  }
  return bits.length ? bits.join('/') : undefined
}

class Stop extends Error { constructor(readonly result: Failure) { super(result.diagnostic.reason) } }

/** Charging for one operation: `charge` reserves under the operation's tag (refused -> Stop). */
interface Charging { workspace: ChartWorkspace; tag: string; charge(bytes: number): void; live(): void }

/**
 * The chart part's `.rels`, streamed through the raw entry (never the package's cached text/XML), held
 * to EMBEDDED_RELS_LIMIT chunk by chunk, comments stripped, tokenized with the bounded tokenizer. Only
 * the requested relationship's Target and TargetMode are kept, charged before they are retained; each
 * chunk and its decoded text are charged while they are tokenized, the crossing reservation for the
 * whole stream. Undefined when the part or the id is absent.
 */
async function readRelationship(pkg: OfficePackage, owner: string, rId: string, op: Charging): Promise<{ target?: string; external: boolean } | undefined> {
  const slash = owner.lastIndexOf('/')
  const relsPath = `${slash < 0 ? '' : owner.slice(0, slash) + '/'}_rels/${owner.slice(slash + 1)}.rels`
  const entry = rawEntry(pkg, relsPath)
  if (!entry || entry.dir) return undefined
  const { workspace } = op, crossingTag = `${op.tag}:rels${CHART_CROSSING_TAG}`, chunkTag = `${op.tag}:rels${CHART_CHUNK_TAG}`
  const malformed = () => new Stop(unavailable('chart-data-rels-malformed', "The chart's relationships could not be read"))
  try {
    if (!workspace.reserve(EMBEDDED_RELS_CROSSING, crossingTag)) throw new Stop(REFUSED())
    let found: { target?: string; external: boolean } | undefined, depth = 0
    const tokenizer = new XmlTokenizer({
      // The grammar is checked here, before the tokenizer pushes the name or keeps the attributes past this token.
      open(qname, names, values, count, selfClosing) {
        if (qname.length > RELS_NAME_CHARS || localName(qname) !== (depth === 0 ? 'Relationships' : 'Relationship') || depth >= RELS_DEPTH) throw malformed()
        if (count > RELS_ATTRIBUTES) throw malformed()
        for (let i = 0; i < count; i++) if (names[i].length > RELS_NAME_CHARS || values[i].length > RELS_VALUE_CHARS) throw malformed()
        if (!selfClosing) depth++
        if (found || localName(qname) !== 'Relationship') return
        let id: string | undefined, target: string | undefined, mode: string | undefined
        for (let i = 0; i < count; i++) {
          if (names[i] === 'Id') id = values[i]
          else if (names[i] === 'Target') target = values[i]
          else if (names[i] === 'TargetMode') mode = values[i]
        }
        if (id !== rId) return
        op.charge(32 + 2 * (target?.length ?? 0))
        found = { ...(target !== undefined ? { target } : {}), external: mode === 'External' }
      },
      close() { depth-- },
      text(text) { if (/\S/.test(text)) throw malformed() },
    }, { maxTokenChars: RELS_TOKEN_CHARS, maxTextSegmentChars: RELS_TEXT_CHARS, maxDepth: RELS_DEPTH + 1, maxAttributes: RELS_ATTRIBUTES + 1 })
    const comments = new CommentStripper(), decoder = new TextDecoder('utf-8', { fatal: true })
    let accepted = 0
    const write = (text: string) => {
      try { tokenizer.write(text) } catch (error) {
        if (error instanceof Stop) throw error
        if (error instanceof PackageError && error.code === 'resource-limit') throw new Stop(failure('refused', 'chart-rels-limit', "The chart's relationships are larger than the viewer reads", RELS_TOKEN_CHARS))
        throw malformed()
      }
    }
    await streamEntry(entry, chunk => {
      op.live()
      if (chunk.length > CHART_CHUNK_LIMIT) throw new Stop(REFUSED())
      if (accepted + chunk.length > EMBEDDED_RELS_LIMIT) throw new Stop(failure('refused', 'chart-rels-limit', "The chart's relationships are larger than the viewer reads", EMBEDDED_RELS_LIMIT))
      accepted += chunk.length
      // The chunk and its decoded text, released once the tokenizer has consumed them.
      if (!workspace.reserve(3 * chunk.length, chunkTag)) throw new Stop(REFUSED())
      try {
        let text: string
        try { text = decoder.decode(chunk, { stream: true }) } catch { throw malformed() }
        write(comments.strip(text))
      } finally { workspace.release(chunkTag) }
    })
    op.live()
    let tail: string
    try { tail = decoder.decode() } catch { throw malformed() }
    write(comments.strip(tail, true))
    try { tokenizer.end() } catch { throw malformed() }
    return found
  } catch (error) {
    if (error instanceof Stop) throw error
    // An inflate failure of the relationships part.
    throw malformed()
  } finally {
    workspace.release(chunkTag)
    workspace.release(crossingTag)
  }
}

/** The relationship's target, checked: internal, inside the package, present. */
async function resolveTarget(pkg: OfficePackage, chartPath: string, rId: string, op: Charging): Promise<EmbeddedTarget> {
  const rel = await readRelationship(pkg, chartPath, rId, op)
  if (!rel?.target) return unavailable('chart-data-missing', "The chart's data relationship is missing")
  if (rel.external) return unavailable('chart-data-external', 'The chart data is linked to an external workbook, which the viewer does not open')
  const path = internalPath(chartPath, rel.target)
  if (!path) return unavailable('chart-data-outside-package', 'The chart data relationship points outside the document package')
  const entry = rawEntry(pkg, path)
  if (!entry || entry.dir) return unavailable('chart-data-missing', "The chart's embedded workbook is missing")
  return { kind: 'internal', path }
}

/**
 * Chart part `.rels` -> the `c:externalData` relationship's internal target, through the bounded,
 * admitted relationships reader. Its work is charged to `workspace` and refunded on return.
 */
export async function resolveExternalData(context: Pick<ChartParseContext, 'pkg' | 'partPath' | 'workspace'>, rId: string): Promise<EmbeddedTarget> {
  const { pkg, partPath, workspace } = context, tag = `embedded-rels:${partPath}!${rId}#${++reads}`
  try {
    return await resolveTarget(pkg, partPath, rId, { workspace, tag, live() {}, charge(bytes) { if (!workspace.reserve(bytes, tag)) throw new Stop(REFUSED()) } })
  } catch (error) {
    if (error instanceof Stop) return error.result
    throw error
  } finally { workspace.release(tag) }
}
const rawEntry = (pkg: OfficePackage, path: string): JSZip.JSZipObject | null => {
  const entry = (pkg as { entry?: (path: string) => JSZip.JSZipObject | null }).entry
  try { return entry ? entry.call(pkg, path) : null } catch { return null }
}

/** One embedded read at a time per document: the queue is keyed by the document's chart facet. */
const lanes = new WeakMap<object, Promise<void>>()
/** True once `p` settles, false as soon as `signal` aborts (whichever is first). */
export function untilAbort<T>(p: Promise<T>, signal?: AbortSignal): Promise<{ value: T } | undefined> {
  if (!signal) return p.then(value => ({ value }))
  if (signal.aborted) return Promise.resolve(undefined)
  return new Promise<{ value: T } | undefined>((resolve, reject) => {
    const onAbort = () => resolve(undefined)
    signal.addEventListener('abort', onAbort, { once: true })
    p.then(value => { signal.removeEventListener('abort', onAbort); resolve({ value }) }, error => { signal.removeEventListener('abort', onAbort); reject(error) })
  })
}
async function acquire(workspace: ChartWorkspace, signal?: AbortSignal): Promise<(() => void) | undefined> {
  const key = (workspace as Partial<ScopedChartWorkspace>).facet ?? workspace
  const previous = lanes.get(key) ?? Promise.resolve()
  let free!: () => void
  const mine = new Promise<void>(resolve => { free = resolve })
  lanes.set(key, previous.then(() => mine))
  // A read cancelled while queued passes its turn on at once and holds nothing.
  if (!(await untilAbort(previous, signal)) || signal?.aborted) { free(); return undefined }
  let freed = false
  return () => { if (!freed) { freed = true; free() } }
}

const inventoryLimits = (profile: ChartProfile): ZipLimits => ({
  maxEntries: EMBEDDED_ENTRY_LIMIT,
  maxCentralDirectoryBytes: EMBEDDED_ENTRY_LIMIT * (46 + NAME_BYTES + 2 * 1024),
  maxNameBytes: NAME_BYTES,
  maxPartExpandedBytes: EMBEDDED_DECLARED_LIMIT[profile],
  maxTotalExpandedBytes: EMBEDDED_DECLARED_LIMIT[profile],
  // Ratio alone never refuses: the reader counts actual inflated bytes against its budgets.
  maxRatio: Number.MAX_SAFE_INTEGER,
})
function inventoryFailure(error: unknown, profile: ChartProfile): Failure {
  if (!(error instanceof PackageError) || error.code !== 'resource-limit') return CORRUPT()
  if (error.reason === 'entry-limit') return failure('refused', 'embedded-workbook-entries', "The chart's embedded workbook has more parts than the viewer reads", EMBEDDED_ENTRY_LIMIT)
  if (error.reason === 'central-directory-limit' || error.reason === 'name-limit') return failure('refused', 'embedded-workbook-directory', "The chart's embedded workbook has a directory larger than the viewer reads")
  return failure('refused', 'embedded-workbook-declared', "The chart's embedded workbook declares more data than the viewer reads", EMBEDDED_DECLARED_LIMIT[profile])
}

/**
 * Resolve, inflate, admit and inventory one chart's embedded workbook. The
 * caller must `release()` a returned workbook once it has read it: that
 * refunds its charges and lets the document's next embedded read start.
 * Relationship resolution is part of the producer: it runs in the document's
 * slot, is charged under the producer's tag, and an abort during it returns
 * `cancelled` at once while its charges stay held until it settles.
 */
export async function openEmbeddedWorkbook(context: Pick<ChartParseContext, 'pkg' | 'partPath' | 'profile' | 'workspace' | 'signal'>, rId: string): Promise<EmbeddedWorkbookResult> {
  const { pkg, partPath, profile, workspace, signal } = context
  if (signal?.aborted) return CANCELLED()
  const unlock = await acquire(workspace, signal)
  if (!unlock) return CANCELLED()
  const tag = `embedded-zip:${partPath}!${rId}#${++reads}`
  let abandoned = false
  // The producer's charges are held until it settles, whatever happens to the signal.
  const settled = workspace.hold(produce(pkg, partPath, rId, profile, workspace, tag, () => abandoned, unlock), tag)
    .then(result => { if (result.kind !== 'workbook') unlock(); return result }, () => { unlock(); return CORRUPT() })
  const outcome = await untilAbort(settled, signal)
  if (outcome && !signal?.aborted) return outcome.value
  // Publication stops now; a workbook the producer still completes is released, never handed out.
  abandoned = true
  void settled.then(result => { if (result.kind === 'workbook') result.release() })
  return CANCELLED()
}

async function produce(pkg: OfficePackage, partPath: string, rId: string, profile: ChartProfile, workspace: ChartWorkspace, tag: string, abandoned: () => boolean, unlock: () => void): Promise<EmbeddedWorkbookResult> {
  const cap = EMBEDDED_WORKBOOK_LIMIT[profile]
  const chunkTag = `${tag}:chunks`, crossingTag = `${tag}${CHART_CROSSING_TAG}`
  let charged = 0
  const charge = (bytes: number, into = tag) => {
    if (!workspace.reserve(bytes, into)) throw new Stop(REFUSED())
    if (into === tag) charged += bytes
  }
  const live = () => { if (abandoned()) throw new Stop(CANCELLED()) }
  try {
    const target = await resolveTarget(pkg, partPath, rId, { workspace, tag, live, charge: bytes => charge(bytes) })
    live()
    if (target.kind !== 'internal') return target
    const path = target.path, entry = rawEntry(pkg, path)
    if (!entry) return unavailable('chart-data-missing', "The chart's embedded workbook is missing")
    // The in-flight inflate chunk (pako emits at most 16 KiB; a larger one is refused).
    charge(CHART_CHUNK_LIMIT, crossingTag)
    const chunks: Uint8Array[] = []
    let length = 0
    await streamEntry(entry, chunk => {
      live()
      if (chunk.length > CHART_CHUNK_LIMIT) throw new Stop(REFUSED())
      // Actual inflated bytes, never the declared size: a forged header cannot pass the cap.
      if (length + chunk.length > cap) throw new Stop(failure('refused', 'embedded-workbook-limit', "The chart's embedded workbook is larger than the viewer reads", cap))
      charge(chunk.length, chunkTag)
      chunks.push(chunk)
      length += chunk.length
    })
    live()
    charge(length)
    const bytes = new Uint8Array(length)
    let at = 0
    for (const chunk of chunks) { bytes.set(chunk, at); at += chunk.length }
    chunks.length = 0
    workspace.release(chunkTag)
    // The central directory, before anything loads the embedding.
    charge(EMBEDDED_INVENTORY_BYTES)
    let inventory: ZipInventory
    try { inventory = await readZipInventory(bytesRangeSource(bytes), inventoryLimits(profile)) } catch (error) { throw new Stop(inventoryFailure(error, profile)) }
    live()
    const owner = `${tag}:workbook`
    workspace.transfer(tag, owner, charged)
    let held: Uint8Array | undefined = bytes, released = false
    return {
      kind: 'workbook', path, inventory, charged,
      get bytes() { if (!held) throw new Error('embedded workbook released'); return held },
      release() { if (released) return; released = true; held = undefined; workspace.release(owner); unlock() },
    }
  } catch (error) {
    if (error instanceof Stop) return error.result
    // An inflate failure (not deflate, truncated, CRC): the embedding is unreadable.
    return CORRUPT()
  } finally {
    workspace.release(chunkTag)
    workspace.release(crossingTag)
  }
}
