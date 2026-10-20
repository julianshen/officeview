/**
 * The bounded, values-only reader of a DOCX/PPTX chart's embedded workbook
 * (spec 6.4; lazy chunk). It is not `parseXlsx`: after A3a's admitted,
 * uncached embedding (`openEmbeddedWorkbook`), it loads the embedding's zip
 * directory and streams only these parts through the bounded tokenizer:
 *
 * 1. `xl/workbook.xml`: the sheets, `workbookPr@date1904` and only the
 *    defined names a chart asks for (formula length capped);
 * 2. `xl/_rels/workbook.xml.rels`: only the referenced sheets' targets
 *    (package-internal; external or escaping targets are never followed);
 * 3. `xl/styles.xml`: `numFmts` and `cellXfs` only (the read stops after them);
 * 4. only the sheets the chart's references name, keeping only the cells on
 *    the chart's lines and stopping after the last referenced row;
 * 5. `xl/sharedStrings.xml`, last, keeping only the strings those cells use,
 *    while the whole table is held to 8/4 MiB (2 B per unit plus 32 B per
 *    string) and 262,144 strings.
 *
 * Nothing else is followed (drawings, nested embeddings, external links,
 * VBA). Formula cells use their saved `<v>`; a formula without one is an
 * unavailable point. Formulas are never evaluated. Error cells keep their code.
 *
 * Ownership: every allocation is charged to the document chart workspace
 * before it is made: the zip directory structures, each stream's crossing
 * reservation (in-flight chunk, carried token, element stack, text segment)
 * and chunk, text being read, sheets, names, formats, and every kept cell
 * and string. All inflated XML of one embedding is held to 32/16 MiB. A
 * breach refuses the read and refunds everything. Cells (charged at
 * CHART_CELL_BYTES each, at most 262,144 per chart) are held until every
 * consumer of the read has filled its chart.
 *
 * Reuse: requests for one embedding part (per document facet) made while its
 * read waits for the document's embedded-read slot join that read and share
 * its result; nothing is kept after the last consumer is done, so a later
 * request reads again.
 */
import { OfficePackage } from '../core/zip'
import { PackageError } from '../core/package-error'
import { XmlTokenizer, localName, type XmlEvents } from '../core/xml-stream'
import { formatExtended, NO_SECTION } from '../xlsx/cellFormat'
import type { ChartParseContext } from './context'
import type { ChartDiagnostic, ChartRange } from './model'
import { CHART_CHUNK_LIMIT, CommentStripper, streamEntry } from './read'
import { EMBEDDED_DECLARED_LIMIT, failure, internalPath, openEmbeddedWorkbook, resolveExternalData, untilAbort, type EmbeddedFailure, type EmbeddedWorkbook } from './embeddedZip'
import { CHART_CELL_BYTES, chartLine, scanChartNumber, type ChartCell, type ChartCellLookup, type ChartHiddenLookup, type ChartLine } from './fill'
import { CHART_LABEL_LIMIT, parseChartRange } from './parse'
import { stringBytes } from './points'
import { sourceFormat } from './format'
import { CHART_CHUNK_TAG, CHART_CROSSING_TAG, type ChartProfile, type ChartWorkspace, type ScopedChartWorkspace } from './workspace'

const MiB = 1024 * 1024
/** Shared strings, the whole table: UTF-16 bytes plus 32 B per string. */
export const EMBEDDED_SST_BYTES: Readonly<Record<ChartProfile, number>> = Object.freeze({ desktop: 8 * MiB, mobile: 4 * MiB })
export const EMBEDDED_SST_COUNT = 262_144
/** Cells on one chart's lines (64 series x 4,096 points). */
export const EMBEDDED_CELL_LIMIT = 262_144
/** Inflated XML read per embedded package (every part streamed, summed). */
export const EMBEDDED_XML_LIMIT = EMBEDDED_DECLARED_LIMIT
const TOKEN_CHARS = 32 * 1024, NAME_CHARS = 256, DEPTH = 32, ATTRIBUTES = 64, SEGMENT_CHARS = 16 * 1024
/**
 * Reserved once per streamed part, in UTF-16 bytes: the in-flight chunk; the carried token and the last
 * accepted token's names and values the tokenizer still references (one token each); the element stack
 * (names checked against NAME_CHARS before the tokenizer pushes them); the pending text segment and its slice.
 */
export const EMBEDDED_STREAM_CROSSING = CHART_CHUNK_LIMIT + 2 * (2 * TOKEN_CHARS + DEPTH * NAME_CHARS + 2 * SEGMENT_CHARS)
/** Units of one cell value or string read (a longer one is kept cut, its numeric reading unavailable). */
const VALUE_CHARS = 1024
/** Units of a defined name's formula (a longer one is not a usable reference). */
const FORMULA_CHARS = 8192
const SHEET_LIMIT = 4096, NUMFMT_LIMIT = 4096, XF_LIMIT = 65_536, HIDDEN_COL_LIMIT = 16_384
/** JSZip's directory structures for a loaded embedding (bounded by A3a's admitted inventory). */
const LOAD_BASE_BYTES = 64 * 1024, LOAD_ENTRY_BYTES = 2048
const SHEET_RECORD_BYTES = 64, NAME_RECORD_BYTES = 64, XF_BYTES = 8, NUMFMT_RECORD_BYTES = 48, INTERVAL_BYTES = 48, HIDDEN_BYTES = 24
const COLS = 16_384
const key = (row: number, col: number) => row * COLS + col

const CORRUPT = () => failure('unavailable', 'chart-data-corrupt', "The chart's embedded workbook could not be read")
const CANCELLED = () => failure('cancelled', 'chart-cancelled', 'Reading the chart data was cancelled')
const REFUSED = () => failure('refused', 'chart-refused', 'This chart needs more memory than is available for charts')
class Stop extends Error { constructor(readonly result: EmbeddedFailure) { super(result.diagnostic.reason) } }
/** Ends a stream early (the rest of the part is not needed); never a failure. */
class Done extends Error {}

/** One role reference a chart asks for. `single`: one cell (a series name). */
export interface EmbeddedRequest { range: ChartRange; text: boolean; limit: number; single?: boolean }
/** A request's outcome: the concrete range read, or why the role is unavailable. */
export type EmbeddedResolution = { range: ChartRange } | { unavailable: ChartDiagnostic }
export interface EmbeddedCells {
  kind: 'cells'
  /** One per request, in order. */
  resolved: EmbeddedResolution[]
  lookup: ChartCellLookup
  hidden: ChartHiddenLookup
  /** The workbook's date system (`workbookPr@date1904`). */
  epoch: '1900' | '1904'
  /** The first sheet: unqualified references resolve against it. */
  ownerSheet: string
  /** This consumer has filled its chart (idempotent); the last one refunds the cells. */
  done(): void
}
export type EmbeddedCellsResult = EmbeddedCells | EmbeddedFailure

interface Stats { parts: Record<string, number>; cells: number; strings: number; rows: number }
let lastStats: Stats = { parts: {}, cells: 0, strings: 0, rows: 0 }
/** Counters of the most recent read (tests): inflated bytes per part, kept cells and strings, rows visited. */
export const embeddedWorkbookStats = (): Stats => lastStats

interface Consumer { requests: EmbeddedRequest[]; settle(result: EmbeddedCellsResult): void }
interface Batch { consumers: Consumer[]; frozen: boolean }
const batches = new WeakMap<object, Map<string, Batch>>()
let serial = 0
const facetOf = (workspace: ChartWorkspace): object => (workspace as Partial<ScopedChartWorkspace>).facet ?? workspace

/**
 * The cells a chart's references need from its embedded workbook. The caller must call `done()` on a
 * `cells` result once it has filled its chart. A cancelled request returns at once; the read it joined
 * keeps its charges until it settles.
 */
export async function requestEmbeddedCells(context: Pick<ChartParseContext, 'pkg' | 'partPath' | 'profile' | 'workspace' | 'signal'>, rId: string, requests: EmbeddedRequest[]): Promise<EmbeddedCellsResult> {
  const { signal } = context
  if (signal?.aborted) return CANCELLED()
  const target = await resolveExternalData(context, rId)
  if (target.kind !== 'internal') return target
  const facet = facetOf(context.workspace)
  let open = batches.get(facet)
  if (!open) batches.set(facet, open = new Map())
  let consumer!: Consumer
  const mine = new Promise<EmbeddedCellsResult>(resolve => { consumer = { requests, settle: resolve } })
  const joined = open.get(target.path)
  if (joined && !joined.frozen) joined.consumers.push(consumer)
  else {
    const batch: Batch = { consumers: [consumer], frozen: false }
    const list = open
    open.set(target.path, batch)
    void runBatch(batch, context, rId, target.path, () => { batch.frozen = true; if (list.get(target.path) === batch) list.delete(target.path) })
  }
  const outcome = await untilAbort(mine, signal)
  if (outcome && !signal?.aborted) return outcome.value
  // Publication stops now; a result that still arrives is handed back at once.
  void mine.then(result => { if (result.kind === 'cells') result.done() })
  return CANCELLED()
}

/** A stored cell: a number, a text (with its numeric reading), an error, or a formula without a saved value. */
interface Stored { n?: number; text?: string; number?: number | null; truncated?: true; error?: string; unavailable?: true; xf?: number; sst?: number }
interface Plan {
  name: string; path?: string; rId?: string
  vertical: Map<number, Array<[number, number]>>; horizontal: Map<number, Array<[number, number]>>
  minRow: number; maxRow: number
  cells: Map<number, Stored>; hiddenRows: Set<number>; hiddenCols: Array<[number, number]>
}
interface Book { date1904: boolean; sheets: Array<{ name: string; rId?: string }>; names: Array<{ name: string; source?: string; local?: number }> }

/** One read's charging and stream state. */
class Reader {
  inflated = 0
  readonly stats: Stats = { parts: {}, cells: 0, strings: 0, rows: 0 }
  constructor(readonly workspace: ChartWorkspace, readonly tag: string, readonly profile: ChartProfile, private readonly signal?: AbortSignal) {}
  charge(bytes: number, sub: string): void { if (!this.workspace.reserve(bytes, `${this.tag}:${sub}`)) throw new Stop(REFUSED()) }
  release(sub: string): void { this.workspace.release(`${this.tag}:${sub}`) }
  live(): void { if (this.signal?.aborted) throw new Stop(CANCELLED()) }

  /** Streams one inner part through the bounded tokenizer; false when it is absent. `Done` from `events` ends it early. */
  async stream(inner: OfficePackage, path: string, events: XmlEvents): Promise<boolean> {
    const entry = inner.entry(path)
    if (!entry || entry.dir) return false
    const crossing = `stream${CHART_CROSSING_TAG}`, chunkTag = `stream${CHART_CHUNK_TAG}`
    this.charge(EMBEDDED_STREAM_CROSSING, crossing)
    const cap = EMBEDDED_XML_LIMIT[this.profile]
    const malformed = () => new Stop(CORRUPT())
    // Names are checked before the tokenizer pushes them or keeps them past this token.
    const guarded: XmlEvents = {
      open(qname, names, values, count, selfClosing) {
        if (qname.length > NAME_CHARS) throw malformed()
        for (let i = 0; i < count; i++) if (names[i].length > NAME_CHARS) throw malformed()
        events.open(qname, names, values, count, selfClosing)
      },
      close: qname => events.close(qname),
      text: (text, more) => events.text(text, more),
    }
    const tokenizer = new XmlTokenizer(guarded, { maxTokenChars: TOKEN_CHARS, maxTextSegmentChars: SEGMENT_CHARS, maxDepth: DEPTH, maxAttributes: ATTRIBUTES })
    const comments = new CommentStripper(), decoder = new TextDecoder('utf-8', { fatal: true })
    const write = (text: string) => {
      try { tokenizer.write(text) } catch (error) {
        if (error instanceof Stop || error instanceof Done) throw error
        if (error instanceof PackageError && error.code === 'resource-limit') throw new Stop(failure('refused', 'embedded-workbook-markup', "The chart's embedded workbook contains markup larger or deeper than the viewer reads"))
        throw malformed()
      }
    }
    let read = 0
    try {
      await streamEntry(entry, chunk => {
        this.live()
        if (chunk.length > CHART_CHUNK_LIMIT) throw new Stop(REFUSED())
        if (this.inflated + chunk.length > cap) throw new Stop(failure('refused', 'embedded-workbook-xml', "The chart's embedded workbook holds more data than the viewer reads", cap))
        this.inflated += chunk.length
        read += chunk.length
        // The chunk and its decoded text, released once the tokenizer has consumed them.
        this.charge(3 * chunk.length, chunkTag)
        try {
          let text: string
          try { text = decoder.decode(chunk, { stream: true }) } catch { throw malformed() }
          write(comments.strip(text))
        } finally { this.release(chunkTag) }
      })
      this.live()
      let tail: string
      try { tail = decoder.decode() + '' } catch { throw malformed() }
      try { write(comments.strip(tail, true)) } catch (error) { if (error instanceof PackageError) throw malformed(); throw error }
      try { tokenizer.end() } catch { throw malformed() }
    } catch (error) {
      if (error instanceof Done) return true
      if (error instanceof Stop) throw error
      throw malformed()
    } finally {
      this.release(chunkTag)
      this.release(crossing)
      this.stats.parts[path] = (this.stats.parts[path] ?? 0) + read
    }
    return true
  }
}

/** Text of one element being read: kept up to `cap` units, each kept unit charged (transient) before it is kept. */
class Slot {
  text = ''
  units = 0
  constructor(private readonly reader: Reader, private readonly cap: number) {}
  add(text: string): void {
    this.units += text.length
    const room = this.cap - this.text.length
    if (room <= 0) return
    const take = text.length <= room ? text : text.slice(0, room)
    this.reader.charge(2 * take.length, 'text')
    this.text += take
  }
  get over(): boolean { return this.units > this.cap }
}
const attr = (names: readonly string[], values: readonly string[], count: number, name: string): string | undefined => {
  for (let i = 0; i < count; i++) if (localName(names[i]) === name) return values[i]
  return undefined
}
const truthy = (raw: string | undefined) => raw === '1' || raw === 'true'

/** `xl/workbook.xml`: sheets, the epoch, and the defined names in `wanted` (upper case). */
async function readWorkbook(inner: OfficePackage, reader: Reader, wanted: Set<string>): Promise<Book> {
  const book: Book = { date1904: false, sheets: [], names: [] }
  let depth = 0, section = '', name: { name: string; local?: number; slot: Slot } | undefined
  const found = await reader.stream(inner, 'xl/workbook.xml', {
    open(qname, names, values, count, selfClosing) {
      const level = depth, local = localName(qname)
      if (!selfClosing) depth++
      if (level === 1) {
        section = local
        if (local === 'workbookPr') book.date1904 = truthy(attr(names, values, count, 'date1904'))
      } else if (level === 2 && section === 'sheets' && local === 'sheet') {
        if (book.sheets.length >= SHEET_LIMIT) throw new Stop(failure('refused', 'embedded-workbook-sheets', "The chart's embedded workbook has more sheets than the viewer reads", SHEET_LIMIT))
        const sheetName = attr(names, values, count, 'name') ?? '', rId = attr(names, values, count, 'id')
        reader.charge(SHEET_RECORD_BYTES + stringBytes(sheetName) + (rId ? stringBytes(rId) : 0), 'data')
        book.sheets.push({ name: sheetName, ...(rId ? { rId } : {}) })
      } else if (level === 2 && section === 'definedNames' && local === 'definedName' && !selfClosing) {
        const text = attr(names, values, count, 'name') ?? ''
        if (!wanted.has(text.toUpperCase())) return
        const raw = attr(names, values, count, 'localSheetId'), index = raw === undefined ? undefined : Number(raw)
        reader.charge(NAME_RECORD_BYTES + stringBytes(text), 'data')
        name = { name: text, ...(index !== undefined && Number.isSafeInteger(index) && index >= 0 ? { local: index } : {}), slot: new Slot(reader, FORMULA_CHARS) }
      }
    },
    close() {
      depth--
      if (depth === 2 && name) {
        const { slot } = name
        if (!slot.over) reader.charge(stringBytes(slot.text), 'data')
        book.names.push({ name: name.name, ...(slot.over ? {} : { source: slot.text }), ...(name.local !== undefined ? { local: name.local } : {}) })
        name = undefined
        reader.release('text')
      }
      if (depth === 1) section = ''
    },
    text(text) { name?.slot.add(text) },
  })
  if (!found) throw new Stop(CORRUPT())
  return book
}

/** `xl/_rels/workbook.xml.rels`: the targets of the relationship ids in `wanted` (a flat list; anything else is malformed). */
async function readWorkbookRels(inner: OfficePackage, reader: Reader, wanted: Set<string>): Promise<Map<string, string>> {
  const out = new Map<string, string>()
  let depth = 0
  await reader.stream(inner, 'xl/_rels/workbook.xml.rels', {
    open(qname, names, values, count, selfClosing) {
      const level = depth
      if (!selfClosing) depth++
      if (localName(qname) !== (level === 0 ? 'Relationships' : 'Relationship') || level > 1) throw new Stop(CORRUPT())
      if (level !== 1) return
      const id = attr(names, values, count, 'Id'), target = attr(names, values, count, 'Target')
      if (id === undefined || !wanted.has(id) || out.has(id) || target === undefined) return
      // An external target is never followed; an escaping one is not a part.
      if (attr(names, values, count, 'TargetMode') === 'External') return
      const path = internalPath('xl/workbook.xml', target)
      if (!path) return
      reader.charge(stringBytes(id) + stringBytes(path), 'data')
      out.set(id, path)
    },
    close() { depth-- },
    text(text) { if (/\S/.test(text)) throw new Stop(CORRUPT()) },
  })
  return out
}

interface Styles { numFmts: Map<number, string>; xfs: number[] }
/** `xl/styles.xml`: `numFmts` and `cellXfs` only; the stream ends after `cellXfs`. */
async function readStyles(inner: OfficePackage, reader: Reader): Promise<Styles> {
  const styles: Styles = { numFmts: new Map(), xfs: [] }
  let depth = 0, section = ''
  await reader.stream(inner, 'xl/styles.xml', {
    open(qname, names, values, count, selfClosing) {
      const level = depth, local = localName(qname)
      if (!selfClosing) depth++
      if (level === 1) section = local
      else if (level === 2 && section === 'numFmts' && local === 'numFmt' && styles.numFmts.size < NUMFMT_LIMIT) {
        const id = Number(attr(names, values, count, 'numFmtId')), code = attr(names, values, count, 'formatCode')
        if (!Number.isSafeInteger(id) || code === undefined) return
        const kept = code.length > 255 ? 'General' : code
        reader.charge(NUMFMT_RECORD_BYTES + stringBytes(kept), 'data')
        styles.numFmts.set(id, kept)
      } else if (level === 2 && section === 'cellXfs' && local === 'xf' && styles.xfs.length < XF_LIMIT) {
        const id = Number(attr(names, values, count, 'numFmtId') ?? 0)
        reader.charge(XF_BYTES, 'data')
        styles.xfs.push(Number.isSafeInteger(id) && id >= 0 ? id : 0)
      }
    },
    close() { depth--; if (depth === 1) { if (section === 'cellXfs') throw new Done(); section = '' } },
    text() {},
  })
  return styles
}

/** A1-style reference (`B12`) to 0-based coordinates. */
function cellRef(ref: string): { row: number; col: number } | undefined {
  let i = 0, col = 0
  while (i < ref.length && i < 4) { const c = ref.charCodeAt(i) | 0x20; if (c < 97 || c > 122) break; col = col * 26 + c - 96; i++ }
  if (!i || i > 3 || col > COLS) return undefined
  const row = Number(ref.slice(i))
  return Number.isSafeInteger(row) && row >= 1 ? { row: row - 1, col: col - 1 } : undefined
}
const within = (intervals: Array<[number, number]> | undefined, at: number) => { if (intervals) for (const [a, b] of intervals) if (at >= a && at <= b) return true; return false }
const needed = (plan: Plan, row: number, col: number) => within(plan.vertical.get(col), row) || within(plan.horizontal.get(row), col)
const neededRow = (plan: Plan, row: number) => row >= plan.minRow && row <= plan.maxRow && (plan.horizontal.has(row) || [...plan.vertical.values()].some(list => within(list, row)))

/** A text's kept form: the label prefix and its numeric reading (Excel's numeric-text rules). */
function textCell(slot: Slot): Pick<Stored, 'text' | 'number' | 'truncated'> {
  const text = slot.text
  let number: number | null | undefined
  if (slot.over) number = /^[ +\-.0-9eE]*$/.test(text) ? undefined : null
  else number = scanChartNumber(text.length, i => text.charCodeAt(i)) ?? null
  return { text: text.length > CHART_LABEL_LIMIT + 1 ? text.slice(0, CHART_LABEL_LIMIT + 1) : text, ...(number === undefined ? { number: null, truncated: true as const } : { number }) }
}

/** One referenced sheet: only cells on the plan's lines are kept; the stream ends after the last referenced row. */
async function readSheet(inner: OfficePackage, reader: Reader, plan: Plan): Promise<void> {
  let depth = 0, section = '', row = -1, col = -1
  let current: { row: number; col: number; t?: string; xf?: number; formula: boolean; slot?: Slot; inIs: number } | undefined
  let reading: 'v' | 't' | undefined, skipPhonetic = false
  const keep = (stored: Stored, at: number, textBytes: number) => {
    reader.charge(CHART_CELL_BYTES + textBytes, 'data')
    plan.cells.set(at, stored)
    reader.stats.cells++
  }
  const found = await reader.stream(inner, plan.path!, {
    open(qname, names, values, count, selfClosing) {
      const level = depth, local = localName(qname)
      if (!selfClosing) depth++
      if (level === 1) { section = local; return }
      if (section === 'cols' && level === 2 && local === 'col') {
        if (!truthy(attr(names, values, count, 'hidden'))) return
        const min = Number(attr(names, values, count, 'min')) - 1, max = Number(attr(names, values, count, 'max')) - 1
        if (!Number.isSafeInteger(min) || !Number.isSafeInteger(max) || min < 0 || max < min || plan.hiddenCols.length >= HIDDEN_COL_LIMIT) return
        reader.charge(HIDDEN_BYTES, 'data')
        plan.hiddenCols.push([min, Math.min(max, COLS - 1)])
        return
      }
      if (section !== 'sheetData') return
      if (level === 2 && local === 'row') {
        const raw = attr(names, values, count, 'r'), r = raw === undefined ? row + 1 : Number(raw) - 1
        row = Number.isSafeInteger(r) && r >= 0 ? r : row + 1
        col = -1
        // Reading stops after the last referenced row.
        if (row > plan.maxRow) throw new Done()
        reader.stats.rows++
        if (truthy(attr(names, values, count, 'hidden')) && neededRow(plan, row) && !plan.hiddenRows.has(row)) { reader.charge(HIDDEN_BYTES, 'data'); plan.hiddenRows.add(row) }
        return
      }
      if (level === 3 && local === 'c') {
        const ref = attr(names, values, count, 'r'), at = ref ? cellRef(ref) : undefined
        col = at ? at.col : col + 1
        if (at && at.row !== row) row = at.row
        if (row < plan.minRow || !needed(plan, row, col)) { current = undefined; return }
        const xf = Number(attr(names, values, count, 's'))
        current = { row, col, t: attr(names, values, count, 't'), ...(Number.isSafeInteger(xf) && xf > 0 ? { xf } : {}), formula: false, inIs: 0 }
        if (selfClosing) { current = undefined }
        return
      }
      if (!current) return
      if (level === 4 && local === 'f') current.formula = true
      else if (level === 4 && local === 'v' && !selfClosing) { current.slot ??= new Slot(reader, VALUE_CHARS); reading = 'v' }
      else if (level === 4 && local === 'is') current.inIs = level
      else if (current.inIs && local === 'rPh') skipPhonetic = !selfClosing
      else if (current.inIs && local === 't' && !selfClosing && !skipPhonetic) { current.slot ??= new Slot(reader, VALUE_CHARS); reading = 't' }
    },
    close(qname) {
      depth--
      const local = localName(qname)
      if (depth === 1) { section = ''; return }
      if (!current) return
      if (local === 'v' || local === 't') reading = undefined
      else if (local === 'rPh') skipPhonetic = false
      else if (depth === 3 && local === 'c') {
        const cell = current, slot = cell.slot
        current = undefined
        const at = key(cell.row, cell.col), xf = cell.xf !== undefined ? { xf: cell.xf } : {}
        if (cell.t === 'e') { if (slot?.text) keep({ error: slot.text.trim().slice(0, 16), ...xf }, at, 0); else if (cell.formula) keep({ unavailable: true }, at, 0) }
        // A formula without a saved value has no value to show (never evaluated).
        else if (cell.formula && !slot) keep({ unavailable: true }, at, 0)
        else if (!slot) { /* blank */ }
        else if (cell.t === 's') { const index = Number(slot.text.trim()); if (Number.isSafeInteger(index) && index >= 0) keep({ sst: index, ...xf }, at, 0) }
        else if (cell.t === 'str' || cell.t === 'inlineStr' || cell.t === 'd') { const text = textCell(slot); keep({ ...text, ...(cell.t === 'd' ? { number: null } : {}), ...xf }, at, stringBytes(text.text!)) }
        else if (cell.t === 'b') { const text = slot.text.trim() === '1' ? 'TRUE' : 'FALSE'; keep({ text, number: null, ...xf }, at, stringBytes(text)) }
        else { const n = scanChartNumber(slot.text.length, i => slot.text.charCodeAt(i)); if (n !== undefined && !slot.over) keep({ n, ...xf }, at, 0) }
        reader.release('text')
      }
    },
    text(text) { if (reading && current?.slot) current.slot.add(text) },
  })
  if (!found) throw new Stop(CORRUPT())
}

/** `xl/sharedStrings.xml`: the whole table is held to its budgets; only the strings in `wanted` are kept. */
async function readSharedStrings(inner: OfficePackage, reader: Reader, wanted: Set<number>): Promise<Map<number, Pick<Stored, 'text' | 'number' | 'truncated'>>> {
  const out = new Map<number, Pick<Stored, 'text' | 'number' | 'truncated'>>()
  const cap = EMBEDDED_SST_BYTES[reader.profile]
  let depth = 0, index = -1, bytes = 0, units = 0, reading = false, phonetic = false, slot: Slot | undefined
  const found = await reader.stream(inner, 'xl/sharedStrings.xml', {
    open(qname, _names, _values, _count, selfClosing) {
      const level = depth, local = localName(qname)
      if (!selfClosing) depth++
      if (level === 1 && local === 'si') {
        if (++index >= EMBEDDED_SST_COUNT) throw new Stop(failure('refused', 'embedded-sst-count', "The chart's embedded workbook has more shared strings than the viewer reads", EMBEDDED_SST_COUNT))
        units = 0
        slot = wanted.has(index) ? new Slot(reader, VALUE_CHARS) : undefined
        if (selfClosing) { if ((bytes += 32) > cap) throw new Stop(failure('refused', 'embedded-sst-limit', "The chart's embedded workbook has more shared string text than the viewer reads", cap)); if (slot) { out.set(index, textCell(slot)); reader.charge(32, 'data'); slot = undefined } }
      } else if (local === 'rPh') phonetic = !selfClosing
      else if (local === 't' && !selfClosing && !phonetic && depth >= 3) reading = true
    },
    close(qname) {
      depth--
      const local = localName(qname)
      if (local === 't') reading = false
      else if (local === 'rPh') phonetic = false
      else if (depth === 1 && local === 'si') {
        bytes += 2 * units + 32
        if (bytes > cap) throw new Stop(failure('refused', 'embedded-sst-limit', "The chart's embedded workbook has more shared string text than the viewer reads", cap))
        if (slot) {
          const text = textCell(slot)
          reader.charge(stringBytes(text.text!), 'data')
          out.set(index, text)
          reader.stats.strings++
          slot = undefined
          reader.release('text')
        }
      }
    },
    text(text) {
      if (!reading) return
      units += text.length
      if (bytes + 2 * units + 32 > cap) throw new Stop(failure('refused', 'embedded-sst-limit', "The chart's embedded workbook has more shared string text than the viewer reads", cap))
      slot?.add(text)
    },
  })
  if (!found && wanted.size) throw new Stop(CORRUPT())
  return out
}

const general = (n: number) => String(Number(n.toPrecision(15))).replace('e', 'E')

/** Resolves, plans, reads and hands each consumer its cells. Never throws. */
async function runBatch(batch: Batch, context: Pick<ChartParseContext, 'pkg' | 'partPath' | 'profile' | 'workspace' | 'signal'>, rId: string, path: string, freeze: () => void): Promise<void> {
  const { workspace, profile, signal } = context
  const tag = `embedded-cells:${path}#${++serial}`
  const reader = new Reader(workspace, tag, profile, signal)
  lastStats = reader.stats
  let workbook: EmbeddedWorkbook | undefined
  const transient = () => { reader.release('text'); reader.release('load'); reader.release(`stream${CHART_CHUNK_TAG}`); reader.release(`stream${CHART_CROSSING_TAG}`) }
  const fail = (result: EmbeddedFailure) => {
    freeze()
    transient(); reader.release('data'); workbook?.release()
    for (const consumer of batch.consumers) consumer.settle(result)
  }
  try {
    const opened = await openEmbeddedWorkbook(context, rId)
    if (opened.kind !== 'workbook') { fail(opened); return }
    workbook = opened
    reader.live()
    // JSZip's directory structures, admitted before the load (bounded by the admitted inventory).
    reader.charge(LOAD_BASE_BYTES + LOAD_ENTRY_BYTES * opened.inventory.entries.length, 'load')
    let inner: OfficePackage
    // The load is not abortable: its working set is held (`hold`) until it actually settles, whatever the
    // signal does; the retained directory stays charged under `load` until the parts are read.
    reader.charge(LOAD_BASE_BYTES, 'loading')
    try { inner = await workspace.hold(OfficePackage.load(opened.bytes), `${tag}:loading`) } catch { throw new Stop(CORRUPT()) }
    reader.live()
    // From here the consumers are fixed: a later request starts its own read.
    freeze()
    const consumers = batch.consumers
    const wanted = new Set<string>()
    for (const consumer of consumers) for (const request of consumer.requests) if (request.range.definedName) wanted.add(request.range.definedName.toUpperCase())
    const book = await readWorkbook(inner, reader, wanted)
    const ownerSheet = book.sheets[0]?.name ?? ''
    const sheetIndex = (name: string) => book.sheets.findIndex(sheet => sheet.name.toLowerCase() === name.toLowerCase())
    const unavailable = (code: string, reason: string): EmbeddedResolution => ({ unavailable: { code, reason } })
    const external = (sheet: string | undefined) => sheet !== undefined && sheet.includes('[')
    /** A reference to a concrete range on an existing sheet (names: sheet scope first, then workbook; one level). */
    const resolve = (range: ChartRange): EmbeddedResolution => {
      let target = range
      if (range.definedName) {
        if (external(range.sheet)) return unavailable('chart-data-external-name', `Chart data name “${range.definedName}” refers to another workbook`)
        const scope = sheetIndex(range.sheet ?? ownerSheet), upper = range.definedName.toUpperCase()
        const candidates = book.names.filter(entry => entry.name.toUpperCase() === upper)
        const entry = candidates.find(candidate => scope >= 0 && candidate.local === scope) ?? candidates.find(candidate => candidate.local === undefined)
        if (!entry) return unavailable('chart-data-name-missing', `Chart data name “${range.definedName}” is not defined in the embedded workbook`)
        const formula = entry.source?.trim().replace(/^=/, '')
        if (formula !== undefined && formula.includes('[')) return unavailable('chart-data-external-name', `Chart data name “${range.definedName}” refers to another workbook`)
        const parsed = formula !== undefined ? parseChartRange(formula) : undefined
        if (!parsed || parsed.definedName) return unavailable('chart-data-name-unresolved', `Chart data name “${range.definedName}” is not a single worksheet range`)
        target = parsed
      } else if (external(range.sheet)) return unavailable('chart-data-external-reference', 'The chart data refers to another workbook')
      const index = sheetIndex(target.sheet ?? ownerSheet)
      if (index < 0) return unavailable('chart-data-sheet-missing', `The sheet “${(target.sheet ?? ownerSheet).slice(0, 64)}” is not in the embedded workbook`)
      return { range: { ...target, sheet: book.sheets[index].name } }
    }
    // Each consumer's resolutions and lines; a chart over the cell budget is refused (the others still read).
    const plans = new Map<string, Plan>()
    const outcomes = consumers.map(consumer => {
      const resolved = consumer.requests.map(request => resolve(request.range))
      const lines: ChartLine[] = []
      resolved.forEach((resolution, i) => {
        if (!('range' in resolution)) return
        const request = consumer.requests[i], range = resolution.range
        lines.push(request.single ? chartLine({ ...range, r1: range.r0, c1: range.c0 }, range.sheet!, true) : chartLine(range, range.sheet!, request.text, request.limit))
      })
      const cells = lines.reduce((sum, line) => sum + line.count, 0)
      if (cells > EMBEDDED_CELL_LIMIT) return { refused: failure('refused', 'embedded-cells-limit', 'This chart reads more cells than the viewer reads', EMBEDDED_CELL_LIMIT) }
      for (const line of lines) {
        const lower = line.sheet.toLowerCase()
        let plan = plans.get(lower)
        if (!plan) {
          const sheet = book.sheets[sheetIndex(line.sheet)]
          plan = { name: sheet.name, ...(sheet.rId ? { rId: sheet.rId } : {}), vertical: new Map(), horizontal: new Map(), minRow: Infinity, maxRow: -1, cells: new Map(), hiddenRows: new Set(), hiddenCols: [] }
          plans.set(lower, plan)
        }
        reader.charge(INTERVAL_BYTES, 'data')
        const end = line.start + line.count - 1, table = line.vertical ? plan.vertical : plan.horizontal
        const list = table.get(line.fixed) ?? []
        list.push([line.start, end]); table.set(line.fixed, list)
        plan.minRow = Math.min(plan.minRow, line.vertical ? line.start : line.fixed)
        plan.maxRow = Math.max(plan.maxRow, line.vertical ? end : line.fixed)
      }
      return { resolved }
    })
    let styles: Styles = { numFmts: new Map(), xfs: [] }
    if (plans.size) {
      const targets = await readWorkbookRels(inner, reader, new Set([...plans.values()].flatMap(plan => plan.rId ? [plan.rId] : [])))
      for (const plan of plans.values()) {
        const target = plan.rId ? targets.get(plan.rId) : undefined
        if (!target) throw new Stop(CORRUPT())
        plan.path = target
      }
      styles = await readStyles(inner, reader)
      for (const plan of plans.values()) { reader.live(); await readSheet(inner, reader, plan) }
      const strings = new Set<number>()
      for (const plan of plans.values()) for (const stored of plan.cells.values()) if (stored.sst !== undefined) strings.add(stored.sst)
      if (strings.size) {
        const table = await readSharedStrings(inner, reader, strings)
        for (const plan of plans.values()) for (const [at, stored] of plan.cells) {
          if (stored.sst === undefined) continue
          const text = table.get(stored.sst)
          if (text) plan.cells.set(at, { ...text, ...(stored.xf !== undefined ? { xf: stored.xf } : {}) })
          else plan.cells.delete(at)
        }
      }
    }
    // The embedding and its directory are no longer needed: refunded, and the document's next read may start.
    transient()
    workbook.release()
    workbook = undefined
    const epoch: '1900' | '1904' = book.date1904 ? '1904' : '1900'
    const formatOf = (xf: number | undefined): string | undefined => {
      if (xf === undefined) return undefined
      const id = styles.xfs[xf] ?? 0
      return sourceFormat({ numFmtId: id, ...(styles.numFmts.has(id) ? { formatCode: styles.numFmts.get(id)! } : {}) })
    }
    const lookup: ChartCellLookup = (sheet, row, col, text, allocation) => {
      const stored = plans.get(sheet.toLowerCase())?.cells.get(key(row, col))
      if (!stored) return undefined
      if (stored.unavailable) return { value: undefined, unavailable: true }
      if (stored.error !== undefined) return { value: stored.error, error: stored.error }
      const formatCode = formatOf(stored.xf)
      if (stored.text !== undefined) return { value: stored.text, number: stored.number ?? null, ...(stored.truncated ? { truncated: true } : {}), ...(formatCode ? { formatCode } : {}) }
      const out: ChartCell = { value: stored.n, number: stored.n!, ...(formatCode ? { formatCode } : {}) }
      // A label of a formatted number shows the workbook's formatted text; the formatter's buffers are admitted.
      if (text && formatCode && !/^general$/i.test(formatCode)) {
        const formatted = formatExtended(stored.n!, formatCode, { dateSystem: epoch, unicode: { version: 2, source: 'standalone-default' }, locale: 'en-US', timeZone: 'UTC', epochNowMs: 0 }, general, { maxOutput: CHART_LABEL_LIMIT + 1, allocate: bytes => allocation?.allocate(bytes), release: bytes => allocation?.release(bytes) })
        if (typeof formatted === 'string') out.text = formatted
        else if (formatted === NO_SECTION) out.text = ''
      }
      return out
    }
    const hidden: ChartHiddenLookup = (sheet, row, col) => {
      const plan = plans.get(sheet.toLowerCase())
      return Boolean(plan && (plan.hiddenRows.has(row) || within(plan.hiddenCols, col)))
    }
    let remaining = outcomes.filter(outcome => 'resolved' in outcome).length
    const finish = () => { if (--remaining <= 0) { reader.release('data'); plans.clear() } }
    if (!remaining) reader.release('data')
    consumers.forEach((consumer, i) => {
      const outcome = outcomes[i]
      if ('refused' in outcome) { consumer.settle(outcome.refused!); return }
      let done = false
      consumer.settle({ kind: 'cells', resolved: outcome.resolved!, lookup, hidden, epoch, ownerSheet, done: () => { if (!done) { done = true; finish() } } })
    })
  } catch (error) {
    fail(error instanceof Stop ? error.result : CORRUPT())
  }
}
