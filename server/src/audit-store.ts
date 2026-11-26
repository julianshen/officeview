/** Durable audit store (§9.3/§13.2/§13.4, Stage 2b G1): `bun:sqlite` in WAL
 * mode with `synchronous=FULL`, so a write that returned survives a crash.
 * Rows are write-once; per-unit delivery records are written before send;
 * token counters are per key generation and never rewind. */
import { Database } from 'bun:sqlite'
import { chmodSync, closeSync, existsSync, openSync, statSync } from 'node:fs'
import type { AuditRowInput } from './audit'
import type { A11yMode, CopyPolicy, PrintPolicy } from './hooks'
import type { MarkSpec } from './render'
import type { SheetClock } from './sheet/clock'
import type { WireRange, WireSortSpec } from '../../src/stream/sheet-wire'
import type { WireSizes } from './sheet/view-state'
import type { ResolvedFilter } from './sheet/view-spec'

/** §9.3: rows are kept 90 days, then deleted. */
export const AUDIT_RETENTION_SEC = 90 * 24 * 3600

export interface StoredRow extends AuditRowInput {
  docRef: string
  contentPolicy: 'visible-only' | 'full'
  /** The visible mark as rendered, viewer text included (the trace tool's reference must match it). */
  visibleMark: MarkSpec
}
/** One `(unitIndex, scale)` delivered to a sid; `refSha256` is the canonical PNG the mark was applied to. */
export interface DeliveryRecord { sid: string; unitIndex: number; scale: number; status: 'l1' | 'unmarked'; reason?: string; refSha256: string; at: number }
/**
 * Stage 4a (S4 §6.5, F24): one sheet tile delivered to a sid. `sheetId` is the wire id (R12, the decimal index among
 * the visible sheets); `zoomStep` indexes `SHEET_ZOOM_STEPS`; `(col, row)` is the tile on the global 512 device-px grid.
 * `palette` is part of the key because it changes the tile's bytes. `refSha256` is the canonical PNG the mark was
 * applied to. Deduplicated like `DeliveryRecord`.
 */
export interface TileDeliveryRecord { sid: string; sheetId: string; viewGen: number; zoomStep: number; dpr: 1 | 2; palette: 'light' | 'dark'; col: number; row: number; status: 'l1' | 'unmarked'; reason?: string; refSha256: string; at: number }
/** The sheet view a sid's tiles of `(sheetId, viewGen)` were drawn under (R7; Stage 4b: the session's canonical sizes, `[]` for the identity view), with the session clock. Write-once. */
/** A sheet view's audited spec. 4a/4b wrote `{sort: null, filters: [], sizes}`, which reads back as written (F2). Stage 4c
 * (P12) writes the canonical view `{sort, sortFilter, filter, sizes}`, values rules by list `{range, col}` and positions,
 * typed operands in the clear. `sheetViewSpecOf` reads either form as the 4c one. */
export type SheetViewSpec =
  | { sort: null; filters: []; sizes: WireSizes }
  | { sort: WireSortSpec | null; sortFilter: ResolvedFilter | null; filter: ResolvedFilter | null; sizes: WireSizes }
export interface SheetViewRecord { sid: string; sheetId: string; viewGen: number; spec: SheetViewSpec; clock: SheetClock; at: number }
/** Either spec form as the 4c one: a 4b record has no sort, `sortFilter` or filter. */
export function sheetViewSpecOf(spec: SheetViewSpec): { sort: WireSortSpec | null; sortFilter: ResolvedFilter | null; filter: ResolvedFilter | null; sizes: WireSizes } {
  if ('filters' in spec) return { sort: null, sortFilter: null, filter: null, sizes: spec.sizes }
  return { sort: spec.sort, sortFilter: spec.sortFilter, filter: spec.filter, sizes: spec.sizes }
}
/**
 * Stage 4c (P12): one sort, filter, value-list or range-detect request of a sid that passed its shape check, written
 * before the answer. `rules` is `[[col, ruleKind | 'asc' | 'desc'], …]` and never holds an operand; `outcome` is `ok`
 * or the refusal code. Never deduplicated.
 */
export interface SheetOpRecord { sid: string; sheetId: string; kind: 'sort' | 'filter' | 'filter-values' | 'range-detect'; viewGen: number; range: WireRange; rules: Array<[number, string]>; outcome: string; at: number }
/** The only words a `sheet_ops` rule may carry: a sort order or a rule kind. */
export const SHEET_OP_RULE_KINDS: ReadonlySet<string> = new Set(['asc', 'desc', 'text', 'number', 'date', 'datePreset', 'compound', 'values', 'committed'])
/** One text-service answer to a sid (Stage 3a): a copy, or a counted a11y paragraph (Task 8). Never deduplicated.
 * Stage 4b (P8): a sheet copy also carries its wire `sheetId` and `rect`, the bounding source rect of its selection
 * (`fromUnit`/`toUnit` then hold the sheet index); absent on a units-mode record. No text is stored. */
export interface CopyRecord { sid: string; kind: 'copy' | 'a11y'; paragraphs: number; fromUnit: number; toUnit: number; sheetId?: string; rect?: { r0: number; c0: number; r1: number; c1: number }; at: number }
/** One print job answered to a sid (Stage 3b): its page count, unit range and render scale. Never deduplicated. */
export interface PrintRecord { sid: string; pages: number; fromUnit: number; toUnit: number; scale: 1 | 2; at: number }
/** One notes image answered to a sid (Stage 3c): the unit, scale and notes characters drawn (0 for a slide without notes). Never deduplicated (C4). */
export interface NotesRecord { sid: string; unitIndex: number; scale: 1 | 2; charsDrawn: number; at: number }
/** One notes image actually delivered Layer-1 marked (written after the mark, before the send): what trace
 * needs to re-derive it — the unit, scale and the canonical notes PNG the mark was applied to. Distinct from
 * `NotesRecord` (written before the draw, also for images that then failed) and from slide deliveries. */
export interface NotesMarkRecord { sid: string; unitIndex: number; scale: 1 | 2; refSha256: string; at: number }
/**
 * One show send answered to a sid (piece 4 Task 7): a slide bundle (`plan`), a keyframe track (`track`) or a Morph
 * pair (`morph`, unit = the `to` slide), with its asset count and how many of them shipped without Layer-1, and why
 * (`layer-capacity`: no carrier survived the layer's transparent cells; `forensic-off`: the session carries no
 * Layer-1). Written after the marks and before `*-start`. Never deduplicated (each is one charged send).
 */
export interface ShowDeliveryRecord { sid: string; unitIndex: number; kind: 'plan' | 'track' | 'morph'; assetCount: number; unmarked: number; reason?: 'layer-capacity' | 'forensic-off'; at: number }
/**
 * One media answer to a sid (piece 5, spec §7): a delivery (`media-start` and its chunks) or a refusal of a bound
 * mediaId. Rows are per session and viewer (through `sid`): two sessions that share one encode each get their own row.
 * `kind` is always `audio`. `cacheState`: how this session came by the file (`hit`, `joined`, `miss`), absent on a
 * refusal that never reached the cache. `sourceHash`, `ffmpegVersion` and `encodeParamsSha` are the encode key
 * (with the docId of the sid's row); absent when unknown (a record the plan shipped refused). No mark or geometry
 * columns: audio carries no mark (piece 6 adds them with video). Written before `media-start`. Never deduplicated.
 */
export interface MediaDeliveryRecord {
  sid: string; unitIndex: number; mediaId: string; kind: 'audio'; sourceHash?: string; outBytes?: number; durationMs?: number
  reason?: string; at: number; ffmpegVersion?: string; encodeParamsSha?: string; cacheState?: 'hit' | 'joined' | 'miss'
}
/** A resume that changed the sid's resolved text services (the open row keeps the first value, write-once). `media`
 * (piece 5): whether the sid's server serves media (`can.media`); a row stored before piece 5 reads as off. */
export interface ServicesRecord { sid: string; copy: CopyPolicy; textA11y: A11yMode; print: PrintPolicy; notes: boolean; media: boolean; at: number }
export interface DetectionLog { at: number; osUser: string; investigator?: string; imageSha256: string; docRefs: string[]; unitIndex: number; answer: string; sid?: string }

const SCHEMA = `
CREATE TABLE IF NOT EXISTS rows (sid TEXT PRIMARY KEY, started_at INTEGER NOT NULL, key_generation INTEGER NOT NULL, doc_ref TEXT NOT NULL, msg40 TEXT, row TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS rows_doc ON rows (doc_ref);
CREATE INDEX IF NOT EXISTS rows_msg ON rows (key_generation, msg40);
CREATE TABLE IF NOT EXISTS deliveries (sid TEXT NOT NULL REFERENCES rows (sid), unit_index INTEGER NOT NULL, scale REAL NOT NULL, status TEXT NOT NULL, reason TEXT, ref_sha256 TEXT NOT NULL, at INTEGER NOT NULL, UNIQUE (sid, unit_index, scale, ref_sha256, status));
CREATE INDEX IF NOT EXISTS deliveries_unit ON deliveries (unit_index);
CREATE TABLE IF NOT EXISTS tile_deliveries (sid TEXT NOT NULL REFERENCES rows (sid), sheet_id TEXT NOT NULL, view_gen INTEGER NOT NULL, zoom_step INTEGER NOT NULL, dpr INTEGER NOT NULL, palette TEXT NOT NULL, col INTEGER NOT NULL, row INTEGER NOT NULL, status TEXT NOT NULL, reason TEXT, ref_sha256 TEXT NOT NULL, at INTEGER NOT NULL, UNIQUE (sid, sheet_id, view_gen, zoom_step, dpr, palette, col, row, ref_sha256, status));
CREATE INDEX IF NOT EXISTS tile_deliveries_tile ON tile_deliveries (sheet_id, zoom_step, dpr, palette, col, row);
CREATE TABLE IF NOT EXISTS sheet_views (sid TEXT NOT NULL REFERENCES rows (sid), sheet_id TEXT NOT NULL, view_gen INTEGER NOT NULL, spec TEXT NOT NULL, clock TEXT NOT NULL, at INTEGER NOT NULL, UNIQUE (sid, sheet_id, view_gen));
CREATE TABLE IF NOT EXISTS sheet_ops (sid TEXT NOT NULL REFERENCES rows (sid), sheet_id TEXT NOT NULL, kind TEXT NOT NULL, view_gen INTEGER NOT NULL, range TEXT NOT NULL, rules TEXT NOT NULL, outcome TEXT NOT NULL, at INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS sheet_ops_sid ON sheet_ops (sid);
CREATE TABLE IF NOT EXISTS copies (sid TEXT NOT NULL REFERENCES rows (sid), kind TEXT NOT NULL, paragraphs INTEGER NOT NULL, from_unit INTEGER NOT NULL, to_unit INTEGER NOT NULL, at INTEGER NOT NULL, sheet_id TEXT, rect TEXT);
CREATE INDEX IF NOT EXISTS copies_sid ON copies (sid);
CREATE TABLE IF NOT EXISTS prints (sid TEXT NOT NULL REFERENCES rows (sid), pages INTEGER NOT NULL, from_unit INTEGER NOT NULL, to_unit INTEGER NOT NULL, scale REAL NOT NULL, at INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS prints_sid ON prints (sid);
CREATE TABLE IF NOT EXISTS notes_deliveries (sid TEXT NOT NULL REFERENCES rows (sid), unit_index INTEGER NOT NULL, scale REAL NOT NULL, chars_drawn INTEGER NOT NULL, at INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS notes_deliveries_sid ON notes_deliveries (sid);
CREATE TABLE IF NOT EXISTS notes_marks (sid TEXT NOT NULL REFERENCES rows (sid), unit_index INTEGER NOT NULL, scale REAL NOT NULL, ref_sha256 TEXT NOT NULL, at INTEGER NOT NULL, UNIQUE (sid, unit_index, scale, ref_sha256));
CREATE INDEX IF NOT EXISTS notes_marks_unit ON notes_marks (unit_index);
CREATE TABLE IF NOT EXISTS show_deliveries (sid TEXT NOT NULL REFERENCES rows (sid), unit_index INTEGER NOT NULL, kind TEXT NOT NULL, asset_count INTEGER NOT NULL, unmarked INTEGER NOT NULL, reason TEXT, at INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS show_deliveries_sid ON show_deliveries (sid);
CREATE TABLE IF NOT EXISTS media_deliveries (sid TEXT NOT NULL REFERENCES rows (sid), unit_index INTEGER NOT NULL, media_id TEXT NOT NULL, kind TEXT NOT NULL, source_hash TEXT, out_bytes INTEGER, duration_ms INTEGER, reason TEXT, at INTEGER NOT NULL, ffmpeg_version TEXT, encode_params_sha TEXT, cache_state TEXT);
CREATE INDEX IF NOT EXISTS media_deliveries_sid ON media_deliveries (sid);
CREATE INDEX IF NOT EXISTS media_deliveries_media ON media_deliveries (media_id);
CREATE TABLE IF NOT EXISTS services (sid TEXT NOT NULL REFERENCES rows (sid), copy TEXT NOT NULL, text_a11y TEXT NOT NULL, print TEXT, notes INTEGER, media TEXT, at INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS services_sid ON services (sid);
CREATE TABLE IF NOT EXISTS counters (gen INTEGER PRIMARY KEY, next INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS detections (id INTEGER PRIMARY KEY AUTOINCREMENT, at INTEGER NOT NULL, log TEXT NOT NULL);
`

type DeliveryCols = { sid: string; unit_index: number; scale: number; status: 'l1' | 'unmarked'; reason: string | null; ref_sha256: string; at: number }
type MediaCols = { sid: string; unit_index: number; media_id: string; kind: 'audio'; source_hash: string | null; out_bytes: number | null; duration_ms: number | null; reason: string | null; at: number; ffmpeg_version: string | null; encode_params_sha: string | null; cache_state: MediaDeliveryRecord['cacheState'] | null }
const mediaDelivery = (r: MediaCols): MediaDeliveryRecord => ({
  sid: r.sid, unitIndex: r.unit_index, mediaId: r.media_id, kind: r.kind,
  ...(r.source_hash !== null ? { sourceHash: r.source_hash } : {}), ...(r.out_bytes !== null ? { outBytes: r.out_bytes } : {}), ...(r.duration_ms !== null ? { durationMs: r.duration_ms } : {}),
  ...(r.reason !== null ? { reason: r.reason } : {}), at: r.at,
  ...(r.ffmpeg_version !== null ? { ffmpegVersion: r.ffmpeg_version } : {}), ...(r.encode_params_sha !== null ? { encodeParamsSha: r.encode_params_sha } : {}), ...(r.cache_state !== null ? { cacheState: r.cache_state } : {}),
})
type TileCols = { sid: string; sheet_id: string; view_gen: number; zoom_step: number; dpr: 1 | 2; palette: 'light' | 'dark'; col: number; row: number; status: 'l1' | 'unmarked'; reason: string | null; ref_sha256: string; at: number }
const tileDelivery = (d: TileCols): TileDeliveryRecord => ({ sid: d.sid, sheetId: d.sheet_id, viewGen: d.view_gen, zoomStep: d.zoom_step, dpr: d.dpr, palette: d.palette, col: d.col, row: d.row, status: d.status, ...(d.reason !== null ? { reason: d.reason } : {}), refSha256: d.ref_sha256, at: d.at })
const delivery = (d: DeliveryCols): DeliveryRecord => ({ sid: d.sid, unitIndex: d.unit_index, scale: d.scale, status: d.status, ...(d.reason !== null ? { reason: d.reason } : {}), refSha256: d.ref_sha256, at: d.at })

/** Clears group/other permission bits (umask-independent); never adds a bit. */
function tighten(file: string): void {
  if (!existsSync(file)) return
  const mode = statSync(file).mode & 0o777
  if (mode & 0o077) chmodSync(file, mode & 0o700)
}

export class AuditStore {
  private constructor(private readonly db: Database, private readonly readOnly: boolean) {}

  /** Opens the store at `path` (created 0600; db, -wal and -shm tightened to 0600 on every open); `':memory:'` for demo and tests. */
  static open(path: string): AuditStore { return new AuditStore(AuditStore.connect(path), false) }
  /** For the trace tool: rows, deliveries and counters are read-only; only `putDetection` writes. */
  static openReadOnlyForTrace(path: string): AuditStore { return new AuditStore(AuditStore.connect(path), true) }

  private static connect(path: string): Database {
    const disk = path !== ':memory:'
    if (disk) {
      if (!existsSync(path)) closeSync(openSync(path, 'a', 0o600))
      tighten(path) // before open: SQLite's unix VFS creates -wal/-shm with the db file's mode
    }
    const db = new Database(path, { create: true, strict: true })
    db.exec('PRAGMA journal_mode = WAL')
    db.exec('PRAGMA synchronous = FULL')
    db.exec('PRAGMA foreign_keys = ON')
    db.exec(SCHEMA)
    // A store created before Stage 3b has no `services.print` column (CREATE IF NOT EXISTS leaves it alone); add it, null = not recorded.
    const servicesColumns = db.query<{ name: string }, []>('PRAGMA table_info(services)').all().map((c) => c.name)
    if (!servicesColumns.includes('print')) db.exec('ALTER TABLE services ADD COLUMN print TEXT')
    // One created before Stage 3c has no `services.notes` column; add it, null = not recorded (reads as notes off).
    if (!servicesColumns.includes('notes')) db.exec('ALTER TABLE services ADD COLUMN notes INTEGER')
    // One created before piece 5 has no `services.media` column; add it, null = not recorded (reads as media off).
    if (!servicesColumns.includes('media')) db.exec('ALTER TABLE services ADD COLUMN media TEXT')
    // One created before Stage 4b has no `copies.sheet_id`/`copies.rect`; add them, null = a units-mode copy.
    const copiesColumns = db.query<{ name: string }, []>('PRAGMA table_info(copies)').all().map((c) => c.name)
    if (!copiesColumns.includes('sheet_id')) db.exec('ALTER TABLE copies ADD COLUMN sheet_id TEXT')
    if (!copiesColumns.includes('rect')) db.exec('ALTER TABLE copies ADD COLUMN rect TEXT')
    // The sidecars now exist (WAL mode, schema written). Do not rely on the VFS
    // copying the db's mode, nor on the umask: pre-existing or recreated
    // sidecars hold session rows (IPs, viewer ids) until checkpoint.
    if (disk) for (const s of ['-wal', '-shm']) tighten(path + s)
    return db
  }
  private writable(): void { if (this.readOnly) throw new Error('audit store is read-only (trace)') }

  close(): void { this.db.close() }

  /** Write-once: throws on a duplicate sid. */
  putRow(sid: string, row: StoredRow): void {
    this.writable()
    this.db.query('INSERT INTO rows (sid, started_at, key_generation, doc_ref, msg40, row) VALUES (?, ?, ?, ?, ?, ?)')
      .run(sid, row.startedAt, row.keyGeneration, row.docRef, row.forensic?.msg40 ?? null, JSON.stringify(row))
  }
  getRow(sid: string): StoredRow | undefined {
    const r = this.db.query<{ row: string }, [string]>('SELECT row FROM rows WHERE sid = ?').get(sid)
    return r ? JSON.parse(r.row) as StoredRow : undefined
  }
  /** Idempotent on `(sid, unitIndex, scale, refSha256, status)`; throws for a sid with no row. */
  putDelivery(d: DeliveryRecord): void {
    this.writable()
    this.db.query('INSERT OR IGNORE INTO deliveries (sid, unit_index, scale, status, reason, ref_sha256, at) VALUES (?, ?, ?, ?, ?, ?, ?)')
      .run(d.sid, d.unitIndex, d.scale, d.status, d.reason ?? null, d.refSha256, d.at)
  }
  deliveries(sid: string): DeliveryRecord[] {
    return this.db.query<DeliveryCols, [string]>('SELECT * FROM deliveries WHERE sid = ? ORDER BY rowid').all(sid).map(delivery)
  }
  /** Every delivery of `docRef`'s unit `unitIndex`, any status, with its row. */
  candidates(docRef: string, unitIndex: number): Array<{ sid: string; row: StoredRow; delivery: DeliveryRecord }> {
    return this.db.query<DeliveryCols & { row: string }, [string, number]>('SELECT d.*, r.row FROM deliveries d JOIN rows r ON r.sid = d.sid WHERE r.doc_ref = ? AND d.unit_index = ? ORDER BY d.rowid')
      .all(docRef, unitIndex).map((x) => ({ sid: x.sid, row: JSON.parse(x.row) as StoredRow, delivery: delivery(x) }))
  }
  /** Idempotent on `(sid, sheetId, viewGen, zoomStep, dpr, palette, col, row, refSha256, status)`; throws for a sid with no row. */
  putTileDelivery(d: TileDeliveryRecord): void {
    this.writable()
    this.db.query('INSERT OR IGNORE INTO tile_deliveries (sid, sheet_id, view_gen, zoom_step, dpr, palette, col, row, status, reason, ref_sha256, at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)')
      .run(d.sid, d.sheetId, d.viewGen, d.zoomStep, d.dpr, d.palette, d.col, d.row, d.status, d.reason ?? null, d.refSha256, d.at)
  }
  tileDeliveries(sid: string): TileDeliveryRecord[] {
    return this.db.query<TileCols, [string]>('SELECT * FROM tile_deliveries WHERE sid = ? ORDER BY rowid').all(sid).map(tileDelivery)
  }
  /** Write-once on `(sid, sheetId, viewGen)` (a repeat is ignored); throws for a sid with no row. */
  putSheetView(r: SheetViewRecord): void {
    this.writable()
    this.db.query('INSERT OR IGNORE INTO sheet_views (sid, sheet_id, view_gen, spec, clock, at) VALUES (?, ?, ?, ?, ?, ?)')
      .run(r.sid, r.sheetId, r.viewGen, JSON.stringify(r.spec), JSON.stringify(r.clock), r.at)
  }
  sheetView(sid: string, sheetId: string, viewGen: number): SheetViewRecord | undefined {
    const r = this.db.query<{ sid: string; sheet_id: string; view_gen: number; spec: string; clock: string; at: number }, [string, string, number]>('SELECT * FROM sheet_views WHERE sid = ? AND sheet_id = ? AND view_gen = ?').get(sid, sheetId, viewGen)
    return r ? { sid: r.sid, sheetId: r.sheet_id, viewGen: r.view_gen, spec: JSON.parse(r.spec) as SheetViewRecord['spec'], clock: JSON.parse(r.clock) as SheetClock, at: r.at } : undefined
  }
  /** Appends one sort/filter op record (repeats are kept); throws for a rule word outside `SHEET_OP_RULE_KINDS` and for
   * a sid with no row. */
  putSheetOp(r: SheetOpRecord): void {
    this.writable()
    for (const [col, kind] of r.rules) if (!Number.isInteger(col) || !SHEET_OP_RULE_KINDS.has(kind)) throw new Error(`sheet_ops: not a rule kind: ${JSON.stringify(kind).slice(0, 16)}`)
    const range = { r0: r.range.r0, r1: r.range.r1, c0: r.range.c0, c1: r.range.c1, header: r.range.header, totals: r.range.totals }
    this.db.query('INSERT INTO sheet_ops (sid, sheet_id, kind, view_gen, range, rules, outcome, at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)')
      .run(r.sid, r.sheetId, r.kind, r.viewGen, JSON.stringify(range), JSON.stringify(r.rules), r.outcome, r.at)
  }
  sheetOps(sid: string): SheetOpRecord[] {
    return this.db.query<{ sid: string; sheet_id: string; kind: SheetOpRecord['kind']; view_gen: number; range: string; rules: string; outcome: string; at: number }, [string]>('SELECT * FROM sheet_ops WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({ sid: r.sid, sheetId: r.sheet_id, kind: r.kind, viewGen: r.view_gen, range: JSON.parse(r.range) as WireRange, rules: JSON.parse(r.rules) as Array<[number, string]>, outcome: r.outcome, at: r.at }))
  }
  /** Every delivery, any status, of `docRef`'s sheet `sheetId` at `(zoomStep, dpr, palette)` whose tile lies in the
   * inclusive `window` (`tileRangeFor`'s `c0..c1` × `r0..r1`), with its row. `palette` defaults to `'light'` (F24). */
  tileCandidates(docRef: string, sheetId: string, zoomStep: number, dpr: 1 | 2, window: { c0: number; c1: number; r0: number; r1: number }, palette: 'light' | 'dark' = 'light'): Array<{ sid: string; row: StoredRow; delivery: TileDeliveryRecord }> {
    return this.db.query<TileCols & { rowjson: string }, [string, string, number, number, string, number, number, number, number]>(
      'SELECT d.*, r.row AS rowjson FROM tile_deliveries d JOIN rows r ON r.sid = d.sid WHERE r.doc_ref = ? AND d.sheet_id = ? AND d.zoom_step = ? AND d.dpr = ? AND d.palette = ? AND d.col BETWEEN ? AND ? AND d.row BETWEEN ? AND ? ORDER BY d.rowid')
      .all(docRef, sheetId, zoomStep, dpr, palette, window.c0, window.c1, window.r0, window.r1)
      .map((x) => ({ sid: x.sid, row: JSON.parse(x.rowjson) as StoredRow, delivery: tileDelivery(x) }))
  }
  /** Appends one copy/a11y record (repeats are kept); throws for a sid with no row. */
  putCopy(r: CopyRecord): void {
    this.writable()
    const rect = r.rect ? JSON.stringify({ r0: r.rect.r0, c0: r.rect.c0, r1: r.rect.r1, c1: r.rect.c1 }) : null
    this.db.query('INSERT INTO copies (sid, kind, paragraphs, from_unit, to_unit, at, sheet_id, rect) VALUES (?, ?, ?, ?, ?, ?, ?, ?)').run(r.sid, r.kind, r.paragraphs, r.fromUnit, r.toUnit, r.at, r.sheetId ?? null, rect)
  }
  copies(sid: string): CopyRecord[] {
    return this.db.query<{ sid: string; kind: CopyRecord['kind']; paragraphs: number; from_unit: number; to_unit: number; at: number; sheet_id: string | null; rect: string | null }, [string]>('SELECT * FROM copies WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({
        sid: r.sid, kind: r.kind, paragraphs: r.paragraphs, fromUnit: r.from_unit, toUnit: r.to_unit,
        ...(r.sheet_id !== null ? { sheetId: r.sheet_id } : {}), ...(r.rect !== null ? { rect: JSON.parse(r.rect) as NonNullable<CopyRecord['rect']> } : {}),
        at: r.at,
      }))
  }
  /** Appends one print job record (repeats are kept); throws for a sid with no row. */
  putPrint(r: PrintRecord): void {
    this.writable()
    this.db.query('INSERT INTO prints (sid, pages, from_unit, to_unit, scale, at) VALUES (?, ?, ?, ?, ?, ?)').run(r.sid, r.pages, r.fromUnit, r.toUnit, r.scale, r.at)
  }
  prints(sid: string): PrintRecord[] {
    return this.db.query<{ sid: string; pages: number; from_unit: number; to_unit: number; scale: 1 | 2; at: number }, [string]>('SELECT * FROM prints WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({ sid: r.sid, pages: r.pages, fromUnit: r.from_unit, toUnit: r.to_unit, scale: r.scale, at: r.at }))
  }
  /** Appends one notes image record (repeats are kept); throws for a sid with no row. */
  putNotes(r: NotesRecord): void {
    this.writable()
    this.db.query('INSERT INTO notes_deliveries (sid, unit_index, scale, chars_drawn, at) VALUES (?, ?, ?, ?, ?)').run(r.sid, r.unitIndex, r.scale, r.charsDrawn, r.at)
  }
  notesDeliveries(sid: string): NotesRecord[] {
    return this.db.query<{ sid: string; unit_index: number; scale: 1 | 2; chars_drawn: number; at: number }, [string]>('SELECT * FROM notes_deliveries WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({ sid: r.sid, unitIndex: r.unit_index, scale: r.scale, charsDrawn: r.chars_drawn, at: r.at }))
  }
  /** Idempotent on `(sid, unitIndex, scale, refSha256)`; throws for a sid with no row. */
  putNotesMark(r: NotesMarkRecord): void {
    this.writable()
    this.db.query('INSERT OR IGNORE INTO notes_marks (sid, unit_index, scale, ref_sha256, at) VALUES (?, ?, ?, ?, ?)').run(r.sid, r.unitIndex, r.scale, r.refSha256, r.at)
  }
  notesMarks(sid: string): NotesMarkRecord[] {
    return this.db.query<{ sid: string; unit_index: number; scale: 1 | 2; ref_sha256: string; at: number }, [string]>('SELECT * FROM notes_marks WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({ sid: r.sid, unitIndex: r.unit_index, scale: r.scale, refSha256: r.ref_sha256, at: r.at }))
  }
  /** Every marked notes delivery of `docRef`'s unit `unitIndex`, with its row, as a Layer-1 delivery (trace's notes target). */
  notesCandidates(docRef: string, unitIndex: number): Array<{ sid: string; row: StoredRow; delivery: DeliveryRecord }> {
    return this.db.query<{ sid: string; unit_index: number; scale: number; ref_sha256: string; at: number; row: string }, [string, number]>('SELECT n.*, r.row FROM notes_marks n JOIN rows r ON r.sid = n.sid WHERE r.doc_ref = ? AND n.unit_index = ? ORDER BY n.rowid')
      .all(docRef, unitIndex).map((x) => ({ sid: x.sid, row: JSON.parse(x.row) as StoredRow, delivery: { sid: x.sid, unitIndex: x.unit_index, scale: x.scale, status: 'l1' as const, refSha256: x.ref_sha256, at: x.at } }))
  }
  /** Appends one show send record (repeats are kept); throws for a sid with no row. */
  putShow(r: ShowDeliveryRecord): void {
    this.writable()
    this.db.query('INSERT INTO show_deliveries (sid, unit_index, kind, asset_count, unmarked, reason, at) VALUES (?, ?, ?, ?, ?, ?, ?)').run(r.sid, r.unitIndex, r.kind, r.assetCount, r.unmarked, r.reason ?? null, r.at)
  }
  showDeliveries(sid: string): ShowDeliveryRecord[] {
    return this.db.query<{ sid: string; unit_index: number; kind: ShowDeliveryRecord['kind']; asset_count: number; unmarked: number; reason: ShowDeliveryRecord['reason'] | null; at: number }, [string]>('SELECT * FROM show_deliveries WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({ sid: r.sid, unitIndex: r.unit_index, kind: r.kind, assetCount: r.asset_count, unmarked: r.unmarked, ...(r.reason !== null ? { reason: r.reason } : {}), at: r.at }))
  }
  /** Appends one media delivery or refusal (repeats are kept); throws for a sid with no row. */
  putMedia(r: MediaDeliveryRecord): void {
    this.writable()
    this.db.query('INSERT INTO media_deliveries (sid, unit_index, media_id, kind, source_hash, out_bytes, duration_ms, reason, at, ffmpeg_version, encode_params_sha, cache_state) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)')
      .run(r.sid, r.unitIndex, r.mediaId, r.kind, r.sourceHash ?? null, r.outBytes ?? null, r.durationMs ?? null, r.reason ?? null, r.at, r.ffmpegVersion ?? null, r.encodeParamsSha ?? null, r.cacheState ?? null)
  }
  mediaDeliveries(sid: string): MediaDeliveryRecord[] {
    return this.db.query<MediaCols, [string]>('SELECT * FROM media_deliveries WHERE sid = ? ORDER BY rowid').all(sid).map(mediaDelivery)
  }
  /** The latest row for `mediaId` (a session's ids are unique to it), if any. */
  mediaDeliveryById(mediaId: string): MediaDeliveryRecord | undefined {
    const r = this.db.query<MediaCols, [string]>('SELECT * FROM media_deliveries WHERE media_id = ? ORDER BY rowid DESC LIMIT 1').get(mediaId)
    return r ? mediaDelivery(r) : undefined
  }
  /** A table's column names, in order (schema checks). */
  columnsOf(table: string): string[] {
    return this.db.query<{ name: string }, []>(`PRAGMA table_info(${JSON.stringify(table)})`).all().map((c) => c.name)
  }
  /** Appends one services change; throws for a sid with no row. */
  putServices(r: ServicesRecord): void {
    this.writable()
    this.db.query('INSERT INTO services (sid, copy, text_a11y, print, notes, media, at) VALUES (?, ?, ?, ?, ?, ?, ?)').run(r.sid, JSON.stringify(r.copy), r.textA11y, JSON.stringify(r.print), r.notes ? 1 : 0, r.media ? 'on' : 'off', r.at)
  }
  services(sid: string): ServicesRecord[] {
    return this.db.query<{ sid: string; copy: string; text_a11y: A11yMode; print: string | null; notes: number | null; media: string | null; at: number }, [string]>('SELECT * FROM services WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({ sid: r.sid, copy: JSON.parse(r.copy) as CopyPolicy, textA11y: r.text_a11y, print: r.print === null ? false : JSON.parse(r.print) as PrintPolicy, notes: r.notes === 1, media: r.media === 'on', at: r.at }))
  }
  /** The token `takeToken(gen)` would return. */
  nextToken(gen: number): number {
    return this.db.query<{ next: number }, [number]>('SELECT next FROM counters WHERE gen = ?').get(gen)?.next ?? 0
  }
  /** Durably increments `gen`'s counter and returns the previous value. */
  takeToken(gen: number): number {
    this.writable()
    const r = this.db.query<{ next: number }, [number]>('INSERT INTO counters (gen, next) VALUES (?, 1) ON CONFLICT (gen) DO UPDATE SET next = next + 1 RETURNING next').get(gen)!
    return r.next - 1
  }
  /** TEST-ONLY: raise `gen`'s counter to at least `next` (never lowers it). */
  raiseToken(gen: number, next: number): void {
    this.writable()
    this.db.query('INSERT INTO counters (gen, next) VALUES (?, ?) ON CONFLICT (gen) DO UPDATE SET next = max(next, excluded.next)').run(gen, next)
  }
  /** Live keyed messages (hex) of generation `gen`. */
  liveMessages(gen: number): string[] {
    return this.db.query<{ msg40: string }, [number]>('SELECT msg40 FROM rows WHERE key_generation = ? AND msg40 IS NOT NULL').all(gen).map((r) => r.msg40)
  }
  putDetection(r: DetectionLog): void {
    this.db.query('INSERT INTO detections (at, log) VALUES (?, ?)').run(r.at, JSON.stringify(r))
  }
  detections(): DetectionLog[] {
    return this.db.query<{ log: string }, []>('SELECT log FROM detections ORDER BY id').all().map((r) => JSON.parse(r.log) as DetectionLog)
  }
  /** Deletes rows older than the retention window with their deliveries, tile deliveries, sheet views, sheet ops, copies, prints, notes images (audited and marked), show sends, media deliveries and services changes (and so their messages); counters stay. */
  expire(nowSec: number): number {
    this.writable()
    const cutoff = nowSec - AUDIT_RETENTION_SEC
    return this.db.transaction(() => {
      for (const t of ['deliveries', 'tile_deliveries', 'sheet_views', 'sheet_ops', 'copies', 'prints', 'notes_deliveries', 'notes_marks', 'show_deliveries', 'media_deliveries', 'services']) this.db.query(`DELETE FROM ${t} WHERE sid IN (SELECT sid FROM rows WHERE started_at <= ?)`).run(cutoff)
      return this.db.query('DELETE FROM rows WHERE started_at <= ?').run(cutoff).changes
    })()
  }
}
