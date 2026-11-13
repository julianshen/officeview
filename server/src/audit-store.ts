/** Durable audit store (§9.3/§13.2/§13.4, Stage 2b G1): `bun:sqlite` in WAL
 * mode with `synchronous=FULL`, so a write that returned survives a crash.
 * Rows are write-once; per-unit delivery records are written before send;
 * token counters are per key generation and never rewind. */
import { Database } from 'bun:sqlite'
import { chmodSync, closeSync, existsSync, openSync, statSync } from 'node:fs'
import type { AuditRowInput } from './audit'
import type { A11yMode, CopyPolicy, PrintPolicy } from './hooks'
import type { MarkSpec } from './render'

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
/** One text-service answer to a sid (Stage 3a): a copy, or a counted a11y paragraph (Task 8). Never deduplicated. */
export interface CopyRecord { sid: string; kind: 'copy' | 'a11y'; paragraphs: number; fromUnit: number; toUnit: number; at: number }
/** One print job answered to a sid (Stage 3b): its page count, unit range and render scale. Never deduplicated. */
export interface PrintRecord { sid: string; pages: number; fromUnit: number; toUnit: number; scale: 1 | 2; at: number }
/** A resume that changed the sid's resolved text services (the open row keeps the first value, write-once). */
export interface ServicesRecord { sid: string; copy: CopyPolicy; textA11y: A11yMode; print: PrintPolicy; at: number }
export interface DetectionLog { at: number; osUser: string; investigator?: string; imageSha256: string; docRefs: string[]; unitIndex: number; answer: string; sid?: string }

const SCHEMA = `
CREATE TABLE IF NOT EXISTS rows (sid TEXT PRIMARY KEY, started_at INTEGER NOT NULL, key_generation INTEGER NOT NULL, doc_ref TEXT NOT NULL, msg40 TEXT, row TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS rows_doc ON rows (doc_ref);
CREATE INDEX IF NOT EXISTS rows_msg ON rows (key_generation, msg40);
CREATE TABLE IF NOT EXISTS deliveries (sid TEXT NOT NULL REFERENCES rows (sid), unit_index INTEGER NOT NULL, scale REAL NOT NULL, status TEXT NOT NULL, reason TEXT, ref_sha256 TEXT NOT NULL, at INTEGER NOT NULL, UNIQUE (sid, unit_index, scale, ref_sha256, status));
CREATE INDEX IF NOT EXISTS deliveries_unit ON deliveries (unit_index);
CREATE TABLE IF NOT EXISTS copies (sid TEXT NOT NULL REFERENCES rows (sid), kind TEXT NOT NULL, paragraphs INTEGER NOT NULL, from_unit INTEGER NOT NULL, to_unit INTEGER NOT NULL, at INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS copies_sid ON copies (sid);
CREATE TABLE IF NOT EXISTS prints (sid TEXT NOT NULL REFERENCES rows (sid), pages INTEGER NOT NULL, from_unit INTEGER NOT NULL, to_unit INTEGER NOT NULL, scale REAL NOT NULL, at INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS prints_sid ON prints (sid);
CREATE TABLE IF NOT EXISTS services (sid TEXT NOT NULL REFERENCES rows (sid), copy TEXT NOT NULL, text_a11y TEXT NOT NULL, print TEXT, at INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS services_sid ON services (sid);
CREATE TABLE IF NOT EXISTS counters (gen INTEGER PRIMARY KEY, next INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS detections (id INTEGER PRIMARY KEY AUTOINCREMENT, at INTEGER NOT NULL, log TEXT NOT NULL);
`

type DeliveryCols = { sid: string; unit_index: number; scale: number; status: 'l1' | 'unmarked'; reason: string | null; ref_sha256: string; at: number }
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
    if (!db.query<{ name: string }, []>('PRAGMA table_info(services)').all().some((c) => c.name === 'print')) db.exec('ALTER TABLE services ADD COLUMN print TEXT')
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
  /** Appends one copy/a11y record (repeats are kept); throws for a sid with no row. */
  putCopy(r: CopyRecord): void {
    this.writable()
    this.db.query('INSERT INTO copies (sid, kind, paragraphs, from_unit, to_unit, at) VALUES (?, ?, ?, ?, ?, ?)').run(r.sid, r.kind, r.paragraphs, r.fromUnit, r.toUnit, r.at)
  }
  copies(sid: string): CopyRecord[] {
    return this.db.query<{ sid: string; kind: CopyRecord['kind']; paragraphs: number; from_unit: number; to_unit: number; at: number }, [string]>('SELECT * FROM copies WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({ sid: r.sid, kind: r.kind, paragraphs: r.paragraphs, fromUnit: r.from_unit, toUnit: r.to_unit, at: r.at }))
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
  /** Appends one services change; throws for a sid with no row. */
  putServices(r: ServicesRecord): void {
    this.writable()
    this.db.query('INSERT INTO services (sid, copy, text_a11y, print, at) VALUES (?, ?, ?, ?, ?)').run(r.sid, JSON.stringify(r.copy), r.textA11y, JSON.stringify(r.print), r.at)
  }
  services(sid: string): ServicesRecord[] {
    return this.db.query<{ sid: string; copy: string; text_a11y: A11yMode; print: string | null; at: number }, [string]>('SELECT * FROM services WHERE sid = ? ORDER BY rowid').all(sid)
      .map((r) => ({ sid: r.sid, copy: JSON.parse(r.copy) as CopyPolicy, textA11y: r.text_a11y, print: r.print === null ? false : JSON.parse(r.print) as PrintPolicy, at: r.at }))
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
  /** Deletes rows older than the retention window with their deliveries, copies, prints and services changes (and so their messages); counters stay. */
  expire(nowSec: number): number {
    this.writable()
    const cutoff = nowSec - AUDIT_RETENTION_SEC
    return this.db.transaction(() => {
      for (const t of ['deliveries', 'copies', 'prints', 'services']) this.db.query(`DELETE FROM ${t} WHERE sid IN (SELECT sid FROM rows WHERE started_at <= ?)`).run(cutoff)
      return this.db.query('DELETE FROM rows WHERE started_at <= ?').run(cutoff).changes
    })()
  }
}
