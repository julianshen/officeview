/** Stage 2b Task 4: the durable audit store (§9.3/§13.2), delivery records before send. */
import { afterAll, afterEach, describe, expect, it } from 'bun:test'
import { Database } from 'bun:sqlite'
import { createHash } from 'node:crypto'
import { chmodSync, closeSync, existsSync, mkdtempSync, openSync, rmSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { AUDIT_RETENTION_SEC, AuditTable, msgHex } from '../src/audit'
import { AuditStore, type CopyRecord, type DeliveryRecord, type PrintRecord, type StoredRow } from '../src/audit-store'
import { messageFor } from '../src/mark/message'
import { Client, corpus, disposePool, start } from './harness'

afterAll(() => disposePool())
const dirs: string[] = []
afterEach(() => { for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true }) })
const dbPath = () => { const d = mkdtempSync(join(tmpdir(), 'ov-audit-')); dirs.push(d); return join(d, 'audit.db') }

const kMsg = Uint8Array.from({ length: 32 }, (_, i) => 200 - i)
const row = (o: Partial<StoredRow> = {}): StoredRow => ({
  ip: '1.2.3.4', startedAt: 1000, keyGeneration: 1, policyId: 'visible-only@1', markSettings: { visible: 'viewer', forensicRequested: true }, relaxed: false, forensic: null,
  docRef: 'doc-a', contentPolicy: 'visible-only', visibleMark: { kind: 'viewer', text: 'alice' }, services: { copy: false, textA11y: 'off', print: false }, ...o,
})
const sha = (b: Uint8Array) => createHash('sha256').update(b).digest('hex')

describe('AuditStore', () => {
  it('restart keeps rows, deliveries, counters', () => {
    const path = dbPath()
    const t = new AuditTable({ store: AuditStore.open(path) })
    const issued = [t.issue(kMsg, 1, () => 1), t.issue(kMsg, 1, () => 2), t.issue(kMsg, 1, () => 3)]
    const rows = issued.slice(0, 2).map((f, i) => row({ viewerId: `v${i}`, forensic: { codec: 'l1-bch512x40@1', token24: f.token24, salt8: f.salt8, msg40: msgHex(f.msg40) } }))
    t.record('s0', rows[0]!); t.record('s1', rows[1]!)
    t.recordDelivery({ sid: 's0', unitIndex: 0, scale: 1, status: 'l1', refSha256: 'aa' })
    t.recordDelivery({ sid: 's1', unitIndex: 2, scale: 2, status: 'unmarked', reason: 'capacity', refSha256: 'bb' })
    const before = { r0: t.get('s0'), r1: t.get('s1') }
    const store = (t as unknown as { store: AuditStore }).store
    const d0 = store.deliveries('s0'), d1 = store.deliveries('s1')
    store.close()

    const s2 = AuditStore.open(path), t2 = new AuditTable({ store: s2 })
    expect(s2.getRow('s0')).toEqual(rows[0]!); expect(s2.getRow('s1')).toEqual(rows[1]!)
    expect(t2.get('s0')).toEqual(before.r0!); expect(t2.get('s1')).toEqual(before.r1!)
    expect(t2.get('s1')).toMatchObject({ forensicLayers: [], delivered: { marked: 0, unmarked: 1 } })
    expect(s2.deliveries('s0')).toEqual(d0); expect(s2.deliveries('s1')).toEqual(d1)
    expect(d0).toEqual([{ sid: 's0', unitIndex: 0, scale: 1, status: 'l1', refSha256: 'aa', at: expect.any(Number) }])
    expect(d1[0]).toMatchObject({ status: 'unmarked', reason: 'capacity' })
    expect(t2.nextToken(1)).toBe(3); expect(t2.nextToken(2)).toBe(0)
    // A reloaded msg40 is live: token 3's salt-0 message collides only if live, so seed one.
    t2.record('s2', row({ forensic: { codec: 'l1-bch512x40@1', token24: 3, salt8: 0, msg40: msgHex(messageFor(kMsg, 1, 3, 0)) } }))
    s2.close()
    const t3 = new AuditTable({ store: AuditStore.open(path) })
    const draws = [0, 0, 7]
    const f = t3.issue(kMsg, 1, () => draws.shift()!)
    expect(f.token24).toBe(3)
    expect(f.salt8).toBe(7)
    ;(t3 as unknown as { store: AuditStore }).store.close()
  })

  it('write-once survives reopen', () => {
    const path = dbPath()
    const s = AuditStore.open(path); s.putRow('s1', row()); s.close()
    const t = new AuditTable({ store: AuditStore.open(path) })
    expect(() => t.record('s1', row())).toThrow(/write-once/)
    ;(t as unknown as { store: AuditStore }).store.close()
  })

  it('expiry deletes rows, deliveries and messages but not counters', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    const f = t.issue(kMsg, 1)
    t.record('old', row({ startedAt: 0, forensic: { codec: 'l1-bch512x40@1', token24: f.token24, salt8: f.salt8, msg40: msgHex(f.msg40) } }))
    t.record('new', row({ startedAt: 100 }))
    t.recordDelivery({ sid: 'old', unitIndex: 0, scale: 1, status: 'l1', refSha256: 'aa' })
    expect(s.liveMessages(1)).toEqual([msgHex(f.msg40)])
    expect(t.expire(AUDIT_RETENTION_SEC)).toBe(1)
    expect(s.getRow('old')).toBeUndefined(); expect(s.getRow('new')).toBeDefined()
    expect(s.deliveries('old')).toEqual([])
    expect(s.candidates('doc-a', 0)).toEqual([])
    expect(s.liveMessages(1)).toEqual([])
    expect(t.nextToken(1)).toBe(1)
    s.close()
  })

  it('candidates returns every delivery of a docRef unit, all statuses', () => {
    const s = AuditStore.open(':memory:')
    s.putRow('a', row()); s.putRow('b', row({ visibleMark: { kind: 'off' } })); s.putRow('c', row({ docRef: 'doc-b' }))
    const d = (sid: string, unitIndex: number, status: DeliveryRecord['status']): DeliveryRecord => ({ sid, unitIndex, scale: 1, status, ...(status === 'unmarked' ? { reason: 'capacity' } : {}), refSha256: 'r', at: 5 })
    for (const x of [d('a', 0, 'l1'), d('a', 1, 'l1'), d('b', 0, 'unmarked'), d('c', 0, 'l1')]) s.putDelivery(x)
    s.putDelivery(d('a', 0, 'l1')) // a re-request of the same unit: one record
    const c = s.candidates('doc-a', 0)
    expect(c.map((x) => [x.sid, x.delivery.status])).toEqual([['a', 'l1'], ['b', 'unmarked']])
    expect(c[1]!.row.visibleMark).toEqual({ kind: 'off' })
    s.close()
  })

  it('trace store is read-only', () => {
    const path = dbPath()
    const s = AuditStore.open(path); s.putRow('s1', row()); s.close()
    const r = AuditStore.openReadOnlyForTrace(path)
    expect(r.getRow('s1')).toEqual(row())
    expect(() => r.putRow('s2', row())).toThrow(/read-only/)
    expect(() => r.putDelivery({ sid: 's1', unitIndex: 0, scale: 1, status: 'l1', refSha256: 'aa', at: 1 })).toThrow(/read-only/)
    expect(() => r.takeToken(1)).toThrow(/read-only/)
    expect(() => r.expire(AUDIT_RETENTION_SEC * 9)).toThrow(/read-only/)
    const log = { at: 7, osUser: 'ops', investigator: 'Jo', imageSha256: 'ff', docRefs: ['doc-a'], unitIndex: 0, answer: 'match', sid: 's1' }
    r.putDetection(log)
    expect(r.detections()).toEqual([log])
    r.close()
    const s2 = AuditStore.open(path)
    expect(s2.detections()).toEqual([log]); expect(s2.deliveries('s1')).toEqual([])
    s2.close()
  })

  it('file mode is 0600', () => {
    const path = dbPath()
    AuditStore.open(path).close()
    expect(statSync(path).mode & 0o777).toBe(0o600)
  })

  // F2: session rows (IPs, viewer ids) live in the -wal until checkpoint; SQLite
  // must not leave the sidecars to the umask. Every file present is 0600.
  const modes = (path: string) => Object.fromEntries(['', '-wal', '-shm'].filter((x) => existsSync(path + x)).map((x) => [x || 'db', statSync(path + x).mode & 0o777]))
  const withUmask = <T>(mask: number, f: () => T): T => { const prev = process.umask(mask); try { return f() } finally { process.umask(prev) } }
  const write = (s: AuditStore, sid: string) => { s.putRow(sid, row()); s.putDelivery({ sid, unitIndex: 0, scale: 1, status: 'l1', refSha256: 'aa', at: 1 }); s.takeToken(1) }

  it('with umask 022, after writes, the db, -wal and -shm are all 0600 (server and trace opens)', () => withUmask(0o022, () => {
    const path = dbPath()
    const s = AuditStore.open(path)
    write(s, 's0')
    expect(modes(path)).toEqual({ db: 0o600, '-wal': 0o600, '-shm': 0o600 })
    // A trace open alongside the live server keeps them 0600.
    const t = AuditStore.openReadOnlyForTrace(path)
    t.putDetection({ at: 1, osUser: 'u', imageSha256: 'x', docRefs: ['doc-a'], unitIndex: 0, answer: 'no-match' })
    expect(modes(path)).toEqual({ db: 0o600, '-wal': 0o600, '-shm': 0o600 })
    t.close(); s.close()
    // A trace open after the server stopped (sidecars recreated) keeps them 0600.
    const t2 = AuditStore.openReadOnlyForTrace(path)
    t2.putDetection({ at: 2, osUser: 'u', imageSha256: 'y', docRefs: ['doc-a'], unitIndex: 0, answer: 'no-match' })
    expect(modes(path)).toEqual({ db: 0o600, '-wal': 0o600, '-shm': 0o600 })
    t2.close()
  }))

  it('an existing store with looser modes is tightened to 0600 on open (trace open too)', () => withUmask(0o022, () => {
    const path = dbPath()
    const s = AuditStore.open(path); write(s, 's0')
    for (const x of ['', '-wal', '-shm']) chmodSync(path + x, 0o644)
    s.close()
    for (const x of ['-wal', '-shm']) if (!existsSync(path + x)) closeSync(openSync(path + x, 'a', 0o644))
    chmodSync(path, 0o644)
    const t = AuditStore.openReadOnlyForTrace(path)
    expect(modes(path)).toEqual({ db: 0o600, '-wal': 0o600, '-shm': 0o600 })
    t.close()
  }))
})

describe('serving order', () => {
  it('store failure refuses, never sends unrecorded', async () => {
    const bad = AuditStore.open(':memory:')
    bad.putRow = () => { throw new Error('disk full') }
    const h = await start({ audit: new AuditTable({ store: bad }) })
    const c = new Client(h.url)
    await c.opened(); c.send({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1] })
    expect((await c.closed).code).toBe(4009)
    expect(c.msgs.filter((m) => 'binary' in m)).toEqual([])
    h.stop()

    const bad2 = AuditStore.open(':memory:')
    bad2.putDelivery = () => { throw new Error('disk full') }
    const h2 = await start({ audit: new AuditTable({ store: bad2 }) })
    const c2 = new Client(h2.url); await c2.hello(); await c2.ready()
    c2.unit(10, 0)
    const r = await c2.next((m) => 'text' in m && m.text.reqId === 10)
    expect((r as { text: Record<string, unknown> }).text).toMatchObject({ type: 'refused', code: 'policy-refused' })
    await Bun.sleep(50)
    expect(c2.msgs.filter((m) => 'binary' in m && m.binary.kind === 3)).toEqual([])
    c2.close(); h2.stop()
  }, 30_000)

  it('delivery recorded before send with reference hash', async () => {
    const store = AuditStore.open(':memory:')
    const h = await start({ audit: new AuditTable({ store }) })
    const c = new Client(h.url), m = await c.hello(); await c.ready()
    const sid = m.session.sid
    expect(store.getRow(sid)).toMatchObject({ docRef: m.docRef, contentPolicy: 'visible-only', visibleMark: { kind: 'fixed' }, forensic: { token24: 0 } })
    c.unit(10, 0); await c.unitMessages(10)
    expect(store.deliveries(sid)).toEqual([{ sid, unitIndex: 0, scale: 1, status: 'l1', refSha256: sha(h.stream.canonicalOf(sid, 0, 1)!), at: expect.any(Number) }])
    c.close(); h.stop()

    const zip = await JSZip.loadAsync(corpus('pypptx-test.pptx'))
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    zip.file('ppt/presentation.xml', pres.replace(/<p:sldSz[^>]*\/>/, '<p:sldSz cx="952500" cy="476250"/>'))
    const tiny = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    const store2 = AuditStore.open(':memory:')
    const h2 = await start({ doc: tiny, viewerId: 'bob', policy: { visibleMark: 'viewer' }, audit: new AuditTable({ store: store2 }) })
    const c2 = new Client(h2.url), m2 = await c2.hello(); await c2.ready()
    c2.unit(10, 0); await c2.unitMessages(10)
    const sid2 = m2.session.sid
    expect(store2.getRow(sid2)!.visibleMark).toEqual({ kind: 'viewer', text: 'bob' })
    expect(store2.deliveries(sid2)).toEqual([{ sid: sid2, unitIndex: 0, scale: 1, status: 'unmarked', reason: 'capacity', refSha256: sha(h2.stream.canonicalOf(sid2, 0, 1)!), at: expect.any(Number) }])
    c2.close(); h2.stop()
  }, 30_000)
})

describe('text services audit (Stage 3a Task 5)', () => {
  const copyRec = (sid: string, paragraphs: number): Omit<CopyRecord, 'at'> => ({ sid, kind: 'copy', paragraphs, fromUnit: 0, toUnit: 1 })

  it('copies are their own rows: repeats are kept, get() totals copy events only', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    t.record('s1', row({ services: { copy: { maxParagraphs: 3, maxSessionParagraphs: null }, textA11y: 'focused', print: false } }))
    t.recordCopy(copyRec('s1', 2)); t.recordCopy(copyRec('s1', 2))
    s.putCopy({ sid: 's1', kind: 'a11y', paragraphs: 1, fromUnit: 0, toUnit: 0, at: 9 })
    expect(s.copies('s1')).toEqual([{ ...copyRec('s1', 2), at: expect.any(Number) }, { ...copyRec('s1', 2), at: expect.any(Number) }, { sid: 's1', kind: 'a11y', paragraphs: 1, fromUnit: 0, toUnit: 0, at: 9 }])
    expect(t.get('s1')).toMatchObject({ services: { copy: { maxParagraphs: 3, maxSessionParagraphs: null }, textA11y: 'focused' }, servicesChanged: [], copied: { events: 2, paragraphs: 4 } })
    expect(() => t.recordCopy(copyRec('nope', 1))).toThrow(/no audit row/)
    expect(() => t.recordServices('nope', { copy: false, textA11y: 'off', print: false })).toThrow(/no audit row/)
    s.close()
  })

  it('services changes are recorded beside the write-once row, in order, with when', () => {
    const path = dbPath()
    const t = new AuditTable({ store: AuditStore.open(path) })
    t.record('s1', row())
    t.recordServices('s1', { copy: { maxParagraphs: null, maxSessionParagraphs: null }, textA11y: 'unit', print: { maxPages: 2, maxSessionPages: null } })
    t.recordServices('s1', { copy: false, textA11y: 'off', print: false })
    ;(t as unknown as { store: AuditStore }).store.close()
    const s2 = AuditStore.open(path), t2 = new AuditTable({ store: s2 })
    expect(s2.getRow('s1')).toEqual(row())
    expect(t2.get('s1')!.services).toEqual({ copy: false, textA11y: 'off', print: false })
    expect(t2.get('s1')!.servicesChanged).toEqual([
      { copy: { maxParagraphs: null, maxSessionParagraphs: null }, textA11y: 'unit', print: { maxPages: 2, maxSessionPages: null }, at: expect.any(Number) },
      { copy: false, textA11y: 'off', print: false, at: expect.any(Number) },
    ])
    s2.close()
  })

  it('a row stored before Stage 3a reads as services off', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    const { services: _s, ...legacy } = row()
    s.putRow('old', legacy as StoredRow)
    expect(t.get('old')).toMatchObject({ services: { copy: false, textA11y: 'off', print: false }, servicesChanged: [], copied: { events: 0, paragraphs: 0 } })
    s.close()
  })

  it('a services table created before Stage 3b gains its print column; old changes read as print off', () => {
    const path = dbPath()
    const raw = new Database(path, { create: true, strict: true })
    raw.exec('CREATE TABLE services (sid TEXT NOT NULL, copy TEXT NOT NULL, text_a11y TEXT NOT NULL, at INTEGER NOT NULL)')
    raw.query('INSERT INTO services (sid, copy, text_a11y, at) VALUES (?, ?, ?, ?)').run('s1', 'false', 'off', 5)
    raw.close()
    const s = AuditStore.open(path)
    expect(s.services('s1')).toEqual([{ sid: 's1', copy: false, textA11y: 'off', print: false, at: 5 }])
    s.putRow('s1', row())
    s.putServices({ sid: 's1', copy: false, textA11y: 'off', print: { maxPages: 1, maxSessionPages: 3 }, at: 6 })
    expect(s.services('s1').map((r) => r.print)).toEqual([false, { maxPages: 1, maxSessionPages: 3 }])
    s.close()
    AuditStore.open(path).close() // reopening is idempotent
  })

  it('a row stored before Stage 3b (no print in its services) reads as print off', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    s.putRow('old', { ...row(), services: { copy: false, textA11y: 'off' } } as unknown as StoredRow)
    expect(t.get('old')!.services).toEqual({ copy: false, textA11y: 'off', print: false })
    s.close()
  })

  it('expire() removes copies with their row', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    t.record('old', row({ startedAt: 0 })); t.record('new', row({ startedAt: 100 }))
    t.recordCopy(copyRec('old', 1)); t.recordCopy(copyRec('new', 1))
    t.recordServices('old', { copy: false, textA11y: 'off', print: false })
    expect(t.expire(AUDIT_RETENTION_SEC)).toBe(1)
    expect(s.copies('old')).toEqual([]); expect(s.services('old')).toEqual([])
    expect(s.copies('new').length).toBe(1)
    s.close()
  })

  it('trace store refuses copy and services writes', () => {
    const path = dbPath()
    const s = AuditStore.open(path); s.putRow('s1', row()); s.close()
    const r = AuditStore.openReadOnlyForTrace(path)
    expect(() => r.putCopy({ ...copyRec('s1', 1), at: 1 })).toThrow(/read-only/)
    expect(() => r.putServices({ sid: 's1', copy: false, textA11y: 'off', print: false, at: 1 })).toThrow(/read-only/)
    r.close()
  })
})

describe('print audit (Stage 3b Task 2)', () => {
  const printRec = (sid: string, pages: number, scale: 1 | 2 = 1): Omit<PrintRecord, 'at'> => ({ sid, pages, fromUnit: 0, toUnit: pages - 1, scale })

  it('prints are their own rows: kept in order with when, get() totals jobs and pages', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    t.record('s1', row())
    t.recordPrint(printRec('s1', 2)); t.recordPrint(printRec('s1', 3, 2))
    expect(s.prints('s1')).toEqual([{ ...printRec('s1', 2), at: expect.any(Number) }, { ...printRec('s1', 3, 2), at: expect.any(Number) }])
    expect(s.prints('s1')[0]!.at).toBeGreaterThan(0)
    expect(t.get('s1')!.printed).toEqual({ jobs: 2, pages: 5 })
    s.close()
  })

  it('a fresh row has printed {jobs: 0, pages: 0}', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    t.record('s1', row())
    expect(t.get('s1')!.printed).toEqual({ jobs: 0, pages: 0 })
    s.close()
  })

  it('recordPrint for an unknown sid throws', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    expect(() => t.recordPrint(printRec('nope', 1))).toThrow(/no audit row/)
    expect(s.prints('nope')).toEqual([])
    s.close()
  })

  it('expire() removes prints with their row', () => {
    const s = AuditStore.open(':memory:'), t = new AuditTable({ store: s })
    t.record('old', row({ startedAt: 0 })); t.record('new', row({ startedAt: 100 }))
    t.recordPrint(printRec('old', 1)); t.recordPrint(printRec('new', 1))
    expect(t.expire(AUDIT_RETENTION_SEC)).toBe(1)
    expect(s.prints('old')).toEqual([]); expect(s.prints('new').length).toBe(1)
    s.close()
  })

  it('a store created before the prints table gains it on open', () => {
    const path = dbPath()
    const raw = new Database(path, { create: true }); raw.exec('CREATE TABLE rows (sid TEXT PRIMARY KEY, started_at INTEGER NOT NULL, key_generation INTEGER NOT NULL, doc_ref TEXT NOT NULL, msg40 TEXT, row TEXT NOT NULL)'); raw.close()
    const s = AuditStore.open(path); s.putRow('s1', row())
    s.putPrint({ ...printRec('s1', 1), at: 7 })
    expect(s.prints('s1')).toEqual([{ ...printRec('s1', 1), at: 7 }])
    s.close()
  })

  it('trace store refuses print writes', () => {
    const path = dbPath()
    const s = AuditStore.open(path); s.putRow('s1', row()); s.close()
    const r = AuditStore.openReadOnlyForTrace(path)
    expect(() => r.putPrint({ ...printRec('s1', 1), at: 1 })).toThrow(/read-only/)
    r.close()
  })
})
