import { afterAll, describe, expect, it } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { AUDIT_RETENTION_SEC } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { demoHooks, safeDocPath } from '../src/demo-hooks'
import { startServer } from '../src/main'
import { Client } from './harness'

const corpusDir = resolve(import.meta.dir, '../../corpus')
const servers: Array<{ stop(): void }> = []
afterAll(() => { for (const s of servers) s.stop() })
const req = (q: string) => ({ headers: new Headers(), ip: '127.0.0.1', url: new URL(`ws://x/stream?${q}`) })

describe('demo hooks', () => {
  it('resolves bare names inside the doc dir only', () => {
    for (const bad of ['../package.json', '..%2f..%2fpackage.json', 'a/b.docx', 'x.odt', '..docx', '.hidden.docx', '/etc/passwd', 'x.docx\u0000.txt', '']) expect(safeDocPath(corpusDir, bad)).toBeUndefined()
    expect(safeDocPath(corpusDir, 'pypptx-test.pptx')).toBe(resolve(corpusDir, 'pypptx-test.pptx'))
    expect(safeDocPath('', 'pypptx-test.pptx')).toBeUndefined()
  })
  it('URL-decoded traversal and missing files resolve to nothing', async () => {
    const { resolveDocument } = demoHooks({ docDir: corpusDir, demo: true })
    for (const q of ['doc=%2e%2e%2fpackage.json', 'doc=..%2Fpackage.json', 'doc=missing.docx', '']) expect(await resolveDocument(req(q), { policy: {} })).toBeNull()
    expect((await resolveDocument(req('doc=pypptx-test.pptx'), { policy: {} }))?.bytes.length).toBeGreaterThan(0)
  })
  it('a symlink inside the doc dir cannot reach a file outside it (review r2-10)', async () => {
    const { mkdtempSync, symlinkSync, copyFileSync } = await import('node:fs')
    const { tmpdir } = await import('node:os')
    const dir = mkdtempSync(resolve(tmpdir(), 'ovdocs-')), outside = mkdtempSync(resolve(tmpdir(), 'ovprivate-'))
    copyFileSync(resolve(corpusDir, 'pypptx-test.pptx'), resolve(outside, 'secret.pptx'))
    copyFileSync(resolve(corpusDir, 'pypptx-test.pptx'), resolve(dir, 'real.pptx'))
    symlinkSync(resolve(outside, 'secret.pptx'), resolve(dir, 'deck.pptx'))
    const { resolveDocument } = demoHooks({ docDir: dir, demo: true })
    expect(await resolveDocument(req('doc=deck.pptx'), { policy: {} })).toBeNull()
    expect(await resolveDocument(req('doc=real.pptx'), { policy: {} })).not.toBeNull()
  })
  it('?relaxed=1 is honoured only with --demo', async () => {
    expect((await demoHooks({ docDir: corpusDir, demo: true }).authorize(req('relaxed=1')))?.policy).toEqual({ projection: 'relaxed' })
    expect((await demoHooks({ docDir: corpusDir, demo: false }).authorize(req('relaxed=1')))?.policy).toEqual({})
  })
})

describe('demo server end to end', () => {
  it('refuses to start without a key outside --demo, and without an origin', async () => {
    await expect(startServer(['--doc-dir', corpusDir], {})).rejects.toThrow(/OV_STREAM_KEY/)
    await expect(startServer(['--doc-dir', corpusDir], { OV_STREAM_KEY: btoa('k'.repeat(32)), OV_AUDIT_DB: ':memory:' })).rejects.toThrow(/origin/)
  })
  it('opens the OV_AUDIT_DB store outside --demo and expires old rows at startup', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-audit-')), path = join(dir, 'audit.db')
    try {
      const pre = AuditStore.open(path)
      const row = { ip: '1.2.3.4', keyGeneration: 1, policyId: 'p', markSettings: { visible: 'off' as const, forensicRequested: false }, relaxed: false, forensic: null, docRef: 'd', contentPolicy: 'visible-only' as const, visibleMark: { kind: 'off' as const }, services: { copy: false as const, textA11y: 'off' as const, print: false as const } }
      pre.putRow('old', { ...row, startedAt: 0 }); pre.putRow('fresh', { ...row, startedAt: Math.floor(Date.now() / 1000) - AUDIT_RETENTION_SEC + 3600 }); pre.close()
      const s = await startServer(['--port', '0', '--doc-dir', corpusDir, '--origin', 'http://a'], { OV_STREAM_KEY: btoa('k'.repeat(32)), OV_AUDIT_DB: path })
      expect(s.stream.audit.get('old')).toBeUndefined()
      expect(s.stream.audit.get('fresh')).toBeDefined()
      s.stop()
    } finally { rmSync(dir, { recursive: true, force: true }) }
  })
  it('streams every unit of one file per format at scale 1, protected and relaxed', async () => {
    const s = await startServer(['--demo', '--port', '0', '--doc-dir', corpusDir], {}); servers.push(s)
    const base = `ws://localhost:${s.server.port}/stream`
    const counts: string[] = []
    for (const doc of ['pydocx-test.docx', 'poi-56295.xlsx', 'pypptx-test.pptx']) for (const relaxed of [false, true]) {
      const c = new Client(`${base}?doc=${doc}${relaxed ? '&relaxed=1' : ''}`)
      const m = await c.hello(); await c.ready()
      expect(m.policy.protection).toBe(relaxed ? 'commands' : 'final-raster')
      for (let u = 0; u < m.units.count; u++) { c.unit(10 + u, u, 1); await c.unitMessages(10 + u) }
      expect(c.msgs.filter((x) => 'text' in x && x.text.type === 'refused')).toEqual([])
      counts.push(`${doc}${relaxed ? '+relaxed' : ''}:${m.units.count}`)
      c.close()
    }
    console.log(`[demo] ${counts.join(' ')} · pool ${JSON.stringify(s.pool.stats().workers.map((w) => ({ pages: w.pages, parses: w.parses })))}`)
    // ODT and traversal never reach a session.
    for (const [q, code] of [['doc=x.odt', 4003], ['doc=..%2Fpackage.json', 4003]] as const) {
      const c = new Client(`${base}?${q}`); await c.opened(); c.send({ reqId: 1, type: 'hello', client: 'x', recVersions: [1] })
      expect((await c.closed).code).toBe(code)
    }
  }, 60_000)
})
