// @vitest-environment node
/**
 * Manual-device collector (findings 2, 3, 5, 6): same-origin ingest with a
 * per-session token, a streaming body cap, exclusive-create writes under a
 * server-generated name, and rejection of anything that is not an honest,
 * unautomated device result. Runs against a real temp directory and a real
 * node http server.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { createServer, request as httpRequest, type Server } from 'node:http'
import type { AddressInfo } from 'node:net'
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import {
  ALLOWED_DEVICE_BROWSERS,
  MAX_BODY_BYTES,
  createCollector,
  createCollectorMiddleware,
  resolveArtifactPath,
  type Collector,
  type CollectorConfig,
} from '../scripts/harness/collector'
import { buildDeviceResult, clone, makeSample, range, rehash } from './helpers/harnessResult'

const TOKEN = 'tok-0123456789abcdef0123456789abcdef'
let sandbox: string
let outputDir: string

beforeEach(() => {
  sandbox = mkdtempSync(join(tmpdir(), 'collector-'))
  outputDir = join(sandbox, 'out')
  mkdirSync(outputDir)
})

function config(overrides: Partial<CollectorConfig> = {}): CollectorConfig {
  return {
    token: TOKEN,
    outputDir,
    expect: { profile: 'desktop', mode: 'smoke', requestedFixtureId: 'F1' },
    ...overrides,
  }
}

function post(collector: Collector, body: unknown, extra: { token?: string | null; origin?: string | null; host?: string | null } = {}) {
  return collector.ingest({
    token: extra.token === undefined ? TOKEN : extra.token,
    origin: extra.origin ?? null,
    host: extra.host ?? 'localhost:4173',
    body: typeof body === 'string' ? body : JSON.stringify(body),
  })
}

describe('accepting an honest device result', () => {
  it('writes one compact artifact under a server-generated name and answers with the basename only', async () => {
    const collector = createCollector(config({ now: () => new Date('2026-10-10T01:02:03.456Z'), newId: () => 'abcd1234' }))
    const result = await buildDeviceResult()
    const response = await post(collector, result)
    expect(response.status).toBe(200)
    expect(response.body).toEqual({ ok: true, artifact: 'viewer-manual-F1-desktop-2026-10-10T01-02-03-456Z-abcd1234.json' })
    const files = readdirSync(outputDir)
    expect(files).toEqual(['viewer-manual-F1-desktop-2026-10-10T01-02-03-456Z-abcd1234.json'])
    const text = readFileSync(join(outputDir, files[0]), 'utf8')
    expect(text).not.toMatch(/\n\s+"/) // not pretty-printed
    const artifact = JSON.parse(text)
    expect(artifact.run).toMatchObject({ browser: 'safari', collection: 'manual-device-post' })
    expect(artifact.results).toEqual([result])
  })

  it('never puts payload text in the path, even when free-text fields contain traversal sequences', async () => {
    const collector = createCollector(config())
    const result = await buildDeviceResult({ note: '../../../../tmp/pwned', cache: { state: 'cold', note: '..\\..\\x' } })
    expect((await post(collector, result)).status).toBe(200)
    expect(readdirSync(outputDir)).toHaveLength(1)
    expect(readdirSync(outputDir)[0]).toMatch(/^viewer-manual-F1-desktop-[0-9T-]+Z-[0-9a-f]{8}\.json$/)
    expect(readdirSync(sandbox)).toEqual(['out'])
  })
})

describe('traversal payloads are rejected and nothing escapes --output', () => {
  it('rejects ../ in browser.name even with a recomputed hash', async () => {
    const collector = createCollector(config())
    const forged = await rehash(await buildDeviceResult(), (d) => (d.environment.browser.name = '../../../escape'))
    const response = await post(collector, forged)
    expect(response.status).toBe(400)
    expect(readdirSync(outputDir)).toEqual([])
    expect(readdirSync(sandbox)).toEqual(['out'])
  })

  it('rejects traversal in requestedFixtureId, profile and mode', async () => {
    const collector = createCollector(config())
    for (const mutate of [
      (d: Record<string, any>) => (d.environment.requestedFixtureId = '../x'),
      (d: Record<string, any>) => (d.environment.profile = '../../p'),
      (d: Record<string, any>) => (d.environment.mode = '../m'),
    ]) {
      const response = await post(collector, await rehash(await buildDeviceResult(), mutate))
      expect(response.status).toBe(400)
    }
    expect(readdirSync(sandbox)).toEqual(['out'])
    expect(readdirSync(outputDir)).toEqual([])
  })

  it('resolveArtifactPath refuses names that leave the output directory', () => {
    for (const bad of ['../x.json', '..\\x.json', '/etc/passwd', 'a/../../b.json', 'sub/dir.json', '', '.', '..']) {
      expect(() => resolveArtifactPath(outputDir, bad), JSON.stringify(bad)).toThrow(/outside|invalid/i)
    }
    expect(resolveArtifactPath(outputDir, 'ok-name.json')).toBe(join(resolve(outputDir), 'ok-name.json'))
  })
})

describe('exclusive create: nothing is ever overwritten', () => {
  it('retries with a fresh id when the generated name already exists', async () => {
    const ids = ['same0001', 'same0001', 'diff0002']
    const collector = createCollector(config({ now: () => new Date('2026-01-01T00:00:00.000Z'), newId: () => ids.shift() ?? 'zzzz9999' }))
    expect((await post(collector, await buildDeviceResult())).status).toBe(200)
    const first = readdirSync(outputDir)[0]
    const firstBody = readFileSync(join(outputDir, first), 'utf8')
    const second = await buildDeviceResult({}, [makeSample(range(31))])
    const response = await post(collector, second)
    expect(response.status).toBe(200)
    expect(readFileSync(join(outputDir, first), 'utf8')).toBe(firstBody)
    expect(readdirSync(outputDir)).toHaveLength(2)
  })

  it('fails closed (500) and leaves the existing file intact when no free name is found', async () => {
    const collector = createCollector(config({ now: () => new Date('2026-01-01T00:00:00.000Z'), newId: () => 'constant0' }))
    const name = 'viewer-manual-F1-desktop-2026-01-01T00-00-00-000Z-constant0.json'
    writeFileSync(join(outputDir, name), 'precious')
    const response = await post(collector, await buildDeviceResult())
    expect(response.status).toBe(500)
    expect(readFileSync(join(outputDir, name), 'utf8')).toBe('precious')
  })
})

describe('authentication and origin', () => {
  it('rejects a missing or wrong token with 401 and writes nothing', async () => {
    const collector = createCollector(config())
    const result = await buildDeviceResult()
    for (const token of [null, '', 'nope', TOKEN.slice(0, -1), `${TOKEN}x`]) {
      expect((await post(collector, result, { token })).status, String(token)).toBe(401)
    }
    expect(readdirSync(outputDir)).toEqual([])
  })

  it('rejects a cross-origin POST (Origin differs from Host) with 403', async () => {
    const collector = createCollector(config())
    const result = await buildDeviceResult()
    expect((await post(collector, result, { origin: 'http://evil.example', host: 'localhost:4173' })).status).toBe(403)
    expect((await post(collector, result, { origin: 'http://localhost:4173', host: 'localhost:4173' })).status).toBe(200)
  })
})

describe('size and count limits', () => {
  it('rejects a body over the cap with 413 before parsing', async () => {
    const collector = createCollector(config({ maxBodyBytes: 2048 }))
    expect((await post(collector, 'x'.repeat(2049))).status).toBe(413)
    expect(readdirSync(outputDir)).toEqual([])
    expect(createCollector(config()).maxBodyBytes).toBe(MAX_BODY_BYTES)
    expect(MAX_BODY_BYTES).toBe(1024 * 1024)
  })

  it('caps the number of artifacts per session with 429', async () => {
    const collector = createCollector(config({ maxArtifacts: 2 }))
    for (const n of [30, 31]) expect((await post(collector, await buildDeviceResult({}, [makeSample(range(n))]))).status).toBe(200)
    expect((await post(collector, await buildDeviceResult({}, [makeSample(range(32))]))).status).toBe(429)
    expect(readdirSync(outputDir)).toHaveLength(2)
  })

  it('rejects a duplicate result (same hash) with 409', async () => {
    const collector = createCollector(config())
    const result = await buildDeviceResult()
    expect((await post(collector, result)).status).toBe(200)
    expect((await post(collector, result)).status).toBe(409)
    expect(readdirSync(outputDir)).toHaveLength(1)
  })
})

describe('validation and honesty rules', () => {
  it('rejects malformed JSON with 400', async () => {
    expect((await post(createCollector(config()), '{not json')).status).toBe(400)
  })

  it('rejects hash mismatches and dishonest metrics with the validator reasons', async () => {
    const collector = createCollector(config())
    const result = await buildDeviceResult()
    const tampered = clone(result) as Record<string, any>
    tampered.environment.note = 'edited'
    const mismatch = await post(collector, tampered)
    expect(mismatch.status).toBe(400)
    expect(JSON.stringify(mismatch.body)).toMatch(/hash mismatch/)
    const forged = await rehash(result, (d) => (d.metrics.op.p50 = 1))
    const forgedResponse = await post(collector, forged)
    expect(forgedResponse.status).toBe(400)
    expect(JSON.stringify(forgedResponse.body)).toMatch(/derived/)
    expect(readdirSync(outputDir)).toEqual([])
  })

  it('rejects automated results (automation must be none) with 422', async () => {
    const collector = createCollector(config())
    const automated = await buildDeviceResult({ browser: { name: 'chromium', version: '1', engine: 'blink', userAgent: 'x', automation: 'playwright' } })
    const response = await post(collector, automated)
    expect(response.status).toBe(422)
    expect(JSON.stringify(response.body)).toMatch(/automation/)
  })

  it('rejects Playwright WebKit and unknown browser labels', async () => {
    const collector = createCollector(config())
    const emulated = await buildDeviceResult({ browser: { name: 'webkit-playwright', version: '1', engine: 'webkit', userAgent: 'x', automation: 'playwright' } })
    expect((await post(collector, emulated)).status).toBe(422)
    const unknown = await buildDeviceResult({ browser: { name: 'unknown', version: '1', engine: 'unknown', userAgent: 'x', automation: 'none' } })
    const response = await post(collector, unknown)
    expect(response.status).toBe(422)
    expect(JSON.stringify(response.body)).toMatch(/label/)
    expect(ALLOWED_DEVICE_BROWSERS).toEqual(['safari', 'webkit', 'chromium', 'firefox'])
  })

  it('rejects a result whose profile, mode or requested fixture differs from what this session served', async () => {
    const collector = createCollector(config())
    expect((await post(collector, await buildDeviceResult({ profile: 'mobile', viewport: { width: 1024, height: 768 } }))).status).toBe(422)
    expect((await post(collector, await buildDeviceResult({ requestedFixtureId: 'F2' }))).status).toBe(422)
    expect(readdirSync(outputDir)).toEqual([])
  })
})

describe('HTTP middleware (same-origin, no CORS)', () => {
  let server: Server
  let base: string
  let collector: Collector

  beforeEach(async () => {
    collector = createCollector(config({ maxBodyBytes: 64 * 1024 }))
    const middleware = createCollectorMiddleware(collector)
    server = createServer((req, res) => middleware(req, res, () => {
      res.statusCode = 404
      res.end('next')
    }))
    await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
    base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`
  })
  afterEach(async () => {
    server.closeAllConnections()
    await new Promise((done) => server.close(done))
  })

  it('ingests a POST and sends no CORS headers; OPTIONS and GET are not served as CORS', async () => {
    const result = await buildDeviceResult()
    const response = await fetch(`${base}/__harness/collect`, { method: 'POST', headers: { 'x-harness-token': TOKEN, 'content-type': 'application/json' }, body: JSON.stringify(result) })
    expect(response.status).toBe(200)
    expect(response.headers.get('access-control-allow-origin')).toBeNull()
    const options = await fetch(`${base}/__harness/collect`, { method: 'OPTIONS' })
    expect(options.status).toBe(405)
    expect(options.headers.get('access-control-allow-origin')).toBeNull()
    const get = await fetch(`${base}/__harness/collect`)
    expect(get.status).toBe(405)
    expect((await fetch(`${base}/other`)).status).toBe(404)
  })

  it('answers 413 for a declared Content-Length over the cap without reading the body', async () => {
    const response = await fetch(`${base}/__harness/collect`, { method: 'POST', headers: { 'x-harness-token': TOKEN }, body: 'x'.repeat(70 * 1024) })
    expect(response.status).toBe(413)
    expect(readdirSync(outputDir)).toEqual([])
  })

  it('aborts a chunked (no Content-Length) upload that crosses the cap while streaming', async () => {
    const outcome = await new Promise<string>((done) => {
      const req = httpRequest(`${base}/__harness/collect`, { method: 'POST', headers: { 'x-harness-token': TOKEN, 'transfer-encoding': 'chunked' } }, (res) => {
        done(`status ${res.statusCode}`)
        res.resume()
      })
      req.on('error', () => done('connection aborted'))
      const chunk = 'x'.repeat(16 * 1024)
      let sent = 0
      const pump = () => {
        if (sent >= 200 || req.destroyed) return
        sent += 1
        req.write(chunk, pump)
      }
      pump()
    })
    expect(outcome).toMatch(/status 413|connection aborted/)
    expect(readdirSync(outputDir)).toEqual([])
  })

  it('rejects a wrong token over HTTP', async () => {
    const response = await fetch(`${base}/__harness/collect`, { method: 'POST', headers: { 'x-harness-token': 'nope' }, body: '{}' })
    expect(response.status).toBe(401)
  })
})
