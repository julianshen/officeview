/**
 * Manual-device result collector.
 *
 * Mounted by vite.harness.config.ts as a middleware on the SAME ORIGIN as the
 * harness page (no second port, so no CORS and no cross-origin exfiltration
 * path). It is deliberately small and paranoid because on the manual path it
 * listens on the LAN:
 *
 *  - a random per-session token (printed in the device URL) must accompany
 *    every POST, compared in constant time;
 *  - the body is capped at 1 MiB while it streams (not just by Content-Length);
 *  - an Origin that differs from Host is refused;
 *  - the artifact file name is generated here ([a-z0-9-] only) and NEVER built
 *    from payload text; the resolved path is asserted to stay inside --output
 *    and the file is created exclusively (`wx`), so nothing is overwritten;
 *  - the payload must pass the full ingest validation (honesty invariants plus
 *    hash) and must look like an unautomated device run for the profile, mode
 *    and fixture this session served.
 *
 * The logic is node-http free (`ingest`) so tests can drive it directly; the
 * thin `createCollectorMiddleware` adapter does the streaming.
 */
import { createHash, randomBytes, timingSafeEqual } from 'node:crypto'
import { mkdirSync, writeFileSync } from 'node:fs'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { resolve, sep } from 'node:path'
import { COLLECT_PATH } from '../../src/debug/harness/collect'
import { verifyHarnessResult } from '../../src/debug/harness/result'
import type { FixtureId, HarnessMode, HarnessProfile, HarnessResult } from '../../src/debug/harness/types'

export const MAX_BODY_BYTES = 1024 * 1024
export const MAX_ARTIFACTS = 100
/** Labels a page may legitimately report from an unautomated device browser. */
export const ALLOWED_DEVICE_BROWSERS: readonly string[] = ['safari', 'webkit', 'chromium', 'firefox']
const MAX_REPORTED_ERRORS = 20
const NAME_ATTEMPTS = 5

export interface CollectorConfig {
  token: string
  outputDir: string
  /** What this session served; results for anything else are refused. */
  expect: { profile: HarnessProfile; mode: HarnessMode; requestedFixtureId: FixtureId }
  maxBodyBytes?: number
  maxArtifacts?: number
  now?: () => Date
  newId?: () => string
}

export interface IngestRequest {
  token: string | null
  origin: string | null
  host: string | null
  body: string
}

export interface IngestResponse {
  status: number
  body: Record<string, unknown>
}

export interface Collector {
  readonly maxBodyBytes: number
  ingest(request: IngestRequest): Promise<IngestResponse>
}

const SAFE_NAME = /^[A-Za-z0-9][A-Za-z0-9._-]{0,200}$/

/** The only way an artifact path is built: a safe basename inside outputDir. */
export function resolveArtifactPath(outputDir: string, fileName: string): string {
  if (!SAFE_NAME.test(fileName)) throw new Error(`invalid artifact name ${JSON.stringify(fileName)}`)
  const root = resolve(outputDir)
  const full = resolve(root, fileName)
  if (!full.startsWith(root + sep)) throw new Error(`artifact path ${JSON.stringify(fileName)} resolves outside the output directory`)
  return full
}

function tokensMatch(expected: string, provided: string | null): boolean {
  if (provided === null || provided.length === 0) return false
  // Hash both sides so timingSafeEqual sees equal lengths and length leaks nothing.
  const a = createHash('sha256').update(expected).digest()
  const b = createHash('sha256').update(provided).digest()
  return timingSafeEqual(a, b)
}

function fail(status: number, errors: string[]): IngestResponse {
  return { status, body: { ok: false, errors: errors.slice(0, MAX_REPORTED_ERRORS) } }
}

export function createCollector(config: CollectorConfig): Collector {
  const maxBodyBytes = config.maxBodyBytes ?? MAX_BODY_BYTES
  const maxArtifacts = config.maxArtifacts ?? MAX_ARTIFACTS
  const now = config.now ?? (() => new Date())
  const newId = config.newId ?? (() => randomBytes(4).toString('hex'))
  const seenHashes = new Set<string>()
  let written = 0
  mkdirSync(config.outputDir, { recursive: true })

  async function ingest(request: IngestRequest): Promise<IngestResponse> {
    if (!tokensMatch(config.token, request.token)) return fail(401, ['missing or invalid session token'])
    if (request.origin !== null) {
      let originHost: string | null = null
      try {
        originHost = new URL(request.origin).host
      } catch {
        // fall through: unparsable origin is refused below
      }
      if (originHost === null || originHost !== request.host) return fail(403, ['cross-origin posts are refused'])
    }
    if (Buffer.byteLength(request.body) > maxBodyBytes) return fail(413, [`body exceeds the ${maxBodyBytes}-byte limit`])
    if (written >= maxArtifacts) return fail(429, [`this session already stored ${maxArtifacts} artifacts`])

    let payload: unknown
    try {
      payload = JSON.parse(request.body)
    } catch {
      return fail(400, ['request body is not valid JSON'])
    }
    const verification = await verifyHarnessResult(payload)
    if (!verification.ok) return fail(400, verification.errors)

    const result = payload as HarnessResult
    const env = result.environment
    const refusals: string[] = []
    if (env.browser.automation !== 'none') refusals.push(`manual-device results must have automation "none", got "${env.browser.automation}"`)
    if (!ALLOWED_DEVICE_BROWSERS.includes(env.browser.name)) refusals.push(`unexpected device label ${JSON.stringify(env.browser.name)}; allowed: ${ALLOWED_DEVICE_BROWSERS.join(', ')}`)
    if (env.evidence === 'emulated-webkit') refusals.push('emulated WebKit is not device evidence')
    if (env.profile !== config.expect.profile) refusals.push(`result profile ${env.profile} differs from the served profile ${config.expect.profile}`)
    if (env.mode !== config.expect.mode) refusals.push(`result mode ${env.mode} differs from the served mode ${config.expect.mode}`)
    if (env.requestedFixtureId !== config.expect.requestedFixtureId) refusals.push(`result fixture ${env.requestedFixtureId} differs from the served fixture ${config.expect.requestedFixtureId}`)
    if (refusals.length > 0) return fail(422, refusals)

    if (seenHashes.has(result.hash)) return fail(409, ['this result was already ingested'])

    const stamp = now().toISOString().replace(/[:.]/g, '-')
    const text = JSON.stringify({
      schemaVersion: 1,
      run: {
        browser: env.browser.name,
        collection: 'manual-device-post',
        note: 'device-reported labeling; physical-device verification is the operator responsibility (spec 12.2)',
      },
      results: [result],
    })
    // Name = server-controlled fields only (validated enums + clock + random id).
    for (let attempt = 0; attempt < NAME_ATTEMPTS; attempt++) {
      const fileName = `viewer-manual-${env.requestedFixtureId}-${env.profile}-${stamp}-${newId()}.json`
      let path: string
      try {
        path = resolveArtifactPath(config.outputDir, fileName)
      } catch (error) {
        return fail(500, [error instanceof Error ? error.message : String(error)])
      }
      try {
        writeFileSync(path, `${text}\n`, { flag: 'wx' })
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'EEXIST') continue
        return fail(500, ['could not store the artifact'])
      }
      seenHashes.add(result.hash)
      written += 1
      return { status: 200, body: { ok: true, artifact: fileName } }
    }
    return fail(500, ['could not allocate a free artifact name'])
  }

  return { maxBodyBytes, ingest }
}

function send(res: ServerResponse, status: number, body: unknown, extraHeaders: Record<string, string> = {}): void {
  const text = JSON.stringify(body)
  res.writeHead(status, {
    'content-type': 'application/json',
    'content-length': Buffer.byteLength(text),
    'cache-control': 'no-store',
    ...extraHeaders,
  })
  res.end(text)
}

/** Connect-style adapter; streams the body with a hard byte cap. */
export function createCollectorMiddleware(collector: Collector): (req: IncomingMessage, res: ServerResponse, next: () => void) => void {
  return (req, res, next) => {
    const pathname = new URL(req.url ?? '/', 'http://harness.invalid').pathname
    if (pathname !== COLLECT_PATH) {
      next()
      return
    }
    if (req.method !== 'POST') {
      send(res, 405, { ok: false, errors: ['POST only'] }, { allow: 'POST' })
      return
    }
    const max = collector.maxBodyBytes
    let aborted = false
    const tooLarge = (): void => {
      if (aborted) return
      aborted = true
      const text = JSON.stringify({ ok: false, errors: [`body exceeds the ${max}-byte limit`] })
      res.writeHead(413, { 'content-type': 'application/json', 'content-length': Buffer.byteLength(text), connection: 'close' })
      res.end(text, () => req.destroy())
    }
    const declared = Number(req.headers['content-length'])
    if (Number.isFinite(declared) && declared > max) {
      tooLarge()
      return
    }
    const chunks: Buffer[] = []
    let received = 0
    req.on('data', (chunk: Buffer) => {
      if (aborted) return
      received += chunk.length
      if (received > max) {
        chunks.length = 0
        tooLarge()
        return
      }
      chunks.push(chunk)
    })
    req.on('error', () => {
      aborted = true
    })
    req.on('end', () => {
      if (aborted) return
      const token = req.headers['x-harness-token']
      void collector
        .ingest({
          token: typeof token === 'string' ? token : null,
          origin: typeof req.headers.origin === 'string' ? req.headers.origin : null,
          host: typeof req.headers.host === 'string' ? req.headers.host : null,
          body: Buffer.concat(chunks).toString('utf8'),
        })
        .then((response) => send(res, response.status, response.body))
        .catch(() => send(res, 500, { ok: false, errors: ['collector failure'] }))
    })
  }
}
