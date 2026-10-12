/**
 * Serves viewer fixtures to a browser that Playwright does not control (real
 * Safari through safaridriver, VoiceOver runs). The Playwright specs route
 * /__fixtures/* inside the browser; a real browser has no such interception and
 * the harness vite server only serves src/debug, so this is a separate loopback
 * server with CORS, answering a fixed allow-list of names and nothing else.
 *
 * Generated packages use the same builders, parameters and cache directory as
 * tests/browser/viewer-e2e.ts, so both harnesses open identical bytes.
 */
import { existsSync, mkdirSync, readFileSync } from 'node:fs'
import { createServer, type Server } from 'node:http'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { writeFixturePackage } from '../generate-viewer-fixtures'

const GENERATED = join(tmpdir(), 'officeview-viewer-e2e-v1')

export const SERVED_FIXTURES = {
  'f1-small.xlsx': { label: 'F1 committed small package (corners plus 200 sampled cells)' },
  'f1-paged.xlsx': { label: 'F1 sampled corners, nnz 120,000 (crosses the 100,000-cell object cap)' },
  'f2-1k.xlsx': { label: 'F2 tall table at 1,000 rows (reduced scale, object path)' },
  'f2-paged.xlsx': { label: 'F2 tall table at 100,010 rows (reduced scale; crosses the 100,000-row object cap)' },
  'f2-storage.xlsx': { label: 'F2 tall table at 160,000 rows (page bytes exceed the 32 MiB memory store; needs OPFS/IndexedDB)' },
  'f3-paged.xlsx': { label: 'F3 dense wide table at 8 rows x 16,384 columns (131,072 cells)' },
} as const
export type ServedFixture = keyof typeof SERVED_FIXTURES

export function servedFixturePath(root: string, name: ServedFixture): string {
  return name === 'f1-small.xlsx' ? join(root, 'tests', 'fixtures', 'xlsx-viewer', name) : join(GENERATED, name)
}

/** Writes any missing generated package (deterministic bytes, cached per machine). */
export async function ensureServedFixtures(names: readonly ServedFixture[]): Promise<void> {
  mkdirSync(GENERATED, { recursive: true })
  const writers: Partial<Record<ServedFixture, (path: string) => Promise<unknown>>> = {
    'f1-paged.xlsx': path => writeFixturePackage(path, 'F1', { nnz: 120_000 }),
    'f2-1k.xlsx': path => writeFixturePackage(path, 'F2', { rowCount: 1000 }),
    'f2-paged.xlsx': path => writeFixturePackage(path, 'F2', { rowCount: 100_010 }),
    'f2-storage.xlsx': path => writeFixturePackage(path, 'F2', { rowCount: 160_000 }),
    'f3-paged.xlsx': path => writeFixturePackage(path, 'F3', { rowCount: 8 }),
  }
  for (const name of names) {
    const write = writers[name]
    const path = join(GENERATED, name)
    if (write && !existsSync(path)) await write(path)
  }
}

export interface FixtureServer {
  /** The bound port (the requested one, or the OS choice for port 0). */
  port: number
  url(name: ServedFixture): string
  close(): Promise<void>
}

export async function startFixtureServer(root: string, port: number): Promise<FixtureServer> {
  const server: Server = createServer((request, response) => {
    response.setHeader('access-control-allow-origin', '*')
    let name: string
    try {
      name = decodeURIComponent(new URL(request.url ?? '/', 'http://x').pathname.slice(1))
    } catch {
      // A malformed path (e.g. a broken percent escape) is the client's error, never a crash.
      response.statusCode = 400
      response.end('malformed request path')
      return
    }
    if (!(name in SERVED_FIXTURES) || request.method !== 'GET') {
      response.statusCode = 404
      response.end('not a served fixture')
      return
    }
    try {
      const bytes = readFileSync(servedFixturePath(root, name as ServedFixture))
      response.setHeader('content-type', 'application/octet-stream')
      response.setHeader('cache-control', 'no-store')
      response.end(bytes)
    } catch (error) {
      response.statusCode = 500
      response.end(error instanceof Error ? error.message : String(error))
    }
  })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(port, '127.0.0.1', () => resolve())
  })
  const address = server.address()
  const bound = typeof address === 'object' && address ? address.port : port
  return {
    port: bound,
    url: name => `http://127.0.0.1:${bound}/${encodeURIComponent(name)}`,
    close: () => new Promise(resolve => server.close(() => resolve())),
  }
}
