/**
 * Loopback static server for the production-built S10.2 perf page
 * (vite.perf.config.ts output). Every response is cross-origin isolated
 * (COOP same-origin + COEP require-corp) so Chrome exposes
 * performance.measureUserAgentSpecificMemory() to the page. Paths outside the
 * build directory are refused; nothing else in the checkout is served.
 */
import { createServer, type Server } from 'node:http'
import { readFile } from 'node:fs/promises'
import { extname, join, normalize, sep } from 'node:path'

const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.wasm': 'application/wasm',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
}

/** Resolves a request path inside `root`, or null when it would escape it. */
export function resolveStaticPath(root: string, urlPath: string): string | null {
  let decoded: string
  try { decoded = decodeURIComponent(new URL(urlPath, 'http://static.invalid').pathname) } catch { return null }
  if (decoded.includes('\0')) return null
  const full = normalize(join(root, decoded))
  const base = normalize(root.endsWith(sep) ? root : root + sep)
  return full.startsWith(base) ? full : null
}

export interface StaticServer { url: string; close(): Promise<void> }

export async function startStaticServer(root: string, port: number, host = '127.0.0.1'): Promise<StaticServer> {
  const server: Server = createServer((req, res) => {
    const path = resolveStaticPath(root, req.url ?? '/')
    if (!path) { res.writeHead(403).end(); return }
    readFile(path).then(body => {
      res.writeHead(200, {
        'content-type': TYPES[extname(path)] ?? 'application/octet-stream',
        'cache-control': 'no-cache',
        'cross-origin-opener-policy': 'same-origin',
        'cross-origin-embedder-policy': 'require-corp',
      })
      res.end(body)
    }, () => { res.writeHead(404).end() })
  })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(port, host, () => resolve())
  })
  return {
    url: `http://${host}:${port}`,
    close: () => new Promise(resolve => server.close(() => resolve())),
  }
}
