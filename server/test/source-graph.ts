/** Static-import closure over TypeScript sources (dynamic imports followed
 * too — the server bundles nothing, so every reachable module counts). */
import { existsSync, readFileSync, statSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'

const transpiler = new Bun.Transpiler({ loader: 'tsx' })
const EXT = ['.ts', '.tsx', '/index.ts', '/index.tsx', '.js']

function resolveLocal(from: string, spec: string): string | undefined {
  const base = resolve(dirname(from), spec)
  if (existsSync(base) && statSync(base).isFile()) return base
  for (const e of EXT) if (existsSync(base + e)) return base + e
  return undefined
}

export function sourceClosure(entry: string, root: string): { files: string[]; packages: string[] } {
  const files = new Set<string>(), packages = new Set<string>()
  const queue = [resolve(entry)]
  while (queue.length) {
    const f = queue.pop()!
    if (files.has(f)) continue
    files.add(f)
    if (!/\.(ts|tsx|js)$/.test(f)) continue
    for (const imp of transpiler.scanImports(readFileSync(f, 'utf8'))) {
      if (imp.path.startsWith('.')) { const r = resolveLocal(f, imp.path); if (r) queue.push(r) }
      else packages.add(imp.path)
    }
  }
  return { files: [...files].map((f) => relative(root, f)).sort(), packages: [...packages].sort() }
}
