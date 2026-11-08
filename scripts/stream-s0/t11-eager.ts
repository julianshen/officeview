/**
 * T11 (Stage-0 subset, scope call 1): run the §8.1 gates against the CURRENT
 * build and publish the numbers. Building `officeview/core` and its purity
 * assert are Stage 1.
 *
 *   bun scripts/stream-s0/t11-eager.ts
 *
 * Gate 1 (§8.1) = gzip bytes summed over the CHUNK FILES in dist/index.js's
 * static-import closure (index.js + every chunk it statically imports,
 * transitively; dynamic imports are lazy and excluded), ceiling 38,720 B.
 * Today's `scripts/eager-budget.ts` gzips dist/index.js alone — both numbers
 * are recorded so the difference is visible.
 */
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { gzipSync } from 'node:zlib'
import type { ModuleMap } from '../../tests/helpers/eager-graph'
import { writeResult } from './results'

const CEILING = 38720
const DIST = join(import.meta.dir, '..', '..', 'dist')

/** Chunk files in `entry`'s static-import closure (entry included). */
export function staticChunkClosure(map: ModuleMap, entry = 'index.js'): string[] {
  const seen = new Set<string>()
  const queue = [entry]
  while (queue.length > 0) {
    const file = queue.pop()!
    if (seen.has(file) || !map[file]) continue
    seen.add(file)
    queue.push(...map[file].imports)
  }
  return [...seen].sort()
}

function main(): number {
  // Same build the eager gates use; stdout kept out of the result.
  execFileSync('bun', ['run', 'build'], { stdio: 'inherit', env: { ...process.env, OV_MODULE_MAP: '1' } })
  const mapPath = join(DIST, '.module-map.json')
  if (!existsSync(mapPath)) throw new Error('dist/.module-map.json missing after OV_MODULE_MAP=1 build')
  const map = JSON.parse(readFileSync(mapPath, 'utf8')) as ModuleMap
  const chunks = staticChunkClosure(map)
  const perChunk = chunks.map((file) => ({ file, gzipBytes: gzipSync(readFileSync(join(DIST, file))).length }))
  const transitiveGzipBytes = perChunk.reduce((s, c) => s + c.gzipBytes, 0)
  const indexOnlyGzipBytes = gzipSync(readFileSync(join(DIST, 'index.js'))).length
  // Identity of the main bundle's module set (order-independent) for Stage 1's map-identity gate.
  const entry = map['index.js']!
  const moduleMapHash = createHash('sha256')
    .update(JSON.stringify({ modules: [...entry.modules].sort(), imports: [...entry.imports].sort() }))
    .digest('hex')
  let checkEager: string
  let checkEagerOk = true
  try {
    checkEager = execFileSync('bun', ['scripts/check-eager.ts'], { encoding: 'utf8' })
  } catch (e) {
    const err = e as { stdout?: string; stderr?: string }
    checkEager = `${err.stdout ?? ''}${err.stderr ?? ''}`
    checkEagerOk = false
  }
  const green = transitiveGzipBytes <= CEILING && checkEagerOk
  writeResult('t11', {
    status: green ? 'GREEN' : 'RED',
    ceiling: CEILING,
    transitiveGzipBytes,
    indexOnlyGzipBytes,
    perChunk,
    moduleMapHash,
    checkEager: checkEager.trim(),
    corePurity: 'stage1',
    note: 'Gate 1 sums gzip over the static chunk closure of dist/index.js; eager-budget.ts today measures index.js alone (indexOnlyGzipBytes). Core entry + purity assert are Stage 1.',
  })
  console.log(`[t11] ${green ? 'GREEN' : 'RED'} closure ${transitiveGzipBytes} B over ${chunks.length} chunk(s) (index.js alone ${indexOnlyGzipBytes} B), ceiling ${CEILING}`)
  return green ? 0 : 3
}

if (import.meta.main) process.exit(main())
