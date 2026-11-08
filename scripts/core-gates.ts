/**
 * The §8.1 eager gates (spec §2 item 14b ruling).
 *
 *   bun run gates:core                      # OV_MODULE_MAP=1 build, then check
 *   bun scripts/core-gates.ts --baseline <out.json>   # capture a baseline from ./dist
 *
 * 1. dist/core.js static closure ≤ 38,720 B gzip and reaches no engine module.
 * 2. Purity: no React value import and no `document` global in that closure.
 * 3. Main entry: dist/index.js static closure gzip ≤ baseline + recorded
 *    changes + 16 B, and its module set equals the baseline ± recorded changes.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { gzipSync } from 'node:zlib'
import { closureModules, engineModulesIn, moduleDiff, moduleSetHash, purityViolations, staticClosure } from '../tests/helpers/core-gates'
import type { ModuleMap } from '../tests/helpers/eager-graph'

export const CORE_CEILING = 38720

interface RecordedChange { reason: string; added: string[]; removed: string[]; gzipDelta: number }
interface Baseline {
  source: string
  indexClosureGzip: number
  indexModules: string[]
  indexModuleSetHash: string
  allowedGrowth: number
  recordedChanges: RecordedChange[]
}

const ROOT = `${process.cwd().replace(/\\/g, '/')}/`
const DIST = join(process.cwd(), 'dist')

function readMap(file: string): ModuleMap {
  const at = join(DIST, file)
  if (!existsSync(at)) throw new Error(`dist/${file} missing — run with OV_MODULE_MAP=1 bun run build first`)
  return JSON.parse(readFileSync(at, 'utf8')) as ModuleMap
}
const gzipOf = (chunks: string[]) => chunks.reduce((sum, file) => sum + gzipSync(readFileSync(join(DIST, file))).length, 0)

function indexMeasure(): { gzip: number; modules: string[] } {
  const map = readMap('.module-map.json')
  return { gzip: gzipOf(staticClosure(map, 'index.js')), modules: closureModules(map, 'index.js', ROOT) }
}

function writeBaseline(out: string): void {
  const { gzip, modules } = indexMeasure()
  const baseline: Baseline = { source: process.argv[process.argv.indexOf('--baseline') + 2] ?? 'unspecified', indexClosureGzip: gzip, indexModules: modules, indexModuleSetHash: moduleSetHash(modules), allowedGrowth: 16, recordedChanges: [] }
  writeFileSync(out, `${JSON.stringify(baseline, null, 2)}\n`)
  console.log(`baseline: index closure ${gzip} B, ${modules.length} modules → ${out}`)
}

function check(): number {
  let failed = 0
  const fail = (line: string) => { failed++; console.error(`FAIL ${line}`) }

  const coreMap = readMap('.module-map.core.json')
  const coreChunks = staticClosure(coreMap, 'core.js')
  const coreGzip = gzipOf(coreChunks)
  const engine = engineModulesIn(coreMap, coreChunks, ROOT)
  if (coreGzip > CORE_CEILING) fail(`core closure ${coreGzip} B over ${coreChunks.length} chunk(s) > ${CORE_CEILING}`)
  else console.log(`ok   core closure ${coreGzip} B over ${coreChunks.length} chunk(s) ≤ ${CORE_CEILING}`)
  if (engine.length) fail(`engine modules in core closure:\n${engine.map((id) => `       ${id}`).join('\n')}`)
  else console.log('ok   engine modules in core closure: none')

  const impure = coreChunks.flatMap((file) => purityViolations(readFileSync(join(DIST, file), 'utf8')).map((v) => `${file}: ${v}`))
  if (impure.length) fail(`core purity:\n${impure.map((v) => `       ${v}`).join('\n')}`)
  else console.log('ok   core purity: no react imports, no document global')

  const baseline = JSON.parse(readFileSync(join(import.meta.dir, 'eager-baseline.json'), 'utf8')) as Baseline
  const { gzip, modules } = indexMeasure()
  const recordedDelta = baseline.recordedChanges.reduce((s, c) => s + c.gzipDelta, 0)
  const delta = gzip - baseline.indexClosureGzip
  const limit = baseline.indexClosureGzip + recordedDelta + baseline.allowedGrowth
  if (gzip > limit) fail(`index closure ${gzip} B (baseline ${baseline.indexClosureGzip}, Δ ${delta}, recorded ${recordedDelta}, limit ${limit})`)
  else console.log(`ok   index closure ${gzip} B (baseline ${baseline.indexClosureGzip}, Δ ${delta}, recorded ${recordedDelta})`)
  const expected = new Set(baseline.indexModules)
  for (const c of baseline.recordedChanges) { for (const id of c.added) expected.add(id); for (const id of c.removed) expected.delete(id) }
  const diff = moduleDiff([...expected].sort(), modules)
  if (diff.added.length || diff.removed.length) {
    fail(`index module set differs from baseline ± recorded changes:\n${[...diff.added.map((id) => `       + ${id}`), ...diff.removed.map((id) => `       - ${id}`)].join('\n')}`)
  } else console.log(`ok   index module set = baseline ± ${baseline.recordedChanges.length} recorded change(s) (${modules.length} modules)`)
  return failed ? 1 : 0
}

if (import.meta.main) {
  const at = process.argv.indexOf('--baseline')
  if (at !== -1) writeBaseline(process.argv[at + 1]!)
  else process.exit(check())
}
