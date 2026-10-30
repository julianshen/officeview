/**
 * Eager import-graph gate.
 *
 *   OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts
 *
 * Fails if the static-import closure of dist/index.js contains a forbidden
 * module (lucide, motion, the paged chrome, ...; see EAGER_FORBIDDEN), or if
 * a paged worker chunk statically imports the table-style built-ins or the
 * styles-extras parser (PAGED_WORKER_FORBIDDEN).
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { eagerModules, forbiddenEager, pagedWorkerViolations, EAGER_FORBIDDEN, type ModuleMap } from '../tests/helpers/eager-graph'

const file = join(import.meta.dir, '..', 'dist', '.module-map.json')
if (!existsSync(file)) {
  console.error('dist/.module-map.json missing; run OV_MODULE_MAP=1 bun run build first')
  process.exit(1)
}
const map = JSON.parse(readFileSync(file, 'utf8')) as ModuleMap
if (!map['index.js']) {
  console.error('dist/.module-map.json has no index.js entry chunk')
  process.exit(1)
}
const eager = eagerModules(map)
const bad = forbiddenEager(map, EAGER_FORBIDDEN)
console.log(`eager modules: ${eager.size}`)
if (bad.length > 0) {
  console.error(`forbidden modules in the eager graph:\n${bad.map((id) => `  ${id}`).join('\n')}`)
  process.exit(1)
}
const paged = pagedWorkerViolations(map)
if (paged.missing.length > 0) {
  console.error(`paged worker roots missing from the module map: ${paged.missing.join(', ')}`)
  process.exit(1)
}
if (paged.bad.length > 0) {
  console.error(`paged worker chunks statically import lazy style code:\n${paged.bad.map(({ chunk, module }) => `  ${chunk}: ${module}`).join('\n')}`)
  process.exit(1)
}
console.log(`paged worker chunks ok: ${paged.roots.join(', ')}`)
console.log('eager graph ok')
