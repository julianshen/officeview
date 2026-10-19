/**
 * Eager import-graph gate.
 *
 *   OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts
 *
 * Fails if the static-import closure of dist/index.js contains a forbidden
 * module (lucide, motion, the paged chrome, ...; see EAGER_FORBIDDEN).
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { eagerModules, forbiddenEager, EAGER_FORBIDDEN, type ModuleMap } from '../tests/helpers/eager-graph'

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
console.log('eager graph ok')
