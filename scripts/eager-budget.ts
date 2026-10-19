/**
 * Eager-bundle budget gate.
 *
 *   bun run build && bun scripts/eager-budget.ts
 *
 * Prints the Node-zlib gzip size of dist/index.js (the eager entry every host
 * downloads) and exits 1 above the ceiling. Always rebuild first: the
 * checked-in dist/ can be stale.
 */
import { readFileSync } from 'node:fs'
import { gzipSync } from 'node:zlib'
import { join } from 'node:path'

export const EAGER_BUDGET = 38720

const bytes = gzipSync(readFileSync(join(import.meta.dir, '..', 'dist', 'index.js'))).length
console.log(bytes)
if (bytes > EAGER_BUDGET) {
  console.error(`eager budget exceeded: ${bytes} B > ${EAGER_BUDGET} B`)
  process.exit(1)
}
