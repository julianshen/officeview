/** Child process for render-determinism.test.ts: renders one unit in a fresh
 * process (fresh module state, fresh DocumentHost) and prints its PNG's SHA-256. */
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { DocumentHost, type MarkSpec } from '../src/render'

const [name, unit, scale, mark] = process.argv.slice(2) as [string, string, string, string]
const host = new DocumentHost()
const { docKey } = await host.open(name, new Uint8Array(readFileSync(resolve(import.meta.dir, '../../corpus', name))), 'visible-only')
const r = host.render(docKey, { unitIndex: Number(unit), scale: Number(scale), projection: 'final-raster', mark: JSON.parse(mark) as MarkSpec })
if (r.kind !== 'raster') throw new Error('expected raster')
process.stdout.write(createHash('sha256').update(r.png).digest('hex'))
