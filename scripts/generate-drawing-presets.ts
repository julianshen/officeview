/** Run: bun scripts/generate-drawing-presets.ts /local/presetShapeDefinitions.xml [output.json] */
import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { orderedChildren, parseXmlOrdered } from '../src/core/xml'
import type { GeometryDefinition } from '../src/drawing/geometry'
import { parseCustomGeometry } from '../src/drawing/geometry-xml'

const PINNED_SHA256 = 'a7dad593d27bd70536b41da9b761fa16409536cc0c25ef2b6c7a61c5d9b3e738'
const input = process.argv[2]
if (!input) throw new Error('Provide the local Apache POI REL_5_4_1 presetShapeDefinitions.xml input; generation never fetches network data')
const bytes = readFileSync(input)
const actualHash = createHash('sha256').update(bytes).digest('hex')
if (actualHash !== PINNED_SHA256) throw new Error(`Pinned preset input SHA-256 mismatch: expected ${PINNED_SHA256}; received ${actualHash}`)
const root = parseXmlOrdered(bytes.toString('utf8'))
const presets: Record<string, GeometryDefinition> = {}
for (const [name, node] of orderedChildren(root)) {
  if (name === '#text') continue
  if (Object.prototype.hasOwnProperty.call(presets, name)) throw new Error(`Duplicate preset ${name}`)
  const definition = parseCustomGeometry(node)
  // These eight exact formulas in the pinned POI source have a redundant
  // fourth zero operand. Keep the runtime evaluator strict for custom data.
  const circular = ['circularArrow', 'leftCircularArrow', 'leftRightCircularArrow'].includes(name)
  for (const guide of definition.guides) {
    const expected = guide[0] === 'xB' ? '+- xH 0 dxB 0' : guide[0] === 'yB' ? '+- yH 0 dyB 0' : name === 'leftRightCircularArrow' && guide[0] === 'xJ' ? '+- xI 0 dxJ 0' : name === 'leftRightCircularArrow' && guide[0] === 'yJ' ? '+- yI 0 dyJ 0' : undefined
    if (circular && expected && guide[1] === expected) guide[1] = expected.slice(0, -2)
  }
  presets[name] = definition
}
if (Object.keys(presets).length !== 187) throw new Error('Pinned input must contain exactly 187 presets')
const output = process.argv[3] ? resolve(process.argv[3]) : fileURLToPath(new URL('../src/drawing/presets.json', import.meta.url))
writeFileSync(output, `${JSON.stringify(presets, null, 2)}\n`)
console.log(`Normalized 187 drawing presets to ${output}`)
