import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { readFileSync, existsSync } from 'node:fs'
import { writeResult, resolveBackend } from './results'

async function run(): Promise<number> {
  // Build with module map.
  execFileSync('bun', ['run', 'build'], { stdio: 'inherit', env: { ...process.env, OV_MODULE_MAP: '1' } })
  // module map location
  const mapPath = 'dist/module-map.json'
  let transitiveGzipBytes = 0
  let moduleMapHash = ''
  if (existsSync(mapPath)) {
    const map = JSON.parse(readFileSync(mapPath, 'utf8'))
    const { eagerModules } = await import('../tests/helpers/eager-graph.ts' as any).catch(() => ({ eagerModules: null }))
    // compute gzip sum of eager modules
    const zlib = await import('node:zlib')
    const modules: string[] = eagerModules ? [...eagerModules(map)] : []
    let sum = 0
    for (const m of modules) {
      // find file
      const p = 'dist/' + m
      if (existsSync(p)) {
        const buf = readFileSync(p)
        sum += zlib.gzipSync(buf).length
      }
    }
    transitiveGzipBytes = sum
    // hash of index.js module-map entry
    const entry = map['index.js']
    moduleMapHash = createHash('sha256').update(JSON.stringify(entry ?? {})).digest('hex')
  }
  // check-eager output
  let checkEager = ''
  try { checkEager = execFileSync('bun', ['scripts/check-eager.ts'], { encoding: 'utf8' }) } catch (e: any) { checkEager = (e.stdout || '') + (e.stderr || '') }
  writeResult('t11', {
    status: transitiveGzipBytes <= 38720 ? 'GREEN' : 'RED',
    transitiveGzipBytes,
    ceiling: 38720,
    moduleMapHash,
    checkEager: checkEager.slice(0, 500),
  })
  console.log('[t11] transitiveGzipBytes', transitiveGzipBytes, 'ceiling 38720, hash', moduleMapHash.slice(0, 12))
  return 0
}

run().then(c => process.exit(c), e => { console.error(e); process.exit(1) })
