import { describe, it, expect } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { eagerModules, forbiddenEager, EAGER_FORBIDDEN, type ModuleMap } from './helpers/eager-graph'

const MOTION = '/node_modules/motion/dist/es/index.mjs'

// index.js statically imports a.js and lazy-loads b.js.
function synthetic(): ModuleMap {
  return {
    'index.js': { modules: ['/repo/src/index.ts', '/repo/src/components/OfficeDoc.tsx'], imports: ['a.js'], dynamicImports: ['b.js'] },
    'a.js': { modules: ['/repo/src/core/zoom.ts'], imports: [], dynamicImports: [] },
    'b.js': { modules: ['/repo/src/components/paged/PagedChrome.tsx', '/repo/src/components/viewerControls.tsx'], imports: ['index.js'], dynamicImports: [] },
  }
}

describe('eager import graph', () => {
  it('eagerModules follows static imports only', () => {
    const eager = eagerModules(synthetic())
    expect(eager.has('/repo/src/index.ts')).toBe(true)
    expect(eager.has('/repo/src/core/zoom.ts')).toBe(true)
    expect(eager.has('/repo/src/components/paged/PagedChrome.tsx')).toBe(false)
    expect(eager.has('/repo/src/components/viewerControls.tsx')).toBe(false)
  })

  it('flags a deliberate eager motion import (mutation)', () => {
    const eagerMotion = synthetic()
    eagerMotion['a.js'].modules.push(MOTION)
    expect(forbiddenEager(eagerMotion, EAGER_FORBIDDEN)).toEqual([MOTION])

    const lazyMotion = synthetic()
    lazyMotion['b.js'].modules.push(MOTION)
    expect(forbiddenEager(lazyMotion, EAGER_FORBIDDEN)).toEqual([])
  })

  it('forbids every listed eager module but allows paged/metrics.ts', () => {
    const hit = (id: string) => forbiddenEager({ 'index.js': { modules: [id], imports: [], dynamicImports: [] } }, EAGER_FORBIDDEN).length > 0
    for (const id of [
      '/r/node_modules/lucide-react/dist/esm/icons/x.js',
      '/r/node_modules/motion/dist/es/index.mjs',
      '/r/node_modules/framer-motion/dist/es/index.mjs',
      '/r/node_modules/motion-dom/dist/es/index.mjs',
      '/r/src/components/viewerControls.tsx',
      '/r/src/components/Icon.tsx',
      '/r/src/components/motion/Fade.tsx',
      '/r/src/styles/tokens.ts',
      '/r/src/components/paged/PagedChrome.tsx',
      'C:\\r\\src\\components\\paged\\PagedChrome.tsx',
    ]) expect(hit(id), id).toBe(true)
    for (const id of [
      '/r/src/components/paged/metrics.ts',
      '/r/src/components/OfficeDoc.tsx',
      '/r/src/components/xlsx/viewerControlsClient.ts',
    ]) expect(hit(id), id).toBe(false)
  })

  it('EAGER_FORBIDDEN rejects a synthetic module id under each new directory', () => {
    const hit = (id: string) => forbiddenEager({ 'index.js': { modules: [id], imports: [], dynamicImports: [] } }, EAGER_FORBIDDEN).length > 0
    for (const id of [
      '/r/src/slideshow/index.ts',
      '/r/src/slideshow/engine/scheduler.ts',
      '/r/src/pptx/appearance/index.ts',
      '/r/src/pptx/appearance/read.ts?worker',
      'C:\\r\\src\\slideshow\\index.ts',
    ]) expect(hit(id), id).toBe(true)
    // The eager PPTX parser and painter stay allowed.
    for (const id of ['/r/src/pptx/parse.ts', '/r/src/pptx/render.ts', '/r/src/render/paint.ts']) expect(hit(id), id).toBe(false)
  })

  it('rejects a statically imported sparkline painter while allowing the lazy module boundary', () => {
    const hit = (id: string) => forbiddenEager({ 'index.js': { modules: [id], imports: [], dynamicImports: [] } }, EAGER_FORBIDDEN).length > 0
    expect(hit('/repo/src/xlsx/sparklines/paint.ts')).toBe(true)
    expect(hit('/repo/src/xlsx/sparklines/geometry.ts')).toBe(true)
    for (const id of ['/repo/src/xlsx/sparklines/model.ts', '/repo/src/xlsx/sparklines/parseGroups.ts', '/repo/src/xlsx/sparklines/resolve.ts']) expect(hit(id), id).toBe(false)
    const renderImport: ModuleMap = {
      'index.js': { modules: ['/repo/src/index.ts'], imports: ['render.js'], dynamicImports: [] },
      'render.js': { modules: ['/repo/src/xlsx/render.ts', '/repo/src/xlsx/sparklines/paint.ts'], imports: [], dynamicImports: [] },
    }
    expect(forbiddenEager(renderImport, EAGER_FORBIDDEN)).toContain('/repo/src/xlsx/sparklines/paint.ts')
  })

  // The built module map (OV_MODULE_MAP=1 bun run build) is optional; the post-build check-eager gate always runs.
  const moduleMapFile = join(__dirname, '..', 'dist', '.module-map.json')
  it.skipIf(!existsSync(moduleMapFile))('src/slideshow and src/pptx/appearance are not in the eager graph (skipped: no dist/.module-map.json)', () => {
    // A present map must parse and must have the entry chunk.
    const map = JSON.parse(readFileSync(moduleMapFile, 'utf8')) as ModuleMap
    expect(map['index.js'], 'dist/.module-map.json has no index.js entry chunk').toBeDefined()
    expect(Array.isArray(map['index.js'].modules)).toBe(true)
    const eager = [...eagerModules(map)].map(id => id.replace(/\\/g, '/'))
    expect(eager.filter(id => /src\/slideshow\/|src\/pptx\/appearance\//.test(id))).toEqual([])
    expect(forbiddenEager(map, EAGER_FORBIDDEN)).toEqual([])
  })
})
