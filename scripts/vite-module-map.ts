/**
 * Vite plugin: writes dist/.module-map.json, the chunk graph with module
 * provenance, so scripts/check-eager.ts can audit what the eager entry pulls
 * in. Enabled in vite.config.ts only when OV_MODULE_MAP is set.
 */
import { gzipSync } from 'node:zlib'
import type { Plugin } from 'vite'
import type { ModuleMap } from '../tests/helpers/eager-graph'

export function moduleMap(fileName = '.module-map.json'): Plugin {
  return {
    name: 'officeview-module-map',
    generateBundle(_options, bundle) {
      const map: ModuleMap = {}
      for (const [file, out] of Object.entries(bundle)) {
        if (out.type !== 'chunk') continue
        // Each module's rendered code size in its chunk (raw and gzip on its own): Task 28 reads the glyph table's.
        const moduleBytes = Object.fromEntries(Object.entries(out.modules).map(([id, m]) => [id, { raw: m.renderedLength, gzip: m.code ? gzipSync(m.code).length : 0 }]))
        map[file] = { modules: Object.keys(out.modules), imports: [...out.imports], dynamicImports: [...out.dynamicImports], moduleBytes }
      }
      this.emitFile({ type: 'asset', fileName, source: JSON.stringify(map, null, 2) })
    },
  }
}
