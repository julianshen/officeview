/**
 * Vite plugin: writes dist/.module-map.json, the chunk graph with module
 * provenance, so scripts/check-eager.ts can audit what the eager entry pulls
 * in. Enabled in vite.config.ts only when OV_MODULE_MAP is set.
 */
import type { Plugin } from 'vite'
import type { ModuleMap } from '../tests/helpers/eager-graph'

export function moduleMap(): Plugin {
  return {
    name: 'officeview-module-map',
    generateBundle(_options, bundle) {
      const map: ModuleMap = {}
      for (const [file, out] of Object.entries(bundle)) {
        if (out.type !== 'chunk') continue
        map[file] = { modules: Object.keys(out.modules), imports: [...out.imports], dynamicImports: [...out.dynamicImports] }
      }
      this.emitFile({ type: 'asset', fileName: '.module-map.json', source: JSON.stringify(map, null, 2) })
    },
  }
}
