import { expect, it, vi } from 'vitest'
import { parseXml } from '../src/core/xml'
import { ensureCf, ensureCfParser, hasConditionalFormatting, CF_PRESENCE_EXTENSION_URI } from '../src/xlsx/cfLoader'
import { CF_X14_EXTENSION_URI } from '../src/xlsx/cf/parse'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { EAGER_FORBIDDEN, PAGED_WORKER_FORBIDDEN, eagerModules, type ModuleMap } from './helpers/eager-graph'
import { closureModules, moduleSetHash } from './helpers/core-gates'
it('no CF workbook imports no CF implementation or allocates a CF cache', async () => {
  const load = vi.fn(async () => ({}))
  const sheets = [{ name: 'plain', conditionalFormatting: undefined }]
  expect(await ensureCf(sheets, load)).toBeUndefined()
  expect(load).not.toHaveBeenCalled(); expect(Object.keys(sheets[0])).toEqual(['name', 'conditionalFormatting'])
  expect(await ensureCf([{ conditionalFormatting: { rules: [], diagnostics: [] } }], load)).toEqual({})
  expect(load).toHaveBeenCalledTimes(1)
})
it('excludes every CF implementation from both ordinary static closures', () => {
  for (const rules of [EAGER_FORBIDDEN, PAGED_WORKER_FORBIDDEN]) for (const file of ['model', 'limits', 'evidence', 'parse', 'capture', 'dxf', 'predicates', 'dataBar', 'dataBarPaint', 'iconSet', 'iconPaint', 'iconGlyphs', 'formula', 'dependencies']) {
    expect(rules.some(rule => rule.test(`/repo/src/xlsx/cf/${file}.ts`))).toBe(true)
  }
})
it('importing the engine from the cf chunk leaves the index closure, and so indexModuleSetHash, unchanged', () => {
  // Synthetic: the cf chunk is lazy and imports engine modules the index closure already holds.
  const engine = ['/r/src/xlsx/formula/parser.ts', '/r/src/xlsx/formula/functions.ts', '/r/src/xlsx/formula/serial.ts']
  const withoutCf: ModuleMap = {
    'index.js': { modules: ['/r/src/index.ts'], imports: ['engine.js'], dynamicImports: [] },
    'engine.js': { modules: engine, imports: [], dynamicImports: [] },
  }
  const withCf: ModuleMap = {
    ...withoutCf,
    'index.js': { modules: ['/r/src/index.ts'], imports: ['engine.js'], dynamicImports: ['cf.js'] },
    'cf.js': { modules: ['/r/src/xlsx/cf/formula.ts', '/r/src/xlsx/cf/parse.ts'], imports: ['engine.js'], dynamicImports: [] },
  }
  expect(moduleSetHash(closureModules(withCf, 'index.js', '/r/'))).toBe(moduleSetHash(closureModules(withoutCf, 'index.js', '/r/')))
  // The paged worker's roots likewise never reach it statically.
  expect([...eagerModules(withCf)].some(id => id.includes('/xlsx/cf/'))).toBe(false)
})
// A stale `dist/` (built before the formula module existed) has no chunk holding it: skipped, never red.
const builtMap = (): ModuleMap | undefined => {
  const file = join(process.cwd(), 'dist/.module-map.json')
  if (!existsSync(file)) return undefined
  const map = JSON.parse(readFileSync(file, 'utf8')) as ModuleMap
  return Object.values(map).some(chunk => chunk.modules.some(id => /src[\\/]xlsx[\\/]cf[\\/]formula\.ts$/.test(id))) ? map : undefined
}
it.skipIf(builtMap() === undefined)('built: the index closure holds no CF module, and every engine module the cf chunk imports is already in it', () => {
  const map = builtMap()!
  const index = eagerModules(map)
  const norm = (id: string) => id.replace(/\\/g, '/')
  expect([...index].filter(id => /src\/xlsx\/cf\//.test(norm(id)))).toEqual([])
  const holders = Object.keys(map).filter(file => map[file]!.modules.some(id => /src\/xlsx\/cf\/formula\.ts$/.test(norm(id))))
  expect(holders.length).toBeGreaterThan(0)
  for (const holder of holders) {
    expect(eagerModules(map, holder), holder).toBeDefined()
    // Everything the cf chunk statically adds from the engine is a module of the index closure already.
    for (const id of eagerModules(map, holder)) if (/src\/xlsx\/formula\//.test(norm(id))) expect(index.has(id), id).toBe(true)
  }
})
it('shares one ext URI between the eager presence check and the lazy parser', () => {
  expect(CF_PRESENCE_EXTENSION_URI).toBe(CF_X14_EXTENSION_URI)
})
it('loads the rule parser only when CF content is present', async () => {
  const load = vi.fn(async () => ({}))
  expect(await ensureCfParser(false, load)).toBeUndefined()
  expect(load).not.toHaveBeenCalled()
  expect(await ensureCfParser(true, load)).toEqual({})
  expect(load).toHaveBeenCalledTimes(1)
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  const plain = parseXml(`<worksheet xmlns="${NS}"><sheetData/></worksheet>`)
  expect(hasConditionalFormatting(plain)).toBe(false)
  const base = parseXml(`<worksheet xmlns="${NS}"><conditionalFormatting sqref="A1"/></worksheet>`)
  expect(hasConditionalFormatting(base)).toBe(true)
  const x14 = parseXml(
    `<worksheet xmlns="${NS}"><extLst><ext uri="${CF_PRESENCE_EXTENSION_URI}"><conditionalFormattings xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/></ext></extLst></worksheet>`,
  )
  expect(hasConditionalFormatting(x14)).toBe(true)
  expect(hasConditionalFormatting(undefined)).toBe(false)
})
