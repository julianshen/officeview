import { expect, it } from 'vitest'
import { closureModules, engineModulesIn, moduleDiff, moduleSetHash, purityViolations, staticClosure } from '../helpers/core-gates'
import type { ModuleMap } from '../helpers/eager-graph'

const ROOT = '/repo/'
const map: ModuleMap = {
  'core.js': { modules: [`${ROOT}src/core-entry.ts`, `${ROOT}src/stream/source.ts`], imports: ['core-a.js'], dynamicImports: ['core-lazy.js'] },
  'core-a.js': { modules: [`${ROOT}src/stream/crc32.ts`], imports: [], dynamicImports: [] },
  'core-lazy.js': { modules: [`${ROOT}src/xlsx/parse.ts`], imports: [], dynamicImports: [] },
  'index.js': { modules: [`${ROOT}src/index.ts`], imports: ['engine.js'], dynamicImports: [] },
  'engine.js': { modules: [`${ROOT}src/drawing/paint.ts`, `${ROOT}node_modules/jszip/lib/index.js`], imports: [], dynamicImports: [] },
}

it('staticClosure follows imports, not dynamicImports', () => {
  expect(staticClosure(map, 'core.js')).toEqual(['core-a.js', 'core.js'])
  expect(staticClosure(map, 'index.js')).toEqual(['engine.js', 'index.js'])
})

it('engineModulesIn names engine modules in the given chunks only', () => {
  expect(engineModulesIn(map, ['core.js', 'core-a.js'], ROOT)).toEqual([])
  expect(engineModulesIn(map, ['engine.js'], ROOT)).toEqual(['src/drawing/paint.ts'])
  // a dynamic-only engine chunk is not in the static closure
  expect(engineModulesIn(map, staticClosure(map, 'core.js'), ROOT)).toEqual([])
})

it('purityViolations flags React value imports and document use', () => {
  expect(purityViolations('import{a}from"react";')).toEqual(['import "react"'])
  expect(purityViolations('import { jsx } from "react/jsx-runtime"')).toEqual(['import "react/jsx-runtime"'])
  expect(purityViolations('export{x}from"react-dom"')).toEqual(['import "react-dom"'])
  expect(purityViolations('const m=import("react-dom/client")')).toEqual(['import "react-dom/client"'])
  expect(purityViolations('const c=document.createElement("canvas")')).toEqual(['document global'])
})

it('purityViolations ignores strings, comments and look-alikes', () => {
  expect(purityViolations('const s="document"; // document.body\n/* import("react") */')).toEqual([])
  expect(purityViolations("const t='react'; const u=`document`")).toEqual([])
  expect(purityViolations('const documentation=1; obj.document=2; import x from "preact"')).toEqual([])
})

it('closureModules is repo-relative over the static closure; moduleSetHash is order-independent', () => {
  expect(closureModules(map, 'core.js', ROOT)).toEqual(['src/core-entry.ts', 'src/stream/crc32.ts', 'src/stream/source.ts'])
  expect(closureModules(map, 'index.js', '/elsewhere/')[0]).toMatch(/^\/repo\//)
  expect(moduleSetHash(['b', 'a'])).toBe(moduleSetHash(['a', 'b']))
  expect(moduleSetHash(['a'])).toMatch(/^[0-9a-f]{64}$/)
})

it('closureModules relativizes rollup virtual (NUL-prefixed) ids', () => {
  const m: ModuleMap = { 'index.js': { modules: ['\u0000/repo/node_modules/jszip/dist/jszip.min.js?commonjs-module'], imports: [], dynamicImports: [] } }
  expect(closureModules(m, 'index.js', ROOT)).toEqual(['virtual:node_modules/jszip/dist/jszip.min.js?commonjs-module'])
})

it('moduleDiff reports added and removed ids', () => {
  expect(moduleDiff(['a', 'b'], ['b', 'c'])).toEqual({ added: ['c'], removed: ['a'] })
})
