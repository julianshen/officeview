// @vitest-environment node
import { mkdtemp, mkdir, readFile, readdir, realpath, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { build } from 'vite'
import { expect, it } from 'vitest'
import libraryConfig from '../vite.config'
import { chunkCycles, chunkGraph, importersOf } from './helpers/chunk-graph'

it('keeps the README worker bootstrap in a Vite production consumer (worker.format es) while tree-shaking the unused main entry', async () => {
  const root = fileURLToPath(new URL('..', import.meta.url))
  // Canonical paths avoid /tmp -> /private/tmp mismatches in Vite's HTML entry.
  const temporary = await realpath(await mkdtemp(join(tmpdir(), 'officeview-worker-package-')))
  try {
    const pkg = join(temporary, 'package')
    await mkdir(pkg)
    await writeFile(join(pkg, 'package.json'), await readFile(join(root, 'package.json')))
    const dependencies = dirname(dirname(fileURLToPath(import.meta.resolve('react/package.json'))))
    // 'junction' needs no elevated rights on Windows and is a plain symlink elsewhere.
    await symlink(dependencies, join(pkg, 'node_modules'), 'junction')
    await build({
      ...libraryConfig,
      configFile: false,
      root,
      logLevel: 'silent',
      build: { ...libraryConfig.build, outDir: join(pkg, 'dist') },
    })

    // The library's own output: nothing imports the worker entry, and no cycle (stream F item 5).
    const libraryFiles = await readdir(join(pkg, 'dist'))
    const libraryGraph = await chunkGraph(join(pkg, 'dist'), libraryFiles)
    expect(importersOf(libraryGraph, 'worker.js')).toEqual([])
    expect(chunkCycles(libraryGraph).filter(cycle => cycle.includes('worker.js'))).toEqual([])

    const consumer = join(temporary, 'consumer')
    await mkdir(join(consumer, 'node_modules'), { recursive: true })
    await symlink(pkg, join(consumer, 'node_modules/officeview'), 'junction')
    await symlink(join(dependencies, 'react'), join(consumer, 'node_modules/react'), 'junction')
    await symlink(join(dependencies, 'react-dom'), join(consumer, 'node_modules/react-dom'), 'junction')
    await writeFile(join(consumer, 'index.html'), '<script type="module" src="/main.ts"></script>')
    // The README factory, with a named import whose binding is never used, so the
    // main entry must tree-shake away while the worker entry keeps its side effect.
    await writeFile(join(consumer, 'main.ts'), [
      "import { OfficeFile } from 'officeview'",
      "void OfficeFile",
      "new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })",
      '',
    ].join('\n'))
    await writeFile(join(consumer, 'viewer-worker.ts'), "import 'officeview/worker'\n")
    const consumerConfig = { configFile: false as const, root: consumer, logLevel: 'silent' as const, build: { minify: false } }

    await build({ ...consumerConfig, worker: { format: 'es' } })

    const assets = join(consumer, 'dist/assets')
    const files = await readdir(assets)
    const worker = files.find(name => /^viewer-worker-.*\.js$/.test(name))
    expect(worker, 'The module Worker must have an emitted entry').toBeDefined()
    const workerCode = await readFile(join(assets, worker!), 'utf8')
    // The exact bootstrap statement from src/worker/worker-entry.ts
    // (`if (typeof window === 'undefined' && typeof scope.postMessage === 'function') scope.onmessage = ...`)
    // as the library build ships it: dist/worker.js is minified by the library build and the
    // consumer (minify: false) keeps that spelling, so only the minified form is pinned.
    expect(workerCode, 'The README production worker entry must install its message handler').toMatch(
      // Stream F item 5: the bootstrap is a block now (queue + lazy router), so the
      // minifier keeps the `if (...) { ... onmessage = ... }` statement.
      /typeof window\s*>\s*"u"\s*&&\s*typeof\s+([\w$]+)\.postMessage\s*==\s*"function"\s*(?:&&\s*\(|\)\s*\{[^}]*?)\1\.onmessage\s*=/,
    )

    // Stream F item 5 (WebKit): a lazy chunk of the worker graph must never import
    // the worker entry chunk. WebKit does not register a module worker's top-level
    // script in the worker's module map, so such an import evaluates the entry a
    // second time and its bootstrap replaces `onmessage` with a fresh, empty kernel
    // ("Unknown documentId"). Pinned on the consumer graph, where the host's bundler
    // re-chunks the library's modules.
    const graph = await chunkGraph(assets, files)
    expect(importersOf(graph, worker!), 'no chunk may import the module worker entry').toEqual([])
    // Lazy chunks may import each other (ordinary modules in the module map); only
    // a cycle through the entry is the WebKit hazard.
    expect(chunkCycles(graph).filter(cycle => cycle.includes(worker!))).toEqual([])

    const main = files.find(name => /^index-.*\.js$/.test(name))
    expect(main).toBeDefined()
    const mainCode = await readFile(join(assets, main!), 'utf8')
    // The unused named import is tree-shaken: no officeview code or chunk reaches main.
    expect(mainCode).not.toMatch(/OfficeFile|OfficeDoc|interactiveSheets/)
    // Other emitted chunks belong to the worker graph; main must not load any of them.
    expect(mainCode).not.toMatch(/(?:from\s*|import\s*\(\s*)["']\.\/[^"']+\.js["']/)

    // README requirement: Vite's default worker format (iife) cannot code-split the
    // kept worker graph, so the documented Vite setup sets worker.format to 'es'.
    await expect(build(consumerConfig)).rejects.toThrow(/worker\.format/)
  } finally {
    await rm(temporary, { recursive: true, force: true })
  }
}, 60_000)

it('builds the perf page worker without a chunk importing its entry (stream F item 5: WebKit re-evaluated it)', async () => {
  const temporary = await realpath(await mkdtemp(join(tmpdir(), 'officeview-perf-build-')))
  try {
    const { default: perfConfig } = await import('../vite.perf.config')
    const outDir = join(temporary, 'dist')
    await build({ ...perfConfig, configFile: false, logLevel: 'silent', cacheDir: join(temporary, 'cache'), build: { ...perfConfig.build, outDir } })
    const assets = join(outDir, 'assets'), files = await readdir(assets)
    const worker = files.find(name => /^viewerPerfWorker-.*\.js$/.test(name))
    expect(worker).toBeDefined()
    const graph = await chunkGraph(assets, files)
    expect(importersOf(graph, worker!)).toEqual([])
    expect(chunkCycles(graph).filter(cycle => cycle.includes(worker!))).toEqual([])
  } finally {
    await rm(temporary, { recursive: true, force: true })
  }
}, 120_000)
