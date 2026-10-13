// @vitest-environment node
import { mkdtemp, mkdir, readFile, readdir, realpath, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { build } from 'vite'
import { expect, it } from 'vitest'
import libraryConfig from '../vite.config'

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
      /typeof window\s*>\s*"u"\s*&&\s*typeof\s+([\w$]+)\.postMessage\s*==\s*"function"\s*&&\s*\(\s*\1\.onmessage\s*=/,
    )

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
