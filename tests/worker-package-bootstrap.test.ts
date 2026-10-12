// @vitest-environment node
import { mkdtemp, mkdir, readFile, readdir, realpath, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { build } from 'vite'
import { expect, it } from 'vitest'
import libraryConfig from '../vite.config'

it('keeps the README worker bootstrap in a production consumer while tree-shaking an unused main entry', async () => {
  const root = fileURLToPath(new URL('..', import.meta.url))
  // Canonical paths avoid /tmp -> /private/tmp mismatches in Vite's HTML entry.
  const temporary = await realpath(await mkdtemp(join(tmpdir(), 'officeview-worker-package-')))
  try {
    const pkg = join(temporary, 'package')
    await mkdir(pkg)
    await writeFile(join(pkg, 'package.json'), await readFile(join(root, 'package.json')))
    const dependencies = dirname(dirname(fileURLToPath(import.meta.resolve('react/package.json'))))
    await symlink(dependencies, join(pkg, 'node_modules'), 'dir')
    await build({
      ...libraryConfig,
      configFile: false,
      root,
      logLevel: 'silent',
      build: { ...libraryConfig.build, outDir: join(pkg, 'dist') },
    })

    const consumer = join(temporary, 'consumer')
    await mkdir(join(consumer, 'node_modules'), { recursive: true })
    await symlink(pkg, join(consumer, 'node_modules/officeview'), 'dir')
    await symlink(join(dependencies, 'react'), join(consumer, 'node_modules/react'), 'dir')
    await symlink(join(dependencies, 'react-dom'), join(consumer, 'node_modules/react-dom'), 'dir')
    await writeFile(join(consumer, 'index.html'), '<script type="module" src="/main.js"></script>')
    await writeFile(join(consumer, 'main.js'), "import 'officeview'\nnew Worker(new URL('./viewer-worker.js', import.meta.url), { type: 'module' })\n")
    await writeFile(join(consumer, 'viewer-worker.js'), "import 'officeview/worker'\n")
    await build({
      configFile: false,
      root: consumer,
      logLevel: 'silent',
      build: { minify: false },
      worker: { format: 'es' },
    })

    const assets = join(consumer, 'dist/assets')
    const files = await readdir(assets)
    const worker = files.find(name => /^viewer-worker-.*\.js$/.test(name))
    expect(worker, 'The module Worker must have an emitted entry').toBeDefined()
    const workerCode = await readFile(join(assets, worker!), 'utf8')
    expect(workerCode, 'The README production worker entry must install its message handler').toMatch(/\.onmessage\s*=/)
    expect(workerCode).toMatch(/typeof window/)
    expect(workerCode).toMatch(/\.postMessage\(/)

    const main = files.find(name => /^index-.*\.js$/.test(name))
    expect(main).toBeDefined()
    const mainCode = await readFile(join(assets, main!), 'utf8')
    // Only Vite's preload helper and the Worker factory should remain in main.
    expect(Buffer.byteLength(mainCode)).toBeLessThan(2_000)
  } finally {
    await rm(temporary, { recursive: true, force: true })
  }
}, 60_000)
