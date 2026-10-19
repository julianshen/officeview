// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { createServer, isFileServingAllowed } from 'vite'
import { resolve, join } from 'node:path'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { createServer as createHttpServer } from 'node:http'

it('serves the lazy print entry only after warming the paged harness, without opening the source tree', async () => {
  const transport = createHttpServer() // Unbound: Vite attaches upgrade handlers only.
  const cacheDir = await mkdtemp(join(tmpdir(), 'ov-print-harness-'))
  const server = await createServer({
    cacheDir,
    configFile: resolve('vite.harness.config.ts'),
    server: { middlewareMode: true, hmr: { server: transport }, watch: null },
    optimizeDeps: { noDiscovery: true, include: [] },
  })
  try {
    const file = resolve('src/components/paged/printDocument.ts')
    expect(isFileServingAllowed(`/@fs/${file}`, server)).toBe(false)
    await expect(server.transformRequest('/src/components/paged/printDocument.ts')).rejects.toMatchObject({ code: 'ERR_LOAD_URL' })
    const pending: Promise<void>[] = []
    const warm = server.warmupRequest.bind(server)
    vi.spyOn(server, 'warmupRequest').mockImplementation(url => {
      const promise = warm(url)
      pending.push(promise)
      return promise
    })
    await server.transformRequest('/src/debug/viewer-e2e.tsx')
    await Promise.all(pending)
    // Warm-up caches a transform; it intentionally does not open fs.allow.
    const print = await server.transformRequest('/src/components/paged/printDocument.ts')
    expect(print?.code).toContain('printDocument')
    expect(isFileServingAllowed(`/@fs/${resolve('package.json')}`, server)).toBe(false)
  } finally {
    await server.close()
    vi.restoreAllMocks()
    await rm(cacheDir, { recursive: true, force: true })
  }
}, 30_000)
