// @vitest-environment node
/** S10.2 perf-page static server: requests never escape the build directory. */
import { describe, expect, it } from 'vitest'
import { join } from 'node:path'
import { resolveStaticPath } from '../scripts/harness/static-server'
import { isCrossOriginIsolatedPath } from '../scripts/harness/vite-plugin'

describe('resolveStaticPath', () => {
  const root = '/srv/perf-dist'
  it('serves paths inside the root and refuses traversal or NUL', () => {
    expect(resolveStaticPath(root, '/src/debug/viewer-perf.html?x=1')).toBe(join(root, 'src/debug/viewer-perf.html'))
    expect(resolveStaticPath(root, '/assets/a.js')).toBe(join(root, 'assets/a.js'))
    // Dot segments are normalized by URL parsing: the result stays under the root.
    for (const url of ['/../secret', '/%2e%2e/%2e%2e/etc/passwd', '/a/../../../b']) {
      const path = resolveStaticPath(root, url)
      expect(path === null || path.startsWith(`${root}/`)).toBe(true)
    }
    expect(resolveStaticPath(root, '/a%00b')).toBeNull()
  })
})

describe('isCrossOriginIsolatedPath', () => {
  it('isolates only the perf page and its worker on the dev server', () => {
    expect(isCrossOriginIsolatedPath('/src/debug/viewer-perf.html')).toBe(true)
    expect(isCrossOriginIsolatedPath('/src/debug/viewerPerfWorker.ts?worker_file&type=module')).toBe(true)
    expect(isCrossOriginIsolatedPath('/src/debug/viewer-e2e.html')).toBe(false)
    expect(isCrossOriginIsolatedPath('/src/debug/interactive-viewer.html')).toBe(false)
  })
})
