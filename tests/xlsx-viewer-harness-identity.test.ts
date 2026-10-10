// @vitest-environment node
/**
 * Harness server identity (finding N3): the public identity marker must not
 * leak the absolute checkout path, and the LAN-mode vite server must deny
 * files outside the harness's minimum serving set while the page still loads.
 */
import { describe, expect, it } from 'vitest'
import { spawn, type ChildProcess } from 'node:child_process'
import { createServer } from 'node:net'
import { resolve } from 'node:path'
import { checkIdentity, gitIdentity, identityDigest } from '../scripts/harness/identity'

function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const probe = createServer()
    probe.on('error', reject)
    probe.listen(0, '127.0.0.1', () => {
      const port = (probe.address() as { port: number }).port
      probe.close(() => resolve(port))
    })
  })
}

describe('harness identity carries a digest, never an absolute path', () => {
  it('is stable for one checkout and differs across checkouts', () => {
    expect(identityDigest('/repo/a', 'abc123')).toBe(identityDigest('/repo/a', 'abc123'))
    expect(identityDigest('/repo/a', 'abc123')).not.toBe(identityDigest('/repo/b', 'abc123'))
    expect(identityDigest('/repo/a', 'abc123')).not.toBe(identityDigest('/repo/a', 'def456'))
    expect(identityDigest('/repo/a', 'abc123')).toMatch(/^[0-9a-f]{64}$/)
  })

  it('exposes no absolute path on the public marker', () => {
    const identity = gitIdentity(process.cwd())
    expect(identity).not.toHaveProperty('root')
    expect(JSON.stringify(identity)).not.toContain(process.cwd())
    expect(identity).toMatchObject({ service: 'officeview-viewer-harness' })
    expect(identity.rootDigest).toBe(identityDigest(resolve(process.cwd()), identity.commit))
  })

  it('accepts a matching marker and rejects stale, foreign or mismatched ones', () => {
    const identity = gitIdentity(process.cwd())
    expect(checkIdentity({ ...identity }, { commit: identity.commit, rootDigest: identity.rootDigest })).toBeNull()
    const stale = { ...identity, commit: 'deadbee' }
    expect(checkIdentity(stale, { commit: identity.commit, rootDigest: identity.rootDigest })).toMatch(/commit/)
    const foreign = { ...identity, rootDigest: '0'.repeat(64) }
    expect(checkIdentity(foreign, { commit: identity.commit, rootDigest: identity.rootDigest })).toMatch(/checkout/)
    expect(checkIdentity({ service: 'something-else' }, { commit: identity.commit, rootDigest: identity.rootDigest })).toMatch(/not the officeview viewer harness/)
    for (const outcome of [
      checkIdentity(stale, { commit: identity.commit, rootDigest: identity.rootDigest }),
      checkIdentity(foreign, { commit: identity.commit, rootDigest: identity.rootDigest }),
    ]) {
      expect(outcome ?? '').not.toContain(process.cwd())
    }
  })
})

describe('LAN-mode file serving is locked to the harness minimum', () => {
  let child: ChildProcess | null = null

  async function startHarnessServer(): Promise<string> {
    const port = await freePort()
    child = spawn('bunx', ['vite', '--config', 'vite.harness.config.ts', '--port', String(port), '--strictPort', '--host', '127.0.0.1'], {
      cwd: process.cwd(),
      stdio: 'ignore',
    })
    const base = `http://127.0.0.1:${port}`
    const startedAt = Date.now()
    for (;;) {
      try {
        const response = await fetch(`${base}/src/debug/interactive-viewer.html`)
        if (response.ok) return base
      } catch {
        // not up yet
      }
      if (Date.now() - startedAt > 90_000) throw new Error('harness vite server did not start')
      await new Promise((resolve) => setTimeout(resolve, 250))
    }
  }

  it('loads the page and its entry module while denying files outside the allow list', async () => {
    let base = ''
    try {
      base = await startHarnessServer()
      const page = await fetch(`${base}/src/debug/interactive-viewer.html`)
      expect(page.status).toBe(200)
      const entry = await fetch(`${base}/src/debug/interactive-viewer.tsx`)
      expect(entry.status).toBe(200)
      const repo = encodeURI(resolve(process.cwd()))
      for (const forbidden of [`/@fs/${repo}/package.json`, '/src/core/zoom.ts', '/package.json']) {
        const response = await fetch(`${base}${forbidden}`)
        expect(response.status, forbidden).not.toBe(200)
      }
      const identity = await (await fetch(`${base}/__harness-id`)).json()
      expect(identity.root).toBeUndefined()
      expect(JSON.stringify(identity)).not.toContain(resolve(process.cwd()))
    } finally {
      child?.kill('SIGTERM')
      child = null
    }
  }, 120_000)
})
