/**
 * Harness identity: which checkout/commit a running vite harness server was
 * started from. vite answers 200 with index.html for unknown paths, so the
 * check inspects the JSON BODY, never just the status.
 */
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { resolve } from 'node:path'

export const HARNESS_SERVICE = 'officeview-viewer-harness'
export const HARNESS_ID_PATH = '/__harness-id'

export interface HarnessIdentity {
  service: string
  commit: string
  dirty: boolean
  /** SHA-256 over root + commit: proves same checkout without leaking the
   * absolute path to anyone on the LAN who fetches the marker. */
  rootDigest: string
}

/** Digest binding a checkout path to a commit; compared, never displayed. */
export function identityDigest(root: string, commit: string): string {
  return createHash('sha256').update(`${root}\n${commit}`, 'utf8').digest('hex')
}

function git(root: string, args: string[]): string {
  return execFileSync('git', args, { cwd: root, stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim()
}

export function gitIdentity(root: string): HarnessIdentity {
  const absolute = resolve(root)
  try {
    const commit = git(absolute, ['rev-parse', 'HEAD'])
    return { service: HARNESS_SERVICE, commit, dirty: git(absolute, ['status', '--porcelain']).length > 0, rootDigest: identityDigest(absolute, commit) }
  } catch {
    return { service: HARNESS_SERVICE, commit: 'unknown', dirty: false, rootDigest: identityDigest(absolute, 'unknown') }
  }
}

/** The label stored in results: the sha, plus -dirty for uncommitted changes. */
export function commitLabel(identity: Pick<HarnessIdentity, 'commit' | 'dirty'>): string {
  return identity.commit === 'unknown' ? 'unknown' : identity.dirty ? `${identity.commit}-dirty` : identity.commit
}

/** null when `body` is the harness server for this commit and checkout. Absolute
 * paths never appear here: neither side discloses its checkout location. */
export function checkIdentity(body: unknown, expected: Pick<HarnessIdentity, 'commit' | 'rootDigest'>): string | null {
  if (typeof body !== 'object' || body === null || (body as { service?: unknown }).service !== HARNESS_SERVICE) {
    return 'the server on this port is not the officeview viewer harness (no identity marker)'
  }
  const identity = body as Partial<HarnessIdentity>
  if (identity.commit !== expected.commit) return `harness server runs commit ${identity.commit} but this checkout is ${expected.commit} (stale server?)`
  if (identity.rootDigest !== expected.rootDigest) return 'harness server serves a different checkout (identity digest mismatch)'
  return null
}
