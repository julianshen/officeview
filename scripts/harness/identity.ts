/**
 * Harness identity: which checkout/commit a running vite harness server was
 * started from. vite answers 200 with index.html for unknown paths, so the
 * check inspects the JSON BODY, never just the status.
 */
import { execFileSync } from 'node:child_process'
import { resolve } from 'node:path'

export const HARNESS_SERVICE = 'officeview-viewer-harness'
export const HARNESS_ID_PATH = '/__harness-id'

export interface HarnessIdentity {
  service: string
  commit: string
  dirty: boolean
  root: string
}

function git(root: string, args: string[]): string {
  return execFileSync('git', args, { cwd: root, stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim()
}

export function gitIdentity(root: string): HarnessIdentity {
  const absolute = resolve(root)
  try {
    return { service: HARNESS_SERVICE, commit: git(absolute, ['rev-parse', 'HEAD']), dirty: git(absolute, ['status', '--porcelain']).length > 0, root: absolute }
  } catch {
    return { service: HARNESS_SERVICE, commit: 'unknown', dirty: false, root: absolute }
  }
}

/** The label stored in results: the sha, plus -dirty for uncommitted changes. */
export function commitLabel(identity: Pick<HarnessIdentity, 'commit' | 'dirty'>): string {
  return identity.commit === 'unknown' ? 'unknown' : identity.dirty ? `${identity.commit}-dirty` : identity.commit
}

/** null when `body` is the harness server for this commit and checkout. */
export function checkIdentity(body: unknown, expected: Pick<HarnessIdentity, 'commit' | 'root'>): string | null {
  if (typeof body !== 'object' || body === null || (body as { service?: unknown }).service !== HARNESS_SERVICE) {
    return 'the server on this port is not the officeview viewer harness (no identity marker)'
  }
  const identity = body as Partial<HarnessIdentity>
  if (identity.commit !== expected.commit) return `harness server runs commit ${identity.commit} but this checkout is ${expected.commit} (stale server?)`
  if (identity.root !== expected.root) return `harness server serves a different checkout (root ${identity.root}, expected ${expected.root})`
  return null
}
