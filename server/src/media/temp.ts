/**
 * Piece 5 (spec §4.2, plan Global Constraints "Temp files"): the per-process media temp directory and its ledger.
 * `<root>/ov-media-<pid>`, mode 0700, every file 0600 and randomly named. The ledger counts live handles and their
 * bytes, so a caller can assert it is empty after every path.
 */
import { randomBytes } from 'node:crypto'
import { closeSync, fchmodSync, lstatSync, openSync, readdirSync, statSync, unlinkSync } from 'node:fs'
import { chmod, lstat, mkdir, open, rm, unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'

export interface MediaFile { path: string; dispose(): Promise<void> }
export interface MediaTemp {
  readonly dir: string
  create(bytes: Uint8Array): Promise<MediaFile>
  /** A 0600 file a worker wrote directly inside `dir`. */
  adopt(path: string, bytes: number): MediaFile
  /** A pre-created, empty 0600 output path (ffmpeg `-y` keeps the mode); counted at its size once written. */
  reserve(ext: string): MediaFile
  live(): number
  bytes(): number
  /** Removes the whole directory, at shutdown. */
  remove(): Promise<void>
}

const EXT = /^[a-z0-9]{1,8}$/

function uid(): number | undefined { return typeof process.getuid === 'function' ? process.getuid() : undefined }

async function claimDir(dir: string): Promise<void> {
  try {
    await mkdir(dir, { mode: 0o700 })
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e
    // A predictable name in a shared tmp. The caller has already checked `claimed`, so this process holds no live
    // instance here: an existing `ov-media-<our pid>` can only be left by an earlier, dead process that had our pid.
    // Only a real directory we own is reclaimed, emptied and reused; a symlink or anyone else's entry is refused.
    const st = await lstat(dir)
    if (!st.isDirectory() || st.isSymbolicLink() || (uid() !== undefined && st.uid !== uid())) throw new Error(`media temp: ${dir} exists and is not ours`)
    await rm(dir, { recursive: true, force: true })
    await mkdir(dir, { mode: 0o700 })
  }
  await chmod(dir, 0o700) // umask can only narrow; make it exact
  const st = await lstat(dir)
  if (!st.isDirectory() || (st.mode & 0o777) !== 0o700) throw new Error(`media temp: ${dir} is not a 0700 directory`)
}

/** Directories claimed by this process. A second claim throws (rather than sharing an instance whose `remove()`
 * would pull files from under the other holder). Set synchronously before any await, so concurrent calls cannot
 * both pass the check; released by `remove()` or a failed claim. */
const claimed = new Set<string>()

export async function createMediaTemp(root: string = tmpdir()): Promise<MediaTemp> {
  const dir = join(resolve(root), `ov-media-${process.pid}`)
  if (claimed.has(dir)) throw new Error(`media temp: ${dir} is already claimed by this process`)
  claimed.add(dir)
  try { await claimDir(dir) } catch (e) { claimed.delete(dir); throw e }
  return tempOver(dir, async (ledger) => {
    ledger.clear()
    try { await rm(dir, { recursive: true, force: true }) } finally { claimed.delete(dir) }
  })
}

/**
 * A view over a media directory another instance claimed (a render worker's view of the pool's `dir`, plan D21).
 * Workers share `process.pid` but not module state, so a worker must never call `createMediaTemp`: its empty
 * `claimed` set would reclaim, and empty, the live directory. The view claims nothing; `remove()` disposes only the
 * files it holds and leaves the directory to its owner.
 *
 * `prefix` (Task 13, carried from Task 3): every file the view creates is named `<prefix><random>.<ext>`. The pool
 * gives each worker its own prefix and, when that worker is terminated or dies (its ledger and its `remove()` die
 * with it), deletes the files that bear it (`sweepPrefixed`).
 */
export const WORKER_PREFIX = /^w[0-9a-f]{12}-$/
export function attachMediaTemp(dir: string, prefix = ''): MediaTemp {
  if (prefix && !WORKER_PREFIX.test(prefix)) throw new Error(`media temp: bad prefix ${JSON.stringify(prefix)}`)
  const abs = resolve(dir)
  const st = lstatSync(abs)
  if (!st.isDirectory() || st.isSymbolicLink() || (st.mode & 0o777) !== 0o700 || (uid() !== undefined && st.uid !== uid())) throw new Error(`media temp: ${abs} is not our 0700 directory`)
  return tempOver(abs, async (ledger) => {
    const paths = [...ledger.values()].map((e) => e.path)
    ledger.clear()
    await Promise.all(paths.map((p) => unlink(p).catch((e: NodeJS.ErrnoException) => { if (e.code !== 'ENOENT') throw e })))
  }, prefix)
}

/** Deletes every regular file in `dir` whose name starts with `prefix` (a gone worker's files). Synchronous: the pool
 * runs it in the worker's `close` event, before anything waiting on the close resumes. */
export function sweepPrefixed(dir: string, prefix: string): number {
  if (!WORKER_PREFIX.test(prefix)) throw new Error(`media temp: bad prefix ${JSON.stringify(prefix)}`)
  let n = 0
  let names: string[] = []
  try { names = readdirSync(dir) } catch { return 0 }
  for (const name of names) {
    if (!name.startsWith(prefix)) continue
    try { if (lstatSync(join(dir, name)).isFile()) { unlinkSync(join(dir, name)); n++ } } catch { /* gone */ }
  }
  return n
}

type Ledger = Map<symbol, { path: string; bytes?: number }>

function tempOver(dir: string, removeAll: (ledger: Ledger) => Promise<void>, prefix = ''): MediaTemp {
  /** Ledger: one entry per live handle; `bytes` undefined means "reserved, size read when asked". */
  const ledger: Ledger = new Map()

  const handle = (path: string, bytes?: number): MediaFile => {
    const key = Symbol(path)
    ledger.set(key, { path, bytes })
    let disposed = false
    return {
      path,
      async dispose() {
        if (disposed) return
        disposed = true
        try { await unlink(path) } catch (e) { if ((e as NodeJS.ErrnoException).code !== 'ENOENT') throw e } finally { ledger.delete(key) }
      },
    }
  }
  const fresh = (ext: string) => {
    if (!EXT.test(ext)) throw new Error(`media temp: bad extension ${JSON.stringify(ext)}`)
    return join(dir, `${prefix}${randomBytes(16).toString('hex')}.${ext}`)
  }

  return {
    dir,
    async create(bytes) {
      const path = fresh('bin')
      const fh = await open(path, 'wx', 0o600)
      try {
        await fh.chmod(0o600)
        await fh.writeFile(bytes)
      } catch (e) {
        await fh.close(); await unlink(path).catch(() => {})
        throw e
      }
      await fh.close()
      return handle(path, bytes.byteLength)
    },
    adopt(path, bytes) {
      const abs = resolve(path)
      if (dirname(abs) !== dir) throw new Error('media temp: adopt outside the media dir')
      const st = lstatSync(abs)
      if (!st.isFile() || (st.mode & 0o777) !== 0o600 || (uid() !== undefined && st.uid !== uid())) throw new Error('media temp: adopt needs our own 0600 regular file')
      if (!Number.isSafeInteger(bytes) || bytes < 0) throw new Error('media temp: bad byte count')
      return handle(abs, bytes)
    },
    reserve(ext) {
      const path = fresh(ext)
      const fd = openSync(path, 'wx', 0o600)
      try { fchmodSync(fd, 0o600) } finally { closeSync(fd) }
      return handle(path)
    },
    live: () => ledger.size,
    bytes() {
      let n = 0
      for (const e of ledger.values()) {
        if (e.bytes !== undefined) { n += e.bytes; continue }
        try { n += statSync(e.path).size } catch { /* gone */ }
      }
      return n
    },
    remove: () => removeAll(ledger),
  }
}
