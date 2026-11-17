/**
 * Piece 5 (spec §4.2, plan Global Constraints "Temp files"): the per-process media temp directory and its ledger.
 * `<root>/ov-media-<pid>`, mode 0700, every file 0600 and randomly named. The ledger counts live handles and their
 * bytes, so a caller can assert it is empty after every path.
 */
import { randomBytes } from 'node:crypto'
import { closeSync, fchmodSync, lstatSync, openSync, statSync } from 'node:fs'
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
  /** Ledger: one entry per live handle; `bytes` undefined means "reserved, size read when asked". */
  const ledger = new Map<symbol, { path: string; bytes?: number }>()

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
    return join(dir, `${randomBytes(16).toString('hex')}.${ext}`)
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
    async remove() {
      ledger.clear()
      try { await rm(dir, { recursive: true, force: true }) } finally { claimed.delete(dir) }
    },
  }
}
