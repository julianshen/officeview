/**
 * Piece 5 (plan D17): generates the committed media fixtures under `server/test/fixtures/media/`.
 *
 *   bun scripts/stream-p5/make-media-fixtures.ts            # writes the fixtures
 *   bun scripts/stream-p5/make-media-fixtures.ts --check    # regenerates into a temp dir and compares sha256
 *
 * Deterministic on one ffmpeg build: every encode is `+bitexact` (no encoder strings, no creation time) and drops
 * metadata. ffmpeg comes from `OV_FFMPEG` (default `ffmpeg` on PATH). Every file is 2–4 s of `sine` (and, for
 * `av.mp4` only, `testsrc2`), under 1 MB each and ≤ 4 MB in all.
 *
 * Hostile fixtures (all `media-unreadable`):
 * - `trunc.m4a`: the first 4 KiB of `a.m4a`, whose `moov` is at the end (no faststart), so the cut loses it;
 * - `badhdr.m4a`: an `ftyp` box followed by garbage;
 * - `hls.m3u8`, `concat.txt`: playlist and `ffconcat` text naming other files (`file:/etc/hosts`);
 * - `dref.mov`: an audio-only mov with no `mdat`, whose `dref` holds an external alias (`alis`) entry naming
 *   `DREF_SENTINEL`: its samples would be read from that file were external data references enabled.
 */
import { createHash } from 'node:crypto'
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

/** The absolute path `dref.mov`'s external data reference names (the hostile test writes a secret there). */
export const DREF_SENTINEL = '/tmp/ov-p5-dref-sentinel.txt'

const OUT = resolve(import.meta.dir, '../../server/test/fixtures/media')
const FFMPEG = process.env.OV_FFMPEG || 'ffmpeg'
const GLOBAL = ['-nostdin', '-hide_banner', '-loglevel', 'error', '-y']
const OUTPUT = ['-fflags', '+bitexact', '-flags', '+bitexact', '-map_metadata', '-1', '-map_chapters', '-1']
const sine = (sr: number, seconds: number, freq = 440) => ['-f', 'lavfi', '-i', `sine=frequency=${freq}:sample_rate=${sr}:duration=${seconds}`]

function ffmpeg(inputs: string[], output: string[]): void {
  const args = [...GLOBAL, ...inputs, ...OUTPUT, ...output]
  const r = Bun.spawnSync([FFMPEG, ...args], { stdout: 'pipe', stderr: 'pipe' })
  if (r.exitCode !== 0) throw new Error(`ffmpeg ${args.join(' ')} failed: ${r.stderr.toString()}`)
}

// --- ISO BMFF box editing (dref.mov) -----------------------------------------------------------------------------

interface Box { type: string; start: number; size: number; header: number }
function boxes(b: Uint8Array, from: number, to: number): Box[] {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength), out: Box[] = []
  for (let at = from; at + 8 <= to;) {
    const size = v.getUint32(at), type = String.fromCharCode(...b.subarray(at + 4, at + 8))
    if (size < 8 || at + size > to) throw new Error(`bad box ${type} at ${at}`)
    out.push({ type, start: at, size, header: 8 }); at += size
  }
  return out
}
const u32 = (n: number) => { const a = new Uint8Array(4); new DataView(a.buffer).setUint32(0, n); return a }
const cat = (...parts: Uint8Array[]) => { const o = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let i = 0; for (const p of parts) { o.set(p, i); i += p.length } return o }
const ascii = (s: string) => new TextEncoder().encode(s)
const box = (type: string, body: Uint8Array) => cat(u32(8 + body.length), ascii(type), body)

/** Rebuilds `b`'s box at `path` (container types descended by name) with `edit` applied to its body. */
function rewrite(b: Uint8Array, path: string[], edit: (body: Uint8Array) => Uint8Array, from = 0, to = b.length): Uint8Array {
  const [head, ...rest] = path
  const kids = boxes(b, from, to)
  const out: Uint8Array[] = []
  let found = false
  for (const k of kids) {
    const whole = b.subarray(k.start, k.start + k.size)
    if (k.type !== head || found) { out.push(whole); continue }
    found = true
    // `stsd` and `dref` carry a version/flags word and an entry count before their children.
    const pre = head === 'stsd' || head === 'dref' ? 8 : 0
    const bodyStart = k.start + k.header, bodyEnd = k.start + k.size
    if (rest.length === 0) { out.push(box(head!, edit(b.subarray(bodyStart, bodyEnd)))); continue }
    const inner = rewrite(b, rest, edit, bodyStart + pre, bodyEnd)
    out.push(box(head!, cat(b.subarray(bodyStart, bodyStart + pre), inner)))
  }
  if (!found) throw new Error(`box ${head} not found`)
  return cat(...out)
}

/**
 * A Mac alias record (`alis`) as ffmpeg's `mov_read_dref` parses it: 10 skipped bytes, the volume (length byte and
 * 27 bytes), 12 skipped, the file name (length byte and 63 bytes), 16 skipped, `nlvl_from`/`nlvl_to`, 16 skipped,
 * then typed records ending at type -1. Type 2 is the absolute path. No volume and no levels: only the path names it.
 */
function aliasRecord(path: string): Uint8Array {
  const name = path.split('/').pop()!
  const fixed = new Uint8Array(150)
  fixed[10 + 1 + 27 + 12] = name.length
  fixed.set(ascii(name), 10 + 1 + 27 + 12 + 1)
  let p = ascii(path)
  if (p.length & 1) p = cat(p, new Uint8Array(1))
  const rec = (type: number, body: Uint8Array) => { const h = new Uint8Array(4); const v = new DataView(h.buffer); v.setInt16(0, type); v.setUint16(2, body.length); return cat(h, body) }
  return cat(fixed, rec(2, p), rec(-1, new Uint8Array(0)))
}

/** An audio-only mov whose single `dref` entry is an external alias (flags 0) naming `target`, with no `mdat`. */
function drefMov(src: Uint8Array, target: string): Uint8Array {
  const top = boxes(src, 0, src.length)
  const keep = top.filter((k) => k.type === 'ftyp' || k.type === 'moov').map((k) => src.subarray(k.start, k.start + k.size))
  const noMdat = cat(...keep)
  // `alis` full box: version 0, flags 0 (NOT self-contained), then the alias record.
  const entry = box('alis', cat(new Uint8Array([0, 0, 0, 0]), aliasRecord(target)))
  return rewrite(noMdat, ['moov', 'trak', 'mdia', 'minf', 'dinf', 'dref'], () => cat(new Uint8Array([0, 0, 0, 0]), u32(1), entry))
}

// --- Fixtures --------------------------------------------------------------------------------------------------

export function makeMediaFixtures(dir: string): void {
  mkdirSync(dir, { recursive: true })
  const at = (n: string) => join(dir, n)
  ffmpeg(sine(22050, 2), ['-ac', '1', '-c:a', 'aac', '-b:a', '32k', '-f', 'ipod', at('a.m4a')])
  ffmpeg(sine(22050, 3, 523), ['-ac', '1', '-c:a', 'libmp3lame', '-b:a', '32k', '-id3v2_version', '0', '-write_xing', '1', '-f', 'mp3', at('a.mp3')])
  ffmpeg(sine(8000, 2, 660), ['-ac', '1', '-c:a', 'pcm_s16le', '-f', 'wav', at('a.wav')])
  ffmpeg(sine(22050, 2, 330), ['-ac', '1', '-c:a', 'wmav2', '-b:a', '32k', '-f', 'asf', at('a.wma')])
  ffmpeg(['-f', 'lavfi', '-i', 'testsrc2=size=64x48:rate=10:duration=2', ...sine(22050, 2)],
    ['-map', '0:v', '-map', '1:a', '-c:v', 'libx264', '-preset', 'ultrafast', '-tune', 'zerolatency', '-threads', '1', '-x264-params', 'bframes=0', '-pix_fmt', 'yuv420p',
    '-c:a', 'aac', '-b:a', '32k', '-ac', '1', '-f', 'mp4', at('av.mp4')])
  ffmpeg(sine(8000, 2, 880), ['-ac', '1', '-c:a', 'pcm_s16le', '-f', 'mov', at('dref-src.mov')])

  const m4a = readFileSync(at('a.m4a'))
  writeFileSync(at('trunc.m4a'), m4a.subarray(0, 4096))
  const garbage = new Uint8Array(2048)
  for (let i = 0, x = 0x9e3779b9; i < garbage.length; i++) { x = (x * 1664525 + 1013904223) >>> 0; garbage[i] = x >>> 24 }
  writeFileSync(at('badhdr.m4a'), cat(box('ftyp', cat(ascii('M4A '), u32(0), ascii('M4A isom'))), garbage))
  writeFileSync(at('hls.m3u8'), '#EXTM3U\n#EXT-X-VERSION:3\n#EXT-X-TARGETDURATION:2\n#EXTINF:2.0,\nfile:/etc/hosts\n#EXT-X-ENDLIST\n')
  writeFileSync(at('concat.txt'), 'ffconcat version 1.0\nfile \'file:/etc/hosts\'\nduration 2\n')
  writeFileSync(at('dref.mov'), drefMov(readFileSync(at('dref-src.mov')), DREF_SENTINEL))
  rmSync(at('dref-src.mov'))
}

const digest = (dir: string) => Object.fromEntries(readdirSync(dir).sort().map((n) => [n, createHash('sha256').update(readFileSync(join(dir, n))).digest('hex')]))

if (import.meta.main) {
  if (process.argv.includes('--check')) {
    const d = mkdtempSync(join(tmpdir(), 'ov-p5-fixtures-'))
    try {
      makeMediaFixtures(d)
      const want = digest(OUT), got = digest(d)
      const diff = [...new Set([...Object.keys(want), ...Object.keys(got)])].filter((n) => want[n] !== got[n])
      if (diff.length) { console.error(`fixtures differ: ${diff.join(', ')}`); process.exit(1) }
      console.log('fixtures match')
    } finally { rmSync(d, { recursive: true, force: true }) }
  } else {
    makeMediaFixtures(OUT)
    for (const [n, h] of Object.entries(digest(OUT))) console.log(`${h.slice(0, 16)}  ${readFileSync(join(OUT, n)).length.toString().padStart(7)}  ${n}`)
  }
}
