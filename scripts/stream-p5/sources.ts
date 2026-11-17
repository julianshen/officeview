/**
 * Piece-5 Task 1 (feasibility gate) — the ffmpeg side of the measurement harness.
 *
 * Measurement only: a plain subprocess, not the product runner (no sandbox, no D7 demuxer allowlist).
 * The decode and encode argv follow plan D8 verbatim except where noted:
 *   - the sources carry no audio, so the encode has no input 1 and no `-map '1:a:0?'` / AAC options;
 *   - the source decode omits the D7 input options (unknown to this task; the sources are our own files);
 *   - `-threads 2` on the encode (controller ruling X31: the product encode argv carries it).
 * Every intermediate goes to the scratch directory; nothing here is committed.
 */
import { spawnSync } from 'node:child_process'
import { closeSync, existsSync, mkdirSync, openSync, readSync, statSync, writeFileSync } from 'node:fs'
import path from 'node:path'

export const FFMPEG = process.env.OV_FFMPEG ?? '/opt/homebrew/bin/ffmpeg'
export const FPS = 30
export const SECONDS = 6
export const FRAMES = FPS * SECONDS
export const PAGE = { w: 1280, h: 720 } as const

export type SourceName = 'testsrc2' | 'mandelbrot' | 'kenburns'
export const SOURCE_NAMES: readonly SourceName[] = ['testsrc2', 'mandelbrot', 'kenburns']

/** The corpus photo for the Ken Burns source: 720×540 JPEG inside a LibreOffice test deck. */
export const KENBURNS_PHOTO = { deck: 'corpus/lo-tdf104788.pptx', part: 'ppt/media/image1.jpeg' }

function run(args: string[], opts: { stdinFd?: number; maxBuffer?: number } = {}): Buffer {
  const r = spawnSync(FFMPEG, ['-hide_banner', '-v', 'error', '-nostdin', ...args], {
    stdio: [opts.stdinFd ?? 'ignore', 'pipe', 'pipe'],
    maxBuffer: opts.maxBuffer ?? 64 * 1024 * 1024,
  })
  if (r.status !== 0) throw new Error(`ffmpeg ${args.join(' ')} failed (${r.status}): ${r.stderr?.toString().slice(0, 4000)}`)
  return r.stdout
}

export function ffmpegVersion(): string {
  return spawnSync(FFMPEG, ['-version'], { encoding: 'utf8' }).stdout.split('\n')[0]!.trim()
}

/** Writes the 6 s, 30 fps, 1280×720 source as lossless FFV1 (deterministic lavfi / zoompan). */
export function makeSource(name: SourceName, dir: string, repoRoot: string): string {
  mkdirSync(dir, { recursive: true })
  const out = path.join(dir, `${name}.mkv`)
  if (existsSync(out)) return out
  const tail = ['-t', String(SECONDS), '-r', String(FPS), '-c:v', 'ffv1', '-fflags', '+bitexact', '-flags', '+bitexact', '-threads', '1', '-y', out]
  if (name === 'testsrc2') run(['-f', 'lavfi', '-i', `testsrc2=size=1280x720:rate=${FPS}:duration=${SECONDS}`, ...tail])
  else if (name === 'mandelbrot') run(['-f', 'lavfi', '-i', `mandelbrot=size=1280x720:rate=${FPS}`, ...tail])
  else {
    const photo = path.join(dir, 'kenburns-photo.jpeg')
    if (!existsSync(photo)) {
      const r = spawnSync('unzip', ['-p', path.join(repoRoot, KENBURNS_PHOTO.deck), KENBURNS_PHOTO.part], { maxBuffer: 1 << 24 })
      if (r.status !== 0 || !r.stdout.length) throw new Error(`cannot extract ${KENBURNS_PHOTO.part}; run bun scripts/fetch-corpus.ts`)
      writeFileSync(photo, r.stdout)
    }
    // Slow zoom-in to 1.36× over 180 frames, centred: a Ken Burns move with sub-pixel motion every frame.
    run(['-loop', '1', '-framerate', String(FPS), '-i', photo, '-vf',
      `scale=2560:1920:flags=bicubic+bitexact,zoompan=z='1+0.002*on':x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':d=${FRAMES}:s=1280x720:fps=${FPS}`, ...tail])
  }
  return out
}

/** D8 video decode (pinned): fps, bitexact bicubic scale to full-range RGBA. */
export function decodeArgs(input: string, w: number, h: number): string[] {
  return ['-i', input, '-map', '0:v:0', '-vf',
    `fps=${FPS},scale=${w}:${h}:flags=bicubic+accurate_rnd+full_chroma_int+bitexact:in_range=auto:out_range=full,format=rgba`,
    '-fflags', '+bitexact', '-flags', '+bitexact', '-threads', '1', '-f', 'rawvideo', 'pipe:1']
}

/** Decodes `input` to `FRAMES` RGBA frames at w×h. Throws if the frame count differs. */
export function decodeFrames(input: string, w: number, h: number): Buffer {
  const buf = run(decodeArgs(input, w, h), { maxBuffer: w * h * 4 * (FRAMES + 8) })
  if (buf.length !== w * h * 4 * FRAMES) throw new Error(`decode ${input}: ${buf.length / (w * h * 4)} frames, expected ${FRAMES}`)
  return buf
}

export interface EncodeSettings { crf: number; preset: string; qp0?: boolean }

/** D8 video encode from a raw RGBA file streamed on pipe:0, both-output flags, `-t` to the duration. */
export function encodeArgs(w: number, h: number, s: EncodeSettings, out: string): string[] {
  const rate = s.qp0 ? ['-qp', '0'] : ['-crf', String(s.crf)]
  return ['-f', 'rawvideo', '-pix_fmt', 'rgba', '-s', `${w}x${h}`, '-r', String(FPS), '-i', 'pipe:0',
    '-map', '0:v:0', '-c:v', 'libx264', '-preset', s.preset, ...rate, '-threads', '2', '-pix_fmt', 'yuv420p', ...(s.qp0 ? [] : ['-profile:v', 'high']),
    '-colorspace', 'bt709', '-color_primaries', 'bt709', '-color_trc', 'bt709', '-color_range', 'tv',
    '-bsf:v', 'filter_units=remove_types=6', '-movflags', '+faststart',
    '-sn', '-dn', '-map_metadata', '-1', '-map_chapters', '-1', '-metadata:s:v:0', 'encoder=',
    '-fflags', '+bitexact', '-flags:v', '+bitexact', '-t', String(SECONDS), '-y', out]
}

/** Encodes the raw RGBA file; returns the output size and the encode wall time. */
export function encodeRaw(rawPath: string, w: number, h: number, s: EncodeSettings, out: string): { bytes: number; wallMs: number } {
  const fd = openSync(rawPath, 'r')
  const t = performance.now()
  try { run(encodeArgs(w, h, s, out), { stdinFd: fd }) } finally { closeSync(fd) }
  return { bytes: statSync(out).size, wallMs: performance.now() - t }
}

/** One frame of a raw RGBA file. */
export function readFrame(rawPath: string, w: number, h: number, index: number): Uint8ClampedArray {
  const n = w * h * 4, out = Buffer.alloc(n), fd = openSync(rawPath, 'r')
  try { if (readSync(fd, out, 0, n, index * n) !== n) throw new Error(`short read ${rawPath}#${index}`) } finally { closeSync(fd) }
  return new Uint8ClampedArray(out.buffer, out.byteOffset, n)
}
