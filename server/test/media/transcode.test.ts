/** Piece 5 Task 7: the sandboxed, metadata-free AAC/M4A transcoder (plan D5, D7, D8, D23; pre-flight X4, X29 R8). */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test'
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { MEMORY_BYTES, runSandboxed, type SandboxRun } from '../../src/media/sandbox'
import { createMediaTemp, type MediaFile } from '../../src/media/temp'
import { CODEC_WHITELIST, inputArgs, parseProbe, probeArgv, sniffDemuxer, type ProbedMedia } from '../../src/media/probe'
import { AUDIO_ARGV_TEMPLATE, MUXER_TAG_ALLOWLIST, PROBE_REPORTED_OPTIONAL, audioArgv, encodeKeyId, encodeKeyOf, encodeParamsSha } from '../../src/media/params'
import { AUDIO_MAX_BYTES, transcode, wallCapMs, type EncodeResult } from '../../src/media/transcode'
import { ffmpegForTests } from './ffmpeg-gate'

const info = await ffmpegForTests()
const FIX = resolve(import.meta.dir, '../fixtures/media')
const fx = (n: string) => new Uint8Array(readFileSync(join(FIX, n)))
const root = mkdtempSync(join(tmpdir(), 'ov-transcode-'))
const temp = await createMediaTemp(root)
afterAll(async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) })

const never = () => new AbortController().signal
const sha = (b: Uint8Array) => createHash('sha256').update(b).digest('hex')
/** One AAC frame at the output rate: the duration tolerance. */
const AAC_FRAME_MS = (1024 / 48_000) * 1000

/** stdout of a sandboxed run, after BOTH exit and the stream's end. */
async function collect(run: SandboxRun) {
  const parts: Buffer[] = []
  run.stdout.on('data', (b: Buffer) => parts.push(b))
  const drained = new Promise<void>((ok) => { run.stdout.once('end', ok); run.stdout.once('close', ok); run.stdout.once('error', () => ok()) })
  const [res] = await Promise.all([run.done, drained])
  return { ...res, stdout: Buffer.concat(parts).toString('utf8') }
}
const sandboxLimits = () => ({ cpuSeconds: 30, fileBytes: 64 << 20, memoryBytes: MEMORY_BYTES[info!.platform], wallMs: 30_000, stderrBytes: 65536 as const })

/** ffprobe JSON of a file, sandboxed. */
async function ffprobe(path: string, entries: string[]): Promise<any> {
  const r = await collect(runSandboxed(info!, 'ffprobe', ['-v', 'error', '-hide_banner', ...entries, '-of', 'json', `file:${path}`], sandboxLimits(), never()))
  if (r.code !== 0) throw new Error(`ffprobe failed: ${r.stderr}`)
  return JSON.parse(r.stdout)
}

/** A staged source and its entry from the REAL probe (durations differ across ffmpeg versions: never literals). */
interface Source { name: string; bytes: Uint8Array; file: MediaFile; entry: ProbedMedia }
async function stage(name: string, bytes: Uint8Array): Promise<Source> {
  const demuxer = sniffDemuxer(bytes)
  if (!demuxer) throw new Error(`${name}: not sniffed`)
  const file = await temp.create(bytes)
  const r = await collect(runSandboxed(info!, 'ffprobe', probeArgv(demuxer, file.path), sandboxLimits(), never()))
  const p = r.code === 0 ? parseProbe(r.stdout) : undefined
  if (!p) throw new Error(`${name}: probe failed ${r.stderr}`)
  return { name, bytes, file, entry: { k: 0, partPath: `/ppt/media/${name}`, kind: p.kind, hasAudio: p.hasAudio, durationMs: p.durationMs, sourceHash: sha(bytes), demuxer } }
}

async function encode(s: Source, seams?: Parameters<typeof transcode>[3]): Promise<{ r: EncodeResult; out: MediaFile; ms: number }> {
  const out = temp.reserve('m4a')
  const t0 = performance.now()
  const r = await transcode(info!, { entry: s.entry, src: s.file.path, out: out.path }, never(), seams)
  return { r, out, ms: Math.round(performance.now() - t0) }
}

/** A PCM s16le WAV of seeded noise (incompressible: AAC spends its whole bitrate). */
function noiseWav(seconds: number, rate = 48_000, channels = 2): Uint8Array {
  const n = seconds * rate * channels
  const b = Buffer.alloc(44 + n * 2)
  b.write('RIFF', 0); b.writeUInt32LE(36 + n * 2, 4); b.write('WAVE', 8); b.write('fmt ', 12)
  b.writeUInt32LE(16, 16); b.writeUInt16LE(1, 20); b.writeUInt16LE(channels, 22); b.writeUInt32LE(rate, 24)
  b.writeUInt32LE(rate * channels * 2, 28); b.writeUInt16LE(channels * 2, 32); b.writeUInt16LE(16, 34)
  b.write('data', 36); b.writeUInt32LE(n * 2, 40)
  let x = 0x9e3779b9
  for (let i = 0; i < n; i++) { x ^= x << 13; x ^= x >>> 17; x ^= x << 5; b.writeInt16LE(((x >>> 0) % 40_000) - 20_000, 44 + i * 2) }
  return new Uint8Array(b)
}

/** ffmpeg processes whose argv names `path` (pgrep exits 1 when none match). */
function processesNaming(path: string): string {
  try { return execFileSync('/usr/bin/pgrep', ['-fl', path], { encoding: 'utf8' }).trim() } catch { return '' }
}

const AUDIO_FIXTURES = ['a.m4a', 'a.mp3', 'a.wav', 'a.wma'] as const
const ALL_FIXTURES = [...AUDIO_FIXTURES, 'av.mp4'] as const
/** Source tags a real deck's media may carry; none of these strings may survive. */
const SECRETS = ['OVSECRET-TITLE', 'OVSECRET-ARTIST', 'OVSECRET-COMMENT', 'OVSECRET-STREAM', 'OVSECRET-CHAPTER', 'OVSECRET-HANDLER']

/** The ISO BMFF box tree (paths of types), descending into containers; `meta` is a full box (4 bytes of flags). */
const CONTAINERS = new Set(['moov', 'trak', 'mdia', 'minf', 'stbl', 'dinf', 'edts', 'udta', 'meta'])
function boxTree(b: Buffer, from = 0, end = b.length, prefix = ''): Array<{ path: string; at: number; size: number }> {
  const out: Array<{ path: string; at: number; size: number }> = []
  for (let o = from; o + 8 <= end;) {
    const size = b.readUInt32BE(o)
    const type = b.toString('latin1', o + 4, o + 8)
    if (size < 8 || o + size > end) throw new Error(`bad box ${type} at ${o}`)
    const path = prefix ? `${prefix}/${type}` : type
    out.push({ path, at: o, size })
    if (CONTAINERS.has(type)) out.push(...boxTree(b, o + 8 + (type === 'meta' ? 4 : 0), o + size, path))
    o += size
  }
  return out
}
const ALLOWED_BOXES = ['ftyp', 'free', 'mdat', 'moov', 'moov/mvhd', 'moov/trak', 'moov/trak/tkhd', 'moov/trak/edts', 'moov/trak/edts/elst',
  'moov/trak/mdia', 'moov/trak/mdia/mdhd', 'moov/trak/mdia/hdlr', 'moov/trak/mdia/minf', 'moov/trak/mdia/minf/smhd',
  'moov/trak/mdia/minf/dinf', 'moov/trak/mdia/minf/dinf/dref', 'moov/trak/mdia/minf/stbl',
  ...['stsd', 'stts', 'stsc', 'stsz', 'stco', 'sgpd', 'sbgp'].map((t) => `moov/trak/mdia/minf/stbl/${t}`),
  'moov/udta', 'moov/udta/meta', 'moov/udta/meta/hdlr', 'moov/udta/meta/ilst']
/** udta { meta { hdlr 'mdir' 'appl', ilst {} } }: written by the mp4 muxer for every output (measured, 8.1.1). */
const EMPTY_UDTA = '0000003d75647461000000356d657461000000000000002168646c7200000000000000006d6469726170706c00000000000000000000000008696c7374'

// --- pure ------------------------------------------------------------------------------------------------------

describe('audio argv (D7 input rules, D8)', () => {
  it('every encode argv carries an allowlisted -f, -protocol_whitelist file, -enable_drefs 0 for mov, the codec whitelist, and the staged file: path as its only input', () => {
    for (const d of ['mov', 'avi', 'wav', 'asf', 'mp3'] as const) {
      const a = audioArgv('/tmp/ov-media-1/src.bin', d, '/tmp/ov-media-1/out.m4a', 1999)
      const i = a.indexOf('-i')
      expect(a.filter((x) => x === '-i')).toHaveLength(1)
      expect(a[i + 1]).toBe('file:/tmp/ov-media-1/src.bin')
      const before = a.slice(0, i)
      expect(before[before.indexOf('-f') + 1]).toBe(d)
      expect(before[before.indexOf('-protocol_whitelist') + 1]).toBe('file')
      expect(before.includes('-enable_drefs')).toBe(d === 'mov')
      if (d === 'mov') expect(before[before.indexOf('-enable_drefs') + 1]).toBe('0')
      expect(before[before.indexOf('-codec_whitelist') + 1]).toBe(CODEC_WHITELIST)
      // The probe's own input options, verbatim (one source of truth).
      expect(a.slice(i - inputArgs(d, '/x').length + 2, i + 2)).toEqual(inputArgs(d, '/tmp/ov-media-1/src.bin'))
      // Exactly one output, an explicit mp4 muxer over a file: path, last.
      expect(a.slice(-3)).toEqual(['-f', 'mp4', 'file:/tmp/ov-media-1/out.m4a'])
      expect(a.filter((x) => x.startsWith('file:'))).toHaveLength(2)
      // D8 output options.
      const after = a.slice(i + 2)
      for (const [k, v] of [['-map', '0:a:0'], ['-c:a', 'aac'], ['-b:a', '128k'], ['-ac', '2'], ['-ar', '48000'], ['-map_metadata', '-1'], ['-map_chapters', '-1'], ['-metadata:s:a:0', 'encoder='], ['-fflags', '+bitexact'], ['-flags:a', '+bitexact'], ['-movflags', '+faststart'], ['-t', '1.999']] as const) {
        expect(after[after.indexOf(k) + 1]).toBe(v)
      }
      for (const f of ['-vn', '-sn', '-dn']) expect(after).toContain(f)
      expect(after[after.indexOf('-af') + 1]).toContain('apad')
      expect(a.some((x) => /thread/.test(x))).toBe(false) // D5: the audio run sets no thread option
    }
    expect(audioArgv('/a/s', 'wav', '/a/o', 2000)).toContain('2.000')
    expect(audioArgv('/a/s', 'wav', '/a/o', 1_800_000)).toContain('1800.000')
    expect(audioArgv('/a/s', 'wav', '/a/o', 5)).toContain('0.005')
    expect(() => audioArgv('relative', 'wav', '/a/o', 2000)).toThrow()
    expect(() => audioArgv('/a/s', 'hls' as never, '/a/o', 2000)).toThrow()
    expect(() => audioArgv('/a/s', 'wav', 'o.m4a', 2000)).toThrow()
    expect(() => audioArgv('/a/s', 'wav', '/a/o\n', 2000)).toThrow()
    for (const ms of [0, -1, 1.5, 1_800_001, Number.NaN]) expect(() => audioArgv('/a/s', 'wav', '/a/o', ms)).toThrow()
  })

  it('audioArgv is the hashed template with only src, demuxer, out and duration substituted', () => {
    const a = audioArgv('/S', 'wav', '/O', 1234)
    const expanded = AUDIO_ARGV_TEMPLATE.flatMap((t) => t === '{input}' ? inputArgs('wav', '/S') : [t.replace('{t}', '1.234').replace('{out}', '/O')])
    expect(a).toEqual(expanded)
  })

  it('encodeParamsSha changes with the ffmpeg version and with the argv template', () => {
    const a = encodeParamsSha('8.1.1')
    expect(a).toMatch(/^[0-9a-f]{64}$/)
    expect(encodeParamsSha('8.1.1')).toBe(a)
    expect(encodeParamsSha('6.1.1')).not.toBe(a)
    expect(encodeParamsSha('8.1.1', AUDIO_ARGV_TEMPLATE)).toBe(a)
    expect(encodeParamsSha('8.1.1', AUDIO_ARGV_TEMPLATE.map((t) => t === '128k' ? '96k' : t))).not.toBe(a)
    expect(encodeParamsSha('8.1.1', [...AUDIO_ARGV_TEMPLATE, '-shortest'])).not.toBe(a)
  })

  it('encodeKeyOf depends on docId, sourceHash, paramsSha and ffmpegVersion only: two entries of one document with equal sourceHash and different k or partPath give one key; a different docId or ffmpeg version gives a different key', () => {
    const e1: ProbedMedia = { k: 0, partPath: '/ppt/media/media1.mp3', kind: 'audio', hasAudio: true, durationMs: 2000, sourceHash: 'a'.repeat(64), demuxer: 'mp3' }
    const e2: ProbedMedia = { ...e1, k: 3, partPath: '/ppt/media/media9.mp3', durationMs: 1999 }
    const v8 = { version: '8.1.1' }, v6 = { version: '6.1.1' }
    const k1 = encodeKeyOf('doc-1', e1, v8)
    expect(k1).toEqual({ docId: 'doc-1', sourceHash: 'a'.repeat(64), paramsSha: encodeParamsSha('8.1.1'), ffmpegVersion: '8.1.1' })
    expect(Object.keys(k1).sort()).toEqual(['docId', 'ffmpegVersion', 'paramsSha', 'sourceHash'])
    expect(encodeKeyId(encodeKeyOf('doc-1', e2, v8))).toBe(encodeKeyId(k1))
    expect(encodeKeyId(k1)).toMatch(/^[0-9a-f]{64}$/)
    expect(encodeKeyId(encodeKeyOf('doc-2', e1, v8))).not.toBe(encodeKeyId(k1))
    expect(encodeKeyId(encodeKeyOf('doc-1', e1, v6))).not.toBe(encodeKeyId(k1))
    expect(encodeKeyId(encodeKeyOf('doc-1', { sourceHash: 'b'.repeat(64) }, v8))).not.toBe(encodeKeyId(k1))
    // Canonical: field order does not matter.
    expect(encodeKeyId({ ffmpegVersion: k1.ffmpegVersion, paramsSha: k1.paramsSha, sourceHash: k1.sourceHash, docId: k1.docId })).toBe(encodeKeyId(k1))
  })

  it('caps: 32 MiB, wall = 2 × duration + 30 s', () => {
    expect(AUDIO_MAX_BYTES).toBe(32 << 20)
    expect(wallCapMs(2000)).toBe(34_000)
    expect(wallCapMs(1_800_000)).toBe(3_630_000)
  })
})

// --- ffmpeg ----------------------------------------------------------------------------------------------------

describe.skipIf(!info)('transcode (sandboxed ffmpeg)', () => {
  const sources = new Map<string, Source>()
  const outputs = new Map<string, { r: EncodeResult; bytes: Uint8Array; ms: number }>()
  let tagged: Source

  beforeAll(async () => {
    for (const n of ALL_FIXTURES) sources.set(n, await stage(n, fx(n)))
    tagged = await stage('tagged.m4a', await taggedSource())
    sources.set('tagged.m4a', tagged)
    for (const [n, s] of sources) {
      const { r, out, ms } = await encode(s)
      outputs.set(n, { r, bytes: new Uint8Array(readFileSync(out.path)), ms })
      await out.dispose()
    }
    console.log(`encode wall ms: ${[...outputs].map(([n, o]) => `${n}=${o.ms}`).join(' ')}`)
  }, 120_000)

  /** A source carrying every kind of metadata the output must not: global and stream tags, chapters, a non-und
   * language, a custom handler, an encoder string, a creation_time. Made from a.wav by sandboxed ffmpeg. */
  async function taggedSource(): Promise<Uint8Array> {
    const dir = mkdtempSync(join(root, 'tagged-'))
    const meta = join(dir, 'meta.txt'), wav = join(dir, 'a.wav'), out = join(dir, 'tagged.m4a')
    writeFileSync(wav, fx('a.wav'))
    writeFileSync(meta, ';FFMETADATA1\ntitle=OVSECRET-TITLE\nartist=OVSECRET-ARTIST\ncomment=OVSECRET-COMMENT\n'
      + '[CHAPTER]\nTIMEBASE=1/1000\nSTART=0\nEND=1000\ntitle=OVSECRET-CHAPTER\n[CHAPTER]\nTIMEBASE=1/1000\nSTART=1000\nEND=2000\ntitle=OVSECRET-CHAPTER-2\n')
    const r = await collect(runSandboxed(info!, 'ffmpeg', ['-nostdin', '-hide_banner', '-loglevel', 'error', '-y', '-i', `file:${wav}`, '-i', `file:${meta}`,
      '-map', '0:a', '-map_metadata', '1', '-map_chapters', '1', '-c:a', 'aac', '-metadata:s:a:0', 'title=OVSECRET-STREAM', '-metadata:s:a:0', 'language=fra',
      '-metadata:s:a:0', 'handler_name=OVSECRET-HANDLER', '-metadata', 'creation_time=2020-01-02T03:04:05Z', '-f', 'mp4', `file:${out}`], sandboxLimits(), never()))
    if (r.code !== 0) throw new Error(`tagged source: ${r.stderr}`)
    return new Uint8Array(readFileSync(out))
  }

  const ok = (n: string) => {
    const o = outputs.get(n)!
    if (!o.r.ok) throw new Error(`${n}: ${o.r.reason}`)
    return { ...o, r: o.r }
  }
  const outFile = (n: string) => { const p = join(root, `${n}.probe.m4a`); writeFileSync(p, ok(n).bytes); return p }

  it('the tagged source really carries the metadata (positive control for the no-metadata test)', async () => {
    const j = await ffprobe(tagged.file.path, ['-show_format', '-show_streams', '-show_chapters'])
    expect(j.format.tags.title).toBe('OVSECRET-TITLE')
    expect(j.format.tags.creation_time).toBeDefined()
    expect(j.streams[0].tags.language).toBe('fra')
    expect(j.streams[0].tags.handler_name).toBe('OVSECRET-HANDLER')
    expect(j.chapters.length).toBe(2)
    expect(Buffer.from(tagged.bytes).includes('Lavf')).toBe(true)
  })

  it('output duration equals durationMs within one AAC frame, and EncodeResult.durationMs === entry.durationMs', async () => {
    for (const [n, s] of sources) {
      const { r } = ok(n)
      expect(r.durationMs).toBe(s.entry.durationMs)
      expect(r.contentType).toBe('audio/mp4')
      expect(r.paramsSha).toBe(encodeParamsSha(info!.version))
      expect(r.ffmpegVersion).toBe(info!.version)
      expect(r.bytes).toBe(ok(n).bytes.length)
      const j = await ffprobe(outFile(n), ['-show_entries', 'format=duration:stream=duration'])
      for (const d of [j.format.duration, j.streams[0].duration]) {
        expect(Math.abs(Number(d) * 1000 - s.entry.durationMs)).toBeLessThanOrEqual(AAC_FRAME_MS)
      }
    }
  })

  it('no metadata: ffprobe -show_format -show_streams shows only MUXER_TAG_ALLOWLIST tags at their pinned values; no chapters, subtitle/data/attachment streams, encoder string or creation_time', async () => {
    // Pinned exactly (no version branch): a 6.x difference fails here instead of slipping through.
    expect(MUXER_TAG_ALLOWLIST).toEqual({
      format: { major_brand: ['isom'], minor_version: ['512'], compatible_brands: ['isomiso2mp41'] },
      stream: { language: ['und'], handler_name: ['SoundHandler'] },
    })
    for (const n of sources.keys()) {
      const j = await ffprobe(outFile(n), ['-show_format', '-show_streams', '-show_chapters', '-show_programs'])
      expect(j.chapters ?? []).toEqual([])
      expect(j.programs ?? []).toEqual([])
      expect(j.streams).toHaveLength(1)
      expect(j.streams.map((s: any) => s.codec_type)).toEqual(['audio'])
      expect(j.format.nb_streams).toBe(1)
      const check = (tags: Record<string, string> | undefined, allow: Record<string, readonly string[]>, optional: Record<string, readonly string[]> = {}) => {
        // Every allowlisted key is present at its pinned value; a probe-reported optional key may appear
        // only at its pinned value; nothing else.
        const seen = Object.keys(tags ?? {}).filter((k) => !(k in optional)).sort()
        expect(seen).toEqual(Object.keys(allow).sort())
        for (const [k, v] of Object.entries(tags ?? {})) expect(allow[k] ?? optional[k]).toContain(v)
      }
      check(j.format.tags, MUXER_TAG_ALLOWLIST.format)
      check(j.streams[0].tags, MUXER_TAG_ALLOWLIST.stream, PROBE_REPORTED_OPTIONAL.stream)
      const raw = Buffer.from(ok(n).bytes)
      for (const s of ['Lavf', 'Lavc', 'encoder', 'creation_time', 'libavformat', 'FFmpeg', ...SECRETS]) expect(raw.includes(s)).toBe(false)
      // The box tree holds only the boxes an audio-only mp4 needs (no chpl, tref, covr, uuid, a second trak…), and
      // its one udta is the muxer's constant empty item list (meta/hdlr 'mdir' 'appl'/ilst with no items).
      const boxes = boxTree(raw)
      for (const b of boxes) expect(ALLOWED_BOXES).toContain(b.path)
      expect(boxes.filter((b) => b.path === 'moov/trak')).toHaveLength(1)
      const udta = boxes.filter((b) => b.path === 'moov/udta')
      expect(udta).toHaveLength(1)
      expect(raw.subarray(udta[0]!.at, udta[0]!.at + udta[0]!.size).toString('hex')).toBe(EMPTY_UDTA)
    }
  })

  it('no 64-byte window of the source appears in the output', () => {
    for (const [n, s] of sources) {
      const out = Buffer.from(ok(n).bytes)
      const windows = new Set<string>()
      for (let i = 0; i + 64 <= out.length; i++) windows.add(out.subarray(i, i + 64).toString('latin1'))
      const src = Buffer.from(s.bytes)
      let checked = 0
      for (let i = 0; i + 64 <= src.length; i++) {
        const w = src.subarray(i, i + 64)
        if (new Set(w).size < 32) continue // X29 R8: a low-entropy window carries no recoverable content
        checked++
        if (windows.has(w.toString('latin1'))) throw new Error(`${n}: source window at ${i} appears in the output`)
      }
      expect(checked).toBeGreaterThan(0)
    }
  })

  it('av.mp4 used as audio: the output has one AAC stream and no video stream', async () => {
    expect(sources.get('av.mp4')!.entry.kind).toBe('video')
    const j = await ffprobe(outFile('av.mp4'), ['-show_streams'])
    expect(j.streams.map((s: any) => [s.codec_type, s.codec_name])).toEqual([['audio', 'aac']])
    const raw = Buffer.from(ok('av.mp4').bytes)
    for (const s of ['avc1', 'avcC', 'x264', 'vide']) expect(raw.includes(s)).toBe(false)
  })

  it('audio path: AAC in M4A for mp3, wav, m4a and wma sources', async () => {
    for (const n of AUDIO_FIXTURES) {
      const j = await ffprobe(outFile(n), ['-show_format', '-show_streams'])
      expect(j.format.format_name).toContain('mp4')
      expect(j.streams.map((s: any) => [s.codec_type, s.codec_name, s.profile, s.sample_rate, s.channels])).toEqual([['audio', 'aac', 'LC', '48000', 2]])
      // faststart: moov before mdat.
      const raw = Buffer.from(ok(n).bytes)
      expect(raw.indexOf('moov')).toBeLessThan(raw.indexOf('mdat'))
    }
  })

  it('two encodes of one source are byte-identical', async () => {
    for (const [n, s] of sources) {
      const { r, out } = await encode(s)
      try {
        expect(r.ok).toBe(true)
        expect(statSync(out.path).mode & 0o777).toBe(0o600) // Global Constraints: temp files are 0600
        expect(sha(new Uint8Array(readFileSync(out.path)))).toBe(sha(ok(n).bytes))
      } finally { await out.dispose() }
    }
  })

  it('a high-bitrate source over 32 MiB → media-too-large, never truncated, output deleted', async () => {
    const s = await stage('noise.wav', noiseWav(30))
    try {
      // Positive control: uncapped, the same source encodes to more than the seam's cap.
      const big = await encode(s)
      expect(big.r.ok).toBe(true)
      if (big.r.ok) expect(big.r.bytes).toBeGreaterThan(256 << 10)
      await big.out.dispose()
      if (!big.r.ok) throw new Error('positive control failed')
      // The final-size branch (under the fsize slack, so no SIGXFSZ): one byte under the real size is refused…
      const under = await encode(s, { maxBytes: big.r.bytes - 1 })
      expect(under.r).toEqual({ ok: false, reason: 'media-too-large' })
      expect(existsSync(under.out.path)).toBe(false)
      await under.out.dispose()
      // …and exactly the real size is not.
      const exact = await encode(s, { maxBytes: big.r.bytes })
      expect(exact.r.ok).toBe(true)
      await exact.out.dispose()
      const { r, out } = await encode(s, { maxBytes: 256 << 10 })
      expect(r).toEqual({ ok: false, reason: 'media-too-large' })
      expect(existsSync(out.path)).toBe(false)
      await out.dispose()
      expect(processesNaming(out.path)).toBe('')
    } finally { await s.file.dispose() }
  }, 60_000)

  it('wall cap → media-timeout and the group is killed', async () => {
    const s = await stage('long.wav', noiseWav(120, 48_000, 2))
    try {
      const t0 = performance.now()
      const { r, out } = await encode(s, { wallCapMs: 150 })
      expect(r).toEqual({ ok: false, reason: 'media-timeout' })
      expect(performance.now() - t0).toBeLessThan(5000)
      expect(existsSync(out.path)).toBe(false)
      expect(processesNaming(out.path)).toBe('')
      await out.dispose()
    } finally { await s.file.dispose() }
  }, 60_000)

  it('abort rejects (never a cacheable refusal), kills the group and deletes the output', async () => {
    const s = await stage('abort.wav', noiseWav(120, 48_000, 2))
    try {
      const out = temp.reserve('m4a')
      const ac = new AbortController()
      const p = transcode(info!, { entry: s.entry, src: s.file.path, out: out.path }, ac.signal)
      setTimeout(() => ac.abort(), 100)
      await expect(p).rejects.toThrow()
      expect(existsSync(out.path)).toBe(false)
      expect(processesNaming(out.path)).toBe('')
      await out.dispose()
    } finally { await s.file.dispose() }
  }, 60_000)

  it('an undecodable source → media-unreadable, output deleted', async () => {
    const bytes = fx('trunc.m4a')
    const file = await temp.create(bytes)
    const entry: ProbedMedia = { k: 0, partPath: '/ppt/media/t.m4a', kind: 'audio', hasAudio: true, durationMs: 2000, sourceHash: sha(bytes), demuxer: 'mov' }
    const out = temp.reserve('m4a')
    try {
      const r = await transcode(info!, { entry, src: file.path, out: out.path }, never())
      expect(r).toEqual({ ok: false, reason: 'media-unreadable' })
      expect(existsSync(out.path)).toBe(false)
    } finally { await out.dispose(); await file.dispose() }
  })

  it('leaves no temp file behind', async () => {
    for (const s of sources.values()) await s.file.dispose()
    expect(temp.live()).toBe(0)
  })
})
