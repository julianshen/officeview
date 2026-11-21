/** Piece 5 Task 3: sniffing (D7), the sandboxed duration probe, the per-document media table (D21) and audio uses (D22). */
import { afterAll, describe, expect, it } from 'bun:test'
import { createHash, randomBytes } from 'node:crypto'
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { createServer } from 'node:net'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import JSZip from 'jszip'
import { buildPptxDeck, timingXmlOf, type PptxShapeSpec } from '../../../src/testdata/ooxml-builders'
import { flyInDeck } from '../../../src/testdata/show-decks'
import { MEMORY_BYTES, runSandboxed } from '../../src/media/sandbox'
import { createMediaTemp } from '../../src/media/temp'
import { AUDIO_MAX_MS, PART_MAX_BYTES, PROBE_DECK_MS, PROBE_PART_MS, durationMsOf, inputArgs, parseProbe, probeArgv, sniffDemuxer, type DocMediaEntry, type ProbedMedia } from '../../src/media/probe'
import { DocumentHost } from '../../src/render'
import { Pool, RenderFailed } from '../../src/pool'
import { showBuildGrant } from '../../src/show/grant'
import { ffmpegForTests } from './ffmpeg-gate'
import { rejection } from '../fixtures/rejection'

const info = await ffmpegForTests()
const FIX = resolve(import.meta.dir, '../fixtures/media')
const fx = (n: string) => new Uint8Array(readFileSync(join(FIX, n)))
const root = mkdtempSync(join(tmpdir(), 'ov-probe-'))
/** One media temp for the whole file (a second claim of the same root throws). */
const temp = await createMediaTemp(root)
afterAll(async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) })

/** The fixed path `dref.mov`'s external alias names (scripts/stream-p5/make-media-fixtures.ts). */
const DREF_SENTINEL = '/tmp/ov-p5-dref-sentinel.txt'

// --- decks ------------------------------------------------------------------------------------------------------

type Use = 'audio' | 'narration' | 'video' | 'none'
interface MediaShape { spid: number; file: string | Uint8Array; ext?: string; rel?: 'audio' | 'video'; use: Use; hidden?: boolean }
const px = (n: number) => String(Math.round(n * 9525))
function mediaNode(id: number, spid: number, use: Use): string {
  if (use === 'none') return ''
  const tag = use === 'video' ? 'video' : 'audio'
  return `<p:${tag}${use === 'narration' ? ' isNarration="1"' : ''}><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:${tag}>`
}
async function mediaDeck(slides: MediaShape[][]): Promise<Uint8Array> {
  let id = 3
  return buildPptxDeck({
    slides: slides.map((shapes) => {
      const nodes = shapes.map((s) => mediaNode(id++, s.spid, s.use)).join('')
      return {
        shapes: shapes.map((s, i): PptxShapeSpec => ({ id: s.spid, name: `M${s.spid}`, prst: 'rect', off: [px(20 + 60 * i), px(20)], ext: [px(50), px(50)], fill: '336699' })),
        media: shapes.map((s) => ({ kind: s.rel ?? (s.use === 'video' ? 'video' : 'audio'), data: typeof s.file === 'string' ? fx(s.file) : s.file, ext: s.ext ?? (typeof s.file === 'string' ? s.file.split('.').pop()! : 'bin'), shapeId: s.spid })),
        ...(nodes ? { timingXml: timingXmlOf(nodes) } : {}),
      }
    }),
  })
}

/** A host with media on, over the file's temp directory. */
function mediaHost(seams?: DocumentHost['mediaSeams']) {
  const host = new DocumentHost()
  host.configureMedia(info ? { info, dir: temp.dir } : undefined)
  if (seams) host.mediaSeams = seams
  return host
}
const probed = (e: DocMediaEntry | undefined): ProbedMedia => { if (!e || 'refused' in e) throw new Error(`not probed: ${JSON.stringify(e)}`); return e }
const refusal = (e: DocMediaEntry | undefined) => (!e ? 'absent' : 'refused' in e ? e.refused : 'probed')

// --- pure ------------------------------------------------------------------------------------------------------

describe('sniffing (D7)', () => {
  it('sniffs each fixture to its demuxer; stubs, hls, concat, text → undefined', () => {
    expect(sniffDemuxer(fx('a.m4a'))).toBe('mov')
    expect(sniffDemuxer(fx('av.mp4'))).toBe('mov')
    expect(sniffDemuxer(fx('dref.mov'))).toBe('mov')
    expect(sniffDemuxer(fx('a.wav'))).toBe('wav')
    expect(sniffDemuxer(fx('a.wma'))).toBe('asf')
    expect(sniffDemuxer(fx('a.mp3'))).toBe('mp3')
    expect(sniffDemuxer(new Uint8Array([0x49, 0x44, 0x33, 4, 0]))).toBe('mp3')
    const avi = new Uint8Array(64); avi.set(new TextEncoder().encode('RIFF\0\0\0\0AVI '))
    expect(sniffDemuxer(avi)).toBe('avi')
    for (const n of ['hls.m3u8', 'concat.txt']) expect(sniffDemuxer(fx(n))).toBeUndefined()
    for (const s of ['#EXTM3U\n', 'ffconcat version 1.0\n', 'concat:a.mp3|b.mp3', 'file:/etc/hosts', 'hello world, plain text']) expect(sniffDemuxer(new TextEncoder().encode(s))).toBeUndefined()
    expect(sniffDemuxer(new Uint8Array(10))).toBeUndefined()
    expect(sniffDemuxer(new Uint8Array(0))).toBeUndefined()
  })
  it('the ftyp, RIFF and GUID checks read only the first 64 bytes', () => {
    const late = new Uint8Array(200); late.set(new TextEncoder().encode('ftyp'), 100)
    expect(sniffDemuxer(late)).toBeUndefined()
  })
})

describe('probe argv (Task 2 review: reads are not confined by the sandbox)', () => {
  it('an allowlisted -f, -protocol_whitelist file, and only the staged absolute file: path as input', () => {
    for (const d of ['mov', 'avi', 'wav', 'asf', 'mp3'] as const) {
      const a = probeArgv(d, '/tmp/ov-media-1/abc.bin')
      const i = a.indexOf('-i')
      expect(a.slice(i)).toEqual(['-i', 'file:/tmp/ov-media-1/abc.bin'])
      expect(a.filter((x) => x === '-i')).toHaveLength(1)
      expect(a[a.indexOf('-f') + 1]).toBe(d)
      expect(a.indexOf('-f')).toBeLessThan(i)
      expect(a[a.indexOf('-protocol_whitelist') + 1]).toBe('file')
      expect(a.indexOf('-protocol_whitelist')).toBeLessThan(i)
      expect(a.includes('-enable_drefs')).toBe(d === 'mov')
      if (d === 'mov') expect(a[a.indexOf('-enable_drefs') + 1]).toBe('0')
      expect(a[a.indexOf('-codec_whitelist') + 1]).toContain('pcm_s16le')
    }
    expect(() => inputArgs('hls' as never, '/tmp/x.bin')).toThrow()
    expect(() => inputArgs('mov', 'relative.bin')).toThrow()
    expect(() => inputArgs('mov', '/tmp/a\nb')).toThrow()
  })
  it('durations are floor(container ms), parsed exactly', () => {
    expect(durationMsOf('2.000000')).toBe(2000)
    expect(durationMsOf('1.999000')).toBe(1999)
    expect(durationMsOf('1.9999')).toBe(1999)
    expect(durationMsOf('0.0005')).toBe(0)
    expect(durationMsOf('3')).toBe(3000)
    for (const v of ['N/A', 'inf', '-1.0', '1e3', '', undefined, 2]) expect(durationMsOf(v)).toBeUndefined()
  })
  it('parseProbe needs a read packet, a finite positive duration; cover art is not video', () => {
    const j = (streams: unknown[], duration: unknown = '2.000000') => JSON.stringify({ streams, format: { duration } })
    expect(parseProbe(j([{ codec_type: 'audio', nb_read_packets: '45' }]))).toEqual({ kind: 'audio', hasAudio: true, durationMs: 2000 })
    expect(parseProbe(j([{ codec_type: 'video', nb_read_packets: '20' }, { codec_type: 'audio', nb_read_packets: '45' }]))).toEqual({ kind: 'video', hasAudio: true, durationMs: 2000 })
    expect(parseProbe(j([{ codec_type: 'video', nb_read_packets: '1', disposition: { attached_pic: 1 } }, { codec_type: 'audio', nb_read_packets: '9' }]))?.kind).toBe('audio')
    expect(parseProbe(j([{ codec_type: 'audio' }]))).toBeUndefined() // the dref case: header only
    expect(parseProbe(j([{ codec_type: 'audio', nb_read_packets: '0' }]))).toBeUndefined()
    expect(parseProbe(j([{ codec_type: 'audio', nb_read_packets: '3' }], 'N/A'))).toBeUndefined()
    expect(parseProbe(j([{ codec_type: 'audio', nb_read_packets: '3' }], '0.000000'))).toBeUndefined()
    expect(parseProbe('{}')).toBeUndefined()
    expect(parseProbe('not json')).toBeUndefined()
  })
})

// --- with ffmpeg -----------------------------------------------------------------------------------------------

describe.skipIf(!info)('the sandboxed probe (D7, D21)', () => {
  it('a render worker can spawn the sandboxed ffprobe', async () => {
    // D21's unverified risk, first: Pool → render-worker.ts → runSandboxed → ffprobe, with `info` from the init.
    const pool = new Pool({ workers: 1, media: { info: info!, dir: temp.dir } })
    try {
      await pool.open('w', await mediaDeck([[{ spid: 3, file: 'a.m4a', use: 'audio' }]]), 'visible-only')
      const r = await pool.mediaRefs('s1', 'w', 'visible-only', 0)
      expect(r.refs).toHaveLength(1)
      const e = probed(r.refs[0]!.entry)
      expect(e).toMatchObject({ kind: 'audio', hasAudio: true, durationMs: 2000, demuxer: 'mov', sourceHash: createHash('sha256').update(fx('a.m4a')).digest('hex') })
      expect(pool.stats().workers[0]!.probes).toBe(1)
    } finally { pool.dispose() }
    expect(temp.live()).toBe(0)
  })

  it('probes a.m4a: audio, hasAudio, durationMs = floor(container ms); and each fixture', async () => {
    const host = mediaHost()
    const files = ['a.m4a', 'a.mp3', 'a.wav', 'a.wma', 'av.mp4']
    const deck = await mediaDeck([files.map((file, i) => ({ spid: 3 + i, file, use: 'audio' as const }))])
    const { docKey } = await host.open('fx', deck, 'visible-only')
    const table = host.mediaTable(docKey)
    const parts = await table.parts()
    const of = async (ext: string) => probed(await table.entry(parts.find((p) => p.endsWith(`.${ext}`))!, { external: false }))
    expect(await of('m4a')).toMatchObject({ kind: 'audio', hasAudio: true, durationMs: 2000, demuxer: 'mov' })
    expect(await of('wav')).toMatchObject({ kind: 'audio', hasAudio: true, durationMs: 2000, demuxer: 'wav' })
    // mp3 (Xing header) and asf durations are estimated by the demuxer; 8.1.1 reports 3000 and 1999, and CI's 6.1.1
    // may differ by a frame, so these are bounded, not pinned.
    const mp3 = await of('mp3')
    expect(mp3).toMatchObject({ kind: 'audio', hasAudio: true, demuxer: 'mp3' })
    expect(Math.abs(mp3.durationMs - 3000)).toBeLessThanOrEqual(60)
    const wma = await of('wma')
    expect(wma).toMatchObject({ kind: 'audio', hasAudio: true, demuxer: 'asf' })
    expect(Math.abs(wma.durationMs - 2000)).toBeLessThanOrEqual(60)
    expect(await of('mp4')).toMatchObject({ kind: 'video', hasAudio: true, durationMs: 2000, demuxer: 'mov' })
    for (const p of parts) expect(Number.isInteger(probed(table.peek(p)).durationMs)).toBe(true)
    expect(temp.live()).toBe(0)
  })

  it('hostile inputs are media-unreadable and read no other file', async () => {
    const secret = `ov-sentinel-${randomBytes(12).toString('hex')}`
    // Long enough that the mov's sample offsets fall inside it; the tail is 'Q's, a recognisable pattern.
    writeFileSync(DREF_SENTINEL, secret + 'Q'.repeat(65536))
    try {
      const seen: Array<{ argv: string[]; stdout: string; stderr: string }> = []
      const host = mediaHost({ observe: (r) => seen.push(r) })
      const files: Array<[string, string]> = [['trunc.m4a', 'm4a'], ['badhdr.m4a', 'm4a'], ['hls.m3u8', 'm3u8'], ['concat.txt', 'txt'], ['dref.mov', 'mov']]
      const deck = await mediaDeck([files.map(([file, ext], i) => ({ spid: 3 + i, file, ext, use: 'audio' as const }))])
      const { docKey } = await host.open('hostile', deck, 'visible-only')
      const table = host.mediaTable(docKey)
      await host.mediaRefs(docKey, 0) // loads the slide, then probes up to its highest k
      const parts = await table.parts()
      expect(parts).toHaveLength(5)
      for (const p of parts) expect(refusal(table.peek(p))).toBe('media-unreadable')
      // hls and concat never reach ffprobe (the sniffer refuses them); trunc, badhdr and dref do, and are refused.
      expect(seen).toHaveLength(3)
      for (const r of seen) {
        expect(r.stdout + r.stderr).not.toContain(secret)
        const input = r.argv[r.argv.length - 1]!
        expect(input.startsWith('file:')).toBe(true)
        expect(dirname(input.slice('file:'.length))).toBe(temp.dir)
        expect(r.argv[r.argv.indexOf('-f') + 1]).toBe('mov')
        expect(r.argv[r.argv.indexOf('-protocol_whitelist') + 1]).toBe('file')
      }
      expect(temp.live()).toBe(0)

      // Positive control: the fixture really is an external reference to the sentinel. ffmpeg recognises it with
      // drefs off, and with drefs and absolute paths on (never in production) it reads the sentinel's bytes.
      const staged = await temp.create(fx('dref.mov'))
      try {
        const limits = { cpuSeconds: 5, fileBytes: 1 << 16, memoryBytes: MEMORY_BYTES[info!.platform], wallMs: 10_000, stderrBytes: 65536 as const }
        const run = async (argv: string[]) => {
          const r = runSandboxed(info!, 'ffprobe', argv, limits, new AbortController().signal)
          let out = ''
          r.stdout.on('data', (b: Buffer) => { out += b.toString('utf8') })
          const end = new Promise<void>((ok) => r.stdout.once('close', ok))
          const [res] = await Promise.all([r.done, end])
          return { ...res, stdout: out }
        }
        const off = await run(['-v', 'warning', '-hide_banner', '-f', 'mov', '-protocol_whitelist', 'file', '-enable_drefs', '0', '-show_entries', 'format=duration', '-of', 'json', '-i', `file:${staged.path}`])
        expect(off.stderr).toContain('external track')
        expect(off.stderr).toContain(DREF_SENTINEL)
        const on = await run(['-v', 'warning', '-hide_banner', '-f', 'mov', '-enable_drefs', '1', '-use_absolute_path', '1', '-show_packets', '-show_data', '-of', 'json', '-i', `file:${staged.path}`])
        expect(on.stderr).toContain('absolute path')
        expect(on.stdout).toContain('5151 5151 5151 5151') // the packets are the sentinel's bytes
      } finally { await staged.dispose() }
    } finally { rmSync(DREF_SENTINEL, { force: true }) }
  })

  const corpus = resolve(import.meta.dir, '../../../corpus/lo-bnc591147.pptx')
  it.skipIf(!existsSync(corpus))('lo-bnc591147 10-byte stubs → media-unreadable', async () => {
    const seen: unknown[] = []
    const host = mediaHost({ observe: (r) => seen.push(r) })
    const { docKey } = await host.open('bnc', new Uint8Array(readFileSync(corpus)), 'visible-only')
    const table = host.mediaTable(docKey)
    const parts = await table.parts()
    expect(parts.length).toBeGreaterThan(0)
    for (const p of parts) expect(refusal(await table.entry(p, { external: false }))).toBe('media-unreadable')
    expect(seen).toHaveLength(0) // refused by the sniffer, never handed to ffprobe
  })

  it('external target → media-linked; never opens a socket', async () => {
    let connections = 0
    const server = createServer((s) => { connections++; s.destroy() })
    await new Promise<void>((ok) => server.listen(0, '127.0.0.1', ok))
    try {
      const port = (server.address() as { port: number }).port
      const seen: unknown[] = []
      const host = mediaHost({ observe: (r) => seen.push(r) })
      const { docKey } = await host.open('ext', await mediaDeck([[{ spid: 3, file: 'a.wav', use: 'audio' }]]), 'visible-only')
      const table = host.mediaTable(docKey)
      for (const url of [`http://127.0.0.1:${port}/a.mp3`, `file:///etc/hosts`, 'ppt/media/showMedia1.wav']) {
        expect(await table.entry(url, { external: true })).toEqual({ k: -1, partPath: url, refused: 'media-linked' })
      }
      await new Promise((r) => setTimeout(r, 100))
      expect(connections).toBe(0)
      expect(seen).toHaveLength(0)
      expect(table.probes).toBe(0)
    } finally { server.close() }
  })

  it('deck budget: after 10 s spent, later parts are media-unreadable', async () => {
    // A dilated clock (×20) and a 150 ms hold per probe: each probe costs ≥ 3 s of budget.
    const t0 = performance.now()
    const host = mediaHost({ now: () => (performance.now() - t0) * 20, hangMs: 150 })
    const deck = await mediaDeck([Array.from({ length: 6 }, (_, i) => ({ spid: 3 + i, file: 'a.wav', use: 'audio' as const }))])
    const { docKey } = await host.open('budget', deck, 'visible-only')
    const table = host.mediaTable(docKey)
    await host.mediaRefs(docKey, 0)
    const states = (await table.parts()).map((p) => refusal(table.peek(p)))
    const ok = states.findIndex((s) => s !== 'probed')
    expect(ok).toBeGreaterThanOrEqual(1)
    expect(ok).toBeLessThanOrEqual(4)
    expect(states.slice(ok).every((s) => s === 'media-unreadable')).toBe(true)
    expect(table.spentMs).toBeGreaterThanOrEqual(PROBE_DECK_MS)
    expect(temp.live()).toBe(0)
  }, 20_000)

  it('a probe held past PROBE_PART_MS is killed → media-unreadable', async () => {
    const host = mediaHost({ hangMs: PROBE_PART_MS + 500 })
    const { docKey } = await host.open('hang', await mediaDeck([[{ spid: 3, file: 'a.wav', use: 'audio' }]]), 'visible-only')
    const table = host.mediaTable(docKey)
    const [p] = await table.parts()
    expect(refusal(await table.entry(p!, { external: false }))).toBe('media-unreadable')
    expect(table.spentMs).toBeGreaterThanOrEqual(PROBE_PART_MS - 50)
    expect(table.spentMs).toBeLessThan(PROBE_PART_MS + 1500)
  })

  it('a 31-minute audio → media-too-long', async () => {
    const out = await temp.create(new Uint8Array(0))
    try {
      const limits = { cpuSeconds: 30, fileBytes: 8 << 20, memoryBytes: MEMORY_BYTES[info!.platform], wallMs: 30_000, stderrBytes: 65536 as const }
      const gen = runSandboxed(info!, 'ffmpeg', ['-nostdin', '-hide_banner', '-loglevel', 'error', '-y', '-f', 'lavfi', '-i', 'anullsrc=r=8000:cl=mono', '-t', String(AUDIO_MAX_MS / 1000 + 60), '-c:a', 'aac', '-b:a', '16k', '-f', 'ipod', `file:${out.path}`], limits, new AbortController().signal)
      gen.stdout.resume()
      expect((await gen.done).code).toBe(0)
      const host = mediaHost()
      const { docKey } = await host.open('long', await mediaDeck([[{ spid: 3, file: new Uint8Array(readFileSync(out.path)), ext: 'm4a', use: 'audio' }]]), 'visible-only')
      const table = host.mediaTable(docKey)
      const [p] = await table.parts()
      expect(refusal(await table.entry(p!, { external: false }))).toBe('media-too-long')
    } finally { await out.dispose() }
  })

  it('mediaRefs lists audio uses only: a video node\'s part is absent; av.mp4 used as a narration is present', async () => {
    const host = mediaHost()
    const deck = await mediaDeck([[{ spid: 3, file: 'av.mp4', use: 'video' }, { spid: 4, file: 'av.mp4', use: 'narration' }, { spid: 5, file: 'a.wav', use: 'none' }]])
    const { docKey } = await host.open('d22', deck, 'visible-only')
    const r = await host.mediaRefs(docKey, 0)
    expect(r.unit).toBe(0)
    expect(r.hidden).toBe(false)
    expect(r.refs.map((x) => x.spid)).toEqual(['4'])
    const e = probed(r.refs[0]!.entry)
    expect(e).toMatchObject({ kind: 'video', hasAudio: true, durationMs: 2000 })
    expect(r.refs[0]!.hiddenShape).toBe(false)
    const parts = await host.mediaTable(docKey).parts()
    expect(parts).toHaveLength(3)
    // The video's own part was probed (k order, budget independence) but is not listed.
    const videoPart = parts.find((p) => p !== e.partPath && p.endsWith('.mp4'))!
    expect(r.refs.some((x) => x.partPath === videoPart)).toBe(false)
    expect(host.mediaTable(docKey).peek(videoPart)).toBeDefined()
  })

  it('mediaRefs under full marks an audio use on a hidden shape', async () => {
    const host = mediaHost()
    const deck = await mediaDeck([[{ spid: 3, file: 'a.wav', use: 'audio' }, { spid: 4, file: 'a.wav', use: 'audio' }]])
    // Hide shape 4 (its own cNvPr), as PowerPoint writes it.
    const zip = await JSZip.loadAsync(deck)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    zip.file('ppt/slides/slide1.xml', xml.replace(/<p:cNvPr id="4"([^>]*?)(\/?)>/, '<p:cNvPr id="4"$1 hidden="1"$2>'))
    const hiddenDeck = await zip.generateAsync({ type: 'uint8array' })
    const full = await host.open('hid', hiddenDeck, 'full')
    const r = await host.mediaRefs(full.docKey, 0)
    expect(r.refs.map((x) => [x.spid, x.hiddenShape])).toEqual([['3', false], ['4', true]])
    // Visible-only strips the hidden shape: its audio use has no shape and is not listed.
    const vis = await host.open('hid', hiddenDeck, 'visible-only')
    expect((await host.mediaRefs(vis.docKey, 0)).refs.map((x) => x.spid)).toEqual(['3'])
  })

  it('the table is cached per document: two sessions probe once', async () => {
    const pool = new Pool({ workers: 2, media: { info: info!, dir: temp.dir } })
    try {
      await pool.open('once', await mediaDeck([[{ spid: 3, file: 'a.wav', use: 'audio' }, { spid: 4, file: 'a.m4a', use: 'audio' }]]), 'visible-only')
      const [a, b] = await Promise.all([pool.mediaRefs('s1', 'once', 'visible-only', 0), pool.mediaRefs('s2', 'once', 'visible-only', 0)])
      expect(a.refs).toEqual(b.refs)
      expect(a.refs).toHaveLength(2)
      expect(pool.stats().workers.reduce((n, w) => n + w.probes, 0)).toBe(2)
      await pool.mediaRefs('s3', 'once', 'visible-only', 0)
      expect(pool.stats().workers.reduce((n, w) => n + w.probes, 0)).toBe(2)
    } finally { pool.dispose() }
  })

  it('load(unit) probes in k order up to the unit\'s highest k; peek is sync; digest(unit) is unchanged by probing a later slide', async () => {
    const host = mediaHost()
    const deck = await mediaDeck([[{ spid: 3, file: 'a.wav', use: 'audio' }], [{ spid: 3, file: 'a.m4a', use: 'audio' }], [{ spid: 3, file: 'a.mp3', use: 'audio' }]])
    const { docKey } = await host.open('order', deck, 'visible-only')
    const table = host.mediaTable(docKey)
    const parts = await table.parts()
    expect(parts).toEqual([...parts].sort())
    await host.mediaRefs(docKey, 1) // loads slide 1 (k 0 and 1)
    expect(table.peek(parts[0]!)).toBeDefined()
    expect(table.peek(parts[1]!)).toBeDefined()
    expect(table.peek(parts[2]!)).toBeUndefined()
    expect(table.peek(parts[0]!) instanceof Promise).toBe(false)
    const d1 = table.digest(1)
    await host.mediaRefs(docKey, 2)
    expect(table.peek(parts[2]!)).toBeDefined()
    expect(table.digest(1)).toBe(d1)
    expect(table.digest(2)).not.toBe(d1)
    expect(table.probes).toBe(3)
  })

  it('mediaPart writes the projected part 0600 into dir; a part absent from the projection is refused', async () => {
    const pool = new Pool({ workers: 1, media: { info: info!, dir: temp.dir } })
    try {
      await pool.open('part', await mediaDeck([[{ spid: 3, file: 'a.wav', use: 'audio' }]]), 'visible-only')
      const refs = await pool.mediaRefs('s1', 'part', 'visible-only', 0)
      const partPath = refs.refs[0]!.partPath
      const w = await pool.mediaPart('s1', 'part', 'visible-only', partPath, temp.dir)
      expect(dirname(w.path)).toBe(temp.dir)
      expect(statSync(w.path).mode & 0o777).toBe(0o600)
      expect(new Uint8Array(readFileSync(w.path))).toEqual(fx('a.wav'))
      expect(w.bytes).toBe(fx('a.wav').length)
      expect(w.sourceHash).toBe(probed(refs.refs[0]!.entry).sourceHash)
      const f = temp.adopt(w.path, w.bytes)
      expect(temp.live()).toBe(1)
      await f.dispose()
      expect(temp.live()).toBe(0)
      expect(await rejection(pool.mediaPart('s1', 'part', 'visible-only', 'ppt/media/absent.wav', temp.dir))).toMatchObject({ refusal: 'media-part' })
      expect(await rejection(pool.mediaPart('s1', 'part', 'visible-only', 'ppt/slides/slide1.xml', temp.dir))).toBeInstanceOf(RenderFailed)
      expect(await rejection(pool.mediaPart('s1', 'part', 'visible-only', partPath, root))).toMatchObject({ refusal: 'media-part' })
    } finally { pool.dispose() }
  })

  it('a part over 512 MiB → media-too-large unread', async () => {
    const deck = await mediaDeck([[{ spid: 3, file: 'a.wav', use: 'audio' }]])
    // Declare the part's size past the cap in the central directory (what JSZip reads); its bytes stay small.
    const z = new Uint8Array(deck)
    const v = new DataView(z.buffer)
    const name = new TextEncoder().encode('ppt/media/showMedia1.wav')
    let patched = 0
    for (let i = 0; i + 46 < z.length; i++) {
      if (v.getUint32(i, true) !== 0x02014b50) continue
      const n = v.getUint16(i + 28, true)
      if (n === name.length && name.every((c, j) => z[i + 46 + j] === c)) { v.setUint32(i + 24, PART_MAX_BYTES + 1, true); patched++ }
    }
    expect(patched).toBe(1)
    const seen: unknown[] = []
    const host = mediaHost({ observe: (r) => seen.push(r) })
    // `full`: the projection is the identity, so the package is not rewritten (a rewrite would inflate the part).
    const { docKey } = await host.open('big', z, 'full')
    const table = host.mediaTable(docKey)
    expect(refusal(await table.entry('ppt/media/showMedia1.wav', { external: false }))).toBe('media-too-large')
    expect(table.probes).toBe(0)
    expect(seen).toHaveLength(0)
  })
})

describe('media off', () => {
  it('with no ffmpeg (info undefined), every entry is media-unavailable', async () => {
    const host = new DocumentHost()
    host.configureMedia(undefined)
    const { docKey } = await host.open('off', await mediaDeck([[{ spid: 3, file: 'a.wav', use: 'audio' }]]), 'visible-only')
    const r = await host.mediaRefs(docKey, 0)
    expect(r.refs.map((x) => refusal(x.entry))).toEqual(['media-unavailable'])
    expect(host.mediaTable(docKey).probes).toBe(0)
  })
})

describe('the show-plan keys carry media (D21)', () => {
  it('a show-plan job with media true and one with media false do not share a pool cache entry', async () => {
    const pool = new Pool({ workers: 1, testHooks: true, showDeadlineMs: 4000 })
    const job = { docRef: 'fly', contentPolicy: 'full' as const, unitIndex: 0, scale: 1 as const, motion: 'full' as const, mark: { kind: 'off' as const }, strictMarks: false }
    const grant = showBuildGrant({ w: 960, h: 720 }, 1)
    try {
      await pool.open('fly', await flyInDeck(), 'full')
      const refused = await pool.showPlan('s1', { ...job, media: false, hangMs: 5000 }, grant)
      expect('refused' in refused && refused.refused).toBe('show-plan-timeout')
      const on = await pool.showPlan('s1', { ...job, media: true }, grant)
      expect('refused' in on ? on.refused : 'built').toBe('built')
      const again = await pool.showPlan('s1', { ...job, media: false }, grant)
      expect('refused' in again && again.refused).toBe('show-plan-timeout')
    } finally { pool.dispose() }
  }, 20_000)
})

