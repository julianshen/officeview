/** Piece 5 Task 2: the ffmpeg startup probe fails closed (plan D5–D7, spec §4.2). */
import { describe, expect, it } from 'bun:test'
import { AUDIO_DECODERS, MIN_MAJOR, REQUIRED_DEMUXERS, parseCodecList, parseDemuxerList, parseVersion, probeFfmpeg } from '../../src/media/ffmpeg'
import { ffmpegForTests } from './ffmpeg-gate'

describe('probeFfmpeg: fail closed', () => {
  it('absent binary → absent', async () => {
    expect(await probeFfmpeg({ OV_FFMPEG: '/nonexistent' })).toEqual({ ok: false, reason: 'absent' })
  })
  it('absent ffprobe → absent', async () => {
    expect(await probeFfmpeg({ OV_FFMPEG: process.execPath, OV_FFPROBE: '/nonexistent/ffprobe' })).toEqual({ ok: false, reason: 'absent' })
  })
  it('OV_MEDIA=off → off', async () => {
    expect(await probeFfmpeg({ OV_MEDIA: 'off' })).toEqual({ ok: false, reason: 'off' })
    expect(await probeFfmpeg({ OV_MEDIA: 'off', OV_FFMPEG: '/nonexistent' })).toEqual({ ok: false, reason: 'off' })
  })
  it('an OV_MEDIA value other than on/off → off', async () => {
    expect(await probeFfmpeg({ OV_MEDIA: 'yes' })).toEqual({ ok: false, reason: 'off' })
  })
})

describe('probe parsers', () => {
  it('parseVersion reads the release major and refuses git and unparsable builds', () => {
    expect(parseVersion('ffmpeg version 8.1.1 Copyright (c) 2000-2026 the FFmpeg developers\n')).toEqual({ version: '8.1.1', major: 8 })
    expect(parseVersion('ffmpeg version n8.0 Copyright')).toEqual({ version: 'n8.0', major: 8 })
    expect(parseVersion('ffmpeg version 6.1.1-3ubuntu5 Copyright')).toEqual({ version: '6.1.1-3ubuntu5', major: 6 })
    expect(parseVersion('ffmpeg version N-118000-gabcdef Copyright')).toBeUndefined()
    expect(parseVersion('not ffmpeg')).toBeUndefined()
  })
  it('parseDemuxerList splits comma names and keeps demux-capable rows only', () => {
    const out = 'Formats:\n D.. = Demuxing supported\n .E. = Muxing supported\n ---\n D   asf             ASF\n  E  asf_stream      ASF stream\n DE  mov,mp4,m4a,3gp,3g2,mj2 QuickTime / MOV\n'
    const s = parseDemuxerList(out)
    expect([...s].sort()).toEqual(['3g2', '3gp', 'asf', 'm4a', 'mj2', 'mov', 'mp4'])
  })
  it('parseDemuxerList reads ffmpeg 6 output: two flag columns, no device column', () => {
    // Verbatim rows of `ffmpeg -hide_banner -demuxers` from ffmpeg 6.1.1-3ubuntu5 (ubuntu:24.04 apt).
    const out = 'File formats:\n D. = Demuxing supported\n .E = Muxing supported\n --\n D  3dostr          3DO STR\n'
      + ' D  asf             ASF (Advanced / Active Streaming Format)\n  E asf_stream      ASF (Advanced / Active Streaming Format)\n'
      + ' D  avi             AVI (Audio Video Interleaved)\n D  mov,mp4,m4a,3gp,3g2,mj2 QuickTime / MOV\n'
      + ' D  mp3             MP2/3 (MPEG audio layer 2/3)\n DE wav             WAV / WAVE (Waveform Audio)\n'
    const s = parseDemuxerList(out)
    for (const d of REQUIRED_DEMUXERS) expect(s.has(d)).toBe(true)
    expect(s.has('asf_stream')).toBe(false)
    expect(s.has('3dostr')).toBe(true)
  })
  it('parseCodecList keeps audio rows only', () => {
    const out = 'Encoders:\n A..... = Audio\n ------\n V....D h264_x  X\n A....D aac                  AAC\n A....D aac_at               aac (AudioToolbox)\n'
    expect([...parseCodecList(out, 'A')].sort()).toEqual(['aac', 'aac_at'])
  })
  it('the version floor is major 6 (CI: apt 6.1.1 on ubuntu-24.04)', () => {
    expect(MIN_MAJOR).toBe(6)
    expect(parseVersion('ffmpeg version 6.1.1-3ubuntu5 Copyright')!.major).toBeGreaterThanOrEqual(MIN_MAJOR)
    expect(parseVersion('ffmpeg version 5.1.6-0+deb12u1 Copyright')!.major).toBeLessThan(MIN_MAJOR)
  })
  it('the required lists are D7 (demuxers) and the audio half of its codec whitelist (no video decoders, no libx264)', () => {
    expect([...REQUIRED_DEMUXERS]).toEqual(['mov', 'avi', 'wav', 'asf', 'mp3'])
    expect([...AUDIO_DECODERS]).toEqual(['aac', 'mp3', 'mp3float', 'pcm_s16le', 'pcm_s24le', 'pcm_u8', 'pcm_f32le', 'adpcm_ms', 'wmav2', 'alac'])
  })
})

describe('probeFfmpeg: present', () => {
  it('probe ok reports version and platform', async () => {
    const info = await ffmpegForTests()
    if (!info) return // skipped: no usable ffmpeg here (OV_MEDIA_REQUIRED=1 turns this into a failure)
    expect(parseVersion(`ffmpeg version ${info.version} `)!.major).toBeGreaterThanOrEqual(MIN_MAJOR)
    expect(info.platform).toBe(process.platform as 'linux' | 'darwin')
    expect(info.ffmpeg.startsWith('/')).toBe(true)
    expect(info.ffprobe.startsWith('/')).toBe(true)
  })
})
