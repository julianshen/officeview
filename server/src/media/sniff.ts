/**
 * Piece 5 (plan D7): pick the ffmpeg demuxer from a part's first 64 bytes. Only the allowlist is ever named with
 * `-f`; anything else (playlists, `ffconcat`, `concat:`/`file:` text, the corpus 10-byte stubs) is undefined, which
 * the caller refuses as `media-unreadable`. The video containers stay on the list because a part used as audio may
 * be one (an audio track taken from a video file).
 */
export type Demuxer = 'mov' | 'avi' | 'wav' | 'asf' | 'mp3'
export const DEMUXERS: readonly Demuxer[] = Object.freeze(['mov', 'avi', 'wav', 'asf', 'mp3'] as const)
export const SNIFF_BYTES = 64

const ASF_GUID = [0x30, 0x26, 0xb2, 0x75, 0x8e, 0x66, 0xcf, 0x11]
const tag = (b: Uint8Array, at: number, s: string) => b.length >= at + s.length && [...s].every((c, i) => b[at + i] === c.charCodeAt(0))

export function sniffDemuxer(head: Uint8Array): Demuxer | undefined {
  const b = head.subarray(0, SNIFF_BYTES)
  if (tag(b, 4, 'ftyp')) return 'mov'
  if (tag(b, 0, 'RIFF') && tag(b, 8, 'AVI ')) return 'avi'
  if (tag(b, 0, 'RIFF') && tag(b, 8, 'WAVE')) return 'wav'
  if (b.length >= ASF_GUID.length && ASF_GUID.every((v, i) => b[i] === v)) return 'asf'
  if (tag(b, 0, 'ID3')) return 'mp3'
  // An MPEG audio frame sync: 0xFF then 0xEx or 0xFx.
  if (b.length >= 2 && b[0] === 0xff && (b[1]! & 0xe0) === 0xe0) return 'mp3'
  return undefined
}
