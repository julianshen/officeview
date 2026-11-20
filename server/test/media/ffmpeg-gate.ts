/**
 * TEST-ONLY ffmpeg gate (plan D6, pre-flight X27): not product code. Gated files do
 * `const info = await ffmpegForTests()` at top level, then `describe.skipIf(!info)`.
 * Under `OV_MEDIA_REQUIRED=1` a failing probe rejects, so those files fail instead of skipping.
 */
import { probeFfmpeg, type FfmpegInfo } from '../../src/media/ffmpeg'

let memo: Promise<FfmpegInfo | undefined> | undefined

export function ffmpegForTests(): Promise<FfmpegInfo | undefined> {
  memo ??= (async () => {
    const { OV_FFMPEG, OV_FFPROBE, OV_MEDIA } = process.env
    const p = await probeFfmpeg({ OV_FFMPEG, OV_FFPROBE, OV_MEDIA })
    if (p.ok) return p.info
    if (process.env.OV_MEDIA_REQUIRED === '1') throw new Error(`OV_MEDIA_REQUIRED=1 but the ffmpeg probe failed: ${p.reason}`)
    return undefined
  })()
  return memo
}
