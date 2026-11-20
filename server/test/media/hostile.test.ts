/**
 * Piece 5 Task 13 (spec §4.2, §9; plan D5, D7): hostile media embedded in a deck, through the harness. Each fixture
 * is refused `media-unreadable` (on its plan record and on a `media-request`), no ffmpeg encode is started, and no
 * output file (or staged part) is left in the media temp dir.
 *
 * Not re-run here (X29 R5): the client-side kind-9 vectors (a compressed kind 9, a 4 MiB + 1 B chunk, an unsolicited
 * chunk, an overrun of `media-start` `bytes`). The harness reassembler accepts every reqId, so it cannot exercise a
 * client's rejection; Task 9's `frame-kind9.test.ts` and Task 10's `remote-media.test.ts` own them.
 */
import { afterAll, describe, expect, it } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { parseMediaRefused, parseMediaShowOk } from '../../../src/stream/media-wire'
import { narrationDeck } from '../../../src/testdata/show-decks'
import { disposePool, start } from '../harness'
import { planOf, session, settled, sidOf, spawnLog, withPart } from './battery'
import { ffmpegForTests } from './ffmpeg-gate'

const info = await ffmpegForTests()
afterAll(() => disposePool())
const FIX = resolve(import.meta.dir, '../fixtures/media')
/** The Task 3 hostile fixtures (probe.test.ts): a truncated and a bad-header m4a, an HLS playlist, an ffconcat list,
 * and a mov whose track is an external data reference. */
const HOSTILE = ['trunc.m4a', 'badhdr.m4a', 'hls.m3u8', 'concat.txt', 'dref.mov'] as const

describe.skipIf(!info)('hostile media (Task 13)', () => {
  for (const fixture of HOSTILE) {
    it(`${fixture} → media-unreadable, no ffmpeg output file left`, async () => {
      const h = await start({ doc: await withPart(await narrationDeck(), new Uint8Array(readFileSync(join(FIX, fixture)))), media: {} })
      const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
      const c = await session(h)
      try {
        const { media } = await planOf(c)
        expect(media).toHaveLength(1)
        expect(media[0]!.refused).toBe('media-unreadable')
        // Never encoded ahead.
        expect(parseMediaShowOk(await c.mediaShow(0)).queued).toBe(0)
        const r = await c.mediaRequest(media[0]!.mediaId, { ms: 30_000 })
        expect(r.start).toBeUndefined()
        expect(r.chunks).toHaveLength(0)
        expect(parseMediaRefused(r.end).reason).toBe('media-unreadable')
        expect(h.stream.audit.mediaDeliveries(sidOf(c)).map((x) => [x.reason, x.cacheState])).toEqual([['media-unreadable', undefined]])
        expect(log.ffmpeg()).toEqual([])
        expect(h.media!.cache.stats()).toMatchObject({ entries: 0, inflight: 0 })
        expect(readdirSync(h.media!.temp.dir)).toEqual([])
        await settled(h, [sidOf(c)])
      } finally { log.stop(); c.close(); await h.stop() }
    }, 60_000)
  }
})
