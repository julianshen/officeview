/**
 * The one-command D4 capture pipeline (scripts/slideshow-capture-export.ts):
 * PowerPoint movie export (AppleScript through a mocked runner; PowerPoint is
 * never launched here), burst detection, row mapping, crop, provenance, and
 * one ffmpeg integration test on a synthesized CFR movie (skipped without
 * ffmpeg + libx264).
 */
import { spawnSync } from 'node:child_process'
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import {
  ANCHOR_FRAMES, MOVIE_STABLE_S, analyzeMovie, appleScriptFor, buildProvenance, compareDecks, cropFor, detectBursts, expectRows, exportDecks, extractDecks, extractFrame,
  mapBursts, mergeRuns, parseProbe, pngSize, realRunner, seekCheck, seekSeconds, statusChanges, steadyFrameTimes, type Burst, type ProcResult, type RowExpect, type Runner,
} from '../scripts/slideshow-capture-export'
import { parseProvenance, readRefs } from '../scripts/slideshow-compare'
import type { CapturePlan, CaptureStep } from '../scripts/slideshow-capture-plan'

const tmp: string[] = []
const scratch = () => { const d = mkdtempSync(join(tmpdir(), 'ov-capture-')); tmp.push(d); return d }
afterEach(() => { for (const d of tmp.splice(0)) rmSync(d, { recursive: true, force: true }) })

const ok = (stdout = ''): ProcResult => ({ code: 0, stdout: Buffer.from(stdout), stderr: '', timedOut: false })

function frames(slide: number, step: number, span: CaptureStep['spanMs']): CaptureStep['frames'] {
  if (typeof span !== 'number') return [{ pct: 0, ms: 0, file: `${slide}-${step}-0.png` }]
  const pcts = span === 0 ? [100] : [0, 25, 50, 75, 100]
  return pcts.map(pct => ({ pct, ms: Math.round(span * pct / 100), file: `${slide}-${step}-${pct}.png` }))
}
function step(slide: number, n: number, span: CaptureStep['spanMs'], cls: string[] = [], transition?: string): CaptureStep {
  return {
    slide, step: n, trigger: n === 0 ? 'arrive' : 'click', effects: cls.map((c, i) => ({ cls: c, id: i + 1, label: c })), spanMs: span, frames: frames(slide, n, span),
    ...(transition ? { transition: { element: transition, ns: 'p', durationMs: typeof span === 'number' ? span : 0 } } : {}),
  }
}
function plan(slides: CaptureStep[][]): CapturePlan {
  return { deck: 'd', slideSize: { cx: 12192000, cy: 6858000 }, slides: slides.map((steps, i) => ({ slide: i + 1, hidden: false, advClick: true, interactive: [], steps })) }
}
const burst = (first: number, last: number): Burst => ({ first, last })

describe('AppleScript export', () => {
  it('opens the staged copy, sets 1920x1080 movie settings, saves as movie and closes without saving, under a 1800 s timeout', () => {
    const s = appleScriptFor('/c/in.pptx', '/c/out.mov')
    expect(s).toMatch(/^with timeout of 1800 seconds/)
    expect(s).toContain('tell application "Microsoft PowerPoint"')
    expect(s).toContain('open POSIX file "/c/in.pptx"')
    expect(s).toContain('set movie frame width of s to 1920')
    expect(s).toContain('set movie frame height of s to 1080')
    expect(s).toContain('save p in (POSIX file "/c/out.mov") as save as movie')
    expect(s).toContain('close p saving no')
    expect(s).not.toMatch(/as save as MP4/i)
    expect(s.trimEnd()).toMatch(/end timeout$/)
  })

  it('waits for the background export: closes only after the movie size has held still for MOVIE_STABLE_S seconds, checked from outside PowerPoint', () => {
    const s = appleScriptFor('/c/in.pptx', '/c/out.mov')
    const save = s.indexOf('as save as movie'), wait = s.indexOf(`repeat while stable < ${MOVIE_STABLE_S}`), close = s.lastIndexOf('close p saving no')
    expect(save).toBeGreaterThan(0)
    expect(wait).toBeGreaterThan(save)
    expect(close).toBeGreaterThan(wait)
    expect(s).toContain('tell current application to set sz to (do shell script "stat -f %z " & quoted form of "/c/out.mov") as integer')
  })

  it.skipIf(process.platform !== 'darwin' || !existsSync('/Applications/Microsoft PowerPoint.app'))('compiles against PowerPoint\'s dictionary', () => {
    const out = join(scratch(), 'x.scpt')
    const r = spawnSync('osacompile', ['-o', out], { input: appleScriptFor('/c/in.pptx', '/c/out.mov') })
    expect(r.stderr.toString()).toBe('')
    expect(r.status).toBe(0)
  })

  it('PowerPoint error -9074 (it could not open the staged deck): the reason says to answer its dialog and rerun the deck', async () => {
    const root = scratch()
    const corpus = join(root, 'corpus')
    spawnSync('mkdir', ['-p', corpus])
    writeFileSync(join(corpus, 'lo-a.pptx'), 'A')
    const run: Runner = async () => ({ code: 1, stdout: Buffer.alloc(0), stderr: '72:200: execution error: Microsoft PowerPoint got an error. (-9074)', timedOut: false })
    const res = await exportDecks({ decks: ['a'], corpus, staging: join(root, 's'), videos: join(root, 'v'), timeoutMs: 5000, dryRun: false }, run, () => {})
    expect(res[0].reason).toMatch(/\(-9074\) \(PowerPoint could not open .*a\.pptx: answer any dialog it shows, such as a "Grant File Access" sheet, then rerun with --deck a\)$/)
  })

  it('quotes backslashes and double quotes in paths', () => {
    expect(appleScriptFor('/a "b"\\c.pptx', '/o.mov')).toContain('open POSIX file "/a \\"b\\"\\\\c.pptx"')
  })

  it('a timed-out deck is reported and the loop continues with the next deck; corpus files are only read', async () => {
    const root = scratch()
    const corpus = join(root, 'corpus'), staging = join(root, 'staging'), videos = join(root, 'videos')
    for (const d of [corpus]) spawnSync('mkdir', ['-p', d])
    writeFileSync(join(corpus, 'lo-a.pptx'), 'A')
    writeFileSync(join(corpus, 'lo-b.pptx'), 'B')
    const calls: Array<{ cmd: string; input?: string; timeoutMs?: number }> = []
    const run: Runner = async (cmd, _args, opts) => {
      calls.push({ cmd, input: opts?.input, timeoutMs: opts?.timeoutMs })
      if (cmd !== 'osascript') return ok()
      if (opts?.input?.includes(join(staging, 'a.pptx'))) return { code: null, stdout: Buffer.alloc(0), stderr: '', timedOut: true }
      writeFileSync(join(staging, 'b.mov'), 'movie')
      return ok('16.106\t26020821')
    }
    const log: string[] = []
    const res = await exportDecks({ decks: ['a', 'b'], corpus, staging, videos, timeoutMs: 5000, dryRun: false }, run, m => log.push(m))
    expect(res[0]).toMatchObject({ deck: 'a', ok: false })
    expect(res[0].reason).toMatch(/timed out after 5 s/)
    expect(res[0].reason).toMatch(/Grant File Access/)
    expect(res[1]).toMatchObject({ deck: 'b', ok: true, movie: join(videos, 'b.mov') })
    expect(readFileSync(join(videos, 'b.mov'), 'utf8')).toBe('movie')
    expect(JSON.parse(readFileSync(join(videos, 'b.export.json'), 'utf8'))).toMatchObject({ version: '16.106', build: '26020821' })
    expect(calls.filter(c => c.cmd === 'osascript').map(c => c.timeoutMs)).toEqual([5000, 5000])
    expect(readFileSync(join(corpus, 'lo-a.pptx'), 'utf8')).toBe('A')
    expect(existsSync(join(staging, 'b.pptx'))).toBe(false)
  })

  it('a deck whose script succeeds but writes no movie fails with the script output', async () => {
    const root = scratch()
    spawnSync('mkdir', ['-p', join(root, 'corpus')])
    writeFileSync(join(root, 'corpus', 'lo-a.pptx'), 'A')
    const run: Runner = async () => ({ code: 1, stdout: Buffer.alloc(0), stderr: 'execution error: -1708', timedOut: false })
    const [r] = await exportDecks({ decks: ['a'], corpus: join(root, 'corpus'), staging: join(root, 's'), videos: join(root, 'v'), timeoutMs: 1000, dryRun: false }, run, () => {})
    expect(r.ok).toBe(false)
    expect(r.reason).toMatch(/-1708/)
  })

  it('--dry-run prints the script and runs nothing', async () => {
    const root = scratch()
    spawnSync('mkdir', ['-p', join(root, 'corpus')])
    writeFileSync(join(root, 'corpus', 'lo-a.pptx'), 'A')
    const run: Runner = async () => { throw new Error('must not run') }
    const log: string[] = []
    const [r] = await exportDecks({ decks: ['a'], corpus: join(root, 'corpus'), staging: join(root, 's'), videos: join(root, 'v'), timeoutMs: 1000, dryRun: true }, run, m => log.push(m))
    expect(r.ok).toBe(true)
    expect(log.join('\n')).toContain('save p in (POSIX file')
    expect(existsSync(join(root, 's'))).toBe(false)
  })

  it('the real runner kills a process at its timeout', async () => {
    const r = await realRunner('sleep', ['5'], { timeoutMs: 100 })
    expect(r.timedOut).toBe(true)
  })
})

describe('probe', () => {
  const json = (r: string, avg: string) => JSON.stringify({ streams: [{ width: 1920, height: 1080, r_frame_rate: r, avg_frame_rate: avg, nb_frames: '300' }] })
  it('reads size and fps of a CFR movie', () => {
    expect(parseProbe(json('30/1', '30/1'))).toEqual({ ok: true, width: 1920, height: 1080, fps: 30, fpsText: '30/1', frames: 300 })
    expect(parseProbe(json('30000/1001', '30000/1001'))).toMatchObject({ ok: true, fps: 30000 / 1001 })
  })
  it('PowerPoint 30 fps in 1/600 s ticks (r_frame_rate 600/1): fps is the average, frame times to be checked', () => {
    expect(parseProbe(json('600/1', '9120/301'))).toMatchObject({ ok: true, fps: 9120 / 301, fpsText: '30.299', tick: 600 })
  })
  it('steady frame times: gaps of 19 and 20 ticks pass; a held frame or a dropped tick run fails', () => {
    const ticks = (gaps: number[]) => gaps.reduce<number[]>((t, g) => [...t, t[t.length - 1] + g / 600], [0])
    expect(steadyFrameTimes(ticks([20, 20, 19, 20, 20, 19, 20]), 600)).toBe(true)
    expect(steadyFrameTimes(ticks([20, 20, 19, 600, 20]), 600)).toBe(false)
    expect(steadyFrameTimes(ticks([20, 22, 20]), 600)).toBe(false)
    expect(steadyFrameTimes([0], 600)).toBe(false)
  })
  it('refuses a variable-frame-rate movie', () => {
    const r = parseProbe(json('60/1', '5734/97'))
    expect(r.ok).toBe(false)
    expect(!r.ok && r.reason).toMatch(/variable frame rate/)
  })
})

describe('burst detection', () => {
  const W = 16, H = 9, N = W * H
  const movie = (count: number, lum = 50) => new Uint8Array(count * N).fill(lum)
  const px = (_g: Uint8Array, frame: number, i: number) => frame * N + i

  it('ignores codec noise and finds a slow fade on a small region at its first and last changed frames', () => {
    const g = movie(60)
    // A keyframe at frame 10 re-quantizes a third of the pixels by one level, for good.
    for (let f = 10; f < 60; f++) for (let i = 0; i < N; i += 3) g[px(g, f, i)] += 1
    // 14 pixels brighten one level per frame over frames 30–45: too little per
    // frame for the frame-to-frame test, caught by the 250 ms lag test.
    for (let f = 30; f < 60; f++) for (let i = 1; i < 15; i++) g[px(g, f, i)] = (i % 3 === 0 ? 51 : 50) + Math.min(f, 45) - 29
    expect(detectBursts(g, W, H, 30)).toEqual([burst(30, 45)])
  })

  it('finds a small sharp local change (a wipe edge) and a cut', () => {
    const g = movie(60)
    for (let f = 5; f < 60; f++) for (let i = 0; i < 4; i++) g[px(g, f, i)] = 200
    g.fill(120, px(g, 30, 0))
    expect(detectBursts(g, W, H, 30)).toEqual([burst(5, 5), burst(30, 30)])
  })

  it('merges gaps shorter than the gap threshold and keeps longer gaps apart', () => {
    const a = Array(100).fill(false)
    for (const i of [10, 11, 12, 15, 16, 40, 41]) a[i] = true
    // 30 fps: frames 12 → 15 is a 2-frame gap (67 ms), merged; 16 → 40 is 800 ms, split
    expect(mergeRuns(a, 30, 250)).toEqual([burst(10, 16), burst(40, 41)])
  })
})

describe('row mapping', () => {
  const fps = 30

  it('media-only click steps expect no burst; arrivals and visible clicks expect one; slide 1 is anchored at the movie start', () => {
    const rows = expectRows(plan([[step(1, 0, 0), step(1, 1, 0, ['mediacall'])], [step(2, 0, 500, [], 'fade'), step(2, 1, 1000, ['entr'])]]))
    expect(rows.map(r => r.expect)).toEqual(['anchor', 'none', 'burst', 'burst'])
  })

  it('t0 is the last frame before the first changed frame; frame n = t0 + round(ms * fps / 1000); span 0 takes the settled frame', () => {
    const rows = expectRows(plan([[step(1, 0, 0)], [step(2, 0, 0)], [step(3, 0, 1000, [], 'fade'), step(3, 1, 500, ['entr'])]]))
    const m = mapBursts(rows, [burst(60, 61), burst(150, 179), burst(300, 314)], fps, 600)
    expect(m.ok).toBe(true)
    if (!m.ok) return
    const at = Object.fromEntries(m.frames.map(f => [f.file, f.index]))
    expect(at['1-0-100.png']).toBe(59) // settled: the last frame before slide 2's change
    expect(at['2-0-100.png']).toBe(149)
    expect(at['3-0-0.png']).toBe(149) // t0 = 150 - 1
    expect(at['3-0-25.png']).toBe(149 + Math.round(250 * fps / 1000))
    expect(at['3-0-100.png']).toBe(179)
    expect(at['3-1-0.png']).toBe(299)
    expect(at['3-1-50.png']).toBe(299 + 8) // 250 ms = 7.5 frames, rounds to 8
    expect(at['3-1-100.png']).toBe(314)
  })

  it('a burst at the very start belongs to slide 1 (an opening transition) and t0 clamps to frame 0', () => {
    const rows = expectRows(plan([[step(1, 0, 1000, [], 'fade'), step(1, 1, 500, ['entr'])]]))
    const m = mapBursts(rows, [burst(1, 30), burst(100, 115)], fps, 300)
    expect(m.ok && m.frames.find(f => f.file === '1-0-50.png')?.index).toBe(15)
    expect(m.ok && m.frames.find(f => f.file === '1-1-0.png')?.index).toBe(99)
    expect(ANCHOR_FRAMES).toBeGreaterThanOrEqual(1)
  })

  it('an opening transition after a black hold is slide 1\'s, with t0 the last black frame', () => {
    const rows = expectRows(plan([[step(1, 0, 1000, [], 'fade'), step(1, 1, 500, ['entr'])]]))
    expect(mapBursts(rows, [burst(10, 39), burst(100, 115)], fps, 300).ok).toBe(false) // without the dark start: one burst too many
    const m = mapBursts(rows, [burst(10, 39), burst(100, 115)], fps, 300, true)
    expect(m.ok && m.frames.find(f => f.file === '1-0-0.png')?.index).toBe(9)
    expect(m.ok && m.frames.find(f => f.file === '1-0-100.png')?.index).toBe(39)
  })

  it('an opening transition the export did not play skips those frames instead of mislabelling the still slide', () => {
    const rows = expectRows(plan([[step(1, 0, 1000, [], 'fade'), step(1, 1, 500, ['entr'])]]))
    const m = mapBursts(rows, [burst(100, 115)], fps, 300)
    expect(m.ok).toBe(true)
    if (!m.ok) return
    expect(m.frames.map(f => f.file)).toEqual(['1-1-0.png', '1-1-25.png', '1-1-50.png', '1-1-75.png', '1-1-100.png'])
    expect(m.skipped.map(s => s.file)).toContain('1-0-0.png')
    expect(m.skipped[0].reason).toMatch(/opening transition/)
  })

  it('a media-only step takes the settled frame of the step before it', () => {
    const rows = expectRows(plan([[step(1, 0, 0), step(1, 1, 0, ['mediacall'])], [step(2, 0, 0)]]))
    const m = mapBursts(rows, [burst(90, 90)], fps, 200)
    expect(m.ok && m.frames).toEqual([{ file: '1-0-100.png', index: 89 }, { file: '1-1-100.png', index: 89 }, { file: '2-0-100.png', index: 199 }])
  })

  it('refuses when the burst count does not match the rows, listing both', () => {
    const rows = expectRows(plan([[step(1, 0, 0), step(1, 1, 500, ['entr']), step(1, 2, 500, ['entr'])]]))
    const m = mapBursts(rows, [burst(30, 60)], fps, 200)
    expect(m.ok).toBe(false)
    if (m.ok) return
    expect(m.reason).toMatch(/1 activity burst\(s\) but 2 row\(s\) expect one/)
    expect(m.reason).toContain('1-1 click span 500 ms')
    expect(m.reason).toContain('1-2 click span 500 ms')
    expect(m.reason).toMatch(/burst 1: frames 30–60/)
  })

  it('refuses a burst far longer than its row (two steps merged)', () => {
    const rows = expectRows(plan([[step(1, 0, 0), step(1, 1, 500, ['entr'])]]))
    const m = mapBursts(rows, [burst(30, 90)], fps, 200)
    expect(!m.ok && m.reason).toMatch(/1-1: its burst lasts 2033 ms but the row's span is 500 ms/)
  })

  it('refuses when a frame time runs into the next step', () => {
    const rows = expectRows(plan([[step(1, 0, 0), step(1, 1, 1000, ['entr']), step(1, 2, 0, ['entr'])]]))
    const m = mapBursts(rows, [burst(30, 40), burst(50, 50)], fps, 200)
    expect(!m.ok && m.reason).toMatch(/1-1-75\.png at frame 52 is in the next step/)
  })

  it('a step with a start delay: t0 is the first changed frame less the delay', () => {
    // 1000 ms span, first 500 ms a delay: the change shows at frame 165, so the click was at frame 150
    const rows = expectRows(plan([[step(1, 0, 0), { ...step(1, 1, 1000, ['entr']), leadMs: 500 }]]))
    const m = mapBursts(rows, [burst(165, 179)], fps, 400)
    expect(m.ok).toBe(true)
    if (!m.ok) return
    const at = Object.fromEntries(m.frames.map(f => [f.file, f.index]))
    expect(at['1-1-0.png']).toBe(149)
    expect(at['1-1-25.png']).toBe(157) // in the delay, before the change
    expect(at['1-1-100.png']).toBe(179)
    expect(m.warnings).toEqual([])
  })

  it('refuses a start delay that reaches back into the previous burst', () => {
    const rows = expectRows(plan([[step(1, 0, 0), step(1, 1, 500, ['entr']), { ...step(1, 2, 2500, ['entr']), leadMs: 2000 }]]))
    const m = mapBursts(rows, [burst(30, 44), burst(80, 94)], fps, 400)
    expect(m.ok || m.reason).toMatch(/1-2: its 2000 ms start delay puts the trigger at frame 19, inside the previous burst \(ends frame 44\)/)
  })

  it('warns when a burst lasts noticeably shorter than the row span', () => {
    const rows = expectRows(plan([[step(1, 0, 0), step(1, 1, 1000, ['entr'])]]))
    const m = mapBursts(rows, [burst(30, 44)], fps, 200)
    expect(m.ok && m.warnings.join()).toMatch(/1-1: burst lasts 500 ms, row span 1000 ms/)
  })
})

describe('compare step of all', () => {
  const opts = { corpus: '/alt/corpus', refs: '/r', records: '/rec.json', dryRun: false }
  it('compares the deck from the selected corpus directory, by path', async () => {
    const calls: string[][] = []
    await compareDecks(['tdf105739'], opts, () => {}, async args => { calls.push(args); return 0 })
    expect(calls).toEqual([['/alt/corpus/lo-tdf105739.pptx', '--write', '--refs', '/r', '--records', '/rec.json']])
  })
  it('returns the decks whose comparison failed and keeps going', async () => {
    const failed = await compareDecks(['a', 'b', 'c'], opts, () => {}, async args => args[0].includes('lo-b.') ? 1 : 0)
    expect(failed).toEqual(['b'])
  })
  it('dry run only logs', async () => {
    const lines: string[] = []
    const failed = await compareDecks(['a'], { ...opts, dryRun: true }, l => lines.push(l), async () => { throw new Error('ran') })
    expect(failed).toEqual([])
    expect(lines.join()).toMatch(/\[dry-run\] would run: bun scripts\/slideshow-compare.ts \/alt\/corpus\/lo-a.pptx --write/)
  })
})

describe('crop', () => {
  it('16:9 movie of a 16:9 deck: full frame', () => expect(cropFor(1920, 1080, 12192000, 6858000)).toEqual({ w: 1920, h: 1080, x: 0, y: 0 }))
  it('16:9 movie of a 4:3 deck: the pillarbox crop 1440x1080 at x=240', () => expect(cropFor(1920, 1080, 9144000, 6858000)).toEqual({ w: 1440, h: 1080, x: 240, y: 0 }))
  it('a movie already at the deck aspect: full frame', () => expect(cropFor(1440, 1080, 9144000, 6858000)).toEqual({ w: 1440, h: 1080, x: 0, y: 0 }))
  it('a 4:3 movie of a 16:9 deck: letterbox crop', () => expect(cropFor(1440, 1080, 12192000, 6858000)).toEqual({ w: 1440, h: 810, x: 0, y: 135 }))
})

describe('provenance and frame files', () => {
  it('builds a provenance that passes parseProvenance and says it is a movie export', () => {
    const p = buildProvenance({ version: '16.106', build: '26020821', os: 'macOS 26.5.2', resolution: [1440, 1080], capturedBy: 'Julian Shen', date: '2026-11-01', fps: 30, movie: [1920, 1080] })
    expect(p.powerpoint).toBe('Microsoft PowerPoint for Mac 16.106 (26020821)')
    expect(p.notes).toContain('PowerPoint movie export (File → Export, save as movie), not a live slide show')
    expect(p.notes).toContain('30 fps')
    expect(parseProvenance(JSON.parse(JSON.stringify(p))).ok).toBe(true)
  })

  it('seeks half a frame early so the first frame at or after the seek is frame n', () => {
    expect(seekSeconds(0, 30)).toBe('0.000000')
    expect(seekSeconds(30, 30)).toBe('0.983333')
  })

  it('checks that a seek landed on the analysed frame', () => {
    const n = 160 * 90
    const gray = new Uint8Array(10 * n)
    for (let i = 0; i < 10; i++) gray.fill(i * 20, i * n, (i + 1) * n)
    expect(seekCheck(gray, 4, gray.slice(4 * n, 5 * n))).toBeUndefined()
    expect(seekCheck(gray, 4, gray.slice(5 * n, 6 * n))).toMatch(/landed on a different frame \(closest: 5/)
    expect(seekCheck(gray, 4, new Uint8Array(0))).toMatch(/no frame/)
  })

  it('reads a PNG size from its IHDR', () => {
    const b = Buffer.alloc(24)
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]).copy(b)
    b.write('IHDR', 12, 'ascii'); b.writeUInt32BE(1440, 16); b.writeUInt32BE(1080, 20)
    expect(pngSize(b)).toEqual([1440, 1080])
    expect(pngSize(Buffer.from('nope'))).toBeUndefined()
  })

  it('lists catalogue rows whose status moved', () => {
    const before = new Map([['a', 'unverified'], ['b', 'unverified'], ['c', 'verified']] as const)
    const after = new Map([['a', 'verified'], ['b', 'failed'], ['c', 'verified']] as const)
    expect(statusChanges(before, after)).toEqual([{ row: 'a', from: 'unverified', to: 'verified' }, { row: 'b', from: 'unverified', to: 'failed' }])
  })
})

const hasX264 = (() => {
  const r = spawnSync('ffmpeg', ['-hide_banner', '-encoders'], { encoding: 'utf8' })
  return r.status === 0 && /libx264/.test(r.stdout) && spawnSync('ffprobe', ['-version']).status === 0
})()

describe.skipIf(!hasX264)('ffmpeg integration (synthesized CFR movie)', () => {
  it('finds the cut and the fade at their known frames and extracts frames at t0 + ms', async () => {
    const dir = scratch()
    const mov = join(dir, 'synth.mov')
    // 30 fps, 4 s: luma 40 until 1.0 s (cut to 200 at frame 30), fade 200 → 50 over
    // 2.0–2.5 s (frames 61–75 change), then static. x264 with 1 s GOPs adds
    // real compression noise at every keyframe.
    const lum = "if(lt(T,1),40,if(lt(T,2),200,if(lt(T,2.5),200-(T-2)*300,50)))"
    const gen = spawnSync('ffmpeg', ['-v', 'error', '-y', '-f', 'lavfi', '-i', 'color=c=black:s=320x180:r=30:d=4',
      '-vf', `geq=lum='${lum}':cb=128:cr=128`, '-c:v', 'libx264', '-g', '30', '-pix_fmt', 'yuv420p', mov])
    expect(gen.status, gen.stderr?.toString()).toBe(0)

    const crop = cropFor(320, 180, 12192000, 6858000)
    const a = await analyzeMovie(mov, crop, realRunner)
    if (!a.ok) throw new Error(a.reason)
    expect(a.probe.fps).toBe(30)
    expect(a.bursts).toEqual([burst(30, 30), burst(61, 75)])

    const rows: RowExpect[] = expectRows(plan([[step(1, 0, 0)], [step(2, 0, 0), step(2, 1, 500, ['entr'])]]))
    const m = mapBursts(rows, a.bursts, a.probe.fps, a.frameCount)
    if (!m.ok) throw new Error(m.reason)
    const at = Object.fromEntries(m.frames.map(f => [f.file, f.index]))
    expect(at['2-1-0.png']).toBe(60)
    expect(at['2-1-25.png']).toBe(64)

    // The extracted PNG is exactly frame 64: luma 200 - (64/30 - 2) * 300 = 160.
    const out = join(dir, '2-1-25.png')
    const landed = await extractFrame(mov, at['2-1-25.png'], a.probe.fps, crop, out, realRunner)
    expect(seekCheck(a.gray, 64, landed)).toBeUndefined()
    expect(pngSize(readFileSync(out))).toEqual([320, 180])
    const { loadImage, createCanvas } = await import('canvas')
    const img = await loadImage(out)
    const ctx = createCanvas(img.width, img.height).getContext('2d')
    ctx.drawImage(img, 0, 0)
    const g = ctx.getImageData(160, 90, 1, 1).data[1]
    // limited-range luma 160 is full-range (160 - 16) * 255 / 219 ≈ 168; neighbours are ±10 away
    expect(Math.abs(g - 168)).toBeLessThanOrEqual(3)
  }, 30_000)

  it('codec noise on a still, detailed slide is not a burst', async () => {
    const dir = scratch()
    const png = join(dir, 'still.png'), mov = join(dir, 'still.mov')
    // testsrc2: text, gradients and sharp edges; CRF 28 is noisier than any export.
    expect(spawnSync('ffmpeg', ['-v', 'error', '-y', '-f', 'lavfi', '-i', 'testsrc2=s=1920x1080', '-frames:v', '1', png]).status).toBe(0)
    expect(spawnSync('ffmpeg', ['-v', 'error', '-y', '-loop', '1', '-i', png, '-t', '4', '-r', '30', '-c:v', 'libx264', '-crf', '28', '-g', '30', '-pix_fmt', 'yuv420p', mov]).status).toBe(0)
    const a = await analyzeMovie(mov, cropFor(1920, 1080, 12192000, 6858000), realRunner)
    expect(a.ok && a.bursts).toEqual([])
  }, 30_000)
})

const CORPUS = join(dirname(fileURLToPath(import.meta.url)), '../corpus')
describe.skipIf(!hasX264 || !existsSync(join(CORPUS, 'lo-tdf105739.pptx')))('extract end to end (synthesized movie for a corpus deck)', () => {
  it('writes the deck\'s frames and a provenance.json that readRefs accepts', async () => {
    const dir = scratch()
    const videos = join(dir, 'videos'), refs = join(dir, 'refs')
    spawnSync('mkdir', ['-p', videos])
    // tdf105739: one 16:9 slide, no transition, one settled frame.
    const gen = spawnSync('ffmpeg', ['-v', 'error', '-y', '-f', 'lavfi', '-i', 'color=c=gray:s=320x180:r=30:d=2', '-c:v', 'libx264', '-pix_fmt', 'yuv420p', join(videos, 'tdf105739.mov')])
    expect(gen.status).toBe(0)
    writeFileSync(join(videos, 'tdf105739.export.json'), JSON.stringify({ version: '16.106', build: '26020821' }))
    const run: Runner = (cmd, args, opts) =>
      cmd === 'git' ? Promise.resolve(ok('Test Owner\n')) : cmd === 'sw_vers' ? Promise.resolve(ok(args[0] === '-productVersion' ? '26.5.2\n' : '25F84\n')) : realRunner(cmd, args, opts)
    const [r] = await extractDecks({ decks: ['tdf105739'], corpus: CORPUS, videos, refs, gapMs: 250, dryRun: false }, run, () => {})
    expect(r.refused).toBeUndefined()
    expect(r.written).toBe(1)
    const set = await readRefs(refs)
    expect(set.problems).toEqual([])
    expect(set.frames.map(f => `${f.deck}/${f.slide}-${f.step}-${f.pct}`)).toEqual(['tdf105739/1-0-100'])
    expect(set.provenance.get('tdf105739')).toMatchObject({ powerpoint: 'Microsoft PowerPoint for Mac 16.106 (26020821)', os: 'macOS 26.5.2 (25F84)', resolution: [320, 180], capturedBy: 'Test Owner' })
  }, 30_000)
})

describe.skipIf(!existsSync(join(CORPUS, 'lo-tdf104788.pptx')))('capture plan start delay', () => {
  it('a click effect that starts 500 ms after the click: leadMs 500; effects that start at once carry none', async () => {
    const { default: JSZip } = await import('jszip')
    const { planDeck } = await import('../scripts/slideshow-capture-plan')
    const zip = await JSZip.loadAsync(readFileSync(join(CORPUS, 'lo-tdf104788.pptx')))
    const xml = await zip.file('ppt/slides/slide6.xml')!.async('string')
    const at = 'presetID="37" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="'
    expect(xml).toContain(`${at}0"/>`)
    zip.file('ppt/slides/slide6.xml', xml.replace(`${at}0"/>`, `${at}500"/>`))
    const delayed = await planDeck('tdf104788', await zip.generateAsync({ type: 'uint8array' }))
    const original = await planDeck('tdf104788', new Uint8Array(readFileSync(join(CORPUS, 'lo-tdf104788.pptx'))))
    const step = (p: CapturePlan, n: number) => p.slides.find(s => s.slide === 6)!.steps.find(s => s.step === n)!
    expect(step(delayed, 1).leadMs).toBe(500)
    expect(step(original, 1).leadMs).toBeUndefined()
    expect(step(delayed, 2).leadMs).toBeUndefined()
  })
})
