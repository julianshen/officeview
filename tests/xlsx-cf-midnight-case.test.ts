// @vitest-environment node
/**
 * The across-midnight follow-up (brief cf-midnight, Phase 1 tasks 4-7): `stage1-clock-across-midnight`, pre-registered and
 * registrable, its capture and its decision, before any capture. Checked here:
 * - the layout: every pair of readings separated by >= 3 decisive cells at instant 2, controls equal in every world,
 *   instant 3 the post-recalculation world, markers and literal serials following the night the case is built for;
 * - production against the model (the rehearsal): the generated workbook loaded at 23:55 on D and prepared at 00:05 on
 *   D+1 under each injected reading paints that reading's predictions (full-recalc: its TODAY() cells withheld, the
 *   pre-committed narrowing), unregistered it withholds the clock and TODAY() cells and paints the clock-free column, and
 *   on the load day every world paints instant 1's predictions;
 * - the decision: each planted reading selects exactly itself; a no-match plant, a redraw-dependent screen, a recalculating
 *   readout, a wrong night or an unsound instant 3 register nothing;
 * - the capture run (dry, injected clock, Excel, ioreg and screenshots): the order of its steps, its refusals and aborts.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { deflateSync } from 'node:zlib'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { prepareCfTextScope } from '../src/xlsx/cf/text'
import { cfDocumentEvaluationOf } from '../src/xlsx/cf/formula'
import { CF_ACROSS_MIDNIGHT_CAP, injectCfAcrossMidnightReadings } from '../src/xlsx/cf/midnight'
import { withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { generateCaseArchive } from '../scripts/cf-refs/generate'
import { CF_CASES } from '../scripts/cf-refs/cases'
import {
  MIDNIGHT_CASE_ID, MIDNIGHT_CELLS, MIDNIGHT_INSTANT1, MIDNIGHT_INSTANT3, MIDNIGHT_MARKER, MIDNIGHT_READINGS, MIDNIGHT_RULES, MIDNIGHT_SENTINELS, MIDNIGHT_TOKEN_FILL, MIDNIGHT_WORLDS,
  decideAcrossMidnight, midnightDecisive, midnightMarkers, midnightPredictions, midnightSeparating, midnightSerial,
  type MidnightCapture, type MidnightReading, type MidnightToken, type MidnightWorld,
} from '../scripts/cf-refs/midnight-case'
import { MidnightAbort, midnightReadoutScript, midnightSchedule, midnightSetupScript, runMidnightCapture, screenTokens, type MidnightCaptureDeps } from '../scripts/cf-refs/capture-midnight'
import { decodePngSync } from '../scripts/cf-refs/verify'

const ZONE = 'Asia/Taipei'
const previousTz = process.env.TZ
beforeAll(() => { process.env.TZ = ZONE })
afterAll(() => { if (previousTz === undefined) delete process.env.TZ; else process.env.TZ = previousTz })
const NIGHT = '2026-12-03'
const template = CF_CASES.find(c => c.id === MIDNIGHT_CASE_ID)!
const caseFor = (iso: string) => ({ ...template, clockRequirement: { ...template.clockRequirement!, anchorLocalDate: iso, anchorSerial: midnightSerial(iso), staleSavedTodaySerial: midnightSerial(iso) - 5 }, markers: midnightMarkers(iso) })

describe('the layout (midnight-case.ts)', () => {
  it('is the last Stage 1 case, registrable, owning clock:across-midnight, well under 300 cells', () => {
    const stage1 = CF_CASES.filter(c => c.stage === 1)
    expect(stage1.at(-1)!.id).toBe(MIDNIGHT_CASE_ID)
    expect(template.acrossMidnight?.capabilityId).toBe(CF_ACROSS_MIDNIGHT_CAP)
    expect(template.question).toContain('registrable')
    expect(template.markers.every(m => m.capabilityId.startsWith(`${CF_ACROSS_MIDNIGHT_CAP}:`))).toBe(true)
    expect(new Set(template.markers.map(m => m.capabilityId)).size).toBe(MIDNIGHT_RULES.length)
    expect(MIDNIGHT_CELLS.length * 2 + 4).toBeLessThan(300)
  })
  it('separates every pair of readings at instant 2 at >= 3 decisive cells', () => {
    const pairs = [['frozen', 'rules-tick'], ['frozen', 'full-recalc'], ['rules-tick', 'full-recalc']] as const
    expect(Object.fromEntries(pairs.map(([a, b]) => [`${a}/${b}`, midnightSeparating(a, b)]))).toEqual({
      'frozen/rules-tick': ['B3', 'B4', 'C2', 'C3', 'D4', 'D5', 'E3', 'E4', 'H2', 'H3', 'H4', 'H5'],
      'frozen/full-recalc': ['B3', 'B4', 'C2', 'C3', 'D4', 'D5', 'E3', 'E4', 'F2', 'F3', 'F4', 'F5', 'G2', 'G3'],
      'rules-tick/full-recalc': ['F2', 'F3', 'F4', 'F5', 'G2', 'G3', 'H2', 'H3', 'H4', 'H5'],
    })
    for (const [a, b] of pairs) expect(midnightSeparating(a, b).length).toBeGreaterThanOrEqual(3)
    expect(MIDNIGHT_CELLS.filter(midnightDecisive).length).toBe(18)
  })
  it('predicts every cell under every world (the recipe table)', () => {
    const row = (w: MidnightWorld) => MIDNIGHT_CELLS.map(c => `${c.ref}:${midnightPredictions(w)[c.ref]}`).filter(s => !s.endsWith(':none')).join(' ')
    expect(row(MIDNIGHT_INSTANT1)).toBe('B3:red C2:green D4:yellow E3:red F2:green F4:yellow G3:red H2:red H4:yellow J3:green J5:green')
    expect(row(MIDNIGHT_WORLDS.frozen)).toBe(row(MIDNIGHT_INSTANT1))
    expect(row(MIDNIGHT_WORLDS['rules-tick'])).toBe('B4:red C3:green D5:yellow E4:red F2:green F4:yellow G3:red H3:green H5:green J3:green J5:green')
    expect(row(MIDNIGHT_WORLDS['full-recalc'])).toBe('B4:red C3:green D5:yellow E4:red F3:green F5:yellow G2:red H2:red H4:yellow J3:green J5:green')
    expect(row(MIDNIGHT_INSTANT3)).toBe(row(MIDNIGHT_WORLDS['full-recalc']))
  })
  it('controls (far dates under a clock rule, the clock-free column, the never-matching dates) are the same in every world and instant', () => {
    const controls = MIDNIGHT_CELLS.filter(c => !midnightDecisive(c))
    expect(controls.map(c => c.ref)).toEqual(['B2', 'B5', 'C4', 'C5', 'D2', 'D3', 'E2', 'E5', 'I2', 'I3', 'I4', 'I5', 'J2', 'J3', 'J4', 'J5'])
    for (const c of controls) {
      const t = new Set([MIDNIGHT_INSTANT1, MIDNIGHT_INSTANT3, ...MIDNIGHT_READINGS.map(r => MIDNIGHT_WORLDS[r])].map(w => midnightPredictions(w)[c.ref]))
      expect(t.size, c.ref).toBe(1)
    }
  })
  it('the generated workbook follows the night it is built for: literal dates, cellIs serials and markers', async () => {
    for (const iso of [NIGHT, '2027-01-31']) {
      const doc = await parseXlsx(await OfficePackage.load(await generateCaseArchive(MIDNIGHT_CASE_ID, iso)), { evaluation: { now: () => Date.parse(`${iso}T23:50:00+08:00`) } })
      const sheet = doc.sheets[0]!
      const D = midnightSerial(iso)
      const value = (ref: string) => sheet.rows.flatMap(r => r.cells).find(c => c.ref === ref)?.value
      expect([value('B2'), value('B3'), value('B4'), value('B5')]).toEqual([D - 1, D, D + 1, D + 2])
      expect([value('F2'), value('F4'), value('F5')]).toEqual([D, D + 1, D - 1])
      expect(sheet.conditionalFormatting!.rules.length).toBe(MIDNIGHT_RULES.length)
      const f3 = midnightMarkers(iso).find(m => m.sqref === 'F3')!
      expect(f3.formulas).toEqual([String(D + 1)])
      const fill = sheet.rows.flatMap(r => r.cells).filter(c => MIDNIGHT_MARKER.refs.includes(c.ref as 'A1')).map(c => c.style?.fillColor?.replace('#', '').toUpperCase())
      expect(fill).toEqual(['13F7A9', '13F7A9'].map(h => expect.stringContaining(h)))
    }
  })
})

/* ------------------------------------------------------------------ */
/* Production against the model.                                       */
/* ------------------------------------------------------------------ */

const at = (iso: string, hhmm: string) => Date.parse(`${iso}T${hhmm}:00+08:00`)
const next = (iso: string) => new Date(Date.UTC(Number(iso.slice(0, 4)), Number(iso.slice(5, 7)) - 1, Number(iso.slice(8, 10)) + 1)).toISOString().slice(0, 10)
const TOKEN_OF: Record<string, MidnightToken> = Object.fromEntries(Object.entries(MIDNIGHT_TOKEN_FILL).map(([t, f]) => [f, t as MidnightToken]))
/** Production's text scope at `instant` (the resident route of copy, find and accessibility): a token per ref, or `withheld`. */
async function production(iso: string, instant: number): Promise<Record<string, MidnightToken | 'withheld'>> {
  const doc = await parseXlsx(await OfficePackage.load(await generateCaseArchive(MIDNIGHT_CASE_ID, iso)), { evaluation: { now: () => at(iso, '23:50') } })
  const sheet = doc.sheets[0]!
  const scope = await prepareCfTextScope(sheet, { rects: [{ r0: 0, r1: 11, c0: 0, c1: 10 }], signal: new AbortController().signal, budget: { reserve() {}, release() {} }, documentSheets: doc.sheets, evaluation: cfDocumentEvaluationOf(doc), now: () => instant } as never)
  try {
    return Object.fromEntries(MIDNIGHT_CELLS.map(c => {
      const m = /^([A-Z])(\d+)$/.exec(c.ref)!, row = Number(m[2]) - 1, col = m[1]!.charCodeAt(0) - 65
      if (scope!.describe(row, col)?.key === 'cf-unavailable') return [c.ref, 'withheld']
      const fill = scope!.effectiveStyle(row, col, undefined)?.fillColor?.replace('#', '').toUpperCase()
      return [c.ref, fill ? TOKEN_OF[fill.length === 6 ? `FF${fill}` : fill] ?? 'none' : 'none']
    }))
  } finally { scope?.dispose() }
}
async function withWorld<T>(world: 'withhold' | MidnightReading, fn: () => Promise<T>): Promise<T> {
  if (world === 'withhold') return fn()
  injectCfAcrossMidnightReadings({ [CF_ACROSS_MIDNIGHT_CAP]: world })
  try { return await withInjectedCfEvidence([CF_ACROSS_MIDNIGHT_CAP], fn) } finally { injectCfAcrossMidnightReadings(null) }
}
const TODAY_CELLS = MIDNIGHT_CELLS.filter(c => 'today' in c.value).map(c => c.ref)
const CLOCK_RULE_CELLS = MIDNIGHT_CELLS.filter(c => { const r = MIDNIGHT_RULES[c.rule]!.rule; return r.kind !== 'cellIs' }).map(c => c.ref)

describe('production against the model (the rehearsal)', () => {
  for (const world of ['withhold', ...MIDNIGHT_READINGS] as const) {
    it(`${world}: the load day paints instant 1; 00:05 on D+1 paints the world (narrowed where pre-committed)`, async () => {
      await withWorld(world, async () => {
        const load = await production(NIGHT, at(NIGHT, '23:55'))
        expect(load).toEqual(midnightPredictions(MIDNIGHT_INSTANT1))
        const after = await production(NIGHT, at(next(NIGHT), '00:05'))
        const expected: Record<string, MidnightToken | 'withheld'> = world === 'withhold'
          ? Object.fromEntries(MIDNIGHT_CELLS.map(c => [c.ref, CLOCK_RULE_CELLS.includes(c.ref) || TODAY_CELLS.includes(c.ref) ? 'withheld' : midnightPredictions(MIDNIGHT_INSTANT1)[c.ref]!]))
          : world === 'full-recalc'
            // No route recalculates a cell after load: its TODAY() cells stay withheld (the pre-committed narrowing).
            ? Object.fromEntries(MIDNIGHT_CELLS.map(c => [c.ref, TODAY_CELLS.includes(c.ref) ? 'withheld' : midnightPredictions(MIDNIGHT_WORLDS['full-recalc'])[c.ref]!]))
            : { ...midnightPredictions(MIDNIGHT_WORLDS[world]) }
        expect(after).toEqual(expected)
      })
    })
  }
})

/* ------------------------------------------------------------------ */
/* The decision.                                                        */
/* ------------------------------------------------------------------ */

type Mutable<T> = { -readonly [K in keyof T]: T[K] }
function plant(world: MidnightWorld, edit: (c: Mutable<NonNullable<MidnightCapture['instants'][number]>>) => void = () => {}): MidnightCapture {
  const obs = (w: MidnightWorld, date: string, now: string) => ({ localDateBefore: date, localDateAfter: date, displayFormat: midnightPredictions(w), screen: midnightPredictions(w), nowBefore: now, nowAfter: now, cellValues: {} })
  const recalc = world.cellDay === 1
  const i2 = obs(world, next(NIGHT), recalc ? '46359.0003' : '46358.99')
  edit(i2)
  return { anchorLocalDate: NIGHT, timeZone: ZONE, instants: [obs(MIDNIGHT_INSTANT1, NIGHT, '46358.99'), i2, obs(MIDNIGHT_INSTANT3, next(NIGHT), '46359.0042')] }
}
describe('the decision (decideAcrossMidnight)', () => {
  for (const r of MIDNIGHT_READINGS) it(`a planted ${r} registers ${r} and only it`, () => {
    expect(decideAcrossMidnight(plant(MIDNIGHT_WORLDS[r]))).toMatchObject({ kind: 'register', reading: r })
  })
  it('a no-match plant (one decisive cell of each reading off) registers nothing: the withhold stays', () => {
    const d = decideAcrossMidnight(plant(MIDNIGHT_WORLDS['rules-tick'], i2 => { i2.displayFormat = { ...i2.displayFormat, B4: 'none' }; i2.screen = i2.displayFormat }))
    expect(d).toMatchObject({ kind: 'nothing' })
  })
  it('a screen that disagrees with its DisplayFormat (redraw-dependent) registers nothing', () => {
    expect(decideAcrossMidnight(plant(MIDNIGHT_WORLDS['rules-tick'], i2 => { i2.screen = midnightPredictions(MIDNIGHT_WORLDS.frozen) }))).toMatchObject({ kind: 'nothing', reason: expect.stringContaining('redraw') })
    expect(decideAcrossMidnight(plant(MIDNIGHT_WORLDS.frozen, i2 => { delete (i2 as { screen?: unknown }).screen }))).toMatchObject({ kind: 'nothing', reason: expect.stringContaining('unmeasurable') })
  })
  it('a readout that recalculated, a full-recalc with no recalculation, or a frozen after one register nothing', () => {
    expect(decideAcrossMidnight(plant(MIDNIGHT_WORLDS.frozen, i2 => { i2.nowAfter = '46359.0001' }))).toMatchObject({ kind: 'nothing' })
    expect(decideAcrossMidnight(plant(MIDNIGHT_WORLDS['full-recalc'], i2 => { i2.nowBefore = '46358.99'; i2.nowAfter = '46358.99' }))).toMatchObject({ kind: 'nothing' })
    expect(decideAcrossMidnight(plant(MIDNIGHT_WORLDS.frozen, i2 => { i2.nowBefore = '46359.0001'; i2.nowAfter = '46359.0001' }))).toMatchObject({ kind: 'nothing' })
  })
  it('the wrong night, a control off, or an unsound instant 3 register nothing', () => {
    expect(decideAcrossMidnight(plant(MIDNIGHT_WORLDS.frozen, i2 => { i2.localDateAfter = NIGHT }))).toMatchObject({ kind: 'nothing' })
    expect(decideAcrossMidnight(plant(MIDNIGHT_WORLDS.frozen, i2 => { i2.displayFormat = { ...i2.displayFormat, J3: 'none' }; i2.screen = i2.displayFormat }))).toMatchObject({ kind: 'nothing' })
    const unsound = plant(MIDNIGHT_WORLDS.frozen)
    const bad: MidnightCapture = { ...unsound, instants: [unsound.instants[0], unsound.instants[1], { ...unsound.instants[2]!, displayFormat: midnightPredictions(MIDNIGHT_INSTANT1) }] }
    expect(decideAcrossMidnight(bad)).toMatchObject({ kind: 'nothing', reason: expect.stringContaining('instant 3') })
    const noI3: MidnightCapture = { ...unsound, instants: [unsound.instants[0], unsound.instants[1], undefined] }
    expect(decideAcrossMidnight(noI3)).toMatchObject({ kind: 'register', reading: 'frozen' })
  })
})

/* ------------------------------------------------------------------ */
/* The capture run, dry (injected clock, Excel, ioreg, screenshots).    */
/* ------------------------------------------------------------------ */

/** A PNG of a 1100x600 'screen' with the A1:K12 grid at (100, 60), cells 80x40, the markers and each cell's token fill. */
function screenPng(tokens: Readonly<Record<string, MidnightToken>>, markers = true): Uint8Array {
  const W = 1100, H = 600, x0 = 100, y0 = 60, cw = 80, ch = 40
  const rgb = Buffer.alloc(W * H * 3, 0xff)
  const paint = (col: number, row: number, hex: string) => {
    const [r, g, b] = [0, 1, 2].map(i => parseInt(hex.slice(2 + 2 * i, 4 + 2 * i), 16))
    for (let y = y0 + row * ch + 1; y < y0 + (row + 1) * ch; y++) for (let x = x0 + col * cw + 1; x < x0 + (col + 1) * cw; x++) { const i = (y * W + x) * 3; rgb[i] = r!; rgb[i + 1] = g!; rgb[i + 2] = b! }
  }
  if (markers) { paint(0, 0, MIDNIGHT_MARKER.fillRgb); paint(10, 11, MIDNIGHT_MARKER.fillRgb) }
  for (const [ref, t] of Object.entries(tokens)) if (t !== 'none') { const m = /^([A-Z])(\d+)$/.exec(ref)!; paint(m[1]!.charCodeAt(0) - 65, Number(m[2]) - 1, MIDNIGHT_TOKEN_FILL[t]) }
  // The marker cells span the grid's corners exactly: extend the first marker to the cell's top-left edge pixel.
  for (let x = x0; x < x0 + cw; x++) { const i = (y0 * W + x) * 3; rgb.set([0x13, 0xf7, 0xa9], i) }
  for (let y = y0; y < y0 + ch; y++) { const i = (y * W + x0) * 3; rgb.set([0x13, 0xf7, 0xa9], i) }
  const raw = Buffer.alloc((W * 3 + 1) * H)
  for (let y = 0; y < H; y++) { raw[y * (W * 3 + 1)] = 0; rgb.copy(raw, y * (W * 3 + 1) + 1, y * W * 3, (y + 1) * W * 3) }
  const crc = (buf: Buffer) => { let c = ~0; for (const b of buf) { c ^= b; for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1)) } return ~c >>> 0 }
  const chunk = (type: string, data: Buffer) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]) }
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(W, 0); ihdr.writeUInt32BE(H, 4); ihdr[8] = 8; ihdr[9] = 2
  return new Uint8Array(Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]))
}
const ARGB_RGB16 = (argb: string) => `{${[0, 1, 2].map(i => parseInt(argb.slice(2 + 2 * i, 4 + 2 * i), 16) * 257).join(', ')}}`
/** An Excel stand-in: answers the readout from the world of the current instant (`world()`), recording every script. */
function fakeRun(options: { world: (now: number) => MidnightWorld; idle?: (now: number) => number; startAt?: number; overshootInstant2?: number; markers?: boolean }) {
  const D = midnightSerial(NIGHT)
  let clock = options.startAt ?? at(NIGHT, '22:00')
  const scripts: Array<{ at: number; lines: string[] }> = [], shots: Array<{ at: number; path: string }> = [], files = new Map<string, Uint8Array>(), log: string[] = []
  let opened = 0, closed = 0, recalculated = false
  const deps: MidnightCaptureDeps = {
    now: () => clock,
    sleepUntil: async t => { if (t > clock) clock = t + (t === midnightSchedule(NIGHT, ZONE).instant2 ? options.overshootInstant2 ?? 0 : 0) },
    ambientZone: () => ZONE,
    idleSeconds: () => options.idle?.(clock) ?? 3600,
    displayLocked: () => false,
    excel: lines => {
      scripts.push({ at: clock, lines }); clock += 2_000
      if (lines.includes('calculate')) recalculated = true
      if (!lines.some(l => l.includes('display format'))) return 'ok'
      const w = recalculated ? MIDNIGHT_INSTANT3 : options.world(clock)
      const now = recalculated || w.cellDay === 1 ? String(D + 1.0001) : String(D + 0.99)
      const p = midnightPredictions(w)
      return [`nowBefore:::${now}:::`, ...MIDNIGHT_CELLS.map(c => `${c.ref}:::v:::${p[c.ref] === 'none' ? 'missing value' : ARGB_RGB16(MIDNIGHT_TOKEN_FILL[p[c.ref] as 'red'])}`), `${MIDNIGHT_SENTINELS.today}:::${D}:::`, `nowAfter:::${now}:::`].join('|||')
    },
    openWorkbook: (_path, beforeEach) => { beforeEach('the open'); opened++; clock += 5_000; beforeEach('a poll of Excel while it opens the workbook'); clock += 5_000; return `${MIDNIGHT_CASE_ID}.xlsx` },
    closeWorkbook: () => { closed++ },
    screenshot: path => { shots.push({ at: clock, path }); clock += 1_000; files.set(path, screenPng(midnightPredictions(recalculated ? MIDNIGHT_INSTANT3 : options.world(clock)), options.markers ?? true)) },
    readFile: path => files.get(path) ?? new Uint8Array(Buffer.from('saved-workbook')),
    writeFile: () => {},
    keepAwake: () => () => { log.push('released') },
    log: m => { log.push(m) },
  }
  return { deps, scripts, shots, log, counts: () => ({ opened, closed }) }
}

describe('the capture run (dry)', () => {
  const c = caseFor(NIGHT)
  it('schedules setup 23:45 and instant 1 23:55 on D, instant 2 00:05 on D+1 (Asia/Taipei)', () => {
    const s = midnightSchedule(NIGHT, ZONE)
    expect(new Date(s.setup).toISOString()).toBe('2026-12-03T15:45:00.000Z')
    expect(new Date(s.instant1).toISOString()).toBe('2026-12-03T15:55:00.000Z')
    expect(new Date(s.instant2).toISOString()).toBe('2026-12-03T16:05:00.000Z')
  })
  it('the screenshot classifier reads each cell\'s token, and refuses a screen without the markers', () => {
    const p = midnightPredictions(MIDNIGHT_WORLDS['rules-tick'])
    expect(screenTokens(decodePngSync(screenPng(p))!)).toEqual(p)
    expect(screenTokens(decodePngSync(screenPng(p, false))!)).toBeUndefined()
  })
  for (const r of MIDNIGHT_READINGS) it(`a planted ${r} Excel: the run orders its steps and its decision registers ${r}`, async () => {
    const midnight = at(next(NIGHT), '00:00')
    const run = fakeRun({ world: now => now < midnight ? MIDNIGHT_INSTANT1 : MIDNIGHT_WORLDS[r] })
    const { sidecar } = await runMidnightCapture(c, run.deps)
    expect(sidecar.decision).toMatchObject({ kind: 'register', reading: r })
    // Setup, readout 1, readout 2, readout 3 (calculate first): nothing else touches Excel.
    expect(run.scripts.map(s => s.lines.includes('calculate') ? 'calculate' : s.lines.some(l => l.includes('display format')) ? 'readout' : 'setup')).toEqual(['setup', 'readout', 'readout', 'calculate'])
    expect(run.scripts[0]!.lines).toEqual(midnightSetupScript(`${MIDNIGHT_CASE_ID}.xlsx`, 'Sheet1'))
    expect(run.scripts[1]!.lines).toEqual(midnightReadoutScript(`${MIDNIGHT_CASE_ID}.xlsx`, 'Sheet1', [...MIDNIGHT_CELLS.map(x => x.ref), MIDNIGHT_SENTINELS.today], false))
    // Each instant's first screenshot precedes its readout; instant 2 starts on D+1.
    expect(run.shots[0]!.at).toBeLessThan(run.scripts[1]!.at)
    expect(run.shots[2]!.at).toBeLessThan(run.scripts[2]!.at)
    expect(run.scripts[2]!.at).toBeGreaterThanOrEqual(midnight)
    expect(run.counts()).toEqual({ opened: 1, closed: 1 })
    expect(run.log).toContain('released')
  })
  it('refuses to start on another night than the case was built for, or after the setup time; opens nothing', async () => {
    for (const startAt of [at('2026-12-02', '22:00'), at(NIGHT, '23:50')]) {
      const run = fakeRun({ world: () => MIDNIGHT_INSTANT1, startAt })
      await expect(runMidnightCapture(c, run.deps)).rejects.toBeInstanceOf(MidnightAbort)
      expect(run.counts().opened).toBe(0)
    }
  })
  it('aborts, never retries, when the user is active before instant 2\'s readout; the workbook is closed', async () => {
    const midnight = at(next(NIGHT), '00:00')
    const run = fakeRun({ world: () => MIDNIGHT_INSTANT1, idle: now => (now >= midnight ? 5 : 3600) })
    await expect(runMidnightCapture(c, run.deps)).rejects.toThrow(/user is active/)
    expect(run.scripts.filter(s => s.lines.some(l => l.includes('display format'))).length).toBe(1)
    expect(run.counts()).toEqual({ opened: 1, closed: 1 })
  })
  it('aborts, never retries, when the user becomes active between the open and the setup script; the workbook is closed (fix round 1, P2)', async () => {
    const setup = midnightSchedule(NIGHT, ZONE).setup
    // Idle at the open; active from 5 s after it (the poll while Excel opens and the setup script).
    const run = fakeRun({ world: () => MIDNIGHT_INSTANT1, idle: now => (now >= setup + 4_000 ? 3 : 3600) })
    await expect(runMidnightCapture(c, run.deps)).rejects.toThrow(/setup: the user is active.*poll/)
    expect(run.scripts.length, 'no setup script, no readout').toBe(0)
    expect(run.counts()).toEqual({ opened: 1, closed: 1 })
    // Active only after the poll: the setup script is still guarded.
    const late = fakeRun({ world: () => MIDNIGHT_INSTANT1, idle: now => (now >= setup + 9_000 ? 3 : 3600) })
    await expect(runMidnightCapture(c, late.deps)).rejects.toThrow(/setup: the user is active.*setup script/)
    expect(late.scripts.length).toBe(0)
    expect(late.counts()).toEqual({ opened: 1, closed: 1 })
  })
  it('aborts when instant 2 does not fall on exactly the next local day', async () => {
    const run = fakeRun({ world: () => MIDNIGHT_INSTANT1, overshootInstant2: 86_400_000 })
    await expect(runMidnightCapture(c, run.deps)).rejects.toThrow(/instant 2: the local date/)
  })
  it('a screen without markers makes the run register nothing (unmeasurable)', async () => {
    const run = fakeRun({ world: now => now < at(next(NIGHT), '00:00') ? MIDNIGHT_INSTANT1 : MIDNIGHT_WORLDS.frozen, markers: false })
    const { sidecar } = await runMidnightCapture(c, run.deps)
    expect(sidecar.decision).toMatchObject({ kind: 'nothing' })
  })
})

/* ------------------------------------------------------------------ */
/* The committed harness under every injected registration: same day.  */
/* ------------------------------------------------------------------ */

describe('the committed clock cases under every injected registration (the same day: nothing changes)', () => {
  const ids = ['stage1-timeperiod-today-clock', 'stage1-timeperiod-today-clock-2', 'stage5-clock-today', MIDNIGHT_CASE_ID]
  for (const world of ['withhold', ...MIDNIGHT_READINGS] as const) {
    it(`${world}: cf-compare outcomes equal the unregistered ones, interactive and paged`, async () => {
      const { compareCfCase } = await import('../scripts/cf-compare')
      const { readFileSync } = await import('node:fs')
      const { resolve } = await import('node:path')
      const root = resolve(__dirname, '..')
      const manifest = JSON.parse(readFileSync(resolve(root, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as { cases: Array<{ id: string }> }
      const outcomes = async () => {
        const out: Record<string, string> = {}
        for (const id of ids) for (const mode of ['interactive', 'paged'] as const) {
          const r = await compareCfCase(manifest.cases.find(c => c.id === id)! as never, mode, { rootDir: root })
          out[`${id} ${mode}`] = `${r.outcome}${r.outputHash ? ` ${r.outputHash}` : ''}`
        }
        return out
      }
      const base = await outcomes()
      expect(Object.values(base).every(o => /^(pass|recipe-pending)/.test(o)), JSON.stringify(base)).toBe(true)
      expect(await withWorld(world, outcomes)).toEqual(base)
    }, 300_000)
  }
})
