/**
 * The across-midnight follow-up (brief cf-midnight): the two-instant capture of `stage1-clock-across-midnight`, one run
 * across one local midnight. The recipe is docs/validation/xlsx-cf-recipes-stage1.md §4; the decision `decideAcrossMidnight`
 * (midnight-case.ts). The orchestrator starts it in the background on the evening the workbook was built for:
 *
 *   bun scripts/cf-refs/capture-midnight.ts --case stage1-clock-across-midnight
 *   bun scripts/cf-refs/capture-midnight.ts --case stage1-clock-across-midnight --dry-run   (prints the schedule, touches nothing)
 *
 * It waits until each time itself. It never changes the system clock or a setting; it holds a display-awake assertion
 * (`caffeinate -d -i`) for its own lifetime only. Before every Excel interaction it checks that the user is idle
 * (`ioreg` HIDIdleTime) and the display unlocked, and ABORTS (never retries) otherwise: a capture cannot be retried.
 *
 * Schedule (case zone, which must be the host's zone):
 * - setup at 23:45 on D: open the saved workbook, activate it and Sheet1, zoom 100, scroll to A1, select M20 (outside the
 *   range). The last interaction before instant 1.
 * - instant 1 at 23:55 on D, instant 2 at 00:05 on D+1, each in this order: idle and lock check; local date check; the
 *   screenshot (`screencapture` of the main display: no Excel interaction); the readout (one script: the NOW() sentinel,
 *   every designated cell's value and DisplayFormat fill, the NOW() sentinel again; no `calculate`, no select, no scroll);
 *   the second screenshot (recorded only: did the readout repaint?); local date check.
 * - instant 3 right after instant 2: idle check; `calculate` (Calculate Now); the readout; the screenshot.
 * Between setup and instant 3 nothing touches Excel but the instant 1 readout, which is read-only; the NOW() sentinel
 * shows whether any recalculation happened (the decision refuses a readout that moved it).
 * Then the workbook is closed without saving, the sidecar and the PNGs written, and the decision printed. Nothing is
 * registered by this script.
 */
import { createHash } from 'node:crypto'
import { execFileSync, spawn } from 'node:child_process'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CfManifest, CfOracleCase } from './cases'
import {
  MIDNIGHT_CELLS, MIDNIGHT_SENTINELS, MIDNIGHT_TOKEN_FILL, decideAcrossMidnight,
  type MidnightCapture, type MidnightInstantObservation, type MidnightToken,
} from './midnight-case'
import { parseExcelColorToArgb } from './capture'
import { isDisplayLocked } from './build'
import { decodePngSync, type DecodedPng } from './verify'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
/** Seconds of user inactivity required before every Excel interaction. */
export const MIDNIGHT_MIN_IDLE_SECONDS = 120
/** Local times (hh:mm) of the steps; instant 2 is on the next day. */
export const MIDNIGHT_TIMES = { setup: '23:45', instant1: '23:55', instant2: '00:05' } as const
const MARKER_TOLERANCE = 6, FILL_TOLERANCE = 12

export interface MidnightCaptureDeps {
  now(): number
  sleepUntil(epochMs: number): Promise<void>
  /** The host's IANA zone (Excel reads the host's local time). */
  ambientZone(): string
  idleSeconds(): number
  displayLocked(): boolean
  /** Runs one script inside `tell application "Microsoft Excel"`. */
  excel(lines: string[]): string
  openWorkbook(path: string): string
  closeWorkbook(name: string): void
  /** Screenshot of the main display to `path` (no interaction with any application). */
  screenshot(path: string): void
  readFile(path: string): Uint8Array
  writeFile(path: string, data: string | Uint8Array): void
  /** Starts the display-awake assertion; returns its release. */
  keepAwake(): () => void
  log(message: string): void
}

/** The local date (YYYY-MM-DD) of an instant in a zone. */
export function localDateOf(epochMs: number, zone: string): string {
  return new Intl.DateTimeFormat('en-CA', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(epochMs))
}
const nextIso = (iso: string): string => {
  const [y, m, d] = iso.split('-').map(Number)
  return new Date(Date.UTC(y!, m! - 1, d! + 1)).toISOString().slice(0, 10)
}
/** The instant of a local wall time in a zone with a fixed offset (the case zones have no daylight saving time). */
export function zonedInstant(iso: string, hhmm: string, zone: string): number {
  const guess = Date.parse(`${iso}T${hhmm}:00Z`)
  const offsetMs = (at: number): number => {
    const p = new Intl.DateTimeFormat('en-US', { timeZone: zone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).formatToParts(new Date(at))
    const v = (t: string) => Number(p.find(x => x.type === t)!.value)
    return Date.UTC(v('year'), v('month') - 1, v('day'), v('hour'), v('minute'), v('second')) - at
  }
  return guess - offsetMs(guess)
}
export function midnightSchedule(anchorIso: string, zone: string): { setup: number; instant1: number; instant2: number } {
  return {
    setup: zonedInstant(anchorIso, MIDNIGHT_TIMES.setup, zone),
    instant1: zonedInstant(anchorIso, MIDNIGHT_TIMES.instant1, zone),
    instant2: zonedInstant(nextIso(anchorIso), MIDNIGHT_TIMES.instant2, zone),
  }
}

const q = (s: string) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')
/** The setup script: the last interaction before instant 1. */
export function midnightSetupScript(workbook: string, sheet: string): string[] {
  return [
    `set wb to workbook "${q(workbook)}"`,
    'activate object wb',
    `set ws to worksheet "${q(sheet)}" of wb`,
    'activate object ws',
    'set zoom of active window to 100',
    'set scroll row of active window to 1',
    'set scroll column of active window to 1',
    'select range "M20" of ws',
    'return "ok"',
  ]
}
/** The readout: the NOW() sentinel, every cell's value and DisplayFormat fill, the sentinel again. `calculate` only at instant 3. */
export function midnightReadoutScript(workbook: string, sheet: string, refs: readonly string[], calculateFirst: boolean): string[] {
  const cell = (r: string): string[] => [
    `set v to ""`, `try`, `  set v to ((value of range "${q(r)}" of ws) as string)`, `end try`,
    `set df to "missing value"`, `try`, `  set df to (color of interior object of display format of range "${q(r)}" of ws as string)`, `end try`,
    `set end of outList to ("${q(r)}" & ":::" & v & ":::" & df)`,
  ]
  const sentinel = (label: string): string[] => [`set s to ""`, `try`, `  set s to ((value of range "${MIDNIGHT_SENTINELS.now}" of ws) as string)`, `end try`, `set end of outList to ("${label}" & ":::" & s & ":::")`]
  return [
    `set wb to workbook "${q(workbook)}"`,
    `set ws to worksheet "${q(sheet)}" of wb`,
    ...(calculateFirst ? ['calculate'] : []),
    'set outList to {}',
    ...sentinel('nowBefore'),
    ...refs.flatMap(cell),
    ...sentinel('nowAfter'),
    `set AppleScript's text item delimiters to "|||"`,
    'return (outList as string)',
  ]
}
export function parseMidnightReadout(out: string): { nowBefore: string; nowAfter: string; cells: Record<string, { value: string; fill: string | null }> } {
  const cells: Record<string, { value: string; fill: string | null }> = {}
  let nowBefore = '', nowAfter = ''
  for (const item of out.trim().split('|||')) {
    const [ref = '', value = '', fill = ''] = item.split(':::')
    if (ref === 'nowBefore') nowBefore = value
    else if (ref === 'nowAfter') nowAfter = value
    else if (ref) cells[ref] = { value, fill: parseExcelColorToArgb(fill) }
  }
  return { nowBefore, nowAfter, cells }
}
/** A DisplayFormat fill as a token: a DXF fill, white or no fill (none), anything else unknown (never a match). */
export function tokenOfFill(argb: string | null): MidnightToken | 'unknown' {
  if (argb === null || argb === 'FFFFFFFF') return 'none'
  for (const [t, fill] of Object.entries(MIDNIGHT_TOKEN_FILL)) if (fill === argb) return t as MidnightToken
  return 'unknown'
}

/**
 * The designated cells' tokens in a screenshot: the marker fill's bounding box is the range's A1 top-left to K12
 * bottom-right (uniform columns and rows), each cell sampled left of its text; undefined when the markers are not found
 * exactly (the screenshot is then unmeasurable).
 */
export function screenTokens(png: DecodedPng): Record<string, MidnightToken | 'unknown'> | undefined {
  const target = [0x13, 0xf7, 0xa9]
  let x0 = Infinity, y0 = Infinity, x1 = -1, y1 = -1, n = 0
  for (let y = 0; y < png.height; y++) for (let x = 0; x < png.width; x++) {
    const i = (y * png.width + x) * 4
    if (Math.abs(png.data[i]! - target[0]!) <= MARKER_TOLERANCE && Math.abs(png.data[i + 1]! - target[1]!) <= MARKER_TOLERANCE && Math.abs(png.data[i + 2]! - target[2]!) <= MARKER_TOLERANCE) {
      n++; x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y)
    }
  }
  if (n === 0) return undefined
  const cw = (x1 - x0 + 1) / 11, ch = (y1 - y0 + 1) / 12
  // Two marker cells: between 1.5 and 2 cells of marker pixels (gridlines and antialiasing take a little).
  if (cw < 8 || ch < 8 || n < 1.5 * cw * ch || n > 2.05 * cw * ch) return undefined
  const out: Record<string, MidnightToken | 'unknown'> = {}
  for (const c of MIDNIGHT_CELLS) {
    const m = /^([A-Z])(\d+)$/.exec(c.ref)!
    const col = m[1]!.charCodeAt(0) - 65, row = Number(m[2]) - 1
    let r = 0, g = 0, b = 0, k = 0
    for (let y = Math.round(y0 + (row + 0.35) * ch); y <= Math.round(y0 + (row + 0.65) * ch); y++) for (let x = Math.round(x0 + (col + 0.1) * cw); x <= Math.round(x0 + (col + 0.3) * cw); x++) {
      const i = (y * png.width + x) * 4
      r += png.data[i]!; g += png.data[i + 1]!; b += png.data[i + 2]!; k++
    }
    const avg = [r / k, g / k, b / k]
    const near = (hex: string) => [0, 1, 2].every(j => Math.abs(avg[j]! - parseInt(hex.slice(2 + 2 * j, 4 + 2 * j), 16)) <= FILL_TOLERANCE)
    out[c.ref] = near('FFFFFFFF') ? 'none' : (Object.entries(MIDNIGHT_TOKEN_FILL).find(([, f]) => near(f))?.[0] as MidnightToken | undefined) ?? 'unknown'
  }
  return out
}

export class MidnightAbort extends Error {}

/** One instant: checks, screenshot, readout, screenshot, checks. */
async function instant(deps: MidnightCaptureDeps, c: CfOracleCase, workbook: string, label: string, expectedDate: string, shots: readonly string[], calculate: boolean): Promise<MidnightInstantObservation & { startedAt: string; endedAt: string; raw: unknown }> {
  const zone = c.clockRequirement!.requiredTimeZone
  const guard = (what: string) => {
    if (deps.displayLocked()) throw new MidnightAbort(`${label}: the display is locked (${what})`)
    const idle = deps.idleSeconds()
    if (!(idle >= MIDNIGHT_MIN_IDLE_SECONDS)) throw new MidnightAbort(`${label}: the user is active (idle ${idle}s < ${MIDNIGHT_MIN_IDLE_SECONDS}s) before ${what}`)
  }
  const startedAt = deps.now()
  const localDateBefore = localDateOf(startedAt, zone)
  if (localDateBefore !== expectedDate) throw new MidnightAbort(`${label}: the local date is ${localDateBefore}, not ${expectedDate}`)
  const refs = MIDNIGHT_CELLS.map(x => x.ref)
  let screen: Record<string, MidnightToken | 'unknown'> | undefined, screenAfter: Record<string, MidnightToken | 'unknown'> | undefined
  if (!calculate) { deps.screenshot(resolve(REPO_ROOT, shots[0]!)); screen = screenTokens(decodePngSync(deps.readFile(resolve(REPO_ROOT, shots[0]!)))!) }
  guard('the readout')
  const raw = deps.excel(midnightReadoutScript(workbook, c.sheet, [...refs, MIDNIGHT_SENTINELS.today], calculate))
  const readout = parseMidnightReadout(raw)
  const shotAfter = calculate ? shots[0]! : shots[1]!
  deps.screenshot(resolve(REPO_ROOT, shotAfter))
  screenAfter = screenTokens(decodePngSync(deps.readFile(resolve(REPO_ROOT, shotAfter)))!)
  if (calculate) { screen = screenAfter; screenAfter = undefined }
  const endedAt = deps.now()
  const localDateAfter = localDateOf(endedAt, zone)
  const displayFormat = Object.fromEntries(refs.map(r => [r, tokenOfFill(readout.cells[r]?.fill ?? null)])) as Record<string, MidnightToken>
  const cellValues = Object.fromEntries(Object.entries(readout.cells).map(([r, v]) => [r, v.value]))
  return {
    localDateBefore, localDateAfter, displayFormat,
    ...(screen ? { screen: screen as Record<string, MidnightToken> } : {}), ...(screenAfter ? { screenAfter: screenAfter as Record<string, MidnightToken> } : {}),
    nowBefore: readout.nowBefore, nowAfter: readout.nowAfter, cellValues,
    startedAt: new Date(startedAt).toISOString(), endedAt: new Date(endedAt).toISOString(), raw,
  }
}

/** The whole run. Throws `MidnightAbort` (after closing the workbook it opened) on any refused check. */
export async function runMidnightCapture(c: CfOracleCase, deps: MidnightCaptureDeps): Promise<{ capture: MidnightCapture; sidecar: Record<string, unknown> }> {
  if (!c.acrossMidnight || !c.clockRequirement) throw new MidnightAbort(`${c.id}: not an across-midnight case`)
  const zone = c.clockRequirement.requiredTimeZone, D = c.clockRequirement.anchorLocalDate, D1 = nextIso(D)
  if (deps.ambientZone() !== zone) throw new MidnightAbort(`the host zone is ${deps.ambientZone()}, not the case zone ${zone}`)
  const start = deps.now()
  if (localDateOf(start, zone) !== D) throw new MidnightAbort(`the host's local date is ${localDateOf(start, zone)}, but the case was built for ${D}: rebuild it for tonight first`)
  const at = midnightSchedule(D, zone)
  if (start > at.setup) throw new MidnightAbort(`started at ${new Date(start).toISOString()}, after the setup time ${MIDNIGHT_TIMES.setup}`)
  const shots = c.acrossMidnight.screenshotPaths
  const release = deps.keepAwake()
  let workbook: string | undefined
  try {
    await deps.sleepUntil(at.setup)
    if (deps.displayLocked()) throw new MidnightAbort('setup: the display is locked')
    if (!(deps.idleSeconds() >= MIDNIGHT_MIN_IDLE_SECONDS)) throw new MidnightAbort('setup: the user is active')
    if (localDateOf(deps.now(), zone) !== D) throw new MidnightAbort('setup: the local date changed')
    workbook = deps.openWorkbook(resolve(REPO_ROOT, c.saved))
    deps.excel(midnightSetupScript(workbook, c.sheet))
    deps.log(`setup done at ${new Date(deps.now()).toISOString()}`)
    await deps.sleepUntil(at.instant1)
    const i1 = await instant(deps, c, workbook, 'instant 1', D, [shots[0]!, shots[1]!], false)
    deps.log(`instant 1 done at ${i1.endedAt}`)
    await deps.sleepUntil(at.instant2)
    const i2 = await instant(deps, c, workbook, 'instant 2', D1, [shots[2]!, shots[3]!], false)
    deps.log(`instant 2 done at ${i2.endedAt}`)
    let i3: Awaited<ReturnType<typeof instant>> | undefined
    try {
      i3 = await instant(deps, c, workbook, 'instant 3', D1, [shots[4]!], true)
    } catch (e) {
      if (!(e instanceof MidnightAbort)) throw e
      deps.log(`instant 3 not taken: ${e.message}`)
    }
    const strip = (o: Awaited<ReturnType<typeof instant>>): MidnightInstantObservation => ({ localDateBefore: o.localDateBefore, localDateAfter: o.localDateAfter, displayFormat: o.displayFormat, ...(o.screen ? { screen: o.screen } : {}), ...(o.screenAfter ? { screenAfter: o.screenAfter } : {}), nowBefore: o.nowBefore, nowAfter: o.nowAfter, cellValues: o.cellValues })
    const capture: MidnightCapture = { anchorLocalDate: D, timeZone: zone, instants: [strip(i1), strip(i2), i3 ? strip(i3) : undefined] }
    const sha = (p: string) => createHash('sha256').update(deps.readFile(resolve(REPO_ROOT, p))).digest('hex')
    const sidecar = {
      caseId: c.id, anchorLocalDate: D, timeZone: zone, savedSha256: createHash('sha256').update(deps.readFile(resolve(REPO_ROOT, c.saved))).digest('hex'),
      instants: [i1, i2, ...(i3 ? [i3] : [])],
      screenshots: Object.fromEntries(shots.filter((_p, i) => i < 4 || i3).map(p => [p, sha(p)])),
      decision: decideAcrossMidnight(capture),
    }
    return { capture, sidecar }
  } finally {
    if (workbook) { try { deps.closeWorkbook(workbook) } catch (e) { deps.log(`close failed: ${String(e)}`) } }
    release()
  }
}

/** The production dependencies (real clock, osascript, ioreg, screencapture, caffeinate). */
export function realMidnightDeps(): MidnightCaptureDeps {
  return {
    now: () => Date.now(),
    sleepUntil: async (t: number) => { while (Date.now() < t) await new Promise(r => setTimeout(r, Math.min(30_000, Math.max(50, t - Date.now())))) },
    ambientZone: () => Intl.DateTimeFormat().resolvedOptions().timeZone,
    idleSeconds: () => {
      try {
        const out = execFileSync('ioreg', ['-c', 'IOHIDSystem', '-d', '4'], { encoding: 'utf8', timeout: 5_000 })
        const m = /"HIDIdleTime" = (\d+)/.exec(out)
        return m ? Number(m[1]) / 1e9 : 0
      } catch { return 0 }
    },
    displayLocked: () => isDisplayLocked(),
    excel: lines => execFileSync('osascript', ['-e', 'tell application "Microsoft Excel"', ...lines.flatMap(l => ['-e', l]), '-e', 'end tell'], { encoding: 'utf8', timeout: 120_000 }).trim(),
    openWorkbook: path => {
      const name = path.split('/').pop()!
      execFileSync('open', ['-a', 'Microsoft Excel', path])
      const deadline = Date.now() + 60_000
      for (;;) {
        const names = execFileSync('osascript', ['-e', 'tell application "Microsoft Excel" to get name of every workbook'], { encoding: 'utf8' }).trim()
        if (names.split(', ').includes(name)) return name
        if (Date.now() > deadline) throw new MidnightAbort(`open ${name} timed out`)
        execFileSync('sleep', ['0.5'])
      }
    },
    closeWorkbook: name => { execFileSync('osascript', ['-e', `tell application "Microsoft Excel" to close workbook "${q(name)}" saving no`], { timeout: 30_000 }) },
    screenshot: path => { mkdirSync(dirname(path), { recursive: true }); execFileSync('screencapture', ['-x', '-m', '-t', 'png', path], { timeout: 30_000 }) },
    readFile: path => new Uint8Array(readFileSync(path)),
    writeFile: (path, data) => { mkdirSync(dirname(path), { recursive: true }); writeFileSync(path, data) },
    keepAwake: () => {
      const child = spawn('caffeinate', ['-d', '-i', '-w', String(process.pid)], { stdio: 'ignore', detached: false })
      return () => { try { child.kill() } catch { /* already gone */ } }
    },
    log: m => console.log(`[cf-refs/capture-midnight] ${m}`),
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2)
  const id = args.includes('--case') ? args[args.indexOf('--case') + 1] : 'stage1-clock-across-midnight'
  const manifest = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  const c = manifest.cases.find(x => x.id === id)
  if (!c?.acrossMidnight || !c.clockRequirement) throw new Error(`${id}: not an across-midnight case`)
  const at = midnightSchedule(c.clockRequirement.anchorLocalDate, c.clockRequirement.requiredTimeZone)
  console.log(`[cf-refs/capture-midnight] ${id} built for ${c.clockRequirement.anchorLocalDate} (${c.clockRequirement.requiredTimeZone}): setup ${new Date(at.setup).toISOString()}, instant 1 ${new Date(at.instant1).toISOString()}, instant 2 ${new Date(at.instant2).toISOString()}, instant 3 right after`)
  if (args.includes('--dry-run')) process.exit(0)
  const deps = realMidnightDeps()
  try {
    const { sidecar } = await runMidnightCapture(c, deps)
    deps.writeFile(resolve(REPO_ROOT, c.acrossMidnight.sidecarPath), `${JSON.stringify(sidecar, null, 2)}\n`)
    console.log(`[cf-refs/capture-midnight] decision: ${JSON.stringify(sidecar.decision)}`)
  } catch (e) {
    console.error(`[cf-refs/capture-midnight] ABORTED: ${(e as Error).message} (do not retry; record it)`)
    process.exit(2)
  }
}
