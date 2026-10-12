/**
 * S13 VoiceOver check (AC-A11Y-01 screen-reader leg): real macOS VoiceOver
 * reading the interactive viewer in real Safari.
 *
 *   bun scripts/harness/voiceover-check.ts --preflight     # checks only; never starts VoiceOver
 *   bun scripts/harness/voiceover-check.ts                 # starts VoiceOver (it speaks and takes focus)
 *       [--driver-port 4453] [--harness-port 4191] [--fixture-port 4194] [--out /tmp/officeview-voiceover-check]
 *
 * Safari is driven through safaridriver (page focus and keys go to WebKit);
 * VoiceOver observes what Safari's accessibility tree announces and each
 * step's phrases are read with AppleScript (`content of last phrase`, polled
 * through a short window because a later phrase replaces an earlier one).
 * Expected announcements come from spec §9 (grid name, chrome-inclusive counts
 * 1048577×16385, "B3, value …" cell names, header labels, selected/sort state,
 * the "Column filter" popup name). VoiceOver's own wording around them
 * (roles, positions) is not pinned; the raw phrases are kept in the receipt.
 *
 * Requirements (checked first, before VoiceOver is started):
 *  - the screen is unlocked;
 *  - VoiceOver Utility > General > "Allow VoiceOver to be controlled with
 *    AppleScript" is on (com.apple.VoiceOver4/default SCREnableAppleScript);
 *  - Safari remote automation is enabled (sudo safaridriver --enable, plus
 *    Safari > Settings > Developer > "Allow Remote Automation").
 * VoiceOver is always turned off in a finally block, then verified off.
 */
import { execFileSync } from 'node:child_process'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ensureServedFixtures, startFixtureServer, type FixtureServer } from './fixture-server'
import { gitIdentity } from './identity'
import { ensureHarnessServer, type ServerHandle } from './vite-server'
import { Keys, SAFARI_ENABLE_HELP, SCREEN_LOCKED_HELP, SafariDriver, looksLikeAutomationDisabled, macScreenLocked, type ElementRef } from './webdriver'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const LABEL = 'real macOS VoiceOver reading real Safari (page driven through safaridriver)'

function option(name: string, fallback: string): string {
  const index = process.argv.indexOf(`--${name}`)
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback
}
const PREFLIGHT_ONLY = process.argv.includes('--preflight')
const DRIVER_PORT = Number(option('driver-port', '4453'))
const HARNESS_PORT = Number(option('harness-port', '4191'))
const FIXTURE_PORT = Number(option('fixture-port', '4194'))
const OUT = option('out', '/tmp/officeview-voiceover-check')

const APPLESCRIPT_HELP = [
  'VoiceOver is not allowed to be controlled with AppleScript. Turn it on once, then rerun:',
  '  open VoiceOver Utility (/System/Library/CoreServices/VoiceOver Utility.app) > General >',
  '  "Allow VoiceOver to be controlled with AppleScript".',
].join('\n')

function osascript(script: string, timeoutMs = 10_000): string {
  return execFileSync('osascript', ['-e', script], { encoding: 'utf8', timeout: timeoutMs }).trim()
}
function tryRun(file: string, args: string[]): string | null {
  try { return execFileSync(file, args, { encoding: 'utf8', timeout: 10_000 }).trim() } catch { return null }
}
function tryQuiet(file: string, args: string[]): string | null {
  try { return execFileSync(file, args, { encoding: 'utf8', timeout: 10_000, stdio: ['ignore', 'pipe', 'ignore'] }).trim() } catch { return null }
}
const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms))

/**
 * Reads the AppleScript-control preference without launching VoiceOver:
 * true/false from the legacy defaults domain, or null when it is not readable
 * (on macOS 26 VoiceOver keeps its settings in a TCC-protected group
 * container). Null is resolved after start: the first AppleScript call then
 * either answers or fails with the permission error, and VoiceOver is stopped.
 */
export function voiceOverAppleScriptEnabled(): boolean | null {
  const value = tryQuiet('defaults', ['read', 'com.apple.VoiceOver4/default', 'SCREnableAppleScript'])
  if (value === null) return null
  return value === '1' || value === 'true'
}
function voiceOverRunning(): boolean {
  return tryRun('pgrep', ['-x', 'VoiceOver']) !== null
}
function voiceOverRunningPerSystemEvents(): boolean | string {
  try {
    return osascript('tell application "System Events" to (name of processes) contains "VoiceOver"') === 'true'
  } catch (error) {
    return `System Events query failed: ${error instanceof Error ? error.message.split('\n')[0] : String(error)}`
  }
}

async function startVoiceOver(): Promise<void> {
  execFileSync('open', ['/System/Library/CoreServices/VoiceOver.app'])
  const started = Date.now()
  for (;;) {
    let lastError = ''
    if (voiceOverRunning()) {
      try { osascript('tell application "VoiceOver" to return content of last phrase'); return } catch (error) { lastError = error instanceof Error ? error.message : String(error) }
    }
    // -1743/-1728 or "not allowed": AppleScript control is off (the finally block stops VoiceOver).
    if (/not allowed|-1743|AppleScript/i.test(lastError) && Date.now() - started > 3000) throw new Error(`${APPLESCRIPT_HELP}\n(AppleScript said: ${lastError.trim()})`)
    if (Date.now() - started > 20_000) throw new Error(`VoiceOver did not start (or did not answer AppleScript) within 20 s${lastError ? `: ${lastError.trim()}` : ''}`)
    await sleep(300)
  }
}

/** Always leaves VoiceOver off: AppleScript quit, then the process is checked and stopped. */
async function stopVoiceOver(): Promise<Record<string, unknown>> {
  const record: Record<string, unknown> = {}
  try { osascript('tell application "VoiceOver" to quit') ; record.quit = 'applescript quit sent' } catch (error) { record.quit = `applescript quit failed: ${error instanceof Error ? error.message.split('\n')[0] : String(error)}` }
  for (let i = 0; i < 20 && voiceOverRunning(); i++) await sleep(250)
  if (voiceOverRunning()) {
    tryRun('pkill', ['-x', 'VoiceOver'])
    record.pkill = true
    for (let i = 0; i < 20 && voiceOverRunning(); i++) await sleep(250)
  }
  record.runningAfter = voiceOverRunning()
  record.systemEventsSaysRunning = voiceOverRunningPerSystemEvents()
  return record
}

/** Every distinct `content of last phrase` seen within `windowMs`. */
async function phrasesFor(windowMs: number): Promise<string[]> {
  const seen: string[] = []
  const end = Date.now() + windowMs
  while (Date.now() < end) {
    try {
      const phrase = osascript('tell application "VoiceOver" to return content of last phrase', 3000)
      if (phrase && seen[seen.length - 1] !== phrase) seen.push(phrase)
    } catch { /* VoiceOver busy */ }
    await sleep(200)
  }
  return seen
}

interface Expectation { name: string; pattern: RegExp }
interface StepResult { name: string; pass: boolean; phrases: string[]; expectations: Array<{ name: string; pattern: string; matched: boolean }>; error?: string }

async function main(): Promise<number> {
  mkdirSync(OUT, { recursive: true })
  const receipt: Record<string, unknown> = {
    label: LABEL,
    startedAt: new Date().toISOString(),
    machine: { macOS: tryRun('sw_vers', ['-productVersion']), build: tryRun('sw_vers', ['-buildVersion']) },
    safari: tryRun('plutil', ['-extract', 'CFBundleShortVersionString', 'raw', '/Applications/Safari.app/Contents/Info.plist']),
    commit: tryRun('git', ['-C', ROOT, 'rev-parse', 'HEAD']),
  }
  const writeReceipt = (extra: Record<string, unknown>) => {
    const file = join(OUT, 'receipt.json')
    writeFileSync(file, `${JSON.stringify({ ...receipt, ...extra, finishedAt: new Date().toISOString() }, null, 2)}\n`)
    console.log(`receipt: ${file}`)
  }

  // ---- preflight: nothing below starts VoiceOver until every check passes ----
  const problems: string[] = []
  if (process.platform !== 'darwin') problems.push('VoiceOver needs macOS.')
  if (macScreenLocked()) problems.push(SCREEN_LOCKED_HELP)
  const appleScript = voiceOverAppleScriptEnabled()
  if (appleScript === false) problems.push(APPLESCRIPT_HELP)
  if (voiceOverRunning()) problems.push('VoiceOver is already running; turn it off (Cmd+F5) so this check owns its start and stop.')
  receipt.preflight = { problems, appleScriptEnabled: appleScript ?? 'unreadable (TCC-protected); confirmed by the first AppleScript call after start', screenLocked: macScreenLocked() }
  if (problems.length) {
    console.error(problems.join('\n\n'))
    writeReceipt({ ran: false })
    return 2
  }
  if (PREFLIGHT_ONLY) {
    console.log(`preflight ok (VoiceOver not started): screen unlocked; AppleScript control ${appleScript === true ? 'allowed' : 'setting not readable here, confirmed when the check starts VoiceOver'}.`)
    writeReceipt({ ran: false, preflightOnly: true })
    return 0
  }

  const driver = new SafariDriver(DRIVER_PORT)
  let harness: ServerHandle | null = null
  let fixtures: FixtureServer | null = null
  let voiceOverStarted = false
  const steps: StepResult[] = []
  try {
    try {
      await driver.start()
      await driver.newSession()
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      console.error(`Could not open a Safari WebDriver session: ${message}`)
      if (looksLikeAutomationDisabled(message)) console.error(`\n${SAFARI_ENABLE_HELP}`)
      writeReceipt({ ran: false, setupError: message })
      return 2
    }
    await ensureServedFixtures(['f2-1k.xlsx'])
    harness = await ensureHarnessServer({ root: ROOT, port: HARNESS_PORT, host: '127.0.0.1', identity: gitIdentity(ROOT), allowReuse: true }, 60_000)
    fixtures = await startFixtureServer(ROOT, FIXTURE_PORT)
    await driver.navigate(`http://127.0.0.1:${HARNESS_PORT}/src/debug/viewer-e2e.html?${new URLSearchParams({ src: fixtures.url('f2-1k.xlsx') })}`)
    await driver.waitFor(`const s = [...document.querySelectorAll('[role=status]')].map(e => e.textContent).find(t => t && t.includes('complete body')); return document.querySelector('[role=grid]') && s`, 120_000, 'f2-1k to open')
    const goTo = async (address: string) => {
      const input = await driver.execute<ElementRef>(`return document.querySelector('input[aria-label="Go To cell"]')`)
      await driver.clear(input)
      await driver.type(input, address)
      await driver.press(Keys.Enter)
      await driver.waitFor(`const g = document.querySelector('[role=grid]'), id = g.getAttribute('aria-activedescendant'), c = id && document.getElementById(id); return c && c.getAttribute('aria-label').split(',')[0] === arguments[0]`, 30_000, `active ${address}`, address)
    }
    await goTo('B2')

    await startVoiceOver()
    voiceOverStarted = true
    receipt.voiceOverStartedAt = new Date().toISOString()
    await sleep(2500) // VoiceOver's own start-up speech

    const step = async (name: string, act: () => Promise<unknown>, expectations: Expectation[], windowMs = 2500) => {
      const result: StepResult = { name, pass: false, phrases: [], expectations: [] }
      try {
        await act()
        result.phrases = await phrasesFor(windowMs)
        const text = result.phrases.join(' | ')
        result.expectations = expectations.map(e => ({ name: e.name, pattern: String(e.pattern), matched: e.pattern.test(text) }))
        result.pass = result.expectations.every(e => e.matched)
      } catch (error) {
        result.error = error instanceof Error ? error.message : String(error)
      }
      steps.push(result)
      console.log(`${result.pass ? 'PASS' : 'FAIL'} ${name}: ${JSON.stringify(result.phrases)}`)
    }
    const press = (key: string, modifiers: string[] = []) => () => driver.press(key, modifiers)

    await step('focus the grid', () => driver.execute(`document.querySelector('[role=grid]').focus(); return true`), [
      { name: 'grid name "Worksheet F2"', pattern: /Worksheet F2/ },
      { name: 'chrome-inclusive row count 1048577', pattern: /1,?048,?577/ },
      { name: 'chrome-inclusive column count 16385', pattern: /16,?385/ },
    ], 4000)
    await step('ArrowDown announces the new active descendant', press(Keys.ArrowDown), [{ name: 'cell name "B3, value …"', pattern: /B3, value/ }])
    await step('ArrowRight announces the next cell', press(Keys.ArrowRight), [{ name: 'cell name "C3, value …"', pattern: /C3, value/ }])
    await step('Shift+ArrowDown extends and reports selection', press(Keys.ArrowDown, [Keys.Shift]), [
      { name: 'extension target "C4"', pattern: /C4/ },
      { name: 'selected state', pattern: /selected/i },
    ])
    await step('Alt+ArrowDown opens the column filter popup', press(Keys.ArrowDown, [Keys.Alt]), [{ name: 'popup name "Column filter"', pattern: /Column filter/ }], 3000)
    await step('Escape closes the popup and returns to the grid', press(Keys.Escape), [{ name: 'grid or active cell re-announced', pattern: /Worksheet F2|C\d+, value/ }])
    await step('F2 enters header control mode', press(Keys.F2), [{ name: 'column select control "Select Column C"', pattern: /Select Column C/ }])
    await step('Tab reaches the sort control', press(Keys.Tab), [{ name: 'sort control "Cycle sort column C"', pattern: /Cycle sort column C/ }])
    await step('Enter applies a header sort', press(Keys.Enter), [{ name: 'sort state "ascending" / applied', pattern: /ascending|View-only sort applied/i }], 4000)
    const failed = steps.filter(s => !s.pass).length
    receipt.summary = { steps: steps.length, failed }
    return failed ? 1 : 0
  } catch (error) {
    receipt.error = error instanceof Error ? error.stack ?? error.message : String(error)
    console.error(receipt.error)
    return 2
  } finally {
    if (voiceOverStarted || voiceOverRunning()) receipt.voiceOverStop = await stopVoiceOver()
    await driver.quit()
    await fixtures?.close()
    harness?.stop()
    writeReceipt({ ran: voiceOverStarted, steps })
    if (voiceOverRunning()) console.error('WARNING: VoiceOver is still running. Press Cmd+F5 to turn it off.')
  }
}

process.exit(await main())
