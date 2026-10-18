/**
 * S13 actual-Safari check: the core S10 viewer flows in the Safari build
 * installed on this Mac, driven through `safaridriver` (W3C WebDriver over
 * HTTP). This is real Safari, not Playwright WebKit.
 *
 *   bun scripts/harness/safari-check.ts [--driver-port 4451] [--harness-port 4191]
 *                                       [--fixture-port 4193] [--out /tmp/officeview-safari-check]
 *
 * It starts (or reuses, after an identity check) the vite harness server that
 * the Playwright suite uses, a loopback fixture server, and safaridriver; then
 * opens src/debug/viewer-e2e.html on F1 (resident), F1/F2/F3 worker-backed
 * paged packages and the F2 storage package. Each step records pass/fail, a
 * detail object and a screenshot; the JSON receipt lands in --out.
 *
 * Copy is asserted on the payload captured by the e2e page's clipboard stub
 * (the page replaces navigator.clipboard.writeText); the system clipboard is
 * never read. Timings are observations under WebDriver, not S10.2 receipts.
 *
 * Exit codes: 0 every step passed, 1 a step failed, 2 setup failed (for
 * example remote automation not enabled or the screen locked — the message
 * says what to do).
 */
import { execFileSync } from 'node:child_process'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createFixture } from '../../tests/helpers/xlsx-viewer-fixtures'
import { ensureServedFixtures, startFixtureServer, type FixtureServer, type ServedFixture } from './fixture-server'
import { installTeardown } from './cleanup'
import { gitIdentity } from './identity'
import { ensureHarnessServer, type ServerHandle } from './vite-server'
import { Keys, SAFARI_ENABLE_HELP, SCREEN_LOCKED_HELP, SafariDriver, looksLikeAutomationDisabled, macScreenLocked, type ElementRef } from './webdriver'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))
const LABEL = 'actual Safari via safaridriver (W3C WebDriver) on this Mac'

function option(name: string, fallback: string): string {
  const index = process.argv.indexOf(`--${name}`)
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback
}
const DRIVER_PORT = Number(option('driver-port', '4451'))
const HARNESS_PORT = Number(option('harness-port', '4191'))
const FIXTURE_PORT = Number(option('fixture-port', '4193'))
const OUT = option('out', '/tmp/officeview-safari-check')

function plistVersion(app: string): string | null {
  try {
    return execFileSync('plutil', ['-extract', 'CFBundleShortVersionString', 'raw', `${app}/Contents/Info.plist`], { encoding: 'utf8' }).trim()
  } catch {
    return null
  }
}
function command(file: string, args: string[]): string | null {
  try { return execFileSync(file, args, { encoding: 'utf8' }).trim() } catch { return null }
}

/** In-page helpers prepended to every script (WebDriver runs a function body). */
const H = `
const byLabel = (sel, name, prefix) => [...document.querySelectorAll(sel)].find(e => {
  const l = (e.getAttribute('aria-label') ?? e.textContent ?? '').trim()
  return prefix ? l.startsWith(name) : l === name
}) ?? null
const active = () => {
  const g = document.querySelector('[role=grid]'), id = g && g.getAttribute('aria-activedescendant')
  const c = id ? document.getElementById(id) : null
  return c ? { id: c.id, label: c.textContent ?? '' } : null
}
const statusText = () => [...document.querySelectorAll('[role=status]')].map(e => e.textContent ?? '').find(t => t.includes('Visible body')) ?? null
const e2e = () => window.__viewerE2E
`

interface StepResult {
  name: string
  pass: boolean
  ms: number
  detail: Record<string, unknown>
  error?: string
  screenshot?: string
}

class Check {
  readonly steps: StepResult[] = []
  constructor(private readonly driver: SafariDriver, private readonly fixtures: FixtureServer) {}

  run = async (name: string, body: (detail: Record<string, unknown>) => Promise<void>): Promise<void> => {
    const detail: Record<string, unknown> = {}
    const started = Date.now()
    const step: StepResult = { name, pass: false, ms: 0, detail }
    try {
      await body(detail)
      step.pass = true
    } catch (error) {
      step.error = error instanceof Error ? error.message : String(error)
    }
    step.ms = Date.now() - started
    try {
      const file = join(OUT, `${String(this.steps.length + 1).padStart(2, '0')}-${name.replace(/[^a-z0-9]+/gi, '-').toLowerCase()}.png`)
      writeFileSync(file, await this.driver.screenshot())
      step.screenshot = file
    } catch (error) {
      detail.screenshotError = error instanceof Error ? error.message : String(error)
    }
    this.steps.push(step)
    console.log(`${step.pass ? 'PASS' : 'FAIL'} ${name} (${step.ms} ms)${step.error ? ` — ${step.error}` : ''}`)
  }

  page = <T>(body: string, ...args: unknown[]) => this.driver.execute<T>(`${H}\n${body}`, ...args)
  until = <T>(body: string, timeoutMs: number, describe: string, ...args: unknown[]) => this.driver.waitFor<T>(`${H}\n${body}`, timeoutMs, describe, ...args)

  /** Opens a fixture; `settle` (default) also waits for the import to finish
   * ("complete body", no "(loading)") so later steps never race the import. */
  async open(name: ServedFixture, worker: boolean, settle = true): Promise<string> {
    const query = new URLSearchParams({ src: this.fixtures.url(name) })
    if (worker) query.set('worker', '1')
    await this.driver.navigate(`http://127.0.0.1:${HARNESS_PORT}/src/debug/viewer-e2e.html?${query}`)
    return this.until<string>(`
      const alert = document.querySelector('[data-testid=e2e-error]')
      if (alert) return 'ERROR: ' + alert.textContent
      if (e2e() && e2e().error) return 'ERROR: ' + e2e().error
      const s = statusText()
      if (!document.querySelector('[role=grid]') || !s) return null
      return !arguments[0] || (s.includes('complete body') && !s.includes('(loading)')) ? s : null`, 180_000, `${name} to open`, settle).then(text => {
      if (text.startsWith('ERROR: ')) throw new Error(text)
      return text
    })
  }

  async element(body: string, describe: string, ...args: unknown[]): Promise<ElementRef> {
    const found = await this.page<ElementRef | null>(body, ...args)
    if (!found) throw new Error(`element not found: ${describe}`)
    return found
  }

  /** A real WebDriver click on an element that exists and is enabled (popup
   * controls stay disabled while their value list loads, and a click on a
   * disabled button is a silent no-op); a synthetic el.click() fallback is
   * recorded, never hidden. */
  async click(body: string, describe: string, detail: Record<string, unknown>, ...args: unknown[]): Promise<void> {
    await this.until(`const el = (() => { ${body} })(); return Boolean(el && !el.disabled)`, 60_000, `${describe} to be enabled`, ...args)
    const el = await this.element(body, describe, ...args)
    try {
      await this.driver.click(el)
    } catch (error) {
      const fallbacks = (detail.syntheticClickFallbacks as string[] | undefined) ?? []
      fallbacks.push(`${describe}: ${error instanceof Error ? error.message : String(error)}`)
      detail.syntheticClickFallbacks = fallbacks
      await this.driver.execute('arguments[0].click()', el)
    }
  }

  async goTo(address: string, detail: Record<string, unknown>): Promise<string> {
    const input = await this.element(`return document.querySelector('input[aria-label="Go To cell"]')`, 'Go To cell')
    await this.driver.clear(input)
    await this.driver.type(input, address)
    await this.click(`return byLabel('button', 'Go To')`, 'Go To button', detail)
    return this.activeName(address)
  }

  /**
   * Waits (bounded) until the active cell's name is no longer pending. A
   * worker-backed cell legitimately reads "…, value Loading…" until its page
   * arrives; the labels seen on the way are returned so the caller can check
   * the transition was pending → value and never a blank value.
   */
  async settledName(address: string, timeoutMs = 30_000): Promise<{ label: string; sequence: string[] }> {
    const sequence: string[] = []
    const started = Date.now()
    for (;;) {
      const label = await this.page<string | null>(`const a = active(); return a && a.label.split(',')[0] === arguments[0] ? a.label : null`, address)
      if (label !== null && sequence[sequence.length - 1] !== label) sequence.push(label)
      if (label !== null && !/value Loading…/.test(label)) return { label, sequence }
      if (Date.now() - started > timeoutMs) throw new Error(`active cell ${address} still pending after ${timeoutMs} ms; seen ${JSON.stringify(sequence)}`)
      await new Promise(resolve => setTimeout(resolve, 150))
    }
  }

  activeName(address: string, timeoutMs = 30_000): Promise<string> {
    return this.until<string>(`const a = active(); return a && a.label.split(',')[0] === arguments[0] && a.label`, timeoutMs, `active cell ${address}`, address)
  }

  focusGrid(): Promise<unknown> { return this.page(`document.querySelector('[role=grid]').focus(); return document.activeElement === document.querySelector('[role=grid]')`) }

  async search(text: string): Promise<void> {
    const input = await this.element(`return document.querySelector('[aria-label="Search document"]')`, 'Search document')
    await this.driver.clear(input)
    await this.driver.type(input, text)
  }
}

async function main(): Promise<number> {
  mkdirSync(OUT, { recursive: true })
  const startedAt = new Date().toISOString()
  const safariVersion = plistVersion('/Applications/Safari.app')
  const receipt: Record<string, unknown> = {
    label: LABEL,
    startedAt,
    machine: { macOS: command('sw_vers', ['-productVersion']), build: command('sw_vers', ['-buildVersion']), arch: command('uname', ['-m']) },
    safari: { appVersion: safariVersion, safaridriver: command('/usr/bin/safaridriver', ['--version']) },
    commit: command('git', ['-C', ROOT, 'rev-parse', 'HEAD']),
    ports: { driver: DRIVER_PORT, harness: HARNESS_PORT, fixtures: FIXTURE_PORT },
  }
  const writeReceipt = (extra: Record<string, unknown>) => {
    const file = join(OUT, 'receipt.json')
    writeFileSync(file, `${JSON.stringify({ ...receipt, ...extra, finishedAt: new Date().toISOString() }, null, 2)}\n`)
    console.log(`receipt: ${file}`)
  }

  if (macScreenLocked()) {
    console.error(SCREEN_LOCKED_HELP)
    writeReceipt({ setupError: 'screen locked', steps: [] })
    return 2
  }
  const driver = new SafariDriver(DRIVER_PORT)
  let harness: ServerHandle | null = null
  let fixtures: FixtureServer | null = null
  // One teardown for the finally block and for Ctrl+C / SIGTERM / SIGHUP.
  const teardown = installTeardown(async () => {
    await driver.quit()
    await fixtures?.close()
    harness?.stop()
  })
  try {
    try {
      await driver.start()
      receipt.session = await driver.newSession()
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      console.error(`Could not open a Safari WebDriver session: ${message}`)
      if (looksLikeAutomationDisabled(message)) console.error(`\n${SAFARI_ENABLE_HELP}`)
      writeReceipt({ setupError: message, automationDisabled: looksLikeAutomationDisabled(message), steps: [] })
      return 2
    }
    const names: ServedFixture[] = ['f1-small.xlsx', 'f1-paged.xlsx', 'f2-paged.xlsx', 'f3-paged.xlsx', 'f2-storage.xlsx']
    await ensureServedFixtures(names)
    harness = await ensureHarnessServer({ root: ROOT, port: HARNESS_PORT, host: '127.0.0.1', identity: gitIdentity(ROOT), allowReuse: true }, 60_000)
    fixtures = await startFixtureServer(ROOT, FIXTURE_PORT)
    const check = new Check(driver, fixtures)
    const { run } = check

    await run('capability detection (main thread and dedicated worker)', async detail => {
      await check.open('f1-small.xlsx', false)
      const caps = await driver.executeAsync<Record<string, unknown>>(`
        const done = arguments[arguments.length - 1]
        const out = {}
        const step = async (name, fn) => { try { out[name] = await fn() } catch (e) { out[name] = 'error: ' + (e && e.name) + ': ' + (e && e.message) } }
        ;(async () => {
          out.userAgent = navigator.userAgent
          out.worker = typeof Worker === 'function'
          out.offscreenCanvas = typeof OffscreenCanvas === 'function'
          out.offscreen2d = await (async () => { try { return Boolean(new OffscreenCanvas(4, 4).getContext('2d')) } catch (e) { return 'error: ' + e.message } })()
          out.indexedDB = typeof indexedDB === 'object' && indexedDB !== null
          await step('indexedDBOpen', () => new Promise((res, rej) => { const r = indexedDB.open('officeview-safari-check-probe', 1); r.onsuccess = () => { r.result.close(); indexedDB.deleteDatabase('officeview-safari-check-probe'); res(true) }; r.onerror = () => rej(r.error) }))
          out.opfsGetDirectory = Boolean(navigator.storage && navigator.storage.getDirectory)
          await step('storageEstimate', async () => { const e = await navigator.storage.estimate(); return { quota: e.quota, usage: e.usage } })
          await step('opfsMainThreadCreateWritable', async () => {
            const dir = await navigator.storage.getDirectory()
            const handle = await dir.getFileHandle('officeview-safari-check-main.bin', { create: true })
            if (typeof handle.createWritable !== 'function') return 'createWritable is not a function'
            const w = await handle.createWritable(); await w.write(new Uint8Array([1, 2, 3])); await w.close()
            const size = (await handle.getFile()).size
            await dir.removeEntry('officeview-safari-check-main.bin')
            return 'wrote ' + size + ' bytes'
          })
          const src = "onmessage = async () => { const out = { offscreenCanvas: typeof OffscreenCanvas === 'function', indexedDB: typeof indexedDB === 'object' && indexedDB !== null }; try { const dir = await navigator.storage.getDirectory(); const h = await dir.getFileHandle('officeview-safari-check-worker.bin', { create: true }); out.createSyncAccessHandle = typeof h.createSyncAccessHandle === 'function'; out.createWritable = typeof h.createWritable === 'function'; if (out.createSyncAccessHandle) { const a = await h.createSyncAccessHandle(); const n = a.write(new Uint8Array(1 << 20), { at: 0 }); a.flush(); out.syncWrite = n + ' bytes, size ' + a.getSize(); a.close() } await dir.removeEntry('officeview-safari-check-worker.bin') } catch (e) { out.opfsError = e.name + ': ' + e.message } postMessage(out) }"
          await step('dedicatedWorker', () => new Promise((res, rej) => {
            const w = new Worker(URL.createObjectURL(new Blob([src], { type: 'text/javascript' })))
            const t = setTimeout(() => { w.terminate(); rej(new Error('worker probe timed out')) }, 15000)
            w.onmessage = e => { clearTimeout(t); w.terminate(); res(e.data) }
            w.onerror = e => { clearTimeout(t); rej(new Error(e.message)) }
            w.postMessage(0)
          }))
          done(out)
        })()`)
      Object.assign(detail, caps)
      if (caps.worker !== true) throw new Error('Worker is unavailable')
    })

    for (const fixture of [{ name: 'f1-small.xlsx', worker: false }, { name: 'f1-paged.xlsx', worker: true }] as const) {
      const kind = fixture.worker ? 'worker-backed paged' : 'resident'
      await run(`F1 ${kind}: open, Go To XFD1048576, find a never-painted corner, copy`, async detail => {
        detail.status = await check.open(fixture.name, fixture.worker)
        const workers = await check.page<{ created: number; open: number }>(`return { created: e2e().workersCreated, open: e2e().counts().open ?? 0 }`)
        detail.workers = workers
        if (fixture.worker ? workers.open !== 1 : workers.created !== 0) throw new Error(`unexpected worker use ${JSON.stringify(workers)}`)
        await check.goTo('XFD1048576', detail)
        // The worker read may still be in flight: "Loading…" is the correct
        // pending state, so wait for the value, then check the transition.
        const settled = await check.settledName('XFD1048576')
        const corner = settled.label
        detail.corner = corner
        detail.cornerSequence = settled.sequence
        const blank = settled.sequence.find(label => /, value\s*$/.test(label) || /, value ;/.test(label))
        if (blank) throw new Error(`corner showed a blank value on the way: ${blank}`)
        if (settled.sequence.some(label => !/value (Loading…|17179869183)/.test(label))) throw new Error(`unexpected corner labels ${JSON.stringify(settled.sequence)}`)
        if (!/^XFD1048576, value 17179869183/.test(corner)) throw new Error(`corner label ${corner}`)
        const grid = await check.page<Record<string, string | null>>(`const g = document.querySelector('[role=grid]'); const a = active(); return { rowcount: g.getAttribute('aria-rowcount'), colcount: g.getAttribute('aria-colcount'), id: a && a.id, name: g.getAttribute('aria-label') }`)
        detail.grid = grid
        if (grid.rowcount !== '1048577' || grid.colcount !== '16385' || !grid.id?.endsWith('-cell-17179869183')) throw new Error(`grid ${JSON.stringify(grid)}`)
        await check.search('A1048576')
        detail.find = await check.until<string>(`const a = active(); return a && /^A1048576, value A1048576/.test(a.label) && a.id`, 60_000, 'find to reach A1048576')
        if (!String(detail.find).endsWith('-cell-17179852800')) throw new Error(`find landed on ${detail.find}`)
        await check.click(`return byLabel('button', 'Copy selection')`, 'Copy selection', detail)
        detail.copyPayload = await check.until<string>(`const c = e2e().clipboard; return c.length && c[c.length - 1]`, 30_000, 'copy payload')
        if (detail.copyPayload !== 'A1048576') throw new Error(`copy payload ${JSON.stringify(detail.copyPayload)}`)
        detail.copyNote = 'payload captured by the e2e page clipboard stub; system clipboard not read'
      })
    }

    await run('keyboard navigation and zoom (F1 resident)', async detail => {
      await check.open('f1-small.xlsx', false)
      await check.goTo('B2', detail)
      detail.focused = await check.focusGrid()
      const steps: Array<[string, string[], string]> = [
        [Keys.ArrowDown, [], 'B3'], [Keys.ArrowRight, [], 'C3'], [Keys.Enter, [], 'C4'], [Keys.Home, [], 'A4'],
        // Ctrl+Home, not Cmd+Home: Safari reserves Cmd+Home (it navigated this
        // tab to the Safari home page in the S13 run); the viewer accepts both.
        [Keys.Home, [Keys.Control], 'A1'], [Keys.ArrowDown, [Keys.Meta], 'A1048576'], [Keys.ArrowRight, [Keys.Meta], 'XFD1048576'], [Keys.ArrowLeft, [], 'XFC1048576'],
      ]
      const seen: string[] = []
      for (const [key, modifiers, want] of steps) {
        await driver.press(key, modifiers)
        seen.push(await check.activeName(want, 15_000))
      }
      detail.sequence = seen
      const zoom = () => check.page<string>(`return byLabel('button', 'Reset zoom').textContent`)
      detail.zoomBefore = await zoom()
      await check.focusGrid()
      await driver.press('+')
      detail.zoomAfterPlus = await check.until<string>(`const t = byLabel('button', 'Reset zoom').textContent; return t !== '100%' && t`, 10_000, 'zoom in')
      await driver.press('0')
      detail.zoomAfterReset = await check.until<string>(`const t = byLabel('button', 'Reset zoom').textContent; return t === '100%' && t`, 10_000, 'zoom reset')
      if (detail.zoomBefore !== '100%') throw new Error(`zoom started at ${detail.zoomBefore}`)
    })

    await run('F2 worker-backed paged (100,010 rows): header sort with acknowledgment timeline, value filter apply/cancel/clear', async detail => {
      const rows = 100_010
      const handle = createFixture('F2', { rowCount: rows })
      let alpha = 0
      for (let row = 1; row < rows; row++) if (handle.cellAt(row, 3)?.value === 'alpha') alpha++
      detail.status = await check.open('f2-paged.xlsx', true)
      // In-page timeline: the click time, then each status / aria-sort change.
      await check.page(`
        const t = window.__safariAck = { click: null, events: [] }
        const button = byLabel('button', 'Cycle sort column G')
        button.addEventListener('click', () => { if (t.click === null) t.click = performance.now() }, { capture: true })
        new MutationObserver(() => {
          const s = statusText(), h = byLabel('[role=columnheader]', 'Column G', true)
          const sort = h && h.getAttribute('aria-sort')
          const last = t.events[t.events.length - 1]
          if (!last || last.status !== s || last.sort !== sort) t.events.push({ at: performance.now(), status: s, sort })
        }).observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['aria-sort', 'aria-busy'] })
        return true`)
      await check.click(`return byLabel('button', 'Cycle sort column G')`, 'Cycle sort column G', detail)
      await check.until(`return (statusText() ?? '').includes('View-only sort applied')`, 120_000, 'ascending sort applied')
      await check.click(`return byLabel('button', 'Cycle sort column G')`, 'Cycle sort column G (2)', detail)
      await check.until(`const h = byLabel('[role=columnheader]', 'Column G', true); return h && h.getAttribute('aria-sort') === 'descending'`, 120_000, 'descending sort')
      const timeline = await check.page<{ click: number; events: Array<{ at: number; status: string | null; sort: string | null }> }>(`return window.__safariAck`)
      detail.acknowledgment = {
        firstChangeAfterClickMs: timeline.events.length ? Math.round(timeline.events[0].at - timeline.click) : null,
        events: timeline.events.slice(0, 8).map(e => ({ ms: Math.round(e.at - timeline.click), sort: e.sort, status: e.status?.slice(0, 120) })),
        note: 'in-page performance.now() from the click listener; observation only, not an S10.2 latency receipt',
      }
      const last = rows - 2
      const g2 = await check.goTo('G2', detail)
      detail.sortedG2 = g2
      if (g2 !== `G2, value ${(2 * last + 1).toFixed(2)}; original G${last + 1}`) throw new Error(`sorted G2 ${g2}`)

      await check.goTo('A1', detail)
      const dialog = `const d = document.querySelector('[role=dialog][aria-label="Column filter"]')`
      await check.click(`return byLabel('button', 'Filter column D', true)`, 'Filter column D', detail)
      await check.until(`${dialog}; return d && byLabel('button', 'Deselect All')`, 30_000, 'filter popup')
      await check.click(`return byLabel('button', 'Deselect All')`, 'Deselect All', detail)
      const checkbox = (value: string) => `return [...document.querySelectorAll('[role=dialog] input[type=checkbox]')].find(c => (c.getAttribute('aria-label') ?? c.closest('label')?.textContent ?? '').trim() === ${JSON.stringify(value)}) ?? null`
      await check.click(checkbox('alpha'), 'alpha checkbox', detail)
      await check.click(`return byLabel('[role=dialog] button', 'OK')`, 'OK', detail)
      const want = `Visible body ${alpha} / complete body ${rows - 1}`
      detail.filtered = await check.until<string>(`const s = statusText(); return s && s.includes(arguments[0]) && s`, 120_000, want, want)
      const filterCount = () => check.page<number>(`return e2e().counts().filter ?? 0`)
      const afterApply = await filterCount()
      await check.click(`return byLabel('button', 'Filter column D', true)`, 'Filter column D (cancel)', detail)
      await check.until(`${dialog}; return d && byLabel('[role=dialog] button', 'Cancel')`, 30_000, 'filter popup (cancel)')
      await check.click(checkbox('beta'), 'beta checkbox', detail)
      await check.click(`return byLabel('[role=dialog] button', 'Cancel')`, 'Cancel', detail)
      await check.until(`${dialog}; return !d`, 10_000, 'popup closed')
      detail.afterCancel = await check.page<string>(`return statusText()`)
      if (!String(detail.afterCancel).includes(want) || await filterCount() !== afterApply) throw new Error('Cancel changed the view')
      await check.click(`return byLabel('button', 'Filter column D', true)`, 'Filter column D (clear)', detail)
      await check.until(`${dialog}; return d && byLabel('[role=dialog] button', 'Clear all filters')`, 30_000, 'filter popup (clear)')
      await check.click(`return byLabel('[role=dialog] button', 'Clear all filters')`, 'Clear all filters', detail)
      const all = `Visible body ${rows - 1} / complete body ${rows - 1}`
      detail.cleared = await check.until<string>(`const s = statusText(); return s && s.includes(arguments[0]) && s`, 120_000, all, all)
      detail.workerCommands = await check.page(`return e2e().counts()`)
    })

    await run('F3 worker-backed dense 8x16384: horizontal recenter and edge navigation', async detail => {
      detail.status = await check.open('f3-paged.xlsx', true)
      const visited: Record<string, unknown> = {}
      for (const address of ['XFD1', 'A1', 'XFD8', 'H3']) {
        await check.goTo(address, detail)
        visited[address] = await check.page(`
          const headers = new Set([...document.querySelectorAll('[role=columnheader]')].map(h => h.getAttribute('aria-colindex')))
          const cells = [...document.querySelectorAll('[role=gridcell]')]
          return { cells: cells.length, orphans: cells.filter(c => !headers.has(c.getAttribute('aria-colindex'))).length, header: Boolean(byLabel('[role=columnheader]', 'Column ' + arguments[0].replace(/\\d+$/, ''), true)) }`, address)
        const v = visited[address] as { orphans: number; header: boolean; cells: number }
        if (v.orphans !== 0 || !v.header || v.cells >= 2000) throw new Error(`${address}: ${JSON.stringify(v)}`)
      }
      detail.visited = visited
      await check.goTo('XFD8', detail)
      await check.focusGrid()
      await driver.press(Keys.ArrowLeft)
      detail.left = await check.activeName('XFC8')
      if (!/^XFC8, value \S/.test(String(detail.left))) throw new Error(`XFC8 not loaded: ${detail.left}`)
    })

    await run('F2 storage package (160,000 rows, needs OPFS/IndexedDB): WebKit OPFS-write outcome', async detail => {
      detail.openStatus = await check.open('f2-storage.xlsx', true, false)
      // Wait for a terminal state: complete, a stated incomplete/refusal, or an alert.
      detail.finalStatus = await check.until<string>(`
        const alert = document.querySelector('[role=alert]')
        if (alert) return 'ALERT: ' + alert.textContent
        const s = statusText() ?? ''
        if (s.includes('(incomplete)') || s.includes('could not be fully loaded')) return s
        if (/complete body 160000/.test(s) && !s.includes('(loading)')) return s
        const refusal = [...document.querySelectorAll('[role=status]')].map(e => e.textContent).find(t => t && /needs browser storage|bytes free/.test(t))
        return refusal ?? null`, 240_000, 'a terminal storage outcome').catch(async error => {
        detail.statusAtTimeout = await check.page<string | null>(`return statusText()`)
        throw error
      })
      const text = String(detail.finalStatus)
      detail.outcome = /complete body 160000/.test(text) && !text.includes('incomplete') ? 'complete: pages persisted (no write restriction hit)'
        : text.includes('could not be fully loaded') ? 'documented restriction: storage write failed and is stated (incomplete), never an empty-sheet lie'
          : `other: ${text.slice(0, 160)}`
      detail.workerCommands = await check.page(`return e2e().counts()`)
      // Pass = a truthful state: complete, or visibly diagnosed. A silent
      // "Visible body 0 / complete body 0" would fail.
      if (/Visible body 0 \/ complete body 0/.test(text) && !text.includes('could not be fully loaded')) throw new Error(`silent empty sheet: ${text}`)
    })

    const failed = check.steps.filter(step => !step.pass).length
    writeReceipt({ browserVersion: driver.capabilities.browserVersion, platformName: driver.capabilities.platformName, steps: check.steps, summary: { steps: check.steps.length, failed } })
    return failed ? 1 : 0
  } catch (error) {
    const message = error instanceof Error ? error.stack ?? error.message : String(error)
    console.error(message)
    writeReceipt({ setupError: message, steps: [] })
    return 2
  } finally {
    await teardown.run()
    teardown.dispose()
  }
}

process.exit(await main())
