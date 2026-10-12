/**
 * Minimal W3C WebDriver client for real Safari through `safaridriver`
 * (HTTP + JSON over fetch; no client library). Used by the S13 actual-Safari
 * check and the VoiceOver check. Nothing here emulates Safari: every call goes
 * to the Safari build installed on this Mac.
 */
import { execFileSync, spawn, type ChildProcess } from 'node:child_process'

const ELEMENT_KEY = 'element-6066-11e4-a52e-4f735466cecf'
export type ElementRef = { [ELEMENT_KEY]: string }

/** WebDriver key codes (W3C §17.4.2). */
export const Keys = {
  Enter: '', Escape: '', Tab: '', Home: '', End: '',
  ArrowLeft: '', ArrowUp: '', ArrowRight: '', ArrowDown: '',
  Shift: '', Control: '', Alt: '', Meta: '', F2: '',
} as const

/** Per-request ceiling for ordinary commands (page scripts carry their own shorter waits). */
export const COMMAND_TIMEOUT_MS = 120_000
/** Ceiling for the session delete during teardown. */
export const QUIT_TIMEOUT_MS = 3_000

export class WebDriverError extends Error {
  constructor(readonly status: number, readonly error: string, message: string) {
    super(`${error}: ${message}`)
  }
}

/** The exact steps the user must take once per Mac to allow automation. */
export const SAFARI_ENABLE_HELP = [
  'Safari remote automation is not enabled. Run these once, then rerun this check:',
  '  1. sudo safaridriver --enable            (asks for an administrator password)',
  '  2. Safari > Settings > Advanced: turn on "Show features for web developers",',
  '     then Safari > Settings > Developer (or the Develop menu): turn on "Allow Remote Automation".',
].join('\n')

export function looksLikeAutomationDisabled(message: string): boolean {
  return /remote automation|allow remote automation|safaridriver --enable|not enabled|session could not be created/i.test(message)
}

async function wait(ms: number): Promise<void> {
  await new Promise(resolve => setTimeout(resolve, ms))
}

export class SafariDriver {
  private child: ChildProcess | null = null
  private output = ''
  sessionId: string | null = null
  capabilities: Record<string, unknown> = {}

  constructor(readonly port: number, private readonly binary = '/usr/bin/safaridriver') {}

  get base(): string { return `http://127.0.0.1:${this.port}` }

  /** Starts `safaridriver -p <port>` and waits until /status answers. */
  async start(timeoutMs = 15_000): Promise<void> {
    if (process.platform !== 'darwin') throw new Error(`safaridriver needs macOS (this is ${process.platform})`)
    this.child = spawn(this.binary, ['-p', String(this.port)], { stdio: ['ignore', 'pipe', 'pipe'] })
    // A missing binary or a spawn failure arrives as an 'error' event; without
    // a listener it would crash the process instead of reaching the caller.
    let spawnError: Error | null = null
    this.child.once('error', error => { spawnError = error })
    const capture = (chunk: Buffer) => { this.output = (this.output + chunk.toString()).slice(-4096) }
    this.child.stdout?.on('data', capture)
    this.child.stderr?.on('data', capture)
    const started = Date.now()
    for (;;) {
      if (spawnError) { this.child = null; throw new Error(`safaridriver failed to start (${this.binary}): ${(spawnError as Error).message}`) }
      if (this.child.exitCode !== null) throw new Error(`safaridriver exited (${this.child.exitCode}): ${this.output.trim()}`)
      try {
        const response = await fetch(`${this.base}/status`)
        if (response.ok) return
      } catch {
        // not listening yet
      }
      if (Date.now() - started > timeoutMs) throw new Error(`safaridriver did not answer on port ${this.port}: ${this.output.trim()}`)
      await wait(200)
    }
  }

  /** One WebDriver call; `timeoutMs` aborts a request that safaridriver never answers. */
  async command<T = unknown>(method: 'GET' | 'POST' | 'DELETE', path: string, body?: unknown, timeoutMs = COMMAND_TIMEOUT_MS): Promise<T> {
    const response = await fetch(`${this.base}${path}`, {
      method,
      headers: body === undefined ? undefined : { 'content-type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body),
      signal: AbortSignal.timeout(timeoutMs),
    })
    const text = await response.text()
    let json: { value?: unknown } = {}
    try { json = JSON.parse(text) as { value?: unknown } } catch { /* non-JSON error body */ }
    const value = json.value as { error?: string; message?: string } | undefined
    if (!response.ok || (value && typeof value === 'object' && 'error' in value && value.error)) {
      throw new WebDriverError(response.status, value?.error ?? `HTTP ${response.status}`, value?.message ?? text.slice(0, 500))
    }
    return json.value as T
  }

  private s(path: string): string {
    if (!this.sessionId) throw new Error('no WebDriver session')
    return `/session/${this.sessionId}${path}`
  }

  async newSession(): Promise<Record<string, unknown>> {
    const value = await this.command<{ sessionId: string; capabilities: Record<string, unknown> }>('POST', '/session', {
      capabilities: { alwaysMatch: { browserName: 'safari' } },
    })
    this.sessionId = value.sessionId
    this.capabilities = value.capabilities
    return value.capabilities
  }

  async navigate(url: string): Promise<void> { await this.command('POST', this.s('/url'), { url }) }

  async execute<T = unknown>(script: string, ...args: unknown[]): Promise<T> {
    return this.command<T>('POST', this.s('/execute/sync'), { script, args })
  }

  async executeAsync<T = unknown>(script: string, ...args: unknown[]): Promise<T> {
    return this.command<T>('POST', this.s('/execute/async'), { script, args })
  }

  /** Polls a page predicate (a function body returning a truthy value). */
  async waitFor<T>(script: string, timeoutMs = 30_000, describe = script, ...args: unknown[]): Promise<T> {
    const started = Date.now()
    let last: unknown
    for (;;) {
      try {
        last = await this.execute(script, ...args)
        if (last) return last as T
      } catch (error) {
        last = error instanceof Error ? error.message : String(error)
      }
      if (Date.now() - started > timeoutMs) throw new Error(`timed out after ${timeoutMs} ms waiting for ${describe}; last=${JSON.stringify(last)}`)
      await wait(150)
    }
  }

  async click(element: ElementRef): Promise<void> {
    await this.command('POST', this.s(`/element/${element[ELEMENT_KEY]}/click`), {})
  }

  async clear(element: ElementRef): Promise<void> {
    await this.command('POST', this.s(`/element/${element[ELEMENT_KEY]}/clear`), {})
  }

  async type(element: ElementRef, text: string): Promise<void> {
    await this.command('POST', this.s(`/element/${element[ELEMENT_KEY]}/value`), { text })
  }

  /** Presses a chord through the Actions API: modifiers down, key down/up, modifiers up. */
  async press(key: string, modifiers: string[] = []): Promise<void> {
    const actions = [
      ...modifiers.map(value => ({ type: 'keyDown', value })),
      { type: 'keyDown', value: key }, { type: 'keyUp', value: key },
      ...modifiers.slice().reverse().map(value => ({ type: 'keyUp', value })),
    ]
    await this.command('POST', this.s('/actions'), { actions: [{ type: 'key', id: 'keyboard', actions }] })
    await this.command('DELETE', this.s('/actions'))
  }

  async screenshot(): Promise<Buffer> {
    return Buffer.from(await this.command<string>('GET', this.s('/screenshot')), 'base64')
  }

  /** Bounded: a wedged session delete is abandoned after QUIT_TIMEOUT_MS and the driver is killed anyway. */
  async quit(): Promise<void> {
    if (this.sessionId) {
      try { await this.command('DELETE', this.s(''), undefined, QUIT_TIMEOUT_MS) } catch { /* already gone, or wedged */ }
      this.sessionId = null
    }
    if (this.child && this.child.exitCode === null) this.child.kill('SIGTERM')
    this.child = null
  }
}

/**
 * True when the macOS login session's screen is locked. A locked screen makes
 * Safari present nothing (black WebDriver screenshots) and WebDriver clicks
 * report success without reaching the page, and VoiceOver would read the lock
 * screen, so both checks refuse to run rather than record meaningless steps.
 */
export function macScreenLocked(): boolean | null {
  try {
    const out = execFileSync('ioreg', ['-n', 'Root', '-d1', '-a'], { encoding: 'utf8' })
    const match = /<key>CGSSessionScreenIsLocked<\/key>\s*<(true|false)\/>/.exec(out)
    return match ? match[1] === 'true' : false
  } catch {
    return null
  }
}

export const SCREEN_LOCKED_HELP = 'The Mac screen is locked. Unlock it and keep the display awake (for example `caffeinate -d` in another terminal), then rerun.'
