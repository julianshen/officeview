/**
 * Start (or safely reuse) the harness vite server for the bench CLI.
 * Reuse requires the identity marker for THIS checkout and commit: any other
 * HTTP answer (a stale server, another checkout, a foreign app that returns
 * 200 for every path) is refused instead of silently measured.
 */
import { spawn } from 'node:child_process'
import { HARNESS_ID_PATH, checkIdentity, type HarnessIdentity } from './identity'

export interface ServerOptions {
  root: string
  port: number
  /** 127.0.0.1 for automation; 0.0.0.0 only for the manual LAN mode. */
  host: string
  identity: HarnessIdentity
  env?: Record<string, string>
  allowReuse: boolean
}

export interface ServerHandle {
  reused: boolean
  stop(): void
}

const OUTPUT_TAIL = 8192

async function probe(port: number): Promise<{ state: 'down' } | { state: 'up'; body: unknown }> {
  let response: Response
  try {
    response = await fetch(`http://127.0.0.1:${port}${HARNESS_ID_PATH}`)
  } catch {
    return { state: 'down' }
  }
  const text = await response.text()
  try {
    return { state: 'up', body: response.ok ? JSON.parse(text) : text }
  } catch {
    return { state: 'up', body: text }
  }
}

export async function ensureHarnessServer(options: ServerOptions, timeoutMs = 30_000): Promise<ServerHandle> {
  const existing = await probe(options.port)
  if (existing.state === 'up') {
    const problem = checkIdentity(existing.body, options.identity)
    if (problem) throw new Error(`port ${options.port} is already in use and cannot be reused: ${problem}`)
    if (!options.allowReuse) throw new Error(`port ${options.port} already serves a harness; stop it first (this mode needs its own server)`)
    return { reused: true, stop: () => undefined }
  }

  const child = spawn('bunx', ['vite', '--config', 'vite.harness.config.ts', '--port', String(options.port), '--strictPort', '--host', options.host], {
    cwd: options.root,
    env: { ...process.env, ...options.env },
    stdio: ['ignore', 'pipe', 'pipe'],
    detached: true,
  })
  let output = ''
  const capture = (chunk: Buffer): void => {
    output = (output + chunk.toString()).slice(-OUTPUT_TAIL)
  }
  child.stdout?.on('data', capture)
  child.stderr?.on('data', capture)
  let exited = false
  child.on('exit', () => {
    exited = true
  })
  const stop = (): void => {
    if (exited || child.pid === undefined) return
    try {
      process.kill(-child.pid, 'SIGTERM')
    } catch {
      child.kill('SIGTERM')
    }
  }
  const startedAt = Date.now()
  for (;;) {
    const state = await probe(options.port)
    if (state.state === 'up') {
      const problem = checkIdentity(state.body, options.identity)
      if (!problem) return { reused: false, stop }
      stop()
      throw new Error(`vite started but the harness identity check failed: ${problem}\n${output}`)
    }
    if (exited) throw new Error(`vite exited before serving the harness:\n${output}`)
    if (Date.now() - startedAt > timeoutMs) {
      stop()
      throw new Error(`timed out waiting for the harness server on port ${options.port}:\n${output}`)
    }
    await new Promise((resolve) => setTimeout(resolve, 250))
  }
}
