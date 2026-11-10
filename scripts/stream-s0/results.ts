/**
 * Stage-0 Task 1 — results writer.
 *
 * Every probe/plan task writes one JSON result to
 * `docs/validation/streaming-s0/<task>.json` with the Global-Constraints
 * envelope `{ gate, status: GREEN|RED|REPORTED, backend, gitSha, generatedAt, …data }`.
 * `REPORTED` = a measurement with no pass/fail target at Stage 0.
 */
import * as canvasModule from 'canvas'
import { execSync } from 'node:child_process'
import { mkdirSync, writeFileSync } from 'node:fs'
import path from 'node:path'

const RESULTS_DIR_DEFAULT = path.join(process.cwd(), 'docs', 'validation', 'streaming-s0')

export type RunStatus = 'GREEN' | 'RED' | 'REPORTED'

export interface ResultBody {
  status: RunStatus
  [key: string]: unknown
}

export interface ResultBodyWithEnvelope extends ResultBody {
  gate: string
  backend: string
  gitSha: string
  generatedAt: string
}

function currentGitSha(): string {
  return execSync('git rev-parse HEAD', { encoding: 'utf8', cwd: process.cwd() }).trim()
}

/**
 * The real backend the `canvas` specifier resolved to (round-10: module
 * identity, not an env var). Under the napi vitest alias / the bun preload
 * the specifier resolves to the shim, which exports `STREAM_S0_BACKEND`; the
 * real `canvas` package never does.
 */
export function resolveBackend(): string {
  return 'STREAM_S0_BACKEND' in canvasModule ? '@napi-rs/canvas' : 'node-canvas'
}

function resultsDir(): string {
  return process.env.S0_RESULTS_DIR && process.env.S0_RESULTS_DIR.length > 0
    ? process.env.S0_RESULTS_DIR
    : RESULTS_DIR_DEFAULT
}

const STATUSES: ReadonlyArray<RunStatus> = ['GREEN', 'RED', 'REPORTED']

/** Write `docs/validation/streaming-s0/<task>.json` and return the envelope. */
export function writeResult(task: string, body: ResultBody): ResultBodyWithEnvelope {
  if (!STATUSES.includes(body.status)) {
    throw new TypeError(`writeResult: status must be one of ${STATUSES.join(', ')} (got ${String(body.status)})`)
  }
  const out: ResultBodyWithEnvelope = {
    ...body,
    // The caller's `gate`, if any, is authoritative — the writer never
    // overrides a caller-supplied value (round-10 finding 7). Unset, the
    // file-name `task` is the gate.
    gate: 'gate' in body && typeof body.gate === 'string' && body.gate.length > 0 ? body.gate : task,
    backend: resolveBackend(),
    gitSha: currentGitSha(),
    generatedAt: new Date().toISOString(),
  }
  mkdirSync(resultsDir(), { recursive: true })
  writeFileSync(path.join(resultsDir(), `${out.gate}.json`), JSON.stringify(out, null, 2) + '\n')
  return out
}
