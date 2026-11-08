/**
 * Stage-0 Task 1 — results writer.
 *
 * Every probe/plan task writes one JSON result to
 * `docs/validation/streaming-s0/<task>.json` with the Global-Constraints
 * envelope `{ gate, status: GREEN|RED|REPORTED, backend, gitSha, generatedAt, …data }`.
 * `REPORTED` = a measurement with no pass/fail target at Stage 0.
 */
import { execSync } from 'node:child_process'
import { mkdirSync, writeFileSync } from 'node:fs'
import path from 'node:path'

const RESULTS_DIR = path.join(process.cwd(), 'docs', 'validation', 'streaming-s0')

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

function currentBackend(): string {
  return process.env.S0_BACKEND === 'napi' ? '@napi-rs/canvas' : 'node-canvas'
}

/**
 * Resolve the backend under test: the napi vitest config and the napi runs set
 * `S0_BACKEND=napi`; everything else counts as the node-canvas backend.
 */
export function resolveBackend(env: string | undefined = process.env.S0_BACKEND): string {
  return env === 'napi' ? '@napi-rs/canvas' : 'node-canvas'
}

/** Write `docs/validation/streaming-s0/<task>.json` and return the envelope. */
export function writeResult(task: string, body: ResultBody): ResultBodyWithEnvelope {
  const out: ResultBodyWithEnvelope = {
    ...body,
    gate: task,
    backend: currentBackend(),
    gitSha: currentGitSha(),
    generatedAt: new Date().toISOString(),
  }
  mkdirSync(RESULTS_DIR, { recursive: true })
  writeFileSync(path.join(RESULTS_DIR, `${task}.json`), JSON.stringify(out, null, 2) + '\n')
  return out
}
