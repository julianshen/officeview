/**
 * Piece-5 Task 1 — results writer. One small JSON of numbers (no frames, no videos):
 * `docs/validation/streaming-s3-p5-feasibility.json` (or `--out <path>`).
 *
 * `deterministic` holds every measured outcome and must be identical across runs; `timing` holds
 * wall times, fps and machine load, which vary by run and are excluded from the determinism diff.
 */
import { execSync } from 'node:child_process'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'

export interface Envelope { gate: string; status: 'GREEN' | 'RED'; gitSha: string; generatedAt: string; deterministic: unknown; timing: unknown }

const r4 = (x: unknown): unknown =>
  typeof x === 'number' ? (Number.isInteger(x) ? x : Number(x.toPrecision(10)))
    : Array.isArray(x) ? x.map(r4)
      : x && typeof x === 'object' ? Object.fromEntries(Object.entries(x).map(([k, v]) => [k, r4(v)]))
        : x

export function writeResult(out: string, body: Omit<Envelope, 'gitSha' | 'generatedAt'>): Envelope {
  const env: Envelope = { ...body, deterministic: r4(body.deterministic), timing: r4(body.timing), gitSha: execSync('git rev-parse HEAD', { encoding: 'utf8' }).trim(), generatedAt: new Date().toISOString() }
  mkdirSync(path.dirname(out), { recursive: true })
  writeFileSync(out, JSON.stringify(env, null, 1) + '\n')
  return env
}

export function readResult(file: string): Envelope { return JSON.parse(readFileSync(file, 'utf8')) as Envelope }
