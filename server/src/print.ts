/**
 * Stage 3b print (§4.2, rulings P1–P7): the whole-job plan and the per-sid
 * driver. A job is checked whole before any frame; its pages render at most
 * PRINT_CONCURRENCY at a time and go out strictly in ascending unit order.
 */
import { encodeRasterBody } from '../../src/stream/raster'
import { unitSurfaceAllowed, type UnitSize } from './render'

export const PRINT_RATE_PER_MIN = 10
export const PRINT_CONCURRENCY = 2
/** §14 PRINT_RASTER (kind 8) encoded cap; its decoded cap is the UNIT_RASTER one (64 MiB). */
export const PRINT_ENCODED_CAP = 8 * 1024 * 1024
/** The whole job's decoded budget (Σ w·h·4 over its pages). */
export const PRINT_JOB_MAX_DECODED = 1024 * 1024 * 1024

/** P1: scale 2, else 1. Every page of `from..to` must pass `unitSurfaceAllowed`
 * at the scale, and the job's Σ w·h·4 must stay within PRINT_JOB_MAX_DECODED. */
export function planPrint(units: readonly UnitSize[], from: number, to: number): { scale: 1 | 2 } | { refused: 'too-large' } {
  for (const scale of [2, 1] as const) {
    let total = 0, fits = true
    for (let i = from; i <= to && fits; i++) {
      const u = units[i]!, w = Math.ceil(u.w * scale), h = Math.ceil(u.h * scale)
      if (!unitSurfaceAllowed(w, h)) fits = false
      else total += w * h * 4
    }
    if (fits && total <= PRINT_JOB_MAX_DECODED) return { scale }
  }
  return { refused: 'too-large' }
}

/** The 0-based inclusive range of a `print` request, or undefined (malformed). */
export function printArgs(msg: Record<string, unknown>, unitCount: number): { from: number; to: number } | undefined {
  const { from, to } = msg
  if (!Number.isInteger(from) || !Number.isInteger(to)) return undefined
  const f = from as number, t = to as number
  return f >= 0 && f <= t && t < unitCount ? { from: f, to: t } : undefined
}

/** PRINT_RASTER body: `u32 LE unitIndex` + the UNIT_RASTER body (`u32 w, u32 h, u32 pngLen, PNG`). */
export function encodePrintBody(unitIndex: number, w: number, h: number, png: Uint8Array): Uint8Array {
  const raster = encodeRasterBody(w, h, png), out = new Uint8Array(4 + raster.length)
  new DataView(out.buffer).setUint32(0, unitIndex, true)
  out.set(raster, 4)
  return out
}

/** The sid's running job (P3: one at a time). `stopped`: cancel, teardown or a failed page. */
export interface PrintJob { readonly reqId: number; stopped: boolean }
/** Per-sid print state on the session entry: created at open, kept on resume, gone on teardown. */
export interface PrintSlot {
  /** Pages charged to this sid so far (a running sum: a page printed twice counts twice). */
  printedPages: number
  print?: PrintJob
}

/** One page's framed messages, a refusal, or undefined (the job or its socket is gone). */
export type PageResult = Uint8Array[][] | 'rate-limited' | 'policy-refused' | undefined
export interface PrintRun {
  job: PrintJob
  from: number
  to: number
  /** Renders, marks, frames and records one page. */
  page(unitIndex: number): Promise<PageResult>
  /** The issuing socket still owns the sid (in `stream`, same entry) and the job is not stopped. */
  alive(): boolean
  send(messages: Uint8Array[][]): void
  done(): void
  refuse(code: 'rate-limited' | 'policy-refused'): void
}

/** Drives one job: page j starts only while j < (next to send) + PRINT_CONCURRENCY,
 * so at most that many pages are in flight or buffered. The first failure sends one
 * refusal and stops the job; a stop sends nothing more. Returns once every started
 * page has settled, so the caller may free the sid's slot. */
export async function runPrint(r: PrintRun): Promise<void> {
  const n = r.to - r.from + 1
  const pending = new Map<number, Promise<PageResult>>()
  let started = 0
  try {
    for (let next = 0; next < n; next++) {
      if (!r.alive()) return
      while (started < n && started < next + PRINT_CONCURRENCY) {
        const i = started++
        pending.set(i, r.page(r.from + i).catch((): PageResult => 'policy-refused'))
      }
      const res = await pending.get(next)!
      pending.delete(next)
      if (!r.alive() || res === undefined) return
      if (typeof res === 'string') { r.job.stopped = true; return r.refuse(res) }
      r.send(res)
    }
    r.done()
  } finally {
    r.job.stopped = true
    await Promise.allSettled(pending.values())
  }
}
