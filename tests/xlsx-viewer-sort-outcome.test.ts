import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { Page } from '@playwright/test'
import { measureSortOutcome, recordSortOutcome } from '../scripts/harness/viewerSortOutcome'

// Execute precisely the serialized browser callbacks, against jsdom. No browser is started.
const page = { evaluate: async (fn: (arg: unknown) => unknown, arg: unknown) => window.eval(`(${fn.toString()})`)(arg) } as unknown as Pick<Page, 'evaluate'>
const header = () => document.querySelector('[role=columnheader]')!
const status = () => document.querySelector('[data-officeview-ui=viewer-status]')!
const before = 'Visible body 1000000 / complete body 1000000'

beforeEach(() => {
  vi.useFakeTimers()
  document.body.innerHTML = `<div role="columnheader" aria-label="Column G" aria-sort="none"></div><div data-officeview-ui="viewer-status">${before}</div>`
})
afterEach(() => { vi.useRealTimers(); document.body.innerHTML = '' })

describe('sort operation outcomes', () => {
  it('records a real result separately from its acknowledgment', async () => {
    const result = await measureSortOutcome(page, async () => {
      header().setAttribute('aria-sort', 'ascending')
      status().textContent = `${before} — View-only sort applied; formulas unchanged.`
    })
    expect(result).toMatchObject({ outcome: 'completed', ariaSort: 'ascending' })
    const series = { operations: [], endToEnd: [] }, errors: string[] = []
    recordSortOutcome(series, result, errors)
    expect(series.endToEnd).toEqual([result.elapsedMs])
    expect(errors).toEqual([])
    expect(vi.getTimerCount()).toBe(0)
  })

  it('ends immediately on refusal and excludes it from successful end-to-end timings', async () => {
    const result = await measureSortOutcome(page, async () => {
      status().textContent = `${before} — Operation refused: main memory budget exceeded.`
    })
    expect(result).toMatchObject({ outcome: 'refused', elapsedMs: 0, ariaSort: 'none' })
    expect(result.status).toContain('main memory budget exceeded')
    const series = { operations: [], endToEnd: [] }, errors: string[] = []
    recordSortOutcome(series, result, errors)
    expect(series.endToEnd).toEqual([])
    expect(errors).toEqual([expect.stringContaining('sort sample 1: refused')])
    expect(errors[0]).toContain('main memory budget exceeded')
    expect(vi.getTimerCount()).toBe(0)
  })

  it('does not mistake the pending acknowledgment for a terminal refusal', async () => {
    let settled = false
    const pending = measureSortOutcome(page, async () => {
      status().textContent = `${before} — Sorting column G…`
      setTimeout(() => { status().textContent = `${before} — Operation refused: resource limit.` }, 25)
    })
    void pending.then(() => { settled = true })
    await vi.advanceTimersByTimeAsync(24)
    expect(settled).toBe(false)
    await vi.advanceTimersByTimeAsync(1)
    expect(await pending).toMatchObject({ outcome: 'refused', elapsedMs: 25 })
    expect(vi.getTimerCount()).toBe(0)
  })

  it('records a bounded timeout with the last status, never a completed 120-second sort', async () => {
    const pending = measureSortOutcome(page, async () => { status().textContent = `${before} — Sorting column G…` }, 120_000)
    await vi.advanceTimersByTimeAsync(120_000)
    const result = await pending
    expect(result).toMatchObject({ outcome: 'timeout', elapsedMs: 120_000, ariaSort: 'none' })
    expect(result.status).toContain('Sorting column G')
    const series = { operations: [], endToEnd: [] }, errors: string[] = []
    recordSortOutcome(series, result, errors)
    expect(series.endToEnd).toEqual([])
    expect(errors).toEqual([expect.stringContaining('sort sample 1: timeout')])
    expect(vi.getTimerCount()).toBe(0)
  })

  it('ignores a refusal left over from an earlier operation and unrelated status changes', async () => {
    status().textContent = `${before} — Operation refused: earlier attempt.`
    const pending = measureSortOutcome(page, async () => {}, 100)
    await vi.advanceTimersByTimeAsync(100)
    expect((await pending).outcome).toBe('timeout')
    status().textContent = before
    const unrelated = measureSortOutcome(page, async () => { status().textContent = `${before} — Font loaded.` }, 100)
    await vi.advanceTimersByTimeAsync(100)
    expect((await unrelated).outcome).toBe('timeout')
  })

  it('recognizes the same refusal again after this operation visibly enters pending', async () => {
    const refusal = `${before} — Operation refused: same budget limit.`
    status().textContent = refusal
    const pending = measureSortOutcome(page, async () => {
      status().textContent = `${before} — Sorting column G… ${refusal}`
      setTimeout(() => { status().textContent = refusal }, 25)
    }, 100)
    await vi.advanceTimersByTimeAsync(100)
    expect(await pending).toMatchObject({ outcome: 'refused' })
  })

  it('cleans up observation when the driver click fails', async () => {
    const failed = measureSortOutcome(page, async () => { throw new Error('button detached') })
    await expect(failed).rejects.toThrow('button detached')
    expect(vi.getTimerCount()).toBe(0)
  })
})
