// @vitest-environment node
/**
 * S10.2 review round 2, finding 4: the memory session's periodic probes must
 * never leak a rejection (a closed page rejects CDP calls) and must stop on
 * every exit path. Vitest fails the run on any unhandled rejection.
 */
import { describe, expect, it } from 'vitest'
import { startSampler } from '../scripts/harness/sampler'

const wait = (ms: number) => new Promise(resolve => setTimeout(resolve, ms))

describe('startSampler', () => {
  it('reports a rejecting probe while running, never overlaps itself, and stops for good', async () => {
    const errors: unknown[] = []
    let calls = 0, active = 0, maxActive = 0
    const stop = startSampler(2, async () => {
      calls++; active++; maxActive = Math.max(maxActive, active)
      await wait(5)
      active--
      throw new Error('Target page, context or browser has been closed')
    }, error => errors.push(error))
    await wait(40)
    stop()
    const after = calls
    await wait(30)
    expect(errors.length).toBeGreaterThan(0)
    expect((errors[0] as Error).message).toMatch(/has been closed/)
    expect(maxActive).toBe(1)
    expect(calls).toBe(after) // no probe starts after stop
  })

  it('a probe still in flight when stopped settles silently, even if it rejects', async () => {
    const errors: unknown[] = []
    let release: (error: Error) => void = () => {}
    const stop = startSampler(1, () => new Promise<void>((_, reject) => { release = reject }), error => errors.push(error))
    await wait(10)
    stop()
    release(new Error('closed during shutdown'))
    await wait(10)
    expect(errors).toEqual([])
  })
})
