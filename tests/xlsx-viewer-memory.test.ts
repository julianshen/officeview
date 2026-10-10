import { spawnSync } from 'node:child_process'
import { describe, expect, it } from 'vitest'

// Real retained-memory checks (review items 2 and 5). The modules' own diagnostics count
// only what they choose to count, so these run each scenario in a child `bun` process that
// can force gc and read process.memoryUsage(); see the probe for the measurement method.
const MiB = 1024 * 1024
const hasBun = spawnSync('bun', ['--version']).status === 0
function probe(scenario: string): { externalBytes: number; heapBytes: number } {
  const run = spawnSync('bun', ['tests/probes/xlsx-viewer-memory-probe.ts', scenario], { encoding: 'utf8', timeout: 120_000 })
  if (run.status !== 0) throw new Error(`probe ${scenario} failed: ${run.stderr}`)
  return JSON.parse(run.stdout.trim().split('\n').pop()!)
}

describe.skipIf(!hasBun)('retained typed-array memory in a forced-gc child process', () => {
  it('the full-capacity identity projection retains no typed arrays and no row objects', () => {
    const { externalBytes, heapBytes } = probe('identity')
    // Baseline noise is ~65 KiB; Float64Array(MAX_ROWS) alone would be 8 MiB.
    expect(externalBytes).toBeLessThan(256 * 1024)
    // A million row/cell objects would be tens of MiB; a few KiB of closures is expected.
    expect(heapBytes).toBeLessThan(1 * MiB)
  }, 130_000)
  it('a 1M-row band order retains exactly order + inverse (8 MiB), not the caller input', () => {
    const { externalBytes } = probe('order')
    expect(externalBytes).toBeGreaterThanOrEqual(8 * MiB)
    expect(externalBytes).toBeLessThan(8 * MiB + 512 * 1024)
  }, 130_000)
  it('the full variable-band projection stays inside the spec 12.4 28-30 MiB workspace', () => {
    const { externalBytes } = probe('projection')
    expect(externalBytes).toBeGreaterThanOrEqual(28 * MiB)
    expect(externalBytes).toBeLessThanOrEqual(30 * MiB)
  }, 130_000)
})
