import { describe, expect, it } from 'vitest'
import { AxisMetricsError, createAxisMetrics } from '../src/xlsx/view/axes'
import { MAX_ROWS, MAX_COLS } from '../src/xlsx/view/coordinates'
import type { AxisMetrics, AxisRun } from '../src/xlsx/view/contracts'

describe('sparse axis metrics (spec §§3, 5.3, AC-VIEW-03)', () => {
  it('rejects a positive cell that cannot advance the Float64 logical prefix', () => {
    expect(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, overrides: [
      { start: 0, end: 0, sizeCss: 25_165_824 }, { start: 1, end: 1, sizeCss: 1e-10 },
    ] })).toThrow(/precision|represent/i)
  })

  it('rejects sub-ULP cells even when their uniform segment total advances', () => {
    expect(() => createAxisMetrics({ logicalCount: 102, defaultSizeCss: 1, overrides: [
      { start: 0, end: 0, sizeCss: 25_165_824 }, { start: 1, end: 100, sizeCss: 1e-10 },
    ] })).toThrow(/precision|represent/i)
  })

  it('checks resolution at the far end of a uniform span without imposing a global size floor', () => {
    const base = 2 ** 25 - 2 ** -28
    expect(() => createAxisMetrics({ logicalCount: 5, defaultSizeCss: 2 ** -28,
      overrides: [{ start: 0, end: 0, sizeCss: base }],
    })).toThrow(/precision|represent/i)
    const subnormal = createAxisMetrics({ logicalCount: 3, defaultSizeCss: Number.MIN_VALUE })
    expect(subnormal.start(1)).toBe(Number.MIN_VALUE)
    expect(subnormal.start(2)).toBe(2 * Number.MIN_VALUE)
    expect(subnormal.locate(Number.MIN_VALUE)).toBe(1)
    const tiny = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 1e-6 })
    expect(tiny.locate(tiny.start(MAX_ROWS - 1))).toBe(MAX_ROWS - 1)
    const exactBoundary = createAxisMetrics({ logicalCount: 2, defaultSizeCss: 2 ** -28,
      overrides: [{ start: 0, end: 0, sizeCss: base }],
    })
    expect(exactBoundary.totalCss()).toBe(2 ** 25)
    expect(exactBoundary.locate(base)).toBe(1)
    const collapsed = createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1,
      overrides: [{ start: 0, end: 0, sizeCss: 25_165_824 }, { start: 1, end: 1, sizeCss: 1e-10 }],
      collapsed: [{ start: 1, end: 1 }],
    })
    expect(collapsed.size(1)).toBe(0)
    expect(collapsed.locate(collapsed.start(2))).toBe(2)
  })

  it('covers the last row and XFD with constant-size blank geometry and Float64 precision', () => {
    const rows = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 })
    const cols = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 })
    expect(rows.totalCss()).toBe(25_165_824)
    expect(rows.start(MAX_ROWS - 1)).toBe(25_165_800)
    expect(rows.locate(25_165_800.25)).toBe(MAX_ROWS - 1)
    expect(cols.start(MAX_COLS - 1)).toBe(1_048_512)
    expect(cols.locate(cols.start(MAX_COLS - 1) + 0.125)).toBe(MAX_COLS - 1)
    expect(rows.diagnostics.segmentCount).toBe(1)
    expect(rows.diagnostics.summaryBytes).toBeLessThan(64)
    expect([...rows.visibleBetween(25_165_800.25, rows.totalCss()).runs]).toEqual([
      { start: MAX_ROWS - 1, end: MAX_ROWS - 1, startCss: 25_165_800, sizeCss: 24 },
    ])
  })

  it('canonicalizes sparse overrides and gives collapsed intervals precedence', () => {
    const axis = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 10,
      overrides: [{ start: 700_000, end: 900_000, sizeCss: 3 }, { start: 2, end: 4, sizeCss: 5 },
        { start: 4, end: 6, sizeCss: 5 }],
      collapsed: [{ start: 3, end: 5 }, { start: 4, end: 5 }],
    })
    expect(axis.overrides).toEqual([
      { start: 2, end: 2, sizeCss: 5 }, { start: 3, end: 5, sizeCss: 0 },
      { start: 6, end: 6, sizeCss: 5 }, { start: 700_000, end: 900_000, sizeCss: 3 },
    ])
    expect([axis.start(2), axis.start(3), axis.start(6), axis.start(7)]).toEqual([20, 25, 25, 30])
    expect(axis.size(4)).toBe(0)
    expect(axis.totalCss()).toBe(MAX_ROWS * 10 - 40 - 200_001 * 7)
    expect(axis.diagnostics.segmentCount).toBe(7)
    expect(axis.diagnostics.summaryBytes).toBeLessThan(400)
    expect([...axis.visibleBetween(24.5, 30).runs]).toEqual([
      { start: 2, end: 2, startCss: 20, sizeCss: 5 },
      { start: 6, end: 6, startCss: 25, sizeCss: 5 },
    ])
  })

  it('locates fractional offsets and exact half-open boundaries across zero-size spans', () => {
    const axis = createAxisMetrics({ logicalCount: 8, defaultSizeCss: 0.1,
      overrides: [{ start: 3, end: 4, sizeCss: 0 }, { start: 6, end: 6, sizeCss: 0.25 }],
    })
    expect(axis.locate(0.05)).toBe(0)
    expect(axis.locate(axis.start(3))).toBe(5)
    expect(axis.locate(axis.start(6))).toBe(6)
    expect(axis.locate(axis.start(7))).toBe(7)
    expect(axis.locate(axis.totalCss())).toBeNull()
    expect([...axis.visibleBetween(axis.start(5), axis.start(6)).runs]).toEqual([
      { start: 5, end: 5, startCss: axis.start(5), sizeCss: 0.1 },
    ])
  })

  it('answers all-hidden axes without visiting their million collapsed coordinates', () => {
    for (const options of [
      { logicalCount: MAX_ROWS, defaultSizeCss: 0 },
      { logicalCount: MAX_ROWS, defaultSizeCss: 24, collapsed: [{ start: 0, end: MAX_ROWS - 1 }] },
    ]) {
      const axis = createAxisMetrics(options)
      expect(axis.totalCss()).toBe(0)
      expect(axis.start(MAX_ROWS - 1)).toBe(0)
      expect(axis.locate(0)).toBeNull()
      expect([...axis.visibleBetween(-1, 1).runs]).toEqual([])
      expect(axis.diagnostics.lastWindowVisitedSegments).toBe(0)
      expect(axis.diagnostics.segmentCount).toBe(1)
    }
  })

  it('compresses a million tiny cells into one run without a visible-cell array', () => {
    const axis = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 0.000_001 })
    const window = axis.visibleBetween(0, 2)
    expect(window.count).toBe(MAX_ROWS)
    expect([...window.runs]).toEqual([{ start: 0, end: MAX_ROWS - 1, startCss: 0, sizeCss: 0.000_001 }])
    expect(window.nextCursor).toBeUndefined()
    expect(axis.diagnostics.lastWindowVisitedSegments).toBe(1)
  })

  it('bounds variable runs and resumes precisely while skipping huge hidden spans', () => {
    const axis = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 1, maxWindowRuns: 2,
      overrides: [{ start: 1, end: 1, sizeCss: 2 }, { start: 3, end: 3, sizeCss: 3 }],
      collapsed: [{ start: 5, end: MAX_ROWS - 2 }],
    })
    let page = axis.visiblePageBetween(0, 100)
    const runs = [...page.runs]
    let count = page.count
    let pages = 1
    while (page.nextCursor) {
      expect(axis.diagnostics.lastWindowVisitedSegments).toBeLessThanOrEqual(2)
      page = axis.continueVisibleBetween(0, 100, page.nextCursor)
      runs.push(...page.runs)
      count += page.count
      pages++
      expect(pages).toBeLessThan(5)
    }
    expect(runs.map(run => [run.start, run.end])).toEqual([[0, 0], [1, 1], [2, 2], [3, 3], [4, 4], [MAX_ROWS - 1, MAX_ROWS - 1]])
    expect(count).toBe(6)
    expect(pages).toBe(3)
    const cursor = axis.visiblePageBetween(0, 100).nextCursor!
    expect(() => axis.continueVisibleBetween(0, 99, cursor)).toThrow(RangeError)
    expect(() => createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 1 }).continueVisibleBetween(0, 100, cursor)).toThrow(RangeError)
    expect(() => axis.continueVisibleBetween(0, 100, 'bad')).toThrow(RangeError)
  })

  it('uses logarithmic lookup summaries for many size changes', () => {
    const axis = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24,
      overrides: Array.from({ length: 10_000 }, (_, i) => ({ start: i * 100, end: i * 100, sizeCss: 12 })),
    })
    expect(axis.locate(axis.start(800_500) + 1)).toBe(800_500)
    expect(axis.diagnostics.lastLocateComparisons).toBeLessThan(20)
    expect([...axis.visibleBetween(axis.start(800_501), axis.start(800_505)).runs]).toEqual([
      { start: 800_501, end: 800_504, startCss: axis.start(800_501), sizeCss: 24 },
    ])
    expect(axis.diagnostics.lastWindowVisitedSegments).toBe(1)
  })

  it('validates counts, intervals, sizes, boundaries and cursor limits', () => {
    for (const logicalCount of [-1, -0, 0.5, NaN, Infinity, MAX_ROWS + 1]) {
      expect(() => createAxisMetrics({ logicalCount, defaultSizeCss: 1 })).toThrow(RangeError)
    }
    for (const defaultSizeCss of [-1, NaN, Infinity, Number.MAX_VALUE]) {
      expect(() => createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss })).toThrow(RangeError)
    }
    for (const interval of [{ start: -1, end: 1 }, { start: 1.5, end: 2 }, { start: 2, end: 1 }, { start: 0, end: 3 }]) {
      expect(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, collapsed: [interval] })).toThrow(RangeError)
    }
    expect(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, overrides: [{ start: 0, end: 0, sizeCss: -1 }] })).toThrow(RangeError)
    expect(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, overrides: [
      { start: 0, end: 1, sizeCss: 2 }, { start: 1, end: 2, sizeCss: 3 },
    ] })).toThrow(RangeError)
    for (const maxWindowRuns of [0, -1, 0.5, Infinity, MAX_ROWS + 1]) {
      expect(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, maxWindowRuns })).toThrow(RangeError)
    }
    const axis = createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1 })
    for (const index of [-1, -0, 0.5, NaN, 4]) {
      expect(() => axis.start(index)).toThrow(RangeError)
      expect(() => axis.size(index)).toThrow(RangeError)
    }
    for (const css of [NaN, Infinity, -Infinity]) {
      expect(() => axis.locate(css)).toThrow(RangeError)
      expect(() => axis.visibleBetween(0, css)).toThrow(RangeError)
    }
    expect(axis.locate(-1)).toBeNull()
    expect(axis.visibleBetween(4, 5).count).toBe(0)
    expect(axis.visibleBetween(2, 2).count).toBe(0)
    expect(() => axis.visibleBetween(2, 1)).toThrow(RangeError)
    expect([...axis.visibleBetween(-1, 0.5).runs]).toEqual([{ start: 0, end: 0, startCss: 0, sizeCss: 1 }])
    const empty = createAxisMetrics({ logicalCount: 0, defaultSizeCss: 1 })
    expect(empty.locate(0)).toBeNull()
    expect(empty.visibleBetween(0, 1).count).toBe(0)
  })

  it('allows the terminal start prefix while keeping size inside the index domain', () => {
    const axis = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 })
    expect(axis.start(MAX_ROWS)).toBe(axis.totalCss())
    expect(() => axis.size(MAX_ROWS)).toThrow(RangeError)
    const empty = createAxisMetrics({ logicalCount: 0, defaultSizeCss: 0 })
    expect(empty.start(0)).toBe(0)
    expect(() => empty.size(0)).toThrow(RangeError)
  })

  it('keeps locate diagnostics scoped to the last locate call across window queries', () => {
    const axis = createAxisMetrics({ logicalCount: 100, defaultSizeCss: 1,
      overrides: [{ start: 50, end: 59, sizeCss: 2 }],
    })
    expect(axis.locate(51)).toBe(50)
    const comparisons = axis.diagnostics.lastLocateComparisons
    axis.visibleBetween(0, 10)
    expect(axis.diagnostics.lastLocateComparisons).toBe(comparisons)
  })

  it('agrees with an independent dense oracle for randomized dimensions, hits and paged windows', () => {
    let seed = 17
    const random = () => {
      seed = seed * 16807 % 2147483647
      return seed
    }
    for (let trial = 0; trial < 20; trial++) {
      const authoredSizes = Array.from({ length: 127 }, () => random() % 9 / 4)
      const collapsed = authoredSizes.flatMap((_, start) => random() % 6 === 0 ? [{ start, end: start }] : [])
      const sizes = [...authoredSizes]
      for (const interval of collapsed) sizes[interval.start] = 0
      const prefix = [0]
      for (const size of sizes) prefix.push(prefix[prefix.length - 1] + size)
      const axis = createAxisMetrics({ logicalCount: sizes.length, defaultSizeCss: 1, maxWindowRuns: 3,
        overrides: authoredSizes.map((sizeCss, start) => ({ start, end: start, sizeCss })), collapsed,
      })
      expect(axis.totalCss()).toBe(prefix[sizes.length])
      for (let index = 0; index < sizes.length; index++) {
        expect(axis.start(index)).toBe(prefix[index])
        expect(axis.size(index)).toBe(sizes[index])
        for (const css of [prefix[index], prefix[index] + 0.125]) {
          const expected = sizes.findIndex((size, row) => size > 0 && prefix[row] <= css && prefix[row + 1] > css)
          expect(axis.locate(css)).toBe(expected < 0 ? null : expected)
        }
      }
      for (let query = 0; query < 10; query++) {
        const lower = random() % 400 / 4 - 2
        const upper = lower + random() % 200 / 4
        const expected = sizes.flatMap((size, row) => size > 0 && lower < upper
          && prefix[row] < upper && prefix[row + 1] > lower ? [row] : [])
        const actual: number[] = []
        let page = axis.visiblePageBetween(lower, upper)
        let count = 0
        let pages = 0
        for (;;) {
          expect(++pages).toBeLessThan(128)
          expect([...page.runs].length).toBeLessThanOrEqual(3)
          expect(axis.diagnostics.lastWindowVisitedSegments).toBeLessThanOrEqual(3)
          count += page.count
          for (const run of page.runs) {
            for (let row = run.start; row <= run.end; row++) actual.push(row)
          }
          if (!page.nextCursor) break
          page = axis.continueVisibleBetween(lower, upper, page.nextCursor)
        }
        expect(actual).toEqual(expected)
        expect(count).toBe(expected.length)
      }
    }
  })
})

describe('contract-typed lazy windows (review item 1)', () => {
  it('a consumer typed against AxisMetrics sees every run of 1,000 alternating hidden/visible rows', () => {
    const axis: AxisMetrics = createAxisMetrics({ logicalCount: 1000, defaultSizeCss: 1,
      collapsed: Array.from({ length: 500 }, (_, i) => ({ start: i * 2 + 1, end: i * 2 + 1 })) })
    const window = axis.visibleBetween(0, 450)
    const runs = [...window.runs]
    expect(runs).toHaveLength(450)
    expect(runs.map(run => run.start)).toEqual(Array.from({ length: 450 }, (_, i) => i * 2))
    expect(window.count).toBe(450)
    expect(window.nextCursor).toBeUndefined()
  })

  it('iterates a 1M alternating layout exactly once per run, bounded per page, and is re-iterable', () => {
    const maxWindowRuns = 64
    const concrete = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 1, maxWindowRuns,
      overrides: Array.from({ length: MAX_ROWS / 2 }, (_, i) => ({ start: i * 2 + 1, end: i * 2 + 1, sizeCss: 2 })) })
    const axis: AxisMetrics = concrete
    const window = axis.visibleBetween(0, axis.totalCss())
    expect(window.count).toBe(MAX_ROWS)
    expect(window.nextCursor).toBeUndefined()
    for (let pass = 0; pass < 2; pass++) {
      let expectedStart = 0
      let expectedCss = 0
      let peakVisited = 0
      let runsSeen = 0
      let exact = true
      // expect() per run is far too slow for a million runs; fold into one boolean.
      for (const run of window.runs as Iterable<AxisRun>) {
        if (run.start !== expectedStart || run.end !== expectedStart || run.startCss !== expectedCss) exact = false
        expectedStart++
        expectedCss += run.sizeCss
        runsSeen++
        peakVisited = Math.max(peakVisited, concrete.diagnostics.lastWindowVisitedSegments)
      }
      expect(exact).toBe(true)
      expect(runsSeen).toBe(MAX_ROWS)
      expect(expectedCss).toBe(axis.totalCss())
      expect(peakVisited).toBeLessThanOrEqual(maxWindowRuns)
    }
  }, 30_000)

  it('starts inside the window mid-layout and supports abandoning iteration early', () => {
    const axis = createAxisMetrics({ logicalCount: 2000, defaultSizeCss: 1, maxWindowRuns: 4,
      collapsed: Array.from({ length: 1000 }, (_, i) => ({ start: i * 2 + 1, end: i * 2 + 1 })) })
    const first: number[] = []
    for (const run of axis.visibleBetween(100.5, 1000).runs) {
      first.push(run.start)
      if (first.length === 10) break
    }
    expect(first).toEqual([200, 202, 204, 206, 208, 210, 212, 214, 216, 218])
  })

  it('reports count equal to the number of covered indices for randomized windows', () => {
    let seed = 99
    const random = () => (seed = seed * 16807 % 2147483647)
    for (let trial = 0; trial < 30; trial++) {
      const sizes = Array.from({ length: 90 }, () => random() % 5 / 2)
      const axis = createAxisMetrics({ logicalCount: sizes.length, defaultSizeCss: 1, maxWindowRuns: 3,
        overrides: sizes.map((sizeCss, start) => ({ start, end: start, sizeCss })) })
      for (let query = 0; query < 20; query++) {
        const lower = random() % 300 / 4 - 3
        const upper = lower + random() % 120 / 4
        const window = axis.visibleBetween(lower, upper)
        let total = 0
        for (const run of window.runs) total += run.end - run.start + 1
        expect(window.count).toBe(total)
      }
    }
  })

  it('keeps the eager bounded page API for cursor consumers', () => {
    const axis = createAxisMetrics({ logicalCount: 100, defaultSizeCss: 1, maxWindowRuns: 3,
      collapsed: Array.from({ length: 50 }, (_, i) => ({ start: i * 2 + 1, end: i * 2 + 1 })) })
    const page = axis.visiblePageBetween(0, 50)
    expect([...page.runs]).toHaveLength(3)
    expect(page.count).toBe(3)
    expect(page.nextCursor).toBeDefined()
    expect(axis.continueVisibleBetween(0, 50, page.nextCursor!).count).toBe(3)
  })
})

describe('typed construction errors (review item 4)', () => {
  const codeOf = (build: () => unknown): string | undefined => {
    try { build() } catch (error) {
      expect(error).toBeInstanceOf(AxisMetricsError)
      expect(error).toBeInstanceOf(RangeError)
      expect((error as Error).name).toBe('AxisMetricsError')
      return (error as AxisMetricsError).code
    }
    return undefined
  }
  it('reports CONTRADICTORY_OVERRIDES for overlapping overrides with different sizes', () => {
    expect(codeOf(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, overrides: [
      { start: 0, end: 1, sizeCss: 2 }, { start: 1, end: 2, sizeCss: 3 },
    ] }))).toBe('CONTRADICTORY_OVERRIDES')
  })
  it('reports PRECISION_EXCEEDED for cells Float64 cannot tell apart', () => {
    expect(codeOf(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, overrides: [
      { start: 0, end: 0, sizeCss: 25_165_824 }, { start: 1, end: 1, sizeCss: 1e-10 },
    ] }))).toBe('PRECISION_EXCEEDED')
  })
  it('reports TOTAL_NOT_FINITE when the axis total overflows', () => {
    expect(codeOf(() => createAxisMetrics({ logicalCount: 4, defaultSizeCss: 1,
      overrides: [{ start: 0, end: 3, sizeCss: Number.MAX_VALUE }] }))).toBe('TOTAL_NOT_FINITE')
  })
  it('reports INVALID_METADATA for out-of-domain counts, sizes and intervals', () => {
    expect(codeOf(() => createAxisMetrics({ logicalCount: MAX_ROWS + 1, defaultSizeCss: 1 }))).toBe('INVALID_METADATA')
    expect(codeOf(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: NaN }))).toBe('INVALID_METADATA')
    expect(codeOf(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, maxWindowRuns: 0 }))).toBe('INVALID_METADATA')
    expect(codeOf(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, collapsed: [{ start: 0, end: 3 }] }))).toBe('INVALID_METADATA')
    expect(codeOf(() => createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1, overrides: [{ start: 0, end: 0, sizeCss: -1 }] }))).toBe('INVALID_METADATA')
  })
  it('keeps query-time domain errors as plain RangeError', () => {
    const axis = createAxisMetrics({ logicalCount: 3, defaultSizeCss: 1 })
    try { axis.start(9) } catch (error) { expect(error).not.toBeInstanceOf(AxisMetricsError) }
  })
})
