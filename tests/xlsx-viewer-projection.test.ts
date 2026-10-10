import { describe, expect, it } from 'vitest'
import { AxisMetricsError, createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { MAX_ROWS, MAX_COLS } from '../src/xlsx/view/coordinates'
import { createIdentityOrder, createRowOrder, pinRowPermutation } from '../src/xlsx/view/order'
import { CompactRowVisibility, createRowVisibility, createIntervalSet, visibilityFromHidden, isVisibleSource, normalizeFilter, projectVisibility } from '../src/xlsx/view/visibility'
import type { AxisMetrics, FilterSpec, SheetMetrics } from '../src/xlsx/view/contracts'

function shuffled(length: number, seed: number): Uint32Array {
  const rows = Uint32Array.from({ length }, (_, i) => i)
  let state = seed
  for (let i = length - 1; i > 0; i--) {
    state = (state * 1664525 + 1013904223) % 4294967296
    const j = state % (i + 1)
    const value = rows[i]; rows[i] = rows[j]; rows[j] = value
  }
  return rows
}

describe('row identity and permutations (AC-VIEW-01)', () => {
  it('constructs arithmetic identity with zero typed row storage', () => {
    const order = createIdentityOrder()
    expect(order.kind).toBe('identity')
    expect(order.bands).toHaveLength(0)
    expect(order.diagnostics).toEqual({ typedArrayBytes: 0, bandCount: 0 })
    expect(order.displayToSource(MAX_ROWS - 1)).toBe(MAX_ROWS - 1)
    expect(order.sourceToDisplay(16385)).toBe(16385)
    expect(() => order.displayToSource(MAX_ROWS)).toThrow()
  })
  it.each([[0, 3, 1, 2], [0, 2, 1]])('keeps display slots distinct for order %j', (...rows) => {
    const values = rows as number[]
    const order = createRowOrder([{ r0: 0, r1: values.length - 1, sourceAtOffset: new Uint32Array(values) }])
    values.forEach((source, slot) => {
      expect(order.displayToSource(slot)).toBe(source)
      expect(order.sourceToDisplay(source)).toBe(slot)
    })
    expect(order.displayToSource(16385)).toBe(16385)
  })
  it('normalizes band order, snapshots inputs and preserves identity outside bands', () => {
    const input = new Uint32Array([12, 10, 11])
    const order = createRowOrder([{ r0: 10, r1: 12, sourceAtOffset: input }, { r0: 2, r1: 3, sourceAtOffset: new Uint32Array([3, 2]) }])
    input[0] = 10
    expect(order.bands.map(b => b.r0)).toEqual([2, 10])
    expect(order.displayToSource(10)).toBe(12)
    expect(order.displayToSource(9)).toBe(9)
    expect(order.sourceToDisplay(11)).toBe(12)
  })
  it('rejects malformed, duplicate, outside, fractional and overlapping bands', () => {
    for (const band of [
      { r0: -1, r1: 0, sourceAtOffset: new Uint32Array([0, 1]) },
      { r0: 0.5, r1: 1, sourceAtOffset: new Uint32Array([0]) },
      { r0: 1, r1: 0, sourceAtOffset: new Uint32Array() },
      { r0: 0, r1: 1, sourceAtOffset: new Uint32Array([0]) },
      { r0: 0, r1: 1, sourceAtOffset: new Uint32Array([0, 0]) },
      { r0: 0, r1: 1, sourceAtOffset: new Uint32Array([0, 2]) },
    ]) expect(() => createRowOrder([band])).toThrow()
    expect(() => createRowOrder([{ r0: 0, r1: 1, sourceAtOffset: new Uint32Array([1, 0]) }, { r0: 1, r1: 2, sourceAtOffset: new Uint32Array([2, 1]) }])).toThrow()
  })
  it('is bijective over seeded permutations and a full million-row band', () => {
    for (const length of [7, 257, MAX_ROWS]) {
      const started = performance.now()
      const input = shuffled(length, 23)
      const order = createRowOrder([{ r0: 0, r1: length - 1, sourceAtOffset: input }])
      let bijective = true
      for (let slot = 0; slot < length; slot++) {
        if (order.displayToSource(slot) !== input[slot] || order.sourceToDisplay(input[slot]) !== slot) { bijective = false; break }
      }
      expect(bijective).toBe(true)
      expect(order.diagnostics.typedArrayBytes).toBe(length * 8)
      if (length === MAX_ROWS) console.info('S2.1 full-band order ledger', { elapsedMs: performance.now() - started, retainedBytes: order.diagnostics.typedArrayBytes, callerInputBytes: input.byteLength })
    }
  })
  it('pins hidden/header/totals slots while permuting visible candidates', () => {
    const pins = createRowVisibility(6, slot => slot === 0 || slot === 2 || slot === 5)
    const band = pinRowPermutation(0, 5, pins, new Uint32Array([4, 1, 3]))
    expect(Array.from(band.sourceAtOffset)).toEqual([0, 4, 2, 1, 3, 5])
    const order = createRowOrder([band])
    expect(order.sourceToDisplay(2)).toBe(2)
    expect(() => pinRowPermutation(0, 5, pins, new Uint32Array([4, 2, 3]))).toThrow()
    expect(() => pinRowPermutation(0, 5, pins, new Uint32Array([4, 1]))).toThrow()
  })
})

describe('hierarchical display visibility and source snapshots (AC-SEM-01)', () => {
  it('matches naive rank/select/strict-next across word and summary boundaries', () => {
    const predicate = (slot: number) => slot % 7 === 0 || slot === 31 || slot === 32 || slot === 1023 || slot === 1024
    const mask = createRowVisibility(2081, predicate)
    const visible = Array.from({ length: 2081 }, (_, i) => i).filter(predicate)
    for (let slot = 0; slot < 2081; slot++) {
      expect(mask.visible(slot)).toBe(predicate(slot))
      expect(mask.rank(slot)).toBe(visible.filter(row => row < slot).length)
      expect(mask.next(slot, 1)).toBe(visible.find(row => row > slot) ?? null)
      expect(mask.next(slot, -1)).toBe(visible.filter(row => row < slot).at(-1) ?? null)
    }
    visible.forEach((slot, ordinal) => expect(mask.select(ordinal)).toBe(slot))
    expect(mask.rank(2081)).toBe(visible.length)
    expect(mask.select(visible.length)).toBeNull()
  })
  it('handles full-capacity all-visible/all-hidden masks, padding and boundary validation', () => {
    const visible = createRowVisibility()
    const hidden = visibilityFromHidden(createIntervalSet([{ r0: 0, r1: MAX_ROWS - 1 }]))
    expect(visible.words.byteLength).toBe(131072)
    expect(visible.summaries.byteLength).toBeLessThanOrEqual(8192)
    expect(visible.rank(MAX_ROWS)).toBe(MAX_ROWS)
    expect(visible.select(MAX_ROWS - 1)).toBe(MAX_ROWS - 1)
    expect(hidden.rank(MAX_ROWS)).toBe(0)
    expect(hidden.select(0)).toBeNull()
    expect(hidden.next(MAX_ROWS - 1, -1)).toBeNull()
    expect(() => visible.rank(MAX_ROWS + 1)).toThrow()
    expect(() => visible.visible(NaN)).toThrow()
    const odd = createRowVisibility(33)
    expect(odd.words[1]).toBe(1)
    expect(odd.rank(33)).toBe(33)
    console.info('S2.1 visibility ledger', { maskBytes: visible.words.byteLength, summaryBytes: visible.summaries.byteLength })
  })
  it('coalesces immutable snapshot intervals without per-hidden-row objects', () => {
    const input = [{ r0: 7, r1: 9 }, { r0: 1, r1: 4 }, { r0: 4, r1: 6 }]
    const set = createIntervalSet(input)
    input[0].r0 = 0
    expect(set.intervals).toEqual([{ r0: 1, r1: 9 }])
    expect(set.contains(5)).toBe(true)
    expect(set.contains(0)).toBe(false)
    expect(visibilityFromHidden(set, 16).rank(16)).toBe(7)
    expect(() => createIntervalSet([{ r0: 5, r1: 4 }])).toThrow()
  })
  it('re-evaluates only body rows; protects headers/totals/outside; restores snapshot on Clear', () => {
    const baseHidden = createIntervalSet([{ r0: 0, r1: 0 }, { r0: 2, r1: 2 }, { r0: 4, r1: 5 }])
    const filter: FilterSpec = { range: { r0: 0, r1: 4, c0: 0, c1: 2, headerRows: 1, totalsRows: 1 }, columns: new Map([[1, { kind: 'number', op: 'gt', first: 0 }]]) }
    const visited: number[] = []
    const match = (row: number) => { visited.push(row); return row !== 3 }
    const applied = projectVisibility(baseHidden, filter, createIdentityOrder(), match, 8)
    expect(Array.from({ length: 8 }, (_, row) => applied.visibility.visible(row))).toEqual([false, true, true, false, false, false, true, true])
    expect(visited).toEqual([1, 2, 3])
    expect(isVisibleSource(2, { baseHidden, filter }, () => true)).toBe(true)
    const replacement = { ...filter, range: { ...filter.range, r0: 5, r1: 7, headerRows: 0 as const, totalsRows: 0 as const } }
    const replaced = projectVisibility(baseHidden, replacement, createIdentityOrder(), () => true, 8)
    expect(replaced.visibility.visible(2)).toBe(false)
    expect(replaced.visibility.visible(5)).toBe(true)
    const cleared = projectVisibility(baseHidden, { ...filter, columns: new Map() }, createIdentityOrder(), () => { throw new Error('No rule callback on Clear') }, 8)
    expect(cleared.filter).toBeNull()
    expect(normalizeFilter({ ...filter, columns: new Map() })).toBeNull()
    expect(Array.from({ length: 8 }, (_, row) => cleared.visibility.visible(row))).toEqual([false, true, false, true, false, false, true, true])
  })
})

describe('model boundary validation and filter permutation ownership', () => {
  it('rejects fractional candidate inputs instead of truncating them into a valid permutation', () => {
    const pins = createRowVisibility(3, () => false)
    expect(() => pinRowPermutation(0, 2, pins, [0.5, 1.5, 2] as unknown as Uint32Array)).toThrow()
    expect(() => pinRowPermutation(0, 2, pins, [0, 1, 2] as unknown as Uint32Array)).toThrow()
    expect(() => new CompactRowVisibility(3, [7] as unknown as Uint32Array)).toThrow()
  })
  it('rejects signed-zero mask dimensions', () => {
    expect(() => createRowVisibility(-0)).toThrow()
  })
  it('maps source snapshots and matches through a nonidentity order without visiting protected rows', () => {
    const order = createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: new Uint32Array([0, 3, 1, 2]) }])
    const baseHidden = createIntervalSet([{ r0: 1, r1: 1 }])
    const filter: FilterSpec = { range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    const calls: number[] = []
    const applied = projectVisibility(baseHidden, filter, order, row => { calls.push(row); return row === 1 }, 4)
    expect(calls).toEqual([1, 2])
    expect(Array.from({ length: 4 }, (_, slot) => applied.visibility.visible(slot))).toEqual([true, true, true, false])
    const cleared = projectVisibility(baseHidden, null, order, () => { throw new Error('No callback on Clear') }, 4)
    expect(Array.from({ length: 4 }, (_, slot) => cleared.visibility.visible(slot))).toEqual([true, true, false, true])
    const hiddenPredicate = isVisibleSource(1, { baseHidden, filter: null }, () => true)
    expect(hiddenPredicate).toBe(false)
  })
  it('matches seeded naive masks for empty, padding and block boundaries', () => {
    for (const count of [0, 1, 31, 32, 33, 1023, 1024, 1025]) {
      const mask = createRowVisibility(count, row => ((row * 13 + 7) % 23) < 5)
      const rows = Array.from({ length: count }, (_, row) => row).filter(row => ((row * 13 + 7) % 23) < 5)
      let valid = mask.rank(count) === rows.length
      for (let ordinal = 0; ordinal < rows.length; ordinal++) if (mask.select(ordinal) !== rows[ordinal]) valid = false
      expect(valid).toBe(true)
      expect(mask.select(-1)).toBeNull()
      expect(mask.select(NaN)).toBeNull()
      expect(mask.select(rows.length)).toBeNull()
    }
  })
})

describe('full-capacity source/display/logical projection (AC-VIEW-02/03)', () => {
  it('maps the corner and former row/column caps without allocating dense identity geometry', () => {
    const sourceRows = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 })
    const cols = createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 })
    const sheet = createSheetMetrics({ rows: sourceRows, cols })
    const corner = { row: MAX_ROWS - 1, col: MAX_COLS - 1 }
    expect(sheet.sourceCellToLogical(corner)).toEqual({ x: (MAX_COLS - 1) * 64, y: (MAX_ROWS - 1) * 24, width: 64, height: 24 })
    expect(sheet.sourceCellToLogical({ row: 16385, col: 4096 })?.y).toBe(16385 * 24)
    expect(sheet.logicalPointToSource((MAX_COLS - 1) * 64 + 0.1, (MAX_ROWS - 1) * 24 + 0.1)).toEqual(corner)
    expect(sheet.diagnostics.sectionCount).toBe(1)
    expect(sheet.diagnostics.densePrefixBytes).toBe(0)
    expect(sheet.diagnostics.typedArrayBytes).toBeLessThan(1024)
    expect(sheet.rows.totalCss()).toBe(25165824)
  })
  it('moves the maximum source row to the top without a geometry prefix for equal heights', () => {
    const values = Uint32Array.from({ length: MAX_ROWS }, (_, i) => i === 0 ? MAX_ROWS - 1 : i - 1)
    const order = createRowOrder([{ r0: 0, r1: MAX_ROWS - 1, sourceAtOffset: values }])
    const sheet = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order })
    expect(sheet.sourceCellToLogical({ row: MAX_ROWS - 1, col: 0 })?.y).toBe(0)
    expect(sheet.logicalPointToSource(1, 1)).toEqual({ row: MAX_ROWS - 1, col: 0 })
    expect(sheet.visibleRankToDisplay(1)).toBe(1)
    expect(sheet.visibleRankToSource(1)).toBe(0)
    expect(sheet.sourceToVisibleRank(0)).toBe(1)
    expect(sheet.diagnostics.densePrefixBytes).toBe(0)
  })
  it('moves source dimensions with rows, collapsing hidden rows/columns and skipping zero heights in ranks', () => {
    const order = createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: new Uint32Array([0, 3, 1, 2]) }])
    const sheet = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 5, defaultSizeCss: 10, overrides: [{ start: 1, end: 1, sizeCss: 20 }, { start: 2, end: 2, sizeCss: 0 }] }), cols: createAxisMetrics({ logicalCount: 3, defaultSizeCss: 5, collapsed: [{ start: 1, end: 1 }] }), order })
    expect(sheet.sourceCellToLogical({ row: 1, col: 2 })).toEqual({ x: 5, y: 20, width: 5, height: 20 })
    expect(sheet.sourceCellToLogical({ row: 2, col: 0 })).toBeNull()
    expect(sheet.sourceCellToLogical({ row: 0, col: 1 })).toBeNull()
    expect(sheet.logicalPointToSource(5, 17)).toEqual({ row: 3, col: 2 })
    expect(sheet.logicalPointToSource(5, 30)).toEqual({ row: 1, col: 2 })
    expect(sheet.logicalPointToSource(-1, 0)).toBeNull()
    expect(sheet.logicalPointToSource(10, 0)).toBeNull()
    expect(sheet.logicalPointToSource(0, 50)).toBeNull()
    expect(sheet.visibleRankToDisplay(3)).toBe(4)
    expect(sheet.sourceToVisibleRank(2)).toBeNull()
    expect(sheet.rows.totalCss()).toBe(50)
    expect(sheet.diagnostics.densePrefixBytes).toBeGreaterThan(0)
    expect(sheet.rows.overrides.map(interval => [interval.start, interval.end, interval.sizeCss])).toEqual([[2, 2, 20], [3, 3, 0]])
  })
  it('uses visibility for all geometry and rank mappings, with no phantom hidden hit target', () => {
    const sheet = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 5, defaultSizeCss: 10 }), cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 5 }), visibility: createRowVisibility(5, slot => slot % 2 === 0) })
    expect(sheet.sourceCellToLogical({ row: 1, col: 0 })).toBeNull()
    expect(sheet.sourceCellToLogical({ row: 2, col: 0 })?.y).toBe(10)
    expect(sheet.logicalPointToSource(0, 10)).toEqual({ row: 2, col: 0 })
    expect(sheet.displayToVisibleRank(2)).toBe(1)
    expect(sheet.visibleRankToDisplay(1)).toBe(2)
    expect(sheet.visibleRankToDisplay(3)).toBeNull()
    expect(sheet.rows.totalCss()).toBe(30)
  })
  it('returns null/empty analytically for completely collapsed axes', () => {
    const sheet = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 0 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 0 }) })
    expect(sheet.sourceCellToLogical({ row: 16385, col: 4096 })).toBeNull()
    expect(sheet.logicalPointToSource(0, 0)).toBeNull()
    expect(sheet.visibleRankToSource(0)).toBeNull()
    expect(Array.from(sheet.rows.visibleBetween(0, 1).runs)).toEqual([])
    expect(sheet.diagnostics.lastWindowVisitedSlots).toBe(0)
    expect(sheet.diagnostics.densePrefixBytes).toBe(0)
  })
  it('bounds projected windows and retains Float64 precision in sparse identity spans', () => {
    const sheet = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 0.001 }), cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 64 }) })
    const window = sheet.rows.visibleBetween(0, 2000)
    expect(window.count).toBe(MAX_ROWS)
    expect(Array.from(window.runs)).toHaveLength(1)
    expect(sheet.diagnostics.lastWindowVisitedSlots).toBeLessThanOrEqual(256)
    expect(sheet.logicalPointToSource(0, sheet.rows.start(MAX_ROWS - 1))).toEqual({ row: MAX_ROWS - 1, col: 0 })
  })
})

describe('projection precision, cursor and allocation regressions', () => {
  it('rejects a reordered tiny positive height that cannot be represented after its new prefix', () => {
    const rows = createAxisMetrics({ logicalCount: 2, defaultSizeCss: 25165824, overrides: [{ start: 0, end: 0, sizeCss: 1e-10 }] })
    const cols = createAxisMetrics({ logicalCount: 1, defaultSizeCss: 64 })
    const order = createRowOrder([{ r0: 0, r1: 1, sourceAtOffset: new Uint32Array([1, 0]) }])
    expect(() => createSheetMetrics({ rows, cols, order })).toThrow(/represent|precision/i)
  })
  it('rejects malformed projection extents and out-of-domain cell mappings', () => {
    const rows = createAxisMetrics({ logicalCount: 4, defaultSizeCss: 10 })
    const cols = createAxisMetrics({ logicalCount: 2, defaultSizeCss: 64 })
    expect(() => createSheetMetrics({ rows, cols: createAxisMetrics({ logicalCount: MAX_COLS + 1, defaultSizeCss: 1 }) })).toThrow()
    expect(() => createSheetMetrics({ rows, cols, visibility: createRowVisibility(3) })).toThrow()
    expect(() => createSheetMetrics({ rows, cols, order: createRowOrder([{ r0: 2, r1: 4, sourceAtOffset: new Uint32Array([4, 3, 2]) }]) })).toThrow()
    const sheet = createSheetMetrics({ rows, cols })
    expect(() => sheet.sourceCellToLogical({ row: 4, col: 0 })).toThrow()
    expect(() => sheet.sourceCellToLogical({ row: 0, col: 2 })).toThrow()
  })
  it('matches a seeded dense oracle for moved variable heights, masks and partial bands', () => {
    for (let seed = 1; seed <= 20; seed++) {
      const count = 71
      const heights = Array.from({ length: count }, (_, source) => [0, 0.25, 1, 5][(source * 13 + seed) % 4])
      const rows = createAxisMetrics({ logicalCount: count, defaultSizeCss: 1, overrides: heights.map((height, source) => ({ start: source, end: source, sizeCss: height })) })
      const sourceAtOffset = shuffled(57, seed)
      for (let i = 0; i < sourceAtOffset.length; i++) sourceAtOffset[i] += 7
      const order = createRowOrder([{ r0: 7, r1: 63, sourceAtOffset }])
      const visibility = createRowVisibility(count, slot => (slot * 7 + seed) % 11 > 2)
      const sheet = createSheetMetrics({ rows, cols: createAxisMetrics({ logicalCount: 3, defaultSizeCss: 64 }), order, visibility })
      let css = 0, ordinal = 0
      for (let displaySlot = 0; displaySlot < count; displaySlot++) {
        const source = order.displayToSource(displaySlot)
        const height = visibility.visible(displaySlot) ? heights[source] : 0
        expect(sheet.rows.start(displaySlot)).toBe(css)
        expect(sheet.rows.size(displaySlot)).toBe(height)
        const box = sheet.sourceCellToLogical({ row: source, col: 2 })
        if (height) {
          expect(box).toEqual({ x: 128, y: css, width: 64, height })
          expect(sheet.logicalPointToSource(129, css)).toEqual({ row: source, col: 2 })
          expect(sheet.logicalPointToSource(129, css + height / 2)).toEqual({ row: source, col: 2 })
          expect(sheet.displayToVisibleRank(displaySlot)).toBe(ordinal)
          expect(sheet.visibleRankToSource(ordinal++)).toBe(source)
        } else {
          expect(box).toBeNull()
          expect(sheet.sourceToVisibleRank(source)).toBeNull()
        }
        css += height
      }
      expect(sheet.rows.totalCss()).toBe(css)
      expect(sheet.visibleRankToDisplay(ordinal)).toBeNull()
    }
  })
  it('pages fragmented display windows without omissions or stale cursor ownership', () => {
    const rows = createAxisMetrics({ logicalCount: 900, defaultSizeCss: 1 })
    const cols = createAxisMetrics({ logicalCount: 1, defaultSizeCss: 64 })
    const sheet = createSheetMetrics({ rows, cols, visibility: createRowVisibility(900, slot => slot % 2 === 0) })
    const visited: number[] = []
    let page = sheet.rows.visiblePageBetween(0, 450)
    const cursor = page.nextCursor!
    expect(cursor).toBeDefined()
    for (;;) {
      for (const run of page.runs) for (let row = run.start; row <= run.end; row++) visited.push(row)
      expect(sheet.diagnostics.lastWindowVisitedSlots).toBeLessThanOrEqual(256)
      if (!page.nextCursor) break
      page = sheet.rows.continueVisibleBetween(0, 450, page.nextCursor)
    }
    expect(visited).toEqual(Array.from({ length: 450 }, (_, i) => i * 2))
    expect(() => sheet.rows.continueVisibleBetween(0, 449, cursor)).toThrow()
    expect(() => createSheetMetrics({ rows, cols }).rows.continueVisibleBetween(0, 450, cursor)).toThrow()
    expect(() => sheet.rows.continueVisibleBetween(0, 450, 'bad')).toThrow()
  })
  it('accounts for full variable-band workspace and lazy fragmented metadata separately', () => {
    const half = MAX_ROWS / 2
    const values = Uint32Array.from({ length: MAX_ROWS }, (_, slot) => slot % 2 ? slot / 2 - 0.5 : half + slot / 2)
    const order = createRowOrder([{ r0: 0, r1: MAX_ROWS - 1, sourceAtOffset: values }])
    const visibility = createRowVisibility()
    const sourceRows = createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24, overrides: [{ start: half, end: MAX_ROWS - 1, sizeCss: 48 }] })
    const sheet = createSheetMetrics({ rows: sourceRows, cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order, visibility })
    const maskBytes = visibility.words.byteLength + visibility.summaries.byteLength
    const before = sheet.diagnostics.typedArrayBytes + order.diagnostics.typedArrayBytes + maskBytes
    expect(sheet.diagnostics.densePrefixBytes).toBe((MAX_ROWS + 1) * 8)
    expect(before).toBeLessThan(30 * 1024 * 1024)
    const page = sheet.rows.visiblePageBetween(0, sheet.rows.totalCss())
    expect(page.count).toBe(256)
    expect(page.nextCursor).toBeDefined()
    expect(sheet.diagnostics.lastWindowVisitedSlots).toBe(256)
    expect(sheet.rows.overrides.length).toBe(half)
    const after = sheet.diagnostics.typedArrayBytes + order.diagnostics.typedArrayBytes + maskBytes
    expect(after - before).toBe(half * 8)
    expect(after).toBeLessThan(34 * 1024 * 1024)
    console.info('S2.1 full projection ledger', { beforeBytes: before, afterBytes: after, callerInputBytes: values.byteLength })
  })
})

it('keeps lazy display override arrays enumerable and serializable like the frozen array contract', () => {
  const sheet = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 4, defaultSizeCss: 24, overrides: [{ start: 1, end: 1, sizeCss: 48 }] }), cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 64 }) })
  expect(sheet.rows.overrides).toEqual([{ start: 1, end: 1, sizeCss: 48 }])
  expect(Object.keys(sheet.rows.overrides)).toEqual(['0'])
  expect(JSON.parse(JSON.stringify(sheet.rows.overrides))).toEqual([{ start: 1, end: 1, sizeCss: 48 }])
  expect(() => { (sheet.rows.overrides as { start: number; end: number; sizeCss: number }[])[0] = { start: 0, end: 1, sizeCss: 1 } }).toThrow()
})

describe('contract-typed lazy projected windows (review item 1)', () => {
  it('a SheetMetrics consumer sees all 450 runs of 1,000 alternating hidden/visible display rows', () => {
    const sheet: SheetMetrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 1000, defaultSizeCss: 1 }),
      cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 64 }), visibility: createRowVisibility(1000, slot => slot % 2 === 0) })
    const window = sheet.rows.visibleBetween(0, 450)
    const rows: number[] = []
    for (const run of window.runs) for (let row = run.start; row <= run.end; row++) rows.push(row)
    expect(rows).toEqual(Array.from({ length: 450 }, (_, i) => i * 2))
    expect(window.count).toBe(450)
    expect(window.nextCursor).toBeUndefined()
  })
  it('iterates a 1M-row dense variable band completely, exactly once, and twice', () => {
    const half = MAX_ROWS / 2
    const values = Uint32Array.from({ length: MAX_ROWS }, (_, slot) => slot % 2 ? slot / 2 - 0.5 : half + slot / 2)
    const order = createRowOrder([{ r0: 0, r1: MAX_ROWS - 1, sourceAtOffset: values }])
    const concrete = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24, overrides: [{ start: half, end: MAX_ROWS - 1, sizeCss: 48 }] }),
      cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 64 }), order })
    const sheet: SheetMetrics = concrete
    const window = sheet.rows.visibleBetween(0, sheet.rows.totalCss())
    expect(window.count).toBe(MAX_ROWS)
    for (let pass = 0; pass < 2; pass++) {
      let next = 0, css = 0, exact = true, peak = 0
      for (const run of window.runs) {
        if (run.start !== next || run.startCss !== css) exact = false
        css += (run.end - run.start + 1) * run.sizeCss
        next = run.end + 1
        peak = Math.max(peak, concrete.diagnostics.lastWindowVisitedSlots)
      }
      expect(exact).toBe(true)
      expect(next).toBe(MAX_ROWS)
      expect(css).toBe(sheet.rows.totalCss())
      expect(peak).toBeLessThanOrEqual(256)
    }
  }, 60_000)
  it('reports count equal to covered rows for randomized masked, reordered windows', () => {
    for (let seed = 1; seed <= 15; seed++) {
      const count = 83
      const heights = Array.from({ length: count }, (_, source) => [0, 0.5, 1, 3][(source * 7 + seed) % 4])
      const order = createRowOrder([{ r0: 5, r1: 70, sourceAtOffset: Uint32Array.from(shuffled(66, seed), row => row + 5) }])
      const sheet = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: count, defaultSizeCss: 1, overrides: heights.map((height, row) => ({ start: row, end: row, sizeCss: height })) }),
        cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 64 }), order, visibility: createRowVisibility(count, slot => (slot * 5 + seed) % 9 > 1) })
      for (let query = 0; query < 25; query++) {
        const lower = ((query * 37 + seed * 11) % 140) / 2 - 2, upper = lower + ((query * 53 + seed) % 60) / 2
        const rows: number[] = []
        for (const run of sheet.rows.visibleBetween(lower, upper).runs) for (let row = run.start; row <= run.end; row++) rows.push(row)
        const expected = Array.from({ length: count }, (_, slot) => slot).filter(slot => {
          const size = sheet.rows.size(slot), start = sheet.rows.start(slot)
          return size > 0 && lower < upper && start < upper && start + size > lower
        })
        expect(rows).toEqual(expected)
        expect(sheet.rows.visibleBetween(lower, upper).count).toBe(expected.length)
      }
    }
  })
})

describe('typed projection construction errors (review item 4)', () => {
  const cols = createAxisMetrics({ logicalCount: 1, defaultSizeCss: 64 })
  const codeOf = (build: () => unknown): string | undefined => {
    try { build() } catch (error) {
      expect(error).toBeInstanceOf(AxisMetricsError)
      expect(error).toBeInstanceOf(RangeError)
      return (error as AxisMetricsError).code
    }
    return undefined
  }
  it('reports PRECISION_EXCEEDED when a reorder pushes a tiny cell past Float64 resolution', () => {
    const rows = createAxisMetrics({ logicalCount: 2, defaultSizeCss: 25165824, overrides: [{ start: 0, end: 0, sizeCss: 1e-10 }] })
    const order = createRowOrder([{ r0: 0, r1: 1, sourceAtOffset: new Uint32Array([1, 0]) }])
    expect(codeOf(() => createSheetMetrics({ rows, cols, order }))).toBe('PRECISION_EXCEEDED')
  })
  it('reports TOTAL_NOT_FINITE when the display total overflows Float64', () => {
    const huge: AxisMetrics = { logicalCount: 2, defaultSizeCss: Number.MAX_VALUE, overrides: [], totalCss: () => Infinity,
      start: () => 0, size: () => Number.MAX_VALUE, locate: () => null, visibleBetween: () => ({ runs: [], count: 0 }) }
    expect(codeOf(() => createSheetMetrics({ rows: huge, cols }))).toBe('TOTAL_NOT_FINITE')
  })
})
