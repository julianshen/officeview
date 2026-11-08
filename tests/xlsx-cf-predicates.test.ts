import { describe, expect, it } from 'vitest'
import { parseXml } from '../src/core/xml'
import type {
  CfDefinition,
  CfLiteral,
  CfRule,
  CfStamp,
  SourceRect,
} from '../src/xlsx/cf/model'
import type { XlsxCell } from '../src/xlsx/types'
import { createCfMemory, createCfWorkBudget, type CfWorkBudget } from '../src/xlsx/cf/limits'
import { parseCfSheet } from '../src/xlsx/cf/parse'
import { createCfStatistics, createResidentCfSource, type CfRangeStatistics } from '../src/xlsx/cf/statistics'
import { cfRuntimeEvidence, evaluatePredicate, type CfEvidenceGate } from '../src/xlsx/cf/predicates'

const SHEET = 'sheet1'
const stamp = (gen = 1): CfStamp => ({
  owner: 't5', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1,
})
const rect = (firstRow: number, lastRow: number, firstCol = 0, lastCol = 0): SourceRect => ({ firstRow, lastRow, firstCol, lastCol })
const frozenBudget = (): CfWorkBudget => createCfWorkBudget({ now: () => 0 })
// The production gate (backed by the reviewed evidence allowlist, currently
// empty) must deny every measurement-dependent variant below.
const denyGate: CfEvidenceGate = cfRuntimeEvidence
const allowGate: CfEvidenceGate = { allows: () => true }

const num = (value: number): CfLiteral => ({ kind: 'number', value })
const str = (value: string): CfLiteral => ({ kind: 'string', value })
const bool = (value: boolean): CfLiteral => ({ kind: 'boolean', value })
const blank: CfLiteral = { kind: 'blank' }
const err = (value: string): CfLiteral => ({ kind: 'error', value })

function rule<D extends object>(definition: D, ranges: readonly SourceRect[] = [rect(0, 0)]): CfRule {
  return {
    id: 0, priority: 1, documentOrder: 0, ranges, stopIfTrue: false,
    writes: new Set(['ALL']),
    definition: { dialect: 'base', formulas: [], ...definition } as unknown as CfDefinition,
  }
}
const cellIs = (operator: 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual' | 'between' | 'notBetween', operands: CfLiteral[]): CfRule =>
  rule({ kind: 'cellIs', operator, operands })

/** Parse one rule's definition the way production does (anchor B2, priority 1). */
function parsedDefinition(inner: string, sqref = 'B2'): CfDefinition {
  const sheet = parseXml(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="${sqref}">${inner}</conditionalFormatting></worksheet>`,
  )
  const lease = parseCfSheet(sheet, { colors: {}, dxfNodes: [] }, createCfMemory({}), {})
  try {
    expect(lease.value.rules).toHaveLength(1)
    return lease.value.rules[0]!.definition
  } finally {
    lease.release()
  }
}

const colName = (c: number): string => {
  let n = c + 1, s = ''
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26) }
  return s
}
const xcell = (row: number, col: number, value: XlsxCell['value']): XlsxCell => ({
  ref: `${colName(col)}${row + 1}`, row, col, value, styleIndex: 0,
})

/** Read every value of a scope through a chunked source (paged-shaped delivery). */
async function readChunked(cells: XlsxCell[], scope: SourceRect): Promise<CfLiteral[]> {
  const out: CfLiteral[] = []
  const memory = createCfMemory({})
  const full = createResidentCfSource(cells, SHEET, stamp(), memory)
  const budget = frozenBudget()
  const signal = new AbortController().signal
  // Re-chunk the resident pages into 3-cell pages to prove chunk independence.
  for await (const page of full.pages(SHEET, [scope], budget, signal)) {
    const values = page.value.values
    for (let s = 0; s < values.length; s += 3) out.push(...values.slice(s, s + 3))
    page.release()
  }
  expect(memory.currentBytes).toBe(0)
  return out
}

describe('CF predicate evaluation', () => {
  it('literal threshold 100 separates 99 100 101 (manifest stage1-cellis-literals)', () => {
    const eq = cellIs('equal', [num(100)])
    const at = { row: 0, col: 0 }
    // B5 (99) does not match, B6 (100) matches; 101 behaves like 99.
    expect(evaluatePredicate(eq, at, num(99), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(eq, at, num(100), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(eq, at, num(101), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
  })

  it('cellIs relational matrix matches manifest coordinates', () => {
    const at = { row: 0, col: 0 }
    const cases: Array<[ReturnType<typeof cellIs>, CfLiteral, 'ready' | 'no-match']> = [
      // D: greaterThan 100. D6 (100) false, D7 (101) true.
      [cellIs('greaterThan', [num(100)]), num(100), 'no-match'],
      [cellIs('greaterThan', [num(100)]), num(101), 'ready'],
      // E: lessThan 100. E2 (-1) true, E6 (100) false.
      [cellIs('lessThan', [num(100)]), num(-1), 'ready'],
      [cellIs('lessThan', [num(100)]), num(100), 'no-match'],
      // F: greaterThanOrEqual 100. F5 (99) false, F6 (100) true.
      [cellIs('greaterThanOrEqual', [num(100)]), num(99), 'no-match'],
      [cellIs('greaterThanOrEqual', [num(100)]), num(100), 'ready'],
      // G: lessThanOrEqual 100. G6 (100) true, G7 (101) false.
      [cellIs('lessThanOrEqual', [num(100)]), num(100), 'ready'],
      [cellIs('lessThanOrEqual', [num(100)]), num(101), 'no-match'],
      // H: between 1..3. H2 (-1) false, H4 (1) true.
      [cellIs('between', [num(1), num(3)]), num(-1), 'no-match'],
      [cellIs('between', [num(1), num(3)]), num(1), 'ready'],
      // J: notBetween 1..3. J2 (-1) true, J4 (1) false.
      [cellIs('notBetween', [num(1), num(3)]), num(-1), 'ready'],
      [cellIs('notBetween', [num(1), num(3)]), num(1), 'no-match'],
      // C: notEqual 100. C2 (-1) true, C6 (100) false.
      [cellIs('notEqual', [num(100)]), num(-1), 'ready'],
      [cellIs('notEqual', [num(100)]), num(100), 'no-match'],
      // L: equal string. L8 ("100") true, L6 numeric 100 vs "100" false.
      [cellIs('equal', [str('100')]), str('100'), 'ready'],
      [cellIs('equal', [str('100')]), num(100), 'no-match'],
      // M: equal boolean. M9 (TRUE) true, M10 (FALSE vs TRUE) false.
      [cellIs('equal', [bool(true)]), bool(true), 'ready'],
      [cellIs('equal', [bool(true)]), bool(false), 'no-match'],
      // O: empty string. O12 ("" == "") true, O6 (100 vs "") false.
      [cellIs('equal', [str('')]), str(''), 'ready'],
      [cellIs('equal', [str('')]), num(100), 'no-match'],
      // N6: numeric 100 vs error operand is unrepresentable; error cell vs number is false.
      [cellIs('equal', [num(100)]), err('#N/A'), 'no-match'],
    ]
    for (const [r, value, expected] of cases) {
      const outcome = evaluatePredicate(r, at, value, undefined, denyGate)
      expect(outcome.kind, JSON.stringify({ def: r.definition, value })).toBe(expected)
      if (expected === 'ready') expect(outcome).toMatchObject({ value: true })
    }
  })

  it('reversed between operands are gated on operand-normalization evidence', () => {
    // Manifest I2/I4/K2/K4 assume [min,max] normalisation, which is unmeasured:
    // the production gate withholds, the allow gate reproduces the coordinates.
    const at = { row: 0, col: 0 }
    const betweenRev = cellIs('between', [num(3), num(1)])
    const notBetweenRev = cellIs('notBetween', [num(3), num(1)])
    for (const r of [betweenRev, notBetweenRev]) {
      expect(evaluatePredicate(r, at, num(-1), undefined, denyGate).kind).toBe('unresolved')
      expect(evaluatePredicate(r, at, num(1), undefined, denyGate).kind).toBe('unresolved')
    }
    expect(evaluatePredicate(betweenRev, at, num(-1), undefined, allowGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(betweenRev, at, num(1), undefined, allowGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(notBetweenRev, at, num(-1), undefined, allowGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(notBetweenRev, at, num(1), undefined, allowGate)).toMatchObject({ kind: 'no-match' })
  })

  it('cellIs missing combinations stay unresolved instead of guessing', () => {
    const at = { row: 0, col: 0 }
    const unresolved: Array<[ReturnType<typeof cellIs>, CfLiteral]> = [
      // Mixed-type relational operators are unmeasured.
      [cellIs('greaterThan', [str('100')]), num(100)],
      [cellIs('greaterThan', [num(100)]), str('100')],
      [cellIs('lessThan', [num(1)]), bool(true)],
      [cellIs('greaterThanOrEqual', [num(100)]), bool(false)],
      // Blank cells are unmeasured for every operator.
      [cellIs('equal', [num(100)]), blank],
      [cellIs('equal', [num(0)]), blank],
      [cellIs('equal', [str('')]), blank],
      [cellIs('lessThan', [num(100)]), blank],
      [cellIs('notBetween', [num(1), num(3)]), blank],
      // Non-numeric between operands and cells.
      [cellIs('between', [str('a'), str('z')]), str('m')],
      [cellIs('between', [num(1), num(3)]), str('2')],
      [cellIs('between', [num(1), num(3)]), bool(true)],
      [cellIs('between', [num(1), num(3)]), blank],
      // notEqual is only decided for number pairs (C2/C6); everything else is missing.
      [cellIs('notEqual', [str('100')]), str('100')],
      [cellIs('notEqual', [num(100)]), str('100')],
      // Cross-type equal pairs without manifest rows.
      [cellIs('equal', [num(100)]), bool(true)],
      [cellIs('equal', [str('100')]), bool(true)],
      [cellIs('equal', [bool(true)]), num(1)],
      [cellIs('equal', [str('100')]), blank],
    ]
    for (const [r, value] of unresolved) {
      expect(evaluatePredicate(r, at, value, undefined, denyGate).kind, JSON.stringify({ def: r.definition, value })).toBe('unresolved')
    }
  })

  it('nonliteral B1 stays unresolved on both adapters', async () => {
    const definition = parsedDefinition('<cfRule type="cellIs" operator="greaterThan" priority="1"><formula>B1</formula></cfRule>')
    expect(definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    const r = rule(definition, [rect(0, 1, 1, 1)])
    // Resident-shaped delivery: one page per rect.
    const residentCells = [xcell(0, 1, 5), xcell(1, 1, 150)]
    const residentValues = await readChunked(residentCells, rect(0, 1, 1, 1))
    expect(residentValues).toHaveLength(2)
    // Chunked delivery of the same cells.
    const chunked = await readChunked(residentCells, rect(0, 1, 1, 1))
    expect(chunked).toHaveLength(2)
    for (const value of [...residentValues, ...chunked]) {
      expect(evaluatePredicate(r, { row: 1, col: 1 }, value, undefined, denyGate).kind).toBe('unresolved')
    }
  })

  it('canonical blank formula is gated by its evidence variant', () => {
    // Manifest stage1-blanks-errors: LEN(TRIM(B2))=0 matches blanks and "" but not 0.
    // Without reviewed evidence the evaluation must not run.
    const canonical = parsedDefinition('<cfRule type="containsBlanks" priority="1"><formula>LEN(TRIM(B2))=0</formula></cfRule>', 'B2')
    expect(canonical).toMatchObject({ kind: 'predicate', type: 'containsBlanks' })
    const r = rule(canonical, [rect(1, 1, 1, 1)])
    for (const value of [blank, str(''), num(0)] as CfLiteral[]) {
      const gated = evaluatePredicate(r, { row: 1, col: 1 }, value, undefined, denyGate)
      expect(gated).toMatchObject({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
    }
    // The allow-all gate proves the underlying semantics match the manifest once measured.
    expect(evaluatePredicate(r, { row: 1, col: 1 }, blank, undefined, allowGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(r, { row: 1, col: 1 }, str(''), undefined, allowGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(r, { row: 1, col: 1 }, num(0), undefined, allowGate)).toMatchObject({ kind: 'no-match' })
    // Ruling C3: Absent-formula blanks refuse as cf-unsupported-rule (Excel paints nothing on absent-formula rules).
    const absent = parsedDefinition('<cfRule type="containsBlanks" priority="1"/>', 'B2')
    expect(absent).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    const absentNot = parsedDefinition('<cfRule type="notContainsBlanks" priority="1"/>', 'B2')
    expect(absentNot).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
  })

  it('case-insensitive text matching is gated on its evidence variant', () => {
    // Manifest B3/C4-adjacent behaviour needs case folding, which is unmeasured:
    // exact matches stay decided, fold-only matches withhold without evidence.
    const at = { row: 0, col: 0 }
    const contains = rule({ kind: 'text', type: 'containsText', needle: 'ab' })
    expect(evaluatePredicate(contains, at, str('AbC'), undefined, allowGate)).toMatchObject({ kind: 'ready', value: true })
    const begins = rule({ kind: 'text', type: 'beginsWith', needle: 'ab' })
    expect(evaluatePredicate(begins, at, str('AbC'), undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(begins, at, str('AbC'), undefined, allowGate)).toMatchObject({ kind: 'ready', value: true })
    const ends = rule({ kind: 'text', type: 'endsWith', needle: 'BC' })
    expect(evaluatePredicate(ends, at, str('abc'), undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(ends, at, str('abc'), undefined, allowGate)).toMatchObject({ kind: 'ready', value: true })
    const notContains = rule({ kind: 'text', type: 'notContainsText', needle: 'ab' })
    expect(evaluatePredicate(notContains, at, str('AbC'), undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(notContains, at, str('AbC'), undefined, allowGate)).toMatchObject({ kind: 'no-match' })
  })

  it('text predicates match manifest stage1-text-predicates coordinates', () => {
    const at = { row: 0, col: 0 }
    const contains = rule({ kind: 'text', type: 'containsText', needle: 'ab' })
    expect(evaluatePredicate(contains, at, str('abc'), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    // B3 ("AbC") needs case-insensitive matching: gated below, unresolved here.
    expect(evaluatePredicate(contains, at, str('AbC'), undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(contains, at, str('ábc'), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    const notContains = rule({ kind: 'text', type: 'notContainsText', needle: 'ab' })
    expect(evaluatePredicate(notContains, at, str('abc'), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(notContains, at, str('ábc'), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    const begins = rule({ kind: 'text', type: 'beginsWith', needle: 'ab' })
    expect(evaluatePredicate(begins, at, str('abc'), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(begins, at, str(' ABC '), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    const ends = rule({ kind: 'text', type: 'endsWith', needle: 'bc' })
    expect(evaluatePredicate(ends, at, str('abc'), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(ends, at, str('a"b'), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    // H6: numeric cells evaluate on the raw value, not a formatted display string.
    const begins00 = rule({ kind: 'text', type: 'beginsWith', needle: '00' })
    expect(evaluatePredicate(begins00, at, num(123), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    // Whitespace and escaped-quote needles are significant.
    const ws = rule({ kind: 'text', type: 'containsText', needle: ' ABC ' })
    expect(evaluatePredicate(ws, at, str('abc'), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(ws, at, str(' ABC '), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    const quote = rule({ kind: 'text', type: 'containsText', needle: 'a"b' })
    expect(evaluatePredicate(quote, at, str('abc'), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(quote, at, str('a"b'), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    // Divergent formulas never reach evaluation: the parser already refused them.
    const divergent = parsedDefinition('<cfRule type="containsText" text="ab" priority="1"><formula>NOT(ISERROR(SEARCH("zz",J2)))</formula></cfRule>', 'J2')
    expect(divergent).toMatchObject({ kind: 'unresolved' })
    expect(evaluatePredicate(rule(divergent), at, str('abc'), undefined, denyGate).kind).toBe('unresolved')
    // Empty needles and non-string cells are missing combinations.
    expect(evaluatePredicate(rule({ kind: 'text', type: 'containsText', needle: '' }), at, str('abc'), undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(contains, at, bool(true), undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(contains, at, err('#N/A'), undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(contains, at, blank, undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(notContains, at, num(123), undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(notContains, at, blank, undefined, denyGate).kind).toBe('unresolved')
  })

  it('error predicates match manifest stage1-blanks-errors coordinates', () => {
    const at = { row: 0, col: 0 }
    const isErr = rule({ kind: 'predicate', type: 'containsErrors' })
    expect(evaluatePredicate(isErr, at, num(0), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(isErr, at, err('#N/A'), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(isErr, at, err('#DIV/0!'), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    const notErr = rule({ kind: 'predicate', type: 'notContainsErrors' })
    expect(evaluatePredicate(notErr, at, num(0), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(notErr, at, err('#N/A'), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    // Non-error values are definitionally decided; only blank cells are unmeasured.
    for (const value of [str('abc'), bool(true)] as CfLiteral[]) {
      expect(evaluatePredicate(isErr, at, value, undefined, denyGate)).toMatchObject({ kind: 'no-match' })
      expect(evaluatePredicate(notErr, at, value, undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
    }
    expect(evaluatePredicate(isErr, at, blank, undefined, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(notErr, at, blank, undefined, denyGate).kind).toBe('unresolved')
  })

  it('evaluates all ten timePeriod operators against clock context', () => {
    const at = { row: 0, col: 0 }
    const D = 46332 // 2026-11-06 (Friday)
    const ctx = { clockSerial: D, dateSystem: '1900' as const }

    const periods: Array<{ period: import('../src/xlsx/cf/model').CfTimePeriod; pos: number[]; neg: number[] }> = [
      { period: 'today', pos: [D, D + 0.5, D + 0.999], neg: [D - 1, D + 1] },
      { period: 'yesterday', pos: [D - 1, D - 1 + 0.25], neg: [D, D - 2] },
      { period: 'tomorrow', pos: [D + 1, D + 1 + 0.75], neg: [D, D + 2] },
      { period: 'last7Days', pos: [D, D - 6, D - 3], neg: [D - 7, D + 1] },
      { period: 'thisWeek', pos: [D, 46327, 46333], neg: [46326, 46334] },
      { period: 'lastWeek', pos: [D - 7, 46320, 46326], neg: [46319, 46327] },
      { period: 'nextWeek', pos: [D + 7, 46334, 46340], neg: [46333, 46341] },
      { period: 'thisMonth', pos: [D, 46330, 46355], neg: [46301, 46362] },
      { period: 'lastMonth', pos: [46301, 46326], neg: [D, 46362] },
      { period: 'nextMonth', pos: [46362, 46380], neg: [D, 46301] },
    ]

    for (const { period, pos, neg } of periods) {
      const r = rule({ kind: 'timePeriod', period })
      for (const val of pos) {
        const out = evaluatePredicate(r, at, num(val), undefined, allowGate, ctx)
        expect(out, `${period} on ${val}`).toEqual({ kind: 'ready', value: true })
      }
      for (const val of neg) {
        const out = evaluatePredicate(r, at, num(val), undefined, allowGate, ctx)
        expect(out, `${period} on ${val}`).toEqual({ kind: 'no-match' })
      }
      // Non-numbers fail match
      for (const val of [str('2026-11-06'), blank, err('#VALUE!')] as CfLiteral[]) {
        const out = evaluatePredicate(r, at, val, undefined, allowGate, ctx)
        expect(out).toEqual({ kind: 'no-match' })
      }
    }

    // 1904 date system variant
    const D1904 = 44870
    const ctx1904 = { clockSerial: D1904, dateSystem: '1904' as const }
    const today1904 = rule({ kind: 'timePeriod', period: 'today' })
    expect(evaluatePredicate(today1904, at, num(D1904), undefined, allowGate, ctx1904)).toEqual({ kind: 'ready', value: true })
    expect(evaluatePredicate(today1904, at, num(D1904 - 1), undefined, allowGate, ctx1904)).toEqual({ kind: 'no-match' })

    // Denied gate refuses
    const denyTimeGate: CfEvidenceGate = { allows: () => false }
    expect(evaluatePredicate(today1904, at, num(D1904), undefined, denyTimeGate, ctx1904)).toEqual({
      kind: 'unresolved',
      reason: 'cf-unmeasured-capability',
    })
  })

  it('top10 ranks the manifest set with ties included', () => {
    // Manifest stage1-top10-rank-edge values {1,2,2,3,9} plus "999", TRUE, #N/A, "".
    const sorted = [1, 2, 2, 3, 9]
    const stats: CfRangeStatistics = {
      scopeKey: 't', numericCount: 5, sum: 17, mean: 3.4, min: 1, max: 9, stddevPop: undefined,
      sortedAsc: Float64Array.from(sorted), frequencies: new Map(), declaredCells: 9, spanCells: 9, storedCells: 9,
    }
    const at = { row: 0, col: 0 }
    const top2 = rule({ kind: 'top10', rank: 2, bottom: false, percent: false })
    expect(evaluatePredicate(top2, at, num(3), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(top2, at, num(9), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(top2, at, num(2), stats, denyGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(top2, at, str('999'), stats, denyGate)).toMatchObject({ kind: 'no-match' })
    const bottom2 = rule({ kind: 'top10', rank: 2, bottom: true, percent: false })
    expect(evaluatePredicate(bottom2, at, num(1), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(bottom2, at, num(2), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(bottom2, at, num(9), stats, denyGate)).toMatchObject({ kind: 'no-match' })
    // Rank 99 selects every numeric cell; percent rounds up per the manifest recipe.
    const top99 = rule({ kind: 'top10', rank: 99, bottom: false, percent: false })
    expect(evaluatePredicate(top99, at, num(1), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    const pct1 = rule({ kind: 'top10', rank: 1, bottom: false, percent: true })
    expect(evaluatePredicate(pct1, at, num(9), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(pct1, at, num(1), stats, denyGate)).toMatchObject({ kind: 'no-match' })
    const pct40 = rule({ kind: 'top10', rank: 40, bottom: false, percent: true })
    expect(evaluatePredicate(pct40, at, num(3), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(pct40, at, num(9), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    // Excel oracle: ECMA-376 specifies default rank="10". Both omitted rank and rank="0"
    // resolve to rank=10, matching Excel 16.114 on stage1-top10-rank-edge.
    const noRank = parsedDefinition('<cfRule type="top10" priority="1"/>')
    const zeroRank = parsedDefinition('<cfRule type="top10" priority="1" rank="0"/>')
    expect(noRank).toMatchObject({ kind: 'top10', rank: 10, bottom: false, percent: false })
    expect(zeroRank).toMatchObject({ kind: 'top10', rank: 10, bottom: false, percent: false })
    expect(evaluatePredicate(rule(noRank), at, num(1), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(rule(zeroRank), at, num(9), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })

    const zeroRankPercent = parsedDefinition('<cfRule type="top10" priority="1" percent="1" rank="0"/>')
    expect(zeroRankPercent).toMatchObject({ kind: 'top10', rank: 10, bottom: false, percent: true })
    // Top 10% of 5 numbers {1,2,2,3,9} = ceil(5*10/100) = 1 cell: matches 9 (F6), but not 1, 2, or 3 (F2..F5).
    expect(evaluatePredicate(rule(zeroRankPercent), at, num(9), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(rule(zeroRankPercent), at, num(3), stats, denyGate)).toMatchObject({ kind: 'no-match' })
  })

  it('aboveAverage follows the manifest mean without stddev', () => {
    // Manifest stage1-above-average-stddev: {0,2,4}, mean 2. Plain
    // comparisons are settled; stddev selection is gated separately below.
    const stats: CfRangeStatistics = {
      scopeKey: 't', numericCount: 3, sum: 6, mean: 2, min: 0, max: 4, stddevPop: Math.sqrt(8 / 3),
      sortedAsc: Float64Array.from([0, 2, 4]), frequencies: new Map(), declaredCells: 3, spanCells: 3, storedCells: 3,
    }
    const at = { row: 0, col: 0 }
    const above = rule({ kind: 'aboveAverage', above: true, equal: false })
    expect(evaluatePredicate(above, at, num(2), stats, denyGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(above, at, num(4), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    const below = rule({ kind: 'aboveAverage', above: false, equal: false })
    expect(evaluatePredicate(below, at, num(0), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(below, at, num(2), stats, denyGate)).toMatchObject({ kind: 'no-match' })
    const eqAbove = rule({ kind: 'aboveAverage', above: true, equal: true })
    expect(evaluatePredicate(eqAbove, at, num(2), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    const eqBelow = rule({ kind: 'aboveAverage', above: false, equal: true })
    expect(evaluatePredicate(eqBelow, at, num(2), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(eqBelow, at, num(4), stats, denyGate)).toMatchObject({ kind: 'no-match' })
    // stdDev with equalAverage is normatively invalid if hand-constructed; on parse it repairs to plain equalAverage (Ruling R3).
    const invalid = rule({ kind: 'aboveAverage', above: true, equal: true, stdDev: 1 })
    expect(evaluatePredicate(invalid, at, num(11), stats, denyGate)).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })

    // R3: stdDev<=0 or stdDev with equalAverage are repaired on parse to plain aboveAverage (preserving equalAverage)
    const stdDev0 = parsedDefinition('<cfRule type="aboveAverage" priority="1" stdDev="0"/>')
    expect(stdDev0).toMatchObject({ kind: 'aboveAverage', above: true, equal: false })
    expect(evaluatePredicate(rule(stdDev0), at, num(4), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })

    const stdDevNeg = parsedDefinition('<cfRule type="aboveAverage" priority="1" stdDev="-1"/>')
    expect(stdDevNeg).toMatchObject({ kind: 'aboveAverage', above: true, equal: false })
    expect(evaluatePredicate(rule(stdDevNeg), at, num(4), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })

    const stdDevEq = parsedDefinition('<cfRule type="aboveAverage" priority="1" stdDev="1" equalAverage="1"/>')
    expect(stdDevEq).toMatchObject({ kind: 'aboveAverage', above: true, equal: true })
    expect(evaluatePredicate(rule(stdDevEq), at, num(2), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
  })

  it('stddev never falls back to population stddev (fails closed without sample stddev)', () => {
    // Ruling I2: When stats has only stddevPop and no stddevSample, stddev rules fail closed (no-match)
    // rather than using population stddev.
    const at = { row: 0, col: 0 }
    const pop: CfRangeStatistics = {
      scopeKey: 'p', numericCount: 4, sum: 21, mean: 5.25, min: 0, max: 11, stddevPop: 5.2619,
      sortedAsc: Float64Array.from([0, 0, 10, 11]), frequencies: new Map(), declaredCells: 4, spanCells: 4, storedCells: 4,
    }
    const std1 = rule({ kind: 'aboveAverage', above: true, equal: false, stdDev: 1 })
    expect(evaluatePredicate(std1, at, num(11), pop, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(std1, at, num(0), pop, denyGate).kind).toBe('unresolved')
    // Under allowGate, missing stddevSample yields no-match, NOT matching via population stddev
    expect(evaluatePredicate(std1, at, num(11), pop, allowGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(std1, at, num(0), pop, allowGate)).toMatchObject({ kind: 'no-match' })
    const std1below = rule({ kind: 'aboveAverage', above: false, equal: false, stdDev: 1 })
    const popBelow: CfRangeStatistics = {
      scopeKey: 'pb', numericCount: 4, sum: 21, mean: 5.25, min: -2, max: 11, stddevPop: 5.4486,
      sortedAsc: Float64Array.from([-2, 2, 10, 11]), frequencies: new Map(), declaredCells: 4, spanCells: 4, storedCells: 4,
    }
    expect(evaluatePredicate(std1below, at, num(-2), popBelow, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(std1below, at, num(-2), popBelow, allowGate)).toMatchObject({ kind: 'no-match' })
  })

  it('stddev uses sample standard deviation (Excel STDEV.S) and fails closed on range errors', async () => {
    // Stage 1 stage1-above-average-stddev values: {0,0,10,11}, mean 5.25.
    // Sample stddev = sqrt(110.75 / 3) = 6.0759. Threshold for stdDev=1: 5.25 + 6.0759 = 11.3259.
    // 11 < 11.3259: F5 does not match (Excel oracle leaves F5 white).
    const at = { row: 0, col: 0 }
    const sampleStats: CfRangeStatistics = {
      scopeKey: 's', numericCount: 4, sum: 21, mean: 5.25, min: 0, max: 11,
      stddevPop: 5.2619,
      stddevSample: 6.0759,
      sortedAsc: Float64Array.from([0, 0, 10, 11]), frequencies: new Map(), declaredCells: 4, spanCells: 4, storedCells: 4,
    }
    const std1 = rule({ kind: 'aboveAverage', above: true, equal: false, stdDev: 1 })
    // Under sample stddev, 11 is not > 11.3259 -> no-match
    expect(evaluatePredicate(std1, at, num(11), sampleStats, allowGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(std1, at, num(12), sampleStats, allowGate)).toMatchObject({ kind: 'ready', value: true })

    // When range has errors, sample stddev is undefined (STDEV.S fails closed), yielding no-match.
    // Ruling I2: Even when stddevPop is defined, sample stddev must not fall back to population stddev.
    const errorStats: CfRangeStatistics = {
      ...sampleStats,
      stddevSample: undefined,
      stddevPop: 5.2619,
    }
    expect(evaluatePredicate(std1, at, num(12), errorStats, allowGate)).toMatchObject({ kind: 'no-match' })

    // Exercise with real prepareStatistics over an error-bearing range:
    const mem = createCfMemory({})
    const statSession = createCfStatistics(mem)
    const errorCells: XlsxCell[] = [
      { row: 0, col: 0, ref: 'A1', value: 10, valueIsError: false, styleIndex: 0 },
      { row: 1, col: 0, ref: 'A2', value: 20, valueIsError: false, styleIndex: 0 },
      { row: 2, col: 0, ref: 'A3', value: '#N/A', valueIsError: true, styleIndex: 0 },
    ]
    const errorSrc = createResidentCfSource(errorCells, SHEET, stamp(1), mem)
    const [realOut] = await statSession.prepareStatistics([[rect(0, 2)]], stamp(1), errorSrc, frozenBudget())
    expect(realOut!.kind).toBe('ready')
    if (realOut!.kind === 'ready') {
      const realStats = realOut.value.value
      expect(realStats.stddevSample).toBeUndefined()
      expect(realStats.stddevPop).toBeDefined()
      expect(evaluatePredicate(std1, at, num(100), realStats, allowGate)).toMatchObject({ kind: 'no-match' })
      realOut.value.release()
    }
  })

  it('aboveAverage on statistics crossing hidden rows gates on layout:hidden-row-statistics on any range', () => {
    // Non-fixture range B10:B15 (rows 9..14, col 1):
    const at = { row: 10, col: 1 }
    const r = rule({ kind: 'aboveAverage', above: true, equal: false }, [rect(9, 14, 1, 1)])
    const statsWithHidden: CfRangeStatistics = {
      scopeKey: 'h', numericCount: 5, sum: 150, mean: 30, min: 10, max: 50, stddevPop: 10,
      sortedAsc: Float64Array.from([10, 20, 30, 40, 50]), frequencies: new Map(), declaredCells: 6, spanCells: 6, storedCells: 5,
      hasHiddenRows: true,
    }
    const statsWithoutHidden: CfRangeStatistics = {
      ...statsWithHidden,
      hasHiddenRows: false,
    }
    // Gated when hasHiddenRows is true
    expect(evaluatePredicate(r, at, num(40), statsWithHidden, denyGate)).toMatchObject({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
    expect(evaluatePredicate(r, at, num(40), statsWithHidden, allowGate)).toMatchObject({ kind: 'ready', value: true })
    // Ungated when hasHiddenRows is false (even with denyGate)
    expect(evaluatePredicate(r, at, num(40), statsWithoutHidden, denyGate)).toMatchObject({ kind: 'ready', value: true })
  })

  it('duplicateValues and uniqueValues follow the manifest sets', () => {
    // Manifest stage1-duplicate-scope singles: numeric 1/1 duplicate.
    // Text equivalence is gated below; blanks and errors never match.
    const freq = new Map<string, number>([['n:1', 2], ['s:a', 2], ['s:á', 1]])
    const stats: CfRangeStatistics = {
      scopeKey: 't', numericCount: 3, sum: 3, mean: 1, min: 1, max: 1, stddevPop: 0,
      sortedAsc: Float64Array.from([1, 1, 1]), frequencies: freq, declaredCells: 9, spanCells: 9, storedCells: 9,
    }
    const at = { row: 0, col: 0 }
    const dup = rule({ kind: 'predicate', type: 'duplicateValues' })
    expect(evaluatePredicate(dup, at, num(1), stats, denyGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(dup, at, blank, stats, denyGate)).toMatchObject({ kind: 'no-match' })
    const uniq = rule({ kind: 'predicate', type: 'uniqueValues' })
    expect(evaluatePredicate(uniq, at, num(1), stats, denyGate)).toMatchObject({ kind: 'no-match' })
  })

  it('duplicate text equivalence is gated on its evidence variant', () => {
    // Manifest B5/B6 ("A"/"a" duplicate case-insensitively) and C7 ("á" unique)
    // need text equivalence, which is unmeasured: withheld without evidence.
    const freq = new Map<string, number>([['s:a', 2], ['s:á', 1]])
    const stats: CfRangeStatistics = {
      scopeKey: 't', numericCount: 0, sum: 0, mean: undefined, min: undefined, max: undefined,
      stddevPop: undefined, sortedAsc: Float64Array.from([]), frequencies: freq, declaredCells: 3, spanCells: 3, storedCells: 3,
    }
    const at = { row: 0, col: 0 }
    const dup = rule({ kind: 'predicate', type: 'duplicateValues' })
    const uniq = rule({ kind: 'predicate', type: 'uniqueValues' })
    for (const value of [str('A'), str('a'), str('á')] as CfLiteral[]) {
      expect(evaluatePredicate(dup, at, value, stats, denyGate).kind).toBe('unresolved')
      expect(evaluatePredicate(uniq, at, value, stats, denyGate).kind).toBe('unresolved')
    }
    expect(evaluatePredicate(dup, at, str('A'), stats, allowGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(dup, at, str('a'), stats, allowGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(dup, at, str('á'), stats, allowGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(uniq, at, num(1), stats, allowGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(uniq, at, str('á'), stats, allowGate)).toMatchObject({ kind: 'ready', value: true })
  })

  it('multi-rectangle union scope is gated on its evidence variant', () => {
    // Manifest D3/F2 (disjoint union) and H3/I3 (overlap without double-count)
    // require union semantics, which are unmeasured: withheld without evidence.
    const freq = new Map<string, number>([['n:20', 2], ['n:77', 1], ['n:88', 2]])
    const stats: CfRangeStatistics = {
      scopeKey: 'u', numericCount: 5, sum: 0, mean: 0, min: 0, max: 0, stddevPop: 0,
      sortedAsc: Float64Array.from([]), frequencies: freq, declaredCells: 6, spanCells: 6, storedCells: 5,
    }
    const scope = [rect(0, 1), rect(0, 1, 2, 2)]
    const at = { row: 0, col: 0 }
    const dup = rule({ kind: 'predicate', type: 'duplicateValues' }, scope)
    const uniq = rule({ kind: 'predicate', type: 'uniqueValues' }, scope)
    expect(evaluatePredicate(dup, at, num(20), stats, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(uniq, at, num(77), stats, denyGate).kind).toBe('unresolved')
    expect(evaluatePredicate(dup, at, num(20), stats, allowGate)).toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(dup, at, num(77), stats, allowGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(uniq, at, num(77), stats, allowGate)).toMatchObject({ kind: 'ready', value: true })
    // top10/aboveAverage on multi-rectangle scopes stay unresolved: no manifest rows.
    const top = rule({ kind: 'top10', rank: 2, bottom: false, percent: false }, scope)
    expect(evaluatePredicate(top, at, num(20), stats, allowGate).kind).toBe('unresolved')
    const avg = rule({ kind: 'aboveAverage', above: true, equal: false }, scope)
    expect(evaluatePredicate(avg, at, num(20), stats, allowGate).kind).toBe('unresolved')
  })

  it('unresolved definitions propagate their reason and missing stats refuse', () => {
    const at = { row: 0, col: 0 }
    const unsupported = rule({ kind: 'unresolved', type: 'expression', reason: 'cf-unsupported-rule' })
    expect(evaluatePredicate(unsupported, at, num(5), undefined, denyGate)).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    const invalid = rule({ kind: 'unresolved', type: 'top10', reason: 'cf-invalid-rule' })
    expect(evaluatePredicate(invalid, at, num(5), undefined, denyGate)).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
    // A statistics rule without prepared statistics cannot decide.
    const top = rule({ kind: 'top10', rank: 2, bottom: false, percent: false })
    expect(evaluatePredicate(top, at, num(5), undefined, denyGate)).toMatchObject({ kind: 'unresolved', reason: 'cf-missing-statistics' })
  })

  it('cells outside every range never match', () => {
    const r = rule({ kind: 'cellIs', operator: 'equal', operands: [num(100)] }, [rect(0, 1)])
    expect(evaluatePredicate(r, { row: 9, col: 9 }, num(100), undefined, denyGate)).toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(r, { row: 1, col: 0 }, num(100), undefined, denyGate)).toMatchObject({ kind: 'ready', value: true })
  })
})
