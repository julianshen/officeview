/**
 * B1 reference geometry and resumable traversal regressions.
 *
 * Finite B1 scope only: ordered unions/intersections, whole-axis geometry,
 * shared per-endpoint anchored translation, ROWS/COLUMNS helpers.
 * Native measured rows (NATIVE-REFERENCES.json, verified): SUM((B2,B4))=40,
 * SUM(B2:B4 B3:B5)=50, ROWS(A:A)=1048576, COLUMNS(1:1)=16384 with the exact
 * fixture B2=10/B3=20/B4=30/B5=40. Later families (lookup/conditional/date/
 * text/array) are excluded. No native calculation claimed beyond cited rows.
 */
import { describe, expect, it } from 'vitest'
import { evaluateFormula } from '../src/xlsx/formula/evaluator'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import { parseFormula } from '../src/xlsx/formula/parser'
import type {CellAddress,
  DefinedNameMetadata,
  ElementValue,
  EvaluationContext,
  ReferenceNode,
  SheetId,
  TableMetadata } from '../src/xlsx/formula/types'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'
import * as functionsModule from '../src/xlsx/formula/functions'
import type { AstNode } from '../src/xlsx/formula/types'
import { isEvaluationError } from '../src/xlsx/formula/evaluator'
import {
  createReferenceServices,
  type ReferenceServiceDeps,
  type SparseCellStore,
  type StoredCell,
} from '../src/xlsx/formula/refs'
import { translateSharedFormula } from '../src/xlsx/formula/shared'

// Exact native fixture coordinates: B2=10, B3=20, B4=30, B5=40.
function fixtureCtx(): EvaluationContext {
  const values = new Map<string, number>([
    ['1:1', 10], // B2
    ['1:2', 20], // B3
    ['1:3', 30], // B4
    ['1:4', 40], // B5
  ])
  return {
    getCellValue: (_sheet: string | undefined, col: number, row: number) =>
      values.get(`${col}:${row}`) ?? null,
  }
}

describe('B1: native measured union/intersection aggregates', () => {
  it('SUM((B2,B4)) is 40 for B2=10 and B4=30', () => {
    expect(evaluateFormula('SUM((B2,B4))', fixtureCtx())).toBe(40)
  })

  it('SUM(B2:B4 B3:B5) is 50 for overlap B3:B4=20+30', () => {
    expect(evaluateFormula('SUM(B2:B4 B3:B5)', fixtureCtx())).toBe(50)
  })
})

describe('B1: native measured whole-axis dimensions', () => {
  it('ROWS(A:A) is 1048576', () => {
    expect(evaluateFormula('ROWS(A:A)', fixtureCtx())).toBe(1048576)
  })

  it('COLUMNS(1:1) is 16384', () => {
    expect(evaluateFormula('COLUMNS(1:1)', fixtureCtx())).toBe(16384)
  })
})

describe('B1: shared per-endpoint anchored translation', () => {
  it('keeps the absolute from-anchor in $A:C shifted by one column', () => {
    expect(translateSharedFormula('$A:C', 1, 0).formula).toBe('$A:D')
  })

  it('keeps the absolute to-anchor in A:$C shifted by one column', () => {
    expect(translateSharedFormula('A:$C', 1, 0).formula).toBe('B:$C')
  })

  it('keeps the absolute from-anchor in $1:3 shifted by one row', () => {
    expect(translateSharedFormula('$1:3', 0, 1).formula).toBe('$1:4')
  })

  it('translates 3D targets while keeping the sheet run', () => {
    expect(translateSharedFormula('Sheet1:Sheet2!A1', 1, 0).formula).toBe('Sheet1:Sheet2!B1')
  })

  it('translates union operands independently', () => {
    expect(translateSharedFormula('SUM((B2,B4))', 1, 0).formula).toBe('SUM((C2,C4))')
  })
})

// Structural + synthetic integration controls below. Provenance: native
// measured rows keep their NATIVE-REFERENCES.json labels; documented rows
// keep native-b1-r1/FORMULAS.json labels (primary-source-derived, fixture
// import failed — not native measured); everything else is synthetic and
// labeled as such. Later families stay excluded.

interface FakeCell {
  sheet: string
  col: number
  row: number
  value: ElementValue
  origin?: 'input' | 'formula' | 'spill'
}

function fakeServices(
  cells: FakeCell[],
  opts: {
    sheets?: Array<{ sheetId: SheetId; name: string; workbookIndex: number }>
    names?: DefinedNameMetadata[]
    tables?: TableMetadata[] | undefined
    issues?: Array<{ feature: string; message: string }>
    reads?: { count: number }
    suspendOnceAt?: { col: number; row: number }
    prepareSpy?: { called: boolean }
    gateReadsOnPrepare?: boolean
  } = {},
): {
  services: ReturnType<typeof createReferenceServices>
  deps: ReferenceServiceDeps
  gen: { n: number }
  backend: (sheet: string | undefined, col: number, row: number) => ElementValue
} {
  const gen = { n: 0 }
  const sheets = opts.sheets ?? [{ sheetId: 'S', name: 'S', workbookIndex: 0 }]
  const byId = new Map(sheets.map((s) => [s.sheetId, s]))
  const byName = new Map(sheets.map((s) => [s.name.toLowerCase(), s]))
  const stored: StoredCell[] = cells.map((c) => {
    const id = byName.get(c.sheet.toLowerCase())?.sheetId ?? c.sheet
    return {
      address: { sheetId: id, col: c.col, row: c.row },
      value: c.value,
      origin: c.origin ?? 'input',
    }
  })
  const suspended = new Set<string>()
  const store: SparseCellStore = {
    read: (address: CellAddress) => {
      if (opts.reads) opts.reads.count++
      if (opts.gateReadsOnPrepare && !opts.prepareSpy?.called) throw new Error('enumerated before prepare')
      const key = `${address.sheetId}:${address.col}:${address.row}`
      if (
        opts.suspendOnceAt &&
        address.col === opts.suspendOnceAt.col &&
        address.row === opts.suspendOnceAt.row &&
        !suspended.has(key)
      ) {
        suspended.add(key)
        throw new Error('injected test suspension')
      }
      const hit = stored.find(
        (s) => s.address.sheetId === address.sheetId && s.address.col === address.col && s.address.row === address.row,
      )
      return hit ? { value: hit.value, origin: hit.origin } : undefined
    },
    stored: () => stored.filter((s) => !(s.value === null && s.origin === 'input')),
  }
  const deps: ReferenceServiceDeps = {
    generation: () => gen.n,
    sheets,
    sheetIdOfName: (name: string) => byName.get(name.toLowerCase())?.sheetId,
    sheetNameOfId: (id: SheetId) => byId.get(id)?.name,
    store,
    definedNames: opts.names ?? [],
    // Explicit undefined inventory means unavailable (never known-empty).
    tables: 'tables' in opts ? opts.tables : [],
    onIssue: (feature: string, message: string) => {
      opts.issues?.push({ feature, message })
    },
  }
  return { services: createReferenceServices(deps), deps, gen, backend }

  function backend(sheet: string | undefined, col: number, row: number): ElementValue {
    // Shared backend: legacy callbacks read the same cells as services.
    const id = sheet !== undefined ? byName.get(sheet.toLowerCase())?.sheetId : undefined
    const hit = stored.find(
      (s) => (id === undefined || s.address.sheetId === id) && s.address.col === col && s.address.row === row,
    )
    return hit ? hit.value : null
  }
}

function evalCtx(
  services: ReturnType<typeof createReferenceServices>,
  opts: {
    sheet?: string
    cell?: { col: number; row: number }
    backend?: (sheet: string | undefined, col: number, row: number) => ElementValue
  } = {},
): EvaluationContext {
  return {
    currentSheet: opts.sheet ?? 'S',
    ...(opts.cell ? { currentCell: { ...opts.cell, absCol: false, absRow: false } } : {}),
    ...(opts.backend ? { getCellValue: opts.backend } : {}),
    flatArgs: new WeakMap(),
    references: services,
  }
}

describe('B1: grammar extras (duplicates, empty intersection, spill, @, quoting)', () => {
  it('SUM((B2,B2)) counts duplicate occurrences twice', () => {
    expect(evaluateFormula('SUM((B2,B2))', fixtureCtx())).toBe(20)
  })

  it('empty intersection is typed #NULL!', () => {
    expect(evaluateFormula('SUM(A1:A2 C3:C4)', fixtureCtx())).toBe('#NULL!')
  })

  it('spill suffix parses; unavailable ownership emits a capability diagnostic, not a genuine #REF!', () => {
    expect(parseFormula('A1#')).toEqual({
      type: 'spill',
      anchor: { col: 0, row: 0, absCol: false, absRow: false },
    })
    // Spill lifecycle is a later stage: the vehicle is #NAME? with a marked
    // capability diagnostic (SEMANTICS: unavailable services preserve valid
    // caches with a specific issue — never an invented genuine error).
    const ctx: EvaluationContext = { ...fixtureCtx(), unsupportedFeatures: new Set<string>() }
    expect(evaluateFormula('A1#', ctx)).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('spill-unavailable')).toBe(true)
  })

  it('@ projects a vertical vector on the current row', () => {
    const ctx: EvaluationContext = {
      ...fixtureCtx(),
      currentCell: { col: 0, row: 2, absCol: false, absRow: false },
    }
    // Zero-based row 2 is Excel row 3: B2:B4 at the current row is B3=20.
    expect(evaluateFormula('@B2:B4', ctx)).toBe(20)
  })

  it('quoted sheet references keep ordinary values', () => {
    expect(evaluateFormula("SUM(' Data '!B2:B4)", fixtureCtx())).toBe(60)
  })

  it('range error literals propagate instead of evaluating', () => {
    expect(evaluateFormula('SUM(#N/A,B2)', fixtureCtx())).toBe('#N/A')
  })
})

describe('B1 documented helpers (native-b1-r1/FORMULAS.json, primary-source-derived)', () => {
  it('B1-026 ROW() at F2 is 2', () => {
    expect(
      evaluateFormula('ROW()', { currentCell: { col: 5, row: 1, absCol: false, absRow: false } }),
    ).toBe(2)
  })

  it('B1-027 COLUMN() at F3 is 6', () => {
    expect(
      evaluateFormula('COLUMN()', { currentCell: { col: 5, row: 2, absCol: false, absRow: false } }),
    ).toBe(6)
  })

  it('B1-028 ROW(Names!$B$4) is 4', () => {
    expect(evaluateFormula('ROW(Names!$B$4)', fixtureCtx())).toBe(4)
  })

  it('B1-029 COLUMN(Names!$C$4) is 3', () => {
    expect(evaluateFormula('COLUMN(Names!$C$4)', fixtureCtx())).toBe(3)
  })

  it('B1-043 NoHeader[#Headers] is typed #REF!', () => {
    const tables: TableMetadata[] = [{
      id: '1',
      name: 'NoHeader',
      displayName: 'NoHeader',
      sheetId: 'S',
      partPath: 'xl/tables/table1.xml',
      extent: { sheetId: 'S', firstCol: 0, firstRow: 0, cols: 2, rows: 2 },
      headerRowCount: 0,
      totalsRowCount: 0,
      columns: [
        { id: '1', name: 'A', index: 0 },
        { id: '2', name: 'B', index: 1 },
      ],
    }]
    const { services } = fakeServices([], { tables })
    expect(evaluateFormula('NoHeader[#Headers]', evalCtx(services))).toBe('#REF!')
  })

  it('NA() is tagged #N/A', () => {
    expect(evaluateFormula('NA()', fixtureCtx())).toBe('#N/A')
  })
})

describe('B1: defined-name binding (synthetic; native source/value rows pending)', () => {
  const names: DefinedNameMetadata[] = [
    { name: 'MyName', localSheetIndex: 1, source: 'Data!$A$1', baseProvenance: 'unknown' },
    { name: 'MyName', source: 'Data!$B$2', baseProvenance: 'unknown' },
    { name: 'Plain', source: '1+1', baseProvenance: 'unknown' },
    { name: 'Relative', source: 'Data!A1', baseProvenance: 'unknown' },
  ]
  const sheets = [
    { sheetId: 'c', name: 'Chart1', workbookIndex: 0 },
    { sheetId: 'd', name: 'Data', workbookIndex: 1 },
  ]

  it('local scope shadows global through original workbook order', () => {
    const { services } = fakeServices(
      [{ sheet: 'Data', col: 0, row: 0, value: 42 }],
      { sheets, names },
    )
    const binding = services.bindName({ name: 'MyName' }, { currentSheet: 'Data' })
    expect(binding).toMatchObject({ name: 'MyName', source: 'Data!$A$1', scopeSheetId: 'd' })
  })

  it('explicit qualification selects that local scope', () => {
    const { services } = fakeServices([], { sheets, names })
    const binding = services.bindName({ name: 'MyName', sheet: 'Chart1' }, { currentSheet: 'Data' })
    // Chart1 has no local MyName: falls back to the workbook-global binding.
    expect(binding).toMatchObject({ source: 'Data!$B$2' })
  })

  it('qualified absolute names evaluate through the binding', () => {
    const { services } = fakeServices(
      [
        { sheet: 'Data', col: 0, row: 0, value: 42 },
        { sheet: 'Data', col: 1, row: 1, value: 7 },
      ],
      { sheets, names },
    )
    expect(evaluateFormula('Data!MyName', evalCtx(services, { sheet: 'Data' }))).toBe(42)
  })

  it('value-denoting names evaluate their source expression', () => {
    const f = fakeServices(
      [{ sheet: 'Data', col: 1, row: 1, value: 10 }],
      { sheets, names },
    )
    // Bare identifiers stay grammar errors (Core retention); qualified names
    // bind: Data has no local Plain, so the workbook-global source applies.
    // Legacy callbacks and reference services share one backend: B2 reads 10
    // through getCellValue while Plain=1+1 binds through metadata.
    expect(
      evaluateFormula('SUM(Data!Plain,B2)', evalCtx(f.services, { sheet: 'Data', backend: f.backend })),
    ).toBe(12)
  })

  it('relative names with unknown base are a provenance gate, never guessed', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const { services } = fakeServices(
      [{ sheet: 'Data', col: 0, row: 0, value: 42 }],
      { sheets, names, issues },
    )
    // Unknown relative bases gate with a diagnostic (caches preserved),
    // never a guessed offset nor an invented genuine error.
    const ctx = evalCtx(services, { sheet: 'Data' })
    expect(evaluateFormula('Data!Relative', ctx)).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('relative-name')).toBe(true)
    expect(issues.some((i) => i.feature === 'relative-name')).toBe(true)
  })

  it('unknown names stay typed #NAME?', () => {
    const { services } = fakeServices([], { sheets, names })
    expect(evaluateFormula('Data!Nope', evalCtx(services, { sheet: 'Data' }))).toBe('#NAME?')
  })
})

describe('B1: table evaluation (synthetic geometry; native rows pending)', () => {
  const tables: TableMetadata[] = [{
    id: '1',
    name: 'Table1',
    displayName: 'Table1',
    sheetId: 'S',
    partPath: 'xl/tables/table1.xml',
    extent: { sheetId: 'S', firstCol: 0, firstRow: 0, cols: 2, rows: 4 },
    headerRowCount: 1,
    totalsRowCount: 0,
    columns: [
      { id: '1', name: 'Amount', index: 0 },
      { id: '2', name: 'Tax', index: 1 },
    ],
  }]

  it('SUM(Table1[Amount]) aggregates the data body only', () => {
    const { services } = fakeServices(
      [
        { sheet: 'S', col: 0, row: 0, value: 'Amount' },
        { sheet: 'S', col: 0, row: 1, value: 10 },
        { sheet: 'S', col: 0, row: 2, value: 20 },
        { sheet: 'S', col: 0, row: 3, value: 30 },
      ],
      { tables },
    )
    expect(evaluateFormula('SUM(Table1[Amount])', evalCtx(services))).toBe(60)
  })

  it('unknown tables are #NAME? and unknown columns are #REF!', () => {
    const { services } = fakeServices([], { tables })
    expect(evaluateFormula('SUM(Nope[Amount])', evalCtx(services))).toBe('#NAME?')
    expect(evaluateFormula('SUM(Table1[Nope])', evalCtx(services))).toBe('#REF!')
  })

  it('unavailable table inventory emits a capability diagnostic, not a silent empty sum', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const { services } = fakeServices([], { tables: undefined, issues })
    // SEMANTICS: unavailable services preserve valid caches with a specific
    // issue — the vehicle is #NAME? with a marked diagnostic, never an
    // invented genuine error and never a silent zero.
    const ctx = evalCtx(services)
    expect(evaluateFormula('SUM(Table1[Amount])', ctx)).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('table-unavailable')).toBe(true)
    expect(issues.some((i) => i.feature === 'table-unavailable')).toBe(true)
  })

  it('unavailable tables preserve workbook caches with a specific diagnostic', () => {
    const doc = {
      sheets: [{
        name: 'S',
        sourcePartPath: 's',
        rows: [{ index: 0, cells: [
          { ref: 'A1', row: 0, col: 0, value: 99, styleIndex: 0, hasCachedValue: true, formula: '=SUM(Table1[Amount])' },
        ] as XlsxCell[] }],
      }],
      images: [],
      drawingCoverage: [],
      tables: undefined,
    } as never as XlsxDocument
    evaluateWorkbookFormulas(doc as never, { forceRecalc: true })
    expect(doc.sheets[0].rows[0].cells[0].value).toBe(99)
    const diags = (doc as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics ?? []
    expect(diags.some((d) => d.feature === 'table-unavailable')).toBe(true)
  })

  it('unavailable spills preserve workbook caches with a specific diagnostic', () => {
    // C2 phase2 maintenance: B1's generic 'spill-unavailable' capability gate is
    // superseded by the workbook spill service, which reports the specific
    // 'spill-provenance' gate for an anchor with no verified dynamic/legacy
    // provenance (failed/nonspilling-anchor error remains unmeasured). The
    // paired behavior is unchanged: the valid cache is retained with a specific
    // diagnostic, never an invented genuine error.
    const doc = {
      sheets: [{
        name: 'S',
        sourcePartPath: 's',
        rows: [{ index: 0, cells: [
          { ref: 'B2', row: 1, col: 1, value: 5, styleIndex: 0, hasCachedValue: true, formula: '=A1#' },
        ] as XlsxCell[] }],
      }],
      images: [],
      drawingCoverage: [],
      tables: [],
    } as never as XlsxDocument
    evaluateWorkbookFormulas(doc as never, { forceRecalc: true })
    expect(doc.sheets[0].rows[0].cells[0].value).toBe(5)
    const diags = (doc as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics ?? []
    expect(diags.some((d) => d.feature === 'spill-provenance')).toBe(true)
  })
})

describe('B1: 3D references (synthetic; native rows pending)', () => {
  it('resolves sheet runs in original workbook order', () => {
    const { services } = fakeServices(
      [
        { sheet: 'S1', col: 0, row: 0, value: 10 },
        { sheet: 'S2', col: 0, row: 0, value: 20 },
      ],
      {
        sheets: [
          { sheetId: 's1', name: 'S1', workbookIndex: 0 },
          { sheetId: 's2', name: 'S2', workbookIndex: 1 },
        ],
      },
    )
    const resolved = services.resolve(
      { type: 'ref3d', sheets: { fromSheet: 'S1', toSheet: 'S2' }, target: { type: 'cell', ref: { col: 0, row: 0, absCol: false, absRow: false } } },
      { currentSheet: 'S1' },
    )
    expect(resolved).toMatchObject({
      areas: [
        { sheetId: 's1', firstCol: 0, firstRow: 0, cols: 1, rows: 1 },
        { sheetId: 's2', firstCol: 0, firstRow: 0, cols: 1, rows: 1 },
      ],
    })
  })

  it('SUM across a sheet run aggregates every sheet once', () => {
    const { services } = fakeServices(
      [
        { sheet: 'S1', col: 0, row: 0, value: 10 },
        { sheet: 'S2', col: 0, row: 0, value: 20 },
      ],
      {
        sheets: [
          { sheetId: 's1', name: 'S1', workbookIndex: 0 },
          { sheetId: 's2', name: 'S2', workbookIndex: 1 },
        ],
      },
    )
    expect(evaluateFormula('SUM(S1:S2!A1)', evalCtx(services, { sheet: 'S1' }))).toBe(30)
  })

  it('reversed sheet runs keep an explicit unknown/provenance gate', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const { services } = fakeServices(
      [],
      {
        sheets: [
          { sheetId: 's1', name: 'S1', workbookIndex: 0 },
          { sheetId: 's2', name: 'S2', workbookIndex: 1 },
        ],
        issues,
      },
    )
    // Reversed-run native semantics are unmeasured: no computed error is
    // assumed solely to pass. The gate preserves caches with a diagnostic.
    const ctx = evalCtx(services, { sheet: 'S1' })
    expect(evaluateFormula('SUM(S2:S1!A1)', ctx)).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('reversed-3d')).toBe(true)
    expect(issues.some((i) => i.feature === 'reversed-3d')).toBe(true)
  })

  it('unknown sheet runs are computed #REF!', () => {
    const { services } = fakeServices(
      [],
      {
        sheets: [
          { sheetId: 's1', name: 'S1', workbookIndex: 0 },
          { sheetId: 's2', name: 'S2', workbookIndex: 1 },
        ],
      },
    )
    expect(evaluateFormula('SUM(S1:Missing!A1)', evalCtx(services, { sheet: 'S1' }))).toBe('#REF!')
  })
})

describe('B1: sparse geometry and cursors (synthetic structural controls)', () => {
  it('whole-column areas keep full geometry without allocation', () => {
    const { services } = fakeServices([{ sheet: 'S', col: 0, row: 499999, value: 0 }])
    const resolved = services.resolve(
      { type: 'wholeCol', ref: { from: { index: 0, absolute: false }, to: { index: 0, absolute: false } } } as ReferenceNode,
      { currentSheet: 'S' },
    )
    expect(resolved).toMatchObject({ areas: [{ sheetId: 'S', firstCol: 0, firstRow: 0, cols: 1, rows: 1048576 }] })
  })

  it('positional reads reach row 500000 with a single counted read', () => {
    const reads = { count: 0 }
    const { services } = fakeServices(
      [{ sheet: 'S', col: 0, row: 499999, value: 5 }],
      { reads },
    )
    const resolved = services.resolve(
      { type: 'wholeCol', ref: { from: { index: 0, absolute: false }, to: { index: 0, absolute: false } } } as ReferenceNode,
      { currentSheet: 'S' },
    )
    if (!resolved || 'kind' in resolved === false || (resolved as { kind?: string }).kind !== 'resolved-reference') {
      throw new Error('expected a resolved reference')
    }
    const ref = resolved as import('../src/xlsx/formula/types').ResolvedRef
    expect(ref.readAt(0, 499999, 0)).toBe(5)
    expect(ref.readAt(0, 0, 0)).toBe(null)
    expect(reads.count).toBe(2)
    expect(ref.readAt(0, 1048576, 0)).toEqual({ kind: 'formula-error', code: '#REF!' })
  })

  it('aggregates read populated cells only, never blank enumeration', () => {
    const reads = { count: 0 }
    const { services } = fakeServices(
      [
        { sheet: 'S', col: 0, row: 0, value: 10 },
        { sheet: 'S', col: 0, row: 1, value: 20 },
        { sheet: 'S', col: 0, row: 2, value: '', origin: 'formula' },
      ],
      { reads },
    )
    expect(evaluateFormula('SUM(A1:A10)', evalCtx(services))).toBe(30)
    expect(reads.count).toBe(3)
  })

  it('countAbsent is structural and excludes formula empty text', () => {
    const { services } = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'S', col: 0, row: 1, value: 20 },
      { sheet: 'S', col: 0, row: 2, value: '', origin: 'formula' },
    ])
    const resolved = services.resolve(
      {
        type: 'range',
        ref: {
          from: { col: 0, row: 0, absCol: false, absRow: false },
          to: { col: 0, row: 9, absCol: false, absRow: false },
        },
      },
      { currentSheet: 'S' },
    )
    if (!resolved || (resolved as { kind?: string }).kind !== 'resolved-reference') {
      throw new Error('expected a resolved reference')
    }
    expect((resolved as import('../src/xlsx/formula/types').ResolvedRef).countAbsent()).toBe(7)
  })

  it('peek-then-commit resumes after suspension without skips or duplicates', () => {
    const { services } = fakeServices(
      [
        { sheet: 'S', col: 0, row: 0, value: 10 },
        { sheet: 'S', col: 0, row: 1, value: 20 },
        { sheet: 'S', col: 0, row: 2, value: 30 },
      ],
      { suspendOnceAt: { col: 0, row: 1 } },
    )
    const resolved = services.resolve(
      {
        type: 'range',
        ref: {
          from: { col: 0, row: 0, absCol: false, absRow: false },
          to: { col: 0, row: 2, absCol: false, absRow: false },
        },
      },
      { currentSheet: 'S' },
    )
    if (!resolved || (resolved as { kind?: string }).kind !== 'resolved-reference') {
      throw new Error('expected a resolved reference')
    }
    const ref = resolved as import('../src/xlsx/formula/types').ResolvedRef
    const cursor = ref.openCursor('populated')
    const seen: unknown[] = []
    for (;;) {
      let entry
      try {
        entry = ref.peek(cursor)
      } catch {
        entry = ref.peek(cursor)
      }
      if (entry === undefined) break
      seen.push(entry.value)
      ref.advance(cursor)
    }
    expect(seen).toEqual([10, 20, 30])
  })

  it('generation changes invalidate live cursors', () => {
    const { services, gen } = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }])
    const resolved = services.resolve(
      {
        type: 'range',
        ref: {
          from: { col: 0, row: 0, absCol: false, absRow: false },
          to: { col: 0, row: 0, absCol: false, absRow: false },
        },
      },
      { currentSheet: 'S' },
    )
    if (!resolved || (resolved as { kind?: string }).kind !== 'resolved-reference') {
      throw new Error('expected a resolved reference')
    }
    const ref = resolved as import('../src/xlsx/formula/types').ResolvedRef
    const cursor = ref.openCursor('populated')
    gen.n++
    expect(() => ref.peek(cursor)).toThrow(/stale|generation/)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 groups 1/10/11 — cursor commit/prepare/generation shipping regressions
// (ported from sealed B1/r1 spec-review probes.test.ts + cache-hooks.test.ts;
// provenance: contract/structural controls, no invented native values)
// ---------------------------------------------------------------------------
describe('B1 r2 group 1: cursor commits the exact peeked area/address (leading/middle empty, zero-size, both modes, suspension)', () => {
  it('leading empty union area is skipped in populated mode: peek/advance walk A1=10, A2=7 then EOF', () => {
    const { services } = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'S', col: 0, row: 1, value: 7 },
    ])
    const resolved = services.resolve(parseFormula('(C1:C2,A1:A2)') as ReferenceNode, evalCtx(services, { sheet: 'S', cell: { col: 1, row: 1 } }))
    if (!resolved || (resolved as { kind?: string }).kind !== 'resolved-reference') throw new Error('expected resolved reference')
    const ref = resolved as import('../src/xlsx/formula/types').ResolvedRef
    const cursor = ref.openCursor('populated')
    const e1 = ref.peek(cursor)
    expect(e1).toMatchObject({ address: { col: 0, row: 0 }, value: 10 })
    ref.advance(cursor)
    const e2 = ref.peek(cursor)
    expect(e2).toMatchObject({ address: { col: 0, row: 1 }, value: 7 })
    ref.advance(cursor)
    expect(ref.peek(cursor)).toBeUndefined()
  })

  it('leading empty union aggregate does not duplicate the next populated value: SUM((B1:B2,A1:A2)) = 7 (sealed probes row 31)', () => {
    const withCells = fakeServices([{ sheet: 'S', col: 0, row: 1, value: 7 }])
    expect(evaluateFormula('SUM((B1:B2,A1:A2))', evalCtx(withCells.services, { sheet: 'S', cell: { col: 1, row: 1 } }))).toBe(7)
  })

  it('middle empty union area: SUM((A1,C1:C2,A2)) = 17 with duplicate-occurrence fidelity', () => {
    const f = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'S', col: 0, row: 1, value: 7 },
    ])
    expect(evaluateFormula('SUM((A1,C1:C2,A2))', evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))).toBe(17)
  })

  it('zero-size area in populated mode is skipped: (Table1[#Data],A1) peeks A1=10 then EOF (sealed probes row 32)', () => {
    const table: TableMetadata = {
      id: 'one', name: 'Table1', displayName: 'Table1', sheetId: 'S', partPath: 'table.xml',
      extent: { sheetId: 'S', firstCol: 0, firstRow: 0, cols: 5, rows: 1 },
      headerRowCount: 1, totalsRowCount: 0,
      columns: [{ id: '1', name: 'Amount', index: 0 }],
    }
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], { tables: [table] })
    const resolved = f.services.resolve(parseFormula('(Table1[#Data],A1)') as ReferenceNode, evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))
    if (!resolved || (resolved as { kind?: string }).kind !== 'resolved-reference') throw new Error('expected resolved reference')
    const ref = resolved as import('../src/xlsx/formula/types').ResolvedRef
    const cursor = ref.openCursor('populated')
    expect(ref.peek(cursor)?.value).toBe(10)
    ref.advance(cursor)
    expect(ref.peek(cursor)).toBeUndefined()
  })

  it('all-mode cursor includes the zero-size/empty geometry without duplication', () => {
    const table: TableMetadata = {
      id: 'one', name: 'Table1', displayName: 'Table1', sheetId: 'S', partPath: 'table.xml',
      extent: { sheetId: 'S', firstCol: 0, firstRow: 0, cols: 5, rows: 1 },
      headerRowCount: 1, totalsRowCount: 0,
      columns: [{ id: '1', name: 'Amount', index: 0 }],
    }
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], { tables: [table] })
    const resolved = f.services.resolve(parseFormula('(Table1[#Data],A1)') as ReferenceNode, evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))
    if (!resolved || (resolved as { kind?: string }).kind !== 'resolved-reference') throw new Error('expected resolved reference')
    const ref = resolved as import('../src/xlsx/formula/types').ResolvedRef
    const cursor = ref.openCursor('all')
    const e1 = ref.peek(cursor) // zero-size area contributes no entries in all mode too
    expect(e1).toMatchObject({ address: { col: 0, row: 0 }, value: 10 })
    ref.advance(cursor)
    expect(ref.peek(cursor)).toBeUndefined()
  })

  it('suspension across a leading-empty union boundary preserves the exact peeked entry (retry, no skip/duplicate)', () => {
    const f = fakeServices(
      [
        { sheet: 'S', col: 0, row: 0, value: 10 },
        { sheet: 'S', col: 0, row: 1, value: 20 },
        { sheet: 'S', col: 0, row: 2, value: 30 },
      ],
      { suspendOnceAt: { col: 0, row: 1 } },
    )
    const resolved = f.services.resolve(parseFormula('(C1:C2,A1:A3)') as ReferenceNode, evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))
    if (!resolved || (resolved as { kind?: string }).kind !== 'resolved-reference') throw new Error('expected resolved reference')
    const ref = resolved as import('../src/xlsx/formula/types').ResolvedRef
    const cursor = ref.openCursor('populated')
    const seen: unknown[] = []
    for (;;) {
      let entry
      try {
        entry = ref.peek(cursor)
      } catch {
        entry = ref.peek(cursor) // retry after suspension: same entry, committed once
      }
      if (entry === undefined) break
      seen.push(entry.value)
      ref.advance(cursor)
    }
    expect(seen).toEqual([10, 20, 30])
  })

  it('workbook blank-first union computes 17 into the cached cell, not 27 (cache-hooks sealed control)', () => {
    const formula = '=SUM((C1:C2,A1:A2))'
    const doc = {
      sheets: [
        {
          name: 'S', sheetId: 's', sourcePartPath: 's',
          rows: [
            { index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 10, styleIndex: 0 } as XlsxCell] },
            { index: 1, cells: [{ ref: 'A2', col: 0, row: 1, value: 7, styleIndex: 0 } as XlsxCell] },
            { index: 2, cells: [{ ref: 'B3', col: 1, row: 2, value: 77, styleIndex: 0, hasCachedValue: true, formula } as unknown as XlsxCell] },
          ],
        },
      ],
      tables: [], images: [], drawingCoverage: [],
    } as unknown as XlsxDocument
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const cell = (doc.sheets[0].rows[2].cells as XlsxCell[])[0]
    expect(cell.value).toBe(17)
  })
})

describe('B1 r2 group 10: approved prepare service is invoked before open/enumerate/count/read', () => {
  it('injected prepare hook runs before the sparse enumeration throws-gate (sealed cache-hooks control)', () => {
    let prepared = false
    const ref = {
      kind: 'resolved-reference',
      areas: [{ sheetId: 's', firstCol: 0, firstRow: 0, cols: 1, rows: 1 }],
      generation: 1,
      openCursor: () => ({ id: 'a', mode: 'populated', generation: 1 }),
      peek: () => {
        if (!prepared) throw new Error('enumerated before prepare')
        return undefined
      },
      advance: () => {},
      readAt: () => null,
      countAbsent: () => 1,
    }
    const ctx = {
      flatArgs: new WeakMap(),
      references: {
        resolve: () => ref,
        bindName: () => ({ kind: 'formula-error', code: '#NAME?' }),
        prepare: () => {
          prepared = true
        },
      },
    } as unknown as EvaluationContext
    expect(() => evaluateFormula('SUM(A1:A2)', ctx)).not.toThrow()
    expect(prepared).toBe(true)
  })

  it('shipping harness (coordinator correction): overriding services.prepare directly is honored before gated enumeration — no deps.prepare ABI needed for scaffolding', () => {
    const prepareSpy = { called: false }
    const f = fakeServices(
      [
        { sheet: 'S', col: 0, row: 0, value: 10 },
        { sheet: 'S', col: 0, row: 1, value: 20 },
      ],
      { gateReadsOnPrepare: true, prepareSpy },
    )
    const services = {
      ...f.services,
      prepare: () => {
        prepareSpy.called = true
      },
    }
    expect(evaluateFormula('SUM(A1:A2)', evalCtx(services, { sheet: 'S', cell: { col: 1, row: 1 } }))).toBe(30)
    expect(prepareSpy.called).toBe(true)
  })
})

describe('B1 r2 group 11: generation invalidates readAt/countAbsent/memo as well as cursors', () => {
  function resolvedWholeCol() {
    const { services, gen } = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }])
    const resolved = services.resolve(
      { type: 'wholeCol', ref: { from: { index: 0, absolute: false }, to: { index: 0, absolute: false } } } as ReferenceNode,
      { currentSheet: 'S' },
    )
    if (!resolved || (resolved as { kind?: string }).kind !== 'resolved-reference') throw new Error('expected resolved reference')
    return { ref: resolved as import('../src/xlsx/formula/types').ResolvedRef, gen }
  }

  it('readAt throws after a generation change (sealed probes row 28)', () => {
    const { ref, gen } = resolvedWholeCol()
    ref.readAt(0, 0, 0)
    gen.n++
    expect(() => ref.readAt(0, 0, 0)).toThrow(/generation|stale/)
  })

  it('countAbsent throws after a generation change (sealed probes row 29)', () => {
    const { ref, gen } = resolvedWholeCol()
    ref.countAbsent()
    gen.n++
    expect(() => ref.countAbsent()).toThrow(/generation|stale/)
  })

  it('peek also throws after a generation change (existing cursor guard retained)', () => {
    const { ref, gen } = resolvedWholeCol()
    const cursor = ref.openCursor('populated')
    gen.n++
    expect(() => ref.peek(cursor)).toThrow(/generation|stale/)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 11 (generation drain-resume): two-generation suspended drain with
// changed footprint — flattened once, suspended, footprint changed, resumed on
// the SAME args WeakMap key. Contract: fresh items (no old values), no stale
// throw, caller-supplied memo ownership untouched.
// ---------------------------------------------------------------------------
class TestMarkerGlobal extends Error {}

describe('B1 r2 group 11: suspended drain re-resolves after a generation change (same args key, no stale throw, no old items)', () => {

  function twoGenFixture() {
    let cells: StoredCell[] = [
      { address: { sheetId: 's', col: 0, row: 0 }, value: 10, origin: 'input' },
      { address: { sheetId: 's', col: 0, row: 1 }, value: 20, origin: 'input' },
    ]
    let suspendsLeft = 1 // one injected suspension on the A2 read
    const gen = { n: 1 }
    const marker = new TestMarkerGlobal('injected test suspension')
    const store = {
      stored: () => cells,
      read: (address: { sheetId: string; col: number; row: number }) => {
        if (address.col === 0 && address.row === 1 && suspendsLeft > 0) {
          suspendsLeft--
          throw marker
        }
        const hit = cells.find(
          (c) => c.address.sheetId === address.sheetId && c.address.col === address.col && c.address.row === address.row,
        )
        return hit ? { value: hit.value, origin: hit.origin } : undefined
      },
    }
    const services = createReferenceServices({
      generation: () => gen.n,
      sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
      sheetIdOfName: (n) => (n.toLowerCase() === 's' ? 's' : undefined),
      sheetNameOfId: () => 'S',
      definedNames: [],
      tables: [],
      store,
    })
    const callerMemo = new WeakMap()
    const ctx = {
      references: services,
      currentSheet: 'S',
      currentCell: { col: 1, row: 1, absCol: false, absRow: false },
      flatArgs: new WeakMap(),
      nodeValues: callerMemo,
    } as unknown as EvaluationContext
    return {
      services,
      ctx,
      callerMemo,
      marker,
      changeFootprint() {
        cells = [
          { address: { sheetId: 's', col: 0, row: 0 }, value: 100, origin: 'input' },
          { address: { sheetId: 's', col: 0, row: 1 }, value: 200, origin: 'input' },
          { address: { sheetId: 's', col: 0, row: 2 }, value: 300, origin: 'input' },
        ]
        gen.n++
      },
    }
  }

  it('two-generation suspended drain: fresh 100/200 (sum 300), no old 10, caller nodeValues identity retained', () => {
    const { flattenArgs } = functionsModule
    const ast = parseFormula('SUM(A1:A2)') as { type: string; args: AstNode[] }
    expect(ast.type).toBe('call')
    const f = twoGenFixture()
    // First flatten: pushes A1=10, then store.read A2 throws the marker.
    let caught: unknown
    try {
      flattenArgs(ast.args, f.ctx, () => 0)
    } catch (e) {
      caught = e
    }
    expect(caught).toBe(f.marker) // exact marker identity
    // Change footprint (values/areas), advance generation, retry SAME args key.
    f.changeFootprint()
    const fresh = flattenArgs(ast.args, f.ctx, () => 0)
    expect(fresh.map((i) => i.value)).toEqual([100, 200]) // no old 10 anywhere
    // Caller-supplied memo ownership retained by identity.
    expect(f.ctx.nodeValues).toBe(f.callerMemo)
  })

  it('prepare throwing the pending marker propagates with identity; retry on the SAME args succeeds', () => {
    const { flattenArgs } = functionsModule
    const ast = parseFormula('SUM(A1:A2)') as { type: string; args: AstNode[] }
    const marker = new TestMarkerGlobal('prepare pending')
    let cells: StoredCell[] = [
      { address: { sheetId: 's', col: 0, row: 0 }, value: 10, origin: 'input' },
      { address: { sheetId: 's', col: 0, row: 1 }, value: 20, origin: 'input' },
    ]
    const services = createReferenceServices({
      generation: () => 1,
      sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
      sheetIdOfName: (n) => (n.toLowerCase() === 's' ? 's' : undefined),
      sheetNameOfId: () => 'S',
      definedNames: [],
      tables: [],
      store: {
        stored: () => cells,
        read: (address: { sheetId: string; col: number; row: number }) => {
          const hit = cells.find(
            (c) => c.address.sheetId === address.sheetId && c.address.col === address.col && c.address.row === address.row,
          )
          return hit ? { value: hit.value, origin: hit.origin } : undefined
        },
      },
    })
    let throwInPrepare = true
    const services2 = {
      ...services,
      prepare: () => {
        if (throwInPrepare) throw marker
      },
    }
    const ctx = {
      references: services2,
      currentSheet: 'S',
      currentCell: { col: 1, row: 1, absCol: false, absRow: false },
      flatArgs: new WeakMap(),
    } as unknown as EvaluationContext
    let caught: unknown
    try {
      flattenArgs(ast.args, ctx, () => 0)
    } catch (e) {
      caught = e
    }
    expect(caught).toBe(marker) // exact exception identity, uncaught by production
    throwInPrepare = false
    const fresh = flattenArgs(ast.args, ctx, () => 0)
    expect(fresh.map((i) => i.value)).toEqual([10, 20])
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 13 (root-authorized): cell-cell whitespace intersections parse
// and evaluate typed #NULL! / same-cell value (sealed intersection-
// authorization controls; Core obsolete expectation updated minimally)
// ---------------------------------------------------------------------------
describe('B1 r2 group 13: authorized cell-cell intersection grammar', () => {
  it('SUM(A1 B1) parses an intersect and evaluates typed #NULL! for empty overlap', () => {
    const ast = parseFormula('SUM(A1 B1)') as { type: string; args?: unknown[] }
    expect(ast.type).toBe('call')
    expect(ast.args?.[0]).toMatchObject({ type: 'intersect' })
    expect(evaluateFormula('SUM(A1 B1)', { getCellValue: () => 7 })).toBe('#NULL!')
  })
  it('SUM(A1 A1) parses an intersect and evaluates 7 with A1=7', () => {
    const ast = parseFormula('SUM(A1 A1)') as { type: string; args?: unknown[] }
    expect(ast.type).toBe('call')
    expect(ast.args?.[0]).toMatchObject({ type: 'intersect' })
    expect(evaluateFormula('SUM(A1 A1)', { getCellValue: () => 7 })).toBe(7)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 12: helper/domain/unknown-sheet validation. Unknown qualified
// name must not fall back to a global; the exact native name-error code is
// UNMEASURED — the contract-approved control is the explicit capability
// diagnostic + no-global-fallback, not an invented genuine #REF! assertion.
// Coordinate invalidity -> typed #REF! is contract-approved (ABI).
// ---------------------------------------------------------------------------
describe('B1 r2 group 12: helper argument validation without needless value reads', () => {
  it('ROW(Missing!A1) does not report geometry from an unknown sheet', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }])
    const r = evaluateFormula('ROW(Missing!A1)', evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))
    expect(r).not.toBe(1)
  })

  it('Missing!Named does not fall back to the unrelated global name; capability diagnostic preserved', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const f = fakeServices([], { names: [{ name: 'Named', source: '5', baseProvenance: 'unknown' }], issues })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    const r = evaluateFormula('Missing!Named', ctx)
    expect(r).not.toBe(5)
    // Explicit capability diagnostic (not an invented genuine error code) and
    // a tagged error result (valid-cache policy preserved by the evaluator).
    expect(ctx.unsupportedFeatures?.has('qualified-name-sheet-unavailable')).toBe(true)
    expect(issues.length).toBeGreaterThan(0)
  })

  it('resolve of an out-of-domain coordinate (col -1) is typed #REF! [ABI contract-approved]', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }])
    const r = f.services.resolve(
      { type: 'cell', ref: { col: -1, row: 0, absCol: false, absRow: false, sheet: undefined } },
      { currentSheet: 'S' },
    )
    expect(isEvaluationError(r) && r.code === '#REF!').toBe(true)
  })

  it('resolve of an out-of-domain row (1048576 0-based) is typed #REF! [ABI contract-approved]', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }])
    const r = f.services.resolve(
      { type: 'cell', ref: { col: 0, row: 1048576, absCol: false, absRow: false, sheet: undefined } },
      { currentSheet: 'S' },
    )
    expect(isEvaluationError(r) && r.code === '#REF!').toBe(true)
  })

  it('helper unknown-sheet grid argument yields no invented number', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }])
    const r = evaluateFormula('COLUMNS(Missing!A1:C1)', evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))
    expect(r).not.toBe(3)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 2 (ported sealed cache-hooks + probes): prohibited 3D operators
// and ambiguous @$ (2D/multi-area/missing context) gate at the evaluated frame
// with capability diagnostics and valid-cache retention; dead branches stay
// silent. Ordinary vertical/horizontal vectors remain supported. No invented
// native ambiguous-projection result.
// ---------------------------------------------------------------------------
describe('B1 r2 group 2: 3D intersection AND ambiguous $G gates with valid-cache retention', () => {
  it('SUM 3D x 2D inequality intersects to gate with diagnostic, not computed silence', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const f3 = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'T', col: 0, row: 0, value: 20 },
    ], { sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }], issues })
    void f3
    const ctx = evalCtx(f3.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    evaluateFormula("SUM('S:T'!A1 S!A1:A2)", ctx)
    expect(ctx.unsupportedFeatures.has('3d-intersection')).toBe(true)
    expect(issues.some((i) => i.feature === '3d-intersection')).toBe(true)
  })

  it('@ 3D run gates with a specific provenance marker (both sheet identities present)', () => {
    const f2 = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'T', col: 0, row: 0, value: 20 },
    ], { sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }] })
    const ctx = evalCtx(f2.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    const outcome = evaluateFormula("@'S:T'!A1", ctx)
    expect(outcome).toBe('#NAME?')
    expect(ctx.unsupportedFeatures.has('implicit-intersection-3d')).toBe(true)
  })
  it('@ 2D range (ambiguous projection) gates instead of computed silence', () => {
    const f2 = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }, { sheet: 'S', col: 0, row: 1, value: 20 }],
      {})
    const ctx = evalCtx(f2.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    evaluateFormula('@A1:B2', ctx)
    expect(ctx.unsupportedFeatures.size).toBeGreaterThan(0)
  })

  it('multi-area @ union gates (ambiguous multi-area projection)', () => {
    const f2 = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }, { sheet: 'S', col: 0, row: 1, value: 20 }], {})
    const ctx = evalCtx(f2.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    evaluateFormula('@(A1:A2,A3:A4)', ctx)
    expect(ctx.unsupportedFeatures.size).toBeGreaterThan(0)
  })

  it('3D gate through NAME nesting: qualified source walks the binding source', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const f3 = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'T', col: 0, row: 0, value: 20 },
    ], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names: [{ name: 'Ref3d', source: "'S:T'!A1", baseProvenance: 'unknown' }],
      issues,
    })
    const ctx = evalCtx(f3.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    evaluateFormula('Ref3d S!A1', ctx)
    expect(ctx.unsupportedFeatures.has('3d-intersection')).toBe(true)
    expect(issues.some((i) => i.feature === '3d-intersection')).toBe(true)
  })

  it('3D gate through UNION nesting', () => {
    const f3 = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'T', col: 0, row: 0, value: 20 },
    ], { sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }] })
    const ctx = evalCtx(f3.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    evaluateFormula("SUM((('S:T'!A1),S!A2) S!A1)", ctx)
    expect(ctx.unsupportedFeatures.has('3d-intersection')).toBe(true)
  })


  it('missing currentCell context: diagnosed API-context typed error, no fabricated context', () => {
    const f2 = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {})
    const ctx = evalCtx(f2.services, { sheet: 'S' }) // no currentCell
    ctx.unsupportedFeatures = new Set<string>()
    const r = evaluateFormula('A1:A2', ctx)
    // The outcome is a diagnosed API-context error (never an invented A1
    // context or a computed value), and the missing-cell-context capability
    // is marked specifically.
    expect(isEvaluationError(r) ? r.code : r).not.toBe(10)
    expect(ctx.unsupportedFeatures.has('projection-context-missing')).toBe(true)
  })

  it('@ ordinary vertical vector projection stays supported (no gate)', () => {
    const f2 = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }, { sheet: 'S', col: 0, row: 1, value: 20 }],
      {})
    // currentCell on the same row as A1 -> projection picks A1
    const ctx = evalCtx(f2.services, { sheet: 'S', cell: { col: 5, row: 0 } })
    expect(evaluateFormula('@A1:A2', ctx)).toBe(10)
  })

  it('live/@@ standoff: dead branches stay silent (IF(FALSE, gate-like, 8) = 8, no diagnostic)', () => {
    const f2 = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {})
    const ctx = evalCtx(f2.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    const issues: unknown[] = []
    void issues
    expect(evaluateFormula('IF(FALSE,@A1:B2,8)', ctx)).toBe(8)
    expect(ctx.unsupportedFeatures.size).toBe(0)
  })
})

describe('B1 r2 group 2 workbook cache retention (evaluated-frame validity + diagnostics)', () => {
  function doc(formula: string): XlsxDocument {
    return {
      sheets: [
        {
          name: 'S', sheetId: 's', sourcePartPath: 's',
          rows: [
            { index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 10, styleIndex: 0 } as XlsxCell] },
            { index: 1, cells: [{ ref: 'A2', col: 0, row: 1, value: 7, styleIndex: 0 } as XlsxCell] },
            { index: 2, cells: [{ ref: 'B3', col: 1, row: 2, value: 77, styleIndex: 0, hasCachedValue: true, formula } as unknown as XlsxCell] },
          ],
        },
        {
          name: 'T', sheetId: 't', sourcePartPath: 't',
          rows: [{ index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 20, styleIndex: 0 } as XlsxCell] }],
        },
      ],
      tables: [], images: [], drawingCoverage: [],
    } as unknown as XlsxDocument
  }

  it.each([
    "=@'S:T'!A1",
    '=@A1:B2',
    "=SUM('S:T'!A1 S!A1:A2)",
    '=Missing!Named',
  ])('retains valid cached 77 with diagnostics for %s', (formula) => {
    const d = doc(formula)
    evaluateWorkbookFormulas(d, { forceRecalc: true })
    const cell = (d.sheets[0].rows[2].cells as XlsxCell[])[0]
    expect(cell.value).toBe(77)
    expect((d as unknown as { diagnostics?: unknown[] }).diagnostics?.length).toBeGreaterThan(0)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 12 finals: positive domain bounds and case-insensitive sheet
// identity through the services helper path (required-not-closed tracked).
// ---------------------------------------------------------------------------
describe('B1 r2 group 12 finals: positive bounds and case identity', () => {
  it('edge-valid coordinates resolve: max col/row (16383, 1048575) is in-domain', () => {
    const f = fakeServices([{ sheet: 'S', col: 16383, row: 1048575, value: 1 }])
    const r = f.services.resolve(
      { type: 'cell', ref: { col: 16383, row: 1048575, absCol: false, absRow: false, sheet: undefined } },
      { currentSheet: 'S' },
    )
    expect(r && (r as { kind?: string }).kind === 'resolved-reference').toBe(true)
  })
  it('case-insensitive sheet identity: ROW(s!A5) resolves on sheet S', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 4, value: 1 }])
    const r = evaluateFormula('ROW(s!A5)', evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 9 } }))
    expect(r).toBe(5)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 9: standalone whole-axis aggregates never enumerate/allocate
// blank geometry; a specific unavailable-capability diagnostic is emitted
// instead (no scanning, no clamping, no million-cell allocation). ROWS/COLUMNS
// geometry stays available.
// ---------------------------------------------------------------------------
describe('B1 r2 group 9: standalone whole-axis/huge aggregates diagnose unavailable sparse population', () => {
  it('SUM(A:A) getCellValue-only: specific sparse-population-unavailable marker + internal convention result; reads bounded (instrumentation only, never a product cap)', () => {
    let reads = 0
    const ctx = {
      getCellValue: () => {
        if (++reads > 50) throw new Error('blank enumeration')
        return null
      },
    } as unknown as EvaluationContext & { unsupportedFeatures?: Set<string> }
    ctx.unsupportedFeatures = new Set<string>()
    const r = evaluateFormula('SUM(A:A)', ctx)
    // The specific capability diagnostic is asserted (a silent clamp or a
    // genuine invented #NUM! must NOT pass):
    expect(ctx.unsupportedFeatures.has('sparse-population-unavailable')).toBe(true)
    // evaluateFormula is the PUBLIC primitive API: tagged internal errors
    // surface as their code string (evaluateFormulaInternal would test the
    // tagged shape — not used here).
    expect(r).toBe('#NAME?')
    // read51 keeps throwing as instrumentation: proving no enumeration
    // happened; it is NOT a product cap.
    expect(reads).toBeLessThanOrEqual(50)
  })

  it('SUM(A1:A1048576) with a huge finite range: same diagnosed convention (no silent clamp, no invented genuine #NUM!)', () => {
    let reads = 0
    const ctx = {
      getCellValue: () => {
        if (++reads > 50) throw new Error('blank enumeration')
        return null
      },
    } as unknown as EvaluationContext & { unsupportedFeatures?: Set<string> }
    ctx.unsupportedFeatures = new Set<string>()
    const r = evaluateFormula('SUM(A1:A1048576)', ctx)
    expect(ctx.unsupportedFeatures.has('sparse-population-unavailable')).toBe(true)
    expect(r).toBe('#NAME?')
    expect(reads).toBeLessThanOrEqual(50)
  })

  it('legacy FINITE range behavior unchanged: small standalone ranges still read cells', () => {
    const values = new Map<string, number>([
      ['0:0', 10], ['0:1', 20], ['0:2', 30],
    ])
    const ctx = {
      getCellValue: (_s: string | undefined, c: number, r: number) => values.get(`${c}:${r}`) ?? null,
    } as unknown as EvaluationContext
    expect(evaluateFormula('SUM(A1:A3)', ctx)).toBe(60)
  })

  it('ROWS/COLUMNS geometry stays available without sparse capability (retention)', () => {
    const ctx = { getCellValue: () => null } as unknown as EvaluationContext
    expect(evaluateFormula('ROWS(A:A)', ctx)).toBe(1048576)
    expect(evaluateFormula('COLUMNS(1:1)', ctx)).toBe(16384)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 3: bounded, cycle-aware, cooperative defined-name traversal.
// Deep chains evaluate without stack overflow; cyclic names stay bounded (an
// honest capability gate, not a caught RangeError). Native cycle VALUE and
// initialization semantics remain UNKNOWN and are not asserted.
// ---------------------------------------------------------------------------
describe('B1 r2 group 3: cooperative name chains and bounded cycles', () => {
  it('3000-name chain stays cooperative without stack overflow (sealed probes row 33)', () => {
    const names = Array.from({ length: 3000 }, (_, i) => ({
      name: `Name_${i}`,
      source: i === 2999 ? '5' : `Name_${i + 1}`,
      baseProvenance: 'unknown' as const,
    }))
    const f = fakeServices([], { names })
    expect(() => evaluateFormula('Name_0', evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))).not.toThrow()
    expect(evaluateFormula('Name_0', evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))).toBe(5)
  })

  it('cyclic names are a bounded capability gate, not JS recursion (sealed probes row 34 + cache-hooks RangeError)', () => {
    const f = fakeServices([], { names: [{ name: 'LoopName', source: 'LoopName', baseProvenance: 'unknown' }] })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    let result: unknown
    expect(() => {
      result = evaluateFormula('LoopName', ctx)
    }).not.toThrow()
    // Bounded outcome: internal unavailable-capability convention; the
    // actual internal #NAME? result is ASSERTED (native circular
    // value/initialization semantics remain UNKNOWN).
    expect(result).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('name-cycle')).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 3 (stable binding identity): repeated bindName for the same
// stable name+scope+source+generation yields the SAME cached AST identity;
// changing generation or source refreshes properly (bounded cache, no reparse
// per retry).
// ---------------------------------------------------------------------------
describe('B1 r2 group 3: stable binding identity', () => {
  it('bindName twice on the same name/scope/source/generation returns the same AST reference', () => {
    const f = fakeServices([], { names: [{ name: 'Name_0', source: '5', baseProvenance: 'unknown' }] })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    const a = f.services.bindName({ name: 'Name_0' }, ctx)
    const b = f.services.bindName({ name: 'Name_0' }, ctx)
    if (isEvaluationError(a) || isEvaluationError(b)) throw new Error('bindName failed')
    expect((b as import('../src/xlsx/formula/types').NameBinding).ast).toBe((a as import('../src/xlsx/formula/types').NameBinding).ast)
  })

  it('changing generation or definition source refreshes the binding cache properly', () => {
    const gen = { n: 1 }
    let cells: Array<{ address: { sheetId: string; col: number; row: number }; value: ElementValue; origin: 'input' }> = []
    const services = createReferenceServices({
      generation: () => gen.n,
      sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
      sheetIdOfName: (n) => (n.toLowerCase() === 's' ? 's' : undefined),
      sheetNameOfId: () => 'S',
      definedNames: [{ name: 'Src1', source: '7', baseProvenance: 'unknown' }],
      tables: [],
      store: {
        stored: () => cells,
        read: (address) => {
          const hit = cells.find((c) => c.address.col === address.col && c.address.row === address.row)
          return hit ? { value: hit.value, origin: hit.origin } : undefined
        },
      },
    })
    const ctx = {
      references: services,
      currentSheet: 'S',
      currentCell: { col: 1, row: 1, absCol: false, absRow: false },
      flatArgs: new WeakMap(),
    } as unknown as EvaluationContext
    const a = services.bindName({ name: 'Src1' }, ctx)
    gen.n++ // new generation refreshes the cache
    const b = services.bindName({ name: 'Src1' }, ctx)
    if (isEvaluationError(a) || isEvaluationError(b)) throw new Error('bindName failed')
    expect((b as import('../src/xlsx/formula/types').NameBinding).ast).not.toBe((a as import('../src/xlsx/formula/types').NameBinding).ast)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 3 (buried-cycle/scope controls, TEST-ONLY first): bounded cycles
// buried in calls/unions; scope-suffixed traversal keys; workbook cache77 +
// dead-branch silence; false-cycle-free retry after suspension.
// ---------------------------------------------------------------------------
describe('B1 r2 group 3: buried cycles, scope keys, suspension survival', () => {
  function buriedFixture(sources: Array<{ name: string; source: string; localSheetIndex?: number }>) {
    return fakeServices(
      [
        { sheet: 'S', col: 0, row: 0, value: 10 },
        { sheet: 'T', col: 0, row: 0, value: 4 },
      ],
      {
        sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
        names: sources.map((s) => ({ ...s, baseProvenance: 'unknown' as const })),
      },
    )
  }

  it('cycle buried in a CALL (LoopSum = SUM(LoopSum)) is bounded: #NAME? + name-cycle marker, no crash/hang', () => {
    const f = buriedFixture([{ name: 'LoopSum', source: '=SUM(LoopSum)' }])
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    let r: unknown
    expect(() => {
      r = evaluateFormula('LoopSum', ctx)
    }).not.toThrow()
    expect(r).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('name-cycle')).toBe(true)
  })

  it('cycle buried in a UNION (LoopUni = SUM((LoopUni,A1))) is bounded the same way', () => {
    const f = buriedFixture([{ name: 'LoopUni', source: '=SUM((LoopUni,A1))' }])
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    let r: unknown
    expect(() => {
      r = evaluateFormula('LoopUni', ctx)
    }).not.toThrow()
    expect(r).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('name-cycle')).toBe(true)
  })

  it('scope-suffixed traversal: a chain through DIFFERENT-scoped same-spelling names does not report a false cycle', () => {
    // local-s Rel -> T!Rel (qualified to sheet T) -> T-local value 4; the two
    // 'rel' bindings live in different scopes and BOTH sheet identities are
    // walked — a scope-blind key would report a false cycle.
    const f = buriedFixture([
      { name: 'Rel', source: 'T!Rel', localSheetIndex: 0 },
      { name: 'Rel', source: '4', localSheetIndex: 1 },
    ])
    const r = evaluateFormula('Rel', evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))
    expect(r).toBe(4) // chained through the T-local scope, no false cycle
  })

  it('workbook cache77 retention with a buried cycle: cached value stays, diagnostic recorded', () => {
    const formula = '=SUM(LoopSum)'
    const doc = {
      sheets: [
        {
          name: 'S', sheetId: 's', sourcePartPath: 's',
          rows: [
            { index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 10, styleIndex: 0 } as XlsxCell] },
            { index: 1, cells: [] as XlsxCell[] },
            { index: 2, cells: [{ ref: 'B3', col: 1, row: 2, value: 77, styleIndex: 0, hasCachedValue: true, formula } as unknown as XlsxCell] },
          ],
        },
      ],
      // Genuine cycle shape requires the definition inventory (existing
      // metadata ABI): without it the unknown name is a genuine unknownName,
      // not a cycle, and must NOT force retention.
      definedNames: [{ name: 'LoopSum', source: 'SUM(LoopSum)', baseProvenance: 'unknown', namesComplete: true }],
      tables: [], images: [], drawingCoverage: [],
    } as unknown as XlsxDocument
    // The definition inventory is workbook-level; provide LoopSum via the
    // workbook API path is out of this fixture — the cycle gate still marks
    // the evaluated frame (bindName fails closed #NAME? for the unknown
    // definition) so the cached value must be retained with diagnostics.
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const cell = (doc.sheets[0].rows[2].cells as XlsxCell[])[0]
    expect(cell.value).toBe(77)
    expect((doc as unknown as { diagnostics?: unknown[] }).diagnostics?.length).toBeGreaterThan(0)
  })

  it('dead-branch silence with a buried cycle: IF(FALSE, SUM(LoopSum), 8) = 8 and no marker', () => {
    const f = buriedFixture([{ name: 'LoopSum', source: '=SUM(LoopSum)' }])
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    expect(evaluateFormula('IF(FALSE, SUM(LoopSum), 8)', ctx)).toBe(8)
    expect(ctx.unsupportedFeatures?.has('name-cycle') ?? false).toBe(false)
  })

  it('false-cycle-free retry: a name evaluation that suspends once and is retried does NOT report a false cycle', () => {
    class PendingMarker extends Error {}
    let cells: Array<{ address: { sheetId: string; col: number; row: number }; value: ElementValue; origin: 'input' }> = [
      { address: { sheetId: 's', col: 0, row: 0 }, value: 5, origin: 'input' },
    ]
    let suspends = 1
    const pendingMarker = new PendingMarker('injected pending')
    const services = createReferenceServices({
      generation: () => 1,
      sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
      sheetIdOfName: (n) => (n.toLowerCase() === 's' ? 's' : undefined),
      sheetNameOfId: () => 'S',
      definedNames: [{ name: 'CellRef', source: '$A$1', baseProvenance: 'unknown' }],
      tables: [],
      store: {
        stored: () => cells,
        read: (address: { sheetId: string; col: number; row: number }) => {
          if (address.col === 0 && address.row === 0 && suspends > 0) {
            suspends--
            throw pendingMarker
          }
          const hit = cells.find((c) => c.address.col === address.col && c.address.row === address.row)
          return hit ? { value: hit.value, origin: hit.origin } : undefined
        },
      },
    })
    const ctx = {
      references: services,
      // TEST-ONLY instrumentation: the legacy cell path is the only store.read
      // route reachable from this standalone driver; the callback throws the
      // SAME PendingMarker first, then the real value (same given memo shape).
      getCellValue: (_sheet: string | undefined, col: number, row: number) => {
        if (col === 0 && row === 0 && suspends > 0) {
          suspends--
          throw pendingMarker
        }
        return 5
      },
      currentSheet: 'S',
      currentCell: { col: 1, row: 1, absCol: false, absRow: false },
      flatArgs: new WeakMap(),
    } as unknown as EvaluationContext & { unsupportedFeatures?: Set<string> }
    ctx.unsupportedFeatures = new Set<string>()
    let caught: unknown
    try {
      evaluateFormula('CellRef', ctx)
    } catch (e) {
      caught = e
    }
    expect(caught instanceof PendingMarker).toBe(true)
    // Retry: no false name-cycle marker; value resolves.
    const r = evaluateFormula('CellRef', ctx)
    expect(r).toBe(5)
    expect(ctx.unsupportedFeatures.has('name-cycle')).toBe(false)
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 8: shared quoted-3D roundtrip (spaces/apostrophes) + unquoted
// wholeCol/wholeRow 3D targets with independent per-endpoint anchors.
// ---------------------------------------------------------------------------
describe('B1 r2 group 8: shared quoted-3D roundtrip and unquoted whole-axis targets', () => {
  it("SUM('First Sheet:Last Sheet'!$A1) shared-translation roundtrips with exact quoted-run spelling", () => {
    const translated = translateSharedFormula("SUM('First Sheet:Last Sheet'!$A1)", 1, 1)
    // exact quoted-run spelling: spaces and the joined quoted run preserved,
    // row anchor shifted $A1 -> $A2
    expect(translated.formula).toBe("SUM('First Sheet:Last Sheet'!$A2)")
    expect(parseFormula(translated.formula)).toEqual(translated.ast)
  })

  it("shared quoted sheet with apostrophe interior roundtrips with exact spelling", () => {
    const src = "SUM('O''Brien-Sheet:Fin'!$A1)"
    const translated = translateSharedFormula(src, 1, 2)
    expect(translated.formula).toBe("SUM('O''Brien-Sheet:Fin'!$A3)")
    expect(parseFormula(translated.formula)).toEqual(translated.ast)
  })

  it('unquoted wholeCol 3D target parses as a ref3d wholeCol with per-endpoint anchors', () => {
    const ast = parseFormula('SUM(First:Last!$A:C)') as { type: string; args?: unknown[] }
    expect(ast.type).toBe('call')
    const arg = (ast.args?.[0] ?? {}) as {
      type?: string
      sheets?: { fromSheet: string; toSheet: string }
      target?: { type: string; ref: { from: { index: number; absolute: boolean }; to: { index: number; absolute: boolean } } }
    }
    expect(arg.type).toBe('ref3d')
    expect(arg.sheets).toMatchObject({ fromSheet: 'First', toSheet: 'Last' })
    expect(arg.target).toMatchObject({
      type: 'wholeCol',
      ref: { from: { index: 0, absolute: true }, to: { index: 2, absolute: false } },
    })
  })

  it('unquoted wholeRow 3D target parses as a ref3d wholeRow with per-endpoint anchors', () => {
    const ast = parseFormula('SUM(First:Last!$1:3)') as { type: string; args?: unknown[] }
    expect(ast.type).toBe('call')
    const arg = (ast.args?.[0] ?? {}) as {
      type?: string
      sheets?: { fromSheet: string; toSheet: string }
      target?: { type: string; ref: { from: { index: number; absolute: boolean }; to: { index: number; absolute: boolean } } }
    }
    expect(arg.type).toBe('ref3d')
    expect(arg.sheets).toMatchObject({ fromSheet: 'First', toSheet: 'Last' })
    expect(arg.target).toMatchObject({
      type: 'wholeRow',
      ref: { from: { index: 0, absolute: true }, to: { index: 2, absolute: false } },
    })
  })

  it('shared translation AND reparse shift whole-axis 3D targets per endpoint anchors', () => {
    // $A:C with deltaCol 1 => $A:D (from index 0 absolute TRUE; to index 3 absolute FALSE)
    const t1 = translateSharedFormula('SUM(First:Last!$A:C)', 1, 2)
    const a1 = parseFormula(t1.formula) as {
      args?: Array<{ target?: { ref?: { from?: { index: number; absolute: boolean }; to?: { index: number; absolute: boolean } } } }>
    }
    expect(t1.formula).toContain('!$A:D')
    expect(a1.args?.[0]?.target?.ref?.from).toEqual({ index: 0, absolute: true })
    expect(a1.args?.[0]?.target?.ref?.to).toEqual({ index: 3, absolute: false })
    // $1:3 with deltaRow 2 => $1:5
    const t2 = translateSharedFormula('SUM(First:Last!$1:3)', 2, 2)
    const a2 = parseFormula(t2.formula) as {
      args?: Array<{ target?: { type?: string; ref?: { from?: { index: number; absolute: boolean }; to?: { index: number; absolute: boolean } } } }>
    }
    expect(t2.formula).toContain('!$1:5')
    expect(a2.args?.[0]?.target?.ref?.from).toEqual({ index: 0, absolute: true })
    expect(a2.args?.[0]?.target?.ref?.to).toEqual({ index: 4, absolute: false })
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 6: table bracket escape-aware scanning and @ shorthand this-row
// ---------------------------------------------------------------------------
describe('B1 r2 group 6: table escapes and @ shorthand this-row', () => {
  const roundTrip = (src: string, wantTable: Record<string, unknown>) => {
    it(`${src}`, () => {
      const ast = parseFormula(src) as { type: string; args?: unknown[]; ref?: unknown }
      expect(ast.type).not.toBe('error')
      // Drill into the table argument and match the exact decode contract.
      const arg = (ast.type === 'call' ? (ast.args?.[0] ?? {}) : ast) as Record<string, unknown>
      const ref = arg.ref as Record<string, unknown>
      for (const key of Object.keys(wantTable)) {
        expect(JSON.stringify((ref ?? arg)[key])).toBe(JSON.stringify(wantTable[key]))
      }
    })
  }

  describe('escaped header characters decode', () => {
    roundTrip("SUM(Table1['#Tag])", { items: ['data'], columns: { kind: 'single', name: '#Tag' } })
    roundTrip("Table1['@Rate]", { items: ['data'], columns: { kind: 'single', name: '@Rate' } })
    roundTrip("Table1[A']B]", { items: ['data'], columns: { kind: 'single', name: 'A]B' } })
    roundTrip("Table1[A'[B]", { items: ['data'], columns: { kind: 'single', name: 'A[B' } })
    // @ shorthand this-row: Table1[@Amount] == [[#This Row],[Amount]]
    roundTrip("SUM(Table1[@Amount])", { items: ['thisRow'], columns: { kind: 'single', name: 'Amount' } })
    // doubled apostrophe inside the name decodes to ONE literal apostrophe
    roundTrip("Table1[A''B]", { items: ['data'], columns: { kind: 'single', name: "A'B" } })
    it('raw source preservation: table token keeps the exact verbatim source', () => {
      const src = "Table1[A']B]" // apostrophe-only escape (actual syntax)
      const ast = parseFormula(src) as {
        type: string
        ref?: { raw?: string; table?: string; items?: string[]; columns?: { kind: string; name: string } }
      }
      expect(ast.type).toBe('table')
      expect(ast.ref?.raw).toBe(src) // exact TableSyntax.raw contract
      // table copy immutability: shared translation leaves a table operand
      // verbatim (no row anchors apply inside the table)
      const translated = translateSharedFormula(src, 1, 2)
      expect(translated.formula).toBe(src)
      expect(parseFormula(translated.formula)).toMatchObject(ast as object)
    })
    it('raw prefix preservation: quoted-# table raw token verbatim + immutable copy', () => {
      const src = "SUM(Table1['#Tag])"
      const ast = parseFormula(src) as { type: string; args?: Array<{ ref?: { raw?: string } }> }
      expect(ast.type).toBe('call')
      expect(ast.args?.[0]?.ref?.raw).toBe("Table1['#Tag]")
      const translated = translateSharedFormula(src, 1, 2)
      expect(translated.formula).toBe(src)
    })
    it('raw prefix preservation: @ shorthand raw token verbatim + immutable copy', () => {
      const src = 'SUM(Table1[@Amount])'
      const ast = parseFormula(src) as { type: string; args?: Array<{ ref?: { raw?: string } }> }
      expect(ast.type).toBe('call')
      expect(ast.args?.[0]?.ref?.raw).toBe('Table1[@Amount]')
      const translated = translateSharedFormula(src, 1, 2)
      expect(translated.formula).toBe(src)
    })
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 3 union bridge: TEST-ONLY full controls (workbook actual
// definition inventory, cache77/dead-branch/scope-alias/suspension provided
// memo/ordinary names).
// ---------------------------------------------------------------------------
describe('B1 r2 group 3 union bridge controls', () => {
  it('workbook union-hidden cycle with real definition inventory: cache77 stays + name-cycle diagnostic (not CALL-only)', () => {
    // ROOT AUDIT: cached formula is the OWNING name =LoopUni so the driver
    // ENTERS the name frame before the union member lookup.
    const formula = '=LoopUni'
    const doc = {
      sheets: [
        {
          name: 'S', sheetId: 's', sourcePartPath: 's',
          rows: [
            { index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 10, styleIndex: 0 } as XlsxCell] },
            { index: 1, cells: [] as XlsxCell[] },
            { index: 2, cells: [{ ref: 'B3', col: 1, row: 2, value: 77, styleIndex: 0, hasCachedValue: true, formula } as unknown as XlsxCell] },
          ],
        },
      ],
      definedNames: [{ name: 'LoopUni', source: 'SUM((LoopUni,A1))', baseProvenance: 'unknown', namesComplete: true }],
      tables: [], images: [], drawingCoverage: [],
    } as unknown as XlsxDocument
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const cell = (doc.sheets[0].rows[2].cells as XlsxCell[])[0]
    expect(cell.value).toBe(77)
    const diags = (doc as unknown as { diagnostics?: Array<{ feature?: string }> }).diagnostics ?? []
    expect(diags.some((d) => d.feature === 'name-cycle')).toBe(true)
  })

  it('dead branch IF(FALSE,SUM((LoopUni,A1)),8) = 8 with NO union cycle mark', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], { names: [{ name: 'LoopUni', source: 'SUM((LoopUni,A1))', baseProvenance: 'unknown' }] })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    expect(evaluateFormula('IF(FALSE, SUM((LoopUni,A1)), 8)', ctx)).toBe(8)
    expect(ctx.unsupportedFeatures.has('name-cycle')).toBe(false)
  })

  it('same SELECTED definition identity: qualified/unqualified/case spellings share ONE binding source AST (no false cycle), provided real WeakMap memo untouched', () => {
    const f = fakeServices(
      [ { sheet: 'S', col: 0, row: 0, value: 10 }, { sheet: 'T', col: 0, row: 0, value: 4 } ],
      {
        sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
        names: [
          { name: 'Rel', source: 'T!rel', localSheetIndex: 0, baseProvenance: 'unknown' },
          { name: 'Rel', source: '4', localSheetIndex: 1, baseProvenance: 'unknown' },
        ],
      },
    )
    const provided = new WeakMap<AstNode, unknown>()
    const sentinel = parseFormula('4') as AstNode
    provided.set(sentinel, 42) // sentinel AST entry preserved across replay
    const ctx = {
      references: f.services,
      currentSheet: 'S',
      currentCell: { col: 1, row: 1, absCol: false, absRow: false },
      flatArgs: new WeakMap(),
      nodeValues: provided,
    } as unknown as EvaluationContext & { nodeValues: WeakMap<AstNode, unknown> }
    // qualified + case variants of the T-local definition share its identity
    const b1 = f.services.bindName({ name: 'Rel', sheet: 'T' }, ctx)
    const b2 = f.services.bindName({ name: 'rel', sheet: 'T' }, ctx)
    expect(isEvaluationError(b1) || isEvaluationError(b2)).toBe(false)
    if (!isEvaluationError(b1) && !isEvaluationError(b2)) {
      expect(b2.ast).toBe(b1.ast) // same selected definition -> SAME cached AST identity
    }
    expect(evaluateFormula('Rel', ctx)).toBe(4) // no false cycle through case/qualifier spellings
    expect(evaluateFormula('rel', ctx)).toBe(4)
    expect(ctx.nodeValues).toBe(provided) // memo ownership untouched
    expect(provided.get(sentinel)).toBe(42) // sentinel entry preserved after replay
  })

  it('ordinary value names stay ordinary: Five=5 / reference R=$A$1 read, no false name-cycle marker', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 5 }], {
      names: [
        { name: 'Five', source: '5', baseProvenance: 'unknown' },
        { name: 'R', source: '$A$1', baseProvenance: 'unknown' },
      ],
    })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    expect(evaluateFormula('Five', ctx)).toBe(5)
    expect(evaluateFormula('R', ctx)).toBe(5)
    expect(evaluateFormula('SUM(R,R)', evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } }))).toBe(10)
    expect(ctx.unsupportedFeatures.has('name-cycle')).toBe(false)
  })

  it('suspension chain: PendingName=SUM(R,LoopUni); first R read throws SAME marker, retry reaches the ACTUAL union cycle (LoopUni=SUM((LoopUni,A1))); real WeakMap + seed-before/after', () => {
    class PendingU extends Error {}
    let suspends = 1 // ONE actual suspension: the FIRST A1 read throws the marker
    const pendingMarker = new PendingU('injected pending union case')
    // cells: A1=5; store.read for A1 suspends ONCE then returns the value
    let cells: StoredCell[] = [
      { address: { sheetId: 's', col: 0, row: 0 }, value: 15 as ElementValue, origin: 'input' as const },
    ]
    const services = createReferenceServices({
      generation: () => 1,
      sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
      sheetIdOfName: (n) => (n.toLowerCase() === 's' ? 's' : undefined),
      sheetNameOfId: () => 'S',
      definedNames: [
        { name: 'R', source: '$A$1', baseProvenance: 'unknown' },
        { name: 'LoopUni', source: 'SUM((LoopUni,A1))', baseProvenance: 'unknown' },
        { name: 'PendingName', source: 'SUM(R,LoopUni)', baseProvenance: 'unknown' },
      ],
      tables: [],
      store: {
        stored: () => cells,
        read: (address: { sheetId: string; col: number; row: number }) => {
          if (address.col === 0 && address.row === 0) {
            if (suspends > 0) { suspends--; throw pendingMarker }
            const hit = cells.find((c) => c.address.col === address.col && c.address.row === address.row)
            return hit ? { value: hit.value, origin: hit.origin } : undefined
          }
          return undefined
        },
      },
    })

    // real WeakMap caller memo, seeded BEFORE evaluation with a sentinel AST
    const provided = new WeakMap<AstNode, unknown>()
    const sentinel = parseFormula('5')
    provided.set(sentinel as AstNode, 42)
    const ctx = {
      references: services,
      currentSheet: 'S',
      currentCell: { col: 1, row: 1, absCol: false, absRow: false },
      flatArgs: new WeakMap(),
      nodeValues: provided,
    } as unknown as EvaluationContext & { unsupportedFeatures?: Set<string> }
    ctx.unsupportedFeatures = new Set<string>()
    let caught: unknown
    try {
      evaluateFormula('PendingName', ctx)
    } catch (e) {
      caught = e
    }
    expect(caught).toBe(pendingMarker) // FIRST A1 read suspends with exact identity
    // retry: pending gone; the union member LoopUni re-enters the ACTIVE
    // binding (LoopUni is on the evaluating path) and the bounded cycle gate
    // fires (native circular value unmeasured; internal #NAME? convention).
    expect(evaluateFormula('PendingName', ctx)).toBe('#NAME?')
    expect(ctx.unsupportedFeatures.has('name-cycle')).toBe(true)
    expect(ctx.nodeValues).toBe(provided)      // memo identity kept
    expect(provided.get(sentinel as AstNode)).toBe(42) // sentinel preserved AFTER replay
  })
})

// ---------------------------------------------------------------------------
// B1 r2 group 3 deep3D (DEEP-3D.md counterexamples + reviewed cross-module ref
// finding): discovery reaches 3D through deep alias chains / same-spelling
// scopes / union second siblings, without RangeError and without spelling
// heuristics. Ordinary SUM 3D aggregates and @ vectors stay positive.
// ---------------------------------------------------------------------------
describe('B1 r2 group 3 deep3D discovery controls', () => {
  it('3000-name absolute chain ending in a 3D sheet run reaches the intersection gate without RangeError', () => {
    const names = Array.from({ length: 3000 }, (_, i) => ({
      name: `N_${i}`,
      source: i === 2999 ? "'S:T'!$A$1" : `N_${i + 1}`,
      baseProvenance: 'unknown' as const,
    }))
    const issues: Array<{ feature: string; message: string }> = []
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names,
      issues,
    })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    let r: unknown
    expect(() => {
      r = evaluateFormula('SUM(N_0 S!$A$1)', ctx)
    }).not.toThrow()
    expect(r).toBe('#NAME?')
    expect(ctx.unsupportedFeatures.has('3d-intersection')).toBe(true)
    expect(issues.some((i) => i.feature === '3d-intersection')).toBe(true)
  })

  it('3000-name absolute chain ending on SHEET S only: no 3D gate (detector-only control)', () => {
    const names = Array.from({ length: 3000 }, (_, i) => ({
      name: `M_${i}`,
      source: i === 2999 ? 'S!$A$1' : `M_${i + 1}`,
      baseProvenance: 'unknown' as const,
    }))
    const issues: Array<{ feature: string; message: string }> = []
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {
      names,
      issues,
    })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    // Observe discovery reaching the final binding; full alias-intersection
    // cooperation is outside this detector-only control.
    let finalSourceReads = 0
    Object.defineProperty(names[2999], 'source', {
      get: () => { finalSourceReads++; return 'S!$A$1' },
    })
    evaluateFormula('SUM(M_0 S!$A$1)', ctx)
    expect(finalSourceReads).toBeGreaterThan(0)
    expect(evaluateFormula('SUM(S!$A$1 S!$A$1)', ctx)).toBe(10)
    expect(ctx.unsupportedFeatures.has('3d-intersection')).toBe(false)
    expect(issues.some((i) => i.feature === '3d-intersection')).toBe(false)
  })

  it('same-spelling DEREF-scoped names: S-local Rel → T-local Rel (3D source) is discovered, not skipped by spelling', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names: [
        { name: 'Rel', source: 'T!Rel', localSheetIndex: 0, baseProvenance: 'unknown' },
        { name: 'Rel', source: "'S:T'!$A$1", localSheetIndex: 1, baseProvenance: 'unknown' },
      ],
      issues,
    })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    evaluateFormula('SUM(Rel S!$A$1)', ctx)
    expect(ctx.unsupportedFeatures.has('3d-intersection')).toBe(true)
    expect(issues.some((i) => i.feature === '3d-intersection')).toBe(true)
  })

  it('union with same-spelling SIBLINGS: SECOND binding (T-local 3D twin) is discovered', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names: [
        { name: 'Twin', source: 'S!$A$1', localSheetIndex: 0, baseProvenance: 'unknown' },
        { name: 'Twin', source: "'S:T'!$A$1", localSheetIndex: 1, baseProvenance: 'unknown' },
      ],
      issues,
    })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    evaluateFormula('SUM((S!Twin,T!Twin) S!$A$1)', ctx)
    expect(ctx.unsupportedFeatures.has('3d-intersection')).toBe(true)
    expect(issues.some((i) => i.feature === '3d-intersection')).toBe(true)
  })

  it('live vs dead versus cache77: intersection gate marks live frame; IF(FALSE,…,8) silent; workbook retained 77', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names: [{ name: 'DeepN', source: "'S:T'!$A$1", baseProvenance: 'unknown' }],
    })
    const liveCtx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    liveCtx.unsupportedFeatures = new Set<string>()
    evaluateFormula('SUM(DeepN S!$A$1)', liveCtx)
    expect(liveCtx.unsupportedFeatures.has('3d-intersection')).toBe(true)
    const deadCtx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    deadCtx.unsupportedFeatures = new Set<string>()
    expect(evaluateFormula('IF(FALSE, SUM(DeepN S!$A$1), 8)', deadCtx)).toBe(8)
    expect(deadCtx.unsupportedFeatures.has('3d-intersection')).toBe(false)
    // workbook actual definition inventory with valid cached 77 retains the cache
    const doc: XlsxDocument = {
      sheets: [
        {
          name: 'S', sheetId: 's', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], sourcePartPath: 's',
          rows: [
            { index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 10, styleIndex: 0 }] },
            { index: 2, cells: [{ ref: 'B3', col: 1, row: 2, value: 77, styleIndex: 0, hasCachedValue: true, formula: '=SUM(DeepN S!$A$1)' }] },
          ],
        },
        { name: 'T', sheetId: 't', workbookIndex: 1, rows: [], cols: [], merges: [], mergeRanges: [] },
      ],
      definedNames: [{ name: 'DeepN', source: "'S:T'!$A$1", baseProvenance: 'unknown' }],
      tables: [], images: [], drawingCoverage: [],
    }
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const cell = doc.sheets[0].rows[1].cells[0]
    expect(cell.value).toBe(77)
    const diags = doc.diagnostics ?? []
    expect(diags.some((d) => d.feature === '3d-intersection')).toBe(true)
  })

  it('ACTUAL @ name operand gates 3D discovery: @DeepN marks implicitly 3D before projection; @A1:A2 vertical stays supported', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const f = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'S', col: 0, row: 1, value: 20 },
    ], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names: [{ name: 'DeepN', source: "'S:T'!$A$1", baseProvenance: 'unknown' }],
      issues,
    })
    const atnCtx = evalCtx(f.services, { sheet: 'S', cell: { col: 5, row: 0 } })
    atnCtx.unsupportedFeatures = new Set<string>()
    atnCtx.reportFormulaIssue = (issue) => issues.push({ feature: issue.feature ?? '', message: issue.message })
    expect(evaluateFormula('@DeepN', atnCtx)).toBe('#NAME?')
    expect(atnCtx.unsupportedFeatures.has('implicit-intersection-3d')).toBe(true)
    expect(issues.some((i) => i.feature === 'implicit-intersection-3d')).toBe(true)
    const vecCtx = evalCtx(f.services, { sheet: 'S', cell: { col: 5, row: 0 } })
    expect(evaluateFormula('@A1:A2', vecCtx)).toBe(10)
  })

  it('ACTUAL @ union operand gates 3D discovery: @((S!Twin,T!Twin)) marks specifically', () => {
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names: [
        { name: 'Twin', source: 'S!$A$1', localSheetIndex: 0, baseProvenance: 'unknown' },
        { name: 'Twin', source: "'S:T'!$A$1", localSheetIndex: 1, baseProvenance: 'unknown' },
      ],
    })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 1, row: 1 } })
    ctx.unsupportedFeatures = new Set<string>()
    const issues: string[] = []
    ctx.reportFormulaIssue = (issue) => issues.push(issue.feature ?? '')
    expect(evaluateFormula('@(S!Twin,T!Twin)', ctx)).toBe('#NAME?')
    expect(ctx.unsupportedFeatures.has('implicit-intersection-3d')).toBe(true)
    expect(issues).toContain('implicit-intersection-3d')
  })

  it('a cyclic first union sibling does not hide a later 3D sibling', () => {
    const issues: Array<{ feature: string; message: string }> = []
    const f = fakeServices([], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names: [{ name: 'Cycle', source: 'Cycle', baseProvenance: 'unknown' }],
      issues,
    })
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 0, row: 0 } })
    expect(evaluateFormula("SUM((Cycle,'S:T'!$A$1) S!$A$1)", ctx)).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('3d-intersection')).toBe(true)
    expect(issues.some((issue) => issue.feature === '3d-intersection')).toBe(true)
  })

  it('actual @ direct and deep-name operands gate before reads; dead @ and value-name branches stay silent; ordinary SUM 3D stays supported', () => {
    const reads = { count: 0 }
    const issues: string[] = []
    const f = fakeServices([
      { sheet: 'S', col: 0, row: 0, value: 10 },
      { sheet: 'T', col: 0, row: 0, value: 20 },
    ], {
      sheets: [{ sheetId: 'S', name: 'S', workbookIndex: 0 }, { sheetId: 'T', name: 'T', workbookIndex: 1 }],
      names: [
        ...Array.from({ length: 3000 }, (_, i) => ({ name: `At_${i}`, source: i === 2999 ? "'S:T'!$A$1" : `At_${i + 1}`, baseProvenance: 'unknown' as const })),
        { name: 'Quiet', source: "IF(FALSE,@'S:T'!$A$1,8)", baseProvenance: 'unknown' },
      ],
      reads,
    })
    for (const formula of ["@'S:T'!$A$1", '@At_0']) {
      const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 0, row: 0 } })
      ctx.reportFormulaIssue = (issue) => issues.push(issue.feature ?? '')
      expect(evaluateFormula(formula, ctx)).toBe('#NAME?')
      expect(ctx.unsupportedFeatures?.has('implicit-intersection-3d')).toBe(true)
    }
    expect(issues).toEqual(['implicit-intersection-3d', 'implicit-intersection-3d'])
    expect(reads.count).toBe(0)
    for (const formula of ['IF(FALSE,@At_0,8)', 'Quiet']) {
      const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 0, row: 0 } })
      ctx.reportFormulaIssue = (issue) => issues.push(issue.feature ?? '')
      expect(evaluateFormula(formula, ctx)).toBe(8)
      expect(ctx.unsupportedFeatures?.has('implicit-intersection-3d') ?? false).toBe(false)
    }
    expect(issues).toHaveLength(2)
    expect(reads.count).toBe(0)
    expect(evaluateFormula("SUM('S:T'!$A$1)", evalCtx(f.services))).toBe(30)
  })

  it.each(["@'S:T'!$A$1", '@DeepN', '@(S!Twin,T!Twin)'])('workbook actual %s retains cache77 with specific 3D issue and dead branch8', (formula) => {
    const live: XlsxCell = { ref: 'B3', col: 1, row: 2, value: 77, styleIndex: 0, hasCachedValue: true, formula }
    const dead: XlsxCell = { ref: 'B4', col: 1, row: 3, value: 77, styleIndex: 0, hasCachedValue: true, formula: `IF(FALSE,${formula},8)` }
    const doc: XlsxDocument = {
      sheets: [
        { name: 'S', sheetId: 's', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [{ index: 2, cells: [live] }, { index: 3, cells: [dead] }] },
        { name: 'T', sheetId: 't', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows: [] },
      ],
      definedNames: [
        { name: 'DeepN', source: "'S:T'!$A$1", baseProvenance: 'unknown' },
        { name: 'Twin', source: 'S!$A$1', localSheetIndex: 0, baseProvenance: 'unknown' },
        { name: 'Twin', source: "'S:T'!$A$1", localSheetIndex: 1, baseProvenance: 'unknown' },
      ],
      tables: [],
    }
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(live.value).toBe(77)
    expect(dead.value).toBe(8)
    expect(doc.diagnostics?.some((issue) => issue.feature === 'implicit-intersection-3d')).toBe(true)
    const deadDoc: XlsxDocument = { ...doc, diagnostics: [], sheets: [{ ...doc.sheets[0], rows: [{ index: 3, cells: [{ ...dead, value: 77 }] }] }, doc.sheets[1]] }
    evaluateWorkbookFormulas(deadDoc, { forceRecalc: true })
    expect(deadDoc.sheets[0].rows[0].cells[0].value).toBe(8)
    expect(deadDoc.diagnostics).toEqual([])
    expect(doc.diagnostics?.some((issue) => issue.feature === 'projection-ambiguous')).toBe(false)
  })

  it.each(['@A1', '@A1:A2', '@R', 'A1:A2'])('scalar %s prepares before reads and preserves pending identity on retry', (formula) => {
    const prepareSpy = { called: false }
    const f = fakeServices([{ sheet: 'S', col: 0, row: 0, value: 10 }], {
      names: [{ name: 'R', source: '$A$1', baseProvenance: 'unknown' }],
      prepareSpy,
      gateReadsOnPrepare: true,
    })
    const pending = new Error('pending scalar preparation')
    let suspends = true
    f.services.prepare = () => {
      if (suspends) { suspends = false; throw pending }
      prepareSpy.called = true
    }
    const ctx = evalCtx(f.services, { sheet: 'S', cell: { col: 0, row: 0 } })
    let caught: unknown
    try { evaluateFormula(formula, ctx) } catch (error) { caught = error }
    expect(caught).toBe(pending)
    expect(prepareSpy.called).toBe(false)
    expect(evaluateFormula(formula, ctx)).toBe(10)
    expect(prepareSpy.called).toBe(true)
  })
})
