/**
 * C2 phase1 arrays RED — literal matrix grammar, scalar/array API + modes,
 * singleton broadcast, the six finite array functions, and returned
 * reference/matrix consumption.
 *
 * Native-measured control (root A-M1): {1,2}+{10;20;30} is a 3x2 matrix with
 * row-major 11,12/21,22/31,32 while the scalar API stays 11. Documented
 * profiles come from primary Microsoft function pages (root FACTS.json):
 * FILTER Boolean-convertible include / empty #CALC! / structured sources;
 * SORT same-shape, sort_index default1, order1/-1, by_col rows/columns, table
 * sources; SORTBY one-row/one-column keys, order1/-1, multiple key/order
 * pairs, table sources; UNIQUE by_col/exactly_once flags, no approximate mode.
 * Unmeasured non-singleton mismatch stays an honest capability gate, never an
 * invented #VALUE!/padding, and no tie-order or mixed-type sort profile is
 * asserted (both explicitly unsettled in FACTS.json).
 */
import { describe, it, expect } from 'vitest'
import {
  evaluateFormula,
  evaluateFormulaInternal,
  formulaError,
  isEvaluationError,
} from '../src/xlsx/formula/evaluator'
import { parseFormula } from '../src/xlsx/formula/parser'
import { createReferenceServices, type StoredCell } from '../src/xlsx/formula/refs'
import type { AstNode, DefinedNameMetadata, ElementValue, EvaluationContext, EvaluationValue, TableMetadata } from '../src/xlsx/formula/types'
import * as formulaBarrel from '../src/xlsx/formula'

interface MatrixLike {
  kind: 'formula-matrix'
  rows: number
  cols: number
  values: readonly (readonly ElementValue[])[]
}
type ArrayEval = (input: string | AstNode, ctx?: EvaluationContext) => unknown
const maybeArrayEval = (formulaBarrel as unknown as { evaluateArrayFormula?: ArrayEval }).evaluateArrayFormula

function evalArray(input: string | AstNode, ctx?: EvaluationContext): unknown {
  if (typeof maybeArrayEval !== 'function') throw new Error('evaluateArrayFormula export is missing')
  return maybeArrayEval(input, ctx)
}
function asMatrix(value: unknown): MatrixLike {
  if (typeof value !== 'object' || value === null || (value as { kind?: string }).kind !== 'formula-matrix') {
    throw new Error(`expected matrix, received ${JSON.stringify(value)}`)
  }
  return value as MatrixLike
}
function plain(matrix: MatrixLike): ElementValue[][] {
  return matrix.values.map((row) => [...row])
}

interface Fixture {
  ctx: EvaluationContext
  issues: string[]
  marker: Error
  next: () => void
}

function gridFixture(
  grid: ElementValue[][],
  opts?: { pendingAt?: { col: number; row: number }; tables?: TableMetadata[]; definedNames?: readonly DefinedNameMetadata[]; currentCell?: { col: number; row: number } },
): Fixture {
  const cells: StoredCell[] = []
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      cells.push({ address: { sheetId: 's', col: c, row: r }, value: grid[r][c], origin: 'input' })
    }
  }
  let generation = 1
  let pending = !!opts?.pendingAt
  const marker = new Error('array pending')
  const issues: string[] = []
  const services = createReferenceServices({
    generation: () => generation,
    sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
    sheetIdOfName: (name) => (name.toLowerCase() === 's' ? 's' : undefined),
    sheetNameOfId: () => 'S',
    definedNames: opts?.definedNames ?? [],
    tables: opts?.tables,
    store: {
      stored: function* () {
        for (const cell of cells) yield cell
      },
      read: (address) => {
        if (pending && address.col === opts?.pendingAt?.col && address.row === opts?.pendingAt?.row) {
          pending = false
          throw marker
        }
        const hit = cells.find((x) => x.address.col === address.col && x.address.row === address.row)
        return hit ? { value: hit.value, origin: hit.origin } : undefined
      },
    },
    onIssue: (feature) => issues.push(feature),
  })
  const ctx: EvaluationContext = {
    typedValues: true,
    references: services,
    currentSheet: 'S',
    currentAddress: { sheetId: 's', col: opts?.currentCell?.col ?? 9, row: opts?.currentCell?.row ?? 0 },
    currentCell: { col: opts?.currentCell?.col ?? 9, row: opts?.currentCell?.row ?? 0, absCol: false, absRow: false },
    nodeValues: new WeakMap(),
    flatArgs: new WeakMap(),
    functionWork: new WeakMap(),
    getCellValue: (_sheet, col, row) => grid[row]?.[col] ?? null,
    reportFormulaIssue: (issue) => issues.push(issue.feature ?? issue.kind),
  }
  return { ctx, issues, marker, next: () => { generation++ } }
}

// A1:A3 keys, B1:B3 values, C1:C3 labels (row 0 is the A1 row).
const lookupGrid: ElementValue[][] = [
  [1, 10, 11],
  [2, 20, 21],
  [3, 30, 31],
]
const amountTable: TableMetadata = {
  id: 't1',
  name: 'T',
  displayName: 'T',
  sheetId: 's',
  partPath: 'xl/tables/table1.xml',
  extent: { sheetId: 's', firstCol: 0, firstRow: 0, cols: 2, rows: 4 },
  headerRowCount: 1,
  totalsRowCount: 0,
  columns: [
    { id: 'c1', name: 'Item', index: 0 },
    { id: 'c2', name: 'Amount', index: 1 },
  ],
}
const tableGrid: ElementValue[][] = [
  ['Item', 'Amount'],
  ['a', 30],
  ['b', 10],
  ['c', 20],
]

describe('C2 phase1 rectangular literal matrix grammar', () => {
  it('parses {1,2;3,4} as a 2x2 grammar-only arrayConst', () => {
    const node = parseFormula('{1,2;3,4}')
    expect(node.type).toBe('arrayConst')
    if (node.type !== 'arrayConst') return
    expect(node.rows).toEqual([
      [1, 2],
      [3, 4],
    ])
  })

  it('accepts strings, booleans and typed error literals without registry metadata', () => {
    const node = parseFormula('{1,"a",TRUE;4,"b",FALSE}')
    expect(node.type).toBe('arrayConst')
    if (node.type !== 'arrayConst') return
    expect(node.rows).toEqual([
      [1, 'a', true],
      [4, 'b', false],
    ])
    const errors = parseFormula('{#N/A,1}')
    expect(errors.type).toBe('arrayConst')
    if (errors.type !== 'arrayConst') return
    const first = errors.rows[0][0]
    expect(isEvaluationError(first)).toBe(true)
    if (isEvaluationError(first)) expect(first.code).toBe('#N/A')
  })

  it('keeps quoted commas and semicolons inside string elements', () => {
    const node = parseFormula('{"a,b";"c;d"}')
    expect(node.type).toBe('arrayConst')
    if (node.type !== 'arrayConst') return
    expect(node.rows).toEqual([['a,b'], ['c;d']])
  })

  it.each([
    ['ragged rows', '{1,2;3}'],
    ['nested array', '{1,{2,3}}'],
    ['omitted cell', '{1,,2}'],
    ['empty literal', '{}'],
    ['trailing separator', '{1,2;}'],
    ['unclosed literal', '{1,2'],
  ])('rejects %s without invented zero fill', (_label, source) => {
    expect(parseFormula(source).type).toBe('error')
  })
})

describe('C2 phase1 public scalar/array API', () => {
  it('returns a rectangular MatrixValue from evaluateArrayFormula', () => {
    const matrix = asMatrix(evalArray('{1,2;3,4}'))
    expect(matrix.rows).toBe(2)
    expect(matrix.cols).toBe(2)
    expect(plain(matrix)).toEqual([
      [1, 2],
      [3, 4],
    ])
  })

  it('keeps the public scalar API at the primitive top-left element', () => {
    expect(evaluateFormula('{1,2;3,4}')).toBe(1)
  })

  it('lifts a successful scalar to 1x1', () => {
    const matrix = asMatrix(evalArray('42'))
    expect(matrix.rows).toBe(1)
    expect(matrix.cols).toBe(1)
    expect(plain(matrix)).toEqual([[42]])
  })

  it('keeps a root error tagged through the array API', () => {
    const result = evalArray('1/0')
    expect(isEvaluationError(result)).toBe(true)
    if (isEvaluationError(result)) expect(result.code).toBe('#DIV/0!')
  })

  it('measured singleton broadcast {1,2}+{10;20;30} is 3x2 while scalar stays 11', () => {
    const fixture = gridFixture([])
    const matrix = asMatrix(evalArray('{1,2}+{10;20;30}', fixture.ctx))
    expect(matrix.rows).toBe(3)
    expect(matrix.cols).toBe(2)
    expect(plain(matrix)).toEqual([
      [11, 12],
      [21, 22],
      [31, 32],
    ])
    expect(evaluateFormula('{1,2}+{10;20;30}')).toBe(11)
  })

  it('honest-gates an unmeasured non-singleton mismatch instead of inventing VALUE/padding', () => {
    const fixture = gridFixture([])
    const result = evalArray('{1,2}+{1,2,3}', fixture.ctx)
    expect(isEvaluationError(result)).toBe(true)
    expect(fixture.ctx.unsupportedFeatures?.has('array-shape-mismatch-unverified')).toBe(true)
  })

  it('applies comparison operators elementwise in array mode', () => {
    const matrix = asMatrix(evalArray('{1,2;3,4}>2'))
    expect(plain(matrix)).toEqual([
      [false, false],
      [true, true],
    ])
  })

  it('applies concatenation elementwise with a scalar', () => {
    const matrix = asMatrix(evalArray('{1,2}&"x"'))
    expect(plain(matrix)).toEqual([['1x', '2x']])
  })

  it('applies unary operators elementwise in array mode', () => {
    const matrix = asMatrix(evalArray('-{1,2;3,4}'))
    expect(plain(matrix)).toEqual([
      [-1, -2],
      [-3, -4],
    ])
  })

  it('explicit @ projects a matrix to its top-left in both modes', () => {
    expect(evaluateFormula('@{1,2;3,4}')).toBe(1)
    const matrix = asMatrix(evalArray('@{1,2;3,4}'))
    expect(matrix.rows).toBe(1)
    expect(matrix.cols).toBe(1)
    expect(plain(matrix)).toEqual([[1]])
  })
})

describe('C2 phase1 six finite array functions', () => {
  it.each([
    ['SEQUENCE(2,3)', [[1, 2, 3], [4, 5, 6]]],
    ['SEQUENCE(3,1,10,10)', [[10], [20], [30]]],
    ['SEQUENCE(1)', [[1]]],
    ['SEQUENCE(2,2,0,5)', [[0, 5], [10, 15]]],
    ['TRANSPOSE({1,2;3,4})', [[1, 3], [2, 4]]],
    ['UNIQUE({1;2;2;3})', [[1], [2], [3]]],
    ['UNIQUE({1;2;2;3},FALSE,TRUE)', [[1], [3]]],
    ['UNIQUE({1,2,2,3},TRUE)', [[1, 2, 3]]],
    ['UNIQUE({1,2,2,3},TRUE,TRUE)', [[1, 3]]],
    ['SORT({3;1;2})', [[1], [2], [3]]],
    ['SORT({3;1;2},1,-1)', [[3], [2], [1]]],
    ['SORT({3,1,2},1,1,TRUE)', [[1, 2, 3]]],
    ['SORTBY({1;2;3},{30;10;20})', [[2], [3], [1]]],
    ['SORTBY({1;2;3},{30;10;20},-1)', [[1], [3], [2]]],
    ['SORTBY({1;2;3},{2;1;1},1,{3;2;1},-1)', [[2], [3], [1]]],
    ['FILTER({1;2;3},{TRUE;FALSE;TRUE})', [[1], [3]]],
    ['FILTER({1;2;3},{1;0;1})', [[1], [3]]],
    ['FILTER({1,2,3},{TRUE,FALSE,TRUE})', [[1, 3]]],
  ] as const)('%s returns the documented rectangle', (source, expected) => {
    const matrix = asMatrix(evalArray(source))
    expect(plain(matrix)).toEqual(expected)
  })

  it('FILTER empty output without if_empty is documented CALC', () => {
    expect(evaluateFormula('FILTER({1;2},{FALSE;FALSE})')).toBe('#CALC!')
  })

  it('FILTER empty output uses the supplied fallback', () => {
    const matrix = asMatrix(evalArray('FILTER({1;2},{FALSE;FALSE},"none")'))
    expect(plain(matrix)).toEqual([['none']])
  })

  it('FILTER propagates an include error and rejects a non-convertible include', () => {
    expect(evaluateFormula('FILTER({1;2},{#N/A;TRUE})')).toBe('#N/A')
    expect(evaluateFormula('FILTER({1;2},{"x";TRUE})')).toBe('#VALUE!')
  })

  it('FILTER rejects an include whose height/width does not match the source', () => {
    expect(evaluateFormula('FILTER({1;2},{TRUE;FALSE;TRUE})')).toBe('#VALUE!')
  })

  it('SORT keeps the input shape and rejects an undocumented sort order', () => {
    const matrix = asMatrix(evalArray('SORT({3,1;2,4},1,1)'))
    expect(matrix.rows).toBe(2)
    expect(matrix.cols).toBe(2)
    expect(evaluateFormula('SORT({3;1;2},1,2)')).toBe('#VALUE!')
  })

  it('SORT/SORTBY resolve structured table sources instead of mislabelling them', () => {
    const sortFixture = gridFixture(tableGrid, { tables: [amountTable] })
    expect(plain(asMatrix(evalArray('SORT(T[Amount])', sortFixture.ctx)))).toEqual([[10], [20], [30]])
    expect(plain(asMatrix(evalArray('SORTBY(T[Item],T[Amount])', sortFixture.ctx)))).toEqual([['b'], ['c'], ['a']])
  })

  it('array materialization resource policy is a diagnosed outcome, not a native shape error', () => {
    const fixture = gridFixture([])
    fixture.ctx.maxArrayCells = 4
    const result = evalArray('SEQUENCE(3,3)', fixture.ctx)
    expect(isEvaluationError(result)).toBe(true)
    expect(fixture.ctx.unsupportedFeatures?.has('array-resource-limit')).toBe(true)
  })
})

describe('C2 phase1 returned reference/matrix consumption', () => {
  it('SUM(INDEX(B1:B3,2)) consumes the selected value', () => {
    const fixture = gridFixture(lookupGrid)
    expect(evaluateFormula('SUM(INDEX(B1:B3,2))', fixture.ctx)).toBe(20)
  })

  it('SUM(INDEX(B1:B3,0,1)) sums the full selected axis', () => {
    const fixture = gridFixture(lookupGrid)
    expect(evaluateFormula('SUM(INDEX(B1:B3,0,1))', fixture.ctx)).toBe(60)
  })

  it('public INDEX zero-axis keeps the matrix shape in array mode and top-left in scalar mode', () => {
    const fixture = gridFixture([
      [1, 2],
      [3, 4],
    ])
    expect(evaluateFormula('INDEX(A1:B2,0,2)', fixture.ctx)).toBe(2)
    const matrix = asMatrix(evalArray('INDEX(A1:B2,0,2)', fixture.ctx))
    expect(matrix.rows).toBe(2)
    expect(matrix.cols).toBe(1)
    expect(plain(matrix)).toEqual([[2], [4]])
  })

  it('XLOOKUP keeps the full selected row and column', () => {
    const rowFixture = gridFixture(lookupGrid)
    const row = asMatrix(evalArray('XLOOKUP(2,A1:A3,B1:C3)', rowFixture.ctx))
    expect(row.rows).toBe(1)
    expect(row.cols).toBe(2)
    expect(plain(row)).toEqual([[20, 21]])

    const columnFixture = gridFixture([
      [1, 2, 3],
      [10, 20, 30],
      [11, 21, 31],
    ])
    const column = asMatrix(evalArray('XLOOKUP(2,A1:C1,A2:C3)', columnFixture.ctx))
    expect(column.rows).toBe(2)
    expect(column.cols).toBe(1)
    expect(plain(column)).toEqual([[20], [21]])
  })

  it('nested array/lookup function arguments retain matrix shape', () => {
    const fixture = gridFixture(lookupGrid)
    expect(evaluateFormula('SUM(TRANSPOSE(INDEX(B1:B3,0,1)))', fixture.ctx)).toBe(60)
    expect(plain(asMatrix(evalArray('SORT(INDEX(B1:B3,0,1),1,-1)', fixture.ctx)))).toEqual([[30], [20], [10]])
  })

  it('no longer gates returned-matrix consumption as an unavailable capability', () => {
    const fixture = gridFixture(lookupGrid)
    expect(evaluateFormula('SUM(INDEX(B1:B3,0,1))', fixture.ctx)).toBe(60)
    expect(fixture.ctx.unsupportedFeatures?.has('lookup-result-consumption-unavailable') ?? false).toBe(false)
  })

  it('closes the carried literal INDEX({10;20;30},2)=20 control', () => {
    expect(evaluateFormula('INDEX({10;20;30},2)')).toBe(20)
    expect(evaluateFormula('SUM(INDEX({10;20;30},2))')).toBe(20)
    expect(evaluateFormula('SUM(INDEX({10;20;30},0,1))')).toBe(60)
    const whole = asMatrix(evalArray('INDEX({10;20;30},0,1)'))
    expect(plain(whole)).toEqual([[10], [20], [30]])
  })

  it('applies the documented one-row INDEX omission rule to array values', () => {
    expect(evaluateFormula('INDEX({1,2,3},2)')).toBe(2)
    expect(evaluateFormula('INDEX({1,2,3},0)')).toBe(1)
    const whole = asMatrix(evalArray('INDEX({1,2,3},0)'))
    expect(plain(whole)).toEqual([[1, 2, 3]])
  })
})

describe('C2 phase1 scalar/array modes and public scope restoration', () => {
  it('never reuses a scalar/top-left memo for array arguments and restores the caller memo', () => {
    const fixture = gridFixture([])
    const ast = parseFormula('{1,2;3,4}')
    const memo = new WeakMap<AstNode, unknown>()
    memo.set(ast, 99)
    fixture.ctx.nodeValues = memo as NonNullable<EvaluationContext['nodeValues']>
    const matrix = asMatrix(evalArray(ast, fixture.ctx))
    expect(plain(matrix)).toEqual([
      [1, 2],
      [3, 4],
    ])
    expect(fixture.ctx.nodeValues).toBe(memo)
    expect(memo.get(ast)).toBe(99)
  })

  it('uses intentionally supplied mode stores without clearing them', () => {
    const fixture = gridFixture([])
    const ast = parseFormula('{1,2;3,4}')
    const arrayStore = { nodeValues: new WeakMap<AstNode, unknown>(), flatArgs: new WeakMap(), functionWork: new WeakMap() }
    const scalarStore = { nodeValues: new WeakMap<AstNode, unknown>(), flatArgs: new WeakMap(), functionWork: new WeakMap() }
    fixture.ctx.modeStores = {
      scalar: scalarStore as never,
      array: arrayStore as never,
    }
    const matrix = asMatrix(evalArray(ast, fixture.ctx))
    expect(plain(matrix)).toEqual([
      [1, 2],
      [3, 4],
    ])
    expect(arrayStore.nodeValues.has(ast)).toBe(true)
  })

  it('restores the caller context on a scheduler Pending suspension', () => {
    const fixture = gridFixture([[1], [2]], { pendingAt: { col: 0, row: 0 } })
    const memo = new WeakMap<AstNode, unknown>()
    fixture.ctx.nodeValues = memo as NonNullable<EvaluationContext['nodeValues']>
    let caught: unknown
    try {
      evalArray('INDEX(A1:A2,1)', fixture.ctx)
    } catch (error) {
      caught = error
    }
    expect(caught).toBe(fixture.marker)
    expect(fixture.ctx.mode).toBeUndefined()
    expect(fixture.ctx.nodeValues).toBe(memo)
  })

  it('recognizes modern SPILL/CALC literals while quoted lookalikes stay text', () => {
    expect(evaluateFormula('=#CALC!')).toBe('#CALC!')
    expect(evaluateFormula('=#SPILL!')).toBe('#SPILL!')
    expect(evaluateFormula('"#SPILL!"')).toBe('#SPILL!')
    const tagged = evaluateFormulaInternal(parseFormula('#CALC!'))
    expect(isEvaluationError(tagged)).toBe(true)
    if (isEvaluationError(tagged)) expect(tagged.code).toBe('#CALC!')
  })
})

describe('C2 phase1 element flag preservation', () => {
  it('TRANSPOSE preserves a genuine error element and error-looking text distinctly', () => {
    const matrix = asMatrix(evalArray('TRANSPOSE({1,#N/A})'))
    expect(plain(matrix).length).toBe(2)
    expect(matrix.values[0][0]).toBe(1)
    expect(isEvaluationError(matrix.values[1][0])).toBe(true)
    if (isEvaluationError(matrix.values[1][0])) expect(matrix.values[1][0].code).toBe('#N/A')
  })

  it('keeps error-looking text as text and out of the tagged error set', () => {
    const fixture = gridFixture([[1], ['#N/A']])
    expect(evaluateFormula('INDEX(A1:A2,2)', fixture.ctx)).toBe('#N/A')
    const tagged = evaluateFormulaInternal(parseFormula('INDEX(A1:A2,2)'), fixture.ctx)
    expect(isEvaluationError(tagged)).toBe(false)
  })

  it('quoted fallback text is not converted into a genuine error', () => {
    expect(evaluateFormula('FILTER({1},{FALSE},"#CALC!")')).toBe('#CALC!')
  })
})

describe('C2 phase1 array-mode reference/name geometry (root-preview port)', () => {
  const grid: ElementValue[][] = [
    [1, 10],
    [2, 20],
    [3, 30],
  ]
  const names: readonly DefinedNameMetadata[] = [{ name: 'Nums', source: '$A$1:$A$3', baseProvenance: 'unknown' }]
  it.each([
    ['A1:B3', [[1, 10], [2, 20], [3, 30]]],
    ['A1:A3+10', [[11], [12], [13]]],
    ['FILTER(B1:B3,A1:A3>1)', [[20], [30]]],
    ['SORT(Nums+1)', [[2], [3], [4]]],
    ['TRANSPOSE(A1:A3+1)', [[2, 3, 4]]],
  ] as const)('array mode preserves full reference geometry in %s', (source, values) => {
    const fixture = gridFixture(grid, { definedNames: names })
    const saved = fixture.ctx.nodeValues
    expect(evalArray(source, fixture.ctx)).toEqual({ kind: 'formula-matrix', rows: values.length, cols: values[0].length, values })
    expect(fixture.ctx.nodeValues).toBe(saved)
    expect(fixture.ctx.mode).toBeUndefined()
  })
})

describe('C2 phase1 documented array boundaries (root-preview port)', () => {
  const dataNames: readonly DefinedNameMetadata[] = [{ name: 'Data', source: '$A$1:$A$4', baseProvenance: 'unknown' }]
  const errorGrid: ElementValue[][] = [[1], [formulaError('#N/A')], [3], ['#N/A']]
  it('MS-XLSX negative numerical constants are legal array atoms', () => {
    expect(evalArray('{-1,-.25;-2E+3,4}')).toEqual({ kind: 'formula-matrix', rows: 2, cols: 2, values: [[-1, -0.25], [-2000, 4]] })
  })
  it('documented ISERROR(Data) maps each genuine element without coercing error-looking text', () => {
    const fixture = gridFixture(errorGrid, { definedNames: dataNames })
    expect(evalArray('ISERROR(Data)', fixture.ctx)).toEqual({ kind: 'formula-matrix', rows: 4, cols: 1, values: [[false], [true], [false], [false]] })
  })
  it('documented IFERROR array output replaces only genuine errors', () => {
    const fixture = gridFixture(errorGrid, { definedNames: dataNames })
    expect(evalArray('IFERROR(Data,0)', fixture.ctx)).toEqual({ kind: 'formula-matrix', rows: 4, cols: 1, values: [[1], [0], [3], ['#N/A']] })
  })
  it('Microsoft SUM(IF(ISERROR(Data),"",Data)) example follows elementwise branch selection', () => {
    const fixture = gridFixture(errorGrid, { definedNames: dataNames })
    expect(evalArray('SUM(IF(ISERROR(Data),"",Data))', fixture.ctx)).toEqual({ kind: 'formula-matrix', rows: 1, cols: 1, values: [[4]] })
  })
})

describe('C2 phase1 owned array-store generation restart (root-preview port)', () => {
  it.each(['A1:A3', 'TRANSPOSE(A1:A3)'])('owned array API refreshes its complete materialization after one callback generation change: %s', (source) => {
    let generation = 1
    let once = true
    const values = [1, 2, 3]
    const services = createReferenceServices({
      generation: () => generation,
      sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
      sheetIdOfName: () => 's',
      sheetNameOfId: () => 'S',
      definedNames: [],
      tables: [],
      store: {
        stored: () => values.map((value, row) => ({ address: { sheetId: 's', row, col: 0 }, value, origin: 'input' as const })),
        read: (address) => {
          if (once && address.row === 1) {
            once = false
            values.splice(0, 3, 100, 200, 300)
            generation++
          }
          return { value: values[address.row], origin: 'input' as const }
        },
      },
    })
    const ctx: EvaluationContext = { references: services, typedValues: true, currentSheet: 'S', currentCell: { row: 0, col: 3, absCol: false, absRow: false } }
    const expected = source.startsWith('TRANSPOSE')
      ? { kind: 'formula-matrix', rows: 1, cols: 3, values: [[100, 200, 300]] }
      : { kind: 'formula-matrix', rows: 3, cols: 1, values: [[100], [200], [300]] }
    expect(evalArray(source, ctx)).toEqual(expected)
    expect(ctx.mode).toBeUndefined()
    expect(ctx.nodeValues).toBeUndefined()
  })
})

describe('C2 phase1 array-value lookup inputs (root-preview port)', () => {
  const numsNames: readonly DefinedNameMetadata[] = [{ name: 'Nums', source: '{10;20;30}', baseProvenance: 'unknown' }]
  it.each([
    ['INDEX(Nums,2)', [[20]]],
    ['TRANSPOSE(Nums)', [[10, 20, 30]]],
    ['FILTER(Nums,{TRUE;FALSE;TRUE})', [[10], [30]]],
  ] as const)('value names containing arrays remain valid array arguments: %s', (source, values) => {
    const fixture = gridFixture([], { definedNames: numsNames })
    expect(evalArray(source, fixture.ctx)).toEqual({ kind: 'formula-matrix', rows: values.length, cols: values[0].length, values })
  })
  it.each([
    ['MATCH(2,{1;2;3},0)', 2],
    ['XMATCH(2,{1;2;3})', 2],
    ['XLOOKUP(2,{1;2;3},{10;20;30})', 20],
    ['VLOOKUP(2,{1,10;2,20;3,30},2,FALSE)', 20],
    ['HLOOKUP(2,{1,2,3;10,20,30},2,FALSE)', 20],
  ] as const)('documented lookup array inputs work without workbook metadata: %s', (source, value) => {
    expect(evalArray(source)).toEqual({ kind: 'formula-matrix', rows: 1, cols: 1, values: [[value]] })
  })
})

describe('C2 phase1 native literal lookup inputs (root-preview port)', () => {
  it.each([
    ['MATCH(2,{3;2;2;1},-1)', 2],
    ['MATCH("t*",{"one";"two";"three"},0)', 2],
    ['XMATCH(2,{1;2;2;3},0,2)', 2],
    ['XMATCH(2,{3;2;2;1},0,-2)', 3],
    ['XLOOKUP(2,{1;2;2;3},{10;20;30;40},,0,2)', 20],
    ['XLOOKUP(2,{3;2;2;1},{40;30;20;10},,0,-2)', 20],
    ['XLOOKUP("t*",{"one";"two";"three"},{1;2;3},,2)', 2],
    ['HLOOKUP(2,{1,2,3;10,20,30},2,FALSE)', 20],
    ['INDEX({1,2;3,4},2)', 3],
    ['INDEX({1,2;3,4},0,2)', 2],
  ] as const)('existing native scalar projection from actual literal input: %s', (source, value) => {
    expect(evaluateFormula(source)).toBe(value)
  })
  it('primary-documented INDEX 2D array form with only row_num returns the entire selected row', () => {
    expect(evalArray('INDEX({1,2;3,4},2)')).toEqual({ kind: 'formula-matrix', rows: 1, cols: 2, values: [[3, 4]] })
  })
})

describe('C2 phase1 conditional shape safety (root-preview port)', () => {
  it('documented IFNA per-cell replacement catches NA only and preserves typed/text boundaries', () => {
    expect(evalArray('IFNA({1,#N/A,#DIV/0!,"#N/A"},0)')).toEqual({ kind: 'formula-matrix', rows: 1, cols: 4, values: [[1, 0, formulaError('#DIV/0!'), '#N/A']] })
  })
  it.each(['IFERROR({1;#N/A;#N/A},{8;9})', 'IF({TRUE;FALSE;TRUE},{10;20},0)'])('unmeasured incompatible conditional dimensions diagnose rather than pad null: %s', (source) => {
    const ctx: EvaluationContext = {}
    const value = evalArray(source, ctx)
    expect(isEvaluationError(value)).toBe(true)
    expect(ctx.unsupportedFeatures?.size ?? 0).toBeGreaterThan(0)
  })
  it.each([
    ['IF({FALSE;FALSE;FALSE},{10;20},0)', [[0], [0], [0]]],
    ['IFERROR({1;2;3},{8;9})', [[1], [2], [3]]],
  ] as const)('unreached incompatible branch does not gate: %s', (source, values) => {
    const ctx: EvaluationContext = {}
    expect(evalArray(source, ctx)).toEqual({ kind: 'formula-matrix', rows: 3, cols: 1, values })
    expect(ctx.unsupportedFeatures?.size ?? 0).toBe(0)
  })
})

describe('C2 phase1 consumer mode stores (root-preview port)', () => {
  const modeGrid: ElementValue[][] = [[1], [2], [3]]
  it.each([
    ['TRANSPOSE(A1:A3+1)', 2],
    ['SUM(TRANSPOSE(A1:A3+1))', 9],
  ] as const)('scalar API array consumer evaluates its expression argument in array mode: %s', (source, value) => {
    const fixture = gridFixture(modeGrid, { currentCell: { col: 5, row: 1 } })
    expect(evaluateFormula(source, fixture.ctx)).toBe(value)
  })
  it('array consumer bypasses intentionally supplied scalar argument memo while preserving it', () => {
    const fixture = gridFixture(modeGrid, { currentCell: { col: 5, row: 1 } })
    const ast = parseFormula('TRANSPOSE(A1:A3+1)')
    if (ast.type !== 'call') throw new Error('bad fixture')
    const arg = ast.args[0]
    const memo = new WeakMap<AstNode, EvaluationValue>()
    memo.set(arg, 99)
    fixture.ctx.nodeValues = memo
    fixture.ctx.modeStores = {
      scalar: { nodeValues: memo, flatArgs: new WeakMap(), functionWork: new WeakMap() },
      array: { nodeValues: new WeakMap(), flatArgs: new WeakMap(), functionWork: new WeakMap() },
    }
    expect(evaluateFormulaInternal(ast, fixture.ctx)).toBe(2)
    expect(fixture.ctx.nodeValues).toBe(memo)
    expect(memo.get(arg)).toBe(99)
    expect(fixture.ctx.mode).toBeUndefined()
  })
})

describe('C2 phase1 mode ownership (root-preview port)', () => {
  const nameGrid: ElementValue[][] = [[1], [2], [3]]
  const nameDefs: readonly DefinedNameMetadata[] = [{ name: 'Nums', source: '$A$1:$A$3+1', baseProvenance: 'unknown' }]
  it('same bound source evaluated in array and scalar modes keeps default memos distinct', () => {
    const fixture = gridFixture(nameGrid, { definedNames: nameDefs, currentCell: { col: 5, row: 1 } })
    fixture.ctx.nodeValues = undefined
    expect(evaluateFormula('TRANSPOSE(Nums)+Nums', fixture.ctx)).toBe(5)
    expect(fixture.ctx.mode).toBeUndefined()
    expect(fixture.ctx.nodeValues).toBeUndefined()
  })
  it('owned scalar-to-array child inherits root generation restart ownership', () => {
    const values = [1, 2, 3]
    let generation = 1
    let once = true
    const services = createReferenceServices({
      generation: () => generation,
      sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
      sheetIdOfName: () => 's',
      sheetNameOfId: () => 'S',
      definedNames: [],
      tables: [],
      store: {
        stored: () => values.map((value, row) => ({ address: { sheetId: 's', row, col: 0 }, value, origin: 'input' as const })),
        read: (address) => {
          if (once && address.row === 1) { once = false; values.splice(0, 3, 100, 200, 300); generation++ }
          return { value: values[address.row], origin: 'input' as const }
        },
      },
    })
    const ctx: EvaluationContext = { references: services, typedValues: true, currentSheet: 'S', currentCell: { col: 5, row: 1, absCol: false, absRow: false }, functionWork: new WeakMap(), flatArgs: new WeakMap(), getCellValue: (_s, _c, row) => values[row] ?? null }
    expect(evaluateFormula('SUM(TRANSPOSE(A1:A3+1))', ctx)).toBe(603)
    expect(ctx.mode).toBeUndefined()
    expect(ctx.nodeValues).toBeUndefined()
  })
  it('array consumer honors an explicitly supplied evaluateInMode callback', () => {
    const fixture = gridFixture(nameGrid, { definedNames: nameDefs, currentCell: { col: 5, row: 1 } })
    fixture.ctx.nodeValues = undefined
    let calls = 0
    const hook: NonNullable<EvaluationContext['evaluateInMode']> = (_node, mode) => {
      calls++
      expect(mode).toBe('array')
      return { kind: 'formula-matrix', rows: 3, cols: 1, values: [[2], [3], [4]] }
    }
    fixture.ctx.evaluateInMode = hook
    expect(evaluateFormula('TRANSPOSE(A1:A3+1)', fixture.ctx)).toBe(2)
    expect(calls).toBe(1)
    expect(fixture.ctx.evaluateInMode).toBe(hook)
  })
  it('external evaluateInMode Pending identity escapes and hook restores', () => {
    const fixture = gridFixture(nameGrid, { definedNames: nameDefs, currentCell: { col: 5, row: 1 } })
    fixture.ctx.nodeValues = undefined
    const marker = { externalModePending: 1 }
    let first = true
    const hook: NonNullable<EvaluationContext['evaluateInMode']> = () => {
      if (first) { first = false; throw marker }
      return { kind: 'formula-matrix', rows: 3, cols: 1, values: [[2], [3], [4]] }
    }
    fixture.ctx.evaluateInMode = hook
    let caught: unknown
    try { evaluateFormula('TRANSPOSE(A1:A3+1)', fixture.ctx) } catch (error) { caught = error }
    expect(caught).toBe(marker)
    expect(fixture.ctx.evaluateInMode).toBe(hook)
    expect(fixture.ctx.mode).toBeUndefined()
    expect(evaluateFormula('TRANSPOSE(A1:A3+1)', fixture.ctx)).toBe(2)
  })
})

describe('C2 phase1 private matrix preparation boundary (root-preview port)', () => {
  it.each([
    'INDEX(Nums,2)',
    'XLOOKUP(2,{1;2;3},{10;20;30})',
    'VLOOKUP(2,{1,10;2,20;3,30},2,FALSE)',
    'HLOOKUP(2,{1,2,3;10,20,30},2,FALSE)',
  ])('private in-memory values never invoke physical workbook preparation: %s', (source) => {
    let prepareCalls = 0
    const services = createReferenceServices({
      generation: () => 1,
      sheets: [{ sheetId: 's', name: 'S', workbookIndex: 0 }],
      sheetIdOfName: () => 's',
      sheetNameOfId: () => 'S',
      definedNames: [{ name: 'Nums', source: '{10;20;30}', baseProvenance: 'unknown' }],
      tables: [],
      store: { stored: () => [], read: () => undefined },
    })
    services.prepare = () => { prepareCalls++; throw new Error('in-memory value view reached physical workbook hook') }
    const ctx: EvaluationContext = { references: services, typedValues: true, currentSheet: 'S', currentCell: { row: 0, col: 5, absCol: false, absRow: false } }
    expect(evalArray(source, ctx)).toEqual({ kind: 'formula-matrix', rows: 1, cols: 1, values: [[20]] })
    expect(prepareCalls).toBe(0)
  })
})

describe('C2 phase1 standalone array consumers (root-preview port)', () => {
  it.each([
    ['SUM(TRANSPOSE({1;2}+1))', 5],
    ['SUM(TRANSPOSE(SEQUENCE(2,3)+1))', 27],
    ['SUM({1;2}+1)', 5],
  ] as const)('standalone scalar aggregate preserves its array argument: %s', (source, value) => {
    expect(evaluateFormula(source)).toBe(value)
    expect(evalArray(source)).toEqual({ kind: 'formula-matrix', rows: 1, cols: 1, values: [[value]] })
  })
  it.each([
    ['ROWS({1,2,3;4,5,6})', 2],
    ['COLUMNS({1,2,3;4,5,6})', 3],
    ['ROWS(SEQUENCE(2,3))', 2],
  ] as const)('documented dimension helper accepts array/array-formula inputs: %s', (source, value) => {
    expect(evaluateFormula(source)).toBe(value)
  })
})

describe('C2 phase1 grammar boundary guards', () => {
  it('does not invent an arrayConst for a single literal', () => {
    expect(parseFormula('1').type).toBe('number')
  })

  it('does not let a literal error token be a formula error node in scalar position', () => {
    expect(evaluateFormulaInternal(parseFormula('{#N/A}'))).toEqual(formulaError('#N/A'))
  })
})
