import { expect, test } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { buildXlsx, type XlsxCellSpec, type XlsxSheetSpec } from '../src/testdata/ooxml-builders'
import { parseXlsx } from '../src/xlsx/parse'
import { evaluateFormula } from '../src/xlsx/formula/evaluator'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import { translateSharedFormula } from '../src/xlsx/formula/shared'
import { renderSheet } from '../src/xlsx/render'
import type { XlsxDocument } from '../src/xlsx/types'

async function parse(sheets: XlsxSheetSpec[], strings: string[] = [], fullCalc = false) {
  let bytes = await buildXlsx(sheets, strings)
  if (fullCalc) {
    const zip = await JSZip.loadAsync(bytes)
    const xml = await zip.file('xl/workbook.xml')!.async('string')
    zip.file('xl/workbook.xml', xml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>'))
    bytes = await zip.generateAsync({ type: 'uint8array' })
  }
  return parseXlsx(await OfficePackage.load(bytes))
}
function oneRow(cells: XlsxCellSpec[]): XlsxSheetSpec[] {
  return [{ name: 'Sheet1', rows: [{ r: 1, cells }] }]
}
function values(doc: XlsxDocument, sheet = 0) {
  return Object.fromEntries(doc.sheets[sheet].rows.flatMap(r => r.cells.map(c => [c.ref, c.value])))
}

test('absent string/boolean caches calculate while explicit empty/false caches remain', async () => {
  const sheets = oneRow([
    { ref: 'A1', t: 'str', formula: '"computed"' },
    { ref: 'B1', t: 'b', formula: 'TRUE' },
    { ref: 'C1', t: 'str', v: '', formula: '"computed"' },
    { ref: 'D1', t: 'b', v: 0, formula: 'TRUE' },
    { ref: 'E1', t: 'str', v: '', formula: '"forced"', ca: true },
  ])
  const doc = await parse(sheets)
  expect(values(doc)).toEqual({ A1: 'computed', B1: true, C1: '', D1: false, E1: 'forced' })
  evaluateWorkbookFormulas(doc)
  expect(values(doc)).toEqual({ A1: 'computed', B1: true, C1: '', D1: false, E1: 'forced' })
  evaluateWorkbookFormulas(doc, { forceRecalc: true })
  expect(values(doc)).toEqual({ A1: 'computed', B1: true, C1: 'computed', D1: true, E1: 'forced' })
  expect(values(await parse(sheets, [], true))).toEqual(values(doc))
})

test('shared string/boolean followers distinguish absent caches from empty/false', async () => {
  const doc = await parse([{ name: 'Sheet1', rows: [
    { r: 1, cells: [
      { ref: 'A1', t: 'str', formula: '"computed"', sharedFormula: { si: 0, ref: 'A1:A3' } },
      { ref: 'B1', t: 'b', formula: 'TRUE', sharedFormula: { si: 1, ref: 'B1:B3' } },
    ] },
    { r: 2, cells: [{ ref: 'A2', t: 'str', sharedFormula: { si: 0 } }, { ref: 'B2', t: 'b', sharedFormula: { si: 1 } }] },
    { r: 3, cells: [{ ref: 'A3', t: 'str', v: '', sharedFormula: { si: 0 } }, { ref: 'B3', t: 'b', v: 0, sharedFormula: { si: 1 } }] },
  ] }])
  expect(values(doc)).toEqual({ A1: 'computed', B1: true, A2: 'computed', B2: true, A3: '', B3: false })
})

for (const order of ['ascending', 'descending', 'interleaved'] as const) {
  test(`2000-cell dependency chain computes every cell with ${order} source order`, async () => {
    let rows: XlsxSheetSpec['rows'] = Array.from({ length: 2000 }, (_, i) => ({ r: i + 1, cells: [
      i === 1999 ? { ref: 'A2000', v: 1 } : { ref: `A${i + 1}`, formula: `A${i + 2}+1` },
    ] }))
    if (order === 'descending') rows.reverse()
    if (order === 'interleaved') rows = [...rows.filter(r => r.r % 2 === 0), ...rows.filter(r => r.r % 2)]
    const doc = await parse([{ name: 'Sheet1', rows }])
    for (const cell of doc.sheets[0].rows.flatMap(r => r.cells)) expect(cell.value).toBe(2001 - cell.row - 1)
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    expect(values(doc).A1).toBe(2000)
  })
}

test('deep shared and cross-sheet dependencies retain lazy branches and cached cycle breaks', async () => {
  const rows: XlsxSheetSpec['rows'] = Array.from({ length: 700 }, (_, i) => ({ r: i + 1, cells: [
    i === 699 ? { ref: 'A700', formula: 'Data!A1+1' } : i === 0
      ? { ref: 'A1', formula: 'A2+1', sharedFormula: { si: 0, ref: 'A1:A699' } }
      : { ref: `A${i + 1}`, sharedFormula: { si: 0 } },
  ] }))
  rows[0].cells.push({ ref: 'B1', formula: 'IF(TRUE,A1,B1)' }, { ref: 'C1', formula: 'IFERROR(A1,C1)' },
    { ref: 'D1', v: 9, formula: 'E1' }, { ref: 'E1', formula: 'D1+1' })
  const doc = await parse([{ name: 'Sheet1', rows }, { name: 'Data', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])
  expect(values(doc)).toMatchObject({ A1: 701, A700: 2, B1: 701, C1: 701, D1: 9, E1: 10 })
})

test('actual cycle suffix alone becomes zero even after a deep dependent prefix', async () => {
  const rows: XlsxSheetSpec['rows'] = Array.from({ length: 700 }, (_, i) => ({ r: i + 1, cells: [
    { ref: `A${i + 1}`, formula: i === 699 ? 'A699+1' : `A${i + 2}+1` },
  ] }))
  const doc = await parse([{ name: 'Sheet1', rows }])
  expect(values(doc)).toMatchObject({ A1: 698, A698: 1, A699: 0, A700: 0 })
})

test('shared serialization preserves same-precedence grouping and floating evaluation order', async () => {
  const doc = await parse([{ name: 'Sheet1', rows: [
    { r: 1, cells: [
      { ref: 'A1', formula: '1E200*(1E200/1E200)', sharedFormula: { si: 0, ref: 'A1:A2' } },
      { ref: 'B1', formula: '1E16+(-1E16+1)', sharedFormula: { si: 1, ref: 'B1:B2' } },
    ] },
    { r: 2, cells: [{ ref: 'A2', sharedFormula: { si: 0 } }, { ref: 'B2', sharedFormula: { si: 1 } }] },
  ] }])
  expect(values(doc)).toEqual({ A1: 1E200, B1: 0, A2: 1E200, B2: 0 })
  expect(translateSharedFormula('$A$1+A$1+$A1+A1', 1, 1).formula).toBe('$A$1+B$1+$A2+B2')
  expect(evaluateFormula(translateSharedFormula('1+(', 1, 1).formula)).toBe('#NAME?')
})

test('error-looking literal and calculated text stays text in operations and IFERROR', () => {
  expect(evaluateFormula('IFERROR("#N/A","fallback")')).toBe('#N/A')
  expect(evaluateFormula('LEN("#N/A")')).toBe(4)
  expect(evaluateFormula('IFERROR("#"&"N/A","fallback")')).toBe('#N/A')
  expect(evaluateFormula('IFERROR(CONCAT("#","N/A"),"fallback")')).toBe('#N/A')
  expect(evaluateFormula('IFERROR(#N/A,"fallback")')).toBe('fallback')
})

test('cached shared/string text and real errors keep distinct identity through references', async () => {
  const doc = await parse(oneRow([
    { ref: 'A1', t: 's', v: 0 }, { ref: 'B1', t: 'str', v: '#N/A', formula: '1' },
    { ref: 'C1', t: 'e', v: '#N/A' }, { ref: 'D1', formula: '"#"&"N/A"' },
    { ref: 'E1', formula: 'IFERROR(A1,"fallback")' }, { ref: 'F1', formula: 'IFERROR(B1,"fallback")' },
    { ref: 'G1', formula: 'IFERROR(C1,"fallback")' }, { ref: 'H1', formula: 'LEN(D1)' },
    { ref: 'I1', formula: 'C1' }, { ref: 'J1', formula: 'IFERROR(I1,"fallback")' },
    { ref: 'K1', t: 'e', v: '#N/A', formula: '"#N/A"', ca: true }, { ref: 'L1', formula: 'LEN(K1)' },
    { ref: 'M1', formula: 'CONCAT(A1,B1)' },
  ]), ['#N/A'])
  expect(values(doc)).toEqual({ A1: '#N/A', B1: '#N/A', C1: '#N/A', D1: '#N/A', E1: '#N/A', F1: '#N/A', G1: 'fallback', H1: 4, I1: '#N/A', J1: 'fallback', K1: '#N/A', L1: 4, M1: '#N/A#N/A' })
  evaluateWorkbookFormulas(doc)
  expect(values(doc).H1).toBe(4)
  expect(values(doc).J1).toBe('fallback')
  expect(JSON.parse(JSON.stringify(doc)).sheets[0].rows[0].cells.every((c: { value: unknown }) => typeof c.value !== 'object')).toBe(true)
})

test('rendered formula text is left aligned and real cached/calculated errors centered', async () => {
  const doc = await parse([{ name: 'Sheet1', cols: '<col min="1" max="4" width="12"/>', rows: [{ r: 1, cells: [
    { ref: 'A1', formula: '"#N/A"' }, { ref: 'B1', formula: '#N/A' },
    { ref: 'C1', t: 'str', v: '#N/A', formula: '1' }, { ref: 'D1', t: 'e', v: '#N/A' },
  ] }] }])
  const painted: { text: string; x: number }[] = []
  const noop = () => {}
  const ctx = { save: noop, restore: noop, beginPath: noop, rect: noop, clip: noop, fillRect: noop, stroke: noop,
    moveTo: noop, lineTo: noop, translate: noop, measureText: (text: string) => ({ width: text.length * 7 }),
    fillText: (text: string, x: number) => painted.push({ text, x }), font: '', fillStyle: '', strokeStyle: '', lineWidth: 1, textBaseline: 'alphabetic' }
  renderSheet(doc.sheets[0], ctx as unknown as CanvasRenderingContext2D)
  expect(painted.filter(p => p.text === '#N/A').map(p => p.x)).toEqual([3, 89 + (89 - 28) / 2, 178 + 3, 267 + (89 - 28) / 2])
})

test('AND/OR flatten ranges, ignore reference text/blanks and propagate later real errors', async () => {
  const doc = await parse([{ name: 'Sheet1', rows: [
    { r: 1, cells: [{ ref: 'A1', t: 'b', v: 1 }, { ref: 'B1', t: 'str', v: 'text' }, { ref: 'C1', t: 'e', v: '#N/A' },
      { ref: 'D1', formula: 'AND(A1:A2)' }, { ref: 'E1', formula: 'OR(A1:A2)' },
      { ref: 'F1', formula: 'AND(A1:B2)' }, { ref: 'G1', formula: 'OR(A1:B2)' },
      { ref: 'H1', formula: 'AND(B1:B2)' }, { ref: 'I1', formula: 'OR(B1:B2)' },
      { ref: 'J1', formula: 'AND(FALSE,C1)' }, { ref: 'K1', formula: 'OR(TRUE,C1)' },
      { ref: 'L1', formula: 'AND(A1,A2)' }, { ref: 'M1', formula: 'OR(A1,A2)' },
      { ref: 'N1', formula: 'AND(B1,TRUE)' }, { ref: 'O1', formula: 'OR(B1,FALSE)' },
      { ref: 'P1', formula: 'IF(FALSE,C1,7)' }, { ref: 'Q1', formula: 'AND(1,2)' }, { ref: 'R1', formula: 'OR(0,0)' },
    ] },
    { r: 2, cells: [{ ref: 'A2', t: 'b', v: 1 }] },
  ] }])
  expect(values(doc)).toMatchObject({ D1: true, E1: true, F1: true, G1: true, H1: '#VALUE!', I1: '#VALUE!', J1: '#N/A', K1: '#N/A', L1: true, M1: true, N1: true, O1: false, P1: 7, Q1: true, R1: false })
})

test('range fanout resolves uncached dependencies with linear reference reads', async () => {
  // Observe the real evaluator callback, rather than timing machine-dependent work.
  const { FUNCTIONS } = await import('../src/xlsx/formula/functions')
  const original = FUNCTIONS.SUM
  const contexts = new WeakSet<object>()
  let reads = 0
  FUNCTIONS.SUM = (args, ctx, evalNode) => {
    if (ctx && !contexts.has(ctx)) {
      contexts.add(ctx)
      const get = ctx.getCellValue!
      ctx.getCellValue = (...coordinates) => { reads++; return get(...coordinates) }
    }
    return original(args, ctx, evalNode)
  }
  try {
    const count = 1200
    const doc = await parse([{ name: 'Sheet1', rows: [
      { r: 1, cells: [{ ref: 'B1', formula: `SUM(A1:A${count})` }, { ref: 'A1', formula: '1' }] },
      ...Array.from({ length: count - 1 }, (_, i) => ({ r: i + 2, cells: [{ ref: `A${i + 2}`, formula: '1' }] })),
    ] }])
    expect(values(doc).B1).toBe(count)
    expect(reads).toBeLessThanOrEqual(count * 3)
  } finally {
    FUNCTIONS.SUM = original
  }
})

test('recursive legacy-context resource guard reports NUM instead of a false cycle zero', () => {
  expect(evaluateFormula('1+1', { evalDepth: 512 })).toBe('#NUM!')
})

test('typed custom contexts distinguish text and tagged errors while legacy callbacks stay compatible', async () => {
  const { formulaError } = await import('../src/xlsx/formula/evaluator')
  const typed = { typedValues: true, getCellValue: (_sheet: string | undefined, col: number) => col === 0 ? '#N/A' : formulaError('#N/A'),
    getRangeValues: () => [['#N/A', formulaError('#N/A')]] }
  expect(evaluateFormula('LEN(A1)', typed)).toBe(4)
  expect(evaluateFormula('IFERROR(A1,"fallback")', typed)).toBe('#N/A')
  expect(evaluateFormula('IFERROR(B1,"fallback")', typed)).toBe('fallback')
  expect(evaluateFormula('IFERROR(CONCAT(A1:B1),"fallback")', typed)).toBe('fallback')
  const legacy = { getCellValue: () => '#N/A' }
  expect(evaluateFormula('IFERROR(A1,"fallback")', legacy)).toBe('fallback')
  expect(evaluateFormula('LEN(A1)', legacy)).toBe('#N/A')
})

test('iterative frames preserve authored sheet names including spaces and exclamation marks', async () => {
  const doc = await parse([
    { name: ' Sheet ! ', rows: [{ r: 1, cells: [{ ref: 'A1', formula: 'A2+1' }] }, { r: 2, cells: [{ ref: 'A2', v: 3 }] }] },
    { name: 'Summary', rows: [{ r: 1, cells: [{ ref: 'A1', formula: "' Sheet ! '!A1+1" }] }] },
  ])
  expect(values(doc).A1).toBe(4)
  expect(values(doc, 1).A1).toBe(5)
})

test('formula display and text search/copy expose primitive calculated values and real errors', async () => {
  const { buildTextIndex, findMatches } = await import('../src/core/search')
  const { textForRange } = await import('../src/core/selection')
  const doc = await parse(oneRow([
    { ref: 'A1', t: 'str', formula: '"computed-text"' }, { ref: 'B1', formula: 'LEN(A1)' },
    { ref: 'C1', formula: '"#N/A"&" text"' }, { ref: 'D1', formula: '1/0' },
  ]))
  const index = await buildTextIndex([{ spec: { widthPx: 600, heightPx: 40 }, paint: ctx => renderSheet(doc.sheets[0], ctx) }])
  expect(findMatches(index, 'computed-text')).toHaveLength(1)
  expect(findMatches(index, '#N/A text')).toHaveLength(1)
  expect(findMatches(index, '#DIV/0!')).toHaveLength(1)
  const line = index.pages[0].lines[0]
  expect(line.text).toBe('computed-text13#N/A text#DIV/0!')
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 0, charIndex: line.text.length } })).toBe(line.text)
  expect(line.text).not.toContain('[object Object]')
})

test('multiple range arguments retain completed flattening work across later dependencies', async () => {
  const { FUNCTIONS } = await import('../src/xlsx/formula/functions')
  const original = FUNCTIONS.SUM
  const contexts = new WeakSet<object>()
  let visits = 0
  FUNCTIONS.SUM = (args, ctx, evalNode) => {
    if (ctx && !contexts.has(ctx)) {
      contexts.add(ctx)
      const getRange = ctx.getRangeValues!
      ctx.getRangeValues = (...coordinates) => getRange(...coordinates).map(row => new Proxy(row, {
        get(target, key, receiver) {
          if (key === Symbol.iterator) return function* () { for (const value of target) { visits++; yield value } }
          return Reflect.get(target, key, receiver)
        },
      }))
    }
    return original(args, ctx, evalNode)
  }
  try {
    const count = 800
    const rows = Array.from({ length: count }, (_, i) => ({ r: i + 1, cells: [
      { ref: `A${i + 1}`, formula: '1' }, { ref: `B${i + 1}`, formula: '2' },
    ] }))
    rows[0].cells.unshift({ ref: 'C1', formula: `SUM(A1:A${count},B1:B${count})` })
    const doc = await parse([{ name: 'Sheet1', rows }])
    expect(values(doc).C1).toBe(count * 3)
    expect(visits).toBeLessThanOrEqual(count * 4)
  } finally {
    FUNCTIONS.SUM = original
  }
})

test('many scalar reference arguments retain completed flattening work across later dependencies', async () => {
  const { FUNCTIONS } = await import('../src/xlsx/formula/functions')
  const original = FUNCTIONS.SUM
  let visits = 0
  FUNCTIONS.SUM = (args, ctx, evalNode) => original(args, ctx, (node, context) => {
    visits++
    return evalNode(node, context)
  })
  try {
    const count = 800
    const rows = Array.from({ length: count }, (_, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, formula: '1' }] }))
    rows[0].cells.unshift({ ref: 'B1', formula: `SUM(${rows.map(row => `A${row.r}`).join(',')})` })
    const doc = await parse([{ name: 'Sheet1', rows }])
    expect(values(doc).B1).toBe(count)
    expect(visits).toBeLessThanOrEqual(count * 3)
  } finally {
    FUNCTIONS.SUM = original
  }
})

for (const code of ['#N/A', '#DIV/0!', '#REF!']) {
  for (const errorKind of [undefined, true, false]) {
    test(`legacy non-formula ${code} source with error metadata ${String(errorKind)} preserves its reference contract`, () => {
      const doc: XlsxDocument = JSON.parse(JSON.stringify({ sheets: [{
        name: 'Legacy', merges: [], mergeRanges: [], cols: [], rows: [{ index: 0, cells: [
          { ref: 'A1', col: 0, row: 0, value: code, styleIndex: 0, ...(errorKind === undefined ? {} : { valueIsError: errorKind }) },
          { ref: 'B1', col: 1, row: 0, value: null, formula: 'IFERROR(A1,"fallback")', styleIndex: 0 },
          { ref: 'C1', col: 2, row: 0, value: null, formula: 'SUM(A1,1)', styleIndex: 0 },
          { ref: 'D1', col: 3, row: 0, value: null, formula: 'LEN(A1)', styleIndex: 0 },
        ] }],
      }] }))
      const expected = errorKind === false
        ? { A1: code, B1: code, C1: 1, D1: code.length }
        : { A1: code, B1: 'fallback', C1: code, D1: code }
      evaluateWorkbookFormulas(doc)
      expect(values(doc)).toEqual(expected)
      expect(doc.sheets[0].rows[0].cells.slice(1).map(cell => cell.valueIsError)).toEqual([false, errorKind !== false, errorKind !== false])
      evaluateWorkbookFormulas(doc, { forceRecalc: true })
      expect(values(doc)).toEqual(expected)
      expect(doc.sheets[0].rows[0].cells.every(cell => typeof cell.value !== 'object')).toBe(true)
    })
  }
}
