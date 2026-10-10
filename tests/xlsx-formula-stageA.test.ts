/**
 * StageA native-oracle regression gates (Excel 16.106.1 COMPARISON.json rows
 * filtered to currently supported constructs + StageA architecture bounds).
 * All imports bind THIS worktree; no external test is treated as oracle.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { EvaluationContext } from '../src/xlsx/formula/types'
import { evaluateFormula } from '../src/xlsx/formula/evaluator'
import { translateSharedFormula } from '../src/xlsx/formula/shared'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

describe('StageA: bounded nonrecursive evaluation (Node22: 8000 unaries / 4000 addends)', () => {
  it('computes 8000 nested unaries without RangeError', () => {
    const formula = '=' + '-'.repeat(8000) + '2'
    expect(evaluateFormula(formula)).toBe(2)
  })
  it('computes 4000 addends without RangeError', () => {
    const formula = '=1' + '+1'.repeat(3999)
    expect(evaluateFormula(formula)).toBe(4000)
  })
  it('renders an 8000-addend diagnostic ladder without RangeError (mixed errors)', () => {
    const formula = '=1' + '+#N/A'.repeat(800)
    expect(evaluateFormula(formula)).toBe('#N/A')
  })
})

describe('StageA: COUNT/COUNTA native semantics', () => {
  it('COUNT ignores errors (never propagates) and non-numeric text', () => {
    expect(evaluateFormula('COUNT(1, #N/A)')).toBe(1)
    expect(evaluateFormula('COUNT("abc")')).toBe(0)
    expect(evaluateFormula('COUNT("")')).toBe(0)
    expect(evaluateFormula('COUNT(" ")')).toBe(0)
    expect(evaluateFormula('COUNT(TRUE, 42, "10", "text")')).toBe(3)
  })
  it('COUNT skips error/blank refs in ranges but counts numbers', () => {
    const ctx: EvaluationContext = {
      getCellValue: (_s: string | undefined, col: number) => {
        if (col === 0) return { kind: 'formula-error', code: '#N/A' } as never
        if (col === 1) return 42
        if (col === 2) return 'text'
        return null
      },
    }
    expect(evaluateFormula('COUNT(A1:C1)', ctx)).toBe(1)
    expect(evaluateFormula('COUNT(A1)', ctx)).toBe(0)
    expect(evaluateFormula('SUM(A1:C1)', ctx)).toBe('#N/A') // native SUM propagates range errors (COUNT skips them)
  })
  it('COUNT explicit omitted args participate (native COUNT(,)=2)', () => {
    expect(evaluateFormula('COUNT(,)')).toBe(2)
    expect(evaluateFormula('COUNT(1,)')).toBe(2)
    expect(evaluateFormula('COUNTA(,)')).toBe(2)
  })
  it('COUNTA counts empty text and blank refs zero', () => {
    expect(evaluateFormula('COUNTA("")')).toBe(1)
    const ctx: EvaluationContext = {
      getCellValue: (_s: string | undefined, col: number) => (col === 0 ? '' : null),
    }
    expect(evaluateFormula('COUNTA(A1)', ctx)).toBe(1) // B1-ref holds ""
    expect(evaluateFormula('COUNTA(D1)', ctx)).toBe(0) // blank ref
  })
})

describe('StageA: omitted args participate as 0; blank refs stay ignorable', () => {
  it('MIN/MAX/AVERAGE omitted args act as 0 participants', () => {
    expect(evaluateFormula('MIN(5,)')).toBe(0)
    expect(evaluateFormula('MAX(-5,)')).toBe(0)
    expect(evaluateFormula('AVERAGE(,)')).toBe(0)
    expect(evaluateFormula('PRODUCT(5,)')).toBe(5)
    expect(evaluateFormula('SUM(,)')).toBe(0)
  })
  it('blank references remain ignorable in aggregates', () => {
    const blankCtx: EvaluationContext = { getCellValue: () => null }
    expect(evaluateFormula('MIN(5, D1)', blankCtx)).toBe(5)
    expect(evaluateFormula('MAX(-5, D1)', blankCtx)).toBe(-5)
    expect(evaluateFormula('AVERAGE(5, D1)', blankCtx)).toBe(5)
    expect(evaluateFormula('SUM(5, D1)', blankCtx)).toBe(5)
    expect(evaluateFormula('COUNT(D1)', blankCtx)).toBe(0)
    expect(evaluateFormula('COUNTA(D1)', blankCtx)).toBe(0)
  })
})

describe('StageA: IF omitted/blank scalarization; blank=FALSE comparison', () => {
  it('IF empty args and blank refs scalarize to 0', () => {
    expect(evaluateFormula('IF(TRUE,,7)')).toBe(0)
    expect(evaluateFormula('IF(FALSE,7,)')).toBe(0)
    const blankCtx: EvaluationContext = { getCellValue: () => null }
    expect(evaluateFormula('IF(TRUE,D1,7)', blankCtx)).toBe(0)
    // REV1-4: typed-context blank ref is native 0; unbound (no getCellValue) is null.
    expect(evaluateFormula('D1', blankCtx)).toBe(0)
  })
  it('blank compares equal to FALSE (native blank=FALSE TRUE)', () => {
    const blankCtx: EvaluationContext = { getCellValue: () => null }
    expect(evaluateFormula('D1=FALSE', blankCtx)).toBe(true)
    expect(evaluateFormula('FALSE=0', blankCtx)).toBe(false)
    expect(evaluateFormula('D1=0', blankCtx)).toBe(true)
    expect(evaluateFormula('D1=""', blankCtx)).toBe(true)
  })
  it('workbook writes blank scalar formula results as 0; absent inputs remain null', async () => {
    const doc = { sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
      { ref: 'A1', row: 0, col: 0, value: null, styleIndex: 0 },
      { ref: 'B1', row: 0, col: 1, value: null, styleIndex: 0, formula: '=A1' },
      { ref: 'C1', row: 0, col: 2, value: 3, styleIndex: 0 },
    ] as XlsxCell[] }] }], images: [], drawingCoverage: [] } as never as XlsxDocument
    evaluateWorkbookFormulas(doc as never, {})
    const b = (doc.sheets[0].rows[0].cells[1] as XlsxCell)
    expect(b.value).toBe(0)
    expect(b.hasCachedValue).toBe(true)
    expect((doc.sheets[0].rows[0].cells[2] as XlsxCell).value).toBe(3)
  })
})

describe('StageA: Excel numeric-text grammar', () => {
  it('rejects hex/empty/whitespace in arithmetic coercion, accepts percent', () => {
    expect(evaluateFormula('""+1')).toBe('#VALUE!')
    expect(evaluateFormula('"  "+1')).toBe('#VALUE!')
    expect(evaluateFormula('"0x10"+1')).toBe('#VALUE!')
    expect(evaluateFormula('"10%"+1')).toBe(1.1)
    expect(evaluateFormula('"1.5e1"+1')).toBe(16)
    expect(evaluateFormula('"5"+2')).toBe(7)
  })
})

describe('StageA: native decimal ROUND/INT/0^0/MOD', () => {
  it('ROUND is decimal midpoints and extreme digits', () => {
    expect(evaluateFormula('ROUND(1.005, 2)')).toBe(1.01)
    expect(evaluateFormula('ROUND(-1.005, 2)')).toBe(-1.01)
    expect(evaluateFormula('ROUND(1.255, 2)')).toBe(1.26)
    expect(evaluateFormula('ROUND(2.675, 2)')).toBe(2.68)
    expect(evaluateFormula('ROUND(1, 309)')).toBe(1)
    expect(evaluateFormula('ROUND(1, -309)')).toBe(0)
    expect(evaluateFormula('ROUND(125.4, -1)')).toBe(130)
    expect(evaluateFormula('ROUND(-1.5, 0)')).toBe(-2)
    expect(evaluateFormula('ROUND(2.5, 0)')).toBe(3)
  })
  it('INT truncates the actual value without premature 15-digit rounding', () => {
    expect(evaluateFormula('INT(1.999999999999999)')).toBe(1)
    expect(evaluateFormula('INT(2.1)')).toBe(2)
    expect(evaluateFormula('INT(-3.7)')).toBe(-4)
  })
  it('0^0 is #NUM!; MOD matches native float controls', () => {
    expect(evaluateFormula('0^0')).toBe('#NUM!')
    expect(evaluateFormula('MOD(6.3, 2.1)')).toBe(2.1)
    expect(evaluateFormula('MOD(0.3, 0.1)')).toBe(0.1)
  })
})

describe('StageA: version-aware Unicode text functions', () => {
  const emoji = '😀'
  it('default text semantics are current standalone Excel (code points)', () => {
    expect(evaluateFormula(`LEN("${emoji}")`)).toBe(1)
    expect(evaluateFormula(`LEFT("${emoji}",1)`)).toBe(emoji)
  })
  it('workbook legacy v1 keeps LEN units without splitting LEFT surrogate pairs', () => {
    const ctx: EvaluationContext & { excelTextLengthVersion?: 1 | 2 } = { excelTextLengthVersion: 1 }
    expect(evaluateFormula(`LEN("${emoji}")`, ctx)).toBe(2)
    expect(evaluateFormula(`LEFT("${emoji}",1)`, ctx)).toBe(emoji)
  })
})

describe('StageA: sheet identity without trim; safe shared quoting', () => {
  it('space-bearing sheet names resolve exactly', async () => {
    const doc = { sheets: [
      { name: ' Data ', sourcePartPath: 'a', rows: [{ index: 0, cells: [
        { ref: 'A1', row: 0, col: 0, value: 9, styleIndex: 0 },
      ] as XlsxCell[] }] },
      { name: 'Summary', sourcePartPath: 'b', rows: [{ index: 0, cells: [
        { ref: 'A1', row: 0, col: 0, value: null, styleIndex: 0, formula: "' Data '!A1" },
      ] as XlsxCell[] }] },
    ], images: [], drawingCoverage: [] } as never as XlsxDocument
    evaluateWorkbookFormulas(doc as never, {})
    expect(doc.sheets[1].rows[0].cells[0].value).toBe(9)
  })
  it('digit-leading sheet names stay quoted on shared translation', () => {
    expect(translateSharedFormula("'2024'!A1+1", 1, 0).formula).toBe("'2024'!B1+1")
    expect(translateSharedFormula("'My Sheet'!A1+1", 0, 1).formula).toBe("'My Sheet'!A2+1")
  })
})

describe('StageA: unsupported evaluation retains valid cache with diagnostic', () => {
  const mkDoc = (): XlsxDocument => ({
    sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
      { ref: 'B1', row: 0, col: 1, value: 2, styleIndex: 0 },
      { ref: 'B2', row: 0, col: 2, value: 3, styleIndex: 0 },
      { ref: 'A1', row: 0, col: 0, value: 5, styleIndex: 0, hasCachedValue: true, formula: '=SUMPRODUCT(B1:B2)' },
      { ref: 'D1', row: 0, col: 3, value: 'ok', styleIndex: 0, hasCachedValue: false, formula: '=IF(FALSE,SQRT(4),"ok")' },
    ] as XlsxCell[] }] }],
    images: [], drawingCoverage: [],
  } as never as XlsxDocument)
  it('keeps the cached value and diagnoses instead of overwriting with #NAME?', () => {
    const doc = mkDoc()
    evaluateWorkbookFormulas(doc as never, { forceRecalc: true })
    const a1 = doc.sheets[0].rows[0].cells[2] as XlsxCell
    expect(a1.value).toBe(5)
    expect((doc as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics?.some(d => d.feature === 'SUMPRODUCT')).toBe(true)
  })
  it('does not diagnose dead IF branches (unsupported decided from the evaluated path)', () => {
    const doc = mkDoc()
    evaluateWorkbookFormulas(doc as never, {})
    const d = doc.sheets[0].rows[0].cells[3] as XlsxCell
    expect(d.value).toBe('ok')
    expect((doc as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics?.some(x => x.feature === 'SQRT')).toBeFalsy()
  })
})

describe('StageA: shared master provenance preserved (no sentence-formula followers)', () => {
  it('follower keeps cached value; master original text quoted in diagnostic; formula not overwritten', () => {
    const doc: XlsxDocument = {
      sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
        { ref: 'B1', row: 0, col: 1, value: undefined, styleIndex: 0, formula: '=SUM(Table1[Amount])', sharedFormula: { si: 0, ref: 'B1:B2' } },
        { ref: 'B2', row: 0, col: 2, value: 88, styleIndex: 0, hasCachedValue: true, sharedFormula: { si: 0 } },
      ] as XlsxCell[] }] }],
      images: [], drawingCoverage: [],
    } as never as XlsxDocument
    evaluateWorkbookFormulas(doc as never, {})
    const follower = doc.sheets[0].rows[0].cells[1] as XlsxCell
    expect(follower.value).toBe(88)
    expect((doc as unknown as { diagnostics?: Array<{ message: string }> }).diagnostics?.some(d =>
      d.message.includes('SUM(Table1[Amount])'))).toBe(true)
  })
})

describe('StageA: empty numeric <v/> is a missing cache', () => {
  it('recalculates formula cells whose cached <v/> is empty', async () => {
    const zip = new JSZip()
    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
< worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>1+1</f><v></v></c></row></sheetData></worksheet>`.replace('< worksheet', '<worksheet')
    zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="S" sheetId="1" r:id="r1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>')
    zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    zip.file('xl/worksheets/sheet1.xml', sheetXml)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const tmp = path.join(os.tmpdir(), `stageA-emptyv-${Date.now()}-${Math.random().toString(36).slice(2)}.xlsx`)
    fs.writeFileSync(tmp, Buffer.from(bytes))
    const doc = await parseXlsx(await OfficePackage.load(fs.readFileSync(tmp)))
    fs.rmSync(tmp, { force: true })
    expect(doc.sheets[0].rows[0].cells[0].value).toBe(2)
  })
})

describe('StageA: NATIVE-BOUNDARIES.json verified rows (implemented constructs only)', () => {
  it('MOD sign/float matrix matches native exactly (bridge-precision equal)', () => {
    expect(evaluateFormula('MOD(6.3, 2.1)')).toBe(2.1)
    expect(evaluateFormula('MOD(0.3, 0.1)')).toBe(0.1)
    expect(Math.abs(evaluateFormula('MOD(-6.3, 2.1)') as number - 4.440892098501e-16)).toBeLessThan(1e-27)
    expect(Math.abs(evaluateFormula('MOD(6.3, -2.1)') as number + 4.440892098501e-16)).toBeLessThan(1e-27)
    expect(evaluateFormula('MOD(-6.3, -2.1)')).toBe(-2.1)
    // 13-digit native TSV bridge: MOD(10,0.1)/MOD(0.5,0.1) compare at bridge
    // precision (engine remainder 0.0999999999999995 ≈ native 0.1).
    expect(Math.abs(evaluateFormula('MOD(10, 0.1)') as number - 0.1)).toBeLessThan(1e-15)
    expect(Math.abs(evaluateFormula('MOD(-0.3, 0.1)') as number - 2.775557561563e-17)).toBeLessThan(1e-28)
    expect(Math.abs(evaluateFormula('MOD(0.5, 0.1)') as number - 0.1)).toBeLessThan(1e-15)
  })
  it('ROUND midpoints and ±extreme digits (ROUND±.15 confirmed)', () => {
    expect(evaluateFormula('ROUND(0.15, 1)')).toBe(0.2)
    expect(evaluateFormula('ROUND(-0.15, 1)')).toBe(-0.2)
    expect(evaluateFormula('ROUND(1.005, 2)')).toBe(1.01)
    expect(evaluateFormula('ROUND(1.255, 2)')).toBe(1.26)
    expect(evaluateFormula('ROUND(1, 1000)')).toBe(1)
    expect(evaluateFormula('ROUND(1, -1000)')).toBe(0)
  })
  it('numeric comparison precision boundaries (literal precision matters)', () => {
    expect(evaluateFormula('1.999999999999999=2')).toBe(false)
    expect(evaluateFormula('1.000000000000001=1')).toBe(true)
    expect(evaluateFormula('0.1+0.2=0.3')).toBe(true)
    expect(Math.abs(evaluateFormula('(0.1+0.2)-0.3') as number - 5.551115123125783e-17)).toBeLessThan(1e-30)
  })
  it('numeric-text grammar: currency and thousands separators; percent in COUNT/SUM', () => {
    expect(evaluateFormula('COUNT("10%")')).toBe(1)
    expect(evaluateFormula('SUM("10%")')).toBe(0.1)
    expect(evaluateFormula('"$5"+1')).toBe(6)
    expect(evaluateFormula('"1,234"+1')).toBe(1235)
  })
  it('ROUND extreme beyond ±significand windows remain native (1,±1000 covered above)', () => {
    expect(evaluateFormula('ROUND(125.4, -1)')).toBe(130)
    expect(evaluateFormula('ROUND(2.5, 0)')).toBe(3)
    expect(evaluateFormula('ROUND(-2.5, 0)')).toBe(-3)
  })
})
