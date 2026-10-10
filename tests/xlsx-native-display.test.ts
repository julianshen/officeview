import { describe, expect, it } from 'vitest'
import { evaluateFormula } from '../src/xlsx/formula/evaluator'
import { formatValue, renderSheet } from '../src/xlsx/render'
import type { XlsxSheet } from '../src/xlsx/types'

describe('native Excel function-specific Unicode profile', () => {
  const legacy = { excelTextLengthVersion: 1 as const }
  it('LEFT retains the full emoji even when LEN and MID use UTF-16 units', () => {
    expect(evaluateFormula('LEFT("😀x",1)', legacy)).toBe('😀')
    expect(evaluateFormula('LEN("😀")', legacy)).toBe(2)
    expect(evaluateFormula('MID("😀x",2,1)', legacy)).toBe('\uDE00')
  })
  it('RIGHT retains the full emoji in the same legacy workbook profile', () => {
    expect(evaluateFormula('RIGHT("x😀",1)', legacy)).toBe('😀')
    expect(evaluateFormula('RIGHT("x😀",0)', legacy)).toBe('')
  })
})

describe('native Excel General and time display profiles', () => {
  it('retains General precision instead of rounding to two decimal places', () => {
    expect(formatValue(3.14159265358979, 0)).toBe('3.141592654')
    expect(formatValue(0.000123456789, 0)).toBe('0.000123457')
    expect(formatValue(123456789012, 0)).toBe('1.23457E+11')
  })
  it.each([
    [-3.14159265358979,'-3.141592654'],[12345678901,'12345678901'],
    [123456789012,'1.23457E+11'],[0.000000123456789,'1.23457E-07'],
    [0.0000000123456789,'1.23457E-08'],[0.123456789012345,'0.123456789'],
    [12345.6789012345,'12345.6789'],[0,'0'],[1e308,'1E+308'],
    [0.0001,'0.0001'],[0.00001,'0.00001'],[0.000001,'0.000001'],
    [0.0000001,'0.0000001'],[0.0000123456789,'1.23457E-05'],
    [1e-8,'0.00000001'],[1e-9,'0.000000001'],[1e-10,'1E-10'],
  ] as const)('matches supplementary wide-column native General %s', (value,text) => {
    expect(formatValue(value,0)).toBe(text)
  })
  it('keeps the character budget when rounding crosses an integer boundary', () => {
    expect(formatValue(99999999999.9,0)).toBe('1E+11')
  })
  it.each([[18,'12:34 PM'],[19,'12:34:56 PM'],[20,'12:34'],[21,'12:34:56']] as const)('renders built-in time %s as %s', (id,text) => {
    expect(formatValue(0.524259259259259, id)).toBe(text)
  })
  it('renders date-plus-time 22 through the shared formatter', () => {
    expect(formatValue(45351.524259259259, 22)).toBe('2/29/24 12:34')
  })
  it('uses the cell width when choosing General precision', () => {
    const painted:string[]=[]
    const sheet:XlsxSheet={name:'S',cols:[],merges:[],mergeRanges:[],rows:[{index:0,cells:[{ref:'A1',row:0,col:0,styleIndex:0,value:3.14159265358979}]}]}
    // Fixed glyph widths pin the layout result independently of host fonts.
    const ctx=new Proxy({font:'',measureText:(s:string)=>({width:s.length*7}),fillText:(s:string)=>painted.push(s)},{get:(o,k)=>k in o?o[k as keyof typeof o]:()=>{}}) as unknown as CanvasRenderingContext2D
    renderSheet(sheet,ctx,{colWidthsPx:[50],rowHeightsPx:[20],widthPx:50,heightPx:20})
    expect(painted).toEqual(['3.1416'])
  })
})
