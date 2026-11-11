/**
 * Deterministic raw SpreadsheetML generator for XLSX Conditional Formatting Stage 1 oracle workbooks.
 *
 *   bun scripts/cf-refs/generate.ts [--case <id>] [--anchor-date YYYY-MM-DD]
 *
 * Rules:
 * - Writes `tests/fixtures/xlsx-cf/inputs/<id>.xlsx`.
 * - Seeds `tests/fixtures/xlsx-cf/<id>.xlsx` ONLY when no verified capture or Excel-resaved output exists yet
 *   (`resavedByExcel !== true` and `savedSha256 === inputSha256`), preserving pending Excel-resaved artifacts.
 * - When `--case <id>` (`onlyCaseId`) is specified, only that case's files and manifest entry are touched;
 *   all other cases' saved workbooks and manifest entries remain untouched.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { deflateRawSync } from 'node:zlib'
import {
  CF_CASES,
  CF_STAGE3_FIXTURES,
  DEFAULT_CLOCK_ANCHOR_DATE,
  isoDateToExcelSerial,
  s3ControlOffset,
  s3ControlRef,
  s3Range,
  type CfBarColors,
  type CfBarCfvo,
  type CfManifest,
  type CfOracleCase,
  type S3BarRule,
  type S3Fixture,
  type S3Value,
} from './cases'
import { verifyCfCase } from './verify'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
export const CF_FIXTURE_DIR = join(REPO_ROOT, 'tests/fixtures/xlsx-cf')
export const CF_INPUT_DIR = join(CF_FIXTURE_DIR, 'inputs')
export const CF_REFS_DIR = join(CF_FIXTURE_DIR, 'refs')
export const CF_MANIFEST_PATH = join(CF_FIXTURE_DIR, 'manifest.json')

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const NS_X14AC = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const X14_CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

const sha256Hex = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')

const colLetters = (col1: number): string => {
  let s = ''
  for (let c = col1; c > 0; c = Math.floor((c - 1) / 26)) {
    s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  }
  return s
}

type CellSpec =
  | { ref: string; v: number | string | boolean | null; s?: number; f?: string; err?: undefined }
  | { ref: string; err: string; s?: number; v?: undefined; f?: undefined }

interface SheetPayload {
  dimension: string
  date1904?: boolean
  cells: readonly CellSpec[]
  merges?: readonly string[]
  hiddenCols?: readonly number[]
  hiddenRows?: readonly number[]
  colWidths?: ReadonlyArray<{ min: number; max: number; width: number }>
  tablePartXml?: string
  pivotParts?: {
    pivotTableXml: string
    pivotCacheDefXml: string
    pivotCacheRecordsXml: string
    pivotSrcSheetXml: string
  }
  legacyCfXml: string
  x14CfXml?: string
  dxfsXml: string
  /** Stage 3: per-row heights in points (default: 18, the Stage 1/2 value). */
  rowHeightsPt?: ReadonlyMap<number, number>
  defaultRowHeightPt?: number
  /** Stage 3: `false` writes `showGridLines="0"` (default on). */
  showGridLines?: boolean
  /** Stage 3: raw `<autoFilter>` element, written after `<sheetData>`. */
  autoFilterXml?: string
  /** Stage 3: add a date-formatted `cellXfs` entry (style 4). */
  dateStyle?: boolean
}

const DEFAULT_DXFS_XML = [
  '<dxfs count="3">',
  '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
  '<dxf><font><color rgb="FF006100"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>',
  '<dxf><font><color rgb="FF9C6500"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>',
  '</dxfs>',
].join('')

function replicateWithControlBlock(
  headers: readonly string[],
  inputRows: ReadonlyArray<{ label: string; values: ReadonlyArray<CellSpec['v'] | { err: string } | { f: string; v: number }>; s?: number }>,
  controlRowOffset: number,
): CellSpec[] {
  const out: CellSpec[] = []
  // Header row 1 and control header row (1 + controlRowOffset)
  headers.forEach((h, cIdx) => {
    const col = colLetters(cIdx + 1)
    out.push({ ref: `${col}1`, v: h })
    out.push({ ref: `${col}${1 + controlRowOffset}`, v: `${h}_Ctrl` })
  })

  inputRows.forEach((row, rIdx) => {
    const r = rIdx + 2
    const rc = r + controlRowOffset
    out.push({ ref: `A${r}`, v: row.label })
    out.push({ ref: `A${rc}`, v: `${row.label}_ctrl` })
    row.values.forEach((cellVal, colOffset) => {
      const col = colLetters(colOffset + 2)
      for (const targetRow of [r, rc]) {
        const ref = `${col}${targetRow}`
        if (cellVal !== null && typeof cellVal === 'object' && 'err' in cellVal) {
          out.push({ ref, err: cellVal.err, s: row.s })
        } else if (cellVal !== null && typeof cellVal === 'object' && 'f' in cellVal) {
          out.push({ ref, f: cellVal.f, v: cellVal.v, s: row.s })
        } else if (cellVal !== undefined) {
          out.push({ ref, v: cellVal, s: row.s })
        }
      }
    })
  })
  return out
}

function casePayload(id: string, anchorDateIso = DEFAULT_CLOCK_ANCHOR_DATE): SheetPayload {
  switch (id) {
    case 'stage1-cellis-literals': {
      const headers = [
        'Input',
        'Eq100',
        'Ne100',
        'Gt100',
        'Lt100',
        'Ge100',
        'Le100',
        'Btw1_3',
        'RevBtw3_1',
        'NotBtw1_3',
        'RevNotBtw3_1',
        'EqStr100',
        'EqTrue',
        'EqNA',
        'EqEmptyStr',
      ]
      const rawInputs: Array<{ label: string; val: number | string | boolean | null | { err: string } }> = [
        { label: 'v=-1', val: -1 },
        { label: 'v=0', val: 0 },
        { label: 'v=1', val: 1 },
        { label: 'v=99', val: 99 },
        { label: 'v=100', val: 100 },
        { label: 'v=101', val: 101 },
        { label: 'str="100"', val: '100' },
        { label: 'bool=TRUE', val: true },
        { label: 'bool=FALSE', val: false },
        { label: 'err=#N/A', val: { err: '#N/A' } },
        { label: 'empty=""', val: '' },
        { label: 'blank', val: null },
      ]
      const rows = rawInputs.map(item => ({
        label: item.label,
        values: Array.from({ length: 14 }, () => item.val),
      }))
      return {
        dimension: 'A1:O26',
        cells: replicateWithControlBlock(headers, rows, 13),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B13"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C13"><cfRule type="cellIs" dxfId="1" priority="2" operator="notEqual"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D13"><cfRule type="cellIs" dxfId="0" priority="3" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E13"><cfRule type="cellIs" dxfId="1" priority="4" operator="lessThan"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F13"><cfRule type="cellIs" dxfId="0" priority="5" operator="greaterThanOrEqual"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G13"><cfRule type="cellIs" dxfId="1" priority="6" operator="lessThanOrEqual"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H13"><cfRule type="cellIs" dxfId="2" priority="7" operator="between"><formula>1</formula><formula>3</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I13"><cfRule type="cellIs" dxfId="2" priority="8" operator="between"><formula>3</formula><formula>1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="J2:J13"><cfRule type="cellIs" dxfId="1" priority="9" operator="notBetween"><formula>1</formula><formula>3</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="K2:K13"><cfRule type="cellIs" dxfId="1" priority="10" operator="notBetween"><formula>3</formula><formula>1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="L2:L13"><cfRule type="cellIs" dxfId="0" priority="11" operator="equal"><formula>"100"</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="M2:M13"><cfRule type="cellIs" dxfId="1" priority="12" operator="equal"><formula>TRUE</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="N2:N13"><cfRule type="cellIs" dxfId="2" priority="13" operator="equal"><formula>#N/A</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="O2:O13"><cfRule type="cellIs" dxfId="0" priority="14" operator="equal"><formula>""</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-text-predicates': {
      const headers = [
        'Input',
        'Contains',
        'NotContains',
        'Begins',
        'Ends',
        'WsNeedle',
        'QuoteNeedle',
        'Fmt00000',
        'AbsentF',
        'DivF',
        'X14Canon',
        'X14Div',
      ]
      const rawInputs: Array<{ label: string; val: number | string | null; s?: number }> = [
        { label: 'lower', val: 'abc' },
        { label: 'mixed', val: 'AbC' },
        { label: 'accent', val: 'ábc' },
        { label: 'padded', val: ' ABC ' },
        { label: 'numFmt', val: 123, s: 2 },
        { label: 'quoted', val: 'a"b' },
        { label: 'emptyStr', val: '' },
        { label: 'blank', val: null },
      ]
      const rows = rawInputs.map(item => ({
        label: item.label,
        values: Array.from({ length: 11 }, () => item.val),
        s: item.s,
      }))
      return {
        dimension: 'A1:L18',
        cells: replicateWithControlBlock(headers, rows, 9),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B9"><cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH("ab",B2)))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C9"><cfRule type="notContainsText" dxfId="1" priority="2" operator="notContains" text="ab"><formula>ISERROR(SEARCH("ab",C2))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D9"><cfRule type="beginsWith" dxfId="2" priority="3" operator="beginsWith" text="ab"><formula>LEFT(D2,2)="ab"</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E9"><cfRule type="endsWith" dxfId="0" priority="4" operator="endsWith" text="bc"><formula>RIGHT(E2,2)="bc"</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F9"><cfRule type="containsText" dxfId="1" priority="5" operator="containsText" text=" ABC "><formula>NOT(ISERROR(SEARCH(" ABC ",F2)))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G9"><cfRule type="containsText" dxfId="2" priority="6" operator="containsText" text="a&quot;b"><formula>NOT(ISERROR(SEARCH("a""b",G2)))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H9"><cfRule type="beginsWith" dxfId="0" priority="7" operator="beginsWith" text="00"><formula>LEFT(H2,2)="00"</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I9"><cfRule type="containsText" dxfId="1" priority="8" operator="containsText" text="ab"/></conditionalFormatting>',
          '<conditionalFormatting sqref="J2:J9"><cfRule type="containsText" dxfId="2" priority="9" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH("zz",J2)))</formula></cfRule></conditionalFormatting>',
        ].join(''),
        x14CfXml: [
          '<x14:conditionalFormattings>',
          '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">',
          '<x14:cfRule type="containsText" priority="10" operator="containsText" text="ab" id="{11111111-2222-3333-4444-555555555501}">',
          '<xm:f>NOT(ISERROR(SEARCH("ab",K2)))</xm:f>',
          '<x14:dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></x14:dxf>',
          '</x14:cfRule>',
          '<xm:sqref>K2:K9</xm:sqref>',
          '</x14:conditionalFormatting>',
          '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">',
          '<x14:cfRule type="containsText" priority="11" operator="containsText" text="ab" id="{11111111-2222-3333-4444-555555555502}">',
          '<xm:f>NOT(ISERROR(SEARCH("zz",L2)))</xm:f>',
          '<x14:dxf><font><color rgb="FF006100"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></x14:dxf>',
          '</x14:cfRule>',
          '<xm:sqref>L2:L9</xm:sqref>',
          '</x14:conditionalFormatting>',
          '</x14:conditionalFormattings>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-blanks-errors': {
      const headers = [
        'Input',
        'Blank_Can',
        'Blank_Abs',
        'Blank_Div',
        'NotBlank_Can',
        'NotBlank_Abs',
        'NotBlank_Div',
        'Err_Can',
        'Err_Abs',
        'Err_Div',
        'NotErr_Can',
        'NotErr_Abs',
        'NotErr_Div',
        'X14_Blank',
        'X14_NotErr',
      ]
      const rawInputs: Array<{ label: string; val: number | string | boolean | null | { err: string } }> = [
        { label: 'unstoredBlank', val: null },
        { label: 'emptyStr', val: '' },
        { label: 'space', val: ' ' },
        { label: 'tab', val: '\t' },
        { label: 'zero', val: 0 },
        { label: 'false', val: false },
        { label: 'errNA', val: { err: '#N/A' } },
        { label: 'errDiv0', val: { err: '#DIV/0!' } },
      ]
      const rows = rawInputs.map(item => ({
        label: item.label,
        values: Array.from({ length: 14 }, () => item.val),
      }))
      return {
        dimension: 'A1:O18',
        cells: replicateWithControlBlock(headers, rows, 9),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B9"><cfRule type="containsBlanks" dxfId="0" priority="1"><formula>LEN(TRIM(B2))=0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C9"><cfRule type="containsBlanks" dxfId="0" priority="2"/></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D9"><cfRule type="containsBlanks" dxfId="0" priority="3"><formula>ISBLANK(D2)</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E9"><cfRule type="notContainsBlanks" dxfId="1" priority="4"><formula>LEN(TRIM(E2))&gt;0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F9"><cfRule type="notContainsBlanks" dxfId="1" priority="5"/></conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G9"><cfRule type="notContainsBlanks" dxfId="1" priority="6"><formula>NOT(ISBLANK(G2))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H9"><cfRule type="containsErrors" dxfId="2" priority="7"><formula>ISERROR(H2)</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I9"><cfRule type="containsErrors" dxfId="2" priority="8"/></conditionalFormatting>',
          '<conditionalFormatting sqref="J2:J9"><cfRule type="containsErrors" dxfId="2" priority="9"><formula>ISNA(J2)</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="K2:K9"><cfRule type="notContainsErrors" dxfId="1" priority="10"><formula>NOT(ISERROR(K2))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="L2:L9"><cfRule type="notContainsErrors" dxfId="1" priority="11"/></conditionalFormatting>',
          '<conditionalFormatting sqref="M2:M9"><cfRule type="notContainsErrors" dxfId="1" priority="12"><formula>NOT(ISNA(M2))</formula></cfRule></conditionalFormatting>',
        ].join(''),
        x14CfXml: [
          '<x14:conditionalFormattings>',
          '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">',
          '<x14:cfRule type="containsBlanks" priority="13" id="{22222222-3333-4444-5555-666666666601}">',
          '<xm:f>LEN(TRIM(N2))=0</xm:f>',
          '<x14:dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></x14:dxf>',
          '</x14:cfRule>',
          '<xm:sqref>N2:N9</xm:sqref>',
          '</x14:conditionalFormatting>',
          '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">',
          '<x14:cfRule type="notContainsErrors" priority="14" id="{22222222-3333-4444-5555-666666666602}">',
          '<xm:f>NOT(ISERROR(O2))</xm:f>',
          '<x14:dxf><font><color rgb="FF006100"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></x14:dxf>',
          '</x14:cfRule>',
          '<xm:sqref>O2:O9</xm:sqref>',
          '</x14:conditionalFormatting>',
          '</x14:conditionalFormattings>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-timeperiod-today-clock': {
      const D = isoDateToExcelSerial(anchorDateIso, '1900')
      const headers = [
        'Offset',
        'Today',
        'Yesterday',
        'Tomorrow',
        'Last7Days',
        'ThisWeek',
        'LastWeek',
        'NextWeek',
        'ThisMonth',
        'LastMonth',
        'NextMonth',
      ]
      const rawInputs: Array<{ label: string; val: number | { f: string; v: number } }> = [
        { label: 'D-2', val: D - 2 },
        { label: 'D-1_UTC', val: D - 1 },
        { label: 'D_Local', val: D },
        { label: 'D+0.5', val: D + 0.5 },
        { label: 'D+0.999', val: D + 0.999 },
        { label: 'D+1', val: D + 1 },
        { label: 'D+2', val: D + 2 },
        { label: 'D-6_7d', val: D - 6 },
        { label: 'D-7_LastWk', val: D - 7 },
        { label: 'D+7_NextWk', val: D + 7 },
        { label: 'D-31_LastMo', val: D - 31 },
        { label: 'D+30_NextMo', val: D + 30 },
        { label: 'YearEdge', val: D - 310 },
        { label: 'StaleTodayCache', val: { f: 'TODAY()', v: D - 5 } },
      ]
      const rows = rawInputs.map(item => ({
        label: item.label,
        values: Array.from({ length: 10 }, () => item.val),
      }))
      return {
        dimension: 'A1:K30',
        date1904: false,
        cells: replicateWithControlBlock(headers, rows, 15),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B15"><cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="today"><formula>FLOOR(B2,1)=TODAY()</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C15"><cfRule type="timePeriod" dxfId="1" priority="2" timePeriod="yesterday"><formula>FLOOR(C2,1)=TODAY()-1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D15"><cfRule type="timePeriod" dxfId="2" priority="3" timePeriod="tomorrow"><formula>FLOOR(D2,1)=TODAY()+1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E15"><cfRule type="timePeriod" dxfId="0" priority="4" timePeriod="last7Days"><formula>AND(TODAY()-FLOOR(E2,1)&lt;=6,FLOOR(E2,1)&lt;=TODAY())</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F15"><cfRule type="timePeriod" dxfId="1" priority="5" timePeriod="thisWeek"><formula>AND(TODAY()-ROUNDDOWN(F2,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(F2,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G15"><cfRule type="timePeriod" dxfId="2" priority="6" timePeriod="lastWeek"><formula>AND(TODAY()-ROUNDDOWN(G2,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(G2,0)&lt;(WEEKDAY(TODAY())+7))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H15"><cfRule type="timePeriod" dxfId="0" priority="7" timePeriod="nextWeek"><formula>AND(ROUNDDOWN(H2,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(H2,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I15"><cfRule type="timePeriod" dxfId="1" priority="8" timePeriod="thisMonth"><formula>AND(MONTH(I2)=MONTH(TODAY()),YEAR(I2)=YEAR(TODAY()))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="J2:J15"><cfRule type="timePeriod" dxfId="2" priority="9" timePeriod="lastMonth"><formula>AND(MONTH(J2)=MONTH(EDATE(TODAY(),0-1)),YEAR(J2)=YEAR(EDATE(TODAY(),0-1)))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="K2:K15"><cfRule type="timePeriod" dxfId="0" priority="10" timePeriod="nextMonth"><formula>AND(MONTH(K2)=MONTH(EDATE(TODAY(),0+1)),YEAR(K2)=YEAR(EDATE(TODAY(),0+1)))</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-timeperiod-today-clock-2': {
      const D = isoDateToExcelSerial(anchorDateIso, '1904')
      const headers = ['Offset1904', 'Today1904', 'Yesterday1904', 'Tomorrow1904', 'ThisWeek1904', 'ThisMonth1904']
      const rawInputs: Array<{ label: string; val: number | { f: string; v: number } }> = [
        { label: 'D-1_UTC', val: D - 1 },
        { label: 'D_Local', val: D },
        { label: 'D+0.5', val: D + 0.5 },
        { label: 'D+1', val: D + 1 },
        { label: 'D-40', val: D - 40 },
        { label: 'StaleTodayCache', val: { f: 'TODAY()', v: D - 5 } },
      ]
      const rows = rawInputs.map(item => ({
        label: item.label,
        values: Array.from({ length: 5 }, () => item.val),
      }))
      return {
        dimension: 'A1:F16',
        date1904: true,
        cells: replicateWithControlBlock(headers, rows, 8),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B7"><cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="today"><formula>FLOOR(B2,1)=TODAY()</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C7"><cfRule type="timePeriod" dxfId="1" priority="2" timePeriod="yesterday"><formula>FLOOR(C2,1)=TODAY()-1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D7"><cfRule type="timePeriod" dxfId="2" priority="3" timePeriod="tomorrow"><formula>FLOOR(D2,1)=TODAY()+1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E7"><cfRule type="timePeriod" dxfId="1" priority="4" timePeriod="thisWeek"><formula>AND(TODAY()-ROUNDDOWN(E2,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(E2,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F7"><cfRule type="timePeriod" dxfId="0" priority="5" timePeriod="thisMonth"><formula>AND(MONTH(F2)=MONTH(TODAY()),YEAR(F2)=YEAR(TODAY()))</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-top10-rank-edge': {
      const headers = [
        'Item',
        'Top2',
        'Bot2Ties',
        'Top99',
        'Bot99',
        'Pct0',
        'Pct1',
        'Pct40',
        'Pct100',
        'OmitRank',
        'ZeroRank',
      ]
      const rawInputs: Array<{ label: string; val: number | string | boolean | { err: string } }> = [
        { label: 'n1', val: 1 },
        { label: 'n2a', val: 2 },
        { label: 'n2b', val: 2 },
        { label: 'n3', val: 3 },
        { label: 'n9', val: 9 },
        { label: 'str999', val: '999' },
        { label: 'boolTrue', val: true },
        { label: 'errNA', val: { err: '#N/A' } },
        { label: 'emptyStr', val: '' },
      ]
      const rows = rawInputs.map(item => ({
        label: item.label,
        values: Array.from({ length: 10 }, () => item.val),
      }))
      return {
        dimension: 'A1:K20',
        cells: replicateWithControlBlock(headers, rows, 10),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B10"><cfRule type="top10" dxfId="0" priority="1" rank="2"/></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C10"><cfRule type="top10" dxfId="1" priority="2" bottom="1" rank="2"/></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D10"><cfRule type="top10" dxfId="2" priority="3" rank="99"/></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E10"><cfRule type="top10" dxfId="0" priority="4" bottom="1" rank="99"/></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F10"><cfRule type="top10" dxfId="1" priority="5" percent="1" rank="0"/></conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G10"><cfRule type="top10" dxfId="2" priority="6" percent="1" rank="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H10"><cfRule type="top10" dxfId="0" priority="7" percent="1" rank="40"/></conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I10"><cfRule type="top10" dxfId="1" priority="8" percent="1" rank="100"/></conditionalFormatting>',
          '<conditionalFormatting sqref="J2:J10"><cfRule type="top10" dxfId="2" priority="9"/></conditionalFormatting>',
          '<conditionalFormatting sqref="K2:K10"><cfRule type="top10" dxfId="0" priority="10" rank="0"/></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-above-average-stddev': {
      const headers = [
        'Row',
        'Above',
        'Below',
        'EqAbove',
        'EqBelow',
        'Std1_Abv',
        'Std1_Bel',
        'Std2_Abv',
        'Std3_Abv',
        'Std0_Abv',
        'StdNeg1',
        'Std4_Abv',
        'Std1_EqAbv',
      ]
      const datasetA: Array<number | string | boolean | null | { err: string }> = [
        0,
        2,
        4,
        '10',
        true,
        { err: '#DIV/0!' },
        '',
        null,
        null,
        null,
        null,
      ]
      const datasetB: Array<number | string | boolean | null | { err: string }> = [
        0,
        0,
        10,
        11,
        '20',
        false,
        { err: '#N/A' },
        null,
        null,
        null,
        null,
      ]
      const datasetG: Array<number | string | boolean | null | { err: string }> = [
        -2,
        2,
        10,
        11,
        '20',
        false,
        { err: '#N/A' },
        null,
        null,
        null,
        null,
      ]
      const datasetH: Array<number | null> = [0, 0, 0, 0, 0, 0, 0, 100, null, null, null]
      const datasetI: Array<number> = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 100]
      const rows = datasetA.map((valA, idx) => {
        const valB = datasetB[idx]
        const valG = datasetG[idx]
        const valH = datasetH[idx]
        const valI = datasetI[idx]
        return {
          label: `r${idx + 1}`,
          values: [valA, valA, valA, valA, valB, valG, valH, valI, valB, valB, valI, valB],
        }
      })
      return {
        dimension: 'A1:M25',
        cells: replicateWithControlBlock(headers, rows, 12),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B12"><cfRule type="aboveAverage" dxfId="0" priority="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C12"><cfRule type="aboveAverage" dxfId="1" priority="2" aboveAverage="0"/></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D12"><cfRule type="aboveAverage" dxfId="2" priority="3" equalAverage="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E12"><cfRule type="aboveAverage" dxfId="0" priority="4" aboveAverage="0" equalAverage="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F12"><cfRule type="aboveAverage" dxfId="1" priority="5" stdDev="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G12"><cfRule type="aboveAverage" dxfId="2" priority="6" aboveAverage="0" stdDev="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H12"><cfRule type="aboveAverage" dxfId="0" priority="7" stdDev="2"/></conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I12"><cfRule type="aboveAverage" dxfId="1" priority="8" stdDev="3"/></conditionalFormatting>',
          '<conditionalFormatting sqref="J2:J12"><cfRule type="aboveAverage" dxfId="2" priority="9" stdDev="0"/></conditionalFormatting>',
          '<conditionalFormatting sqref="K2:K12"><cfRule type="aboveAverage" dxfId="0" priority="10" stdDev="-1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="L2:L12"><cfRule type="aboveAverage" dxfId="1" priority="11" stdDev="4"/></conditionalFormatting>',
          '<conditionalFormatting sqref="M2:M12"><cfRule type="aboveAverage" dxfId="2" priority="12" equalAverage="1" stdDev="1"/></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-duplicate-scope': {
      const headers = ['Row', 'DupSingle', 'UniqSingle', 'DisjointA', 'DisjointGap', 'DisjointB', 'OverlapGap', 'OverlapDup', 'OverlapUniq']
      const singleCol: Array<number | string | null | { err: string }> = [
        1,
        1,
        '1',
        'A',
        'a',
        'á',
        ' a ',
        { err: '#N/A' },
        { err: '#N/A' },
        '',
        null,
      ]
      const disjointA: Array<number | null> = [10, 20, 30, null, null, null, null, null, null, null, null]
      const disjointB: Array<number | null> = [20, 40, 50, null, null, null, null, null, null, null, null]
      const overlapVals: Array<number | null> = [66, 77, 88, 88, null, null, null, null, null, null, null]
      const rows = singleCol.map((sv, idx) => ({
        label: `r${idx + 1}`,
        values: [
          sv,
          sv,
          disjointA[idx],
          null,
          disjointB[idx],
          null,
          overlapVals[idx],
          overlapVals[idx],
        ],
      }))
      return {
        dimension: 'A1:I24',
        cells: replicateWithControlBlock(headers, rows, 12),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B12"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C12"><cfRule type="uniqueValues" dxfId="1" priority="2"/></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D4 F2:F4"><cfRule type="duplicateValues" dxfId="2" priority="3"/></conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H4 H3:H5"><cfRule type="duplicateValues" dxfId="0" priority="4"/></conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I4 I3:I5"><cfRule type="uniqueValues" dxfId="1" priority="5"/></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-priority-composition': {
      const headers = ['Val', 'FillWins', 'OrthoMerge', 'BareStop', 'EqualPrio', 'CrossVisual', 'UnresStub']
      const rows = [
        { label: 'v=25', values: [25, 25, 25, 25, 25, 25] },
        { label: 'v=15_g14', values: [15, 15, 15, 15, 15, 14] },
        { label: 'v=5', values: [5, 5, 5, 5, 5, 5] },
        { label: 'v=0', values: [0, 0, 0, 0, 0, 0] },
      ]
      return {
        dimension: 'A1:G10',
        cells: replicateWithControlBlock(headers, rows, 5),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B5">',
          '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule>',
          '<cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C5">',
          '<cfRule type="cellIs" dxfId="5" priority="3" operator="greaterThan"><formula>10</formula></cfRule>',
          '<cfRule type="cellIs" dxfId="2" priority="4" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D5">',
          '<cfRule type="cellIs" priority="5" stopIfTrue="1" operator="greaterThan"><formula>20</formula></cfRule>',
          '<cfRule type="cellIs" dxfId="3" priority="6" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E5">',
          '<cfRule type="cellIs" dxfId="0" priority="7" operator="greaterThan"><formula>10</formula></cfRule>',
          '<cfRule type="cellIs" dxfId="1" priority="7" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F5">',
          '<cfRule type="cellIs" dxfId="4" priority="8" operator="greaterThan"><formula>20</formula></cfRule>',
          '<cfRule type="colorScale" priority="9"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '<cfRule type="dataBar" priority="10"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>',
          '<cfRule type="iconSet" priority="11"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G5">',
          '<cfRule type="expression" dxfId="1" priority="12" stopIfTrue="1"><formula>MOD(G2,2)=1</formula></cfRule>',
          '<cfRule type="cellIs" dxfId="0" priority="13" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
        ].join(''),
        dxfsXml: [
          '<dxfs count="6">',
          '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>',
          '<dxf><font><b/><color rgb="FF006100"/></font></dxf>',
          '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>',
          '<dxf><font><b/><color rgb="FF9C0006"/></font></dxf>',
          '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CD"/></patternFill></fill></dxf>',
          '</dxfs>',
        ].join(''),
      }
    }

    case 'stage1-dxf-layering': {
      const headers = ['Row', 'ClearFontBorder', 'ClearFillNone', 'TableLayer', 'PivotLayer', 'ThemeNeighbor']
      const vals = [150, 200, 50, 180, 190]
      const rows = vals.map((v, idx) => ({
        label: `r${idx + 1}`,
        values: [v, v, v, v, v],
        s: 1,
      }))
      return {
        dimension: 'A1:F12',
        cells: replicateWithControlBlock(headers, rows, 6),
        tablePartXml: [
          XML_DECL,
          `<table xmlns="${NS_MAIN}" id="1" name="Table1" displayName="Table1" ref="D1:D6" totalsRowShown="0">`,
          '<tableColumns count="1"><tableColumn id="1" name="TableLayer"/></tableColumns>',
          '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
          '</table>',
        ].join(''),
        pivotParts: {
          pivotTableXml: [
            XML_DECL,
            `<pivotTableDefinition xmlns="${NS_MAIN}" name="PivotTable1" cacheId="1" dataOnRows="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="1" applyAlignmentFormats="0" applyWidthHeightFormats="0" dataCaption="Values" createdVersion="5" updatedVersion="5" minRefreshableVersion="3">`,
            '<location ref="E1:E6" firstHeaderRow="0" firstDataRow="1" firstDataCol="0"/>',
            '<pivotFields count="1"><pivotField dataField="1" showAll="0"/></pivotFields>',
            '<dataFields count="1"><dataField name="Sum of PivotLayer" fld="0" baseField="0" baseItem="0"/></dataFields>',
            '<pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="0" showLastColumn="1"/>',
            '</pivotTableDefinition>',
          ].join(''),
          pivotCacheDefXml: [
            XML_DECL,
            `<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" r:id="rIdCacheRec1" refreshedBy="OfficeView" createdVersion="5" refreshedVersion="5" minRefreshableVersion="3" recordCount="5">`,
            '<cacheSource type="worksheet"><worksheetSource ref="A1:A6" sheet="PivotSrc"/></cacheSource>',
            '<cacheFields count="1"><cacheField name="PivotLayer" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="50" maxValue="200"/></cacheField></cacheFields>',
            '</pivotCacheDefinition>',
          ].join(''),
          pivotCacheRecordsXml: [
            XML_DECL,
            `<pivotCacheRecords xmlns="${NS_MAIN}" count="5">`,
            '<r><n v="150"/></r><r><n v="200"/></r><r><n v="50"/></r><r><n v="180"/></r><r><n v="190"/></r>',
            '</pivotCacheRecords>',
          ].join(''),
          pivotSrcSheetXml: [
            XML_DECL,
            `<worksheet xmlns="${NS_MAIN}"><dimension ref="A1:A6"/><sheetData>`,
            '<row r="1"><c r="A1" t="inlineStr"><is><t>PivotLayer</t></is></c></row>',
            '<row r="2"><c r="A2"><v>150</v></c></row>',
            '<row r="3"><c r="A3"><v>200</v></c></row>',
            '<row r="4"><c r="A4"><v>50</v></c></row>',
            '<row r="5"><c r="A5"><v>180</v></c></row>',
            '<row r="6"><c r="A6"><v>190</v></c></row>',
            '</sheetData></worksheet>',
          ].join(''),
        },
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B6"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C6"><cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D6"><cfRule type="cellIs" dxfId="2" priority="3" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E6"><cfRule type="cellIs" dxfId="3" priority="4" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F6"><cfRule type="cellIs" dxfId="4" priority="5" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: [
          '<dxfs count="5">',
          '<dxf><font><b val="0"/><u val="none"/><strike val="0"/><color rgb="FF006100"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill><border><top style="thin"><color rgb="FF006100"/></top><bottom style="thin"><color rgb="FF006100"/></bottom></border></dxf>',
          '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="none"/></fill></dxf>',
          '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><font><color rgb="FF9C6500"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>',
          '<dxf><fill><patternFill patternType="solid"><bgColor theme="4" tint="0.39997558519241921"/></patternFill></fill><border><top style="medium"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF9C0006"/></bottom></border></dxf>',
          '</dxfs>',
        ].join(''),
      }
    }

    case 'stage1-dxf-layering-2': {
      const headers = ['Row', 'GradFill', 'PatFill', 'DiagBorder', 'Alignment', 'Protection']
      const vals = [10, 20, 30]
      const rows = vals.map((v, idx) => ({
        label: `u${idx + 1}`,
        values: [v, v, v, v, v],
      }))
      return {
        dimension: 'A1:F6',
        cells: replicateWithControlBlock(headers, rows, 3),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C4"><cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D4"><cfRule type="cellIs" dxfId="2" priority="3" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E4"><cfRule type="cellIs" dxfId="3" priority="4" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F4"><cfRule type="cellIs" dxfId="4" priority="5" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: [
          '<dxfs count="5">',
          '<dxf><fill><gradientFill degree="90"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill></dxf>',
          '<dxf><fill><patternFill patternType="gray125"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><border><diagonal style="thin"><color rgb="FF000000"/></diagonal></border></dxf>',
          '<dxf><alignment horizontal="center"/></dxf>',
          '<dxf><protection locked="0"/></dxf>',
          '</dxfs>',
        ].join(''),
      }
    }

    case 'stage1-blank-merge-order':
      return {
        dimension: 'A1:Z14',
        cells: [
          { ref: 'A1', v: 'H1' },
          { ref: 'B1', v: 'H2' },
          { ref: 'C1', v: 'H3' },
          { ref: 'D1', v: 'H4' },
          { ref: 'E1', v: 'H5' },
          { ref: 'F1', v: 'H6' },
          { ref: 'G1', v: 'H7' },
          { ref: 'H1', v: 'H8' },
          { ref: 'M1', v: 'H8' },
          { ref: 'A2', v: 'Row2' },
          // B2 is unstored blank (CF target); B10 is unstored blank control
          // Z5 is distant unstored blank (CF target); Z12 is unstored blank control
          { ref: 'A4', v: 100 },
          { ref: 'E4', v: 100 },
          { ref: 'A11', v: 100 },
          { ref: 'F11', v: 100 },
          // Hidden-row statistics block A5:A8 (row 6 hidden with outlier 1000)
          { ref: 'A5', v: 10 },
          { ref: 'A6', v: 1000 },
          { ref: 'A7', v: 20 },
          { ref: 'A8', v: 30 },
          { ref: 'A12', v: 10 },
          { ref: 'A13', v: 30 },
          // Sorted source rows block J2:J5
          { ref: 'J2', v: 90 },
          { ref: 'J3', v: 80 },
          { ref: 'J4', v: 20 },
          { ref: 'J5', v: 10 },
          { ref: 'J10', v: 90 },
          { ref: 'J13', v: 10 },
        ],
        merges: ['A4:C4', 'E4:G4'],
        hiddenCols: [3, 5],
        hiddenRows: [6],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2"><cfRule type="containsBlanks" dxfId="0" priority="1"><formula>LEN(TRIM(B2))=0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="Z5"><cfRule type="containsBlanks" dxfId="2" priority="2"><formula>LEN(TRIM(Z5))=0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="A4:C4"><cfRule type="cellIs" dxfId="1" priority="3" operator="greaterThan"><formula>50</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E4:G4"><cfRule type="cellIs" dxfId="0" priority="4" operator="greaterThan"><formula>50</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="A5:A8"><cfRule type="aboveAverage" dxfId="1" priority="5"/></conditionalFormatting>',
          '<conditionalFormatting sqref="J2:J5"><cfRule type="top10" dxfId="2" priority="6" rank="2"/></conditionalFormatting>',
          '<conditionalFormatting sqref="M1:M1048576"><cfRule type="containsBlanks" dxfId="0" priority="7"><formula>LEN(TRIM(M1))=0</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: [
          '<dxfs count="3">',
          '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><font><color rgb="FF006100"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill><border><top style="thin"><color rgb="FF006100"/></top><bottom style="thin"><color rgb="FF006100"/></bottom></border></dxf>',
          '<dxf><font><color rgb="FF9C6500"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>',
          '</dxfs>',
        ].join(''),
      }

    case 'stage1-numfmt-accessibility':
      return {
        dimension: 'A1:G10',
        colWidths: [{ min: 6, max: 6, width: 3 }],
        cells: [
          { ref: 'A1', v: 'Case' },
          { ref: 'B1', v: 'CF_0.00_Fill' },
          { ref: 'C1', v: 'Ctrl_0.0' },
          { ref: 'D1', v: 'CF_0.000_ValOnly' },
          { ref: 'E1', v: 'Ctrl_ValOnly' },
          { ref: 'F1', v: 'NarrowHash_CF' },
          { ref: 'G1', v: 'SpillNeighbor' },
          { ref: 'A2', v: 'match_1.25' },
          { ref: 'B2', v: 1.25, s: 3 },
          { ref: 'C2', v: 1.25, s: 3 },
          { ref: 'D2', v: 1.25, s: 3 },
          { ref: 'E2', v: 1.25, s: 3 },
          { ref: 'F2', v: 123456.78, s: 3 },
          { ref: 'A3', v: 'match_2.5' },
          { ref: 'B3', v: 2.5, s: 3 },
          { ref: 'C3', v: 2.5, s: 3 },
          { ref: 'D3', v: 2.5, s: 3 },
          { ref: 'E3', v: 2.5, s: 3 },
          { ref: 'F3', v: 2.5, s: 3 },
          { ref: 'A4', v: 'nomatch_0.5' },
          { ref: 'B4', v: 0.5, s: 3 },
          { ref: 'C4', v: 0.5, s: 3 },
          { ref: 'D4', v: 0.5, s: 3 },
          { ref: 'E4', v: 0.5, s: 3 },
          { ref: 'F4', v: 0.5, s: 3 },
          { ref: 'A5', v: 'spillText' },
          { ref: 'B5', v: 'LongSpillTextAcrossCells' },
          { ref: 'C5', v: null },
          { ref: 'F7', v: 123456.78, s: 3 },
          { ref: 'F8', v: 0.5, s: 3 },
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B5"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D5"><cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F5"><cfRule type="cellIs" dxfId="2" priority="3" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: [
          '<dxfs count="3">',
          '<dxf><font><b/><color rgb="FF9C0006"/></font><numFmt numFmtId="166" formatCode="0.00"/><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><numFmt numFmtId="167" formatCode="0.000"/><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><numFmt numFmtId="168" formatCode="&quot;NT$&quot;#,##0.00"/><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>',
          '</dxfs>',
        ].join(''),
      }

    case 'stage2-scale-interpolation': {
      const headers = ['Val', 'Precondition', 'RoundingLadder', 'ColorSpace', 'SpecPalette']
      const rows = [
        { label: 'r2', values: [10, 0, 0, 1] },
        { label: 'r3', values: [20, 1, 50, 2] },
        { label: 'r4', values: [30, 3, 100, 3] },
        { label: 'r5', values: [40, 5, null, 4] },
        { label: 'r6', values: [50, 255, null, 5] },
        { label: 'r7', values: [60, 509, null, 5] },
        { label: 'r8', values: [70, 510, null, 6] },
        { label: 'r9', values: [null, null, null, 7] },
        { label: 'r10', values: [null, null, null, 8] },
        { label: 'r11', values: [null, null, null, 9] },
      ]
      return {
        dimension: 'A1:E22',
        cells: replicateWithControlBlock(headers, rows, 11),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B8">',
          '<cfRule type="colorScale" priority="1">',
          '<colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale>',
          '</cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C8">',
          '<cfRule type="colorScale" priority="2">',
          '<colorScale><cfvo type="num" val="0"/><cfvo type="num" val="510"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale>',
          '</cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D4">',
          '<cfRule type="colorScale" priority="3">',
          '<colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/></colorScale>',
          '</cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E11">',
          '<cfRule type="colorScale" priority="4">',
          '<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>',
          '</cfRule>',
          '</conditionalFormatting>',
        ].join(''),
        dxfsXml: '<dxfs count="0"/>',
      }
    }

    case 'stage2-scale-thresholds': {
      const headers = ['Val', 'MinMax', 'NumBounds', 'Percent', 'PercentileIncExc', 'PercentOOB', 'LegacyMinVal0', 'MidListMin', 'NumValRef']
      const rows = [
        { label: 'r2', values: [1, 10, 10, 1, 0, 10, 10, 10] },
        { label: 'r3', values: [2, 30, 25, 2, 25, 30, 30, 30] },
        { label: 'r4', values: [3, 50, 50, 3, 50, 50, 50, 50] },
        { label: 'r5', values: [4, 70, 75, 4, 75, 70, 70, 70] },
        { label: 'r6', values: [10, 90, 100, 10, 100, 90, 90, 90] },
      ]
      return {
        dimension: 'A1:I12',
        cells: replicateWithControlBlock(headers, rows, 6),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B6"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C6"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="num" val="20"/><cfvo type="num" val="80"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D6"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="percent" val="25"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E6"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="percentile" val="30"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F6"><cfRule type="colorScale" priority="5"><colorScale><cfvo type="percent" val="-10"/><cfvo type="percentile" val="150"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G6"><cfRule type="colorScale" priority="6"><colorScale><cfvo type="min" val="0"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H6"><cfRule type="colorScale" priority="7"><colorScale><cfvo type="max"/><cfvo type="min"/><color rgb="FF63BE7B"/><color rgb="FFF8696B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I6"><cfRule type="colorScale" priority="8"><colorScale><cfvo type="num" val="$A$1"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: '<dxfs count="0"/>',
      }
    }

    case 'stage2-scale-degenerate': {
      const headers = ['Val', 'AllEqual', 'SingleCell', 'InvertedNum', 'MidAboveMax', 'OutOfBounds']
      const rows = [
        { label: 'r2', values: [5, 42, 10, 10, 5] },
        { label: 'r3', values: [5, null, 30, 30, 20] },
        { label: 'r4', values: [5, null, 50, 50, 50] },
        { label: 'r5', values: [5, null, 70, 70, 80] },
        { label: 'r6', values: [5, null, 100, 100, 95] },
      ]
      return {
        dimension: 'A1:F12',
        cells: replicateWithControlBlock(headers, rows, 6),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B6"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D6"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="num" val="100"/><cfvo type="num" val="10"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E6"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="num" val="200"/><cfvo type="num" val="100"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F6"><cfRule type="colorScale" priority="5"><colorScale><cfvo type="num" val="20"/><cfvo type="num" val="80"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: '<dxfs count="0"/>',
      }
    }

    case 'stage2-scale-participation': {
      const headers = ['Val', 'Types', 'ErrorPresence']
      const rows = [
        { label: 'r2', values: [10, 10] },
        { label: 'r3', values: [50, 20] },
        { label: 'r4', values: [90, 30] },
        { label: 'r5', values: [45000, 40], s: 3 }, // date formatted serial
        { label: 'r6', values: [true, 50] },
        { label: 'r7', values: [false, 60] },
        { label: 'r8', values: ['5', 70] },
        { label: 'r9', values: [{ err: '#N/A' }, { err: '#N/A' }] },
        { label: 'r10', values: [{ err: '#DIV/0!' }, 80] },
        { label: 'r11', values: ['', 90] },
      ]
      return {
        dimension: 'A1:C22',
        cells: replicateWithControlBlock(headers, rows, 11),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B11"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C11"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: '<dxfs count="0"/>',
      }
    }

    case 'stage2-scale-colors': {
      const headers = ['Val', 'ThemeTint', 'Indexed', 'Alpha00', 'Auto', 'CustomTheme']
      const rows = [
        { label: 'r2', values: [10, 10, 10, 10, 10] },
        { label: 'r3', values: [30, 30, 30, 30, 30] },
        { label: 'r4', values: [50, 50, 50, 50, 50] },
        { label: 'r5', values: [70, 70, 70, 70, 70] },
        { label: 'r6', values: [90, 90, 90, 90, 90] },
      ]
      return {
        dimension: 'A1:F12',
        cells: replicateWithControlBlock(headers, rows, 6),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B6"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color theme="4" tint="0.4"/><color theme="5" tint="-0.4"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C6"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color indexed="10"/><color indexed="12"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D6"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="00F8696B"/><color rgb="0063BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E6"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="max"/><color auto="1"/><color auto="1"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F6"><cfRule type="colorScale" priority="5"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF336699"/><color rgb="FF99CC33"/></colorScale></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: '<dxfs count="0"/>',
      }
    }

    case 'stage2-scale-scope': {
      const headers = ['Val', 'Disjoint', 'DupeBlock', 'HiddenRow', 'FilterHidden', 'MergeScope', 'WholeCol']
      const rows = [
        { label: 'r2', values: [10, 10, 10, 10, 10, 10] },
        { label: 'r3', values: [20, 20, 20, 20, 20, 20] },
        { label: 'r4', values: [30, 30, 30, 30, 30, 30] },
        { label: 'r5', values: [40, 40, 40, 40, 40, 40] },
        { label: 'r6', values: [50, 50, 50, 50, 50, 50] },
      ]
      const cells = replicateWithControlBlock(headers, rows, 11)
      cells.push({ ref: 'B8', v: 100 })
      cells.push({ ref: 'C8', v: 100 })
      cells.push({ ref: 'B9', v: 100 })
      cells.push({ ref: 'C9', v: 100 })
      cells.push({ ref: 'B18', v: 100 })
      cells.push({ ref: 'C18', v: 100 })
      cells.push({ ref: 'B19', v: 100 })
      cells.push({ ref: 'C19', v: 100 })
      return {
        dimension: 'A1:F20',
        cells,
        hiddenRows: [4],
        merges: ['B8:C9'],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B4 D2:D4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="B2:B6"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C6"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D6"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="B8:C9"><cfRule type="colorScale" priority="5"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F1:F1048576"><cfRule type="colorScale" priority="6"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="B2:B6"><cfRule type="colorScale" priority="7"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: '<dxfs count="0"/>',
      }
    }

    case 'stage2-scale-composition': {
      const headers = ['Val', 'DxfP1ScaleP2', 'ScaleP1DxfP2', 'EqualPrio', 'FontDxfMerge', 'MatchStopAhead', 'NonMatchStopAhead', 'ScaleStop', 'Overlapping', 'TextLowerDxf', 'OwnFillTable', 'UnresStub']
      const rows = [
        { label: 'v=25', values: [25, 25, 25, 25, 25, 25, 25, 25, 'text', 25, 25] },
        { label: 'v=15', values: [15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15] },
        { label: 'v=5', values: [5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5] },
        { label: 'v=0', values: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] },
      ]
      return {
        dimension: 'A1:L10',
        cells: replicateWithControlBlock(headers, rows, 5),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B5">',
          '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule>',
          '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C5">',
          '<cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '<cfRule type="cellIs" dxfId="1" priority="4" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D5">',
          '<cfRule type="cellIs" dxfId="0" priority="5" operator="greaterThan"><formula>10</formula></cfRule>',
          '<cfRule type="colorScale" priority="5"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E5">',
          '<cfRule type="cellIs" dxfId="2" priority="6" operator="greaterThan"><formula>0</formula></cfRule>',
          '<cfRule type="colorScale" priority="7"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F5">',
          '<cfRule type="cellIs" priority="8" stopIfTrue="1" operator="greaterThan"><formula>10</formula></cfRule>',
          '<cfRule type="colorScale" priority="9"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="G2:G5">',
          '<cfRule type="cellIs" priority="10" stopIfTrue="1" operator="greaterThan"><formula>100</formula></cfRule>',
          '<cfRule type="colorScale" priority="11"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="H2:H5">',
          '<cfRule type="colorScale" priority="12" stopIfTrue="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '<cfRule type="cellIs" dxfId="0" priority="13" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="I2:I5">',
          '<cfRule type="colorScale" priority="14"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '<cfRule type="colorScale" priority="15"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/><color rgb="FFFFFF00"/></colorScale></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="J2:J5">',
          '<cfRule type="colorScale" priority="16"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '<cfRule type="cellIs" dxfId="0" priority="17" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="K2:K5">',
          '<cfRule type="colorScale" priority="18"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="L2:L5">',
          '<cfRule type="expression" dxfId="0" priority="19" stopIfTrue="1"><formula>MOD(L2,2)=1</formula></cfRule>',
          '<cfRule type="colorScale" priority="20"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
          '</conditionalFormatting>',
        ].join(''),
        dxfsXml: [
          '<dxfs count="3">',
          '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>',
          '<dxf><font><b/><color rgb="FF006100"/></font></dxf>',
          '</dxfs>',
        ].join(''),
      }
    }

    case 'stage2-scale-unresolved-shapes': {
      const headers = ['Val', '2cfvo3colors', 'CfvoNoType', 'FormulaCfvo', 'ActivePresent', 'X14Standalone']
      const rows = [
        { label: 'r2', values: [10, 10, 10, 10, 10] },
        { label: 'r3', values: [30, 30, 30, 30, 30] },
        { label: 'r4', values: [50, 50, 50, 50, 50] },
        { label: 'r5', values: [70, 70, 70, 70, 70] },
        { label: 'r6', values: [90, 90, 90, 90, 90] },
      ]
      return {
        dimension: 'A1:F12',
        cells: replicateWithControlBlock(headers, rows, 6),
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B6"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="C2:C6"><cfRule type="colorScale" priority="2"><colorScale><cfvo/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D6"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="formula" val="SUM(A1:A5)"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="E2:E6"><cfRule type="colorScale" priority="4" activePresent="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
        ].join(''),
        x14CfXml: [
          '<x14:conditionalFormattings>',
          '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">',
          '<x14:cfRule type="colorScale" priority="5" id="{33333333-4444-5555-6666-777777777701}">',
          '<x14:colorScale><x14:cfvo type="min"/><x14:cfvo type="max"/><x14:color rgb="FFF8696B"/><x14:color rgb="FF63BE7B"/></x14:colorScale>',
          '</x14:cfRule>',
          '<xm:sqref>F2:F6</xm:sqref>',
          '</x14:conditionalFormatting>',
          '</x14:conditionalFormattings>',
        ].join(''),
        dxfsXml: '<dxfs count="0"/>',
      }
    }

    default: {
      const fx = CF_STAGE3_FIXTURES.find(f => f.id === id)
      if (fx) return stage3Payload(fx)
      throw new Error(`Unknown CF case id: ${id}`)
    }
  }
}

// ---- Stage 3: data bar fixtures (the shared `S3Fixture` spec in cases.ts drives both markers and XML) ----

const colorAttrs = (spec: string): string => {
  if (spec.startsWith('theme:')) {
    const [, theme, tint] = spec.split(':')
    return `theme="${theme}"${tint !== undefined ? ` tint="${tint}"` : ''}`
  }
  if (spec.startsWith('indexed:')) return `indexed="${spec.split(':')[1]}"`
  return `rgb="${spec}"`
}

const attrString = (attrs: Readonly<Record<string, string>> | undefined): string =>
  Object.entries(attrs ?? {}).map(([k, v]) => ` ${k}="${esc(v)}"`).join('')

const baseCfvoXml = (cfvo: readonly CfBarCfvo[]): string =>
  cfvo.map(c => `<cfvo type="${esc(c.type)}"${c.val !== undefined ? ` val="${esc(c.val)}"` : ''}/>`).join('')

const x14CfvoXml = (cfvo: readonly CfBarCfvo[]): string =>
  cfvo.map(c => (c.f !== undefined ? `<x14:cfvo type="${esc(c.type)}"><xm:f>${esc(c.f)}</xm:f></x14:cfvo>` : `<x14:cfvo type="${esc(c.type)}"/>`)).join('')

const x14ColorXml = (colors: CfBarColors, includeFill: boolean): string =>
  [
    includeFill && colors.fill ? `<x14:fillColor ${colorAttrs(colors.fill)}/>` : '',
    colors.border ? `<x14:borderColor ${colorAttrs(colors.border)}/>` : '',
    colors.negativeFill ? `<x14:negativeFillColor ${colorAttrs(colors.negativeFill)}/>` : '',
    colors.negativeBorder ? `<x14:negativeBorderColor ${colorAttrs(colors.negativeBorder)}/>` : '',
    colors.axis ? `<x14:axisColor ${colorAttrs(colors.axis)}/>` : '',
  ].join('')

const deriveX14 = (base: readonly CfBarCfvo[]): CfBarCfvo[] =>
  base.map(c => (c.val !== undefined ? { type: c.type, f: c.val } : { type: c.type }))

/** Excel's canonical rule shapes: paired = base `<dataBar>` + cfRule-level `<x14:id>` and an x14 companion without priority. */
function barRuleXml(r: S3BarRule, guid: string): { legacy?: string; x14?: string } {
  const base = r.baseCfvo ?? [{ type: 'min' }, { type: 'max' }]
  const x14Cfvo = r.x14Cfvo ?? deriveX14(base)
  const colors = r.colors ?? {}
  const stop = r.stopIfTrue ? ' stopIfTrue="1"' : ''
  if (r.pairing === 'standalone') {
    return {
      x14: `<x14:cfRule type="dataBar" priority="${r.priority}"${stop} id="${guid}"><x14:dataBar${attrString(r.x14Attrs)}>${x14CfvoXml(x14Cfvo)}${x14ColorXml(colors, true)}</x14:dataBar></x14:cfRule>`,
    }
  }
  const fill = colors.fill ? `<color ${colorAttrs(colors.fill)}/>` : ''
  const legacy = `<cfRule type="dataBar" priority="${r.priority}"${stop}><dataBar${attrString(r.baseAttrs)}>${baseCfvoXml(base)}${fill}</dataBar>${r.pairing === 'paired' ? idExtXml(guid) : ''}</cfRule>`
  if (r.pairing === 'base-only') return { legacy }
  return {
    legacy,
    x14: `<x14:cfRule type="dataBar" id="${guid}"><x14:dataBar${attrString(r.x14Attrs)}>${x14CfvoXml(x14Cfvo)}${x14ColorXml(colors, false)}</x14:dataBar></x14:cfRule>`,
  }
}

const idExtXml = (guid: string): string =>
  `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${NS_X14}"><x14:id>${guid}</x14:id></ext></extLst>`

function stage3Payload(fx: S3Fixture): SheetPayload {
  const off = s3ControlOffset(fx)
  const cells: CellSpec[] = []
  const put = (ref: string, v: S3Value, ctrlRef: string, labelSuffix = ''): void => {
    void labelSuffix
    for (const target of [ref, ctrlRef]) {
      if (v === null) continue
      if (typeof v === 'object') {
        if ('err' in v) cells.push({ ref: target, err: v.err })
        else if ('f' in v) cells.push({ ref: target, f: v.f, v: v.v })
        else cells.push({ ref: target, v: v.date, s: 4 })
      } else {
        cells.push({ ref: target, v })
      }
    }
  }
  cells.push({ ref: 'A1', v: 'Val' }, { ref: `A${1 + off}`, v: 'Val_Ctrl' })
  for (const c of fx.cols) {
    cells.push({ ref: `${c.col}1`, v: c.header }, { ref: `${c.col}${1 + off}`, v: `${c.header}_Ctrl` })
  }
  for (let r = 2; r <= fx.n + 1; r++) {
    cells.push({ ref: `A${r}`, v: `r${r}` }, { ref: `A${r + off}`, v: `r${r}_ctrl` })
  }
  for (const c of fx.cols) {
    c.values.forEach((v, i) => put(`${c.col}${i + 2}`, v, s3ControlRef(fx, `${c.col}${i + 2}`)))
  }
  for (const e of fx.extraCells ?? []) put(e.ref, e.v, s3ControlRef(fx, e.ref))

  const heights = new Map<number, number>()
  for (const [row, h] of Object.entries(fx.rowHeights ?? {})) {
    heights.set(Number(row), h)
    heights.set(Number(row) + off, h)
  }

  // Group legacy rules into <conditionalFormatting> blocks (consecutive rules sharing a sqref, unless separate).
  const legacyBlocks: Array<{ sqref: string; rules: string[] }> = []
  const x14Blocks: string[] = []
  fx.rules.forEach((r, i) => {
    const guid = `{3B3B00${(CF_STAGE3_FIXTURES.indexOf(fx) + 1).toString(16).toUpperCase().padStart(2, '0')}-0000-4000-8000-${(i + 1).toString(16).toUpperCase().padStart(12, '0')}}`
    const out = r.kind === 'bar' ? barRuleXml(r, guid) : r.where === 'legacy' ? { legacy: r.xml } : { x14: r.xml }
    if (out.legacy) {
      const last = legacyBlocks[legacyBlocks.length - 1]
      if (last && last.sqref === r.sqref && !r.separateBlock) last.rules.push(out.legacy)
      else legacyBlocks.push({ sqref: r.sqref, rules: [out.legacy] })
    }
    if (out.x14) {
      x14Blocks.push(`<x14:conditionalFormatting xmlns:xm="${NS_XM}">${out.x14}<xm:sqref>${esc(r.sqref)}</xm:sqref></x14:conditionalFormatting>`)
    }
  })

  return {
    dimension: s3Range(fx),
    cells,
    ...(fx.merges ? { merges: fx.merges } : {}),
    ...(fx.hiddenRows ? { hiddenRows: fx.hiddenRows } : {}),
    colWidths: fx.cols.map(c => ({ min: colIdxOf(c.col), max: colIdxOf(c.col), width: c.width ?? 30 })),
    legacyCfXml: legacyBlocks.map(b => `<conditionalFormatting sqref="${esc(b.sqref)}">${b.rules.join('')}</conditionalFormatting>`).join(''),
    ...(x14Blocks.length > 0 ? { x14CfXml: `<x14:conditionalFormattings>${x14Blocks.join('')}</x14:conditionalFormattings>` } : {}),
    dxfsXml: fx.dxfsXml ?? '<dxfs count="0"/>',
    rowHeightsPt: heights,
    defaultRowHeightPt: 30,
    showGridLines: fx.gridlines === true,
    ...(fx.autoFilterXml ? { autoFilterXml: fx.autoFilterXml } : {}),
    ...(fx.dateStyle ? { dateStyle: true } : {}),
  }
}

const colIdxOf = (letters: string): number => {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}


function buildWorksheetXml(p: SheetPayload): string {
  const byRow = new Map<number, CellSpec[]>()
  for (const c of p.cells) {
    const m = /^([A-Z]+)(\d+)$/.exec(c.ref)
    if (!m) throw new Error(`Invalid cell ref ${c.ref}`)
    const r = Number(m[2])
    const list = byRow.get(r) ?? []
    list.push(c)
    byRow.set(r, list)
  }
  const colIdx = (ref: string): number => {
    const letters = /^([A-Z]+)/.exec(ref)?.[1] ?? 'A'
    let n = 0
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64)
    return n
  }
  const hiddenRowSet = new Set(p.hiddenRows ?? [])
  const maxRowFromDim = Number(/:([A-Z]+)(\d+)$/.exec(p.dimension)?.[2] ?? 10)
  for (let r = 1; r <= maxRowFromDim; r++) {
    if (!byRow.has(r)) byRow.set(r, [])
  }
  const sortedRows = [...byRow.keys()].sort((a, b) => a - b)
  const rowsXml = sortedRows
    .map(r => {
      const cells = (byRow.get(r) ?? []).sort((a, b) => colIdx(a.ref) - colIdx(b.ref))
      const cellsXml = cells
        .map(c => {
          const sAttr = c.s !== undefined ? ` s="${c.s}"` : ''
          if (c.err !== undefined) {
            return `<c r="${c.ref}"${sAttr} t="e"><v>${esc(c.err)}</v></c>`
          }
          if (c.v === null) return ''
          if (c.f !== undefined) {
            return `<c r="${c.ref}"${sAttr}><f>${esc(c.f)}</f><v>${c.v}</v></c>`
          }
          if (typeof c.v === 'number') {
            return `<c r="${c.ref}"${sAttr}><v>${c.v}</v></c>`
          }
          if (typeof c.v === 'boolean') {
            return `<c r="${c.ref}"${sAttr} t="b"><v>${c.v ? 1 : 0}</v></c>`
          }
          return `<c r="${c.ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${esc(c.v)}</t></is></c>`
        })
        .filter(Boolean)
        .join('')
      const hiddenAttr = hiddenRowSet.has(r) ? ' hidden="1"' : ''
      const rowHt = p.rowHeightsPt?.get(r) ?? p.defaultRowHeightPt ?? 18
      return `<row r="${r}" ht="${rowHt}" customHeight="1"${hiddenAttr}>${cellsXml}</row>`
    })
    .join('')

  const hiddenColSet = new Set(p.hiddenCols ?? [])
  const customWidthMap = new Map<number, number>()
  for (const cw of p.colWidths ?? []) {
    for (let c = cw.min; c <= cw.max; c++) customWidthMap.set(c, cw.width)
  }
  const maxColLetters = /:([A-Z]+)\d+$/.exec(p.dimension)?.[1] ?? 'Z'
  const maxCol1 = colIdx(`${maxColLetters}1`)
  const colEntries: string[] = []
  for (let c = 1; c <= maxCol1; c++) {
    const w = customWidthMap.get(c) ?? 14
    const hiddenAttr = hiddenColSet.has(c) ? ' hidden="1"' : ''
    colEntries.push(`<col min="${c}" max="${c}" width="${w}" customWidth="1"${hiddenAttr}/>`)
  }
  const colsXml = `<cols>${colEntries.join('')}</cols>`

  const mergesXml =
    p.merges && p.merges.length > 0
      ? `<mergeCells count="${p.merges.length}">${p.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`
      : ''

  const tablePartsXml = p.tablePartXml ? '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>' : ''

  const extLstXml = p.x14CfXml
    ? `<extLst><ext uri="${X14_CF_EXT_URI}" xmlns:x14="${NS_X14}">${p.x14CfXml}</ext></extLst>`
    : ''

  return [
    XML_DECL,
    `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" mc:Ignorable="x14ac" xmlns:x14ac="${NS_X14AC}" xmlns:x14="${NS_X14}" xmlns:xm="${NS_XM}">`,
    `<dimension ref="${p.dimension}"/>`,
    `<sheetViews><sheetView workbookViewId="0" showGridLines="${p.showGridLines === false ? 0 : 1}"/></sheetViews>`,
    `<sheetFormatPr defaultRowHeight="${p.defaultRowHeightPt ?? 18}" defaultColWidth="14"/>`,
    colsXml,
    `<sheetData>${rowsXml}</sheetData>`,
    p.autoFilterXml ?? '',
    mergesXml,
    p.legacyCfXml,
    tablePartsXml,
    extLstXml,
    '</worksheet>',
  ].join('')
}

function buildStylesXml(dxfsXml: string, dateStyle = false): string {
  return [
    XML_DECL,
    `<styleSheet xmlns="${NS_MAIN}">`,
    '<numFmts count="2">',
    '<numFmt numFmtId="164" formatCode="00000"/>',
    '<numFmt numFmtId="165" formatCode="0.0"/>',
    '</numFmts>',
    '<fonts count="2">',
    '<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>',
    '<font><b/><u/><strike/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>',
    '</fonts>',
    '<fills count="3">',
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>',
    '</fills>',
    '<borders count="1">',
    '<border><left/><right/><top/><bottom/><diagonal/></border>',
    '</borders>',
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    `<cellXfs count="${dateStyle ? 5 : 4}">`,
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    ...(dateStyle ? ['<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'] : []),
    '</cellXfs>',
    dxfsXml,
    '</styleSheet>',
  ].join('')
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let i = 0; i < 256; i++) {
    let c = i
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[i] = c >>> 0
  }
  return table
})()

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff
  for (let i = 0; i < bytes.byteLength; i++) {
    c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8)
  }
  return (c ^ 0xffffffff) >>> 0
}

/** Deterministic synchronous ZIP writer (fixed 2026-01-01 DOS timestamp, sorted entry order). */
function writeDeterministicZipSync(files: Record<string, string>): Uint8Array {
  const encoder = new TextDecoder()
  void encoder
  const textEncoder = new TextEncoder()
  const entries = Object.keys(files)
    .sort()
    .map(name => {
      const nameBytes = textEncoder.encode(name)
      const rawBytes = textEncoder.encode(files[name])
      const compBytes = new Uint8Array(deflateRawSync(rawBytes, { level: 6 }))
      const crc = crc32(rawBytes)
      return { nameBytes, rawBytes, compBytes, crc }
    })

  const dosTime = 0
  const dosDate = ((2026 - 1980) << 9) | (1 << 5) | 1

  let localTotal = 0
  let cdTotal = 0
  for (const e of entries) {
    localTotal += 30 + e.nameBytes.byteLength + e.compBytes.byteLength
    cdTotal += 46 + e.nameBytes.byteLength
  }

  const out = new Uint8Array(localTotal + cdTotal + 22)
  const view = new DataView(out.buffer)
  let localOffset = 0
  const offsets: number[] = []

  for (const e of entries) {
    offsets.push(localOffset)
    view.setUint32(localOffset, 0x04034b50, true)
    view.setUint16(localOffset + 4, 20, true)
    view.setUint16(localOffset + 6, 0, true)
    view.setUint16(localOffset + 8, 8, true)
    view.setUint16(localOffset + 10, dosTime, true)
    view.setUint16(localOffset + 12, dosDate, true)
    view.setUint32(localOffset + 14, e.crc, true)
    view.setUint32(localOffset + 18, e.compBytes.byteLength, true)
    view.setUint32(localOffset + 22, e.rawBytes.byteLength, true)
    view.setUint16(localOffset + 26, e.nameBytes.byteLength, true)
    view.setUint16(localOffset + 28, 0, true)
    out.set(e.nameBytes, localOffset + 30)
    out.set(e.compBytes, localOffset + 30 + e.nameBytes.byteLength)
    localOffset += 30 + e.nameBytes.byteLength + e.compBytes.byteLength
  }

  let cdOffset = localOffset
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i]
    view.setUint32(cdOffset, 0x02014b50, true)
    view.setUint16(cdOffset + 4, 20, true)
    view.setUint16(cdOffset + 6, 20, true)
    view.setUint16(cdOffset + 8, 0, true)
    view.setUint16(cdOffset + 10, 8, true)
    view.setUint16(cdOffset + 12, dosTime, true)
    view.setUint16(cdOffset + 14, dosDate, true)
    view.setUint32(cdOffset + 16, e.crc, true)
    view.setUint32(cdOffset + 20, e.compBytes.byteLength, true)
    view.setUint32(cdOffset + 24, e.rawBytes.byteLength, true)
    view.setUint16(cdOffset + 28, e.nameBytes.byteLength, true)
    view.setUint16(cdOffset + 30, 0, true)
    view.setUint16(cdOffset + 32, 0, true)
    view.setUint16(cdOffset + 34, 0, true)
    view.setUint16(cdOffset + 36, 0, true)
    view.setUint32(cdOffset + 38, 0, true)
    view.setUint32(cdOffset + 42, offsets[i], true)
    out.set(e.nameBytes, cdOffset + 46)
    cdOffset += 46 + e.nameBytes.byteLength
  }

  view.setUint32(cdOffset, 0x06054b50, true)
  view.setUint16(cdOffset + 4, 0, true)
  view.setUint16(cdOffset + 6, 0, true)
  view.setUint16(cdOffset + 8, entries.length, true)
  view.setUint16(cdOffset + 10, entries.length, true)
  view.setUint32(cdOffset + 12, cdTotal, true)
  view.setUint32(cdOffset + 16, localTotal, true)
  view.setUint16(cdOffset + 20, 0, true)
  return out
}

export async function generateCaseArchive(id: string, anchorDateIso = DEFAULT_CLOCK_ANCHOR_DATE): Promise<Uint8Array> {
  const p = casePayload(id, anchorDateIso)
  const wbPrXml = p.date1904 ? '<workbookPr date1904="1"/>' : '<workbookPr date1904="0"/>'
  const files: Record<string, string> = {
    '[Content_Types].xml': [
      XML_DECL,
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
      ...(p.pivotParts
        ? [
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>',
            '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
            '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>',
          ]
        : []),
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
      ...(p.tablePartXml
        ? ['<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>']
        : []),
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
      '</Types>',
    ].join(''),
    '_rels/.rels': [
      XML_DECL,
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/>`,
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
      `<Relationship Id="rId3" Type="${NS_R}/extended-properties" Target="docProps/app.xml"/>`,
      '</Relationships>',
    ].join(''),
    'docProps/app.xml': [
      XML_DECL,
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">',
      '<Application>Microsoft Macintosh Excel</Application>',
      '<AppVersion>16.0300</AppVersion>',
      '</Properties>',
    ].join(''),
    'docProps/core.xml': [
      XML_DECL,
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
      `<dc:title>${esc(id)}</dc:title>`,
      '<dc:creator>OfficeView CF Oracle Generator</dc:creator>',
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-11-06T00:00:00Z</dcterms:created>',
      '</cp:coreProperties>',
    ].join(''),
    'xl/workbook.xml': [
      XML_DECL,
      `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`,
      wbPrXml,
      p.pivotParts
        ? '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="PivotSrc" sheetId="2" state="hidden" r:id="rId3"/></sheets><pivotCaches><pivotCache cacheId="1" r:id="rId4"/></pivotCaches>'
        : '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>',
      '</workbook>',
    ].join(''),
    'xl/_rels/workbook.xml.rels': [
      XML_DECL,
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      `<Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/>`,
      `<Relationship Id="rId2" Type="${NS_R}/styles" Target="styles.xml"/>`,
      ...(p.pivotParts
        ? [
            `<Relationship Id="rId3" Type="${NS_R}/worksheet" Target="worksheets/sheet2.xml"/>`,
            `<Relationship Id="rId4" Type="${NS_R}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>`,
          ]
        : []),
      '</Relationships>',
    ].join(''),
    'xl/styles.xml': buildStylesXml(p.dxfsXml, p.dateStyle === true),
    'xl/worksheets/sheet1.xml': buildWorksheetXml(p),
  }
  if (p.tablePartXml || p.pivotParts) {
    if (p.tablePartXml) {
      files['xl/tables/table1.xml'] = p.tablePartXml
    }
    if (p.pivotParts) {
      files['xl/worksheets/sheet2.xml'] = p.pivotParts.pivotSrcSheetXml
      files['xl/pivotTables/pivotTable1.xml'] = p.pivotParts.pivotTableXml
      files['xl/pivotTables/_rels/pivotTable1.xml.rels'] = [
        XML_DECL,
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        `<Relationship Id="rIdCacheDef1" Type="${NS_R}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>`,
        '</Relationships>',
      ].join('')
      files['xl/pivotCache/pivotCacheDefinition1.xml'] = p.pivotParts.pivotCacheDefXml
      files['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = [
        XML_DECL,
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        `<Relationship Id="rIdCacheRec1" Type="${NS_R}/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>`,
        '</Relationships>',
      ].join('')
      files['xl/pivotCache/pivotCacheRecords1.xml'] = p.pivotParts.pivotCacheRecordsXml
    }
    files['xl/worksheets/_rels/sheet1.xml.rels'] = [
      XML_DECL,
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      ...(p.tablePartXml ? [`<Relationship Id="rIdTable1" Type="${NS_R}/table" Target="../tables/table1.xml"/>`] : []),
      ...(p.pivotParts ? [`<Relationship Id="rIdPivot1" Type="${NS_R}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>`] : []),
      '</Relationships>',
    ].join('')
  }
  return writeDeterministicZipSync(files)
}

export interface GenerateStage1Options {
  rootDir?: string
  /** When set, only this case's files and manifest entry are generated/updated; other cases remain untouched. */
  onlyCaseId?: string
  /** When set, only cases belonging to this stage are generated/updated; other stages' cases remain untouched. */
  stage?: number
  /** Anchor local date (`YYYY-MM-DD` in `Asia/Taipei`) for clock-dependent fixtures. */
  anchorLocalDate?: string
  /** Regenerate `measured` cases too, dropping their captures and observed expectations (`--force`). */
  force?: boolean
}

export async function generateAllStage1(options: GenerateStage1Options | string = {}): Promise<CfManifest> {
  const opts: GenerateStage1Options = typeof options === 'string' ? { rootDir: options } : options
  const rootDir = opts.rootDir ?? REPO_ROOT
  const onlyCaseId = opts.onlyCaseId
  const targetStage = opts.stage
  const anchorDateIso = opts.anchorLocalDate ?? DEFAULT_CLOCK_ANCHOR_DATE

  const fixtureDir = resolve(rootDir, 'tests/fixtures/xlsx-cf')
  const inputDir = join(fixtureDir, 'inputs')
  const refsDir = join(fixtureDir, 'refs')
  const manifestPath = join(fixtureDir, 'manifest.json')
  mkdirSync(inputDir, { recursive: true })
  mkdirSync(refsDir, { recursive: true })

  let existingManifest: CfManifest | undefined
  if (existsSync(manifestPath)) {
    try {
      existingManifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as CfManifest
    } catch {
      existingManifest = undefined
    }
  }

  const cases: CfOracleCase[] = []
  for (const template of CF_CASES) {
    const prev = existingManifest?.cases.find(x => x.id === template.id)
    if (onlyCaseId && template.id !== onlyCaseId && prev) {
      // Do not touch other cases when rebuilding a single selected case
      cases.push(prev)
      continue
    }
    if (targetStage !== undefined && template.stage !== targetStage && prev) {
      // Do not touch other stages when targeting a specific stage
      cases.push(prev)
      continue
    }

    if (prev?.status === 'measured' && !opts.force) {
      // A measured case holds committed Excel evidence (captures, observed expectations, the resave receipt). It is
      // never reset by regeneration, even when it no longer verifies: only an explicit --force clobbers it.
      const regenerated = sha256Hex(await generateCaseArchive(template.id, prev.clockRequirement?.anchorLocalDate ?? anchorDateIso))
      if (regenerated !== prev.expected.inputSha256) {
        console.warn(`[cf-refs/generate] ${template.id} is measured: kept as is although its template now generates a different input (pass --force to regenerate and drop its evidence)`)
      }
      cases.push(prev)
      continue
    }

    const effectiveAnchorDate =
      template.clockRequirement && opts.anchorLocalDate
        ? anchorDateIso
        : prev?.clockRequirement?.anchorLocalDate ?? template.clockRequirement?.anchorLocalDate ?? anchorDateIso

    const zipBytes = await generateCaseArchive(template.id, effectiveAnchorDate)
    const inputAbs = resolve(rootDir, template.input)
    const savedAbs = resolve(rootDir, template.saved)
    writeFileSync(inputAbs, zipBytes)
    const inputSha256 = sha256Hex(zipBytes)

    const prevVerify = prev ? verifyCfCase(prev, { rootDir }) : null
    const sameInputAsPrev = prev?.expected.inputSha256 === inputSha256
    const wasResavedByExcel =
      sameInputAsPrev &&
      (prev?.resavedByExcel === true ||
        (prev !== undefined &&
          existsSync(savedAbs) &&
          Boolean(prev.expected.savedSha256) &&
          Boolean(prev.expected.inputSha256) &&
          prev.expected.savedSha256 !== prev.expected.inputSha256))
    const preserveExistingSaved =
      existsSync(savedAbs) && ((prevVerify?.kind === 'verified' && sameInputAsPrev) || wasResavedByExcel)

    if (!preserveExistingSaved) {
      writeFileSync(savedAbs, zipBytes)
    }
    const savedBytes = new Uint8Array(readFileSync(savedAbs))
    const savedSha256 = sha256Hex(savedBytes)

    const updatedClockReq = template.clockRequirement
      ? {
          ...template.clockRequirement,
          anchorLocalDate: effectiveAnchorDate,
          anchorSerial: isoDateToExcelSerial(effectiveAnchorDate, template.clockRequirement.dateSystem),
          staleSavedTodaySerial: isoDateToExcelSerial(effectiveAnchorDate, template.clockRequirement.dateSystem) - 5,
        }
      : undefined

    cases.push({
      ...template,
      ...(updatedClockReq ? { clockRequirement: updatedClockReq } : {}),
      resavedByExcel: preserveExistingSaved ? (prev?.resavedByExcel ?? wasResavedByExcel) : false,
      ...(preserveExistingSaved && prev?.resaveReceipt ? { resaveReceipt: prev.resaveReceipt } : {}),
      status: preserveExistingSaved && prev ? prev.status : template.status,
      reason: preserveExistingSaved && prev ? prev.reason : template.reason,
      captures: preserveExistingSaved && prev ? prev.captures : template.captures,
      expected: {
        ...(preserveExistingSaved && prev && prevVerify?.kind === 'verified' ? prev.expected : template.expected),
        inputSha256,
        savedSha256,
      },
    })
  }

  const manifest: CfManifest = {
    version: 1,
    excelBuild: '16.114',
    locale: 'zh-TW',
    dateSystem: '1900',
    hostTimeZone: 'Asia/Taipei',
    generatedAt: '2026-11-06T00:00:00Z',
    cases,
  }
  writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + '\n')
  return manifest
}

export const generateAllCases = generateAllStage1

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2)
  const onlyCaseId = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  const stage = args.includes('--stage') ? Number(args[args.indexOf('--stage') + 1]) : undefined
  const anchorLocalDate = args.includes('--anchor-date') ? args[args.indexOf('--anchor-date') + 1] : undefined
  const force = args.includes('--force')
  const manifest = await generateAllStage1({ onlyCaseId, stage, anchorLocalDate, force })
  console.log(`[cf-refs/generate] generated ${manifest.cases.length} oracle cases -> ${CF_MANIFEST_PATH}`)
}
