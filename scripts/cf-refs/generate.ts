/**
 * Deterministic raw SpreadsheetML generator for XLSX Conditional Formatting Stage 1 oracle workbooks.
 *
 *   bun scripts/cf-refs/generate.ts
 *
 * Writes `tests/fixtures/xlsx-cf/inputs/<id>.xlsx`, seeds `tests/fixtures/xlsx-cf/<id>.xlsx` when no verified
 * Excel-resaved output exists yet, and updates `tests/fixtures/xlsx-cf/manifest.json` with SHA-256 digests.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { CF_STAGE1_CASES, type CfManifest, type CfOracleCase } from './cases'
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
const X14_CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AA7539E650}'
const ZIP_DATE = new Date(Date.UTC(2026, 0, 1))

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

const sha256Hex = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')

type CellSpec =
  | { ref: string; v: number | string | boolean | null; s?: number; f?: string; err?: undefined }
  | { ref: string; err: string; s?: number; v?: undefined; f?: undefined }

interface SheetPayload {
  dimension: string
  cells: readonly CellSpec[]
  merges?: readonly string[]
  hiddenCols?: readonly number[]
  legacyCfXml: string
  x14CfXml?: string
  dxfsXml: string
}

const DEFAULT_DXFS_XML = [
  '<dxfs count="3">',
  '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
  '<dxf><font><color rgb="FF006100"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>',
  '<dxf><font><color rgb="FF9C6500"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>',
  '</dxfs>',
].join('')

function casePayload(id: string): SheetPayload {
  switch (id) {
    case 'stage1-cellis-literals':
      return {
        dimension: 'A1:G10',
        cells: [
          { ref: 'A1', v: 'Input' },
          { ref: 'B1', v: 'GT100_CF' },
          { ref: 'C1', v: 'GT100_Ctrl' },
          { ref: 'D1', v: 'RevBtw_CF' },
          { ref: 'E1', v: 'RevBtw_Ctrl' },
          { ref: 'F1', v: 'StrEq_CF' },
          { ref: 'G1', v: 'StrEq_Ctrl' },
          ...([-1, 0, 1, 99, 100, 101] as const).flatMap((val, idx): CellSpec[] => {
            const r = idx + 2
            return [
              { ref: `A${r}`, v: `v=${val}` },
              { ref: `B${r}`, v: val },
              { ref: `C${r}`, v: val },
              { ref: `D${r}`, v: val },
              { ref: `E${r}`, v: val },
              { ref: `F${r}`, v: idx === 4 ? '100' : val },
              { ref: `G${r}`, v: idx === 4 ? '100' : val },
            ]
          }),
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B7"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D7"><cfRule type="cellIs" dxfId="1" priority="2" operator="between"><formula>3</formula><formula>1</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F7"><cfRule type="cellIs" dxfId="2" priority="3" operator="equal"><formula>"100"</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }

    case 'stage1-text-predicates':
      return {
        dimension: 'A1:G9',
        cells: [
          { ref: 'A1', v: 'Case' },
          { ref: 'B1', v: 'Contains_CF' },
          { ref: 'C1', v: 'Contains_Ctrl' },
          { ref: 'D1', v: 'Begins_CF' },
          { ref: 'E1', v: 'Begins_Ctrl' },
          { ref: 'F1', v: 'X14Div_CF' },
          { ref: 'G1', v: 'X14Div_Ctrl' },
          ...(
            [
              ['lower', 'abc'],
              ['mixed', 'AbC'],
              ['accent', 'ábc'],
              ['padded', ' ABC '],
              ['numFmt', 123],
              ['emptyStr', ''],
              ['blank', null],
            ] as const
          ).flatMap(([label, val], idx): CellSpec[] => {
            const r = idx + 2
            const s = label === 'numFmt' ? 2 : undefined
            return [
              { ref: `A${r}`, v: label },
              { ref: `B${r}`, v: val, s },
              { ref: `C${r}`, v: val, s },
              { ref: `D${r}`, v: val, s },
              { ref: `E${r}`, v: val, s },
              { ref: `F${r}`, v: val, s },
              { ref: `G${r}`, v: val, s },
            ]
          }),
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B8"><cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH("ab",B2)))</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D8"><cfRule type="beginsWith" dxfId="1" priority="2" operator="beginsWith" text="00"><formula>LEFT(D2,2)="00"</formula></cfRule></conditionalFormatting>',
        ].join(''),
        x14CfXml: `<x14:conditionalFormatting xmlns:xm="${NS_XM}"><x14:cfRule type="containsText" operator="containsText" priority="3" id="{11111111-2222-3333-4444-555555555501}"><xm:f>NOT(ISERROR(SEARCH("zz",F2)))</xm:f><xm:f>"ab"</xm:f><x14:dxf><font><color rgb="FF9C6500"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>F2:F8</xm:sqref></x14:conditionalFormatting>`,
        dxfsXml: DEFAULT_DXFS_XML,
      }

    case 'stage1-blanks-errors':
      return {
        dimension: 'A1:G10',
        cells: [
          { ref: 'A1', v: 'Sample' },
          { ref: 'B1', v: 'Blanks_CF' },
          { ref: 'C1', v: 'Blanks_Ctrl' },
          { ref: 'D1', v: 'Errors_CF' },
          { ref: 'E1', v: 'Errors_Ctrl' },
          { ref: 'F1', v: 'NoErr_CF' },
          { ref: 'G1', v: 'NoErr_Ctrl' },
          { ref: 'A2', v: 'unstored' },
          { ref: 'A3', v: 'emptyStr' },
          { ref: 'B3', v: '' },
          { ref: 'C3', v: '' },
          { ref: 'D3', v: '' },
          { ref: 'E3', v: '' },
          { ref: 'F3', v: '' },
          { ref: 'G3', v: '' },
          { ref: 'A4', v: 'space' },
          { ref: 'B4', v: ' ' },
          { ref: 'C4', v: ' ' },
          { ref: 'D4', v: ' ' },
          { ref: 'E4', v: ' ' },
          { ref: 'F4', v: ' ' },
          { ref: 'G4', v: ' ' },
          { ref: 'A5', v: 'tab' },
          { ref: 'B5', v: '\t' },
          { ref: 'C5', v: '\t' },
          { ref: 'D5', v: '\t' },
          { ref: 'E5', v: '\t' },
          { ref: 'F5', v: '\t' },
          { ref: 'G5', v: '\t' },
          { ref: 'A6', v: 'zero' },
          { ref: 'B6', v: 0 },
          { ref: 'C6', v: 0 },
          { ref: 'D6', v: 0 },
          { ref: 'E6', v: 0 },
          { ref: 'F6', v: 0 },
          { ref: 'G6', v: 0 },
          { ref: 'A7', v: 'false' },
          { ref: 'B7', v: false },
          { ref: 'C7', v: false },
          { ref: 'D7', v: false },
          { ref: 'E7', v: false },
          { ref: 'F7', v: false },
          { ref: 'G7', v: false },
          { ref: 'A8', v: 'na' },
          { ref: 'B8', err: '#N/A' },
          { ref: 'C8', err: '#N/A' },
          { ref: 'D8', err: '#N/A' },
          { ref: 'E8', err: '#N/A' },
          { ref: 'F8', err: '#N/A' },
          { ref: 'G8', err: '#N/A' },
          { ref: 'A9', v: 'div0' },
          { ref: 'B9', err: '#DIV/0!' },
          { ref: 'C9', err: '#DIV/0!' },
          { ref: 'D9', err: '#DIV/0!' },
          { ref: 'E9', err: '#DIV/0!' },
          { ref: 'F9', err: '#DIV/0!' },
          { ref: 'G9', err: '#DIV/0!' },
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B9"><cfRule type="containsBlanks" dxfId="0" priority="1"><formula>LEN(TRIM(B2))=0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D9"><cfRule type="containsErrors" dxfId="1" priority="2"><formula>ISERROR(D2)</formula></cfRule></conditionalFormatting>',
        ].join(''),
        x14CfXml: `<x14:conditionalFormatting xmlns:xm="${NS_XM}"><x14:cfRule type="notContainsErrors" priority="3" id="{11111111-2222-3333-4444-555555555502}"><xm:f>NOT(ISERROR(F2))</xm:f><x14:dxf><font><color rgb="FF9C6500"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>F2:F9</xm:sqref></x14:conditionalFormatting>`,
        dxfsXml: DEFAULT_DXFS_XML,
      }

    case 'stage1-timeperiod-today-clock': {
      // Base serial around 2026-11-06 (1900 date system serial 46332)
      const d = 46332
      const offsets = [
        ['D-2', d - 2],
        ['D-1', d - 1],
        ['D', d],
        ['D+1', d + 1],
        ['D+2', d + 2],
        ['D+0.5', d + 0.5],
        ['D+0.999', d + 0.999],
        ['D-7', d - 7],
        ['D+7', d + 7],
        ['StaleTODAY', d - 5],
      ] as const
      return {
        dimension: 'A1:E12',
        cells: [
          { ref: 'A1', v: 'Offset' },
          { ref: 'B1', v: 'Today_CF' },
          { ref: 'C1', v: 'Today_Ctrl' },
          { ref: 'D1', v: 'ThisWeek_CF' },
          { ref: 'E1', v: 'ThisWeek_Ctrl' },
          ...offsets.flatMap(([label, val], idx): CellSpec[] => {
            const r = idx + 2
            const f = label === 'StaleTODAY' ? 'TODAY()' : undefined
            return [
              { ref: `A${r}`, v: label },
              { ref: `B${r}`, v: val, f },
              { ref: `C${r}`, v: val, f },
              { ref: `D${r}`, v: val, f },
              { ref: `E${r}`, v: val, f },
            ]
          }),
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B11"><cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="today"><formula>FLOOR(B2,1)=TODAY()</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D11"><cfRule type="timePeriod" dxfId="1" priority="2" timePeriod="thisWeek"><formula>AND(TODAY()-ROUNDDOWN(D2,0)&lt;=(WEEKDAY(TODAY())-1),ROUNDDOWN(D2,0)-TODAY()&lt;=(7-WEEKDAY(TODAY())))</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }
    }

    case 'stage1-top10-rank-edge':
      return {
        dimension: 'A1:G9',
        cells: [
          { ref: 'A1', v: 'Row' },
          { ref: 'B1', v: 'Top2_CF' },
          { ref: 'C1', v: 'Top2_Ctrl' },
          { ref: 'D1', v: 'Bot2_CF' },
          { ref: 'E1', v: 'Bot2_Ctrl' },
          { ref: 'F1', v: 'Top40Pct_CF' },
          { ref: 'G1', v: 'Top40Pct_Ctrl' },
          ...(
            [
              ['n1', 1],
              ['n2a', 2],
              ['n2b', 2],
              ['n3', 3],
              ['n9', 9],
              ['str', '999'],
              ['bool', true],
            ] as const
          ).flatMap(([label, val], idx): CellSpec[] => {
            const r = idx + 2
            return [
              { ref: `A${r}`, v: label },
              { ref: `B${r}`, v: val },
              { ref: `C${r}`, v: val },
              { ref: `D${r}`, v: val },
              { ref: `E${r}`, v: val },
              { ref: `F${r}`, v: val },
              { ref: `G${r}`, v: val },
            ]
          }),
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B8"><cfRule type="top10" dxfId="0" priority="1" rank="2"/></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D8"><cfRule type="top10" dxfId="1" priority="2" bottom="1" rank="2"/></conditionalFormatting>',
          '<conditionalFormatting sqref="F2:F8"><cfRule type="top10" dxfId="2" priority="3" percent="1" rank="40"/></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }

    case 'stage1-above-average-stddev':
      return {
        dimension: 'A1:G8',
        cells: [
          { ref: 'A1', v: 'Row' },
          { ref: 'B1', v: 'StdDev1_CF' },
          { ref: 'C1', v: 'StdDev1_Ctrl' },
          { ref: 'D1', v: 'BelowEq_CF' },
          { ref: 'E1', v: 'BelowEq_Ctrl' },
          // Dataset [0, 0, 10, 11] in B2:B5 distinguishes population sigma (threshold 10.5119) vs sample s (threshold 11.3259)
          { ref: 'A2', v: 'p1' },
          { ref: 'B2', v: 0 },
          { ref: 'C2', v: 0 },
          { ref: 'D2', v: 0 },
          { ref: 'E2', v: 0 },
          { ref: 'A3', v: 'p2' },
          { ref: 'B3', v: 0 },
          { ref: 'C3', v: 0 },
          { ref: 'D3', v: 2 },
          { ref: 'E3', v: 2 },
          { ref: 'A4', v: 'p3' },
          { ref: 'B4', v: 10 },
          { ref: 'C4', v: 10 },
          { ref: 'D4', v: 4 },
          { ref: 'E4', v: 4 },
          { ref: 'A5', v: 'p4_probe11' },
          { ref: 'B5', v: 11 },
          { ref: 'C5', v: 11 },
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B5"><cfRule type="aboveAverage" dxfId="0" priority="1" aboveAverage="1" stdDev="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D4"><cfRule type="aboveAverage" dxfId="1" priority="2" aboveAverage="0" equalAverage="1"/></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }

    case 'stage1-duplicate-scope':
      return {
        dimension: 'A1:G10',
        cells: [
          { ref: 'A1', v: 'Kind' },
          { ref: 'B1', v: 'Dupe_CF' },
          { ref: 'C1', v: 'Dupe_Ctrl' },
          { ref: 'D1', v: 'Disj1_CF' },
          { ref: 'E1', v: 'Disj1_Ctrl' },
          { ref: 'F1', v: 'Disj2_CF' },
          { ref: 'G1', v: 'Disj2_Ctrl' },
          { ref: 'A2', v: 'num1_a' },
          { ref: 'B2', v: 1 },
          { ref: 'C2', v: 1 },
          { ref: 'D2', v: 77 },
          { ref: 'E2', v: 77 },
          { ref: 'F2', v: 10 },
          { ref: 'G2', v: 10 },
          { ref: 'A3', v: 'num1_b' },
          { ref: 'B3', v: 1 },
          { ref: 'C3', v: 1 },
          { ref: 'D3', v: 20 },
          { ref: 'E3', v: 20 },
          { ref: 'F3', v: 77 },
          { ref: 'G3', v: 77 },
          { ref: 'A4', v: 'str1' },
          { ref: 'B4', v: '1' },
          { ref: 'C4', v: '1' },
          { ref: 'D4', v: 30 },
          { ref: 'E4', v: 30 },
          { ref: 'F4', v: 40 },
          { ref: 'G4', v: 40 },
          { ref: 'A5', v: 'upperA' },
          { ref: 'B5', v: 'A' },
          { ref: 'C5', v: 'A' },
          { ref: 'A6', v: 'lowerA' },
          { ref: 'B6', v: 'a' },
          { ref: 'C6', v: 'a' },
          { ref: 'A7', v: 'accentA' },
          { ref: 'B7', v: 'á' },
          { ref: 'C7', v: 'á' },
          { ref: 'A8', v: 'paddedA' },
          { ref: 'B8', v: ' a ' },
          { ref: 'C8', v: ' a ' },
          { ref: 'A9', v: 'errNA' },
          { ref: 'B9', err: '#N/A' },
          { ref: 'C9', err: '#N/A' },
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B9"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D4 F2:F4"><cfRule type="duplicateValues" dxfId="1" priority="2"/></conditionalFormatting>',
        ].join(''),
        dxfsXml: DEFAULT_DXFS_XML,
      }

    case 'stage1-priority-composition':
      return {
        dimension: 'A1:G8',
        cells: [
          { ref: 'A1', v: 'Row' },
          { ref: 'B1', v: 'Compose_CF' },
          { ref: 'C1', v: 'Compose_Ctrl' },
          { ref: 'D1', v: 'BareStop_CF' },
          { ref: 'E1', v: 'BareStop_Ctrl' },
          { ref: 'A2', v: 'pos10' },
          { ref: 'B2', v: 10 },
          { ref: 'C2', v: 10 },
          { ref: 'D2', v: 10 },
          { ref: 'E2', v: 10 },
          { ref: 'A3', v: 'neg5' },
          { ref: 'B3', v: -5 },
          { ref: 'C3', v: -5 },
          { ref: 'D3', v: -5 },
          { ref: 'E3', v: -5 },
          { ref: 'A4', v: 'pos20' },
          { ref: 'B4', v: 20 },
          { ref: 'C4', v: 20 },
          { ref: 'D4', v: 20 },
          { ref: 'E4', v: 20 },
        ],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2:B4">',
          '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule>',
          '<cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
          '<conditionalFormatting sqref="D2:D4">',
          '<cfRule type="cellIs" priority="3" stopIfTrue="1" operator="greaterThan"><formula>0</formula></cfRule>',
          '<cfRule type="cellIs" dxfId="2" priority="4" operator="greaterThan"><formula>0</formula></cfRule>',
          '</conditionalFormatting>',
        ].join(''),
        dxfsXml: [
          '<dxfs count="3">',
          '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><font><b/></font><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>',
          '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>',
          '</dxfs>',
        ].join(''),
      }

    case 'stage1-dxf-layering':
      return {
        dimension: 'A1:G8',
        cells: [
          { ref: 'A1', v: 'Row' },
          { ref: 'B1', v: 'Clears_CF' },
          { ref: 'C1', v: 'Clears_Ctrl' },
          { ref: 'A2', v: 'styled1' },
          { ref: 'B2', v: 25, s: 1 },
          { ref: 'C2', v: 25, s: 1 },
          { ref: 'A3', v: 'styled2' },
          { ref: 'B3', v: 50, s: 1 },
          { ref: 'C3', v: 50, s: 1 },
          { ref: 'A4', v: 'neg' },
          { ref: 'B4', v: -10, s: 1 },
          { ref: 'C4', v: -10, s: 1 },
          { ref: 'A5', v: 'zero' },
          { ref: 'B5', v: 0, s: 1 },
          { ref: 'C5', v: 0, s: 1 },
        ],
        legacyCfXml:
          '<conditionalFormatting sqref="B2:B5"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
        dxfsXml: [
          '<dxfs count="1">',
          '<dxf>',
          '<font><b val="0"/><u val="none"/><strike val="0"/><color rgb="FF9C0006"/></font>',
          '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>',
          '<border><top style="thin"><color rgb="FF9C0006"/></top><bottom style="thin"><color rgb="FF9C0006"/></bottom></border>',
          '</dxf>',
          '</dxfs>',
        ].join(''),
      }

    case 'stage1-blank-merge-order':
      return {
        dimension: 'A1:H6',
        cells: [
          { ref: 'A1', v: 'H1' },
          { ref: 'B1', v: 'H2' },
          { ref: 'C1', v: 'H3' },
          { ref: 'D1', v: 'H4' },
          { ref: 'E1', v: 'H5' },
          { ref: 'F1', v: 'H6' },
          { ref: 'G1', v: 'H7' },
          { ref: 'H1', v: 'H8' },
          // Row 2 has A2 populated, B2 unstored blank with CF, D2 unstored blank control
          { ref: 'A2', v: 'SparseRow' },
          // Row 4 has A4:C4 merged with anchor A4 = 42, and D4 = 42 unmerged control
          { ref: 'A4', v: 42 },
          { ref: 'D4', v: 42 },
        ],
        merges: ['A4:C4'],
        legacyCfXml: [
          '<conditionalFormatting sqref="B2"><cfRule type="containsBlanks" dxfId="0" priority="1"><formula>LEN(TRIM(B2))=0</formula></cfRule></conditionalFormatting>',
          '<conditionalFormatting sqref="A4:C4"><cfRule type="cellIs" dxfId="1" priority="2" operator="equal"><formula>42</formula></cfRule></conditionalFormatting>',
        ].join(''),
        dxfsXml: [
          '<dxfs count="2">',
          '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
          '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill><border><left style="thin"><color rgb="FF006100"/></left><right style="thin"><color rgb="FF006100"/></right><top style="thin"><color rgb="FF006100"/></top><bottom style="thin"><color rgb="FF006100"/></bottom></border></dxf>',
          '</dxfs>',
        ].join(''),
      }

    case 'stage1-numfmt-accessibility':
      return {
        dimension: 'A1:E6',
        cells: [
          { ref: 'A1', v: 'Label' },
          { ref: 'B1', v: 'NumFmt_CF' },
          { ref: 'C1', v: 'NumFmt_Ctrl' },
          { ref: 'A2', v: 'pos_1_25' },
          { ref: 'B2', v: 1.25, s: 3 },
          { ref: 'C2', v: 1.25, s: 3 },
          { ref: 'A3', v: 'neg_0_25' },
          { ref: 'B3', v: 0.25, s: 3 },
          { ref: 'C3', v: 0.25, s: 3 },
          { ref: 'A4', v: 'pos_2_5' },
          { ref: 'B4', v: 2.5, s: 3 },
          { ref: 'C4', v: 2.5, s: 3 },
        ],
        legacyCfXml:
          '<conditionalFormatting sqref="B2:B4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>',
        dxfsXml: [
          '<dxfs count="1">',
          '<dxf>',
          '<font><color rgb="FF9C0006"/></font>',
          '<numFmt numFmtId="165" formatCode="0.00"/>',
          '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>',
          '</dxf>',
          '</dxfs>',
        ].join(''),
      }

    default:
      throw new Error(`Unknown Stage 1 CF case id: ${id}`)
  }
}

function buildStylesXml(dxfsXml: string): string {
  // Cell style indices:
  // s=0: Default Calibri 11
  // s=1: Bold + Single Underline + Strike (for stage1-dxf-layering clear testing)
  // s=2: numFmtId 164 ("00000" for stage1-text-predicates formatted number)
  // s=3: numFmtId 2 ("0.00" / "0.0" -> numFmtId 166 "0.0" for stage1-numfmt-accessibility)
  return [
    XML_DECL,
    `<styleSheet xmlns="${NS_MAIN}">`,
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="00000"/><numFmt numFmtId="166" formatCode="0.0"/></numFmts>',
    '<fonts count="2">',
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
    '<font><b/><u/><strike/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
    '</fonts>',
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    '<cellXfs count="4">',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    '</cellXfs>',
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    dxfsXml,
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>',
    '</styleSheet>',
  ].join('')
}

const colIndex0 = (letters: string): number => {
  let n = 0
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64)
  return n - 1
}

function buildWorksheetXml(payload: SheetPayload): string {
  const byRow = new Map<number, Array<{ col: number; spec: CellSpec }>>()
  for (const spec of payload.cells) {
    const m = /^([A-Z]+)(\d+)$/.exec(spec.ref)!
    const col = colIndex0(m[1])
    const row = Number(m[2])
    const list = byRow.get(row) ?? []
    list.push({ col, spec })
    byRow.set(row, list)
  }
  const maxRow = Math.max(10, ...byRow.keys())
  const rowsXml: string[] = []
  for (let r = 1; r <= maxRow; r++) {
    const cells = (byRow.get(r) ?? []).sort((a, b) => a.col - b.col)
    const cellsXml = cells
      .map(({ spec }) => {
        const sAttr = spec.s ? ` s="${spec.s}"` : ''
        const fXml = spec.f ? `<f>${esc(spec.f)}</f>` : ''
        if (spec.err !== undefined) {
          return `<c r="${spec.ref}"${sAttr} t="e">${fXml}<v>${esc(spec.err)}</v></c>`
        }
        if (spec.v === null || spec.v === undefined) {
          return sAttr ? `<c r="${spec.ref}"${sAttr}/>` : ''
        }
        if (typeof spec.v === 'number') {
          return `<c r="${spec.ref}"${sAttr}>${fXml}<v>${spec.v}</v></c>`
        }
        if (typeof spec.v === 'boolean') {
          return `<c r="${spec.ref}"${sAttr} t="b">${fXml}<v>${spec.v ? 1 : 0}</v></c>`
        }
        return `<c r="${spec.ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${esc(spec.v)}</t></is></c>`
      })
      .join('')
    rowsXml.push(`<row r="${r}" ht="18" customHeight="1">${cellsXml}</row>`)
  }
  const mergeXml = payload.merges?.length
    ? `<mergeCells count="${payload.merges.length}">${payload.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`
    : ''
  const extXml = payload.x14CfXml
    ? `<extLst><ext uri="${X14_CF_EXT_URI}" xmlns:x14="${NS_X14}"><x14:conditionalFormattings>${payload.x14CfXml}</x14:conditionalFormattings></ext></extLst>`
    : ''
  return [
    XML_DECL,
    `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" mc:Ignorable="x14ac" xmlns:x14ac="${NS_X14AC}" xmlns:x14="${NS_X14}" xmlns:xm="${NS_XM}">`,
    `<dimension ref="${payload.dimension}"/>`,
    '<sheetViews><sheetView showGridLines="1" workbookViewId="0"/></sheetViews>',
    '<sheetFormatPr defaultRowHeight="18"/>',
    '<cols><col min="1" max="8" width="14" customWidth="1"/></cols>',
    `<sheetData>${rowsXml.join('')}</sheetData>`,
    mergeXml,
    payload.legacyCfXml,
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    extXml,
    '</worksheet>',
  ].join('')
}

export async function generateCaseArchive(c: CfOracleCase): Promise<Uint8Array> {
  const payload = casePayload(c.id)
  const zip = new JSZip()
  const files: Record<string, string> = {
    '[Content_Types].xml': `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels': `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `${XML_DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><workbookPr defaultThemeVersion="166925"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="17500"/></bookViews><sheets><sheet name="${esc(c.sheet)}" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`,
    'xl/_rels/workbook.xml.rels': `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${NS_R}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': buildStylesXml(payload.dxfsXml),
    'xl/worksheets/sheet1.xml': buildWorksheetXml(payload),
  }
  for (const [name, content] of Object.entries(files)) {
    zip.file(name, content, { date: ZIP_DATE, createFolders: false })
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

export async function generateAllStage1(): Promise<CfManifest> {
  mkdirSync(CF_FIXTURE_DIR, { recursive: true })
  mkdirSync(CF_INPUT_DIR, { recursive: true })
  mkdirSync(CF_REFS_DIR, { recursive: true })

  let existingManifest: CfManifest | undefined
  if (existsSync(CF_MANIFEST_PATH)) {
    try {
      existingManifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    } catch {
      existingManifest = undefined
    }
  }
  const previousById = new Map((existingManifest?.cases ?? []).map(c => [c.id, c]))

  const updatedCases: CfOracleCase[] = []
  for (const baseCase of CF_STAGE1_CASES) {
    const inputBytes = await generateCaseArchive(baseCase)
    const inputSha256 = sha256Hex(inputBytes)
    const inputAbs = resolve(REPO_ROOT, baseCase.input)
    const savedAbs = resolve(REPO_ROOT, baseCase.saved)
    writeFileSync(inputAbs, inputBytes)

    const prev = previousById.get(baseCase.id)
    const prevVerified = prev ? verifyCfCase(prev, { rootDir: REPO_ROOT }) : undefined
    if (prev && prevVerified?.kind === 'verified' && prev.expected.inputSha256 === inputSha256 && existsSync(savedAbs)) {
      updatedCases.push(prev)
      continue
    }

    writeFileSync(savedAbs, inputBytes)
    const savedSha256 = sha256Hex(new Uint8Array(readFileSync(savedAbs)))
    updatedCases.push({
      ...baseCase,
      expected: {
        ...baseCase.expected,
        inputSha256,
        savedSha256,
      },
    })
  }

  const manifest: CfManifest = {
    schemaVersion: 1,
    excelVersion: '16.114',
    uiLocale: 'zh-TW',
    generatedAt: '2026-11-06T00:00:00.000Z',
    cases: updatedCases,
  }
  writeFileSync(CF_MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n')
  return manifest
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const manifest = await generateAllStage1()
  console.log(`Generated ${manifest.cases.length} Stage 1 CF oracle workbooks and updated ${CF_MANIFEST_PATH}`)
}
