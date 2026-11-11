import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { cpSync, existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'

const hasOsacompile = process.platform === 'darwin' && spawnSync('which', ['osacompile']).status === 0
import { compareCfCase, stage3ObservedBars } from '../scripts/cf-compare'
import { auditStage3Manifest, renderVerifiedCapabilitiesSource } from '../scripts/cf-refs/audit'
import { assertCanCloseWorkbook, filterCasesForStage, isDisplayLocked, parseScreenLockState, shouldSkipCaseOnResume } from '../scripts/cf-refs/build'
import {
  CF_CAPABILITIES,
  CF_CASES,
  CF_STAGE1_CAPABILITIES,
  CF_STAGE1_CASES,
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_CASES,
  CF_STAGE3_CAPABILITIES,
  CF_STAGE3_CASES,
  CF_STAGE3_NONDISCRIMINATING,
  barCapabilityKind,
  type CfCellExpectation,
  type CfManifest,
  type CfOracleCase,
  type CfPartMarker,
  type CfResaveReceipt,
} from '../scripts/cf-refs/cases'
import {
  buildColorProbeScript,
  buildColorScaleCriteriaProbeScript,
  buildDataBarSettingsProbeScript,
  buildBarSidecarExtras,
  buildDecorationProbeScript,
  buildTextProbeScript,
  captureAndCloseOwned,
  captureUnderLockGuard,
  staleCaptureSidecar,
  classifyBarCapture,
  parseDataBarSettingsReadback,
  parseExcelColorToArgb,
  verifyOnlyCaptureState,
} from '../scripts/cf-refs/capture'
import { hasEvidence, VERIFIED_CAPABILITIES, VERIFIED_STAGE1_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES } from '../src/xlsx/cf/evidence'
import { CF_MANIFEST_PATH, generateAllStage1 } from '../scripts/cf-refs/generate'
import {
  type CfObservedBar,
  type ComputeGeometryOptions,
  type WorksheetRangeGeometry,
  computeWorksheetRangeGeometry,
  decodePngSync,
  BAR_PIXEL_PROBE_UNOBSERVABLE,
  discriminateBar,
  evaluateBarPixelProbe,
  isCapabilityVerified,
  isCellWithinRange,
  isSupportedExcelVersion,
  measureBarGeometry,
  measureCaseBars,
  observeCellsFromCapture,
  parseA1,
  readZipPartsSync,
  verifyCfCase,
} from '../scripts/cf-refs/verify'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const ZIP_DATE = new Date(Date.UTC(2026, 0, 1))

const sha256 = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')

async function makeSyntheticXlsx(opts: {
  sheetRuleXml?: string
  workbookExtraXml?: string
  dxfFillRgb?: string
  dxfCount?: number
  repaired?: boolean
}): Promise<Uint8Array> {
  const sheetRule =
    opts.sheetRuleXml ??
    '<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>'
  const dxfFill = opts.dxfFillRgb ?? 'FFFFC7CE'
  const dxfsXml =
    opts.dxfCount === 0
      ? '<dxfs count="0"/>'
      : `<dxfs count="1"><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="${dxfFill}"/></patternFill></fill></dxf></dxfs>`
  const zip = new JSZip()
  const files: Record<string, string> = {
    '[Content_Types].xml': `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels': `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `${XML_DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>${opts.workbookExtraXml ?? ''}</workbook>`,
    'xl/_rels/workbook.xml.rels': `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${NS_R}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `${XML_DECL}<styleSheet xmlns="${NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>${dxfsXml}</styleSheet>`,
    'xl/worksheets/sheet1.xml': `${XML_DECL}<worksheet xmlns="${NS_MAIN}"><dimension ref="A1:B3"/><sheetFormatPr defaultRowHeight="18" defaultColWidth="14"/><sheetData><row r="1" ht="18"><c r="A1" t="inlineStr"><is><t>Label</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2" ht="18"><c r="A2" t="inlineStr"><is><t>Pos</t></is></c><c r="B2"><v>150</v></c></row><row r="3" ht="18"><c r="A3"><v>50</v></c><c r="B3"><v>50</v></c></row></sheetData>${sheetRule}</worksheet>`,
  }
  if (opts.repaired) {
    files['xl/repairedRecords/repairedRecord1.xml'] = `${XML_DECL}<repairedRecord/>`
  }
  for (const [name, content] of Object.entries(files)) {
    zip.file(name, content, { date: ZIP_DATE, createFolders: false })
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/**
 * Renders a synthetic `168x54` PNG for `A1:B3` (2 cols x 84px, 3 rows x 18px)
 * where `B2` (`x: 84..168, y: 18..36`) has `#ffc7ce` and `#9c0006` bold font when `positiveHighlight` is true.
 */
function makeCapturePng(positiveHighlight: boolean): Uint8Array {
  const canvas = createCanvas(168, 54)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, 168, 54)
  if (positiveHighlight) {
    ctx.fillStyle = '#ffc7ce'
    ctx.fillRect(84, 18, 84, 18)
    ctx.fillStyle = '#9c0006'
    ctx.font = 'bold 11px sans-serif'
    ctx.fillText('150', 94, 32)
  } else {
    ctx.fillStyle = '#000000'
    ctx.font = '11px sans-serif'
    ctx.fillText('150', 94, 32)
  }
  ctx.fillStyle = '#000000'
  ctx.font = '11px sans-serif'
  ctx.fillText('50', 94, 50)
  ctx.fillText('50', 10, 50)
  return new Uint8Array(canvas.toBuffer('image/png'))
}

/** Renders a synthetic capture matching every `expected.cells` entry of a `CfOracleCase` for testing. */
function makeCaseMatchingPng(
  c: CfOracleCase,
  savedBytes: Uint8Array,
  paintMatchedFills: boolean,
  geomOptions?: ComputeGeometryOptions,
): Uint8Array {
  const parts = readZipPartsSync(savedBytes)!
  const geom = computeWorksheetRangeGeometry(parts, c.range, geomOptions)!
  const width = geom.expectedWidthPt
  const height = geom.expectedHeightPt
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, width, height)

  const totalW = geom.expectedWidthPt
  const totalH = geom.expectedHeightPt
  for (const cell of c.expected.cells) {
    const pos = parseA1(cell.ref)!
    const colIdx = pos.col - geom.minCol
    const rowIdx = pos.row - geom.minRow
    const colBefore = geom.colWidthsPt.slice(0, colIdx).reduce((a, b) => a + b, 0)
    const colSpan = geom.colWidthsPt[colIdx]
    const rowBefore = geom.rowHeightsPt.slice(0, rowIdx).reduce((a, b) => a + b, 0)
    const rowSpan = geom.rowHeightsPt[rowIdx]
    const x0 = Math.round((colBefore / totalW) * width)
    const x1 = Math.max(x0 + 1, Math.round(((colBefore + colSpan) / totalW) * width))
    const y0 = Math.round((rowBefore / totalH) * height)
    const y1 = Math.max(y0 + 1, Math.round(((rowBefore + rowSpan) / totalH) * height))

    if (paintMatchedFills && cell.matched) {
      if (cell.fillRgb) {
        const hex6 = cell.fillRgb.replace(/^#/, '').slice(-6)
        ctx.fillStyle = `#${hex6}`
        ctx.fillRect(x0, y0, x1 - x0, y1 - y0)
      }
      if (cell.borderTop) {
        ctx.strokeStyle = '#006100'
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(x0, y0 + 0.5)
        ctx.lineTo(x1, y0 + 0.5)
        ctx.stroke()
      }
      if (cell.borderBottom) {
        ctx.strokeStyle = '#006100'
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(x0, y1 - 0.5)
        ctx.lineTo(x1, y1 - 0.5)
        ctx.stroke()
      }
      const displayText = cell.formattedText ?? cell.ref
      const fontHex = cell.fontRgb ? cell.fontRgb.replace(/^#/, '').slice(-6) : '000000'
      ctx.fillStyle = `#${fontHex}`
      ctx.font = cell.bold ? 'bold 11px sans-serif' : '11px sans-serif'
      ctx.fillText(displayText, x0 + 4, y0 + 13)
      const textW = ctx.measureText(displayText).width
      if (cell.underline === 'single' || cell.underline === 'double') {
        ctx.strokeStyle = `#${fontHex}`
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(x0 + 4, y0 + 15)
        ctx.lineTo(x0 + 4 + textW, y0 + 15)
        ctx.stroke()
      }
      if (cell.strikethrough === true) {
        ctx.strokeStyle = `#${fontHex}`
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(x0 + 4, y0 + 10)
        ctx.lineTo(x0 + 4 + textW, y0 + 10)
        ctx.stroke()
      }
    } else {
      const displayText = cell.formattedText ?? cell.ref
      ctx.fillStyle = '#000000'
      ctx.font = '11px sans-serif'
      ctx.fillText(displayText, x0 + 4, y0 + 13)
      const textW = ctx.measureText(displayText).width
      if (cell.underline === 'single' || cell.underline === 'double') {
        ctx.strokeStyle = '#000000'
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(x0 + 4, y0 + 15)
        ctx.lineTo(x0 + 4 + textW, y0 + 15)
        ctx.stroke()
      }
      if (cell.strikethrough === true) {
        ctx.strokeStyle = '#000000'
        ctx.lineWidth = 1
        ctx.beginPath()
        ctx.moveTo(x0 + 4, y0 + 10)
        ctx.lineTo(x0 + 4 + textW, y0 + 10)
        ctx.stroke()
      }
    }

    if (cell.controlRef) {
      const cPos = parseA1(cell.controlRef)
      if (cPos) {
        const cColIdx = cPos.col - geom.minCol
        const cRowIdx = cPos.row - geom.minRow
        if (cColIdx >= 0 && cRowIdx >= 0 && cColIdx < geom.colWidthsPt.length && cRowIdx < geom.rowHeightsPt.length) {
          const cColBefore = geom.colWidthsPt.slice(0, cColIdx).reduce((a, b) => a + b, 0)
          const cRowBefore = geom.rowHeightsPt.slice(0, cRowIdx).reduce((a, b) => a + b, 0)
          const cx0 = Math.round((cColBefore / totalW) * width)
          const cy0 = Math.round((cRowBefore / totalH) * height)
          // For matched cells, control cell has unformatted text (controlRef), not the CF formatted text
          // For unmatched cells, control cell has identical text as the cell (so they match)
          const ctrlText = cell.matched
            ? cell.formattedText !== undefined
              ? cell.controlRef
              : cell.ref
            : (cell.formattedText ?? cell.ref)
          ctx.fillStyle = '#000000'
          ctx.font = '11px sans-serif'
          ctx.fillText(ctrlText, cx0 + 4, cy0 + 13)
          const ctrlW = ctx.measureText(ctrlText).width
          const cUnderline = (cell as any).controlUnderline
          if (cUnderline === 'single' || cUnderline === 'double') {
            ctx.strokeStyle = '#000000'
            ctx.lineWidth = 1
            ctx.beginPath()
            ctx.moveTo(cx0 + 4, cy0 + 15)
            ctx.lineTo(cx0 + 4 + ctrlW, cy0 + 15)
            ctx.stroke()
          }
          if ((cell as any).controlStrikethrough === true) {
            ctx.strokeStyle = '#000000'
            ctx.lineWidth = 1
            ctx.beginPath()
            ctx.moveTo(cx0 + 4, cy0 + 10)
            ctx.lineTo(cx0 + 4 + ctrlW, cy0 + 10)
            ctx.stroke()
          }
        }
      }
    }
  }
  return new Uint8Array(canvas.toBuffer('image/png'))
}

/** Old painter counterexample from r2-probe.ts: only solid rectangles or 4px black bars, no text/font/borders. */
function makeFillBarOnlyPng(c: CfOracleCase, savedBytes: Uint8Array): Uint8Array {
  const parts = readZipPartsSync(savedBytes)!
  const geom = computeWorksheetRangeGeometry(parts, c.range)!
  const width = geom.expectedWidthPt
  const height = geom.expectedHeightPt
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, width, height)

  const totalW = geom.expectedWidthPt
  const totalH = geom.expectedHeightPt
  for (const cell of c.expected.cells) {
    if (!cell.matched) continue
    const pos = parseA1(cell.ref)!
    const colIdx = pos.col - geom.minCol
    const rowIdx = pos.row - geom.minRow
    const colBefore = geom.colWidthsPt.slice(0, colIdx).reduce((a, b) => a + b, 0)
    const colSpan = geom.colWidthsPt[colIdx]
    const rowBefore = geom.rowHeightsPt.slice(0, rowIdx).reduce((a, b) => a + b, 0)
    const rowSpan = geom.rowHeightsPt[rowIdx]
    const x0 = Math.floor((colBefore / totalW) * width)
    const x1 = Math.floor(((colBefore + colSpan) / totalW) * width)
    const y0 = Math.floor((rowBefore / totalH) * height)
    const y1 = Math.floor(((rowBefore + rowSpan) / totalH) * height)
    if (cell.fillRgb) {
      const hex6 = cell.fillRgb.replace(/^#/, '').slice(-6)
      ctx.fillStyle = `#${hex6}`
      ctx.fillRect(x0, y0, x1 - x0, y1 - y0)
    } else {
      ctx.fillStyle = '#000000'
      ctx.fillRect(x0 + 4, y0 + 6, Math.max(4, x1 - x0 - 8), 4)
    }
  }
  return new Uint8Array(canvas.toBuffer('image/png'))
}

describe('CF oracle verifier', () => {
  const dirs: string[] = []
  afterEach(() => {
    while (dirs.length) rmSync(dirs.pop()!, { recursive: true, force: true })
  })

  it('rejects stripped rules, wrong-part markers and unchanged control captures', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-oracle-'))
    dirs.push(dir)
    mkdirSync(join(dir, 'refs'), { recursive: true })

    // 0. Valid measured baseline fixture
    const validXlsx = await makeSyntheticXlsx({})
    const validXlsxHash = sha256(validXlsx)
    writeFileSync(join(dir, 'valid.xlsx'), validXlsx)
    const validPng = makeCapturePng(true)
    const validPngHash = sha256(validPng)
    writeFileSync(join(dir, 'refs/valid.png'), validPng)
    writeFileSync(
      join(dir, 'refs/valid.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        width: 168,
        height: 54,
        geometrySource: 'excel-measured',
        colWidthsPt: [84, 84],
        rowHeightsPt: [18, 18, 18],
        resavedByExcel: true,
        resaveSavedSha256: validXlsxHash,
        sheet: 'Sheet1',
        range: 'A1:B3',
        savedSha256: validXlsxHash,
        pngSha256: validPngHash,
        capturedAt: '2026-11-06T08:00:00Z',
      }),
    )

    const baseCase: CfOracleCase = {
      id: 'syn-cellis',
      stage: 1,
      question: 'Synthetic test case',
      input: 'valid.xlsx',
      saved: 'valid.xlsx',
      resavedByExcel: true,
      resaveReceipt: {
        openedRepairFree: true,
        excelVersion: '16.114',
        inputSha256: validXlsxHash,
        savedSha256: validXlsxHash,
        resavedAt: '2026-11-06T08:00:00Z',
      },
      sheet: 'Sheet1',
      range: 'A1:B3',
      requiredAppearances: ['screen'],
      markers: [
        {
          part: 'xl/worksheets/sheet1.xml',
          extension: 'legacy',
          capabilityId: 'syn:cellIs-gt',
          sqref: 'B2:B3',
          ruleType: 'cellIs',
          operator: 'greaterThan',
          priority: 1,
          formulas: ['100'],
          dxf: {
            fillRgb: 'FFFFC7CE',
            fontRgb: 'FF9C0006',
            bold: true,
          },
        },
      ],
      expected: {
        inputSha256: validXlsxHash,
        savedSha256: validXlsxHash,
        provenance: 'excel-observed',
        positiveRefs: ['B2'],
        negativeRefs: ['B3'],
        cells: [
          { ref: 'B2', controlRef: 'B3', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true },
          { ref: 'B3', controlRef: 'A3', matched: false },
        ],
      },
      captures: [
        {
          appearance: 'screen',
          pngPath: 'refs/valid.png',
          sidecarPath: 'refs/valid.json',
          pngSha256: validPngHash,
        },
      ],
      status: 'measured',
      tolerance: { meanAbs: 0.015, rowMax: 0.08, strong: 12 },
    }

    const verified = verifyCfCase(baseCase, { rootDir: dir })
    expect(verified.kind).toBe('verified')
    if (verified.kind === 'verified') {
      expect(verified.savedSha256).toBe(validXlsxHash)
      expect(verified.positiveContrastMaxDelta).toBeGreaterThan(40)
      expect(verified.minPositiveContrastDelta).toBeGreaterThan(40)
    }

    // 1. Mutated archive: remove target rule from worksheet
    const strippedXlsx = await makeSyntheticXlsx({ sheetRuleXml: '' })
    const strippedHash = sha256(strippedXlsx)
    writeFileSync(join(dir, 'stripped.xlsx'), strippedXlsx)
    const strippedResult = verifyCfCase(
      {
        ...baseCase,
        id: 'syn-stripped',
        saved: 'stripped.xlsx',
        expected: { ...baseCase.expected, savedSha256: strippedHash },
      },
      { rootDir: dir },
    )
    expect(strippedResult).toMatchObject({
      kind: 'refusal',
      code: 'stripped-rule',
      structuralValid: false,
    })

    // 2a. Mutated archive: move identical rule XML text to another package part (xl/workbook.xml)
    const movedRuleXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>'
    const wrongPartXlsx = await makeSyntheticXlsx({
      sheetRuleXml: '',
      workbookExtraXml: movedRuleXml,
    })
    const wrongPartHash = sha256(wrongPartXlsx)
    writeFileSync(join(dir, 'wrong-part.xlsx'), wrongPartXlsx)
    const wrongPartResult = verifyCfCase(
      {
        ...baseCase,
        id: 'syn-wrong-part',
        saved: 'wrong-part.xlsx',
        expected: { ...baseCase.expected, savedSha256: wrongPartHash },
      },
      { rootDir: dir },
    )
    expect(wrongPartResult).toMatchObject({
      kind: 'refusal',
      code: 'wrong-part-marker',
      structuralValid: false,
    })

    // 2b. Mutated archive: break rule -> target -> DXF pairing (mismatched DXF fill)
    const badDxfXlsx = await makeSyntheticXlsx({ dxfFillRgb: 'FF00FF00' })
    const badDxfHash = sha256(badDxfXlsx)
    writeFileSync(join(dir, 'bad-dxf.xlsx'), badDxfXlsx)
    const badDxfResult = verifyCfCase(
      {
        ...baseCase,
        id: 'syn-bad-dxf',
        saved: 'bad-dxf.xlsx',
        expected: { ...baseCase.expected, savedSha256: badDxfHash },
      },
      { rootDir: dir },
    )
    expect(badDxfResult).toMatchObject({
      kind: 'refusal',
      code: 'dxf-pairing-failed',
      structuralValid: false,
    })

    // 2c. Mutated archive: Excel repair log present
    const repairedXlsx = await makeSyntheticXlsx({ repaired: true })
    const repairedHash = sha256(repairedXlsx)
    writeFileSync(join(dir, 'repaired.xlsx'), repairedXlsx)
    const repairedResult = verifyCfCase(
      {
        ...baseCase,
        id: 'syn-repaired',
        saved: 'repaired.xlsx',
        expected: { ...baseCase.expected, savedSha256: repairedHash },
      },
      { rootDir: dir },
    )
    expect(repairedResult).toMatchObject({
      kind: 'refusal',
      code: 'repair-log-present',
      structuralValid: false,
    })

    // 3. Mutated capture: reuse unformatted baseline image (positive B2 and negative B3 look identical)
    const baselinePng = makeCapturePng(false)
    const baselinePngHash = sha256(baselinePng)
    writeFileSync(join(dir, 'refs/unchanged.png'), baselinePng)
    writeFileSync(
      join(dir, 'refs/unchanged.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        width: 168,
        height: 54,
        geometrySource: 'excel-measured',
        colWidthsPt: [84, 84],
        rowHeightsPt: [18, 18, 18],
        resavedByExcel: true,
        resaveSavedSha256: validXlsxHash,
        sheet: 'Sheet1',
        range: 'A1:B3',
        savedSha256: validXlsxHash,
        pngSha256: baselinePngHash,
        capturedAt: '2026-11-06T08:00:00Z',
      }),
    )
    const unchangedResult = verifyCfCase(
      {
        ...baseCase,
        id: 'syn-unchanged-control',
        captures: [
          {
            appearance: 'screen',
            pngPath: 'refs/unchanged.png',
            sidecarPath: 'refs/unchanged.json',
            pngSha256: baselinePngHash,
          },
        ],
      },
      { rootDir: dir },
    )
    expect(unchangedResult.kind).toBe('refusal')
    if (unchangedResult.kind === 'refusal') {
      expect(['cell-expectation-mismatch', 'unchanged-control-capture']).toContain(unchangedResult.code)
      expect(unchangedResult.structuralValid).toBe(true)
    }

    // 4. Synthetic fixture with correct structural markers but pending host capture
    const pendingCase: CfOracleCase = {
      ...baseCase,
      id: 'syn-pending',
      status: 'recipe-pending',
      captures: [],
      reason: 'Awaiting host Excel capture window',
    }
    const pendingResult = verifyCfCase(pendingCase, { rootDir: dir })
    expect(pendingResult).toMatchObject({
      kind: 'refusal',
      status: 'recipe-pending',
      code: 'semantic-evidence-pending',
      structuralValid: true,
    })

    // 5. Resume guard: skips ONLY verified matching-input cases
    expect(shouldSkipCaseOnResume(baseCase, validXlsxHash, dir)).toBe(true)
    expect(shouldSkipCaseOnResume(baseCase, 'different-input-hash', dir)).toBe(false)
    expect(shouldSkipCaseOnResume(pendingCase, validXlsxHash, dir)).toBe(false)

    // 6. Unrelated workbook guard: refuses to close pre-existing open workbooks
    expect(() => assertCanCloseWorkbook('UserBudget.xlsx', new Set(['UserBudget.xlsx']))).toThrow(
      /Refusing to close unrelated pre-existing open workbook/,
    )
    expect(() => assertCanCloseWorkbook('.work-stage1.xlsx', new Set(['UserBudget.xlsx']))).not.toThrow()

    // 7. Task 8 exercises the real raster route; synthetic capture pixels
    // are a comparator fixture, never Excel-fidelity evidence.
    const cmpRes = await compareCfCase(baseCase, 'static', { rootDir: dir })
    expect(cmpRes.outcome).not.toBe('implementation-pending')
    expect(cmpRes.diff).toBeDefined()
    const print = await compareCfCase(baseCase, 'print', { rootDir: dir })
    expect(print).toMatchObject({ outcome: 'fail', reason: expect.stringContaining('No printer capture') })
  })

  it('verifies all Stage 1 manifest cases have valid structural markers and honest status', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const stage1Cases = manifest.cases.filter(c => c.stage === 1 || !c.stage)
    expect(stage1Cases.map(c => c.id)).toEqual(CF_STAGE1_CASES.map(c => c.id))
    expect(stage1Cases).toHaveLength(13)

    for (const c of stage1Cases) {
      expect(c.expected.inputSha256).toMatch(/^[0-9a-f]{64}$/)
      expect(c.expected.savedSha256).toMatch(/^[0-9a-f]{64}$/)
      const result = verifyCfCase(c, { rootDir: REPO_ROOT })
      if (result.kind === 'refusal') {
        expect(result.structuralValid).toBe(true)
        expect(['semantic-evidence-pending', 'unsupported-capability']).toContain(result.code)
      } else {
        expect(result.status).toBe('measured')
      }
    }
  })

  it('scale marker pins cfvo and colour order exactly', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-scale-pin-'))
    dirs.push(dir)

    const baseMarker: CfPartMarker = {
      part: 'xl/worksheets/sheet1.xml',
      extension: 'legacy',
      capabilityId: 'scale:pin-test',
      sqref: 'B2:B3',
      ruleType: 'colorScale',
      priority: 1,
      formulas: [],
      scale: {
        cfvo: [{ type: 'min' }, { type: 'percentile', val: '30' }, { type: 'max' }],
        colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
      },
    }

    const makeScaleCase = (filename: string, hash: string, markers = [baseMarker]): CfOracleCase => ({
      id: 'scale-pin-case',
      stage: 2,
      question: 'Pins exact cfvo sequence and color sequence in scale marker',
      input: filename,
      saved: filename,
      sheet: 'Sheet1',
      range: 'A1:B3',
      requiredAppearances: ['screen'],
      markers,
      expected: {
        inputSha256: hash,
        savedSha256: hash,
        provenance: 'seed-hypothesis',
        positiveRefs: ['B2'],
        negativeRefs: ['B3'],
        cells: [
          { ref: 'B2', controlRef: 'A2', matched: true, fillSource: 'unverified' },
          { ref: 'B3', controlRef: 'A3', matched: false, fillSource: 'unverified' },
        ],
      },
      captures: [],
      status: 'recipe-pending',
      reason: 'Testing scale marker pinning',
      tolerance: { meanAbs: 0.015, rowMax: 0.08, strong: 12 },
    })

    // 0. Canonical valid scale archive: structuralValid should be true
    const validScaleXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="30"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const validBytes = await makeSyntheticXlsx({ sheetRuleXml: validScaleXml, dxfCount: 0 })
    const validHash = sha256(validBytes)
    writeFileSync(join(dir, 'valid.xlsx'), validBytes)
    const validRes = verifyCfCase(makeScaleCase('valid.xlsx', validHash), { rootDir: dir })
    expect(validRes).toMatchObject({ structuralValid: true })

    // 1. Reordered colours: FF63BE7B -> FFFFEB84 -> FFF8696B
    const reorderedColorsXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="30"/><cfvo type="max"/><color rgb="FF63BE7B"/><color rgb="FFFFEB84"/><color rgb="FFF8696B"/></colorScale></cfRule></conditionalFormatting>'
    const reorderedBytes = await makeSyntheticXlsx({ sheetRuleXml: reorderedColorsXml, dxfCount: 0 })
    const reorderedHash = sha256(reorderedBytes)
    writeFileSync(join(dir, 'reordered.xlsx'), reorderedBytes)
    const reorderedRes = verifyCfCase(makeScaleCase('reordered.xlsx', reorderedHash), { rootDir: dir })
    expect(reorderedRes).toMatchObject({
      kind: 'refusal',
      code: 'stripped-rule',
      structuralValid: false,
    })

    // 2. Missing cfvo: only min and max (2 cfvos instead of 3)
    const missingCfvoXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const missingBytes = await makeSyntheticXlsx({ sheetRuleXml: missingCfvoXml, dxfCount: 0 })
    const missingHash = sha256(missingBytes)
    writeFileSync(join(dir, 'missing.xlsx'), missingBytes)
    const missingRes = verifyCfCase(makeScaleCase('missing.xlsx', missingHash), { rootDir: dir })
    expect(missingRes).toMatchObject({
      kind: 'refusal',
      code: 'stripped-rule',
      structuralValid: false,
    })

    // 3. Mutated val: "30" changed to "3"
    const mutatedValXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="3"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const mutatedValBytes = await makeSyntheticXlsx({ sheetRuleXml: mutatedValXml, dxfCount: 0 })
    const mutatedValHash = sha256(mutatedValBytes)
    writeFileSync(join(dir, 'mutated-val.xlsx'), mutatedValBytes)
    const mutatedValRes = verifyCfCase(makeScaleCase('mutated-val.xlsx', mutatedValHash), { rootDir: dir })
    expect(mutatedValRes).toMatchObject({
      kind: 'refusal',
      code: 'stripped-rule',
      structuralValid: false,
    })

    // 4. Rule in wrong part: moved to xl/workbook.xml
    const wrongPartBytes = await makeSyntheticXlsx({ sheetRuleXml: '', workbookExtraXml: validScaleXml, dxfCount: 0 })
    const wrongPartHash = sha256(wrongPartBytes)
    writeFileSync(join(dir, 'wrong-part.xlsx'), wrongPartBytes)
    const wrongPartRes = verifyCfCase(makeScaleCase('wrong-part.xlsx', wrongPartHash), { rootDir: dir })
    expect(wrongPartRes).toMatchObject({
      kind: 'refusal',
      code: 'wrong-part-marker',
      structuralValid: false,
    })

    // 5. Semantic equivalence (Lesson 2): attribute order switched, lowercase hex
    const semanticEquivXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo val="30" type="percentile"/><cfvo type="max"/><color rgb="fff8696b"/><color rgb="ffffeb84"/><color rgb="ff63be7b"/></colorScale></cfRule></conditionalFormatting>'
    const semanticBytes = await makeSyntheticXlsx({ sheetRuleXml: semanticEquivXml, dxfCount: 0 })
    const semanticHash = sha256(semanticBytes)
    writeFileSync(join(dir, 'semantic.xlsx'), semanticBytes)
    const semanticRes = verifyCfCase(makeScaleCase('semantic.xlsx', semanticHash), { rootDir: dir })
    expect(semanticRes).toMatchObject({ structuralValid: true })
  })

  it('scale fill requires trusted fill evidence', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-scale-fill-'))
    dirs.push(dir)

    const baseMarker: CfPartMarker = {
      part: 'xl/worksheets/sheet1.xml',
      extension: 'legacy',
      capabilityId: 'scale:fill-test',
      sqref: 'B2:B3',
      ruleType: 'colorScale',
      priority: 1,
      formulas: [],
      scale: {
        cfvo: [{ type: 'min' }, { type: 'max' }],
        colors: ['FFFF0000', 'FF00FF00'],
      },
    }

    const scaleXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>'
    const xlsxBytes = await makeSyntheticXlsx({ sheetRuleXml: scaleXml, dxfCount: 0 })
    const xlsxHash = sha256(xlsxBytes)
    writeFileSync(join(dir, 'scale.xlsx'), xlsxBytes)

    const validPng = makeCapturePng(true)
    const validPngHash = sha256(validPng)
    writeFileSync(join(dir, 'scale.png'), validPng)

    const makeCase = (fillSource?: 'excel-display-format' | 'pixel-matched' | 'unverified'): CfOracleCase => ({
      id: 'scale-fill-case',
      stage: 2,
      question: 'Scale fill trust evidence',
      input: 'scale.xlsx',
      saved: 'scale.xlsx',
      sheet: 'Sheet1',
      range: 'A1:B3',
      requiredAppearances: ['screen'],
      markers: [baseMarker],
      expected: {
        inputSha256: xlsxHash,
        savedSha256: xlsxHash,
        provenance: 'excel-observed',
        positiveRefs: ['B2'],
        negativeRefs: ['B3'],
        cells: [
          { ref: 'B2', controlRef: 'A2', matched: true, fillRgb: 'FFFF0000', scaleApplied: true, fillSource },
          { ref: 'B3', controlRef: 'A3', matched: false, scaleApplied: false, fillSource },
        ],
      },
      resavedByExcel: true,
      resaveReceipt: {
        openedRepairFree: true,
        excelVersion: '16.114',
        inputSha256: xlsxHash,
        savedSha256: xlsxHash,
        resavedAt: '2026-11-06T08:00:00Z',
      },
      captures: [
        {
          appearance: 'screen',
          pngPath: 'scale.png',
          sidecarPath: 'scale.json',
          pngSha256: validPngHash,
        },
      ],
      status: 'measured',
      tolerance: { meanAbs: 0.015, rowMax: 0.08, strong: 12 },
    })

    const writeSidecar = (
      geomSource = 'excel-measured',
      observedArgb = 'FFFF0000',
      fillSource: 'excel-display-format' | 'pixel-matched' = 'excel-display-format',
    ) => {
      writeFileSync(
        join(dir, 'scale.json'),
        JSON.stringify({
          excelVersion: '16.114',
          route: 'copy-picture',
          appearance: 'screen',
          scale: 1,
          pxPerPt: 1,
          width: 168,
          height: 54,
          colWidthsPt: [84, 84],
          rowHeightsPt: [18, 18, 18],
          geometrySource: geomSource,
          resavedByExcel: true,
          resaveSavedSha256: xlsxHash,
          sheet: 'Sheet1',
          range: 'A1:B3',
          savedSha256: xlsxHash,
          pngSha256: validPngHash,
          capturedAt: '2026-11-06T08:00:00Z',
          observedCells: [
            { ref: 'B2', fillSource, fillRgb: observedArgb },
            { ref: 'B3', fillSource, fillRgb: 'FFFFFFFF' },
          ],
        }),
      )
    }

    // 1. Missing fillSource or 'unverified' -> fill-unobserved
    writeSidecar('excel-measured')
    const unverifiedRes = verifyCfCase(makeCase('unverified'), { rootDir: dir })
    expect(unverifiedRes).toMatchObject({
      kind: 'refusal',
      code: 'fill-unobserved',
    })

    const missingSourceRes = verifyCfCase(makeCase(undefined), { rootDir: dir })
    expect(missingSourceRes).toMatchObject({
      kind: 'refusal',
      code: 'fill-unobserved',
    })

    // 2. Geometry must be excel-measured before pixel sample (T1-A(a))
    writeSidecar('calc-estimated')
    const unverifiedGeomRes = verifyCfCase(makeCase('pixel-matched'), { rootDir: dir })
    expect(unverifiedGeomRes).toMatchObject({
      kind: 'refusal',
      code: 'geometry-mismatch',
    })

    // 3. OD-5 calibrated pixel-matched with excel-measured geometry verifies
    writeSidecar('excel-measured', 'FFFFC7CE', 'pixel-matched')
    // validPng has #ffc7ce at B2, target FFFFC7CE
    const cPixelMatched = makeCase('pixel-matched')
    cPixelMatched.expected.cells[0].fillRgb = 'FFFFC7CE'
    const pixelMatchedRes = verifyCfCase(cPixelMatched, { rootDir: dir })
    expect(pixelMatchedRes.kind).toBe('verified')

    // 4. Trusted ARGB (fillSource: 'excel-display-format') must agree with screen pixel (delta <= 12); refuses with fill-pixel-mismatch otherwise
    writeSidecar('excel-measured', 'FFFF0000')
    const trustedArgbCase = makeCase('excel-display-format')
    trustedArgbCase.expected.cells[0].fillRgb = 'FFFF0000'
    // validPng has #ffc7ce (delta > 12 from FFFF0000), so display-format and screen pixel mismatch
    const trustedRes = verifyCfCase(trustedArgbCase, { rootDir: dir })
    expect(trustedRes).toMatchObject({
      kind: 'refusal',
      code: 'fill-pixel-mismatch',
    })

    // 4b. When display-format ARGB agrees with screen pixel within 12, it verifies
    writeSidecar('excel-measured', 'FFFFC7CE')
    const agreedCase = makeCase('excel-display-format')
    agreedCase.expected.cells[0].fillRgb = 'FFFFC7CE'
    const agreedRes = verifyCfCase(agreedCase, { rootDir: dir })
    expect(agreedRes.kind).toBe('verified')

    // 5. Trusted ARGB that differs from expectation gets cell-expectation-mismatch
    writeSidecar('excel-measured', 'FFFFC7CE') // observed FFFFC7CE matches pixel, but differs from expected FFFF0000
    const mismatchArgbCase = makeCase('excel-display-format')
    mismatchArgbCase.expected.cells[0].fillRgb = 'FFFF0000'
    const mismatchRes = verifyCfCase(mismatchArgbCase, { rootDir: dir })
    expect(mismatchRes).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 6. OD-6: trusted ARGB + resave receipt with NO PNG promotes semantic capability unconditionally
    const od6NoPngCase = makeCase('excel-display-format')
    od6NoPngCase.captures = []
    const od6Res = verifyCfCase(od6NoPngCase, { rootDir: dir })
    expect(od6Res.kind).toBe('verified')

    // 7. Precondition probe failure refuses scale:* capabilities unless OD-5 or OD-6 engages
    const manifestFixture: CfManifest = {
      version: 1,
      excelBuild: '16.114',
      locale: 'zh-TW',
      dateSystem: '1900',
      hostTimeZone: 'Asia/Taipei',
      generatedAt: '2026-11-06T00:00:00Z',
      cases: [
        {
          id: 'probe-case',
          stage: 2,
          question: 'Precondition probe',
          input: 'scale.xlsx',
          saved: 'scale.xlsx',
          sheet: 'Sheet1',
          range: 'A1:B3',
          requiredAppearances: ['screen'],
          markers: [
            {
              part: 'xl/worksheets/sheet1.xml',
              extension: 'legacy',
              capabilityId: 'scale:interpolation-precondition-probe',
              sqref: 'B2:B3',
              ruleType: 'colorScale',
              priority: 1,
              formulas: [],
            },
          ],
          expected: {
            inputSha256: xlsxHash,
            savedSha256: xlsxHash,
            provenance: 'seed-hypothesis',
            positiveRefs: ['B2'],
            negativeRefs: ['B3'],
            cells: [{ ref: 'B2', controlRef: 'A2', matched: true }, { ref: 'B3', controlRef: 'A3', matched: false }],
          },
          captures: [],
          status: 'recipe-pending', // Refused!
          tolerance: { meanAbs: 0.015, rowMax: 0.08, strong: 12 },
        },
        od6NoPngCase, // Has OD-6 fallback
      ],
    }

    // OD-6 fallback engages -> capability verified despite probe failure
    expect(isCapabilityVerified('scale:fill-test', manifestFixture, { rootDir: dir })).toBe(true)

    // Without OD-5 or OD-6 fallback, capability is refused when probe fails
    const noFallbackCase = makeCase('unverified')
    manifestFixture.cases[1] = noFallbackCase
    expect(isCapabilityVerified('scale:fill-test', manifestFixture, { rootDir: dir })).toBe(false)
  })

  it('stage filtering in generate, build, and capture touches only the specified stage', async () => {
    // 1. filterCasesForStage enforces that --stage 2 returns strictly stage 2 buildable cases
    const stage2Cases = filterCasesForStage(CF_CASES as CfOracleCase[], 2)
    expect(stage2Cases.length).toBe(7)
    expect(stage2Cases.every(c => c.stage === 2)).toBe(true)
    expect(stage2Cases.some(c => c.stage === 1)).toBe(false)
    expect((CF_CASES as CfOracleCase[]).filter(c => c.stage === 2)).toHaveLength(8)

    // 2. filterCasesForStage with --stage 1 returns strictly stage 1 buildable cases
    const stage1Cases = filterCasesForStage(CF_CASES as CfOracleCase[], 1)
    expect(stage1Cases.length).toBe(12)
    expect(stage1Cases.every(c => c.stage === 1)).toBe(true)
    expect(stage1Cases.some(c => c.stage === 2)).toBe(false)
    expect((CF_CASES as CfOracleCase[]).filter(c => c.stage === 1)).toHaveLength(13)

    // 3. generateAllStage1 with stage: 2 preserves existing Stage 1 cases in manifest
    const dir = mkdtempSync(join(tmpdir(), 'cf-stage-filter-'))
    const fixtureDir = join(dir, 'tests/fixtures/xlsx-cf')
    mkdirSync(join(fixtureDir, 'inputs'), { recursive: true })
    mkdirSync(join(fixtureDir, 'refs'), { recursive: true })

    const fakeStage1Case: CfOracleCase = {
      ...(CF_CASES.find(c => c.id === 'stage1-cellis-literals') as CfOracleCase),
      expected: {
        ...(CF_CASES.find(c => c.id === 'stage1-cellis-literals') as CfOracleCase).expected,
        inputSha256: 'deadbeef11111111111111111111111111111111111111111111111111111111',
      },
    }
    const initialManifest: CfManifest = {
      version: 1,
      excelBuild: '16.114',
      locale: 'zh-TW',
      dateSystem: '1900',
      hostTimeZone: 'Asia/Taipei',
      generatedAt: '2026-11-06T00:00:00Z',
      cases: [fakeStage1Case],
    }
    writeFileSync(join(fixtureDir, 'manifest.json'), JSON.stringify(initialManifest, null, 2))

    // Generating stage: 2 must NOT overwrite the fake Stage 1 case
    const genManifest = await generateAllStage1({ rootDir: dir, stage: 2 })
    const retainedStage1 = genManifest.cases.find(c => c.id === 'stage1-cellis-literals')
    expect(retainedStage1?.expected.inputSha256).toBe('deadbeef11111111111111111111111111111111111111111111111111111111')
  })

  // osacompile ships with macOS only: skipped (by name) on Linux CI, strict wherever it exists.
  it.skipIf(!hasOsacompile)('compiles all probe AppleScripts under osacompile with zero syntax errors', () => {
    for (const c of CF_CASES) {
      const refs = c.expected.cells.map(x => x.ref)
      const scripts = [
        buildTextProbeScript(`${c.id}.xlsx`, c.sheet, refs),
        buildDecorationProbeScript(`${c.id}.xlsx`, c.sheet, refs),
        buildColorProbeScript(`${c.id}.xlsx`, c.sheet, refs),
        buildColorScaleCriteriaProbeScript(`${c.id}.xlsx`, c.sheet),
        ...(c.stage === 3 ? [buildDataBarSettingsProbeScript(`${c.id}.xlsx`, c.sheet)] : []),
      ]
      for (const lines of scripts) {
        const full = ['tell application "Microsoft Excel"', ...lines, 'end tell'].join('\n')
        const res = spawnSync('osacompile', ['-o', '/dev/null'], { input: full, encoding: 'utf8' })
        expect(res.status, `osacompile failed for ${c.id}: ${res.stderr}`).toBe(0)
      }
    }
  }, 120_000)

  it('correctly validates cell coordinates within range even with hidden rows (stage2-scale-scope)', () => {
    const scopeCase = CF_CASES.find(c => c.id === 'stage2-scale-scope')!
    const file = readFileSync(resolve(REPO_ROOT, scopeCase.input))
    const parts = readZipPartsSync(new Uint8Array(file))!
    const geom = computeWorksheetRangeGeometry(parts, scopeCase.range)!

    // B4 (row 4 hidden) and control B15 are both inside capture range A1:F20
    expect(isCellWithinRange('B4', geom)).toBe(true)
    expect(isCellWithinRange('B15', geom)).toBe(true)
    expect(isCellWithinRange('A1', geom)).toBe(true)
    expect(isCellWithinRange('F20', geom)).toBe(true)
    expect(isCellWithinRange('G20', geom)).toBe(false)
    expect(isCellWithinRange('F21', geom)).toBe(false)

    // verifyCfCase structural validation must NOT refuse stage2-scale-scope on B4/B15
    const res = verifyCfCase(scopeCase, { rootDir: REPO_ROOT })
    expect(res).toMatchObject({ structuralValid: true })
    if (res.kind === 'refusal') {
      expect(res.code).not.toBe('cell-expectation-mismatch')
    }
  })

  it('guards every cell and criterion query inside per-item try blocks in probe AppleScripts', () => {
    const testRefs = ['B2', 'B8']
    const colorScript = buildColorProbeScript('wb.xlsx', 'Sheet1', testRefs).join('\n')
    const decScript = buildDecorationProbeScript('wb.xlsx', 'Sheet1', testRefs).join('\n')
    const textScript = buildTextProbeScript('wb.xlsx', 'Sheet1', testRefs).join('\n')
    const critScript = buildColorScaleCriteriaProbeScript('wb.xlsx', 'Sheet1').join('\n')

    // Every cell query in color probe must be guarded by try ... end try with "missing value" fallback
    for (const ref of testRefs) {
      expect(colorScript).toMatch(new RegExp(`set dfColor to "missing value"[\\s\\S]*?try[\\s\\S]*?display format of range "${ref}"[\\s\\S]*?end try`))
      expect(colorScript).toMatch(new RegExp(`set baseColor to "missing value"[\\s\\S]*?try[\\s\\S]*?interior object of range "${ref}"[\\s\\S]*?end try`))
    }

    // Every cell query in decoration probe must be guarded by try ... end try with "missing value" fallback
    for (const ref of testRefs) {
      expect(decScript).toMatch(new RegExp(`set uVal to "missing value"[\\s\\S]*?try[\\s\\S]*?underline of font object of display format of range "${ref}"[\\s\\S]*?end try`))
      expect(decScript).toMatch(new RegExp(`set sVal to "missing value"[\\s\\S]*?try[\\s\\S]*?strikethrough of font object of display format of range "${ref}"[\\s\\S]*?end try`))
    }

    // Text probe must also guard string value query per cell
    for (const ref of testRefs) {
      expect(textScript).toMatch(new RegExp(`set tVal to ""[\\s\\S]*?try[\\s\\S]*?string value of range "${ref}"[\\s\\S]*?end try`))
    }

    // Criteria probe must guard criterion type query
    expect(critScript).toMatch(/set cType to "missing value"[\s\S]*?try[\s\S]*?color scale criterion type of crit[\s\S]*?end try/)
  })

  it('accepts Excel canonical forms for min/max cfvo val omission and alpha normalization', async () => {
    // 1. Thresholds case G2:G6 (type="min" val="0") matches Excel-resaved G2:G6 where val is dropped
    const thresholdsCase = CF_CASES.find(c => c.id === 'stage2-scale-thresholds')!
    const canonThresholdsXml =
      '<conditionalFormatting sqref="G2:G6"><cfRule type="colorScale" priority="6"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const dir = mkdtempSync(join(tmpdir(), 'cf-canon-'))
    const canonThresholdsBytes = await makeSyntheticXlsx({ sheetRuleXml: canonThresholdsXml, dxfCount: 0 })
    writeFileSync(join(dir, 'canon-thresholds.xlsx'), canonThresholdsBytes)
    const markerLegacy = thresholdsCase.markers.find(m => m.capabilityId === 'scale:thresholds-legacy-min-val-0')!
    const testCase = {
      ...thresholdsCase,
      markers: [markerLegacy],
      input: 'canon-thresholds.xlsx',
      saved: 'canon-thresholds.xlsx',
      expected: { ...thresholdsCase.expected, inputSha256: sha256(canonThresholdsBytes), savedSha256: sha256(canonThresholdsBytes) },
    }
    const resThresholds = verifyCfCase(testCase, { rootDir: dir })
    expect(resThresholds).toMatchObject({ structuralValid: true })

    // 2. Colors case D2:D6 (00F8696B) matches Excel-resaved D2:D6 where alpha is normalized to FFF8696B
    const colorsCase = CF_CASES.find(c => c.id === 'stage2-scale-colors')!
    const canonColorsXml =
      '<conditionalFormatting sqref="D2:D6"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const canonColorsBytes = await makeSyntheticXlsx({ sheetRuleXml: canonColorsXml, dxfCount: 0 })
    writeFileSync(join(dir, 'canon-colors.xlsx'), canonColorsBytes)
    const markerAlpha = colorsCase.markers.find(m => m.capabilityId === 'scale:colors-alpha-00')!
    const testColorsCase = {
      ...colorsCase,
      markers: [markerAlpha],
      input: 'canon-colors.xlsx',
      saved: 'canon-colors.xlsx',
      expected: { ...colorsCase.expected, inputSha256: sha256(canonColorsBytes), savedSha256: sha256(canonColorsBytes) },
    }
    const resColors = verifyCfCase(testColorsCase, { rootDir: dir })
    expect(resColors).toMatchObject({ structuralValid: true })

    // 3. Composition case H2:H5 (scale:composition-scale-stopIfTrue) matches Excel-resaved rule where stopIfTrue is dropped
    const compCase = CF_CASES.find(c => c.id === 'stage2-scale-composition')!
    const canonCompXml =
      '<conditionalFormatting sqref="H2:H5"><cfRule type="colorScale" priority="12"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const canonCompBytes = await makeSyntheticXlsx({ sheetRuleXml: canonCompXml, dxfCount: 0 })
    writeFileSync(join(dir, 'canon-comp.xlsx'), canonCompBytes)
    const markerStop = compCase.markers.find(m => m.capabilityId === 'scale:composition-scale-stopIfTrue')!
    const testCompCase = {
      ...compCase,
      markers: [markerStop],
      input: 'canon-comp.xlsx',
      saved: 'canon-comp.xlsx',
      expected: { ...compCase.expected, inputSha256: sha256(canonCompBytes), savedSha256: sha256(canonCompBytes) },
    }
    const resComp = verifyCfCase(testCompCase, { rootDir: dir })
    expect(resComp).toMatchObject({ structuralValid: true })
  })

  it('does not falsely refuse unmatched cells in colour-scale cases when controlRef is a scale cell', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-unmatched-scale-'))
    dirs.push(dir)

    const baseMarker: CfPartMarker = {
      part: 'xl/worksheets/sheet1.xml',
      extension: 'legacy',
      capabilityId: 'scale:unmatched-test',
      sqref: 'B2:B3',
      ruleType: 'colorScale',
      priority: 1,
      formulas: [],
      scale: {
        cfvo: [{ type: 'min' }, { type: 'max' }],
        colors: ['FFFF0000', 'FF00FF00'],
      },
    }

    const scaleXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>'
    const xlsxBytes = await makeSyntheticXlsx({ sheetRuleXml: scaleXml, dxfCount: 0 })
    const xlsxHash = sha256(xlsxBytes)
    writeFileSync(join(dir, 'scale.xlsx'), xlsxBytes)

    const validPng = makeCapturePng(true)
    const validPngHash = sha256(validPng)
    writeFileSync(join(dir, 'scale.png'), validPng)

    writeFileSync(
      join(dir, 'scale.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        width: 168,
        height: 54,
        colWidthsPt: [84, 84],
        rowHeightsPt: [18, 18, 18],
        geometrySource: 'excel-measured',
        resavedByExcel: true,
        resaveSavedSha256: xlsxHash,
        sheet: 'Sheet1',
        range: 'A1:B3',
        savedSha256: xlsxHash,
        pngSha256: validPngHash,
        capturedAt: '2026-11-06T08:00:00Z',
        observedCells: [
          { ref: 'B2', fillSource: 'excel-display-format', fillRgb: 'FFFFC7CE' },
          { ref: 'B3', fillSource: 'excel-display-format', fillRgb: 'FFFFFFFF' },
        ],
      }),
    )

    // Unmatched cell B3 has controlRef pointing to scale cell B2.
    // The verifier must validate B3 against default background without refusing due to contrast against B2.
    const testCase: CfOracleCase = {
      id: 'scale-unmatched-case',
      stage: 2,
      question: 'Scale unmatched cell control comparison',
      input: 'scale.xlsx',
      saved: 'scale.xlsx',
      sheet: 'Sheet1',
      range: 'A1:B3',
      requiredAppearances: ['screen'],
      markers: [baseMarker],
      expected: {
        inputSha256: xlsxHash,
        savedSha256: xlsxHash,
        provenance: 'excel-observed',
        positiveRefs: ['B2'],
        negativeRefs: ['B3'],
        cells: [
          { ref: 'B2', controlRef: 'A2', matched: true, fillRgb: 'FFFFC7CE', scaleApplied: true, fillSource: 'excel-display-format' },
          { ref: 'B3', controlRef: 'B2', matched: false, scaleApplied: false, fillSource: 'unverified' },
        ],
      },
      resavedByExcel: true,
      resaveReceipt: {
        openedRepairFree: true,
        excelVersion: '16.114',
        inputSha256: xlsxHash,
        savedSha256: xlsxHash,
        resavedAt: '2026-11-06T08:00:00Z',
      },
      captures: [
        {
          appearance: 'screen',
          pngPath: 'scale.png',
          sidecarPath: 'scale.json',
          pngSha256: validPngHash,
        },
      ],
      status: 'measured',
      tolerance: { meanAbs: 0.015, rowMax: 0.08, strong: 12 },
    }

    const verified = verifyCfCase(testCase, { rootDir: dir })
    expect(verified.kind).toBe('verified')
  })

  it('calibrates parseExcelColorToArgb against DXF_RED FFFFC7CE across all return shapes', () => {
    const DXF_RED = 'FFFFC7CE'

    // 1. AppleScript 16-bit RGBColor record / list {65535, 51118, 52942}
    expect(parseExcelColorToArgb([65535, 51118, 52942])).toBe(DXF_RED)
    expect(parseExcelColorToArgb('{65535, 51118, 52942}')).toBe(DXF_RED)
    expect(parseExcelColorToArgb('65535, 51118, 52942')).toBe(DXF_RED)

    // 2. 8-bit RGB [255, 199, 206]
    expect(parseExcelColorToArgb([255, 199, 206])).toBe(DXF_RED)
    expect(parseExcelColorToArgb('{255, 199, 206}')).toBe(DXF_RED)
    expect(parseExcelColorToArgb('255, 199, 206')).toBe(DXF_RED)

    // 3. Classic VBA BGR integer (0xCE_C7_FF = 13551615)
    expect(parseExcelColorToArgb(13551615)).toBe(DXF_RED)
    expect(parseExcelColorToArgb('13551615')).toBe(DXF_RED)

    // 4. Hex string #FFC7CE or FFFFC7CE
    expect(parseExcelColorToArgb('#FFC7CE')).toBe(DXF_RED)
    expect(parseExcelColorToArgb('FFFFC7CE')).toBe(DXF_RED)
    expect(parseExcelColorToArgb('ffc7ce')).toBe(DXF_RED)

    // 5. Extremes and primary colors
    expect(parseExcelColorToArgb([65535, 0, 0])).toBe('FFFF0000')
    expect(parseExcelColorToArgb([0, 65535, 0])).toBe('FF00FF00')
    expect(parseExcelColorToArgb([0, 0, 65535])).toBe('FF0000FF')
    expect(parseExcelColorToArgb([65535, 65535, 65535])).toBe('FFFFFFFF')
    expect(parseExcelColorToArgb([0, 0, 0])).toBe('FF000000')

    // 6. Null / none / empty / invalid shapes
    expect(parseExcelColorToArgb(null)).toBeNull()
    expect(parseExcelColorToArgb(undefined)).toBeNull()
    expect(parseExcelColorToArgb('none')).toBeNull()
    expect(parseExcelColorToArgb('')).toBeNull()
    expect(parseExcelColorToArgb('missing value')).toBeNull()
  })

  it('decodes Excel standard red F8696B and green 63BE7B from raw AppleScript return shapes', () => {
    const EXCEL_RED = 'FFF8696B'
    const EXCEL_GREEN = 'FF63BE7B'

    // Comma-delimited list string from AppleScript with text item delimiters set to ","
    expect(parseExcelColorToArgb('248, 105, 107')).toBe(EXCEL_RED)
    expect(parseExcelColorToArgb('248,105,107')).toBe(EXCEL_RED)
    expect(parseExcelColorToArgb('{248, 105, 107}')).toBe(EXCEL_RED)
    expect(parseExcelColorToArgb([248, 105, 107])).toBe(EXCEL_RED)
    expect(parseExcelColorToArgb('99, 190, 123')).toBe(EXCEL_GREEN)

    // Undelimited channel strings are ambiguous and must be refused as unparseable null
    expect(parseExcelColorToArgb('248105107')).toBeNull()
    expect(parseExcelColorToArgb('99190123')).toBeNull()
  })

  it('refuses with no-cells-checked when zero colour cells are verified', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-no-cells-'))
    dirs.push(dir)

    const baseMarker: CfPartMarker = {
      part: 'xl/worksheets/sheet1.xml',
      extension: 'legacy',
      capabilityId: 'scale:no-cells-test',
      sqref: 'B2:B3',
      ruleType: 'colorScale',
      priority: 1,
      formulas: [],
      scale: {
        cfvo: [{ type: 'min' }, { type: 'max' }],
        colors: ['FFFF0000', 'FF00FF00'],
      },
    }

    const scaleXml =
      '<conditionalFormatting sqref="B2:B3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>'
    const xlsxBytes = await makeSyntheticXlsx({ sheetRuleXml: scaleXml, dxfCount: 0 })
    const xlsxHash = sha256(xlsxBytes)
    writeFileSync(join(dir, 'scale.xlsx'), xlsxBytes)

    // Unmatched cells B2, B3 must not be highlighted so they pass background checks
    const validPng = makeCapturePng(false)
    const validPngHash = sha256(validPng)
    writeFileSync(join(dir, 'scale.png'), validPng)

    writeFileSync(
      join(dir, 'scale.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        width: 168,
        height: 54,
        colWidthsPt: [84, 84],
        rowHeightsPt: [18, 18, 18],
        geometrySource: 'excel-measured',
        resavedByExcel: true,
        resaveSavedSha256: xlsxHash,
        sheet: 'Sheet1',
        range: 'A1:B3',
        savedSha256: xlsxHash,
        pngSha256: validPngHash,
        capturedAt: '2026-11-06T08:00:00Z',
        observedCells: [],
      }),
    )

    // A case where cells has no matched scale cells that undergo color comparison
    const testCase: CfOracleCase = {
      id: 'scale-no-cells-case',
      stage: 2,
      question: 'Scale zero checked cells test',
      input: 'scale.xlsx',
      saved: 'scale.xlsx',
      sheet: 'Sheet1',
      range: 'A1:B3',
      requiredAppearances: ['screen'],
      markers: [baseMarker],
      expected: {
        inputSha256: xlsxHash,
        savedSha256: xlsxHash,
        provenance: 'excel-observed',
        positiveRefs: [],
        negativeRefs: ['B2', 'B3'],
        cells: [
          { ref: 'B2', controlRef: 'A2', matched: false, scaleApplied: false, fillSource: 'unverified' },
          { ref: 'B3', controlRef: 'A3', matched: false, scaleApplied: false, fillSource: 'unverified' },
        ],
      },
      resavedByExcel: true,
      resaveReceipt: {
        openedRepairFree: true,
        excelVersion: '16.114',
        inputSha256: xlsxHash,
        savedSha256: xlsxHash,
        resavedAt: '2026-11-06T08:00:00Z',
      },
      captures: [
        {
          appearance: 'screen',
          pngPath: 'scale.png',
          sidecarPath: 'scale.json',
          pngSha256: validPngHash,
        },
      ],
      status: 'measured',
      tolerance: { meanAbs: 0.015, rowMax: 0.08, strong: 12 },
    }

    const res = verifyCfCase(testCase, { rootDir: dir })
    expect(res).toMatchObject({
      kind: 'refusal',
      code: 'no-cells-checked',
    })
  })

  it('skips fill-vs-pixel comparison for hidden cells and marks them hidden: true without counting towards checked cells', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-hidden-cell-'))
    dirs.push(dir)

    const baseMarker: CfPartMarker = {
      part: 'xl/worksheets/sheet1.xml',
      extension: 'legacy',
      capabilityId: 'scale:hidden-cell-test',
      sqref: 'B2:B4',
      ruleType: 'colorScale',
      priority: 1,
      formulas: [],
      scale: {
        cfvo: [{ type: 'min' }, { type: 'max' }],
        colors: ['FFFFC7CE', 'FF63BE7B'],
      },
    }

    const scaleXml =
      '<conditionalFormatting sqref="B2:B4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFC7CE"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const xlsxBytes = await makeSyntheticXlsx({ sheetRuleXml: scaleXml, dxfCount: 0 })
    const xlsxHash = sha256(xlsxBytes)
    writeFileSync(join(dir, 'scale.xlsx'), xlsxBytes)

    // B2 is rendered red (#ffc7ce); B4 is hidden (row height 0) so it has no pixels
    const validPng = makeCapturePng(true)
    const validPngHash = sha256(validPng)
    writeFileSync(join(dir, 'scale.png'), validPng)

    writeFileSync(
      join(dir, 'scale.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        width: 168,
        height: 54,
        colWidthsPt: [84, 84],
        rowHeightsPt: [18, 18, 18, 0], // row 4 has height 0 (hidden)
        geometrySource: 'excel-measured',
        resavedByExcel: true,
        resaveSavedSha256: xlsxHash,
        sheet: 'Sheet1',
        range: 'A1:B4',
        savedSha256: xlsxHash,
        pngSha256: validPngHash,
        capturedAt: '2026-11-06T08:00:00Z',
        observedCells: [
          { ref: 'B2', fillSource: 'excel-display-format', fillRgb: 'FFFFC7CE' },
          { ref: 'B4', fillSource: 'excel-display-format', fillRgb: 'FF63BE7B' }, // disagrees with white pixel
        ],
      }),
    )

    const testCase: CfOracleCase = {
      id: 'scale-hidden-cell-case',
      stage: 2,
      question: 'Scale hidden cell handling',
      input: 'scale.xlsx',
      saved: 'scale.xlsx',
      sheet: 'Sheet1',
      range: 'A1:B4',
      requiredAppearances: ['screen'],
      markers: [baseMarker],
      expected: {
        inputSha256: xlsxHash,
        savedSha256: xlsxHash,
        provenance: 'excel-observed',
        positiveRefs: ['B2', 'B4'],
        negativeRefs: [],
        cells: [
          { ref: 'B2', controlRef: 'A2', matched: true, scaleApplied: true, fillSource: 'excel-display-format' },
          { ref: 'B4', controlRef: 'A4', matched: true, scaleApplied: true, fillSource: 'excel-display-format' },
        ],
      },
      resavedByExcel: true,
      resaveReceipt: {
        openedRepairFree: true,
        excelVersion: '16.114',
        inputSha256: xlsxHash,
        savedSha256: xlsxHash,
        resavedAt: '2026-11-06T08:00:00Z',
      },
      captures: [
        {
          appearance: 'screen',
          pngPath: 'scale.png',
          sidecarPath: 'scale.json',
          pngSha256: validPngHash,
        },
      ],
      status: 'measured',
      tolerance: { meanAbs: 0.015, rowMax: 0.08, strong: 12 },
    }

    // 1. Verification succeeds because B4's fill-vs-pixel mismatch is skipped
    const res = verifyCfCase(testCase, { rootDir: dir })
    expect(res.kind).toBe('verified')

    // 2. If all scale cells are hidden (e.g. only B4), verifyCfCase must refuse with no-cells-checked
    const allHiddenCase: CfOracleCase = {
      ...testCase,
      expected: {
        ...testCase.expected,
        positiveRefs: ['B4'],
        cells: [
          { ref: 'B4', controlRef: 'A4', matched: true, scaleApplied: true, fillSource: 'excel-display-format' },
        ],
      },
    }
    const allHiddenRes = verifyCfCase(allHiddenCase, { rootDir: dir })
    expect(allHiddenRes).toMatchObject({
      kind: 'refusal',
      code: 'no-cells-checked',
    })
  })

  it('preserves seed hypothesis cell expectations without circular copy of observed display-format fill', () => {
    // Stage 2 scale cases must retain their authored expected cells without circular copy of observed fillRgb
    const s2Cases = CF_CASES.filter(c => c.stage === 2 && c.markers.some(m => m.ruleType === 'colorScale'))
    expect(s2Cases.length).toBeGreaterThan(0)
    for (const c of s2Cases) {
      // Authored seed hypothesis expectations for scale cells must never pre-populate observed fillRgb
      for (const cell of c.expected.cells) {
        if (cell.matched && cell.scaleApplied) {
          expect(cell.fillRgb, `Case ${c.id} scale cell ${cell.ref} has circular or authored fillRgb`).toBeUndefined()
        }
      }
    }
  })

  it('expects whole-column scale cells F13..F17 in stage2-scale-scope as matched against unformatted column E', () => {
    const scopeCase = CF_CASES.find(c => c.id === 'stage2-scale-scope')!
    expect(scopeCase).toBeDefined()

    const wholeColRefs = ['F13', 'F14', 'F15', 'F16', 'F17']
    for (const ref of wholeColRefs) {
      expect(scopeCase.expected.positiveRefs).toContain(ref)
      expect(scopeCase.expected.negativeRefs).not.toContain(ref)
    }

    for (let r = 2; r <= 6; r++) {
      const fCell = scopeCase.expected.cells.find(cell => cell.ref === `F${r}`)
      expect(fCell).toBeDefined()
      expect(fCell!.matched).toBe(true)
      expect(fCell!.scaleApplied).toBe(true)
      expect(fCell!.controlRef).toBe(`E${r}`)
    }

    for (let r = 13; r <= 17; r++) {
      const fCell = scopeCase.expected.cells.find(cell => cell.ref === `F${r}`)
      expect(fCell).toBeDefined()
      expect(fCell!.matched).toBe(true)
      expect(fCell!.scaleApplied).toBe(true)
      expect(fCell!.controlRef).toBe(`E${r}`)
    }
  })

  it('rejects review-t1-r1 counterexamples: text-only capture, inverted/blue expectations, white printer capture, unbound sidecar, and threshold 100->1000', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-r1-'))
    dirs.push(dir)
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const c0 = structuredClone(manifest.cases[0])
    const savedBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, c0.saved)))
    const savedHash = sha256(savedBytes)

    // First verify that a properly sized, per-cell matching synthetic capture for c0 (with provenance='excel-observed') passes
    const validMatchingPng = makeCaseMatchingPng(c0, savedBytes, true)
    const validMatchingHash = sha256(validMatchingPng)
    writeFileSync(join(dir, 'matching.png'), validMatchingPng)
    const c0Geom = computeWorksheetRangeGeometry(readZipPartsSync(savedBytes)!, c0.range)!
    const baseSide = {
      sheet: c0.sheet,
      range: c0.range,
      pxPerPt: 1,
      scale: 1,
      route: 'copy-picture',
      appearance: 'screen',
      excelVersion: '16.114',
      geometrySource: 'excel-measured',
      colWidthsPt: c0Geom.colWidthsPt,
      rowHeightsPt: c0Geom.rowHeightsPt,
      resavedByExcel: true,
      resaveSavedSha256: savedHash,
      savedSha256: savedHash,
      pngSha256: validMatchingHash,
      capturedAt: '2026-11-06T08:00:00Z',
    }
    writeFileSync(join(dir, 'matching.json'), JSON.stringify(baseSide))

    // Without provenance='excel-observed', promoting seed-hypothesis to measured is refused
    expect(
      verifyCfCase(
        {
          ...c0,
          status: 'measured',
          expected: {
            ...c0.expected,
            provenance: 'seed-hypothesis',
          },
          resavedByExcel: true,
          resaveReceipt: {
            openedRepairFree: true,
            excelVersion: '16.114',
            inputSha256: c0.expected.inputSha256,
            savedSha256: savedHash,
            resavedAt: '2026-11-06T08:00:00Z',
          },
          captures: [
            {
              appearance: 'screen',
              pngPath: join(dir, 'matching.png'),
              sidecarPath: join(dir, 'matching.json'),
              pngSha256: validMatchingHash,
            },
          ],
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'unobserved-expectation',
    })

    const observedC0: CfOracleCase = {
      ...c0,
      status: 'measured',
      resavedByExcel: true,
      resaveReceipt: {
        openedRepairFree: true,
        excelVersion: '16.114',
        inputSha256: c0.expected.inputSha256,
        savedSha256: savedHash,
        resavedAt: '2026-11-06T08:00:00Z',
      },
      expected: { ...c0.expected, provenance: 'excel-observed' },
      captures: [
        {
          appearance: 'screen',
          pngPath: join(dir, 'matching.png'),
          sidecarPath: join(dir, 'matching.json'),
          pngSha256: validMatchingHash,
        },
      ],
    }
    expect(verifyCfCase(observedC0, { rootDir: REPO_ROOT }).kind).toBe('verified')

    // Counterexample 2a: White image with only black "101" text (no CF styling), even at exact range dimensions
    const parts = readZipPartsSync(savedBytes)!
    const geom = computeWorksheetRangeGeometry(parts, c0.range)!
    const exactW = geom.expectedWidthPt
    const exactH = geom.expectedHeightPt
    const textCanvas = createCanvas(exactW, exactH)
    const ctx = textCanvas.getContext('2d')
    ctx.fillStyle = 'white'
    ctx.fillRect(0, 0, exactW, exactH)
    ctx.fillStyle = 'black'
    ctx.font = '14px Arial'
    ctx.fillText('101', 80, 100)
    const textOnlyPng = new Uint8Array(textCanvas.toBuffer('image/png'))
    const textOnlyHash = sha256(textOnlyPng)
    writeFileSync(join(dir, 'text.png'), textOnlyPng)
    writeFileSync(join(dir, 'text.json'), JSON.stringify({ ...baseSide, pngSha256: textOnlyHash }))

    const textOnlyRes = verifyCfCase(
      {
        ...observedC0,
        captures: [
          {
            appearance: 'screen',
            pngPath: join(dir, 'text.png'),
            sidecarPath: join(dir, 'text.json'),
            pngSha256: textOnlyHash,
          },
        ],
      },
      { rootDir: REPO_ROOT },
    )
    expect(textOnlyRes).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // Counterexample 2b: Inverted expected matches and blue fillRgb against valid capture
    const invertedCase: CfOracleCase = {
      ...observedC0,
      expected: {
        ...observedC0.expected,
        cells: observedC0.expected.cells.map(x => ({ ...x, matched: !x.matched, fillRgb: 'FF0000FF' })),
      },
    }
    expect(verifyCfCase(invertedCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // Counterexample 2c: Valid screen capture + all-white printer capture
    const whitePng = makeCaseMatchingPng(c0, savedBytes, false)
    const whiteHash = sha256(whitePng)
    writeFileSync(join(dir, 'white-printer.png'), whitePng)
    writeFileSync(
      join(dir, 'white-printer.json'),
      JSON.stringify({ ...baseSide, appearance: 'printer', pngSha256: whiteHash }),
    )
    const whitePrinterRes = verifyCfCase(
      {
        ...observedC0,
        captures: [
          ...observedC0.captures,
          {
            appearance: 'printer',
            pngPath: join(dir, 'white-printer.png'),
            sidecarPath: join(dir, 'white-printer.json'),
            pngSha256: whiteHash,
          },
        ],
      },
      { rootDir: REPO_ROOT },
    )
    expect(whitePrinterRes.kind).toBe('refusal')
    if (whitePrinterRes.kind === 'refusal') {
      expect(['cell-expectation-mismatch', 'unchanged-control-capture']).toContain(whitePrinterRes.code)
    }

    // Counterexample 4a: Unbound sidecar with excelVersion 16.1 and missing hashes/route/scale/timestamp
    writeFileSync(
      join(dir, 'unbound-side.json'),
      JSON.stringify({ sheet: c0.sheet, range: c0.range, pxPerPt: 1, excelVersion: '16.1' }),
    )
    expect(
      verifyCfCase(
        {
          ...observedC0,
          captures: [
            {
              appearance: 'screen',
              pngPath: join(dir, 'matching.png'),
              sidecarPath: join(dir, 'unbound-side.json'),
              pngSha256: validMatchingHash,
            },
          ],
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'sidecar-mismatch',
    })

    // Counterexample 4b: Wrong capture dimensions (420x200 for A1:O26)
    const wrongDimCanvas = createCanvas(420, 200)
    const wrongDimPng = new Uint8Array(wrongDimCanvas.toBuffer('image/png'))
    const wrongDimHash = sha256(wrongDimPng)
    writeFileSync(join(dir, 'wrong-dim.png'), wrongDimPng)
    writeFileSync(join(dir, 'wrong-dim.json'), JSON.stringify({ ...baseSide, pngSha256: wrongDimHash }))
    expect(
      verifyCfCase(
        {
          ...observedC0,
          captures: [
            {
              appearance: 'screen',
              pngPath: join(dir, 'wrong-dim.png'),
              sidecarPath: join(dir, 'wrong-dim.json'),
              pngSha256: wrongDimHash,
            },
          ],
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'geometry-mismatch',
    })

    // Counterexample 3a: Structural formula threshold changed from <formula>100</formula> to <formula>1000</formula>
    const z = await JSZip.loadAsync(savedBytes)
    const sheetXml = await z.file('xl/worksheets/sheet1.xml')!.async('string')
    z.file('xl/worksheets/sheet1.xml', sheetXml.replace('<formula>100</formula>', '<formula>1000</formula>'))
    const mutatedBytes = await z.generateAsync({ type: 'uint8array' })
    const mutatedPath = join(dir, 'mutated-1000.xlsx')
    writeFileSync(mutatedPath, mutatedBytes)
    expect(
      verifyCfCase(
        {
          ...c0,
          saved: mutatedPath,
          status: 'recipe-pending',
          expected: { ...c0.expected, savedSha256: sha256(mutatedBytes) },
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'stripped-rule',
      structuralValid: false,
    })

    // Counterexample 3b: Mutating between's second operand <formula>3</formula><formula>1</formula> -> <formula>3</formula><formula>9</formula>
    const zBetween = await JSZip.loadAsync(savedBytes)
    zBetween.file(
      'xl/worksheets/sheet1.xml',
      sheetXml.replace('<formula>3</formula><formula>1</formula>', '<formula>3</formula><formula>9</formula>'),
    )
    const mutatedBetweenBytes = await zBetween.generateAsync({ type: 'uint8array' })
    const mutatedBetweenPath = join(dir, 'mutated-between.xlsx')
    writeFileSync(mutatedBetweenPath, mutatedBetweenBytes)
    expect(
      verifyCfCase(
        {
          ...c0,
          saved: mutatedBetweenPath,
          status: 'recipe-pending',
          expected: { ...c0.expected, savedSha256: sha256(mutatedBetweenBytes) },
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'stripped-rule',
      structuralValid: false,
    })

    // Counterexample 3c: Adding a formula to an absent-formula rule (formulas: [])
    const textCase = structuredClone(manifest.cases.find(c => c.id === 'stage1-text-predicates')!)
    const textSavedBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, textCase.saved)))
    const zText = await JSZip.loadAsync(textSavedBytes)
    const textSheetXml = await zText.file('xl/worksheets/sheet1.xml')!.async('string')
    zText.file(
      'xl/worksheets/sheet1.xml',
      textSheetXml.replace(
        '<cfRule type="containsText" dxfId="1" priority="8" operator="containsText" text="ab"/>',
        '<cfRule type="containsText" dxfId="1" priority="8" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH("ab",I2)))</formula></cfRule>',
      ),
    )
    const mutatedAbsentBytes = await zText.generateAsync({ type: 'uint8array' })
    const mutatedAbsentPath = join(dir, 'mutated-absent-formula.xlsx')
    writeFileSync(mutatedAbsentPath, mutatedAbsentBytes)
    expect(
      verifyCfCase(
        {
          ...textCase,
          saved: mutatedAbsentPath,
          status: 'recipe-pending',
          expected: { ...textCase.expected, savedSha256: sha256(mutatedAbsentBytes) },
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'stripped-rule',
      structuralValid: false,
    })
  })

  it('covers the full prescribed Stage 1 matrix, clock window enforcement, atomic fail-closed docs, and non-destructive single-case rebuild', async () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest

    // 1. Verify full matrix coverage across CF_STAGE1_CASES & CF_STAGE1_CAPABILITIES
    const cellIsCase = manifest.cases.find(c => c.id === 'stage1-cellis-literals')!
    const cellIsOps = new Set(cellIsCase.markers.map(m => m.operator))
    for (const op of [
      'equal',
      'notEqual',
      'greaterThan',
      'lessThan',
      'greaterThanOrEqual',
      'lessThanOrEqual',
      'between',
      'notBetween',
    ] as const) {
      expect(cellIsOps.has(op), `missing cellIs operator ${op}`).toBe(true)
    }

    const textCase = manifest.cases.find(c => c.id === 'stage1-text-predicates')!
    const textTypes = new Set(textCase.markers.map(m => m.ruleType))
    for (const t of ['containsText', 'notContainsText', 'beginsWith', 'endsWith'] as const) {
      expect(textTypes.has(t), `missing text ruleType ${t}`).toBe(true)
    }
    expect(textCase.markers.some(m => m.formulas.length === 0)).toBe(true)

    const blanksCase = manifest.cases.find(c => c.id === 'stage1-blanks-errors')!
    const blankTypes = new Set(blanksCase.markers.map(m => m.ruleType))
    for (const t of ['containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'] as const) {
      expect(blankTypes.has(t), `missing blanks/errors ruleType ${t}`).toBe(true)
    }

    const timeCases = manifest.cases.filter(c => c.id.startsWith('stage1-timeperiod-today-clock'))
    expect(timeCases).toHaveLength(2)
    expect(timeCases.map(c => c.clockRequirement?.dateSystem)).toEqual(['1900', '1904'])
    const timePeriods = new Set(timeCases.flatMap(c => c.markers.map(m => m.timePeriod)))
    for (const tp of [
      'today',
      'yesterday',
      'tomorrow',
      'last7Days',
      'thisWeek',
      'lastWeek',
      'nextWeek',
      'thisMonth',
      'lastMonth',
      'nextMonth',
    ] as const) {
      expect(timePeriods.has(tp), `missing timePeriod ${tp}`).toBe(true)
    }

    const top10Case = manifest.cases.find(c => c.id === 'stage1-top10-rank-edge')!
    expect(top10Case.markers.some(m => m.rank === 'omitted')).toBe(true)
    expect(top10Case.markers.some(m => m.rank === 0 && !m.percent)).toBe(true)
    expect(top10Case.markers.some(m => m.percent && m.rank === 0)).toBe(true)
    expect(top10Case.markers.some(m => m.percent && m.rank === 1)).toBe(true)
    expect(top10Case.markers.some(m => m.percent && m.rank === 40)).toBe(true)
    expect(top10Case.markers.some(m => m.percent && m.rank === 100)).toBe(true)

    const avgCase = manifest.cases.find(c => c.id === 'stage1-above-average-stddev')!
    for (const sd of [1, 2, 3, 4]) {
      expect(avgCase.markers.some(m => m.stdDev === sd), `missing stdDev=${sd}`).toBe(true)
    }
    expect(avgCase.markers.some(m => m.capabilityId === 'aboveAverage:stdDev-0')).toBe(true)
    expect(avgCase.markers.some(m => m.capabilityId === 'aboveAverage:stdDev-negative')).toBe(true)
    expect(avgCase.markers.some(m => m.capabilityId === 'aboveAverage:stdDev-with-equalAverage')).toBe(true)

    const dupCase = manifest.cases.find(c => c.id === 'stage1-duplicate-scope')!
    expect(dupCase.markers.some(m => m.sqref === 'D2:D4 F2:F4')).toBe(true)
    expect(dupCase.markers.some(m => m.sqref === 'H2:H4 H3:H5')).toBe(true)

    const prioCase = manifest.cases.find(c => c.id === 'stage1-priority-composition')!
    expect(prioCase.markers.filter(m => m.priority === 7)).toHaveLength(2)
    expect(prioCase.markers.some(m => m.ruleType === 'colorScale')).toBe(true)
    expect(prioCase.markers.some(m => m.ruleType === 'dataBar')).toBe(true)
    expect(prioCase.markers.some(m => m.ruleType === 'iconSet')).toBe(true)
    expect(prioCase.markers.some(m => m.ruleType === 'expression' && m.stopIfTrue === true)).toBe(true)

    const blankMergeCase = manifest.cases.find(c => c.id === 'stage1-blank-merge-order')!
    expect(blankMergeCase.markers.some(m => m.sqref === 'B2')).toBe(true)
    expect(blankMergeCase.markers.some(m => m.sqref === 'Z5')).toBe(true)
    expect(blankMergeCase.markers.some(m => m.sqref === 'M1:M1048576')).toBe(true)

    // Check capability verification status matches manifest case status:
    // 12 resaved cases are verified (110 capabilities verified); 1 unsupported (5 capabilities) is refused
    expect(CF_STAGE1_CAPABILITIES.length).toBeGreaterThanOrEqual(90)
    let verifiedCapCount = 0
    let refusedCapCount = 0
    for (const cap of CF_STAGE1_CAPABILITIES) {
      if (isCapabilityVerified(cap.capabilityId, manifest, { rootDir: REPO_ROOT })) {
        verifiedCapCount++
      } else {
        refusedCapCount++
      }
    }
    expect(verifiedCapCount).toBe(110)
    expect(refusedCapCount).toBe(5)

    // 5. Clock window and fixture expiration enforcement
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-clock-'))
    dirs.push(dir)
    const clockCase = structuredClone(manifest.cases.find(c => c.id === 'stage1-timeperiod-today-clock')!)
    const clockSavedBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, clockCase.saved)))
    const clockSavedHash = sha256(clockSavedBytes)
    const clockPng = makeCaseMatchingPng(clockCase, clockSavedBytes, true)
    const clockPngHash = sha256(clockPng)
    writeFileSync(join(dir, 'clock.png'), clockPng)
    const clockGeom = computeWorksheetRangeGeometry(readZipPartsSync(clockSavedBytes)!, clockCase.range)!
    const clockReceipt: CfResaveReceipt = {
      openedRepairFree: true,
      excelVersion: '16.114',
      inputSha256: clockCase.expected.inputSha256,
      savedSha256: clockSavedHash,
      resavedAt: '2026-11-06T08:00:00Z',
    }

    // 5a: Daytime capture (16:00 Asia/Taipei, outside 00:05-07:55 window) -> refused with clock-window-invalid
    writeFileSync(
      join(dir, 'clock-daytime.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        geometrySource: 'excel-measured',
        colWidthsPt: clockGeom.colWidthsPt,
        rowHeightsPt: clockGeom.rowHeightsPt,
        sheet: clockCase.sheet,
        range: clockCase.range,
        hostTimeZone: 'Asia/Taipei',
        resavedByExcel: true,
        resaveSavedSha256: clockSavedHash,
        savedSha256: clockSavedHash,
        pngSha256: clockPngHash,
        capturedAt: '2026-11-06T08:00:00Z',
        excelTodaySerial: 46332,
        staleTodayCacheRefreshed: true,
        observedTodayBranch: 'local',
      }),
    )
    expect(
      verifyCfCase(
        {
          ...clockCase,
          status: 'measured',
          resavedByExcel: true,
          resaveReceipt: clockReceipt,
          expected: { ...clockCase.expected, provenance: 'excel-observed' },
          captures: [
            {
              appearance: 'screen',
              pngPath: join(dir, 'clock.png'),
              sidecarPath: join(dir, 'clock-daytime.json'),
              pngSha256: clockPngHash,
            },
          ],
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'clock-window-invalid',
    })

    // 5b: Later day's 00:05-07:55 window (2026-11-08 01:30 Asia/Taipei = 2026-11-07T17:30:00Z) against 2026-11-06 fixture -> refused with clock-fixture-expired
    writeFileSync(
      join(dir, 'clock-expired.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        geometrySource: 'excel-measured',
        colWidthsPt: clockGeom.colWidthsPt,
        rowHeightsPt: clockGeom.rowHeightsPt,
        sheet: clockCase.sheet,
        range: clockCase.range,
        hostTimeZone: 'Asia/Taipei',
        resavedByExcel: true,
        resaveSavedSha256: clockSavedHash,
        savedSha256: clockSavedHash,
        pngSha256: clockPngHash,
        capturedAt: '2026-11-07T17:30:00Z',
        excelTodaySerial: 46334,
        staleTodayCacheRefreshed: true,
        observedTodayBranch: 'local',
      }),
    )
    expect(
      verifyCfCase(
        {
          ...clockCase,
          status: 'measured',
          resavedByExcel: true,
          resaveReceipt: clockReceipt,
          expected: { ...clockCase.expected, provenance: 'excel-observed' },
          captures: [
            {
              appearance: 'screen',
              pngPath: join(dir, 'clock.png'),
              sidecarPath: join(dir, 'clock-expired.json'),
              pngSha256: clockPngHash,
            },
          ],
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'clock-fixture-expired',
    })

    // 5c: Valid 00:05-07:55 Asia/Taipei capture on anchor date -> verified
    const anchorReq = clockCase.clockRequirement!
    const [aY, aM, aD] = anchorReq.anchorLocalDate.split('-').map(Number)
    const validTaipeiDate = new Date(Date.UTC(aY!, aM! - 1, aD!, 1, 30, 0))
    const validCapturedAt = new Date(validTaipeiDate.getTime() - 8 * 3600_000).toISOString()
    writeFileSync(
      join(dir, 'clock-valid.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        geometrySource: 'excel-measured',
        colWidthsPt: clockGeom.colWidthsPt,
        rowHeightsPt: clockGeom.rowHeightsPt,
        sheet: clockCase.sheet,
        range: clockCase.range,
        hostTimeZone: 'Asia/Taipei',
        resavedByExcel: true,
        resaveSavedSha256: clockSavedHash,
        savedSha256: clockSavedHash,
        pngSha256: clockPngHash,
        capturedAt: validCapturedAt,
        excelTodaySerial: anchorReq.anchorSerial,
        staleTodayCacheRefreshed: true,
        observedTodayBranch: 'local',
      }),
    )
    expect(
      verifyCfCase(
        {
          ...clockCase,
          status: 'measured',
          resavedByExcel: true,
          resaveReceipt: clockReceipt,
          expected: { ...clockCase.expected, provenance: 'excel-observed' },
          captures: [
            {
              appearance: 'screen',
              pngPath: join(dir, 'clock.png'),
              sidecarPath: join(dir, 'clock-valid.json'),
              pngSha256: clockPngHash,
            },
          ],
        },
        { rootDir: REPO_ROOT },
      ).kind,
    ).toBe('verified')

    // 6. Atomic fail-closed validation documentation check
    const validationDoc = readFileSync(resolve(REPO_ROOT, 'docs/validation/xlsx-cf.md'), 'utf8')
    expect(validationDoc).not.toContain('stopIfTrue ignored')
    expect(validationDoc).not.toContain('Refused rules never partially apply DXF properties or halt lower-priority rules')
    expect(validationDoc).toContain('known `not-matched`')
    expect(validationDoc).toContain('stop-tainted')
    expect(validationDoc).toContain('property-tainted')

    // 7. Non-destructive single-case rebuild: preserves Excel-resaved pending workbooks for other cases
    const tempRepo = mkdtempSync(join(tmpdir(), 'ov-cf-rebuild-'))
    dirs.push(tempRepo)
    cpSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf'), join(tempRepo, 'tests/fixtures/xlsx-cf'), { recursive: true })
    const savedCaseAPath = join(tempRepo, 'tests/fixtures/xlsx-cf/stage1-cellis-literals.xlsx')
    const zResaved = await JSZip.loadAsync(readFileSync(savedCaseAPath))
    zResaved.file('docProps/custom.xml', `${XML_DECL}<Properties/>`)
    const resavedBytes = await zResaved.generateAsync({ type: 'uint8array' })
    writeFileSync(savedCaseAPath, resavedBytes)
    const resavedHash = sha256(resavedBytes)

    const tempManifestPath = join(tempRepo, 'tests/fixtures/xlsx-cf/manifest.json')
    const tempManifest = JSON.parse(readFileSync(tempManifestPath, 'utf8')) as CfManifest
    tempManifest.cases[0] = {
      ...tempManifest.cases[0],
      resavedByExcel: true,
      status: 'recipe-pending',
      reason: 'Excel resaved; awaiting capture',
      expected: {
        ...tempManifest.cases[0].expected,
        savedSha256: resavedHash,
      },
    }
    writeFileSync(tempManifestPath, JSON.stringify(tempManifest, null, 2))

    // Rebuild only case B (stage1-text-predicates), then also run full generateAllStage1
    const afterSingle = await generateAllStage1({ rootDir: tempRepo, onlyCaseId: 'stage1-text-predicates' })
    expect(afterSingle.cases[0].expected.savedSha256).toBe(resavedHash)
    expect(afterSingle.cases[0].resavedByExcel).toBe(true)
    expect(sha256(new Uint8Array(readFileSync(savedCaseAPath)))).toBe(resavedHash)

    const afterFull = await generateAllStage1({ rootDir: tempRepo })
    expect(afterFull.cases[0].expected.savedSha256).toBe(resavedHash)
    expect(afterFull.cases[0].resavedByExcel).toBe(true)
    expect(sha256(new Uint8Array(readFileSync(savedCaseAPath)))).toBe(resavedHash)
  }, 30000)

  it('rejects review-t1-r2 counterexamples: fill-only/bar-only PNG without text/font/border, mutated non-fill expectations, unobserved capabilities, missing pivot parts, missing Excel resave provenance, proportional vs ECMA-376 column geometry, and Excel version 16.120..199', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-r2-'))
    dirs.push(dir)
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest

    // 1. Every marker in every capture-bearing Stage 1 case must have expected.cells observations inside its sqref
    for (const c of manifest.cases) {
      if (c.status === 'unsupported') continue
      for (const m of c.markers) {
        const covered = c.expected.cells.filter(cell => {
          const pos = parseA1(cell.ref)
          if (!pos) return false
          return m.sqref.trim().split(/\s+/).some(token => {
            const parts = token.split(':')
            const a = parseA1(parts[0] ?? '')
            const b = parseA1(parts[1] ?? parts[0] ?? '')
            if (!a || !b) return false
            return (
              pos.row >= Math.min(a.row, b.row) &&
              pos.row <= Math.max(a.row, b.row) &&
              pos.col >= Math.min(a.col, b.col) &&
              pos.col <= Math.max(a.col, b.col)
            )
          })
        })
        expect(covered.length, `marker ${m.capabilityId} (${m.sqref}) in ${c.id} has no expected.cells`).toBeGreaterThan(0)
      }
    }

    // 2. stage1-dxf-layering must contain real xl/pivotTables/pivotTable1.xml for dxf:pivot-style-layering
    const dxfLayering = manifest.cases.find(x => x.id === 'stage1-dxf-layering')!
    const dxfParts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, dxfLayering.saved))))!
    expect(dxfParts.has('xl/pivotTables/pivotTable1.xml')).toBe(true)

    // 3. r2-probe counterexample: solid-fill / black-bar PNG without text/font/border must be refused
    const cNumFmt = structuredClone(manifest.cases.find(x => x.id === 'stage1-numfmt-accessibility')!)
    const numFmtBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, cNumFmt.saved)))
    const numFmtHash = sha256(numFmtBytes)
    const numFmtGeomEarly = computeWorksheetRangeGeometry(readZipPartsSync(numFmtBytes)!, cNumFmt.range)!
    const fillBarOnlyPng = makeFillBarOnlyPng(cNumFmt, numFmtBytes)
    writeFileSync(join(dir, 'numfmt-fillbar.png'), fillBarOnlyPng)
    const fillBarHash = sha256(fillBarOnlyPng)
    const numFmtReceipt: CfResaveReceipt = {
      openedRepairFree: true,
      excelVersion: '16.114',
      inputSha256: cNumFmt.expected.inputSha256,
      savedSha256: numFmtHash,
      resavedAt: '2026-11-06T08:00:00Z',
    }
    const fillBarCase: CfOracleCase = {
      ...cNumFmt,
      status: 'measured',
      resavedByExcel: true,
      resaveReceipt: numFmtReceipt,
      expected: { ...cNumFmt.expected, provenance: 'excel-observed' },
      captures: (cNumFmt.requiredAppearances ?? ['screen']).map(appearance => {
        const side = join(dir, `numfmt.${appearance}.json`)
        writeFileSync(
          side,
          JSON.stringify({
            sheet: cNumFmt.sheet,
            range: cNumFmt.range,
            appearance,
            route: 'copy-picture',
            scale: 1,
            pxPerPt: 1,
            excelVersion: '16.114',
            geometrySource: 'excel-measured',
            colWidthsPt: numFmtGeomEarly.colWidthsPt,
            rowHeightsPt: numFmtGeomEarly.rowHeightsPt,
            resavedByExcel: true,
            resaveSavedSha256: numFmtHash,
            savedSha256: numFmtHash,
            pngSha256: fillBarHash,
            capturedAt: '2026-11-06T08:00:00Z',
          }),
        )
        return {
          appearance,
          pngPath: join(dir, 'numfmt-fillbar.png'),
          sidecarPath: side,
          pngSha256: fillBarHash,
        }
      }),
    }
    const fillBarResult = verifyCfCase(fillBarCase, { rootDir: REPO_ROOT })
    expect(fillBarResult).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 4. Producer-to-verifier path: sidecar observedCells generated by observeCellsFromCapture
    const validNumFmtPng = makeCaseMatchingPng(cNumFmt, numFmtBytes, true)
    const validNumFmtHash = sha256(validNumFmtPng)
    writeFileSync(join(dir, 'numfmt-valid.png'), validNumFmtPng)

    const numFmtParts = readZipPartsSync(numFmtBytes)!
    const geomNumFmt = computeWorksheetRangeGeometry(numFmtParts, cNumFmt.range)!
    const validTexts: Record<string, string> = {}
    for (const cell of cNumFmt.expected.cells) {
      if (cell.formattedText !== undefined) validTexts[cell.ref] = cell.formattedText
    }
    const validDecs: Record<string, { underline?: 'none' | 'single' | 'double'; strikethrough?: boolean }> = {}
    for (const cell of cNumFmt.expected.cells) {
      validDecs[cell.ref] = {
        underline: cell.underline ?? 'none',
        strikethrough: cell.strikethrough ?? false,
      }
    }
    const validObservedCells = observeCellsFromCapture(
      decodePngSync(validNumFmtPng)!,
      geomNumFmt,
      cNumFmt,
      { observedCellTexts: validTexts, observedCellDecorations: validDecs },
    )
    expect(validObservedCells).not.toBeNull()

    const validNumFmtCase: CfOracleCase = {
      ...cNumFmt,
      status: 'measured',
      resavedByExcel: true,
      resaveReceipt: numFmtReceipt,
      expected: { ...cNumFmt.expected, provenance: 'excel-observed' },
      captures: (cNumFmt.requiredAppearances ?? ['screen']).map(appearance => {
        const side = join(dir, `numfmt-valid.${appearance}.json`)
        writeFileSync(
          side,
          JSON.stringify({
            sheet: cNumFmt.sheet,
            range: cNumFmt.range,
            appearance,
            route: 'copy-picture',
            scale: 1,
            pxPerPt: 1,
            excelVersion: '16.114',
            geometrySource: 'excel-measured',
            colWidthsPt: geomNumFmt.colWidthsPt,
            rowHeightsPt: geomNumFmt.rowHeightsPt,
            resavedByExcel: true,
            resaveSavedSha256: numFmtHash,
            savedSha256: numFmtHash,
            pngSha256: validNumFmtHash,
            capturedAt: '2026-11-06T08:00:00Z',
            observedCells: validObservedCells,
          }),
        )
        return {
          appearance,
          pngPath: join(dir, 'numfmt-valid.png'),
          sidecarPath: side,
          pngSha256: validNumFmtHash,
        }
      }),
    }
    expect(verifyCfCase(validNumFmtCase, { rootDir: REPO_ROOT }).kind).toBe('verified')

    // 4a: Mutate font color to blue
    const blueFontCase: CfOracleCase = {
      ...validNumFmtCase,
      expected: {
        ...validNumFmtCase.expected,
        cells: validNumFmtCase.expected.cells.map(x => (x.fontRgb ? { ...x, fontRgb: 'FF0000FF' } : x)),
      },
    }
    expect(verifyCfCase(blueFontCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 4b: Mutate bold expectation to false
    const noBoldCase: CfOracleCase = {
      ...validNumFmtCase,
      expected: {
        ...validNumFmtCase.expected,
        cells: validNumFmtCase.expected.cells.map(x => (x.bold ? { ...x, bold: false } : x)),
      },
    }
    expect(verifyCfCase(noBoldCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 4c: Mutate formattedText to WRONG TEXT
    const wrongTextCase: CfOracleCase = {
      ...validNumFmtCase,
      expected: {
        ...validNumFmtCase.expected,
        cells: validNumFmtCase.expected.cells.map(x =>
          x.formattedText ? { ...x, formattedText: 'WRONG TEXT' } : x,
        ),
      },
    }
    expect(verifyCfCase(wrongTextCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 4d: Expect unexpected top border
    const unexpectedBorderCase: CfOracleCase = {
      ...validNumFmtCase,
      expected: {
        ...validNumFmtCase.expected,
        cells: validNumFmtCase.expected.cells.map(x => ({ ...x, borderTop: true })),
      },
    }
    expect(verifyCfCase(unexpectedBorderCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 4e: Raw observer output without observedCellTexts has no formattedText; must be refused
    const rawObserved = observeCellsFromCapture(decodePngSync(validNumFmtPng)!, geomNumFmt, cNumFmt)
    const rawSidePath = join(dir, 'numfmt-raw-observed.json')
    writeFileSync(
      rawSidePath,
      JSON.stringify({
        sheet: cNumFmt.sheet,
        range: cNumFmt.range,
        appearance: 'screen',
        route: 'copy-picture',
        scale: 1,
        pxPerPt: 1,
        excelVersion: '16.114',
        geometrySource: 'excel-measured',
        colWidthsPt: geomNumFmt.colWidthsPt,
        rowHeightsPt: geomNumFmt.rowHeightsPt,
        resavedByExcel: true,
        resaveSavedSha256: numFmtHash,
        savedSha256: numFmtHash,
        pngSha256: validNumFmtHash,
        capturedAt: '2026-11-06T08:00:00Z',
        observedCells: rawObserved,
      }),
    )
    const rawProducerCase: CfOracleCase = {
      ...validNumFmtCase,
      captures: validNumFmtCase.captures.map(cap =>
        cap.appearance === 'screen' ? { ...cap, sidecarPath: rawSidePath } : cap,
      ),
    }
    expect(verifyCfCase(rawProducerCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
      reason: expect.stringContaining('has no observed formattedText for it'),
    })

    // 4f: Absent underline and strikethrough added to ordinary text must be refused
    const addedDecorationsCase: CfOracleCase = {
      ...validNumFmtCase,
      expected: {
        ...validNumFmtCase.expected,
        cells: validNumFmtCase.expected.cells.map(x =>
          x.matched ? { ...x, underline: 'double', strikethrough: true } : x,
        ),
      },
    }
    expect(verifyCfCase(addedDecorationsCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 4g: Mutated underline expectation (single vs none) must be refused
    const underlineMismatchCase: CfOracleCase = {
      ...validNumFmtCase,
      expected: {
        ...validNumFmtCase.expected,
        cells: validNumFmtCase.expected.cells.map(x => (x.ref === 'B2' ? { ...x, underline: 'single' } : x)),
      },
    }
    expect(verifyCfCase(underlineMismatchCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 4h: Mutated strikethrough expectation (true vs false) must be refused
    const strikeMismatchCase: CfOracleCase = {
      ...validNumFmtCase,
      expected: {
        ...validNumFmtCase.expected,
        cells: validNumFmtCase.expected.cells.map(x => (x.ref === 'B2' ? { ...x, strikethrough: true } : x)),
      },
    }
    expect(verifyCfCase(strikeMismatchCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'cell-expectation-mismatch',
    })

    // 5. Excel resave provenance: must have resavedByExcel: true and valid resaveReceipt
    const unresavedCase: CfOracleCase = {
      ...validNumFmtCase,
      resavedByExcel: false,
      resaveReceipt: undefined,
    }
    expect(verifyCfCase(unresavedCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'resave-provenance-missing',
    })

    // 6. Excel version check accepts 16.120..199 and 16.200+, rejects < 16.114
    expect(isSupportedExcelVersion('16.114')).toBe(true)
    expect(isSupportedExcelVersion('16.120')).toBe(true)
    expect(isSupportedExcelVersion('16.199')).toBe(true)
    expect(isSupportedExcelVersion('16.200')).toBe(true)
    expect(isSupportedExcelVersion('17.0')).toBe(true)
    expect(isSupportedExcelVersion('16.113')).toBe(false)
    expect(isSupportedExcelVersion('16.1')).toBe(false)
    expect(isSupportedExcelVersion('15.0')).toBe(false)
    expect(isSupportedExcelVersion('invalid')).toBe(false)

    // 7. Column geometry: sidecar with mismatched colWidthsPt is refused
    const badGeomSidePath = join(dir, 'numfmt-bad-geom.json')
    writeFileSync(
      badGeomSidePath,
      JSON.stringify({
        sheet: cNumFmt.sheet,
        range: cNumFmt.range,
        appearance: 'screen',
        route: 'copy-picture',
        scale: 1,
        pxPerPt: 1,
        excelVersion: '16.114',
        geometrySource: 'excel-measured',
        resavedByExcel: true,
        resaveSavedSha256: numFmtHash,
        savedSha256: numFmtHash,
        pngSha256: validNumFmtHash,
        capturedAt: '2026-11-06T08:00:00Z',
        colWidthsPt: [100, 100, 100, 100, 100, 100, 100], // Mismatched widths (sum 700 !== 420)
        rowHeightsPt: geomNumFmt.rowHeightsPt,
      }),
    )
    const badGeomCase: CfOracleCase = {
      ...validNumFmtCase,
      captures: validNumFmtCase.captures.map(cap =>
        cap.appearance === 'screen' ? { ...cap, sidecarPath: badGeomSidePath } : cap,
      ),
    }
    expect(verifyCfCase(badGeomCase, { rootDir: REPO_ROOT })).toMatchObject({
      kind: 'refusal',
      code: 'geometry-mismatch',
    })

    // 8. Real producer measured geometry baseline with author recorded 522pt example:
    // Stored width 14 = 84pt, narrow width 3 = 18pt; cols = [84, 84, 84, 84, 84, 18, 84], sum = 522pt
    const measuredCols522 = [84, 84, 84, 84, 84, 18, 84]
    const measured522Png = makeCaseMatchingPng(cNumFmt, numFmtBytes, true, {
      colWidthsPt: measuredCols522,
    })
    const measured522Hash = sha256(measured522Png)
    const measured522PngPath = join(dir, 'measured-522.png')
    writeFileSync(measured522PngPath, measured522Png)

    const measured522Geom = computeWorksheetRangeGeometry(numFmtParts, cNumFmt.range, {
      colWidthsPt: measuredCols522,
    })!
    expect(measured522Geom.expectedWidthPt).toBe(522)

    const measured522ObsCells = observeCellsFromCapture(
      decodePngSync(measured522Png)!,
      measured522Geom,
      cNumFmt,
      {
        observedCellTexts: validTexts,
        observedCellDecorations: validDecs,
      },
    )!

    const measured522Case: CfOracleCase = {
      ...validNumFmtCase,
      captures: validNumFmtCase.captures.map(cap => {
        const sidePath = join(dir, `measured-522.${cap.appearance}.json`)
        writeFileSync(
          sidePath,
          JSON.stringify({
            sheet: cNumFmt.sheet,
            range: cNumFmt.range,
            appearance: cap.appearance,
            route: 'copy-picture',
            scale: 1,
            pxPerPt: 1,
            excelVersion: '16.114',
            geometrySource: 'excel-measured',
            resavedByExcel: true,
            resaveSavedSha256: numFmtHash,
            savedSha256: numFmtHash,
            pngSha256: measured522Hash,
            capturedAt: '2026-11-06T08:00:00Z',
            width: 522,
            height: measured522Geom.expectedHeightPt,
            colWidthsPt: measuredCols522,
            rowHeightsPt: measured522Geom.rowHeightsPt,
            observedCells: measured522ObsCells,
          }),
        )
        return {
          appearance: cap.appearance,
          pngPath: measured522PngPath,
          sidecarPath: sidePath,
          pngSha256: measured522Hash,
        }
      }),
    }
    const measured522Res = verifyCfCase(measured522Case, { rootDir: REPO_ROOT })
    expect(measured522Res).toMatchObject({
      kind: 'verified',
      id: 'stage1-numfmt-accessibility',
      status: 'measured',
    })

    // 8b. Mismatched column count in excel-measured sidecar is refused
    const badColCountSidePath = join(dir, 'measured-bad-col-count.json')
    writeFileSync(
      badColCountSidePath,
      JSON.stringify({
        sheet: cNumFmt.sheet,
        range: cNumFmt.range,
        appearance: 'screen',
        route: 'copy-picture',
        scale: 1,
        pxPerPt: 1,
        excelVersion: '16.114',
        geometrySource: 'excel-measured',
        resavedByExcel: true,
        resaveSavedSha256: numFmtHash,
        savedSha256: numFmtHash,
        pngSha256: measured522Hash,
        capturedAt: '2026-11-06T08:00:00Z',
        width: 522,
        height: measured522Geom.expectedHeightPt,
        colWidthsPt: [84, 84, 84, 84, 84, 18], // 6 columns instead of 7
        rowHeightsPt: measured522Geom.rowHeightsPt,
        observedCells: measured522ObsCells,
      }),
    )
    expect(
      verifyCfCase(
        {
          ...measured522Case,
          captures: measured522Case.captures.map(cap =>
            cap.appearance === 'screen' ? { ...cap, sidecarPath: badColCountSidePath } : cap,
          ),
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({
      kind: 'refusal',
      code: 'geometry-mismatch',
    })

    // 9. Valid decorated baseline: cell B2 has single underline and strikethrough drawn and observed
    const decoratedCaseDef: CfOracleCase = {
      ...validNumFmtCase,
      expected: {
        ...validNumFmtCase.expected,
        provenance: 'excel-observed',
        cells: validNumFmtCase.expected.cells.map(x =>
          x.ref === 'B2' ? { ...x, underline: 'single' as const, strikethrough: true } : x,
        ),
      },
    }
    const decPng = makeCaseMatchingPng(decoratedCaseDef, numFmtBytes, true)
    const decPngHash = sha256(decPng)
    const decPngPath = join(dir, 'numfmt-decorated.png')
    writeFileSync(decPngPath, decPng)

    const decObsCells = observeCellsFromCapture(decodePngSync(decPng)!, geomNumFmt, decoratedCaseDef, {
      observedCellTexts: validTexts,
      observedCellDecorations: {
        ...validDecs,
        B2: { underline: 'single', strikethrough: true },
      },
    })!
    const b2Obs = decObsCells.find(o => o.ref === 'B2')!
    expect(b2Obs.hasUnderlineInk).toBe(true)
    expect(b2Obs.hasStrikethroughInk).toBe(true)
    expect(b2Obs.underline).toBe('single')
    expect(b2Obs.strikethrough).toBe(true)

    const validDecoratedCase: CfOracleCase = {
      ...validNumFmtCase,
      expected: decoratedCaseDef.expected,
      captures: validNumFmtCase.captures.map(cap => {
        const sidePath = join(dir, `numfmt-decorated.${cap.appearance}.json`)
        writeFileSync(
          sidePath,
          JSON.stringify({
            sheet: cNumFmt.sheet,
            range: cNumFmt.range,
            appearance: cap.appearance,
            route: 'copy-picture',
            scale: 1,
            pxPerPt: 1,
            excelVersion: '16.114',
            geometrySource: 'excel-measured',
            colWidthsPt: geomNumFmt.colWidthsPt,
            rowHeightsPt: geomNumFmt.rowHeightsPt,
            resavedByExcel: true,
            resaveSavedSha256: numFmtHash,
            savedSha256: numFmtHash,
            pngSha256: decPngHash,
            capturedAt: '2026-11-06T08:00:00Z',
            observedCells: decObsCells,
          }),
        )
        return {
          appearance: cap.appearance,
          pngPath: decPngPath,
          sidecarPath: sidePath,
          pngSha256: decPngHash,
        }
      }),
    }
    const decRes = verifyCfCase(validDecoratedCase, { rootDir: REPO_ROOT })
    expect(decRes).toMatchObject({
      kind: 'verified',
      id: 'stage1-numfmt-accessibility',
      status: 'measured',
    })

    // 10. Valid undecorated baseline with dense glyph strings (reviewer r4-ink.ts repro strings)
    // Ensures arbitrary glyph density across 111, HHH, EEE, 000, 888, ___, ABC, LLL never triggers false decoration
    for (const testStr of ['111', 'HHH', 'EEE', '000', '888', '___', 'ABC', 'LLL']) {
      const denseCanvas = createCanvas(160, 22)
      const denseCtx = denseCanvas.getContext('2d')
      denseCtx.fillStyle = '#ffffff'
      denseCtx.fillRect(0, 0, 160, 22)
      denseCtx.fillStyle = '#000000'
      denseCtx.font = '11px Arial'
      denseCtx.fillText(testStr, 4, 15)
      denseCtx.fillText(testStr, 84, 15)
      const densePng = decodePngSync(new Uint8Array(denseCanvas.toBuffer('image/png')))!
      const denseGeom: WorksheetRangeGeometry = {
        minCol: 0,
        maxCol: 1,
        minRow: 0,
        maxRow: 0,
        colWidthsPt: [80, 80],
        rowHeightsPt: [22],
        colWidthChars: [10, 10],
        expectedWidthPt: 160,
        expectedHeightPt: 22,
      }
      const denseCaseDef = {
        expected: {
          cells: [{ ref: 'A1', controlRef: 'B1', matched: true, underline: 'none' as const, strikethrough: false }],
        },
      } as unknown as CfOracleCase

      const denseObs = observeCellsFromCapture(densePng, denseGeom, denseCaseDef, {
        observedCellTexts: { A1: testStr, B1: testStr },
        observedCellDecorations: { A1: { underline: 'none', strikethrough: false } },
      })![0]
      expect(denseObs.hasUnderlineInk).toBe(false)
      expect(denseObs.hasStrikethroughInk).toBe(false)
    }

    // 11. Valid cleared baseline: cell B2 in stage1-dxf-layering explicitly clears base bold/underline/strike
    const cDxf = structuredClone(manifest.cases.find(x => x.id === 'stage1-dxf-layering')!)
    const dxfBytes = new Uint8Array(readFileSync(join(REPO_ROOT, cDxf.saved)))
    const dxfHash = sha256(dxfBytes)
    cDxf.status = 'measured'
    cDxf.expected.provenance = 'excel-observed'
    cDxf.resavedByExcel = true
    cDxf.resaveReceipt = {
      openedRepairFree: true,
      resavedAt: '2026-11-06T07:00:00Z',
      excelVersion: '16.114',
      inputSha256: cDxf.expected.inputSha256,
      savedSha256: dxfHash,
    }
    // Cell B2 has underline="none", strikethrough=false. Control B8 has base formatting with underline="single", strikethrough=true
    const dxfCellsWithControlDecs = cDxf.expected.cells.map(x =>
      x.ref === 'B2'
        ? {
            ...x,
            controlUnderline: 'single' as const,
            controlStrikethrough: true,
          }
        : x,
    )
    const dxfCaseWithControlDecs: CfOracleCase = {
      ...cDxf,
      expected: { ...cDxf.expected, cells: dxfCellsWithControlDecs },
    }
    const dxfPng = makeCaseMatchingPng(dxfCaseWithControlDecs, dxfBytes, true)
    const dxfPngHash = sha256(dxfPng)
    const dxfPngPath = join(dir, 'dxf-layering-cleared.png')
    writeFileSync(dxfPngPath, dxfPng)

    const dxfGeom = computeWorksheetRangeGeometry(readZipPartsSync(dxfBytes)!, cDxf.range)!
    const dxfObsCells = observeCellsFromCapture(
      decodePngSync(dxfPng)!,
      dxfGeom,
      cDxf,
      {
        observedCellDecorations: {
          B2: { underline: 'none', strikethrough: false },
        },
      },
    )!
    const b2DxfObs = dxfObsCells.find(o => o.ref === 'B2')!
    expect(b2DxfObs.hasUnderlineInk).toBe(false)
    expect(b2DxfObs.hasStrikethroughInk).toBe(false)

    const clearedCase: CfOracleCase = {
      ...cDxf,
      captures: (cDxf.requiredAppearances ?? ['screen']).map(appearance => {
        const sidePath = join(dir, `dxf-layering-cleared.${appearance}.json`)
        writeFileSync(
          sidePath,
          JSON.stringify({
            sheet: cDxf.sheet,
            range: cDxf.range,
            appearance,
            route: 'copy-picture',
            scale: 1,
            pxPerPt: 1,
            excelVersion: '16.114',
            geometrySource: 'excel-measured',
            colWidthsPt: dxfGeom.colWidthsPt,
            rowHeightsPt: dxfGeom.rowHeightsPt,
            resavedByExcel: true,
            resaveSavedSha256: dxfHash,
            savedSha256: dxfHash,
            pngSha256: dxfPngHash,
            capturedAt: '2026-11-06T08:00:00Z',
            observedCells: dxfObsCells,
          }),
        )
        return {
          appearance,
          pngPath: dxfPngPath,
          sidecarPath: sidePath,
          pngSha256: dxfPngHash,
        }
      }),
    }
    const clearedRes = verifyCfCase(clearedCase, { rootDir: REPO_ROOT })
    expect(clearedRes).toMatchObject({
      kind: 'verified',
      id: 'stage1-dxf-layering',
      status: 'measured',
    })
  })

  it('enforces T1-A: unverified geometry never becomes measured; trusted display-format overrides pixel advice', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-r5-'))
    dirs.push(dir)
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const cNumFmt = structuredClone(manifest.cases.find(x => x.id === 'stage1-numfmt-accessibility')!)
    const numFmtBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, cNumFmt.saved)))
    const numFmtHash = sha256(numFmtBytes)
    const numFmtReceipt: CfResaveReceipt = {
      openedRepairFree: true,
      excelVersion: '16.114',
      inputSha256: cNumFmt.expected.inputSha256,
      savedSha256: numFmtHash,
      resavedAt: '2026-11-06T08:00:00Z',
    }
    const numFmtParts = readZipPartsSync(numFmtBytes)!
    const geomNumFmt = computeWorksheetRangeGeometry(numFmtParts, cNumFmt.range)!
    const validTexts: Record<string, string> = {}
    for (const cell of cNumFmt.expected.cells) {
      if (cell.formattedText !== undefined) validTexts[cell.ref] = cell.formattedText
    }
    const validDecs: Record<string, { underline?: 'none' | 'single' | 'double'; strikethrough?: boolean }> = {}
    for (const cell of cNumFmt.expected.cells) {
      validDecs[cell.ref] = {
        underline: cell.underline ?? 'none',
        strikethrough: cell.strikethrough ?? false,
      }
    }
    const validPng = makeCaseMatchingPng(cNumFmt, numFmtBytes, true)
    const validHash = sha256(validPng)
    const validPngPath = join(dir, 'r5-valid.png')
    writeFileSync(validPngPath, validPng)
    const validObs = observeCellsFromCapture(decodePngSync(validPng)!, geomNumFmt, cNumFmt, {
      observedCellTexts: validTexts,
      observedCellDecorations: validDecs,
    })!
    const baseMeasured: CfOracleCase = {
      ...cNumFmt,
      status: 'measured',
      resavedByExcel: true,
      resaveReceipt: numFmtReceipt,
      expected: { ...cNumFmt.expected, provenance: 'excel-observed' },
      captures: (cNumFmt.requiredAppearances ?? ['screen']).map(appearance => {
        const side = join(dir, `r5-valid.${appearance}.json`)
        writeFileSync(
          side,
          JSON.stringify({
            sheet: cNumFmt.sheet,
            range: cNumFmt.range,
            appearance,
            route: 'copy-picture',
            scale: 1,
            pxPerPt: 1,
            excelVersion: '16.114',
            geometrySource: 'excel-measured',
            colWidthsPt: geomNumFmt.colWidthsPt,
            rowHeightsPt: geomNumFmt.rowHeightsPt,
            resavedByExcel: true,
            resaveSavedSha256: numFmtHash,
            savedSha256: numFmtHash,
            pngSha256: validHash,
            capturedAt: '2026-11-06T08:00:00Z',
            observedCells: validObs,
          }),
        )
        return { appearance, pngPath: validPngPath, sidecarPath: side, pngSha256: validHash }
      }),
    }
    expect(verifyCfCase(baseMeasured, { rootDir: REPO_ROOT }).kind).toBe('verified')

    // T1-A(a,c): explicit geometrySource 'unverified' must never become measured,
    // even when pixel samples match the fallback model.
    const unverifiedSide = join(dir, 'r5-unverified.json')
    writeFileSync(
      unverifiedSide,
      JSON.stringify({
        sheet: cNumFmt.sheet,
        range: cNumFmt.range,
        appearance: 'screen',
        route: 'copy-picture',
        scale: 1,
        pxPerPt: 1,
        excelVersion: '16.114',
        geometrySource: 'unverified',
        colWidthsPt: geomNumFmt.colWidthsPt,
        rowHeightsPt: geomNumFmt.rowHeightsPt,
        resavedByExcel: true,
        resaveSavedSha256: numFmtHash,
        savedSha256: numFmtHash,
        pngSha256: validHash,
        capturedAt: '2026-11-06T08:00:00Z',
        observedCells: validObs,
      }),
    )
    expect(
      verifyCfCase(
        {
          ...baseMeasured,
          captures: baseMeasured.captures.map(cap =>
            cap.appearance === 'screen' ? { ...cap, sidecarPath: unverifiedSide } : cap,
          ),
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({ kind: 'refusal', code: 'geometry-mismatch' })

    // T1-A(a,c): missing geometrySource must also fail closed (no silent fallback to measured).
    const missingGeomSide = join(dir, 'r5-missing-geom.json')
    writeFileSync(
      missingGeomSide,
      JSON.stringify({
        sheet: cNumFmt.sheet,
        range: cNumFmt.range,
        appearance: 'screen',
        route: 'copy-picture',
        scale: 1,
        pxPerPt: 1,
        excelVersion: '16.114',
        resavedByExcel: true,
        resaveSavedSha256: numFmtHash,
        savedSha256: numFmtHash,
        pngSha256: validHash,
        capturedAt: '2026-11-06T08:00:00Z',
        observedCells: validObs,
      }),
    )
    expect(
      verifyCfCase(
        {
          ...baseMeasured,
          captures: baseMeasured.captures.map(cap =>
            cap.appearance === 'screen' ? { ...cap, sidecarPath: missingGeomSide } : cap,
          ),
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({ kind: 'refusal', code: 'geometry-mismatch' })

    // T1-A(b): trusted excel-display-format observation overrides disagreeing pixel advice.
    // PNG has no strikethrough ink, but trusted observation says strikethrough=true.
    const strikeExpectation: CfOracleCase = {
      ...baseMeasured,
      expected: {
        ...baseMeasured.expected,
        cells: baseMeasured.expected.cells.map(x =>
          x.ref === 'B2' ? { ...x, strikethrough: true } : x,
        ),
      },
    }
    const strikeObs = observeCellsFromCapture(decodePngSync(validPng)!, geomNumFmt, strikeExpectation, {
      observedCellTexts: validTexts,
      observedCellDecorations: { ...validDecs, B2: { underline: 'none', strikethrough: true } },
    })!
    const b2StrikeObs = strikeObs.find(o => o.ref === 'B2')!
    expect(b2StrikeObs.decorationSource).toBe('excel-display-format')
    expect(b2StrikeObs.hasStrikethroughInk).toBe(false)
    const strikeCaptures = (strikeExpectation.requiredAppearances ?? ['screen']).map(appearance => {
      const side = join(dir, `r5-strike-trusted.${appearance}.json`)
      writeFileSync(
        side,
        JSON.stringify({
          sheet: cNumFmt.sheet,
          range: cNumFmt.range,
          appearance,
          route: 'copy-picture',
          scale: 1,
          pxPerPt: 1,
          excelVersion: '16.114',
          geometrySource: 'excel-measured',
          colWidthsPt: geomNumFmt.colWidthsPt,
          rowHeightsPt: geomNumFmt.rowHeightsPt,
          resavedByExcel: true,
          resaveSavedSha256: numFmtHash,
          savedSha256: numFmtHash,
          pngSha256: validHash,
          capturedAt: '2026-11-06T08:00:00Z',
          observedCells: strikeObs,
        }),
      )
      return { appearance, pngPath: validPngPath, sidecarPath: side, pngSha256: validHash }
    })
    expect(
      verifyCfCase({ ...strikeExpectation, captures: strikeCaptures }, { rootDir: REPO_ROOT }).kind,
    ).toBe('verified')

    // T1-A(b): missing trusted observation stays explicitly unverified (no pixel authority).
    const rawObs = observeCellsFromCapture(decodePngSync(validPng)!, geomNumFmt, strikeExpectation, {
      observedCellTexts: validTexts,
    })!
    const b2Raw = rawObs.find(o => o.ref === 'B2')!
    expect(b2Raw.decorationSource).toBe('unverified')
    const rawSide = join(dir, 'r5-strike-untrusted.json')
    writeFileSync(
      rawSide,
      JSON.stringify({
        sheet: cNumFmt.sheet,
        range: cNumFmt.range,
        appearance: 'screen',
        route: 'copy-picture',
        scale: 1,
        pxPerPt: 1,
        excelVersion: '16.114',
        geometrySource: 'excel-measured',
        colWidthsPt: geomNumFmt.colWidthsPt,
        rowHeightsPt: geomNumFmt.rowHeightsPt,
        resavedByExcel: true,
        resaveSavedSha256: numFmtHash,
        savedSha256: numFmtHash,
        pngSha256: validHash,
        capturedAt: '2026-11-06T08:00:00Z',
        observedCells: rawObs,
      }),
    )
    expect(
      verifyCfCase(
        {
          ...strikeExpectation,
          captures: strikeExpectation.captures.map(cap =>
            cap.appearance === 'screen' ? { ...cap, sidecarPath: rawSide } : cap,
          ),
        },
        { rootDir: REPO_ROOT },
      ),
    ).toMatchObject({ kind: 'refusal', code: 'unobserved-expectation' })
  })

  it('verifies the 7 Excel 16.114 resaved Stage 1 captures and rejects mutated geometry/fill counterexamples', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const resavedIds = [
      'stage1-cellis-literals',
      'stage1-text-predicates',
      'stage1-blanks-errors',
      'stage1-duplicate-scope',
      'stage1-dxf-layering',
      'stage1-blank-merge-order',
      'stage1-numfmt-accessibility',
    ]

    for (const id of resavedIds) {
      const orig = manifest.cases.find(c => c.id === id)!
      const candidate: CfOracleCase = {
        ...orig,
        status: 'measured',
        expected: {
          ...orig.expected,
          provenance: 'excel-observed',
        },
      }
      const check = verifyCfCase(candidate, { rootDir: REPO_ROOT })
      expect(check.kind, `${id} must verify: ${check.kind === 'refusal' ? check.reason : ''}`).toBe('verified')
    }

    // Vacuity 1: 1-row geometry mutation must fail
    const literalsOrig = manifest.cases.find(c => c.id === 'stage1-cellis-literals')!
    const mutatedGeomCase: CfOracleCase = {
      ...literalsOrig,
      status: 'measured',
      expected: {
        ...literalsOrig.expected,
        provenance: 'excel-observed',
        // Shift B6 (positive fill) down by one row to B7 (unformatted control)
        cells: literalsOrig.expected.cells.map(c => c.ref === 'B6' ? { ...c, ref: 'B7' } : c),
      },
    }
    const geomCheck = verifyCfCase(mutatedGeomCase, { rootDir: REPO_ROOT })
    expect(geomCheck.kind).toBe('refusal')
    expect(geomCheck.kind === 'refusal' ? geomCheck.code : '').toBe('cell-expectation-mismatch')

    // Vacuity 2: Fills removed from capture must fail
    const testDir = mkdtempSync(join(tmpdir(), 'vacuity-fill-'))
    try {
      const origPngAbs = resolve(REPO_ROOT, literalsOrig.captures[0]!.pngPath)
      const origPngBytes = new Uint8Array(readFileSync(origPngAbs))
      const decoded = decodePngSync(origPngBytes)!
      // Overwrite the B6 box with white (remove fill)
      // B6 is at x: 84..168, y: 90..108
      for (let y = 90; y < 108; y++) {
        for (let x = 84; x < 168; x++) {
          const idx = (y * decoded.width + x) * 4
          decoded.data[idx] = 255
          decoded.data[idx + 1] = 255
          decoded.data[idx + 2] = 255
          decoded.data[idx + 3] = 255
        }
      }
      // Re-encode or simulate by mutating the sidecar / png hash
      const mutatedPngPath = join(testDir, 'mutated.png')
      // Create canvas with mutated data
      const canvas = createCanvas(decoded.width, decoded.height)
      const ctx = canvas.getContext('2d')
      const imgData = ctx.createImageData(decoded.width, decoded.height)
      imgData.data.set(decoded.data)
      ctx.putImageData(imgData, 0, 0)
      const mutatedBytes = new Uint8Array(canvas.toBuffer('image/png'))
      writeFileSync(mutatedPngPath, mutatedBytes)
      const mutatedPngHash = sha256(mutatedBytes)

      const origSidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, literalsOrig.captures[0]!.sidecarPath), 'utf8')) as Record<string, unknown>
      const mutatedSidecarPath = join(testDir, 'mutated.json')
      writeFileSync(mutatedSidecarPath, JSON.stringify({ ...origSidecar, pngSha256: mutatedPngHash }))

      const noFillCase: CfOracleCase = {
        ...literalsOrig,
        status: 'measured',
        captures: [{ ...literalsOrig.captures[0]!, pngPath: mutatedPngPath, pngSha256: mutatedPngHash, sidecarPath: mutatedSidecarPath }],
        expected: { ...literalsOrig.expected, provenance: 'excel-observed' },
      }
      const fillCheck = verifyCfCase(noFillCase, { rootDir: REPO_ROOT })
      expect(fillCheck.kind).toBe('refusal')
      expect(fillCheck.kind === 'refusal' ? fillCheck.code : '').toBe('cell-expectation-mismatch')
    } finally {
      rmSync(testDir, { recursive: true, force: true })
    }
  })

  it('structurally verifies stage2-scale-unresolved-shapes as unsupported-capability refusal', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const unresolvedCase = manifest.cases.find(c => c.id === 'stage2-scale-unresolved-shapes')!
    expect(unresolvedCase).toBeDefined()
    expect(unresolvedCase.status).toBe('unsupported')

    const res = verifyCfCase(unresolvedCase, { rootDir: REPO_ROOT })
    expect(res.kind).toBe('refusal')
    if (res.kind === 'refusal') {
      expect(res.structuralValid).toBe(true)
      expect(res.status).toBe('unsupported')
      expect(res.code).toBe('unsupported-capability')
    }
  })
  // -------------------------------------------------------------------------------------------------------------------
  // Stage 3 data bars (Task 18)
  // -------------------------------------------------------------------------------------------------------------------

  const X14_DECL = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
  const XM_DECL = 'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
  const GUID1 = '{3B3B0001-0000-4000-8000-000000000001}'

  /** Excel's canonical paired form (cfRule-level `<x14:id>` plus an x14 companion without priority). */
  const pairedBarXml = (
    o: {
      link?: boolean
      second?: boolean
      baseCfvo?: string
      baseDataBarAttrs?: string
      baseColor?: string
      x14Attrs?: string
      x14Cfvo?: string
      x14Extra?: string
    } = {},
  ): string => {
    const base = `<cfRule type="dataBar" priority="1"><dataBar${o.baseDataBarAttrs ?? ' minLength="10" maxLength="90"'}>${o.baseCfvo ?? '<cfvo type="min"/><cfvo type="max"/>'}${o.baseColor ?? '<color rgb="FF0000FF"/>'}</dataBar>${o.link === false ? '' : `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ${X14_DECL}><x14:id>${GUID1}</x14:id></ext></extLst>`}</cfRule>`
    const second = o.second
      ? `<cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ${X14_DECL}><x14:id>${GUID1}</x14:id></ext></extLst></cfRule>`
      : ''
    const companion = `<x14:cfRule type="dataBar" id="${GUID1}"><x14:dataBar${o.x14Attrs ?? ' minLength="0" maxLength="100" gradient="0"'}>${o.x14Cfvo ?? '<x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo><x14:cfvo type="autoMax"/>'}${o.x14Extra ?? ''}<x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>`
    return (
      `<conditionalFormatting sqref="B2:B3">${base}${second}</conditionalFormatting>` +
      `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${X14_DECL}><x14:conditionalFormattings><x14:conditionalFormatting ${XM_DECL}>${companion}<xm:sqref>B2:B3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
    )
  }

  const barPart = 'xl/worksheets/sheet1.xml'
  const pairedBaseMarker: CfPartMarker = {
    part: barPart,
    extension: 'legacy',
    capabilityId: 'bar:pairing-paired',
    sqref: 'B2:B3',
    ruleType: 'dataBar',
    priority: 1,
    formulas: [],
    bar: { pairing: 'paired', cfvo: [{ type: 'min' }, { type: 'max' }], attrs: { minLength: '10', maxLength: '90' }, colors: { fill: 'FF0000FF' } },
  }
  const pairedCompanionMarker: CfPartMarker = {
    part: barPart,
    extension: 'x14',
    capabilityId: 'bar:settings-paired-precedence',
    sqref: 'B2:B3',
    ruleType: 'dataBar',
    formulas: [],
    bar: {
      pairing: 'paired',
      cfvo: [{ type: 'num', f: '20' }, { type: 'autoMax' }],
      attrs: { minLength: '0', maxLength: '100', gradient: '0', negativeBarColorSameAsPositive: '0' },
      colors: { negativeFill: 'FFFF0000', axis: 'FF000000' },
    },
  }
  const TOL = { meanAbs: 0.015, rowMax: 0.08, strong: 12 }

  it('stage 3 cases join the manifest and stage 1-2 cases are unchanged', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    expect(manifest.cases.map(c => c.id)).toEqual(CF_CASES.map(c => c.id))
    // Stage 1 and 2 catalogs and capability ids are exactly what Stage 2 shipped.
    expect(CF_STAGE1_CAPABILITIES).toHaveLength(115)
    expect(CF_STAGE2_CAPABILITIES).toHaveLength(46)
    expect(CF_STAGE1_CASES).toHaveLength(13)
    expect(CF_STAGE2_CASES).toHaveLength(8)
    expect(CF_CAPABILITIES.length).toBe(CF_STAGE1_CAPABILITIES.length + CF_STAGE2_CAPABILITIES.length + CF_STAGE3_CAPABILITIES.length)
    expect(new Set(CF_CAPABILITIES.map(c => c.capabilityId)).size).toBe(CF_CAPABILITIES.length)
    expect(CF_CASES.slice(0, 21).map(c => c.stage)).toEqual([...Array(13).fill(1), ...Array(8).fill(2)])

    expect(CF_STAGE3_CASES.map(c => c.id)).toEqual([
      'stage3-bar-settings-probe',
      'stage3-bar-lengths',
      'stage3-bar-degenerate',
      'stage3-bar-axis-negative',
      'stage3-bar-rendering',
      'stage3-bar-geometry',
      'stage3-bar-showvalue',
      'stage3-bar-participation',
      'stage3-bar-scope',
      'stage3-bar-composition',
      'stage3-bar-pairing',
      'stage3-bar-unresolved-shapes',
    ])
    for (const c of manifest.cases.filter(x => x.stage === 3)) {
      // Invariants that hold before, during and after capture, as the harness writes the manifest.
      expect(['recipe-pending', 'unsupported', 'measured']).toContain(c.status)
      expect(c.expected.inputSha256).toMatch(/^[0-9a-f]{64}$/)
      expect(c.expected.savedSha256).toMatch(/^[0-9a-f]{64}$/)
      if (c.resavedByExcel === true) {
        expect(c.expected.savedSha256).toBe(sha256(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))))
      }
      let hasSidecar = false
      for (const cap of c.captures) {
        if (cap.pngSha256 === undefined) continue
        const png = resolve(REPO_ROOT, cap.pngPath)
        expect(existsSync(png), cap.pngPath).toBe(true)
        expect(existsSync(resolve(REPO_ROOT, cap.sidecarPath)), cap.sidecarPath).toBe(true)
        expect(sha256(new Uint8Array(readFileSync(png)))).toBe(cap.pngSha256)
        hasSidecar = true
      }
      for (const cell of c.expected.cells) {
        if (!cell.barEvidence) continue
        expect(['pixel-measured', 'non-discriminating', 'unverified']).toContain(cell.barEvidence.settings)
        expect(['pixel-measured', 'non-discriminating', 'unverified']).toContain(cell.barEvidence.geometry)
        if (cell.barEvidence.settings === 'excel-object-model' || cell.barEvidence.geometry === 'pixel-measured') {
          expect(hasSidecar, `${c.id}:${cell.ref} claims evidence without a capture sidecar`).toBe(true)
        }
      }
      const res = verifyCfCase(c, { rootDir: REPO_ROOT })
      if (res.kind === 'refusal') {
        expect(res.structuralValid).toBe(true)
        expect(['semantic-evidence-pending', 'unsupported-capability']).toContain(res.code)
      } else {
        expect(res.status).toBe('measured')
        expect(c.status).toBe('measured')
        expect(res.savedSha256).toBe(c.expected.savedSha256)
      }
    }
    // Every Stage 3 capability is a bar:* capability with a ruling citation, and production holds none of them.
    for (const entry of CF_STAGE3_CAPABILITIES) {
      expect(entry.capabilityId.startsWith('bar:')).toBe(true)
      expect(entry.failClosedDisposition).toMatch(/S3-[A-S]|Always unsupported/)
      expect(hasEvidence(entry.capabilityId)).toBe(false)
      expect(VERIFIED_CAPABILITIES.has(entry.capabilityId)).toBe(false)
    }
    expect(CF_STAGE3_CAPABILITIES.filter(e => barCapabilityKind(e.capabilityId) === 'settings').map(e => e.capabilityId).sort()).toEqual([
      'bar:pairing-base-only',
      'bar:pairing-paired',
      'bar:pairing-standalone',
      'bar:settings-paired-precedence',
      'bar:settings-precondition-probe',
      'bar:settings-schema-defaults',
      'bar:showvalue-copy-text',
    ])
    for (const id of Object.keys(CF_STAGE3_NONDISCRIMINATING)) {
      expect(CF_STAGE3_CAPABILITIES.some(e => e.capabilityId === id)).toBe(true)
    }
  })

  it('bar marker pins cfvo, attributes, colours and pairing exactly', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-bar-pin-'))
    dirs.push(dir)
    const makeBarCase = (filename: string, hash: string, markers: CfPartMarker[] = [pairedBaseMarker, pairedCompanionMarker]): CfOracleCase => ({
      id: 'bar-pin-case',
      stage: 3,
      question: 'Pins exact bar cfvo, attributes, colours and pairing',
      input: filename,
      saved: filename,
      sheet: 'Sheet1',
      range: 'A1:B3',
      requiredAppearances: ['screen'],
      markers,
      expected: {
        inputSha256: hash,
        savedSha256: hash,
        provenance: 'seed-hypothesis',
        positiveRefs: ['B2'],
        negativeRefs: ['B3'],
        cells: [
          { ref: 'B2', controlRef: 'A2', matched: true, barApplied: true, fillSource: 'unverified' },
          { ref: 'B3', controlRef: 'A3', matched: false, barApplied: false, fillSource: 'unverified' },
        ],
      },
      captures: [],
      status: 'recipe-pending',
      reason: 'Testing bar marker pinning',
      tolerance: TOL,
    })
    let n = 0
    const run = async (xml: string, markers?: CfPartMarker[]) => {
      const bytes = await makeSyntheticXlsx({ sheetRuleXml: xml, dxfCount: 0 })
      const name = `bar-${n++}.xlsx`
      writeFileSync(join(dir, name), bytes)
      return verifyCfCase(makeBarCase(name, sha256(bytes), markers), { rootDir: dir })
    }

    // 0. Canonical paired archive is structurally valid.
    expect(await run(pairedBarXml())).toMatchObject({ kind: 'refusal', code: 'semantic-evidence-pending', structuralValid: true })

    // The six mutations the verifier must refuse.
    const mutations: Array<[string, string]> = [
      ['reordered cfvo', pairedBarXml({ baseCfvo: '<cfvo type="max"/><cfvo type="min"/>' })],
      ['minLength 0 -> 10', pairedBarXml({ x14Attrs: ' minLength="10" maxLength="100" gradient="0"' })],
      ['dropped <x14:id>', pairedBarXml({ link: false })],
      ['<x14:fillColor> added to a paired companion', pairedBarXml({ x14Extra: '<x14:fillColor rgb="FF0000FF"/>' })],
      ['xm:f "20" -> "2"', pairedBarXml({ x14Cfvo: '<x14:cfvo type="num"><xm:f>2</xm:f></x14:cfvo><x14:cfvo type="autoMax"/>' })],
      ['negativeBarColorSameAsPositive flipped', pairedBarXml({ x14Attrs: ' minLength="0" maxLength="100" gradient="0" negativeBarColorSameAsPositive="1"' })],
      ['colour changed', pairedBarXml({ baseColor: '<color rgb="FF00FF00"/>' })],
      ['duplicate GUID link', pairedBarXml({ second: true })],
    ]
    for (const [label, xml] of mutations) {
      const res = await run(xml)
      expect(res, label).toMatchObject({ kind: 'refusal', code: 'stripped-rule', structuralValid: false })
    }

    // Semantics, not raw strings: attribute order, hex case and explicit schema defaults all verify.
    const semantic = pairedBarXml({
      baseDataBarAttrs: ' maxLength="90" showValue="1" minLength="10"',
      baseColor: '<color rgb="ff0000ff"/>',
      x14Attrs: ' gradient="0" maxLength="100" negativeBarColorSameAsPositive="0" minLength="0" axisPosition="automatic"',
    })
    expect(await run(semantic)).toMatchObject({ kind: 'refusal', code: 'semantic-evidence-pending', structuralValid: true })
    // An unpinned non-default attribute is a mismatch, not a silent pass.
    expect(await run(pairedBarXml({ baseDataBarAttrs: ' minLength="10" maxLength="90" showValue="0"' }))).toMatchObject({ code: 'stripped-rule', structuralValid: false })

    // Standalone x14 needs <x14:fillColor> and @priority; base-only must not be linked.
    const standaloneMarker: CfPartMarker = {
      part: barPart, extension: 'x14', capabilityId: 'bar:pairing-standalone', sqref: 'B2:B3', ruleType: 'dataBar', priority: 2, formulas: [],
      bar: { pairing: 'standalone', cfvo: [{ type: 'min' }, { type: 'max' }], attrs: { gradient: '0' }, colors: { fill: 'FF0000FF' } },
    }
    const standaloneXml = (fill: string, priority = ' priority="2"') =>
      `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${X14_DECL}><x14:conditionalFormattings><x14:conditionalFormatting ${XM_DECL}><x14:cfRule type="dataBar"${priority} id="${GUID1}"><x14:dataBar gradient="0"><x14:cfvo type="min"/><x14:cfvo type="max"/>${fill}</x14:dataBar></x14:cfRule><xm:sqref>B2:B3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
    expect(await run(standaloneXml('<x14:fillColor rgb="FF0000FF"/>'), [standaloneMarker])).toMatchObject({ code: 'semantic-evidence-pending', structuralValid: true })
    expect(await run(standaloneXml(''), [standaloneMarker])).toMatchObject({ code: 'stripped-rule', structuralValid: false })
    expect(await run(standaloneXml('<x14:fillColor rgb="FF0000FF"/>', ''), [standaloneMarker])).toMatchObject({ code: 'stripped-rule', structuralValid: false })
    const baseOnlyMarker: CfPartMarker = {
      part: barPart, extension: 'legacy', capabilityId: 'bar:pairing-base-only', sqref: 'B2:B3', ruleType: 'dataBar', priority: 1, formulas: [],
      bar: { pairing: 'base-only', cfvo: [{ type: 'min' }, { type: 'max' }], attrs: {}, colors: { fill: 'FF0000FF' } },
    }
    const baseOnlyXml = '<conditionalFormatting sqref="B2:B3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar></cfRule></conditionalFormatting>'
    expect(await run(baseOnlyXml, [baseOnlyMarker])).toMatchObject({ code: 'semantic-evidence-pending', structuralValid: true })
    expect(await run(pairedBarXml({ baseDataBarAttrs: '' }), [baseOnlyMarker])).toMatchObject({ code: 'stripped-rule', structuralValid: false })
  })

  /** A 240x90 capture of `A1:B3` (col A 40 px, col B 200 px, rows 30 px): painted by `paint` over white. */
  const barGeom: WorksheetRangeGeometry = {
    minRow: 0, maxRow: 2, minCol: 0, maxCol: 1,
    rowHeightsPt: [30, 30, 30], colWidthChars: [6, 28], colWidthsPt: [40, 200],
    expectedHeightPt: 90, expectedWidthPt: 240,
  }
  const paintBarPng = (paint: (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>) => void): Uint8Array => {
    const canvas = createCanvas(240, 90)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 240, 90)
    paint(ctx)
    return new Uint8Array(canvas.toBuffer('image/png'))
  }
  const measureOpts = { controlOf: () => 'B3', tolerancePx: 1, caseId: 'bar-measure-test' }

  it('bar edges are located by column scan within +-1 px', () => {
    // Solid bar at px 40..119 in B1 (cell left edge is 40).
    const solid = decodePngSync(paintBarPng(ctx => { ctx.fillStyle = '#0000ff'; ctx.fillRect(40, 0, 80, 30) }))!
    const solidRes = measureBarGeometry(solid, barGeom, ['B1'], measureOpts)
    expect(solidRes).toBeInstanceOf(Map)
    const bar = (solidRes as ReadonlyMap<string, CfObservedBar>).get('B1')!
    expect(bar.cellLeftPx).toBe(40)
    expect(bar.cellWidthPx).toBe(200)
    expect(bar.bar).toMatchObject({ leftPx: 40, rightPx: 119 })

    // A 1 px black axis at px 80, over the bar.
    const axis = decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#0000ff'; ctx.fillRect(40, 0, 80, 30)
      ctx.fillStyle = '#000000'; ctx.fillRect(80, 0, 1, 30)
    }))!
    const axisRes = measureBarGeometry(axis, barGeom, ['B1'], { ...measureOpts, axisArgb: 'FF000000' }) as ReadonlyMap<string, CfObservedBar>
    expect(axisRes.get('B1')).toMatchObject({ axisPx: 80, bar: { leftPx: 40, rightPx: 119 } })

    // Scanlines (y = 7, 15, 22) that disagree by 2 px refuse.
    const unstable = decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#0000ff'
      ctx.fillRect(40, 0, 80, 10)
      ctx.fillRect(42, 10, 78, 10)
      ctx.fillRect(40, 20, 80, 10)
    }))!
    expect(measureBarGeometry(unstable, barGeom, ['B1'], measureOpts)).toMatchObject({ kind: 'refusal', code: 'bar-edge-unstable' })
    // 1 px disagreement is inside the tolerance.
    const within = decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#0000ff'
      ctx.fillRect(40, 0, 80, 10)
      ctx.fillRect(41, 10, 79, 10)
      ctx.fillRect(40, 20, 80, 10)
    }))!
    expect(measureBarGeometry(within, barGeom, ['B1'], measureOpts)).toBeInstanceOf(Map)

    // A fade with no strong delta (< 64 per channel) is no bar.
    const fade = decodePngSync(paintBarPng(ctx => { ctx.fillStyle = '#e0e0ff'; ctx.fillRect(40, 0, 80, 30) }))!
    const fadeRes = measureBarGeometry(fade, barGeom, ['B1'], measureOpts) as ReadonlyMap<string, CfObservedBar>
    expect(fadeRes.get('B1')!.bar).toBeNull()

    // A 2 px inset: leftPx is cellLeftPx + 2.
    const inset = decodePngSync(paintBarPng(ctx => { ctx.fillStyle = '#0000ff'; ctx.fillRect(42, 0, 76, 30) }))!
    const insetBar = (measureBarGeometry(inset, barGeom, ['B1'], measureOpts) as ReadonlyMap<string, CfObservedBar>).get('B1')!
    expect(insetBar.bar!.leftPx).toBe(insetBar.cellLeftPx + 2)
    expect(insetBar.bar!.rightPx).toBe(117)
    // An unknown cell refuses instead of inventing a measurement.
    expect(measureBarGeometry(solid, barGeom, ['Z9'], measureOpts)).toMatchObject({ kind: 'refusal' })
  })

  it('bar edges ignore the control cell\'s value text and still see a bar lighter than text', () => {
    // Control B3 shows value text (dark pixels on the middle scanline only) to the right of the bar; the bar cell hides its value.
    const withText = decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#0000ff'; ctx.fillRect(40, 0, 80, 30)
      ctx.fillStyle = '#000000'; ctx.fillRect(160, 72, 10, 6)
    }))!
    const res = measureBarGeometry(withText, barGeom, ['B1'], measureOpts)
    expect(res).toBeInstanceOf(Map)
    expect((res as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toMatchObject({ leftPx: 40, rightPx: 119 })
    // A bar lighter than text (but >= 64 from the background) still measures.
    const light = decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#9090ff'; ctx.fillRect(40, 0, 80, 30)
      ctx.fillStyle = '#000000'; ctx.fillRect(160, 72, 10, 6)
    }))!
    const lightRes = measureBarGeometry(light, barGeom, ['B1'], measureOpts) as ReadonlyMap<string, CfObservedBar>
    expect(lightRes.get('B1')!.bar).toMatchObject({ leftPx: 40, rightPx: 119 })
  })

  it('bar edges use the bar cell\'s own background and never read through value text', () => {
    const barOf = (png: Uint8Array) => measureBarGeometry(decodePngSync(png)!, barGeom, ['B1'], measureOpts)
    // Own fill (light, >= 64 from white) with a solid bar at 60..119; the control is white.
    const own = barOf(paintBarPng(ctx => {
      ctx.fillStyle = '#e0a0a0'; ctx.fillRect(40, 0, 200, 30)
      ctx.fillStyle = '#0000ff'; ctx.fillRect(60, 0, 60, 30)
    }))
    expect(own).toBeInstanceOf(Map)
    expect((own as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toMatchObject({ leftPx: 60, rightPx: 119 })
    // A uniformly painted cell (no inset anywhere) is a filled cell with no bar: Excel's bars keep a vertical inset
    // (bar:geometry-insets), so the old full-width fallback that read such a fill as a bar is retired.
    const uniform = barOf(paintBarPng(ctx => { ctx.fillStyle = '#0000ff'; ctx.fillRect(40, 0, 200, 30) }))
    expect((uniform as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toBeNull()
    // A full-width bar drawn as Excel draws it (2 px vertical inset) measures through to the cell's last bar column.
    const full = barOf(paintBarPng(ctx => { ctx.fillStyle = '#0000ff'; ctx.fillRect(42, 2, 197, 25) }))
    expect((full as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toMatchObject({ leftPx: 42, rightPx: 238 })
    // Perimeter samples that disagree beyond the strong delta leave no consistent background.
    const split = barOf(paintBarPng(ctx => {
      ctx.fillStyle = '#e0a0a0'; ctx.fillRect(40, 0, 100, 30)
      ctx.fillStyle = '#20c020'; ctx.fillRect(140, 0, 100, 30)
      ctx.fillStyle = '#0000ff'; ctx.fillRect(90, 8, 30, 14)
    }))
    expect(split).toMatchObject({ kind: 'refusal', code: 'bar-background-unresolved' })

    // Value text: glyphs over the bar's last columns in the bar cell, the same glyphs over white in the control.
    const glyphOverEnd = barOf(paintBarPng(ctx => {
      ctx.fillStyle = '#0000ff'; ctx.fillRect(42, 2, 197, 26)
      ctx.fillStyle = '#000000'; ctx.fillRect(232, 5, 8, 20); ctx.fillRect(232, 65, 8, 20)
    }))
    expect(glyphOverEnd).toMatchObject({ kind: 'refusal', code: 'bar-edge-under-text' })
    // A bar that ends well left of the glyph band measures its true edges with the mask present.
    const clear = barOf(paintBarPng(ctx => {
      ctx.fillStyle = '#0000ff'; ctx.fillRect(40, 0, 80, 30)
      ctx.fillStyle = '#000000'; ctx.fillRect(200, 5, 8, 20); ctx.fillRect(200, 65, 8, 20)
    }))
    expect((clear as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toMatchObject({ leftPx: 40, rightPx: 119 })
  })

  it('a near-full bar in a filled cell refuses instead of measuring an inverted sliver, whatever the gap (round-3 Minor a)', () => {
    const fillWithBar = (barWidth: number) => decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#e0a0a0'; ctx.fillRect(40, 0, 200, 30)
      ctx.fillStyle = '#0000ff'; ctx.fillRect(40, 0, barWidth, 30)
    }))!
    // Gaps of 2, 4 and 10 px at the far edge of a filled cell: the corners 1 px inside the cell split, so the fill and
    // the bar cannot be told apart and the measurement refuses (the old < 4 px sliver rule failed open at 4+ px).
    for (const gap of [2, 4, 10]) {
      expect(measureBarGeometry(fillWithBar(200 - gap), barGeom, ['B1'], measureOpts), `gap ${gap}`).toMatchObject({ kind: 'refusal', code: 'bar-background-unresolved' })
    }
    // A 1 px gap leaves all four corners on the bar: the sliver rule refuses it.
    expect(measureBarGeometry(fillWithBar(199), barGeom, ['B1'], measureOpts)).toMatchObject({ kind: 'refusal', code: 'bar-background-unresolved' })
    // The same near-full bar drawn as Excel draws it (vertical inset) leaves the fill on every corner and measures exactly.
    const excelLike = decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#e0a0a0'; ctx.fillRect(40, 0, 200, 30)
      ctx.fillStyle = '#0000ff'; ctx.fillRect(42, 2, 190, 25)
    }))!
    expect((measureBarGeometry(excelLike, barGeom, ['B1'], measureOpts) as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toMatchObject({ leftPx: 42, rightPx: 231 })
  })

  it('a bar left of a filled neighbour, an Excel dotted axis and a zero-length bar measure cleanly (run 4: composition D3, axis-negative B4)', () => {
    // Composition D3: a vertically inset bar over a scale fill (the corners sit on the fill, two of the old corners sat on the bar).
    const overFill = decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#9f9c74'; ctx.fillRect(40, 0, 200, 30)
      ctx.fillStyle = '#0000ff'; ctx.fillRect(42, 2, 106, 26)
    }))!
    expect((measureBarGeometry(overFill, barGeom, ['B1'], measureOpts) as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toMatchObject({ leftPx: 42, rightPx: 147 })
    // A scale-filled cell whose value is the minimum: no bar, not a full-width one.
    const filledNoBar = decodePngSync(paintBarPng(ctx => { ctx.fillStyle = '#f8696b'; ctx.fillRect(40, 0, 200, 30) }))!
    expect((measureBarGeometry(filledNoBar, barGeom, ['B1'], measureOpts) as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toBeNull()
    // The left neighbour's fill painted over the shared gridline column is not bar ink.
    const bleed = decodePngSync(paintBarPng(ctx => {
      ctx.fillStyle = '#63be7b'; ctx.fillRect(0, 0, 41, 30)
      ctx.fillStyle = '#0000ff'; ctx.fillRect(42, 2, 60, 26)
    }))!
    expect((measureBarGeometry(bleed, barGeom, ['B1'], measureOpts) as ReadonlyMap<string, CfObservedBar>).get('B1')!.bar).toMatchObject({ leftPx: 42, rightPx: 101 })
    // Excel's dotted axis (3 on, 1 off) at x 100: a zero-length bar is "no bar" with its axis, not an unstable edge.
    const dotted = (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>) => {
      ctx.fillStyle = '#000000'
      for (let y = 0; y < 30; y += 4) ctx.fillRect(100, y, 1, 3)
    }
    const zero = decodePngSync(paintBarPng(dotted))!
    expect((measureBarGeometry(zero, barGeom, ['B1'], { ...measureOpts, axisArgb: 'FF000000' }) as ReadonlyMap<string, CfObservedBar>).get('B1')).toMatchObject({ bar: null, axisPx: 100 })
    // A positive bar starting after the axis: the axis column is never part of the bar.
    const positive = decodePngSync(paintBarPng(ctx => { dotted(ctx); ctx.fillStyle = '#0000ff'; ctx.fillRect(101, 2, 60, 26) }))!
    expect((measureBarGeometry(positive, barGeom, ['B1'], { ...measureOpts, axisArgb: 'FF000000' }) as ReadonlyMap<string, CfObservedBar>).get('B1')).toMatchObject({ bar: { leftPx: 101, rightPx: 160 }, axisPx: 100 })
  })

  it('--verify-only skips a case with no capture on disk and fails loudly when a claimed capture is missing', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-verify-only-'))
    dirs.push(dir)
    const cap = { appearance: 'screen' as const, pngPath: 'x.png', sidecarPath: 'x.json' }
    const c = { ...CF_STAGE3_CASES[1]!, captures: [cap] }
    expect(verifyOnlyCaptureState(c, dir)).toBe('skip')
    expect(() => verifyOnlyCaptureState({ ...c, captures: [{ ...cap, pngSha256: 'a'.repeat(64) }] }, dir)).toThrow(/x\.png/)
    writeFileSync(join(dir, 'x.png'), 'p'); writeFileSync(join(dir, 'x.json'), '{}')
    expect(verifyOnlyCaptureState(c, dir)).toBe('verify')
  })

  it('non-discriminating bar variant keeps its gate closed (Lesson 3)', () => {
    const measured = { B2: { leftPx: 84, rightPx: 143 }, B4: { leftPx: 84, rightPx: 190 } }
    // Two hypotheses 2 px apart at every designated cell cannot be told apart.
    const close = {
      B2: [{ id: 'a', leftPx: 84, rightPx: 143 }, { id: 'b', leftPx: 84, rightPx: 145 }],
      B4: [{ id: 'a', leftPx: 84, rightPx: 190 }, { id: 'b', leftPx: 84, rightPx: 192 }],
    }
    expect(discriminateBar(close, measured, 1)).toMatchObject({ status: 'non-discriminating' })
    // 3 px apart with exactly one hypothesis matching discriminates.
    const apart = {
      B2: [{ id: 'a', leftPx: 84, rightPx: 143 }, { id: 'b', leftPx: 84, rightPx: 146 }],
      B4: [{ id: 'a', leftPx: 84, rightPx: 190 }, { id: 'b', leftPx: 84, rightPx: 193 }],
    }
    expect(discriminateBar(apart, measured, 1)).toEqual({ status: 'discriminating', matched: 'a' })
    // Separated enough but neither hypothesis matches: bar-hypothesis-none.
    expect(discriminateBar(apart, { B2: { leftPx: 84, rightPx: 160 }, B4: { leftPx: 84, rightPx: 160 } }, 1)).toMatchObject({ status: 'none' })
    // A single hypothesis cannot discriminate anything.
    expect(discriminateBar({ B2: [{ id: 'a', leftPx: 84, rightPx: 143 }] }, { B2: measured.B2 }, 1)).toMatchObject({ status: 'non-discriminating' })
    // The static registry candidate is withheld and production holds no evidence for it.
    expect(CF_STAGE3_NONDISCRIMINATING['bar:bounds-automax']).toMatch(/autoMax/)
    expect(hasEvidence('bar:bounds-automax')).toBe(false)
    expect(VERIFIED_CAPABILITIES.has('bar:bounds-automax')).toBe(false)
  })

  describe('measured bar cases', () => {
    const geometryCap = 'bar:lengths-explicit-0-100'
    const baseOnlyBarXml = '<conditionalFormatting sqref="B2"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar></cfRule></conditionalFormatting>'
    const barMarkerFor = (capabilityId: string): CfPartMarker => ({
      part: barPart, extension: 'legacy', capabilityId, sqref: 'B2', ruleType: 'dataBar', priority: 1, formulas: [],
      bar: { pairing: 'base-only', cfvo: [{ type: 'min' }, { type: 'max' }], attrs: {}, colors: { fill: 'FF0000FF' } },
    })
    const BAR_LEFT = 84
    const BAR_RIGHT = 143
    const makeBarPng = (): Uint8Array => {
      const canvas = createCanvas(168, 54)
      const ctx = canvas.getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, 168, 54)
      ctx.fillStyle = '#0000ff'
      ctx.fillRect(BAR_LEFT, 18, BAR_RIGHT - BAR_LEFT + 1, 18)
      return new Uint8Array(canvas.toBuffer('image/png'))
    }

    type Overrides = {
      capabilityId?: string
      evidence?: CfCellExpectation['barEvidence']
      hypotheses?: Record<string, Array<{ id: string; leftPx: number; rightPx: number }>> | null
      readback?: unknown
      lockAfter?: string
      barApplied?: boolean
      calibration?: CfOracleCase['barCalibration']
      xml?: string
      markers?: CfPartMarker[]
    }
    const build = async (dir: string, o: Overrides): Promise<{ c: CfOracleCase; savedHash: string }> => {
      const xlsx = await makeSyntheticXlsx({ sheetRuleXml: o.xml ?? baseOnlyBarXml, dxfCount: 0 })
      const hash = sha256(xlsx)
      const png = makeBarPng()
      const pngHash = sha256(png)
      writeFileSync(join(dir, 'bar.xlsx'), xlsx)
      writeFileSync(join(dir, 'bar.png'), png)
      const cap = o.capabilityId ?? geometryCap
      const sidecar = {
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        zoom: 100,
        dpr: 1,
        width: 168,
        height: 54,
        colWidthsPt: [84, 84],
        rowHeightsPt: [18, 18, 18],
        geometrySource: 'excel-measured',
        resavedByExcel: true,
        resaveSavedSha256: hash,
        sheet: 'Sheet1',
        range: 'A1:B3',
        savedSha256: hash,
        pngSha256: pngHash,
        capturedAt: '2026-11-11T08:00:00Z',
        lockBefore: 'unlocked',
        lockAfter: o.lockAfter ?? 'unlocked',
        observedBars: [{ ref: 'B2', cellLeftPx: 84, cellWidthPx: 84, bar: { leftPx: BAR_LEFT, rightPx: BAR_RIGHT, topPx: 18, bottomPx: 35 } }],
        ...(o.hypotheses === null
          ? {}
          : { barHypotheses: { [cap]: o.hypotheses ?? { B2: [{ id: 'h-a', leftPx: BAR_LEFT, rightPx: BAR_RIGHT }, { id: 'h-b', leftPx: BAR_LEFT, rightPx: BAR_RIGHT + 6 }] } } }),
        ...(o.readback !== undefined ? { barSettingsReadback: o.readback } : {}),
      }
      writeFileSync(join(dir, 'bar.json'), JSON.stringify(sidecar))
      const c: CfOracleCase = {
        id: 'bar-measured-case',
        stage: 3,
        question: 'Bar evidence kinds',
        input: 'bar.xlsx',
        saved: 'bar.xlsx',
        sheet: 'Sheet1',
        range: 'A1:B3',
        requiredAppearances: ['screen'],
        markers: o.markers ?? [barMarkerFor(cap)],
        ...(o.calibration ? { barCalibration: o.calibration } : {}),
        expected: {
          inputSha256: hash,
          savedSha256: hash,
          provenance: 'excel-observed',
          positiveRefs: ['B2'],
          negativeRefs: ['B3'],
          cells: [
            { ref: 'B2', controlRef: 'B3', matched: true, barApplied: o.barApplied ?? true, fillSource: 'unverified', barEvidence: o.evidence ?? { settings: 'unverified', geometry: 'pixel-measured' } },
            { ref: 'B3', controlRef: 'B2', matched: false, barApplied: false, fillSource: 'unverified' },
          ],
        },
        resavedByExcel: true,
        resaveReceipt: { openedRepairFree: true, excelVersion: '16.114', inputSha256: hash, savedSha256: hash, resavedAt: '2026-11-11T08:00:00Z' },
        captures: [{ appearance: 'screen', pngPath: 'bar.png', sidecarPath: 'bar.json', pngSha256: pngHash }],
        status: 'measured',
        tolerance: TOL,
      }
      return { c, savedHash: hash }
    }

    it('a failed calibration refuses', async () => {
      const dir = mkdtempSync(join(tmpdir(), 'ov-cf-bar-probe-'))
      dirs.push(dir)
      // Calibration: the PNG has no gridlines, so gridline-to-gridline widths cannot be measured.
      const failed = await build(dir, { capabilityId: 'bar:pixel-calibration', hypotheses: null, calibration: { columns: ['A', 'B'], scanRow: 3, solidRef: 'B2', solidArgb: 'FF00FF00' } })
      expect(verifyCfCase(failed.c, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'bar-calibration-failed' })
    })

    it('a lock observed around a capture refuses and the verifier refuses a locked sidecar', async () => {
      const dir = mkdtempSync(join(tmpdir(), 'ov-cf-bar-lock-'))
      dirs.push(dir)
      const locked = await build(dir, { lockAfter: 'locked' })
      expect(verifyCfCase(locked.c, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'display-locked' })
    })

    it('geometry capabilities need the calibration and settings capabilities need the probe (cross-case gate)', async () => {
      const dir = mkdtempSync(join(tmpdir(), 'ov-cf-bar-cross-'))
      dirs.push(dir)
      const geo = await build(dir, {})
      const manifest: CfManifest = { version: 1, excelBuild: '16.114', locale: 'zh-TW', dateSystem: '1900', hostTimeZone: 'Asia/Taipei', generatedAt: '2026-11-11T00:00:00Z', cases: [geo.c] }
      // Verified in isolation, but bar:pixel-calibration is not verified anywhere: the gate stays closed.
      expect(isCapabilityVerified(geometryCap, manifest, { rootDir: dir })).toBe(false)
      expect(isCapabilityVerified('bar:pixel-calibration', manifest, { rootDir: dir })).toBe(false)
    })
  })

  it('capture refuses on a locked or unknown display', () => {
    const plist = (body: string): string => `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0"><array><dict><key>IOKitBuildVersion</key><string>x</string>${body}</dict></array></plist>`
    expect(parseScreenLockState(plist('<key>CGSSessionScreenIsLocked</key><true/>'))).toBe('locked')
    expect(parseScreenLockState(plist('<key>CGSSessionScreenIsLocked</key>\n\t<true/>'))).toBe('locked')
    expect(parseScreenLockState(plist('<key>CGSSessionScreenIsLocked</key><false/>'))).toBe('unlocked')
    expect(parseScreenLockState(plist(''))).toBe('unlocked')
    // Malformed input is never "unlocked".
    expect(parseScreenLockState('')).toBe('locked')
    expect(parseScreenLockState('not a plist')).toBe('locked')
    expect(parseScreenLockState(plist('<key>CGSSessionScreenIsLocked</key><integer>1</integer>'))).toBe('locked')
    expect(parseScreenLockState(plist('<key>CGSSessionScreenIsLocked</key>'))).toBe('locked')

    // An ioreg failure is a locked display, not an unlocked one.
    expect(isDisplayLocked(() => { throw new Error('ioreg: not found') })).toBe(true)
    expect(isDisplayLocked(() => plist('<key>CGSSessionScreenIsLocked</key><false/>'))).toBe(false)
    expect(isDisplayLocked(() => plist('<key>CGSSessionScreenIsLocked</key><true/>'))).toBe(true)
    expect(isDisplayLocked(() => '')).toBe(true)

    // The lock is checked immediately before and after every capture.
    const events: string[] = []
    const states = ['unlocked', 'unlocked']
    const ok = captureUnderLockGuard({ isLocked: () => states.shift() === 'locked', run: () => { events.push('run') }, discard: () => { events.push('discard') } })
    expect(ok).toEqual({ kind: 'captured' })
    expect(events).toEqual(['run'])

    const beforeEvents: string[] = []
    const lockedBefore = captureUnderLockGuard({ isLocked: () => true, run: () => { beforeEvents.push('run') }, discard: () => { beforeEvents.push('discard') } })
    expect(lockedBefore).toMatchObject({ kind: 'refused', code: 'display-locked', phase: 'before' })
    expect(beforeEvents).toEqual([])

    const afterEvents: string[] = []
    const afterStates = [false, true]
    const lockedAfter = captureUnderLockGuard({ isLocked: () => afterStates.shift() === true, run: () => { afterEvents.push('run') }, discard: () => { afterEvents.push('discard') } })
    expect(lockedAfter).toMatchObject({ kind: 'refused', code: 'display-locked', phase: 'after' })
    expect(afterEvents).toEqual(['run', 'discard'])

    // A capture that throws is discarded too.
    const thrown: string[] = []
    expect(() => captureUnderLockGuard({ isLocked: () => false, run: () => { throw new Error('boom') }, discard: () => { thrown.push('discard') } })).toThrow('boom')
    expect(thrown).toEqual(['discard'])
  })

  it('probes data bar settings with every property read guarded by try', () => {
    const script = buildDataBarSettingsProbeScript('wb.xlsx', 'Sheet1').join('\n')
    expect(script).toContain('set wb to workbook "wb.xlsx"')
    for (const prop of [
      'databar axis position of fc',
      'databar fill type of fc',
      'databar border type of databar border of fc',
      'min point percent of fc',
      'max point percent of fc',
      'databar direction of fc',
      'format condition show value of fc',
      'negative bar color type of negative bar format of fc',
      'format condition priority of fc',
    ]) {
      expect(script, prop).toMatch(new RegExp(`try[\\s\\S]*?${prop}[\\s\\S]*?end try`))
    }
    expect(script).toMatch(/set axisVal to "missing value"/)
    expect(script).toContain('|||')
    expect(script).toContain(':::')
  })

  it('parses the settings readback into one row per format condition', () => {
    const out = ['1', 'databar axis midpoint', 'databar fill solid', 'databar border solid', '0', '100', '-4152', 'false', 'databar color type color'].join(':::') + '|||' +
      ['2', 'databar axis automatic', 'databar fill gradient', 'databar border none', '10', '90', '-5002', 'true', 'databar color type color'].join(':::')
    const rows = parseDataBarSettingsReadback(out)
    expect(rows).toHaveLength(2)
    expect(rows[0]).toMatchObject({ priority: 1, axis: 'databar axis midpoint', fillType: 'databar fill solid', borderType: 'databar border solid', minPointPercent: '0', maxPointPercent: '100', direction: '-4152', showValue: 'false', negativeColorType: 'databar color type color' })
    expect(parseDataBarSettingsReadback('')).toEqual([])
  })

  it('audit --stage 3 --write is idempotent (Lesson 4)', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const first = renderVerifiedCapabilitiesSource(manifest, 3, REPO_ROOT)
    const second = renderVerifiedCapabilitiesSource(manifest, 3, REPO_ROOT)
    expect(second).toBe(first)
    // Stage 1/2 sets are exactly what is committed; Stage 3 holds nothing without captures.
    const listed = (name: string): string[] => {
      const m = new RegExp(`export const ${name}: ReadonlySet<string> = new Set\\(\\[([\\s\\S]*?)\\]\\)`).exec(first)
      expect(m, name).not.toBeNull()
      return [...m![1]!.matchAll(/'([^']+)'/g)].map(x => x[1]!)
    }
    expect(listed('VERIFIED_STAGE1_CAPABILITIES')).toEqual([...VERIFIED_STAGE1_CAPABILITIES].sort())
    expect(listed('VERIFIED_STAGE2_CAPABILITIES')).toEqual([...VERIFIED_STAGE2_CAPABILITIES].sort())
    const reported = auditStage3Manifest(manifest, { rootDir: REPO_ROOT }).perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()
    expect(listed('VERIFIED_STAGE3_CAPABILITIES')).toEqual(reported)
    expect(listed('VERIFIED_CAPABILITIES')).toEqual([...new Set([...VERIFIED_STAGE1_CAPABILITIES, ...VERIFIED_STAGE2_CAPABILITIES, ...reported])].sort())
    // Idempotence is about rendering, not the committed file: the test never writes verifiedCapabilities.ts.
    const committed = readFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
    for (const name of ['VERIFIED_STAGE1_CAPABILITIES', 'VERIFIED_STAGE2_CAPABILITIES']) {
      const block = (text: string): string => new RegExp(`export const ${name}: ReadonlySet<string> = new Set\\(\\[[\\s\\S]*?\\]\\)`).exec(text)![0]
      expect(block(first), name).toBe(block(committed))
    }
    // Stage 2's own rendering is unchanged by the Stage 3 work (no VERIFIED_STAGE3 export).
    expect(renderVerifiedCapabilitiesSource(manifest, 2, REPO_ROOT)).not.toContain('VERIFIED_STAGE3_CAPABILITIES')

    const report = auditStage3Manifest(manifest, { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    expect(report.stageComplete).toBe(false)
    expect(report.summary.cases).toBe(12)
    expect(report.summary.capabilities).toBe(CF_STAGE3_CAPABILITIES.length)
    expect(report.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()).toEqual(reported)
    expect(report.blockedCapabilities.length).toBeGreaterThan(0)
    expect(report.perCapability.filter(p => p.disposition === 'refused-unsupported').every(p => p.capabilityId.startsWith('bar:unresolved-'))).toBe(true)
  })
  it('stage3ObservedBars reads the recorded px and ignores malformed entries', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-bar-observed-'))
    dirs.push(dir)
    const mk = (stage: 2 | 3): CfOracleCase => ({
      id: 'observed-bars-case', stage, question: 'q', input: 'x.xlsx', saved: 'x.xlsx', sheet: 'Sheet1', range: 'A1:B3', markers: [],
      expected: { inputSha256: '', savedSha256: '', provenance: 'seed-hypothesis', positiveRefs: [], negativeRefs: [], cells: [] },
      captures: [{ appearance: 'screen', pngPath: 'x.png', sidecarPath: 'x.json' }], status: 'recipe-pending', tolerance: TOL,
    })
    writeFileSync(join(dir, 'x.json'), JSON.stringify({
      observedBars: [
        { ref: 'b2', cellLeftPx: 84, cellWidthPx: 84, bar: { leftPx: 84, rightPx: 143, topPx: 18, bottomPx: 35 }, axisPx: 90 },
        { ref: 'B3', cellLeftPx: 84, cellWidthPx: 84, bar: null },
        { ref: 'B4', cellLeftPx: 84, cellWidthPx: 84, bar: { leftPx: 'x' } },
        { cellLeftPx: 1 },
      ],
    }))
    const observed = stage3ObservedBars(mk(3), dir)!
    expect([...observed.keys()].sort()).toEqual(['B2', 'B3'])
    expect(observed.get('B2')).toMatchObject({ axisPx: 90, bar: { leftPx: 84, rightPx: 143 } })
    expect(observed.get('B3')!.bar).toBeNull()
    expect(stage3ObservedBars(mk(2), dir)).toBeUndefined()
    expect(stage3ObservedBars({ ...mk(3), captures: [] }, dir)).toBeUndefined()
    writeFileSync(join(dir, 'x.json'), '{not json')
    expect(stage3ObservedBars(mk(3), dir)).toBeUndefined()
  })
  /**
   * The committed Excel capture of `stage3-bar-lengths` (real PNG, real resaved workbook) run through the capture
   * pipeline and the verifier in a temp dir: labels come from the pipeline, the verifier recomputes them.
   */
  const realLengthsCase = (dir: string) => {
    const tpl = CF_STAGE3_CASES.find(x => x.id === 'stage3-bar-lengths')!
    const man = (JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest).cases.find(x => x.id === tpl.id)!
    const bytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, man.saved)))
    const png = new Uint8Array(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs/stage3-bar-lengths.png')))
    const committed = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs/stage3-bar-lengths.json'), 'utf8')) as Record<string, unknown>
    writeFileSync(join(dir, 'lengths.xlsx'), bytes)
    writeFileSync(join(dir, 'lengths.input.xlsx'), readFileSync(resolve(REPO_ROOT, man.input)))
    writeFileSync(join(dir, 'lengths.png'), png)
    const parts = readZipPartsSync(bytes)!
    const geom = computeWorksheetRangeGeometry(parts, tpl.range, { colWidthsPt: committed.colWidthsPt as number[], rowHeightsPt: committed.rowHeightsPt as number[] })!
    const decoded = decodePngSync(png)!
    const base: CfOracleCase = { ...man, markers: tpl.markers, input: 'lengths.input.xlsx', saved: 'lengths.xlsx', captures: [{ appearance: 'screen', pngPath: 'lengths.png', sidecarPath: 'lengths.json', pngSha256: sha256(png) }] }
    const { extras, observed, evaluation, refusal } = buildBarSidecarExtras(base, parts, decoded, geom, { lockBefore: 'unlocked', lockAfter: 'unlocked' })
    expect(refusal).toBeUndefined()
    const labels = classifyBarCapture(base, evaluation!)
    const cells: CfCellExpectation[] = tpl.expected.cells.map(cell => {
      if (!cell.matched) return cell
      const m = observed!.get(cell.ref)
      return { ...cell, ...(m ? { barApplied: m.bar !== null } : {}), barEvidence: labels.get(cell.ref) ?? { settings: 'unverified', geometry: 'unverified' } }
    })
    const c: CfOracleCase = { ...base, expected: { ...man.expected, cells } }
    const sidecar = { ...committed, ...extras }
    const write = (sc: Record<string, unknown>) => writeFileSync(join(dir, 'lengths.json'), JSON.stringify(sc))
    write(sidecar)
    return { c, sidecar, write, evaluation: evaluation!, labels }
  }

  it('the capture pipeline feeds a verifier that recomputes every bar claim from the hash-bound PNG (real lengths capture)', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-bar-pipeline-'))
    dirs.push(dir)
    const { c, sidecar, write, evaluation, labels } = realLengthsCase(dir)
    // Predicted px come from the saved workbook's values, the markers and the measured cell boxes (Excel's 2 px / 1 px
    // inset model); 0/100 discriminates from 10/90 on the off-midpoint cells.
    expect(evaluation['bar:lengths-explicit-0-100']).toMatchObject({ status: 'discriminating', matched: 'effective' })
    expect(labels.get('C3')).toEqual({ settings: 'unverified', geometry: 'pixel-measured' })
    // num 20/80 equals this data's min/max: both answers predict the same px (non-discriminating, Lesson 3).
    expect(evaluation['bar:bounds-num']).toMatchObject({ status: 'non-discriminating' })
    expect(labels.get('E3')).toEqual({ settings: 'unverified', geometry: 'non-discriminating' })
    // Excel's default 10/90 matches neither 10/90 linear on the span nor 0/100: a finding, never a claim.
    expect(evaluation['bar:lengths-default-10-90']).toMatchObject({ status: 'none' })
    expect(labels.get('B3')).toEqual({ settings: 'unverified', geometry: 'unverified' })

    const res = verifyCfCase(c, { rootDir: dir })
    expect(res, JSON.stringify(res).slice(0, 400)).toMatchObject({ kind: 'verified' })
    if (res.kind !== 'verified') return
    expect([...res.verifiedCapabilities].sort()).toEqual(['bar:bounds-automin', 'bar:bounds-min-max', 'bar:bounds-percentile', 'bar:lengths-all-negative', 'bar:lengths-explicit-0-100'])
    expect([...(res.nonDiscriminatingCapabilities ?? [])].sort()).toEqual(['bar:bounds-automax', 'bar:bounds-num', 'bar:bounds-percent'])
    expect(res.refusedCapabilities).toEqual([expect.objectContaining({ capabilityId: 'bar:lengths-default-10-90', code: 'bar-hypothesis-none' })])
    expect(res.matchedHypotheses?.['bar:bounds-percentile']).toBe('inc')

    const relabel = (ref: string, ev: CfCellExpectation['barEvidence']): CfOracleCase => ({ ...c, expected: { ...c.expected, cells: c.expected.cells.map(x => (x.ref === ref ? { ...x, barEvidence: ev } : x)) } })
    const relabelCap = (refs: string[], ev: CfCellExpectation['barEvidence']): CfOracleCase => ({ ...c, expected: { ...c.expected, cells: c.expected.cells.map(x => (refs.includes(x.ref) ? { ...x, barEvidence: ev } : x)) } })
    // The retired object-model kind refuses, wherever it is claimed.
    expect(verifyCfCase(relabel('C3', { settings: 'excel-object-model', geometry: 'pixel-measured' }), { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'unobserved-expectation' })
    // A geometry capability claimed in the settings field refuses.
    expect(verifyCfCase(relabelCap(['C2', 'C3', 'C5', 'C6'], { settings: 'pixel-measured', geometry: 'unverified' }), { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'unobserved-expectation' })
    // A non-discriminating capability relabelled pixel-measured refuses; an underclaim refuses too.
    expect(verifyCfCase(relabelCap(['E2', 'E3', 'E4', 'E5', 'E6'], { settings: 'unverified', geometry: 'pixel-measured' }), { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'unobserved-expectation' })
    expect(verifyCfCase(relabelCap(['C2', 'C3', 'C5', 'C6'], { settings: 'unverified', geometry: 'non-discriminating' }), { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'unobserved-expectation' })
    // Claiming the refused 10/90 family refuses the case with bar-hypothesis-none.
    expect(verifyCfCase(relabelCap(['B2', 'B3', 'B5', 'B6'], { settings: 'unverified', geometry: 'pixel-measured' }), { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'bar-hypothesis-none' })
    // barApplied must follow the PNG.
    expect(verifyCfCase({ ...c, expected: { ...c.expected, cells: c.expected.cells.map(x => (x.ref === 'C3' ? { ...x, barApplied: false } : x)) } }, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'cell-expectation-mismatch' })

    // Sidecar hypotheses are recomputed, never trusted: a moved prediction refuses, and so does a missing one.
    const hyps = sidecar.barHypotheses as Record<string, Record<string, Array<{ id: string; rightPx: number }>>>
    const moved = JSON.parse(JSON.stringify(hyps)) as typeof hyps
    moved['bar:lengths-explicit-0-100']!.C3![1]!.rightPx += 9
    write({ ...sidecar, barHypotheses: moved })
    expect(verifyCfCase(c, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
    const { ['bar:lengths-explicit-0-100']: _gone, ...without } = hyps
    write({ ...sidecar, barHypotheses: without })
    expect(verifyCfCase(c, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'unobserved-expectation' })
    // Recorded vertical edges and the axis are compared too (review Minor 1).
    for (const field of ['topPx', 'bottomPx'] as const) {
      write({ ...sidecar, observedBars: (sidecar.observedBars as Array<{ ref: string; bar: Record<string, number> | null }>).map(b => (b.ref === 'C3' && b.bar ? { ...b, bar: { ...b.bar, [field]: b.bar[field]! + 1 } } : b)) })
      expect(verifyCfCase(c, { rootDir: dir }), field).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
    }
    write({ ...sidecar, observedBars: (sidecar.observedBars as Array<{ ref: string }>).map(b => (b.ref === 'C3' ? { ...b, axisPx: 300 } : b)) })
    expect(verifyCfCase(c, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
    // Recorded px that disagree with the PNG refuse.
    write({ ...sidecar, observedBars: (sidecar.observedBars as Array<{ ref: string; bar: { rightPx: number } | null }>).map(b => (b.ref === 'C3' && b.bar ? { ...b, bar: { ...b.bar, rightPx: b.bar.rightPx + 5 } } : b)) })
    expect(verifyCfCase(c, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
  })

  it('the pixel precondition probe passes on its five observed settings and refuses when any one is unobserved or equals the control (Ruling S3-PROBE5)', () => {
    const man = (JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest).cases.find(x => x.id === 'stage3-bar-settings-probe')!
    const tpl = CF_STAGE3_CASES.find(x => x.id === man.id)!
    const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs/stage3-bar-settings-probe.json'), 'utf8')) as Record<string, number[]>
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, man.saved))))!
    const geom = computeWorksheetRangeGeometry(parts, man.range, { colWidthsPt: sc.colWidthsPt!, rowHeightsPt: sc.rowHeightsPt! })!
    const pngBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs/stage3-bar-settings-probe.png')))
    const png = decodePngSync(pngBytes)!
    const c: CfOracleCase = { ...man, markers: tpl.markers, expected: { ...man.expected, cells: tpl.expected.cells } }
    const measured = measureCaseBars(c, png, geom) as ReadonlyMap<string, CfObservedBar>
    // Ruling S3-PROBE5: the gate is fill, border, direction, axis and value text; lengths and negative colours are
    // verified by their own cases, so nothing else blocks the probe.
    expect(BAR_PIXEL_PROBE_UNOBSERVABLE).toEqual({})
    expect(evaluateBarPixelProbe(c, png, geom, measured)).toMatchObject({ status: 'discriminating' })
    // One setting equal to the control: the probe measured against the control's own cells.
    const same: CfOracleCase = { ...c, markers: c.markers.map(m => (m.bar && m.priority === 2 ? { ...m, sqref: 'B2:B6' } : m)) }
    expect(evaluateBarPixelProbe(same, png, geom, measured)).toMatchObject({ status: 'unobserved', reason: expect.stringMatching(/equals the default control/) })
    // One setting unobserved: B2's border painted over with its fill, so the probe's border is not seen consistently.
    const patched = decodePngSync(pngBytes)!
    const b2 = measured.get('B2')!.bar!
    const fillAt = (x: number, y: number) => (y * patched.width + x) * 4
    for (const x of [b2.leftPx, b2.rightPx]) {
      for (let y = b2.topPx; y <= b2.bottomPx; y++) for (let k = 0; k < 4; k++) patched.data[fillAt(x, y) + k] = patched.data[fillAt(b2.leftPx + 3, y) + k]!
    }
    expect(evaluateBarPixelProbe(c, patched, geom, measured)).toMatchObject({ status: 'unobserved', reason: expect.stringMatching(/probe border is not observed consistently/) })
    // Fewer than two probe bars: nothing to observe.
    const empty = new Map([...measured].map(([ref, o]) => [ref, ref.startsWith('B') ? { ...o, bar: null } : o]))
    expect(evaluateBarPixelProbe(c, png, geom, empty)).toMatchObject({ status: 'unobserved' })
  })

  it('Excel\'s canonical resave of a standalone x14 bar (paired base + companion) is structurally accepted (stage3-bar-pairing.rejected.xlsx)', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-bar-standalone-'))
    dirs.push(dir)
    const rejected = new Uint8Array(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/stage3-bar-pairing.rejected.xlsx')))
    writeFileSync(join(dir, 'rejected.xlsx'), rejected)
    const tpl = CF_STAGE3_CASES.find(x => x.id === 'stage3-bar-pairing')!
    const h = sha256(rejected)
    const c: CfOracleCase = { ...tpl, input: 'rejected.xlsx', saved: 'rejected.xlsx', captures: [], expected: { ...tpl.expected, inputSha256: h, savedSha256: h } }
    // Every marker, the standalone one included, is found: only the evidence is pending.
    expect(verifyCfCase(c, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'semantic-evidence-pending', structuralValid: true })
    // The same file with the companion's link GUID broken is not the canonical form and refuses.
    const zip = await JSZip.loadAsync(rejected)
    const xml = (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).replace('<x14:id>{3B3B000B-0000-4000-8000-000000000002}</x14:id>', '<x14:id>{3B3B000B-0000-4000-8000-0000000000FF}</x14:id>')
    zip.file('xl/worksheets/sheet1.xml', xml)
    const broken = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    writeFileSync(join(dir, 'broken.xlsx'), broken)
    const hb = sha256(broken)
    expect(verifyCfCase({ ...c, input: 'broken.xlsx', saved: 'broken.xlsx', expected: { ...c.expected, inputSha256: hb, savedSha256: hb } }, { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'stripped-rule' })
  })

  it('--verify-only never rebinds a kept capture to a regenerated saved workbook', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-stale-'))
    dirs.push(dir)
    const c = { ...CF_STAGE3_CASES[1]!, captures: [{ appearance: 'screen' as const, pngPath: 'x.png', sidecarPath: 'x.json' }] }
    expect(staleCaptureSidecar(c, dir, 'b'.repeat(64))).toBe(false)
    writeFileSync(join(dir, 'x.json'), JSON.stringify({ savedSha256: 'a'.repeat(64) }))
    expect(staleCaptureSidecar(c, dir, 'a'.repeat(64))).toBe(false)
    expect(staleCaptureSidecar(c, dir, 'b'.repeat(64))).toBe(true)
  })

  it('capture closes the workbook it opened, never a pre-existing one, even when the capture throws', () => {
    const closed: string[] = []
    const close = (name: string) => { closed.push(name) }
    captureAndCloseOwned('/x/y/stage3-bar-lengths.xlsx', new Set(), () => {}, close)
    expect(closed).toEqual(['stage3-bar-lengths.xlsx'])
    expect(() => captureAndCloseOwned('/x/y/a.xlsx', new Set(), () => { throw new Error('boom') }, close)).toThrow('boom')
    expect(closed).toEqual(['stage3-bar-lengths.xlsx', 'a.xlsx'])
    captureAndCloseOwned('/x/y/mine.xlsx', new Set(['mine.xlsx']), () => {}, close)
    expect(closed).toEqual(['stage3-bar-lengths.xlsx', 'a.xlsx'])
  })

  it('generate never resets a measured case unless --force is passed (round-3 Minor b)', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-gen-measured-'))
    dirs.push(dir)
    mkdirSync(join(dir, 'tests/fixtures/xlsx-cf/inputs'), { recursive: true })
    const real = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const measured = real.cases.find(x => x.id === 'stage3-bar-lengths')!
    // A measured entry whose saved workbook is gone (so it no longer verifies) must still be kept verbatim.
    writeFileSync(join(dir, 'tests/fixtures/xlsx-cf/manifest.json'), JSON.stringify({ ...real, cases: [measured] }))
    const kept = await generateAllStage1({ rootDir: dir, onlyCaseId: measured.id, stage: 3 })
    expect(kept.cases.find(x => x.id === measured.id)).toEqual(measured)
    // A recipe-pending case keeps its kept captures too (review Minor 4)...
    const pending = { ...measured, status: 'recipe-pending' as const, reason: 'kept capture refused' }
    writeFileSync(join(dir, 'tests/fixtures/xlsx-cf/manifest.json'), JSON.stringify({ ...real, cases: [pending] }))
    const keptPending = await generateAllStage1({ rootDir: dir, onlyCaseId: measured.id, stage: 3 })
    expect(keptPending.cases.find(x => x.id === measured.id)!.captures).toEqual(measured.captures)
    // ...and only --force drops either.
    const forced = await generateAllStage1({ rootDir: dir, onlyCaseId: measured.id, stage: 3, force: true })
    const after = forced.cases.find(x => x.id === measured.id)!
    expect(after.status).toBe('recipe-pending')
    expect(after.captures[0]!.pngSha256).toBeUndefined()
  })
})
