import { spawnSync } from 'node:child_process'
import * as verifiedExports from '../src/xlsx/cf/verifiedCapabilities'
import { createHash } from 'node:crypto'
import { cpSync, existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'

const hasOsacompile = process.platform === 'darwin' && spawnSync('which', ['osacompile']).status === 0
import { compareCfCase, stage3ObservedBars, stage4ObservedIcons } from '../scripts/cf-compare'
import { auditStage3Manifest, auditStage4Manifest, CF_STAGE4_NEVER_PUBLISH, renderIconSelectionsBlock, renderVerifiedCapabilitiesSource, stage4Publication } from '../scripts/cf-refs/audit'
import { assertBuildDisplayUnlocked, assertCanCloseWorkbook, filterCasesForStage, isDisplayLocked, parseScreenLockState, probeAndBuildCase, shouldSkipCaseOnResume } from '../scripts/cf-refs/build'
import {
  CF_CAPABILITIES,
  CF_CASES,
  CF_STAGE1_CAPABILITIES,
  CF_STAGE1_CASES,
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_CASES,
  CF_STAGE3_CAPABILITIES,
  CF_STAGE3_UNSUPPORTED,
  CF_STAGE3_CASES,
  CF_STAGE3_NONDISCRIMINATING,
  CF_STAGE4_CAPABILITIES,
  CF_STAGE4_CASES,
  CF_STAGE4_NONDISCRIMINATING,
  CF_STAGE4_UNSUPPORTED,
  S4_ICON_FAMILIES,
  barCapabilityKind,
  iconCapabilityKind,
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
  buildIconOmReachabilityScript,
  buildIconSidecarExtras,
  parseIconOmReachability,
  buildBarSidecarExtras,
  buildDecorationProbeScript,
  buildTextProbeScript,
  PROBE_CHUNK_CELLS,
  probeRefChunks,
  captureAndCloseOwned,
  captureUnderLockGuard,
  staleCaptureSidecar,
  classifyBarCapture,
  parseDataBarSettingsReadback,
  parseExcelColorToArgb,
  verifyOnlyCaptureState,
  verifyOnlyCaseOrder,
} from '../scripts/cf-refs/capture'
import { hasEvidence, VERIFIED_CAPABILITIES, VERIFIED_STAGE1_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES } from '../src/xlsx/cf/evidence'
import { CF_MANIFEST_PATH, EXCEL_AUTHORED_ICONSETS_PATH, diffAuthoredIconTemplates, generateAllStage1, generateCaseArchive } from '../scripts/cf-refs/generate'
import {
  type CfObservedBar,
  type CfObservedIcon,
  type DecodedPng,
  buildIconClassHypotheses,
  buildIconGlyphKey,
  buildIconScenarioHypotheses,
  dialectAllAbsent,
  type IconClassHypothesis,
  classifyIconCrop,
  discriminateIcon,
  evaluateIconCapabilities,
  iconCropRect,
  iconTwinSetClass,
  iconTwinSetMembers,
  iconSidecarMalformed,
  iconGlyphFaintRecord,
  evaluateIconGeometryCapability,
  measureBarExtentInCell,
  replicationShortfall,
  ICON_GEOMETRY_RULES_VERSION,
  type IconGeometryCell,
  ICON_NO_ICON_MEMBER,
  measureCaseIcons,
  observedIconClass,
  type IconGlyphKey,
  loadIconGlyphKey,
  discriminateIconBox,
  measureIconGeometry,
  type ComputeGeometryOptions,
  type WorksheetRangeGeometry,
  computeWorksheetRangeGeometry,
  decodePngSync,
  BAR_PIXEL_PROBE_UNOBSERVABLE,
  discriminateBar,
  evaluateBarPixelProbe,
  isCapabilityVerified,
  isCellInSqref,
  isCellWithinRange,
  isSupportedExcelVersion,
  measureBarGeometry,
  measureCaseBars,
  observeCellsFromCapture,
  parseA1,
  readZipPartsSync,
  verifyCfCase,
} from '../scripts/cf-refs/verify'

import {
  CF_STAGE5_CAPABILITIES,
  CF_STAGE5_CASES,
  CF_STAGE5_FIXTURES,
  CF_STAGE5_GROUPS,
  CF_STAGE5_NONDISCRIMINATING,
  CF_STAGE5_REUSED_CAPTURES,
  CF_STAGE5_UNSUPPORTED,
  type S5Configuration,
} from '../scripts/cf-refs/cases'
import { auditStage5Manifest, CF_STAGE5_PRODUCTION_READINGS, stage5Publication, type Stage5AuditReport } from '../scripts/cf-refs/audit'
import { buildExplicitCalculationObservationScript, buildManualObservationScript, buildStage5CaptureScripts } from '../scripts/cf-refs/capture'
import { buildCaseBuildScripts } from '../scripts/cf-refs/build'
import { evaluateFormulaCapabilities, normalizeCfFormula, verifyMarkerInArchive, type FormulaObservation } from '../scripts/cf-refs/verify'

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
      // The per-cell probes run in chunks of PROBE_CHUNK_CELLS, as capture.ts runs them (osacompile overflows near 500).
      const scripts = [
        ...probeRefChunks(refs).flatMap(chunk => [
          buildTextProbeScript(`${c.id}.xlsx`, c.sheet, chunk),
          buildDecorationProbeScript(`${c.id}.xlsx`, c.sheet, chunk),
          buildColorProbeScript(`${c.id}.xlsx`, c.sheet, chunk),
        ]),
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

  it('probe chunks keep every cell in order and leave every case of up to 100 cells as one chunk', () => {
    expect(PROBE_CHUNK_CELLS).toBe(100)
    const refs = Array.from({ length: 510 }, (_, i) => `A${i + 1}`)
    const chunks = probeRefChunks(refs)
    expect(chunks.map(c => c.length)).toEqual([100, 100, 100, 100, 100, 10])
    expect(chunks.flat()).toEqual(refs)
    expect(probeRefChunks(refs.slice(0, 88))).toEqual([refs.slice(0, 88)])
    expect(probeRefChunks([])).toEqual([[]])
    // Every Stage 1-3 case before Task 22b fits one chunk, so its probes are byte-identical to the unchunked ones.
    for (const c of CF_CASES.filter(x => x.id !== 'stage3-bar-lengths-default' && x.stage < 4)) expect(c.expected.cells.length, c.id).toBeLessThanOrEqual(PROBE_CHUNK_CELLS)
    // Stage 4 postdates the chunking: exactly these cases span several chunks, and each still probes every cell in order.
    // (fix round 3: stage4-icon-degenerate lost its H column and now fits one chunk.)
    // (Task 28 round 2: stage4-icon-rows holds 64 designated cells and their 64 controls.)
    const STAGE4_MULTI_CHUNK = ['stage4-icon-glyph-key', 'stage4-icon-families', 'stage4-icon-thresholds', 'stage4-icon-reverse-custom', 'stage4-icon-rows']
    expect(CF_CASES.filter(x => x.stage === 4 && x.expected.cells.length > PROBE_CHUNK_CELLS).map(x => x.id)).toEqual(STAGE4_MULTI_CHUNK)
    for (const c of CF_CASES.filter(x => x.stage === 4)) {
      const refs = c.expected.cells.map(x => x.ref)
      expect(probeRefChunks(refs).flat(), c.id).toEqual(refs)
      expect(probeRefChunks(refs).every(chunk => chunk.length <= PROBE_CHUNK_CELLS), c.id).toBe(true)
    }
  })

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
    const mutated1000Xml = sheetXml.replace('<formula>100</formula>', '<formula>1000</formula>')
    // A counterexample must actually mutate the saved workbook (as 3c below asserts).
    expect(mutated1000Xml).not.toBe(sheetXml)
    z.file('xl/worksheets/sheet1.xml', mutated1000Xml)
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
    const mutatedBetweenXml = sheetXml.replace('<formula>3</formula><formula>1</formula>', '<formula>3</formula><formula>9</formula>')
    expect(mutatedBetweenXml).not.toBe(sheetXml)
    zBetween.file('xl/worksheets/sheet1.xml', mutatedBetweenXml)
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
    // A counterexample must actually mutate the saved workbook (Task 18c fix 1: Excel's resave renumbered the dxfIds,
    // so a literal search for dxfId="1" silently changed nothing).
    const absentRule = /<cfRule type="containsText" dxfId="(\d+)" priority="8" operator="containsText" text="ab"\/>/
    expect(textSheetXml).toMatch(absentRule)
    const mutatedAbsentXml = textSheetXml.replace(absentRule, (_m, dxfId: string) => `<cfRule type="containsText" dxfId="${dxfId}" priority="8" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH("ab",I2)))</formula></cfRule>`)
    expect(mutatedAbsentXml).not.toBe(textSheetXml)
    zText.file('xl/worksheets/sheet1.xml', mutatedAbsentXml)
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
    // Stage 4 (Task 24) appends its own registry after Stage 3's.
    expect(CF_CAPABILITIES.length).toBe(CF_STAGE1_CAPABILITIES.length + CF_STAGE2_CAPABILITIES.length + CF_STAGE3_CAPABILITIES.length + CF_STAGE4_CAPABILITIES.length + CF_STAGE5_CAPABILITIES.length)
    expect(new Set(CF_CAPABILITIES.map(c => c.capabilityId)).size).toBe(CF_CAPABILITIES.length)
    expect(CF_CASES.slice(0, 21).map(c => c.stage)).toEqual([...Array(13).fill(1), ...Array(8).fill(2)])

    expect(CF_STAGE3_CASES.map(c => c.id)).toEqual([
      'stage3-bar-settings-probe',
      'stage3-bar-lengths-default',
      'stage3-bar-auto-bounds',
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
      // Task 19b: the automatic colours Excel writes on resave (Ruling S3-F as amended), recipe-pending until captured.
      'stage3-bar-auto-colors',
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
    // Every Stage 3 capability is a bar:* capability with a ruling citation. Since Task 23 production holds exactly the
    // published Stage 3 set (the audit's verified result, asserted in tests/xlsx-cf-stage3.test.ts), and never a
    // non-discriminating, unsupported or unresolved one.
    const published = Reflect.get(verifiedExports, 'VERIFIED_STAGE3_CAPABILITIES') as ReadonlySet<string>
    expect(published.size).toBe(51)
    for (const entry of CF_STAGE3_CAPABILITIES) {
      expect(entry.capabilityId.startsWith('bar:')).toBe(true)
      expect(entry.failClosedDisposition).toMatch(/S3-[A-S]|Always unsupported/)
      expect(hasEvidence(entry.capabilityId), entry.capabilityId).toBe(published.has(entry.capabilityId))
      expect(VERIFIED_CAPABILITIES.has(entry.capabilityId), entry.capabilityId).toBe(published.has(entry.capabilityId))
      if (entry.capabilityId.startsWith('bar:unresolved-') || CF_STAGE3_UNSUPPORTED[entry.capabilityId] || CF_STAGE3_NONDISCRIMINATING[entry.capabilityId]) {
        expect(hasEvidence(entry.capabilityId), entry.capabilityId).toBe(false)
      }
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
  }, 60_000) // hashes every Stage 3 workbook/PNG and runs verifyCfCase on all 15 cases: ~2 s alone, 5.9 s on the CI runner

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
    // A static registry candidate is withheld and production holds no evidence for it. (bar:bounds-automax left the
    // registry in Task 22c: stage3-bar-auto-bounds now owns it with data that discriminates, S3-AUTO-BOUNDS.)
    expect(CF_STAGE3_NONDISCRIMINATING['bar:bounds-num']).toMatch(/num 20\/80/)
    expect(hasEvidence('bar:bounds-num')).toBe(false)
    expect(VERIFIED_CAPABILITIES.has('bar:bounds-num')).toBe(false)
    expect(CF_STAGE3_NONDISCRIMINATING['bar:bounds-automax']).toBeUndefined()
    // Verified by stage3-bar-auto-bounds and published by Task 23.
    expect(hasEvidence('bar:bounds-automax')).toBe(true)
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

  it('audit --write renders every stage from its own audit, whatever --stage, and is idempotent (Lesson 4)', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const committed = readFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
    const first = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: committed })
    const second = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: committed })
    expect(second).toBe(first)
    const listed = (text: string, name: string): string[] | undefined => {
      const m = new RegExp(`export const ${name}: ReadonlySet<string> = new Set\\(\\[([\\s\\S]*?)\\]\\)`).exec(text)
      return m ? [...m[1]!.matchAll(/'([^']+)'/g)].map(x => x[1]!) : undefined
    }
    // Task 18c fix 1: a Stage-1-only rendering dropped VERIFIED_STAGE2_CAPABILITIES and broke evidence.ts. Every
    // rendering now carries the Stage 1 and Stage 2 exports, each from its own audit, and they equal what is committed.
    expect(listed(first, 'VERIFIED_STAGE1_CAPABILITIES')).toEqual([...VERIFIED_STAGE1_CAPABILITIES].sort())
    expect(listed(first, 'VERIFIED_STAGE2_CAPABILITIES')).toEqual([...VERIFIED_STAGE2_CAPABILITIES].sort())
    // Ruling S3-WRITE: without --publish-stage3 the Stage 3 export stays as committed (none today), so the write
    // reproduces the committed file byte for byte.
    expect(listed(first, 'VERIFIED_STAGE3_CAPABILITIES')).toEqual(listed(committed, 'VERIFIED_STAGE3_CAPABILITIES'))
    expect(listed(first, 'VERIFIED_CAPABILITIES')).toEqual([...VERIFIED_CAPABILITIES].sort())
    expect(first).toBe(committed)
    // A committed Stage 3 export is carried over verbatim, never re-derived, without --publish-stage3.
    const withS3 = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { publishStage3: true })
    const reported = auditStage3Manifest(manifest, { rootDir: REPO_ROOT }).perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()
    expect(reported.length).toBeGreaterThan(0)
    expect(listed(withS3, 'VERIFIED_STAGE3_CAPABILITIES')).toEqual(reported)
    expect(listed(withS3, 'VERIFIED_CAPABILITIES')).toEqual([...new Set([...VERIFIED_STAGE1_CAPABILITIES, ...VERIFIED_STAGE2_CAPABILITIES, ...reported])].sort())
    const keptS3 = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: withS3.replace(`  '${reported[0]}',\n`, '') })
    expect(listed(keptS3, 'VERIFIED_STAGE3_CAPABILITIES')).toEqual(reported.slice(1))
    expect(renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { publishStage3: true })).toBe(withS3)

    const report = auditStage3Manifest(manifest, { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    // Capture run 8 (stage3-bar-auto-bounds) completes Stage 3 again: nothing is blocked or unaccounted.
    expect(report.stageComplete).toBe(true)
    expect(report.summary.cases).toBe(15)
    expect(report.summary.capabilities).toBe(CF_STAGE3_CAPABILITIES.length)
    expect(report.blockedCapabilities).toEqual([])
    expect(report.perCapability.filter(p => p.disposition === 'verified')).toHaveLength(51)
    expect(report.barAutoBounds).toEqual({ autoMin: { rule: 'zero-anchored', settledBy: ['bar:bounds-automin'] }, autoMax: { rule: 'zero-anchored', settledBy: ['bar:bounds-automax'] } })
    expect(report.perCapability.filter(p => p.disposition === 'refused-unsupported').every(p => p.capabilityId.startsWith('bar:unresolved-') || CF_STAGE3_UNSUPPORTED[p.capabilityId] !== undefined)).toBe(true)
  }, 90_000) // about ten whole-stage audits: 15–25 s alone since capture run 7 (a 510-cell case), up to 50 s beside the full parallel suite
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
    // Excel's default 10/90 is not 10/90 linear on the span (Task 18b); at this one width every other candidate rule
    // (5/85 or 10/90 minus 9 pt, on x14, solid or all bars) predicts the same px (Task 22b): non-discriminating here,
    // decided by stage3-bar-lengths-default.
    expect(evaluation['bar:lengths-default-10-90']).toMatchObject({ status: 'non-discriminating' })
    expect(labels.get('B3')).toEqual({ settings: 'unverified', geometry: 'non-discriminating' })

    const res = verifyCfCase(c, { rootDir: dir })
    expect(res, JSON.stringify(res).slice(0, 400)).toMatchObject({ kind: 'verified' })
    if (res.kind !== 'verified') return
    expect([...res.verifiedCapabilities].sort()).toEqual(['bar:bounds-automin', 'bar:bounds-min-max', 'bar:bounds-percentile', 'bar:lengths-all-negative', 'bar:lengths-explicit-0-100'])
    expect([...(res.nonDiscriminatingCapabilities ?? [])].sort()).toEqual(['bar:bounds-automax', 'bar:bounds-num', 'bar:bounds-percent', 'bar:lengths-default-10-90'])
    expect(res.refusedCapabilities ?? []).toEqual([])
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
    // Claiming the non-discriminating 10/90 family as pixel-measured refuses the case.
    expect(verifyCfCase(relabelCap(['B2', 'B3', 'B5', 'B6'], { settings: 'unverified', geometry: 'pixel-measured' }), { rootDir: dir })).toMatchObject({ kind: 'refusal', code: 'unobserved-expectation' })
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

  it('Excel\'s legacy resave of the Stage 1 x14 rules is accepted and a stripped or altered twin refuses (Task 18c)', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-x14-legacy-'))
    dirs.push(dir)
    for (const id of ['stage1-text-predicates', 'stage1-blanks-errors']) {
      const rejected = new Uint8Array(readFileSync(resolve(REPO_ROOT, `tests/fixtures/xlsx-cf/${id}.rejected.xlsx`)))
      const tpl = CF_CASES.find(x => x.id === id)!
      const run = (bytes: Uint8Array, name: string) => {
        writeFileSync(join(dir, name), bytes)
        const h = sha256(bytes)
        return verifyCfCase({ ...tpl, input: name, saved: name, captures: [], expected: { ...tpl.expected, inputSha256: h, savedSha256: h } }, { rootDir: dir })
      }
      // The real resave: every marker, the x14 ones included, is found as Excel's legacy rule.
      expect(run(rejected, `${id}.xlsx`), id).toMatchObject({ kind: 'refusal', code: 'semantic-evidence-pending', structuralValid: true })
      const zip = await JSZip.loadAsync(rejected)
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      const mutate = async (from: string, to: string, label: string) => {
        expect(sheet.includes(from), `${id} fixture has ${from}`).toBe(true)
        zip.file('xl/worksheets/sheet1.xml', sheet.replace(from, to))
        return run(new Uint8Array(await zip.generateAsync({ type: 'uint8array' })), `${id}-${label}.xlsx`)
      }
      if (id === 'stage1-text-predicates') {
        // The x14 canonical twin's formula altered, its rule stripped, and its DXF swapped all refuse.
        expect(await mutate('<formula>NOT(ISERROR(SEARCH("ab",K2)))</formula>', '<formula>NOT(ISERROR(SEARCH("ac",K2)))</formula>', 'formula')).toMatchObject({ kind: 'refusal', code: 'stripped-rule' })
        expect(await mutate(/<conditionalFormatting sqref="K2:K9">.*?<\/conditionalFormatting>/.exec(sheet)![0], '', 'stripped')).toMatchObject({ kind: 'refusal', code: 'stripped-rule' })
        expect(await mutate('<cfRule type="containsText" dxfId="1" priority="10"', '<cfRule type="containsText" dxfId="0" priority="10"', 'dxf')).toMatchObject({ kind: 'refusal', code: 'dxf-pairing-failed' })
      } else {
        expect(await mutate('<cfRule type="notContainsErrors" dxfId="0" priority="14">', '<cfRule type="notContainsErrors" dxfId="0" priority="15">', 'priority')).toMatchObject({ kind: 'refusal', code: 'stripped-rule' })
      }
    }
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
    const original = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    const xml = original.replace('<x14:id>{3B3B000B-0000-4000-8000-000000000002}</x14:id>', '<x14:id>{3B3B000B-0000-4000-8000-0000000000FF}</x14:id>')
    // The GUID swap must actually change the workbook, or the refusal below proves nothing.
    expect(xml).not.toBe(original)
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

describe('Stage 4 icon oracle (Task 24)', () => {
  const s4dirs: string[] = []
  afterEach(() => {
    for (const d of s4dirs.splice(0)) rmSync(d, { recursive: true, force: true })
  })
  const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
  const XM = 'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
  const x14Block = (rule: string, sqref = 'B2:B3'): string =>
    `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${X14}><x14:conditionalFormattings><x14:conditionalFormatting ${XM}>${rule}<xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`

  it('stage 4 cases join the manifest and stage 1–3 cases are unchanged', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    expect(manifest.cases.map(c => c.id)).toEqual(CF_CASES.map(c => c.id))
    // Stage 1-3 catalogues and capability ids are exactly what Stage 3 shipped.
    expect(CF_STAGE1_CAPABILITIES).toHaveLength(115)
    expect(CF_STAGE2_CAPABILITIES).toHaveLength(46)
    expect(CF_STAGE3_CAPABILITIES).toHaveLength(73)
    expect(CF_STAGE1_CASES).toHaveLength(13)
    expect(CF_STAGE2_CASES).toHaveLength(8)
    expect(CF_STAGE3_CASES).toHaveLength(15)
    expect(CF_CASES.slice(0, 36).map(c => c.id)).toEqual([...CF_STAGE1_CASES, ...CF_STAGE2_CASES, ...CF_STAGE3_CASES].map(c => c.id))
    expect(CF_CAPABILITIES.length).toBe(CF_STAGE1_CAPABILITIES.length + CF_STAGE2_CAPABILITIES.length + CF_STAGE3_CAPABILITIES.length + CF_STAGE4_CAPABILITIES.length + CF_STAGE5_CAPABILITIES.length)
    expect(new Set(CF_CAPABILITIES.map(c => c.capabilityId)).size).toBe(CF_CAPABILITIES.length)
    for (const c of manifest.cases.filter(x => x.stage !== 4)) expect(c.markers.some(m => m.icon !== undefined), c.id).toBe(false)

    // The brief's matrix, case by case: the capability ids Tasks 25-27 consume.
    const glyphCaps = S4_ICON_FAMILIES.flatMap(f => Array.from({ length: f.size }, (_, id) => `icon:glyph-${f.name}-${id}`))
    expect(glyphCaps).toHaveLength(75)
    const expected: Record<string, string[]> = {
      'stage4-icon-om-reachability': ['icon:om-reachability-probe'],
      'stage4-icon-authored': ['icon:dialect-base', 'icon:dialect-standalone'],
      'stage4-icon-dialect': ['icon:dialect-base-x14-family', 'icon:dialect-paired'],
      'stage4-icon-glyph-key': ['icon:glyph-key', 'icon:pixel-calibration', ...glyphCaps],
      'stage4-icon-families': ['icon:families-base', 'icon:families-x14', 'icon:scope-whole-column', 'icon:settings-schema-defaults'],
      'stage4-icon-thresholds': ['icon:thresholds-gte-false', 'icon:thresholds-gte-true', 'icon:thresholds-mixed', 'icon:thresholds-num', 'icon:thresholds-percent', 'icon:thresholds-percentile-inc-exc'],
      'stage4-icon-degenerate': ['icon:bounds-out-of-domain', 'icon:degenerate-equal-values', 'icon:degenerate-non-monotonic', 'icon:degenerate-single-cell', 'icon:thresholds-floor', 'icon:thresholds-min-max'],
      'stage4-icon-reverse-custom': ['icon:custom-mapping', 'icon:custom-noicons', 'icon:reverse'],
      // Fix round 8: stage4-icon-geometry-2 owns the box and placement gates and the bar/icon showValue conflict.
      'stage4-icon-geometry-2': ['icon:composition-bar-showvalue', 'icon:geometry-box', 'icon:geometry-vertical-align'],
      'stage4-icon-geometry': ['icon:geometry-narrow-clip', 'icon:geometry-short-row-clip', 'icon:geometry-snapping'],
      'stage4-icon-inset': ['icon:inset-center', 'icon:inset-left', 'icon:inset-overflow', 'icon:inset-right-aligned', 'icon:noicons-inset', 'icon:noicons-showvalue'],
      'stage4-icon-showvalue': ['icon:showvalue-canvas-suppression', 'icon:showvalue-copy-text'],
      'stage4-icon-participation': ['icon:participation-dateserial', 'icon:participation-error-presence', 'icon:participation-types'],
      'stage4-icon-scope': ['icon:scope-autofilter-hidden-row', 'icon:scope-disjoint', 'icon:scope-duplicate-blocks', 'icon:scope-hidden-row', 'icon:scope-merged-cell'],
      'stage4-icon-composition': ['icon:composition-dxf-fill', 'icon:composition-matching-stopIfTrue-ahead', 'icon:composition-noicons-slot', 'icon:composition-overlapping-icons', 'icon:composition-unresolved-stub', 'icon:composition-with-bar', 'icon:composition-with-scale'],
      'stage4-icon-rtl': ['icon:placement-rtl-sheet'],
      'stage4-icon-unresolved-shapes': ['icon:unresolved-active-present', 'icon:unresolved-base-x14-set', 'icon:unresolved-cfvo-count', 'icon:unresolved-cfvo-no-type', 'icon:unresolved-custom-count', 'icon:unresolved-formula-cfvo', 'icon:unresolved-iconid-range', 'icon:unresolved-noicons-id', 'icon:unresolved-unknown-set', 'icon:unresolved-xmf-reference', 'icon:attr-percent-false', 'icon:reverse-custom', 'icon:composition-icon-stopIfTrue'],
      // Task 28 round 4 (Ruling S4-BOTTOM-REPLICATE): placed ahead of stage4-icon-rows, it owns the row-offset capability;
      // stage4-icon-rows (round 2, S4-ROWS-CAPTURE) stays captured and recorded.
      'stage4-icon-rows-bottom': ['icon:geometry-row-offset'],
      'stage4-icon-rows': [],
    }
    expect(CF_STAGE4_CASES.map(c => c.id)).toEqual(Object.keys(expected))
    for (const [id, caps] of Object.entries(expected)) {
      expect(CF_STAGE4_CAPABILITIES.filter(e => e.caseId === id).map(e => e.capabilityId).sort(), id).toEqual([...caps].sort())
    }
    // The family table matches the viewer's ICON_SETS, read as source text (the oracle never imports the code under test).
    const parseSrc = readFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/parse.ts'), 'utf8')
    const viewerSets = new Map([...(/const ICON_SETS[^[]*\[([\s\S]*?)\]\)/.exec(parseSrc)?.[1] ?? '').matchAll(/\['([^']+)', (\d)\]/g)].map(m => [m[1]!, Number(m[2])]))
    expect(new Map(S4_ICON_FAMILIES.map(f => [f.name, f.size]))).toEqual(viewerSets)

    for (const c of manifest.cases.filter(x => x.stage === 4)) {
      expect(['recipe-pending', 'unsupported', 'measured']).toContain(c.status)
      expect(c.expected.inputSha256).toMatch(/^[0-9a-f]{64}$/)
      expect(c.expected.savedSha256).toMatch(/^[0-9a-f]{64}$/)
      if (c.resavedByExcel === true) expect(c.expected.savedSha256).toBe(sha256(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))))
      let hasSidecar = false
      for (const cap of c.captures) {
        if (cap.pngSha256 === undefined) continue
        expect(existsSync(resolve(REPO_ROOT, cap.pngPath)), cap.pngPath).toBe(true)
        expect(sha256(new Uint8Array(readFileSync(resolve(REPO_ROOT, cap.pngPath))))).toBe(cap.pngSha256)
        hasSidecar = true
      }
      for (const cell of c.expected.cells) {
        if (!cell.iconEvidence) continue
        expect(['pixel-classified', 'non-discriminating', 'unverified']).toContain(cell.iconEvidence.bucket)
        expect(['pixel-measured', 'non-discriminating', 'unverified']).toContain(cell.iconEvidence.geometry)
        if (cell.iconEvidence.bucket !== 'unverified' || cell.iconEvidence.geometry !== 'unverified') {
          expect(hasSidecar, `${c.id}:${cell.ref} claims icon evidence without a capture`).toBe(true)
        }
      }
      const res = verifyCfCase(c, { rootDir: REPO_ROOT })
      if (res.kind === 'refusal') {
        expect(res.structuralValid, `${c.id}: ${res.reason}`).toBe(true)
        expect(['semantic-evidence-pending', 'unsupported-capability']).toContain(res.code)
      } else {
        expect(c.status).toBe('measured')
      }
    }
    // Every Stage 4 capability is an icon:* id citing the S4 ruling it lifts, and production holds exactly the ones the
    // Task 30 publication gate publishes (none before it).
    const gated = new Set(stage4Publication(auditStage4Manifest(manifest, { rootDir: REPO_ROOT })).published)
    for (const entry of CF_STAGE4_CAPABILITIES) {
      expect(entry.capabilityId.startsWith('icon:')).toBe(true)
      expect(entry.failClosedDisposition, entry.capabilityId).toMatch(/S4-[A-R]|Always unsupported/)
      expect(hasEvidence(entry.capabilityId), entry.capabilityId).toBe(gated.has(entry.capabilityId))
      expect(VERIFIED_CAPABILITIES.has(entry.capabilityId), entry.capabilityId).toBe(gated.has(entry.capabilityId))
    }
    for (const id of Object.keys(CF_STAGE4_NONDISCRIMINATING)) expect(CF_STAGE4_CAPABILITIES.some(e => e.capabilityId === id), id).toBe(true)
    expect(CF_STAGE4_CAPABILITIES.filter(e => iconCapabilityKind(e.capabilityId) === 'record').map(e => e.capabilityId).sort()).toEqual(
      [...Object.keys(CF_STAGE4_NONDISCRIMINATING).filter(id => iconCapabilityKind(id) === 'record'), ...expected['stage4-icon-unresolved-shapes']!.filter(id => id.startsWith('icon:unresolved-'))].sort(),
    )
  }, 60_000)

  it('icon marker pins set, cfvo, gte, custom mapping and dialect exactly', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-icon-pin-'))
    s4dirs.push(dir)
    const part = 'xl/worksheets/sheet1.xml'
    const customMarker: CfPartMarker = {
      part, extension: 'x14', capabilityId: 'icon:custom-noicons', sqref: 'B2:B3', ruleType: 'iconSet', priority: 1, formulas: [],
      icon: {
        dialect: 'x14', set: '3Arrows',
        custom: [{ set: '3Stars', iconId: 0 }, { set: '3Flags', iconId: 2 }, { set: 'NoIcons', iconId: 0 }],
        cfvo: [{ type: 'percent', f: '0' }, { type: 'num', f: '20' }, { type: 'percent', f: '67' }],
        attrs: {},
      },
    }
    const baseMarker: CfPartMarker = {
      part, extension: 'legacy', capabilityId: 'icon:reverse', sqref: 'B2:B3', ruleType: 'iconSet', priority: 1, formulas: [],
      icon: { dialect: 'base', set: '4Arrows', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '25' }, { type: 'percent', val: '50', gte: false }, { type: 'percent', val: '75' }], attrs: { reverse: '1' } },
    }
    const keyMarker: CfPartMarker = {
      part, extension: 'x14', capabilityId: 'icon:glyph-3Arrows-1', sqref: 'B2', ruleType: 'iconSet', priority: 1, formulas: [],
      icon: { dialect: 'x14', set: '3Arrows', custom: [0, 1, 2].map(() => ({ set: '3Arrows', iconId: 1 })), cfvo: [{ type: 'num', f: '0' }, { type: 'num', f: '1' }, { type: 'num', f: '2' }], attrs: { percent: '(omitted)' } },
    }
    const customXml = (o: { setAttrs?: string; cfvo?: string; icons?: string } = {}): string =>
      x14Block(`<x14:cfRule type="iconSet" priority="1" id="{3B3B4F01-0000-4000-8000-000000000001}"><x14:iconSet${o.setAttrs ?? ' iconSet="3Arrows" custom="1"'}>${o.cfvo ?? '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>'}${o.icons ?? '<x14:cfIcon iconSet="3Stars" iconId="0"/><x14:cfIcon iconSet="3Flags" iconId="2"/><x14:cfIcon iconSet="NoIcons" iconId="0"/>'}</x14:iconSet></x14:cfRule>`)
    const baseXml = (o: { setAttrs?: string; cfvo?: string } = {}): string =>
      `<conditionalFormatting sqref="B2:B3"><cfRule type="iconSet" priority="1"><iconSet${o.setAttrs ?? ' iconSet="4Arrows" reverse="1"'}>${o.cfvo ?? '<cfvo type="percent" val="0"/><cfvo type="percent" val="25"/><cfvo type="percent" val="50" gte="0"/><cfvo type="percent" val="75"/>'}</iconSet></cfRule></conditionalFormatting>`
    const keyXml = (setAttrs = ' iconSet="3Arrows" custom="1"'): string =>
      x14Block(`<x14:cfRule type="iconSet" priority="1" id="{3B3B4F03-0000-4000-8000-000000000001}"><x14:iconSet${setAttrs}><x14:cfvo type="num"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>1</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>2</xm:f></x14:cfvo><x14:cfIcon iconSet="3Arrows" iconId="1"/><x14:cfIcon iconSet="3Arrows" iconId="1"/><x14:cfIcon iconSet="3Arrows" iconId="1"/></x14:iconSet></x14:cfRule>`, 'B2')
    let n = 0
    const run = async (xml: string, marker: CfPartMarker) => {
      const bytes = await makeSyntheticXlsx({ sheetRuleXml: xml, dxfCount: 0 })
      const name = `icon-${n++}.xlsx`
      writeFileSync(join(dir, name), bytes)
      const hash = sha256(bytes)
      const c: CfOracleCase = {
        id: 'icon-pin-case', stage: 4, question: 'Pins exact icon set semantics', input: name, saved: name, sheet: 'Sheet1', range: 'A1:B3',
        requiredAppearances: ['screen'], markers: [marker],
        expected: {
          inputSha256: hash, savedSha256: hash, provenance: 'seed-hypothesis', positiveRefs: ['B2'], negativeRefs: ['A2'],
          cells: [
            { ref: 'B2', controlRef: 'A2', matched: true, iconApplied: true, fillSource: 'unverified' },
            { ref: 'A2', controlRef: 'B2', matched: false, iconApplied: false, fillSource: 'unverified' },
          ],
        },
        captures: [], status: 'recipe-pending', reason: 'Testing icon marker pinning', tolerance: { meanAbs: 0.015, rowMax: 0.08, strong: 12 },
      }
      return verifyCfCase(c, { rootDir: dir })
    }
    const pending = { kind: 'refusal', code: 'semantic-evidence-pending', structuralValid: true }
    const stripped = { kind: 'refusal', code: 'stripped-rule', structuralValid: false }
    // Canonical archives verify structurally.
    expect(await run(customXml(), customMarker)).toMatchObject(pending)
    expect(await run(baseXml(), baseMarker)).toMatchObject(pending)
    expect(await run(keyXml(), keyMarker)).toMatchObject(pending)

    const mutations: Array<[string, string, CfPartMarker]> = [
      ['gte true -> false on cfvo[2]', customXml({ cfvo: '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo><x14:cfvo type="percent" gte="0"><xm:f>67</xm:f></x14:cfvo>' }), customMarker],
      ['reordered cfvo', customXml({ cfvo: '<x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>' }), customMarker],
      ['cfIcon[1] iconId 2 -> 1', customXml({ icons: '<x14:cfIcon iconSet="3Stars" iconId="0"/><x14:cfIcon iconSet="3Flags" iconId="1"/><x14:cfIcon iconSet="NoIcons" iconId="0"/>' }), customMarker],
      ['NoIcons -> 3Arrows:0', customXml({ icons: '<x14:cfIcon iconSet="3Stars" iconId="0"/><x14:cfIcon iconSet="3Flags" iconId="2"/><x14:cfIcon iconSet="3Arrows" iconId="0"/>' }), customMarker],
      ['reverse flipped', baseXml({ setAttrs: ' iconSet="4Arrows"' }), baseMarker],
      ['showValue flipped', baseXml({ setAttrs: ' iconSet="4Arrows" reverse="1" showValue="0"' }), baseMarker],
      ['xm:f "20" -> "2"', customXml({ cfvo: '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>2</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>' }), customMarker],
      ['standalone x14 moved to the base dialect', '<conditionalFormatting sqref="B2:B3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows" custom="1"><cfvo type="percent" val="0"/><cfvo type="num" val="20"/><cfvo type="percent" val="67"/><cfIcon iconSet="3Stars" iconId="0"/><cfIcon iconSet="3Flags" iconId="2"/><cfIcon iconSet="NoIcons" iconId="0"/></iconSet></cfRule></conditionalFormatting>', customMarker],
      ['glyph key percent form changed from the authored (omitted) form', keyXml(' iconSet="3Arrows" custom="1" percent="1"'), keyMarker],
      ['set name changed', baseXml({ setAttrs: ' iconSet="4ArrowsGray" reverse="1"' }), baseMarker],
      ['custom flag dropped', customXml({ setAttrs: ' iconSet="3Arrows"' }), customMarker],
    ]
    for (const [label, xml, marker] of mutations) expect(await run(xml, marker), label).toMatchObject(stripped)

    // Semantics, not raw strings: attribute order, explicit schema defaults and numeric spelling all verify.
    expect(await run(baseXml({ setAttrs: ' reverse="true" showValue="1" percent="1" iconSet="4Arrows"', cfvo: '<cfvo type="percent" val="0" gte="1"/><cfvo type="percent" val="25.0"/><cfvo type="percent" val="50" gte="false"/><cfvo type="percent" val="75"/>' }), baseMarker)).toMatchObject(pending)
    // An omitted iconSet is 3TrafficLights1, and only that.
    const defaultsMarker: CfPartMarker = { ...baseMarker, icon: { dialect: 'base', set: '3TrafficLights1', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }], attrs: {} } }
    const defaultsXml = (setAttrs: string) => `<conditionalFormatting sqref="B2:B3"><cfRule type="iconSet" priority="1"><iconSet${setAttrs}><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>`
    expect(await run(defaultsXml(''), defaultsMarker)).toMatchObject(pending)
    expect(await run(defaultsXml(' iconSet="3TrafficLights1"'), defaultsMarker)).toMatchObject(pending)
    expect(await run(defaultsXml(' iconSet="3Arrows"'), defaultsMarker)).toMatchObject(stripped)
    // The key's form pin also refuses an explicit percent="1" when the marker pins an omitted form, and accepts the
    // written form only when that is what it pins.
    expect(await run(keyXml(' iconSet="3Arrows" custom="1" percent="1"'), { ...keyMarker, icon: { ...keyMarker.icon!, attrs: { percent: '1' } } })).toMatchObject(pending)
    expect(await run(keyXml(), { ...keyMarker, icon: { ...keyMarker.icon!, attrs: { percent: '1' } } })).toMatchObject(stripped)
  })

  /** A synthetic capture: white background painted by `paint`. */
  const paintPng = (w: number, h: number, paint: (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>) => void): DecodedPng => {
    const canvas = createCanvas(w, h)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, w, h)
    paint(ctx)
    return decodePngSync(new Uint8Array(canvas.toBuffer('image/png')))!
  }
  const COLOURS = ['#ff0000', '#00a000', '#0000ff', '#ffc000', '#a000a0', '#00a0a0', '#808080', '#000000', '#ff8080']

  it('glyph key refuses a mismatched or ambiguous crop', async () => {
    // A 3x3 key of distinct 16x16 crops (nine glyph classes, each a differently coloured and placed square).
    const crop = (i: number): DecodedPng => paintPng(16, 16, ctx => { ctx.fillStyle = COLOURS[i]!; ctx.fillRect(2 + (i % 3), 2 + Math.floor(i / 3), 9, 9) })
    const key = new Map(Array.from({ length: 9 }, (_, i) => [`class-${i}`, crop(i)] as const))
    // A capture holding a pasted copy of crop 2 at (40, 10).
    const pasted = paintPng(80, 40, ctx => { ctx.fillStyle = COLOURS[2]!; ctx.fillRect(40 + 2 + 2, 10 + 2, 9, 9) })
    const at = { x: 40, y: 10, width: 16, height: 16 }
    expect(classifyIconCrop(pasted, at, key, { tolerance: 8 })).toEqual({ glyphClass: 'class-2' })
    // Small antialiasing noise (within +-8 per channel) still classifies.
    const noisy = paintPng(80, 40, ctx => { ctx.fillStyle = COLOURS[2]!; ctx.fillRect(44, 12, 9, 9); ctx.fillStyle = '#f9f9f9'; ctx.fillRect(41, 11, 1, 1) })
    expect(classifyIconCrop(noisy, at, key, { tolerance: 8 })).toEqual({ glyphClass: 'class-2' })
    // One strong pixel away from every key crop: no class.
    const off = paintPng(80, 40, ctx => { ctx.fillStyle = COLOURS[2]!; ctx.fillRect(44, 12, 9, 9); ctx.fillStyle = '#000000'; ctx.fillRect(41, 24, 1, 1) })
    expect(classifyIconCrop(off, at, key, { tolerance: 8 })).toMatchObject({ kind: 'refusal', code: 'icon-class-none' })
    // Two key entries identical to the crop under different glyph classes: ambiguous, never a pick.
    const twin = new Map([...key, ['class-2-twin', crop(2)] as const])
    expect(classifyIconCrop(pasted, at, twin, { tolerance: 8 })).toMatchObject({ kind: 'refusal', code: 'icon-class-ambiguous' })
    // A key entry that differs from the match only within the strong delta (no pixel >= 64) is not separated: ambiguous.
    const faint = paintPng(16, 16, ctx => { ctx.fillStyle = COLOURS[2]!; ctx.fillRect(4, 2, 9, 9); ctx.fillStyle = '#d0d0d0'; ctx.fillRect(0, 15, 1, 1) })
    expect(classifyIconCrop(pasted, at, new Map([...key, ['class-faint', faint] as const]), { tolerance: 8 })).toMatchObject({ kind: 'refusal', code: 'icon-class-ambiguous' })
    // A crop outside the capture refuses.
    expect(classifyIconCrop(pasted, { x: 70, y: 30, width: 16, height: 16 }, key, { tolerance: 8 })).toMatchObject({ kind: 'refusal', code: 'icon-class-none' })

    // Building the key from a capture: a key cell whose resaved XML lost its rule refuses with icon-key-mismatch.
    const keyCase = CF_STAGE4_CASES.find(c => c.id === 'stage4-icon-glyph-key')!
    const savedParts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, keyCase.saved))))!
    const geom = computeWorksheetRangeGeometry(savedParts, keyCase.range)!
    const blank = paintPng(geom.expectedWidthPt, geom.expectedHeightPt, () => {})
    const sheet = new TextDecoder().decode(savedParts.get('xl/worksheets/sheet1.xml')!)
    const lost = new Map(savedParts)
    lost.set('xl/worksheets/sheet1.xml', new TextEncoder().encode(sheet.replace(/<x14:conditionalFormatting [^>]*><x14:cfRule type="iconSet" priority="2" [\s\S]*?<\/x14:conditionalFormatting>/, '')))
    expect(buildIconGlyphKey(keyCase, lost, blank, geom)).toMatchObject({ kind: 'refusal', code: 'icon-key-mismatch' })
    // With every rule present, a capture that shows no icon in a key cell is not a key either.
    expect(buildIconGlyphKey(keyCase, savedParts, blank, geom)).toMatchObject({ kind: 'refusal', code: 'icon-key-mismatch' })
  })

  /** A 120x60 capture of A1:B3 (col A 60 px, col B 60 px, rows 20 px). */
  const iconGeom: WorksheetRangeGeometry = {
    minRow: 0, maxRow: 2, minCol: 0, maxCol: 1,
    rowHeightsPt: [20, 20, 20], colWidthChars: [8, 8], colWidthsPt: [60, 60],
    expectedHeightPt: 60, expectedWidthPt: 120,
  }
  const textAt = (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>, x: number, y: number) => {
    // A fixed glyph pattern (three strokes) standing in for value text.
    ctx.fillStyle = '#000000'
    ctx.fillRect(x, y, 2, 8); ctx.fillRect(x + 4, y, 2, 8); ctx.fillRect(x + 8, y + 2, 3, 6)
  }

  it('icon edges and text ink are located by scan within ±1 px', () => {
    const opts = { controlOf: (ref: string) => (ref === 'A1' ? 'A3' : 'B3'), tolerancePx: 1 }
    // A 16x16 icon at px 3..18 in A1 (cell 0..59); the control A3 is blank (number format ;;;).
    const icon16 = paintPng(120, 60, ctx => { ctx.fillStyle = '#00a000'; ctx.fillRect(3, 2, 16, 16) })
    const res = measureIconGeometry(icon16, iconGeom, ['A1'], opts)
    expect(res).toBeInstanceOf(Map)
    const a1 = (res as ReadonlyMap<string, CfObservedIcon>).get('A1')!
    expect(a1).toMatchObject({ ref: 'A1', cellLeftPx: 0, cellTopPx: 0, cellWidthPx: 60, cellHeightPx: 20, icon: { leftPx: 3, rightPx: 18, topPx: 2, bottomPx: 17 } })
    expect(a1.textInkLeftPx).toBeUndefined()
    // Text ink: the icon cell's text starts at 24, the control's at 4 (B1 is 60..119, its control B3): a shift of 20.
    const inset = paintPng(120, 60, ctx => {
      ctx.fillStyle = '#00a000'; ctx.fillRect(63, 2, 16, 16)
      textAt(ctx, 60 + 24, 6)
      textAt(ctx, 60 + 4, 46)
    })
    const b1 = (measureIconGeometry(inset, iconGeom, ['B1'], opts) as ReadonlyMap<string, CfObservedIcon>).get('B1')!
    expect(b1.icon).toMatchObject({ leftPx: 63, rightPx: 78 })
    expect(b1.textInkLeftPx).toBe(84)
    expect(b1.textShiftPx).toBe(20)
    // showValue="0": the control shows the text, the icon cell does not.
    const hidden = paintPng(120, 60, ctx => { ctx.fillStyle = '#00a000'; ctx.fillRect(63, 2, 16, 16); textAt(ctx, 64, 46) })
    const h1 = (measureIconGeometry(hidden, iconGeom, ['B1'], opts) as ReadonlyMap<string, CfObservedIcon>).get('B1')!
    expect(h1.icon).toMatchObject({ leftPx: 63, rightPx: 78 })
    expect(h1.textInkLeftPx).toBeUndefined()
    // No icon: the cell shows exactly its control's text at the same offset.
    const plain = paintPng(120, 60, ctx => { textAt(ctx, 64, 6); textAt(ctx, 64, 46) })
    const p1 = (measureIconGeometry(plain, iconGeom, ['B1'], opts) as ReadonlyMap<string, CfObservedIcon>).get('B1')!
    expect(p1.icon).toBeNull()
    expect(p1.textShiftPx).toBe(0)
    // Scanlines (25/50/75 percent of the icon's height) disagreeing by 2 px refuse.
    const unstable = paintPng(120, 60, ctx => {
      ctx.fillStyle = '#00a000'
      ctx.fillRect(3, 2, 16, 6); ctx.fillRect(5, 8, 14, 5); ctx.fillRect(3, 13, 16, 5)
    })
    expect(measureIconGeometry(unstable, iconGeom, ['A1'], opts)).toMatchObject({ kind: 'refusal', code: 'icon-edge-unstable' })
    // 1 px disagreement is inside the tolerance.
    const within = paintPng(120, 60, ctx => {
      ctx.fillStyle = '#00a000'
      ctx.fillRect(3, 2, 16, 6); ctx.fillRect(4, 8, 15, 5); ctx.fillRect(3, 13, 16, 5)
    })
    expect(measureIconGeometry(within, iconGeom, ['A1'], opts)).toBeInstanceOf(Map)
    // A non-rectangular glyph is stable against its own glyph-key crop (same shape, rigidly placed).
    const triangle = (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>, x: number, y: number) => {
      ctx.fillStyle = '#00a000'
      for (let r = 0; r < 16; r++) ctx.fillRect(x + 8 - Math.floor(r / 2), y + r, 2 * Math.floor(r / 2) + 1, 1)
    }
    const tri = paintPng(120, 60, ctx => triangle(ctx, 3, 2))
    expect(measureIconGeometry(tri, iconGeom, ['A1'], opts)).toMatchObject({ kind: 'refusal', code: 'icon-edge-unstable' })
    const reference = paintPng(16, 16, ctx => triangle(ctx, 0, 0))
    const triRes = measureIconGeometry(tri, iconGeom, ['A1'], { ...opts, referenceOf: () => reference })
    // Rows r = 0..15 span x 3+8-floor(r/2) .. 3+8+floor(r/2): the widest rows give the box 4..18.
    expect((triRes as ReadonlyMap<string, CfObservedIcon>).get('A1')!.icon).toMatchObject({ leftPx: 4, rightPx: 18, topPx: 2, bottomPx: 17 })
    // In record mode (the capture path) an unstable glyph is a per-cell finding, never a whole-case refusal.
    const recorded = (measureIconGeometry(tri, iconGeom, ['A1'], { ...opts, unstable: 'record' }) as ReadonlyMap<string, CfObservedIcon>).get('A1')!
    expect(recorded).toMatchObject({ edgeStable: false, icon: { leftPx: 4, rightPx: 18 } })
    expect((triRes as ReadonlyMap<string, CfObservedIcon>).get('A1')!.edgeStable).toBeUndefined()
    // An unknown cell refuses instead of inventing a measurement.
    expect(measureIconGeometry(icon16, iconGeom, ['Z9'], opts)).toMatchObject({ kind: 'refusal' })
  })

  it('non-discriminating icon variant keeps its gate closed (Lesson 3)', () => {
    // Production evidence for an id exists only because its owning Excel capture discriminates it (Task 30 publishes the
    // gated audit result); these synthetic non-discriminating runs never lift a gate on their own.
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const ownDiscriminates = (id: string): boolean => {
      const owner = CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === id)!.caseId
      const res = verifyCfCase(manifest.cases.find(c => c.id === owner)!, { rootDir: REPO_ROOT })
      return res.kind === 'verified' && res.verifiedCapabilities.includes(id) && !(res.nonDiscriminatingCapabilities ?? []).includes(id) && res.matchedHypotheses?.[id] !== undefined
    }
    const classOf = (set: string, iconId: number): string | null => (set === 'NoIcons' ? null : `${set}:${iconId}`)
    // gte: num 0/20/40 over off-boundary values only. "gte" and "gt" predict the same bucket everywhere.
    const gteRule = { set: '3Arrows', cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '20', gte: true }, { type: 'num', val: '40', gte: true }], attrs: {} }
    const offBoundary = new Map([['B2', 5], ['B3', 30], ['B4', 75]])
    const gteHyps = buildIconClassHypotheses('icon:thresholds-gte-true', gteRule, offBoundary, classOf)!
    expect(discriminateIcon(gteHyps, { B2: '3Arrows:0', B3: '3Arrows:1', B4: '3Arrows:2' })).toMatchObject({ status: 'non-discriminating' })
    expect(hasEvidence('icon:thresholds-gte-true')).toBe(ownDiscriminates('icon:thresholds-gte-true'))
    // The exact-boundary cell 20 discriminates gte from gt.
    const boundary = new Map([...offBoundary, ['B5', 20]])
    expect(discriminateIcon(buildIconClassHypotheses('icon:thresholds-gte-true', gteRule, boundary, classOf)!, { B2: '3Arrows:0', B3: '3Arrows:1', B4: '3Arrows:2', B5: '3Arrows:1' })).toEqual({ status: 'discriminating', matched: 'as-written' })
    // ... and refuses when the capture contradicts every hypothesis.
    expect(discriminateIcon(buildIconClassHypotheses('icon:thresholds-gte-true', gteRule, boundary, classOf)!, { B2: '3Arrows:2', B3: '3Arrows:1', B4: '3Arrows:2', B5: '3Arrows:1' })).toMatchObject({ status: 'none' })
    // reverse judged only on the middle bucket of a 3-icon set: reversed and not reversed agree.
    const revRule = { set: '3Arrows', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }], attrs: { reverse: '1' } }
    const scope = new Map([['C2', 0], ['C3', 50], ['C4', 100]])
    const middleOnly = buildIconClassHypotheses('icon:reverse', revRule, scope, classOf, ['C3'])!
    expect(Object.keys(middleOnly)).toEqual(['C3'])
    expect(discriminateIcon(middleOnly, { C3: '3Arrows:1' })).toMatchObject({ status: 'non-discriminating' })
    expect(hasEvidence('icon:reverse')).toBe(ownDiscriminates('icon:reverse'))
    // The end buckets discriminate.
    expect(discriminateIcon(buildIconClassHypotheses('icon:reverse', revRule, scope, classOf, ['C2', 'C4'])!, { C2: '3Arrows:2', C4: '3Arrows:0' })).toEqual({ status: 'discriminating', matched: 'reversed' })
    // A custom set whose end icons are one glyph class is non-discriminating for reverse-custom.
    const sameEnds = { set: '3Arrows', cfvo: revRule.cfvo, attrs: { reverse: '1' }, custom: [{ set: '3Flags', iconId: 0 }, { set: '3Flags', iconId: 1 }, { set: '3Flags', iconId: 0 }] }
    expect(discriminateIcon(buildIconClassHypotheses('icon:reverse-custom', sameEnds, scope, classOf, ['C2', 'C4'])!, { C2: '3Flags:0', C4: '3Flags:0' })).toMatchObject({ status: 'non-discriminating' })

    // Box hypotheses 2 px apart at every designated cell cannot be told apart.
    const measured = { B2: { leftPx: 3, rightPx: 18, topPx: 2, bottomPx: 17 } }
    const close = { B2: [{ id: 'fixed', leftPx: 3, rightPx: 18, topPx: 2, bottomPx: 17 }, { id: 'font', leftPx: 3, rightPx: 20, topPx: 2, bottomPx: 17 }] }
    expect(discriminateIconBox(close, measured, 1)).toMatchObject({ status: 'non-discriminating' })
    expect(hasEvidence('icon:geometry-box')).toBe(ownDiscriminates('icon:geometry-box'))
    // 3 px apart with exactly one hypothesis matching: the capability is verified.
    const apart = { B2: [{ id: 'fixed', leftPx: 3, rightPx: 18, topPx: 2, bottomPx: 17 }, { id: 'font', leftPx: 3, rightPx: 21, topPx: 2, bottomPx: 17 }] }
    expect(discriminateIconBox(apart, measured, 1)).toEqual({ status: 'discriminating', matched: 'fixed' })
    expect(discriminateIconBox(apart, { B2: { leftPx: 3, rightPx: 25, topPx: 2, bottomPx: 17 } }, 1)).toMatchObject({ status: 'none' })
  })

  it('audit --stage 4 --write is idempotent (Lesson 4)', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const committed = readFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
    const listed = (text: string, name: string): string[] | undefined => {
      const m = new RegExp(`export const ${name}: ReadonlySet<string> = new Set\\(\\[([\\s\\S]*?)\\]\\)`).exec(text)
      return m ? [...m[1]!.matchAll(/'([^']+)'/g)].map(x => x[1]!) : undefined
    }
    // Two runs, byte-identical; without --publish-stage4 the Stage 4 export is carried over verbatim (published by
    // Task 30), so the write reproduces the committed file and the Stage 1-4 sets are unchanged.
    const first = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: committed })
    expect(renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: committed })).toBe(first)
    expect(first).toBe(committed)
    for (const name of ['VERIFIED_STAGE1_CAPABILITIES', 'VERIFIED_STAGE2_CAPABILITIES', 'VERIFIED_STAGE3_CAPABILITIES', 'VERIFIED_STAGE4_CAPABILITIES', 'VERIFIED_CAPABILITIES']) {
      expect(listed(first, name), name).toEqual(listed(committed, name))
    }
    // A committed Stage 4 export is carried over verbatim, never re-derived.
    const unpublished = committed
      .replace(/export const VERIFIED_STAGE4_CAPABILITIES: [^\n]*\n[\s\S]*?\n\]\)\n\n/, '')
      .replace(/export const VERIFIED_STAGE4_ICON_SELECTIONS: [^\n]*\n[\s\S]*?\n\}\n\n/, '')
    expect(listed(unpublished, 'VERIFIED_STAGE4_CAPABILITIES')).toBeUndefined()
    expect(unpublished).not.toContain('VERIFIED_STAGE4_ICON_SELECTIONS')
    const withS4 = unpublished.replace(/(export const VERIFIED_CAPABILITIES)/, "export const VERIFIED_STAGE4_CAPABILITIES: ReadonlySet<string> = new Set([\n  'icon:pixel-calibration',\n])\n\n$1")
    const kept = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: withS4 })
    expect(listed(kept, 'VERIFIED_STAGE4_CAPABILITIES')).toEqual(['icon:pixel-calibration'])
    expect(renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: kept })).toBe(kept)
    // --publish-stage4 publishes the audit's own verified set, idempotently (fix round 6: evidence-driven, so it holds for
    // whatever the committed captures verify; publishing itself is Task 30 and is only rendered in memory here).
    const report = auditStage4Manifest(manifest, { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    expect(report.summary.cases).toBe(19)
    expect(report.summary.capabilities).toBe(CF_STAGE4_CAPABILITIES.length)
    expect(report.perCapability.filter(p => p.disposition === 'unaccounted')).toEqual([])
    expect(report.stageComplete).toBe(report.blockedCapabilities.length === 0)
    // Refused-unsupported is exactly the unresolved-shapes case (its own shapes plus those Excel canonicalises).
    const unsupportedCases = new Set(CF_STAGE4_CASES.filter(c => c.status === 'unsupported').map(c => c.id))
    expect(report.perCapability.filter(p => p.disposition === 'refused-unsupported').map(p => p.capabilityId).sort()).toEqual(CF_STAGE4_CAPABILITIES.filter(e => unsupportedCases.has(e.caseId) || CF_STAGE4_UNSUPPORTED[e.capabilityId] !== undefined).map(e => e.capabilityId).sort())
    const verified = report.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId)
    // Every verified id is backed by its owning case's own verification and the cross-case gates.
    const byCase = new Map<string, ReturnType<typeof verifyCfCase>>()
    for (const id of verified) {
      const owner = CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === id)!.caseId
      if (!byCase.has(owner)) byCase.set(owner, verifyCfCase(manifest.cases.find(c => c.id === owner)!, { rootDir: REPO_ROOT }))
      const res = byCase.get(owner)!
      expect(res.kind === 'verified' && res.verifiedCapabilities.includes(id), id).toBe(true)
      expect(isCapabilityVerified(id, manifest, { rootDir: REPO_ROOT }), id).toBe(true)
    }
    // Non-vacuous in both directions: the glyph-key gate decides whether anything can verify at all.
    const keyMeasured = manifest.cases.find(c => c.id === 'stage4-icon-glyph-key')!.status === 'measured'
    if (keyMeasured) expect(verified).toEqual(expect.arrayContaining(['icon:pixel-calibration', 'icon:glyph-key']))
    else expect(verified).toEqual([])
    expect(verified.length + report.blockedCapabilities.length).toBeGreaterThan(50)
    // Task 30: --publish-stage4 publishes the gated subset (`stage4Publication`) and its selections, idempotently; the
    // ruling deny-list is verified yet absent.
    const gate = stage4Publication(report)
    const published = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: committed, publishStage4: true })
    expect(listed(published, 'VERIFIED_STAGE4_CAPABILITIES')).toEqual(gate.published)
    for (const id of Object.keys(CF_STAGE4_NEVER_PUBLISH)) {
      expect(verified, id).toContain(id)
      expect(listed(published, 'VERIFIED_STAGE4_CAPABILITIES'), id).not.toContain(id)
    }
    expect(published).toContain(renderIconSelectionsBlock(gate.selections))
    expect(renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: published, publishStage4: true })).toBe(published)
    // Without the flag a committed publication (set and selections) is carried over verbatim.
    expect(renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: published })).toBe(published)
    expect(listed(published, 'VERIFIED_STAGE3_CAPABILITIES')).toEqual(listed(committed, 'VERIFIED_STAGE3_CAPABILITIES'))
    // The union is the Stage 1-3 sets plus exactly the published Stage 4 set.
    const s123 = ['VERIFIED_STAGE1_CAPABILITIES', 'VERIFIED_STAGE2_CAPABILITIES', 'VERIFIED_STAGE3_CAPABILITIES'].flatMap(n => listed(committed, n) ?? [])
    expect(listed(published, 'VERIFIED_CAPABILITIES')).toEqual([...new Set([...s123, ...gate.published])].sort())
  }, 120_000)

  it('builds the icon object-model reachability probe with every read guarded', () => {
    const lines = buildIconOmReachabilityScript('stage4-icon-om-reachability.xlsx', 'Sheet1')
    const text = lines.join('\n')
    expect(text).toContain('workbook "stage4-icon-om-reachability.xlsx"')
    for (const read of ['format condition 1 of', 'icon criteria', 'reverse icon set order', 'show icon only', 'properties']) expect(text, read).toContain(read)
    // Every read sits in its own try, so one unreachable class records an error instead of aborting the probe.
    expect((text.match(/\btry\b/g) ?? []).length).toBeGreaterThanOrEqual(5)
    expect(parseIconOmReachability('ok:::1|||error:::-1728 Can\'t get icon criteria|||missing value')).toEqual([
      { status: 'ok', detail: '1' },
      { status: 'error', detail: '-1728 Can\'t get icon criteria' },
    ])
  })

  it.skipIf(!hasOsacompile)('compiles the icon object-model reachability probe under osacompile (Lesson 8)', () => {
    const full = ['tell application "Microsoft Excel"', ...buildIconOmReachabilityScript('stage4-icon-om-reachability.xlsx', 'Sheet1'), 'end tell'].join('\n')
    const res = spawnSync('osacompile', ['-o', '/dev/null'], { input: full, encoding: 'utf8' })
    expect(res.status, res.stderr).toBe(0)
  })

  it('stage4ObservedIcons reads the recorded px and glyph classes and ignores malformed entries', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-icon-observed-'))
    s4dirs.push(dir)
    const c = { ...CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-thresholds')!, captures: [{ appearance: 'screen' as const, pngPath: 'x.png', sidecarPath: 'x.json' }] }
    expect(stage4ObservedIcons(c, dir)).toBeUndefined()
    writeFileSync(join(dir, 'x.json'), JSON.stringify({ observedIcons: [
      { ref: 'b2', cellLeftPx: 0, cellTopPx: 0, cellWidthPx: 60, cellHeightPx: 20, icon: { leftPx: 3, topPx: 2, rightPx: 18, bottomPx: 17, glyphClass: '3Arrows:0' } },
      { ref: 'B3', cellLeftPx: 0, cellTopPx: 20, cellWidthPx: 60, cellHeightPx: 20, icon: null, textInkLeftPx: 40 },
      { ref: 'B4', cellLeftPx: 0, cellTopPx: 40, cellWidthPx: 60, cellHeightPx: 20, icon: { leftPx: 3, rightPx: 18 } },
      { ref: 'B5', cellLeftPx: 'x' },
    ] }))
    const got = stage4ObservedIcons(c, dir)!
    expect([...got.keys()]).toEqual(['B2', 'B3'])
    expect(got.get('B2')!.icon!.glyphClass).toBe('3Arrows:0')
    expect(stage4ObservedIcons(CF_STAGE3_CASES[0]!, dir)).toBeUndefined()
  })

  it('icon template diff reports every way Excel\'s dialog output differs from the generator (Lesson 12)', async () => {
    const generated = readZipPartsSync(await generateCaseArchive('stage4-icon-authored'))!
    expect(diffAuthoredIconTemplates(generated, new Map(generated))).toEqual([])
    const sheet = new TextDecoder().decode(generated.get('xl/worksheets/sheet1.xml')!)
    const mutate = (xml: string): Map<string, Uint8Array> => new Map([...generated, ['xl/worksheets/sheet1.xml', new TextEncoder().encode(xml)]])
    // An explicit percent="1" is a form difference (the glyph key pins the dialog's form), not a semantic one.
    expect(diffAuthoredIconTemplates(generated, mutate(sheet.replace('<x14:iconSet iconSet="3Stars">', '<x14:iconSet iconSet="3Stars" percent="1">'))).join('\n')).toMatch(/B2:B6: <iconSet> attributes/)
    // A gte flag, a cfIcon and a dropped rule are each reported, and every line names the template to change.
    const gte = diffAuthoredIconTemplates(generated, mutate(sheet.replace('<cfvo type="percent" val="33" gte="0"/>', '<cfvo type="percent" val="33"/>')))
    expect(gte.join('\n')).toMatch(/H2:H6: cfvo/)
    expect(gte.every(d => d.includes('iconRuleXml'))).toBe(true)
    expect(diffAuthoredIconTemplates(generated, mutate(sheet.replace('<x14:cfIcon iconSet="NoIcons" iconId="0"/>', '<x14:cfIcon iconSet="3Arrows" iconId="1"/>'))).join('\n')).toMatch(/E2:E6: cfIcon/)
    expect(diffAuthoredIconTemplates(generated, mutate(sheet.replace(/<conditionalFormatting sqref="F2:F6">[\s\S]*?<\/conditionalFormatting>/, ''))).join('\n')).toMatch(/F2:F6: Excel's authored workbook has no icon set rule/)
  })

  it.skipIf(!existsSync(EXCEL_AUTHORED_ICONSETS_PATH))('icon templates write what Excel\'s dialog writes (excel-authored-iconsets.xlsx)', async () => {
    const authored = readZipPartsSync(new Uint8Array(readFileSync(EXCEL_AUTHORED_ICONSETS_PATH)))!
    expect(diffAuthoredIconTemplates(readZipPartsSync(await generateCaseArchive('stage4-icon-authored'))!, authored)).toEqual([])
  })

  it('a capture of non-rectangular, unclassifiable glyphs records per-cell findings instead of refusing the case', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-icon-geom-'))
    s4dirs.push(dir) // no manifest here: no glyph key, so no cell classifies (as geometry and inset cells never do)
    const c = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-geometry')!
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
    const geom = computeWorksheetRangeGeometry(parts, c.range)!
    const iconRefs = c.expected.cells.filter(x => x.matched && x.iconApplied).map(x => x.ref)
    const png = paintPng(geom.expectedWidthPt, geom.expectedHeightPt, ctx => {
      ctx.fillStyle = '#00a000'
      for (const ref of iconRefs) {
        const p = parseA1(ref)!
        const x0 = Math.round(geom.colWidthsPt.slice(0, p.col - geom.minCol).reduce((a, b) => a + b, 0))
        const y0 = Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow).reduce((a, b) => a + b, 0))
        const h = Math.min(8, Math.round(geom.rowHeightsPt[p.row - geom.minRow]!) - 3)
        for (let r = 0; r < h; r++) ctx.fillRect(x0 + 2 + h - r, y0 + 2 + r, 2 * r + 1, 1) // a triangle: never rectangular
      }
    })
    const res = buildIconSidecarExtras(c, parts, png, geom, { lockBefore: 'unlocked', lockAfter: 'unlocked', rootDir: dir })
    expect(res.refusal).toBeUndefined()
    expect(res.observed!.size).toBe(iconRefs.length)
    const unstable = [...res.observed!.values()].filter(o => o.edgeStable === false)
    expect(unstable.length).toBeGreaterThan(0)
    for (const o of res.observed!.values()) if (o.icon) expect(o.icon.glyphClass).toBe('unclassified')
  })

  /** Fix round 1: hypotheses predicted from the real saved workbook, with a class per (set, iconId). */
  const s4classOf = (set: string, iconId: number): string | null => (set === 'NoIcons' ? null : `${set}:${iconId}`)
  const scenario = (caseId: string, cap: string) => {
    const c = CF_STAGE4_CASES.find(x => x.id === caseId)!
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
    return buildIconScenarioHypotheses(cap, c, parts, s4classOf)
  }
  /** Whether two hypotheses predict a different class at some designated cell. */
  const separates = (h: Record<string, IconClassHypothesis[]> | undefined, a: string, b: string): boolean =>
    h !== undefined && Object.values(h).some(list => list.find(x => x.id === a)?.glyphClass !== list.find(x => x.id === b)?.glyphClass)

  it('every bucket capability of the authored data has competing readings that predict different classes (S4-DIALECT-EVIDENCE, concern 2)', () => {
    const rows: Array<[string, string, string, string]> = [
      ['stage4-icon-authored', 'icon:dialect-base', 'as-modelled', 'no-icon'],
      ['stage4-icon-authored', 'icon:dialect-base', 'as-modelled', 'reversed-order'],
      ['stage4-icon-authored', 'icon:dialect-standalone', 'as-modelled', 'reversed-order'],
      ['stage4-icon-families', 'icon:settings-schema-defaults', 'as-modelled', 'family:3Arrows'],
      ['stage4-icon-families', 'icon:settings-schema-defaults', 'as-modelled', 'no-icon'],
      // (stage4-icon-dialect is unsupported since fix round 7, S4-DIALECT-UNSUPPORTED: it designates no cell.)
      ['stage4-icon-scope', 'icon:scope-disjoint', 'union', 'per-range'],
      ['stage4-icon-scope', 'icon:scope-hidden-row', 'exclude-hidden', 'include-hidden'],
      ['stage4-icon-scope', 'icon:scope-autofilter-hidden-row', 'exclude-hidden', 'include-hidden'],
      ['stage4-icon-scope', 'icon:scope-merged-cell', 'all-stored', 'merged-excluded'],
      ['stage4-icon-families', 'icon:scope-whole-column', 'whole-column', 'data-block-only'],
      ['stage4-icon-scope', 'icon:scope-duplicate-blocks', 'first-paints', 'second-paints'],
      ['stage4-icon-participation', 'icon:participation-types', 'no-icon-for-non-numbers', 'coerce-booleans'],
      ['stage4-icon-participation', 'icon:participation-dateserial', 'evaluated', 'no-icon'],
      ['stage4-icon-participation', 'icon:participation-error-presence', 'error-ignored', 'rule-suppressed'],
      ['stage4-icon-degenerate', 'icon:degenerate-equal-values', 'as-model', 'lowest'],
      ['stage4-icon-degenerate', 'icon:degenerate-single-cell', 'as-model', 'lowest'],
      ['stage4-icon-degenerate', 'icon:degenerate-non-monotonic', 'as-model', 'sorted'],
      ['stage4-icon-degenerate', 'icon:degenerate-non-monotonic', 'as-model', 'first-fail'],
      ['stage4-icon-degenerate', 'icon:bounds-out-of-domain', 'as-model', 'clamped'],
      ['stage4-icon-degenerate', 'icon:thresholds-floor', 'below-floor-bucket0', 'below-floor-no-icon'],
      ['stage4-icon-degenerate', 'icon:thresholds-min-max', 'max-reached', 'max-unreachable'],
      ['stage4-icon-composition', 'icon:composition-overlapping-icons', 'first-paints', 'second-paints'],
      ['stage4-icon-composition', 'icon:composition-with-bar', 'both-paint', 'icon-withheld'],
      ['stage4-icon-composition', 'icon:composition-matching-stopIfTrue-ahead', 'halted', 'not-halted'],
      ['stage4-icon-composition', 'icon:composition-noicons-slot', 'first-paints', 'second-paints'],
      ['stage4-icon-composition', 'icon:composition-unresolved-stub', 'halted', 'not-halted'],
    ]
    for (const [caseId, cap, a, b] of rows) expect(separates(scenario(caseId, cap), a, b), `${cap}: ${a} vs ${b}`).toBe(true)
    // Composition over a fill is the one shape no data can discriminate under the S4-A rule (registry, one-line reason).
    for (const id of ['icon:composition-with-scale', 'icon:composition-dxf-fill']) expect(CF_STAGE4_NONDISCRIMINATING[id], id).toMatch(/fill/)
    // A reading with no competitor on its cells stays non-discriminating: a dialect on all-equal buckets cannot tell order.
    const flat = buildIconScenarioHypotheses('icon:dialect-base', { ...CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-authored')!, expected: { ...CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-authored')!.expected, cells: [{ ref: 'F4', controlRef: 'F10', matched: true, iconApplied: true }] } }, readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/stage4-icon-authored.xlsx'))))!, s4classOf)!
    expect(discriminateIcon(flat, { F4: '3Arrows:1' })).toMatchObject({ status: 'non-discriminating' })
  })

  it('a dialect capture where every marked cell shows no icon refuses the case (S4-DIALECT-EVIDENCE)', () => {
    const c = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-authored')!
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
    const hyps = buildIconScenarioHypotheses('icon:dialect-base', c, parts, s4classOf)!
    const none = Object.fromEntries(Object.keys(hyps).map(ref => [ref, null]))
    expect(dialectAllAbsent('icon:dialect-base', none)).toBe(true)
    expect(dialectAllAbsent('icon:dialect-base', { ...none, F2: '3Arrows:2' })).toBe(false)
    // The dialect case records the outcome instead: no icon there is an answer, not a refusal.
    expect(dialectAllAbsent('icon:dialect-base-x14-family', none)).toBe(false)
  })

  it('stage4-icon-dialect authors an x14-only family on a base rule and a paired base + x14 rule', () => {
    const c = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-dialect')!
    // Ruling S4-DIALECT-UNSUPPORTED (fix round 7): the case records Excel's rejected resave and lifts nothing.
    expect(c.status).toBe('unsupported')
    expect(c.markers.map(m => [m.capabilityId, m.extension, m.icon?.set, m.priority ?? null])).toEqual([
      ['icon:dialect-base-x14-family', 'legacy', '3Stars', 1],
      ['icon:dialect-paired', 'legacy', '3Arrows', 2],
      ['icon:dialect-paired', 'x14', '3Triangles', null],
    ])
    const sheet = new TextDecoder().decode(readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.input))))!.get('xl/worksheets/sheet1.xml')!)
    const guid = /<x14:id>(\{[^}]+\})<\/x14:id>/.exec(sheet)?.[1]
    expect(guid).toBeDefined()
    expect(sheet).toContain(`<x14:cfRule type="iconSet" id="${guid}">`)
    const entry = (JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest).cases.find(x => x.id === c.id)!
    expect(verifyCfCase(entry, { rootDir: REPO_ROOT })).toMatchObject({ kind: 'refusal', structuralValid: true })
  })

  it('a measured Stage 4 case verifies and audits from its stored capture alone, and re-derives with no Excel run (concern 3)', () => {
    const root = mkdtempSync(join(tmpdir(), 'ov-cf-icon-e2e-'))
    s4dirs.push(root)
    const c0 = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-glyph-key')!
    mkdirSync(join(root, 'tests/fixtures/xlsx-cf/inputs'), { recursive: true })
    mkdirSync(join(root, 'tests/fixtures/xlsx-cf/refs'), { recursive: true })
    for (const f of [c0.input, c0.saved]) cpSync(resolve(REPO_ROOT, f), join(root, f))
    const savedBytes = new Uint8Array(readFileSync(join(root, c0.saved)))
    const parts = readZipPartsSync(savedBytes)!
    const geom = computeWorksheetRangeGeometry(parts, c0.range)!
    // A synthetic "Excel" capture: gridlines, one distinct glyph per key pair inside the left band, the solid DXF cell.
    const pairs = S4_ICON_FAMILIES.flatMap(f => Array.from({ length: f.size }, (_, id) => `${f.name}:${id}`))
    const boxOf = (ref: string) => { const p = parseA1(ref)!; return { x0: Math.round(geom.colWidthsPt.slice(0, p.col - geom.minCol).reduce((a, b) => a + b, 0)), y0: Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow).reduce((a, b) => a + b, 0)), w: Math.round(geom.colWidthsPt[p.col - geom.minCol]!), h: Math.round(geom.rowHeightsPt[p.row - geom.minRow]!) } }
    const W = geom.expectedWidthPt, H = geom.expectedHeightPt
    const canvas = createCanvas(W + 1, H + 1)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W + 1, H + 1)
    ctx.fillStyle = '#d4d4d4'
    let x = 0; for (const w of geom.colWidthsPt) { ctx.fillRect(Math.round(x), 0, 1, H + 1); x += w } ctx.fillRect(Math.round(x), 0, 1, H + 1)
    let y = 0; for (const h of geom.rowHeightsPt) { ctx.fillRect(0, Math.round(y), W + 1, 1); y += h } ctx.fillRect(0, Math.round(y), W + 1, 1)
    const k2 = boxOf('K2'); ctx.fillStyle = '#c6efce'; ctx.fillRect(k2.x0 + 1, k2.y0 + 1, k2.w - 1, k2.h - 1)
    ctx.fillStyle = '#202080'
    for (const m of c0.markers.filter(mm => mm.capabilityId.startsWith('icon:glyph-'))) {
      const i = pairs.indexOf(`${m.icon!.custom![0]!.set}:${m.icon!.custom![0]!.iconId}`)
      const b = boxOf(m.sqref)
      ctx.fillRect(b.x0 + 3 + (i % 10) * 2, b.y0 + 3 + Math.floor(i / 10), 6, 6)
    }
    const pngBytes = new Uint8Array(canvas.toBuffer('image/png'))
    const cap = c0.captures[0]!
    writeFileSync(join(root, cap.pngPath), pngBytes)
    const decoded = decodePngSync(pngBytes)!
    const savedSha256 = sha256(savedBytes)
    const inputSha256 = sha256(new Uint8Array(readFileSync(join(root, c0.input))))
    const sidecarFor = () => {
      const extra = buildIconSidecarExtras(c0, parts, decoded, geom, { lockBefore: 'unlocked', lockAfter: 'unlocked', rootDir: root })
      expect(extra.refusal).toBeUndefined()
      return JSON.stringify({
        excelVersion: '16.114', route: 'copy-picture', appearance: 'screen', scale: 1, pxPerPt: 1, width: decoded.width, height: decoded.height,
        colWidthsPt: geom.colWidthsPt, rowHeightsPt: geom.rowHeightsPt, geometrySource: 'excel-measured', resavedByExcel: true, resaveSavedSha256: savedSha256,
        sheet: c0.sheet, range: c0.range, hostTimeZone: 'Asia/Taipei', savedSha256, pngSha256: sha256(pngBytes), capturedAt: '2026-11-14T12:00:00Z', ...extra.extras,
      }, null, 2) + '\n'
    }
    writeFileSync(join(root, cap.sidecarPath), sidecarFor())
    const measured: CfOracleCase = {
      ...c0,
      status: 'measured',
      resavedByExcel: true,
      resaveReceipt: { resavedAt: '2026-11-14T12:00:00Z', excelVersion: '16.114', inputSha256, savedSha256, openedRepairFree: true },
      captures: [{ ...cap, pngSha256: sha256(pngBytes) }],
      expected: { ...c0.expected, inputSha256, savedSha256, provenance: 'excel-observed' },
    }
    const manifest: CfManifest = { version: 1, excelBuild: '16.114', locale: 'zh-TW', dateSystem: '1900', hostTimeZone: 'Asia/Taipei', generatedAt: '2026-11-06T00:00:00Z', cases: [measured] }
    writeFileSync(join(root, 'tests/fixtures/xlsx-cf/manifest.json'), JSON.stringify(manifest))
    const res = verifyCfCase(measured, { rootDir: root })
    expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
    const verified = (res as { verifiedCapabilities: readonly string[] }).verifiedCapabilities
    expect(verified).toEqual(expect.arrayContaining(['icon:pixel-calibration', 'icon:glyph-key', 'icon:glyph-3Arrows-0', 'icon:glyph-5Boxes-4']))
    const audited = auditStage4Manifest(manifest, { rootDir: root }).perCapability
    for (const id of ['icon:pixel-calibration', 'icon:glyph-key']) expect(audited.find(p => p.capabilityId === id)!.disposition, id).toBe('verified')
    // Re-deriving the sidecar from the stored PNG (what capture --verify-only does with newer hypotheses) needs no Excel
    // and verifies the same evidence; a tampered sidecar refuses.
    writeFileSync(join(root, cap.sidecarPath), sidecarFor())
    expect(verifyCfCase(measured, { rootDir: root })).toMatchObject({ kind: 'verified' })
    const tampered = JSON.parse(readFileSync(join(root, cap.sidecarPath), 'utf8'))
    tampered.observedIcons[0].icon.leftPx += 1
    writeFileSync(join(root, cap.sidecarPath), JSON.stringify(tampered))
    expect(verifyCfCase(measured, { rootDir: root })).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
  }, 60_000)

  it('fix round 3 (item 8): shapes Excel canonicalises on resave leave the capturable cases, which then author only rules Excel kept', () => {
    const sheetOf = (path: string) => new TextDecoder().decode(readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, path))))!.get('xl/worksheets/sheet1.xml')!)
    // Every icon rule (base or x14), with priority and GUIDs dropped: the semantic shape Excel either kept or rewrote.
    const rulesOf = (xml: string) => [
      ...[...xml.matchAll(/<conditionalFormatting sqref="([^"]+)">([\s\S]*?)<\/conditionalFormatting>/g)].map(m => `${m[1]} ${m[2]!.replace(/<extLst>[\s\S]*?<\/extLst>/g, '')}`),
      ...[...xml.matchAll(/<x14:conditionalFormatting [^>]*>([\s\S]*?)<xm:sqref>([^<]+)<\/xm:sqref>/g)].map(m => `${m[2]} ${m[1]}`),
    ].map(r => r.replace(/ priority="\d+"/g, '').replace(/ id="\{[^}]+\}"/g, ''))
    for (const id of ['stage4-icon-degenerate', 'stage4-icon-reverse-custom']) {
      const c = CF_STAGE4_CASES.find(x => x.id === id)!
      const authored = rulesOf(sheetOf(c.input))
      const excelKept = new Set(rulesOf(sheetOf(`tests/fixtures/xlsx-cf/${id}.rejected.xlsx`)))
      expect(authored.length, id).toBeGreaterThan(5)
      for (const r of authored) expect(excelKept.has(r), `${id}: Excel rewrote ${r}`).toBe(true)
      // Priorities stay contiguous (no gap for Excel to renumber).
      expect(c.markers.filter(m => m.priority !== undefined).map(m => m.priority).sort((a, b) => a! - b!), id).toEqual(c.markers.filter(m => m.priority !== undefined).map((_, i) => i + 1))
    }
    const unresolved = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-unresolved-shapes')!
    const xml = sheetOf(unresolved.input)
    expect(xml).toMatch(/<conditionalFormatting sqref="M2:M6"><cfRule type="iconSet" priority="12"><iconSet iconSet="3Arrows" percent="0">/)
    expect(xml).toMatch(/<x14:iconSet iconSet="3Arrows" custom="1" reverse="1">[\s\S]*?<xm:sqref>N2:N6<\/xm:sqref>/)
    for (const cap of ['icon:attr-percent-false', 'icon:reverse-custom', 'icon:dialect-base-x14-family', 'icon:dialect-paired']) {
      expect(CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === cap)!.failClosedDisposition, cap).toMatch(/c63de53a/)
    }
  })

  it('Ruling S4-ICON-STOPIFTRUE (fix round 6): an icon rule with stopIfTrue is unauthorable and is recorded with the unresolved shapes', () => {
    const sheetOf = (path: string) => new TextDecoder().decode(readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, path))))!.get('xl/worksheets/sheet1.xml')!)
    const comp = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-composition')!
    expect(comp.markers.some(m => m.ruleType === 'iconSet' && m.stopIfTrue)).toBe(false)
    expect(sheetOf(comp.input)).not.toMatch(/<cfRule type="iconSet"[^>]*stopIfTrue/)
    // Priorities stay contiguous after the move.
    expect(comp.markers.filter(m => m.priority !== undefined).map(m => m.priority!).sort((a, b) => a - b)).toEqual(comp.markers.filter(m => m.priority !== undefined).map((_, i) => i + 1))
    const un = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-unresolved-shapes')!
    expect(un.markers.filter(m => m.capabilityId === 'icon:composition-icon-stopIfTrue').map(m => [m.ruleType, m.stopIfTrue === true])).toEqual([['iconSet', true], ['cellIs', false]])
    expect(sheetOf(un.input)).toMatch(/<conditionalFormatting sqref="O2:O6"><cfRule type="iconSet" priority="14" stopIfTrue="1">/)
    expect(CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === 'icon:composition-icon-stopIfTrue')!.failClosedDisposition).toMatch(/S4-ICON-STOPIFTRUE/)
  })

  it('fix round 7: S4-DIALECT-UNSUPPORTED and S4-MINMAX make the dialect and min/max capabilities refused-unsupported', async () => {
    const dialect = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-dialect')!
    expect(dialect.status).toBe('unsupported')
    // S4-MINMAX uses the S3-CLOSE registry: the degenerate capture keeps its other evidence (no recapture).
    expect(Object.keys(CF_STAGE4_UNSUPPORTED)).toEqual(['icon:thresholds-min-max'])
    expect(CF_STAGE4_UNSUPPORTED['icon:thresholds-min-max']).toMatch(/S4-MINMAX[\s\S]*3Arrows:2/)
    const owner = (cap: string) => CF_STAGE4_CASES.find(c => c.id === CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === cap)!.caseId)!
    for (const cap of ['icon:dialect-base-x14-family', 'icon:dialect-paired']) expect(owner(cap).status, cap).toBe('unsupported')
    for (const cap of ['icon:dialect-base-x14-family', 'icon:dialect-paired', 'icon:thresholds-min-max']) {
      expect(CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === cap)!.failClosedDisposition, cap).toMatch(/S4-DIALECT-UNSUPPORTED|S4-MINMAX/)
    }
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    const dispositions = auditStage4Manifest(manifest, { rootDir: REPO_ROOT }).perCapability
    for (const cap of ['icon:dialect-base-x14-family', 'icon:dialect-paired', 'icon:thresholds-min-max']) {
      expect(dispositions.find(p => p.capabilityId === cap)!.disposition, cap).toBe('refused-unsupported')
      expect(isCapabilityVerified(cap, manifest, { rootDir: REPO_ROOT }), cap).toBe(false)
    }
  }, 120_000)

  describe('fix round 8: stage4-icon-geometry-2 discriminates the paint gates and the bar/icon showValue conflict', () => {
    const G2 = () => CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-geometry-2')!
    it('the new case owns geometry-box, vertical-align and composition-bar-showvalue, with a rectangular x14 glyph, ahead of stage4-icon-geometry', () => {
      const c = G2()
      expect(c).toBeDefined()
      const ids = CF_STAGE4_CASES.map(x => x.id)
      expect(ids.indexOf('stage4-icon-geometry-2')).toBe(ids.indexOf('stage4-icon-geometry') - 1)
      for (const cap of ['icon:geometry-box', 'icon:geometry-vertical-align', 'icon:composition-bar-showvalue']) {
        expect(CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === cap)!.caseId, cap).toBe('stage4-icon-geometry-2')
      }
      // The box and placement cells draw 5Boxes:2 (two filled boxes: a rectangle whose three scanlines agree).
      for (const m of c.markers.filter(x => x.capabilityId.startsWith('icon:geometry-'))) expect([m.icon?.set, m.icon?.dialect], m.sqref).toEqual(['5Boxes', 'x14'])
    })

    /** A synthetic Excel capture of the new case: rectangular glyphs per `box`, text patterns and short right-to-left bars. */
    const capture = (o: { side: (fontPt: number, rowPt: number) => number; place: (cellV: 'top' | 'center' | 'bottom', rowPt: number, side: number) => number; hidden: (col: string) => boolean }) => {
      const c = G2()
      const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.input))))!
      const geom = computeWorksheetRangeGeometry(parts, c.range)!
      const boxIn = (ref: string) => { const p = parseA1(ref)!; return { x0: Math.round(geom.colWidthsPt.slice(0, p.col - geom.minCol).reduce((a, b) => a + b, 0)), y0: Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow).reduce((a, b) => a + b, 0)), w: Math.round(geom.colWidthsPt[p.col - geom.minCol]!), h: Math.round(geom.rowHeightsPt[p.row - geom.minRow]!) } }
      const W = geom.expectedWidthPt, H = geom.expectedHeightPt
      const canvas = createCanvas(W + 1, H + 1)
      const ctx = canvas.getContext('2d')
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W + 1, H + 1)
      const FONT: Record<string, number> = { B: 8, C: 11, D: 20 }
      const VERT: Record<string, 'top' | 'center' | 'bottom'> = { E: 'top', F: 'center', G: 'bottom' }
      for (const e of c.expected.cells.filter(x => x.matched)) {
        const col = /^[A-Z]+/.exec(e.ref)![0]
        const b = boxIn(e.ref)
        if ('BCDEFG'.includes(col)) {
          const side = o.side(FONT[col] ?? 11, b.h)
          ctx.fillStyle = '#202080'
          ctx.fillRect(b.x0 + 3, b.y0 + o.place(VERT[col] ?? 'bottom', b.h, side), side, side)
        } else {
          // H..J: icon at the left, text after its inset unless hidden, the control's text un-inset, a short bar at the right.
          ctx.fillStyle = '#202080'; ctx.fillRect(b.x0 + 3, b.y0 + b.h - 13, 11, 11)
          if (!o.hidden(col)) textAt(ctx, b.x0 + 4 + 12, b.y0 + b.h - 10)
          const cb = boxIn(e.controlRef)
          textAt(ctx, cb.x0 + 4, cb.y0 + cb.h - 10)
          ctx.fillStyle = '#0000ff'; ctx.fillRect(b.x0 + b.w - 6, b.y0 + 2, 4, b.h - 3)
        }
      }
      const decoded = decodePngSync(new Uint8Array(canvas.toBuffer('image/png')))!
      const observed = measureCaseIcons(c, decoded, geom) as ReadonlyMap<string, CfObservedIcon>
      return evaluateIconCapabilities(c, parts, observed, undefined)
    }
    const bottom = (_: string, rowPt: number, side: number) => rowPt - 2 - side

    it('geometry-box discriminates fixed, font and row boxes on the new case', () => {
      expect(capture({ side: () => 11, place: bottom, hidden: () => false })['icon:geometry-box']).toMatchObject({ status: 'discriminating', matched: 'fixed' })
      expect(capture({ side: f => f, place: bottom, hidden: () => false })['icon:geometry-box']).toMatchObject({ status: 'discriminating', matched: 'font' })
      expect(capture({ side: (_, r) => r - 6, place: bottom, hidden: () => false })['icon:geometry-box']).toMatchObject({ status: 'discriminating', matched: 'row' })
    })

    it('geometry-vertical-align discriminates top, centre, bottom and the cell alignment on the new case', () => {
      const at = (v: 'top' | 'center' | 'bottom', rowPt: number, side: number) => (v === 'top' ? 1 : v === 'bottom' ? rowPt - 2 - side : Math.round((rowPt - side) / 2))
      for (const [truth, place] of [
        ['cell-align', (v: 'top' | 'center' | 'bottom', r: number, s: number) => at(v, r, s)],
        ['top', (_: string, r: number, s: number) => at('top', r, s)],
        ['center', (_: string, r: number, s: number) => at('center', r, s)],
        ['bottom', (_: string, r: number, s: number) => at('bottom', r, s)],
      ] as const) {
        expect(capture({ side: () => 11, place: place as never, hidden: () => false })['icon:geometry-vertical-align'], truth).toMatchObject({ status: 'discriminating', matched: truth })
      }
    })

    it('composition-bar-showvalue discriminates any-hides, priority-wins, bar-wins and icon-wins', () => {
      // H: bar(showValue 0) ahead of icon(1); I: bar(1) ahead of icon(0); J: icon(1) ahead of bar(0).
      const truths: Record<string, Record<string, boolean>> = {
        'any-hides': { H: true, I: true, J: true },
        'priority-wins': { H: true, I: false, J: false },
        'bar-wins': { H: true, I: false, J: true },
        'icon-wins': { H: false, I: true, J: false },
      }
      for (const [truth, hide] of Object.entries(truths)) {
        expect(capture({ side: () => 11, place: bottom, hidden: col => hide[col]! })['icon:composition-bar-showvalue'], truth).toMatchObject({ status: 'discriminating', matched: truth })
      }
    })
  })

  describe('fix round 7: icon geometry, inset and showValue hypotheses (S4-K/S4-L, S3-A style)', () => {
    type Facts = IconGeometryCell['facts']
    /** A synthetic measured cell at 1 px/pt: cell box from its column/row, icon box relative to the cell's outer box. */
    const cellAt = (ref: string, facts: Partial<Facts> & { rowPt: number }, icon: { l: number; t: number; w: number; h: number } | null, o: { widthPx?: number; text?: number; stable?: boolean } = {}): IconGeometryCell => {
      const p = parseA1(ref)!
      const x0 = 100 + p.col * 80, y0 = 20 + p.row * 70, W = o.widthPx ?? 72
      const observed: CfObservedIcon = {
        ref, cellLeftPx: x0, cellTopPx: y0, cellWidthPx: W, cellHeightPx: facts.rowPt,
        icon: icon ? { leftPx: x0 + icon.l, topPx: y0 + icon.t, rightPx: x0 + icon.l + icon.w - 1, bottomPx: y0 + icon.t + icon.h - 1, glyphClass: 'icon-class-none' } : null,
        ...(o.text !== undefined ? { textInkLeftPx: x0 + 40, textInkRightPx: x0 + 50, textShiftPx: o.text } : {}),
        ...(o.stable === false ? { edgeStable: false as const } : {}),
      }
      return { ref, facts: { fontPt: 11, vertical: 'bottom', horizontal: 'left', showValue: true, ...facts }, observed }
    }
    const FONTS: Record<string, number> = { B: 8, C: 11, D: 20 }
    const ROWS: Record<string, number> = { '2': 12, '3': 30, '4': 60 }
    const boxGrid = (w: (fontPt: number, rowPt: number) => number, stable: (ref: string) => boolean = () => true) =>
      ['B', 'C', 'D'].flatMap(col => ['2', '3', '4'].map(r => {
        const ref = `${col}${r}`, fontPt = FONTS[col]!, rowPt = ROWS[r]!
        const side = w(fontPt, rowPt)
        return cellAt(ref, { fontPt, rowPt }, { l: 2, t: rowPt - 1 - side, w: side, h: side }, { stable: stable(ref) })
      }))

    it('geometry-box discriminates the font-scaled box from the fixed and row boxes, and only on stable cells', () => {
      // Font-scaled: side = font pt (a 20 pt box cannot fit the 12 pt row, so that cell is no evidence for "font").
      const font = boxGrid(f => f)
      expect(evaluateIconGeometryCapability('icon:geometry-box', font, font)).toMatchObject({ status: 'discriminating', matched: 'font' })
      // Fixed 11 px everywhere.
      const fixed = boxGrid(() => 11)
      expect(evaluateIconGeometryCapability('icon:geometry-box', fixed, fixed)).toMatchObject({ status: 'discriminating', matched: 'fixed' })
      // Row-scaled: side = row less 4 px.
      const row = boxGrid((_, r) => r - 4)
      expect(evaluateIconGeometryCapability('icon:geometry-box', row, row)).toMatchObject({ status: 'discriminating', matched: 'row' })
      // Only the 11 pt column is edge-stable: fixed and font cannot be told apart (honestly non-discriminating).
      const onlyC = boxGrid(f => f, ref => ref.startsWith('C'))
      expect(evaluateIconGeometryCapability('icon:geometry-box', onlyC, onlyC).status).toBe('non-discriminating')
      // A designated cell with no icon refutes every box reading.
      const gone = boxGrid(f => f).map(c => (c.ref === 'C3' ? { ...c, observed: { ...c.observed, icon: null } } : c))
      expect(evaluateIconGeometryCapability('icon:geometry-box', gone, gone).status).toBe('none')
    })

    it('geometry-vertical-align separates top, centre, bottom and the cell alignment', () => {
      const side = 11
      const place = (v: 'top' | 'center' | 'bottom', rowPt: number) => (v === 'top' ? 1 : v === 'bottom' ? rowPt - 1 - side : Math.round((rowPt - side) / 2))
      const grid = (how: (cellV: 'top' | 'center') => 'top' | 'center' | 'bottom') => (['E', 'H'] as const).flatMap(col => ['3', '4'].map(r => {
        const cellV = col === 'E' ? 'top' as const : 'center' as const
        const rowPt = ROWS[r]!
        return cellAt(`${col}${r}`, { rowPt, vertical: cellV }, { l: 2, t: place(how(cellV), rowPt), w: side, h: side })
      }))
      const follows = grid(v => v)
      expect(evaluateIconGeometryCapability('icon:geometry-vertical-align', follows, follows)).toMatchObject({ status: 'discriminating', matched: 'cell-align' })
      const bottom = grid(() => 'bottom')
      expect(evaluateIconGeometryCapability('icon:geometry-vertical-align', bottom, bottom)).toMatchObject({ status: 'discriminating', matched: 'bottom' })
      // Only centred cells: centre and cell-align agree.
      const centred = follows.filter(c => c.ref.startsWith('H'))
      expect(evaluateIconGeometryCapability('icon:geometry-vertical-align', centred, centred).status).toBe('non-discriminating')
    })

    it('narrow and short-row hosts: clipped, shrunk or no icon, against a same-font reference box', () => {
      const ref11 = [cellAt('M3', { rowPt: 30 }, { l: 2, t: 18, w: 11, h: 11 }, { widthPx: 188 })]
      const narrowClip = [cellAt('K3', { rowPt: 30 }, { l: 2, t: 18, w: 9, h: 11 }, { widthPx: 12 })]
      expect(evaluateIconGeometryCapability('icon:geometry-narrow-clip', narrowClip, [...narrowClip, ...ref11])).toMatchObject({ status: 'discriminating', matched: 'clip' })
      const narrowShrink = [cellAt('K3', { rowPt: 30 }, { l: 2, t: 22, w: 7, h: 7 }, { widthPx: 12 })]
      expect(evaluateIconGeometryCapability('icon:geometry-narrow-clip', narrowShrink, [...narrowShrink, ...ref11])).toMatchObject({ status: 'discriminating', matched: 'shrink' })
      const narrowNone = [cellAt('K3', { rowPt: 30 }, null, { widthPx: 12 })]
      expect(evaluateIconGeometryCapability('icon:geometry-narrow-clip', narrowNone, [...narrowNone, ...ref11])).toMatchObject({ status: 'discriminating', matched: 'no-icon' })
      // Short row (6 pt): a clip keeps the reference width, a shrink does not.
      const shortClip = [cellAt('L5', { rowPt: 6 }, { l: 2, t: 1, w: 11, h: 4 })]
      expect(evaluateIconGeometryCapability('icon:geometry-short-row-clip', shortClip, [...shortClip, ...ref11])).toMatchObject({ status: 'discriminating', matched: 'clip' })
      const shortShrink = [cellAt('L5', { rowPt: 6 }, { l: 2, t: 1, w: 4, h: 4 })]
      expect(evaluateIconGeometryCapability('icon:geometry-short-row-clip', shortShrink, [...shortShrink, ...ref11])).toMatchObject({ status: 'discriminating', matched: 'shrink' })
      // No reference box in the case: clip and shrink cannot be told apart.
      expect(evaluateIconGeometryCapability('icon:geometry-short-row-clip', shortShrink, shortShrink).status).toBe('non-discriminating')
    })

    it('snapping is non-discriminating when no measured gridline lands at a fractional pt', () => {
      // Ruling S4-SNAP-SHARED (fix round 9): registry non-discriminating by construction; the paint gate will use
      // bar:geometry-snapping (S4-PROD-ALIGN).
      expect(CF_STAGE4_NONDISCRIMINATING['icon:geometry-snapping']).toMatch(/S4-SNAP-SHARED[\s\S]*bar:geometry-snapping/)
      const cells = [cellAt('M2', { rowPt: 12 }, { l: 2, t: 1, w: 11, h: 8 }, { widthPx: 188 }), cellAt('N2', { rowPt: 12 }, { l: 2, t: 1, w: 11, h: 8 }, { widthPx: 195 })]
      expect(evaluateIconGeometryCapability('icon:geometry-snapping', cells, cells)).toMatchObject({ status: 'non-discriminating', reason: expect.stringMatching(/fractional/) })
    })

    it('text insets: left, right-aligned, centred and overflow readings, and the NoIcons inset and showValue', () => {
      const icon = { l: 2, t: 7, w: 11, h: 11 }
      const at = (ref: string, horizontal: Facts['horizontal'], text: number | undefined, ic: typeof icon | null = icon, more: Partial<Facts> = {}) => cellAt(ref, { rowPt: 20, horizontal, ...more }, ic, text === undefined ? {} : { text })
      const left = [at('B2', 'left', 12), at('B3', 'left', 12)]
      expect(evaluateIconGeometryCapability('icon:inset-left', left, left)).toMatchObject({ status: 'discriminating', matched: 'inset' })
      const leftNone = [at('B2', 'left', 0), at('B3', 'left', 1)]
      expect(evaluateIconGeometryCapability('icon:inset-left', leftNone, leftNone)).toMatchObject({ status: 'discriminating', matched: 'none' })
      expect(evaluateIconGeometryCapability('icon:inset-right-aligned', [at('C2', 'right', 0)], [at('C2', 'right', 0)])).toMatchObject({ status: 'discriminating', matched: 'unshifted' })
      expect(evaluateIconGeometryCapability('icon:inset-right-aligned', [at('C2', 'right', -12)], [at('C2', 'right', -12)])).toMatchObject({ status: 'discriminating', matched: 'shifted-left' })
      expect(evaluateIconGeometryCapability('icon:inset-center', [at('D2', 'center', 0)], [at('D2', 'center', 0)])).toMatchObject({ status: 'discriminating', matched: 'full-cell' })
      expect(evaluateIconGeometryCapability('icon:inset-center', [at('D2', 'center', 6)], [at('D2', 'center', 6)])).toMatchObject({ status: 'discriminating', matched: 'inset-box' })
      // Overflow: an icon plus overflowing ink wider than the reference icon; or the icon withheld (production's fallback).
      const pool = [at('B2', 'left', 12)]
      const wide = [at('E2', 'right', undefined, { l: 2, t: 4, w: 24, h: 12 })]
      expect(evaluateIconGeometryCapability('icon:inset-overflow', wide, [...wide, ...pool])).toMatchObject({ status: 'discriminating', matched: 'icon-and-text' })
      const withheld = [at('E2', 'right', 0, null)]
      expect(evaluateIconGeometryCapability('icon:inset-overflow', withheld, [...withheld, ...pool])).toMatchObject({ status: 'discriminating', matched: 'icon-withheld' })
      // NoIcons: production reads "no inset" and "text shown"; Excel may inset or hide.
      const noIconsInset = [at('F2', 'left', 12, null)]
      expect(evaluateIconGeometryCapability('icon:noicons-inset', noIconsInset, noIconsInset)).toMatchObject({ status: 'discriminating', matched: 'inset' })
      const noIconsFlat = [at('F2', 'left', 0, null)]
      expect(evaluateIconGeometryCapability('icon:noicons-inset', noIconsFlat, noIconsFlat)).toMatchObject({ status: 'discriminating', matched: 'no-inset' })
      // S4-REPLICATE (fix round 10): a showValue verdict needs two agreeing cells per configuration (column).
      const hidden = [at('G2', 'left', undefined, null, { showValue: false }), at('G3', 'left', undefined, null, { showValue: false })]
      expect(evaluateIconGeometryCapability('icon:noicons-showvalue', hidden, hidden)).toMatchObject({ status: 'discriminating', matched: 'text-hidden' })
      const shown = [at('G2', 'left', 0, null, { showValue: false }), at('G3', 'left', 1, null, { showValue: false })]
      expect(evaluateIconGeometryCapability('icon:noicons-showvalue', shown, shown)).toMatchObject({ status: 'discriminating', matched: 'text-shown' })
      expect(evaluateIconGeometryCapability('icon:noicons-showvalue', hidden.slice(0, 1), hidden.slice(0, 1))).toMatchObject({ status: 'non-discriminating', reason: expect.stringMatching(/S4-REPLICATE/) })
    })

    it('fix round 9: a sidecar derived under older geometry rules leaves its geometry capabilities unverified, never refuses the case', () => {
      const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
      // A measured case with geometry capabilities (fix round 10: geometry-2 is regenerated, so the inset case stands in).
      const c = manifest.cases.find(x => x.id === 'stage4-icon-inset')!
      expect(c.status).toBe('measured')
      const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { iconGeometryRulesVersion?: number }
      const res = verifyCfCase(c, { rootDir: REPO_ROOT })
      expect(res.kind, res.kind === 'refusal' ? `${res.code}: ${res.reason}` : '').toBe('verified')
      if ((sidecar.iconGeometryRulesVersion ?? 1) !== ICON_GEOMETRY_RULES_VERSION) {
        const v = res as { verifiedCapabilities: readonly string[] }
        for (const cap of ['icon:inset-left', 'icon:inset-right-aligned', 'icon:noicons-inset']) expect(v.verifiedCapabilities, cap).not.toContain(cap)
      }
    }, 60_000)

    it('fix round 9 (regression test on Excel\'s stored numbers, not independent evidence): a font-scaled box is judged within +-1 px in measured px', () => {
      // Excel 16.115 (stage4-icon-geometry-2): 5Boxes:2 ink widths 7 / 10 / 20 px at 8 / 11 / 20 pt, the same in 30 and
      // 60 pt rows. Some scale k puts every cell within 1 px of k x font (k in [0.95, 1]): font fits; fixed and row do not.
      const cells = ['B', 'C', 'D'].flatMap(col => ['2', '3'].map(r => {
        const fontPt = ({ B: 8, C: 11, D: 20 } as Record<string, number>)[col]!, rowPt = r === '2' ? 30 : 60
        const w = ({ B: 7, C: 10, D: 20 } as Record<string, number>)[col]!
        return cellAt(`${col}${r}`, { fontPt, rowPt }, { l: 3, t: rowPt - 2 - Math.round(w / 2), w, h: Math.round(w / 2) })
      }))
      expect(evaluateIconGeometryCapability('icon:geometry-box', cells, cells)).toMatchObject({ status: 'discriminating', matched: 'font' })
      // Widths 7 / 10 / 22: the best scale leaves a 1.35 px residual (undecided); 3 / 16 / 20 refutes every reading.
      const off = cells.map(c => (c.ref.startsWith('D') ? cellAt(c.ref, { ...c.facts }, { l: 3, t: 1, w: 22, h: 11 }) : c))
      expect(evaluateIconGeometryCapability('icon:geometry-box', off, off).status).toBe('non-discriminating')
      const far = cells.map(c => (c.ref.startsWith('B') ? cellAt(c.ref, { ...c.facts }, { l: 3, t: 1, w: 3, h: 2 }) : c.ref.startsWith('C') ? cellAt(c.ref, { ...c.facts }, { l: 3, t: 1, w: 16, h: 8 }) : c))
      expect(evaluateIconGeometryCapability('icon:geometry-box', far, far).status).toBe('none')
    })

    it('fix round 10 (mechanism test): the free-scale band: a ~1.3 px best residual is undecided, a ~2 px one refutes', () => {
      const grid = (w: Record<string, number>) => ['B', 'C', 'D'].flatMap(col => ['2', '3'].map(r => {
        const fontPt = ({ B: 8, C: 11, D: 20 } as Record<string, number>)[col]!, rowPt = r === '2' ? 30 : 60
        return cellAt(`${col}${r}`, { fontPt, rowPt }, { l: 3, t: 2, w: w[col]!, h: 4 })
      }))
      // 7 / 10 / 22 px: min over k of max |w - k x font| is about 1.35 px, inside the (1, 1.5) band: font undecided.
      const near = grid({ B: 7, C: 10, D: 22 })
      expect(evaluateIconGeometryCapability('icon:geometry-box', near, near)).toMatchObject({ status: 'non-discriminating' })
      // 7 / 14 / 20 px: the best residual is about 1.95 px, past 1.5: font refuted, and with fixed and row refuted, none.
      const far = grid({ B: 7, C: 14, D: 20 })
      expect(evaluateIconGeometryCapability('icon:geometry-box', far, far)).toMatchObject({ status: 'none' })
    })

    it('fix round 10 (mechanism test): ink inside the measured bar extent is no text evidence, and each configuration needs two agreeing cells', () => {
      const conflict = (barShow: boolean, iconShow: boolean, barFirst: boolean) => ({ barShow, iconShow, barFirst })
      const at = (ref: string, k: ReturnType<typeof conflict>, text: { shift: number; left: number } | undefined, bar?: { l: number; r: number }) => {
        const c = cellAt(ref, { rowPt: 30, horizontal: 'left', conflict: k }, { l: 2, t: 15, w: 11, h: 12 }, { widthPx: 180 })
        const x0 = c.observed.cellLeftPx
        const withText = text ? { ...c, observed: { ...c.observed, textShiftPx: text.shift, textInkLeftPx: x0 + text.left, textInkRightPx: x0 + text.left + 10 } } : c
        return bar ? { ...withText, barExtent: { leftPx: x0 + bar.l, rightPx: x0 + bar.r } } : withText
      }
      const H = conflict(false, true, true), I = conflict(true, false, true), J = conflict(false, true, false)
      const bar = { l: 166, r: 178 }
      // Rows 2 and 5 (no bar): H hidden (only the icon matches the text pattern), I and J shown after the inset; row 3 (a
      // short bar at the right): the rightmost match lies inside the measured bar, so those cells say nothing.
      const clean = [at('H2', H, { shift: -1, left: 3 }), at('I2', I, { shift: 12, left: 16 }), at('J2', J, { shift: 12, left: 16 }),
        at('H5', H, { shift: -1, left: 3 }), at('I5', I, { shift: 12, left: 16 }), at('J5', J, { shift: 12, left: 16 })]
      const barred = [at('H3', H, { shift: 165, left: 169 }, bar), at('I3', I, { shift: 165, left: 169 }, bar), at('J3', J, { shift: 165, left: 169 }, bar)]
      const got = evaluateIconGeometryCapability('icon:composition-bar-showvalue', [...clean, ...barred], [...clean, ...barred])
      expect(got).toMatchObject({ status: 'discriminating', matched: 'priority-wins' })
      expect(got.geometry!['H3']!['priority-wins']).toEqual({ verdict: 'excluded' })
      expect(got.textExclusions).toEqual(['H3', 'I3', 'J3'].map(ref => ({ ref, reason: expect.stringMatching(/bar/) })))
      // S4-REPLICATE: one clean cell per configuration (row 2 only) is not enough.
      const single = [...clean.slice(0, 3), ...barred]
      expect(evaluateIconGeometryCapability('icon:composition-bar-showvalue', single, single)).toMatchObject({ status: 'non-discriminating', reason: expect.stringMatching(/S4-REPLICATE/) })
      // Ink in the right half but outside any measured bar is text evidence (the old half-cell rule would have dropped it):
      // a shown value there refutes the readings that hide it.
      const farText = [at('H2', H, { shift: 120, left: 124 }), at('H5', H, { shift: 120, left: 124 }), ...clean.slice(1, 3), ...clean.slice(4)]
      expect(evaluateIconGeometryCapability('icon:composition-bar-showvalue', farText, farText).geometry!['H2']!['priority-wins']).toEqual({ verdict: 'fail' })
    })

    it('fix round 10: the bar extent of a right-to-left bar is measured from the capture, anchored at the inner right edge', () => {
      // A 180 x 30 cell: an icon at the left, text after it, and a 12 px solid bar flush with the inner right edge.
      const png = paintPng(400, 120, ctx => {
        ctx.fillStyle = '#202080'; ctx.fillRect(102, 37, 11, 12)
        ctx.fillStyle = '#000000'; ctx.fillRect(116, 40, 2, 8)
        ctx.fillStyle = '#0000ff'; ctx.fillRect(266, 22, 12, 26)
      })
      expect(measureBarExtentInCell(png, { x: 101, y: 21, width: 178, height: 28 }, 'right')).toEqual({ leftPx: 266, rightPx: 277 })
      // No bar: nothing tall enough touches the anchored edge.
      expect(measureBarExtentInCell(png, { x: 101, y: 21, width: 160, height: 28 }, 'right')).toBeUndefined()
    })

    it('fix round 10: S4-REPLICATE counts decisive cells per configuration', () => {
      expect(replicationShortfall(['H2', 'H5', 'I2', 'I5'])).toBeUndefined()
      expect(replicationShortfall(['H2', 'H5', 'I2'])).toMatch(/S4-REPLICATE: configuration I has 1/)
    })

    it('showvalue canvas suppression decides only on cells whose text and icon band cannot overlap', () => {
      const icon = { l: 2, t: 7, w: 11, h: 11 }
      // General (right-aligned) number, showValue 0: the only ink matching the text pattern is far left (the icon): hidden.
      const general = ['B2', 'B3'].map(ref => cellAt(ref, { rowPt: 20, horizontal: 'right', showValue: false }, null, { text: -164 }))
      expect(evaluateIconGeometryCapability('icon:showvalue-canvas-suppression', general, general)).toMatchObject({ status: 'discriminating', matched: 'text-hidden' })
      const visible = ['B2', 'B3'].map(ref => cellAt(ref, { rowPt: 20, horizontal: 'right', showValue: false }, icon, { text: 0 }))
      expect(evaluateIconGeometryCapability('icon:showvalue-canvas-suppression', visible, visible)).toMatchObject({ status: 'discriminating', matched: 'text-shown' })
      // S4-REPLICATE: one decisive cell in the configuration is not enough.
      expect(evaluateIconGeometryCapability('icon:showvalue-canvas-suppression', general.slice(0, 1), general.slice(0, 1)).status).toBe('non-discriminating')
      // Centred: a centred icon can sit where the text would, so the cell is no evidence either way.
      const centred = [cellAt('C2', { rowPt: 20, horizontal: 'center', showValue: false }, null, { text: -1 })]
      expect(evaluateIconGeometryCapability('icon:showvalue-canvas-suppression', centred, centred).status).toBe('non-discriminating')
    })
  })

  describe('Ruling S4-NEARTWIN (fix round 3): near-twin glyph classes', () => {
    const PAIRS = S4_ICON_FAMILIES.flatMap(f => Array.from({ length: f.size }, (_, id) => `${f.name}:${id}`))
    const keyCase = () => CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-glyph-key')!
    /** A pixel colour `delta` below white on every channel. */
    const grey = (delta: number) => { const v = (255 - delta).toString(16).padStart(2, '0'); return `#${v}${v}${v}` }
    /** `as`: draw another pair's glyph; `faint`: plus one grey pixel; `onlyGrey`: draw the glyph square in grey of that delta instead. */
    type Variant = { as?: string; faint?: { delta: number; dx?: number }; onlyGrey?: number }
    /** Cell boxes of a geometry (gridline at x0/y0). */
    const boxIn = (geom: WorksheetRangeGeometry, ref: string) => { const p = parseA1(ref)!; return { x0: Math.round(geom.colWidthsPt.slice(0, p.col - geom.minCol).reduce((a, b) => a + b, 0)), y0: Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow).reduce((a, b) => a + b, 0)), w: Math.round(geom.colWidthsPt[p.col - geom.minCol]!), h: Math.round(geom.rowHeightsPt[p.row - geom.minRow]!) } }
    /** White page with gridlines; `cells` draws each glyph cell (its box, and the pair index whose square it shows). */
    const drawSheet = (geom: WorksheetRangeGeometry, paint: (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>) => void): Uint8Array => {
      const W = geom.expectedWidthPt, H = geom.expectedHeightPt
      const canvas = createCanvas(W + 1, H + 1)
      const ctx = canvas.getContext('2d')
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W + 1, H + 1)
      ctx.fillStyle = '#d4d4d4'
      let x = 0; for (const w of geom.colWidthsPt) { ctx.fillRect(Math.round(x), 0, 1, H + 1); x += w } ctx.fillRect(Math.round(x), 0, 1, H + 1)
      let y = 0; for (const h of geom.rowHeightsPt) { ctx.fillRect(0, Math.round(y), W + 1, 1); y += h } ctx.fillRect(0, Math.round(y), W + 1, 1)
      paint(ctx)
      return new Uint8Array(canvas.toBuffer('image/png'))
    }
    const glyph = (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>, b: { x0: number; y0: number }, i: number) => {
      ctx.fillStyle = '#202080'
      ctx.fillRect(b.x0 + 3 + (i % 10) * 2, b.y0 + 3 + Math.floor(i / 10), 6, 6)
    }
    /** The synthetic key capture: one distinct glyph per pair; a `variants` pair copies another pair's glyph plus a faint pixel. */
    const keyCapture = (variants: Record<string, Variant> = {}) => {
      const c0 = keyCase()
      const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c0.saved))))!
      const geom = computeWorksheetRangeGeometry(parts, c0.range)!
      const png = drawSheet(geom, ctx => {
        const k2 = boxIn(geom, 'K2'); ctx.fillStyle = '#c6efce'; ctx.fillRect(k2.x0 + 1, k2.y0 + 1, k2.w - 1, k2.h - 1)
        for (const m of c0.markers.filter(mm => mm.capabilityId.startsWith('icon:glyph-'))) {
          const pair = `${m.icon!.custom![0]!.set}:${m.icon!.custom![0]!.iconId}`
          const v = variants[pair]
          const b = boxIn(geom, m.sqref)
          if (v?.onlyGrey !== undefined) {
            const i = PAIRS.indexOf(pair)
            ctx.fillStyle = grey(v.onlyGrey); ctx.fillRect(b.x0 + 3 + (i % 10) * 2, b.y0 + 3 + Math.floor(i / 10), 6, 6)
          } else glyph(ctx, b, PAIRS.indexOf(v?.as ?? pair))
          if (v?.faint) { ctx.fillStyle = grey(v.faint.delta); ctx.fillRect(b.x0 + 24 + (v.faint.dx ?? 0), b.y0 + 15, 1, 1) }
        }
      })
      return { c0, parts, geom, png, decoded: decodePngSync(png)! }
    }

    it('glyph key admits a near-twin pair as two classes with a twin relation', () => {
      // 5Quarters:4 draws 4RedToBlack:0's glyph plus one pixel of max delta 32: unseparated (8 < 32 < 64), not identical.
      const { c0, parts, geom, decoded } = keyCapture({ '5Quarters:4': { as: '4RedToBlack:0', faint: { delta: 32 } } })
      const key = buildIconGlyphKey(c0, parts, decoded, geom)
      expect('kind' in key ? `${key.code}: ${key.reason}` : 'built').toBe('built')
      const k = key as IconGlyphKey
      expect(k.classes.has('4RedToBlack:0')).toBe(true)
      expect(k.classes.has('5Quarters:4')).toBe(true)
      expect(k.classOfPair.get('5Quarters:4')).toBe('5Quarters:4')
      expect([...(k.twins.get('4RedToBlack:0') ?? [])]).toEqual(['5Quarters:4'])
      expect([...(k.twins.get('5Quarters:4') ?? [])]).toEqual(['4RedToBlack:0'])
      expect(k.twins.size).toBe(2)
      // A crop of either twin classifies to the twin set, never to a pick; a third, separated class still classifies alone.
      const rect = (ref: string) => iconCropRect(decoded, geom, ref)!
      const refOf = (pair: string) => c0.markers.find(m => m.icon?.custom?.[0] && `${m.icon.custom[0].set}:${m.icon.custom[0].iconId}` === pair)!.sqref
      const twinSet = iconTwinSetClass(['4RedToBlack:0', '5Quarters:4'])
      expect(twinSet).toBe(iconTwinSetClass(['5Quarters:4', '4RedToBlack:0']))
      expect(iconTwinSetMembers(twinSet)).toEqual(['4RedToBlack:0', '5Quarters:4'])
      expect(classifyIconCrop(decoded, rect(refOf('4RedToBlack:0')), k.classes, { tolerance: 8, twins: k.twins })).toMatchObject({ glyphClass: twinSet })
      expect(classifyIconCrop(decoded, rect(refOf('5Quarters:4')), k.classes, { tolerance: 8, twins: k.twins })).toMatchObject({ glyphClass: twinSet })
      expect(classifyIconCrop(decoded, rect(refOf('3Flags:1')), k.classes, { tolerance: 8, twins: k.twins })).toEqual({ glyphClass: '3Flags:1' })
      // Without the twin relation the same crop stays ambiguous (unchanged behaviour).
      expect(classifyIconCrop(decoded, rect(refOf('4RedToBlack:0')), k.classes, { tolerance: 8 })).toMatchObject({ kind: 'refusal', code: 'icon-class-ambiguous' })
      // Unseparated from a non-twin class still refuses: drop the relation for one direction only.
      const oneWay = new Map([['5Quarters:4', new Set(['4RedToBlack:0'])]])
      expect(classifyIconCrop(decoded, rect(refOf('4RedToBlack:0')), k.classes, { tolerance: 8, twins: oneWay })).toMatchObject({ kind: 'refusal', code: 'icon-class-ambiguous' })
    }, 60_000)

    it('glyph key still refuses a crop identical to two classes', () => {
      // 3Flags:1 = 3Flags:0 + a pixel of delta 12 (a twin); 3Flags:2 = 3Flags:0 + a pixel of delta 6 at the same spot:
      // within +-8 of both, so it is the same as two classes.
      const { c0, parts, geom, decoded } = keyCapture({ '3Flags:1': { as: '3Flags:0', faint: { delta: 12 } }, '3Flags:2': { as: '3Flags:0', faint: { delta: 6 } } })
      expect(buildIconGlyphKey(c0, parts, decoded, geom)).toMatchObject({ kind: 'refusal', code: 'icon-class-ambiguous' })
    }, 60_000)

    it('a twin-set observation verifies when exactly one predicted class is in the set', () => {
      const tw = iconTwinSetClass(['A', 'B'])
      const hyps = { B2: [{ id: 'h1', glyphClass: 'A' }, { id: 'h2', glyphClass: 'C' }], B3: [{ id: 'h1', glyphClass: 'C' }, { id: 'h2', glyphClass: 'C' }] }
      expect(discriminateIcon(hyps, { B2: tw, B3: 'C' })).toEqual({ status: 'discriminating', matched: 'h1' })
      // A twin set holding no predicted class matches no hypothesis (a finding, never a pick).
      expect(discriminateIcon(hyps, { B2: iconTwinSetClass(['B', 'D']), B3: 'C' })).toMatchObject({ status: 'none' })
    })

    it('a twin-set observation is non-discriminating when both twins are predicted', () => {
      const tw = iconTwinSetClass(['A', 'B'])
      // Ruling S4-TWINSET-CONSISTENCY: two hypotheses that differ only at the ambiguous cell, both inside its set, are
      // both consistent: non-discriminating.
      const only = { B2: [{ id: 'h1', glyphClass: 'A' }, { id: 'h2', glyphClass: 'B' }], B3: [{ id: 'h1', glyphClass: 'C' }, { id: 'h2', glyphClass: 'C' }] }
      expect(discriminateIcon(only, { B2: tw, B3: 'C' }).status).toBe('non-discriminating')
    })

    describe('Ruling S4-TWINSET-CONSISTENCY (fix round 5): discrimination by set consistency', () => {
      const tw = iconTwinSetClass(['A', 'B'])
      it('two hypotheses that differ only at an ambiguous cell, both inside its set, are non-discriminating', () => {
        const hyps = { B2: [{ id: 'h1', glyphClass: 'A' }, { id: 'h2', glyphClass: 'B' }], B3: [{ id: 'h1', glyphClass: 'C' }, { id: 'h2', glyphClass: 'C' }] }
        expect(discriminateIcon(hyps, { B2: tw, B3: 'C' })).toMatchObject({ status: 'non-discriminating' })
      })
      it('a clean separating cell beside the ambiguous one discriminates for the right hypothesis', () => {
        const hyps = { B2: [{ id: 'h1', glyphClass: 'A' }, { id: 'h2', glyphClass: 'B' }], B3: [{ id: 'h1', glyphClass: 'C' }, { id: 'h2', glyphClass: 'D' }] }
        expect(discriminateIcon(hyps, { B2: tw, B3: 'C' })).toEqual({ status: 'discriminating', matched: 'h1' })
        expect(discriminateIcon(hyps, { B2: tw, B3: 'D' })).toEqual({ status: 'discriminating', matched: 'h2' })
      })
      it('a hypothesis whose prediction lies outside the observed set at any cell is inconsistent', () => {
        const hyps = { B2: [{ id: 'h1', glyphClass: 'A' }, { id: 'h2', glyphClass: 'E' }], B3: [{ id: 'h1', glyphClass: 'C' }, { id: 'h2', glyphClass: 'C' }] }
        expect(discriminateIcon(hyps, { B2: tw, B3: 'C' })).toEqual({ status: 'discriminating', matched: 'h1' })
        // Both outside: none consistent (icon-hypothesis-none upstream).
        expect(discriminateIcon({ B2: [{ id: 'h1', glyphClass: 'E' }, { id: 'h2', glyphClass: 'F' }] }, { B2: tw })).toMatchObject({ status: 'none' })
        // A null prediction is outside a set without no-icon.
        expect(discriminateIcon({ B2: [{ id: 'h1', glyphClass: null }, { id: 'h2', glyphClass: 'A' }] }, { B2: tw })).toEqual({ status: 'discriminating', matched: 'h2' })
      })
    })

    it('a twin-set observation does not refuse the case', () => {
      const c = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-thresholds')!
      const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.input))))!
      const empty = new Map<string, DecodedPng>()
      const key: IconGlyphKey = { crops: empty, classes: empty, classOfPair: new Map(PAIRS.map(x => [x, x])), twins: new Map(), faint: new Set() }
      const obs = (ref: string, glyphClass: string | null): CfObservedIcon => ({ ref, cellLeftPx: 0, cellTopPx: 0, cellWidthPx: 60, cellHeightPx: 20, icon: glyphClass === null ? null : { leftPx: 3, topPx: 2, rightPx: 18, bottomPx: 17, glyphClass } })
      const cap = 'icon:thresholds-num'
      const probe = evaluateIconCapabilities(c, parts, new Map(), key)[cap]!
      expect(probe.hypotheses).toBeDefined()
      const hyps = probe.hypotheses!
      const winner = hyps[Object.keys(hyps)[0]!]![0]!.id
      const predictedBy = (ref: string) => hyps[ref]!.find(h => h.id === winner)!.glyphClass
      const asModelled = new Map(Object.keys(hyps).map(ref => [ref, obs(ref, predictedBy(ref))]))
      const baseline = evaluateIconCapabilities(c, parts, asModelled, key)[cap]!
      expect(baseline.status).toBe('discriminating')
      // One separating designated cell observed as the twin set {its predicted class, an unpredicted twin}: same outcome,
      // no refusal.
      const at = Object.keys(hyps).find(ref => predictedBy(ref) !== null && hyps[ref]!.some(h => h.glyphClass !== null && h.glyphClass !== predictedBy(ref)))!
      expect(at).toBeDefined()
      const p = predictedBy(at)!
      const twinObserved = new Map(asModelled)
      twinObserved.set(at, obs(at, iconTwinSetClass([p, 'Z:twin'])))
      const got = evaluateIconCapabilities(c, parts, twinObserved, key)[cap]!
      expect(got.status).toBe('discriminating')
      expect(got.matched).toBe(baseline.matched)
      // A twin set holding two predicted classes never refuses the case; the clean cells still decide (Ruling
      // S4-TWINSET-CONSISTENCY), so the same hypothesis wins.
      const other = hyps[at]!.find(h => h.glyphClass !== null && h.glyphClass !== p)?.glyphClass
      expect(other).toBeDefined()
      twinObserved.set(at, obs(at, iconTwinSetClass([p, other!])))
      const both = evaluateIconCapabilities(c, parts, twinObserved, key)[cap]!
      expect(both.status).toBe('discriminating')
      expect(both.matched).toBe(baseline.matched)
    })

    it('S4-FAINT: a faint key glyph (d = 34 against its control) is admitted as a class twinned with no icon', () => {
      const { c0, parts, geom, decoded } = keyCapture({ '5Boxes:0': { onlyGrey: 34 } })
      const key = buildIconGlyphKey(c0, parts, decoded, geom)
      expect('kind' in key ? `${key.code}: ${key.reason}` : 'built').toBe('built')
      const k = key as IconGlyphKey
      expect(k.classes.has('5Boxes:0')).toBe(true)
      expect([...k.faint]).toEqual(['5Boxes:0'])
      expect(iconGlyphFaintRecord(k)).toEqual(['5Boxes:0'])
      // The faint key cell itself measures blank at the 64 ink threshold, yet classifies to {no icon, 5Boxes:0}, never to
      // plain null; a strong glyph still classifies alone.
      const measured = measureCaseIcons(c0, decoded, geom, k) as ReadonlyMap<string, CfObservedIcon>
      const refOf = (pair: string) => c0.markers.find(m => m.icon?.custom?.[0] && `${m.icon.custom[0].set}:${m.icon.custom[0].iconId}` === pair)!.sqref
      const faintCell = measured.get(refOf('5Boxes:0'))!
      expect(faintCell.icon).toBeNull()
      expect(observedIconClass(faintCell)).toBe(iconTwinSetClass(['5Boxes:0', ICON_NO_ICON_MEMBER]))
      expect(observedIconClass(measured.get(refOf('3Flags:1'))!)).toBe('3Flags:1')
      // A blank control is plain no icon.
      const ctrl = c0.expected.cells.find(x => x.ref === refOf('5Boxes:0'))!.controlRef
      expect(observedIconClass({ ...faintCell, ref: ctrl, faintGlyphClass: undefined })).toBeNull()
    }, 60_000)

    it('S4-FAINT: a key glyph within tolerance of its control (d = 5) still refuses icon-key-mismatch', () => {
      const { c0, parts, geom, decoded } = keyCapture({ '5Boxes:0': { onlyGrey: 5 } })
      expect(buildIconGlyphKey(c0, parts, decoded, geom)).toMatchObject({ kind: 'refusal', code: 'icon-key-mismatch' })
    }, 60_000)

    it('S4-FAINT: a faint twin set verifies against a non-faint competitor and is non-discriminating against a no-icon one', () => {
      const tw = iconTwinSetClass(['5Boxes:0', ICON_NO_ICON_MEMBER])
      // Authored-style: D2 predicts 5Boxes:0 (as modelled) against 3Arrows:0 (another reading); D3 agrees for both.
      expect(discriminateIcon({ D2: [{ id: 'as-modelled', glyphClass: '5Boxes:0' }, { id: 'other', glyphClass: '3Arrows:0' }], D3: [{ id: 'as-modelled', glyphClass: '5Boxes:1' }, { id: 'other', glyphClass: '5Boxes:1' }] }, { D2: tw, D3: '5Boxes:1' })).toEqual({ status: 'discriminating', matched: 'as-modelled' })
      // Against a reading that predicts no icon at D2, the faint set holds both outcomes: on D2 alone, withheld, never
      // refused; a clean D3 that refutes the no-icon reading decides it (Ruling S4-TWINSET-CONSISTENCY).
      expect(discriminateIcon({ D2: [{ id: 'as-modelled', glyphClass: '5Boxes:0' }, { id: 'no-icon', glyphClass: null }] }, { D2: tw }).status).toBe('non-discriminating')
      expect(discriminateIcon({ D2: [{ id: 'as-modelled', glyphClass: '5Boxes:0' }, { id: 'no-icon', glyphClass: null }], D3: [{ id: 'as-modelled', glyphClass: '5Boxes:1' }, { id: 'no-icon', glyphClass: null }] }, { D2: tw, D3: '5Boxes:1' })).toEqual({ status: 'discriminating', matched: 'as-modelled' })
      // A null prediction alone inside the set resolves to no icon.
      expect(discriminateIcon({ D2: [{ id: 'a', glyphClass: null }, { id: 'b', glyphClass: '3Arrows:0' }] }, { D2: tw })).toEqual({ status: 'discriminating', matched: 'a' })
    })

    it('verify-only re-derives dependent Stage 4 sidecars when the glyph key changes', () => {
      const root = mkdtempSync(join(tmpdir(), 'ov-cf-icon-twin-'))
      s4dirs.push(root)
      mkdirSync(join(root, 'tests/fixtures/xlsx-cf/inputs'), { recursive: true })
      mkdirSync(join(root, 'tests/fixtures/xlsx-cf/refs'), { recursive: true })
      const { c0, parts, geom, png, decoded } = keyCapture({ '5Quarters:4': { as: '4RedToBlack:0', faint: { delta: 32 } } })
      const dep = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-authored')!
      for (const f of [c0.input, c0.saved, dep.input, dep.saved]) cpSync(resolve(REPO_ROOT, f), join(root, f))
      const depParts = readZipPartsSync(new Uint8Array(readFileSync(join(root, dep.saved))))!
      const depGeom = computeWorksheetRangeGeometry(depParts, dep.range)!
      // The dependent capture shows 3Arrows:0's synthetic glyph in every designated icon cell (controls stay blank).
      const ruled = (ref: string) => dep.markers.some(m => isCellInSqref(ref, m.sqref))
      const depPng = drawSheet(depGeom, ctx => { for (const e of dep.expected.cells.filter(x => x.iconApplied === true && ruled(x.ref))) glyph(ctx, boxIn(depGeom, e.ref), PAIRS.indexOf('3Arrows:0')) })
      const depDecoded = decodePngSync(depPng)!
      const measuredCase = (c: CfOracleCase, pngBytes: Uint8Array): CfOracleCase => {
        const savedSha256 = sha256(new Uint8Array(readFileSync(join(root, c.saved))))
        const inputSha256 = sha256(new Uint8Array(readFileSync(join(root, c.input))))
        return { ...c, status: 'measured', resavedByExcel: true, resaveReceipt: { resavedAt: '2026-11-15T00:00:00Z', excelVersion: '16.114', inputSha256, savedSha256, openedRepairFree: true }, captures: [{ ...c.captures[0]!, pngSha256: sha256(pngBytes) }], expected: { ...c.expected, inputSha256, savedSha256, provenance: 'excel-observed' } }
      }
      const writeSidecar = (c: CfOracleCase, p: Map<string, Uint8Array>, d: DecodedPng, g: WorksheetRangeGeometry, pngBytes: Uint8Array, withExtras: boolean) => {
        const extra = buildIconSidecarExtras(c, p, d, g, { lockBefore: 'unlocked', lockAfter: 'unlocked', rootDir: root })
        const savedSha256 = sha256(new Uint8Array(readFileSync(join(root, c.saved))))
        writeFileSync(join(root, c.captures[0]!.sidecarPath), JSON.stringify({
          excelVersion: '16.114', route: 'copy-picture', appearance: 'screen', scale: 1, pxPerPt: 1, width: d.width, height: d.height,
          colWidthsPt: g.colWidthsPt, rowHeightsPt: g.rowHeightsPt, geometrySource: 'excel-measured', resavedByExcel: true, resaveSavedSha256: savedSha256,
          sheet: c.sheet, range: c.range, hostTimeZone: 'Asia/Taipei', savedSha256, pngSha256: sha256(pngBytes), capturedAt: '2026-11-15T00:00:00Z',
          ...(withExtras ? extra.extras : { zoom: 100, dpr: 1, lockBefore: 'unlocked', lockAfter: 'unlocked' }),
        }, null, 2) + '\n')
        return extra
      }
      writeFileSync(join(root, c0.captures[0]!.pngPath), png)
      writeFileSync(join(root, dep.captures[0]!.pngPath), depPng)
      const keyEntry = { ...measuredCase(c0, png), status: 'recipe-pending' as const }
      const depEntry = measuredCase(dep, depPng)
      // Manifest order as committed: the dependent precedes the key.
      const manifest: CfManifest = { version: 1, excelBuild: '16.114', locale: 'zh-TW', dateSystem: '1900', hostTimeZone: 'Asia/Taipei', generatedAt: '2026-11-06T00:00:00Z', cases: [depEntry, keyEntry] }
      writeFileSync(join(root, 'tests/fixtures/xlsx-cf/manifest.json'), JSON.stringify(manifest))

      // Today's state: the key sidecar was written by a refused key build (no glyphClasses / glyphTwins), so no key is
      // restored and the dependent was derived keyless.
      writeSidecar(c0, parts, decoded, geom, png, false)
      expect(loadIconGlyphKey(root)).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
      writeSidecar(depEntry, depParts, depDecoded, depGeom, depPng, true)
      const keyless = JSON.parse(readFileSync(join(root, dep.captures[0]!.sidecarPath), 'utf8')) as { observedIcons: CfObservedIcon[] }
      expect(keyless.observedIcons.map(o => o.icon?.glyphClass ?? null)).toContain('unclassified')
      expect(verifyCfCase(depEntry, { rootDir: root })).toMatchObject({ kind: 'verified' })

      // verify-only processes the glyph key first, whatever the manifest order.
      expect(verifyOnlyCaseOrder(manifest.cases, 4).map(i => manifest.cases[i]!.id)).toEqual(['stage4-icon-glyph-key', 'stage4-icon-authored'])
      // Re-deriving the key sidecar persists the twin relation; the key is then restored from it, twins included.
      writeSidecar(c0, parts, decoded, geom, png, true)
      const keySidecar = JSON.parse(readFileSync(join(root, c0.captures[0]!.sidecarPath), 'utf8')) as { glyphTwins?: Record<string, string[]> }
      expect(keySidecar.glyphTwins).toEqual({ '4RedToBlack:0': ['5Quarters:4'], '5Quarters:4': ['4RedToBlack:0'] })
      const restored = loadIconGlyphKey(root)
      expect(restored && !('kind' in restored) ? [...restored.twins.get('4RedToBlack:0')!] : restored).toEqual(['5Quarters:4'])
      // The key change invalidates the keyless dependent sidecar; re-deriving it (no Excel) classifies and verifies.
      expect(verifyCfCase(depEntry, { rootDir: root })).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
      writeSidecar(depEntry, depParts, depDecoded, depGeom, depPng, true)
      const keyed = JSON.parse(readFileSync(join(root, dep.captures[0]!.sidecarPath), 'utf8')) as { observedIcons: CfObservedIcon[] }
      expect(keyed.observedIcons.some(o => o.icon?.glyphClass === '3Arrows:0')).toBe(true)
      expect(keyed.observedIcons.every(o => o.icon?.glyphClass !== 'unclassified')).toBe(true)
      expect(verifyCfCase(depEntry, { rootDir: root })).toMatchObject({ kind: 'verified' })

      // Fix round 6 (review minor): a malformed twin relation or twin-set string in a sidecar refuses sidecar-mismatch,
      // named as malformed, before any comparison.
      const keyPath = join(root, c0.captures[0]!.sidecarPath)
      const goodKey = readFileSync(keyPath, 'utf8')
      for (const bad of [{ '4RedToBlack:0': [] }, { '4RedToBlack:0': [''] }, { '4RedToBlack:0': ['5Quarters:4'] }, { '': ['x'] }]) {
        writeFileSync(keyPath, JSON.stringify({ ...JSON.parse(goodKey), glyphTwins: bad }))
        const got = loadIconGlyphKey(root)
        expect(got, JSON.stringify(bad)).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
        expect((got as { reason: string }).reason, JSON.stringify(bad)).toMatch(/malformed/)
      }
      writeFileSync(keyPath, goodKey)
      const depPath = join(root, dep.captures[0]!.sidecarPath)
      const goodDep = readFileSync(depPath, 'utf8')
      for (const bad of ['twin-set:', 'twin-set:3Arrows:0', 'twin-set:3Arrows:0|']) {
        const tampered = JSON.parse(goodDep) as { observedIcons: CfObservedIcon[] }
        tampered.observedIcons.find(o => o.icon)!.icon!.glyphClass = bad
        writeFileSync(depPath, JSON.stringify(tampered))
        const got = verifyCfCase(depEntry, { rootDir: root })
        expect(got, bad).toMatchObject({ kind: 'refusal', code: 'sidecar-mismatch' })
        expect((got as { reason: string }).reason, bad).toMatch(/malformed/)
      }
      writeFileSync(depPath, goodDep)
      expect(verifyCfCase(depEntry, { rootDir: root })).toMatchObject({ kind: 'verified' })
    }, 120_000)

    it('iconTwinSetMembers refuses a twin-set string with fewer than two non-empty members (review minor)', () => {
      expect(iconTwinSetMembers(iconTwinSetClass(['A', 'B']))).toEqual(['A', 'B'])
      for (const bad of ['twin-set:', 'twin-set:A', 'twin-set:A|', 'twin-set:|B', 'twin-set:A|A', 'twin-set:||']) expect(iconTwinSetMembers(bad), bad).toBeUndefined()
      // Fix round 9 (review minor): only the canonical, strictly increasing order is a set.
      expect(iconTwinSetMembers('twin-set:B|A')).toBeUndefined()
      expect(iconTwinSetMembers('twin-set:A|C|B')).toBeUndefined()
      expect(iconSidecarMalformed({ observedIcons: [{ ref: 'B2', icon: { glyphClass: 'twin-set:B|A' } }] })).toMatch(/twin set/)
      expect(iconTwinSetMembers('3Arrows:0')).toBeUndefined()
      expect(iconTwinSetMembers(null)).toBeUndefined()
      // A malformed observation is no set at all: it matches no hypothesis (not even one predicting '').
      expect(discriminateIcon({ B2: [{ id: 'h1', glyphClass: '' }, { id: 'h2', glyphClass: 'A' }] }, { B2: 'twin-set:' })).toMatchObject({ status: 'none' })
    })
  })

  it('the icon template diff catches a template that departs from an independently written Excel form', async () => {
    // The 3Stars rule of stage4-icon-authored as Excel's x14 form is expected to be written, typed here by hand (not
    // produced by iconRuleXml): a wrong template fails this test.
    const excelForm = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData/><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="iconSet" priority="1" id="{00000000-0000-4000-8000-000000000001}"><x14:iconSet iconSet="3Stars"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>B2:B6</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>'
    const authored = new Map([['xl/worksheets/sheet1.xml', new TextEncoder().encode(excelForm)]])
    const diffs = diffAuthoredIconTemplates(readZipPartsSync(await generateCaseArchive('stage4-icon-authored'))!, authored)
    expect(diffs.filter(d => d.startsWith('B2:B6'))).toEqual([])
  })

  it('build refuses on a locked display before it opens any workbook (review fix 3)', () => {
    const c = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-om-reachability')!
    expect(() => assertBuildDisplayUnlocked(() => true)).toThrow(/display-locked/)
    expect(() => assertBuildDisplayUnlocked(() => false)).not.toThrow()
    // probeAndBuildCase checks first: with a locked display it throws before copying or opening the input.
    expect(() => probeAndBuildCase(c, new Set(), false, () => true)).toThrow(/display-locked/)
  })
})

describe('Stage 5 formula oracle (Task 31)', () => {
  const s5dirs: string[] = []
  afterEach(() => {
    for (const d of s5dirs.splice(0)) rmSync(d, { recursive: true, force: true })
  })
  const loadManifest = (): CfManifest => JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest

  it('stage 5 cases join the manifest and stage 1–4 cases are unchanged', async () => {
    const manifest = loadManifest()
    expect(manifest.cases.map(c => c.id)).toEqual(CF_CASES.map(c => c.id))
    // Stage 1-4 catalogues and capability ids are exactly what Stage 4 shipped.
    expect([CF_STAGE1_CAPABILITIES.length, CF_STAGE2_CAPABILITIES.length, CF_STAGE3_CAPABILITIES.length, CF_STAGE4_CAPABILITIES.length]).toEqual([115, 46, 73, 145])
    expect([CF_STAGE1_CASES.length, CF_STAGE2_CASES.length, CF_STAGE3_CASES.length, CF_STAGE4_CASES.length]).toEqual([13, 8, 15, 19])
    expect(CF_CASES.slice(0, 55).map(c => c.id)).toEqual([...CF_STAGE1_CASES, ...CF_STAGE2_CASES, ...CF_STAGE3_CASES, ...CF_STAGE4_CASES].map(c => c.id))
    expect(CF_CASES.slice(55).map(c => c.id)).toEqual(CF_STAGE5_CASES.map(c => c.id))
    expect(CF_CAPABILITIES.length).toBe(CF_STAGE1_CAPABILITIES.length + CF_STAGE2_CAPABILITIES.length + CF_STAGE3_CAPABILITIES.length + CF_STAGE4_CAPABILITIES.length + CF_STAGE5_CAPABILITIES.length)
    expect(new Set(CF_CAPABILITIES.map(c => c.capabilityId)).size).toBe(CF_CAPABILITIES.length)
    // Stage 1-4 generated inputs are byte-identical to fc61c6c6's generator (digest of every case's archive hash).
    const all = createHash('sha256')
    for (const c of CF_CASES.filter(x => x.stage < 5)) all.update(`${c.id}:${sha256(await generateCaseArchive(c.id))}\n`)
    expect(all.digest('hex')).toBe('c1daae80753c0dea2a10f782ac118a221f2df7fe1de83156e1a4bc08bcd89b36')

    // The brief's matrix, row by row (CF_STAGE5_GROUPS maps a row to its workbooks).
    const required: Record<string, string[]> = {
      'stage5-formula-authored': ['formula:dialect-base', 'formula:dialect-x14'],
      'stage5-formula-truthiness': ['formula:truth-number', 'formula:truth-boolean', 'formula:truth-text', 'formula:truth-blank', 'formula:truth-error', 'formula:truth-range', 'formula:truth-range-nested'],
      'stage5-formula-anchor': ['formula:anchor-single', 'formula:anchor-mixed-abs', 'formula:anchor-multi-range'],
      'stage5-formula-edge': ['formula:relative-wrap'],
      'stage5-formula-functions': ['formula:fn-logical', 'formula:fn-information', 'formula:fn-text', 'formula:fn-math', 'formula:fn-aggregate', 'formula:fn-conditional-aggregate', 'formula:fn-date', 'formula:fn-row-column'],
      'stage5-formula-hoist': ['formula:hoist-aggregate', 'formula:hoist-countif-frequency'],
      'stage5-cellis-operands': ['formula:cellis-operand-reference', 'formula:cellis-operand-relative', 'formula:cellis-operand-expression', 'formula:cellis-operand-text', 'formula:cellis-operand-blank', 'formula:cellis-operand-error'],
      'stage5-predicate-formulas': ['formula:predicate-divergent', 'formula:predicate-timeperiod-divergent'],
      'stage5-clock-today': ['formula:fn-today'],
      'stage5-calc-manual': ['formula:calc-manual-saved-values', 'formula:calc-manual-today'],
      'stage5-composition': ['formula:composition-stub-successor-dxf', 'formula:composition-stub-successor-scale', 'formula:composition-stub-successor-bar', 'formula:composition-stub-successor-icon'],
      'stage5-cross-sheet': ['formula:cross-sheet-reference', 'formula:cross-sheet-quoted-name', 'formula:cross-sheet-missing', 'formula:cellis-operand-cross-sheet'],
      'stage5-cfvo-formulas': ['scale:thresholds-formula', 'bar:bounds-formula', 'icon:thresholds-formula', 'formula:cfvo-reference', 'formula:cfvo-expression', 'formula:cfvo-relative', 'formula:cfvo-nonnumeric'],
      'stage5-cfvo-scope': ['formula:cfvo-scope-multi-range'],
      'stage5-x14-scale': ['scale:dialect-standalone'],
      'stage5-active-present': ['formula:active-present-scale', 'formula:active-present-bar', 'formula:active-present-icon'],
      'stage5-unresolved-shapes': [],
    }
    expect(Object.keys(CF_STAGE5_GROUPS).sort()).toEqual(Object.keys(required).sort())
    const grouped = Object.values(CF_STAGE5_GROUPS).flat()
    expect([...grouped].sort()).toEqual(CF_STAGE5_CASES.map(c => c.id).sort())
    for (const [row, caps] of Object.entries(required)) {
      const owned = new Set(CF_STAGE5_CAPABILITIES.filter(e => CF_STAGE5_GROUPS[row]!.includes(e.caseId) || CF_STAGE5_REUSED_CAPTURES.some(r => r.capabilityId === e.capabilityId && row === 'stage5-composition')).map(e => e.capabilityId))
      for (const cap of caps) expect(owned.has(cap), `${row} registers ${cap}`).toBe(true)
    }
    // At least one composition-* capability of its own beside the reused stub successors.
    expect(CF_STAGE5_CAPABILITIES.filter(e => e.caseId === 'stage5-composition' && /^formula:composition-(?!stub)/.test(e.capabilityId)).length).toBeGreaterThan(3)
    // The unresolved shapes lift nothing: one workbook per shape, every capability refused-unsupported.
    for (const id of CF_STAGE5_GROUPS['stage5-unresolved-shapes']!) {
      const c = CF_STAGE5_CASES.find(x => x.id === id)!
      expect(c.status, id).toBe('unsupported')
      expect(c.markers, id).toHaveLength(1)
    }
    for (const id of [...CF_STAGE5_GROUPS['stage5-active-present']!]) expect(CF_STAGE5_CASES.find(x => x.id === id)!.markers.filter(m => m.activePresent === true), id).toHaveLength(1)
    // Reused Stage 1-4 captures name existing cases and cells.
    for (const r of CF_STAGE5_REUSED_CAPTURES) {
      const c = CF_CASES.find(x => x.id === r.caseId)
      expect(c?.stage, r.caseId).toBeLessThan(5)
      for (const ref of r.refs) expect(c!.expected.cells.some(e => e.ref === ref), `${r.caseId}:${ref}`).toBe(true)
    }
    // Every Stage 5 case is preparation only: nothing is claimed.
    for (const c of manifest.cases.filter(x => x.stage === 5)) {
      expect(['recipe-pending', 'unsupported'], c.id).toContain(c.status)
      expect(c.expected.provenance, c.id).toBe('seed-hypothesis')
      expect(c.expected.cells.every(e => e.fillSource === 'unverified' && e.matched === false), c.id).toBe(true)
      expect(c.captures.every(cap => cap.pngSha256 === undefined), c.id).toBe(true)
      expect(c.markers.every(m => m.part === 'xl/worksheets/sheet1.xml'), c.id).toBe(true)
      const res = verifyCfCase(c, { rootDir: REPO_ROOT })
      expect(res.kind, c.id).toBe('refusal')
      expect(res.kind === 'refusal' && res.structuralValid, `${c.id}: ${res.kind === 'refusal' ? res.reason : ''}`).toBe(true)
    }
    for (const id of Object.keys({ ...CF_STAGE5_UNSUPPORTED, ...CF_STAGE5_NONDISCRIMINATING })) expect(CF_STAGE5_CAPABILITIES.some(e => e.capabilityId === id), id).toBe(true)
  }, 120_000)

  it('stage 5 hypotheses are fixed before capture and pairwise discriminating by design (Lessons 3, 14)', () => {
    for (const fx of CF_STAGE5_FIXTURES) {
      const c = CF_STAGE5_CASES.find(x => x.id === fx.id)!
      expect(c.formulaConfigurations, fx.id).toEqual(fx.configurations)
      const byCap = new Map<string, S5Configuration[]>()
      for (const cfg of fx.configurations) byCap.set(cfg.capabilityId, [...(byCap.get(cfg.capabilityId) ?? []), cfg])
      for (const [cap, cfgs] of byCap) {
        const live = cfgs.filter(x => x.record === undefined)
        if (live.length === 0) continue
        const names = Object.keys(live[0]!.hypotheses).sort()
        expect(names.length, `${fx.id} ${cap}`).toBeGreaterThanOrEqual(2)
        for (const cfg of live) expect(Object.keys(cfg.hypotheses).sort(), `${fx.id} ${cfg.id} names`).toEqual(names)
        for (const cfg of cfgs) {
          for (const preds of Object.values(cfg.hypotheses)) expect(Object.keys(preds).sort(), `${fx.id} ${cfg.id} predicts every ref`).toEqual([...cfg.refs].sort())
        }
        // Every pair of hypotheses predicts different tokens at >= 2 refs of one configuration (twin tokens differ by twin
        // cell, which the capture then decides; S4-REPLICATE).
        for (let i = 0; i < names.length; i++) {
          for (let j = i + 1; j < names.length; j++) {
            const separated = live.some(cfg => cfg.refs.filter(ref => cfg.hypotheses[names[i]!]![ref] !== cfg.hypotheses[names[j]!]![ref]).length >= 2)
            expect(separated, `${fx.id} ${cap}: ${names[i]} vs ${names[j]} separated at >= 2 cells of one configuration`).toBe(true)
          }
        }
      }
      // Each designated ref has a same-content control no rule covers, inside the capture range.
      const cellRefs = new Set(c.expected.cells.map(e => e.ref))
      for (const cfg of fx.configurations) {
        for (const raw of cfg.refs) {
          const ref = raw.replace(/^(pre|calc|post):/, '')
          expect(cellRefs.has(ref), `${fx.id} ${cfg.id} ${ref} is an expected cell`).toBe(true)
          const exp = c.expected.cells.find(e => e.ref === ref)!
          expect(c.markers.some(m => isCellInSqref(exp.controlRef, m.sqref)), `${fx.id}: control ${exp.controlRef} of ${ref} is uncovered`).toBe(false)
        }
      }
    }
  })

  it('formula marker pins rule formula, operands, cfvo f, dialect and activePresent exactly', () => {
    const X14NS = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
    const sheet = (legacy: string, x14 = ''): string =>
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${X14NS}><sheetData/>${legacy}${x14 ? `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>${x14}</x14:conditionalFormattings></ext></extLst>` : ''}</worksheet>`
    const styles = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>'
    const parts = (legacy: string, x14 = '', part = 'xl/worksheets/sheet1.xml'): Map<string, Uint8Array> =>
      new Map([[part, new TextEncoder().encode(sheet(legacy, x14))], ['xl/styles.xml', new TextEncoder().encode(styles)]])
    const expr = (f: string, sqref = 'B2:D9'): string => `<conditionalFormatting sqref="${sqref}"><cfRule type="expression" dxfId="0" priority="1"><formula>${f}</formula></cfRule></conditionalFormatting>`
    const between = (a: string, b: string): string => `<conditionalFormatting sqref="H2:H9"><cfRule type="cellIs" dxfId="0" priority="1" operator="between"><formula>${a}</formula><formula>${b}</formula></cfRule></conditionalFormatting>`
    const exprMarker: CfPartMarker = { part: 'xl/worksheets/sheet1.xml', extension: 'legacy', capabilityId: 'formula:dialect-base', sqref: 'B2:D9', ruleType: 'expression', priority: 1, formulas: ['$B2>$C2'], dxf: { fillRgb: 'FFFFC7CE' }, formula: { kind: 'expression', anchor: 'B2', dialect: 'base' } }
    const ok = (p: Map<string, Uint8Array>, m: CfPartMarker): boolean => verifyMarkerInArchive(p, m).ok
    expect(ok(parts(expr('$B2&gt;$C2')), exprMarker)).toBe(true)
    // Semantics, not raw strings: Excel's canonical spelling (case, spaces, unneeded sheet quotes) still matches.
    expect(ok(parts(expr('$b2 &gt; $c2')), exprMarker)).toBe(true)
    expect(normalizeCfFormula("='Sheet2'!$a$1 > 0")).toBe(normalizeCfFormula('Sheet2!$A$1>0'))
    expect(normalizeCfFormula("'My Sheet'!A1>0")).toBe("'MY SHEET'!A1>0")
    expect(normalizeCfFormula('"a b"=A1')).toBe('"a b"=A1')
    // Each mutation is refused.
    expect(ok(parts(expr('B2&gt;$C2')), exprMarker), '$B2 -> B2').toBe(false)
    expect(ok(parts(expr('$B2&gt;=$C2')), exprMarker), '> -> >=').toBe(false)
    const betweenMarker: CfPartMarker = { ...exprMarker, capabilityId: 'formula:cellis-operand-reference', sqref: 'H2:H9', ruleType: 'cellIs', operator: 'between', formulas: ['$F$1', '$F$2'], formula: { kind: 'operand', dialect: 'base' } }
    expect(ok(parts(between('$F$1', '$F$2')), betweenMarker)).toBe(true)
    expect(ok(parts(between('$F$2', '$F$1')), betweenMarker), 'operand order swapped').toBe(false)
    const scaleXml = (f: string): string => `<conditionalFormatting sqref="I2:I9"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="${f}"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>`
    const scaleMarker: CfPartMarker = { part: 'xl/worksheets/sheet1.xml', extension: 'legacy', capabilityId: 'formula:cfvo-reference', sqref: 'I2:I9', ruleType: 'colorScale', priority: 1, formulas: [], scale: { cfvo: [{ type: 'num', f: '$F$1' }, { type: 'max' }], colors: ['FFF8696B', 'FF63BE7B'], dialect: 'base' }, formula: { kind: 'cfvo', dialect: 'base' } }
    expect(ok(parts(scaleXml('$F$1')), scaleMarker)).toBe(true)
    expect(ok(parts(scaleXml('$F$2')), scaleMarker), 'cfvo.f $F$1 -> $F$2').toBe(false)
    const x14Scale = (attrs = ''): string => `<x14:conditionalFormatting><x14:cfRule type="colorScale" priority="1" id="{00000000-0000-4000-8000-000000000001}"${attrs}><x14:colorScale><x14:cfvo type="num"><xm:f>Sheet2!$A$1</xm:f></x14:cfvo><x14:cfvo type="max"/><x14:color rgb="FFF8696B"/><x14:color rgb="FF63BE7B"/></x14:colorScale></x14:cfRule><xm:sqref>M2:M9</xm:sqref></x14:conditionalFormatting>`
    const x14Marker: CfPartMarker = { part: 'xl/worksheets/sheet1.xml', extension: 'x14', capabilityId: 'scale:dialect-standalone', sqref: 'M2:M9', ruleType: 'colorScale', priority: 1, formulas: [], scale: { cfvo: [{ type: 'num', f: 'Sheet2!$A$1' }, { type: 'max' }], colors: ['FFF8696B', 'FF63BE7B'], dialect: 'x14' }, formula: { kind: 'cfvo', dialect: 'x14' } }
    expect(ok(parts('', x14Scale()), x14Marker)).toBe(true)
    const baseTwin = `<conditionalFormatting sqref="M2:M9"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="Sheet2!$A$1"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>`
    expect(ok(parts(baseTwin), x14Marker), 'x14 scale moved to base').toBe(false)
    const activeMarker: CfPartMarker = { ...x14Marker, capabilityId: 'formula:active-present-scale', activePresent: true, formulas: ['$A2>0'], formula: { kind: 'activation', dialect: 'x14' } }
    const activeXml = (attrs: string): string => x14Scale(attrs).replace('<x14:colorScale>', '<xm:f>$A2&gt;0</xm:f><x14:colorScale>')
    expect(ok(parts('', activeXml(' activePresent="1"')), activeMarker)).toBe(true)
    expect(ok(parts('', activeXml('')), activeMarker), 'activePresent dropped').toBe(false)
    // sqref order is pinned once the marker pins Excel's resaved order; an unpinned marker accepts either order.
    const multi = (sqref: string): string => expr('$A5&gt;0', sqref)
    const pinned: CfPartMarker = { ...exprMarker, capabilityId: 'formula:anchor-multi-range', sqref: 'D5:D8 B2:B4', formulas: ['$A5>0'] }
    expect(ok(parts(multi('D5:D8 B2:B4')), pinned)).toBe(true)
    expect(ok(parts(multi('B2:B4 D5:D8')), pinned), 'sqref order changed').toBe(false)
    expect(ok(parts(multi('B2:B4 D5:D8')), { ...pinned, sqrefOrder: 'any' })).toBe(true)
    // A formula moved to another part is refused as such.
    const moved = verifyMarkerInArchive(parts(expr('$B2&gt;$C2'), '', 'xl/worksheets/sheet2.xml'), exprMarker)
    expect(moved.ok === false && moved.code).toBe('missing-part')
    const both = new Map([...parts(''), ['xl/worksheets/sheet2.xml', new TextEncoder().encode(sheet(expr('$B2&gt;$C2')))]])
    const wrong = verifyMarkerInArchive(both, exprMarker)
    expect(wrong.ok === false && wrong.code).toBe('wrong-part-marker')
    // The marker's own dialect must agree with where it lives.
    expect(ok(parts(expr('$B2&gt;$C2')), { ...exprMarker, formula: { kind: 'expression', anchor: 'B2', dialect: 'x14' } })).toBe(false)
  })

  describe('semantic DisplayFormat evidence needs two decisive cells and a control (Lesson 3)', () => {
    const RED = 'FFFFC7CE'
    const GREEN = 'FFC6EFCE'
    const obs = (fills: Record<string, string>): FormulaObservation => ({
      fill: ref => (fills[ref] === undefined ? undefined : { fill: fills[ref]!, strike: false }),
      same: () => undefined,
    })
    const cfg = (refs: string[], hypotheses: Record<string, Record<string, string>>, extra: Partial<S5Configuration> = {}): S5Configuration => ({ id: 'cfg', capabilityId: 'formula:truth-number', observable: 'fill', refs, hypotheses, ...extra })
    const controls = { B2: 'P2', B3: 'P3', B4: 'P4', B5: 'P5' }
    const unformatted = { P2: 'none', P3: 'none', P4: 'none', P5: 'none' }

    it('hypotheses that agree at every designated cell are unverified-nondiscriminating', () => {
      const ev = evaluateFormulaCapabilities([cfg(['B2', 'B3'], { a: { B2: RED, B3: 'none' }, b: { B2: RED, B3: 'none' } })], controls, obs({ B2: RED, B3: 'none', ...unformatted }))
      expect(ev['formula:truth-number']!.status).toBe('non-discriminating')
    })
    it('one decisive cell is a replicationShortfall refusal', () => {
      const ev = evaluateFormulaCapabilities([cfg(['B2', 'B3'], { a: { B2: RED, B3: 'none' }, b: { B2: RED, B3: RED } })], controls, obs({ B2: RED, B3: 'none', ...unformatted }))
      expect(ev['formula:truth-number']!.status).toBe('refused')
      expect(ev['formula:truth-number']!.reason).toMatch(/S4-REPLICATE/)
    })
    it('two agreeing decisive cells with one hypothesis matching verify it', () => {
      const ev = evaluateFormulaCapabilities([cfg(['B2', 'B3', 'B4'], { a: { B2: RED, B3: RED, B4: 'none' }, b: { B2: 'none', B3: 'none', B4: 'none' } })], controls, obs({ B2: RED, B3: RED, B4: 'none', ...unformatted }))
      expect(ev['formula:truth-number']).toMatchObject({ status: 'discriminating', matched: 'a' })
    })
    it('a positive whose control is also formatted is refused', () => {
      const ev = evaluateFormulaCapabilities([cfg(['B2', 'B3'], { a: { B2: RED, B3: RED }, b: { B2: 'none', B3: 'none' } })], controls, obs({ B2: RED, B3: RED, P2: RED, P3: 'none' }))
      expect(ev['formula:truth-number']).toMatchObject({ status: 'refused', code: 'formula-control-formatted' })
    })
    it('no matching hypothesis is a finding, record configurations never decide, and an unobserved cell refuses', () => {
      const none = evaluateFormulaCapabilities([cfg(['B2', 'B3'], { a: { B2: RED, B3: RED }, b: { B2: 'none', B3: 'none' } })], controls, obs({ B2: GREEN, B3: RED, ...unformatted }))
      expect(none['formula:truth-number']).toMatchObject({ status: 'refused', code: 'formula-hypothesis-none' })
      const rec = evaluateFormulaCapabilities([cfg(['B2', 'B3'], { a: { B2: RED, B3: RED }, b: { B2: 'none', B3: 'none' } }, { record: 'refused configuration' })], controls, obs({ B2: RED, B3: RED, ...unformatted }))
      expect(rec['formula:truth-number']!.status).toBe('no-hypotheses')
      const unobserved = evaluateFormulaCapabilities([cfg(['B2', 'B3'], { a: { B2: RED, B3: RED }, b: { B2: 'none', B3: 'none' } })], controls, obs({ B2: RED, ...unformatted }))
      expect(unobserved['formula:truth-number']).toMatchObject({ status: 'refused', code: 'fill-unobserved' })
    })
    it('twin tokens decide by the observed twin, and agreeing twins do not discriminate', () => {
      const twin = (fills: Record<string, string>): FormulaObservation => obs(fills)
      const c1 = cfg(['B2', 'B3'], { evaluated: { B2: 'twin:C2', B3: 'twin:C3' }, ignored: { B2: 'twin:D2', B3: 'twin:D3' } })
      const ok = evaluateFormulaCapabilities([c1], controls, twin({ B2: 'FF111111', B3: 'FF222222', C2: 'FF111111', C3: 'FF222222', D2: 'FF999999', D3: 'FF888888', ...unformatted }))
      expect(ok['formula:truth-number']).toMatchObject({ status: 'discriminating', matched: 'evaluated' })
      const flat = evaluateFormulaCapabilities([c1], controls, twin({ B2: 'FF111111', B3: 'FF222222', C2: 'FF111111', C3: 'FF222222', D2: 'FF111111', D3: 'FF222222', ...unformatted }))
      expect(flat['formula:truth-number']!.status).toBe('non-discriminating')
    })
    it('prefixed refs read their own observation set, and their controls under the same prefix', () => {
      // post: (live edit) and pre: (manual calculation) refs and their controls come from their own tables.
      const tables: Record<string, Record<string, string>> = { 'post:': { B3: RED, B5: RED, P3: 'none', P5: 'none' }, 'pre:': { B2: RED, B4: RED, P2: 'none', P4: 'none' } }
      const prefixed: FormulaObservation = {
        fill: ref => { const m = /^(post:|pre:)?(.*)$/.exec(ref)!; const v = tables[m[1] ?? '']?.[m[2]!]; return v === undefined ? undefined : { fill: v, strike: false } },
        same: () => undefined,
      }
      const live = evaluateFormulaCapabilities([cfg(['post:B3', 'post:B5'], { documented: { 'post:B3': RED, 'post:B5': RED }, stale: { 'post:B3': 'none', 'post:B5': 'none' } })], controls, prefixed)
      expect(live['formula:truth-number']).toMatchObject({ status: 'discriminating', matched: 'documented' })
      const pre = evaluateFormulaCapabilities([cfg(['pre:B2', 'pre:B4'], { saved: { 'pre:B2': RED, 'pre:B4': RED }, recomputed: { 'pre:B2': 'none', 'pre:B4': 'none' } })], controls, prefixed)
      expect(pre['formula:truth-number']).toMatchObject({ status: 'discriminating', matched: 'saved' })
      // A dynamic-refresh observation that skipped the controls refuses (the capture reads both).
      delete tables['post:']!.P3
      expect(evaluateFormulaCapabilities([cfg(['post:B3', 'post:B5'], { documented: { 'post:B3': RED, 'post:B5': RED }, stale: { 'post:B3': 'none', 'post:B5': 'none' } })], controls, prefixed)['formula:truth-number']).toMatchObject({ code: 'fill-unobserved' })
    })
    it('the capture reads every refreshed cell and its control, and the manual route every pre:/calc: cell and its control', async () => {
      const { dynamicRefreshRefs } = await import('../scripts/cf-refs/capture')
      const { manualObservationRefs } = await import('../scripts/cf-refs/build')
      const cross = CF_STAGE5_CASES.find(x => x.id === 'stage5-cross-sheet')!
      const ctrl = (c: CfOracleCase, ref: string): string => c.expected.cells.find(e => e.ref === ref)!.controlRef
      expect(dynamicRefreshRefs(cross)).toEqual(['B3', ctrl(cross, 'B3'), 'B5', ctrl(cross, 'B5')])
      const script = buildStage5CaptureScripts(cross, 'x.xlsx').dynamicRefresh[0]!.join('\n')
      for (const ref of dynamicRefreshRefs(cross)) expect(script).toContain(`range "${ref}" of ws`)
      expect(script).not.toMatch(/\bsave\b|^\s*calculate\b/m)
      const manual = CF_STAGE5_CASES.find(x => x.id === 'stage5-calc-manual')!
      const refs = manualObservationRefs(manual)
      for (const r of ['C2', 'C5', 'H2', 'H5']) expect(refs, r).toEqual(expect.arrayContaining([r, ctrl(manual, r)]))
    })
    it('replicationShortfall groups by an explicit configuration key when given one', () => {
      expect(replicationShortfall(['C1', 'D1'])).toMatch(/configuration C has 1/)
      expect(replicationShortfall(['C1', 'D1'], 2, () => 'cols')).toBeUndefined()
    })
  })

  it('a measured Stage 5 case verifies from its stored capture alone, and edited hypotheses or sidecars refuse', async () => {
    const { evaluateFormulaCase } = await import('../scripts/cf-refs/verify')
    const root = mkdtempSync(join(tmpdir(), 'ov-cf-s5-e2e-'))
    s5dirs.push(root)
    const t = CF_STAGE5_CASES.find(c => c.id === 'stage5-formula-edge')!
    const bytes = await generateCaseArchive(t.id)
    for (const d of ['tests/fixtures/xlsx-cf/inputs', 'tests/fixtures/xlsx-cf/refs']) mkdirSync(join(root, d), { recursive: true })
    writeFileSync(join(root, t.input), bytes)
    writeFileSync(join(root, t.saved), bytes)
    const parts = readZipPartsSync(bytes)!
    const geom = computeWorksheetRangeGeometry(parts, t.range)!
    const canvas = createCanvas(geom.expectedWidthPt, geom.expectedHeightPt)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, geom.expectedWidthPt, geom.expectedHeightPt)
    const png = new Uint8Array(canvas.toBuffer('image/png'))
    writeFileSync(join(root, t.captures[0]!.pngPath), png)
    // Excel "shows" the wrap reading: the overflowing targets format, the controls stay white.
    const fills: Record<string, string> = { A3: 'FFFFC7CE', A4: 'FFFFC7CE', C1: 'FFC6EFCE', D1: 'FFC6EFCE' }
    const formulaObservations = Object.fromEntries(t.expected.cells.map(e => [e.ref, { displayFormatFill: fills[e.ref] ?? 'FFFFFFFF', baseFill: 'FFFFFFFF', strikethrough: false }]))
    const h = sha256(bytes)
    const c: CfOracleCase = {
      ...t,
      status: 'measured',
      resavedByExcel: true,
      resaveReceipt: { resavedAt: '2026-11-16T00:00:00Z', excelVersion: '16.114', inputSha256: h, savedSha256: h, openedRepairFree: true },
      expected: { ...t.expected, inputSha256: h, savedSha256: h, provenance: 'excel-observed', cells: t.expected.cells.map(e => ({ ...e, fillSource: 'excel-display-format' as const, matched: fills[e.ref] !== undefined })) },
      captures: [{ ...t.captures[0]!, pngSha256: sha256(png) }],
    }
    const recorded = { formulaObservations }
    const sidecar = {
      excelVersion: '16.114', route: 'copy-picture', appearance: 'screen', scale: 1, pxPerPt: 1, width: geom.expectedWidthPt, height: geom.expectedHeightPt,
      colWidthsPt: geom.colWidthsPt, rowHeightsPt: geom.rowHeightsPt, geometrySource: 'excel-measured', resavedByExcel: true, resaveSavedSha256: h,
      sheet: c.sheet, range: c.range, hostTimeZone: 'Asia/Taipei', savedSha256: h, pngSha256: sha256(png), capturedAt: '2026-11-16T00:00:00Z',
      lockBefore: 'unlocked', lockAfter: 'unlocked', ...recorded, formulaEvaluation: evaluateFormulaCase(c, recorded, decodePngSync(png)!, geom),
    }
    const sidecarAbs = join(root, c.captures[0]!.sidecarPath)
    writeFileSync(sidecarAbs, JSON.stringify(sidecar))
    const res = verifyCfCase(c, { rootDir: root })
    expect(res.kind, res.kind === 'refusal' ? res.reason : '').toBe('verified')
    expect(res.kind === 'verified' && res.verifiedCapabilities).toEqual(['formula:relative-wrap'])
    expect(res.kind === 'verified' && res.matchedHypotheses).toEqual({ 'formula:relative-wrap': 'wrap' })
    // The Stage 5 gates hold it back: formula:truth-boolean (and formula:truth-error for the #REF! reading) are unverified.
    const manifest: CfManifest = { version: 1, excelBuild: '16.114', locale: 'zh-TW', dateSystem: '1900', hostTimeZone: 'Asia/Taipei', generatedAt: '', cases: [c] }
    expect(isCapabilityVerified('formula:relative-wrap', manifest, { rootDir: root })).toBe(false)
    // Lesson 14: hypotheses edited after the capture refuse, whatever the sidecar says.
    const edited = { ...c, formulaConfigurations: c.formulaConfigurations!.map(x => ({ ...x, hypotheses: { ...x.hypotheses, extra: x.hypotheses.wrap! } })) }
    const refusedEdit = verifyCfCase(edited, { rootDir: root })
    expect(refusedEdit.kind === 'refusal' && refusedEdit.code).toBe('unobserved-expectation')
    // A sidecar evaluation that differs from the recomputation refuses.
    writeFileSync(sidecarAbs, JSON.stringify({ ...sidecar, formulaEvaluation: {} }))
    const tampered = verifyCfCase(c, { rootDir: root })
    expect(tampered.kind === 'refusal' && tampered.code).toBe('sidecar-mismatch')
    // A locked display before the capture refuses.
    writeFileSync(sidecarAbs, JSON.stringify({ ...sidecar, lockBefore: 'locked' }))
    expect((verifyCfCase(c, { rootDir: root }) as { code?: string }).code).toBe('display-locked')
  })

  it('audit --stage 5 --write is idempotent (Lesson 4)', () => {
    const manifest = loadManifest()
    const committed = readFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
    const listed = (text: string, name: string): string[] | undefined => {
      const m = new RegExp(`export const ${name}: ReadonlySet<string> = new Set\\(\\[([\\s\\S]*?)\\]\\)`).exec(text)
      return m ? [...m[1]!.matchAll(/'([^']+)'/g)].map(x => x[1]!) : undefined
    }
    const first = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: committed })
    expect(renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: committed })).toBe(first)
    // Without --publish-stage5 the Stage 5 export is carried over verbatim: absent at first, so the file is unchanged.
    expect(first).toBe(committed)
    expect(listed(first, 'VERIFIED_STAGE5_CAPABILITIES')).toBeUndefined()
    const withS5 = committed.replace(/(export const VERIFIED_CAPABILITIES)/, "export const VERIFIED_STAGE5_CAPABILITIES: ReadonlySet<string> = new Set([\n  'formula:truth-boolean',\n])\n\n$1")
    const kept = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: withS5 })
    expect(listed(kept, 'VERIFIED_STAGE5_CAPABILITIES')).toEqual(['formula:truth-boolean'])
    expect(renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: kept })).toBe(kept)
    for (const name of ['VERIFIED_STAGE1_CAPABILITIES', 'VERIFIED_STAGE2_CAPABILITIES', 'VERIFIED_STAGE3_CAPABILITIES', 'VERIFIED_STAGE4_CAPABILITIES']) expect(listed(kept, name), name).toEqual(listed(committed, name))
    // --publish-stage5 writes the export only when the stage is complete; today nothing is captured, so it changes nothing.
    const report = auditStage5Manifest(manifest, { rootDir: REPO_ROOT })
    expect(report.ok, JSON.stringify(report.perCase.filter(p => !p.structuralValid))).toBe(true)
    expect(report.summary.cases).toBe(CF_STAGE5_CASES.length)
    expect(report.summary.capabilities).toBe(CF_STAGE5_CAPABILITIES.length)
    expect(report.perCapability.filter(p => p.disposition === 'unaccounted')).toEqual([])
    expect(report.perCapability.filter(p => p.disposition === 'verified')).toEqual([])
    expect(report.stageComplete).toBe(false)
    expect(report.blocked).toEqual(report.blockedCapabilities)
    expect(report.blocked.length).toBeGreaterThan(40)
    const unsupportedCases = new Set(CF_STAGE5_CASES.filter(c => c.status === 'unsupported').map(c => c.id))
    expect(report.perCapability.filter(p => p.disposition === 'refused-unsupported').map(p => p.capabilityId).sort()).toEqual(CF_STAGE5_CAPABILITIES.filter(e => unsupportedCases.has(e.caseId) || CF_STAGE5_UNSUPPORTED[e.capabilityId] !== undefined).map(e => e.capabilityId).sort())
    const published = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: committed, publishStage5: true })
    expect(published).toBe(committed)
    expect(renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, { committedSource: kept, publishStage5: true })).toBe(kept)
    // Stage 1-4 exports unchanged by any Stage 5 option.
    for (const name of ['VERIFIED_STAGE1_CAPABILITIES', 'VERIFIED_STAGE2_CAPABILITIES', 'VERIFIED_STAGE3_CAPABILITIES', 'VERIFIED_STAGE4_CAPABILITIES', 'VERIFIED_CAPABILITIES']) expect(listed(published, name), name).toEqual(listed(committed, name))
    // A complete synthetic report publishes exactly its gated set, sorted.
    const complete: Stage5AuditReport = { ...report, ok: true, stageComplete: true, blocked: [], blockedCapabilities: [], perCapability: [{ capabilityId: 'formula:truth-range', caseId: 'stage5-formula-truthiness', disposition: 'verified' }], matchedReadings: { 'formula:truth-range': 'range-implicit' } }
    expect(stage5Publication(complete, { 'formula:truth-range': 'range-implicit' }).published).toEqual(['formula:truth-range'])
  }, 120_000)

  it('a publication requires production\'s reading to equal Excel\'s matched reading', () => {
    const base: Stage5AuditReport = { ...auditStage5Manifest({ ...loadManifest(), cases: [] }, { rootDir: REPO_ROOT }), ok: true, stageComplete: true, blocked: [], blockedCapabilities: [] }
    const report = (matched: string): Stage5AuditReport => ({ ...base, perCapability: [{ capabilityId: 'formula:truth-range', caseId: 'stage5-formula-truthiness', disposition: 'verified' }], matchedReadings: { 'formula:truth-range': matched } })
    const prod = { 'formula:truth-range': 'range-implicit' }
    expect(stage5Publication(report('range-implicit'), prod)).toMatchObject({ published: ['formula:truth-range'] })
    const mismatch = stage5Publication(report('range-top-left'), prod)
    expect(mismatch.published).toEqual([])
    expect(mismatch.withheld['formula:truth-range']).toMatch(/reading-mismatch: Excel matched range-top-left, production implements range-implicit/)
    // Production implements no Stage 5 reading yet (Tasks 32-38 add them): every matched reading is withheld.
    expect(stage5Publication(report('range-implicit')).withheld['formula:truth-range']).toMatch(/reading-unimplemented/)
    expect(Object.keys(CF_STAGE5_PRODUCTION_READINGS)).toEqual([])
    // An unverified capability never publishes, whatever its reading.
    const unverified = stage5Publication({ ...report('range-implicit'), perCapability: [{ capabilityId: 'formula:truth-range', caseId: 'stage5-formula-truthiness', disposition: 'supported-structural' }] }, prod)
    expect(unverified.published).toEqual([])
  })

  it('the manual-calculation route emits no premature calculate (round-1 #4)', () => {
    const c = CF_STAGE5_CASES.find(x => x.id === 'stage5-calc-manual')!
    expect(c.calcObservation).toBe('pre-calculate')
    const refs = ['A2', 'C2']
    const manual = buildManualObservationScript('stage5-calc-manual.xlsx', 'Sheet1', refs)
    const firstObservation = manual.findIndex(l => /display format|value of range/.test(l))
    expect(firstObservation).toBeGreaterThan(0)
    expect(manual.slice(0, firstObservation).some(l => /^\s*calculate\b/.test(l))).toBe(false)
    expect(manual.some(l => /^\s*calculate\b/.test(l))).toBe(false)
    expect(manual.join('\n')).toMatch(/calculation/)
    expect(manual.join('\n')).toMatch(/count of workbooks/)
    const explicit = buildExplicitCalculationObservationScript('stage5-calc-manual.xlsx', 'Sheet1', refs)
    expect(explicit.filter(l => /^\s*calculate\b/.test(l))).toHaveLength(1)
    // Build: the observations come first, each its own script, and the resave carries no calculate.
    const scripts = buildCaseBuildScripts(c, '.work-stage5-calc-manual.xlsx', '/tmp/x.candidate.xlsx', refs)
    expect(scripts.map(s => s.kind)).toEqual(['manual-observation', 'explicit-calculation', 'resave'])
    expect(scripts[0]!.lines).toEqual(buildManualObservationScript('.work-stage5-calc-manual.xlsx', 'Sheet1', refs))
    expect(scripts[2]!.lines.some(l => /^\s*calculate\b/.test(l)), 'resave keeps build.ts:180').toBe(false)
    // Mutant: an ordinary case keeps the unconditional calculate, so the check has teeth.
    const ordinary = buildCaseBuildScripts(CF_STAGE5_CASES.find(x => x.id === 'stage5-formula-anchor')!, 'w.xlsx', '/tmp/y.xlsx', refs)
    expect(ordinary.map(s => s.kind)).toEqual(['resave'])
    expect(ordinary[0]!.lines.some(l => /^\s*calculate\b/.test(l))).toBe(true)
    // Capture: the clock probe skips its calculate for a pre-calculate case only.
    const cap = buildStage5CaptureScripts(c, 'stage5-calc-manual.xlsx')
    expect(cap.clock.some(l => /^\s*calculate\b/.test(l))).toBe(false)
    const clockCase = CF_STAGE5_CASES.find(x => x.id === 'stage5-clock-today')!
    expect(buildStage5CaptureScripts(clockCase, 'stage5-clock-today.xlsx').clock.some(l => /^\s*calculate\b/.test(l))).toBe(true)
  })

  it.skipIf(!hasOsacompile)('compiles the Stage 5 probe scripts under osacompile (Lesson 8)', () => {
    const scripts: string[][] = []
    for (const c of CF_STAGE5_CASES) {
      const refs = c.expected.cells.map(e => e.ref)
      const s = buildStage5CaptureScripts(c, `${c.id}.xlsx`)
      scripts.push(s.clock, ...s.dynamicRefresh)
      scripts.push(...buildCaseBuildScripts(c, `.work-${c.id}.xlsx`, `/tmp/${c.id}.candidate.xlsx`, refs.slice(0, 50)).map(x => x.lines))
    }
    for (const lines of scripts) {
      if (lines.length === 0) continue
      const full = ['tell application "Microsoft Excel"', ...lines, 'end tell'].join('\n')
      const res = spawnSync('osacompile', ['-o', '/dev/null'], { input: full, encoding: 'utf8' })
      expect(res.status, res.stderr).toBe(0)
    }
  }, 120_000)

  it.skipIf(!existsSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/excel-authored-formulas.xlsx')))('formula templates write what Excel\'s dialog writes (excel-authored-formulas.xlsx)', async () => {
    const { diffAuthoredFormulaTemplates, EXCEL_AUTHORED_FORMULAS_PATH } = await import('../scripts/cf-refs/generate')
    const authored = readZipPartsSync(new Uint8Array(readFileSync(EXCEL_AUTHORED_FORMULAS_PATH)))!
    const generated = readZipPartsSync(await generateCaseArchive('stage5-formula-authored'))!
    expect(diffAuthoredFormulaTemplates(generated, authored)).toEqual([])
  })
})
