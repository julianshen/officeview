import { createHash } from 'node:crypto'
import { cpSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { compareCfCase } from '../scripts/cf-compare'
import { assertCanCloseWorkbook, shouldSkipCaseOnResume } from '../scripts/cf-refs/build'
import {
  CF_CAPABILITIES,
  CF_CASES,
  CF_STAGE1_CAPABILITIES,
  CF_STAGE1_CASES,
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_CASES,
  type CfManifest,
  type CfOracleCase,
  type CfPartMarker,
  type CfResaveReceipt,
} from '../scripts/cf-refs/cases'
import { parseExcelColorToArgb } from '../scripts/cf-refs/capture'
import { CF_MANIFEST_PATH, generateAllStage1 } from '../scripts/cf-refs/generate'
import {
  type ComputeGeometryOptions,
  type WorksheetRangeGeometry,
  computeWorksheetRangeGeometry,
  decodePngSync,
  isCapabilityVerified,
  isSupportedExcelVersion,
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

  it('stage 2 cases join the manifest and stage 1 cases are unchanged', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    expect(manifest.cases.map(c => c.id)).toEqual(CF_CASES.map(c => c.id))
    expect(CF_STAGE1_CAPABILITIES.length).toBeGreaterThanOrEqual(90)
    expect(CF_STAGE2_CAPABILITIES.length).toBeGreaterThan(0)
    expect(CF_CAPABILITIES.length).toBe(CF_STAGE1_CAPABILITIES.length + CF_STAGE2_CAPABILITIES.length)
    expect(CF_STAGE2_CASES).toHaveLength(8)
    for (const c of manifest.cases.filter(c => c.stage === 2)) {
      expect(['recipe-pending', 'unsupported']).toContain(c.status)
      expect(c.expected.inputSha256).toMatch(/^[0-9a-f]{64}$/)
      expect(c.expected.savedSha256).toMatch(/^[0-9a-f]{64}$/)
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

    const writeSidecar = (geomSource = 'excel-measured', observedArgb = 'FFFF0000') => {
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
            { ref: 'B2', fillSource: 'excel-display-format', fillRgb: observedArgb },
            { ref: 'B3', fillSource: 'excel-display-format', fillRgb: 'FFFFFFFF' },
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
    writeSidecar('excel-measured')
    // validPng has #ffc7ce at B2, target FFFFC7CE
    const cPixelMatched = makeCase('pixel-matched')
    cPixelMatched.expected.cells[0].fillRgb = 'FFFFC7CE'
    const pixelMatchedRes = verifyCfCase(cPixelMatched, { rootDir: dir })
    expect(pixelMatchedRes.kind).toBe('verified')

    // 4. Trusted ARGB (fillSource: 'excel-display-format') verifies even when pixel advice disagrees
    writeSidecar('excel-measured', 'FFFF0000')
    const trustedArgbCase = makeCase('excel-display-format')
    trustedArgbCase.expected.cells[0].fillRgb = 'FFFF0000'
    // validPng has #ffc7ce (delta > 28 from FFFF0000), but trusted ARGB overrides pixel advice
    const trustedRes = verifyCfCase(trustedArgbCase, { rootDir: dir })
    expect(trustedRes.kind).toBe('verified')

    // 5. Trusted ARGB that differs from expectation gets cell-expectation-mismatch
    writeSidecar('excel-measured', 'FF00FF00') // observed green vs expected red
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
})
