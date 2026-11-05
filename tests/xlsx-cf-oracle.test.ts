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
  CF_STAGE1_CAPABILITIES,
  CF_STAGE1_CASES,
  type CfManifest,
  type CfOracleCase,
} from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH, generateAllStage1 } from '../scripts/cf-refs/generate'
import {
  computeWorksheetRangeGeometry,
  isCapabilityVerified,
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
    'xl/worksheets/sheet1.xml': `${XML_DECL}<worksheet xmlns="${NS_MAIN}"><dimension ref="A1:B3"/><sheetFormatPr defaultRowHeight="18" defaultColWidth="14"/><sheetData><row r="1" ht="18"><c r="A1" t="inlineStr"><is><t>Label</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2" ht="18"><c r="A2" t="inlineStr"><is><t>Pos</t></is></c><c r="B2"><v>150</v></c></row><row r="3" ht="18"><c r="A3" t="inlineStr"><is><t>Neg</t></is></c><c r="B3"><v>50</v></c></row></sheetData>${sheetRule}</worksheet>`,
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
 * Renders a synthetic `140x54` PNG for `A1:B3` (2 cols x 70px, 3 rows x 18px)
 * where `B2` (`x: 70..140, y: 18..36`) has `#ffc7ce` when `positiveHighlight` is true.
 */
function makeCapturePng(positiveHighlight: boolean): Uint8Array {
  const canvas = createCanvas(140, 54)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, 140, 54)
  if (positiveHighlight) {
    ctx.fillStyle = '#ffc7ce'
    ctx.fillRect(70, 18, 70, 18)
  }
  return new Uint8Array(canvas.toBuffer('image/png'))
}

/** Renders a synthetic capture matching every `expected.cells` entry of a `CfOracleCase` for testing. */
function makeCaseMatchingPng(c: CfOracleCase, savedBytes: Uint8Array, paintMatchedFills: boolean): Uint8Array {
  const parts = readZipPartsSync(savedBytes)!
  const geom = computeWorksheetRangeGeometry(parts, c.range)!
  const width = Math.round((geom.minExpectedWidthPt + geom.maxExpectedWidthPt) / 2)
  const height = geom.expectedHeightPt
  const canvas = createCanvas(width, height)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, width, height)

  if (paintMatchedFills) {
    const totalW = geom.colWidthChars.reduce((a, b) => a + b, 0)
    const totalH = geom.rowHeightsPt.reduce((a, b) => a + b, 0)
    for (const cell of c.expected.cells) {
      if (!cell.matched) continue
      const pos = parseA1(cell.ref)!
      const colIdx = pos.col - geom.minCol
      const rowIdx = pos.row - geom.minRow
      const colBefore = geom.colWidthChars.slice(0, colIdx).reduce((a, b) => a + b, 0)
      const colSpan = geom.colWidthChars[colIdx]
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
        width: 140,
        height: 54,
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
        width: 140,
        height: 54,
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

    // 7. Comparator reports implementation-pending before Task 8 wires CF rendering
    const cmpRes = await compareCfCase(baseCase, 'static', { rootDir: dir })
    expect(cmpRes.outcome).toBe('implementation-pending')
  })

  it('verifies all Stage 1 manifest cases have valid structural markers and honest status', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    expect(manifest.cases.map(c => c.id)).toEqual(CF_STAGE1_CASES.map(c => c.id))
    expect(manifest.cases).toHaveLength(13)

    for (const c of manifest.cases) {
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
    const baseSide = {
      sheet: c0.sheet,
      range: c0.range,
      pxPerPt: 1,
      scale: 1,
      route: 'copy-picture',
      appearance: 'screen',
      excelVersion: '16.114',
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
    const exactW = Math.round((geom.minExpectedWidthPt + geom.maxExpectedWidthPt) / 2)
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
    for (const sd of [1, 2, 3, 0, -1, 4]) {
      expect(avgCase.markers.some(m => m.stdDev === sd), `missing stdDev=${sd}`).toBe(true)
    }
    expect(avgCase.markers.some(m => m.stdDev === 1 && m.equalAverage === true)).toBe(true)

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

    // Every capability in CF_STAGE1_CAPABILITIES is currently refused because captures are pending/unsupported
    expect(CF_STAGE1_CAPABILITIES.length).toBeGreaterThanOrEqual(90)
    for (const cap of CF_STAGE1_CAPABILITIES) {
      expect(isCapabilityVerified(cap.capabilityId, manifest, { rootDir: REPO_ROOT })).toBe(false)
    }

    // 5. Clock window and fixture expiration enforcement
    const dir = mkdtempSync(join(tmpdir(), 'ov-cf-clock-'))
    dirs.push(dir)
    const clockCase = structuredClone(manifest.cases.find(c => c.id === 'stage1-timeperiod-today-clock')!)
    const clockSavedBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, clockCase.saved)))
    const clockSavedHash = sha256(clockSavedBytes)
    const clockPng = makeCaseMatchingPng(clockCase, clockSavedBytes, true)
    const clockPngHash = sha256(clockPng)
    writeFileSync(join(dir, 'clock.png'), clockPng)

    // 5a: Daytime capture (16:00 Asia/Taipei, outside 00:05-07:55 window) -> refused with clock-window-invalid
    writeFileSync(
      join(dir, 'clock-daytime.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        sheet: clockCase.sheet,
        range: clockCase.range,
        hostTimeZone: 'Asia/Taipei',
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
        sheet: clockCase.sheet,
        range: clockCase.range,
        hostTimeZone: 'Asia/Taipei',
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

    // 5c: Valid 00:05-07:55 Asia/Taipei capture on 2026-11-06 (2026-11-05T17:30:00Z = 2026-11-06 01:30 Taipei) -> verified
    writeFileSync(
      join(dir, 'clock-valid.json'),
      JSON.stringify({
        excelVersion: '16.114',
        route: 'copy-picture',
        appearance: 'screen',
        scale: 1,
        pxPerPt: 1,
        sheet: clockCase.sheet,
        range: clockCase.range,
        hostTimeZone: 'Asia/Taipei',
        savedSha256: clockSavedHash,
        pngSha256: clockPngHash,
        capturedAt: '2026-11-05T17:30:00Z',
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
  })
})
