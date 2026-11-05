import { createHash } from 'node:crypto'
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { compareCfCase } from '../scripts/cf-compare'
import { assertCanCloseWorkbook, shouldSkipCaseOnResume } from '../scripts/cf-refs/build'
import { CF_STAGE1_CASES, type CfManifest, type CfOracleCase } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { verifyCfCase } from '../scripts/cf-refs/verify'

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
    'xl/worksheets/sheet1.xml': `${XML_DECL}<worksheet xmlns="${NS_MAIN}"><dimension ref="A1:B3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Label</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Pos</t></is></c><c r="B2"><v>150</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Neg</t></is></c><c r="B3"><v>50</v></c></row></sheetData>${sheetRule}</worksheet>`,
  }
  if (opts.repaired) {
    files['xl/repairedRecords/repairedRecord1.xml'] = `${XML_DECL}<repairedRecord/>`
  }
  for (const [name, content] of Object.entries(files)) {
    zip.file(name, content, { date: ZIP_DATE, createFolders: false })
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

function makeCapturePng(positiveHighlight: boolean): Uint8Array {
  const canvas = createCanvas(120, 60)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, 120, 60)
  if (positiveHighlight) {
    ctx.fillStyle = '#ffc7ce'
    ctx.fillRect(62, 22, 56, 16)
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
        scale: 1,
        pxPerPt: 1,
        sheet: 'Sheet1',
        range: 'A1:B3',
        savedSha256: validXlsxHash,
        pngSha256: validPngHash,
      }),
    )

    const baseCase: CfOracleCase = {
      id: 'syn-cellis',
      stage: 1,
      input: 'valid.xlsx',
      saved: 'valid.xlsx',
      sheet: 'Sheet1',
      range: 'A1:B3',
      markers: [
        {
          part: 'xl/worksheets/sheet1.xml',
          extension: 'legacy',
          sqref: 'B2:B3',
          ruleType: 'cellIs',
          operator: 'greaterThan',
          priority: 1,
          formula: '100',
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
        positiveRefs: ['B2'],
        negativeRefs: ['B3'],
        cells: [
          { ref: 'B2', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true },
          { ref: 'B3', matched: false },
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
    }

    const verified = verifyCfCase(baseCase, { rootDir: dir })
    expect(verified.kind).toBe('verified')
    if (verified.kind === 'verified') {
      expect(verified.savedSha256).toBe(validXlsxHash)
      expect(verified.positiveContrastMaxDelta).toBeGreaterThan(40)
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
        scale: 1,
        pxPerPt: 1,
        sheet: 'Sheet1',
        range: 'A1:B3',
        savedSha256: validXlsxHash,
        pngSha256: baselinePngHash,
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
    expect(unchangedResult).toMatchObject({
      kind: 'refusal',
      code: 'unchanged-control-capture',
      structuralValid: true,
    })

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

  it('verifies all 11 Stage 1 manifest cases have valid structural markers and honest status', () => {
    const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
    expect(manifest.cases.map(c => c.id)).toEqual(CF_STAGE1_CASES.map(c => c.id))
    expect(manifest.cases).toHaveLength(11)

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
})
