import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  auditStage3Manifest,
  barAutoColoursAcrossCases,
  renderAutoBarColoursBlock,
  renderVerifiedCapabilitiesSource,
  verifiedAutoBarColours,
} from '../scripts/cf-refs/audit'
import { CF_STAGE3_CAPABILITIES, CF_STAGE3_CASES, type CfManifest, type CfOracleCase } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import {
  BAR_AUTO_COLOUR_CAPABILITIES,
  computeWorksheetRangeGeometry,
  decodePngSync,
  judgeBarAutoColour,
  measureCaseBars,
  readZipPartsSync,
  sampleBarAutoColours,
  type CfObservedBar,
  type DecodedPng,
} from '../scripts/cf-refs/verify'

const REPO_ROOT = resolve(__dirname, '..')
const manifest = (): CfManifest => JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest

/** A measured case with its saved parts, its screen PNG and the Excel-measured geometry. */
function load(id: string) {
  const man = manifest().cases.find(x => x.id === id)!
  const tpl = CF_STAGE3_CASES.find(x => x.id === id)!
  const c: CfOracleCase = { ...man, markers: tpl.markers, expected: { ...man.expected, cells: tpl.expected.cells } }
  const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, man.saved))))!
  const cap = man.captures.find(x => x.appearance === 'screen')!
  const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, cap.sidecarPath), 'utf8')) as Record<string, number[]>
  const geom = computeWorksheetRangeGeometry(parts, man.range, { colWidthsPt: sc.colWidthsPt!, rowHeightsPt: sc.rowHeightsPt! })!
  const png = decodePngSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, cap.pngPath))))!
  return { c, parts, geom, png }
}

/** The same saved parts with the sheet XML rewritten (an in-memory counterfactual, never written to disk). */
function withSheet(parts: Map<string, Uint8Array>, edit: (xml: string) => string): Map<string, Uint8Array> {
  const xml = new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml')!)
  const next = edit(xml)
  expect(next).not.toBe(xml)
  return new Map([...parts, ['xl/worksheets/sheet1.xml', new TextEncoder().encode(next)]])
}
const autoAll = (xml: string) =>
  xml.replace(/<x14:axisColor rgb="FF000000"\/>/g, '<x14:axisColor auto="1"/>').replace(/<x14:negativeFillColor rgb="FFFF0000"\/>/g, '<x14:negativeFillColor auto="1"/>')

const clonePng = (png: DecodedPng): DecodedPng => ({ ...png, data: new Uint8ClampedArray(png.data) })

describe('Stage 3 automatic bar colours (Ruling S3-F as amended, Task 19b)', () => {
  it('only the automatic axis and negative fill occur, and no kept capture draws either', () => {
    expect(Object.keys(BAR_AUTO_COLOUR_CAPABILITIES).sort()).toEqual(['bar:color-auto-axis', 'bar:color-auto-negative-fill'])
    expect(CF_STAGE3_CAPABILITIES.filter(e => e.capabilityId.startsWith('bar:color-auto-')).map(e => e.caseId))
      .toEqual(['stage3-bar-auto-colors', 'stage3-bar-auto-colors'])
    const across = barAutoColoursAcrossCases(manifest(), REPO_ROOT)
    for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES)) {
      expect(across[cap]?.samples ?? [], cap).toEqual([])
      expect(judgeBarAutoColour(across[cap]).status, cap).toBe('no-hypotheses')
    }
  })

  it('the sampler recovers the known colours of real pixels when their elements read as automatic', () => {
    // axis-negative draws an explicit black axis and red negative fill: relabelled auto in memory, the sampler must
    // return exactly what the PNG shows there.
    const { c, parts, geom, png } = load('stage3-bar-axis-negative')
    const auto = withSheet(parts, autoAll)
    const measured = measureCaseBars(c, png, geom, auto) as ReadonlyMap<string, CfObservedBar>
    const obs = sampleBarAutoColours(c, auto, png, geom, measured)
    const axis = judgeBarAutoColour(obs['bar:color-auto-axis'])
    expect(axis).toMatchObject({ status: 'discriminating', matched: 'FF000000' })
    expect(obs['bar:color-auto-axis']!.samples.map(s => s.ref)).toContain('B4')
    const neg = judgeBarAutoColour(obs['bar:color-auto-negative-fill'])
    expect(neg).toMatchObject({ status: 'discriminating', matched: 'FFFF0000' })
    expect(obs['bar:color-auto-negative-fill']!.samples.map(s => s.ref)).toEqual(expect.arrayContaining(['B2', 'B3']))
    // Only auto="1" elements are sampled: the saved workbook as Excel wrote it carries none, so nothing is sampled.
    const explicit = sampleBarAutoColours(c, parts, png, geom, measureCaseBars(c, png, geom, parts) as ReadonlyMap<string, CfObservedBar>)
    expect(explicit['bar:color-auto-axis']).toBeUndefined()
    expect(explicit['bar:color-auto-negative-fill']).toBeUndefined()
  })

  it('a non-black automatic axis is located in its own colour and never read as bar ink', () => {
    const { c, parts, geom, png } = load('stage3-bar-axis-negative')
    const original = measureCaseBars(c, png, geom, parts) as ReadonlyMap<string, CfObservedBar>
    // Recolour Excel's dotted black axis to grey wherever the measurer found it.
    const grey = clonePng(png)
    for (const o of original.values()) {
      if (o.axisPx === undefined) continue
      for (let y = 0; y < grey.height; y++) {
        const i = (y * grey.width + o.axisPx) * 4
        if (grey.data[i] === 0 && grey.data[i + 1] === 0 && grey.data[i + 2] === 0) grey.data.set([0x80, 0x80, 0x80], i)
      }
    }
    const auto = withSheet(parts, autoAll)
    const remeasured = measureCaseBars(c, grey, geom, auto) as ReadonlyMap<string, CfObservedBar>
    expect(remeasured).toBeInstanceOf(Map)
    for (const [ref, o] of original) expect(remeasured.get(ref), ref).toEqual(o)
    expect(judgeBarAutoColour(sampleBarAutoColours(c, auto, grey, geom, remeasured)['bar:color-auto-axis'])).toMatchObject({ status: 'discriminating', matched: 'FF808080' })
    // Assuming black instead reads the grey axis as ink: the measurement is not the same.
    const assumed = measureCaseBars(c, grey, geom)
    expect(assumed instanceof Map && [...original].every(([ref, o]) => JSON.stringify(assumed.get(ref)) === JSON.stringify(o))).toBe(false)
  })

  it('an automatic colour that is not one colour refuses, and one too close to another colour of the bar does not discriminate', () => {
    const { c, parts, geom, png } = load('stage3-bar-axis-negative')
    const auto = withSheet(parts, autoAll)
    const measured = measureCaseBars(c, png, geom, auto) as ReadonlyMap<string, CfObservedBar>
    const obs = sampleBarAutoColours(c, auto, png, geom, measured)['bar:color-auto-axis']!
    const split = { ...obs, samples: obs.samples.map((s, i) => (i === 0 ? { ...s, argb: 'FF404040' } : s)) }
    expect(judgeBarAutoColour(split)).toMatchObject({ status: 'none', reason: expect.stringMatching(/not one colour/) })
    // 16 per channel from a bar colour is inside the separation (17): no discrimination.
    expect(judgeBarAutoColour(obs, ['FF101010'])).toMatchObject({ status: 'non-discriminating' })
    expect(judgeBarAutoColour(obs, ['FF111111'])).toMatchObject({ status: 'discriminating', matched: 'FF000000' })
    expect(judgeBarAutoColour({ ...obs, others: undefined })).toMatchObject({ status: 'non-discriminating' })
    expect(judgeBarAutoColour(undefined)).toMatchObject({ status: 'no-hypotheses' })
  })

  it('the audit blocks both capabilities and exports no colour until a verified capture lands', () => {
    const report = auditStage3Manifest(manifest(), { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES)) {
      expect(report.perCapability.find(p => p.capabilityId === cap)?.disposition, cap).toBe('supported-structural')
      expect(report.blockedCapabilities).toContain(cap)
    }
    expect(report.autoBarColours).toEqual({})
  })

  it('a colour is exported only for a verified capability whose captures agree', () => {
    const verified = [{ capabilityId: 'bar:color-auto-axis', caseId: 'x', disposition: 'verified' as const }, { capabilityId: 'bar:color-auto-negative-fill', caseId: 'x', disposition: 'supported-structural' as const }]
    const agree = { 'bar:color-auto-axis': { samples: [{ ref: 'x:B4', argb: 'FF000000' }, { ref: 'y:C4', argb: 'FF020202' }], unsampled: [], others: ['FF0000FF'] },
      'bar:color-auto-negative-fill': { samples: [{ ref: 'x:B2', argb: 'FFFF0000' }], unsampled: [], others: ['FF0000FF'] } }
    expect(verifiedAutoBarColours(verified, agree)).toEqual({ 'bar:color-auto-axis': 'FF020202' })
    // Captures that disagree across cases never export, even when the owning case verified.
    const disagree = { ...agree, 'bar:color-auto-axis': { ...agree['bar:color-auto-axis'], samples: [{ ref: 'x:B4', argb: 'FF000000' }, { ref: 'y:C4', argb: 'FF808080' }, { ref: 'z:D4', argb: 'FF808080' }] } }
    expect(verifiedAutoBarColours(verified, disagree)).toEqual({})
    expect(renderAutoBarColoursBlock({ 'bar:color-auto-negative-fill': 'FFFF0000', 'bar:color-auto-axis': 'FF000000' })).toBe(
      "export const VERIFIED_STAGE3_AUTO_BAR_COLORS: Readonly<Record<string, string>> = {\n  'bar:color-auto-axis': 'FF000000',\n  'bar:color-auto-negative-fill': 'FFFF0000',\n}\n",
    )
  })

  it('rendering without --publish-stage3 is byte-identical to the committed file; publishing carries the colour export', () => {
    const committed = readFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
    expect(renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: committed })).toBe(committed)
    expect(committed).not.toContain('VERIFIED_STAGE3_AUTO_BAR_COLORS')
    const published = renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { publishStage3: true })
    expect(published).toContain('export const VERIFIED_STAGE3_AUTO_BAR_COLORS: Readonly<Record<string, string>> = {\n}\n')
    // A published colour block is carried over verbatim by a later --write without --publish-stage3.
    const withColour = published.replace('VERIFIED_STAGE3_AUTO_BAR_COLORS: Readonly<Record<string, string>> = {\n}', "VERIFIED_STAGE3_AUTO_BAR_COLORS: Readonly<Record<string, string>> = {\n  'bar:color-auto-axis': 'FF000000',\n}")
    expect(renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: withColour })).toBe(withColour)
  })
})
