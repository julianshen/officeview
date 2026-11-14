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
import { autoBarColourEvidence, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import {
  BAR_AUTO_COLOUR_CAPABILITIES,
  computeWorksheetRangeGeometry,
  decodePngSync,
  judgeBarAutoColour,
  measureBarGeometry,
  measureCaseBars,
  readZipPartsSync,
  sampleBarAutoColours,
  type CfObservedBar,
  type DecodedPng,
} from '../scripts/cf-refs/verify'

const REPO_ROOT = resolve(__dirname, '..')
const manifest = (): CfManifest => JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
/** The manifest as it stood before run 6 captured the grey-cell case: that case recipe-pending, its capture spec not
 * yet taken (no recorded hashes). */
const precapture = (): CfManifest => {
  const m = manifest()
  return {
    ...m,
    cases: m.cases.map(c => (c.id === 'stage3-bar-auto-colors'
      ? { ...c, status: 'recipe-pending' as const, captures: c.captures.map(x => ({ appearance: x.appearance, pngPath: x.pngPath, sidecarPath: x.sidecarPath })) }
      : c)),
  }
}
/** These tests run whole-manifest audits (every Stage 3 PNG re-decoded and re-measured), several per test. */
const AUDIT_TIMEOUT_MS = 30_000
const WHITE = { 'bar:color-auto-axis': 'FFFFFFFF', 'bar:color-auto-negative-fill': 'FFFFFFFF' }

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

/**
 * The real axis-negative capture repainted the way run 5 shows Excel draws automatic colours: red negative bars and
 * the black axis become opaque white; with `grey`, neutral pixels (background and antialiased text, composited over
 * white) are re-composited over the FFA6A6A6 cell fill: v * A6 / FF.
 */
function paintExcelWhite(png: DecodedPng, original: ReadonlyMap<string, CfObservedBar>, grey: boolean): DecodedPng {
  const out = clonePng(png)
  const d = out.data
  const axisCols = new Set([...original.values()].map(o => o.axisPx).filter((x): x is number => x !== undefined))
  for (let i = 0; i < d.length; i += 4) {
    const x = (i / 4) % out.width
    const [r, g, b] = [d[i]!, d[i + 1]!, d[i + 2]!]
    if (r === 255 && g === 0 && b === 0) d.set([255, 255, 255], i)
    else if (r === 0 && g === 0 && b === 0 && axisCols.has(x)) d.set([255, 255, 255], i)
    else if (grey && r === g && g === b) d.set([0, 0, 0].map(() => Math.round((r * 0xa6) / 0xff)), i)
  }
  return out
}

const clonePng = (png: DecodedPng): DecodedPng => ({ ...png, data: new Uint8ClampedArray(png.data) })

describe('Stage 3 automatic bar colours (Ruling S3-F as amended, Task 19b)', () => {
  it('only the automatic axis and negative fill occur, and only the grey-cell capture draws them (white)', () => {
    expect(Object.keys(BAR_AUTO_COLOUR_CAPABILITIES).sort()).toEqual(['bar:color-auto-axis', 'bar:color-auto-negative-fill'])
    expect(CF_STAGE3_CAPABILITIES.filter(e => e.capabilityId.startsWith('bar:color-auto-')).map(e => e.caseId))
      .toEqual(['stage3-bar-auto-colors', 'stage3-bar-auto-colors'])
    // Run 6 (50292100): every sample comes from the grey-cell case, and all of them are white.
    const across = barAutoColoursAcrossCases(manifest(), REPO_ROOT)
    for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES)) {
      expect(across[cap]!.samples.length, cap).toBeGreaterThan(0)
      expect(across[cap]!.samples.every(x => x.ref.startsWith('stage3-bar-auto-colors:') && x.argb === 'FFFFFFFF'), cap).toBe(true)
      expect(judgeBarAutoColour(across[cap]), cap).toMatchObject({ status: 'discriminating', matched: 'FFFFFFFF' })
    }
    // Before that capture no kept capture drew either colour (every other resave is all-positive or explicit).
    const before = barAutoColoursAcrossCases(precapture(), REPO_ROOT)
    for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES)) {
      expect(before[cap]?.samples ?? [], cap).toEqual([])
      expect(judgeBarAutoColour(before[cap]).status, cap).toBe('no-hypotheses')
    }
  }, AUDIT_TIMEOUT_MS)

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

  it('on a grey static cell fill, Excel\'s white automatic colours are ink and sampled as FFFFFFFF (Ruling S3-AUTO-GREY)', () => {
    const { c, parts, geom, png } = load('stage3-bar-axis-negative')
    const original = measureCaseBars(c, png, geom, parts) as ReadonlyMap<string, CfObservedBar>
    const auto = withSheet(parts, autoAll)
    // Run 5: Excel paints the automatic negative fill and axis opaque white. Paint the real capture that way, first on
    // the grey cell fill the case now carries, then on white. The capture's neutral pixels (background and antialiased
    // black text, composited over white) are re-composited over the grey: v * A6 / FF.
    const paint = (grey: boolean): DecodedPng => paintExcelWhite(png, original, grey)
    const onGrey = paint(true)
    const measured = measureCaseBars(c, onGrey, geom, auto)
    expect(measured).toBeInstanceOf(Map)
    const bars = measured as ReadonlyMap<string, CfObservedBar>
    // The white negative bars are measured where Excel drew them (a white axis next to them may add its column).
    for (const ref of ['B2', 'B3', 'C2', 'C3', 'D3', 'E2', 'E3']) {
      const was = original.get(ref)!.bar!
      const now = bars.get(ref)!.bar!
      expect(Math.abs(now.leftPx - was.leftPx) + Math.abs(now.rightPx - was.rightPx), ref).toBeLessThanOrEqual(1)
    }
    // A white negative bar is never mistaken for an axis: the axis stays the one-column line of the zero cells.
    expect(bars.get('B4')).toMatchObject({ bar: null, axisPx: original.get('B4')!.axisPx })
    const obs = sampleBarAutoColours(c, auto, onGrey, geom, bars)
    expect(judgeBarAutoColour(obs['bar:color-auto-axis'])).toMatchObject({ status: 'discriminating', matched: 'FFFFFFFF' })
    expect(judgeBarAutoColour(obs['bar:color-auto-negative-fill'])).toMatchObject({ status: 'discriminating', matched: 'FFFFFFFF' })
    // The grey background counts as another colour of the bar: white clears it by 89, above the 17 separation.
    expect(obs['bar:color-auto-axis']!.others).toContain('FFA6A6A6')
    // On white cells the same white is no ink at all: nothing to sample, so the capabilities stay blocked (run 5).
    const onWhite = paint(false)
    const whiteBars = measureCaseBars(c, onWhite, geom, auto)
    const white = sampleBarAutoColours(c, auto, onWhite, geom, whiteBars instanceof Map ? whiteBars : undefined)
    expect(judgeBarAutoColour(white['bar:color-auto-axis']).status).not.toBe('discriminating')
    expect(judgeBarAutoColour(white['bar:color-auto-negative-fill']).status).not.toBe('discriminating')
  })

  it('a white bar beside a white axis measures from the evidenced axis position, and refuses without it', () => {
    const { c, parts, geom, png } = load('stage3-bar-axis-negative')
    const original = measureCaseBars(c, png, geom, parts) as ReadonlyMap<string, CfObservedBar>
    const auto = withSheet(parts, autoAll)
    const axis = original.get('B4')!.axisPx!
    const b2 = original.get('B2')!.bar!
    expect(b2.rightPx).toBe(axis - 1)
    for (const width of [1, 2]) {
      // Shorten B2's white negative bar to `width` px ending beside the white axis (the rest back to the grey fill).
      const narrow = paintExcelWhite(png, original, true)
      for (let y = b2.topPx; y <= b2.bottomPx; y++) for (let x = b2.leftPx; x <= axis - 1 - width; x++) narrow.data.set([0xa6, 0xa6, 0xa6], (y * narrow.width + x) * 4)
      // Evidence: B4's zero cell puts the axis at its column, so the run splits into axis + bar with exact edges.
      const measured = measureCaseBars(c, narrow, geom, auto)
      expect(measured, `${width} px`).toBeInstanceOf(Map)
      expect((measured as ReadonlyMap<string, CfObservedBar>).get('B2'), `${width} px`).toMatchObject({ bar: { leftPx: axis - width, rightPx: axis - 1 }, axisPx: axis })
      // No position evidence, an unknown position, or a position elsewhere: refuse, never a silently wrong edge.
      const refs = c.expected.cells.filter(x => x.matched).map(x => x.ref)
      const controlOf = (ref: string): string => c.expected.cells.find(x => x.ref === ref)!.controlRef
      for (const axisAt of [undefined, () => undefined, () => axis - 40]) {
        expect(measureBarGeometry(narrow, geom, refs, { controlOf, axisArgb: 'FFFFFFFF', tolerancePx: 1, ...(axisAt ? { axisAt } : {}) }), `${width} px`)
          // (A wrong position refuses at the first cell it contradicts, B2 or the zero cell B4.)
          .toMatchObject({ kind: 'refusal', code: 'bar-edge-unstable', reason: expect.stringMatching(/^Bar cell B[24]: /) })
      }
    }
    // A wide axis-coloured run is refused too when nothing places the axis.
    const refs = c.expected.cells.filter(x => x.matched).map(x => x.ref)
    const controlOf = (ref: string): string => c.expected.cells.find(x => x.ref === ref)!.controlRef
    expect(measureBarGeometry(paintExcelWhite(png, original, true), geom, refs, { controlOf, axisArgb: 'FFFFFFFF', tolerancePx: 1 }))
      .toMatchObject({ kind: 'refusal', code: 'bar-edge-unstable' })
  })

  it('the automatic-colours case puts a static grey fill under every bar and control cell', () => {
    const tpl = CF_STAGE3_CASES.find(x => x.id === 'stage3-bar-auto-colors')!
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, tpl.input))))!
    const styles = new TextDecoder().decode(parts.get('xl/styles.xml')!)
    const sheet = new TextDecoder().decode(parts.get('xl/worksheets/sheet1.xml')!)
    expect(styles).toContain('<fill><patternFill patternType="solid"><fgColor rgb="FFA6A6A6"/><bgColor indexed="64"/></patternFill></fill>')
    expect(styles).toContain('<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>')
    // The grey is not a CF DXF.
    expect(styles).not.toMatch(/<dxf>[^]*A6A6A6/)
    for (const cell of tpl.expected.cells) {
      for (const ref of [cell.ref, cell.controlRef]) expect(sheet, ref).toMatch(new RegExp(`<c r="${ref}"[^>]* s="4"`))
    }
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

  it('the audit verifies both colours as white from the grey-cell capture, and production resolves exactly them', () => {
    const report = auditStage3Manifest(manifest(), { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES)) {
      expect(report.perCapability.find(p => p.capabilityId === cap)?.disposition, cap).toBe('verified')
      expect(report.blockedCapabilities).not.toContain(cap)
    }
    expect(report.autoBarColours).toEqual(WHITE)
    // Task 23 published the audit's colours: production resolves each, only through its capability.
    for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES)) expect(autoBarColourEvidence(cap), cap).toBe(0xffffffff)
    withExactCfEvidence([], () => {
      for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES)) expect(autoBarColourEvidence(cap), cap).toBeUndefined()
    })
    // Without a verified capture the audit blocks both and exports no colour.
    const before = auditStage3Manifest(precapture(), { rootDir: REPO_ROOT })
    expect(before.ok).toBe(true)
    for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES)) {
      expect(before.perCapability.find(p => p.capabilityId === cap)?.disposition, cap).toBe('supported-structural')
      expect(before.blockedCapabilities).toContain(cap)
    }
    expect(before.autoBarColours).toEqual({})
  }, AUDIT_TIMEOUT_MS)

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
    // Publishing (Task 23) carries both measured colours, and the committed file is that publication.
    const published = renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { publishStage3: true })
    expect(published).toContain(renderAutoBarColoursBlock(WHITE))
    expect(committed).toBe(published)
    // Before the capture, publishing carries an empty colour block.
    expect(renderVerifiedCapabilitiesSource(precapture(), REPO_ROOT, { publishStage3: true }))
      .toContain('export const VERIFIED_STAGE3_AUTO_BAR_COLORS: Readonly<Record<string, string>> = {\n}\n')
    // A published colour block is carried over verbatim by a later --write without --publish-stage3.
    const withColour = published.replace(renderAutoBarColoursBlock(WHITE), renderAutoBarColoursBlock({ 'bar:color-auto-axis': 'FF000000' }))
    expect(withColour).not.toBe(published)
    expect(renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: withColour })).toBe(withColour)
  }, AUDIT_TIMEOUT_MS)
})
