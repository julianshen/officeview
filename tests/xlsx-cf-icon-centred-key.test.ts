// @vitest-environment node
/**
 * The centred glyph key (owner decision 2026-11-19): Stage 5 icon cells are resaved by Excel as Normal, vertical=center,
 * so their icons draw at cell top + 4 px, while the Stage 4 glyph key is bottom-placed (+7). A second key,
 * `stage4-icon-glyph-key-centred`, is captured at the Stage 5 placement, and each Stage 5 crop is classified against the
 * key whose placement its saved cell style has (a rule fixed before any capture). Classification stays position-exact:
 * no shifted search, no predicted offset.
 *
 * These tests pin the generated input against the Stage 5 icon fixtures and the key-selection rule. They are not
 * evidence of what Excel draws; that comes from the capture.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import { describe, expect, it, vi } from 'vitest'
import { CF_CASES, CF_STAGE4_CAPABILITIES, CF_STAGE4_CASES, CF_STAGE5_CASES, CF_ICON_KEY_CASES, CF_STAGE5_UNSUPPORTED, type CfOracleCase } from '../scripts/cf-refs/cases'
import { generateCaseArchive } from '../scripts/cf-refs/generate'
import { CF_STAGE5_UNCONFIRMABLE } from '../scripts/cf-refs/audit'
import { CF_FORMULA_PRODUCTION_READINGS } from '../src/xlsx/cf/formulaReadings'
import { verifyOnlyCaseOrder } from '../scripts/cf-refs/capture'
import {
  ICON_CROP_WIDTH_PX, type CfObservedIcon, type DecodedPng, type IconGlyphKey, computeWorksheetRangeGeometry, decodePngSync, evaluateFormulaCase, formulaObservationOf, iconKeyPlacement,
  loadIconGlyphKey, parseA1, readZipPartsSync, selectStage5IconKey, stage5IconKeyOf, stage5IconKeys, verifyCfCase,
} from '../scripts/cf-refs/verify'

const ROOT = resolve(__dirname, '..')
const CENTRED = 'stage4-icon-glyph-key-centred'
const KEY_COLS = ['B', 'C', 'D', 'E', 'F']
const dec = new TextDecoder()

/** Everything about one cell that could affect where Excel draws its icon, read from a workbook's parts. */
function placementFacts(parts: Map<string, Uint8Array>, ref: string) {
  const sheet = dec.decode(parts.get('xl/worksheets/sheet1.xml')!)
  const styles = dec.decode(parts.get('xl/styles.xml')!)
  const cell = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`).exec(sheet)
  const s = cell ? /\ss="(\d+)"/.exec(cell[1]!)?.[1] : undefined
  const t = cell ? /\st="(\w+)"/.exec(cell[1]!)?.[1] : undefined
  const v = cell?.[2] ? /<v>([^<]*)<\/v>/.exec(cell[2])?.[1] : undefined
  const xfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)![1]!.match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)!
  const fonts = /<fonts[^>]*>([\s\S]*?)<\/fonts>/.exec(styles)![1]!.match(/<font>[\s\S]*?<\/font>/g)!
  const xf = xfs[Number(s ?? '0')]!
  const fontId = Number(/fontId="(\d+)"/.exec(xf)![1])
  const row = Number(/\d+/.exec(ref)![0])
  const rowTag = new RegExp(`<row r="${row}"[^>]*>`).exec(sheet)?.[0] ?? ''
  const ht = /\sht="([\d.]+)"/.exec(rowTag)?.[1]
  const col = parseA1(ref)!.col + 1
  const cols = [...sheet.matchAll(/<col min="(\d+)" max="(\d+)" width="([\d.]+)"/g)]
  const width = cols.find(m => Number(m[1]) <= col && col <= Number(m[2]))?.[3]
  return { s, t, v, xf, numFmtId: /numFmtId="(\d+)"/.exec(xf)![1], font: fonts[fontId], ht, width }
}

const archive = async (id: string) => readZipPartsSync(await generateCaseArchive(id))!
const saved = (id: string) => readZipPartsSync(new Uint8Array(readFileSync(resolve(ROOT, CF_CASES.find(c => c.id === id)!.saved))))!

// Each admission runs a full verifyCfCase of one or both key cases: 6-11 s per test on the CI runner (under 1 s locally).
vi.setConfig({ testTimeout: 120_000 })

describe('the centred glyph key case', () => {
  it('is a Stage 4 key case with the same 75 single-pair rules as the Stage 4 key, appended after every other Stage 4 case', () => {
    const c = CF_STAGE4_CASES.find(x => x.id === CENTRED)
    expect(c).toBeDefined()
    expect(c!.stage).toBe(4)
    expect(c!.status).toBe('recipe-pending')
    expect(CF_STAGE4_CASES[CF_STAGE4_CASES.length - 1]!.id).toBe(CENTRED)
    const key = CF_STAGE4_CASES.find(x => x.id === 'stage4-icon-glyph-key')!
    const glyphs = (cc: typeof c) => cc!.markers.filter(m => m.capabilityId.startsWith('icon:glyph-')).map(m => ({ ...m, part: undefined }))
    expect(glyphs(c)).toHaveLength(75)
    expect(glyphs(c)).toEqual(glyphs(key))
    expect(c!.iconCalibration).toEqual(key.iconCalibration)
    expect(CF_ICON_KEY_CASES).toEqual({ bottom: 'stage4-icon-glyph-key', centre: CENTRED })
  })

  it('never owns a capability: every glyph and calibration id stays with the Stage 4 key (first in manifest order)', () => {
    const c = CF_STAGE4_CASES.find(x => x.id === CENTRED)!
    for (const m of c.markers) expect(CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === m.capabilityId)?.caseId, m.capabilityId).toBe('stage4-icon-glyph-key')
    expect(CF_STAGE4_CAPABILITIES.some(e => e.caseId === CENTRED)).toBe(false)
    const ids = CF_CASES.map(x => x.id)
    expect(ids.indexOf(CENTRED)).toBeGreaterThan(ids.indexOf('stage4-icon-glyph-key'))
  })

  it('verify-only re-derives both key sidecars before any dependent Stage 4 case', () => {
    const order = verifyOnlyCaseOrder(CF_STAGE4_CASES, 4).map(i => CF_STAGE4_CASES[i]!.id)
    expect(order.slice(0, 2)).toEqual(['stage4-icon-glyph-key', CENTRED])
  })
})

describe('the generated centred-key input matches the Stage 5 icon cells in everything that could affect placement', () => {
  // stage5-cfvo-scope P12..P15 and stage5-cfvo-formulas Z2..Z9 are icon twins on the default style (the re-deferred
  // icon:thresholds-formula and formula:cfvo-scope-multi-range twins).
  const STAGE5_ICON_CELLS: Array<[string, string]> = [['stage5-cfvo-scope', 'P12'], ['stage5-cfvo-formulas', 'Z2'], ['stage5-composition', 'K2']]

  it('key and control cells: default style (no s, xf 0), visible General numbers, 20 pt rows, 14-character columns, the Normal font', async () => {
    const parts = await archive(CENTRED)
    const c = CF_STAGE4_CASES.find(x => x.id === CENTRED)!
    const keyRefs = c.markers.filter(m => m.capabilityId.startsWith('icon:glyph-')).map(m => m.sqref)
    const refs = [...keyRefs, ...keyRefs.map(r => c.expected.cells.find(e => e.ref === r)!.controlRef)]
    expect(refs).toHaveLength(150)
    for (const ref of refs) {
      const f = placementFacts(parts, ref)
      expect({ ref, s: f.s, t: f.t, v: f.v, numFmtId: f.numFmtId, ht: f.ht, width: f.width }).toEqual({ ref, s: undefined, t: undefined, v: '10', numFmtId: '0', ht: '20', width: '14' })
    }
    for (const [id, ref] of STAGE5_ICON_CELLS) {
      const s5 = placementFacts(await archive(id), ref)
      const k = placementFacts(parts, 'B2')
      // The composition input carries the explicit bottom xf Excel folds into Normal on resave (§14.z): compare the
      // generated inputs only where Excel keeps them apart; the saved workbooks are compared below.
      if (id !== 'stage5-composition') expect({ s: k.s, xf: k.xf }, `${id} ${ref}`).toEqual({ s: s5.s, xf: s5.xf })
      expect({ font: k.font, ht: k.ht, width: k.width, numFmtId: k.numFmtId, t: k.t }, `${id} ${ref}`).toEqual({ font: s5.font, ht: s5.ht, width: s5.width, numFmtId: s5.numFmtId, t: s5.t })
    }
  })

  it('the Excel-saved Stage 5 icon cells are what the key reproduces: no s, xf 0 vertical=center, 20 pt, 14 characters', () => {
    for (const [id, ref] of STAGE5_ICON_CELLS) {
      const f = placementFacts(saved(id), ref)
      expect({ s: f.s, ht: f.ht, width: f.width, numFmtId: f.numFmtId }, `${id} ${ref}`).toEqual({ s: undefined, ht: '20', width: '14', numFmtId: '0' })
      expect(f.xf, `${id} ${ref}`).toMatch(/<alignment vertical="center"\/>/)
    }
  })

  it('the visible numbers stay clear of the left 28 px icon crop', async () => {
    const parts = await archive(CENTRED)
    const c = CF_STAGE4_CASES.find(x => x.id === CENTRED)!
    const geom = computeWorksheetRangeGeometry(parts, c.range)!
    for (const col of KEY_COLS) {
      const widthPx = geom.colWidthsPt[parseA1(`${col}2`)!.col - geom.minCol]!
      // "10" right-aligned in 11 pt Calibri: at most 2 x 7 px of digits and a 3 px right margin.
      expect(widthPx - 3 - 2 * 7, col).toBeGreaterThan(ICON_CROP_WIDTH_PX)
    }
  })
})

/** The saved parts with cell `ref`'s style or row height replaced (a synthetic resave, for the selection rule). */
function withXf(parts: Map<string, Uint8Array>, xf: string): Map<string, Uint8Array> {
  const out = new Map(parts)
  const styles = dec.decode(parts.get('xl/styles.xml')!)
  out.set('xl/styles.xml', new TextEncoder().encode(styles.replace(/(<cellXfs[^>]*>)(<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>))/, `$1${xf}`)))
  return out
}
function withRowHt(parts: Map<string, Uint8Array>, row: number, ht: number): Map<string, Uint8Array> {
  const out = new Map(parts)
  const sheet = dec.decode(parts.get('xl/worksheets/sheet1.xml')!)
  out.set('xl/worksheets/sheet1.xml', new TextEncoder().encode(sheet.replace(new RegExp(`(<row r="${row}"[^>]*?) ht="[\\d.]+"`), `$1 ht="${ht}"`)))
  return out
}

describe('key selection: the vertical alignment and row height of the cell\'s style in the Excel-saved workbook', () => {
  it('reads centre at 20 pt on the saved Stage 5 icon cells and bottom at 20 pt on the saved Stage 4 key cells', () => {
    expect(iconKeyPlacement(saved('stage5-cfvo-scope'), 'P12')).toBe('center-20')
    expect(iconKeyPlacement(saved('stage5-cfvo-formulas'), 'Z2')).toBe('center-20')
    expect(iconKeyPlacement(saved('stage5-composition'), 'K2')).toBe('center-20')
    // xf 1 (number format ;;;) has no <alignment>: bottom, as Excel drew it (+7).
    expect(iconKeyPlacement(saved('stage4-icon-glyph-key'), 'B2')).toBe('bottom-20')
  })

  it('anything else has no key: top or justified alignment, or a row other than 20 pt', () => {
    const s5 = saved('stage5-cfvo-scope')
    expect(iconKeyPlacement(withXf(s5, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="top"/></xf>'), 'P12')).toBeUndefined()
    expect(iconKeyPlacement(withXf(s5, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="justify"/></xf>'), 'P12')).toBeUndefined()
    expect(iconKeyPlacement(withXf(s5, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="bottom"/></xf>'), 'P12')).toBe('bottom-20')
    expect(iconKeyPlacement(withRowHt(s5, 12, 15), 'P12')).toBeUndefined()
    expect(iconKeyPlacement(withRowHt(s5, 12, 20.25), 'P12')).toBeUndefined()
    expect(iconKeyPlacement(withRowHt(s5, 13, 15), 'P12')).toBe('center-20')
    // A cell the sheet does not hold, or a workbook without styles, has no key.
    const noStyles = new Map(s5)
    noStyles.delete('xl/styles.xml')
    expect(iconKeyPlacement(noStyles, 'P12')).toBeUndefined()
  })

  it('picks the centred key for centre-20, the Stage 4 key for bottom-20, and nothing otherwise, never falling back', () => {
    const bottom = { classes: new Map() } as unknown as IconGlyphKey
    const centre = { classes: new Map() } as unknown as IconGlyphKey
    expect(selectStage5IconKey({ bottom, centre }, 'center-20')).toBe(centre)
    expect(selectStage5IconKey({ bottom, centre }, 'bottom-20')).toBe(bottom)
    expect(selectStage5IconKey({ bottom, centre }, undefined)).toBeUndefined()
    expect(selectStage5IconKey({ bottom }, 'center-20')).toBeUndefined()
    expect(selectStage5IconKey({ centre }, 'bottom-20')).toBeUndefined()
    const of = stage5IconKeyOf({ bottom, centre }, saved('stage5-cfvo-scope'))
    expect(of('P12')).toBe(centre)
    expect(stage5IconKeyOf({ bottom }, saved('stage5-cfvo-scope'))('P12')).toBeUndefined()
    // Without the saved parts no placement is known: no key.
    expect(stage5IconKeyOf({ bottom, centre }, undefined)('P12')).toBeUndefined()
  })

  it('the committed captures restore both keys, each admitted at its own placement', () => {
    const keys = stage5IconKeys(ROOT)
    expect(keys.bottom?.classes.size).toBe(54)
    expect(keys.centre?.classes.size).toBe(54)
    expect(keys.centre).not.toBe(keys.bottom)
    const centredSaved = saved(CENTRED)
    const c = CF_STAGE4_CASES.find(x => x.id === CENTRED)!
    for (const m of c.markers.filter(x => x.capabilityId.startsWith('icon:glyph-'))) expect(iconKeyPlacement(centredSaved, m.sqref), m.sqref).toBe('center-20')
    // A second call is served from the cache (nothing on disk changed).
    expect(stage5IconKeys(ROOT)).toBe(keys)
  })
})

describe('the centred key self-check (Excel 16.115 capture, 6d13e093)', () => {
  const manifest = () => JSON.parse(readFileSync(resolve(ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as { cases: CfOracleCase[] }
  const sidecar = (id: string) => JSON.parse(readFileSync(resolve(ROOT, `tests/fixtures/xlsx-cf/refs/${id}.json`), 'utf8')) as {
    glyphClasses: Record<string, string>; glyphTwins: Record<string, string[]>; glyphFaint: string[]; observedIcons: CfObservedIcon[]
  }

  it('verifies as a key: calibration, icon:glyph-key and all 75 pairs, classified into distinct classes', () => {
    const c = manifest().cases.find(x => x.id === CENTRED)!
    expect(c.status).toBe('measured')
    const res = verifyCfCase(c, { rootDir: ROOT })
    expect(res.kind).toBe('verified')
    const caps = res.kind === 'verified' ? res.verifiedCapabilities : []
    expect(caps).toContain('icon:pixel-calibration')
    expect(caps).toContain('icon:glyph-key')
    expect(caps.filter(x => x.startsWith('icon:glyph-') && x !== 'icon:glyph-key')).toHaveLength(75)
    expect(Object.keys(sidecar(CENTRED).glyphClasses)).toHaveLength(75)
  })

  it('partitions the 75 pairs exactly as the Stage 4 key does: the same class names, twins and faint class', () => {
    const a = sidecar(CENTRED), b = sidecar('stage4-icon-glyph-key')
    expect(a.glyphClasses).toEqual(b.glyphClasses)
    expect(a.glyphTwins).toEqual(b.glyphTwins)
    expect(a.glyphFaint).toEqual(b.glyphFaint)
  })

  it('every key glyph sits exactly 3 px higher than in the Stage 4 key (top + 4 at its frame), its left edge unchanged', () => {
    const a = sidecar(CENTRED), b = sidecar('stage4-icon-glyph-key')
    const keyRefs = CF_STAGE4_CASES.find(x => x.id === CENTRED)!.markers.filter(m => m.capabilityId.startsWith('icon:glyph-')).map(m => m.sqref)
    for (const ref of keyRefs) {
      const oa = a.observedIcons.find(o => o.ref === ref)!, ob = b.observedIcons.find(o => o.ref === ref)!
      // The faint 5Boxes:0 (B21) holds no measurable frame in either key (Ruling S4-FAINT); every other cell does.
      expect(oa.icon === null, ref).toBe(ob.icon === null)
      if (!oa.icon || !ob.icon) { expect(ref).toBe('B21'); continue }
      expect(oa.icon!.topPx - oa.cellTopPx, ref).toBe(ob.icon!.topPx - ob.cellTopPx - 3)
      expect(oa.icon!.leftPx - oa.cellLeftPx, ref).toBe(ob.icon!.leftPx - ob.cellLeftPx)
    }
  })
})

describe('a missing centred key fails closed (unmeasurable, never a Stage 4 match)', () => {
  const key = loadIconGlyphKey(ROOT) as IconGlyphKey
  const crop = (b: number): DecodedPng => key.classes.get(key.classOfPair.get(`3Arrows:${b}`)!)!
  /** A synthetic capture of the case's range on the given parts, with the Stage 4 key's crop pasted at P12 exactly. */
  const paint = (parts: Map<string, Uint8Array>) => {
    const c = CF_STAGE5_CASES.find(x => x.id === 'stage5-cfvo-scope')!
    const geom = computeWorksheetRangeGeometry(parts, c.range)!
    const canvas = createCanvas(geom.expectedWidthPt, geom.expectedHeightPt)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, geom.expectedWidthPt, geom.expectedHeightPt)
    const p = parseA1('P12')!
    const x0 = Math.round(geom.colWidthsPt.slice(0, p.col - geom.minCol).reduce((a, b) => a + b, 0))
    const y0 = Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow).reduce((a, b) => a + b, 0))
    const k = crop(2)
    const img = ctx.createImageData(k.width, k.height)
    img.data.set(k.data)
    ctx.putImageData(img, x0 + 1, y0 + 1)
    return { c, geom, png: decodePngSync(new Uint8Array(canvas.toBuffer('image/png')))! }
  }

  it('a centred saved cell showing the Stage 4 key crop exactly is unmeasurable without the centred key', () => {
    const parts = saved('stage5-cfvo-scope')
    const { c, geom, png } = paint(parts)
    const obs = formulaObservationOf(c, {}, png, geom, stage5IconKeyOf({ bottom: key }, parts))
    expect(obs.visual!('P12', 'icon')).toBeUndefined()
  })

  it('the same pixels classify on a bottom-placed cell (the mechanism the centred key will use)', () => {
    const bottomParts = withXf(saved('stage5-cfvo-scope'), '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="bottom"/></xf>')
    const { c, geom, png } = paint(bottomParts)
    const obs = formulaObservationOf(c, {}, png, geom, stage5IconKeyOf({ bottom: key }, bottomParts))
    expect(obs.visual!('P12', 'icon')).toMatchObject({ kind: 'icon', cls: key.classOfPair.get('3Arrows:2') })
    // And as the centred key: a centre-20 cell classifies against whatever key is installed for that placement.
    const centreParts = saved('stage5-cfvo-scope')
    const centred = paint(centreParts)
    expect(formulaObservationOf(centred.c, {}, centred.png, centred.geom, stage5IconKeyOf({ centre: key }, centreParts)).visual!('P12', 'icon')).toMatchObject({ kind: 'icon' })
  })

  it('evaluateFormulaCase accepts the key pair and selects per cell from the saved parts', () => {
    const parts = saved('stage5-cfvo-scope')
    const { c, geom, png } = paint(parts)
    const one = { ...c, formulaConfigurations: c.formulaConfigurations!.filter(x => x.id === 'scope-icon') }
    const withBottomOnly = evaluateFormulaCase(one, {}, png, geom, parts, { bottom: key })
    const withSingleKey = evaluateFormulaCase(one, {}, png, geom, parts, key)
    expect(withSingleKey).toEqual(withBottomOnly)
  })
})


describe('Stage 5 re-evaluated with the centred key (the four Task 39 re-deferrals)', () => {
  const evaluation = (id: string) => (JSON.parse(readFileSync(resolve(ROOT, `tests/fixtures/xlsx-cf/refs/${id}.json`), 'utf8')) as { formulaEvaluation: Record<string, { status: string; matched?: string; code?: string; observed?: Record<string, Record<string, string>> }> }).formulaEvaluation

  it('the three icon readings Excel now shows classify at the centred frame and match production: the re-deferrals are lifted', () => {
    const f = evaluation('stage5-cfvo-formulas')['icon:thresholds-formula']!
    expect(f).toMatchObject({ status: 'discriminating', matched: 'evaluated' })
    expect(CF_FORMULA_PRODUCTION_READINGS['icon:thresholds-formula']).toBe('evaluated')
    const comp = evaluation('stage5-composition')
    expect(comp['formula:composition-stop-icon']).toMatchObject({ status: 'discriminating', matched: 'halts' })
    expect(CF_FORMULA_PRODUCTION_READINGS['formula:composition-stop-icon']).toBe('halts')
    expect(comp['formula:composition-indirect-ahead-icon']).toMatchObject({ status: 'discriminating', matched: 'halts' })
    // Every icon witness is a classified glyph at cell-relative (2, 4): the centred frame.
    for (const obs of [f.observed!['cfvo-icon-formula']!, comp['formula:composition-stop-icon']!.observed!['stop-icon']!]) {
      for (const v of Object.values(obs)) if (v.startsWith('icon:')) expect(v).toMatch(/^icon:3Arrows:[012]@2,4\|/)
    }
    for (const id of ['icon:thresholds-formula', 'formula:composition-stop-icon', 'formula:composition-indirect-ahead-icon']) expect(CF_STAGE5_UNSUPPORTED[id], id).toBeUndefined()
    // Lifting the measurement re-deferral does not lift the publication gates: stop-icon stays fixture-unconfirmable and
    // indirect-ahead-icon names no production reading.
    expect(CF_STAGE5_UNCONFIRMABLE['formula:composition-stop-icon']).toBeDefined()
    expect(CF_FORMULA_PRODUCTION_READINGS['formula:composition-indirect-ahead-icon']).toBeUndefined()
  })

  it('formula:cfvo-scope-multi-range now measures and matches neither reading: kept deferred as a finding', () => {
    const e = evaluation('stage5-cfvo-scope')['formula:cfvo-scope-multi-range']!
    expect(e).toMatchObject({ status: 'refused', code: 'formula-hypothesis-none' })
    // Excel drew the top icon (3Arrows:2) on all twelve designated cells and the top colour on every scale cell.
    expect(new Set(Object.values(e.observed!['scope-icon']!))).toEqual(new Set(['icon:3Arrows:2@2,4|none']))
    expect(CF_STAGE5_UNSUPPORTED['formula:cfvo-scope-multi-range']).toMatch(/no hypothesis/)
  })
})
