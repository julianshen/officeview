// @vitest-environment node
/**
 * Final review I5 (Ruling F-2): capabilities whose committed fixture cannot discriminate their question are
 * UNVERIFIED, and the src gate for each really withholds that shape. Every shape below paints when its
 * capability is allowed and is withheld under the production evidence (the verified set).
 */
import { afterEach, describe, expect, it } from 'vitest'
import { parseXml } from '../src/core/xml'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { parseCfSheet, MEASURED_CF_OFFICE_THEME, type CfStylesSource } from '../src/xlsx/cf/parse'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { VERIFIED_CAPABILITIES } from '../src/xlsx/cf/verifiedCapabilities'
import { injectExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { XlsxSheet } from '../src/xlsx/types'

afterEach(() => injectExactCfEvidence(null))

const NON_DISCRIMINATING = [
  'scale:colors-custom-theme', 'scale:scope-disjoint', 'scale:scope-hidden-row', 'scale:scope-autofilter-hidden-row', 'scale:scope-merged-cell',
] as const

const ALL = { allows: () => true }
const PRODUCTION = undefined // default: cfRuntimeEvidence (VERIFIED_CAPABILITIES)
const pt = (type: 'min' | 'max', argb: number) => ({ bound: { type }, argb })
const POINTS = [pt('min', 0xfff8696b), pt('max', 0xff63be7b)]

function sheetOf(opts: { ranges: Array<[number, number, number, number]>; values: Record<string, number | string>; hiddenRows?: number[]; merges?: Array<[number, number, number, number]> }): XlsxSheet {
  const byRow = new Map<number, Array<Record<string, unknown>>>()
  for (const [ref, value] of Object.entries(opts.values)) {
    const col = ref.charCodeAt(0) - 65, row = Number(ref.slice(1)) - 1
    const cell: Record<string, unknown> = { row, col, ref, styleIndex: 0, value }
    if (typeof value === 'string' && value.startsWith('#')) cell.valueIsError = true
    byRow.set(row, [...(byRow.get(row) ?? []), cell])
  }
  return {
    name: 'G', sheetId: 'g', cols: [], merges: [],
    mergeRanges: (opts.merges ?? []).map(([minRow, maxRow, minCol, maxCol]) => ({ minRow, maxRow, minCol, maxCol })),
    rows: [...byRow.entries()].sort((a, b) => a[0] - b[0]).map(([index, cells]) => ({ index, cells, ...(opts.hiddenRows?.includes(index) ? { hidden: true } : {}) })),
    conditionalFormatting: { diagnostics: [], rules: [{
      id: 1, priority: 1, documentOrder: 0, stopIfTrue: false,
      ranges: opts.ranges.map(([firstRow, lastRow, firstCol, lastCol]) => ({ firstRow, lastRow, firstCol, lastCol })),
      writes: new Set(['colorScale', 'fill']), definition: { kind: 'colorScale', points: POINTS, dialect: 'base', formulas: [] },
    }] },
  } as unknown as XlsxSheet
}

async function fillAt(sheet: XlsxSheet, evidence: { allows: (cap: string) => boolean } | undefined, ref: [number, number]): Promise<string | undefined> {
  const batch = await prepareCfPaint(sheet, { memory: createCfMemory(), ...(evidence ? { evidence } : {}) })
  const fill = batch?.layer.styleAt(ref[0], ref[1], undefined)?.fillColor
  batch?.dispose()
  return fill
}

describe('I5 the non-discriminating capabilities are unverified in production', () => {
  for (const id of NON_DISCRIMINATING) it(`${id} is not in VERIFIED_CAPABILITIES`, () => {
    expect(VERIFIED_CAPABILITIES.has(id)).toBe(false)
  })
  it('capabilities that the fixtures DO discriminate stay verified', () => {
    for (const id of ['scale:colors-theme-tint', 'scale:colors-indexed', 'scale:scope-duplicate-blocks', 'scale:scope-whole-column', 'scale:participation-error-presence', 'scale:interpolation-color-space']) {
      expect(VERIFIED_CAPABILITIES.has(id), id).toBe(true)
    }
  })
})

describe('I5 src gates withhold the unmeasured scope shapes', () => {
  it('scale:scope-disjoint: a scale over two disjoint rectangles', async () => {
    const sheet = () => sheetOf({ ranges: [[0, 2, 0, 0], [0, 2, 2, 2]], values: { A1: 1, A2: 2, A3: 3, C1: 10, C2: 20, C3: 30 } })
    expect(await fillAt(sheet(), ALL, [0, 0])).toBe('FFF8696B')
    expect(await fillAt(sheet(), PRODUCTION, [0, 0])).toBeUndefined()
  })

  it('scale:scope-hidden-row (and autofilter, indistinguishable): a hidden row inside the scope', async () => {
    const sheet = () => sheetOf({ ranges: [[0, 3, 0, 0]], values: { A1: 1, A2: 99, A3: 3, A4: 5 }, hiddenRows: [1] })
    expect(await fillAt(sheet(), ALL, [0, 0])).toBe('FFF8696B')
    expect(await fillAt(sheet(), PRODUCTION, [0, 0])).toBeUndefined()
    // either id alone is not enough: an autofilter writes ordinary hidden-row records
    expect(await fillAt(sheet(), { allows: c => c !== 'scale:scope-autofilter-hidden-row' }, [0, 0])).toBeUndefined()
    expect(await fillAt(sheet(), { allows: c => c !== 'scale:scope-hidden-row' }, [0, 0])).toBeUndefined()
  })

  it('scale:scope-merged-cell: a merge intersecting the scope', async () => {
    const sheet = () => sheetOf({ ranges: [[0, 3, 0, 1]], values: { A1: 1, A2: 2, A3: 3, B3: 4 }, merges: [[0, 1, 0, 1]] })
    expect(await fillAt(sheet(), ALL, [2, 0])).toBeDefined()
    expect(await fillAt(sheet(), PRODUCTION, [2, 0])).toBeUndefined()
  })

  it('scale:participation-error-presence is enforced: an error in the scope needs its capability', async () => {
    const sheet = () => sheetOf({ ranges: [[0, 3, 0, 0]], values: { A1: 1, A2: '#DIV/0!', A3: 3, A4: 5 } })
    expect(await fillAt(sheet(), ALL, [0, 0])).toBe('FFF8696B')
    expect(await fillAt(sheet(), { allows: c => c !== 'scale:participation-error-presence' }, [0, 0])).toBeUndefined()
    expect(await fillAt(sheet(), PRODUCTION, [0, 0])).toBe('FFF8696B')
  })

  it('a plain contiguous scope with no hidden rows, merges or errors still paints under production evidence', async () => {
    const sheet = sheetOf({ ranges: [[0, 3, 0, 0]], values: { A1: 1, A2: 2, A3: 3, A4: 5 } })
    expect(await fillAt(sheet, PRODUCTION, [0, 0])).toBe('FFF8696B')
  })
})

describe('I5 scale:colors-custom-theme gate (parse)', () => {
  const xml = (colors: string) => `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A1:A9"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/>${colors}</colorScale></cfRule></conditionalFormatting></worksheet>`
  const parse = (colors: string, theme: Record<string, string> | undefined, allowed: (cap: string) => boolean) => {
    injectExactCfEvidence(['scale:thresholds-min-max', 'scale:colors-theme-tint', 'scale:colors-indexed', 'scale:colors-alpha-00'].filter(allowed))
    const styles: CfStylesSource = { colors: { ...(theme ? { theme } : {}) }, dxfNodes: [] }
    return parseCfSheet(parseXml(xml(colors)), styles, createCfMemory()).value.rules[0]!
  }
  const THEME_COLORS = `<color theme="4" tint="0.4"/><color theme="5" tint="-0.4"/>`
  const custom = { ...MEASURED_CF_OFFICE_THEME, accent1: '112233' }

  it('the measured Office palette resolves theme colours with theme-tint evidence alone', () => {
    expect(parse(THEME_COLORS, MEASURED_CF_OFFICE_THEME, () => true).definition.kind).toBe('colorScale')
  })
  it('a non-Office palette withholds theme-coloured scales (no scale:colors-custom-theme evidence exists)', () => {
    expect(parse(THEME_COLORS, custom, () => true).definition.kind).toBe('unresolved')
  })
  it('a workbook with no theme part falls back to a palette that was never measured: withheld', () => {
    expect(parse(THEME_COLORS, undefined, () => true).definition.kind).toBe('unresolved')
  })
  it('explicit rgb colours are unaffected by the theme', () => {
    expect(parse(`<color rgb="FF336699"/><color rgb="FF99CC33"/>`, custom, () => true).definition.kind).toBe('colorScale')
  })
})
