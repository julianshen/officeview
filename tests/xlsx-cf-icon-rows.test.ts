// @vitest-environment node
/**
 * Task 28 round 2, Ruling S4-ROWS-CAPTURE: `stage4-icon-rows` measures the icon frame across row heights, so the frame
 * the glyph key fixes at 20 pt can be extended, at the measured heights only, to default (15 pt) and other rows.
 *
 * These are mechanism tests (as fix round 8's): synthetic captures, drawn as each offset rule predicts, with the
 * glyph-key crops pasted at the predicted place, prove that the case and the harness discriminate the rules. They are
 * not evidence of what Excel does; Excel's answer comes from the capture.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics } from '../src/xlsx/render'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { iconRecorder, withIcons } from './helpers/cf-icons'
import { AT_CAPTURE } from './helpers/cf-glyph-measure'
import { CF_STAGE4_CAPABILITIES, CF_STAGE4_CASES, iconCapabilityKind } from '../scripts/cf-refs/cases'
import {
  ICON_GEOMETRY_READINGS, type CfObservedIcon, type IconGeometryCell, type IconGlyphKey, computeWorksheetRangeGeometry, decodePngSync, evaluateIconCapabilities, evaluateIconGeometryCapability,
  loadIconGlyphKey, measureCaseIcons, parseA1, readZipPartsSync,
} from '../scripts/cf-refs/verify'

const ROOT = resolve(__dirname, '..')
const CASE = () => CF_STAGE4_CASES.find(c => c.id === 'stage4-icon-rows')!
const CAP = 'icon:geometry-row-offset'
const ROWS_PT = [15, 12.75, 16.5, 18, 20, 22.5, 25, 30]
/** Column → (glyph class at value 10 under num 0/5/8, vertical alignment, horizontal alignment). */
const COLUMNS: Record<string, { cls: string; vertical: 'bottom' | 'center'; horizontal: 'right' | 'left' }> = {
  B: { cls: '3Arrows:2', vertical: 'bottom', horizontal: 'right' },
  C: { cls: '3TrafficLights1:2', vertical: 'bottom', horizontal: 'right' },
  D: { cls: '3Symbols:2', vertical: 'bottom', horizontal: 'right' },
  E: { cls: '3Arrows:2', vertical: 'bottom', horizontal: 'left' },
  F: { cls: '3Arrows:2', vertical: 'center', horizontal: 'right' },
  G: { cls: '3TrafficLights1:2', vertical: 'center', horizontal: 'right' },
  H: { cls: '3Symbols:2', vertical: 'center', horizontal: 'right' },
  I: { cls: '3Arrows:2', vertical: 'center', horizontal: 'left' },
}

describe('S4-ROWS-CAPTURE: the stage4-icon-rows case', () => {
  it('owns icon:geometry-row-offset over 11 pt Calibri rows of 15, 12.75, 16.5, 18, 20, 22.5, 25 and 30 pt', () => {
    const c = CASE()
    expect(c).toBeDefined()
    expect(c.status).toBe('recipe-pending')
    // Round 4 (S4-BOTTOM-REPLICATE): stage4-icon-rows-bottom, placed ahead, now owns the capability; this case stays recorded.
    expect(CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === CAP)?.caseId).toBe('stage4-icon-rows-bottom')
    expect(iconCapabilityKind(CAP)).toBe('geometry')
    expect(ICON_GEOMETRY_READINGS[CAP]).toEqual(['anchor-constant', 'centred', 'row-scaled'])
    // Base 3Arrows, 3TrafficLights1 and 3Symbols rules, every value 10 in bucket 2.
    const markers = c.markers.filter(m => m.capabilityId === CAP)
    expect(markers.map(m => [m.sqref, m.icon?.set, m.icon?.dialect])).toEqual(Object.entries(COLUMNS).map(([col, v]) => [`${col}2:${col}9`, v.cls.split(':')[0], 'base']))
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(ROOT, c.input))))!
    const geom = computeWorksheetRangeGeometry(parts, c.range)!
    expect(geom.rowHeightsPt.slice(1, 9)).toEqual(ROWS_PT)
    expect(c.expected.cells.filter(x => x.matched)).toHaveLength(64)
  })
})

describe('S4-ROWS-CAPTURE: the case discriminates the offset rules (mechanism tests)', () => {
  const key = loadIconGlyphKey(ROOT) as IconGlyphKey
  /** The key frame's top at the key geometry: cell y 7 (`CF_ICON_RENDER_MODEL.keyFrame`), 12 px tall. */
  const KEY_TOP = 7, FRAME_H = 12
  /** A synthetic capture: each designated cell shows its class's key crop with the frame top at `frameTop`. */
  const capture = (frameTop: (vertical: 'bottom' | 'center', rowPx: number) => number) => {
    const c = CASE()
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(ROOT, c.input))))!
    const geom = computeWorksheetRangeGeometry(parts, c.range)!
    const at = (ref: string) => {
      const p = parseA1(ref)!
      const x0 = Math.round(geom.colWidthsPt.slice(0, p.col - geom.minCol).reduce((a, b) => a + b, 0))
      const y0 = Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow).reduce((a, b) => a + b, 0))
      const y1 = Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow + 1).reduce((a, b) => a + b, 0))
      return { x0, y0, h: y1 - y0 }
    }
    const W = geom.expectedWidthPt, H = geom.expectedHeightPt
    const canvas = createCanvas(Math.round(W) + 1, Math.round(H) + 1)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height)
    const img = ctx.getImageData(0, 0, canvas.width, canvas.height)
    for (const e of c.expected.cells.filter(x => x.matched)) {
      const col = /^[A-Z]+/.exec(e.ref)![0]
      const spec = COLUMNS[col]!
      const crop = key.classes.get(spec.cls)!
      const b = at(e.ref)
      const dy = frameTop(spec.vertical, b.h) - KEY_TOP
      // Crop pixel (x, y) is cell (x + 1, y + 1) at the key; here shifted by dy. Only the glyph's own rows move.
      for (let y = 0; y < crop.height; y++) for (let x = 0; x < crop.width; x++) {
        const k = (y * crop.width + x) * 4
        if (crop.data[k]! > 247 && crop.data[k + 1]! > 247 && crop.data[k + 2]! > 247) continue
        const X = b.x0 + x + 1, Y = b.y0 + y + 1 + dy
        if (Y <= b.y0 || Y >= b.y0 + b.h - 1) continue
        const o = (Y * canvas.width + X) * 4
        img.data[o] = crop.data[k]!; img.data[o + 1] = crop.data[k + 1]!; img.data[o + 2] = crop.data[k + 2]!; img.data[o + 3] = 255
      }
    }
    ctx.putImageData(img, 0, 0)
    const decoded = decodePngSync(new Uint8Array(canvas.toBuffer('image/png')))!
    const observed = measureCaseIcons(c, decoded, geom, key) as ReadonlyMap<string, CfObservedIcon>
    return { observed, result: evaluateIconCapabilities(c, parts, observed, key)[CAP] }
  }
  const rules: Record<string, (v: 'bottom' | 'center', rowPx: number) => number> = {
    // The key's frame bottom stays 1 px above the next gridline (bottom), or the frame stays centred (center).
    'anchor-constant': (v, h) => (v === 'bottom' ? h - 1 - FRAME_H : Math.round((h - FRAME_H) / 2)),
    // The frame centred in the row whatever the cell's alignment.
    centred: (_, h) => Math.round((h - FRAME_H) / 2),
    // The key's offsets scaled with the row from 20 pt. A shorter row cannot hold a frame scaled down from the key's
    // (the reading makes no claim there); those cells are drawn as far up as the frame goes, 1 px under the top line.
    'row-scaled': (_, h) => (h >= 20 ? Math.round((KEY_TOP * h) / 20) : 2),
  }

  it('every designated cell is measured against its class key crop (shifted classification), with stable edges', () => {
    const { observed } = capture(rules['anchor-constant']!)
    expect(observed.size).toBe(64)
    for (const [ref, o] of observed) {
      // The 12.75 pt row (13 px, an 11 px inner box) cannot hold the 12 px frame: no claim there (shrink, unmeasured).
      if (o.cellHeightPx - 2 < 12) continue
      const spec = COLUMNS[/^[A-Z]+/.exec(ref)![0]!]!
      expect(o.icon?.glyphClass, ref).toBe(spec.cls)
      expect(o.edgeStable, ref).not.toBe(false)
    }
  })

  for (const truth of Object.keys(rules)) {
    it(`a capture drawn as ${truth} discriminates ${truth}`, () => {
      const r = capture(rules[truth]!).result
      expect(r).toMatchObject({ status: 'discriminating', matched: truth })
    })
  }

  it('a capture that follows no rule (bottom gap 1 at 20 pt, 3 elsewhere) is refuted by every reading', () => {
    const odd = (v: 'bottom' | 'center', h: number) => (v === 'bottom' ? h - (h === 20 ? 1 : 5) - FRAME_H : Math.round((h - FRAME_H) / 2))
    expect(capture(odd).result.status).not.toBe('discriminating')
  })
})

/**
 * Task 28 round 4 (S4-ROW-EXACT): production on Excel's own resaved workbook (cd5ae569). Cells without a style index
 * read the resaved Normal style (11 pt Calibri, `vertical="center"`); the frame paints only on rows whose viewer px
 * equal Excel's (15, 18 and 30 pt), where it sits as the sidecar measured.
 */
describe('S4-ROW-EXACT: production on the resaved stage4-icon-rows workbook', () => {
  it('paints exactly the admitted rows, each frame where Excel inked it', async () => {
    const c = CASE()
    const sidecar = JSON.parse(readFileSync(resolve(ROOT, 'tests/fixtures/xlsx-cf/refs/stage4-icon-rows.json'), 'utf8')) as { observedIcons: Array<{ ref: string; cellTopPx: number; icon: null | { topPx: number } }> }
    // The parse gates (base dialect, schema defaults) and each pair's glyph evidence, over the injected icon stubs.
    const caps = ['icon:settings-schema-defaults', 'icon:dialect-base', 'icon:geometry-row-offset', ...['3Arrows', '3TrafficLights1', '3Symbols'].flatMap(set => [0, 1, 2].map(id => `icon:glyph-${set}-${id}`))]
    await withIcons(async () => {
      const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(ROOT, c.saved)))))
      const sheet = doc.sheets[0]!
      // Cells without a style index read the resaved Normal style; E keeps its own xf, with no vertical (bottom).
      const b2 = sheet.rows.find(r => r.index === 1)!.cells.find(x => x.col === 1)!
      expect(b2.style).toMatchObject({ fontSizePt: 11, fontName: 'Calibri', vertical: 'center' })
      const e2 = sheet.rows.find(r => r.index === 1)!.cells.find(x => x.col === 4)!
      expect(e2.style?.vertical).toBeUndefined()
      const metrics = computeMetrics(sheet)
      const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext: () => createCanvas(1, 1).getContext('2d') as never }))!
      const layout = createCellLayoutHost(createCanvas(1, 1).getContext('2d') as never)
      const top = (row: number) => metrics.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0)
      const left = (col: number) => metrics.colWidthsPx.slice(0, col).reduce((a, b) => a + b, 0)
      const hostAt = (row: number, col: number) => {
        const cell = sheet.rows.find(r => r.index === row)?.cells.find(x => x.col === col)
        if (!cell) return undefined
        const body = layout.layoutBody(cell, metrics.colWidthsPx[col]!, metrics.rowHeightsPx[row]!)
        const textWidth = body.text?.kind === 'plain' ? body.text.width : body.text?.kind === 'layout' ? body.text.layout.lines[0]!.segments.reduce((n, g) => n + g.width, 0) : undefined
        return { x: left(col), y: top(row), width: metrics.colWidthsPx[col]!, height: metrics.rowHeightsPx[row]!, cell, ...(textWidth !== undefined ? { textWidth } : {}) }
      }
      const ADMITTED = new Set([1, 4, 8]) // rows 2, 5, 9: 15, 18 and 30 pt
      for (let row = 1; row <= 8; row++) for (let col = 1; col <= 8; col++) {
        expect(Boolean(batch.iconAt(row, col)), `${String.fromCharCode(65 + col)}${row + 1}`).toBe(ADMITTED.has(row))
      }
      // Where the 3Arrows frames paint (B, E, F, I): the frame top (the first fill's top, the outline's apex) against the
      // cell's top gridline, at the capture scale, equals the sidecar's ink top (3Arrows:2 inks from the frame top).
      const rec = iconRecorder(800, 400, createCanvas as never)
      batch.paintIcons(rec.ctx, hostAt, { r0: 1, c0: 1, r1: 8, c1: 8 }, AT_CAPTURE)
      const roundHalfUp = (x: number) => Math.floor(x + 0.5)
      for (const row of ADMITTED) for (const col of [1, 4, 5, 8]) {
        const ref = `${String.fromCharCode(65 + col)}${row + 1}`
        const X0 = roundHalfUp(left(col) * 0.75), X1 = roundHalfUp((left(col) + metrics.colWidthsPx[col]!) * 0.75)
        const Y0 = roundHalfUp(top(row) * 0.75)
        const ops = rec.ops.filter(o => o.op === 'icon' && o.x >= X0 && o.x < X1 && o.y >= Y0 - 1 && o.y < Y0 + 40)
        expect(ops.length, ref).toBeGreaterThan(0)
        const o = sidecar.observedIcons.find(x => x.ref === ref)!
        expect(Math.min(...ops.map(p => p.y)) - Y0, ref).toBe(o.icon!.topPx - o.cellTopPx)
      }
      batch.dispose()
    }, { model: false, glyphs: false, extra: caps })
  })
})

/**
 * Task 28 round 4, Ruling S4-BOTTOM-REPLICATE: bottom placement captured with styles that survive Excel's resave, and
 * the anchor-constant vs centred decision resting on at least two separating columns.
 */
describe('S4-BOTTOM-REPLICATE: stage4-icon-rows-bottom', () => {
  const BOTTOM = () => CF_STAGE4_CASES.find(c => c.id === 'stage4-icon-rows-bottom')!
  const BCOLS: Record<string, { cls: string; horizontal: 'right' | 'left' }> = {
    B: { cls: '3Arrows:2', horizontal: 'right' },
    C: { cls: '3TrafficLights1:2', horizontal: 'right' },
    D: { cls: '3Symbols:2', horizontal: 'right' },
    E: { cls: '3Arrows:2', horizontal: 'left' },
    F: { cls: '3TrafficLights1:2', horizontal: 'left' },
  }

  it('the case is ahead of stage4-icon-rows and owns icon:geometry-row-offset, over the same heights', () => {
    const ids = CF_STAGE4_CASES.map(c => c.id)
    expect(ids.indexOf('stage4-icon-rows-bottom')).toBe(ids.indexOf('stage4-icon-rows') - 1)
    expect(CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === CAP)?.caseId).toBe('stage4-icon-rows-bottom')
    const c = BOTTOM()
    expect(c.status).toBe('recipe-pending')
    expect(c.iconShiftedKey).toBe(true)
    const markers = c.markers.filter(m => m.capabilityId === CAP)
    expect(markers.map(m => [m.sqref, m.icon?.set])).toEqual(Object.entries(BCOLS).map(([col, v]) => [`${col}2:${col}9`, v.cls.split(':')[0]]))
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(ROOT, c.input))))!
    const geom = computeWorksheetRangeGeometry(parts, c.range)!
    expect(geom.rowHeightsPt.slice(1, 9)).toEqual(ROWS_PT)
  })

  it('every designated cell has an explicit xf with horizontal and vertical="bottom" (no cell uses the Normal style)', () => {
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(ROOT, BOTTOM().input))))!
    const dec = new TextDecoder()
    const styles = dec.decode(parts.get('xl/styles.xml')!)
    const xfs = [...(/<cellXfs[^>]*>(.*?)<\/cellXfs>/s.exec(styles)![1]!.matchAll(/<xf\b[^>]*?(?:\/>|>.*?<\/xf>)/gs))].map(m => m[0])
    const sheet = dec.decode(parts.get('xl/worksheets/sheet1.xml')!)
    for (const [col, spec] of Object.entries(BCOLS)) for (let row = 2; row <= 9; row++) {
      const s = Number(new RegExp(`<c r="${col}${row}"(?: s="(\\d+)")?`).exec(sheet)?.[1] ?? 0)
      expect(s, `${col}${row}`).toBeGreaterThan(0)
      expect(xfs[s], `${col}${row}`).toContain(`horizontal="${spec.horizontal}"`)
      expect(xfs[s], `${col}${row}`).toContain('vertical="bottom"')
    }
  })

  it('the resave keeps an explicit horizontal: a left+bottom xf came back as its own xf (stage4-icon-rows E), drawn at the bottom', () => {
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(ROOT, CASE().saved))))!
    const styles = new TextDecoder().decode(parts.get('xl/styles.xml')!)
    expect(styles).toContain('<alignment horizontal="left"/>')
    const sidecar = JSON.parse(readFileSync(resolve(ROOT, 'tests/fixtures/xlsx-cf/refs/stage4-icon-rows.json'), 'utf8')) as { observedIcons: Array<{ ref: string; cellTopPx: number; cellHeightPx: number; icon: null | { bottomPx: number } }> }
    for (const ref of ['E2', 'E4', 'E5', 'E6', 'E7', 'E8', 'E9']) {
      const o = sidecar.observedIcons.find(x => x.ref === ref)!
      expect(o.cellTopPx + o.cellHeightPx - 1 - o.icon!.bottomPx, ref).toBe(1)
    }
  })

  it('replication: the committed stage4-icon-rows capture (one separating column, E) no longer discriminates anchor-constant from centred', () => {
    const c = CASE()
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(ROOT, c.saved))))!
    const sidecar = JSON.parse(readFileSync(resolve(ROOT, 'tests/fixtures/xlsx-cf/refs/stage4-icon-rows.json'), 'utf8')) as { observedIcons: CfObservedIcon[] }
    const observed = new Map(sidecar.observedIcons.map(o => [o.ref, o] as const))
    const res = evaluateIconCapabilities(c, parts, observed, undefined)[CAP]!
    expect(res.status, res.reason).toBe('non-discriminating')
    expect(res.reason).toMatch(/S4-REPLICATE/)
  })

  const key = loadIconGlyphKey(ROOT) as IconGlyphKey
  const captureBottom = (frameTop: (rowPx: number) => number, onlyCols?: readonly string[]) => {
    const c = BOTTOM()
    const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(ROOT, c.input))))!
    const geom = computeWorksheetRangeGeometry(parts, c.range)!
    const W = geom.expectedWidthPt, H = geom.expectedHeightPt
    const canvas = createCanvas(Math.round(W) + 1, Math.round(H) + 1)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height)
    const img = ctx.getImageData(0, 0, canvas.width, canvas.height)
    for (const e of c.expected.cells.filter(x => x.matched)) {
      const col = /^[A-Z]+/.exec(e.ref)![0]
      const p = parseA1(e.ref)!
      const x0 = Math.round(geom.colWidthsPt.slice(0, p.col - geom.minCol).reduce((a, b) => a + b, 0))
      const y0 = Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow).reduce((a, b) => a + b, 0))
      const h = Math.round(geom.rowHeightsPt.slice(0, p.row - geom.minRow + 1).reduce((a, b) => a + b, 0)) - y0
      // Columns outside `onlyCols` are drawn centred (as a column that cannot separate the readings would be).
      const top = !onlyCols || onlyCols.includes(col) ? frameTop(h) : Math.round((h - 12) / 2)
      const crop = key.classes.get(BCOLS[col]!.cls)!
      for (let y = 0; y < crop.height; y++) for (let x = 0; x < crop.width; x++) {
        const k = (y * crop.width + x) * 4
        if (crop.data[k]! > 247 && crop.data[k + 1]! > 247 && crop.data[k + 2]! > 247) continue
        const X = x0 + x + 1, Y = y0 + y + 1 + top - 7
        if (Y <= y0 || Y >= y0 + h - 1) continue
        const o = (Y * canvas.width + X) * 4
        img.data[o] = crop.data[k]!; img.data[o + 1] = crop.data[k + 1]!; img.data[o + 2] = crop.data[k + 2]!; img.data[o + 3] = 255
      }
    }
    ctx.putImageData(img, 0, 0)
    const decoded = decodePngSync(new Uint8Array(canvas.toBuffer('image/png')))!
    const observed = measureCaseIcons(c, decoded, geom, key) as ReadonlyMap<string, CfObservedIcon>
    return evaluateIconCapabilities(c, parts, observed, key)[CAP]!
  }

  it('a capture drawn bottom-anchored on all five columns discriminates anchor-constant (mechanism test)', () => {
    expect(captureBottom(h => h - 13)).toMatchObject({ status: 'discriminating', matched: 'anchor-constant' })
  })
  it('a capture drawn centred discriminates centred (mechanism test)', () => {
    expect(captureBottom(h => Math.round((h - 12) / 2))).toMatchObject({ status: 'discriminating', matched: 'centred' })
  })
  it('replication: one separating column is not enough, two are (mechanism test)', () => {
    // Bottom-anchored cells (gap 1 to the next gridline) on the measured heights, one column per glyph.
    const cells = (cols: readonly string[]): IconGeometryCell[] => cols.flatMap(col => [15, 16, 18, 20, 22, 25, 30].map((h, i) => {
      const ref = `${col}${i + 2}`, x0 = 100, y0 = 20 + i * 40
      return {
        ref, facts: { fontPt: 11, vertical: 'bottom' as const, horizontal: 'right' as const, rowPt: h, showValue: true },
        observed: { ref, cellLeftPx: x0, cellTopPx: y0, cellWidthPx: 72, cellHeightPx: h, icon: { leftPx: x0 + 2, topPx: y0 + h - 13, rightPx: x0 + 12, bottomPx: y0 + h - 2, glyphClass: '3Arrows:2' } },
      }
    }))
    const one = evaluateIconGeometryCapability(CAP, cells(['B']), cells(['B']))
    expect(one.status, one.reason).toBe('non-discriminating')
    expect(one.reason).toMatch(/S4-REPLICATE/)
    expect(evaluateIconGeometryCapability(CAP, cells(['B', 'C']), cells(['B', 'C']))).toMatchObject({ status: 'discriminating', matched: 'anchor-constant' })
  })
})
