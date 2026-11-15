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
import { CF_STAGE4_CAPABILITIES, CF_STAGE4_CASES, iconCapabilityKind } from '../scripts/cf-refs/cases'
import {
  ICON_GEOMETRY_READINGS, type CfObservedIcon, type IconGlyphKey, computeWorksheetRangeGeometry, decodePngSync, evaluateIconCapabilities,
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
    expect(CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === CAP)?.caseId).toBe('stage4-icon-rows')
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
