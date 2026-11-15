// @vitest-environment node
/**
 * Task 28 (OD-9(b), owner decision 2026-11-14): the lazy table of original, hand-authored vector icon glyphs, and the
 * icon frame that lets icons paint where Excel's glyph-key capture pins them down.
 *
 * Everything is compared against `tests/fixtures/xlsx-cf/refs/stage4-icon-glyph-key.{png,json}` (Excel 16.115),
 * loaded through the harness's `loadIconGlyphKey`; its crops are comparison fixtures only (regression tests on
 * Excel's stored numbers, not independent evidence).
 */
import { existsSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { gzipSync } from 'node:zlib'
import { describe, expect, it, vi } from 'vitest'
import {
  CF_ICON_GLYPH_CLASSES, CF_ICON_GLYPH_WITHHELD, admitIconGlyphs, cfDecodedGlyphCount, cfGlyphBytesBound, cfIconGlyphClass, cfIconGlyphHeader, drawIconGlyph, glyphBytes, iconGlyph, type CfIconGlyph,
} from '../src/xlsx/cf/iconGlyphs'
import { CF_ICON_PAINT_CAPABILITIES, CF_ICON_RENDER_MODEL, admitBatchGlyphs, layoutIcon, paintIconCells } from '../src/xlsx/cf/iconPaint'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { paintMemoryFor } from '../src/xlsx/cf/render'
import type { CfIconSetName } from '../src/xlsx/cf/model'
import { loadIconGlyphKey, type IconGlyphKey } from '../scripts/cf-refs/verify'
import { S4_ICON_FAMILIES } from '../scripts/cf-refs/cases'
import { createCanvas } from 'canvas'
import { AT_CAPTURE, KEY_CELL_W, KEY_FONT, checkGlyph, cropInkBox, cropMedian, inkBoxOf, keyHost } from './helpers/cf-glyph-measure'
import { ICON_CLASS_TOLERANCE } from '../scripts/cf-refs/verify'

const ROOT = resolve(__dirname, '..')
const SIDECAR = JSON.parse(readFileSync(join(ROOT, 'tests/fixtures/xlsx-cf/refs/stage4-icon-glyph-key.json'), 'utf8')) as {
  glyphClasses: Record<string, string>; glyphFaint: string[]
  observedIcons: Array<{ ref: string; cellLeftPx: number; cellTopPx: number; icon: null | { leftPx: number; topPx: number; rightPx: number; bottomPx: number } }>
}
const KEY = loadIconGlyphKey(ROOT) as IconGlyphKey
const PAIRS = Object.keys(SIDECAR.glyphClasses)
const CLASSES = [...new Set(Object.values(SIDECAR.glyphClasses))]
const glyphCaps = (): string[] => PAIRS.map(p => { const [set, id] = p.split(':'); return `icon:glyph-${set}-${id}` })
const withGlyphEvidence = <T>(fn: () => T, caps: readonly string[] = glyphCaps()): T => withExactCfEvidence(caps, fn)
const split = (pair: string): [CfIconSetName, number] => { const [set, id] = pair.split(':'); return [set as CfIconSetName, Number(id)] }
/** The key cell of a pair: one row per family from row 2 (`S4_ICON_FAMILIES` order), iconId 0..4 in B..F. */
const keyRef = (pair: string): string => {
  const [set, id] = split(pair)
  return `${'BCDEF'[id]}${S4_ICON_FAMILIES.findIndex(f => f.name === set) + 2}`
}
/** The sidecar's measured ink box of a class's key cell (cell-relative, inclusive); undefined for a faint class. */
const measuredBox = (cls: string) => {
  const o = SIDECAR.observedIcons.find(x => x.ref === keyRef(cls))!
  return o.icon ? { left: o.icon.leftPx - o.cellLeftPx, top: o.icon.topPx - o.cellTopPx, right: o.icon.rightPx - o.cellLeftPx, bottom: o.icon.bottomPx - o.cellTopPx } : undefined
}
/** Every class's glyph, read past the evidence gate. */
const glyphOf = (cls: string): CfIconGlyph => withGlyphEvidence(() => cfIconGlyphClass(cls)!)
const checks = new Map(CLASSES.map(cls => {
  const faint = SIDECAR.glyphFaint.includes(cls)
  const crop = KEY.classes.get(cls)!
  return [cls, checkGlyph(glyphOf(cls), crop, faint ? cropInkBox(crop, ICON_CLASS_TOLERANCE + 1) : measuredBox(cls), faint)] as const
}))

describe('Task 28: the glyph table', () => {
  it('the glyph key is the verified Excel 16.115 capture: 54 classes over 75 pairs', () => {
    expect(KEY && !('kind' in KEY)).toBe(true)
    expect(PAIRS).toHaveLength(75)
    expect(CLASSES).toHaveLength(54)
    expect(KEY.classes.size).toBe(CLASSES.length)
  })

  it('every (set, iconId) maps to the measured glyph class', () => {
    expect(CF_ICON_GLYPH_CLASSES).toEqual(SIDECAR.glyphClasses)
    withGlyphEvidence(() => {
      const entries = new Map<string, CfIconGlyph>()
      for (const pair of PAIRS) {
        const cls = SIDECAR.glyphClasses[pair]!
        const g = CF_ICON_GLYPH_WITHHELD.has(cls) ? cfIconGlyphHeader(cls) : iconGlyph(...split(pair))
        expect(g?.glyphClass, pair).toBe(cls)
        // Pairs that share a crop share one table entry.
        if (entries.has(cls)) expect(g).toBe(entries.get(cls))
        else entries.set(cls, g!)
      }
      // One entry per class; twins (S4-NEARTWIN) and the faint class (S4-FAINT) are entries of their own.
      expect(entries.size).toBe(CLASSES.length)
      for (const [a, b] of [['4RedToBlack:0', '5Quarters:4'], ['4RedToBlack:1', '5Boxes:0']]) expect(entries.get(a)).not.toBe(entries.get(b))
    })
  })

  it('unknown or unverified glyphs are withheld', () => {
    withGlyphEvidence(() => {
      expect(iconGlyph('3Arrows', 3)).toBeUndefined()
      expect(iconGlyph('3Arrows', -1)).toBeUndefined()
      expect(iconGlyph('NoIcons' as CfIconSetName, 0)).toBeUndefined()
      expect(iconGlyph('3Arrows', 0)?.glyphClass).toBe('3Arrows:0')
    })
    // Each pair's own capability gates it: without `icon:glyph-3Arrows-0` the pair withholds, its class-mates do not.
    withGlyphEvidence(() => {
      expect(iconGlyph('3Arrows', 0)).toBeUndefined()
      expect(iconGlyph('4Arrows', 0)?.glyphClass).toBe('3Arrows:0')
    }, glyphCaps().filter(c => c !== 'icon:glyph-3Arrows-0'))
    // Production publishes no glyph capability before Task 30: every pair withholds.
    for (const pair of PAIRS) expect(iconGlyph(...split(pair)), pair).toBeUndefined()
  })

  it('glyph decode is charged and bounded', () => {
    const owner = {}
    const memory = paintMemoryFor(owner)
    const baseline = memory.currentBytes
    const budget = createCfWorkBudget({ now: () => 0 })
    const all = new Set(CLASSES.filter(c => !CF_ICON_GLYPH_WITHHELD.has(c)))
    const cache = admitIconGlyphs(all, memory, budget, cfIconGlyphHeader)!
    expect(cache).toBeDefined()
    const expected = [...all].reduce((n, c) => n + cfGlyphBytesBound(c)!, 0)
    expect(cache.bytes).toBe(expected)
    expect(memory.currentBytes - baseline).toBe(expected)
    // Decoding spends work units (one per op, plus one per class), never wall time.
    expect(budget.units).toBe([...all].reduce((n, c) => n + cfIconGlyphClass(c)!.path.length + 1, 0))

    // Drawing an admitted glyph at each zoom/DPR allocates nothing: the ledger is unchanged.
    const ctx = createCanvas(64, 64).getContext('2d') as unknown as CanvasRenderingContext2D
    for (const [zoom, dpr] of [[0.5, 1], [0.75, 1], [1, 1], [1, 2], [1.5, 2], [2, 3]] as const) {
      for (const cls of all) expect(drawIconGlyph(ctx, cfIconGlyphHeader(cls)!, { x: 1, y: 1, w: 13 * zoom * dpr, h: 12 * zoom * dpr }, { zoom, dpr }, cache)).toBe(true)
      expect(memory.currentBytes - baseline).toBe(expected)
    }
    cache.release()
    expect(memory.currentBytes).toBe(baseline)

    // A refused charge returns false/undefined, draws nothing and leaves the ledger at baseline.
    const tiny = createCfMemory({ limitBytes: 16 })
    expect(admitIconGlyphs(all, tiny, createCfWorkBudget({ now: () => 0 }), cfIconGlyphHeader)).toBeUndefined()
    expect(tiny.currentBytes).toBe(0)
    const refusing = { spend: () => 'refused' as const, resume() {}, units: 0 }
    expect(admitIconGlyphs(all, memory, refusing, cfIconGlyphHeader)).toBeUndefined()
    expect(memory.currentBytes).toBe(baseline)
    const cls = [...all][0]!
    const other = admitIconGlyphs(new Set([[...all][1]!]), memory, createCfWorkBudget({ now: () => 0 }), cfIconGlyphHeader)!
    const blank = createCanvas(20, 20)
    const bctx = blank.getContext('2d') as unknown as CanvasRenderingContext2D
    expect(drawIconGlyph(bctx, cfIconGlyphHeader(cls)!, { x: 0, y: 0, w: 13, h: 12 }, AT_CAPTURE, other)).toBe(false)
    expect([...bctx.getImageData(0, 0, 20, 20).data].every(v => v === 0)).toBe(true)
    other.release()
    expect(memory.currentBytes).toBe(baseline)
  })

  it('original glyphs approximate the measured box and colours (OD-9(b)), labelled pixel-approximated', () => {
    const failing = new Set<string>()
    for (const [cls, check] of checks) {
      if (!check.ok) failing.add(cls)
      if (!CF_ICON_GLYPH_WITHHELD.has(cls)) expect(check.reasons, cls).toEqual([])
    }
    // A glyph failing either check is withheld for all its pairs, and only such a glyph.
    expect([...CF_ICON_GLYPH_WITHHELD].sort()).toEqual([...failing].sort())
    withGlyphEvidence(() => {
      for (const pair of PAIRS) if (failing.has(SIDECAR.glyphClasses[pair]!)) expect(iconGlyph(...split(pair)), pair).toBeUndefined()
    })
    for (const cls of CLASSES) expect(glyphOf(cls).fidelity, cls).toBe('pixel-approximated')
  })

  it('the glyph table ships no Microsoft artwork', () => {
    const src = readFileSync(join(ROOT, 'src/xlsx/cf/iconGlyphs.ts'), 'utf8')
    expect(src).not.toMatch(/data:/)
    expect(src).not.toMatch(/[A-Za-z0-9+/=]{65,}/)
    // No bitmap construction or decoding (the comments may cite the key capture by name; code may not read it).
    expect(src).not.toMatch(/ImageData|createImageData|putImageData|decodePng|Uint8ClampedArray|\bImage\(|drawImage|readFileSync|fetch\(/)
    expect(src).not.toMatch(/from\s+['"][^'"]*tests\/fixtures/)
    expect(src).not.toMatch(/from\s+['"][^'"]*scripts\//)
    for (const cls of CLASSES) expect(glyphOf(cls).source, cls).toBe('original')
  })

  // The built `cf` chunk's record of this module (OV_MODULE_MAP=1 bun run build writes dist/.module-map.json).
  const mapFile = join(ROOT, 'dist', '.module-map.json')
  it.skipIf(!existsSync(mapFile))('glyph table bytes are measured in the built lazy cf chunk', () => {
    type Chunk = { modules: string[]; moduleBytes?: Record<string, { raw: number; gzip: number }> }
    const map = JSON.parse(readFileSync(mapFile, 'utf8')) as Record<string, Chunk>
    const hits = Object.entries(map).flatMap(([file, c]) => c.modules.filter(m => m.endsWith('src/xlsx/cf/iconGlyphs.ts')).map(m => ({ file, m, bytes: c.moduleBytes?.[m] })))
    expect(hits).toHaveLength(1)
    const { bytes } = hits[0]!
    expect(bytes).toBeDefined()
    // Ruling S4-GLYPHCAP: the ceiling before measurement is 16 KiB gzip; the measured size (table + decoder, Task 28,
    // reported in Task 30) is CF_GLYPH_TABLE_GZIP_BYTES, and the assertion holds it to that plus 10 %.
    expect(bytes!.gzip).toBeLessThanOrEqual(16 * 1024)
    expect(bytes!.gzip).toBeLessThanOrEqual(Math.floor(CF_GLYPH_TABLE_GZIP_BYTES * 1.1))
    // Cross-check: the source table alone (not shipped as such) is far below the ceiling too.
    expect(gzipSync(readFileSync(join(ROOT, 'src/xlsx/cf/iconGlyphs.ts'))).length).toBeLessThanOrEqual(16 * 1024)
  })
})

/** Measured 2026-11-15 by `OV_MODULE_MAP=1 npx vite build` (dist/.module-map.json `moduleBytes`): the iconGlyphs.ts
 * module (encoded table, class map, headers, admission-owned decoder) renders 18,192 B, 5,462 B gzip, inside the lazy
 * cf chunk (round 7: the redrawn classes). */
const CF_GLYPH_TABLE_GZIP_BYTES = 5462

describe('Task 28: the icon frame at the glyph key geometry', () => {
  const icon = { ruleId: 1, set: '3Arrows' as const, iconId: 0, bucket: 0, buckets: 3, showValue: true }
  const ALL = (): boolean => true
  it('S4-ROW-EXACT: the key\'s own 20 pt row withholds (Excel draws 20 px; the viewer\'s 27 or 80/3 px rows are 20.25 or 20 px at capture)', () => {
    expect(layoutIcon(icon, keyHost({ rowPt: 20 }), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)).toBeUndefined()
    expect(layoutIcon(icon, keyHost({ rowPt: 20, height: 27 }), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)).toBeUndefined()
    // The frame (13 × 12, 2 px in) lays out on an admitted 15 pt row, whatever the column width.
    expect(layoutIcon(icon, keyHost({ rowPt: 15, height: 20 }), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)?.box).toEqual({ x: 2, y: 2, w: 13, h: 12 })
    expect(layoutIcon(icon, keyHost({ rowPt: 15, height: 20, width: 240 }), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)?.box).toEqual({ x: 2, y: 2, w: 13, h: 12 })
  })

  it('with its evidence, a key-geometry host paints the original glyph where Excel inks it (left-aligned text, as I5/J5)', () => {
    const caps = [...glyphCaps(), ...CF_ICON_PAINT_CAPABILITIES, 'icon:inset-left', 'icon:geometry-row-offset']
    withGlyphEvidence(() => {
      const memory = createCfMemory({})
      for (const pair of ['3Arrows:0', '3TrafficLights1:2', '5Boxes:2']) {
        const [set, iconId] = split(pair)
        const desc = { ruleId: 1, set, iconId, bucket: iconId, buckets: Number(set[0]), showValue: true }
        const admitted = admitBatchGlyphs([desc], memory, () => createCfWorkBudget({ now: () => 0 }))
        // A 15 pt row (admitted): bottom-anchored, the frame top is 15 - 1 - 12 = 2, 5 px above the key's 7.
        const H = 15
        const canvas = createCanvas(KEY_CELL_W, H)
        const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, KEY_CELL_W, H)
        const host = keyHost({ rowPt: H, height: H * 4 / 3, textWidth: 10, cell: { value: 'ab', style: { numFmtId: 0, fontSizePt: 11, fontName: 'Calibri', horizontal: 'left' } } })
        const layouts: boolean[] = []
        paintIconCells(ctx, new Map([[0, desc]]), () => desc, () => host, { r0: 0, c0: 0, r1: 0, c1: 0 }, AT_CAPTURE,
          { admitted, hasBar: () => false, check: () => true, rowPtOf: () => H }, (_r, _c, withheld) => layouts.push(withheld))
        expect(layouts, pair).toEqual([false])
        const drawn = checks.get(SIDECAR.glyphClasses[pair]!)!.drawn!
        expect(inkBoxOf(ctx.getImageData(0, 0, KEY_CELL_W, H).data, KEY_CELL_W, H), pair).toEqual({ ...drawn, top: drawn.top - 5, bottom: drawn.bottom - 5 })
        admitted.release()
      }
      expect(memory.currentBytes).toBe(0)
    }, caps)
  })

  it('every input the key fixes must equal the key: other font sizes, rows, alignments and faces withhold', () => {
    const at = (host: ReturnType<typeof keyHost>, font: { sizePt?: number; vertical?: 'top' | 'center' | 'bottom' } = {}) =>
      layoutIcon(icon, host, { ...KEY_FONT, ...font }, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)
    const row15 = (o: Record<string, unknown> = {}) => keyHost({ rowPt: 15, height: 20, ...o })
    expect(at(row15())).toBeDefined()
    for (const sizePt of [8, 10, 10.5, 12, 20]) expect(at(row15(), { sizePt }), `${sizePt} pt`).toBeUndefined()
    for (const rowPt of [12.75, 19, 21, 35, 60]) expect(at(keyHost({ rowPt, height: rowPt * 4 / 3 })), `${rowPt} pt row`).toBeUndefined()
    expect(at(row15(), { vertical: 'top' }), 'top').toBeUndefined()
    const style = { numFmtId: 0, fontSizePt: 11, fontName: 'Calibri' }
    for (const s of [{ ...style, fontName: 'Arial' }, { ...style, fontName: undefined }, { ...style, bold: true }, { ...style, italic: true }]) {
      expect(at(row15({ cell: { value: 10, style: s } })), JSON.stringify(s)).toBeUndefined()
    }
  })
})

/**
 * Task 28 round 3: the frame at the heights `refs/stage4-icon-rows.json` measured (Excel 16.115, cd5ae569), which
 * matched `anchor-constant`. Every offset is a measured px at 1 px/pt; nothing between the measured heights.
 */
describe('Task 28 round 3: the frame at the measured row heights (stage4-icon-rows)', () => {
  const icon = { ruleId: 1, set: '3Arrows' as const, iconId: 2, bucket: 2, buckets: 3, showValue: true }
  const ALL = (): boolean => true
  const ROWS = SIDECAR_ROWS()
  const font = (vertical: 'bottom' | 'center', horizontal: 'left' | 'right' | 'center') => ({ sizePt: 11, vertical, horizontal })
  const at = (rowPt: number, vertical: 'bottom' | 'center', horizontal: 'left' | 'right' | 'center', height = rowPt * 4 / 3, hostRowPt: number | null = rowPt) =>
    layoutIcon(icon, keyHost({ height, ...(hostRowPt !== null ? { rowPt: hostRowPt } : {}) }), font(vertical, horizontal), AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)

  it('the model carries exactly the sidecar\'s measured placements', () => {
    expect(CF_ICON_RENDER_MODEL.keyFrame?.rows).toEqual({
      15: { excelPx: 15, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left', 'right'] }, center: { anchor: 'top', offsetPt: 2, horizontal: ['left', 'right'] } },
      16.5: { excelPx: 16, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left', 'right'] }, center: { anchor: 'top', offsetPt: 2, horizontal: ['left', 'right'] } },
      18: { excelPx: 18, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left', 'right'] }, center: { anchor: 'top', offsetPt: 3, horizontal: ['left', 'right'] } },
      20: { excelPx: 20, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left', 'right'] }, center: { anchor: 'top', offsetPt: 4, horizontal: ['left', 'right'] } },
      22.5: { excelPx: 22, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left', 'right'] }, center: { anchor: 'top', offsetPt: 5, horizontal: ['left', 'right'] } },
      25: { excelPx: 25, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left', 'right'] }, center: { anchor: 'top', offsetPt: 7, horizontal: ['left', 'right'] } },
      30: { excelPx: 30, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left', 'right'] }, center: { anchor: 'top', offsetPt: 9, horizontal: ['left', 'right'] } },
    })
  })

  it('the placements are the sidecar\'s: frame top = ink top of 3Arrows:2, 3TrafficLights1:2 and 3Symbols:2 (each class\'s key ink top is the frame top)', () => {
    // E (bottom, left numbers) and B..D, F..I (centre: Excel's resave gives the Normal style vertical="center").
    for (const [rowPt, row] of Object.entries(ROWS)) {
      const r = CF_ICON_RENDER_MODEL.keyFrame!.rows[Number(rowPt)]!
      for (const top of row.centre) expect(top, `${rowPt} centre`).toBe(r.center!.offsetPt)
      for (const [top, h] of row.bottom) expect(h - (top + 12), `${rowPt} bottom gap`).toBe(r.bottom!.offsetPt)
    }
  })

  // S4-ROW-EXACT: a measured height paints only where the viewer's row px equals Excel's at the capture scale in every
  // view: the static and resident views round rows to whole px, the paged ones do not, so only heights whose px are whole
  // (pt × 4/3) agree: 15 (20 px), 18 (24 px) and 30 (40 px).
  for (const rowPt of [15, 18, 30]) {
    it(`a ${rowPt} pt row lays the frame out where Excel measured it`, () => {
      const r = CF_ICON_RENDER_MODEL.keyFrame!.rows[rowPt]!
      expect(r.excelPx).toBe(rowPt)
      for (const h of r.bottom!.horizontal) expect(at(rowPt, 'bottom', h)?.box, `bottom ${h}`).toEqual({ x: 2, y: rowPt - 1 - 12, w: 13, h: 12 })
      for (const h of r.center!.horizontal) expect(at(rowPt, 'center', h)?.box, `center ${h}`).toEqual({ x: 2, y: r.center!.offsetPt, w: 13, h: 12 })
      expect(r.bottom!.horizontal, 'round 5: stage4-icon-rows-bottom measured right-aligned numbers too').toEqual(['left', 'right'])
      // An alignment the capture did not measure at this height withholds.
      expect(at(rowPt, 'center', 'center'), 'center text').toBeUndefined()
      expect(at(rowPt, 'bottom', 'center'), 'bottom center text').toBeUndefined()
      // A host whose row is not known to be this height withholds: no row pt, or another pt with the same px.
      expect(at(rowPt, 'center', 'right', rowPt * 4 / 3, null), 'no row pt').toBeUndefined()
      expect(at(rowPt, 'center', 'right', rowPt * 4 / 3, rowPt - 0.2), 'another pt, same px').toBeUndefined()
      // A session-resized row (another px) withholds.
      expect(at(rowPt, 'center', 'right', rowPt * 4 / 3 + 1), 'resized').toBeUndefined()
    })
  }

  it('S4-ROW-EXACT: measured heights whose viewer px are not whole (16.5, 20, 22.5, 25 pt) withhold in every view', () => {
    for (const rowPt of [16.5, 20, 22.5, 25]) for (const height of [rowPt * 4 / 3, Math.round(rowPt * 4 / 3)]) {
      for (const v of ['bottom', 'center'] as const) expect(at(rowPt, v, 'left', height), `${rowPt} pt ${height} px ${v}`).toBeUndefined()
    }
  })

  it('unmeasured heights withhold: 12.75 pt (the icon shrinks) and every height between the measured ones', () => {
    for (const rowPt of [12, 12.75, 14.25, 15.75, 17.25, 19.5, 21, 24, 27, 29.25, 33]) {
      for (const v of ['bottom', 'center'] as const) expect(at(rowPt, v, 'left'), `${rowPt} ${v}`).toBeUndefined()
    }
  })
})

/** Per measured row (pt): the centre cells' ink tops and the bottom cells' (ink top, row px), from the sidecar. */
function SIDECAR_ROWS(): Record<string, { centre: number[]; bottom: Array<[number, number]> }> {
  const sc = JSON.parse(readFileSync(join(ROOT, 'tests/fixtures/xlsx-cf/refs/stage4-icon-rows.json'), 'utf8')) as typeof SIDECAR & { observedIcons: Array<{ ref: string; cellTopPx: number; cellHeightPx: number; icon: null | { topPx: number } }> }
  const PT: Record<string, number> = { 2: 15, 3: 12.75, 4: 16.5, 5: 18, 6: 20, 7: 22.5, 8: 25, 9: 30 }
  const out: Record<string, { centre: number[]; bottom: Array<[number, number]> }> = {}
  for (const o of sc.observedIcons) {
    const [, col, row] = /^([A-Z]+)(\d+)$/.exec(o.ref)!
    const pt = PT[row!]!
    if (pt === 12.75 || !o.icon) continue
    out[pt] ??= { centre: [], bottom: [] }
    if (col === 'E') out[pt]!.bottom.push([o.icon.topPx - o.cellTopPx, o.cellHeightPx])
    else out[pt]!.centre.push(o.icon.topPx - o.cellTopPx)
  }
  return out
}

describe('S4-BOTTOM-REPLICATE: the measured placements stand only on icon:geometry-row-offset', () => {
  it('without its evidence every row withholds', () => {
    const icon = { ruleId: 1, set: '3Arrows' as const, iconId: 2, bucket: 2, buckets: 3, showValue: true }
    const host = keyHost({ rowPt: 15, height: 20 })
    expect(layoutIcon(icon, host, KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, () => true)).toBeDefined()
    expect(layoutIcon(icon, host, KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, c => c !== 'icon:geometry-row-offset')).toBeUndefined()
  })
})

/** Fix round (codex review of rounds 1–5). */
describe('Task 28 fix round', () => {
  it('P1: negative controls fail the approximation check (a solid median-colour box for 3TrafficLights1:0 and each 3Symbols glyph)', () => {
    for (const cls of ['3TrafficLights1:0', '3Symbols:0', '3Symbols:1', '3Symbols:2']) {
      const crop = KEY.classes.get(cls)!
      const b = cropInkBox(crop, 64)!
      const pix: Array<{ x: number; y: number }> = []
      for (let y = b.top; y <= b.bottom; y++) for (let x = b.left; x <= b.right; x++) pix.push({ x, y })
      const m = cropMedian(crop, pix)!
      const argb = (0xff000000 | (m[0] << 16) | (m[1] << 8) | m[2]) >>> 0
      const x0 = b.left - 2, y0 = b.top - 7, x1 = b.right - 1, y1 = b.bottom - 6
      const box: CfIconGlyph = { glyphClass: cls, widthPt: 13, heightPt: 12, source: 'original', path: [['M', x0, y0], ['L', x1, y0], ['L', x1, y1], ['L', x0, y1], ['Z'], ['F', argb]] }
      const check = checkGlyph(box, crop, measuredBox(cls), false)
      expect(check.ok, cls).toBe(false)
      // It fails on the independent checks, not on the box or the candidate-region palette.
      expect(check.reasons.some(r => r.startsWith('reference region') || r.startsWith('ink IoU')), cls).toBe(true)
    }
  })

  it('P1: admission charges before decoding; a refused admission decodes nothing; release drops the decoded glyphs', () => {
    expect(cfDecodedGlyphCount()).toBe(0)
    const all = new Set(CLASSES.filter(c => !CF_ICON_GLYPH_WITHHELD.has(c)))
    // The charge is computed from the encoded table and equals the decoded size.
    for (const cls of all) expect(cfGlyphBytesBound(cls), cls).toBe(glyphBytes(cfIconGlyphClass(cls)!))
    const tiny = createCfMemory({ limitBytes: 16 })
    expect(admitIconGlyphs(all, tiny, createCfWorkBudget({ now: () => 0 }), cfIconGlyphHeader)).toBeUndefined()
    expect(cfDecodedGlyphCount()).toBe(0)
    expect(tiny.currentBytes).toBe(0)
    const refusing = { spend: () => 'refused' as const, resume() {}, units: 0 }
    expect(admitIconGlyphs(all, createCfMemory({}), refusing, cfIconGlyphHeader)).toBeUndefined()
    expect(cfDecodedGlyphCount()).toBe(0)
    const memory = createCfMemory({})
    const cache = admitIconGlyphs(all, memory, createCfWorkBudget({ now: () => 0 }), cfIconGlyphHeader)!
    expect(cfDecodedGlyphCount()).toBe(all.size)
    const one = [...all][0]!
    expect(cache.glyph?.(one)?.path.length).toBeGreaterThan(0)
    cache.release()
    expect(cfDecodedGlyphCount()).toBe(0)
    expect(cache.glyph?.(one)).toBeUndefined()
    expect(memory.currentBytes).toBe(0)
    // The production lookup returns an undecoded header: no ops until admission.
    withGlyphEvidence(() => expect(iconGlyph('3Arrows', 0)?.path).toEqual([]))
    expect(cfDecodedGlyphCount()).toBe(0)
  })

  it('P2: drawing builds no colour strings (prepared and charged at admission)', () => {
    const all = new Set(CLASSES.filter(c => !CF_ICON_GLYPH_WITHHELD.has(c)))
    const cache = admitIconGlyphs(all, createCfMemory({}), createCfWorkBudget({ now: () => 0 }), cfIconGlyphHeader)!
    const ctx = createCanvas(64, 64).getContext('2d') as unknown as CanvasRenderingContext2D
    const toString = vi.spyOn(Number.prototype, 'toString'), padStart = vi.spyOn(String.prototype, 'padStart')
    try {
      for (let i = 0; i < 2; i++) for (const cls of all) expect(drawIconGlyph(ctx, cfIconGlyphHeader(cls)!, { x: 1, y: 1, w: 13, h: 12 }, AT_CAPTURE, cache)).toBe(true)
      expect(toString).not.toHaveBeenCalled()
      expect(padStart).not.toHaveBeenCalled()
    } finally { toString.mockRestore(); padStart.mockRestore(); cache.release() }
  })
})

/**
 * Round 7 (S4-REDRAW) tracing guard: every glyph is built from a few geometric primitives. A pixel trace of these
 * glyphs (60-130 ink pixels each, four vertices a pixel) would need hundreds of vertices and dozens of fills.
 * Measured maxima at round 7: 8 fills (3Symbols2:1, 5Rating:0) and 32 vertices (3Symbols2:1, 5Rating:0).
 */
describe('S4-REDRAW: glyphs stay primitive', () => {
  const MAX_FILLS = 12, MAX_VERTICES = 48
  it(`every glyph has at most ${MAX_FILLS} fills and ${MAX_VERTICES} path vertices (M, L, A)`, () => {
    for (const cls of CLASSES) {
      const g = cfIconGlyphClass(cls)!
      expect(g.path.filter(o => o[0] === 'F').length, cls).toBeLessThanOrEqual(MAX_FILLS)
      expect(g.path.filter(o => o[0] === 'M' || o[0] === 'L' || o[0] === 'A').length, cls).toBeLessThanOrEqual(MAX_VERTICES)
    }
  })
})
