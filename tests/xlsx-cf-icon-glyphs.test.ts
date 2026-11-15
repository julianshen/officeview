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
import { describe, expect, it } from 'vitest'
import {
  CF_ICON_GLYPH_CLASSES, CF_ICON_GLYPH_WITHHELD, admitIconGlyphs, cfIconGlyphClass, drawIconGlyph, glyphBytes, iconGlyph, type CfIconGlyph,
} from '../src/xlsx/cf/iconGlyphs'
import { CF_ICON_PAINT_CAPABILITIES, CF_ICON_RENDER_MODEL, admitBatchGlyphs, layoutIcon, paintIconCells } from '../src/xlsx/cf/iconPaint'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { paintMemoryFor } from '../src/xlsx/cf/render'
import type { CfIconSetName } from '../src/xlsx/cf/model'
import { loadIconGlyphKey, type IconGlyphKey } from '../scripts/cf-refs/verify'
import { S4_ICON_FAMILIES } from '../scripts/cf-refs/cases'
import { createCanvas } from 'canvas'
import { AT_CAPTURE, KEY_CELL_H, KEY_CELL_W, KEY_FONT, checkGlyph, cropInkBox, inkBoxOf, keyHost } from './helpers/cf-glyph-measure'
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
        const g = CF_ICON_GLYPH_WITHHELD.has(cls) ? cfIconGlyphClass(cls) : iconGlyph(...split(pair))
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
    const cache = admitIconGlyphs(all, memory, budget, cfIconGlyphClass)!
    expect(cache).toBeDefined()
    const expected = [...all].reduce((n, c) => n + glyphBytes(cfIconGlyphClass(c)!), 0)
    expect(cache.bytes).toBe(expected)
    expect(memory.currentBytes - baseline).toBe(expected)
    // Decoding spends work units (one per op, plus one per class), never wall time.
    expect(budget.units).toBe([...all].reduce((n, c) => n + cfIconGlyphClass(c)!.path.length + 1, 0))

    // Drawing an admitted glyph at each zoom/DPR allocates nothing: the ledger is unchanged.
    const ctx = createCanvas(64, 64).getContext('2d') as unknown as CanvasRenderingContext2D
    for (const [zoom, dpr] of [[0.5, 1], [0.75, 1], [1, 1], [1, 2], [1.5, 2], [2, 3]] as const) {
      for (const cls of all) expect(drawIconGlyph(ctx, cfIconGlyphClass(cls)!, { x: 1, y: 1, w: 13 * zoom * dpr, h: 12 * zoom * dpr }, { zoom, dpr }, cache)).toBe(true)
      expect(memory.currentBytes - baseline).toBe(expected)
    }
    cache.release()
    expect(memory.currentBytes).toBe(baseline)

    // A refused charge returns false/undefined, draws nothing and leaves the ledger at baseline.
    const tiny = createCfMemory({ limitBytes: 16 })
    expect(admitIconGlyphs(all, tiny, createCfWorkBudget({ now: () => 0 }), cfIconGlyphClass)).toBeUndefined()
    expect(tiny.currentBytes).toBe(0)
    const refusing = { spend: () => 'refused' as const, resume() {}, units: 0 }
    expect(admitIconGlyphs(all, memory, refusing, cfIconGlyphClass)).toBeUndefined()
    expect(memory.currentBytes).toBe(baseline)
    const cls = [...all][0]!
    const other = admitIconGlyphs(new Set([[...all][1]!]), memory, createCfWorkBudget({ now: () => 0 }), cfIconGlyphClass)!
    const blank = createCanvas(20, 20)
    const bctx = blank.getContext('2d') as unknown as CanvasRenderingContext2D
    expect(drawIconGlyph(bctx, cfIconGlyphClass(cls)!, { x: 0, y: 0, w: 13, h: 12 }, AT_CAPTURE, other)).toBe(false)
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
 * module (encoded table, class map and decoder) renders 12,682 B, 3,740 B gzip, inside the lazy cf chunk. */
const CF_GLYPH_TABLE_GZIP_BYTES = 3740

describe('Task 28: the icon frame at the glyph key geometry', () => {
  const icon = { ruleId: 1, set: '3Arrows' as const, iconId: 0, bucket: 0, buckets: 3, showValue: true }
  const ALL = (): boolean => true
  it('production lays the icon out in the key frame: 13 × 12 px, 2 px in from the left gridline, 1 px above the bottom one', () => {
    const laid = layoutIcon(icon, keyHost(), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)
    expect(laid?.box).toEqual({ x: 2, y: 7, w: 13, h: 12 })
    // The viewer's whole-px row (20 pt -> 27 px) is the same row.
    expect(layoutIcon(icon, keyHost({ height: 27 }), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)).toBeDefined()
    // The column width does not move it (stage4-icon-geometry-2 I5/I6/J5/J6, 180 pt columns: ink 2..12 × 7..18).
    expect(layoutIcon(icon, keyHost({ width: 240 }), KEY_FONT, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)?.box).toEqual({ x: 2, y: 7, w: 13, h: 12 })
  })

  it('with its evidence, a key-geometry host paints the original glyph where Excel inks it (left-aligned text, as I5/J5)', () => {
    const caps = [...glyphCaps(), ...CF_ICON_PAINT_CAPABILITIES, 'icon:inset-left']
    withGlyphEvidence(() => {
      const memory = createCfMemory({})
      for (const pair of ['3Arrows:0', '3TrafficLights1:2', '5Boxes:2']) {
        const [set, iconId] = split(pair)
        const desc = { ruleId: 1, set, iconId, bucket: iconId, buckets: Number(set[0]), showValue: true }
        const admitted = admitBatchGlyphs([desc], memory, () => createCfWorkBudget({ now: () => 0 }))
        const canvas = createCanvas(KEY_CELL_W, KEY_CELL_H)
        const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, KEY_CELL_W, KEY_CELL_H)
        const host = keyHost({ textWidth: 10, cell: { value: 'ab', style: { numFmtId: 0, fontSizePt: 11, fontName: 'Calibri', horizontal: 'left' } } })
        const layouts: boolean[] = []
        paintIconCells(ctx, new Map([[0, desc]]), () => desc, () => host, { r0: 0, c0: 0, r1: 0, c1: 0 }, AT_CAPTURE,
          { admitted, hasBar: () => false, check: () => true }, (_r, _c, withheld) => layouts.push(withheld))
        expect(layouts, pair).toEqual([false])
        expect(inkBoxOf(ctx.getImageData(0, 0, KEY_CELL_W, KEY_CELL_H).data, KEY_CELL_W, KEY_CELL_H), pair).toEqual(checks.get(SIDECAR.glyphClasses[pair]!)!.drawn)
        admitted.release()
      }
      expect(memory.currentBytes).toBe(0)
    }, caps)
  })

  it('every input the key fixes must equal the key: other font sizes, rows, alignments and faces withhold', () => {
    const at = (host: ReturnType<typeof keyHost>, font: { sizePt?: number; vertical?: 'top' | 'center' | 'bottom' } = {}) =>
      layoutIcon(icon, host, { ...KEY_FONT, ...font }, AT_CAPTURE, CF_ICON_RENDER_MODEL, ALL)
    expect(at(keyHost())).toBeDefined()
    for (const sizePt of [8, 10, 10.5, 12, 20]) expect(at(keyHost(), { sizePt }), `${sizePt} pt`).toBeUndefined()
    for (const rowPt of [15, 19, 21, 30, 60]) expect(at(keyHost({ height: rowPt * 4 / 3 })), `${rowPt} pt row`).toBeUndefined()
    for (const vertical of ['top', 'center'] as const) expect(at(keyHost(), { vertical }), vertical).toBeUndefined()
    const style = { numFmtId: 0, fontSizePt: 11, fontName: 'Calibri' }
    for (const s of [{ ...style, fontName: 'Arial' }, { ...style, fontName: undefined }, { ...style, bold: true }, { ...style, italic: true }]) {
      expect(at(keyHost({ cell: { value: 10, style: s } })), JSON.stringify(s)).toBeUndefined()
    }
  })
})
