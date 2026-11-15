/**
 * Icon-set fixtures for the Task 29 painting tests. Nothing Stage 4 is published and no glyph table exists (Ruling
 * S4-ORDER), so production withholds every icon. These tests paint under INJECTED stubs only, all three of them:
 * - evidence: EXACT (production's Stage 1/2 sets plus the listed stubs, so the published Stage 3 set never leaks);
 * - a render model (`injectCfIconRenderModel`): hypothetical pt values, never Excel facts;
 * - a glyph table (`injectCfIconGlyphs`): one plain square per `(set, iconId)`, out-of-family ids included, so a
 *   withheld icon is the predicate's doing and not a hole in the stub. No artwork, no icon pixels.
 */
import { VERIFIED_STAGE1_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, injectCfBarLayout, withExactCfEvidence } from '../../src/xlsx/cf/evidence'
import { injectCfIconGlyphs, type CfIconGlyph } from '../../src/xlsx/cf/iconGlyphs'
import { CF_ICON_PAINT_CAPABILITIES, injectCfIconRenderModel, type CfIconRenderModel } from '../../src/xlsx/cf/iconPaint'
import type { CfDefinition, CfIconSetName, CfRule } from '../../src/xlsx/cf/model'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../../src/xlsx/types'
import { BAR_EVAL_EVIDENCE, BAR_LAYOUT, BAR_PAINT_EVIDENCE, BAR_PARSE_EVIDENCE } from './cf-bars'

/** Every evaluation, scope and composition capability these icons can need (stubs only). */
export const ICON_EVAL_EVIDENCE = [
  'icon:families-base', 'icon:families-x14', 'icon:thresholds-num', 'icon:thresholds-percent', 'icon:thresholds-gte-true', 'icon:thresholds-gte-false',
  'icon:thresholds-mixed', 'icon:reverse', 'icon:custom-mapping', 'icon:custom-noicons', 'icon:participation-types', 'icon:participation-dateserial',
  'icon:composition-with-bar', 'icon:composition-with-scale', 'icon:scope-whole-column',
]
/** The paint gates: the model's box, vertical alignment and snapping, and each alignment's inset. */
export const ICON_PAINT_EVIDENCE = [...CF_ICON_PAINT_CAPABILITIES, 'icon:inset-left', 'icon:inset-right-aligned', 'icon:inset-center']
export const ICON_SHOWVALUE_EVIDENCE = 'icon:showvalue-canvas-suppression'

/** A hypothetical model: a 9 pt box centred in the row, 2 pt in from the left gridline, the text 2 pt after it (13 pt
 * for every alignment). */
export const STUB_MODEL: CfIconRenderModel = Object.freeze({
  box: { size: 'fixed' as const, vertical: 'center' as const }, sizePt: 9, offsetsPt: { left: 2, top: 1, bottom: 1 }, gapPt: 2, alignedInsetPt: { right: 13, center: 13 },
})
/** 13 pt in logical px. */
export const STUB_INSET_PX = 13 * 4 / 3
/** Each stub glyph is one square filled with its icon id's colour. */
export const ICON_COLOURS = ['#11aa11', '#cc8800', '#cc0011', '#1188cc', '#8811cc'] as const
const ARGB = [0xff11aa11, 0xffcc8800, 0xffcc0011, 0xff1188cc, 0xff8811cc]
export function stubGlyph(set: CfIconSetName, iconId: number): CfIconGlyph {
  return { glyphClass: `${set}:${iconId}`, widthPt: 9, heightPt: 9, source: 'original', path: [['M', 0, 0], ['L', 9, 0], ['L', 9, 9], ['L', 0, 9], ['Z'], ['F', ARGB[((iconId % 5) + 5) % 5]!]] }
}

export type IconOptions = { paint?: boolean; extra?: readonly string[]; omit?: readonly string[]; glyphs?: false | ((set: CfIconSetName, iconId: number) => CfIconGlyph | undefined); model?: false | CfIconRenderModel }
/** Runs `fn` with exactly the Stage 1/2 production evidence plus the icon (and bar) stubs, the stub model and the stub
 * glyph table; all are restored after. */
export async function withIcons<T>(fn: () => Promise<T> | T, options: IconOptions = {}): Promise<T> {
  injectCfBarLayout(BAR_LAYOUT)
  injectCfIconRenderModel(options.model === false ? null : options.model ?? STUB_MODEL)
  injectCfIconGlyphs(options.glyphs === false ? null : options.glyphs ?? stubGlyph)
  const caps = [...ICON_EVAL_EVIDENCE, ...BAR_EVAL_EVIDENCE, ...BAR_PARSE_EVIDENCE, ...BAR_PAINT_EVIDENCE, ...(options.paint === false ? [] : ICON_PAINT_EVIDENCE), ...(options.extra ?? [])]
    .filter(c => !options.omit?.includes(c))
  try { return await withExactCfEvidence([...VERIFIED_STAGE1_CAPABILITIES, ...VERIFIED_STAGE2_CAPABILITIES, ...caps], async () => fn()) }
  finally { injectCfBarLayout(null); injectCfIconRenderModel(null); injectCfIconGlyphs(null) }
}

type IconDef = Extract<CfDefinition, { kind: 'iconSet' }>
const num = (value: number) => ({ bound: { type: 'num' as const, value }, gte: true })
/** `3Arrows` over num 0 / 40 / 70: 10 → bucket 0, 50 → bucket 1, 90 → bucket 2 (no statistics). */
export function iconDef(overrides: Partial<IconDef> = {}): IconDef {
  return { kind: 'iconSet', dialect: 'base', formulas: [], set: '3Arrows', thresholds: [num(0), num(40), num(70)], reverse: false, showValue: true, percentAttr: true, sheetRightToLeft: false, ...overrides }
}
export function iconRule(range: { firstRow: number; lastRow: number; firstCol: number; lastCol: number }, overrides: Partial<IconDef> = {}, id = 1, priority = id): CfRule {
  return { id, priority, documentOrder: priority - 1, ranges: [range], stopIfTrue: false, writes: new Set(['iconSet']), definition: iconDef(overrides) }
}

/** A one-column icon sheet: `values` down column `col` from row 0 (a `widthChars` column; 200/7 is 200 px), one icon
 * rule over exactly those cells, each cell styled `style`. */
export function iconSheet(values: readonly (number | string)[], options: { col?: number; widthChars?: number; def?: Partial<IconDef>; name?: string; style?: XlsxCellStyle; firstRow?: number; rows?: number; extraRules?: readonly CfRule[] } = {}): XlsxSheet {
  const col = options.col ?? 0, first = options.firstRow ?? 0
  const cells = values.map((value, i): XlsxCell => ({ ref: `${String.fromCharCode(65 + col)}${first + i + 1}`, row: first + i, col, value, styleIndex: 0, hasCachedValue: true, ...(options.style ? { style: options.style } : {}) }))
  const rows = Math.max(options.rows ?? 0, first + values.length)
  return {
    name: options.name ?? 'Icons', sheetId: options.name ?? 'icons',
    rows: Array.from({ length: rows }, (_, r) => ({ index: r, cells: cells.filter(c => c.row === r) })),
    cols: [{ min: col, max: col, widthChars: options.widthChars ?? 200 / 7 }],
    merges: [], mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules: [iconRule({ firstRow: first, lastRow: first + values.length - 1, firstCol: col, lastCol: col }, options.def), ...(options.extraRules ?? [])] },
  } as XlsxSheet
}

/** Records every fill (rects, text, and each icon path's device bounding box) on a real canvas. */
export function iconRecorder(width: number, height: number, createCanvas: (w: number, h: number) => { getContext(kind: '2d'): unknown }) {
  const canvas = createCanvas(Math.max(1, Math.ceil(width)), Math.max(1, Math.ceil(height)))
  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D
  type Op = { op: 'rect' | 'text' | 'icon'; style: string; x: number; y: number; w: number; h: number; text?: string }
  const ops: Op[] = []
  let points: Array<[number, number]> = []
  const dev = (x: number, y: number): [number, number] => { const t = ctx.getTransform(); return [t.a * x + t.c * y + t.e, t.b * x + t.d * y + t.f] }
  const fillRect = ctx.fillRect.bind(ctx), fillText = ctx.fillText.bind(ctx), moveTo = ctx.moveTo.bind(ctx), lineTo = ctx.lineTo.bind(ctx), fill = ctx.fill.bind(ctx), beginPath = ctx.beginPath.bind(ctx)
  ctx.fillRect = (x: number, y: number, w: number, h: number) => {
    const t = ctx.getTransform()
    ops.push({ op: 'rect', style: String(ctx.fillStyle).toLowerCase(), x: t.a * x + t.e, y: t.d * y + t.f, w: t.a * w, h: t.d * h })
    fillRect(x, y, w, h)
  }
  ctx.fillText = (text: string, x: number, y: number) => {
    const t = ctx.getTransform()
    ops.push({ op: 'text', style: String(ctx.fillStyle).toLowerCase(), x: t.a * x + t.e, y: t.d * y + t.f, w: 0, h: 0, text })
    fillText(text, x, y)
  }
  ctx.beginPath = () => { points = []; beginPath() }
  ctx.moveTo = (x: number, y: number) => { points.push(dev(x, y)); moveTo(x, y) }
  ctx.lineTo = (x: number, y: number) => { points.push(dev(x, y)); lineTo(x, y) }
  ;(ctx as { fill: (...args: unknown[]) => void }).fill = (...args: unknown[]) => {
    if (points.length) {
      const xs = points.map(p => p[0]), ys = points.map(p => p[1])
      const x = Math.min(...xs), y = Math.min(...ys)
      ops.push({ op: 'icon', style: String(ctx.fillStyle).toLowerCase(), x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y })
    }
    ;(fill as (...a: unknown[]) => void)(...args)
  }
  return { ctx, ops, canvas }
}
