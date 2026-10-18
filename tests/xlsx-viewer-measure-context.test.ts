/**
 * Repaint memory admitted per record (spec §12.4): the canonical layout's
 * measurement runs on ONE shared per-viewer measurement context (plus the
 * local identity context core/text-metrics derives from it), admitted once
 * before creation; every distinct (font, text) either context shapes is
 * admitted with the shapeCacheBytes model before it is measured.
 */
import { describe, expect, it, vi } from 'vitest'
import { SharedMeasureContext } from '../src/components/xlsx/measureContext'
import { LayoutAdmissionError, type LayoutAdmission } from '../src/xlsx/view/textLayout'
import { TILE_CANVAS_BYTES, shapeCacheBytes } from '../src/xlsx/view/mainLedger'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'

const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
function gate(limit = Infinity) {
  let held = 0, measured = 0
  const api: LayoutAdmission & { total: () => number; measured: () => number } = {
    admit: bytes => { if (held + measured + bytes > limit) return false; held += bytes; return true },
    release: bytes => { held -= bytes }, hold: bytes => { held += bytes }, cacheChanged: () => {},
    measureChanged: delta => { measured += delta },
    total: () => held + measured, measured: () => measured,
  }
  return api
}
const wrapped = (text: string): XlsxCell => ({ ref: 'A1', row: 0, col: 0, value: text, styleIndex: 0, style: { numFmtId: 0, wrapText: true } })
const distinctWords = (n: number) => Array.from({ length: n }, (_, i) => `w${i.toString(36)}x`).join(' ')

describe('shared admitted measurement context', () => {
  it('admits both contexts before creating any canvas, and every distinct measured text before it is shaped', () => {
    const admission = gate()
    const created = vi.spyOn(document, 'createElement')
    const context = new SharedMeasureContext(() => admission)
    const host = context.host(sheet)
    expect(created).not.toHaveBeenCalled()
    const proto = Object.getPrototypeOf(document.createElement('canvas').getContext('2d')!) as CanvasRenderingContext2D
    created.mockClear()
    const seen = new Map<string, Set<string>>()
    const original = proto.measureText
    const spy = vi.spyOn(proto, 'measureText').mockImplementation(function (this: CanvasRenderingContext2D, text: string) {
      let texts = seen.get(this.font); if (!texts) seen.set(this.font, texts = new Set()); texts.add(text)
      return original.call(this, text)
    })
    try { host.layoutBody(wrapped(distinctWords(200)), 64, 2000) } finally { spy.mockRestore() }
    let required = 2 * TILE_CANVAS_BYTES
    for (const texts of seen.values()) required += shapeCacheBytes(texts.size, [...texts].reduce((n, t) => n + t.length, 0))
    expect(admission.measured()).toBeGreaterThanOrEqual(required)
    // One shared pair of contexts serves every later layout (no per-tile context).
    const before = created.mock.calls.length
    host.layoutBody(wrapped(distinctWords(10)), 64, 200)
    context.host(sheet).layoutBody(wrapped('another text'), 64, 200)
    expect(created.mock.calls.length).toBe(before)
    created.mockRestore()
  })

  it('the measurement path alone exceeding the limit refuses before shaping (LayoutAdmissionError, coarse fallback upstream)', () => {
    const admission = gate(2 * TILE_CANVAS_BYTES + 64 * 1024)
    const host = new SharedMeasureContext(() => admission).host(sheet)
    expect(() => host.layoutBody(wrapped(distinctWords(3000)), 64, 100_000)).toThrow(LayoutAdmissionError)
    expect(admission.total()).toBeLessThanOrEqual(2 * TILE_CANVAS_BYTES + 64 * 1024)
  })

  it('trim drops oversized contexts and returns their charge', () => {
    const admission = gate()
    const context = new SharedMeasureContext(() => admission, 16 * 1024)
    context.host(sheet).layoutBody(wrapped(distinctWords(400)), 64, 10_000)
    expect(admission.measured()).toBeGreaterThan(16 * 1024)
    context.trim()
    expect(admission.measured()).toBe(0)
  })
})

describe('confirmation r8', () => {
  it('F1 trim drops every host built on the retired pair before returning its charge (the old context is collectable)', async () => {
    const admission = gate()
    const context = new SharedMeasureContext(() => admission, 16 * 1024)
    const host = context.host(sheet)
    host.layoutBody(wrapped(distinctWords(400)), 64, 10_000)
    const internals = context as unknown as { ctx: object | null; hosts: WeakMap<object, unknown> }
    const retired = new WeakRef(internals.ctx!)
    let hostsBefore: unknown = internals.hosts
    context.trim()
    expect(admission.measured()).toBe(0)
    // No host references the retired pair: the host map was replaced and the
    // handed-out host captures no context.
    expect(internals.hosts).not.toBe(hostsBefore)
    hostsBefore = undefined
    expect(internals.ctx).toBeNull()
    const gc = (globalThis as { gc?: () => void }).gc
    // Under --expose-gc (vitest execArgv) the retired pair must be collectable.
    if (gc) { await new Promise(resolve => setTimeout(resolve, 0)); gc(); expect(retired.deref()).toBeUndefined() }
    if (process.env.OFFICEVIEW_EXPECT_GC) expect(typeof gc).toBe('function')
    // The next layout admits a fresh pair.
    host.layoutBody(wrapped('fresh'), 64, 200)
    expect(admission.measured()).toBeGreaterThanOrEqual(2 * TILE_CANVAS_BYTES)
  })

  it('F3 a font publication rebuilds hosts and their measured-width caches', () => {
    // The newly available font is simulated by doubling native widths.
    const proto = Object.getPrototypeOf(document.createElement('canvas').getContext('2d')!) as CanvasRenderingContext2D
    const original = proto.measureText
    let factor = 1
    const spy = vi.spyOn(proto, 'measureText').mockImplementation(function (this: CanvasRenderingContext2D, text: string) {
      const m = original.call(this, text)
      return { width: m.width * factor, actualBoundingBoxAscent: m.actualBoundingBoxAscent, actualBoundingBoxDescent: m.actualBoundingBoxDescent, fontBoundingBoxAscent: m.fontBoundingBoxAscent, fontBoundingBoxDescent: m.fontBoundingBoxDescent } as TextMetrics
    })
    const admission = gate()
    const context = new SharedMeasureContext(() => admission)
    context.setFontGen(1)
    const host = context.host(sheet)
    const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: 'fontprobe', styleIndex: 0 }
    const first = host.layoutBody(cell, 400, 20)
    try {
      factor = 2
      context.setFontGen(2)
      const second = host.layoutBody(cell, 400, 20)
      const width = (body: typeof first) => (body.text as { width: number }).width
      expect(width(second)).toBeCloseTo(2 * width(first), 5)
    } finally { spy.mockRestore() }
  })

  it('F4 no global state: the 2D prototype is never patched, even when a layout re-enters another context', () => {
    const proto = Object.getPrototypeOf(document.createElement('canvas').getContext('2d')!) as CanvasRenderingContext2D
    const nativeMeasure = proto.measureText
    let inner: SharedMeasureContext | undefined
    const seenDuring: unknown[] = []
    const outerGate = gate()
    const innerGate = gate()
    outerGate.measureChanged = delta => {
      seenDuring.push(proto.measureText)
      if (!inner) { inner = new SharedMeasureContext(() => innerGate); inner.host(sheet).layoutBody(wrapped('inner first'), 64, 200) }
      ;(outerGate as unknown as { _m: number })._m = ((outerGate as unknown as { _m?: number })._m ?? 0) + delta
    }
    const outer = new SharedMeasureContext(() => outerGate)
    outer.host(sheet).layoutBody(wrapped(distinctWords(20)), 64, 2000)
    const outerCharge = outer.bytes
    inner!.host(sheet).layoutBody(wrapped(distinctWords(40)), 64, 2000)
    expect(outer.bytes).toBe(outerCharge)
    expect(proto.measureText).toBe(nativeMeasure)
    expect(seenDuring.every(fn => fn === nativeMeasure)).toBe(true)
  })
})
