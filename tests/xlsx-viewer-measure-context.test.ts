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
