import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { paintCellLayout } from '../src/xlsx/render'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { buildCellRecords, LayoutCache, type CellBodyLayout, type CellLayoutHost, type PlacedCell } from '../src/xlsx/view/textLayout'
import type { CacheIdentity } from '../src/xlsx/view/contracts'

const owner: CacheIdentity = { docId: 'doc', sheetId: 'sheet', valueGen: 0, fontGen: 0, viewGen: 0 }
const placed: PlacedCell = {
  cell: { ref: 'A1', row: 0, col: 0, value: 'text reaching into empty neighbours', styleIndex: 0 },
  source: { row: 0, col: 0 }, x: 0, y: 0, width: 64, height: 24,
}
function host() {
  return createCellLayoutHost(createCanvas(400, 40).getContext('2d') as unknown as CanvasRenderingContext2D)
}
function pixels(body: CellBodyLayout) {
  const canvas = createCanvas(800, 40)
  const ctx = canvas.getContext('2d')
  ctx.translate(200, 0) // Preserve ink to the left of centered/right-aligned cells.
  paintCellLayout(ctx as unknown as CanvasRenderingContext2D, body, { width: 400, height: 40 })
  return ctx.getImageData(0, 0, 800, 40).data
}

describe('slice9: layout-cache allowance churn', () => {
  test('canonical fitting text reuses one body as tile caps change, without aliases or mutation', () => {
    const canonical = host()
    let layouts = 0
    const original = canonical.layoutBody
    canonical.layoutBody = (...args) => { layouts++; return original(...args) }
    const cache = new LayoutCache()
    const first = cache.layout(owner, { ...placed, overflow: { left: 0, right: 400 } }, canonical)
    expect(first.text).toMatchObject({ kind: 'plain' })
    expect(first.text && 'clipX' in first.text).toBe(false)
    Object.freeze(first.text)
    Object.freeze(first)
    const retainedBytes = cache.stats.bytes
    for (let cap = 401; cap < 501; cap++) {
      // A tile-boundary cap changes although all of this body's text still fits.
      const next = cache.layout(owner, { ...placed, overflow: { left: cap % 17, right: cap } }, canonical)
      expect(next).toBe(first)
    }
    expect(layouts).toBe(1)
    expect(cache.stats).toMatchObject({ built: 1, reused: 100, entries: 1, bytes: retainedBytes })
  })

  test('a nearer blocker keeps its real clip, pixels and text-record bounds', () => {
    const canonical = host()
    const cache = new LayoutCache()
    const fit = cache.layout(owner, { ...placed, overflow: { left: 0, right: 400 } }, canonical)
    const blocked = { ...placed, overflow: { left: 0, right: 0 } }
    const clipped = cache.layout(owner, blocked, canonical)
    const direct = canonical.layoutBody(placed.cell, placed.width, placed.height, blocked.overflow)
    expect(clipped).not.toBe(fit)
    expect(clipped).toEqual(direct)
    expect(pixels(clipped)).toEqual(pixels(direct))
    expect(pixels(clipped)).not.toEqual(pixels(fit))
    const record = buildCellRecords([blocked], owner, canonical, cache)[0]
    expect(record.text[0].clip).toEqual({ x: 3, y: 0, width: 61, height: 24 })
    expect(cache.layout(owner, { ...placed, overflow: { left: 1, right: 450 } }, canonical)).toBe(fit)
    expect(cache.stats.built).toBe(2)
  })

  test('generic hosts retain allowance-dependent formatting even when text fits', () => {
    let layouts = 0
    const generic: CellLayoutHost = { layoutBody: (_cell, w, h, overflow) => {
      layouts++
      const text = String(overflow?.right)
      return { x: 0, y: 0, w, h, formattedText: text, text: { kind: 'plain', text, font: '10px sans-serif', color: '#000', tx: 3, ty: 15, width: 8 } }
    } }
    const cache = new LayoutCache()
    const a = cache.layout(owner, { ...placed, overflow: { left: 0, right: 100 } }, generic)
    const b = cache.layout(owner, { ...placed, overflow: { left: 0, right: 200 } }, generic)
    expect(a.formattedText).toBe('100')
    expect(b.formattedText).toBe('200')
    expect(cache.layout(owner, { ...placed, overflow: { left: 0, right: 100 } }, generic)).toBe(a)
    expect(layouts).toBe(2)
  })

  test('replacing an opted-in shared body charges only the retained entry', () => {
    const specialized: CellLayoutHost = {
      layoutBody: (_cell, w, h, overflow) => ({ x: 0, y: 0, w, h, formattedText: String(overflow?.right) }),
      canReuseOverflow: (body, overflow) => body.formattedText === String(overflow?.right),
    }
    const cache = new LayoutCache()
    cache.layout(owner, { ...placed, overflow: { left: 0, right: 100 } }, specialized)
    const bytes = cache.stats.bytes
    const changed = cache.layout(owner, { ...placed, overflow: { left: 0, right: 200 } }, specialized)
    expect(changed.formattedText).toBe('200')
    expect(cache.stats.entries).toBe(1)
    expect(cache.stats.bytes).toBe(bytes)
  })

  test('a freshly certified replacement evicts a colder cell under byte pressure', () => {
    let builds = 0
    const valueFor = (overflow: PlacedCell['overflow']) => String(overflow?.right).repeat(overflow?.right === 200 ? 100 : 1)
    const specialized: CellLayoutHost = {
      layoutBody: (_cell, w, h, overflow) => {
        builds++
        return { x: 0, y: 0, w, h, formattedText: valueFor(overflow) }
      },
      canReuseOverflow: (body, overflow) => body.formattedText === valueFor(overflow),
    }
    const item = (col: number, right: number): PlacedCell => ({
      ...placed, cell: { ...placed.cell, col }, source: { row: 0, col }, overflow: { left: 0, right },
    })
    const cache = new LayoutCache(10, 1400)
    cache.layout(owner, item(0, 100), specialized)
    cache.layout(owner, item(1, 100), specialized)
    const replacement = cache.layout(owner, item(0, 200), specialized)
    expect(cache.stats.entries).toBe(1)
    expect(cache.stats.bytes).toBeLessThanOrEqual(cache.maxBytes)
    expect(cache.layout(owner, item(0, 200), specialized)).toBe(replacement)
    expect(builds).toBe(3)
  })

  test.each(['fits', null])('a canonical body %j fitting its own box also reuses a zero cap', value => {
    const canonical = host()
    const cache = new LayoutCache()
    const item = { ...placed, cell: { ...placed.cell, value } }
    const first = cache.layout(owner, item, canonical)
    expect(cache.layout(owner, { ...item, overflow: { left: 0, right: 0 } }, canonical)).toBe(first)
    expect(cache.stats).toMatchObject({ built: 1, reused: 1, entries: 1 })
  })

  test('a left cap cutting centered error text cannot reuse an unclipped body', () => {
    const canonical = host()
    const cache = new LayoutCache()
    const item = { ...placed, cell: { ...placed.cell, valueIsError: true } }
    const first = cache.layout(owner, { ...item, overflow: { left: 400, right: 400 } }, canonical)
    const blocked = { ...item, overflow: { left: 0, right: 400 } }
    const clipped = cache.layout(owner, blocked, canonical)
    expect(clipped).not.toBe(first)
    expect(clipped).toEqual(canonical.layoutBody(item.cell, item.width, item.height, blocked.overflow))
    expect(pixels(clipped)).not.toEqual(pixels(first))
  })

  test('complex text retains exact clips when an absent allowance and a roomy cap differ', () => {
    const canonical = host()
    const cache = new LayoutCache()
    const complex = { ...placed, cell: { ...placed.cell, style: { numFmtId: 0, horizontal: 'left' as const } } }
    const unextended = cache.layout(owner, complex, canonical)
    const extended = cache.layout(owner, { ...complex, overflow: { left: 0, right: 400 } }, canonical)
    expect(unextended.text?.kind).toBe('layout')
    expect(extended).not.toBe(unextended)
    expect(extended).toEqual(canonical.layoutBody(complex.cell, complex.width, complex.height, { left: 0, right: 400 }))
    expect(pixels(extended)).not.toEqual(pixels(unextended))
  })
})
