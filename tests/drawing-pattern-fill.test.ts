import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseFillDefinition, resolveFill, resolveDrawingStyle, type DrawingStyle } from '../src/drawing/style'
import { paintGeometry } from '../src/drawing/paint'
import { resolvePreset } from '../src/drawing/geometry'
import {
  clearPatternTileCache,
  patternTileCacheSize,
  hasPatternTile,
  getPatternTile,
  paintPatternTile,
  SUPPORTED_PATTERN_PRESETS,
  type PatternPreset,
  getPatternBitmask,
} from '../src/drawing/pattern'

describe('shape pattern fill (<a:pattFill>)', () => {
  test('parses supported pattern preset, fgClr, and bgClr into DrawingFill', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:pattFill prst="dkUpDiag">` +
      `    <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>` +
      `    <a:bgClr><a:srgbClr val="00FF00"/></a:bgClr>` +
      `  </a:pattFill>` +
      `</a:spPr>`
    )
    const fillDef = parseFillDefinition(xml)
    expect(fillDef).toBeDefined()
    expect(fillDef?.kind).toBe('pattern')
    if (fillDef?.kind === 'pattern') {
      expect(fillDef.preset).toBe('dkUpDiag')
      expect(fillDef.fgColor?.kind).toBe('srgb')
      expect(fillDef.bgColor?.kind).toBe('srgb')
    }

    const issues: any[] = []
    const resolved = resolveFill(fillDef, undefined, undefined, issues)
    expect(resolved).toEqual({
      kind: 'pattern',
      preset: 'dkUpDiag',
      fgColor: { r: 255, g: 0, b: 0, a: 1 },
      bgColor: { r: 0, g: 255, b: 0, a: 1 },
    })
    expect(issues).toEqual([])
  })

  test('omitted bgClr defaults to opaque white per ECMA-376 spec', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:pattFill prst="ltUpDiag">` +
      `    <a:fgClr><a:srgbClr val="0000FF"/></a:fgClr>` +
      `  </a:pattFill>` +
      `</a:spPr>`
    )
    const fillDef = parseFillDefinition(xml)
    const issues: any[] = []
    const resolved = resolveFill(fillDef, undefined, undefined, issues)
    expect(resolved).toEqual({
      kind: 'pattern',
      preset: 'ltUpDiag',
      fgColor: { r: 0, g: 0, b: 255, a: 1 },
      bgColor: { r: 255, g: 255, b: 255, a: 1 },
    })
    expect(issues).toEqual([])
  })

  test('omitted fgClr defaults to black per ECMA-376 spec', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:pattFill prst="ltDnDiag">` +
      `    <a:bgClr><a:srgbClr val="FF0000"/></a:bgClr>` +
      `  </a:pattFill>` +
      `</a:spPr>`
    )
    const fillDef = parseFillDefinition(xml)
    const issues: any[] = []
    const resolved = resolveFill(fillDef, undefined, undefined, issues)
    expect(resolved).toEqual({
      kind: 'pattern',
      preset: 'ltDnDiag',
      fgColor: { r: 0, g: 0, b: 0, a: 1 },
      bgColor: { r: 255, g: 0, b: 0, a: 1 },
    })
    expect(issues).toEqual([])
  })

  test('missing prst attribute diagnoses invalid-fill rather than deferred preset', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:pattFill>` +
      `    <a:fgClr><a:srgbClr val="0000FF"/></a:fgClr>` +
      `  </a:pattFill>` +
      `</a:spPr>`
    )
    const fillDef = parseFillDefinition(xml)
    const issues: any[] = []
    const resolved = resolveFill(fillDef, undefined, undefined, issues)
    expect(issues.some(i => i.kind === 'invalid-fill' && i.feature === 'pattFill')).toBe(true)
    expect(resolved).toEqual({
      kind: 'solid',
      color: { r: 0, g: 0, b: 255, a: 1 },
    })
  })

  test('unsupported shape pattern preset diagnoses unsupported-fill and falls back to fg solid', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:pattFill prst="divot">` +
      `    <a:fgClr><a:srgbClr val="0000FF"/></a:fgClr>` +
      `    <a:bgClr><a:srgbClr val="FFFF00"/></a:bgClr>` +
      `  </a:pattFill>` +
      `</a:spPr>`
    )
    const fillDef = parseFillDefinition(xml)
    const issues: any[] = []
    const resolved = resolveFill(fillDef, undefined, undefined, issues)
    expect(issues.some(i => i.kind === 'unsupported-fill' && i.feature === 'divot')).toBe(true)
    expect(resolved).toEqual({
      kind: 'solid',
      color: { r: 0, g: 0, b: 255, a: 1 },
    })
  })

  test('paint tiles supported shape presets using pattern tile onto canvas', () => {
    const canvas = createCanvas(32, 32)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('rect', 32, 32)
    const style: DrawingStyle = {
      fill: {
        kind: 'pattern',
        preset: 'dkUpDiag',
        fgColor: { r: 255, g: 0, b: 0, a: 1 },
        bgColor: { r: 0, g: 255, b: 0, a: 1 },
      },
      issues: [],
    }
    const issues = paintGeometry(ctx as never, geom, style, 32, 32)
    expect(issues).toEqual([])

    const data = ctx.getImageData(0, 0, 32, 32).data
    let redPixels = 0
    let greenPixels = 0
    for (let i = 0; i < data.length; i += 4) {
      const r = data[i], g = data[i + 1], b = data[i + 2], a = data[i + 3]
      if (a > 200 && r > 200 && g < 50 && b < 50) redPixels++
      if (a > 200 && r < 50 && g > 200 && b < 50) greenPixels++
    }
    expect(redPixels).toBeGreaterThan(10)
    expect(greenPixels).toBeGreaterThan(10)
  })

  test('all supported pattern presets parse and render cleanly', () => {
    for (const preset of SUPPORTED_PATTERN_PRESETS) {
      const canvas = createCanvas(16, 16)
      const ctx = canvas.getContext('2d')
      const geom = resolvePreset('rect', 16, 16)
      const style: DrawingStyle = {
        fill: {
          kind: 'pattern',
          preset: preset as PatternPreset,
          fgColor: { r: 255, g: 0, b: 0, a: 1 },
          bgColor: { r: 0, g: 0, b: 255, a: 1 },
        },
        issues: [],
      }
      const issues = paintGeometry(ctx as never, geom, style, 16, 16)
      expect(issues).toEqual([])
      const data = ctx.getImageData(0, 0, 16, 16).data
      let inked = 0
      for (let i = 0; i < data.length; i += 4) {
        if (data[i + 3] > 0) inked++
      }
      expect(inked).toBe(16 * 16)
    }
  })

  test('pattern fill respects shape geometry clipping', () => {
    const canvas = createCanvas(32, 32)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('triangle', 32, 32)
    const style: DrawingStyle = {
      fill: {
        kind: 'pattern',
        preset: 'smGrid',
        fgColor: { r: 255, g: 0, b: 0, a: 1 },
        bgColor: { r: 0, g: 0, b: 255, a: 1 },
      },
      issues: [],
    }
    paintGeometry(ctx as never, geom, style, 32, 32)
    const data = ctx.getImageData(0, 0, 32, 32).data
    // Top-left corner (0,0) of a standard upward triangle should be outside the path and remain transparent
    const topLeftAlpha = data[3]
    expect(topLeftAlpha).toBe(0)
    // Center of triangle (16, 20) should be inked
    const centerIdx = (20 * 32 + 16) * 4
    expect(data[centerIdx + 3]).toBeGreaterThan(200)
  })

  test('pattern fill works with resolveDrawingStyle inheritance from spPr', () => {
    const spPr = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:pattFill prst="lgGrid">` +
      `    <a:fgClr><a:srgbClr val="112233"/></a:fgClr>` +
      `    <a:bgClr><a:srgbClr val="445566"/></a:bgClr>` +
      `  </a:pattFill>` +
      `</a:spPr>`
    )
    const style = resolveDrawingStyle(spPr)
    expect(style.fill).toEqual({
      kind: 'pattern',
      preset: 'lgGrid',
      fgColor: { r: 0x11, g: 0x22, b: 0x33, a: 1 },
      bgColor: { r: 0x44, g: 0x55, b: 0x66, a: 1 },
    })
  })

  test('pattern fill respects group transforms (translation and rotation)', () => {
    const canvas = createCanvas(64, 64)
    const ctx = canvas.getContext('2d')
    ctx.translate(16, 16)
    ctx.rotate(Math.PI / 4)
    const geom = resolvePreset('rect', 16, 16)
    const style: DrawingStyle = {
      fill: {
        kind: 'pattern',
        preset: 'dkDnDiag',
        fgColor: { r: 255, g: 0, b: 0, a: 1 },
        bgColor: { r: 0, g: 255, b: 0, a: 1 },
      },
      issues: [],
    }
    const issues = paintGeometry(ctx as never, geom, style, 16, 16)
    expect(issues).toEqual([])
    const data = ctx.getImageData(0, 0, 64, 64).data
    let inked = 0
    for (let i = 0; i < data.length; i += 4) {
      if (data[i + 3] > 100) inked++
    }
    expect(inked).toBeGreaterThan(20)
  })

  test('pattern LRU cache bounds capacity and verifies true LRU eviction', () => {
    clearPatternTileCache()
    expect(patternTileCacheSize()).toBe(0)

    const canvas = createCanvas(16, 16)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('rect', 16, 16)

    // Insert 64 distinct items
    for (let i = 0; i < 64; i++) {
      const style: DrawingStyle = {
        fill: {
          kind: 'pattern',
          preset: 'smGrid',
          fgColor: { r: i * 4, g: 0, b: 0, a: 1 },
          bgColor: { r: 0, g: i * 4, b: 0, a: 1 },
        },
        issues: [],
      }
      paintGeometry(ctx as never, geom, style, 16, 16)
    }
    expect(patternTileCacheSize()).toBe(64)
    // First inserted item (i=0) should exist
    expect(hasPatternTile('smGrid', 'rgba(0,0,0,1)', 'rgba(0,0,0,1)')).toBe(true)

    // Promote item 0 by painting it again
    const style0: DrawingStyle = {
      fill: {
        kind: 'pattern',
        preset: 'smGrid',
        fgColor: { r: 0, g: 0, b: 0, a: 1 },
        bgColor: { r: 0, g: 0, b: 0, a: 1 },
      },
      issues: [],
    }
    paintGeometry(ctx as never, geom, style0, 16, 16)
    expect(patternTileCacheSize()).toBe(64)

    // Insert 65th distinct item (i=65)
    const style65: DrawingStyle = {
      fill: {
        kind: 'pattern',
        preset: 'smGrid',
        fgColor: { r: 255, g: 255, b: 255, a: 1 },
        bgColor: { r: 100, g: 100, b: 100, a: 1 },
      },
      issues: [],
    }
    paintGeometry(ctx as never, geom, style65, 16, 16)
    expect(patternTileCacheSize()).toBe(64)

    // Item 0 was promoted, so it should still be in cache!
    expect(hasPatternTile('smGrid', 'rgba(0,0,0,1)', 'rgba(0,0,0,1)')).toBe(true)
    // Item 1 (not promoted) should have been evicted!
    expect(hasPatternTile('smGrid', 'rgba(4,0,0,1)', 'rgba(0,4,0,1)')).toBe(false)
  })

  test('pattern tile cache keys include raster size', () => {
    clearPatternTileCache()
    const canvas = createCanvas(16, 16)
    const ctx = canvas.getContext('2d') as any
    const t8 = getPatternTile(ctx, 'check', '#ff0000', '#00ff00', 8)
    const t16 = getPatternTile(ctx, 'check', '#ff0000', '#00ff00', 16)
    expect(t8).toBeDefined()
    expect(t16).toBeDefined()
    expect(t16).not.toBe(t8)
    expect((t8 as unknown as { width: number }).width).toBe(8)
    expect((t16 as unknown as { width: number }).width).toBe(16)
    expect(patternTileCacheSize()).toBe(2)
  })

  test('tile generator paints background only for unknown presets', () => {
    const canvas = createCanvas(8, 8)
    const ctx = canvas.getContext('2d') as any
    paintPatternTile(ctx, 'notAPreset', '#ff0000', '#00ff00', 8)
    const data = ctx.getImageData(0, 0, 8, 8).data
    for (let i = 0; i < data.length; i += 4) {
      expect(data[i]).toBe(0)
      expect(data[i + 1]).toBe(255)
      expect(data[i + 2]).toBe(0)
      expect(data[i + 3]).toBe(255)
    }
  })

  test('pattern fill applied to shape stroke with dashes and arrowheads', () => {
    const canvas = createCanvas(40, 40)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('line', 40, 40)
    const style: DrawingStyle = {
      line: {
        width: 4,
        dash: 'dash',
        headEnd: { type: 'arrow', width: 'med', length: 'med' },
        fill: {
          kind: 'pattern',
          preset: 'dkUpDiag',
          fgColor: { r: 255, g: 0, b: 0, a: 1 },
          bgColor: { r: 0, g: 0, b: 255, a: 1 },
        },
      },
      issues: [],
    }
    const issues = paintGeometry(ctx as never, geom, style, 40, 40)
    expect(issues).toEqual([])
    const data = ctx.getImageData(0, 0, 40, 40).data
    let inked = 0
    for (let i = 0; i < data.length; i += 4) {
      if (data[i + 3] > 100) inked++
    }
    expect(inked).toBeGreaterThan(10)
  })

  test('pattern fill with translucent colors paints with variable alpha', () => {
    const canvas = createCanvas(16, 16)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('rect', 16, 16)
    const style: DrawingStyle = {
      fill: {
        kind: 'pattern',
        preset: 'smGrid',
        fgColor: { r: 255, g: 0, b: 0, a: 0.5 },
        bgColor: { r: 0, g: 255, b: 0, a: 0.2 },
      },
      issues: [],
    }
    const issues = paintGeometry(ctx as never, geom, style, 16, 16)
    expect(issues).toEqual([])
    const data = ctx.getImageData(0, 0, 16, 16).data
    let hasTranslucent = false
    for (let i = 0; i < data.length; i += 4) {
      const a = data[i + 3]
      if (a > 20 && a < 250) hasTranslucent = true
    }
    expect(hasTranslucent).toBe(true)
  })

  test('tile generator generates seamless bitmasks for extended pattern presets (pct10–pct90, horz, vert, grid, check, trellis)', () => {
    const countBits = (mask: readonly number[]) => {
      let count = 0
      for (const row of mask) {
        for (let b = 0; b < 8; b++) {
          if ((row & (1 << b)) !== 0) count++
        }
      }
      return count
    }

    // 1. Percentage presets
    const pcts = [
      { name: 'pct5', expected: 3 },
      { name: 'pct10', expected: 6 },
      { name: 'pct20', expected: 13 },
      { name: 'pct25', expected: 16 },
      { name: 'pct30', expected: 19 },
      { name: 'pct40', expected: 26 },
      { name: 'pct50', expected: 32 },
      { name: 'pct60', expected: 38 },
      { name: 'pct70', expected: 45 },
      { name: 'pct75', expected: 48 },
      { name: 'pct80', expected: 51 },
      { name: 'pct90', expected: 58 },
    ]

    let prevCount = 0
    for (const { name, expected } of pcts) {
      const mask = getPatternBitmask(name)
      expect(mask, `preset ${name} should have a bitmask`).toBeDefined()
      expect(mask!).toHaveLength(8)
      for (const row of mask!) {
        expect(row).toBeGreaterThanOrEqual(0)
        expect(row).toBeLessThanOrEqual(255)
      }
      const fgCount = countBits(mask!)
      expect(Math.abs(fgCount - expected)).toBeLessThanOrEqual(1)
      expect(fgCount).toBeGreaterThan(prevCount)
      prevCount = fgCount
    }

    // 2. horz preset: has horizontal rules across all 8 columns
    const horzMask = getPatternBitmask('horz')
    expect(horzMask).toBeDefined()
    expect(horzMask!).toHaveLength(8)
    const horzFullRows = horzMask!.filter(r => r === 0xff)
    expect(horzFullRows.length).toBeGreaterThanOrEqual(1)

    // 3. vert preset: has vertical rules running down all rows
    const vertMask = getPatternBitmask('vert')
    expect(vertMask).toBeDefined()
    expect(vertMask!).toHaveLength(8)
    // Every row must have identical column bits enabled
    for (const r of vertMask!) {
      expect(r).toBe(vertMask![0])
      expect(r).toBeGreaterThan(0)
    }

    // 4. grid preset: orthogonal grid combining horizontal and vertical rules
    const gridMask = getPatternBitmask('grid')
    expect(gridMask).toBeDefined()
    expect(gridMask!).toHaveLength(8)
    const gridFullRows = gridMask!.filter(r => r === 0xff)
    expect(gridFullRows.length).toBeGreaterThanOrEqual(1)
    const gridCols = gridMask!.filter(r => r !== 0xff)
    expect(gridCols.length).toBeGreaterThan(0)
    for (const r of gridCols) {
      expect(r).toBe(gridCols[0])
    }

    // 5. check preset: 50% checkerboard (32 fg pixels, 32 bg pixels)
    const checkMask = getPatternBitmask('check')
    expect(checkMask).toBeDefined()
    expect(checkMask!).toHaveLength(8)
    expect(countBits(checkMask!)).toBe(32)

    // 6. trellis preset: mesh/lattice structure
    const trellisMask = getPatternBitmask('trellis')
    expect(trellisMask).toBeDefined()
    expect(trellisMask!).toHaveLength(8)
    const trellisCount = countBits(trellisMask!)
    expect(trellisCount).toBeGreaterThan(10)
    expect(trellisCount).toBeLessThan(54)
  })

  test('extended pattern fills paint with correct foreground and background colors and LRU cache reuse', () => {
    clearPatternTileCache()
    expect(patternTileCacheSize()).toBe(0)

    const extendedPresets: PatternPreset[] = [
      'pct10', 'pct25', 'pct50', 'pct75', 'pct90',
      'horz', 'vert', 'grid', 'check', 'trellis',
    ]

    const fgRed = { r: 255, g: 0, b: 0, a: 1 }
    const bgGreen = { r: 0, g: 255, b: 0, a: 1 }

    for (const preset of extendedPresets) {
      const canvas = createCanvas(32, 32)
      const ctx = canvas.getContext('2d')
      const geom = resolvePreset('rect', 32, 32)
      const style: DrawingStyle = {
        fill: {
          kind: 'pattern',
          preset,
          fgColor: fgRed,
          bgColor: bgGreen,
        },
        issues: [],
      }

      // 1. Initial paint: creates and caches tile
      const issues = paintGeometry(ctx as never, geom, style, 32, 32)
      expect(issues).toEqual([])

      const imgData = ctx.getImageData(0, 0, 32, 32).data
      let redCount = 0
      let greenCount = 0
      for (let i = 0; i < imgData.length; i += 4) {
        const r = imgData[i], g = imgData[i + 1], b = imgData[i + 2], a = imgData[i + 3]
        if (a > 200 && r > 200 && g < 50 && b < 50) redCount++
        if (a > 200 && r < 50 && g > 200 && b < 50) greenCount++
      }

      expect(redCount, `${preset} should have red foreground pixels`).toBeGreaterThan(0)
      expect(greenCount, `${preset} should have green background pixels`).toBeGreaterThan(0)
      expect(redCount + greenCount, `${preset} should fully cover pixels`).toBe(32 * 32)

      // Verify specific proportions
      if (preset === 'pct50' || preset === 'check') {
        expect(Math.abs(redCount - greenCount)).toBeLessThanOrEqual(32) // approx 50/50
      } else if (preset === 'pct25') {
        expect(redCount).toBeLessThan(greenCount)
      } else if (preset === 'pct75') {
        expect(redCount).toBeGreaterThan(greenCount)
      }

      // Check tile is present in cache
      expect(hasPatternTile(preset, 'rgba(255,0,0,1)', 'rgba(0,255,0,1)')).toBe(true)

      // 2. Second paint: reuses cached tile without increasing cache size
      const sizeBefore = patternTileCacheSize()
      const canvas2 = createCanvas(32, 32)
      const ctx2 = canvas2.getContext('2d')
      paintGeometry(ctx2 as never, geom, style, 32, 32)
      expect(patternTileCacheSize()).toBe(sizeBefore)
    }

    expect(patternTileCacheSize()).toBe(extendedPresets.length)
  })

  test('unsupported pattern presets degrade gracefully to solid foreground fill with truthful diagnostics', () => {
    // 1. Shape fill degradation and diagnostic
    const unsupportedPresets = ['weave', 'sphere', 'confetti', 'zigZag', 'nonexistentPattern']

    for (const prst of unsupportedPresets) {
      const xml = parseXmlOrdered(
        `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `  <a:pattFill prst="${prst}">` +
        `    <a:fgClr><a:srgbClr val="0000FF"/></a:fgClr>` +
        `    <a:bgClr><a:srgbClr val="FFFF00"/></a:bgClr>` +
        `  </a:pattFill>` +
        `</a:spPr>`
      )
      const fillDef = parseFillDefinition(xml)
      const issues: any[] = []
      const resolved = resolveFill(fillDef, undefined, undefined, issues)

      // Truthful diagnostic recorded
      expect(issues.some(i => i.kind === 'unsupported-fill' && i.feature === prst),
        `Preset ${prst} should record unsupported-fill diagnostic`).toBe(true)

      // Degrades to solid foreground color
      expect(resolved).toEqual({
        kind: 'solid',
        color: { r: 0, g: 0, b: 255, a: 1 },
      })

      // 2. Canvas paint test: renders as 100% solid foreground color without crashing
      const canvas = createCanvas(16, 16)
      const ctx = canvas.getContext('2d')
      const geom = resolvePreset('rect', 16, 16)
      const style: DrawingStyle = {
        fill: resolved,
        issues,
      }
      const paintIssues = paintGeometry(ctx as never, geom, style, 16, 16)
      expect(paintIssues).toEqual([])

      const data = ctx.getImageData(0, 0, 16, 16).data
      let bluePixels = 0
      let otherPixels = 0
      for (let i = 0; i < data.length; i += 4) {
        const r = data[i], g = data[i + 1], b = data[i + 2], a = data[i + 3]
        if (a > 200 && b > 200 && r < 50 && g < 50) {
          bluePixels++
        } else {
          otherPixels++
        }
      }
      expect(bluePixels).toBe(16 * 16)
      expect(otherPixels).toBe(0)
    }
  })
})
