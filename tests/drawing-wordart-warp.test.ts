/**
 * WordArt WARP validation suite (isolated proposal — scratch workspace).
 *
 * The authoritative matrix is the ECMA-376-derived 118-case fixture provided by
 * warp-fix-preparation (independent of production code and of this suite). All
 * expectations here come from that fixture or from first principles, never from
 * the implementation under test.
 *
 * Gates:
 * - PIN  official pinned catalog: reference hash, 11 presets, path topology
 * - G    geometry resolution vs 118 cases (guides @1e-8, paths @1e-6)
 * - B    official defaults + handle bounds + entry clamping edges
 * - E    envelope interior interpolation (centerline between boundaries)
 * - P1   identity warp paints byte-identical to unwarped text
 * - P2   shaped run joins ONE source fill (Arabic connectors survive)
 * - P3   gradients: identity byte equal; mapped gradient follows source u
 * - P4   raster quality: interior alpha coverage / shadow single composite
 * - S1   production index/copy without warp (regression guard)
 * - S2   production hitTest/search/copy on REAL warped ink (ink-sampled)
 * - PARSER  parse-side defaults/clamping flow through text.ts/text-parse.ts
 * - C1   pattern cache is a true hit-promoting LRU observed via real paint
 */
import { createCanvas } from 'canvas'
import { describe, expect, it, vi } from 'vitest'
// Runtime is pinned: vitest resolves transforms/vitest options from the root
// config, so import the matrix by absolute workspace-relative path.
import { TEXT_WARP_CATALOG, OFFICIAL_TEXT_WARP_REFERENCE } from '../src/drawing/text-warp-catalog'
import {
  SUPPORTED_TEXT_WARP_PRESETS, DEFAULT_WARP_ADJUSTMENTS as RESOLVER_DEFAULTS,
  resolveWarp, buildWarpMapping, buildWarpLattice, clampAdjustment,
} from '../src/drawing/text-warp'
import { resolveTextWarpPaths } from '../src/drawing/geometry'
import { DEFAULT_WARP_ADJUSTMENTS as PARSE_DEFAULTS } from '../src/drawing/text'
import { buildTextIndex, findMatches } from '../src/core/search'
import { hitTest, textForRange } from '../src/core/selection'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'
import { paintTextBody } from '../src/drawing/text-paint'
import { RECORD_TEXT } from '../src/core/text-recording'
import type { DrawingTextBody } from '../src/drawing/text'

// Tracked repository-relative fixture: exact copy (SHA-256 83d64c2d…) of the
// independent warp-fix-preparation/authoritative-matrix.json. Portable: static
// JSON import, no /tmp dependency.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
import MATRIX_JSON from './fixtures/warp-authoritative-matrix.json'
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const MATRIX: any = MATRIX_JSON as any
const GUIDE_TOL = MATRIX.tolerances.guideAbsolute
const POINT_TOL = MATRIX.tolerances.pointAbsolutePx
const MARGIN = 8
const identityFont = (family: string): string => family

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

const warpXml = (preset: string, adjustments?: Record<string, number>): string => {
  const gd = Object.entries(adjustments ?? {}).map(([name, value]) =>
    `<a:gd name="${name}" fmla="val ${value}"/>`).join('')
  return `<a:prstTxWarp prst="${preset}"><a:avLst>${gd}</a:avLst></a:prstTxWarp>`
}

/** Body builder. 'styles' come pre-serialized; XML color values are 6-digit
 * hex WITHOUT a leading '#' — the caller must respect DrawingML exactly. */
function warpBody(preset: string | undefined, adjustments: Record<string, number> | undefined, bodyText: string, styles = '', indent = ''): DrawingTextBody {
  const body = parseTextBody(parseXmlOrdered(`
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr ${indent}>${preset ? warpXml(preset, adjustments) : ''}</a:bodyPr>
  <a:p><a:r><a:rPr sz="3600">${styles}</a:rPr><a:t>${bodyText}</a:t></a:r></a:p>
</a:txBody>`))
  body.insetLeftEmu = 0; body.insetRightEmu = 0; body.insetTopEmu = 0; body.insetBottomEmu = 0
  return body
}

interface HostLog { fills: string[]; targetFills: string[]; drawImage: number[] }

/** Test-only OffscreenCanvas host (per the fix guidance): injects a constructor
 * that counts 8×8 tile requests and returns REAL node-canvas surfaces, so the
 * production createPattern path is observed end-to-end. Raster contexts are
 * wrapped so fillText/drawImage can be logged through the real surface. */
function installCountingHost(log?: HostLog): { tile8: number; creations: number; restore: () => void } {
  const spy = { tile8: 0, creations: 0 }
  const realCreateCanvas = createCanvas as unknown as (w: number, h: number) => any
  const Host = function (this: unknown, w: number, h: number) {
    spy.creations++
    if (w === 8 && h === 8) spy.tile8++
    const surface = realCreateCanvas.call(null, w, h)
    if ((w !== 8 || h !== 8) && log) {
      const original = surface.getContext.bind(surface)
      surface.getContext = (kind: string): unknown => {
        const ctx = original(kind) as any
        return wrapContext(ctx, log)
      }
    }
    return surface
  } as unknown as typeof OffscreenCanvas
  ;(globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = Host
  return {
    get tile8() { return spy.tile8 },
    get creations() { return spy.creations },
    restore() { delete (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas },
  }
}

/** Proxy-backed context wrapper: property SETS delegate to the real context
 * (font/accessors survive), gets return logged variants ONLY for the few
 * paint calls under observation. Watches fillText/strokeText/drawImage. */
function wrapContext(ctx: unknown, log?: HostLog, key = 'fillText'): unknown {
  const real = ctx as any
  let wrapping = false
  return new Proxy(real, {
    get(target, prop) {
      if (prop === key && log) {
        return (text: string, ...rest: unknown[]): void => {
          log.fills.push(text)
          return target[prop](text, ...rest)
        }
      }
      if (prop === 'strokeText' && log) {
        return (text: string, ...rest: unknown[]): void => {
          log.targetFills.push(text)
          return target[prop](text, ...rest)
        }
      }
      if (prop === 'drawImage' && log) {
        return (...args: unknown[]): void => {
          log.drawImage.push(1)
          return target.drawImage(...args as [any, number, number])
        }
      }
      if (prop === 'getContext' && log) {
        return (kind: string): unknown => {
          const inner = target.getContext(kind)
          if (!inner || wrapping) return inner
          return inner
        }
      }
      const value = target[prop]
      return typeof value === 'function' ? value.bind(target) : value
    },
    set(target, prop, value) {
      target[prop] = value
      return true
    },
  })
}

/** Wrap a caller context so its own fillText/strokeText/drawImage calls are
 * logged (target fills) while every other property set/get delegates. */
function targetFillWrap(ctx: CanvasRenderingContext2D, log: HostLog): unknown {
  const real = ctx as unknown as Record<string | symbol, unknown>
  return new Proxy(real, {
    get(target, prop) {
      if (prop === 'fillText') return (text: string, ...rest: unknown[]): void => {
        log.targetFills.push(text)
        return (ctx as unknown as { fillText: (t: string, ...r: unknown[]) => void }).fillText(text, ...rest)
      }
      if (prop === 'strokeText') return (text: string, ...rest: unknown[]): void => {
        log.targetFills.push(text)
        return (ctx as unknown as { strokeText: (t: string, ...r: unknown[]) => void }).strokeText(text, ...rest)
      }
      if (prop === 'drawImage') return (...args: unknown[]): void => {
        log.drawImage.push(1)
        return (ctx as unknown as { drawImage: (...r: unknown[]) => void }).drawImage(...args)
      }
      const value = target[prop]
      return typeof value === 'function' ? value.bind(target) : value
    },
    set(target, prop, value) {
      target[prop as string] = value
      return true
    },
  })
}

const paintBody = (body: DrawingTextBody, w: number, h: number): { data: Uint8ClampedArray; ctx: unknown } => {
  const canvas = createCanvas(w, h)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, w, h)
  paintTextBody(body, ctx as unknown as never, MARGIN, MARGIN, w - 2 * MARGIN, h - 2 * MARGIN, identityFont)
  return { data: ctx.getImageData(0, 0, w, h).data, ctx }
}

const inkCount = (data: Uint8ClampedArray): number => {
  let count = 0
  for (let i = 0; i < data.length; i += 4) if (data[i + 3] > 0 && (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250)) count++
  return count
}

type MatrixCommand =
  | { kind: 'moveTo' | 'lnTo' | 'quadBezTo' | 'cubicBezTo'; points: Array<{ x: number; y: number }> }
  | { kind: 'arcTo'; attrs: { wR: number; hR: number; stAng: number; swAng: number } }

// ---------------------------------------------------------------------------
// PIN: official catalog pinning
// ---------------------------------------------------------------------------

describe('WordArt warp official catalog pin', () => {
  it('pinned reference (pinned SHA-256) and eleven presets with full path topology', () => {
    expect(OFFICIAL_TEXT_WARP_REFERENCE.sha256).toBe('89640bd4aeedfa2c578d1d1acc50d569865ebdef315fd4b468087c25e36acada')
    const presets = new Set(MATRIX.catalog.map((entry: { preset: string }) => entry.preset))
    expect(presets.size).toBe(11)
    for (const { preset, pathCount } of MATRIX.catalog as Array<{ preset: string; pathCount: number }>) {
      const definition = (TEXT_WARP_CATALOG as Record<string, { paths: ReadonlyArray<unknown> }>)[preset]
      expect(definition, preset).toBeDefined()
      expect(definition.paths.length, `${preset} pathCount`).toBe(pathCount)
      expect(SUPPORTED_TEXT_WARP_PRESETS.has(preset as never), `${preset} is supported`).toBe(true)
    }
  })
})

// ---------------------------------------------------------------------------
// G: geometry resolution against the authoritative matrix
// ---------------------------------------------------------------------------

function evaluateG1Mismatches(
  resolveFn: (definition: any, width: number, height: number, authored: Record<string, number>) => { values: Record<string, number>; paths: MatrixCommand[][] },
): string[] {
  const mismatches: string[] = []
  for (const c of MATRIX.cases as Array<{ id: string; preset: string; box: { width: number; height: number }; authoredAdjustments: Record<string, number>; expected: { guides: Record<string, number>; bounds: Array<{ guide: string; axis: string; minimum: number; maximum: number }>; paths: MatrixCommand[][]; defaults: Record<string, number> } }>) {
    const definition = (TEXT_WARP_CATALOG as Record<string, never>)[c.preset]
    const got = resolveFn(definition, c.box.width, c.box.height, c.authoredAdjustments)
    for (const [name, expected] of Object.entries(c.expected.guides)) {
      const value = (got.values as Record<string, number>)[name]
      if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(expected - value) > GUIDE_TOL) {
        mismatches.push(`${c.id} guide ${name}: expected ${expected} got ${value}`)
      }
    }
  }
  return mismatches
}

function evaluateG2Mismatches(
  resolveFn: (definition: any, width: number, height: number, authored: Record<string, number>) => { values: Record<string, number>; paths: Array<{ commands: any[] }> },
): string[] {
  const mismatches: string[] = []
  for (const c of MATRIX.cases as Array<{ id: string; preset: string; box: { width: number; height: number }; authoredAdjustments: Record<string, number>; expected: { paths: MatrixCommand[][] } }>) {
    const definition = (TEXT_WARP_CATALOG as Record<string, never>)[c.preset]
    const got = resolveFn(definition, c.box.width, c.box.height, c.authoredAdjustments)
    if (got.paths.length !== c.expected.paths.length) {
      mismatches.push(`${c.id}: path count ${got.paths.length} != ${c.expected.paths.length}`)
      continue
    }
    got.paths.forEach((resolvedPath, pathIndex) => {
      const matrixCommands = c.expected.paths[pathIndex]
      if (resolvedPath.commands.length !== matrixCommands.length) {
        mismatches.push(`${c.id}: path ${pathIndex} command count ${resolvedPath.commands.length} != ${matrixCommands.length}`)
        return
      }
      resolvedPath.commands.forEach((command, ci) => {
        const e = matrixCommands[ci]
        if (command.kind !== e.kind) {
          mismatches.push(`${c.id}: path ${pathIndex} cmd ${ci} kind ${command.kind} != ${e.kind}`)
          return
        }
        if (e.kind === 'arcTo' && command.kind === 'arcTo') {
          for (const attr of ['wR', 'hR', 'stAng', 'swAng'] as const) {
            const val = command.attrs[attr]
            if (typeof val !== 'number' || !Number.isFinite(val) || Math.abs(val - e.attrs[attr]) > POINT_TOL) {
              mismatches.push(`${c.id}: path ${pathIndex} cmd ${ci} ${attr} expected ${e.attrs[attr]} got ${val}`)
            }
          }
        } else if ('values' in command && 'points' in e) {
          const points = e.points
          if (command.values.length !== points.length * 2) {
            mismatches.push(`${c.id}: path ${pathIndex} cmd ${ci} point arity ${command.values.length} vs ${points.length * 2}`)
            return
          }
          points.forEach((p: any, i: number) => {
            const vx = command.values[2 * i]
            const vy = command.values[2 * i + 1]
            if (typeof vx !== 'number' || !Number.isFinite(vx) || typeof vy !== 'number' || !Number.isFinite(vy) || Math.abs(vx - p.x) > POINT_TOL || Math.abs(vy - p.y) > POINT_TOL) {
              mismatches.push(`${c.id}: path ${pathIndex} cmd ${ci} pt ${i} expected (${p.x},${p.y}) got (${vx},${vy})`)
            }
          })
        }
      })
    })
  }
  return mismatches
}

describe('WordArt warp official geometry vs 118-case matrix', () => {
  it('G1: derived guides match for all 118 cases (tolerance from the matrix)', () => {
    const mismatches = evaluateG1Mismatches(resolveTextWarpPaths as any)
    expect(mismatches, mismatches.slice(0, 6).join(' | ')).toEqual([])
  })

  it('G1-negative: mutated NaN guides fail the authoritative finite gate', () => {
    const nanResolve = (def: any, w: number, h: number, authored: any) => {
      const got = resolveTextWarpPaths(def, w, h, authored)
      for (const name of Object.keys(got.values)) got.values[name] = NaN
      return got as any
    }
    const mismatches = evaluateG1Mismatches(nanResolve)
    expect(mismatches.length).toBeGreaterThan(0)
  })

  it('G2: resolved boundary paths match for all 118 cases (206 paths)', () => {
    const mismatches = evaluateG2Mismatches(resolveTextWarpPaths as any)
    expect(mismatches, mismatches.slice(0, 6).join(' | ')).toEqual([])
  })

  it('G2-negative: mutated NaN boundary path values fail the authoritative finite gate', () => {
    const nanResolve = (def: any, w: number, h: number, authored: any) => {
      const got = resolveTextWarpPaths(def, w, h, authored)
      for (const path of got.paths) {
        for (const cmd of path.commands) {
          if (cmd.kind === 'arcTo') {
            for (const attr of ['wR', 'hR', 'stAng', 'swAng'] as const) cmd.attrs[attr] = NaN
          } else if ('values' in cmd) {
            cmd.values = (cmd.values as number[]).map(() => NaN)
          }
        }
      }
      return got as any
    }
    const mismatches = evaluateG2Mismatches(nanResolve)
    expect(mismatches.length).toBeGreaterThan(0)
  })

  it('G3: official defaults, handle bounds and entry-clamped effective inputs', () => {
    const mismatches: string[] = []
    for (const c of MATRIX.cases as Array<{ id: string; preset: string; box: { width: number; height: number }; authoredAdjustments: Record<string, number>; expected: { defaults: Record<string, number>; bounds: Array<{ guide: string; axis: string; minimum: number; maximum: number }> } }>) {
      // Catalog-derived resolver defaults == matrix defaults for every case.
      for (const [name, expected] of Object.entries(c.expected.defaults)) {
        const value = (RESOLVER_DEFAULTS as Record<string, Record<string, number>>)[c.preset][name]
        if (value === undefined || Math.abs(expected - value) > GUIDE_TOL) {
          mismatches.push(`${c.id} default ${name}: expected ${expected} got ${value}`)
        }
      }
      const diagnosis = resolveWarp(c.preset as never, c.box.width, c.box.height, {})
      if (!diagnosis) { mismatches.push(`${c.id}: resolveWarp returned undefined`); continue }
      for (const b of c.expected.bounds) {
        const gotBound = diagnosis.bounds.find(x => x.guide === b.guide)
        if (!gotBound || gotBound.axis !== b.axis || Math.abs(gotBound.min - b.minimum) > GUIDE_TOL || Math.abs(gotBound.max - b.maximum) > GUIDE_TOL) {
          mismatches.push(`${c.id} bound ${b.guide} (${b.axis}) expected [${b.minimum},${b.maximum}] got ${JSON.stringify(gotBound)}`)
        }
      }
      // Entry clamping maps authored out-of-range values onto the official
      // bound, and beyond-bound corner cases produce a clamped-guide issue.
      for (const [name, authored] of Object.entries(c.authoredAdjustments)) {
        if (authored < (diagnosis.bounds.find(x => x.guide === name)?.min ?? -Infinity) ||
            authored > (diagnosis.bounds.find(x => x.guide === name)?.max ?? Infinity)) {
          const clamped = resolveWarp(c.preset as never, c.box.width, c.box.height, { [name]: authored })
          if (clamped) {
            expect(clamped.effective[name], `${c.id} effective ${name} clamps to the official bound`).toBe(clampAdjustment(c.preset as never, name, authored))
            expect(clamped.issues.some(issue => issue.kind === 'clamped-guide'), `${c.id} records a clamped-guide issue`).toBe(true)
          }
        }
      }
    }
    expect(mismatches, mismatches.slice(0, 6).join(' | ')).toEqual([])
  })

  it('B: official bound edges apply at exactly the documented values', () => {
    expect(clampAdjustment('textArchUp', 'adj', 0)).toBe(0)
    expect(clampAdjustment('textArchUp', 'adj', 21599999)).toBe(21599999)
    expect(clampAdjustment('textArchUp', 'adj', 21600000)).toBe(21599999)
    expect(clampAdjustment('textWave1', 'adj1', -1)).toBe(0)
    expect(clampAdjustment('textWave1', 'adj1', 20001)).toBe(20000)
    expect(clampAdjustment('textWave1', 'adj2', -10001)).toBe(-10000)
    expect(clampAdjustment('textWave1', 'adj2', 10001)).toBe(10000)
    expect(clampAdjustment('textInflate', 'adj', 20001)).toBe(20000)
    expect(clampAdjustment('textDeflate', 'adj', 37501)).toBe(37500)
    expect(clampAdjustment('textSlantUp', 'adj', 71432)).toBe(71431)
    expect(clampAdjustment('textSlantDown', 'adj', 28568)).toBe(28569)
    expect(clampAdjustment('textSlantDown', 'adj', 100001)).toBe(100000)
    expect(clampAdjustment('textCurveUp', 'adj', 56339)).toBe(56338)
  })
})

// ---------------------------------------------------------------------------
// E: envelope interior interpolation
// ---------------------------------------------------------------------------

describe('WordArt warp envelope interior mapping', () => {
  const SAMPLED_U = [0, 0.1, 0.25, 0.5, 0.75, 0.9, 1]
  it('two-boundary presets interpolate linearly between exact boundary rows', () => {
    for (const preset of ['textWave1', 'textWave2', 'textInflate', 'textDeflate', 'textSlantUp', 'textSlantDown', 'textCurveUp', 'textCurveDown'] as const) {
      // A case with a NONZERO deformation (default) so interpolation is real.
      const mapping = buildWarpMapping({ preset, adjustments: {} }, 200, 100)
      expect(mapping.family, preset).toBe('envelope')
      for (const u of SAMPLED_U) {
        const top = mapping.map(u, 0)
        const bottom = mapping.map(u, 1)
        const center = mapping.map(u, 0.5)
        // Linear interior requires the centerline to sit between the exact rows.
        const midX = (top.x + bottom.x) / 2, midY = (top.y + bottom.y) / 2
        expect(Math.abs(center.x - midX), `${preset} u=${u} interior x interpolates`).toBeLessThan(1e-6)
        expect(Math.abs(center.y - midY), `${preset} u=${u} interior y interpolates`).toBeLessThan(1e-6)
      }
    }
  })

  it('follow-path presets displace the band normal of a single centerline', () => {
    for (const preset of ['textArchUp', 'textArchDown', 'textCircle'] as const) {
      const mapping = buildWarpMapping({ preset, adjustments: {} }, 200, 100)
      expect(mapping.family, preset).toBe('follow-path')
      for (const u of SAMPLED_U) {
        // v=0 and v=1 ride to opposite sides of the SAME centerline point.
        const hi = mapping.map(u, 1), lo = mapping.map(u, 0)
        const cx = (hi.x + lo.x) / 2, cy = (hi.y + lo.y) / 2
        expect(Math.abs(cx - mapping.boundaryAt(0, u).x), `${preset} u=${u} band normal center`).toBeLessThanOrEqual(1e-6)
        expect(Math.abs(cy - mapping.boundaryAt(0, u).y), `${preset} u=${u} band normal center`).toBeLessThanOrEqual(1e-6)
      }
    }
  })

  it('identity detection: interior mapping is exactly identity at zero handles', () => {
    // textWave1/2 at amplitude handles 0 (matrix amplitude-min) resolve to a
    // numerically identity mapping over the body box.
    for (const preset of ['textWave1', 'textWave2', 'textSlantUp', 'textCurveUp'] as const) {
      const minAdjust: Record<string, number> = preset === 'textWave1' || preset === 'textWave2' ? { adj1: 0, adj2: 0 } : { adj: 0 }
      const mapping = buildWarpMapping({ preset, adjustments: minAdjust }, 200, 100)
      expect(mapping.identity, `${preset} with zero handles is identity`).toBe(true)
      for (const [i, j] of [[0, 0], [0.5, 0.5], [1, 1], [0.25, 0.75]]) {
        const mapped = mapping.map(i, j)
        expect(mapped.x).toBeCloseTo(i * 200, 6)
        expect(mapped.y).toBeCloseTo(j * 100, 6)
      }
      // Lattice respects its region and never exceeds 64×64 cells.
      const lattice = buildWarpLattice(mapping, { u0: 0.2, u1: 0.8, v0: 0.3, v1: 0.7 })
      expect(lattice.cols).toBeLessThanOrEqual(64)
      expect(lattice.rows).toBeLessThanOrEqual(64)
      expect(lattice.points[0][0].x).toBeCloseTo(0.2 * 200, 6)
    }
  })
})

// ---------------------------------------------------------------------------
// P: paint behavior
// ---------------------------------------------------------------------------

describe('WordArt warp paint', () => {
  it('P1: identity warp paints byte-identical to the ordinary path', () => {
    const body = warpBody('textWave1', { adj1: 0, adj2: 0 }, 'سلام', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    expect(body.textWarp?.preset).toBe('textWave1')
    const warped = paintBody(body, 240, 130)
    const plain = paintBody(warpBody(undefined, undefined, 'سلام', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>'), 240, 130)
    const warpMapping = buildWarpMapping({ preset: 'textWave1', adjustments: { adj1: 0, adj2: 0 } }, 224, 114)
    expect(warpMapping.identity, 'zero-handle wave resolves to identity').toBe(true)
    expect(Array.from(warped.data)).toEqual(Array.from(plain.data))
    expect(inkCount(warped.data)).toBeGreaterThan(60)
  })

  it('P2: compatible run groups paint ONE shaped source fill (contextual joins survive)', async () => {
    vi.resetModules()
    const log: HostLog = { fills: [], targetFills: [], drawImage: [] }
    const spy = installCountingHost(log)
    const { paintTextBody: fresh } = await import('../src/drawing/text-paint')
    const canvas = createCanvas(240, 130)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 240, 130)
    const shared = '<a:solidFill><a:srgbClr val="200000"/></a:solidFill><a:effectLst/>'
    const body = parseTextBody(parseXmlOrdered(`
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr>${warpXml('textCurveUp', { adj: 45977 })}</a:bodyPr>
  <a:p>
    <a:r><a:rPr sz="3600">${shared}</a:rPr><a:t>س</a:t></a:r>
    <a:r><a:rPr sz="3600">${shared}</a:rPr><a:t>لام</a:t></a:r>
  </a:p>
</a:txBody>`))
    body.insetLeftEmu = 0; body.insetRightEmu = 0; body.insetTopEmu = 0; body.insetBottomEmu = 0
    fresh(body, targetFillWrap(ctx as unknown as CanvasRenderingContext2D, log) as unknown as never, MARGIN, MARGIN, 224, 114, identityFont)
    expect(log.fills, 'whole compatible run reaches the raster in ONE source call').toEqual(['سلام'])
    expect(log.targetFills, 'mapped ink is rasterized, not re-filled through the target').toEqual([])
    spy.restore()
    vi.resetModules()
  })

  it('P3: gradients stay continuous in source segment coordinates', () => {
    const gradient = '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="0000FF"/></a:gs><a:gs pos="100000"><a:srgbClr val="FF0000"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>'
    // Identity warp → byte-identical to unwarped.
    const identityBody = warpBody('textWave1', { adj1: 0, adj2: 0 }, 'ABC', gradient)
    const plainBody = warpBody(undefined, undefined, 'ABC', gradient)
    const identity = paintBody(identityBody, 240, 130), plain = paintBody(plainBody, 240, 130)
    expect(Array.from(identity.data)).toEqual(Array.from(plain.data))
    // Mapped warp: blue channel must follow the SOURCE x of each mapped call.
    const mapped = paintBody(warpBody('textCurveUp', { adj: 45977 }, 'ABC', gradient), 240, 130)
    const samples: Array<[x: number, blue: number]> = []
    const mappedMapping = buildWarpMapping({ preset: 'textCurveUp', adjustments: {} }, 224, 114)
    for (const sourceU of [0.25, 0.4, 0.5, 0.6, 0.75]) {
      const mappedPoint = mappedMapping.map(sourceU, 0.45)
      const px = Math.round(MARGIN + mappedPoint.x), py = Math.round(MARGIN + mappedPoint.y)
      const index = (py * 240 + px) * 4
      if (mapped.data[index + 3] > 120) samples.push([sourceU, mapped.data[index + 2]])
    }
    expect(samples.length).toBeGreaterThanOrEqual(3)
    for (let i = 1; i < samples.length; i++) {
      expect(samples[i][1], `mapped blue follows source u (${samples[i - 1][0]}→${samples[i][0]})`).toBeLessThanOrEqual(samples[i - 1][1] + 8)
    }
  })

  it('P4: mapped interior stays covered; shadow composites as ONE drawImage', async () => {
    vi.resetModules()
    // The counting host only COUNTS tile creations; the TARGET context is
    // separately wrapped so the observed drawImage/fill calls are strictly
    // caller-facing (raster-internal mesh draws stay internal by design).
    const log: HostLog = { fills: [], targetFills: [], drawImage: [] }
    const spy = installCountingHost()
    const { paintTextBody: fresh } = await import('../src/drawing/text-paint')
    const canvas = createCanvas(240, 130)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 240, 130)
    const body = warpBody('textCurveUp', { adj: 45977 },
      'WORD', `<a:solidFill><a:srgbClr val="100000"/></a:solidFill><a:effectLst><a:outerShdw dist="38100" dir="2640000"><a:srgbClr val="400000"/></a:outerShdw></a:effectLst>`)
    fresh(body, targetFillWrap(ctx as unknown as CanvasRenderingContext2D, log) as unknown as never, MARGIN, MARGIN, 224, 114, identityFont)
    expect(log.drawImage, 'shadow composites through ONE final drawImage').toEqual([1])
    expect(log.targetFills, 'warp ink is rasterized (never target-direct)').toEqual([])
    const data = ctx.getImageData(0, 0, 240, 130).data
    expect(inkCount(data)).toBeGreaterThan(300)
    // Interior coverage: sample mapped source lattice — mapped cells of a solid
    // body must hold meaningful ink (no holes) with bounded alpha deviation.
    const mapping = buildWarpMapping({ preset: 'textCurveUp', adjustments: {} }, 224, 114)
    let covered = 0
    let minAlpha = Infinity
    let maxAlpha = -Infinity
    for (let i = 0; i <= 8; i++) {
      for (let j = 0; j <= 8; j++) {
        const mapped = mapping.map(0.25 + 0.06 * i, 0.3 + 0.05 * j)
        const px = Math.round(MARGIN + mapped.x), py = Math.round(MARGIN + mapped.y)
        const index = (py * 240 + px) * 4
        const alpha = data[index + 3]
        if (alpha > 0) { covered++; minAlpha = Math.min(minAlpha, alpha); maxAlpha = Math.max(maxAlpha, alpha) }
      }
    }
    expect(covered, 'interior lattice keeps ink coverage').toBeGreaterThan(45)
    expect(maxAlpha - minAlpha, 'interior alpha deviation bounded ≤ 48').toBeLessThanOrEqual(48)
    spy.restore()
    vi.resetModules()
  })
})

// ---------------------------------------------------------------------------
// S: source/index/selection unification
// ---------------------------------------------------------------------------

describe('WordArt warp source index', () => {
  it('S1: production index partitions multi-run source text exactly', async () => {
    const body = parseTextBody(parseXmlOrdered(`
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr/><a:p>
    <a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="200000"/></a:solidFill></a:rPr><a:t>AA</a:t></a:r>
    <a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="400000"/></a:solidFill></a:rPr><a:t>Å👩‍💻B</a:t></a:r>
  </a:p></a:txBody>`))
    body.insetLeftEmu = 0; body.insetRightEmu = 0; body.insetTopEmu = 0; body.insetBottomEmu = 0
    const sourceText = body.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('')
    const index = await buildTextIndex([{ spec: { widthPx: 300, heightPx: 140 }, paint: replayCtx => {
      paintTextBody(body, replayCtx as unknown as never, MARGIN, MARGIN, 284, 124, identityFont)
    } }])
    expect(index.pages[0].lines.map(line => line.text)).toEqual([sourceText])
    const full = { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 0, charIndex: sourceText.length } }
    expect(textForRange(index, full)).toBe(sourceText)
    expect(textForRange(index, { start: full.start, end: { ...full.end, charIndex: 2 } })).toBe('AA')
    expect(findMatches(index, 'AA', { caseSensitive: false }))
    expect(findMatches(index, 'AA')).toHaveLength(1)
    expect(findMatches(index, '👩\u200d💻')).toHaveLength(1)
    expect(String.fromCodePoint(0x1F469, 0x200D, 0x1F4BB)).toHaveLength(5)  // UTF-16 family sanity
  })

  it('S2: production hitTest/search/copy hit REAL warped ink only', async () => {
    vi.resetModules()
    const { RECORD_TEXT } = await import('../src/core/text-recording')
    // A real warped paint: nonzero textCurveUp mapping carries mapped visual
    // geometry on the records; recording mode never draws, so the index is a
    // SEPARATE replay of the production recorder.
    const body = warpBody('textCurveUp', { adj: 45977 }, 'AB',
      '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const recorded: Array<{ text: string; x: number; y: number; width: number; logical: unknown }> = []
    const canvas = createCanvas(240, 130)
    const ctx = canvas.getContext('2d')
    const record = (ctx as unknown as Record<PropertyKey, unknown>)
    ;(ctx as unknown as CanvasRenderingContext2D & { [RECORD_TEXT]?: unknown })[RECORD_TEXT] = (text: string, x: number, y: number, width: number, logical: unknown) => {
      void record
      recorded.push({ text, x, y, width, logical })
    }
    // First: REAL warped paint (evidence pixels).
    const ink = paintBody(body, 240, 130)
    const inkSamples: Array<{ x: number; y: number }> = []
    expect(inkCount(ink.data)).toBeGreaterThan(40)
    // Second: SEPARATE recording index from an independent replayed paint of
    // the same body (records carry mapped visual geometry).
    const index = await buildTextIndex([{ spec: { widthPx: 240, heightPx: 130 }, paint: replayCtx => {
      paintTextBody(body, replayCtx as unknown as never, MARGIN, MARGIN, 224, 114, identityFont)
    } }])
    expect(index.pages[0].lines.map(line => line.text)).toEqual(['AB'])
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 2 } })).toBe('AB')
    expect(findMatches(index, 'AB', { caseSensitive: false })).toHaveLength(1)
    const line = index.pages[0].lines[0]
    const span = line.spans[0]
    expect(span.placement?.visual?.polygon?.length, 'warped record exposes mapped polygon').toBeGreaterThan(2)
    expect(span.placement?.visual?.clusters?.length, 'warped record exposes per-cluster centers').toBe(2)
    // Ink samples come from the RECORD's mapped cluster centers, verified as
    // REAL opaque pixels of the independently painted image.
    for (const span of index.pages[0].lines[0].spans) {
      const clusters = span.placement?.visual?.clusters ?? []
      for (const cluster of clusters) {
        inkSamples.push({ x: Math.round(cluster.x), y: Math.round(cluster.y) })
      }
    }
    expect(inkSamples.length).toBeGreaterThanOrEqual(2)
    for (const probe of inkSamples) {
      const index_ = (probe.y * 240 + probe.x) * 4
      expect(ink.data[index_ + 3], `probe (${probe.x},${probe.y}) actually lands on sampled real ink`).toBeGreaterThan(120)
      const hit = hitTest(index, 0, probe.x, probe.y, { strict: true })
      expect(hit, `strict production hitTest resolves real ink at (${probe.x},${probe.y})`).toBeDefined()
    }
    // Control: a probe OUTSIDE the warped ink must not hit under strict mode.
    const outside = hitTest(index, 0, 4, 4, { strict: true })
    expect(outside, 'strict hitTest rejects ink-less corners').toBeUndefined()
    vi.resetModules()
  })
})

// ---------------------------------------------------------------------------
// PARSER: pipeline uses official defaults and entry clamping
// ---------------------------------------------------------------------------

describe('WordArt warp parse pipeline', () => {
  it('parse-side defaults for supported presets come from the official catalog', () => {
    expect(PARSE_DEFAULTS.textWave1).toEqual({ adj1: 12500, adj2: 0 })
    expect(PARSE_DEFAULTS.textWave2).toEqual({ adj1: 12500, adj2: 0 })
    expect(PARSE_DEFAULTS.textArchUp).toEqual({ adj: 10800000 })
    expect(PARSE_DEFAULTS.textArchDown).toEqual({ adj: 0 })   // official archDown default keyed to the flipped path
    expect(PARSE_DEFAULTS.textCircle).toEqual({ adj: 10800000 })
    expect(PARSE_DEFAULTS.textInflate).toEqual({ adj: 18750 })
    expect(PARSE_DEFAULTS.textDeflate).toEqual({ adj: 18750 })
    expect(PARSE_DEFAULTS.textSlantUp).toEqual({ adj: 55555 })
    expect(PARSE_DEFAULTS.textSlantDown).toEqual({ adj: 44445 })
    expect(PARSE_DEFAULTS.textCurveUp).toEqual({ adj: 45977 })
    expect(PARSE_DEFAULTS.textCurveDown).toEqual({ adj: 45977 })
  })

  it('parser-generated models carry official defaults and authored values', () => {
    // Empty avLst → parse model keeps OFFICIAL defaults (wave negative adj2).
    const plain = warpBody('textWave1', undefined, 'X')
    expect(plain.textWarp?.preset).toBe('textWave1')
    expect(plain.textWarp?.adjustments).toEqual({ adj1: 12500, adj2: 0 })
    // Authored negative adj2 flows verbatim from the XML.
    const negative = warpBody('textWave1', { adj1: 12500, adj2: -30000 }, 'X')
    expect(negative.textWarp?.adjustments).toEqual({ adj1: 12500, adj2: -30000 })
    // ...and the WARP ENGINE clamps it at entry without touching the parse
    // model or the solver defaults.
    const clamped = resolveWarp('textWave1', 200, 100, negative.textWarp!.adjustments as Record<string, number>)
    expect(clamped!.effective).toEqual({ adj1: 12500, adj2: -10000 })
    expect(negative.textWarp!.adjustments).toEqual({ adj1: 12500, adj2: -30000 })
    expect(clamped!.issues.some(issue => issue.kind === 'clamped-guide')).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// C1: pattern cache is a TRUE LRU
// ---------------------------------------------------------------------------

describe('WordArt warp painted pattern cache', () => {
  it('C1: hit-promoted keys survive capacity eviction via real paint+createPattern', async () => {
    vi.resetModules()
    const { paintTextBody: fresh } = await import('../src/drawing/text-paint')
    const spy = installCountingHost()
    const paint = (fg: string): void => {
      const canvas = createCanvas(240, 130)
      const ctx = canvas.getContext('2d')
      const body = warpBody(undefined, undefined, 'AA',
        `<a:pattFill prst="dkUpDiag"><a:fgClr><a:srgbClr val="${fg}"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`)
      expect(body.paragraphs[0].runs[0].noFill ?? false, 'pattern body paints from a real parsed model').toBe(false)
      expect(body.paragraphs[0].runs[0].textFill?.kind, 'fixture parses to a pattern fill').toBe('pattern')
      fresh(body, ctx as unknown as never, MARGIN, MARGIN, 224, 114, identityFont)
    }
    paint('100000')                                     // insert A
    expect(spy.tile8, 'first uncached pattern creates one 8×8 tile').toBe(1)
    for (let i = 1; i <= 63; i++) paint((127616 + i).toString(16).padStart(6, '0'))  // fill the 64-key cache
    expect(spy.tile8, 'capacity is 64: no extra inserts').toBe(64)
    paint('100000')                                     // hit A (promoted; NO new creation)
    expect(spy.tile8).toBe(64)
    paint('200000')                                     // insert B → evicts the least-recent, NOT A
    expect(spy.tile8).toBe(65)
    paint('100000')                                     // A must still be served from the cache
    expect(spy.tile8, 'true LRU keeps promoted entries past capacity eviction').toBe(65)
    spy.restore()
    vi.resetModules()
  })

  it('W1: record-side allocation probe detects undefined scratchSurface and falls back without visual divergence', () => {
    const xml = `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"><a:prstTxWarp prst="textSlantUp"><a:avLst><a:gd name="adj" fmla="val 55555"/></a:avLst></a:prstTxWarp></a:bodyPr><a:p><a:r><a:rPr sz="3600"/><a:t>MMMM</a:t></a:r></a:p></a:txBody>`
    const model = parseTextBody(parseXmlOrdered(xml))
    const c = createCanvas(300, 200)
    const ctx: any = c.getContext('2d')
    class LimitedCanvas {
      constructor(w: number, h: number) {
        if (w > 4 || h > 4) throw new Error('bounded allocation failure')
        return createCanvas(w, h) as any
      }
    }
    Object.defineProperty(c, 'constructor', { value: LimitedCanvas })
    const records: any[] = []
    ctx[RECORD_TEXT] = (text: string, x: number, y: number, width: number, logical: any) =>
      records.push({ text, hasVisual: !!logical.visual, x, y, width })

    paintTextBody(model, ctx, 0, 0, 200, 100, f => f)
    expect(records.length).toBeGreaterThan(0)
    // When scratch allocation fails, record pass MUST fall back and not emit visual warped geometry
    expect(records[0].hasVisual).toBe(false)
    expect(model.diagnostics?.some(d => d.feature === 'warp-alloc-failure')).toBe(true)
  })

  it('W2: joined-run single fillText validates physical contiguity before joining members', () => {
    const xml = `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"><a:prstTxWarp prst="textSlantUp"><a:avLst><a:gd name="adj" fmla="val 55555"/></a:avLst></a:prstTxWarp></a:bodyPr><a:p><a:r><a:rPr sz="2400"/><a:t>A</a:t></a:r><a:r><a:rPr sz="2400"/><a:t>B</a:t></a:r></a:p></a:txBody>`
    const model = parseTextBody(parseXmlOrdered(xml))
    const fillCalls: Array<{ text: string; x: number }> = []

    const c = createCanvas(300, 200)
    class TrackCanvas {
      constructor(w: number, h: number) {
        const sc = createCanvas(w, h)
        const sctx = sc.getContext('2d')
        const origFill = sctx.fillText.bind(sctx)
        sctx.fillText = (text, x, y) => {
          fillCalls.push({ text: String(text), x: Number(x) })
          origFill(text, x, y)
        }
        return sc as any
      }
    }
    Object.defineProperty(c, 'constructor', { value: TrackCanvas })
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D

    const nonContiguousLayout = {
      lines: [
        {
          x: 0, y: 0, baseline: 30, height: 40, paragraphIndex: 0,
          segments: [
            { text: 'A', x: 0, width: 20, style: model.paragraphs[0].runs[0], runIndex: 0, sourceStart: 0, sourceEnd: 1 },
            { text: 'B', x: 40, width: 20, style: model.paragraphs[0].runs[1], runIndex: 1, sourceStart: 1, sourceEnd: 2 },
          ]
        }
      ]
    }

    paintTextBody(model, ctx as unknown as never, 0, 0, 200, 100, f => f, undefined, { layout: nonContiguousLayout as any })
    expect(fillCalls.some(call => call.text === 'AB')).toBe(false)
    expect(fillCalls.map(c => ({ text: c.text, x: c.x }))).toEqual([
      { text: 'A', x: 0 },
      { text: 'B', x: 40 }
    ])
  })

  it('W-Extra: non-strict drag selection tests mesh tris and cells for warped text eligibility', () => {
    const index: any = {
      pages: [
        {
          index: 0,
          lines: [
            {
              top: 10,
              bottom: 50,
              text: 'WARP',
              spans: [
                {
                  x: 10,
                  y: 10,
                  width: 80,
                  text: 'WARP',
                  placement: {
                    x: 10,
                    y: 10,
                    width: 80,
                    top: 10,
                    bottom: 50,
                    clip: [{ x: 0, y: 0 }, { x: 100, y: 0 }, { x: 100, y: 100 }, { x: 0, y: 100 }],
                    transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
                    visual: {
                      polygon: [],
                      tris: [
                        [
                          { x: 10, y: 10 },
                          { x: 90, y: 10 },
                          { x: 50, y: 50 },
                        ],
                      ],
                      clusters: [{ x: 50, y: 30 }],
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    }

    const hit = hitTest(index, 0, 50, 20, { strict: false })
    expect(hit).toBeDefined()
    expect(hit?.lineIndex).toBe(0)
  })

  it('W-Extra: textCircle strict hitTest covers 100% of opaque ink (alpha >= 210)', async () => {
    const W = 260, H = 180
    const body = parseTextBody(parseXmlOrdered(`
      <a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="ctr">
          <a:prstTxWarp prst="textCircle"><a:avLst/></a:prstTxWarp>
        </a:bodyPr>
        <a:p>
          <a:r><a:rPr sz="3600" b="1"><a:latin typeface="Times New Roman"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>MMMM</a:t></a:r>
        </a:p>
      </a:txBody>
    `))
    const canvas = createCanvas(W, H)
    const ctx = canvas.getContext('2d')
    paintTextBody(body, ctx as unknown as never, 8, 8, 244, 164, (x: string) => x, undefined, { clip: { x: 8, y: 8, width: 244, height: 164 } })
    const data = ctx.getImageData(0, 0, W, H).data
    const ink: number[] = []
    for (let i = 0; i < data.length; i += 4) {
      if (data[i + 3] >= 210 && (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250)) ink.push(i / 4)
    }
    expect(ink.length).toBeGreaterThan(100)
    const index = await buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: (replayCtx: CanvasRenderingContext2D) => {
      paintTextBody(body, replayCtx as unknown as never, 8, 8, 244, 164, (x: string) => x, undefined, { clip: { x: 8, y: 8, width: 244, height: 164 } })
    } }])
    let hits = 0
    for (const p of ink) {
      const x = (p % W) + 0.5, y = Math.floor(p / W) + 0.5
      if (hitTest(index, 0, x, y, { strict: true })) hits++
    }
    expect(hits / ink.length).toBeGreaterThanOrEqual(0.99)
  })
})


