/**
 * Official ECMA-376 preset TEXT-warp geometry and interior envelope mapping.
 *
 * Phase 5's rigid per-glyph approximation (sine/circle formulas, per-glyph
 * transforms) is fully replaced: this module resolves each supported preset's
 * actual ordered defaults, derived guides, official handle bounds, and boundary
 * paths from the pinned official catalog (text-warp-catalog.ts, generated from
 * presetTextWarpDefinitions.xml, SHA-256 89640bd4…), reuses the shared guide
 * evaluator from geometry.ts through the TEXT-WARP seed set (w/h/l/t/r/b/hc/vc/
 * wd2/hd2/wd3/ss/cd2 — never the shape catalog), and derives a
 * source→warp mapping function F(u,v) over the whole body box.
 *
 * Mapping policy (documented, native fitting NOT verified):
 * - Source horizontal fraction u = sx / bodyWidth maps to the boundary path's
 *   own document-order fraction. Arc commands advance UNIFORM POLAR-ANGLE
 *   fractions (each polar angle then maps to its ellipse point); this is NOT
 *   identical to the authoritative matrix's parameterSamples convention,
 *   which advances uniform ELLIPSE-PARAMETER fractions (22 non-square arc
 *   cases differ at interior samples, e.g. ArchUp 200x100 adj0 t=.25:
 *   polar-angle (144.721360,94.721360) vs ellipse-parameter
 *   (170.710678,85.355339)). Both conventions trace the same official curve
 *   locus; guides/paths/bounds still match the matrix exactly. Native Office
 *   arc-length/kerned fitting is unverified; regression risk is tracked in
 *   REPORT.md.
 * - Two-boundary envelope (Wave1/2, Inflate/Deflate, SlantUp/Down,
 *   CurveUp/Down): F(u,v) = (1−v)·Top(u) + v·Bottom(u) with v = sy/bodyHeight.
 * - Follow-path (ArchUp/Down, Circle): the single boundary is the CENTER of
 *   the deformed band: F(u,v) = curve(u) + n̂(u)·(v−0.5)·h (curve passes the
 *   box's vertical center per the official moveTo; n̂ is the left-normal of the
 *   sweep-direction tangent). Native baseline-vs-centerline placement is a
 *   documented choice, unverified against Office.
 * - Authored adjustments outside the official handle bounds are clamped at
 *   entry (clamped value also diagnosed), so every guide (raw `adj` in trig
 *   guides AND pinned `adval`) stays consistent; out-of-range equivalence with
 *   literal boundary paths is not claimed.
 */
import { resolveTextWarpPaths, textWarpSeedValues, evaluateGuides, type Guide } from './geometry'
import { TEXT_WARP_CATALOG, OFFICIAL_TEXT_WARP_REFERENCE, type TextWarpDefinition, type TextWarpCommand } from './text-warp-catalog'
import type { TextWarp, TextWarpPreset } from './text'

export { OFFICIAL_TEXT_WARP_REFERENCE }

/** The eleven advertised presets stay supported; no disabling, no approximation claim. */
export const SUPPORTED_TEXT_WARP_PRESETS: ReadonlySet<TextWarpPreset> = new Set<TextWarpPreset>(
  Object.keys(TEXT_WARP_CATALOG) as TextWarpPreset[],
)

export interface WarpAxisBounds { guide: string; axis: 'Ang' | 'R' | 'Y' | 'X'; min: number; max: number }

/** Catalog axis slug → official axis label (matching the authoritative matrix). */
const AXIS_LABEL: Record<string, 'Ang' | 'R' | 'Y' | 'X'> = { angle: 'Ang', radius: 'R', y: 'Y', x: 'X' }

export type WarpResolvedCommand =
  | { kind: 'moveTo' | 'lnTo' | 'quadBezTo' | 'cubicBezTo'; values: number[] }
  | { kind: 'arcTo'; attrs: { wR: number; hR: number; stAng: number; swAng: number } }
  | { kind: 'close' }
export interface WarpResolvedPath { commands: WarpResolvedCommand[] }

export interface WarpResolution {
  guides: Record<string, number>
  bounds: ReadonlyArray<WarpAxisBounds>
  paths: ReadonlyArray<WarpResolvedPath>
  /** clamped authored adjustments (name → official-bounds value) */
  effective: Record<string, number>
  issues: ReadonlyArray<{ kind: string; guide?: string; pathIndex?: number; message: string }>
}

function handleBound(value: string | number, seed: Map<string, number>): number {
  const seeded = seed.get(String(value))
  if (seeded !== undefined) return seeded
  const n = Number(value)
  if (!Number.isFinite(n)) throw new Error(`Unresolvable handle bound ${String(value)}`)
  return n
}

/** Official per-guide bounds (ahLst), evaluated against a nominal box. */
function boundsFor(definition: TextWarpDefinition): ReadonlyArray<WarpAxisBounds> {
  const seed = textWarpSeedValues(100, 100)
  return definition.handles.map(handle => ({
    guide: handle.guide,
    axis: AXIS_LABEL[handle.axis],
    min: handleBound(handle.min, seed),
    max: handleBound(handle.max, seed),
  }))
}

/** Clamp authored values into the OFFICIAL handle bounds at entry (policy:
 * keeps pinned `adval` guides and raw `adj` trig guides consistent; literal
 * out-of-range XML equivalence is explicitly not claimed). */
export function clampAdjustment(preset: TextWarpPreset, name: string, value: number): number {
  const definition = TEXT_WARP_CATALOG[preset]
  if (!definition) return value
  const bound = definition.handles.find(h => h.guide === name)
  if (!bound) return value
  const seed = textWarpSeedValues(100, 100)
  const min = Number.isFinite(Number(bound.min)) ? Number(bound.min) : handleBoundChecked(bound.min, seed)
  const max = Number.isFinite(Number(bound.max)) ? Number(bound.max) : handleBoundChecked(bound.max, seed)
  return Math.min(Math.max(min, value), max)
}

function handleBoundChecked(token: string | number, seed: Map<string, number>): number {
  const seeded = seed.get(String(token))
  if (seeded === undefined) throw new Error(`Unresolvable handle bound ${String(token)}`)
  return seeded
}

/** Official DEFAULTS evaluated from avLst FORMULAS (not a hand table). */
export const DEFAULT_WARP_ADJUSTMENTS: Readonly<Record<string, Record<string, number>>> = Object.fromEntries(
  Object.entries(TEXT_WARP_CATALOG).map(([preset, definition]) => {
    const seed = textWarpSeedValues(100, 100)
    const values = evaluateGuides(definition.adjustments as unknown as ReadonlyArray<Guide>, definition.guides as unknown as ReadonlyArray<Guide>, seed, {}, [])
    const out: Record<string, number> = {}
    for (const [name] of definition.adjustments) {
      const value = values.get(name)
      if (value !== undefined) out[name] = value
    }
    return [preset, out]
  }),
)

/** Resolve an official text-warp definition: guides, bounds, and boundary paths
 * with authored adjustments (clamped at entry). */
export function resolveWarp(
  preset: TextWarpPreset, width: number, height: number,
  adjustments: Record<string, number> = {},
): WarpResolution | undefined {
  const definition = TEXT_WARP_CATALOG[preset]
  if (!definition) return undefined
  const issues: Array<{ kind: string; guide?: string; pathIndex?: number; message: string }> = []
  const effective: Record<string, number> = {}
  const officialDefaults = DEFAULT_WARP_ADJUSTMENTS[preset] ?? {}
  for (const handle of definition.handles) {
    const authored = adjustments[handle.guide]
    if (authored === undefined) {
      // Absent authored values use the OFFICIAL default (avLst formula) — the
      // effective table always exposes the complete set of public guides.
      const fallback = officialDefaults[handle.guide]
      if (fallback !== undefined) effective[handle.guide] = fallback
      continue
    }
    const cleaned = Number.isFinite(authored) ? authored : Number(authored)
    if (!Number.isFinite(cleaned)) { issues.push({ kind: 'invalid-guide', guide: handle.guide, message: 'Adjustment is not finite' }); continue }
    const bounds = { min: 0, max: 0 }
    try {
      const seed = textWarpSeedValues(width, height)
      bounds.min = Number.isFinite(Number(handle.min)) ? Number(handle.min) : handleBoundChecked(handle.min, seed)
      bounds.max = Number.isFinite(Number(handle.max)) ? Number(handle.max) : handleBoundChecked(handle.max, seed)
    } catch (error) {
      issues.push({ kind: 'invalid-guide', guide: handle.guide, message: error instanceof Error ? error.message : String(error) })
      continue
    }
    effective[handle.guide] = clampAdjustment(preset, handle.guide, cleaned)
    if (effective[handle.guide] !== cleaned) {
      // A DISTINCT clamped-guide diagnosis (not 'invalid-guide'): the authored
      // value was finite and in the DOUBLE range but outside the OFFICIAL
      // handle bounds — the warp still renders (documented entry-clamp policy)
      // whereas an invalid guide would suppress the whole warp.
      issues.push({ kind: 'clamped-guide', guide: handle.guide, message: `Adjustment ${cleaned} outside official bounds; clamped to ${effective[handle.guide]}` })
    }
  }
  // Passing written-through adjustments by handle name; unnamed values do not
  // participate in this catalog (official handles are the only public guides).
  const { values, paths, issues: pathIssues } = resolveTextWarpPaths(
    definition as unknown as { adjustments: ReadonlyArray<Guide>; guides: ReadonlyArray<Guide>; paths: ReadonlyArray<Parameters<typeof resolveTextWarpPaths>[0]['paths'][number]>; handles: ReadonlyArray<{ guide: string; min: string | number; max: string | number }> },
    width, height, effective, issues as never,
  )
  return { guides: values, bounds: boundsFor(definition), paths: paths as WarpResolvedPath[], effective, issues: [...issues, ...pathIssues] as never }
}

// ---------------------------------------------------------------------------
// Boundary curve model
// ---------------------------------------------------------------------------

interface CurvePoint { x: number; y: number }
interface CurveSegment {
  /** Evaluate at document parameter t in [0,1] of this segment. */
  at(t: number): CurvePoint
  /** Endpoint in document parameter space for chaining. */
  end: CurvePoint
  length: number
  cubic?: { p0: CurvePoint; p1: CurvePoint; p2: CurvePoint; p3: CurvePoint }
  arc?: { c: CurvePoint; wR: number; hR: number; stAng: number; swAng: number }
}

function cubicAt(p0: CurvePoint, p1: CurvePoint, p2: CurvePoint, p3: CurvePoint, t: number): CurvePoint {
  const s = 1 - t
  const w0 = s * s * s, w1 = 3 * s * s * t, w2 = 3 * s * t * t, w3 = t * t * t
  return { x: w0 * p0.x + w1 * p1.x + w2 * p2.x + w3 * p3.x, y: w0 * p0.y + w1 * p1.y + w2 * p2.y + w3 * p3.y }
}

const TAU = Math.PI * 2
const TO_RADIANS = Math.PI / 10800000

/** DrawingML polar angle → ellipse parameter, preserving full turns. */
function ellipseParameter(angleRad: number, wR: number, hR: number): number {
  const base = Math.atan2(wR * Math.sin(angleRad), hR * Math.cos(angleRad))
  return base + Math.round((angleRad - base) / TAU) * TAU
}

function ellipsePoint(c: CurvePoint, wR: number, hR: number, polarRad: number): CurvePoint {
  const param = ellipseParameter(polarRad, wR, hR)
  return { x: c.x + wR * Math.cos(param), y: c.y + hR * Math.sin(param) }
}

function buildSegments(commands: ReadonlyArray<WarpResolvedCommand>): CurveSegment[] {
  const segments: CurveSegment[] = []
  let current: CurvePoint | undefined
  let start: CurvePoint | undefined
  for (const command of commands) {
    if (command.kind === 'moveTo') {
      current = { x: command.values[0], y: command.values[1] }
      start = current
      continue
    }
    if (command.kind === 'close') { current = start; continue }
    if (!current) throw new Error('Boundary path segment without a current point')
    if (command.kind === 'lnTo') {
      const from = current
      const end = { x: command.values[0], y: command.values[1] }
      segments.push(lineSegment(from, end))
      current = end
    } else if (command.kind === 'quadBezTo' || command.kind === 'cubicBezTo') {
      const from = current
      const v = command.values
      const end = { x: v[v.length - 2], y: v[v.length - 1] }
      const p1 = command.kind === 'cubicBezTo'
        ? { x: v[0], y: v[1] }
        : { x: from.x + (2 / 3) * (v[0] - from.x), y: from.y + (2 / 3) * (v[1] - from.y) }
      const p2 = command.kind === 'cubicBezTo'
        ? { x: v[2], y: v[3] }
        : { x: end.x + (2 / 3) * (v[2] - end.x), y: end.y + (2 / 3) * (v[3] - end.y) }
      segments.push({ at: t => cubicAt(from, p1, p2, end, t), end, length: cubicLength(from, p1, p2, end), cubic: { p0: from, p1, p2, p3: end } })
      current = end
    } else if (command.kind === 'arcTo') {
      const { wR, hR, stAng, swAng } = command.attrs
      // Arc center from the official convention: current − (wR·cos st, hR·sin st)
      // at the ellipse-parameter start.
      const from = current
      const startParam = ellipseParameter(stAng * TO_RADIANS, wR, hR)
      const center = { x: from.x - wR * Math.cos(startParam), y: from.y - hR * Math.sin(startParam) }
      const arcEnd = ellipsePoint(center, wR, hR, (stAng + swAng) * TO_RADIANS)
      segments.push({ at: t => ellipsePoint(center, wR, hR, (stAng + t * swAng) * TO_RADIANS), end: arcEnd, length: Math.abs(swAng) * TO_RADIANS * Math.hypot(wR, hR), arc: { c: center, wR, hR, stAng, swAng } })
      current = arcEnd
    }
  }
  // NOTE: text-warp boundary paths are OPEN curves (envelope top/bottom and
  // follow-path arcs); no implicit close is appended. Text-warp path
  // parameterization follows document order exactly as the official pathLst
  // spans it.
  return segments
}

function lineSegment(from: CurvePoint, to: CurvePoint): CurveSegment {
  return { at: t => ({ x: from.x + (to.x - from.x) * t, y: from.y + (to.y - from.y) * t }), end: to, length: Math.hypot(to.x - from.x, to.y - from.y) }
}

function cubicLength(p0: CurvePoint, p1: CurvePoint, p2: CurvePoint, p3: CurvePoint): number {
  // Control-polygon length is a bounded overestimate good enough for segment weighting.
  return Math.hypot(p1.x - p0.x, p1.y - p0.y) + Math.hypot(p2.x - p1.x, p2.y - p1.y) + Math.hypot(p3.x - p2.x, p3.y - p2.y)
}

/** A boundary path flattened into document-parameter space (t fraction across
 * its commands, weighted by each command's control-polygon length). */
interface WarpCurve {
  /** Point at parameter t ∈ [0,1]. */
  at(t: number): CurvePoint
  /** Unit tangent (sweep direction) at t. */
  tangent(t: number): CurvePoint
  /** Exact segment (cubic) checkpoint at t, for matrix parameterSamples probes. */
  checkpoint(t: number): CurvePoint
  length: number
  segmentCount: number
}

function buildCurve(commands: ReadonlyArray<WarpResolvedCommand>): WarpCurve {
  const segments = buildSegments(commands)
  if (!segments.length) throw new Error('Warp boundary path has no segments')
  return curveFromSegments(segments)
}

function curveFromSegments(segments: CurveSegment[]): WarpCurve {
  const weights = segments.map(s => Math.max(s.length, 1e-9))
  const total = weights.reduce((a, b) => a + b, 0)
  const starts: number[] = []
  let acc = 0
  for (const w of weights) { starts.push(acc / total); acc += w }
  const locate = (t: number): { index: number; local: number } => {
    const tt = Math.min(1, Math.max(0, t))
    let index = 0
    while (index < segments.length - 1 && (starts[index + 1] ?? 1) < tt) index++
    const span = (starts[index + 1] ?? 1) - starts[index]
    return { index, local: span > 0 ? (tt - starts[index]) / span : 0 }
  }
  return {
    at(t) { const { index, local } = locate(t); return segments[index].at(local) },
    tangent(t) {
      const { index, local } = locate(t)
      const eps = 1e-4
      const a = segments[index].at(Math.max(0, local - eps))
      const b = segments[index].at(Math.min(1, local + eps))
      const dx = b.x - a.x, dy = b.y - a.y
      const len = Math.hypot(dx, dy) || 1
      return { x: dx / len, y: dy / len }
    },
    checkpoint(t) { const { index, local } = locate(t); return segments[index].at(local) },
    length: total,
    segmentCount: segments.length,
  }
}

// ---------------------------------------------------------------------------
// Source → warp mapping
// ---------------------------------------------------------------------------

export type WarpFamily = 'envelope' | 'follow-path'

export interface WarpMapping {
  readonly preset: TextWarpPreset
  readonly family: WarpFamily
  readonly width: number
  readonly height: number
  /** True when F is numerically indistinguishable from the identity over the
   * box (e.g. SlantUp adj=0, Wave adj1=0&adj2=0). Paint uses the ordinary path
   * then — shaping/gradient behavior is exactly the unwarped one. */
  readonly identity: boolean
  /** Map a source point expressed in BODY fractions u,v ∈ [0,1]. */
  map(u: number, v: number): CurvePoint
  /** Boundary curve accessor (path index) at parameter t (document fraction). */
  boundaryAt(pathIndex: number, t: number): CurvePoint
  /** Tangent (sweep direction) of a boundary path at t. */
  boundaryTangent(pathIndex: number, t: number): CurvePoint
  readonly curves: ReadonlyArray<WarpCurve>
  readonly resolution: WarpResolution
}

function deviationOf(mapping: WarpMapping, u: number, v: number): number {
  const mapped = mapping.map(u, v)
  return Math.hypot(mapped.x - u * mapping.width, mapped.y - v * mapping.height)
}

function isIdentity(mapping: WarpMapping): boolean {
  for (let i = 0; i <= 8; i++) {
    for (let j = 0; j <= 8; j++) {
      if (deviationOf(mapping, i / 8, j / 8) > 1e-6) return false
    }
  }
  return true
}

/** Build the warp mapping for a body box from the official catalog. */
export function buildWarpMapping(warp: TextWarp, width: number, height: number): WarpMapping {
  const resolution = resolveWarp(warp.preset, width, height, warp.adjustments ?? {})
  if (!resolution) throw new Error(`Unsupported warp preset ${String(warp.preset)}`)
  const curves = resolution.paths.map(path => buildCurve(path.commands))
  const definition = TEXT_WARP_CATALOG[warp.preset]
  const family: WarpFamily = definition.paths.length >= 2 ? 'envelope' : 'follow-path'

  if (family === 'envelope') {
    const top = curves[0], bottom = curves[1] ?? curves[0]
    const mapping: WarpMapping = {
      preset: warp.preset, family, width, height, identity: false,
      resolution,
      curves,
      map(u, v) {
        const a = top.at(u), b = bottom.at(u)
        return { x: a.x + (b.x - a.x) * v, y: a.y + (b.y - a.y) * v }
      },
      boundaryAt: (pathIndex, t) => curves[pathIndex].at(t),
      boundaryTangent: (pathIndex, t) => curves[pathIndex].tangent(t),
    }
    return { ...mapping, identity: isIdentity(mapping) }
  }

  const curve = curves[0]
  const mapping: WarpMapping = {
    preset: warp.preset, family, width, height, identity: false,
    resolution,
    curves,
    map(u, v) {
      const t = u
      const p = curve.at(t)
      const tan = curve.tangent(t)
      // Left-normal of the sweep-direction tangent: for a positive sweep over
      // the top (ArchUp) this places v<0 outward (above the apex); for a
      // negative sweep (ArchDown bottom arc, L→R flow) it places v<0 toward
      // the center — both match Office's visual orientation conventions.
      const n = { x: -tan.y, y: tan.x }
      const offset = (v - 0.5) * height
      return { x: p.x + n.x * offset, y: p.y + n.y * offset }
    },
    boundaryAt: (_pathIndex, t) => curve.at(t),
    boundaryTangent: (_pathIndex, t) => curve.tangent(t),
  }
  return { ...mapping, identity: isIdentity(mapping) }
}

/** Adaptive mesh lattice over a body-fraction region. Points are mapped body
 * positions; cols/rows bounded so each cell carries at most one curve wax/wane. */
export interface WarpLattice {
  /** (cols+1) × (rows+1) mapped points, row-major from region (u0,v0). */
  points: ReadonlyArray<ReadonlyArray<CurvePoint>>
  /** Source-space lattice parameter ranges. */
  region: { u0: number; u1: number; v0: number; v1: number }
  cols: number
  rows: number
}

/** Mesh density bounded by cell size (≤ 8px per cell in source space), max 64×64. */
export function buildWarpLattice(
  mapping: WarpMapping,
  region: { u0: number; u1: number; v0: number; v1: number },
): WarpLattice {
  const sourceW = Math.max(1e-6, (region.u1 - region.u0) * mapping.width)
  const sourceH = Math.max(1e-6, (region.v1 - region.v0) * mapping.height)
  const cols = Math.max(1, Math.min(64, Math.ceil(sourceW / 8)))
  const rows = Math.max(1, Math.min(64, Math.ceil(sourceH / 8)))
  const points: Array<Array<CurvePoint>> = []
  for (let j = 0; j <= rows; j++) {
    const row: Array<CurvePoint> = []
    for (let i = 0; i <= cols; i++) {
      const u = region.u0 + ((region.u1 - region.u0) * i) / cols
      const v = region.v0 + ((region.v1 - region.v0) * j) / rows
      row.push(mapping.map(u, v))
    }
    points.push(row)
  }
  return { points, region, cols, rows }
}

// ---------------------------------------------------------------------------
// Compatibility surface (kept for parse-side callers and tests)
// ---------------------------------------------------------------------------

export interface WarpCommandCursor { kind: string; values?: ReadonlyArray<string | number>; attrs?: Partial<Record<'wR' | 'hR' | 'stAng' | 'swAng', string | number>> }
export type { TextWarpDefinition, TextWarpCommand }
