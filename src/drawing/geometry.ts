/** DrawingML geometry evaluated in shape-local pixels. Placement belongs to adapters. */
import { attrs, getChild, type XmlNode } from '../core/xml'
import { parseCustomGeometry, parseGuideList } from './geometry-xml'
import presetData from './presets.json'

export type Guide = [name: string, formula: string]
export type GeometryCommand = [kind: string, ...values: string[]]
export interface GeometryPath {
  width?: string
  height?: string
  fill?: string
  stroke?: boolean
  commands: GeometryCommand[]
}
export interface GeometryDefinition {
  /** The original preset name, including unrecognized names. Absent for custom geometry. */
  preset?: string
  adjustments: Guide[]
  guides: Guide[]
  paths: GeometryPath[]
  textRect?: string[]
}
export type ResolvedCommand =
  | ['moveTo', number, number]
  | ['lnTo', number, number]
  | ['quadBezTo', number, number, number, number]
  | ['cubicBezTo', number, number, number, number, number, number]
  /** Canvas ellipse arguments: cx, cy, rx, ry, start/end parameter radians, anticlockwise. */
  | ['arcTo', number, number, number, number, number, number, boolean]
  | ['close']
export interface ResolvedPath {
  commands: ResolvedCommand[]
  fill: string
  stroke: boolean
}
export interface GeometryIssue {
  kind: 'unknown-preset' | 'invalid-guide' | 'invalid-path' | 'invalid-text-rect' | 'invalid-extents'
  message: string
  pathIndex?: number
  guide?: string
}
export interface ResolvedGeometry {
  paths: ResolvedPath[]
  textRect?: { left: number; top: number; right: number; bottom: number }
  issues: GeometryIssue[]
}

const presets = presetData as unknown as Record<string, GeometryDefinition>
const emptyDefinition = (): GeometryDefinition => ({ adjustments: [], guides: [], paths: [] })
const hasPreset = (name: string): boolean => Object.prototype.hasOwnProperty.call(presets, name)

/**
 * Use XML parsed with parseXmlOrdered when commands repeat/interleave: ordinary
 * XML object trees cannot recover source order once like-named siblings group.
 */
export function parseGeometry(spPr: XmlNode | undefined): GeometryDefinition {
  const custom = getChild(spPr, 'custGeom')
  if (custom) return parseCustomGeometry(custom)
  const preset = getChild(spPr, 'prstGeom')
  const name = preset ? attrs(preset).prst ?? '' : 'rect'
  const source = hasPreset(name) ? presets[name] : emptyDefinition()
  const overrides = parseGuideList(getChild(preset, 'avLst'))
  const overriddenNames = new Set(overrides.map(([name]) => name))
  // Dependencies between replacement formulas follow document order. Preset
  // defaults for names the document does not replace remain available first.
  const adjustments: Guide[] = [
    ...source.adjustments.filter(([name]) => !overriddenNames.has(name)).map(guide => [...guide] as Guide),
    ...overrides.map(guide => [...guide] as Guide),
  ]
  // Definitions are treated as immutable by resolution; arrays here are copied
  // so parsing another shape cannot replace catalog defaults.
  return {
    ...source,
    preset: name,
    adjustments,
    guides: source.guides.map(guide => [...guide] as Guide),
    paths: source.paths.map(path => ({ ...path, commands: path.commands.map(command => [...command] as GeometryCommand) })),
    ...(source.textRect ? { textRect: [...source.textRect] } : {}),
  }
}

const CIRCLE = 21600000
const TAU = Math.PI * 2
const TO_RADIANS = Math.PI / 10800000
const numericToken = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/

function finite(value: number, label: string): number {
  if (!Number.isFinite(value)) throw new Error(`Nonfinite ${label}`)
  return value
}

function tokenValue(token: string, values: Map<string, number>): number {
  if (numericToken.test(token)) return finite(Number(token), `number ${token}`)
  const known = values.get(token)
  if (known !== undefined) return finite(known, `guide ${token}`)
  const fraction = /^(w|h|ss|ls)d([1-9]\d*)$/.exec(token)
  if (fraction) return finite(values.get(fraction[1])! / Number(fraction[2]), `fraction ${token}`)
  const angle = /^(\d+)?cd([1-9]\d*)$/.exec(token)
  if (angle) return finite((Number(angle[1] ?? 1) * CIRCLE) / Number(angle[2]), `angle ${token}`)
  throw new Error(`Unknown geometry token ${JSON.stringify(token)}`)
}

function formulaValue(formula: string, values: Map<string, number>): number {
  const [op, ...tokens] = formula.trim().split(/\s+/)
  const counts: Record<string, number> = { val: 1, '+-': 3, '*/': 3, '+/': 3, '?:': 3, abs: 1, at2: 2, cat2: 3, cos: 2, max: 2, min: 2, mod: 3, pin: 3, sat2: 3, sin: 2, sqrt: 1, tan: 2 }
  if (!Object.prototype.hasOwnProperty.call(counts, op) || tokens.length !== counts[op]) throw new Error(`Invalid guide formula ${JSON.stringify(formula)}`)
  const [x, y, z] = tokens.map(token => tokenValue(token, values))
  let result: number
  switch (op) {
    case 'val': result = x; break
    case '+-': result = x + y - z; break
    case '*/': result = x * y / z; break
    case '+/': result = (x + y) / z; break
    case '?:': result = x > 0 ? y : z; break
    case 'abs': result = Math.abs(x); break
    case 'at2': result = Math.atan2(y, x) / TO_RADIANS; break
    case 'cat2': result = x * Math.cos(Math.atan2(z, y)); break
    case 'cos': result = x * Math.cos(y * TO_RADIANS); break
    case 'max': result = Math.max(x, y); break
    case 'min': result = Math.min(x, y); break
    case 'mod': result = Math.hypot(x, y, z); break
    case 'pin': result = y < x ? x : y > z ? z : y; break
    case 'sat2': result = x * Math.sin(Math.atan2(z, y)); break
    case 'sin': result = x * Math.sin(y * TO_RADIANS); break
    case 'sqrt': result = Math.sqrt(x); break
    case 'tan': result = x * Math.tan(y * TO_RADIANS); break
    default: throw new Error(`Unknown guide operation ${op}`)
  }
  return finite(result, `formula ${formula}`)
}

/** Polar angle -> ellipse parameter angle, preserving turns and sweep direction. */
function ellipseParameter(angle: number, rx: number, ry: number): number {
  const base = Math.atan2(rx * Math.sin(angle), ry * Math.cos(angle))
  return base + Math.round((angle - base) / TAU) * TAU
}

/** Text-warp seed guide set, per ECMA-376 presetTextWarpDefinitions context.
 * Shape geometry uses a different seed set (ls, cd) — routing a text-warp
 * definition through the shape catalog evaluates the wrong values. */
export function textWarpSeedValues(width: number, height: number): Map<string, number> {
  const values = new Map<string, number>(Object.entries({
    l: 0, t: 0, r: width, b: height, w: width, h: height, hc: width / 2, vc: height / 2,
    wd2: width / 2, hd2: height / 2, wd3: width / 3, ss: Math.min(width, height), cd2: CIRCLE / 2,
  }))
  return values
}

/** Evaluate an ordered guide list (optionally honoring overrides for adjustment
 * guides) into the given seed map. Reuses the shared evaluator so text warps and
 * shape geometry agree on every operation and token. */
export function evaluateGuides(
  adjustments: ReadonlyArray<Guide>, guides: ReadonlyArray<Guide>, seed: Map<string, number>,
  overrides: Record<string, number>, issues: Array<{ kind: string; guide?: string; message: string }>,
): Map<string, number> {
  const values = seed
  for (const [guides_, applyOverrides] of [[adjustments, true], [guides, false]] as const) {
    for (const [name, formula] of guides_) {
      try {
        if (!name) throw new Error('Missing guide name')
        const value = applyOverrides && Object.prototype.hasOwnProperty.call(overrides, name)
          ? finite(overrides[name], `adjustment ${name}`)
          : formulaValue(formula, values)
        values.set(name, value)
      } catch (error) {
        values.delete(name)
        issues.push({ kind: 'invalid-guide', guide: name, message: messageOf(error) })
      }
    }
  }
  return values
}

/** Official resolved command form for TEXT-warp boundary paths: point commands
 * carry numeric pixel values; arcTo keeps its DrawingML angle attributes. */
export type WarpResolvedCommand =
  | { kind: 'moveTo' | 'lnTo' | 'quadBezTo' | 'cubicBezTo'; values: number[] }
  | { kind: 'arcTo'; attrs: { wR: number; hR: number; stAng: number; swAng: number } }
  | { kind: 'close' }
export interface WarpResolvedPath { commands: WarpResolvedCommand[] }
export interface WarpResolvedGeometry { values: Record<string, number>; paths: WarpResolvedPath[]; issues: GeometryIssue[] }

function resolveWarpPath(path: GeometryPath, values: Map<string, number>): WarpResolvedPath {
  // Text-warp catalog commands arrive as {kind, values|attrs} objects (from the
  // pinned catalog), NOT the tuple form geometry-xml produces. Normalize here.
  const catalogLike = path.commands as unknown as ReadonlyArray<{ kind: string; values?: readonly (string | number)[]; attrs?: Partial<Record<'wR' | 'hR' | 'stAng' | 'swAng', string | number>> }>
  const sx = 1
  const sy = 1
  const commands: WarpResolvedCommand[] = []
  let current: [number, number] | undefined
  let subpath: [number, number] | undefined
  for (const catalogCommand of catalogLike) {
    const kind = catalogCommand.kind
    if (kind === 'moveTo') {
      const raw = [tokenValue(String(catalogCommand.values?.[0] ?? '0'), values), tokenValue(String(catalogCommand.values?.[1] ?? '0'), values)]
      current = [raw[0], raw[1]]; subpath = current
      commands.push({ kind: 'moveTo', values: [finite(current[0] * sx, 'x'), finite(current[1] * sy, 'y')] })
      continue
    }
    if (!current) throw new Error(`${kind} without a current point`)
    if (kind === 'close') {
      commands.push({ kind: 'close' })
      current = subpath
    } else if (kind === 'arcTo') {
      // Official form: keep DrawingML angle attrs, trace the arc endpoint.
      const attrs = (catalogCommand.attrs ?? {}) as Partial<Record<string, string | number>>
      if (attrs.wR === undefined || attrs.hR === undefined || attrs.stAng === undefined || attrs.swAng === undefined) throw new Error('arcTo needs wR/hR/stAng/swAng')
      const raw = [tokenValue(String(attrs.wR), values), tokenValue(String(attrs.hR), values), tokenValue(String(attrs.stAng), values), tokenValue(String(attrs.swAng), values)]
      const [rx, ry, stAng, swAng] = raw as [number, number, number, number]
      if (rx < 0 || ry < 0) throw new Error('Arc radii must be nonnegative')
      commands.push({ kind: 'arcTo', attrs: { wR: finite(rx * sx, 'wR'), hR: finite(ry * sy, 'hR'), stAng: finite(stAng, 'stAng'), swAng: finite(swAng, 'swAng') } })
      current = continueWarpArc(current, rx, ry, stAng, swAng)
    } else {
      const valuesList = catalogCommand.values ?? []
      const raw = [tokenValue(String(valuesList[0] ?? '0'), values), tokenValue(String(valuesList[1] ?? '0'), values), tokenValue(String(valuesList[2] ?? '0'), values), tokenValue(String(valuesList[3] ?? '0'), values), tokenValue(String(valuesList[4] ?? '0'), values), tokenValue(String(valuesList[5] ?? '0'), values)]
      const degreesToTokens: Record<string, number> = { lnTo: 2, quadBezTo: 4, cubicBezTo: 6 }
      if (!Object.prototype.hasOwnProperty.call(degreesToTokens, kind)) throw new Error(`Invalid path command ${kind}`)
      const n = degreesToTokens[kind]
      const scaled = raw.slice(0, n).map((value, index) => finite(value * (index % 2 === 0 ? sx : sy), 'path coordinate'))
      commands.push({ kind: kind as 'moveTo' | 'lnTo' | 'quadBezTo' | 'cubicBezTo', values: scaled as unknown as number[] })
      current = [raw[n - 2], raw[n - 1]]
    }
  }
  if (!commands.length) throw new Error('Empty geometry path')
  return { commands }
}

/** Trace the implicit arc endpoint the official arcTo semantics imply so the
 * path cursor continues exactly as the DrawingML engine would. */
function continueWarpArc(
  current: [number, number],
  rx: number, ry: number, stAng: number, swAng: number,
): [number, number] {
  const start = ellipseParameter(stAng * TO_RADIANS, rx, ry)
  const cx = current[0] - rx * Math.cos(start)
  const cy = current[1] - ry * Math.sin(start)
  const end = swAng % CIRCLE === 0 ? start + swAng / CIRCLE * TAU : ellipseParameter((stAng + swAng) * TO_RADIANS, rx, ry)
  return [finite(cx + rx * Math.cos(end), 'arc endpoint x'), finite(cy + ry * Math.sin(end), 'arc endpoint y')]
}


/** Resolve an official TEXT-warp definition (avLst/gdLst/pathLst, guide-reference
 * tokens intact) against a body-local box. Adjustment overrides replace avLst
 * values by name; everything evaluates through the shared guide evaluator with
 * the TEXT-WARP seed set (never the shape seed, never the shape catalog).
 * Handles (official adjustment bounds) also evaluate so callers can state
 * clamp policies in OFFICIAL units. */
export function resolveTextWarpPaths(
  definition: { adjustments: ReadonlyArray<Guide>; guides: ReadonlyArray<Guide>; paths: ReadonlyArray<GeometryPath>; handles?: ReadonlyArray<{ guide: string; min: string | number; max: string | number }> },
  width: number, height: number,
  overrides: Record<string, number> = {},
  issues: GeometryIssue[] = [],
): WarpResolvedGeometry {
  const paths: WarpResolvedPath[] = []
  const pushIssue = (issue: GeometryIssue) => { issues.push({ ...issue, ...(issue.pathIndex !== undefined ? { pathIndex: issue.pathIndex } : {}) }) }
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    pushIssue({ kind: 'invalid-extents', message: 'Text-warp extents must be finite and positive' })
    return { values: {}, paths, issues }
  }
  // The seed map carries the seed values themselves (l/t/r/b/w/h/hc/vc/wd2/hd2/
  // wd3/ss/cd2) as evaluated guides; the caller reads the FULL table back.
  const values = textWarpSeedValues(width, height)
  try {
    evaluateGuides(definition.adjustments, definition.guides, values, overrides, issues as GeometryIssue[])
    definition.paths.forEach((path, pathIndex) => {
      try { paths.push(resolveWarpPath(path, values)) }
      catch (error) { pushIssue({ kind: 'invalid-path', pathIndex, message: messageOf(error) }) }
    })
  } catch (error) {
    pushIssue({ kind: 'invalid-path', message: messageOf(error) })
  }
  const flat: Record<string, number> = {}
  for (const [name, value] of values) flat[name] = value
  return { values: flat, paths, issues }
}

function resolvePath(path: GeometryPath, width: number, height: number, pathWidth: number, pathHeight: number, values: Map<string, number>): ResolvedPath {
  const sx = path.width === undefined ? 1 : width / pathWidth
  const sy = path.height === undefined ? 1 : height / pathHeight
  const commands: ResolvedCommand[] = []
  let current: [number, number] | undefined
  let subpath: [number, number] | undefined
  const counts: Record<string, number> = { moveTo: 2, lnTo: 2, quadBezTo: 4, cubicBezTo: 6, arcTo: 4, close: 0 }
  for (const [kind, ...tokens] of path.commands) {
    if (!Object.prototype.hasOwnProperty.call(counts, kind) || tokens.length !== counts[kind]) throw new Error(`Invalid path command ${kind}`)
    const raw = tokens.map(token => tokenValue(token, values))
    if (kind === 'moveTo') {
      current = [raw[0], raw[1]]
      subpath = current
      commands.push(['moveTo', finite(current[0] * sx, 'x'), finite(current[1] * sy, 'y')])
      continue
    }
    if (!current) throw new Error(`${kind} without a current point`)
    if (kind === 'close') {
      commands.push(['close'])
      current = subpath
    } else if (kind === 'arcTo') {
      const [rx, ry, startAngle, sweepAngle] = raw
      if (rx < 0 || ry < 0) throw new Error('Arc radii must be nonnegative')
      const start = ellipseParameter(startAngle * TO_RADIANS, rx, ry)
      const end = sweepAngle % CIRCLE === 0 ? start + sweepAngle / CIRCLE * TAU : ellipseParameter((startAngle + sweepAngle) * TO_RADIANS, rx, ry)
      const cx = current[0] - rx * Math.cos(start)
      const cy = current[1] - ry * Math.sin(start)
      commands.push(['arcTo', finite(cx * sx, 'arc center x'), finite(cy * sy, 'arc center y'), finite(rx * sx, 'arc radius x'), finite(ry * sy, 'arc radius y'), finite(start, 'arc start'), finite(end, 'arc end'), sweepAngle < 0])
      current = [finite(cx + rx * Math.cos(end), 'arc endpoint x'), finite(cy + ry * Math.sin(end), 'arc endpoint y')]
    } else {
      const scaled = raw.map((value, index) => finite(value * (index % 2 === 0 ? sx : sy), 'path coordinate'))
      if (kind === 'lnTo') commands.push(['lnTo', scaled[0], scaled[1]])
      else if (kind === 'quadBezTo') commands.push(['quadBezTo', scaled[0], scaled[1], scaled[2], scaled[3]])
      else commands.push(['cubicBezTo', scaled[0], scaled[1], scaled[2], scaled[3], scaled[4], scaled[5]])
      current = [raw[raw.length - 2], raw[raw.length - 1]]
    }
  }
  if (!commands.length) throw new Error('Empty geometry path')
  return { commands, fill: path.fill ?? 'norm', stroke: path.stroke ?? true }
}

const messageOf = (error: unknown): string => error instanceof Error ? error.message : String(error)

/** Evaluate adjustments and ordered guides in one path's unscaled coordinate space. */
function guideValues(definition: GeometryDefinition, width: number, height: number, overrides: Record<string, number>, issues: GeometryIssue[], pathIndex?: number): Map<string, number> {
  const values = new Map<string, number>(Object.entries({ l: 0, t: 0, r: width, b: height, w: width, h: height, hc: width / 2, vc: height / 2, ss: Math.min(width, height), ls: Math.max(width, height), cd: CIRCLE }))
  for (const [guides, applyOverrides] of [[definition.adjustments, true], [definition.guides, false]] as const) for (const [name, formula] of guides) {
    try {
      if (!name) throw new Error('Missing guide name')
      const value = applyOverrides && Object.prototype.hasOwnProperty.call(overrides, name) ? finite(overrides[name], `adjustment ${name}`) : formulaValue(formula, values)
      values.set(name, value)
    } catch (error) {
      values.delete(name)
      issues.push({ kind: 'invalid-guide', guide: name, message: messageOf(error), ...(pathIndex !== undefined ? { pathIndex } : {}) })
    }
  }
  return values
}

export function resolveGeometry(definition: GeometryDefinition, width: number, height: number, overrides: Record<string, number> = {}): ResolvedGeometry {
  const result: ResolvedGeometry = { paths: [], issues: [] }
  if (!Number.isFinite(width) || !Number.isFinite(height) || width < 0 || height < 0) {
    result.issues.push({ kind: 'invalid-extents', message: 'Geometry extents must be finite and nonnegative' })
    return result
  }
  if (definition.preset !== undefined && !hasPreset(definition.preset)) {
    result.issues.push({ kind: 'unknown-preset', message: `Unknown preset ${JSON.stringify(definition.preset)}` })
    return result
  }
  definition.paths.forEach((path, pathIndex) => {
    try {
      // Path extents are numeric DrawingML coordinates, not guide references.
      const pathWidth = path.width === undefined ? width : tokenValue(path.width, new Map())
      const pathHeight = path.height === undefined ? height : tokenValue(path.height, new Map())
      if (path.width !== undefined && pathWidth <= 0 || path.height !== undefined && pathHeight <= 0) throw new Error('Path coordinate space must be positive')
      const values = guideValues(definition, pathWidth, pathHeight, overrides, result.issues, pathIndex)
      result.paths.push(resolvePath(path, width, height, pathWidth, pathHeight, values))
    }
    catch (error) { result.issues.push({ kind: 'invalid-path', pathIndex, message: messageOf(error) }) }
  })
  if (definition.textRect) {
    try {
      if (definition.textRect.length !== 4) throw new Error('Text rectangle needs four coordinates')
      const values = guideValues(definition, width, height, overrides, result.issues)
      const [left, top, right, bottom] = definition.textRect.map(token => tokenValue(token, values))
      result.textRect = { left, top, right, bottom }
    } catch (error) { result.issues.push({ kind: 'invalid-text-rect', message: messageOf(error) }) }
  }
  return result
}

export function resolvePreset(name: string, width: number, height: number, overrides?: Record<string, number>): ResolvedGeometry {
  const definition = hasPreset(name) ? presets[name] : emptyDefinition()
  return resolveGeometry({ ...definition, preset: name }, width, height, overrides)
}
