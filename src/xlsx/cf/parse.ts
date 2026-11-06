/**
 * Resident + streaming-shared CF rule parser (spec §§3, 5.1–5.6, 6.2, 6.5).
 *
 * Parses all 18 rule types from base `<conditionalFormatting>` and
 * `x14:conditionalFormattings` into normalized `CfRule` entries in
 * `(priority, documentOrder)` order. Stage-1 families with literal operands or
 * absent/canonical formulas become evaluable definitions; every future-family,
 * non-literal, non-canonical, invalid or over-cap rule becomes an explicit
 * unresolved definition with accurate write masks, never a guess.
 *
 * Memory discipline: a single working reservation is admitted BEFORE any
 * normalized copy is built and grown ahead of each element/rule; the finished
 * `CfSheetRules` is adopted into a retained lease. Count/text caps are checked
 * before copying. Numeric policy-cap overflow (rule/range counts, scope,
 * unorderable extensions) yields a sheet-wide sentinel; per-rule bound overflow
 * yields a rule-level refusal. Byte exhaustion throws fail-stop (the `parse.ts`
 * integration converts it to a sheet diagnostic with no CF attached), never a
 * partial model.
 */
import { attrs, elementNamespaceUri, getChildren, textOf, type XmlNode } from '../../core/xml'
import type { SheetColorContext } from '../sheetColors'
import { parseCfDxf } from './dxf'
import {
  CF_LIMITS,
  createCfCaps,
  createCfDiagnostics,
  createCfMemory,
  type CfMemory,
  type CfReservation,
} from './limits'
import type {
  CfCellState,
  CfDefinition,
  CfLease,
  CfLiteral,
  CfOperator,
  CfRule,
  CfRuleOrigin,
  CfRuleType,
  CfSheetRules,
  CfTimePeriod,
  CfWriteMask,
} from './model'
import { X14_MAIN_NS } from '../tableStyles/parseStylesExtras'

export const XM_MAIN_NS = 'http://schemas.microsoft.com/office/excel/2006/main'
export const SPREADSHEETML_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
/** Worksheet `ext` URI carrying `x14:conditionalFormattings`. */
export const CF_X14_EXTENSION_URI = '{78C0D931-6437-407d-A8EE-F0AA7539E650}'
/** Rule `ext` URI carrying the data-bar pairing `<x14:id>`. */
export const DB_X14_EXTENSION_URI = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'

export interface CfStylesSource {
  readonly colors: SheetColorContext
  /** Original `<dxf>` nodes in `dxfs` order (for strict inspection + indexed lookup). */
  readonly dxfNodes: readonly XmlNode[]
}

export interface CfParseMeta {
  readonly sheetId?: string
  readonly partPath?: string
}

/** Workbook-shared parse account, keyed by the per-workbook styles object. */
const sharedMemories = new WeakMap<object, CfMemory>()
/** Retained handles (caps, diagnostics) keyed by their adopted lease. */
const retainedHandles = new WeakMap<object, unknown>()
/** Adopted rule leases keyed by their sheet. */
const liveSheets = new WeakMap<object, CfLease<CfSheetRules>>()

export function cfMemoryFor(scopeKey: object): CfMemory {
  let memory = sharedMemories.get(scopeKey)
  if (!memory) {
    memory = createCfMemory({})
    sharedMemories.set(scopeKey, memory)
  }
  return memory
}

/** Attach parsed rules to a sheet and retain the lease (and its charges) with it. */
export function attachCfSheetRules(sheet: { conditionalFormatting?: CfSheetRules }, lease: CfLease<CfSheetRules>): void {
  sheet.conditionalFormatting = lease.value
  liveSheets.set(sheet, lease)
}

interface ParsedRange {
  firstRow: number
  lastRow: number
  firstCol: number
  lastCol: number
}

const MAX_COL = 16384
const MAX_ROW = 1048576

const RELATIONAL_OPERATORS: ReadonlySet<string> = new Set([
  'equal',
  'notEqual',
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual',
  'between',
  'notBetween',
])
const CELLIS_TEXT_OPERATORS: ReadonlySet<string> = new Set(['containsText', 'notContains', 'beginsWith', 'endsWith'])
const VISUAL_TYPES: ReadonlySet<string> = new Set(['colorScale', 'dataBar', 'iconSet'])
const PERIODS: ReadonlySet<string> = new Set([
  'today',
  'yesterday',
  'tomorrow',
  'last7Days',
  'thisWeek',
  'lastWeek',
  'nextWeek',
  'thisMonth',
  'lastMonth',
  'nextMonth',
])
const KNOWN_TYPES: ReadonlySet<string> = new Set([
  'cellIs',
  'containsText',
  'notContainsText',
  'beginsWith',
  'endsWith',
  'containsBlanks',
  'notContainsBlanks',
  'containsErrors',
  'notContainsErrors',
  'timePeriod',
  'top10',
  'aboveAverage',
  'duplicateValues',
  'uniqueValues',
  'colorScale',
  'dataBar',
  'iconSet',
  'expression',
])
const ICON_SETS: ReadonlyMap<string, number> = new Map([
  ['3Arrows', 3], ['3ArrowsGray', 3], ['3Flags', 3], ['3TrafficLights1', 3], ['3TrafficLights2', 3],
  ['3Signs', 3], ['3Symbols', 3], ['3Symbols2', 3], ['4Arrows', 4], ['4ArrowsGray', 4],
  ['4RedToBlack', 4], ['4Rating', 4], ['4TrafficLights', 4], ['5Arrows', 5], ['5ArrowsGray', 5],
  ['5Rating', 5], ['5Quarters', 5], ['3Triangles', 3], ['3Stars', 3], ['5Boxes', 5],
])
const TEXT_RULE_TYPES = ['containsText', 'notContainsText', 'beginsWith', 'endsWith'] as const
const BLANK_ERROR_TYPES = ['containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'] as const

const isIntText = (v: string | undefined): boolean => v !== undefined && /^\s*-?\d+\s*$/.test(v)
const parseBoolAttr = (v: string | undefined, def: boolean): boolean | undefined => {
  if (v === undefined) return def
  const t = v.trim().toLowerCase()
  if (t === '1' || t === 'true') return true
  if (t === '0' || t === 'false') return false
  return undefined
}

/** Namespace-gated children: local-name match plus strict URI equality. */
function kids(node: XmlNode | undefined, local: string, expectedNs: string | undefined): XmlNode[] {
  if (!node) return []
  return getChildren(node, local).filter(child => elementNamespaceUri(child) === expectedNs)
}

function normPairId(raw: string | undefined): string | undefined {
  if (raw === undefined) return undefined
  const t = raw.trim()
  return t.length > 0 ? t.toUpperCase() : undefined
}

function colLettersToIndex(letters: string): number | undefined {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) return undefined
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n >= 1 && n <= MAX_COL ? n - 1 : undefined
}

function parseSqrefToken(token: string): ParsedRange | undefined {
  const t = token.replace(/\$/g, '')
  let m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(t)
  if (m) {
    const c1 = colLettersToIndex(m[1]!)
    const c2 = m[3] !== undefined ? colLettersToIndex(m[3]!) : c1
    const r1 = Number(m[2])
    const r2 = m[4] !== undefined ? Number(m[4]) : r1
    if (c1 === undefined || c2 === undefined) return undefined
    if (!Number.isSafeInteger(r1) || !Number.isSafeInteger(r2) || r1 < 1 || r1 > MAX_ROW || r2 < 1 || r2 > MAX_ROW) {
      return undefined
    }
    return {
      firstRow: Math.min(r1, r2) - 1,
      lastRow: Math.max(r1, r2) - 1,
      firstCol: Math.min(c1, c2),
      lastCol: Math.max(c1, c2),
    }
  }
  m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(t)
  if (m) {
    const c1 = colLettersToIndex(m[1]!)
    const c2 = colLettersToIndex(m[2]!)
    if (c1 === undefined || c2 === undefined) return undefined
    return { firstRow: 0, lastRow: MAX_ROW - 1, firstCol: Math.min(c1, c2), lastCol: Math.max(c1, c2) }
  }
  m = /^(\d+):(\d+)$/.exec(t)
  if (m) {
    const r1 = Number(m[1])
    const r2 = Number(m[2])
    if (!Number.isSafeInteger(r1) || !Number.isSafeInteger(r2) || r1 < 1 || r1 > MAX_ROW || r2 < 1 || r2 > MAX_ROW) {
      return undefined
    }
    return {
      firstRow: Math.min(r1, r2) - 1,
      lastRow: Math.max(r1, r2) - 1,
      firstCol: 0,
      lastCol: MAX_COL - 1,
    }
  }
  return undefined
}

function parseSqref(sqref: string | undefined): ParsedRange[] | undefined {
  if (sqref === undefined) return undefined
  const tokens = sqref.trim().split(/\s+/).filter(t => t.length > 0)
  if (tokens.length === 0) return undefined
  const out: ParsedRange[] = []
  for (const token of tokens) {
    const range = parseSqrefToken(token)
    if (!range) return undefined
    out.push(range)
  }
  return out
}

function parseLiteral(text: string): CfLiteral | undefined {
  const t = text.trim().replace(/^=/, '').trim()
  if (/^[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$/.test(t)) {
    const num = Number(t)
    return Number.isFinite(num) ? { kind: 'number', value: num } : undefined
  }
  if (t.length >= 2 && t.startsWith('"') && t.endsWith('"')) {
    return { kind: 'string', value: t.slice(1, -1).replace(/""/g, '"') }
  }
  const lower = t.toLowerCase()
  if (lower === 'true') return { kind: 'boolean', value: true }
  if (lower === 'false') return { kind: 'boolean', value: false }
  return undefined
}

const REF_PATTERN = String.raw`\$?[A-Za-z]{1,3}\$?\d+`

/** Remove whitespace outside string literals (whitespace inside needles is significant). */
function stripOutsideStrings(text: string): string {
  let out = ''
  let i = 0
  while (i < text.length) {
    const ch = text[i]!
    if (ch === '"') {
      let j = i + 1
      out += ch
      while (j < text.length) {
        out += text[j]
        if (text[j] === '"') {
          if (text[j + 1] === '"') {
            out += text[j + 1]
            j += 2
            continue
          }
          j++
          break
        }
        j++
      }
      i = j
      continue
    }
    if (ch !== ' ' && ch !== '\t' && ch !== '\n' && ch !== '\r' && ch !== ' ') out += ch
    i++
  }
  return out
}

function canonicalTextFormula(type: string, needle: string, formulas: readonly string[]): boolean {
  if (formulas.length !== 1) return false
  const f = stripOutsideStrings(formulas[0]!)
  const q = `"${needle.replace(/"/g, '""')}"`
  const esc = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const ref = REF_PATTERN
  switch (type) {
    case 'containsText':
      return new RegExp(`^NOT\\(ISERROR\\(SEARCH\\(${esc},${ref}\\)\\)\\)$`, 'i').test(f)
    case 'notContainsText':
      return new RegExp(`^ISERROR\\(SEARCH\\(${esc},${ref}\\)\\)$`, 'i').test(f)
    case 'beginsWith':
      return new RegExp(`^LEFT\\((${ref}),${needle.length}\\)=${esc}$`, 'i').test(f)
    case 'endsWith':
      return new RegExp(`^RIGHT\\((${ref}),${needle.length}\\)=${esc}$`, 'i').test(f)
    default:
      return false
  }
}

function canonicalBlankFormula(type: string, formulas: readonly string[]): boolean {
  if (formulas.length !== 1) return false
  const f = stripOutsideStrings(formulas[0]!)
  const ref = REF_PATTERN
  switch (type) {
    case 'containsBlanks':
      return new RegExp(`^LEN\\(TRIM\\((${ref})\\)\\)=0$`, 'i').test(f)
    case 'notContainsBlanks':
      return new RegExp(`^LEN\\(TRIM\\((${ref})\\)\\)>0$`, 'i').test(f)
    case 'containsErrors':
      return new RegExp(`^ISERROR\\((${ref})\\)$`, 'i').test(f)
    case 'notContainsErrors':
      return new RegExp(`^NOT\\(ISERROR\\((${ref})\\)\\)$`, 'i').test(f)
    default:
      return false
  }
}

function canonicalTimePattern(period: string): RegExp | undefined {
  const r = `(${REF_PATTERN})`
  const patterns: Record<string, string> = {
    today: `FLOOR\\(${r},1\\)=TODAY\\(\\)`,
    yesterday: `FLOOR\\(${r},1\\)=TODAY\\(\\)-1`,
    tomorrow: `FLOOR\\(${r},1\\)=TODAY\\(\\)\\+1`,
    last7Days: `AND\\(TODAY\\(\\)-FLOOR\\(${r},1\\)<=6,FLOOR\\(${r},1\\)<=TODAY\\(\\)\\)`,
    thisWeek: `AND\\(TODAY\\(\\)-ROUNDDOWN\\(${r},0\\)<=\\(WEEKDAY\\(TODAY\\(\\)\\)-1\\),ROUNDDOWN\\(${r},0\\)-TODAY\\(\\)<=\\(7-WEEKDAY\\(TODAY\\(\\)\\)\\)\\)`,
    lastWeek: `AND\\(TODAY\\(\\)-ROUNDDOWN\\(${r},0\\)>\\(WEEKDAY\\(TODAY\\(\\)\\)-1\\),TODAY\\(\\)-ROUNDDOWN\\(${r},0\\)<\\(WEEKDAY\\(TODAY\\(\\)\\)\\+6\\)\\)`,
    nextWeek: `AND\\(ROUNDDOWN\\(${r},0\\)-TODAY\\(\\)>\\(7-WEEKDAY\\(TODAY\\(\\)\\)\\),ROUNDDOWN\\(${r},0\\)-TODAY\\(\\)<\\(15-WEEKDAY\\(TODAY\\(\\)\\)\\)\\)`,
    thisMonth: `AND\\(MONTH\\(${r}\\)=MONTH\\(TODAY\\(\\)\\),YEAR\\(${r}\\)=YEAR\\(TODAY\\(\\)\\)\\)`,
    lastMonth: `AND\\(MONTH\\(${r}\\)=MONTH\\(EDATE\\(TODAY\\(\\),0-1\\)\\),YEAR\\(${r}\\)=YEAR\\(EDATE\\(TODAY\\(\\),0-1\\)\\)\\)`,
    nextMonth: `AND\\(MONTH\\(${r}\\)=MONTH\\(EDATE\\(TODAY\\(\\),0\\+1\\)\\),YEAR\\(${r}\\)=YEAR\\(EDATE\\(TODAY\\(\\),0\\+1\\)\\)\\)`,
  }
  const body = patterns[period]
  return body === undefined ? undefined : new RegExp(`^${body}$`, 'i')
}

/** Crude lexical token count (no evaluation): identifiers, numbers, strings, operators. */
function countFormulaTokens(text: string): number {
  let tokens = 0
  let i = 0
  const push = () => {
    tokens++
    return tokens > CF_LIMITS.formulaTokens
  }
  while (i < text.length) {
    const ch = text[i]!
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++
      continue
    }
    if (ch === '"') {
      let j = i + 1
      while (j < text.length) {
        if (text[j] === '"') {
          if (text[j + 1] === '"') j += 2
          else break
        } else j++
      }
      i = Math.min(text.length, j + 1)
      if (push()) return tokens
      continue
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i + 1
      while (j < text.length && /[A-Za-z0-9_.]/.test(text[j]!)) j++
      i = j
      if (push()) return tokens
      continue
    }
    if (/[0-9.]/.test(ch)) {
      let j = i + 1
      while (j < text.length && /[0-9.]/.test(text[j]!)) j++
      i = j
      if (push()) return tokens
      continue
    }
    i++
    if (push()) return tokens
  }
  return tokens
}

function parseTopRank(
  a: Record<string, string>,
): { rank: number; bottom: boolean; percent: boolean } | undefined | 'invalid' {
  const bottom = parseBoolAttr(a['bottom'], false)
  const percent = parseBoolAttr(a['percent'], false)
  if (bottom === undefined || percent === undefined) return 'invalid'
  if (a['rank'] === undefined) return undefined
  if (!isIntText(a['rank'])) return 'invalid'
  const rank = Number(a['rank'])
  if (!Number.isSafeInteger(rank)) return 'invalid'
  if (rank <= 0) return undefined
  if (percent ? rank > 100 : rank > 1000) return undefined
  return { rank, bottom, percent }
}

function parseAboveAverage(
  a: Record<string, string>,
): { above: boolean; equal: boolean; stdDev?: number } | undefined | 'invalid' {
  const above = parseBoolAttr(a['aboveAverage'], true)
  const equal = parseBoolAttr(a['equalAverage'], false)
  if (above === undefined || equal === undefined) return 'invalid'
  if (a['stdDev'] === undefined) return { above, equal }
  if (!isIntText(a['stdDev'])) return 'invalid'
  const stdDev = Number(a['stdDev'])
  if (!Number.isSafeInteger(stdDev)) return 'invalid'
  if (equal) return 'invalid'
  if (stdDev < 1 || stdDev > 3) return undefined
  return { above, equal, stdDev }
}

/** Total retained string weight under a node (attrs + text), or undefined past budget. */
function boundedTextLength(node: XmlNode | undefined, budget: number): number | undefined {
  if (!node) return 0
  let total = 0
  const stack: XmlNode[] = [node]
  while (stack.length > 0) {
    const current = stack.pop()!
    for (const [k, v] of Object.entries(current)) {
      if (k === '@attrs') {
        for (const av of Object.values(v as Record<string, unknown>)) {
          total += String(av).length
          if (total > budget) return undefined
        }
      } else if (k === '#text') {
        total += String(v).length
        if (total > budget) return undefined
      } else if (v && typeof v === 'object') {
        if (Array.isArray(v)) {
          if (v.length > 512) return undefined
          for (const item of v) if (item && typeof item === 'object') stack.push(item as XmlNode)
        } else {
          stack.push(v as XmlNode)
        }
      }
    }
  }
  return total
}

class SheetFinalizer {
  charged = 0
  retained = 0
  rules: CfRule[] = []
  sentinel: { reason: string } | undefined
  diagnostics: { kind: 'unsupported' | 'parse'; feature: `cf-${string}`; message: string }[] = []
  aborted = false

  constructor(
    readonly hold: CfReservation,
    readonly styles: CfStylesSource,
    readonly memory: CfMemory,
    readonly caps: ReturnType<typeof createCfCaps>,
    readonly diagHandle: ReturnType<typeof createCfDiagnostics>,
    readonly meta: CfParseMeta,
  ) {}

  private where(): string {
    return this.meta.partPath ?? this.meta.sheetId ?? 'worksheet'
  }

  diag(kind: 'unsupported' | 'parse', feature: `cf-${string}`, message: string): void {
    const bound = `${this.where()}: ${message}`.slice(0, 512)
    this.diagnostics.push({ kind, feature, message: bound })
    this.diagHandle.add({ kind, feature, message: bound })
  }

  failSheet(reason: string, feature: `cf-${string}`, message: string): void {
    this.diag('unsupported', feature, message)
    if (!this.sentinel) this.sentinel = { reason }
    this.aborted = true
  }

  /** Admit `delta` more bytes before building; fail-stop on exhaustion. */
  admit(delta: number): void {
    if (!Number.isSafeInteger(delta) || delta < 0) throw new Error('CF admission exhausted')
    const next = this.charged + delta
    if (!this.hold.grow(next)) throw new Error('CF admission exhausted')
    this.charged = next
  }

  keep(delta: number): void {
    this.retained += delta
  }

  reserveMeta(bytes: number): boolean {
    return this.caps.reserveMetadata(bytes).kind === 'ready'
  }
}

interface BodyInputs {
  attrs: Record<string, string>
  typeText: string | undefined
  operatorText: string | undefined
  textAttr: string | undefined
  periodText: string | undefined
  dxfIdText: string | undefined
  formulas: string[]
  hasForeignFormula: boolean
  inlineDxf: XmlNode | undefined
  cfvoCount: number
  cfvoMissingType: boolean
  customIcons: number
  iconProblem: 'cf-unsupported-icon' | undefined
  origin: CfRuleOrigin
}

function buildRuleBody(
  fin: SheetFinalizer,
  inputs: BodyInputs,
  ranges: ParsedRange[],
  documentOrder: number,
  fallbackType: CfRuleType | 'unknown',
): number {
  let actual = 256 + inputs.formulas.reduce((n, f) => n + f.length * 2, 0)
  const pushUnresolved = (reason: string, priority: number, stop: boolean, writes: CfWriteMask): number => {
    fin.diag('unsupported', reason as `cf-${string}`, `rule priority=${priority} type=${inputs.typeText?.trim() ?? ''}: ${reason}`)
    fin.rules.push({
      id: -1,
      priority,
      documentOrder,
      ranges,
      stopIfTrue: stop,
      writes,
      definition: { kind: 'unresolved', type: fallbackType, reason, origin: inputs.origin },
      dxf: undefined,
    })
    return actual
  }
  const invalid = (priority: number, stop: boolean, message: string): number => {
    fin.diag('unsupported', 'cf-invalid-rule', message)
    fin.rules.push({
      id: -1, priority, documentOrder, ranges, stopIfTrue: stop,
      writes: new Set(['ALL']),
      definition: { kind: 'unresolved', type: fallbackType, reason: 'cf-invalid-rule', origin: inputs.origin },
      dxf: undefined,
    })
    return actual
  }

  const priorityText = inputs.attrs['priority']
  const priority = priorityText !== undefined && isIntText(priorityText) && Number(priorityText) >= 1
    ? Number(priorityText)
    : undefined
  if (priority === undefined) {
    return invalid(0, false, `rule priority "${String(priorityText)}" is not a positive integer`)
  }
  // Unrecognized stopIfTrue conservatively halts downstream rules: an unknown
  // halting intent must withhold, never let later rules paint over it.
  const stopRaw = inputs.attrs['stopIfTrue']?.trim().toLowerCase()
  const stop = stopRaw === undefined || stopRaw === '0' || stopRaw === 'false'
    ? false
    : stopRaw === '1' || stopRaw === 'true'
      ? true
      : undefined
  if (stop === undefined) {
    return invalid(priority, true, `stopIfTrue="${String(inputs.attrs['stopIfTrue'])}" is not a boolean`)
  }
  const active = parseBoolAttr(inputs.attrs['activePresent'], false)
  if (active === undefined) {
    return invalid(priority, true, `activePresent="${String(inputs.attrs['activePresent'])}" is not a boolean`)
  }

  // Bounded dxf resolution before definition dispatch (writes feed every outcome).
  let dxfProps: CfCellState['properties'] | undefined
  let dxfWrites: CfWriteMask | undefined
  if (inputs.inlineDxf !== undefined || inputs.dxfIdText !== undefined) {
    let strict:
      | { properties: CfCellState['properties']; writes: CfWriteMask; unsupported: readonly string[] }
      | undefined
    if (inputs.inlineDxf !== undefined) {
      strict = parseCfDxf(inputs.inlineDxf, fin.styles.colors)
    } else {
      const dxfId = isIntText(inputs.dxfIdText) ? Number(inputs.dxfIdText) : -1
      if (dxfId < 0 || dxfId >= fin.styles.dxfNodes.length) {
        return pushUnresolved('cf-unsupported-dxf', priority, stop, new Set(['ALL']))
      }
      strict = parseCfDxf(fin.styles.dxfNodes[dxfId]!, fin.styles.colors)
    }
    if (strict.unsupported.length > 0) {
      return pushUnresolved('cf-unsupported-dxf', priority, stop, new Set(['ALL']))
    }
    dxfProps = strict.properties
    dxfWrites = strict.writes
    actual += dxfWrites.size * 64
    for (const value of Object.values(dxfProps)) actual += String(value).length * 2
  }

  const type = (inputs.typeText ?? '').trim()
  const withDxf = (extra?: CfWriteMask): CfWriteMask => {
    const out = new Set<string>(dxfWrites ?? [])
    if (extra) for (const w of extra) out.add(w)
    return out
  }
  if (!KNOWN_TYPES.has(type)) {
    return pushUnresolved('cf-unsupported-rule', priority, stop, new Set(['ALL']))
  }
  if (active && type !== 'colorScale' && type !== 'dataBar' && type !== 'iconSet') {
    return invalid(priority, stop, `activePresent="1" on non-visual rule type "${type}" is non-conforming`)
  }

  const finishResolved = (definition: Record<string, unknown>): number => {
    fin.rules.push({
      id: -1, priority, documentOrder, ranges, stopIfTrue: stop,
      writes: withDxf(), definition: { ...inputs.origin, ...definition } as CfDefinition, dxf: dxfProps,
    })
    return actual
  }
  const finishUnresolved = (reason: string, writes: CfWriteMask): number => {
    pushUnresolved(reason, priority, stop, withDxf(writes))
    const last = fin.rules[fin.rules.length - 1]!
    ;(last as { dxf: typeof dxfProps }).dxf = dxfProps
    return actual
  }

  switch (type) {
    case 'cellIs': {
      const operator = (inputs.operatorText ?? '').trim()
      if (!RELATIONAL_OPERATORS.has(operator)) {
        if (CELLIS_TEXT_OPERATORS.has(operator) || operator === 'notContainsText') {
          return finishUnresolved('cf-unsupported-rule', withDxf())
        }
        return invalid(priority, true, `cellIs operator "${operator}" is not a known operator`)
      }
      const want = operator === 'between' || operator === 'notBetween' ? 2 : 1
      if (inputs.formulas.length !== want) {
        return invalid(priority, true, `cellIs operator "${operator}" needs ${want} formula(s), found ${inputs.formulas.length}`)
      }
      const operands: CfLiteral[] = []
      for (const text of inputs.formulas) {
        const lit = parseLiteral(text)
        if (!lit) return finishUnresolved('cf-unsupported-rule', withDxf())
        operands.push(lit)
      }
      return finishResolved({ kind: 'cellIs', operator: operator as CfOperator, operands })
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith': {
      const needle = inputs.textAttr
      if (needle === undefined) {
        return invalid(priority, true, `text rule "${type}" is missing its search needle`)
      }
      if (inputs.hasForeignFormula) return finishUnresolved('cf-unsupported-rule', withDxf())
      if (inputs.formulas.length === 0 || canonicalTextFormula(type, needle, inputs.formulas)) {
        return finishResolved({ kind: 'text', type: type as (typeof TEXT_RULE_TYPES)[number], needle })
      }
      return finishUnresolved('cf-unsupported-rule', withDxf())
    }
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors': {
      if (inputs.hasForeignFormula) return finishUnresolved('cf-unsupported-rule', withDxf())
      if (inputs.formulas.length === 0 || canonicalBlankFormula(type, inputs.formulas)) {
        return finishResolved({ kind: 'predicate', type: type as (typeof BLANK_ERROR_TYPES)[number] })
      }
      return finishUnresolved('cf-unsupported-rule', withDxf())
    }
    case 'timePeriod': {
      const period = (inputs.periodText ?? '').trim()
      if (!PERIODS.has(period)) {
        return invalid(priority, true, `timePeriod="${inputs.periodText ?? ''}" is not a known period`)
      }
      if (inputs.hasForeignFormula) return finishUnresolved('cf-unsupported-rule', withDxf())
      if (inputs.formulas.length === 0) {
        return finishResolved({ kind: 'timePeriod', period: period as CfTimePeriod })
      }
    const pattern = canonicalTimePattern(period)
    const f = inputs.formulas.length === 1 ? stripOutsideStrings(inputs.formulas[0]!) : ''
      if (pattern && pattern.test(f)) {
        return finishResolved({ kind: 'timePeriod', period: period as CfTimePeriod })
      }
      return finishUnresolved('cf-unsupported-rule', withDxf())
    }
    case 'top10': {
      const rank = parseTopRank(inputs.attrs)
      if (rank === 'invalid') {
        return invalid(priority, true, 'top10 rank/percent/bottom attributes are not usable')
      }
      if (rank === undefined) return finishUnresolved('cf-unsupported-rule', withDxf())
      return finishResolved({ kind: 'top10', ...rank })
    }
    case 'aboveAverage': {
      const parsed = parseAboveAverage(inputs.attrs)
      if (parsed === 'invalid') {
        return invalid(priority, true, 'aboveAverage attributes are not usable')
      }
      if (parsed === undefined) return finishUnresolved('cf-unsupported-rule', withDxf())
      return finishResolved({ kind: 'aboveAverage', ...parsed })
    }
    case 'duplicateValues':
    case 'uniqueValues': {
      if (inputs.formulas.length > 0 || inputs.hasForeignFormula) {
        return finishUnresolved('cf-unsupported-rule', withDxf())
      }
      return finishResolved({ kind: 'predicate', type: type as 'duplicateValues' | 'uniqueValues' })
    }
    case 'colorScale':
    case 'dataBar':
    case 'iconSet': {
      const slot = new Set([type])
      if (inputs.cfvoMissingType) return finishUnresolved('cf-invalid-cfvo', withDxf(slot))
      if (inputs.iconProblem) return finishUnresolved(inputs.iconProblem, withDxf(slot))
      return finishUnresolved('cf-unsupported-rule', withDxf(slot))
    }
    case 'expression': {
      return finishUnresolved('cf-unsupported-rule', withDxf())
    }
    default: {
      return pushUnresolved('cf-unsupported-rule', priority, stop, new Set(['ALL']))
    }
  }
}

/* ------------------------------------------------------------------ */
/* Bounded rule-input extraction.                                      */
/* ------------------------------------------------------------------ */

type ReadInputs = { ok: true; inputs: BodyInputs; bound: number } | { ok: false; reason: 'cf-resource-limit' }

const KNOWN_ATTRS: ReadonlySet<string> = new Set([
  'type', 'operator', 'text', 'timePeriod', 'dxfId', 'priority', 'stopIfTrue', 'activePresent',
  'rank', 'percent', 'bottom', 'aboveAverage', 'equalAverage', 'stdDev', 'id',
])

function pickAttrs(node: XmlNode): Record<string, string> {
  const out: Record<string, string> = {}
  const a = attrs(node)
  for (const key of KNOWN_ATTRS) {
    const v = a[key]
    if (typeof v === 'string') out[key] = v
  }
  return out
}

function readFormulas(
  fin: SheetFinalizer,
  ruleNode: XmlNode,
  ns: string | undefined,
  tag: 'formula' | 'f',
): { formulas: string[]; foreign: boolean } | undefined {
  const all = getChildren(ruleNode, tag)
  const mine = all.filter(n => elementNamespaceUri(n) === ns)
  const foreign = all.length !== mine.length
  if (mine.length > CF_LIMITS.formulaChildren) {
    fin.diag('unsupported', 'cf-resource-limit', `rule carries ${mine.length} ${tag} children (cap ${CF_LIMITS.formulaChildren})`)
    return undefined
  }
  const formulas: string[] = []
  for (const child of mine) {
    const raw = textOf(child)
    if (raw.length > CF_LIMITS.textUnits) {
      fin.diag('unsupported', 'cf-resource-limit', `${tag} text exceeds ${CF_LIMITS.textUnits} units`)
      return undefined
    }
    const text = raw.trim()
    if (countFormulaTokens(text) > CF_LIMITS.formulaTokens) {
      fin.diag('unsupported', 'cf-resource-limit', `${tag} exceeds ${CF_LIMITS.formulaTokens} tokens`)
      return undefined
    }
    formulas.push(text)
  }
  return { formulas, foreign }
}

interface VisualFacts {
  cfvoCount: number
  cfvoMissingType: boolean
  customIcons: number
  iconProblem: 'cf-unsupported-icon' | undefined
}

function scanRuleVisual(
  fin: SheetFinalizer,
  ruleNode: XmlNode,
  ns: string | undefined,
  type: string,
): VisualFacts | undefined {
  const hosts = kids(ruleNode, type, ns)
  const host = hosts[0]
  if (!host) return { cfvoCount: 0, cfvoMissingType: false, customIcons: 0, iconProblem: undefined }
  let cfvoCount = 0
  let cfvoMissingType = false
  for (const cfvo of kids(host, 'cfvo', ns)) {
    cfvoCount++
    if (cfvoCount > CF_LIMITS.cfvo) {
      fin.diag('unsupported', 'cf-resource-limit', `${type} carries more than ${CF_LIMITS.cfvo} cfvo children`)
      return undefined
    }
    const a = attrs(cfvo)
    if (a['type'] === undefined) cfvoMissingType = true
    if (a['val'] !== undefined && a['val'].length > CF_LIMITS.textUnits) {
      fin.diag('unsupported', 'cf-resource-limit', 'cfvo val exceeds the text cap')
      return undefined
    }
  }
  let customIcons = 0
  let iconProblem: 'cf-unsupported-icon' | undefined
  if (type === 'iconSet' && host) {
    if (attrs(host)['custom'] === '1' || attrs(host)['custom'] === 'true') {
      for (const icon of kids(host, 'cfIcon', ns)) {
        customIcons++
        if (customIcons > CF_LIMITS.customIcons) {
          fin.diag('unsupported', 'cf-resource-limit', `iconSet carries more than ${CF_LIMITS.customIcons} custom icons`)
          return undefined
        }
        const ia = attrs(icon)
        const size = ICON_SETS.get(ia['iconSet'] ?? '')
        if (size === undefined || !isIntText(ia['iconId']) || Number(ia['iconId']) < 0 || Number(ia['iconId']) >= size) {
          iconProblem = 'cf-unsupported-icon'
        }
      }
    }
    const setName = attrs(host)['iconSet']
    if (setName !== undefined && ICON_SETS.get(setName) === undefined) iconProblem = 'cf-unsupported-icon'
  }
  return { cfvoCount, cfvoMissingType, customIcons, iconProblem }
}

function readRuleInputs(
  fin: SheetFinalizer,
  ruleNode: XmlNode,
  ns: string | undefined,
  dialect: 'base' | 'x14',
  pairRaw: string | undefined,
): ReadInputs {
  const a = attrs(ruleNode)
  const type = (a['type'] ?? '').trim()
  const tag = dialect === 'base' ? 'formula' : 'f'
  const read = readFormulas(fin, ruleNode, dialect === 'base' ? ns : XM_MAIN_NS, tag)
  if (!read) return { ok: false, reason: 'cf-resource-limit' }
  let inlineDxf: XmlNode | undefined
  if (dialect === 'x14') {
    const dxfs = kids(ruleNode, 'dxf', X14_MAIN_NS)
    inlineDxf = dxfs[0]
    if (dxfs.length > 1) read.foreign = true
  }
  let cfvoCount = 0
  let cfvoMissingType = false
  let customIcons = 0
  let iconProblem: 'cf-unsupported-icon' | undefined
  if (VISUAL_TYPES.has(type)) {
    const scan = scanRuleVisual(fin, ruleNode, dialect === 'base' ? ns : X14_MAIN_NS, type)
    if (!scan) return { ok: false, reason: 'cf-resource-limit' }
    cfvoCount = scan.cfvoCount
    cfvoMissingType = scan.cfvoMissingType
    customIcons = scan.customIcons
    iconProblem = scan.iconProblem
  }
  const textAttr = a['text']
  if (textAttr !== undefined && textAttr.length > CF_LIMITS.textUnits) {
    fin.diag('unsupported', 'cf-resource-limit', 'text needle exceeds the text cap')
    return { ok: false, reason: 'cf-resource-limit' }
  }
  const guidRaw = dialect === 'base' ? pairRaw : a['id']
  const guid = guidRaw !== undefined && guidRaw.trim().length <= 256 ? guidRaw.trim() : undefined
  const origin: CfRuleOrigin = {
    dialect,
    ...(guid !== undefined ? { guid } : {}),
    ...(a['activePresent'] === '1' ? { activePresent: true as const } : {}),
    formulas: read.formulas,
    ...(a['dxfId'] !== undefined && isIntText(a['dxfId']) ? { dxfId: Number(a['dxfId']) } : {}),
  }
  let dxfBudget = 0
  if (inlineDxf !== undefined) {
    const weighed = boundedTextLength(inlineDxf, CF_LIMITS.textUnits)
    if (weighed === undefined) return { ok: false, reason: 'cf-resource-limit' }
    dxfBudget = weighed
  } else if (a['dxfId'] !== undefined && isIntText(a['dxfId'])) {
    const id = Number(a['dxfId'])
    const node = id >= 0 && id < fin.styles.dxfNodes.length ? fin.styles.dxfNodes[id] : undefined
    if (node) {
      const weighed = boundedTextLength(node, CF_LIMITS.textUnits)
      if (weighed === undefined) return { ok: false, reason: 'cf-resource-limit' }
      dxfBudget = weighed
    }
  }
  const bound =
    256 +
    read.formulas.reduce((n, f) => n + f.length * 2, 0) +
    (textAttr ? textAttr.length * 2 : 0) +
    dxfBudget * 2 +
    256
  return {
    ok: true,
    inputs: {
      attrs: pickAttrs(ruleNode),
      typeText: a['type'],
      operatorText: a['operator'],
      textAttr,
      periodText: a['timePeriod'],
      dxfIdText: a['dxfId'],
        formulas: read.formulas,
        hasForeignFormula: read.foreign,
        inlineDxf,
        cfvoCount,
        cfvoMissingType,
        customIcons,
        iconProblem,
        origin,
      },
      bound,
    }
}

/* ------------------------------------------------------------------ */
/* Element collection, x14 pairing and sheet finalization.             */
/* ------------------------------------------------------------------ */

interface PendingRule {
  kind: 'base' | 'x14'
  node: XmlNode
  element: XmlNode
  ranges: ParsedRange[] | undefined
  sqrefText: string | undefined
  docOrder: number
  pairId: string | undefined
  pairRaw: string | undefined
}

function pushScopeRefusal(
  fin: SheetFinalizer,
  ruleNode: XmlNode,
  kind: 'base' | 'x14',
  documentOrder: number,
  fallbackType: CfRuleType | 'unknown',
): void {
  const a = attrs(ruleNode)
  const priority = a['priority'] !== undefined && isIntText(a['priority']) && Number(a['priority']) >= 1
    ? Number(a['priority'])
    : 0
  fin.diag('unsupported', 'cf-resource-limit', `rule priority=${priority} exceeds the per-rule range cap`)
  fin.admit(256)
  fin.keep(200)
  fin.rules.push({
    id: -1, priority, documentOrder, ranges: [], stopIfTrue: false,
    writes: new Set(['ALL']),
    definition: { kind: 'unresolved', type: fallbackType, reason: 'cf-resource-limit', origin: { dialect: kind, formulas: [] } },
    dxf: undefined,
  })
}

function invalidLinkRule(
  fin: SheetFinalizer,
  priority: number,
  documentOrder: number,
  ranges: ParsedRange[],
  origin: CfRuleOrigin,
  typeText: string,
  message: string,
): void {
  fin.diag('unsupported', 'cf-invalid-x14-link', message)
  fin.admit(256)
  fin.keep(200)
  const type = KNOWN_TYPES.has(typeText.trim()) ? (typeText.trim() as CfRuleType) : 'dataBar'
  fin.rules.push({
    id: -1,
    priority,
    documentOrder,
    ranges,
    stopIfTrue: false,
    writes: new Set(['ALL']),
    definition: { kind: 'unresolved', type, reason: 'cf-invalid-x14-link', origin },
    dxf: undefined,
  })
}

/** Bounded x14 companion visual facts for a structurally clean pair. */
function readPairedExtFacts(
  fin: SheetFinalizer,
  extNode: XmlNode,
): { ok: true; cfvoCount: number; cfvoMissingType: boolean } | { ok: false } {
  const visual = kids(extNode, 'dataBar', X14_MAIN_NS)[0]
  if (!visual) return { ok: true, cfvoCount: 0, cfvoMissingType: false }
  let cfvoCount = 0
  let cfvoMissingType = false
  for (const cfvo of kids(visual, 'cfvo', X14_MAIN_NS)) {
    cfvoCount++
    if (cfvoCount > CF_LIMITS.cfvo) {
      fin.diag('unsupported', 'cf-resource-limit', 'paired x14 dataBar exceeds the cfvo cap')
      return { ok: false }
    }
    const a = attrs(cfvo)
    if (a['type'] === undefined) cfvoMissingType = true
    if (a['val'] !== undefined && a['val'].length > CF_LIMITS.textUnits) {
      fin.diag('unsupported', 'cf-resource-limit', 'paired x14 cfvo val exceeds the text cap')
      return { ok: false }
    }
  }
  return { ok: true, cfvoCount, cfvoMissingType }
}

function collectAndBuild(fin: SheetFinalizer, sheetNs: string | undefined, base: XmlNode[], x14: XmlNode[]): void {
  const pending: PendingRule[] = []
  let order = 0

  const admitElement = (sqrefText: string | undefined, ruleCount: number): ParsedRange[] | 'rule-refusal' | undefined => {
    const ranges = parseSqref(sqrefText)
    if (!ranges) {
      fin.failSheet('cf-invalid-rule', 'cf-sqref', `conditionalFormatting sqref "${String(sqrefText)}" has no usable scope`)
      return undefined
    }
    if (sqrefText !== undefined && sqrefText.length > CF_LIMITS.textUnits * 16) {
      fin.failSheet('cf-invalid-rule', 'cf-sqref', 'conditionalFormatting sqref text is unrepresentable')
      return undefined
    }
    if (ranges.length > CF_LIMITS.rangesPerRule) {
      // Per-rule interval overflow withholds only the affected rules (scope is
      // known); sheet-wide preconditions use the sentinel below.
      return 'rule-refusal'
    }
    const elementBytes = 64 + ranges.length * 32 + (sqrefText ? sqrefText.length * 2 : 0)
    fin.admit(elementBytes + 256)
    fin.keep(elementBytes)
    if (!fin.reserveMeta(elementBytes)) {
      fin.failSheet('cf-resource-limit', 'cf-resource-limit', 'sheet CF metadata exceeds the per-sheet cap')
      return undefined
    }
    if (fin.caps.addRule(ranges.length).kind === 'unresolved') {
      fin.failSheet('cf-resource-limit', 'cf-resource-limit', 'sheet CF exceeds the rule/range caps')
      return undefined
    }
    for (let i = 1; i < ruleCount; i++) {
      if (fin.caps.addRule(0).kind === 'unresolved') {
        fin.failSheet('cf-resource-limit', 'cf-resource-limit', 'sheet CF exceeds the rule cap')
        return undefined
      }
    }
    return ranges
  }

  for (const element of base) {
    if (fin.aborted) break
    if (elementNamespaceUri(element) !== sheetNs) {
      fin.diag('unsupported', 'cf-namespace', 'conditionalFormatting outside the worksheet namespace is skipped')
      continue
    }
    const ruleNodes = kids(element, 'cfRule', sheetNs)
    if (getChildren(element, 'cfRule').length !== ruleNodes.length) {
      fin.diag('unsupported', 'cf-namespace', 'conditionalFormatting holds a wrong-namespace cfRule, skipped')
    }
    if (ruleNodes.length === 0) continue
    const admitted = admitElement(attrs(element).sqref, ruleNodes.length)
    if (admitted === undefined) continue
    if (admitted === 'rule-refusal') {
      for (const ruleNode of ruleNodes) {
        pushScopeRefusal(fin, ruleNode, 'base', order++, KNOWN_TYPES.has((attrs(ruleNode)['type'] ?? '').trim()) ? ((attrs(ruleNode)['type'] ?? '').trim() as CfRuleType) : 'unknown')
      }
      continue
    }
    const ranges = admitted
    for (const ruleNode of ruleNodes) {
      pending.push({ kind: 'base', node: ruleNode, element, ranges, sqrefText: attrs(element).sqref, docOrder: order++, pairId: undefined, pairRaw: undefined })
    }
  }

  const x14Bodies: Array<{ element: XmlNode; node: XmlNode; sqrefText: string | undefined; docOrder: number }> = []
  for (const element of x14) {
    if (fin.aborted) break
    if (elementNamespaceUri(element) !== X14_MAIN_NS) {
      fin.diag('unsupported', 'cf-namespace', 'non-x14 conditionalFormattings content is skipped')
      continue
    }
    const sqrefNodes = kids(element, 'sqref', XM_MAIN_NS)
    if (getChildren(element, 'sqref').length !== sqrefNodes.length) {
      fin.diag('unsupported', 'cf-namespace', 'x14 conditionalFormatting holds a wrong-namespace sqref')
    }
    const sqrefText = sqrefNodes.length > 0 ? textOf(sqrefNodes[0]).trim() : undefined
    const ruleNodes = kids(element, 'cfRule', X14_MAIN_NS)
    if (getChildren(element, 'cfRule').length !== ruleNodes.length) {
      fin.diag('unsupported', 'cf-namespace', 'x14 conditionalFormatting holds a wrong-namespace cfRule, skipped')
    }
    // Ranges resolve after pairing (paired rules inherit base scope).
    if (sqrefText !== undefined && sqrefText.length > CF_LIMITS.textUnits * 16) {
      fin.failSheet('cf-invalid-rule', 'cf-sqref', 'x14 conditionalFormatting sqref text is unrepresentable')
      continue
    }
    for (const ruleNode of ruleNodes) {
      x14Bodies.push({ element, node: ruleNode, sqrefText, docOrder: order++ })
    }
  }
  if (fin.aborted) return

  // Pairing index over normalized GUIDs.
  const basesById = new Map<string, PendingRule[]>()
  const extsById = new Map<string, PendingRule[]>()
  for (const p of pending) {
    const ids: string[] = []
    for (const extList of kids(p.node, 'extLst', sheetNs)) {
      for (const ext of kids(extList, 'ext', sheetNs)) {
        if (attrs(ext).uri !== DB_X14_EXTENSION_URI) continue
        for (const id of kids(ext, 'id', X14_MAIN_NS)) {
          const text = textOf(id).trim()
          if (text.length > 0 && text.length <= 256) ids.push(text)
        }
      }
    }
    if (ids.length > 1) {
      p.pairId = `__ambiguous__${p.docOrder}`
      p.pairRaw = ids[0]
      const list = basesById.get(p.pairId) ?? []
      list.push(p)
      basesById.set(p.pairId, list)
      continue
    }
    const norm = normPairId(ids[0])
    p.pairId = norm
    p.pairRaw = ids[0]
    if (norm !== undefined) {
      const list = basesById.get(norm) ?? []
      list.push(p)
      basesById.set(norm, list)
    }
  }
  for (const x of x14Bodies) {
    const a = attrs(x.node)
    const norm = normPairId(a['id'])
    const p: PendingRule = { kind: 'x14', node: x.node, element: x.element, ranges: undefined, sqrefText: x.sqrefText, docOrder: x.docOrder, pairId: norm, pairRaw: a['id'] }
    pending.push(p)
    if (norm !== undefined) {
      const list = extsById.get(norm) ?? []
      list.push(p)
      extsById.set(norm, list)
    }
  }

  const mergedBases = new Set<PendingRule>()
  const mergedExts = new Map<PendingRule, PendingRule>()
  const mergedExtFacts = new Map<PendingRule, { cfvoCount: number; cfvoMissingType: boolean }>()
  const invalidBases = new Set<PendingRule>()
  const rectsEqual = (a: ParsedRange, b: ParsedRange): boolean =>
    a.firstRow === b.firstRow && a.lastRow === b.lastRow && a.firstCol === b.firstCol && a.lastCol === b.lastCol
  const allIds = new Set([...basesById.keys(), ...extsById.keys()])
  for (const id of allIds) {
    const bases = basesById.get(id) ?? []
    const exts = extsById.get(id) ?? []
    const loneBase = bases.length === 1 ? bases[0]! : undefined
    const loneExt = exts.length === 1 ? exts[0]! : undefined
    const baseType = loneBase ? (attrs(loneBase.node)['type'] ?? '').trim() : ''
    const extAttrs = loneExt ? attrs(loneExt.node) : undefined
    const extType = (extAttrs?.['type'] ?? '').trim()
    let scopeOk = true
    if (loneBase && loneExt && loneBase.ranges) {
      const extSqref = loneExt.sqrefText?.trim()
      if (extSqref) {
        const r = parseSqref(extSqref)
        scopeOk =
          r !== undefined &&
          r.length === loneBase.ranges.length &&
          r.every((rr, i) => rectsEqual(rr, loneBase.ranges![i]!))
      }
    }
    if (
      loneBase && loneExt && baseType === 'dataBar' && extType === 'dataBar' &&
      extAttrs?.['priority'] === undefined &&
      kids(loneExt.node, 'dataBar', X14_MAIN_NS)[0] !== undefined &&
      kids(kids(loneExt.node, 'dataBar', X14_MAIN_NS)[0], 'fillColor', X14_MAIN_NS)[0] === undefined &&
      scopeOk
    ) {
      const facts = readPairedExtFacts(fin, loneExt.node)
      if (facts.ok) {
        mergedBases.add(loneBase)
        mergedExts.set(loneExt, loneBase)
        mergedExtFacts.set(loneBase, { cfvoCount: facts.cfvoCount, cfvoMissingType: facts.cfvoMissingType })
        continue
      }
    }
    for (const b of bases) invalidBases.add(b)
    for (const e of exts) {
      const a = attrs(e.node)
      if (a['priority'] !== undefined && isIntText(a['priority']) && Number(a['priority']) >= 1) {
        fin.diag('unsupported', 'cf-invalid-x14-link', `x14 rule id="${a['id'] ?? ''}" has no clean pairing target; parsed standalone`)
      } else {
        fin.failSheet('cf-invalid-x14-link', 'cf-invalid-x14-link', `x14 rule id="${a['id'] ?? ''}" cannot be ordered without a pairing target`)
      }
    }
  }

  // Resolve x14 scopes after pairing (paired rules inherit base scope).
  const scopeRefused = new Set<PendingRule>()
  {
    const byElement = new Map<XmlNode, PendingRule[]>()
    for (const p of pending) {
      if (p.kind !== 'x14' || mergedExts.has(p)) continue
      const list = byElement.get(p.element!) ?? []
      list.push(p)
      byElement.set(p.element!, list)
    }
    for (const [, rules] of byElement) {
      if (fin.aborted) break
      // Unorderable extensions already raised the sheet sentinel, which aborts.
      const sqrefText = rules[0]!.sqrefText
      const admitted = admitElement(sqrefText, rules.length)
      if (admitted === undefined) continue
      if (admitted === 'rule-refusal') {
        for (const r of rules) {
          const t = (attrs(r.node)['type'] ?? '').trim()
          pushScopeRefusal(fin, r.node, 'x14', r.docOrder, KNOWN_TYPES.has(t) ? (t as CfRuleType) : 'unknown')
          scopeRefused.add(r)
        }
        continue
      }
      for (const r of rules) r.ranges = admitted
    }
  }

  // Finalize in document order.
  const ordered = [...pending].sort((a, b) => a.docOrder - b.docOrder)
  let ruleId = 0
  const pushRuleRefusal = (
    kind: 'base' | 'x14',
    node: XmlNode,
    ranges: ParsedRange[],
    docOrder: number,
    fallbackType: CfRuleType | 'unknown',
  ): void => {
    const a = attrs(node)
    const priority = a['priority'] !== undefined && isIntText(a['priority']) && Number(a['priority']) >= 1
      ? Number(a['priority'])
      : 0
    fin.diag('unsupported', 'cf-resource-limit', `rule priority=${priority} exceeds a per-rule bound`)
    fin.admit(256)
    fin.keep(200)
    fin.rules.push({
      id: -1, priority, documentOrder: docOrder, ranges, stopIfTrue: false,
      writes: new Set(['ALL']),
      definition: { kind: 'unresolved', type: fallbackType, reason: 'cf-resource-limit', origin: { dialect: kind, formulas: [] } },
      dxf: undefined,
    })
  }
  const buildOne = (
    kind: 'base' | 'x14',
    node: XmlNode,
    ns: string | undefined,
    ranges: ParsedRange[],
    docOrder: number,
    pairRaw: string | undefined,
    fallbackType: CfRuleType | 'unknown',
    mergeFacts?: { cfvoCount: number; cfvoMissingType: boolean },
  ): void => {
    const read = readRuleInputs(fin, node, ns, kind, pairRaw)
    if (!read.ok) {
      pushRuleRefusal(kind, node, ranges, docOrder, fallbackType)
      return
    }
    if (mergeFacts) {
      read.inputs.cfvoCount += mergeFacts.cfvoCount
      read.inputs.cfvoMissingType = read.inputs.cfvoMissingType || mergeFacts.cfvoMissingType
      if (read.inputs.cfvoCount > CF_LIMITS.cfvo) {
        pushRuleRefusal(kind, node, ranges, docOrder, fallbackType)
        return
      }
    }
    if (!fin.reserveMeta(128 + read.inputs.formulas.reduce((n, f) => n + f.length * 2, 0))) {
      fin.failSheet('cf-resource-limit', 'cf-resource-limit', 'sheet CF metadata exceeds the per-sheet cap')
      return
    }
    fin.admit(read.bound)
    const actual = buildRuleBody(fin, read.inputs, ranges, docOrder, fallbackType)
    // True-up: estimation slack is bounded; only genuine exhaustion throws.
    if (actual > read.bound) fin.admit(actual - read.bound)
    fin.keep(actual)
  }

  for (const p of ordered) {
    if (fin.aborted) break
    if (mergedExts.has(p) || scopeRefused.has(p)) continue
    // Invariant: every surviving rule has resolved scope (unresolved scope raises the sentinel, which aborts).
    const ranges = p.ranges!
    if (invalidBases.has(p)) {
      const a = attrs(p.node)
      invalidLinkRule(
        fin,
        a['priority'] !== undefined && isIntText(a['priority']) && Number(a['priority']) >= 1 ? Number(a['priority']) : 0,
        p.docOrder,
        ranges,
        { dialect: 'base', ...(p.pairRaw ? { guid: p.pairRaw } : {}), formulas: [] },
        (a['type'] ?? '').trim(),
        `base rule priority="${a['priority'] ?? ''}" has no valid x14 companion`,
      )
      continue
    }
    if (p.kind === 'base') {
      const type = (attrs(p.node)['type'] ?? '').trim()
      const facts = mergedBases.has(p) ? mergedExtFacts.get(p) : undefined
      buildOne('base', p.node, sheetNs, ranges, p.docOrder, p.pairRaw, KNOWN_TYPES.has(type) ? (type as CfRuleType) : 'unknown', facts)
      continue
    }
    const type = (attrs(p.node)['type'] ?? '').trim()
    buildOne('x14', p.node, X14_MAIN_NS, ranges, p.docOrder, p.pairRaw, KNOWN_TYPES.has(type) ? (type as CfRuleType) : 'unknown')
  }
  for (const rule of fin.rules) {
    if (rule.id === -1) (rule as { id: number }).id = ruleId++
  }
}

function finalizeCollected(
  subtrees: CfSubtrees,
  styles: CfStylesSource,
  memory: CfMemory,
  meta: CfParseMeta,
): CfLease<CfSheetRules> {
  const diagHandle = createCfDiagnostics(memory)
  const caps = createCfCaps(memory)
  const holdRes = memory.reserve(256)
  if (holdRes.kind === 'unresolved') throw new Error('CF admission exhausted')
  const fin = new SheetFinalizer(holdRes, styles, memory, caps, diagHandle, meta)
  try {
    if (subtrees.truncated) {
      fin.failSheet('cf-resource-limit', 'cf-resource-limit', 'CF input was truncated or exceeded the streaming capture bound')
    } else {
      collectAndBuild(fin, subtrees.sheetNs, [...subtrees.base], [...subtrees.x14])
    }
    if (fin.diagHandle.refused) throw new Error('CF admission exhausted')
    const value: CfSheetRules = fin.sentinel
      ? { rules: [], unresolvedSheet: { reason: fin.sentinel.reason, writes: 'ALL' }, diagnostics: fin.diagnostics }
      : { rules: fin.rules, diagnostics: fin.diagnostics }
    const total = fin.retained + 256
    if (total > fin.charged && !fin.hold.grow(total)) throw new Error('CF admission exhausted')
    fin.charged = Math.max(fin.charged, total)
    const lease = fin.hold.adopt(value, total)
    retainedHandles.set(lease, { caps, diagHandle })
    return lease
  } catch (error) {
    try {
      caps.dispose()
    } catch {
      /* release-only path; never masks the original failure */
    }
    try {
      diagHandle.dispose()
    } catch {
      /* release-only path; never masks the original failure */
    }
    try {
      holdRes.release()
    } catch {
      /* release-only path; never masks the original failure */
    }
    throw error
  }
}

export interface CfSubtrees {
  readonly base: readonly XmlNode[]
  readonly x14: readonly XmlNode[]
  readonly sheetNs: string | undefined
  /** Truncated input or structural capture overflow: withhold everything. */
  readonly truncated?: boolean
}

/** Shared finalizer for resident nodes and streaming-captured nodes alike. */
export function finalizeCfSubtrees(
  subtrees: CfSubtrees,
  styles: CfStylesSource,
  memory: CfMemory,
  meta: CfParseMeta,
): CfLease<CfSheetRules> {
  return finalizeCollected(subtrees, styles, memory, meta)
}

/**
 * Parse one worksheet's conditional formatting. `worksheet` is the
 * `<worksheet>` element (or an equivalent subtree root); `styles` carries the
 * workbook `dxfs` nodes and colour context; `memory` is the workbook-shared
 * parse account (see `cfMemoryFor`).
 */
export function parseCfSheet(
  worksheet: XmlNode,
  styles: CfStylesSource,
  memory: CfMemory,
  meta: CfParseMeta = {},
): CfLease<CfSheetRules> {
  const sheetNs = elementNamespaceUri(worksheet)
  const base = getChildren(worksheet, 'conditionalFormatting')
  const x14: XmlNode[] = []
  for (const extList of getChildren(worksheet, 'extLst')) {
    if (elementNamespaceUri(extList) !== sheetNs) continue
    for (const ext of getChildren(extList, 'ext')) {
      if (elementNamespaceUri(ext) !== sheetNs) continue
      if (attrs(ext).uri !== CF_X14_EXTENSION_URI) continue
      // Namespace-filtered in collectAndBuild; collect every same-local child first
      // so wrong-namespace content is diagnosed rather than silently skipped.
      for (const cfs of getChildren(ext, 'conditionalFormattings')) {
        x14.push(...getChildren(cfs, 'conditionalFormatting'))
      }
    }
  }
  return finalizeCollected({ base, x14, sheetNs }, styles, memory, meta)
}
