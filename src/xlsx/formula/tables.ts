/** B0 table metadata extraction (additive, no evaluation). */
import { XMLParser } from 'fast-xml-parser'
import { attrs, getChildren, type XmlNode } from '../../core/xml'
import type { TableMetadata, TableSyntax } from './types'

// Local document-element probe using the installed parser. preserveOrder
// reports the actual root QName (prefixes/Unicode preserved); the shared
// parseXmlOrdered helper discards it, and no core/xml change is permitted.
const rootProbe = new XMLParser({ ignoreAttributes: false, preserveOrder: true })

function isUnsignedIntLexical(raw: unknown): raw is string {
  if (typeof raw !== 'string') return false
  const t = raw.trim()
  if (!/^\d+$/.test(t)) return false
  // xsd:unsignedInt domain is 0..4294967295 (W3C xmlschema-2#unsignedInt).
  // Compare digit length first so long strings are never rounded into range;
  // ten or fewer digits are exactly representable for the bound check.
  const digits = t.replace(/^0+/, '') || '0'
  if (digits.length > 10) return false
  return parseInt(digits, 10) <= 4294967295
}

function cellLettersToIndex(letters: string, digits: string): [number, number] {
  let col = 0
  for (const ch of letters) col = col * 26 + (ch.charCodeAt(0) - 64)
  return [parseInt(digits, 10) - 1, col - 1]
}

function extentOf(ref: unknown): { firstRow: number; firstCol: number; cols: number; rows: number } | undefined {
  if (typeof ref !== 'string') return undefined
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref)
  if (!m) return undefined
  const [r1, c1] = cellLettersToIndex(m[1], m[2])
  const [r2, c2] = m[3] !== undefined ? cellLettersToIndex(m[3], m[4]) : [r1, c1]
  const minRow = Math.min(r1, r2)
  const minCol = Math.min(c1, c2)
  const maxRow = Math.max(r1, r2)
  const maxCol = Math.max(c1, c2)
  return { firstRow: minRow, firstCol: minCol, cols: maxCol - minCol + 1, rows: maxRow - minRow + 1 }
}

function parseCount(raw: string | undefined, absentDefault: 0 | 1): 0 | 1 | undefined {
  if (raw === undefined) return absentDefault
  // xsd:unsignedInt lexical: digits only ("00"->0, "01"->1). Boolean words
  // and out-of-domain values are invalid, never guessed.
  if (!/^\s*\d+\s*$/.test(raw)) return undefined
  const n = parseInt(raw, 10)
  return n === 0 ? 0 : n === 1 ? 1 : undefined
}

/**
 * Verify a table part's document element from its source XML via the parsed
 * root QName: any prefix (ASCII or Unicode) is allowed, but the local name
 * must be exactly `table`. Legal BOM/prolog/comments/whitespace are handled
 * by the parser. A truncated ASCII prefix (`table:wrong`) or a non-table
 * Unicode root (`tableé`) never passes, and table-like attributes on an
 * arbitrary root are irrelevant.
 */
export function isTablePartSource(sourceText: string | undefined): boolean {

  if (typeof sourceText !== 'string' || sourceText === '') return false
  let v: unknown
  try {
    v = rootProbe.parse(sourceText.replace(/^﻿/, ''))
  } catch {
    return false
  }
  if (!Array.isArray(v)) return false
  for (const entry of v) {
    if (!entry || typeof entry !== 'object') continue
    const key = Object.keys(entry as Record<string, unknown>).find(
      (k) => !k.startsWith('?') && !k.startsWith('!') && !k.startsWith(':') && !k.startsWith('#'),
    )
    if (key === undefined) continue
    const local = key.includes(':') ? key.slice(key.indexOf(':') + 1) : key
    if (local.includes(':')) return false
    return local === 'table'
  }
  return false
}

/**
 * Build table metadata from a parsed `table` part. The caller supplies the
 * owning worksheet's stable `sheetId` and the resolved source `partPath`
 * (worksheet-owned tableParts relationship target, never a workbook rel).
 */
export function parseTableMetadata(
  table: XmlNode | undefined,
  sheetId: string,
  partPath: string,
): TableMetadata | undefined {
  if (!table) return undefined
  const a = attrs(table)
  // Mandatory bounded metadata is never fabricated: CT_Table requires
  // id/displayName/ref and CT_TableColumn requires id. Required identifiers
  // are xsd:unsignedInt schema values — missing, non-numeric or otherwise
  // invalid essentials keep the inventory unavailable. Valid source spelling
  // is preserved verbatim.
  const id = typeof a.id === 'string' ? a.id : ''
  if (!isUnsignedIntLexical(id)) return undefined
  // Table name is optional; the required displayName is the documented
  // fallback (MS-OE376). The inverse (displayName from optional name) is
  // fabrication and stays invalid.
  const displayName = typeof a.displayName === 'string' ? a.displayName : ''
  if (displayName === '') return undefined
  const nameRaw = typeof a.name === 'string' ? a.name : ''
  const name = nameRaw !== '' ? nameRaw : displayName
  const geo = extentOf(a.ref)
  if (!geo) return undefined
  const extent = { sheetId, firstCol: geo.firstCol, firstRow: geo.firstRow, cols: geo.cols, rows: geo.rows }
  const headerRowCount = parseCount(a.headerRowCount, 1)
  if (headerRowCount === undefined) return undefined
  const totalsRowCount = parseCount(a.totalsRowCount, 0)
  if (totalsRowCount === undefined) return undefined
  // Every CT_TableColumn carries required id and name together at its
  // original position. A column missing either essential field invalidates
  // the whole table part: fields are rejected, never filtered or shifted.
  const columns: Array<{ id: string; name: string; index: number }> = []
  let position = 0
  for (const colsNode of getChildren(table, 'tableColumns')) {
    for (const col of getChildren(colsNode, 'tableColumn')) {
      const ca = attrs(col)
      if (!isUnsignedIntLexical(ca.id)) return undefined
      if (typeof ca.name !== 'string' || ca.name === '') return undefined
      columns.push({ id: ca.id, name: ca.name, index: position })
      position++
    }
  }
  return {
    id,
    name,
    displayName,
    sheetId,
    partPath,
    extent,
    headerRowCount,
    totalsRowCount,
    columns,
  }
}

export type DecodedTableSyntax =
  | { kind: 'single'; syntax: TableSyntax }
  | { kind: 'union'; items: TableSyntax[] }
  | { kind: 'error'; message: string }

type TableItem = 'all' | 'data' | 'headers' | 'totals' | 'thisRow'

function decodeSpecifier(text: string): TableItem | null {
  const key = text.replace(/\s+/g, '').toLowerCase()
  if (key === '#all') return 'all'
  if (key === '#data') return 'data'
  if (key === '#headers') return 'headers'
  if (key === '#totals') return 'totals'
  if (key === '#thisrow') return 'thisRow'
  return null
}

/** Split a bracket body into top-level comma-separated `[..]` segments. */
function splitTopSegments(body: string): string[] | null {
  const out: string[] = []
  let depth = 0
  let start = 0
  for (let i = 0; i < body.length; i++) {
    const ch = body[i]
    if (ch === "'" && "'#@[]".includes(body[i + 1] ?? '\0')) { i++; continue }
    if (ch === '[') depth++
    else if (ch === ']') {
      depth--
      if (depth < 0) return null
    } else if (ch === ',' && depth === 0) {
      out.push(body.slice(start, i))
      start = i + 1
    }
  }
  out.push(body.slice(start))
  return out
}

function unwrapBrackets(segment: string): string | null {
  const t = segment.trim()
  if (t.length < 2 || t[0] !== '[' || t[t.length - 1] !== ']') return null
  return t.slice(1, -1)
}

/** Depth-zero colon index for adjacent-column selectors, or -1. */
function topLevelColon(segment: string): number {
  let depth = 0
  for (let i = 0; i < segment.length; i++) {
    const ch = segment[i]
    if (ch === "'" && "'#@[]".includes(segment[i + 1] ?? '\0')) { i++; continue }
    if (ch === '[') depth++
    else if (ch === ']') depth--
    else if (ch === ':' && depth === 0) return i
  }
  return -1
}

/** Apostrophes guard special header characters; doubled apostrophes are literal. */
function decodeColumnName(part: string): string {
  let name = ''
  for (let i = 0; i < part.length; i++) {
    if (part[i] === "'" && "'#@[]".includes(part[i + 1] ?? '\0')) {
      name += part[++i]
    } else {
      name += part[i]
    }
  }
  return name
}

/**
 * Decode a structured-reference raw source (`Table1[Amount]`,
 * `Table1[[#Headers],[Amount]]`, `[#Data]`) into ABI table syntax.
 * Apostrophe escapes (`''`) decode to `'`. Discrete table columns become
 * separate operands (parser builds the union); colon selectors encode one
 * adjacent column range. A bare `[Column]` selects the data body.
 */
export function decodeTableSyntax(tableName: string | undefined, raw: string): DecodedTableSyntax {
  const fail = (message: string): DecodedTableSyntax => ({ kind: 'error', message })
  const open = raw.indexOf('[')
  if (open < 0) return fail(`#NAME? Invalid structured reference "${raw}"`)
  const body = raw.slice(open)
  if (!body.startsWith('[') || !body.endsWith(']')) {
    return fail(`#NAME? Invalid structured reference "${raw}"`)
  }
  // Depth-balance the whole body first (escape-aware per group 6: a `'`
  // immediately before '[' or ']' is an escaped literal bracket, not a
  // structural token).
  let depth = 0
  for (let bi = 0; bi < body.length; bi++) {
    const ch = body[bi]
    if (ch === "'") {
      const next = body[bi + 1]
      if (next === "'" || next === '[' || next === ']') { bi++; continue }
      continue
    }
    if (ch === '[') depth++
    else if (ch === ']') {
      depth--
      if (depth < 0) return fail(`#NAME? Invalid structured reference "${raw}"`)
    }
  }
  if (depth !== 0) return fail(`#NAME? Invalid structured reference "${raw}"`)
  const shorthand = body[1] === '@'
  const inner = body.slice(shorthand ? 2 : 1, -1)
  // A doubled outer bracket group holds comma-separated inner segments;
  // a single group holds one bare part ([Amount], [#Data]).
  const segments = inner.startsWith('[') ? splitTopSegments(inner) : null
  if (inner.startsWith('[') && (!segments || segments.length === 0)) {
    return fail(`#NAME? Invalid structured reference "${raw}"`)
  }

  const items: TableItem[] = shorthand ? ['thisRow'] : []
  const singles: string[] = []
  let range: { from: string; to: string } | null = null
  const absorbPart = (part: string): boolean => {
    // Group 6 @ shorthand: a bare leading '@' column spec selects THIS ROW
    // ([@Amount] == [[#This Row],[Amount]]); decoded before the '#' branch.
    if (part.startsWith('@')) {
      const name = decodeColumnName(part.slice(1))
      if (name === '') return false
      if (!items.includes('thisRow')) items.push('thisRow')
      singles.push(name)
      return true
    }
    const spec = part.startsWith('#') ? decodeSpecifier(part) : null
    if (part.startsWith('#') && spec === null) return false
    if (spec !== null) {
      if (!items.includes(spec)) items.push(spec)
      return true
    }
    const name = decodeColumnName(part)
    if (name === '') return false
    singles.push(name)
    return true
  }
  if (segments === null) {
    if (inner.indexOf(':') >= 0) return fail(`#NAME? Invalid structured reference "${raw}"`)
    if (!absorbPart(inner)) return fail(`#NAME? Invalid structured reference "${raw}"`)
  } else {
    for (const seg of segments) {
      // Adjacent-column selector [A]:[C] at this level.
      const colon = topLevelColon(seg)
      if (colon >= 0) {
        const left = unwrapBrackets(seg.slice(0, colon))
        const right = unwrapBrackets(seg.slice(colon + 1))
        if (left === null || right === null) return fail(`#NAME? Invalid structured reference "${raw}"`)
        if (left.startsWith('#') || right.startsWith('#')) {
          return fail(`#NAME? Invalid structured reference "${raw}"`)
        }
        if (range !== null || singles.length > 0) return fail(`#NAME? Invalid structured reference "${raw}"`)
        const from = decodeColumnName(left)
        const to = decodeColumnName(right)
        if (from === '' || to === '') return fail(`#NAME? Invalid structured reference "${raw}"`)
        range = { from, to }
        continue
      }
      const part = unwrapBrackets(seg)
      if (part === null) {
        return fail(`#NAME? Invalid structured reference "${raw}"`)
      }
      if (!absorbPart(part)) return fail(`#NAME? Invalid structured reference "${raw}"`)
    }
  }

  const base = {
    ...(tableName !== undefined ? { table: tableName } : {}),
    items: items.length > 0 ? items : (['data'] as TableItem[]),
    raw,
  }
  if (range !== null) {
    return { kind: 'single', syntax: { ...base, columns: { kind: 'range', from: range.from, to: range.to } } }
  }
  if (singles.length === 0) {
    return { kind: 'single', syntax: { ...base, columns: { kind: 'all' } } }
  }
  if (singles.length === 1) {
    return { kind: 'single', syntax: { ...base, columns: { kind: 'single', name: singles[0] } } }
  }
  // Discrete columns are separate table operands sharing items.
  return {
    kind: 'union',
    items: singles.map((name) => ({ ...base, columns: { kind: 'single' as const, name } })),
  }
}
