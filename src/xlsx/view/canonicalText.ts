import { CASE_FOLD_MULTIPLE, CASE_FOLD_RUNS, PROPERTY_PALETTE, PROPERTY_RUNS, PROPERTY_VALUES, TURKIC_FOLD, UNICODE_VERSION } from './unicode/textData'

/** Canonical text-matching representation (spec §7.1, amended 2026-10-11):
 * NFD, remove Default_Ignorable_Code_Point except ZWJ, variation selectors
 * and tag characters (retained so emoji ZWJ/variation/tag sequences stay
 * single graphemes), locale-aware full case folding, NFD again; then extended
 * grapheme clusters of that result. Case folding,
 * ignorables and segmentation use tables pinned to UNICODE_VERSION, never the
 * host ICU. NFD uses String.prototype.normalize: Unicode normalization is
 * stable across versions for assigned characters, so hosts agree except for
 * code points assigned after the host's own Unicode version (left unchanged
 * there). Raw values are never modified; the canonical string is a derived,
 * cacheable search key.
 */
export { UNICODE_VERSION }
/** Bumped whenever canonicalSearchKey output can change for the same Unicode data. */
export const CANONICAL_KEY_VERSION = 1
export type FoldVariant = 'default' | 'turkic'

// Property byte layout, mirrored by scripts/generate-unicode-text-data.ts.
const GCB_MASK = 15
const CR = 1, LF = 2, CONTROL = 3, EXTEND = 4, ZWJ = 5, REGIONAL = 6, PREPEND = 7, SPACING_MARK = 8
const L = 9, V = 10, T = 11, LV = 12, LVT = 13
const EXT_PICT = 16, INCB_SHIFT = 5, INCB_CONSONANT = 1, INCB_LINKER = 2, INCB_EXTEND = 3
const DEFAULT_IGNORABLE = 128
const HANGUL_FIRST = 0xac00, HANGUL_LAST = 0xd7a3

let bmp: Uint8Array | undefined
let astralStarts: Uint32Array, astralValues: Uint8Array
function decodeProperties(): Uint8Array {
  const table = new Uint8Array(0x10000), starts: number[] = [], values: number[] = []
  let start = 0, digits = ''
  for (const char of PROPERTY_RUNS) {
    const index = PROPERTY_PALETTE.indexOf(char)
    if (index < 0) { digits += char; continue }
    start += parseInt(digits, 36)
    digits = ''
    starts.push(start)
    values.push(PROPERTY_VALUES[index])
  }
  const astral: number[] = [], astralValue: number[] = []
  for (let i = 0; i < starts.length; i++) {
    const end = i + 1 < starts.length ? starts[i + 1] : 0x110000
    if (starts[i] < 0x10000) table.fill(values[i], starts[i], Math.min(end, 0x10000))
    if (end > 0x10000) { astral.push(Math.max(starts[i], 0x10000)); astralValue.push(values[i]) }
  }
  astralStarts = Uint32Array.from(astral)
  astralValues = Uint8Array.from(astralValue)
  return table
}
/** Pinned property byte for a code point (see unicode/textData.ts). */
export function unicodeProperty(codePoint: number): number {
  if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff) throw new RangeError('Invalid code point')
  const table = bmp ??= decodeProperties()
  let value: number
  if (codePoint < 0x10000) value = table[codePoint]
  else {
    let low = 0, high = astralStarts.length - 1
    while (low < high) {
      const middle = (low + high + 1) >> 1
      if (astralStarts[middle] <= codePoint) low = middle
      else high = middle - 1
    }
    value = astralValues[low]
  }
  if (codePoint >= HANGUL_FIRST && codePoint <= HANGUL_LAST) value = (value & ~GCB_MASK) | ((codePoint - HANGUL_FIRST) % 28 === 0 ? LV : LVT)
  return value
}

/** Default-ignorables kept by spec §7.1 (clarified 2026-10-11): U+200D ZWJ,
 * variation selectors U+FE00-FE0F/U+E0100-E01EF and tags U+E0020-E007F. */
function retainedIgnorable(point: number): boolean {
  return point === 0x200d || (point >= 0xfe00 && point <= 0xfe0f) || (point >= 0xe0020 && point <= 0xe007f) || (point >= 0xe0100 && point <= 0xe01ef)
}
function removedIgnorable(point: number): boolean {
  return (unicodeProperty(point) & DEFAULT_IGNORABLE) !== 0 && !retainedIgnorable(point)
}

let folds: Map<number, string> | undefined
function decodeFolds(): Map<number, string> {
  const map = new Map<number, string>()
  let start = 0
  for (const run of CASE_FOLD_RUNS.split(',')) {
    const [delta, count, stride, offset] = run.split('.')
    start += parseInt(delta, 36)
    for (let k = 0, size = parseInt(count, 36); k < size; k++) {
      const point = start + k * Number(stride)
      map.set(point, String.fromCodePoint(point + parseInt(offset, 36)))
    }
  }
  for (const entry of CASE_FOLD_MULTIPLE.split(',')) {
    const [point, ...target] = entry.split('.').map(value => parseInt(value, 36))
    map.set(point, String.fromCodePoint(...target))
  }
  return map
}
/** Full case folding (CaseFolding status C+F, or T overrides for Turkic), without context. */
export function fullCaseFold(codePoint: number, turkic: boolean): string | undefined {
  const turkish = turkic ? TURKIC_FOLD[codePoint] : undefined
  if (turkish !== undefined) return String.fromCodePoint(turkish)
  return (folds ??= decodeFolds()).get(codePoint)
}

/** Turkish and Azeri are the only languages with CaseFolding T mappings. */
export function isTurkicLocale(locale: string): boolean {
  if (typeof locale !== 'string') throw new RangeError('Invalid locale')
  const language = locale.split(/[-_]/)[0].toLowerCase()
  return language === 'tr' || language === 'az'
}
export function foldVariantOf(locale: string): FoldVariant {
  return isTurkicLocale(locale) ? 'turkic' : 'default'
}
/** Identity of the key function for caches/persistence (S8.2): canonicalizer
 * version, pinned Unicode data version and fold variant. */
export function canonicalKeyStrategyId(variant: FoldVariant): string {
  if (variant !== 'default' && variant !== 'turkic') throw new RangeError('Invalid fold variant')
  return `canonical-text/v${CANONICAL_KEY_VERSION}/unicode-${UNICODE_VERSION}/${variant}`
}

const beforeDotAbove = new Map<number, boolean>()
/** True when 0 < ccc < 230: canonical ordering moves the mark before U+0307. */
function sortsBeforeDotAbove(point: number): boolean {
  let known = beforeDotAbove.get(point)
  if (known === undefined) {
    const mark = String.fromCodePoint(point)
    known = ('\u0307' + mark).normalize('NFD') === mark + '\u0307'
    if (beforeDotAbove.size < 4096) beforeDotAbove.set(point, known)
  }
  return known
}

let changes: Uint8Array | undefined
/** BMP code units that canonicalization rewrites: 1 ignorable, 2 folded. */
function decodeChanges(): Uint8Array {
  const table = new Uint8Array(0x10000), properties = bmp ??= decodeProperties()
  for (let unit = 0; unit < 0x10000; unit++) if (properties[unit] & DEFAULT_IGNORABLE && !retainedIgnorable(unit)) table[unit] = 1
  for (const point of (folds ??= decodeFolds()).keys()) if (point < 0x10000) table[point] |= 2
  return table
}

const ASCII = /^[\x00-\x7f]*$/
/** The canonical matching form. A lone surrogate passes through unchanged. */
export function canonicalize(text: string, turkic: boolean): string {
  // ASCII has no decompositions or default-ignorables; folding is A-Z only.
  if (ASCII.test(text)) return (turkic ? text.replace(/I/g, '\u0131') : text).toLowerCase()
  const nfd = text.normalize('NFD'), length = nfd.length, rewrite = changes ??= decodeChanges()
  // Copy unchanged runs as slices; per-character concatenation dominates otherwise.
  let folded = '', run = 0
  for (let i = 0; i < length;) {
    const unit = nfd.charCodeAt(i)
    let point = unit, size = 1, ignorable: boolean
    if (unit >= 0xd800 && unit <= 0xdfff) {
      point = nfd.codePointAt(i)!
      if (point === unit) { i++; continue }
      size = 2
      ignorable = removedIgnorable(point)
      if (!ignorable && !folds!.has(point)) { i += 2; continue }
    } else {
      const change = rewrite[unit]
      if (!change && !(turkic && unit === 0x49)) { i++; continue }
      ignorable = (change & 1) !== 0
    }
    folded += nfd.slice(run, i)
    i += size
    run = i
    if (ignorable) continue
    if (turkic && point === 0x49) {
      // SpecialCasing After_I/Before_Dot: in tr/az, I followed by U+0307 with
      // no intervening ccc 0/230 mark is the decomposed İ and folds to i, which
      // keeps NFD-first folding canonically equivalent to folding U+0130.
      let marks = '', j = i, dotted = false
      while (j < length) {
        const next = nfd.codePointAt(j)!, step = next > 0xffff ? 2 : 1
        if (removedIgnorable(next)) { j += step; continue }
        if (next === 0x307) { dotted = true; j += step; break }
        if (!sortsBeforeDotAbove(next)) break
        marks += String.fromCodePoint(next)
        j += step
      }
      if (dotted) { folded += 'i' + marks; i = run = j } else folded += '\u0131'
      continue
    }
    folded += fullCaseFold(point, turkic)!
  }
  if (run === 0) return nfd
  return (folded + nfd.slice(run)).normalize('NFD')
}

/** Derived, cacheable search key for a raw value, given a locale or a fold
 * variant ('default' | 'turkic'). Only the fold variant changes the result, so
 * cache identity is canonicalKeyStrategyId(foldVariantOf(locale)). */
export function canonicalSearchKey(text: string, localeOrVariant: string): string {
  if (typeof text !== 'string') throw new RangeError('Invalid text')
  const turkic = localeOrVariant === 'turkic' || (localeOrVariant !== 'default' && isTurkicLocale(localeOrVariant))
  return canonicalize(text, turkic)
}

/** Writes the end offset of each extended grapheme cluster of `text` into
 * `ends` (capacity at least text.length) and returns the cluster count, under
 * UAX #29 for Unicode 16.0 (GB3-GB13 including GB9c conjunct breaks and GB11
 * emoji sequences). Per spec §7.1 a lone surrogate is always its own unit, so
 * it is treated as Control here (UAX #29 itself would let marks attach to it).
 * Stops after the first end at or beyond `stopAt`. */
export function graphemeEnds(text: string, ends: Int32Array, stopAt = Infinity): number {
  const length = text.length, table = bmp ??= decodeProperties()
  let count = 0, previous = -1, regionalOdd = false, emoji = 0, conjunct = 0
  for (let i = 0; i < length;) {
    const point = text.codePointAt(i)!
    const property = point >= 0xd800 && point <= 0xdfff ? CONTROL
      : point < HANGUL_FIRST ? table[point] : unicodeProperty(point)
    const kind = property & GCB_MASK, incb = (property >> INCB_SHIFT) & 3
    if (previous >= 0) {
      let join: boolean
      if (previous === CR) join = kind === LF
      else if (previous === LF || previous === CONTROL || kind === CR || kind === LF || kind === CONTROL) join = false
      else if (previous === L) join = kind === L || kind === V || kind === LV || kind === LVT || kind === EXTEND || kind === ZWJ || kind === SPACING_MARK
      else join = ((previous === LV || previous === V) && (kind === V || kind === T))
        || ((previous === LVT || previous === T) && kind === T)
        || kind === EXTEND || kind === ZWJ || kind === SPACING_MARK || previous === PREPEND
        || (conjunct === 2 && incb === INCB_CONSONANT)
        || (emoji === 2 && (property & EXT_PICT) !== 0)
        || (previous === REGIONAL && kind === REGIONAL && regionalOdd)
      if (!join) {
        ends[count++] = i
        if (i >= stopAt) return count
      }
    }
    // Context for GB9c, GB11 and GB12/13 after consuming this code point.
    regionalOdd = kind === REGIONAL ? !(previous === REGIONAL && regionalOdd) : false
    emoji = property & EXT_PICT ? 1 : emoji === 1 && kind === EXTEND ? 1 : emoji === 1 && kind === ZWJ ? 2 : 0
    conjunct = incb === INCB_CONSONANT ? 1 : incb === INCB_LINKER ? (conjunct ? 2 : 0) : incb === INCB_EXTEND ? conjunct : 0
    previous = kind
    i += point > 0xffff ? 2 : 1
  }
  if (length) ends[count++] = length
  return count
}

export function graphemeClusters(text: string): string[] {
  const ends = new Int32Array(text.length), clusters: string[] = []
  for (let k = 0, count = graphemeEnds(text, ends), start = 0; k < count; start = ends[k++]) clusters.push(text.slice(start, ends[k]))
  return clusters
}

const SIMPLE_ASCII = /^[\x00-\x0c\x0e-\x7f]*$/
/** True when every code-unit offset of `key` is a grapheme boundary (no marks,
 * joiners, surrogate pairs, Hangul jamo, CR LF, prepend characters ...), so
 * plain substring tests on it already obey the boundary rule. */
export function isSimpleKey(key: string): boolean {
  if (SIMPLE_ASCII.test(key)) return true
  return graphemeEnds(key, new Int32Array(key.length)) === key.length
}
/** Whether `offset` (0..key.length) is an extended-grapheme boundary of `key`.
 * Linear in `offset`; use keyContains/keyBegins/keyEnds to test candidates. */
export function isGraphemeBoundary(key: string, offset: number): boolean {
  if (!Number.isInteger(offset) || offset < 0 || offset > key.length) throw new RangeError('Invalid key offset')
  if (offset === 0 || offset === key.length) return true
  const ends = new Int32Array(key.length), count = graphemeEnds(key, ends, offset)
  return ends[count - 1] === offset
}
/** Canonical-key begins with the canonical query at a grapheme boundary. An
 * empty query is no search term and matches. `simple` (from isSimpleKey) may
 * be passed by a caller that stored it, to skip boundary verification. */
export function keyBegins(key: string, query: string, simple?: boolean): boolean {
  return key.startsWith(query) && (simple === true || isGraphemeBoundary(key, query.length))
}
export function keyEnds(key: string, query: string, simple?: boolean): boolean {
  return key.endsWith(query) && (simple === true || isGraphemeBoundary(key, key.length - query.length))
}
/** Canonical-key contains the canonical query with both endpoints on grapheme
 * boundaries of the key (not plain includes: key('é') includes 'e', but
 * contains('é', 'e') is false). Linear: code-unit KMP plus one segmentation. */
export function keyContains(key: string, query: string, simple?: boolean): boolean {
  if (!query) return true
  if (simple === true || !key.includes(query)) return key.includes(query)
  const length = key.length, ends = new Int32Array(length), count = graphemeEnds(key, ends)
  if (count === length) return true
  const boundary = new Uint8Array(length + 1)
  boundary[0] = 1
  for (let k = 0; k < count; k++) boundary[ends[k]] = 1
  const failure = new Int32Array(query.length)
  for (let i = 1, matched = 0; i < query.length; i++) {
    while (matched > 0 && query.charCodeAt(i) !== query.charCodeAt(matched)) matched = failure[matched - 1]
    if (query.charCodeAt(i) === query.charCodeAt(matched)) matched++
    failure[i] = matched
  }
  for (let i = 0, matched = 0; i < length; i++) {
    const unit = key.charCodeAt(i)
    while (matched > 0 && unit !== query.charCodeAt(matched)) matched = failure[matched - 1]
    if (unit === query.charCodeAt(matched)) matched++
    if (matched === query.length) {
      if (boundary[i + 1] && boundary[i + 1 - matched]) return true
      matched = failure[matched - 1]
    }
  }
  return false
}
