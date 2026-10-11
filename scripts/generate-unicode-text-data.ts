/** Run: bun scripts/generate-unicode-text-data.ts /local/ucd-16.0.0-dir
 *
 * Generates the pinned Unicode tables used by canonical text matching
 * (spec §7.1, plan S8.5) from local copies of these Unicode 16.0.0 files:
 *   https://www.unicode.org/Public/16.0.0/ucd/CaseFolding.txt
 *   https://www.unicode.org/Public/16.0.0/ucd/DerivedCoreProperties.txt
 *   https://www.unicode.org/Public/16.0.0/ucd/auxiliary/GraphemeBreakProperty.txt
 *   https://www.unicode.org/Public/16.0.0/ucd/auxiliary/GraphemeBreakTest.txt
 *   https://www.unicode.org/Public/16.0.0/ucd/emoji/emoji-data.txt
 * Generation never fetches network data and refuses inputs whose SHA-256 differs
 * from the pins below. Raw UCD files are not committed; the outputs are
 * src/xlsx/view/unicode/textData.ts and the compact conformance fixture
 * tests/fixtures/unicode/grapheme-break-test-16.0.0.json. After writing, the
 * production segmenter is checked against every GraphemeBreakTest case.
 */
import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs'
import { join, resolve, dirname } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const VERSION = '16.0.0'
const BASE = `https://www.unicode.org/Public/${VERSION}/ucd/`
const INPUTS = {
  'CaseFolding.txt': { path: 'CaseFolding.txt', sha256: '6f1f9c588eb4a5c718d9e8f93b782685e5c7fec872cf05e8e6878053599e09bb' },
  'DerivedCoreProperties.txt': { path: 'DerivedCoreProperties.txt', sha256: '39d35161f2954497f69e08bdb9e701493f476a3d30222de20028feda36c1dabd' },
  'GraphemeBreakProperty.txt': { path: 'auxiliary/GraphemeBreakProperty.txt', sha256: 'c29360bd6f7132811d701d29069541e827eb44bfc4c8fbde8c370d6982689dc1' },
  'GraphemeBreakTest.txt': { path: 'auxiliary/GraphemeBreakTest.txt', sha256: 'ee2b9354d270ac061b29f09662cafea06341d77e704b8cc6bd72aaeeda363cb5' },
  'emoji-data.txt': { path: 'emoji/emoji-data.txt', sha256: 'f1365a5173eee18e1f98b240cdc492e84a25f1ce7e0c9d1094eb29c41a22696a' },
} as const
type InputName = keyof typeof INPUTS

const dir = process.argv[2]
if (!dir) throw new Error(`Provide a local directory containing the Unicode ${VERSION} files: ${Object.keys(INPUTS).join(', ')}`)
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
function input(name: InputName): string {
  const bytes = readFileSync(join(dir, name))
  const actual = createHash('sha256').update(bytes).digest('hex')
  if (actual !== INPUTS[name].sha256) throw new Error(`Pinned ${name} SHA-256 mismatch: expected ${INPUTS[name].sha256}; received ${actual}`)
  const text = bytes.toString('utf8')
  if (!text.startsWith(`# ${name.replace('.txt', '')}-${VERSION}.txt`) && !(name === 'emoji-data.txt' && text.startsWith('# emoji-data.txt'))) throw new Error(`${name} is not version ${VERSION}`)
  return text
}
/** Yields [first, last, fields...] for every data line. */
function* records(text: string): Generator<[number, number, string[]]> {
  for (const line of text.split('\n')) {
    const data = line.replace(/#.*/, '').trim()
    if (!data) continue
    const fields = data.split(';').map(field => field.trim())
    const [a, b] = fields[0].split('..')
    yield [parseInt(a, 16), parseInt(b ?? a, 16), fields.slice(1)]
  }
}

// Combined per-code-point property byte; the layout is mirrored in canonicalText.ts.
const GCB = ['Other', 'CR', 'LF', 'Control', 'Extend', 'ZWJ', 'Regional_Indicator', 'Prepend', 'SpacingMark', 'L', 'V', 'T', 'LV', 'LVT'] as const
const INCB = ['None', 'Consonant', 'Linker', 'Extend'] as const
const EXT_PICT = 16, INCB_SHIFT = 5, DEFAULT_IGNORABLE = 128
const MAX = 0x10ffff
const props = new Uint8Array(MAX + 1)
const hangulBase = 0xac00, hangulLast = 0xd7a3

let listed = 0
for (const [first, last, [value]] of records(input('GraphemeBreakProperty.txt'))) {
  const code = GCB.indexOf(value as typeof GCB[number])
  if (code < 0) throw new Error(`Unknown Grapheme_Cluster_Break ${value}`)
  for (let cp = first; cp <= last; cp++) {
    if (props[cp] & 15) throw new Error(`Duplicate Grapheme_Cluster_Break at ${cp.toString(16)}`)
    // LV/LVT are algorithmic over the Hangul syllable block; the decoder
    // recomputes them, so verify instead of storing ~800 alternating runs.
    if (code === 12 || code === 13) {
      if (cp < hangulBase || cp > hangulLast || ((cp - hangulBase) % 28 === 0) !== (code === 12)) throw new Error(`Unexpected Hangul syllable property at ${cp.toString(16)}`)
      listed++
      continue
    }
    props[cp] |= code
    listed++
  }
}
for (let cp = hangulBase; cp <= hangulLast; cp++) if (props[cp] & 15) throw new Error('Hangul syllable block is not purely LV/LVT')

let ignorable = 0, incbCount = 0, pictographic = 0
const derived = input('DerivedCoreProperties.txt')
for (const [first, last, fields] of records(derived)) {
  if (fields[0] === 'Default_Ignorable_Code_Point') {
    for (let cp = first; cp <= last; cp++) { props[cp] |= DEFAULT_IGNORABLE; ignorable++ }
  } else if (fields[0] === 'InCB') {
    const code = INCB.indexOf(fields[1] as typeof INCB[number])
    if (code <= 0) throw new Error(`Unknown InCB ${fields[1]}`)
    for (let cp = first; cp <= last; cp++) { props[cp] |= code << INCB_SHIFT; incbCount++ }
  }
}
for (const [first, last, [value]] of records(input('emoji-data.txt'))) {
  if (value !== 'Extended_Pictographic') continue
  for (let cp = first; cp <= last; cp++) { props[cp] |= EXT_PICT; pictographic++ }
}

// Run-length encode the property bytes: base-36 start delta, then one palette
// character naming the run's property byte. No separators are needed because
// palette characters never occur in base-36 digits.
const PALETTE_CHARS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ!#$%&()*+-./:;<=>?@[]^_{|}~'
const palette: number[] = []
let encodedProps = '', runs = 0, previousStart = 0, previous = -1
for (let cp = 0; cp <= MAX; cp++) {
  if (props[cp] === previous) continue
  let index = palette.indexOf(props[cp])
  if (index < 0) { index = palette.length; palette.push(props[cp]) }
  if (index >= PALETTE_CHARS.length) throw new Error('Property palette overflow')
  encodedProps += (cp - previousStart).toString(36) + PALETTE_CHARS[index]
  previousStart = cp
  previous = props[cp]
  runs++
}

// Full case folding = status C plus F; Turkic T entries are kept separately.
const fold = new Map<number, number[]>(), turkic = new Map<number, number[]>()
for (const [first, , [status, mapping]] of records(input('CaseFolding.txt'))) {
  const target = mapping.split(' ').map(value => parseInt(value, 16))
  if (status === 'C' || status === 'F') {
    if (fold.has(first)) throw new Error(`Duplicate full folding at ${first.toString(16)}`)
    fold.set(first, target)
  } else if (status === 'T') turkic.set(first, target)
  else if (status !== 'S') throw new Error(`Unknown CaseFolding status ${status}`)
}
const turkicEntries = [...turkic].map(([cp, target]) => `${cp.toString(16)}>${target.map(value => value.toString(16)).join(' ')}`).join(',')
if (turkicEntries !== '49>131,130>69') throw new Error(`Unexpected Turkic CaseFolding entries: ${turkicEntries}`)
// Single code point mappings: runs of equal delta with stride 1 or 2.
const singles = [...fold].filter(([, target]) => target.length === 1).map(([cp, target]) => [cp, target[0] - cp] as const).sort((a, b) => a[0] - b[0])
const foldRuns: string[] = []
let last = 0
for (let i = 0; i < singles.length;) {
  const [start, delta] = singles[i]
  let stride = 1, count = 1
  if (i + 1 < singles.length && singles[i + 1][1] === delta && singles[i + 1][0] - start <= 2) {
    stride = singles[i + 1][0] - start
    while (i + count < singles.length && singles[i + count][1] === delta && singles[i + count][0] === start + count * stride) count++
  }
  foldRuns.push(`${(start - last).toString(36)}.${count.toString(36)}.${stride}.${delta.toString(36)}`)
  last = start
  i += count
}
const multiples = [...fold].filter(([, target]) => target.length > 1).sort((a, b) => a[0] - b[0])
const foldMultiple = multiples.map(([cp, target]) => `${cp.toString(36)}.${target.map(value => value.toString(36)).join('.')}`).join(',')

const hash = (name: InputName) => `${BASE}${INPUTS[name].path}\n *   SHA-256 ${INPUTS[name].sha256}`
const output = `/** Pinned Unicode ${VERSION} data for canonical text matching (spec §7.1).
 * GENERATED by scripts/generate-unicode-text-data.ts; do not edit by hand.
 * Sources:
 *   ${hash('CaseFolding.txt')}
 *   ${hash('DerivedCoreProperties.txt')}
 *   ${hash('GraphemeBreakProperty.txt')}
 *   ${hash('emoji-data.txt')}
 * Copyright © 1991-2024 Unicode, Inc. See ./LICENSE and ./NOTICE.md.
 *
 * PROPERTY_RUNS: base-36 start delta followed by a PROPERTY_PALETTE character.
 * Property byte: bits 0-3 Grapheme_Cluster_Break (${GCB.map((name, index) => `${index} ${name}`).join(', ')};
 * Hangul LV/LVT are algorithmic and stored as Other), bit 4 Extended_Pictographic,
 * bits 5-6 Indic_Conjunct_Break (${INCB.map((name, index) => `${index} ${name}`).join(', ')}), bit 7
 * Default_Ignorable_Code_Point. ${runs} runs; ${listed} GCB, ${pictographic} ExtPict, ${incbCount} InCB and
 * ${ignorable} Default_Ignorable code points.
 * CASE_FOLD_RUNS: full case folding (status C+F) single-code-point mappings as
 * comma-separated "startDelta.count.stride.delta" base-36 runs; CASE_FOLD_MULTIPLE
 * lists the ${multiples.length} status F expansions as "cp.target..." in base 36.
 * TURKIC_FOLD: status T (tr/az) overrides.
 */
export const UNICODE_VERSION = '${VERSION}'
export const PROPERTY_PALETTE = '${PALETTE_CHARS.slice(0, palette.length)}'
export const PROPERTY_VALUES = [${palette.join(',')}]
export const PROPERTY_RUNS = '${encodedProps}'
export const CASE_FOLD_RUNS = '${foldRuns.join(',')}'
export const CASE_FOLD_MULTIPLE = '${foldMultiple}'
export const TURKIC_FOLD: Readonly<Record<number, number>> = { 0x49: 0x131, 0x130: 0x69 }
`
const dataPath = join(root, 'src/xlsx/view/unicode/textData.ts')
mkdirSync(dirname(dataPath), { recursive: true })
writeFileSync(dataPath, output)

// Compact conformance fixture: clusters separated by '|', code points by '.'.
const cases: string[] = []
for (const line of input('GraphemeBreakTest.txt').split('\n')) {
  const data = line.replace(/#.*/, '').trim()
  if (!data) continue
  const clusters: string[] = []
  let current: string[] = []
  for (const token of data.split(/\s+/)) {
    if (token === '÷') { if (current.length) clusters.push(current.join('.')); current = [] }
    else if (token !== '×') current.push(token.toLowerCase())
  }
  if (current.length) throw new Error(`Unterminated GraphemeBreakTest case: ${line}`)
  cases.push(clusters.join('|'))
}
const fixturePath = join(root, `tests/fixtures/unicode/grapheme-break-test-${VERSION}.json`)
mkdirSync(dirname(fixturePath), { recursive: true })
writeFileSync(fixturePath, JSON.stringify({
  source: `${BASE}${INPUTS['GraphemeBreakTest.txt'].path}`,
  sha256: INPUTS['GraphemeBreakTest.txt'].sha256,
  format: 'clusters separated by |, hexadecimal code points separated by .',
  cases,
}, null, 0) + '\n')

// Round-trip and conformance checks against the module just written.
const canonical = await import(pathToFileURL(join(root, 'src/xlsx/view/canonicalText.ts')).href + `?v=${Date.now()}`)
for (let cp = 0; cp <= MAX; cp++) {
  const expected = cp >= hangulBase && cp <= hangulLast ? ((cp - hangulBase) % 28 === 0 ? 12 : 13) : props[cp] & 15
  if (canonical.unicodeProperty(cp) !== (props[cp] & ~15 | expected)) throw new Error(`Property round-trip mismatch at ${cp.toString(16)}`)
  const target = fold.get(cp)
  const actual: string | undefined = canonical.fullCaseFold(cp, false)
  if ((target ? String.fromCodePoint(...target) : undefined) !== actual) throw new Error(`Case folding round-trip mismatch at ${cp.toString(16)}`)
}
let failures = 0
for (const entry of cases) {
  const clusters = entry.split('|').map(cluster => String.fromCodePoint(...cluster.split('.').map(value => parseInt(value, 16))))
  const actual = canonical.graphemeClusters(clusters.join(''))
  if (JSON.stringify(actual) !== JSON.stringify(clusters)) failures++
}
if (failures) throw new Error(`${failures} of ${cases.length} GraphemeBreakTest cases failed`)
console.log(`Unicode ${VERSION}: ${runs} property runs (${palette.length} values), ${foldRuns.length} fold runs, ${multiples.length} expansions; ${cases.length}/${cases.length} GraphemeBreakTest cases pass`)
