/**
 * Corpus of real-world office files used to find bugs that synthetic fixtures
 * never will. Files are pinned to exact upstream commits so URLs cannot drift,
 * and are fetched on demand rather than committed (avoids vendoring third-party
 * binaries in our repo).
 *
 *   bun scripts/fetch-corpus.ts          # download + verify
 *   bun scripts/fetch-corpus.ts --check  # verify only, fetch nothing
 *
 * Sources (see LICENSE notes per entry):
 *   apache/poi               Apache-2.0 — 130 docx / 366 xlsx / 95 pptx fixtures,
 *                             many of them real Excel/Word bug repros plus
 *                             minimized fuzzer cases (robustness probes)
 *   python-openxml/python-docx  MIT
 *   scanny/python-pptx           MIT  — includes deliberately malformed files
 */

import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

const POI = '93b6a820f4aaefd7b062615f843799947872cb24' // apache/poi
const PYDOCX = 'e45454602b53e8e572b179ccf1c91093ec9f4ed7' // python-openxml/python-docx
const PYPPTX = '278b47b1dedd5b46ee84c286e77cdfb0bf4594be' // scanny/python-pptx

export interface CorpusEntry {
  /** local filename (prefixed with source for traceability) */
  name: string
  url: string
  source: string
  license: string
  /** why this file is in the corpus — the bug class it probes */
  probe: string
}

const poi = (path: string, probe: string): CorpusEntry => ({
  // local name is flat (basename); the URL keeps the upstream subdirectory
  name: `poi-${path.split('/').pop()}`,
  url: `https://raw.githubusercontent.com/apache/poi/${POI}/test-data/${path}`,
  source: 'apache/poi',
  license: 'Apache-2.0',
  probe,
})
const pydocx = (name: string, probe: string): CorpusEntry => ({
  name: `pydocx-${name}`,
  url: `https://raw.githubusercontent.com/python-openxml/python-docx/${PYDOCX}/tests/test_files/${name}`,
  source: 'python-openxml/python-docx',
  license: 'MIT',
  probe,
})
const pypptx = (name: string, probe: string): CorpusEntry => ({
  name: `pypptx-${name}`,
  url: `https://raw.githubusercontent.com/scanny/python-pptx/${PYPPTX}/tests/test_files/${name}`,
  source: 'scanny/python-pptx',
  license: 'MIT',
  probe,
})

export const CORPUS: CorpusEntry[] = [
  // ---- DOCX ----
  poi('document/Bug66263-table.docx', 'tables in a real Word-produced file'),
  poi('document/table-indent.docx', 'table indentation'),
  poi('document/55733.docx', 'Word bug repro'),
  poi('document/comment.docx', 'comments/annotations present'),
  poi('document/unicode-path.docx', 'non-ascii part names'),
  poi('document/rtl.docx', 'right-to-left text'),
  poi('document/crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx', 'minimized fuzzer case (robustness)'),
  pydocx('test.docx', 'general python-docx output'),
  pydocx('having-images.docx', 'embedded images at scale'),
  pydocx('blk-inner-content.docx', 'nested block content'),
  pydocx('sct-inner-content.docx', 'structured document tag content'),

  // ---- XLSX ----
  poi('spreadsheet/59021.xlsx', 'Excel bug repro'),
  poi('spreadsheet/47889.xlsx', 'Excel bug repro'),
  poi('spreadsheet/64508.xlsx', 'Excel bug repro'),
  poi('spreadsheet/66365.xlsx', 'Excel bug repro'),
  poi('spreadsheet/56295.xlsx', 'Excel bug repro'),
  poi('spreadsheet/62272.xlsx', 'Excel bug repro'),
  poi('spreadsheet/link-external-workbook-a.xlsx', 'external workbook links'),
  poi('spreadsheet/59746_NoRowNums.xlsx', 'worksheet without row numbers'),
  poi('spreadsheet/xlsx-jdbc.xlsx', 'JDBC-produced workbook'),

  // ---- PPTX ----
  poi('slideshow/table_test.pptx', 'tables in a real deck'),
  poi('slideshow/picture-transparency.pptx', 'picture transparency'),
  poi('slideshow/crop-to-0.pptx', 'degenerate crop rectangle'),
  poi('slideshow/text-highlight.pptx', 'text highlight runs'),
  poi('slideshow/copy-slide-demo.pptx', 'duplicated slide ids/parts'),
  pypptx('minimal.pptx', 'smallest valid deck'),
  pypptx('test.pptx', 'general python-pptx output'),
  pypptx('test_slides.pptx', 'multi-slide deck'),
  pypptx('missing_rels_item.pptx', 'dangling relationship (robustness)'),
  pypptx('no-core-props.pptx', 'missing core properties (robustness)'),
]

const CORPUS_DIR = join(import.meta.dir, '..', 'corpus')
const LOCK = join(CORPUS_DIR, 'corpus.lock.json')

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

async function fetchBytes(url: string): Promise<Uint8Array> {
  const res = await fetch(url)
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`)
  return new Uint8Array(await res.arrayBuffer())
}

const checkOnly = process.argv.includes('--check')
mkdirSync(CORPUS_DIR, { recursive: true })

let lock: Record<string, string> = {}
if (existsSync(LOCK)) {
  try {
    lock = JSON.parse(readFileSync(LOCK, 'utf8')) as Record<string, string>
  } catch {
    lock = {}
  }
}

let downloaded = 0
let verified = 0
const failures: string[] = []

for (const entry of CORPUS) {
  const path = join(CORPUS_DIR, entry.name)
  const expected = lock[entry.name]

  if (checkOnly) {
    if (!existsSync(path)) {
      failures.push(`${entry.name}: not fetched`)
      continue
    }
    const actual = sha256(readFileSync(path))
    if (expected && expected !== actual) failures.push(`${entry.name}: sha mismatch`)
    else verified++
    continue
  }

  try {
    if (existsSync(path) && expected && sha256(readFileSync(path)) === expected) {
      verified++
      continue
    }
    const bytes = await fetchBytes(entry.url)
    if (bytes.length < 100) throw new Error(`suspiciously small (${bytes.length} bytes)`)
    writeFileSync(path, bytes)
    const digest = sha256(bytes)
    if (expected && expected !== digest) {
      // upstream changed: surface it loudly rather than silently re-pinning
      throw new Error(`sha mismatch — pinned ${expected.slice(0, 12)}…, got ${digest.slice(0, 12)}…`)
    }
    lock[entry.name] = digest
    downloaded++
    console.log(`  fetched ${entry.name} (${bytes.length} bytes) — ${entry.probe}`)
  } catch (e) {
    failures.push(`${entry.name}: ${e instanceof Error ? e.message : String(e)}`)
  }
}

writeFileSync(LOCK, `${JSON.stringify(lock, null, 2)}\n`)

console.log(`\ncorpus: ${verified} verified, ${downloaded} newly downloaded, ${failures.length} failed`)
for (const f of failures) console.log(`  FAILED ${f}`)
console.log(`\nSources: apache/poi (Apache-2.0), python-openxml/python-docx (MIT), scanny/python-pptx (MIT)`)
console.log(`Files live in corpus/ (gitignored); manifest is embedded in this script, hashes in corpus/corpus.lock.json`)
process.exit(failures.length > 0 ? 1 : 0)
