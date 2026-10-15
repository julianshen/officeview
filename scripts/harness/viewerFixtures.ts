/**
 * S10.2 viewer fixtures: true XLSX packages written by the S0.2 deterministic
 * builders (scripts/generate-viewer-fixtures.ts) at an explicit scale, plus two
 * scroll-mix workbooks (wrapped text; styled fills/fonts/borders/formats) built
 * here. Packages are cached by name under the repository's .tmp/; every run
 * re-hashes the bytes it measures, so a receipt names the exact SHA-256, byte
 * count, expanded XML bytes and populated cells of each workbook.
 */
import { createHash } from 'node:crypto'
import { createInflateRaw } from 'node:zlib'
import { createReadStream, existsSync, mkdirSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeFixturePackage } from '../generate-viewer-fixtures'
import { colName, xlsxBytes } from '../../tests/helpers/xlsx-package'

export type ViewerScale = 'full' | 'reduced'
export type ViewerFixtureKey = 'F1' | 'F2' | 'F3' | 'F4' | 'wrapped' | 'wrappedIntl' | 'styled'
export interface ViewerFixture {
  key: ViewerFixtureKey
  path: string
  bytes: number
  sha256: string
  /** Verified against the committed manifest. */
  contentSha256: string
  uncompressedBytes: number
  nnz: number
  rows: number
  cols: number
  /** What was generated, at which scale (never implied full when reduced). */
  scale: string
}

/** Bump when a builder's output changes, so stale cached packages are not reused. */
export const FIXTURE_CACHE_VERSION = 1
const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..')
export function fixtureDir(): string {
  // Tests point the cache elsewhere. The default is the repository's ignored
  // .tmp/, not the OS temp directory: macOS dirhelper deletes $TMPDIR files
  // older than 3 days at 03:35 daily, which emptied the cache between the
  // suite's start-up verification and its import phase (ENOENT on f1-full.xlsx).
  return process.env.OFFICEVIEW_VIEWER_FIXTURE_DIR ?? join(REPO_ROOT, '.tmp', `viewer-perf-fixtures-v${FIXTURE_CACHE_VERSION}`)
}

const SCROLL_ROWS = { full: 20_000, reduced: 2_000 } as const
const SCROLL_COLS = 12

/** Scale definitions: full is the spec §12.5 catalog size; reduced is a labelled smoke size. */
export const SCALES: Record<ViewerScale, Record<'F1' | 'F2' | 'F3' | 'F4', { params: Record<string, number>; label: string }>> = {
  full: {
    F1: { params: { nnz: 100_000 }, label: 'F1 full: 100,000 populated cells across full axes (corners pinned)' },
    F2: { params: {}, label: 'F2 full: 1,048,576 rows x 8 fields' },
    F3: { params: {}, label: 'F3 full: dense 256 rows x 16,384 columns (4,194,304 cells)' },
    F4: { params: {}, label: 'F4 full: 1,000,000 unique strings' },
  },
  reduced: {
    F1: { params: { nnz: 20_000 }, label: 'F1 reduced: 20,000 populated cells (smoke scale)' },
    F2: { params: { rowCount: 100_010 }, label: 'F2 reduced: 100,010 rows x 8 (smoke scale)' },
    F3: { params: { rowCount: 8 }, label: 'F3 reduced: dense 8 rows x 16,384 columns (smoke scale)' },
    F4: { params: { rowCount: 12_000 }, label: 'F4 reduced: 12,000 unique strings (smoke scale)' },
  },
}

const WRAPPED_STYLES = `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs>`
const STYLED_STYLES = `<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00;[Red]\\-&quot;$&quot;#,##0.00"/></numFmts>
<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><color rgb="FF1F3864"/><name val="Calibri"/></font><font><i/><sz val="10"/><color rgb="FF7F6000"/><name val="Georgia"/></font><font><u/><sz val="11"/><color rgb="FFC00000"/><name val="Arial"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF8EA9DB"/></left><right style="thin"><color rgb="FF8EA9DB"/></right><top style="thin"><color rgb="FF8EA9DB"/></top><bottom style="medium"><color rgb="FF2F5597"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="10" fontId="2" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
<xf numFmtId="14" fontId="3" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs>`

const WORDS = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'theta', 'kappa', 'lambda', 'sigma', 'omega', 'quartz', 'harbor', 'meadow', 'lantern']
/** Non-ASCII words (accents, CJK, emoji with ZWJ) so text layout segments graphemes. */
const INTL_WORDS = ['café', 'naïve', 'Ångström', 'größe', '東京', '数据表', 'سلام', 'Δέλτα', 'жёлтый', '👩‍💻', 'emoji🙂', 'São', 'Zürich', 'piñata', 'façade']
const esc = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;')

/** Wrapped text: every cell holds ~110 characters in a 24-character column, wrapping to 4-5 lines in 64px rows. */
function wrappedPackage(rows: number, words: readonly string[] = WORDS): { bytes: Uint8Array; nnz: number } {
  const out: string[] = []
  for (let r = 1; r <= rows; r++) {
    const cells: string[] = []
    for (let c = 0; c < SCROLL_COLS; c++) {
      const picked: string[] = []
      for (let w = 0; w < 16; w++) picked.push(words[(r * 7 + c * 3 + w * 5) % words.length])
      cells.push(`<c r="${colName(c)}${r}" t="inlineStr" s="1"><is><t>${esc(`R${r}C${c + 1} ${picked.join(' ')}`)}</t></is></c>`)
    }
    out.push(`<row r="${r}" ht="64" customHeight="1">${cells.join('')}</row>`)
  }
  return { bytes: xlsxBytes({ beforeSheetData: `<cols><col min="1" max="${SCROLL_COLS}" width="24" customWidth="1"/></cols>`, sheetData: out.join(''), styles: WRAPPED_STYLES }), nnz: rows * SCROLL_COLS }
}

/** Styled: header row plus fills, fonts, borders, currency/percent/date formats on every body cell. */
function styledPackage(rows: number): { bytes: Uint8Array; nnz: number } {
  const out: string[] = []
  const header: string[] = []
  for (let c = 0; c < SCROLL_COLS; c++) header.push(`<c r="${colName(c)}1" t="inlineStr" s="1"><is><t>Field ${c + 1}</t></is></c>`)
  out.push(`<row r="1">${header.join('')}</row>`)
  for (let r = 2; r <= rows; r++) {
    const cells: string[] = []
    for (let c = 0; c < SCROLL_COLS; c++) {
      const ref = `${colName(c)}${r}`
      switch (c % 4) {
        case 0: cells.push(`<c r="${ref}" s="2"><v>${((r * 7919 + c * 104729) % 2_000_000) / 100 - 5000}</v></c>`); break
        case 1: cells.push(`<c r="${ref}" s="3"><v>${((r * 31 + c) % 1000) / 1000}</v></c>`); break
        case 2: cells.push(`<c r="${ref}" s="4"><v>${40000 + ((r + c) % 3650)}</v></c>`); break
        default: cells.push(`<c r="${ref}" t="inlineStr" s="5"><is><t>${WORDS[(r + c) % WORDS.length]} ${r}</t></is></c>`)
      }
    }
    out.push(`<row r="${r}">${cells.join('')}</row>`)
  }
  return { bytes: xlsxBytes({ beforeSheetData: `<cols><col min="1" max="${SCROLL_COLS}" width="16" customWidth="1"/></cols>`, sheetData: out.join(''), styles: STYLED_STYLES }), nnz: rows * SCROLL_COLS }
}

async function fileSha256(path: string): Promise<string> {
  const hash = createHash('sha256')
  for await (const chunk of createReadStream(path)) hash.update(chunk as Buffer)
  return hash.digest('hex')
}

/** Committed identity of every generated package, per generator version. */
export const FIXTURE_MANIFEST_PATH = 'tests/fixtures/xlsx-viewer/viewer-perf-fixtures.json'
export interface FixtureManifestEntry { file: string; contentSha256: string; uncompressedBytes: number; /** Bun build output; informational (deflate bytes vary by zlib build). */ bytes: number; sha256: string; nnz: number; rows: number; cols: number; seed: number | null; params: Record<string, number>; label: string }
export interface FixtureManifest { schema: 'officeview-viewer-perf-fixtures/1'; cacheVersion: number; scales: Record<ViewerScale, Record<ViewerFixtureKey, FixtureManifestEntry>> }
export const FIXTURE_KEYS: readonly ViewerFixtureKey[] = ['F1', 'F2', 'F3', 'F4', 'wrapped', 'wrappedIntl', 'styled']

/**
 * SHA-256 over the package CONTENT: every entry's name and inflated bytes in
 * central-directory order. Deflate output differs between zlib builds (Bun vs
 * Node), so the compressed file hash cannot be the cross-runtime identity; the
 * content digest is. Streams each entry; never holds a sheet in memory.
 */
export async function zipContentSha256(path: string): Promise<{ contentSha256: string; uncompressedBytes: number }> {
  const { open } = await import('node:fs/promises')
  const handle = await open(path, 'r')
  try {
    const size = (await handle.stat()).size
    const tailLen = Math.min(size, 65_557)
    const tail = Buffer.alloc(tailLen)
    await handle.read(tail, 0, tailLen, size - tailLen)
    let eocd = -1
    for (let i = tailLen - 22; i >= 0; i--) if (tail.readUInt32LE(i) === 0x06054b50) { eocd = i; break }
    if (eocd < 0) throw new Error(`${path}: not a ZIP package`)
    const count = tail.readUInt16LE(eocd + 10), cdSize = tail.readUInt32LE(eocd + 12), cdOffset = tail.readUInt32LE(eocd + 16)
    const cd = Buffer.alloc(cdSize)
    await handle.read(cd, 0, cdSize, cdOffset)
    const hash = createHash('sha256')
    let offset = 0, total = 0
    for (let i = 0; i < count; i++) {
      if (cd.readUInt32LE(offset) !== 0x02014b50) throw new Error(`${path}: bad central directory`)
      const method = cd.readUInt16LE(offset + 10), compressed = cd.readUInt32LE(offset + 20), uncompressed = cd.readUInt32LE(offset + 24)
      const nameLen = cd.readUInt16LE(offset + 28), extraLen = cd.readUInt16LE(offset + 30), commentLen = cd.readUInt16LE(offset + 32), local = cd.readUInt32LE(offset + 42)
      const name = cd.subarray(offset + 46, offset + 46 + nameLen)
      offset += 46 + nameLen + extraLen + commentLen
      const header = Buffer.alloc(30)
      await handle.read(header, 0, 30, local)
      const dataStart = local + 30 + header.readUInt16LE(26) + header.readUInt16LE(28)
      hash.update(name); hash.update('\0')
      const raw = createReadStream(path, { start: dataStart, end: dataStart + compressed - 1 })
      const stream = method === 8 ? raw.pipe(createInflateRaw()) : raw
      let seen = 0
      for await (const chunk of stream) { hash.update(chunk as Buffer); seen += (chunk as Buffer).length }
      if (seen !== uncompressed) throw new Error(`${path}: ${name.toString()} inflated to ${seen} B, directory says ${uncompressed}`)
      total += uncompressed
      hash.update('\0')
    }
    return { contentSha256: hash.digest('hex'), uncompressedBytes: total }
  } finally { await handle.close() }
}
async function identityOf(path: string): Promise<{ bytes: number; sha256: string; contentSha256: string; uncompressedBytes: number }> {
  const [sha256, content] = await Promise.all([fileSha256(path), zipContentSha256(path)])
  return { bytes: statSync(path).size, sha256, ...content }
}
/** Writes one package from its generator definition; returns what the definition says it holds. */
async function generate(path: string, key: ViewerFixtureKey, scale: ViewerScale): Promise<Omit<FixtureManifestEntry, 'file' | 'bytes' | 'sha256' | 'contentSha256' | 'uncompressedBytes'>> {
  if (key === 'wrapped' || key === 'wrappedIntl' || key === 'styled') {
    const rows = SCROLL_ROWS[scale]
    writeFileSync(path, (key === 'styled' ? styledPackage(rows) : wrappedPackage(rows, key === 'wrappedIntl' ? INTL_WORDS : WORDS)).bytes)
    const label = `${key} scroll mix: ${rows.toLocaleString('en-US')} rows x ${SCROLL_COLS} columns${key === 'wrappedIntl' ? ', non-ASCII words' : key === 'wrapped' ? ', ASCII words' : ''}${scale === 'reduced' ? ' (smoke scale)' : ''}`
    return { nnz: rows * SCROLL_COLS, rows, cols: SCROLL_COLS, seed: null, params: { rows }, label }
  }
  const def = SCALES[scale][key]
  const written = await writeFixturePackage(path, key, def.params)
  const { createFixture } = await import('../../tests/helpers/xlsx-viewer-fixtures')
  const handle = createFixture(key, def.params)
  return { nnz: written.nnz, rows: handle.rowCount, cols: handle.colCount, seed: handle.seed, params: def.params, label: def.label }
}
const same = (a: { contentSha256: string; uncompressedBytes: number }, b: { contentSha256: string; uncompressedBytes: number }) =>
  a.contentSha256 === b.contentSha256 && a.uncompressedBytes === b.uncompressedBytes

/**
 * Returns every package for `scale`, each verified against the committed
 * manifest (SHA-256, bytes, expanded bytes). A cached file that does not match
 * is regenerated; a regenerated file that still does not match means the
 * generator changed, and the run is refused. Reported shape and populated
 * counts come from the manifest entry the bytes were verified against, never
 * from the current generator definition alone.
 */
export async function ensureViewerFixtures(scale: ViewerScale, manifest: FixtureManifest, keys: readonly ViewerFixtureKey[] = FIXTURE_KEYS): Promise<Record<ViewerFixtureKey, ViewerFixture>> {
  if (manifest.cacheVersion !== FIXTURE_CACHE_VERSION) throw new Error(`fixture manifest is for cache version ${manifest.cacheVersion}, generator is ${FIXTURE_CACHE_VERSION}: regenerate the manifest`)
  const dir = fixtureDir()
  mkdirSync(dir, { recursive: true })
  const out = {} as Record<ViewerFixtureKey, ViewerFixture>
  for (const key of keys) {
    const expected = manifest.scales[scale]?.[key]
    if (!expected) throw new Error(`fixture manifest has no ${scale} entry for ${key}`)
    const path = join(dir, expected.file)
    let actual = existsSync(path) ? await identityOf(path) : null
    if (!actual || !same(actual, expected)) {
      const temp = `${path}.${process.pid}.tmp`
      await generate(temp, key, scale)
      actual = await identityOf(temp)
      if (!same(actual, expected)) {
        rmSync(temp, { force: true })
        throw new Error(`${key} (${scale}): generated package content ${actual.contentSha256} does not match manifest ${expected.contentSha256}; the generator changed: bump FIXTURE_CACHE_VERSION and rewrite ${FIXTURE_MANIFEST_PATH}`)
      }
      renameSync(temp, path)
    }
    // The file's own bytes and hash are what this run opens; content identity
    // and counts are the manifest entry those bytes were verified against.
    out[key] = { key, path, bytes: actual.bytes, sha256: actual.sha256, contentSha256: expected.contentSha256, uncompressedBytes: expected.uncompressedBytes, nnz: expected.nnz, rows: expected.rows, cols: expected.cols, scale: expected.label }
  }
  return out
}

/** Regenerates every package into a fresh directory and returns the manifest that describes them. */
export async function buildFixtureManifest(dir: string): Promise<FixtureManifest> {
  mkdirSync(dir, { recursive: true })
  const scales = {} as FixtureManifest['scales']
  for (const scale of ['full', 'reduced'] as const) {
    scales[scale] = {} as Record<ViewerFixtureKey, FixtureManifestEntry>
    for (const key of FIXTURE_KEYS) {
      const file = `${key.toLowerCase()}-${scale}.xlsx`
      const facts = await generate(join(dir, file), key, scale)
      scales[scale][key] = { file, ...(await identityOf(join(dir, file))), ...facts }
    }
  }
  return { schema: 'officeview-viewer-perf-fixtures/1', cacheVersion: FIXTURE_CACHE_VERSION, scales }
}
