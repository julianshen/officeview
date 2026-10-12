/**
 * S10.2 viewer fixtures: true XLSX packages written by the S0.2 deterministic
 * builders (scripts/generate-viewer-fixtures.ts) at an explicit scale, plus two
 * scroll-mix workbooks (wrapped text; styled fills/fonts/borders/formats) built
 * here. Packages are cached by name under the OS temp directory; every run
 * re-hashes the bytes it measures, so a receipt names the exact SHA-256, byte
 * count, expanded XML bytes and populated cells of each workbook.
 */
import { createHash } from 'node:crypto'
import { createReadStream, existsSync, mkdirSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { writeFixturePackage } from '../generate-viewer-fixtures'
import { colName, xlsxBytes } from '../../tests/helpers/xlsx-package'

export type ViewerScale = 'full' | 'reduced'
export type ViewerFixtureKey = 'F1' | 'F2' | 'F3' | 'F4' | 'wrapped' | 'styled'
export interface ViewerFixture {
  key: ViewerFixtureKey
  path: string
  bytes: number
  sha256: string
  uncompressedBytes: number
  nnz: number
  /** What was generated, at which scale (never implied full when reduced). */
  scale: string
}

/** Bump when a builder's output changes, so stale cached packages are not reused. */
export const FIXTURE_CACHE_VERSION = 1
export function fixtureDir(): string {
  return join(tmpdir(), `officeview-viewer-perf-v${FIXTURE_CACHE_VERSION}`)
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
const esc = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;')

/** Wrapped text: every cell holds ~110 characters in a 24-character column, wrapping to 4-5 lines in 64px rows. */
function wrappedPackage(rows: number): { bytes: Uint8Array; nnz: number } {
  const out: string[] = []
  for (let r = 1; r <= rows; r++) {
    const cells: string[] = []
    for (let c = 0; c < SCROLL_COLS; c++) {
      const words: string[] = []
      for (let w = 0; w < 16; w++) words.push(WORDS[(r * 7 + c * 3 + w * 5) % WORDS.length])
      cells.push(`<c r="${colName(c)}${r}" t="inlineStr" s="1"><is><t>${esc(`R${r}C${c + 1} ${words.join(' ')}`)}</t></is></c>`)
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

/** Uncompressed bytes from the ZIP central directory (sum of entry sizes). */
async function zipUncompressedBytes(path: string): Promise<number> {
  const { readFile } = await import('node:fs/promises')
  const buf = await readFile(path)
  let eocd = -1
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65_557); i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break }
  if (eocd < 0) return 0
  const count = buf.readUInt16LE(eocd + 10)
  let offset = buf.readUInt32LE(eocd + 16), total = 0
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(offset) !== 0x02014b50) break
    total += buf.readUInt32LE(offset + 24)
    offset += 46 + buf.readUInt16LE(offset + 28) + buf.readUInt16LE(offset + 30) + buf.readUInt16LE(offset + 32)
  }
  return total
}

/** Writes (or reuses) every package for `scale` and returns their measured identity. */
export async function ensureViewerFixtures(scale: ViewerScale, keys: readonly ViewerFixtureKey[] = ['F1', 'F2', 'F3', 'F4', 'wrapped', 'styled']): Promise<Record<ViewerFixtureKey, ViewerFixture>> {
  const dir = fixtureDir()
  mkdirSync(dir, { recursive: true })
  const out = {} as Record<ViewerFixtureKey, ViewerFixture>
  for (const key of keys) {
    const path = join(dir, `${key.toLowerCase()}-${scale}.xlsx`)
    let nnz: number
    let label: string
    if (key === 'wrapped' || key === 'styled') {
      const rows = SCROLL_ROWS[scale]
      label = `${key} scroll mix: ${rows.toLocaleString('en-US')} rows x ${SCROLL_COLS} columns${scale === 'reduced' ? ' (smoke scale)' : ''}`
      nnz = rows * SCROLL_COLS
      if (!existsSync(path)) writeFileSync(path, (key === 'wrapped' ? wrappedPackage : styledPackage)(rows).bytes)
    } else {
      const def = SCALES[scale][key]
      label = def.label
      if (!existsSync(path)) await writeFixturePackage(path, key, def.params)
      nnz = -1
    }
    const bytes = statSync(path).size
    const [sha256, uncompressedBytes] = await Promise.all([fileSha256(path), zipUncompressedBytes(path)])
    out[key] = { key, path, bytes, sha256, uncompressedBytes, nnz, scale: label }
  }
  // Builder nnz for the S0.2 fixtures (deterministic; computed only when needed).
  for (const key of keys) {
    if (out[key].nnz >= 0) continue
    const { createFixture } = await import('../../tests/helpers/xlsx-viewer-fixtures')
    out[key].nnz = createFixture(key as 'F1' | 'F2' | 'F3' | 'F4', SCALES[scale][key as 'F1'].params).nnz()
  }
  return out
}
