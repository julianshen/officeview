/** Run: bun scripts/generate-viewer-fixtures.ts --out <dir> [--fixtures F1,F2,...] [--scale small|full] [--string-table inline|shared]
 * Type-check path: bunx tsc --noEmit (scripts/generate-viewer-fixtures.ts is in tsconfig include).
 * Writes TRUE XLSX packages with a streaming ZIP writer: sheet XML is generated
 * row by row and flushed in bounded chunks, so full-size F2/F3/F4 never hold a
 * whole-sheet string or object model in memory. Small committed packages use
 * stored entries; full-size parts use streaming raw deflate (node:zlib, no new
 * dependencies). F4 full runs default to a streamed shared-string table so the
 * S1.1 sharedStrings path is exercised at scale. Manifest (manifest.json) holds
 * only deterministic regenerable fields; this run's measurements go to
 * measured.json (not committed). */
import { createHash } from 'node:crypto'
import { createReadStream, mkdirSync, statSync, writeFileSync } from 'node:fs'
import { open } from 'node:fs/promises'
import { once } from 'node:events'
import { join, resolve } from 'node:path'
import { createDeflateRaw } from 'node:zlib'
import { columnToLetters } from '../src/xlsx/view/coordinates'
import {
  createFixture, encodeXlsxString, f4StringFor, f8Cases, headQueryHashes, sha256Hex,
  type CellRef, type FixtureCell, type FixtureHandle, type FixtureId, type FixtureParams,
} from '../tests/helpers/xlsx-viewer-fixtures'

const encoder = new TextEncoder()

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

function crcUpdate(crc: number, bytes: Uint8Array): number {
  let c = crc ^ -1
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8)
  return (c ^ -1) >>> 0
}

function u16(value: number): Uint8Array {
  return new Uint8Array([value & 0xff, (value >>> 8) & 0xff])
}

function u32(value: number): Uint8Array {
  return new Uint8Array([value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff])
}

/** Fixed DOS timestamp keeps fixture bytes deterministic across runs. */
const DOS_TIME = 0
const DOS_DATE = ((2026 - 1980) << 9) | (10 << 5) | 10

interface ZipEntry {
  name: string
  method: number
  crc: number
  size: number
  compressed: number
  offset: number
}

/** Minimal streaming ZIP writer (stored or raw-deflate entries, data descriptors). */
class StreamingZip {
  private fd: Awaited<ReturnType<typeof open>> | null = null
  private offset = 0
  private entries: ZipEntry[] = []
  private current: { name: string; method: number; crc: number; size: number; compressed: number; offset: number } | null = null

  constructor(private readonly path: string) {}

  async openFile(): Promise<void> {
    this.fd = await open(this.path, 'w')
  }

  private async write(bytes: Uint8Array): Promise<void> {
    await this.fd!.write(bytes)
    this.offset += bytes.length
  }

  async beginFile(name: string, method: 0 | 8): Promise<void> {
    const nameBytes = encoder.encode(name)
    const header = new Uint8Array(30 + nameBytes.length)
    header.set([0x50, 0x4b, 0x03, 0x04], 0)
    header.set(u16(20), 4)
    header.set(u16(0x08), 6)
    header.set(u16(method), 8)
    header.set(u16(DOS_TIME), 10)
    header.set(u16(DOS_DATE), 12)
    header.set(u32(0), 14)
    header.set(u32(0), 18)
    header.set(u32(0), 22)
    header.set(u16(nameBytes.length), 26)
    header.set(u16(0), 28)
    header.set(nameBytes, 30)
    this.current = { name, method, crc: 0, size: 0, compressed: 0, offset: this.offset }
    await this.write(header)
  }

  async appendChunk(bytes: Uint8Array): Promise<void> {
    if (this.current === null) throw new Error('no open zip entry')
    this.current.crc = crcUpdate(this.current.crc, bytes)
    this.current.size += bytes.length
    this.current.compressed += bytes.length
    await this.write(bytes)
  }

  async appendText(text: string): Promise<void> {
    await this.appendChunk(encoder.encode(text))
  }

  /** Stream uncompressed chunks through raw deflate into the open entry. */
  async appendDeflated(chunks: Iterable<Uint8Array>): Promise<void> {
    if (this.current === null) throw new Error('no open zip entry')
    if (this.current.method !== 8) throw new Error('entry was not opened for deflate')
    const entry = this.current
    const deflater = createDeflateRaw({ level: 6 })
    let tail: Promise<void> = Promise.resolve()
    deflater.on('data', (chunk: Uint8Array) => {
      entry.compressed += chunk.length
      tail = tail.then(() => this.write(chunk))
    })
    for (const chunk of chunks) {
      entry.crc = crcUpdate(entry.crc, chunk)
      entry.size += chunk.length
      if (!deflater.write(chunk)) await once(deflater, 'drain')
    }
    deflater.end()
    await once(deflater, 'end')
    await tail
  }

  async endFile(): Promise<void> {
    if (this.current === null) throw new Error('no open zip entry')
    const descriptor = new Uint8Array(16)
    descriptor.set([0x50, 0x4b, 0x07, 0x08], 0)
    descriptor.set(u32(this.current.crc), 4)
    descriptor.set(u32(this.current.compressed), 8)
    descriptor.set(u32(this.current.size), 12)
    await this.write(descriptor)
    this.entries.push({ ...this.current })
    this.current = null
  }

  /** Total uncompressed payload bytes (manifest accounting). */
  payloadBytes(): number {
    return this.entries.reduce((sum, entry) => sum + entry.size, 0)
  }

  async finish(): Promise<void> {
    const centralStart = this.offset
    for (const entry of this.entries) {
      const nameBytes = encoder.encode(entry.name)
      const header = new Uint8Array(46 + nameBytes.length)
      header.set([0x50, 0x4b, 0x01, 0x02], 0)
      header.set(u16(20), 4)
      header.set(u16(20), 6)
      header.set(u16(0x08), 8)
      header.set(u16(entry.method), 10)
      header.set(u16(DOS_TIME), 12)
      header.set(u16(DOS_DATE), 14)
      header.set(u32(entry.crc), 16)
      header.set(u32(entry.compressed), 20)
      header.set(u32(entry.size), 24)
      header.set(u16(nameBytes.length), 28)
      header.set(u16(0), 30)
      header.set(u16(0), 32)
      header.set(u16(0), 34)
      header.set(u16(0), 36)
      header.set(u32(0), 38)
      header.set(u32(entry.offset), 42)
      header.set(nameBytes, 46)
      await this.write(header)
    }
    const centralSize = this.offset - centralStart
    const end = new Uint8Array(22)
    end.set([0x50, 0x4b, 0x05, 0x06], 0)
    end.set(u16(0), 4)
    end.set(u16(0), 6)
    end.set(u16(this.entries.length), 8)
    end.set(u16(this.entries.length), 10)
    end.set(u32(centralSize), 12)
    end.set(u32(centralStart), 16)
    end.set(u16(0), 20)
    await this.write(end)
    await this.fd!.close()
    this.fd = null
  }
}

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function a1Ref(row: number, col: number): string {
  return `${columnToLetters(col)}${row + 1}`
}

function cellXml(ref: string, cell: FixtureCell, shared: { indexOf(text: string): number | null } | null): string {
  const style = cell.styleIndex !== 0 ? ` s="${cell.styleIndex}"` : ''
  const formulaXml = cell.formula !== undefined ? `<f>${escapeXml(cell.formula)}</f>` : ''
  if (cell.formula !== undefined && (cell.value === null || !cell.hasCachedValue)) {
    return `<c r="${ref}"${style}>${formulaXml}</c>`
  }
  if (typeof cell.value === 'number') {
    return `<c r="${ref}"${style}>${formulaXml}<v>${cell.value}</v></c>`
  }
  if (typeof cell.value === 'boolean') {
    return `<c r="${ref}" t="b"${style}>${formulaXml}<v>${cell.value ? 1 : 0}</v></c>`
  }
  if (cell.valueIsError) {
    return `<c r="${ref}" t="e"${style}>${formulaXml}<v>${escapeXml(cell.value ?? '')}</v></c>`
  }
  const text = cell.value ?? ''
  if (shared !== null) {
    const index = shared.indexOf(text)
    if (index !== null) return `<c r="${ref}" t="s"${style}>${formulaXml}<v>${index}</v></c>`
  }
  if (cell.formula !== undefined) {
    return `<c r="${ref}" t="str"${style}>${formulaXml}<v>${encodeXlsxString(text)}</v></c>`
  }
  return `<c r="${ref}" t="inlineStr"${style}><is><t>${encodeXlsxString(text)}</t></is></c>`
}

const CONTENT_TYPES = (sheets: number, shared: boolean): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${Array.from({ length: sheets }, (_, i) => `  <Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n')}
${shared ? '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n' : ''}  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK_RELS = (sheets: number, shared: boolean): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${Array.from({ length: sheets }, (_, i) => `  <Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('\n')}
${shared ? `  <Relationship Id="rId${sheets + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n` : ''}  <Relationship Id="rId${sheets + (shared ? 2 : 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

const WORKBOOK_XML = (names: string[]): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
${names.map((name, i) => `    <sheet name="${escapeXml(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('\n')}
  </sheets>
</workbook>`

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`

interface SheetRowSpec {
  r: number
  ht?: number
  hidden?: boolean
  cells: CellRef[]
}

interface SharedTable {
  count: number
  strings: Iterable<string>
}

interface SheetSpec {
  name: string
  /** Row-major row iterator; must stream (no full row array) for large sheets. */
  rows: Iterable<SheetRowSpec>
  cols?: string
  merges?: string[]
  sharedTable?: SharedTable
  /** String cells reference their row as the shared-string index (F4 positional table). */
  sharedPositional?: boolean
}

/** Bounded sheet-XML chunk generator: callers choose stored or deflated transport. */
function* sheetXmlChunks(sheet: SheetSpec): Generator<string> {
  const sharedIndex = sheet.sharedTable !== undefined && sheet.sharedPositional !== true && Array.isArray(sheet.sharedTable.strings)
    ? new Map(sheet.sharedTable.strings.map((text, position) => [text, position] as const))
    : null
  const shared = sharedIndex === null ? null : {
    indexOf: (text: string): number | null => {
      const found = sharedIndex.get(text)
      return found === undefined ? null : found
    },
  }
  yield '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
  if (sheet.cols !== undefined) yield `  <cols>${sheet.cols}</cols>\n`
  yield '  <sheetData>\n'
  let buffered = ''
  let bufferedCells = 0
  for (const row of sheet.rows) {
    const attrs = ` r="${row.r + 1}"${row.ht !== undefined ? ` ht="${row.ht}" customHeight="1"` : ''}${row.hidden === true ? ' hidden="1"' : ''}`
    const positional = sheet.sharedPositional === true
    buffered += `    <row${attrs}>${row.cells.map(({ row: cellRow, col, cell }) => {
      if (positional && typeof cell.value === 'string' && !cell.valueIsError && cell.formula === undefined) {
        const style = cell.styleIndex !== 0 ? ` s="${cell.styleIndex}"` : ''
        return `<c r="${a1Ref(cellRow, col)}" t="s"${style}><v>${cellRow}</v></c>`
      }
      return cellXml(a1Ref(cellRow, col), cell, shared)
    }).join('')}</row>\n`
    bufferedCells += Math.max(row.cells.length, 1)
    if (bufferedCells >= 2000) {
      yield buffered
      buffered = ''
      bufferedCells = 0
    }
  }
  if (buffered !== '') yield buffered
  yield '  </sheetData>\n'
  if (sheet.merges !== undefined && sheet.merges.length > 0) {
    yield `  <mergeCells count="${sheet.merges.length}">${sheet.merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>\n`
  }
  yield '</worksheet>'
}

async function writePackage(path: string, sheets: SheetSpec[], deflateData: boolean): Promise<{ bytes: number; uncompressedBytes: number }> {
  const zip = new StreamingZip(path)
  await zip.openFile()
  const shared = sheets.some((sheet) => sheet.sharedTable !== undefined)
  await zip.beginFile('[Content_Types].xml', 0)
  await zip.appendText(CONTENT_TYPES(sheets.length, shared))
  await zip.endFile()
  await zip.beginFile('_rels/.rels', 0)
  await zip.appendText(ROOT_RELS)
  await zip.endFile()
  await zip.beginFile('xl/workbook.xml', 0)
  await zip.appendText(WORKBOOK_XML(sheets.map((sheet) => sheet.name)))
  await zip.endFile()
  await zip.beginFile('xl/_rels/workbook.xml.rels', 0)
  await zip.appendText(WORKBOOK_RELS(sheets.length, shared))
  await zip.endFile()
  let sheetNumber = 0
  for (const sheet of sheets) {
    sheetNumber++
    await zip.beginFile(`xl/worksheets/sheet${sheetNumber}.xml`, deflateData ? 8 : 0)
    if (deflateData) {
      function* encoded(): Generator<Uint8Array> {
        for (const chunk of sheetXmlChunks(sheet)) yield encoder.encode(chunk)
      }
      await zip.appendDeflated(encoded())
    } else {
      for (const chunk of sheetXmlChunks(sheet)) await zip.appendText(chunk)
    }
    await zip.endFile()
    if (sheet.sharedTable !== undefined) {
      const table = sheet.sharedTable
      await zip.beginFile('xl/sharedStrings.xml', deflateData ? 8 : 0)
      const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${table.count}" uniqueCount="${table.count}">`
      const footer = '</sst>'
      if (deflateData) {
        function* encodedStrings(): Generator<Uint8Array> {
          yield encoder.encode(header)
          for (const text of table.strings) yield encoder.encode(`<si><t>${encodeXlsxString(text)}</t></si>`)
          yield encoder.encode(footer)
        }
        await zip.appendDeflated(encodedStrings())
      } else {
        await zip.appendText(header)
        for (const text of table.strings) await zip.appendText(`<si><t>${encodeXlsxString(text)}</t></si>`)
        await zip.appendText(footer)
      }
      await zip.endFile()
    }
  }
  await zip.beginFile('xl/styles.xml', 0)
  await zip.appendText(STYLES_XML)
  await zip.endFile()
  const uncompressedBytes = zip.payloadBytes()
  await zip.finish()
  return { bytes: statSync(path).size, uncompressedBytes }
}

function fullRows(handle: FixtureHandle): Iterable<SheetRowSpec> {
  return {
    *[Symbol.iterator]() {
      for (let row = 0; row < handle.rowCount; row++) {
        const cells: CellRef[] = []
        for (let col = 0; col < handle.colCount; col++) {
          const cell = handle.cellAt(row, col)
          if (cell !== null) cells.push({ row, col, cell })
        }
        if (cells.length > 0) yield { r: row, cells }
      }
    },
  }
}

interface SmallPlan {
  file: string
  fixture: FixtureId
  params: FixtureParams
  description: string
}

const SMALL_PLANS: SmallPlan[] = [
  { file: 'f1-small.xlsx', fixture: 'F1', params: { nnz: 200 }, description: 'corners plus 200 sampled cells' },
  { file: 'f5-small.xlsx', fixture: 'F5', params: {}, description: 'origin cells, A1:B2 merge, column widths, hidden row' },
  { file: 'f8-small.xlsx', fixture: 'F8', params: {}, description: 'parity cases with shared strings and metadata row' },
  { file: 'f3sparse-small.xlsx', fixture: 'F3', params: { rowCount: 256, colCount: 512, variant: 'sparse-wide' }, description: 'sparse-wide lattice sample' },
  { file: 'f4-small.xlsx', fixture: 'F4', params: { rowCount: 500 }, description: 'string-domain head incl. lone high/low surrogates' },
]

function smallSheet(plan: SmallPlan): SheetSpec {
  const handle = createFixture(plan.fixture, plan.params)
  if (plan.fixture === 'F1') {
    const cells = handle.enumerateHead(100_000)
    const byRow = new Map<number, CellRef[]>()
    for (const ref of cells) {
      const list = byRow.get(ref.row) ?? []
      list.push(ref)
      byRow.set(ref.row, list)
    }
    const rows: SheetRowSpec[] = [...byRow.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([r, refs]) => ({ r, cells: refs.sort((a, b) => a.col - b.col) }))
    return { name: plan.fixture, rows }
  }
  if (plan.fixture === 'F5') {
    const rows: SheetRowSpec[] = []
    for (let r = 0; r <= 4; r++) {
      const cells: CellRef[] = []
      for (let c = 0; c < 4; c++) {
        const cell = handle.cellAt(r, c)
        if (cell !== null) cells.push({ row: r, col: c, cell })
      }
      rows.push({ r, cells })
    }
    rows.push({ r: 5, hidden: true, cells: [] })
    // Row 5 stays empty on purpose: the handle populates (5,1), which is
    // outside this package's coverage window ({0..4, 0..3}).
    return {
      name: plan.fixture,
      rows,
      cols: '<col min="1" max="3" width="20" customWidth="1"/><col min="5" max="5" width="8" hidden="1"/>',
      merges: ['A1:B2'],
    }
  }
  if (plan.fixture === 'F3' || plan.fixture === 'F4') {
    const rows: SheetRowSpec[] = []
    for (let r = 0; r < handle.rowCount; r++) {
      const cells: CellRef[] = []
      for (let c = 0; c < handle.colCount; c++) {
        const cell = handle.cellAt(r, c)
        if (cell !== null) cells.push({ row: r, col: c, cell })
      }
      if (cells.length > 0) rows.push({ r, cells })
    }
    return { name: plan.fixture, rows }
  }
  const cases = f8Cases()
  const sharedStrings = cases.filter((c) => c.shared === true).map((c) => c.cell.value as string)
  const byRow = new Map<number, CellRef[]>()
  for (const c of cases) {
    const list = byRow.get(c.address.row) ?? []
    list.push({ row: c.address.row, col: c.address.col, cell: c.cell })
    byRow.set(c.address.row, list)
  }
  const rows: SheetRowSpec[] = [...byRow.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([r, refs]) => ({ r, cells: refs.sort((a, b) => a.col - b.col) }))
  rows.push({ r: 5, ht: 30, cells: [] })
  return {
    name: plan.fixture,
    rows,
    sharedTable: sharedStrings.length > 0 ? { count: sharedStrings.length, strings: sharedStrings } : undefined,
  }
}

const FIXTURE_DESCRIPTIONS: Record<FixtureId, string> = {
  F1: 'full-extents sparse corners (100k NNZ plus pinned A1/XFD1/A1048576/XFD1048576)',
  F2: 'tall table (1048576x8 with header, totals, blanks, errors, dates)',
  F3: 'wide table (dense 256x16384 plus sparse-wide lattice variant)',
  F4: 'string domain (1M unique UTF-16 strings with surrogates and accents)',
  F5: 'structural metadata (giant merges, tables, spills, drawings)',
  F6: 'full-axis synthetic provider (all 2^34 addresses, analytic)',
  F7: 'view stress (eight-key sorts, hidden interleavings, tiny sizes)',
  F8: 'compatibility parity (cache, error, date, table, unicode cases)',
}

interface ManifestFile {
  path: string
  bytes: number
  sha256: string
  uncompressedBytes: number
}

interface ManifestFixture {
  id: FixtureId
  description: string
  seed: number
  rowCount: number
  colCount: number
  nnz: number
  used: { r0: number; r1: number; c0: number; c1: number }
  samplesSha256: string
  queryScope: string
  queryHashes: { readRowsHead: string; overviewHead: string }
  package: string | null
  packageFixture: FixtureId | null
  packageParams: FixtureParams | null
  packageNnz: number | null
  coverage: { r0: number; r1: number; c0: number; c1: number } | null
  extraPackages: Array<{
    path: string; fixture: FixtureId; params: FixtureParams | null; nnz: number
    coverage: { r0: number; r1: number; c0: number; c1: number }
  }>
  files: ManifestFile[]
  stringTable?: 'inline' | 'shared'
}

interface MeasuredRun {
  fixture: FixtureId
  file: string
  bytes: number
  uncompressedBytes: number
  sha256: string
  seconds: number
  peakRssMb: number
  maxRssBytes: number
  committed: boolean
}

/** Streaming SHA-256: the file is never fully buffered (see item 12). maxRSS
 * is recorded raw as maxRssBytes: on this platform resourceUsage() reports
 * bytes (125M for a ~104MB process), not kilobytes. */
async function shaFile(path: string): Promise<string> {
  const hash = createHash('sha256')
  for await (const chunk of createReadStream(path)) {
    hash.update(chunk as Uint8Array)
  }
  return hash.digest('hex')
}

function coverageFor(plan: SmallPlan): { r0: number; r1: number; c0: number; c1: number } {
  switch (plan.fixture) {
    case 'F5': return { r0: 0, r1: 4, c0: 0, c1: 3 }
    case 'F3': return { r0: 0, r1: 255, c0: 0, c1: 511 }
    case 'F4': return { r0: 0, r1: 499, c0: 0, c1: 0 }
    default: return { r0: 0, r1: 1048575, c0: 0, c1: 16383 }
  }
}

async function main(): Promise<void> {
  const args = process.argv.slice(2)
  const outIndex = args.indexOf('--out')
  if (outIndex < 0 || outIndex + 1 >= args.length) throw new Error('usage: generate-viewer-fixtures.ts --out <dir> [--fixtures F1,F2,...] [--scale small|full] [--string-table inline|shared]')
  const out = resolve(args[outIndex + 1])
  const fixturesIndex = args.indexOf('--fixtures')
  const scaleIndex = args.indexOf('--scale')
  const tableIndex = args.indexOf('--string-table')
  const scale = scaleIndex >= 0 ? args[scaleIndex + 1] : 'small'
  if (scale !== 'small' && scale !== 'full') throw new Error(`unknown scale ${scale}`)
  const stringTable = tableIndex >= 0 ? args[tableIndex + 1] : 'inline'
  if (stringTable !== 'inline' && stringTable !== 'shared') throw new Error(`unknown string table ${stringTable}`)
  const wanted = (fixturesIndex >= 0 ? args[fixturesIndex + 1].split(',') : ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8']) as FixtureId[]
  mkdirSync(out, { recursive: true })
  const manifest: { version: number; fixtures: ManifestFixture[] } = { version: 1, fixtures: [] }
  const measured: { generatedAt: string; runs: MeasuredRun[] } = { generatedAt: new Date().toISOString(), runs: [] }
  const smallFiles: ManifestFile[] = []
  const extraPackages: ManifestFixture['extraPackages'] = []
  if (scale === 'small') {
    for (const plan of SMALL_PLANS) {
      if (!wanted.includes(plan.fixture)) continue
      const path = join(out, plan.file)
      const started = Date.now()
      const probe = createFixture(plan.fixture, plan.params)
      const spec = smallSheet(plan)
      const result = await writePackage(path, [spec], false)
      const seconds = (Date.now() - started) / 1000
      const sha256 = await shaFile(path)
      const peakRssMb = Math.round(process.memoryUsage().rss / 1048576)
      const maxRssBytes = process.resourceUsage().maxRSS
      smallFiles.push({ path: plan.file, bytes: result.bytes, sha256, uncompressedBytes: result.uncompressedBytes })
      measured.runs.push({ fixture: plan.fixture, file: plan.file, bytes: result.bytes, uncompressedBytes: result.uncompressedBytes, sha256, seconds, peakRssMb, maxRssBytes, committed: true })
      const coverage = coverageFor(plan)
      const coverRange = { r0: coverage.r0, r1: coverage.r1, c0: coverage.c0, c1: coverage.c1 }
      let nnz: number
      if (plan.fixture === 'F5') {
        nnz = probe.scanCells(coverRange, 0, 1_000_000).length
      } else {
        try {
          nnz = probe.enumerateHead(1_000_000).length
        } catch {
          nnz = probe.scanCells(coverRange, 0, 1_000_000).length
        }
      }
      console.log(`wrote ${plan.file}: ${result.bytes} bytes (${result.uncompressedBytes} uncompressed) in ${seconds.toFixed(1)}s`)
      if (plan.file === 'f3sparse-small.xlsx' || plan.file === 'f4-small.xlsx') {
        extraPackages.push({ path: plan.file, fixture: plan.fixture, params: plan.params, nnz, coverage })
      }
    }
  }
  for (const id of wanted) {
    const handle = createFixture(id)
    const files = scale === 'small' ? smallFiles.filter((file) => file.path.startsWith(id.toLowerCase())) : []
    const plan = SMALL_PLANS.find((candidate) => candidate.fixture === id && !candidate.file.startsWith('f3sparse') && !candidate.file.startsWith('f4-small'))
    let packageNnz: number | null = null
    let coverage: ManifestFixture['coverage'] = null
    if (plan !== undefined && scale === 'small') {
      const probe = createFixture(plan.fixture, plan.params)
      coverage = coverageFor(plan)
      packageNnz = plan.fixture === 'F5'
        ? probe.scanCells({ r0: 0, r1: 4, c0: 0, c1: 3 }, 0, 1_000_000).length
        : probe.enumerateHead(1_000_000).length
    }
    if (scale === 'full' && (id === 'F2' || id === 'F3' || id === 'F4')) {
      const file = `${id.toLowerCase()}-full.xlsx`
      const path = join(out, file)
      const started = Date.now()
      const sheets: SheetSpec[] = id === 'F4' && stringTable === 'shared'
        ? [{
          name: id,
          rows: fullRows(handle),
          sharedTable: {
            count: handle.rowCount,
            strings: (function* strings(): Generator<string> {
              for (let row = 0; row < handle.rowCount; row++) yield f4StringFor(handle.seed, row)
            })(),
          },
          sharedPositional: true,
        }]
        : [{ name: id, rows: fullRows(handle) }]
      const result = await writePackage(path, sheets, true)
      const seconds = (Date.now() - started) / 1000
      const sha256 = await shaFile(path)
      const peakRssMb = Math.round(process.memoryUsage().rss / 1048576)
      const maxRssBytes = process.resourceUsage().maxRSS
      measured.runs.push({ fixture: id, file, bytes: result.bytes, uncompressedBytes: result.uncompressedBytes, sha256, seconds, peakRssMb, maxRssBytes, committed: false })
      console.log(`wrote ${file}: ${result.bytes} bytes (${result.uncompressedBytes} uncompressed) in ${seconds.toFixed(0)}s, peak RSS ${peakRssMb} MB`)
    }
    manifest.fixtures.push({
      id,
      description: FIXTURE_DESCRIPTIONS[id],
      seed: handle.seed,
      rowCount: handle.rowCount,
      colCount: handle.colCount,
      nnz: handle.nnz(),
      used: { ...handle.used },
      samplesSha256: sha256Hex(handle.canonicalSamples()),
      queryScope: 'head (readRows r0..r4, overview rows 0..999)',
      queryHashes: await headQueryHashes(handle),
      package: plan !== undefined && scale === 'small' ? plan.file : null,
      packageFixture: plan !== undefined && scale === 'small' ? plan.fixture : null,
      packageParams: plan !== undefined && scale === 'small' ? plan.params : null,
      packageNnz,
      coverage,
      extraPackages: scale === 'small' ? extraPackages.filter((extra) => extra.fixture === id) : [],
      files,
      ...(id === 'F4' && scale === 'full' ? { stringTable: stringTable as 'inline' | 'shared' } : {}),
    })
  }
  writeFileSync(join(out, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`)
  writeFileSync(join(out, 'measured.json'), `${JSON.stringify(measured, null, 2)}\n`)
  console.log(`wrote manifest.json with ${manifest.fixtures.length} fixtures and measured.json with ${measured.runs.length} runs`)
}

const isMain = (import.meta as unknown as { main?: boolean }).main === true
if (isMain) {
  await main()
}
