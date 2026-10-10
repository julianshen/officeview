/** Run: bun scripts/generate-viewer-fixtures.ts --out <dir> [--fixtures F1,F2,...] [--scale small|full]
 * Writes TRUE XLSX packages with a streaming store-only ZIP writer: sheet XML
 * is generated row by row and flushed in bounded chunks, so full-size F2/F3
 * never hold a whole-sheet string or object model in memory. No new
 * dependencies; small committed packages reuse the same writer. Manifest
 * (manifest.json) is written to the out directory alongside the packages. */
import { createHash } from 'node:crypto'
import { mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { open } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { columnToLetters } from '../src/xlsx/view/coordinates'
import {
  createFixture, f8Cases, readOptionsFor, sha256Hex,
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
  crc: number
  size: number
  offset: number
}

/** Minimal streaming ZIP writer (stored entries, data descriptors). */
class StreamingZip {
  private fd: Awaited<ReturnType<typeof open>> | null = null
  private offset = 0
  private entries: ZipEntry[] = []
  private current: { name: string; crc: number; size: number; offset: number } | null = null

  constructor(private readonly path: string) {}

  async openFile(): Promise<void> {
    this.fd = await open(this.path, 'w')
  }

  private async write(bytes: Uint8Array): Promise<void> {
    await this.fd!.write(bytes)
    this.offset += bytes.length
  }

  async beginFile(name: string): Promise<void> {
    const nameBytes = encoder.encode(name)
    const header = new Uint8Array(30 + nameBytes.length)
    header.set([0x50, 0x4b, 0x03, 0x04], 0)
    header.set(u16(20), 4)
    header.set(u16(0x08), 6)
    header.set(u16(0), 8)
    header.set(u16(DOS_TIME), 10)
    header.set(u16(DOS_DATE), 12)
    header.set(u32(0), 14)
    header.set(u32(0), 18)
    header.set(u32(0), 22)
    header.set(u16(nameBytes.length), 26)
    header.set(u16(0), 28)
    header.set(nameBytes, 30)
    this.current = { name, crc: 0, size: 0, offset: this.offset }
    await this.write(header)
  }

  async appendChunk(bytes: Uint8Array): Promise<void> {
    if (this.current === null) throw new Error('no open zip entry')
    this.current.crc = crcUpdate(this.current.crc, bytes)
    this.current.size += bytes.length
    await this.write(bytes)
  }

  async appendText(text: string): Promise<void> {
    await this.appendChunk(encoder.encode(text))
  }

  async endFile(): Promise<void> {
    if (this.current === null) throw new Error('no open zip entry')
    const descriptor = new Uint8Array(16)
    descriptor.set([0x50, 0x4b, 0x07, 0x08], 0)
    descriptor.set(u32(this.current.crc), 4)
    descriptor.set(u32(this.current.size), 8)
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
      header.set(u16(0), 10)
      header.set(u16(DOS_TIME), 12)
      header.set(u16(DOS_DATE), 14)
      header.set(u32(entry.crc), 16)
      header.set(u32(entry.size), 20)
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

/** ST_Xstring encoding: lone surrogates and C0 controls become _xHHHH_,
 * literal _xHHHH_ text is guarded with _x005F so decoding stays single-pass. */
export function encodeXlsxString(value: string): string {
  const guarded = value.replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_')
  let out = ''
  for (let i = 0; i < guarded.length; i++) {
    const unit = guarded.charCodeAt(i)
    const lead = unit >= 0xd800 && unit <= 0xdbff
    const trail = unit >= 0xdc00 && unit <= 0xdfff
    if (lead) {
      const next = guarded.charCodeAt(i + 1)
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += guarded[i] + guarded[i + 1]
        i++
        continue
      }
      out += `_x${unit.toString(16).toUpperCase().padStart(4, '0')}_`
      continue
    }
    if (trail) {
      out += `_x${unit.toString(16).toUpperCase().padStart(4, '0')}_`
      continue
    }
    if (unit < 0x20 && unit !== 0x09 && unit !== 0x0a && unit !== 0x0d) {
      out += `_x${unit.toString(16).toUpperCase().padStart(4, '0')}_`
      continue
    }
    out += guarded[i]
  }
  return escapeXml(out)
}

function a1Ref(row: number, col: number): string {
  return `${columnToLetters(col)}${row + 1}`
}

function cellXml(ref: string, cell: FixtureCell, sharedIndex?: Map<string, number>): string {
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
  const text = encodeXlsxString(cell.value ?? '')
  if (sharedIndex !== undefined && sharedIndex.has(cell.value as string)) {
    return `<c r="${ref}" t="s"${style}>${formulaXml}<v>${sharedIndex.get(cell.value as string)}</v></c>`
  }
  if (cell.formula !== undefined) {
    return `<c r="${ref}" t="str"${style}>${formulaXml}<v>${text}</v></c>`
  }
  return `<c r="${ref}" t="inlineStr"${style}><is><t>${text}</t></is></c>`
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

interface SheetSpec {
  name: string
  /** Row-major row iterator; must stream (no full row array) for large sheets. */
  rows: Iterable<SheetRowSpec>
  cols?: string
  merges?: string[]
  sharedStrings?: string[]
}

async function writePackage(path: string, sheets: SheetSpec[]): Promise<{ bytes: number; uncompressedBytes: number }> {
  const zip = new StreamingZip(path)
  await zip.openFile()
  const shared = sheets.some((sheet) => (sheet.sharedStrings ?? []).length > 0)
  await zip.beginFile('[Content_Types].xml')
  await zip.appendText(CONTENT_TYPES(sheets.length, shared))
  await zip.endFile()
  await zip.beginFile('_rels/.rels')
  await zip.appendText(ROOT_RELS)
  await zip.endFile()
  await zip.beginFile('xl/workbook.xml')
  await zip.appendText(WORKBOOK_XML(sheets.map((sheet) => sheet.name)))
  await zip.endFile()
  await zip.beginFile('xl/_rels/workbook.xml.rels')
  await zip.appendText(WORKBOOK_RELS(sheets.length, shared))
  await zip.endFile()
  let sheetNumber = 0
  for (const sheet of sheets) {
    sheetNumber++
    const sharedIndex = sheet.sharedStrings !== undefined && sheet.sharedStrings.length > 0
      ? new Map(sheet.sharedStrings.map((text, position) => [text, position]))
      : undefined
    await zip.beginFile(`xl/worksheets/sheet${sheetNumber}.xml`)
    await zip.appendText(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n`)
    if (sheet.cols !== undefined) await zip.appendText(`  <cols>${sheet.cols}</cols>\n`)
    await zip.appendText('  <sheetData>\n')
    let buffered = ''
    let bufferedCells = 0
    const flush = async (): Promise<void> => {
      if (buffered !== '') {
        await zip.appendText(buffered)
        buffered = ''
        bufferedCells = 0
      }
    }
    for (const row of sheet.rows) {
      const attrs = ` r="${row.r + 1}"${row.ht !== undefined ? ` ht="${row.ht}" customHeight="1"` : ''}${row.hidden === true ? ' hidden="1"' : ''}`
      buffered += `    <row${attrs}>${row.cells.map(({ row: cellRow, col, cell }) => cellXml(a1Ref(cellRow, col), cell, sharedIndex)).join('')}</row>\n`
      bufferedCells += Math.max(row.cells.length, 1)
      if (bufferedCells >= 2000) await flush()
    }
    await flush()
    await zip.appendText('  </sheetData>\n')
    if (sheet.merges !== undefined && sheet.merges.length > 0) {
      await zip.appendText(`  <mergeCells count="${sheet.merges.length}">${sheet.merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>\n`)
    }
    await zip.appendText('</worksheet>')
    await zip.endFile()
    if (sharedIndex !== undefined && sheetNumber === 1) {
      await zip.beginFile('xl/sharedStrings.xml')
      await zip.appendText(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sheet.sharedStrings!.length}" uniqueCount="${sheet.sharedStrings!.length}">`)
      for (const text of sheet.sharedStrings!) {
        await zip.appendText(`<si><t>${escapeXml(text)}</t></si>`)
      }
      await zip.appendText('</sst>')
      await zip.endFile()
    }
  }
  await zip.beginFile('xl/styles.xml')
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
    for (let r = 0; r < 4; r++) {
      const cells: CellRef[] = []
      for (let c = 0; c < 4; c++) {
        const cell = handle.cellAt(r, c)
        if (cell !== null) cells.push({ row: r, col: c, cell })
      }
      rows.push({ r, cells })
    }
    rows.push({ r: 4, hidden: true, cells: [] })
    return {
      name: plan.fixture,
      rows,
      cols: '<col min="1" max="3" width="20" customWidth="1"/><col min="5" max="5" width="8" hidden="1"/>',
      merges: ['A1:B2'],
    }
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
  return { name: plan.fixture, rows, sharedStrings }
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
  files: ManifestFile[]
  measured: { bytes: number; uncompressedBytes: number; sha256: string; seconds: number; peakRssMb: number; committed: false } | null
}

async function queryHashes(handle: FixtureHandle): Promise<{ readRowsHead: string; overviewHead: string }> {
  const store = handle.createStore()
  const options = readOptionsFor(store)
  const headCols = Math.min(3, handle.colCount - 1)
  const rowsPage = await store.readRows({ r0: 0, r1: 4, c0: 0, c1: headCols }, undefined, options)
  const overviewPage = await store.queryOverview({
    groups: [{ groupId: 'head', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 999 }] }, cols: [{ c0: 0, c1: headCols }] }],
    maxGroups: 2048,
  }, options)
  return {
    readRowsHead: sha256Hex(JSON.stringify(rowsPage)),
    overviewHead: sha256Hex(JSON.stringify(overviewPage)),
  }
}

function shaFile(path: string): string {
  return createHash('sha256').update(readFileSync(path)).digest('hex')
}

async function main(): Promise<void> {
  const args = process.argv.slice(2)
  const outIndex = args.indexOf('--out')
  if (outIndex < 0 || outIndex + 1 >= args.length) throw new Error('usage: generate-viewer-fixtures.ts --out <dir> [--fixtures F1,F2,...] [--scale small|full]')
  const out = resolve(args[outIndex + 1])
  const fixturesIndex = args.indexOf('--fixtures')
  const scaleIndex = args.indexOf('--scale')
  const scale = scaleIndex >= 0 ? args[scaleIndex + 1] : 'small'
  if (scale !== 'small' && scale !== 'full') throw new Error(`unknown scale ${scale}`)
  const wanted = (fixturesIndex >= 0 ? args[fixturesIndex + 1].split(',') : ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8']) as FixtureId[]
  mkdirSync(out, { recursive: true })
  const manifest: { version: number; fixtures: ManifestFixture[] } = { version: 1, fixtures: [] }
  const smallFiles: ManifestFile[] = []
  if (scale === 'small') {
    for (const plan of SMALL_PLANS) {
      if (!wanted.includes(plan.fixture)) continue
      const path = join(out, plan.file)
      const started = Date.now()
      const result = await writePackage(path, [smallSheet(plan)])
      const seconds = (Date.now() - started) / 1000
      const sha256 = shaFile(path)
      smallFiles.push({ path: plan.file, bytes: result.bytes, sha256, uncompressedBytes: result.uncompressedBytes })
      console.log(`wrote ${plan.file}: ${result.bytes} bytes (${result.uncompressedBytes} uncompressed) in ${seconds.toFixed(1)}s`)
    }
  }
  for (const id of wanted) {
    const handle = createFixture(id)
    const files = scale === 'small' ? smallFiles.filter((file) => file.path.startsWith(id.toLowerCase())) : []
    let measured: ManifestFixture['measured'] = null
    let packageNnz: number | null = null
    if (scale === 'full' && (id === 'F2' || id === 'F3' || id === 'F4')) {
      const file = `${id.toLowerCase()}-full.xlsx`
      const path = join(out, file)
      const started = Date.now()
      const startRss = process.memoryUsage().rss
      void startRss
      const result = await writePackage(path, [{ name: id, rows: fullRows(handle) }])
      const seconds = (Date.now() - started) / 1000
      const peakRssMb = Math.round(process.memoryUsage().rss / 1048576)
      measured = { bytes: result.bytes, uncompressedBytes: result.uncompressedBytes, sha256: shaFile(path), seconds, peakRssMb, committed: false }
      console.log(`wrote ${file}: ${result.bytes} bytes (${result.uncompressedBytes} uncompressed) in ${seconds.toFixed(0)}s, peak RSS ${peakRssMb} MB`)
    }
    const plan = SMALL_PLANS.find((candidate) => candidate.fixture === id)
    if (plan !== undefined && scale === 'small') {
      if (plan.fixture === 'F5') {
        packageNnz = 0
        const probe = createFixture(plan.fixture, plan.params)
        for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) {
          if (probe.cellAt(r, c) !== null) packageNnz++
        }
      } else {
        packageNnz = createFixture(plan.fixture, plan.params).enumerateHead(100_000).length
      }
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
      queryHashes: await queryHashes(handle),
      package: plan !== undefined && scale === 'small' ? plan.file : null,
      packageFixture: plan !== undefined && scale === 'small' ? plan.fixture : null,
      packageParams: plan !== undefined && scale === 'small' ? plan.params : null,
      packageNnz,
      files,
      measured,
    })
  }
  writeFileSync(join(out, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`)
  console.log(`wrote manifest.json with ${manifest.fixtures.length} fixtures`)
}

await main()
