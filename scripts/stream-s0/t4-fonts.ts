/**
 * Stage-0 Task 4 — T4 font probes.
 *
 * Produces `docs/validation/streaming-s0/t4.json`:
 *   { gate, status REPORTED, backend, gitSha, generatedAt,
 *     perFile: [{ file, faces: Face[], verifiable, unverifiable,
 *                 unitsThatWouldRaster }],
 *     totals: { files, faces, verifiable, unverifiable,
 *               unitsThatWouldRaster, unitsTotal, rasterFallbackRate },
 *     fixtureProbe: [{ text, widthNode, widthNapi, relDiff }],   // CJK/ligature/BiDi
 *     familiesCausingRaster: [family…] }
 *
 * Face = { id, family, source: 'embedded'|'bundled'|'system', bytes?,
 *          hash96? } — hash96 = first 12 bytes of SHA-256 of face bytes,
 * hex (24 hex chars). A face with no bytes is unverifiable, and §7.4 rule 3
 * says every unit that uses such a face ships as a raster unit — the
 * raster-fallback rate is the number the owner needs.
 *
 * Run under napi (guards like T1): `bun run s0:napi scripts/stream-s0/t4-fonts.ts`.
 */
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, readdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { loadOfficeFile } from '../../src/components/OfficeFile'
import { getPaintables } from '../../src/render/paint'
import { createCanvas, registerFont, type Canvas } from 'canvas'
import { listCorpus, type CorpusFile } from './corpus'
import { wrapCensus } from './census-proxy'
import { writeResult, resolveBackend } from './results'

const ZOOM = 1
const DPR = 1
/** Fixture probe strings: CJK, ligature, Arabic (RTL), Hebrew (RTL). */
const FIXTURE_PROBES: ReadonlyArray<{ text: string; label: string }> = [
  { text: '漢字テスト', label: 'cjk' },
  { text: 'ffi ff fl', label: 'ligature' },
  { text: 'مرحبا', label: 'arabic-rtl' },
  { text: 'שלום', label: 'hebrew-rtl' },
]

export interface Face {
  id: string
  family: string
  source: 'embedded' | 'bundled' | 'system'
  bytes?: Uint8Array
  /** First 12 bytes of SHA-256 of face bytes, hex (24 hex chars). */
  hash96?: string
}

export interface FaceClassification {
  faces: Face[]
  verifiable: number
  unverifiable: number
  unitsThatWouldRaster: number
}

/** hash96: first 12 bytes of SHA-256 of the face bytes, hex (24 chars). */
export function hash96(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex').slice(0, 24)
}

/** Known bundled font assets (repo font files); family → face id. */
function bundledFaces(): Face[] {
  const dir = path.join(process.cwd(), 'tests', 'fixtures', 'fonts')
  if (!existsSync(dir)) return []
  const out: Face[] = []
  for (const name of readdirSync(dir)) {
    if (!/\.(ttf|otf|woff2?)$/i.test(name)) continue
    const bytes = new Uint8Array(readFileSyncBytes(path.join(dir, name)))
    const family = name.replace(/\.[^.]+$/, '')
    out.push({ id: `bundled:${family}`, family, source: 'bundled', bytes, hash96: hash96(bytes) })
  }
  return out
}

/**
 * Proposed server-pinned font pack (owner ruling 2026-11-08, spec §7.4):
 * open, metric-compatible faces the server lays out with and the client
 * holds byte-identically. Document family (lowercase) → pack file. Only
 * metric-compatible substitutes are listed (Calibri Light has none, so it
 * stays a raster cause).
 */
export const PACK_ALIASES: Record<string, string> = {
  calibri: 'Carlito-Regular.ttf',
  carlito: 'Carlito-Regular.ttf',
  cambria: 'Caladea-Regular.ttf',
  caladea: 'Caladea-Regular.ttf',
  arial: 'LiberationSans-Regular.ttf',
  helvetica: 'LiberationSans-Regular.ttf',
  'liberation sans': 'LiberationSans-Regular.ttf',
  'times new roman': 'LiberationSerif-Regular.ttf',
  'liberation serif': 'LiberationSerif-Regular.ttf',
  'courier new': 'LiberationMono-Regular.ttf',
  'liberation mono': 'LiberationMono-Regular.ttf',
}

/** Pack faces present on disk: family → hash96 of the pinned bytes. */
export function packFaces(dir = path.join(process.cwd(), 'fontpack')): Map<string, string> {
  const out = new Map<string, string>()
  for (const [family, file] of Object.entries(PACK_ALIASES)) {
    const at = path.join(dir, file)
    if (existsSync(at)) out.set(family, hash96(readFileSyncBytes(at)))
  }
  return out
}

function readFileSyncBytes(at: string): Uint8Array {
  return new Uint8Array(readFileSyncInternal(at))
}
function readFileSyncInternal(at: string): Buffer {
  return require('node:fs').readFileSync(at) as Buffer
}

/** Extract the family name(s) from a canvas `ctx.font` shorthand string. */
export function fontFamilies(font: string): string[] {
  const quoted = [...font.matchAll(/"([^"]+)"|'([^']+)'/g)].map((m) => m[1] ?? m[2])
  if (quoted.length > 0) return quoted
  // Unquoted: the family list trails the size/weight tokens; drop leading
  // style/size tokens heuristically (numbers, px/pt keywords, weights).
  const tail = font.split(/[\s,]+/).filter((token) => token.length > 0)
  const out: string[] = []
  for (const token of tail) {
    if (/^[\d.]+(px|pt|em|rem|%)?(\/[-\d.]+)?$/i.test(token)) continue
    if (/^(normal|italic|oblique|bold|bolder|lighter|ultra-|extra-|semi-)([-\s].*)?$/i.test(token)) continue
    if (/^(system-ui|sans-serif|serif|monospace|cursive|fantasy)$/i.test(token)) continue
    out.push(token)
  }
  return out
}

/** Register one embedded face for the backend under test (path-based; the
 * bytes are spooled once per face). Returns the registration alias. */
const spooledFaces = new Map<string, string>()
function registerEmbeddedFace(face: { family: string; bytes: Uint8Array; variant: string }, docId: string): string {
  const key = `${docId}:${face.family}:${face.variant}`
  let alias = spooledFaces.get(key)
  if (alias) return alias
  const spool = path.join(tmpdir(), 'stream-s0-fonts')
  mkdirSync(spool, { recursive: true })
  const file = path.join(spool, `${hash96(face.bytes)}-${face.variant.replace(/[^a-z]/g, '')}.fnt`)
  if (!existsSync(file)) writeFileSync(file, face.bytes)
  alias = `${face.family}`
  registerFont(file, { family: alias })
  spooledFaces.set(key, alias)
  return alias
}

/**
 * requiredFaces for one corpus file: embedded faces (with bytes → verifiable)
 * from the parsed doc, bundled faces (repo font assets), and the census's
 * font strings classify the rest as system (unverifiable → raster).
 */
export async function requiredFaces(file: CorpusFile): Promise<FaceClassification & { units: Array<{ unit: number; families: string[] }>; unitsThatWouldRasterWithPack: number }> {
  const doc = await loadOfficeFile(readFileSyncBytes(file.path))
  const embedded = (doc as { embeddedFonts?: Array<{ family: string; variant: string; bytes: Uint8Array }> }).embeddedFonts ?? []
  const faces: Face[] = embedded.map((face, index) => ({
    id: `embedded:${face.family}:${face.variant}:${index}`,
    family: face.family,
    source: 'embedded',
    bytes: face.bytes,
    hash96: hash96(face.bytes),
  }))
  // Register every embedded face so the census sees the real family names the
  // painters emit (and per §7.3/§7.4 the faces really load under napi).
  const docId = path.basename(file.name)
  const aliasByFamily = new Map<string, string>()
  for (const face of embedded) {
    aliasByFamily.set(face.family.toLowerCase(), registerEmbeddedFace(face, docId))
  }
  // Bundled faces register too (they are real bytes; verifiable).
  for (const face of bundledFaces()) {
    const already = faces.some((f) => f.hash96 === face.hash96)
    if (!already) faces.push(face)
    registerEmbeddedFace({ family: face.family, bytes: face.bytes!, variant: 'regular' }, docId)
    aliasByFamily.set(face.family.toLowerCase(), face.family)
  }
  const verifiableNames = new Set(faces.filter((f) => f.bytes).map((f) => f.family.toLowerCase()))

  // Census over the units to classify the fonts each unit uses.
  const paintable = await getPaintables(doc)
  const perUnit: Array<{ unit: number; families: string[] }> = []
  let unitsThatWouldRaster = 0
  let unitsThatWouldRasterWithPack = 0
  const pack = packFaces()
  for (let index = 0; index < paintable.length; index++) {
    const paintable0 = paintable[index]
    const canvas: Canvas = createCanvas(
      Math.max(1, Math.ceil(paintable0.spec.widthPx * ZOOM * DPR)),
      Math.max(1, Math.ceil(paintable0.spec.heightPx * ZOOM * DPR)),
    )
    const { ctx, report } = wrapCensus(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)
    paintable0.paint(ctx)
    void canvas
    const used = (report().props.font?.values ?? [])
      .filter((v) => v.startsWith('string:'))
      .map((v) => v.slice('string:'.length))
      .flatMap(fontFamilies)
      .map((family) => (aliasByFamily.get(family.toLowerCase()) ?? family).toLowerCase())
    const unique = [...new Set(used)]
    perUnit.push({ unit: index, families: unique })
    const wouldRaster = unique.some((family) => !verifiableNames.has(family))
    if (wouldRaster) unitsThatWouldRaster++
    if (unique.some((family) => !verifiableNames.has(family) && !pack.has(family))) unitsThatWouldRasterWithPack++
  }
  // Faces the census saw but that are neither embedded nor bundled: system.
  const seenFamilies = new Set<string>()
  for (const u of perUnit) for (const family of u.families) seenFamilies.add(family)
  for (const family of seenFamilies) {
    if (!faces.some((f) => f.family.toLowerCase() === family)) {
      faces.push({ id: `system:${family}`, family, source: 'system' })
    }
  }
  const verifiable = faces.filter((f) => f.bytes).length
  const unverifiable = faces.length - verifiable
  paintable.dispose()
  return { faces, verifiable, unverifiable, unitsThatWouldRaster, unitsThatWouldRasterWithPack, units: perUnit }
}

async function run(): Promise<number> {
  // T1's guard: the T4 probe must run under the napi preload (the face
  // resolution compares napi against node-canvas; the parent must be napi).
  if (resolveBackend() !== '@napi-rs/canvas') {
    console.error('[t4] REFUSED: run under the napi preload — bun run s0:napi scripts/stream-s0/t4-fonts.ts')
    return 2
  }
  const corpus = listCorpus()
  const perFile: Array<Record<string, unknown>> = []
  const totals = { files: 0, faces: 0, verifiable: 0, unverifiable: 0, unitsThatWouldRaster: 0, unitsThatWouldRasterWithPack: 0, unitsTotal: 0 }
  const rasterCauses = new Map<string, number>()
  const rasterCausesWithPack = new Map<string, number>()
  const pack = packFaces()

  // The CJK/ligature/BiDi advance probe: napi vs node-canvas (the node-canvas
  // arm runs in a child node process — the preload maps `canvas` in-process).
  const nodeChildScript = `
const { createCanvas } = await import('canvas')
const ctx = createCanvas(128, 64).getContext('2d')
const out = {}
for (const [label, text] of JSON.parse(process.argv[1])) {
  ctx.font = '24px sans-serif'
  out[label] = ctx.measureText(text).width
}
console.log(JSON.stringify(out))
`
  const probeJson = JSON.stringify(FIXTURE_PROBES.map((p) => [p.label, p.text]))
  const raw = execFileSync('node', ['--input-type=module', '--eval', nodeChildScript, probeJson], { encoding: 'utf8', timeout: 10_000 })
  const nodeWidths = JSON.parse(raw.trim().split(/\r?\n/).at(-1)!) as Record<string, number>
  const fixtureProbe = FIXTURE_PROBES.map(({ label, text }) => {
    const ctx = createCanvas(128, 64).getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.font = '24px sans-serif'
    const widthNapi = ctx.measureText(text).width
    const widthNode = nodeWidths[label] ?? 0
    const denom = Math.max(widthNapi, widthNode, 1)
    return { label, text, widthNode: Math.round(widthNode * 100) / 100, widthNapi: Math.round(widthNapi * 100) / 100, relDiff: Math.round(Math.abs(widthNode - widthNapi) / denom * 10_000) / 10_000 }
  })

  for (const file of corpus) {
    try {
      const result = await requiredFaces(file)
      totals.files++
      totals.faces += result.faces.length
      totals.verifiable += result.verifiable
      totals.unverifiable += result.unverifiable
      totals.unitsThatWouldRaster += result.unitsThatWouldRaster
      totals.unitsThatWouldRasterWithPack += result.unitsThatWouldRasterWithPack
      totals.unitsTotal += result.units.length
      const rasterFamilies = new Set<string>()
      for (const unit of result.units) {
        for (const family of unit.families) {
          if (!result.faces.some((f) => f.family.toLowerCase() === family && f.bytes)) rasterFamilies.add(family)
        }
      }
      for (const family of rasterFamilies) rasterCauses.set(family, (rasterCauses.get(family) ?? 0) + 1)
      for (const family of rasterFamilies) if (!pack.has(family)) rasterCausesWithPack.set(family, (rasterCausesWithPack.get(family) ?? 0) + 1)
      perFile.push({
        file: file.name,
        // The report carries face METADATA, never the bytes themselves
        // (round-9/10 honesty: an 85 MB JSON of hex-dumped font bytes was
        // committed once already — hash96 identifies the face; the bytes
        // live in the tmp spool only).
        faces: result.faces.map((face) => ({
          id: face.id,
          family: face.family,
          source: face.source,
          bytesLength: face.bytes?.length,
          hash96: face.hash96,
        })),
        verifiable: result.verifiable,
        unverifiable: result.unverifiable,
        unitsThatWouldRaster: result.unitsThatWouldRaster,
        unitsThatWouldRasterWithPack: result.unitsThatWouldRasterWithPack,
        units: result.units.length,
        rasterFamilies: [...new Set(result.units
          .flatMap((u) => u.families)
          .filter((family) => !result.faces.some((f) => f.family.toLowerCase() === family && f.bytes)))],
      })
    } catch (error) {
      perFile.push({ file: file.name, error: (error as Error).message })
    }
  }

  const rasterFallbackRate = totals.unitsTotal > 0 ? Math.round((totals.unitsThatWouldRaster / totals.unitsTotal) * 10_000) / 10_000 : 0
  const rasterFallbackRateWithPack = totals.unitsTotal > 0 ? Math.round((totals.unitsThatWouldRasterWithPack / totals.unitsTotal) * 10_000) / 10_000 : 0
  writeResult('t4', {
    status: 'REPORTED',
    perFile,
    totals: { ...totals, rasterFallbackRate, rasterFallbackRateWithPack },
    // Counts are FILES in which the family forced at least one raster unit.
    familiesCausingRaster: [...rasterCauses.entries()].sort((a, b) => b[1] - a[1]).map(([family, count]) => ({ family, count })),
    familiesCausingRasterWithPack: [...rasterCausesWithPack.entries()].sort((a, b) => b[1] - a[1]).map(([family, count]) => ({ family, count })),
    pinnedPack: [...pack.entries()].map(([family, hash]) => ({ family, file: PACK_ALIASES[family], hash96: hash })),
    fixtureProbe,
    note: 'REPORTED per plan Task 4 — the raster-fallback rate is the number the owner needs (§7.4 rule 3: units using any byte-less face ship as rasters)',
  })
  console.log(`[t4] files ${totals.files} · faces ${totals.faces} · verifiable ${totals.verifiable} · unverifiable ${totals.unverifiable} · units ${totals.unitsTotal} · wouldRaster ${totals.unitsThatWouldRaster} (${rasterFallbackRate}) · with pinned pack ${totals.unitsThatWouldRasterWithPack} (${rasterFallbackRateWithPack})`)
  return 0
}

if (import.meta.main) run().then(
  (code) => process.exit(code),
  (error) => { console.error(error); process.exit(1) },
)
