/**
 * Verifies `fontpack/pack.json` against the bytes on disk (spec §7.4 rule 6:
 * the pack is pinned by hash; server and client must hold identical bytes).
 *
 *   bun scripts/stream-s1/fontpack-lock.ts [dir]          # verify, exit 1 on drift
 *   bun scripts/stream-s1/fontpack-lock.ts --write [dir]  # (re)write pack.json
 */
import { createHash } from 'node:crypto'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

export interface PackFace { id: string; file: string; family: string; weight: 400 | 700; style: 'normal' | 'italic'; hash96: string; sha256: string; source: string }
export interface Pack { version: 1; faces: PackFace[]; aliases: Record<string, string>; licences: string[] }

const GOOGLE = 'https://github.com/google/fonts/raw/main/ofl'
const LIBERATION = 'https://github.com/liberationfonts/liberation-fonts/files/7261482/liberation-fonts-ttf-2.1.5.tar.gz'
const FAMILIES: Array<{ family: string; prefix: string; source: (file: string) => string }> = [
  { family: 'Carlito', prefix: 'Carlito', source: (f) => `${GOOGLE}/carlito/${f}` },
  { family: 'Caladea', prefix: 'Caladea', source: (f) => `${GOOGLE}/caladea/${f}` },
  { family: 'Liberation Sans', prefix: 'LiberationSans', source: (f) => `${LIBERATION}#${f}` },
  { family: 'Liberation Serif', prefix: 'LiberationSerif', source: (f) => `${LIBERATION}#${f}` },
  { family: 'Liberation Mono', prefix: 'LiberationMono', source: (f) => `${LIBERATION}#${f}` },
]
const STYLES = [
  { suffix: 'Regular', weight: 400, style: 'normal' },
  { suffix: 'Bold', weight: 700, style: 'normal' },
  { suffix: 'Italic', weight: 400, style: 'italic' },
  { suffix: 'BoldItalic', weight: 700, style: 'italic' },
] as const
export const PACK_ALIAS_TABLE: Record<string, string> = {
  Calibri: 'Carlito', Cambria: 'Caladea', Arial: 'Liberation Sans', Helvetica: 'Liberation Sans',
  'Times New Roman': 'Liberation Serif', 'Courier New': 'Liberation Mono',
}

const sha256 = (b: Uint8Array) => createHash('sha256').update(b).digest('hex')

export function buildPack(dir: string): Pack {
  const faces: PackFace[] = []
  for (const fam of FAMILIES) for (const st of STYLES) {
    const file = `${fam.prefix}-${st.suffix}.ttf`
    const sha = sha256(readFileSync(join(dir, file)))
    faces.push({ id: `${fam.prefix.replace(/([a-z])([A-Z])/g, '$1-$2').toLowerCase()}-${st.suffix.replace(/([a-z])([A-Z])/g, '$1-$2').toLowerCase()}`, file, family: fam.family, weight: st.weight, style: st.style, hash96: sha.slice(0, 24), sha256: sha, source: fam.source(file) })
  }
  return { version: 1, faces, aliases: PACK_ALIAS_TABLE, licences: ['OFL-Carlito.txt', 'OFL-Caladea.txt', 'LICENSE-Liberation'] }
}

/** Problems with `dir`'s pack.json; empty = verified. */
export function verifyPack(dir: string): string[] {
  const at = join(dir, 'pack.json')
  if (!existsSync(at)) return ['pack.json missing']
  const pack = JSON.parse(readFileSync(at, 'utf8')) as Pack
  const out: string[] = []
  for (const f of pack.faces) {
    const p = join(dir, f.file)
    if (!existsSync(p)) { out.push(`${f.file} missing`); continue }
    const sha = sha256(readFileSync(p))
    if (sha !== f.sha256 || sha.slice(0, 24) !== f.hash96) out.push(`${f.file} hash drift`)
  }
  for (const l of pack.licences) if (!existsSync(join(dir, l))) out.push(`${l} missing`)
  return out
}

if (import.meta.main) {
  const args = process.argv.slice(2)
  const write = args.includes('--write')
  const dir = args.find((a) => !a.startsWith('--')) ?? join(process.cwd(), 'fontpack')
  if (write) { writeFileSync(join(dir, 'pack.json'), JSON.stringify(buildPack(dir), null, 2) + '\n'); console.log('[fontpack] wrote pack.json') }
  const problems = verifyPack(dir)
  for (const p of problems) console.error(`[fontpack] ${p}`)
  console.log(problems.length ? '[fontpack] DRIFT' : '[fontpack] ok')
  process.exit(problems.length ? 1 : 0)
}
