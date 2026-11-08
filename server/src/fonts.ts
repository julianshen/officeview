/**
 * The pinned open font pack (spec §7.4 rule 6) and `requiredFaces`.
 * Faces are registered under their own family and every metric-compatible
 * alias, so the painter's `font = '12pt Calibri'` lays out with Carlito.
 */
import { readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { OpReader, Opcode } from '../../src/stream/ops'
import type { Pack, PackFace } from '../../scripts/stream-s1/fontpack-lock'

export const PACK_DIR = resolve(import.meta.dir, '../../fontpack')

const packs = new Map<string, Pack>()
export function loadPack(dir = PACK_DIR): Pack {
  let p = packs.get(dir)
  if (!p) { p = JSON.parse(readFileSync(join(dir, 'pack.json'), 'utf8')) as Pack; packs.set(dir, p) }
  return p
}

interface FontRegistry { registerFromPath(path: string, alias?: string): unknown }

/** Registers every face under its family and each alias (idempotent per registry). */
const registered = new WeakSet<object>()
export function registerPack(fonts: FontRegistry, pack = loadPack(), dir = PACK_DIR): void {
  if (registered.has(fonts)) return
  for (const face of pack.faces) {
    const path = join(dir, face.file)
    fonts.registerFromPath(path, face.family)
    for (const [alias, target] of Object.entries(pack.aliases)) if (target === face.family) fonts.registerFromPath(path, alias)
  }
  registered.add(fonts)
}

export interface ParsedFont { families: string[]; weight: 400 | 700; italic: boolean }

/** Minimal CSS `font` shorthand parse: [style] [variant] [weight] [stretch] size[/lh] families. */
export function parseFont(font: string): ParsedFont | undefined {
  const m = /^\s*((?:[^\s"',]+\s+)*?)([\d.]+(?:px|pt|em|rem|%)(?:\/[^\s]+)?)\s+(.+)$/i.exec(font)
  if (!m) return undefined
  const pre = m[1]!.trim().toLowerCase().split(/\s+/).filter(Boolean)
  let weight: 400 | 700 = 400, italic = false
  for (const t of pre) {
    if (t === 'italic' || t === 'oblique') italic = true
    else if (t === 'bold' || t === 'bolder' || (/^\d+$/.test(t) && Number(t) >= 600)) weight = 700
  }
  const families = m[3]!.split(',').map((s) => s.trim().replace(/^["']|["']$/g, '')).filter(Boolean)
  return { families, weight, italic }
}

/** The pack face a `font` string resolves to, or undefined when its first
 * family is not a pack family/alias (fail-closed: no system-font fallback
 * is ever counted as pinned). */
export function faceFor(font: string, pack: Pack): PackFace | undefined {
  const p = parseFont(font)
  const first = p?.families[0]?.toLowerCase()
  if (!p || !first) return undefined
  const alias = Object.entries(pack.aliases).find(([a]) => a.toLowerCase() === first)?.[1]
  const family = alias ?? pack.faces.find((f) => f.family.toLowerCase() === first)?.family
  if (!family) return undefined
  return pack.faces.find((f) => f.family === family && f.weight === p.weight && f.style === (p.italic ? 'italic' : 'normal'))
}

export interface FaceUse { faces: string[]; unpinned: string[] }

/** Face ids every GLYPH_RUN in `ops` paints with, and the font strings that
 * resolve to no pack face (those make the unit not command-eligible). */
export function facesUsed(ops: Uint8Array, pack = loadPack()): FaceUse {
  const faces = new Set<string>(), unpinned = new Set<string>()
  const stack: string[] = []
  let font = '10px sans-serif'
  for (const op of new OpReader(ops)) {
    if (op.code === Opcode.STATE_PUSH) stack.push(font)
    else if (op.code === Opcode.STATE_POP) font = stack.pop() ?? font
    else if (op.code === Opcode.STATE_SETUP && op.fields.font !== undefined) font = op.fields.font
    else if (op.code === Opcode.GLYPH_RUN) {
      const face = faceFor(font, pack)
      if (face) faces.add(face.id); else unpinned.add(font)
    }
  }
  return { faces: [...faces].sort(), unpinned: [...unpinned].sort() }
}

/** Manifest `recorder.requiredFaces` (§6.2): identity only, never bytes. */
export function requiredFacesOf(faceIds: Iterable<string>, pack = loadPack()): Array<{ id: string; hash: string; alias?: string }> {
  const ids = new Set(faceIds)
  return pack.faces.filter((f) => ids.has(f.id)).map((f) => {
    const aliases = Object.entries(pack.aliases).filter(([, t]) => t === f.family).map(([a]) => a)
    return { id: f.id, hash: f.hash96, ...(aliases.length ? { alias: aliases.join(',') } : {}) }
  })
}

/** §7.4 rule 4: a unit using any face the client reported missing (or any
 * unpinned font) streams as a raster. */
export function unitNeedsRaster(use: FaceUse, missingFaces: readonly string[]): boolean {
  return use.unpinned.length > 0 || use.faces.some((f) => missingFaces.includes(f))
}

/** Rewrites alias families in a `font` string to their pack family, quoted
 * (`bold 12pt Calibri` → `bold 12pt "Carlito"`). An installed system face of
 * the alias name (Arial, Times New Roman on macOS) would otherwise win over
 * the alias registration, so pinning happens on the name, not the registry. */
export function pinFontString(font: string, pack: Pack): string {
  const m = /^(\s*(?:[^\s"',]+\s+)*?[\d.]+(?:px|pt|em|rem|%)(?:\/[^\s]+)?\s+)(.+)$/i.exec(font)
  if (!m) return font
  let changed = false
  const families = m[2]!.split(',').map((raw) => {
    const name = raw.trim().replace(/^["']|["']$/g, '')
    const target = Object.entries(pack.aliases).find(([a]) => a.toLowerCase() === name.toLowerCase())?.[1]
    if (!target) return raw.trim()
    changed = true
    return `"${target}"`
  })
  return changed ? `${m[1]}${families.join(', ')}` : font
}

/** A 2-D context view whose `font` setter pins alias families. Methods are
 * bound once per name; every other property passes through. */
export function pinFonts<T extends object>(ctx: T, pack: Pack): T {
  const bound = new Map<PropertyKey, unknown>()
  return new Proxy(ctx, {
    get(target, key) {
      const v = Reflect.get(target, key, target)
      if (typeof v !== 'function') return v
      let b = bound.get(key)
      if (!b) { b = (v as (...a: unknown[]) => unknown).bind(target); bound.set(key, b) }
      return b
    },
    set(target, key, value) {
      return Reflect.set(target, key, key === 'font' && typeof value === 'string' ? pinFontString(value, pack) : value, target)
    },
  })
}
