/** Word font-table declarations and bounded, permission-checked font loading. */
import type { OfficePackage } from '../zip'
import { attrs, getChildren, parseXml, type XmlNode } from '../xml'
import { decodeEmbeddedFont, FONT_LIMITS, FontDecodeError } from './decode'
import type { EmbeddedFontFace, FontDiagnostic, FontVariant } from './types'

export interface DocxFontDeclaration {
  family: string
  variant: FontVariant
  relationshipId: string
  fontKey?: string
  subsetted?: boolean
  partPath?: string
  isExternal?: boolean
}

/** Compatibility name retained for existing Word font consumers. */
export type DocxEmbeddedFontDeclaration = DocxFontDeclaration

function internalPath(owner: string, target: string): string | undefined {
  if (!target || /[?#\\]|^[a-z][a-z0-9+.-]*:/i.test(target)) return undefined
  const segments = target.startsWith('/') ? [] : owner.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') { if (!segments.length) return undefined; segments.pop() }
    else if (segment && segment !== '.') segments.push(segment)
  }
  return segments.join('/') || undefined
}

export function parseDocxFontDeclarations(table: XmlNode | undefined, rels?: XmlNode, owner = 'word/fontTable.xml', maxDeclarations: number = FONT_LIMITS.faces): DocxFontDeclaration[] {
  const declarations: DocxFontDeclaration[] = []
  const limit = Number.isSafeInteger(maxDeclarations) && maxDeclarations > 0 ? maxDeclarations : 0
  const relationships = new Map(getChildren(rels, 'Relationship').map(rel => [attrs(rel).Id, attrs(rel)]))
  for (const font of getChildren(table, 'font')) {
    const family = attrs(font).name ?? ''
    for (const [tag, variant] of [['embedRegular', 'regular'], ['embedBold', 'bold'], ['embedItalic', 'italic'], ['embedBoldItalic', 'boldItalic']] as const) {
      for (const face of getChildren(font, tag)) {
        if (declarations.length >= limit) return declarations
        const a = attrs(face), relationshipId = a.id ?? '', rel = relationships.get(relationshipId)
        const isExternal = rel?.TargetMode === 'External'
        const partPath = rel && !isExternal && rel.Type?.endsWith('/font') ? internalPath(owner, rel.Target) : undefined
        declarations.push({ family, variant, relationshipId, ...(a.fontKey !== undefined ? { fontKey: a.fontKey } : {}), ...(a.subsetted === '1' || a.subsetted === 'true' ? { subsetted: true } : {}), ...(partPath ? { partPath } : {}), ...(isExternal ? { isExternal: true } : {}) })
      }
    }
  }
  return declarations
}

/** Reverse GUID bytes for the ECMA-376 embedded-font XOR key. */
export function parseGuidReversedKey(fontKey: string): Uint8Array {
  if (!/^(?:\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/i.test(fontKey)) throw new FontDecodeError('malformed-font', 'Invalid fontKey GUID')
  const hex = fontKey.replace(/[{}-]/g, '')
  return Uint8Array.from({ length: 16 }, (_, i) => parseInt(hex.slice((15 - i) * 2, (16 - i) * 2), 16))
}

/** ECMA-376 font embedding: reverse GUID bytes, then XOR the first 32 bytes. */
export function deobfuscateOdttf(source: Uint8Array, fontKey: string): Uint8Array {
  const key = parseGuidReversedKey(fontKey)
  if (source.length < 32) throw new FontDecodeError('malformed-font', 'Truncated obfuscated font header')
  const bytes = source.slice()
  for (let i = 0; i < 32; i++) bytes[i] ^= key[i % 16]
  return bytes
}

export async function loadDocxEmbeddedFonts(pkg: OfficePackage, fontTablePath = 'word/fontTable.xml'): Promise<{ embeddedFonts: EmbeddedFontFace[]; fontDiagnostics: FontDiagnostic[] }> {
  const embeddedFonts: EmbeddedFontFace[] = [], fontDiagnostics: FontDiagnostic[] = []
  let owner = fontTablePath
  // Bound XML inflation as well as declaration enumeration. Invalid entries
  // receive diagnostics without consuming the valid face budget.
  const readXml = async (path: string) => {
    const bytes = await pkg.bytes(path, FONT_LIMITS.inputBytes)
    return bytes ? parseXml(new TextDecoder().decode(bytes)) : undefined
  }
  let table: XmlNode | undefined, rels: XmlNode | undefined
  try {
    // Word may place its font table at a relationship-selected custom path.
    // Apply the same bounded XML read and internal-target guard there too.
    if (!pkg.has(owner)) {
      const documentRels = await readXml('word/_rels/document.xml.rels')
      for (const rel of getChildren(documentRels, 'Relationship')) {
        const a = attrs(rel)
        if (!a.Type?.endsWith('/fontTable') || a.TargetMode === 'External') continue
        const path = internalPath('word/document.xml', a.Target ?? '')
        if (path && pkg.has(path)) { owner = path; break }
      }
    }
    table = await readXml(owner)
    if (table) {
      const slash = owner.lastIndexOf('/')
      const relsPath = `${slash < 0 ? '' : owner.slice(0, slash) + '/'}_rels/${owner.slice(slash + 1)}.rels`
      rels = await readXml(relsPath)
    }
  }
  catch (error) {
    fontDiagnostics.push({ kind: /limit|budget/i.test(String(error)) ? 'font-limit' : 'malformed-font', family: '', variant: 'regular', partPath: owner, message: String(error) })
    return { embeddedFonts, fontDiagnostics }
  }
  if (!table) return { embeddedFonts, fontDiagnostics }
  const enumerationLimit = FONT_LIMITS.faces * 4
  const declarations = parseDocxFontDeclarations(table, rels, owner, enumerationLimit + 1)
  let converted = 0
  for (const [index, declaration] of declarations.entries()) {
    const { family, variant, relationshipId, partPath } = declaration
    const issue = (kind: FontDiagnostic['kind'], message: string) => fontDiagnostics.push({ family, variant, relationshipId, ...(partPath ? { partPath } : {}), kind, message })
    if (index >= enumerationLimit) { issue('font-limit', 'Embedded font declaration limit exceeded'); break }
    if (!family || !relationshipId) { issue('malformed-font', 'Font family or relationship identifier is missing'); continue }
    if (embeddedFonts.length >= FONT_LIMITS.faces) { issue('font-limit', 'Embedded font face limit exceeded'); break }
    if (declaration.isExternal) { issue('external-font', 'External font relationship is not loaded'); continue }
    if (!partPath || !pkg.has(partPath)) { issue('missing-font', 'Embedded font part is missing or invalid'); continue }
    try {
      const source = await pkg.bytes(partPath, FONT_LIMITS.inputBytes)
      if (!source) { issue('missing-font', 'Embedded font part is missing'); continue }
      const obfuscated = /\.odttf$/i.test(partPath) || declaration.fontKey !== undefined
      const bytes = decodeEmbeddedFont(obfuscated ? deobfuscateOdttf(source, declaration.fontKey ?? '') : source, Math.min(FONT_LIMITS.decodedBytes, FONT_LIMITS.documentBytes - converted))
      converted += bytes.length
      embeddedFonts.push({ family, variant, relationshipId, partPath, bytes })
    } catch (error) { issue(error instanceof FontDecodeError ? error.kind : /limit|budget/i.test(String(error)) ? 'font-limit' : 'malformed-font', error instanceof Error ? error.message : String(error)) }
  }
  return { embeddedFonts, fontDiagnostics }
}
