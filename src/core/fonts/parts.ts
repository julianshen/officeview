import type { OfficePackage } from '../zip'
import { attrs, getChildren, type XmlNode } from '../xml'
import { decodeEmbeddedFont, FONT_LIMITS, FontDecodeError } from './decode'
import type { EmbeddedFontFace, FontDiagnostic, FontVariant } from './types'
function internalPath(part: string, target: string): string | undefined {
  if (!target || /[?#\\]|^[a-z][a-z0-9+.-]*:/i.test(target)) return undefined
  const segments = target.startsWith('/') ? [] : part.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') { if (!segments.length) return undefined; segments.pop() }
    else if (segment && segment !== '.') segments.push(segment)
  }
  return segments.join('/') || undefined
}
export async function parseEmbeddedFonts(pkg: OfficePackage, presentation: XmlNode, rels?: XmlNode): Promise<{ embeddedFonts: EmbeddedFontFace[]; fontDiagnostics: FontDiagnostic[] }> {
  const embeddedFonts: EmbeddedFontFace[] = [], fontDiagnostics: FontDiagnostic[] = []
  const relationships = new Map(getChildren(rels, 'Relationship').map(rel => [attrs(rel).Id, attrs(rel)]))
  let faces = 0, converted = 0
  for (const entry of getChildren(getChildren(presentation, 'embeddedFontLst')[0], 'embeddedFont')) {
    const family = attrs(getChildren(entry, 'font')[0]).typeface ?? ''
    for (const variant of ['regular', 'bold', 'italic', 'boldItalic'] as FontVariant[]) {
      for (const face of getChildren(entry, variant)) {
        const relationshipId = attrs(face).id ?? attrs(face)['r:id']
        const identity = { family, variant, relationshipId }
        const issue = (kind: FontDiagnostic['kind'], message: string, partPath?: string) => fontDiagnostics.push({ ...identity, kind, message, ...(partPath ? { partPath } : {}) })
        if (++faces > FONT_LIMITS.faces) { issue('font-limit', 'Embedded font face limit exceeded'); return { embeddedFonts, fontDiagnostics } }
        if (!family || !relationshipId) { issue('malformed-font', 'Font family or relationship identifier is missing'); continue }
        const rel = relationships.get(relationshipId)
        if (rel?.TargetMode === 'External') { issue('external-font', 'External embedded font relationship is not loaded'); continue }
        const partPath = rel && internalPath('ppt/presentation.xml', rel.Target)
        if (!partPath || !pkg.has(partPath)) { issue('missing-font', 'Embedded font part is missing or invalid', partPath); continue }
        try {
          const source = await pkg.bytes(partPath, FONT_LIMITS.inputBytes)
          if (!source) { issue('missing-font', 'Embedded font part is missing', partPath); continue }
          const bytes = decodeEmbeddedFont(source, Math.min(FONT_LIMITS.decodedBytes, FONT_LIMITS.documentBytes - converted))
          if (converted + bytes.length > FONT_LIMITS.documentBytes) { issue('font-limit', 'Document converted font byte limit exceeded', partPath); continue }
          converted += bytes.length
          embeddedFonts.push({ ...identity, relationshipId, partPath, bytes })
        } catch (error) { issue(error instanceof FontDecodeError ? error.kind : /limit|budget/i.test(String(error)) ? 'font-limit' : 'malformed-font', error instanceof Error ? error.message : String(error), partPath) }
      }
    }
  }
  return { embeddedFonts, fontDiagnostics }
}
