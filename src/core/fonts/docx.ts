/**
 * DOCX (WordprocessingML) embedded font declaration parser and utilities.
 * Handles fontTable.xml and font relationship parsing.
 */
import type { OfficePackage } from '../zip'
import { attrs, getChildren, type XmlNode } from '../xml'
import { decodeEmbeddedFont, FONT_LIMITS, FontDecodeError } from './decode'
import type { EmbeddedFontFace, FontDiagnostic, FontVariant } from './types'

export interface DocxEmbeddedFontDeclaration {
  family: string
  variant: FontVariant
  relationshipId: string
  fontKey?: string
  subsetted?: boolean
  partPath?: string
  isExternal?: boolean
}

function internalPath(part: string, target: string): string | undefined {
  if (!target || /[?#\\]|^[a-z][a-z0-9+.-]*:/i.test(target)) return undefined
  const segments = target.startsWith('/') ? [] : part.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') {
      if (!segments.length) return undefined
      segments.pop()
    } else if (segment && segment !== '.') {
      segments.push(segment)
    }
  }
  return segments.join('/') || undefined
}

const EMBED_VARIANT_MAP: Record<string, FontVariant> = {
  embedRegular: 'regular',
  embedBold: 'bold',
  embedItalic: 'italic',
  embedBoldItalic: 'boldItalic',
}

export function parseDocxFontDeclarations(
  fontTableNode: XmlNode | undefined,
  relsNode?: XmlNode,
  fontTablePath = 'word/fontTable.xml',
  maxDeclarations = FONT_LIMITS.faces * 2
): DocxEmbeddedFontDeclaration[] {
  if (!fontTableNode) return []
  const declarations: DocxEmbeddedFontDeclaration[] = []

  const relationships = new Map<string, { target: string; isExternal: boolean }>()
  if (relsNode) {
    for (const rel of getChildren(relsNode, 'Relationship')) {
      const a = attrs(rel)
      if (a.Id) {
        relationships.set(a.Id, {
          target: a.Target ?? '',
          isExternal: a.TargetMode === 'External',
        })
      }
    }
  }

  // <w:font w:name="..."> or <font name="...">
  const fontNodes = getChildren(fontTableNode, 'font')
  for (const fontNode of fontNodes) {
    if (declarations.length >= maxDeclarations) break
    const fontAttrs = attrs(fontNode)
    const family = fontAttrs.name ?? ''

    for (const [tagName, variant] of Object.entries(EMBED_VARIANT_MAP)) {
      if (declarations.length >= maxDeclarations) break
      const embedNodes = getChildren(fontNode, tagName)
      for (const embedNode of embedNodes) {
        if (declarations.length >= maxDeclarations) break
        const a = attrs(embedNode)
        const relationshipId = a.id ?? ''
        if (!relationshipId && !family) continue

        const fontKey = a.fontKey
        const subsetted = a.subsetted === '1' || a.subsetted === 'true'

        const rel = relationships.get(relationshipId)
        const isExternal = rel?.isExternal === true
        let partPath: string | undefined
        if (rel && !isExternal) {
          partPath = internalPath(fontTablePath, rel.target)
        }

        declarations.push({
          family,
          variant,
          relationshipId,
          ...(fontKey ? { fontKey } : {}),
          ...(subsetted ? { subsetted } : {}),
          ...(partPath ? { partPath } : {}),
          ...(isExternal ? { isExternal: true } : {}),
        })
      }
    }
  }

  return declarations
}

/**
 * Converts a standard GUID string (e.g. "{12345678-ABCD-EF01-2345-6789ABCDEF01}")
 * into the 16-byte reversed key used in ECMA-376 Part 2 / XPS font obfuscation.
 */
export function parseGuidReversedKey(guid: string): Uint8Array {
  const clean = guid.replace(/[{}-]/g, '').trim()
  if (clean.length !== 32 || !/^[0-9a-fA-F]{32}$/.test(clean)) {
    throw new Error(`Invalid GUID font key "${guid}"`)
  }
  const key = new Uint8Array(16)
  for (let i = 0; i < 16; i++) {
    // Reversed byte order: byte 0 is from the last hex pair (index 15)
    key[i] = parseInt(clean.slice((15 - i) * 2, (15 - i) * 2 + 2), 16)
  }
  return key
}

/**
 * Deobfuscates an ODTTF (Obfuscated OpenType Font) stream using the fontKey GUID.
 * Per ECMA-376 Part 2 / XPS font obfuscation:
 * - The first 32 bytes are XOR'd in a single pass with the 16-byte reversed GUID key cycling modulo 16.
 * - All remaining bytes (offset >= 32) remain unmodified.
 */
export function deobfuscateOdttf(bytes: Uint8Array, fontKey: string): Uint8Array {
  const key = parseGuidReversedKey(fontKey)
  const result = bytes.slice()
  const len = Math.min(32, result.length)
  for (let i = 0; i < len; i++) {
    result[i] ^= key[i % 16]
  }
  return result
}

/**
 * Loads, deobfuscates (ODTTF), and decodes embedded fonts declared in a DOCX package's fontTable.xml.
 */
export async function loadDocxEmbeddedFonts(
  pkg: OfficePackage,
  fontTablePath = 'word/fontTable.xml'
): Promise<{ embeddedFonts: EmbeddedFontFace[]; fontDiagnostics: FontDiagnostic[] }> {
  const embeddedFonts: EmbeddedFontFace[] = []
  const fontDiagnostics: FontDiagnostic[] = []

  let actualPath = fontTablePath
  if (!pkg.has(actualPath)) {
    if (pkg.has('word/_rels/document.xml.rels')) {
      try {
        const docRels = await pkg.xml('word/_rels/document.xml.rels')
        for (const rel of getChildren(docRels, 'Relationship')) {
          const a = attrs(rel)
          if (a.Type?.endsWith('/fontTable') && a.Target) {
            const resolved = internalPath('word/document.xml', a.Target)
            if (resolved && pkg.has(resolved)) {
              actualPath = resolved
              break
            }
          }
        }
      } catch {
        // Ignore malformed doc rels
      }
    }
  }

  if (!pkg.has(actualPath)) {
    return { embeddedFonts, fontDiagnostics }
  }

  let fontTableNode: XmlNode | undefined
  try {
    fontTableNode = await pkg.xml(actualPath)
  } catch {
    return { embeddedFonts, fontDiagnostics }
  }
  if (!fontTableNode) {
    return { embeddedFonts, fontDiagnostics }
  }

  const slash = actualPath.lastIndexOf('/')
  const relsPath = `${slash < 0 ? '' : actualPath.slice(0, slash) + '/'}_rels/${actualPath.slice(slash + 1)}.rels`
  let relsNode: XmlNode | undefined
  if (pkg.has(relsPath)) {
    try {
      relsNode = await pkg.xml(relsPath)
    } catch {
      // Ignored
    }
  }

  const declarations = parseDocxFontDeclarations(fontTableNode, relsNode, actualPath)

  let faces = 0
  let converted = 0

  for (const decl of declarations) {
    const identity = {
      family: decl.family,
      variant: decl.variant,
      relationshipId: decl.relationshipId,
    }
    const issue = (kind: FontDiagnostic['kind'], message: string, partPath?: string) => {
      fontDiagnostics.push({
        ...identity,
        kind,
        message,
        ...(partPath ? { partPath } : {}),
      })
    }

    if (!decl.family || !decl.relationshipId) {
      issue('malformed-font', 'Font family or relationship identifier is missing')
      continue
    }

    if (decl.isExternal) {
      issue('external-font', 'External embedded font relationship is not loaded')
      continue
    }

    if (!decl.partPath || !pkg.has(decl.partPath)) {
      issue('missing-font', 'Embedded font part is missing or invalid', decl.partPath)
      continue
    }

    if (++faces > FONT_LIMITS.faces) {
      issue('font-limit', 'Embedded font face limit exceeded', decl.partPath)
      return { embeddedFonts, fontDiagnostics }
    }

    try {
      const source = await pkg.bytes(decl.partPath, FONT_LIMITS.inputBytes)
      if (!source) {
        issue('missing-font', 'Embedded font part is missing', decl.partPath)
        continue
      }

      let fontBytes = source
      if (decl.fontKey) {
        fontBytes = deobfuscateOdttf(source, decl.fontKey)
      }

      const bytes = decodeEmbeddedFont(
        fontBytes,
        Math.min(FONT_LIMITS.decodedBytes, FONT_LIMITS.documentBytes - converted)
      )

      // Defense-in-depth budget enforcement (decodeEmbeddedFont already limits to documentBytes - converted)
      if (converted + bytes.length > FONT_LIMITS.documentBytes) {
        issue('font-limit', 'Document converted font byte limit exceeded', decl.partPath)
        continue
      }

      converted += bytes.length
      embeddedFonts.push({
        family: decl.family,
        variant: decl.variant,
        relationshipId: decl.relationshipId,
        partPath: decl.partPath,
        bytes,
      })
    } catch (error) {
      issue(
        error instanceof FontDecodeError
          ? error.kind
          : /limit|budget/i.test(String(error))
          ? 'font-limit'
          : 'malformed-font',
        error instanceof Error ? error.message : String(error),
        decl.partPath
      )
    }
  }

  return { embeddedFonts, fontDiagnostics }
}


