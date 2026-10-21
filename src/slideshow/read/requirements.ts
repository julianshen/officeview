/**
 * Slideshow requirements table (slideshow spec 6.1): the markup-compatibility
 * namespaces the show reader understands, by URI. `mc:Choice@Requires` names
 * prefixes, which are resolved through the in-scope bindings of the Choice
 * itself, so an aliased prefix works and the literal spelling "p14" bound to
 * another URI does not.
 *
 * The eager `supportedChoiceRequirements` (src/drawing/coverage.ts) applies the
 * same rule, but its URI table has no p14/p15/p159 rows and growing it would
 * grow the eager chunk; this lazy copy keeps the URIs here.
 */
import { attrs, getChildren, namespaceUri, orderedChildren, type XmlNode } from '../../core/xml'

/** Namespace URIs by canonical model prefix: every namespace the catalogue names. */
export const SLIDESHOW_NAMESPACES: Readonly<Record<string, readonly string[]>> = Object.freeze({
  p: Object.freeze(['http://schemas.openxmlformats.org/presentationml/2006/main', 'http://purl.oclc.org/ooxml/presentationml/main']),
  p14: Object.freeze(['http://schemas.microsoft.com/office/powerpoint/2010/main']),
  p15: Object.freeze(['http://schemas.microsoft.com/office/powerpoint/2012/main']),
  p159: Object.freeze(['http://schemas.microsoft.com/office/powerpoint/2015/09/main']),
})

export const SLIDESHOW_REQUIREMENTS: readonly string[] = Object.freeze(Object.keys(SLIDESHOW_NAMESPACES))

const supportedUris = new Set(Object.values(SLIDESHOW_NAMESPACES).flat())

/** Every Requires token must resolve, at the Choice, to a supported URI. */
export function slideshowChoiceSupported(choice: XmlNode): { supported: boolean; unknown?: string } {
  for (const prefix of (attrs(choice).Requires ?? '').trim().split(/\s+/).filter(Boolean)) {
    if (!supportedUris.has(namespaceUri(choice, prefix) ?? '')) return { supported: false, unknown: prefix }
  }
  return { supported: true }
}

/**
 * The branches of an `mc:AlternateContent` in preference order: supported
 * Choices (document order), then the Fallback. `rejected` names the first
 * unknown Requires prefix of each unsupported Choice.
 */
export function compatibilityBranches(alternate: XmlNode): { branches: XmlNode[]; rejected: string[] } {
  const branches: XmlNode[] = []
  const rejected: string[] = []
  for (const [name, node] of orderedChildren(alternate)) {
    if (name !== 'Choice') continue
    const check = slideshowChoiceSupported(node)
    if (check.supported) branches.push(node)
    else rejected.push(check.unknown ?? '')
  }
  branches.push(...getChildren(alternate, 'Fallback'))
  return { branches, rejected }
}
