/**
 * ODF container sniffing, shared by every ODF application (Writer first;
 * Calc/Impress reuse this when they land). An ODF package is a zip whose
 * first entry is an uncompressed `mimetype` file; the manifest is the
 * fallback for forged packages that lack it.
 */
import type { OfficePackage } from '../core/zip'
import { attrs, getChildren, parseXml } from '../core/xml'

export type OdfKind = 'text' | 'spreadsheet' | 'presentation' | 'graphics' | 'formula' | 'chart'

const KIND_BY_MIME: Array<[string, OdfKind]> = [
  ['vnd.oasis.opendocument.text', 'text'],
  ['vnd.oasis.opendocument.spreadsheet', 'spreadsheet'],
  ['vnd.oasis.opendocument.presentation', 'presentation'],
  ['vnd.oasis.opendocument.graphics', 'graphics'],
  ['vnd.oasis.opendocument.formula', 'formula'],
  ['vnd.oasis.opendocument.chart', 'chart'],
]

function kindOf(mime: string): OdfKind | undefined {
  const normalized = mime.trim().toLowerCase()
  for (const [fragment, kind] of KIND_BY_MIME) {
    if (normalized.includes(fragment)) return kind
  }
  return undefined
}

/** Caps for the sniffing reads, which run before any format is known: a real
 * mimetype is a few dozen bytes; the manifest shares the content-types cap. */
export const MIMETYPE_MAX_BYTES = 1024
export const MANIFEST_MAX_BYTES = 4 * 1024 * 1024

/** The ODF application kind of a package, or undefined when it is not ODF.
 * Rejects with PackageError 'resource-limit' when a sniffed part exceeds its cap. */
export async function odfKind(pkg: OfficePackage): Promise<OdfKind | undefined> {
  const mimetype = await pkg.boundedText('mimetype', MIMETYPE_MAX_BYTES)
  if (mimetype) {
    const kind = kindOf(mimetype)
    if (kind) return kind
  }
  // forged packages sometimes lack the mimetype entry — the manifest records
  // the root document's media type under full-path "/"
  const manifestText = await pkg.boundedText('META-INF/manifest.xml', MANIFEST_MAX_BYTES)
  const manifest = manifestText === undefined ? undefined : parseXml(manifestText)
  if (manifest) {
    for (const entry of getChildren(manifest, 'file-entry')) {
      const a = attrs(entry)
      if ((a['full-path'] as string | undefined) === '/') {
        const kind = kindOf((a['media-type'] as string | undefined) ?? '')
        if (kind) return kind
      }
    }
  }
  return undefined
}
