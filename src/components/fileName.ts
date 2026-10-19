/**
 * The display name of an opened file (spec §6.5): File.name, a
 * Content-Disposition filename (RFC 6266: filename* over filename) or the
 * last URL path segment. Eager and small (used by <OfficeFile>); kept out of
 * components/paged/ so the eager allowlist there stays metrics.ts only.
 */
const decode = (s: string): string | undefined => { try { return decodeURIComponent(s) } catch { return undefined } }

export function fileNameFrom(source: unknown, disposition?: string | null): string | undefined {
  if (disposition) {
    const star = /filename\*\s*=\s*[^']*'[^']*'([^;\s]+)/i.exec(disposition)
    const name = (star && decode(star[1])) ?? /filename\s*=\s*(?:"([^"]*)"|([^;\s]+))/i.exec(disposition)?.slice(1).find(Boolean)
    if (name) return name
  }
  // File.name: of the typed sources (DirectByteSource), only a File carries a
  // name; any other string `name` can arrive only through an untyped cast.
  const { name, url } = (source ?? {}) as { name?: unknown; url?: unknown }
  if (typeof name === 'string') return name || undefined
  if (typeof url !== 'string') return undefined
  const segment = url.split(/[?#]/)[0].split('/').pop()
  return segment ? decode(segment) ?? segment : undefined
}
