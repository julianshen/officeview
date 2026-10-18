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
  if (typeof File !== 'undefined' && source instanceof File) return source.name || undefined
  const url = (source as { url?: unknown } | null)?.url
  if (typeof url !== 'string') return undefined
  const segment = url.split(/[?#]/)[0].split('/').pop()
  return segment ? decode(segment) ?? segment : undefined
}
