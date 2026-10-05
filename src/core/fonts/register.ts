import type { EmbeddedFontFace, FontDiagnostic } from './types'
export interface FontRegistrationRequest {
  bytes: Uint8Array
  alias: string
  descriptors: { weight: '400' | '700'; style: 'normal' | 'italic' }
  face: EmbeddedFontFace
}
/** Resolve after the face is ready to measure/paint. Cleanup releases only this face. */
export type RegisterFont = (request: FontRegistrationRequest) => Promise<void | (() => void)>
export type FontResolver = (family: string) => string
interface FontDocument { embeddedFonts?: EmbeddedFontFace[]; fontDiagnostics?: FontDiagnostic[] }
interface Registration {
  refs: number
  ready: Promise<void>
  aliases: Map<string, string>
  diagnostics: FontDiagnostic[]
  cleanups: Array<() => void>
}
const registrations = new WeakMap<FontDocument, Map<RegisterFont, Registration>>()
let sequence = 0
const browserRegister: RegisterFont = async ({ bytes, alias, descriptors }) => {
  const set = globalThis.document?.fonts
  if (typeof FontFace === 'undefined' || !set) throw new Error('FontFace registration unavailable; provide registerFont for headless display')
  const face = new FontFace(alias, bytes.slice().buffer, descriptors)
  await face.load()
  set.add(face)
  return () => { set.delete(face) }
}
export interface FontLease { resolve: FontResolver; diagnostics: FontDiagnostic[]; dispose: () => void }
/** A pending extraction holds a reference too, so another consumer cannot delete its fonts. */
export async function acquireFonts(doc: FontDocument, adapter: RegisterFont = browserRegister): Promise<FontLease> {
  let byAdapter = registrations.get(doc)
  if (!byAdapter) { byAdapter = new Map(); registrations.set(doc, byAdapter) }
  let state = byAdapter.get(adapter)
  if (!state) {
    state = { refs: 0, ready: Promise.resolve(), aliases: new Map(), diagnostics: [], cleanups: [] }
    byAdapter.set(adapter, state)
    const current = state, id = ++sequence, requestedAliases = new Map<string, string>()
    state.ready = Promise.all((doc.embeddedFonts ?? []).map(async face => {
      const key = face.family.toLocaleLowerCase('en-US')
      let alias = requestedAliases.get(key)
      if (!alias) { alias = `OfficeviewFont_${id}_${requestedAliases.size}`; requestedAliases.set(key, alias) }
      const descriptors: FontRegistrationRequest['descriptors'] = {
        weight: face.variant === 'bold' || face.variant === 'boldItalic' ? '700' : '400',
        style: face.variant === 'italic' || face.variant === 'boldItalic' ? 'italic' : 'normal',
      }
      try {
        const cleanup = await adapter({ bytes: face.bytes, alias, descriptors, face })
        if (cleanup) current.cleanups.push(cleanup)
        current.aliases.set(key, alias)
      } catch (error) {
        current.diagnostics.push({ kind: 'font-load-failed', family: face.family, variant: face.variant, partPath: face.partPath, relationshipId: face.relationshipId, message: error instanceof Error ? error.message : String(error) })
      }
    })).then(() => {})
  }
  const current = state
  current.refs++
  let released = false
  const dispose = () => {
    if (released) return
    released = true
    if (--current.refs !== 0) return
    byAdapter!.delete(adapter)
    if (!byAdapter!.size) registrations.delete(doc)
    // Delete registry entries even when a host cleanup fails, and try every face.
    for (const cleanup of current.cleanups.splice(0)) { try { cleanup() } catch { /* host cleanup cannot strand sibling faces */ } }
    current.aliases.clear()
  }
  try {
    await current.ready
    return { resolve: family => current.aliases.get(family.toLocaleLowerCase('en-US')) ?? family, diagnostics: [...(doc.fontDiagnostics ?? []), ...current.diagnostics], dispose }
  } catch (error) { dispose(); throw error }
}
