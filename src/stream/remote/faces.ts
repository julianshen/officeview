/**
 * §7.4 pinned faces on the client. The host serves the pack (`pack.json` +
 * TTFs) at `baseUrl` (decision D1: the pack is not in the npm package). A
 * face counts as present only when its bytes hash to the manifest's hash
 * (SHA-256 truncated to 96 bits) — the manifest is the authority, pack.json
 * only says where a face's bytes live. Anything else is missing, and the
 * server streams the units that need it as rasters (rule 4).
 */
import type { Manifest } from '../manifest'

type Required = Manifest['recorder']['requiredFaces']
interface PackFace { id: string; file: string; family: string; weight: number; style: string }
export interface FaceDescriptor { weight: number; style: string }
export interface ResolveFacesOptions {
  fetch?: typeof fetch
  /** Registers verified bytes under the family the server pinned. Default: `FontFace` + `document.fonts`. */
  register?: (family: string, bytes: Uint8Array, d: FaceDescriptor) => Promise<void>
}

async function hash96(bytes: Uint8Array): Promise<string> {
  const d = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes as BufferSource))
  let hex = ''
  for (const b of d.subarray(0, 12)) hex += b.toString(16).padStart(2, '0')
  return hex
}

async function registerFontFace(family: string, bytes: Uint8Array, d: FaceDescriptor): Promise<void> {
  const fonts = (globalThis as { document?: { fonts?: { add(f: FontFace): void } } }).document?.fonts
  if (typeof FontFace !== 'function' || !fonts) throw new Error('FontFace unavailable')
  const face = new FontFace(family, bytes as BufferSource, { weight: String(d.weight), style: d.style })
  await face.load()
  fonts.add(face)
}

/** Returns the ids of the required faces that are NOT available, hash-verified. */
export async function resolveFaces(required: Required, baseUrl: string | undefined, o: ResolveFacesOptions = {}): Promise<string[]> {
  const ids = required.map((f) => f.id)
  if (!ids.length) return []
  if (!baseUrl) return ids
  const get = o.fetch ?? fetch, register = o.register ?? registerFontFace
  const base = baseUrl.replace(/\/+$/, '')
  let pack: PackFace[]
  try {
    const r = await get(`${base}/pack.json`)
    if (!r.ok) return ids
    const json = await r.json() as { faces?: unknown }
    pack = Array.isArray(json.faces) ? json.faces as PackFace[] : []
  } catch { return ids }
  const results = await Promise.all(required.map(async (f) => {
    const entry = pack.find((p) => p && p.id === f.id)
    if (!entry || typeof entry.file !== 'string' || !/^[\w.-]+\.ttf$/.test(entry.file) || typeof entry.family !== 'string') return f.id
    try {
      const r = await get(`${base}/${entry.file}`)
      if (!r.ok) return f.id
      const bytes = new Uint8Array(await r.arrayBuffer())
      if (await hash96(bytes) !== f.hash) return f.id
      await register(entry.family, bytes, { weight: Number(entry.weight) || 400, style: entry.style === 'italic' ? 'italic' : 'normal' })
      return undefined
    } catch { return f.id }
  }))
  return results.filter((id): id is string => id !== undefined)
}
