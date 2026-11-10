/** Master-key generations (spec §13.4): one 32-byte master per deployment era. Errors never carry key bytes. */
import { readFileSync } from 'node:fs'
import { ConfigError } from './config'
import { deriveKeys, type Keys } from './keys'

export interface Keyring { current: number; keys: ReadonlyMap<number, Uint8Array> }

/** Strict base64 → exactly 32 bytes; `label` names the source in errors (never the value). */
export function decodeMasterKey(b64: string, label: string): Uint8Array {
  let bytes: Uint8Array
  try { bytes = Uint8Array.from(atob(b64.trim()), (c) => c.charCodeAt(0)) } catch { throw new ConfigError('bad-key', `${label} is not base64`) }
  if (bytes.length !== 32) throw new ConfigError('bad-key', `${label} must decode to 32 bytes, got ${bytes.length}`)
  return bytes
}

const isGeneration = (n: number) => Number.isInteger(n) && n >= 1 && n <= 255

export function parseKeyring(json: string): Keyring {
  let doc: unknown
  try { doc = JSON.parse(json) } catch { throw new ConfigError('bad-key', 'keyring is not valid JSON') }
  if (typeof doc !== 'object' || doc === null || Array.isArray(doc)) throw new ConfigError('bad-key', 'keyring must be an object')
  const { current, keys: rawKeys } = doc as { current?: unknown; keys?: unknown }
  if (typeof current !== 'number' || !isGeneration(current)) throw new ConfigError('bad-key', 'keyring "current" must be an integer 1..255')
  if (typeof rawKeys !== 'object' || rawKeys === null || Array.isArray(rawKeys)) throw new ConfigError('bad-key', 'keyring "keys" must be an object')
  const keys = new Map<number, Uint8Array>()
  for (const [name, value] of Object.entries(rawKeys)) {
    const gen = Number(name)
    if (String(gen) !== name || !isGeneration(gen)) throw new ConfigError('bad-key', `keyring generation "${name.slice(0, 16)}" must be an integer 1..255`)
    if (typeof value !== 'string') throw new ConfigError('bad-key', `keyring generation ${gen} must be a base64 string`)
    keys.set(gen, decodeMasterKey(value, `keyring generation ${gen}`))
  }
  if (!keys.has(current)) throw new ConfigError('bad-key', `keyring has no key for current generation ${current}`)
  return { current, keys }
}

export function loadKeyring(env: Record<string, string | undefined>, demo: boolean): { keyring: Keyring; ephemeral: boolean } {
  const path = env.OV_STREAM_KEYRING, single = env.OV_STREAM_KEY
  if (path && single) throw new ConfigError('bad-arg', 'set only one of OV_STREAM_KEYRING and OV_STREAM_KEY')
  if (path) {
    let text: string
    try { text = readFileSync(path, 'utf8') } catch { throw new ConfigError('bad-key', `cannot read OV_STREAM_KEYRING file ${path}`) }
    return { keyring: parseKeyring(text), ephemeral: false }
  }
  if (single) return { keyring: { current: 1, keys: new Map([[1, decodeMasterKey(single, 'OV_STREAM_KEY')]]) }, ephemeral: false }
  if (demo) return { keyring: { current: 1, keys: new Map([[1, crypto.getRandomValues(new Uint8Array(32))]]) }, ephemeral: true }
  throw new ConfigError('missing-key', 'OV_STREAM_KEY or OV_STREAM_KEYRING is required outside --demo')
}

/** Every generation's subkeys, so detection can trace marks from past eras. */
export async function deriveKeyring(ring: Keyring): Promise<Map<number, Keys>> {
  const out = new Map<number, Keys>()
  for (const [gen, master] of ring.keys) out.set(gen, await deriveKeys(master, gen))
  return out
}
