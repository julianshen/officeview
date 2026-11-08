/** HKDF subkeys from one master (spec §13.4) and the `docRef` (§6.2). */

const enc = new TextEncoder()

export interface Keys {
  kDoc: CryptoKey
  kToken: CryptoKey
  kAudit: Uint8Array
}

async function hkdf(master: Uint8Array, info: string): Promise<Uint8Array> {
  const base = await crypto.subtle.importKey('raw', master as BufferSource, 'HKDF', false, ['deriveBits'])
  const bits = await crypto.subtle.deriveBits({ name: 'HKDF', hash: 'SHA-256', salt: enc.encode('officeview-stream'), info: enc.encode(info) }, base, 256)
  return new Uint8Array(bits)
}

const hmacKey = (raw: Uint8Array) => crypto.subtle.importKey('raw', raw as BufferSource, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign', 'verify'])

export async function deriveKeys(master: Uint8Array): Promise<Keys> {
  const [d, t, a] = await Promise.all([hkdf(master, 'kDoc'), hkdf(master, 'kToken'), hkdf(master, 'kAudit')])
  return { kDoc: await hmacKey(d), kToken: await hmacKey(t), kAudit: a }
}

export const hex = (b: Uint8Array) => Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('')

/** HMAC(kDoc, sha256(bytes)) truncated to 96 bits, hex (24 chars). */
export async function docRefOf(kDoc: CryptoKey, bytes: Uint8Array): Promise<string> {
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes as BufferSource))
  const mac = new Uint8Array(await crypto.subtle.sign('HMAC', kDoc, digest))
  return hex(mac.subarray(0, 12))
}
