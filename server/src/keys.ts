/** HKDF subkeys from one master (spec §13.4) and the `docRef` (§6.2). */

const enc = new TextEncoder()

export interface Keys {
  kDoc: CryptoKey
  kToken: CryptoKey
  kAudit: Uint8Array
  /** Layer-1 basis root; per-unit keys are deriveBasisKey(root, docRef, `${unitId}@${bucket}`) (§13.4). */
  kBasisRoot: Uint8Array
  /** Layer-1 message MAC key (§13.2). */
  kMsg: Uint8Array
  /** Bound into msg_40; one generation in Stage 2a (rotation is 2b). */
  keyGeneration: number
}

async function hkdf(master: Uint8Array, info: string): Promise<Uint8Array> {
  const base = await crypto.subtle.importKey('raw', master as BufferSource, 'HKDF', false, ['deriveBits'])
  const bits = await crypto.subtle.deriveBits({ name: 'HKDF', hash: 'SHA-256', salt: enc.encode('officeview-stream'), info: enc.encode(info) }, base, 256)
  return new Uint8Array(bits)
}

const hmacKey = (raw: Uint8Array) => crypto.subtle.importKey('raw', raw as BufferSource, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign', 'verify'])

export async function deriveKeys(master: Uint8Array): Promise<Keys> {
  const [d, t, a, b, m] = await Promise.all([hkdf(master, 'kDoc'), hkdf(master, 'kToken'), hkdf(master, 'kAudit'), hkdf(master, 'kBasis'), hkdf(master, 'kMsg')])
  return { kDoc: await hmacKey(d), kToken: await hmacKey(t), kAudit: a, kBasisRoot: b, kMsg: m, keyGeneration: 1 }
}

export const hex = (b: Uint8Array) => Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('')

/** HMAC(kDoc, sha256(bytes)) truncated to 96 bits, hex (24 chars). */
export async function docRefOf(kDoc: CryptoKey, bytes: Uint8Array): Promise<string> {
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes as BufferSource))
  const mac = new Uint8Array(await crypto.subtle.sign('HMAC', kDoc, digest))
  return hex(mac.subarray(0, 12))
}
