/** DOM-free, bounded SFNT/EOT/MTX conversion. The vendored decoder is isolated. */
import { decompressMtx, parseEotMetadata, withDecoderByteLimit } from './vendor/mtx.mjs'
export const FONT_LIMITS = { faces: 64, inputBytes: 16 * 1024 * 1024, decodedBytes: 32 * 1024 * 1024, documentBytes: 64 * 1024 * 1024 } as const
export class FontDecodeError extends Error {
  constructor(public readonly kind: 'malformed-font' | 'restricted-font' | 'font-limit', message: string) { super(message) }
}
function displayPermissions(fsType: number): void {
  // fsType 4 permits preview/print, 8 permits editable embedding. Both allow
  // this read-only display pipeline; 2 and bitmap-only embedding do not.
  if ((fsType & 2) || (fsType & 0x200)) throw new FontDecodeError('restricted-font', 'Font embedding is restricted')
}
function sfntPermissions(bytes: Uint8Array): void {
  if (bytes.length < 12) throw new Error('Truncated SFNT')
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const signature = view.getUint32(0)
  if (![0x00010000, 0x4f54544f, 0x74727565].includes(signature)) throw new Error('Unsupported font signature')
  const tables = view.getUint16(4)
  if (!tables || tables > 4096 || 12 + tables * 16 > bytes.length) throw new Error('Truncated SFNT directory')
  for (let i = 0; i < tables; i++) {
    const at = 12 + i * 16, offset = view.getUint32(at + 8), length = view.getUint32(at + 12)
    if (offset > bytes.length || length > bytes.length - offset) throw new Error('Truncated SFNT table')
    if (view.getUint32(at) === 0x4f532f32) {
      if (length < 10) throw new Error('Truncated OS/2 permissions')
      displayPermissions(view.getUint16(offset + 8))
    }
  }
}
export function decodeEmbeddedFont(bytes: Uint8Array, maxDecodedBytes: number = FONT_LIMITS.decodedBytes): Uint8Array {
  if (bytes.length > FONT_LIMITS.inputBytes) throw new FontDecodeError('font-limit', 'Font input byte limit exceeded')
  try {
    return withDecoderByteLimit(maxDecodedBytes, () => {
      if (bytes.length < 12) throw new Error('Truncated font header')
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      let decoded: Uint8Array
      if ([0x00010000, 0x4f54544f, 0x74727565].includes(view.getUint32(0))) {
        if (bytes.length > maxDecodedBytes) throw new FontDecodeError('font-limit', 'Decoded font byte budget exceeded')
        decoded = bytes.slice()
      } else if (bytes.length >= 36 && view.getUint16(34, true) === 0x504c) {
        const metadata = parseEotMetadata(bytes)
        displayPermissions(metadata.permissions)
        if (!metadata.compressed && metadata.fontDataSize > maxDecodedBytes) throw new FontDecodeError('font-limit', 'Decoded EOT font byte budget exceeded')
        if (metadata.rootString) throw new FontDecodeError('restricted-font', 'EOT RootString restriction cannot be satisfied by document display')
        decoded = decompressMtx(bytes.subarray(metadata.fontDataOffset, metadata.fontDataOffset + metadata.fontDataSize), { compressed: metadata.compressed, encrypted: metadata.encrypted })
      } else decoded = decompressMtx(bytes)
      if (decoded.length > maxDecodedBytes) throw new FontDecodeError('font-limit', 'Decoded font byte limit exceeded')
      sfntPermissions(decoded)
      return decoded
    })
  } catch (error) {
    if (error instanceof FontDecodeError) throw error
    const message = error instanceof Error ? error.message : String(error)
    throw new FontDecodeError(/budget|limit/i.test(message) ? 'font-limit' : 'malformed-font', message)
  }
}
