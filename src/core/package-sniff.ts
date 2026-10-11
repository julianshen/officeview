/**
 * Package classification for the entry points (design note §4.5).
 *
 * Two levels: a HINT from the first bytes of a download (only a complete
 * `[Content_Types].xml` inside the prefix counts; part names never do), and
 * an AUTHORITATIVE answer from the central directory plus the bounded
 * `[Content_Types].xml`, mirroring the decision the legacy loader's
 * detectAndParse makes. The hint only chooses between today's
 * accumulate-in-memory path and the bounded spool; it never decides which
 * parser runs.
 */
import { createPakoInflater } from './inflater'
import { readPartBytes, type RangeSource, type ZipInventory } from './zip-stream'

export type OfficeKind = 'docx' | 'xlsx' | 'pptx' | 'other'
export type PackageHint = OfficeKind | 'not-zip' | 'undetermined'

const SIG_LOCAL = 0x04034b50
const SIG_DESCRIPTOR = 0x08074b50
const CT_NAME = '[Content_Types].xml'
const CT_MAX_SNIFF_BYTES = 64 * 1024
/** Inflating earlier data-descriptor entries only to find their end: bounded work. */
const SKIP_INFLATE_MAX_OUT = 16 * 1024 * 1024
const latin1 = new TextDecoder('latin1')
const utf8 = new TextDecoder('utf-8')

function u16(b: Uint8Array, at: number): number {
  return b[at] | (b[at + 1] << 8)
}
function u32(b: Uint8Array, at: number): number {
  return (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16)) + b[at + 3] * 0x1000000
}

/**
 * The same substring tests, in the same order, as the legacy detectAndParse
 * (main content types first), with the part-name fallback expressed through
 * the content types' own Override part names for a hint.
 */
export function classifyContentTypes(text: string, has?: (part: string) => boolean): OfficeKind {
  if (text.includes('wordprocessingml.document.main')) return 'docx'
  if (text.includes('spreadsheetml.sheet.main')) return 'xlsx'
  if (text.includes('presentationml.presentation.main')) return 'pptx'
  const present = has ?? ((part: string) => text.includes(`"/${part}"`))
  if (present('word/document.xml')) return 'docx'
  if (present('xl/workbook.xml')) return 'xlsx'
  if (present('ppt/presentation.xml')) return 'pptx'
  return 'other'
}

/** Inflate raw deflate from `data[at..]` within the prefix; null when it does not end inside it. */
function inflateWithin(data: Uint8Array, at: number, maxOut: number, keep: boolean): { consumed: number; out: Uint8Array } | null {
  const inflater = createPakoInflater()
  const buf = new Uint8Array(65_536)
  const parts: Uint8Array[] = []
  let produced = 0
  let pos = at
  try {
    for (;;) {
      const step = inflater.step(data, pos, buf) as { status: string; inConsumed: number; outProduced: number }
      pos += step.inConsumed
      produced += step.outProduced
      if (produced > maxOut) return null
      if (keep && step.outProduced > 0) parts.push(buf.slice(0, step.outProduced))
      if (step.status === 'end') break
      if (step.status === 'need-input') return null
      if (step.inConsumed === 0 && step.outProduced === 0) return null
    }
  } catch {
    return null
  } finally {
    inflater.dispose()
  }
  const out = new Uint8Array(keep ? produced : 0)
  let o = 0
  for (const p of parts) { out.set(p, o); o += p.length }
  return { consumed: pos - at, out }
}

/**
 * Hint from the first bytes of a ZIP download: walk local headers from offset
 * 0 and classify only from a complete `[Content_Types].xml` inside the prefix.
 * Anything that cannot be decided inside the prefix is `undetermined` (spool,
 * then classify authoritatively). RTF is checked by the caller.
 */
export function sniffPackage(prefix: Uint8Array): PackageHint {
  if (prefix.length < 4 || u32(prefix, 0) !== SIG_LOCAL) return 'not-zip'
  let p = 0
  while (p + 30 <= prefix.length && u32(prefix, p) === SIG_LOCAL) {
    const flags = u16(prefix, p + 6)
    const method = u16(prefix, p + 8)
    const csize = u32(prefix, p + 18)
    const nameLength = u16(prefix, p + 26)
    const extraLength = u16(prefix, p + 28)
    const dataStart = p + 30 + nameLength + extraLength
    if (dataStart > prefix.length) return 'undetermined'
    const name = latin1.decode(prefix.subarray(p + 30, p + 30 + nameLength))
    const descriptor = (flags & 0x08) !== 0
    if (csize === 0xffffffff || (flags & 0x01) !== 0 || (method !== 0 && method !== 8)) return 'undetermined'
    if (name === CT_NAME) {
      let text: string | undefined
      if (method === 8) {
        const r = inflateWithin(prefix, dataStart, CT_MAX_SNIFF_BYTES, true)
        if (r) text = utf8.decode(r.out)
      } else if (!descriptor) {
        if (dataStart + csize <= prefix.length) text = utf8.decode(prefix.subarray(dataStart, dataStart + csize))
      } else {
        const s = utf8.decode(prefix.subarray(dataStart, Math.min(prefix.length, dataStart + CT_MAX_SNIFF_BYTES)))
        const end = s.indexOf('</Types>')
        if (end >= 0) text = s.slice(0, end + 8)
      }
      return text === undefined ? 'undetermined' : classifyContentTypes(text)
    }
    if (!descriptor) {
      p = dataStart + csize
      continue
    }
    // An earlier data-descriptor entry: its end is only known by inflating it.
    if (method !== 8) return 'undetermined'
    const r = inflateWithin(prefix, dataStart, SKIP_INFLATE_MAX_OUT, false)
    if (!r) return 'undetermined'
    let q = dataStart + r.consumed
    if (q + 16 > prefix.length) return 'undetermined'
    q += u32(prefix, q) === SIG_DESCRIPTOR ? 16 : 12
    // a ZIP64 descriptor (24 bytes) would not land on a header: undetermined
    if (q + 4 > prefix.length || u32(prefix, q) !== SIG_LOCAL) return 'undetermined'
    p = q
  }
  return 'undetermined'
}

/** Authoritative: central directory + bounded `[Content_Types].xml`, as detectAndParse decides. */
export async function classifyPackage(src: RangeSource, inventory: ZipInventory, signal: AbortSignal, maxContentTypesBytes = 4 * 1024 * 1024): Promise<OfficeKind> {
  const entry = inventory.find(CT_NAME)
  let text = ''
  if (entry) {
    const bytes = await readPartBytes(src, entry, maxContentTypesBytes, { signal, budget: { charge() {} } })
    text = utf8.decode(bytes)
  }
  return classifyContentTypes(text, (part) => inventory.find(part) !== undefined)
}
