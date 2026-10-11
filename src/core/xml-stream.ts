/**
 * Incremental SAX-style XML tokenizer for large OPC parts (design note §3.3, §4.2).
 *
 * Results never depend on where the input was cut: markup is processed only
 * as complete tokens (an incomplete one is carried, capped at maxTokenChars),
 * a reference is decided from at most `&` + 33 characters, and a CR at the
 * end of a chunk is held so CR|LF collapses to one LF. Text is delivered in
 * segments of at most maxTextSegmentChars; segmentation itself is the only
 * chunk-dependent output. Strings stay UTF-16 code units throughout: nothing
 * here re-encodes, so lone surrogates survive. ST_Xstring `_xHHHH_` escapes
 * are a SpreadsheetML consumer concern (XstringDecoder), not XML.
 *
 * Deliberate differences from the fast-xml-parser path (src/core/xml.ts):
 * DOCTYPE is refused (`unsupported`); `&` must start a reference of at most
 * 32 characters ending in `;` in text and attribute values (so `&#;`, `&#x;`
 * and over-long references are errors, while unknown `&name;` stays literal).
 */
import { PackageError, throwIfAborted } from './package-error'

export interface XmlLimits { maxTokenChars: number; maxTextSegmentChars: number; maxDepth: number; maxAttributes: number }
export const DEFAULT_XML_LIMITS: Readonly<XmlLimits> = { maxTokenChars: 1 << 20, maxTextSegmentChars: 262_144, maxDepth: 256, maxAttributes: 128 }

export interface XmlEvents {
  /** Attribute arrays are reused: valid only inside the callback. Self-closing elements get no close(). */
  open(qname: string, attrNames: readonly string[], attrValues: readonly string[], count: number, selfClosing: boolean): void
  close(qname: string): void
  /** Segmented; `more` is false when the run ended at an element boundary. References and CRLF resolved; _xHHHH_ NOT applied. */
  text(text: string, more: boolean): void
}
export interface XmlTokenizerOptions { signal?: AbortSignal }

const LT = 60
const GT = 62
const AMP = 38
const SEMI = 59
const CR = 13
const LF = 10
const TAB = 9
const SP = 32
const SLASH = 47
const BANG = 33
const QM = 63
const EQ = 61
const DQ = 34
const SQ = 39
const HASH = 35
const MAX_REFERENCE_BODY = 32

function isWs(c: number): boolean {
  return c === SP || c === LF || c === TAB || c === CR
}
function isNameStart(c: number): boolean {
  return (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c === 95 || c === 58 || c >= 0x80
}
function isNameChar(c: number): boolean {
  return isNameStart(c) || (c >= 48 && c <= 57) || c === 45 || c === 46
}
function isRefChar(c: number): boolean {
  return (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || (c >= 48 && c <= 57) || c === HASH
}
function isXmlChar(cp: number): boolean {
  return cp === 0x9 || cp === 0xa || cp === 0xd || (cp >= 0x20 && cp <= 0xd7ff) || (cp >= 0xe000 && cp <= 0xfffd) || (cp >= 0x10000 && cp <= 0x10ffff)
}

function parseError(message: string): never {
  throw new PackageError('parse-error', message)
}

const PREDEFINED: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }

/** Decode a complete reference body (between & and ;). */
function decodeReference(body: string): string {
  if (body.charCodeAt(0) === HASH) {
    const hex = body.charCodeAt(1) === 120 /* x */
    const digits = body.slice(hex ? 2 : 1)
    if (digits.length === 0 || !(hex ? /^[0-9A-Fa-f]+$/ : /^[0-9]+$/).test(digits)) parseError(`malformed character reference &${body};`)
    const cp = parseInt(digits, hex ? 16 : 10)
    if (!isXmlChar(cp)) parseError(`character reference &${body}; is not an XML character`)
    return String.fromCodePoint(cp)
  }
  const known = PREDEFINED[body]
  if (known !== undefined) return known
  if (body.includes('#')) parseError(`malformed entity reference &${body};`)
  // unknown entities stay literal, as in the current parser
  return `&${body};`
}

const INCOMPLETE = -1
let lastReference = ''

/** Scan a reference at s[pos] === '&'. Returns the next position (decoded text
 * in lastReference) or INCOMPLETE when s ends before the decision. */
function scanReference(s: string, pos: number): number {
  const len = s.length
  const limit = Math.min(len, pos + 2 + MAX_REFERENCE_BODY)
  let j = pos + 1
  while (j < limit && isRefChar(s.charCodeAt(j))) j++
  const body = j - pos - 1
  if (body > MAX_REFERENCE_BODY) parseError(`reference longer than ${MAX_REFERENCE_BODY} characters`)
  if (j === len) return INCOMPLETE
  if (s.charCodeAt(j) !== SEMI) parseError(body === 0 ? "bare '&' is not a reference" : 'malformed reference: missing ;')
  if (body === 0) parseError('empty reference')
  lastReference = decodeReference(s.slice(pos + 1, j))
  return j + 1
}

/** Attribute value: references resolved, literal CR/CRLF normalized to LF (as the current parser does). */
function decodeAttributeValue(raw: string): string {
  if (raw.indexOf('&') < 0 && raw.indexOf('\r') < 0) return raw
  let out = ''
  let start = 0
  for (let i = 0; i < raw.length;) {
    const c = raw.charCodeAt(i)
    if (c === AMP) {
      out += raw.slice(start, i)
      const next = scanReference(raw, i)
      if (next === INCOMPLETE) parseError('unterminated reference in attribute value')
      out += lastReference
      i = start = next
    } else if (c === CR) {
      out += raw.slice(start, i) + '\n'
      i += raw.charCodeAt(i + 1) === LF ? 2 : 1
      start = i
    } else {
      i++
    }
  }
  return out + raw.slice(start)
}

function normalizeNewlines(s: string): string {
  return s.indexOf('\r') < 0 ? s : s.replace(/\r\n?/g, '\n')
}

export class XmlTokenizer {
  private readonly events: XmlEvents
  private readonly limits: XmlLimits
  private readonly signal?: AbortSignal
  private carry = ''
  private pending = ''
  private textOpen = false
  private readonly stack: string[] = []
  private readonly names: string[] = []
  private readonly values: string[] = []
  private rootClosed = false
  private rootSeen = false
  private bomChecked = false
  private sawContent = false
  private ended = false
  private error: unknown
  private opens = 0

  constructor(events: XmlEvents, limits: Partial<XmlLimits> = {}, options: XmlTokenizerOptions = {}) {
    this.events = events
    this.limits = { ...DEFAULT_XML_LIMITS, ...limits }
    this.signal = options.signal
  }

  /** Characters held between writes (carried token plus undelivered text). */
  get bufferedChars(): number {
    return this.carry.length + this.pending.length
  }

  write(chunk: string): void {
    if (this.error !== undefined) throw this.error
    if (this.ended) throw new PackageError('parse-error', 'write after end')
    try {
      throwIfAborted(this.signal)
      let buf = this.carry.length > 0 ? this.carry + chunk : chunk
      this.carry = ''
      if (!this.bomChecked && buf.length > 0) {
        this.bomChecked = true
        if (buf.charCodeAt(0) === 0xfeff) buf = buf.slice(1)
      }
      this.run(buf)
      if (this.pending.length > 0) this.flushText(true)
    } catch (error) {
      this.error = error
      throw error
    }
  }

  end(): void {
    if (this.error !== undefined) throw this.error
    if (this.ended) return
    this.ended = true
    try {
      if (this.carry.length > 0) {
        const c = this.carry.charCodeAt(0)
        if (c === AMP) parseError('unterminated reference at end of input')
        if (c === LT) parseError('unterminated markup at end of input')
      }
      if (this.stack.length > 0) parseError(`unclosed element <${this.stack[this.stack.length - 1]}> at end of input`)
      if (!this.rootSeen) parseError('no root element')
    } catch (error) {
      this.error = error
      throw error
    }
  }

  private run(buf: string): void {
    const len = buf.length
    let pos = 0
    while (pos < len) {
      const c = buf.charCodeAt(pos)
      if (c === LT) {
        const next = this.markup(buf, pos)
        if (next === INCOMPLETE) { this.hold(buf, pos); return }
        this.sawContent = true
        pos = next
      } else if (this.stack.length === 0) {
        if (!isWs(c)) parseError(this.rootClosed ? 'text after the root element' : 'text before the root element')
        this.sawContent = true
        pos++
      } else {
        pos = this.text(buf, pos)
      }
    }
  }

  private hold(buf: string, pos: number): void {
    this.carry = buf.slice(pos)
    // Reference and CR holds are structurally bounded (at most 34 chars); only
    // markup is subject to maxTokenChars, whatever limit the caller set.
    if (this.carry.charCodeAt(0) === LT && this.carry.length > this.limits.maxTokenChars) {
      throw new PackageError('resource-limit', `markup token exceeds maxTokenChars (${this.limits.maxTokenChars})`)
    }
  }

  private checkToken(start: number, end: number): void {
    if (end - start > this.limits.maxTokenChars) {
      throw new PackageError('resource-limit', `markup token exceeds maxTokenChars (${this.limits.maxTokenChars})`)
    }
  }

  /** Character data inside the root; returns the position of the next '<' or the end. */
  private text(buf: string, pos: number): number {
    const len = buf.length
    let start = pos
    while (pos < len) {
      const c = buf.charCodeAt(pos)
      if (c === LT) break
      if (c === AMP) {
        if (pos > start) this.addText(buf.slice(start, pos))
        const next = scanReference(buf, pos)
        if (next === INCOMPLETE) { this.hold(buf, pos); return len }
        this.addText(lastReference)
        pos = start = next
      } else if (c === CR) {
        if (pos > start) this.addText(buf.slice(start, pos))
        if (pos + 1 === len) { this.hold(buf, pos); return len }
        this.addText('\n')
        pos = start = pos + (buf.charCodeAt(pos + 1) === LF ? 2 : 1)
      } else {
        pos++
      }
    }
    if (pos > start) this.addText(buf.slice(start, pos))
    return pos
  }

  private addText(s: string): void {
    if (s.length === 0) return
    this.pending = this.pending.length > 0 ? this.pending + s : s
    const max = this.limits.maxTextSegmentChars
    while (this.pending.length >= max) {
      let cut = max
      // keep a surrogate pair in one segment when the cap allows it
      const unit = this.pending.charCodeAt(cut - 1)
      if (cut > 1 && unit >= 0xd800 && unit <= 0xdbff) cut--
      this.events.text(this.pending.slice(0, cut), true)
      this.textOpen = true
      this.pending = this.pending.slice(cut)
    }
  }

  private flushText(more: boolean): void {
    if (this.pending.length > 0) {
      const text = this.pending
      this.pending = ''
      this.textOpen = more
      this.events.text(text, more)
    } else if (!more && this.textOpen) {
      this.textOpen = false
      this.events.text('', false)
    }
  }

  private markup(buf: string, pos: number): number {
    if (pos + 1 >= buf.length) return INCOMPLETE
    const c1 = buf.charCodeAt(pos + 1)
    if (c1 === BANG) return this.bang(buf, pos)
    if (c1 === QM) return this.processingInstruction(buf, pos)
    if (c1 === SLASH) return this.endTag(buf, pos)
    return this.startTag(buf, pos)
  }

  private bang(buf: string, pos: number): number {
    if (buf.startsWith('<!--', pos)) {
      const end = buf.indexOf('-->', pos + 4)
      if (end < 0) return INCOMPLETE
      this.checkToken(pos, end + 3)
      return end + 3
    }
    if (buf.startsWith('<![CDATA[', pos)) {
      const end = buf.indexOf(']]>', pos + 9)
      if (end < 0) return INCOMPLETE
      this.checkToken(pos, end + 3)
      if (this.stack.length === 0) parseError('CDATA outside the root element')
      this.addText(normalizeNewlines(buf.slice(pos + 9, end)))
      return end + 3
    }
    if (buf.startsWith('<!DOCTYPE', pos)) throw new PackageError('unsupported', 'DOCTYPE/DTD is not supported in large-source mode')
    const head = buf.slice(pos, pos + 9)
    if (head.length < 9 && ('<!--'.startsWith(head) || '<![CDATA['.startsWith(head) || '<!DOCTYPE'.startsWith(head))) return INCOMPLETE
    parseError('malformed <! markup')
  }

  private processingInstruction(buf: string, pos: number): number {
    const end = buf.indexOf('?>', pos + 2)
    if (end < 0) return INCOMPLETE
    this.checkToken(pos, end + 2)
    let j = pos + 2
    if (j >= end || !isNameStart(buf.charCodeAt(j))) parseError('invalid processing instruction target')
    while (j < end && isNameChar(buf.charCodeAt(j))) j++
    if (j < end && !isWs(buf.charCodeAt(j))) parseError('invalid processing instruction target')
    if (buf.slice(pos + 2, j).toLowerCase() === 'xml' && this.sawContent) parseError('XML declaration allowed only at the start of the document')
    return end + 2
  }

  private endTag(buf: string, pos: number): number {
    const end = buf.indexOf('>', pos + 2)
    if (end < 0) return INCOMPLETE
    this.checkToken(pos, end + 1)
    let j = pos + 2
    if (!isNameStart(buf.charCodeAt(j))) parseError('invalid end tag')
    while (j < end && isNameChar(buf.charCodeAt(j))) j++
    const name = buf.slice(pos + 2, j)
    for (; j < end; j++) if (!isWs(buf.charCodeAt(j))) parseError(`invalid end tag </${name}`)
    const top = this.stack[this.stack.length - 1]
    if (top !== name) parseError(top === undefined ? `end tag </${name}> without an open element` : `end tag </${name}> does not match <${top}>`)
    this.flushText(false)
    this.stack.pop()
    if (this.stack.length === 0) this.rootClosed = true
    this.events.close(name)
    return end + 1
  }

  private startTag(buf: string, pos: number): number {
    const len = buf.length
    let gt = pos + 1
    let quote = 0
    for (; gt < len; gt++) {
      const c = buf.charCodeAt(gt)
      if (quote !== 0) { if (c === quote) quote = 0 } else if (c === DQ || c === SQ) quote = c
      else if (c === GT) break
    }
    if (gt >= len) return INCOMPLETE
    this.checkToken(pos, gt + 1)
    let end = gt
    const selfClosing = buf.charCodeAt(gt - 1) === SLASH && gt - 1 > pos
    if (selfClosing) end = gt - 1
    let k = pos + 1
    if (k >= end || !isNameStart(buf.charCodeAt(k))) parseError('invalid element name')
    while (k < end && isNameChar(buf.charCodeAt(k))) k++
    const name = buf.slice(pos + 1, k)
    const names = this.names
    const values = this.values
    let count = 0
    while (k < end) {
      if (!isWs(buf.charCodeAt(k))) parseError(`unexpected character in <${name}>`)
      while (k < end && isWs(buf.charCodeAt(k))) k++
      if (k === end) break
      const nameStart = k
      if (!isNameStart(buf.charCodeAt(k))) parseError(`invalid attribute name in <${name}>`)
      while (k < end && isNameChar(buf.charCodeAt(k))) k++
      const attr = buf.slice(nameStart, k)
      while (k < end && isWs(buf.charCodeAt(k))) k++
      if (k >= end || buf.charCodeAt(k) !== EQ) parseError(`attribute ${attr} without value`)
      k++
      while (k < end && isWs(buf.charCodeAt(k))) k++
      const q = k < end ? buf.charCodeAt(k) : 0
      if (q !== DQ && q !== SQ) parseError(`unquoted value for attribute ${attr}`)
      const close = buf.indexOf(q === DQ ? '"' : "'", k + 1)
      if (close < 0 || close >= end) parseError(`unterminated value for attribute ${attr}`)
      for (let i = 0; i < count; i++) if (names[i] === attr) parseError(`duplicate attribute ${attr}`)
      if (count >= this.limits.maxAttributes) throw new PackageError('resource-limit', `more than ${this.limits.maxAttributes} attributes on <${name}>`)
      names[count] = attr
      values[count] = decodeAttributeValue(buf.slice(k + 1, close))
      count++
      k = close + 1
    }
    if (this.stack.length === 0) {
      if (this.rootClosed) parseError('more than one root element')
      this.rootSeen = true
    }
    if (this.stack.length + 1 > this.limits.maxDepth) throw new PackageError('resource-limit', `nesting deeper than ${this.limits.maxDepth}`)
    if ((++this.opens & 1023) === 0) throwIfAborted(this.signal)
    this.flushText(false)
    this.events.open(name, names, values, count, selfClosing)
    if (selfClosing) {
      if (this.stack.length === 0) this.rootClosed = true
    } else {
      this.stack.push(name)
    }
    return gt + 1
  }
}

/** UTF-8 bytes to UTF-16 strings across chunk boundaries (stream mode). A
 * UTF-8 BOM is stripped; a UTF-16 BOM is refused (v1, design note §7.4).
 * Non-fatal decoding substitutes U+FFFD as an external-input policy and
 * counts only substitutions, not genuine U+FFFD characters in the input. */
export class Utf8ChunkDecoder {
  private readonly decoder: TextDecoder
  private readonly fatal: boolean
  private head: Uint8Array | null = new Uint8Array(0)
  private prev1 = -1
  private prev2 = -1
  private replacements = 0

  constructor(opts: { fatal?: boolean } = {}) {
    this.fatal = opts.fatal === true
    this.decoder = new TextDecoder('utf-8', { fatal: this.fatal })
  }

  get replacementCount(): number {
    return this.replacements
  }

  decode(bytes: Uint8Array, last = false): string {
    let input = bytes
    if (this.head !== null) {
      const merged = new Uint8Array(this.head.length + bytes.length)
      merged.set(this.head)
      merged.set(bytes, this.head.length)
      if (merged.length < 2 && !last) {
        this.head = merged
        return ''
      }
      if (merged.length >= 2 && ((merged[0] === 0xff && merged[1] === 0xfe) || (merged[0] === 0xfe && merged[1] === 0xff))) {
        throw new PackageError('unsupported', 'UTF-16 encoded XML part is not supported')
      }
      this.head = null
      input = merged
    }
    let out: string
    try {
      out = this.decoder.decode(input, { stream: !last })
    } catch {
      throw new PackageError('parse-error', 'invalid UTF-8 in XML part')
    }
    if (!this.fatal && out.indexOf('�') >= 0) this.count(out, input)
    if (input.length >= 2) { this.prev2 = input[input.length - 2]; this.prev1 = input[input.length - 1] }
    else if (input.length === 1) { this.prev2 = this.prev1; this.prev1 = input[0] }
    return out
  }

  /** U+FFFD in the output minus genuine EF BF BD sequences completed in this chunk. */
  private count(out: string, input: Uint8Array): void {
    let produced = 0
    for (let i = out.indexOf('�'); i >= 0; i = out.indexOf('�', i + 1)) produced++
    let genuine = 0
    const at = (i: number): number => (i >= 0 ? input[i] : i === -1 ? this.prev1 : this.prev2)
    for (let i = 0; i < input.length; i++) {
      if (input[i] === 0xbd && at(i - 1) === 0xbf && at(i - 2) === 0xef) genuine++
    }
    this.replacements += produced - genuine
  }
}

/** Stateful SpreadsheetML ST_Xstring decoder: equals decodeXlsxString on the
 * concatenation for any segmentation. Holds back only a proper prefix of
 * `_xHHHH_` (at most 6 characters), so the closing `_` of one escape never
 * starts the next (`_x005F_x0041_` is the literal `_x0041_`). */
export class XstringDecoder {
  private held = ''

  push(segment: string, final = false): string {
    const s = this.held.length > 0 ? this.held + segment : segment
    this.held = ''
    if (s.indexOf('_') < 0) return s
    let out = ''
    let start = 0
    let i = s.indexOf('_')
    while (i >= 0) {
      if (i + 7 <= s.length) {
        if (isEscapeAt(s, i)) {
          out += s.slice(start, i) + String.fromCharCode(parseInt(s.slice(i + 2, i + 6), 16))
          start = i + 7
          i = s.indexOf('_', start)
          continue
        }
      } else if (!final && isEscapePrefix(s, i)) {
        out += s.slice(start, i)
        this.held = s.slice(i)
        return out
      }
      i = s.indexOf('_', i + 1)
    }
    return out + s.slice(start)
  }
}

function isHex(c: number): boolean {
  return (c >= 48 && c <= 57) || (c >= 65 && c <= 70) || (c >= 97 && c <= 102)
}
function isEscapeAt(s: string, i: number): boolean {
  return s.charCodeAt(i + 1) === 120 && isHex(s.charCodeAt(i + 2)) && isHex(s.charCodeAt(i + 3))
    && isHex(s.charCodeAt(i + 4)) && isHex(s.charCodeAt(i + 5)) && s.charCodeAt(i + 6) === 95
}
/** s.slice(i) (shorter than 7) could still become `_xHHHH_`. */
function isEscapePrefix(s: string, i: number): boolean {
  for (let k = i + 1; k < s.length; k++) {
    const c = s.charCodeAt(k)
    if (k === i + 1 ? c !== 120 : !isHex(c)) return false
  }
  return true
}

const LOCAL_NAME_CACHE_MAX = 4096
const localNames = new Map<string, string>()

/** Prefix strip matching src/core/xml.ts normalize. The cache is bounded:
 * a hostile part can mint any number of distinct qnames. */
export function localName(qname: string): string {
  const colon = qname.indexOf(':')
  if (colon < 0) return qname
  let local = localNames.get(qname)
  if (local === undefined) {
    if (localNames.size >= LOCAL_NAME_CACHE_MAX) localNames.clear()
    local = qname.slice(colon + 1)
    localNames.set(qname, local)
  }
  return local
}

function isNamespaceDeclaration(qname: string): boolean {
  return qname === 'xmlns' || qname.startsWith('xmlns:')
}

/** The consumer-side duplicate check after prefix stripping (design note §3.4):
 * `x:k` and `y:k` are distinct XML attributes but collide as local names. */
export function findLocalNameCollision(names: readonly string[], count: number): string | undefined {
  for (let i = 0; i < count; i++) {
    if (isNamespaceDeclaration(names[i])) continue
    const a = localName(names[i])
    for (let j = i + 1; j < count; j++) {
      if (!isNamespaceDeclaration(names[j]) && localName(names[j]) === a) return a
    }
  }
  return undefined
}

export function assertUniqueLocalAttributes(names: readonly string[], count: number): void {
  const dup = findLocalNameCollision(names, count)
  if (dup !== undefined) throw new PackageError('parse-error', `duplicate attribute ${dup} after prefix stripping`)
}

export interface TokenizeOptions {
  signal?: AbortSignal
  limits?: Partial<XmlLimits>
  fatal?: boolean
  /** Awaited after every chunk: the downstream backpressure point (e.g. a page commit). */
  afterChunk?: () => void | Promise<void>
}

/** Pull bytes (e.g. openPart), decode UTF-8 and tokenize. The pull loop is the
 * caller's: nothing is read ahead while afterChunk is pending. */
export async function tokenizeBytes(source: AsyncIterable<Uint8Array>, events: XmlEvents, options: TokenizeOptions = {}): Promise<{ replacementCount: number }> {
  const decoder = new Utf8ChunkDecoder({ fatal: options.fatal })
  const tokenizer = new XmlTokenizer(events, options.limits, { signal: options.signal })
  for await (const bytes of source) {
    throwIfAborted(options.signal)
    tokenizer.write(decoder.decode(bytes))
    if (options.afterChunk !== undefined) await options.afterChunk()
    throwIfAborted(options.signal)
  }
  tokenizer.write(decoder.decode(new Uint8Array(0), true))
  tokenizer.end()
  return { replacementCount: decoder.replacementCount }
}
