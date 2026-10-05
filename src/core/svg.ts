/**
 * Self-contained SVG preflight and `asvg:svgBlip` location.
 *
 * The scanner is a bounded XML lexer, not an entity-expanding parser: it never
 * hands bytes to a decoder before a verdict. It tokenizes the ORIGINAL markup
 * (start/end tags, comments, CDATA, PIs, text) first, then decodes character
 * references only WITHIN the resulting attribute/style values — decoded
 * characters are never reinterpreted as XML syntax. Active-element matching is
 * derived only from real start-tag tokens, so literal `<script>` text (raw or
 * via `&lt;script&gt;` or CDATA) is text, not an element.
 */
import { attrs, getChildren, orderedChildren, type XmlNode } from './xml'

/** Standard extension GUID that pairs `asvg:svgBlip` with a raster `a:blip`. */
export const SVG_BLIP_EXT_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}'

export type SvgVerdict = { ok: true } | { ok: false; reason: string }

/**
 * A retained SVG candidate (accepted OR rejected) alongside a raster fallback.
 * `bytes` is deliberately not named `data`: the accepted-Task3 identity audit
 * treats any nested `{ data, referenceId }` object as a source image identity.
 */
export interface SvgCandidate {
  bytes: Uint8Array
  partPath?: string
  referenceId?: string
  /** Immutable static preflight verdict; decode preference derives from this. */
  verdict: SvgVerdict
}

/** Parse-time plan and decode-time outcome for one picture identity. */
export interface ImageSelection {
  phase: 'pending' | 'decoded'
  representation: 'svg' | 'raster' | 'none'
  reason?: string
}

/** Element local names that carry active content and are always rejected. */
const ACTIVE_ELEMENTS = new Set([
  'script', 'foreignobject', 'iframe', 'object', 'embed', 'link', 'meta', 'base',
  'html', 'body', 'audio', 'video', 'canvas',
])

/** SMIL elements that can mutate a resource-bearing attribute. */
const SMIL_ELEMENTS = new Set(['set', 'animate', 'animatecolor', 'animatemotion', 'animatetransform'])
/** Mutation values that name a resource. */
const SMIL_VALUE_ATTRS = ['to', 'from', 'values', 'by'] as const

const localName = (name: string): string => (name.includes(':') ? name.slice(name.lastIndexOf(':') + 1) : name)
const XML_ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }
/**
 * XML 1.0 NameStartChar (production 4, colon excluded => NCName) plus the
 * production 4a extras. Prefixes/names are recognised by the real grammar, not
 * an ASCII or Unicode-letter approximation.
 */
const XML_NAME_START = 'A-Za-z_\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF\\u0370-\\u037D\\u037F-\\u1FFF\\u200C-\\u200D\\u2070-\\u218F\\u2C00-\\u2FEF\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD\\u{10000}-\\u{EFFFF}'
const XML_NAME_EXTRA = '\\-.0-9\\u00B7\\u0300-\\u036F\\u203F-\\u2040'
const NCNAME = `[${XML_NAME_START}][${XML_NAME_START}${XML_NAME_EXTRA}]*`
const SVG_ROOT_RE = new RegExp(`<(?:${NCNAME}:)?svg[\\s/>]`, 'u')

/**
 * XML 1.0 Appendix F auto-detection of recognizable non-UTF-8 XML encodings
 * (BOM or no-BOM signature). Returns the encoding when the bytes are recognisable
 * as encoded XML, else undefined. Bounded: only the first bytes are inspected.
 */
function detectEncodedXml(bytes: Uint8Array): 'utf-16le' | 'utf-16be' | 'utf-32le' | 'utf-32be' | 'ebcdic' | 'undecided' | undefined {
  const b = bytes
  // Explicit BOMs (longest first; UTF-32 LE BOM starts with the UTF-16 LE BOM).
  if (b.length >= 4 && b[0] === 0x00 && b[1] === 0x00 && b[2] === 0xfe && b[3] === 0xff) return 'utf-32be'
  if (b.length >= 4 && b[0] === 0xff && b[1] === 0xfe && b[2] === 0x00 && b[3] === 0x00) return 'utf-32le'
  if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) return 'utf-16be'
  if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) return 'utf-16le'
  // Appendix F no-BOM signatures.
  if (b.length >= 4) {
    const [c0, c1, c2, c3] = [b[0], b[1], b[2], b[3]]
    if (c0 === 0x00 && c1 === 0x00 && c2 === 0x00 && c3 === 0x3c) return 'utf-32be'
    if (c0 === 0x3c && c1 === 0x00 && c2 === 0x00 && c3 === 0x00) return 'utf-32le'
    if (c0 === 0x00 && c1 === 0x00 && c2 === 0x3c && c3 === 0x00) return 'utf-32be'
    if (c0 === 0x00 && c1 === 0x3c && c2 === 0x00 && c3 === 0x00) return 'utf-32le'
    if (c0 === 0x00 && c1 === 0x3c && c2 === 0x00 && c3 === 0x3f) return 'utf-16be'
    if (c0 === 0x3c && c1 === 0x00 && c2 === 0x3f && c3 === 0x00) return 'utf-16le'
    if (c0 === 0x4c && c1 === 0x6f && c2 === 0xa7 && c3 === 0x94) return 'ebcdic'
  }
  // No BOM and no declaration signature: scan the WHOLE stream at the candidate
  // code-unit width, skipping legal encoded leading whitespace / BOM / NUL until
  // a decision. An all-whitespace stream is UNDECIDED and fails closed (never
  // silently classified as raster). Width 4 is checked before width 2 so a
  // zero-padded UTF-32 document is not mistaken for UTF-16.
  const u32le = utfXmlScan(b, 4, true)
  if (u32le === 'xml') return 'utf-32le'
  const u32be = utfXmlScan(b, 4, false)
  if (u32be === 'xml') return 'utf-32be'
  const u16le = utfXmlScan(b, 2, true)
  if (u16le === 'xml') return 'utf-16le'
  const u16be = utfXmlScan(b, 2, false)
  if (u16be === 'xml') return 'utf-16be'
  if (u32le === 'undecided' || u32be === 'undecided' || u16le === 'undecided' || u16be === 'undecided') return 'undecided'
  return undefined
}
/**
 * Scan the WHOLE stream at `width` (LE/BE), skipping legal encoded leading
 * whitespace / BOM / NUL until the first content code unit. Returns 'xml' when
 * that unit is a zero-padded '<', 'other' for any other content, and 'undecided'
 * when the whole stream is zero-padded legal whitespace.
 */
function utfXmlScan(bytes: Uint8Array, width: number, le: boolean): 'xml' | 'other' | 'undecided' {
  const limit = bytes.length - (bytes.length % width)
  if (limit < width) return 'other'
  let sawWhitespace = false, zeroPadded = true
  for (let i = 0; i < limit; i += width) {
    let code = 0, highZero = true
    for (let k = 0; k < width; k++) {
      const byte = bytes[i + (le ? k : width - 1 - k)]
      code |= byte << (8 * k)
      if (k > 0 && byte !== 0) highZero = false
    }
    if (code === 0) continue
    const ch = code & 0xffff
    if (ch === 0x20 || ch === 0x09 || ch === 0x0a || ch === 0x0d || ch === 0xfeff || ch === 0xfffe) { sawWhitespace = true; if (!highZero) zeroPadded = false; continue }
    return ch === 0x3c && highZero ? 'xml' : 'other'
  }
  return sawWhitespace && zeroPadded ? 'undecided' : 'other'
}
/** Known raster magics pass through untouched (normal raster behaviour preserved). */
function isRasterMagic(b: Uint8Array): boolean {
  if (b.length >= 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return true
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return true
  if (b.length >= 3 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return true
  if (b.length >= 12 && b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return true
  return false
}

/**
 * Decode XML character references / predefined entities WITHIN one token value.
 * Callers only ever pass an attribute value or a style/text payload, so decoded
 * characters can never be reinterpreted as markup.
 */
export function normalizeXmlCharRefs(text: string): string {
  return text
    .replace(/&#(?:x([0-9a-fA-F]+)|(\d+));/g, (whole, hex: string | undefined, dec: string | undefined) => {
      const code = hex !== undefined ? parseInt(hex, 16) : parseInt(dec ?? '', 10)
      if (!Number.isFinite(code) || code <= 0 || code > 0x10ffff) return whole
      try { return String.fromCodePoint(code) } catch { return whole }
    })
    .replace(/&(amp|lt|gt|quot|apos);/g, (_, name: string) => XML_ENTITIES[name] ?? _)
}

/** Classify an SVG URL attribute/property value (SVG 2 §16.1.2). */
export function classifySvgUrl(raw: string): 'empty' | 'fragment' | 'external' | 'data' {
  const trimmed = raw.trim()
  if (!trimmed) return 'empty'
  // Strip whitespace/control obfuscation before looking at the scheme.
  const normalized = trimmed.replace(/[\s\u0000-\u001f]+/g, '')
  const lower = normalized.toLowerCase()
  if (lower.startsWith('#')) return 'fragment'
  if (lower.startsWith('data:')) return 'data'
  return 'external'
}

function scanCssValue(value: string): string | undefined {
  const lower = value.toLowerCase()
  if (!lower.includes('url(') && !lower.includes('@import') && !lower.includes('@font-face') && !lower.includes('\\')) return undefined
  const stripped = value.replace(/\/\*[\s\S]*?\*\//g, '')
  if (stripped.includes('\\')) return 'css-escape'
  const normalized = stripped.replace(/[\s\u0000-\u001f]+/g, '').toLowerCase()
  if (normalized.includes('@import')) return 'css-import'
  if (normalized.includes('@font-face')) return 'css-font-face'
  const url = /url\(\s*(?:"([^"]*)"|'([^']*)'|([^)]*))\s*\)/g
  let match: RegExpExecArray | null
  while ((match = url.exec(normalized))) {
    const kind = classifySvgUrl(match[1] ?? match[2] ?? match[3] ?? '')
    if (kind === 'external' || kind === 'data') return 'css-url'
  }
  return undefined
}

/** Reject any external/data URL in a SMIL mutation value (`#frag` stays valid). */
function scanSmilValue(value: string): string | undefined {
  for (const token of value.split(';')) {
    const kind = classifySvgUrl(token)
    if (kind === 'external' || kind === 'data') return 'smil-resource'
    const css = scanCssValue(token)
    if (css) return css
  }
  return undefined
}

// ---------------------------------------------------------------------------
// Bounded XML lexer
// ---------------------------------------------------------------------------

interface XmlAttr { name: string; raw: string }
type XmlToken =
  | { kind: 'start'; name: string; attrs: XmlAttr[]; selfClosing: boolean }
  | { kind: 'end'; name: string }
  | { kind: 'comment' }
  | { kind: 'cdata'; text: string }
  | { kind: 'pi'; target: string }
  | { kind: 'text'; raw: string }
  | { kind: 'doctype' }
  | { kind: 'malformed'; reason: string }

/** Quote-aware index of the `>` that ends the tag starting at `lt` (-1 if none). */
function findTagEnd(text: string, lt: number): number {
  let quote = ''
  for (let j = lt + 1; j < text.length; j++) {
    const c = text[j]
    if (quote) { if (c === quote) quote = '' }
    else if (c === '"' || c === "'") quote = c
    else if (c === '>') return j
  }
  return -1
}

/** Parse one raw `<...>` slice. Attribute boundaries use the ORIGINAL quotes. */
function parseTagToken(tag: string): XmlToken {
  if (tag[1] === '/') {
    const end = /^<\/\s*([^\s/>]+)/.exec(tag)
    return end ? { kind: 'end', name: end[1] } : { kind: 'malformed', reason: 'invalid-end-tag' }
  }
  const start = /^<\s*([^\s/>]+)/.exec(tag)
  if (!start) return { kind: 'malformed', reason: 'invalid-start-tag' }
  const attrs: XmlAttr[] = []
  const body = tag.slice(1)
  const attribute = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
  let match: RegExpExecArray | null
  while ((match = attribute.exec(body))) attrs.push({ name: match[1], raw: match[2] ?? match[3] ?? '' })
  return { kind: 'start', name: start[1], attrs, selfClosing: /\/\s*>$/.test(tag) }
}

/** Lex markup vs text/CDATA vs comment/PI. Never decodes references into markup. */
function tokenizeXml(text: string): XmlToken[] {
  const tokens: XmlToken[] = []
  let i = 0, textStart = 0
  const flushText = (end: number): void => { if (end > textStart) tokens.push({ kind: 'text', raw: text.slice(textStart, end) }) }
  while (i < text.length) {
    if (text[i] !== '<') { i++; continue }
    flushText(i)
    if (text.startsWith('<!--', i)) {
      const end = text.indexOf('-->', i + 4)
      if (end < 0) return [...tokens, { kind: 'malformed', reason: 'unterminated-comment' }]
      tokens.push({ kind: 'comment' }); i = end + 3; textStart = i; continue
    }
    if (text.startsWith('<![CDATA[', i)) {
      const end = text.indexOf(']]>', i + 9)
      if (end < 0) return [...tokens, { kind: 'malformed', reason: 'unterminated-cdata' }]
      tokens.push({ kind: 'cdata', text: text.slice(i + 9, end) }); i = end + 3; textStart = i; continue
    }
    if (text.startsWith('<!', i)) {
      const end = text.indexOf('>', i + 2)
      const decl = text.slice(i + 2, end < 0 ? text.length : end).trim().toUpperCase()
      if (decl.startsWith('DOCTYPE')) { tokens.push({ kind: 'doctype' }); i = end < 0 ? text.length : end + 1; textStart = i; continue }
      return [...tokens, { kind: 'malformed', reason: decl.startsWith('ENTITY') ? 'entity' : 'declaration' }]
    }
    if (text.startsWith('<?', i)) {
      const end = text.indexOf('?>', i + 2)
      if (end < 0) return [...tokens, { kind: 'malformed', reason: 'unterminated-pi' }]
      tokens.push({ kind: 'pi', target: text.slice(i + 2, end).trim().split(/[\s?]/)[0]?.toLowerCase() ?? '' }); i = end + 2; textStart = i; continue
    }
    const end = findTagEnd(text, i)
    if (end < 0) return [...tokens, { kind: 'malformed', reason: 'unterminated-tag' }]
    tokens.push(parseTagToken(text.slice(i, end + 1))); i = end + 1; textStart = i
  }
  flushText(text.length)
  return tokens
}

/** Attribute rules for one start tag (references decoded within each value only). */
function scanAttributes(local: string, attrs: XmlAttr[]): string | undefined {
  const decoded = attrs.map(a => ({ local: localName(a.name).toLowerCase(), value: normalizeXmlCharRefs(a.raw) }))
  if (SMIL_ELEMENTS.has(local)) {
    for (const key of SMIL_VALUE_ATTRS) for (const a of decoded) {
      if (a.local !== key) continue
      const reason = scanSmilValue(a.value)
      if (reason) return reason
    }
    const target = decoded.find(a => a.local === 'attributename')?.value.trim().toLowerCase()
    if (target === 'href' || target === 'xlink:href') {
      for (const key of SMIL_VALUE_ATTRS) for (const a of decoded) {
        if (a.local !== key) continue
        for (const token of a.value.split(';')) {
          const kind = classifySvgUrl(token)
          if (kind === 'external' || kind === 'data') return 'smil-href'
        }
      }
    }
  }
  // Every qualified attribute is inspected INDEPENDENTLY: a later `xlink:href`
  // or `x:base` must never shadow an earlier resource-bearing value.
  for (const a of decoded) {
    if (a.local.length > 2 && a.local.startsWith('on')) return `event-handler:${a.local}`
    if (a.local === 'href' || a.local === 'src') {
      const kind = classifySvgUrl(a.value)
      if (kind === 'external' || kind === 'data') return 'external-href'
    }
    // An authored external xml:base re-resolves otherwise-internal fragments.
    if (a.local === 'base') {
      const kind = classifySvgUrl(a.value)
      if (kind === 'external' || kind === 'data') return 'external-base'
    }
    const css = scanCssValue(a.value)
    if (css) return css
  }
  return undefined
}

/** Deterministically reject forbidden external/active/entity content before decoding. */
export function scanSelfContainedSvg(bytes: Uint8Array | undefined): SvgVerdict {
  if (!bytes || bytes.length === 0) return { ok: false, reason: 'svg-empty' }
  if (detectEncodedXml(bytes)) return { ok: false, reason: 'xml-encoding-unsupported' }
  if (bytes[0] === 0x1f && bytes[1] === 0x8b) return { ok: false, reason: 'svgz' }
  let text: string
  try { text = new TextDecoder('utf-8', { fatal: false }).decode(bytes) } catch { return { ok: false, reason: 'svg-undecodable' } }
  if (text.includes('\u0000')) return { ok: false, reason: 'svg-binary' }
  let inStyle = false, styleCss = ''
  for (const token of tokenizeXml(text)) {
    if (token.kind === 'malformed') return { ok: false, reason: token.reason }
    if (token.kind === 'doctype') return { ok: false, reason: 'doctype' }
    if (token.kind === 'pi') { if (token.target === 'xml-stylesheet') return { ok: false, reason: 'xml-stylesheet' }; continue }
    if (token.kind === 'comment') continue
    // Style children are CSS: ordinary text is reference-normalized, CDATA is literal.
    if (token.kind === 'text') { if (inStyle) styleCss += normalizeXmlCharRefs(token.raw); continue }
    if (token.kind === 'cdata') { if (inStyle) styleCss += token.text; continue }
    if (token.kind === 'end') {
      if (inStyle && localName(token.name).toLowerCase() === 'style') { inStyle = false; const css = scanCssValue(styleCss); if (css) return { ok: false, reason: css }; styleCss = '' }
      continue
    }
    // Start tag: active element names come ONLY from real element tokens.
    const local = localName(token.name).toLowerCase()
    if (ACTIVE_ELEMENTS.has(local)) return { ok: false, reason: `active-content:${local}` }
    const attrReason = scanAttributes(local, token.attrs)
    if (attrReason) return { ok: false, reason: attrReason }
    if (local === 'style' && !token.selfClosing) { inStyle = true; styleCss = '' }
  }
  if (inStyle) { const css = scanCssValue(styleCss); if (css) return { ok: false, reason: css } }
  return { ok: true }
}

/** Reliable FULL-stream SVG root detection over decoded text (any legal XML prefix). */
function hasSvgRoot(bytes: Uint8Array): boolean {
  let text: string
  try { text = new TextDecoder('utf-8', { fatal: false }).decode(bytes) } catch { return false }
  return SVG_ROOT_RE.test(text)
}

/** True when bytes/mime/path identify SVG, gzipped SVG or an unsupported XML encoding. */
export function looksLikeSvg(bytes: Uint8Array | undefined, mime?: string, pathHint?: string): boolean {
  if (mime && /svg/i.test(mime)) return true
  if (pathHint && /\.svgz?($|[?#])/i.test(pathHint)) return true
  if (!bytes || bytes.length < 2) return false
  if (detectEncodedXml(bytes)) return true
  if (bytes[0] === 0x1f && bytes[1] === 0x8b) return true
  if (isRasterMagic(bytes)) return false
  return hasSvgRoot(bytes)
}

/** Guard any byte stream: scan it only when it looks like SVG, else pass through. */
export function guardSvgBytes(bytes: Uint8Array | undefined, mime?: string, pathHint?: string): SvgVerdict {
  if (!bytes) return { ok: true }
  if (isRasterMagic(bytes)) return { ok: true }
  if (detectEncodedXml(bytes)) return { ok: false, reason: 'xml-encoding-unsupported' }
  if (mime && /svg/i.test(mime)) return scanSelfContainedSvg(bytes)
  if (pathHint && /\.svgz?($|[?#])/i.test(pathHint)) return scanSelfContainedSvg(bytes)
  if (bytes[0] === 0x1f && bytes[1] === 0x8b) return scanSelfContainedSvg(bytes)
  if (hasSvgRoot(bytes)) return scanSelfContainedSvg(bytes)
  return { ok: true }
}

/**
 * Locate `a:blip/a:extLst/a:ext[@uri=GUID]/<prefix>:svgBlip` by the standard
 * extension GUID (namespace-agnostic: the prefix is arbitrary). `r:embed` and
 * `r:link` are read independently of the raster `r:embed`.
 */
export function findSvgBlip(blip: XmlNode | undefined): { embed?: string; link?: string } | undefined {
  if (!blip) return undefined
  for (const extList of getChildren(blip, 'extLst')) {
    for (const ext of getChildren(extList, 'ext')) {
      const uri = attrs(ext).uri ?? ''
      if (uri.toUpperCase() !== SVG_BLIP_EXT_URI.toUpperCase()) continue
      for (const [name, child] of orderedChildren(ext)) {
        if (name === '#text' || localName(name).toLowerCase() !== 'svgblip') continue
        const a = attrs(child)
        const embed = a.embed ?? a['r:embed']
        const link = a.link ?? a['r:link']
        if (embed || link) return { ...(embed ? { embed } : {}), ...(link ? { link } : {}) }
        return {}
      }
    }
  }
  return undefined
}

/** Preflight helper shared by adapters: attach an SVG candidate (accepted or rejected). */
export function svgCandidate(bytes: Uint8Array, partPath?: string, referenceId?: string): SvgCandidate {
  return { bytes, partPath, referenceId, verdict: scanSelfContainedSvg(bytes) }
}

/**
 * A candidate with NO bytes (missing / external / link-only relationship) that
 * still carries its parse state, so complete-pair cache keys can distinguish it
 * from an ordinary raster and each source keeps its own reason.
 */
export function svgStateCandidate(reason: string, partPath?: string, referenceId?: string): SvgCandidate {
  return { bytes: new Uint8Array(0), partPath, referenceId, verdict: { ok: false, reason } }
}
