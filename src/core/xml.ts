/**
 * Thin, forgiving XML helpers built on fast-xml-parser. OOXML namespaces
 * (w:, a:, p:, c:) are handled by stripping prefixes so lookups are
 * namespace-agnostic.
 */
import { XMLParser, XMLValidator } from 'fast-xml-parser'

export type XmlValue = string | number | boolean | null
export interface XmlAttrs extends Record<string, string> {}
export interface XmlNode {
  '@attrs'?: XmlAttrs
  '#text'?: string
  [k: string]: XmlValue | XmlAttrs | XmlNode | XmlNode[] | undefined
}

// Keys fast-xml-parser reserves.
const ATTRS = '@attrs'
const TEXT = '#text'

function isValidXmlScalar(cp: number): boolean {
  if (cp === 0x9 || cp === 0xA || cp === 0xD) return true
  if (cp >= 0x20 && cp <= 0xD7FF) return true
  if (cp >= 0xE000 && cp <= 0xFFFD) return true
  if (cp >= 0x10000 && cp <= 0x10FFFF) return true
  return false
}

function validateNumericReferences(xml: string): void {
  // Skip ref-looking content inside CDATA, comments, and PI
  const sanitized = xml.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/g, '')
  const regex = /&#(?:([0-9]+)|[xX]([0-9a-fA-F]+));/g
  let match: RegExpExecArray | null
  while ((match = regex.exec(sanitized)) !== null) {
    const cp = match[1] !== undefined ? parseInt(match[1], 10) : parseInt(match[2], 16)
    if (Number.isNaN(cp) || !isValidXmlScalar(cp)) {
      throw new Error(`Invalid XML numeric character reference: ${match[0]} (U+${cp.toString(16).toUpperCase()})`)
    }
  }
}

// Installed fast-xml-parser runtime accepts an empty named-entity object
// (enabling numeric character reference decoding while keeping HTML aliases off),
// while TypeScript declarations type htmlEntities as boolean; tests pin this behavior.
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  attributesGroupName: '@attrs',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  processEntities: true,
  htmlEntities: {} as unknown as boolean,
});

// Same options, but siblings (including text segments) stay in document
// order. ODF mixes bare text with inline elements (<p>hello <span>bold</span>
// world</p>), and the default parser joins those segments into one string,
// destroying run order. ODF content is parsed with this; OOXML keeps the
// default parser (w:t elements already carry order).
// Quirk: in preserveOrder mode fast-xml-parser groups attributes under ':@'
// instead of the configured attributesGroupName.
const orderedParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  attributesGroupName: '@attrs',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  processEntities: true,
  preserveOrder: true,
  htmlEntities: {} as unknown as boolean,
})

/** Document-order children per node, populated only by parseXmlOrdered. */
const orderMap = new WeakMap<XmlNode, Array<[string, XmlNode]>>()
/** In-scope namespace bindings, kept out of the normalized/serialized node. */
const namespaceMap = new WeakMap<XmlNode, ReadonlyMap<string, string>>()

function withNamespaces(parent: ReadonlyMap<string, string>, raw: unknown): ReadonlyMap<string, string> {
  if (!raw || typeof raw !== 'object') return parent
  let local: Map<string, string> | undefined
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    if (key !== 'xmlns' && !key.startsWith('xmlns:')) continue
    local ??= new Map(parent)
    local.set(key === 'xmlns' ? '' : key.slice(6), String(value))
  }
  return local ?? parent
}

/** Resolve a prefix used by MC Requires at this node, honoring local shadowing. */
export function namespaceUri(node: XmlNode | undefined, prefix: string): string | undefined {
  return node ? namespaceMap.get(node)?.get(prefix) : undefined
}

function stripPrefix(k: string): string {
  return k.includes(':') ? k.slice(k.indexOf(':') + 1) : k
}

function collectOrderedAttrs(out: Record<string, string>, v: unknown): void {
  if (!v || typeof v !== 'object') return
  for (const [ak, av] of Object.entries(v as Record<string, unknown>)) {
    const stripped = stripPrefix(ak)
    if (stripped === 'xmlns' || ak.startsWith('xmlns')) continue
    out[stripped] = String(av as XmlValue)
  }
}

/**
 * Normalize one preserveOrder value (always an array of single-key entries
 * for elements) into the shared XmlNode shape, recording sibling order.
 */
function buildOrdered(entries: unknown[], inherited: ReadonlyMap<string, string> = new Map()): XmlNode {
  const node: Record<string, unknown> = {}
  const attrs: Record<string, string> = {}
  const ordered: Array<[string, XmlNode]> = []
  const grouped = new Map<string, XmlNode[]>()
  const push = (name: string, child: XmlNode): void => {
    const list = grouped.get(name) ?? []
    list.push(child)
    grouped.set(name, list)
    ordered.push([name, child])
  }
  // In preserveOrder mode each tag and its ':@' attributes arrive paired in
  // one entry ({tag: [...], ':@': {...}}), so ':@' always belongs to the most
  // recently pushed sibling — never to the accumulating parent.
  let last: XmlNode | undefined
  for (const entry of entries) {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) continue
    for (const [k, v] of Object.entries(entry as Record<string, unknown>)) {
      if (k === ':@') {
        collectOrderedAttrs((last?.[ATTRS] as Record<string, string> | undefined) ?? attrs, v)
        continue
      }
      if (k.startsWith('?')) continue // ?xml prolog
      const nk = stripPrefix(k)
      if (nk === '#text') {
        const text = String(v ?? '')
        node[TEXT] = ((node[TEXT] as string) ?? '') + text
        // segments stay out of `grouped` so the merged string above survives
        if (text.length > 0) ordered.push(['#text', { [ATTRS]: {}, [TEXT]: text }])
        continue
      }
      if (!Array.isArray(v)) continue
      const child = buildOrdered(v as unknown[], withNamespaces(inherited, (entry as Record<string, unknown>)[':@']))
      push(nk, child)
      last = child
    }
  }
  for (const [name, list] of grouped) node[name] = list.length === 1 ? list[0] : list
  node[ATTRS] = attrs
  const result = node as XmlNode
  orderMap.set(result, ordered)
  namespaceMap.set(result, inherited)
  return result
}

/** Parse XML keeping document order (see orderedParser). Throws on malformed input. */
export function parseXmlOrdered(xml: string): XmlNode {
  if (xml.charCodeAt(0) === 0xfeff) xml = xml.slice(1)
  validateNumericReferences(xml)
  const check = XMLValidator.validate(xml)
  if (check !== true) {
    throw new Error(`XML validation failed: ${JSON.stringify(check)}`)
  }
  const parsed = orderedParser.parse(xml) as Array<Record<string, unknown>>
  if (!Array.isArray(parsed)) throw new Error('Unexpected XML root shape')
  for (const entry of parsed) {
    for (const [k, v] of Object.entries(entry)) {
      if (k.startsWith('?')) continue
      if (!Array.isArray(v)) continue
      const rootNode = buildOrdered(v as unknown[], withNamespaces(new Map(), entry[':@']))
      collectOrderedAttrs(rootNode[ATTRS] as Record<string, string>, entry[':@'])
      return rootNode
    }
  }
  throw new Error('Unexpected XML root shape')
}

/**
 * Document-order [name, child] pairs for a node parsed with parseXmlOrdered
 * (whitespace-only text segments included — callers filter). Falls back to
 * elementChildren order for nodes from the default parser.
 */
export function orderedChildren(node: XmlNode | undefined): Array<[string, XmlNode]> {
  if (!node) return []
  const ordered = orderMap.get(node)
  if (ordered) return ordered
  const out: Array<[string, XmlNode]> = []
  const direct = node[TEXT]
  if (typeof direct === 'string' && direct.length > 0) out.push(['#text', { [ATTRS]: {}, [TEXT]: direct }])
  out.push(...elementChildren(node))
  return out
}

/**
 * Normalize:
 *  - lift the parser's '@attrs' group into our attrs record, stripping
 *    namespace prefixes from attribute names
 *  - strip namespace prefixes from element names ("w:p" -> "p")
 */
function normalize(node: unknown, inherited: ReadonlyMap<string, string> = new Map()): XmlNode | XmlNode[] | XmlValue | undefined {
  if (Array.isArray(node)) return node.map(value => normalize(value, inherited)) as XmlNode[]
  if (node === null || node === undefined) return undefined
  if (typeof node !== 'object') return node as XmlValue
  const src = node as Record<string, unknown>
  const local = withNamespaces(inherited, src['@attrs'])
  const out: Record<string, unknown> = {}
  const attrs: Record<string, string> = {}
  for (const [k, v] of Object.entries(src)) {
    if (k === '@attrs' && v && typeof v === 'object') {
      for (const [ak, av] of Object.entries(v as Record<string, unknown>)) {
        const stripped = ak.includes(':') ? ak.slice(ak.indexOf(':') + 1) : ak
        if (stripped === 'xmlns' || ak.startsWith('xmlns')) continue
        attrs[stripped] = String(av as XmlValue)
      }
      continue
    }
    if (k.startsWith('xmlns')) continue
    if (k === TEXT) {
      out[TEXT] = v
      continue
    }
    const nk = k.includes(':') ? k.slice(k.indexOf(':') + 1) : k
    out[nk] = normalize(v, local)
  }
  out[ATTRS] = attrs
  namespaceMap.set(out as XmlNode, local)
  return out as XmlNode
}

/** Parse an XML string into a normalized node tree. Throws on malformed input. */
export function parseXml(xml: string): XmlNode {
  // strip UTF-8 BOM if present (some generators emit it)
  if (xml.charCodeAt(0) === 0xfeff) xml = xml.slice(1)
  validateNumericReferences(xml)
  // fast-xml-parser can leave entity garbage in some edge documents — validate first
  const check = XMLValidator.validate(xml)
  if (check !== true) {
    throw new Error(`XML validation failed: ${JSON.stringify(check)}`)
  }
  const parsed = parser.parse(xml) as Record<string, unknown>
  // The parser adds a "?xml" key for the prolog and a "#text" key for
  // trailing whitespace — neither is the root element.
  const rootKey = Object.keys(parsed).find((k) => k !== '?xml' && k !== '#text')
  if (rootKey === undefined) {
    throw new Error('Unexpected XML root shape')
  }
  const normalized = normalize(parsed[rootKey])
  if (normalized === undefined || Array.isArray(normalized) || typeof normalized !== 'object') {
    throw new Error('Unexpected XML root shape')
  }
  return normalized as XmlNode
}

/** Access a child (or children) by name, tolerating single-vs-array shapes. */
export function getChild(node: XmlNode | undefined, name: string): XmlNode | undefined {
  if (!node) return undefined
  const v = node[name]
  if (Array.isArray(v)) return v[0] as XmlNode | undefined
  return v as XmlNode | undefined
}

/** Wrap a primitive parser value (text-only element) as a node. */
function asNode(v: unknown): XmlNode {
  if (v && typeof v === 'object') return v as XmlNode
  return { [ATTRS]: {}, [TEXT]: v === null || v === undefined ? '' : String(v) }
}

export function getChildren(node: XmlNode | undefined, name: string): XmlNode[] {
  if (!node) return []
  const v = node[name]
  if (Array.isArray(v)) return v.map(asNode)
  if (v === undefined || v === null) return []
  return [asNode(v)]
}

/** Element children (excluding @attrs/#text) of a node. */
export function elementChildren(node: XmlNode | undefined): Array<[string, XmlNode]> {
  if (!node) return []
  const out: Array<[string, XmlNode]> = []
  for (const [k, v] of Object.entries(node)) {
    if (k === ATTRS || k === TEXT) continue
    if (Array.isArray(v)) {
      for (const item of v) out.push([k, asNode(item)])
    } else if (v !== undefined && v !== null) {
      out.push([k, asNode(v)])
    }
  }
  return out
}

export function attrs(node: XmlNode | undefined): XmlAttrs {
  return (node?.[ATTRS] as XmlAttrs) ?? {}
}

export function textOf(node: unknown): string {
  if (node === null || node === undefined) return ''
  if (typeof node !== 'object') return String(node)
  const n = node as XmlNode
  const direct = n[TEXT]
  if (typeof direct === 'string' && direct.length > 0) return direct
  let out = ''
  for (const [, child] of elementChildren(n)) {
    out += textOf(child)
  }
  if (typeof direct === 'string' && direct.length > 0) out = direct + out
  return out
}
