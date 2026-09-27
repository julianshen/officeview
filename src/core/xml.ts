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

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  attributesGroupName: '@attrs',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  processEntities: true,
})

/**
 * Normalize:
 *  - lift the parser's '@attrs' group into our attrs record, stripping
 *    namespace prefixes from attribute names
 *  - strip namespace prefixes from element names ("w:p" -> "p")
 */
function normalize(node: unknown): XmlNode | XmlNode[] | XmlValue | undefined {
  if (Array.isArray(node)) return node.map(normalize) as XmlNode[]
  if (node === null || node === undefined) return undefined
  if (typeof node !== 'object') return node as XmlValue
  const src = node as Record<string, unknown>
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
    out[nk] = normalize(v)
  }
  out[ATTRS] = attrs
  return out as XmlNode
}

/** Parse an XML string into a normalized node tree. Throws on malformed input. */
export function parseXml(xml: string): XmlNode {
  // fast-xml-parser can leave entity garbage in some edge documents — validate first
  const check = XMLValidator.validate(xml)
  if (check !== true) {
    throw new Error(`XML validation failed: ${JSON.stringify(check)}`)
  }
  const parsed = parser.parse(xml) as Record<string, unknown>
  // Drop the synthetic root wrapper when there's exactly one key
  const keys = Object.keys(parsed)
  const rootKey = keys.length === 1 ? keys[0] : keys[keys.length - 1]
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
