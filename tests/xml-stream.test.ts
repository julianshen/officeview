/** S1.1 incremental XML tokenizer (design note §3.3, §4.2; RED tests 9-13). */
import { describe, expect, it } from 'vitest'
import {
  Utf8ChunkDecoder, XmlTokenizer, XstringDecoder, assertUniqueLocalAttributes, findLocalNameCollision, localName, tokenizeBytes,
  type XmlEvents,
} from '../src/core/xml-stream'
import { PackageError } from '../src/core/package-error'
import { parseXml, getChildren, textOf, type XmlNode } from '../src/core/xml'
import { decodeXlsxString } from '../src/xlsx/strings'
import { createPrng, encodeXlsxString, f4StringFor, FIXTURE_SEEDS } from './helpers/xlsx-viewer-fixtures'
import { everyN, key, outcomeOf, randomCuts, recorder, sharedStringsConsumer, type Outcome } from './helpers/xml-recorder'

const LARGE = process.env.IV_LARGE === '1'

function errorCode(o: Outcome): string {
  return o.ok ? 'ok' : o.code
}

function textOfOutcome(o: Outcome): string {
  if (!o.ok) throw new Error(`unexpected ${o.code}: ${o.message}`)
  return o.events.filter((e) => e.kind === 'text').map((e) => (e as { text: string }).text).join('')
}

/** Outcome for whole, all one-cut splits, 1-char, 5-char feeds. */
function allFeedings(doc: string, limits?: Parameters<typeof outcomeOf>[2]): Outcome[] {
  const out = [outcomeOf(doc, [], limits), outcomeOf(doc, everyN(doc, 1), limits), outcomeOf(doc, everyN(doc, 5), limits)]
  for (let i = 1; i < doc.length; i++) out.push(outcomeOf(doc, [i], limits))
  return out
}

function expectSameEverywhere(doc: string, limits?: Parameters<typeof outcomeOf>[2]): Outcome {
  const outcomes = allFeedings(doc, limits)
  const k0 = key(outcomes[0])
  for (const o of outcomes) expect(key(o)).toBe(k0)
  return outcomes[0]
}

describe('events and grammar', () => {
  it('emits open/close/text with attributes, self-closing flag and resolved references', () => {
    const o = outcomeOf('<?xml version="1.0"?>\n<x:r xmlns:x="u"><a b="1" c=\'q&quot;\'>hi &amp; &#65;&#x42;</a><e/></x:r>', [])
    expect(o).toEqual({
      ok: true,
      events: [
        { kind: 'open', name: 'x:r', attrs: [['xmlns:x', 'u']], selfClosing: false },
        { kind: 'open', name: 'a', attrs: [['b', '1'], ['c', 'q"']], selfClosing: false },
        { kind: 'text', text: 'hi & AB' },
        { kind: 'close', name: 'a' },
        { kind: 'open', name: 'e', attrs: [], selfClosing: true },
        { kind: 'close', name: 'x:r' },
      ],
    })
  })

  it('skips comments and PIs, keeps CDATA verbatim, strips a leading BOM', () => {
    expect(textOfOutcome(outcomeOf('﻿<r>x<!-- c -->y<![CDATA[<&>]]><?pi d?>z</r>', []))).toBe('xy<&>z')
  })

  it('normalizes literal CR and CRLF to LF in text, CDATA and attribute values, also across chunks', () => {
    const doc = '<r a="1\r\n2\r3">x\r\ny\rz<![CDATA[p\r\nq]]></r>'
    const o = expectSameEverywhere(doc)
    expect(textOfOutcome(o)).toBe('x\ny\nzp\nq')
    expect(o.ok && o.events[0]).toEqual({ kind: 'open', name: 'r', attrs: [['a', '1\n2\n3']], selfClosing: false })
  })

  it('&#13; keeps a real CR; unknown &name; stays literal; leading zeros are allowed', () => {
    expect(textOfOutcome(outcomeOf('<r>a&#13;b&foo;&#00000000000000065;</r>', []))).toBe('a\rb&foo;A')
  })

  it('reports the more flag: false when a text run ends at markup', () => {
    const calls: Array<[string, boolean]> = []
    const ev: XmlEvents = { open() {}, close() {}, text(t, more) { calls.push([t, more]) } }
    const tok = new XmlTokenizer(ev)
    tok.write('<r>abc')
    tok.write('def</r>')
    tok.end()
    expect(calls.map(([t]) => t).join('')).toBe('abcdef')
    expect(calls[calls.length - 1][1]).toBe(false)
    expect(calls.slice(0, -1).every(([, m]) => m)).toBe(true)
  })
})

describe('chunk-independence (RED test 9)', () => {
  it('&#00000000000000065; is A for whole, 1-char and 5-char feeds; a 36-digit reference is the same parse-error', () => {
    const a = '<r>&#00000000000000065;</r>'
    for (const cuts of [[], everyN(a, 1), everyN(a, 5)]) expect(textOfOutcome(outcomeOf(a, cuts))).toBe('A')
    const long = `<r>&#${'0'.repeat(34)}65;</r>`
    for (const cuts of [[], everyN(long, 1), everyN(long, 5)]) expect(errorCode(outcomeOf(long, cuts))).toBe('parse-error')
  })

  const POOL = [
    '<a>', '</a>', '<b x="1">', '</b>', '<c/>', '<d y=\'2\' z="&amp;&#65;"/>', 'text', ' ', '\n', '\r\n', '\r', 'é', '😀',
    '&amp;', '&lt;', '&gt;', '&quot;', '&apos;', '&#65;', '&#x41;', '&#00000000000000065;', '&foo;', '&#13;', '_x000D_',
    '<!-- c -->', '<![CDATA[x<&y]]>', '<?pi data?>', '<e f="a>b"/>', '<g\r\nh="1"\r\n/>',
    // malformed
    '&', '&amp', '&#;', '&#xZ;', '&#123456789012345;', `&#${'1'.repeat(37)};`, '&#xD800;', '&#x110000;', '&a#b;',
    '<!--', '<![CDATA[', '<?', '</a', '<e f="1', '<e f=1/>', '<e f="1" f="2"/>', '<!DOCTYPE x>', '<', '>', ']]>',
    '<1x/>', '<!x>', '<?xml version="1.0"?>', '</>', '<a/ >', '<e f="1"g="2"/>', `&${'n'.repeat(33)};`, `&${'n'.repeat(32)};`,
  ]

  function makeDoc(rand: () => number): string {
    const n = 2 + Math.floor(rand() * 12)
    let body = ''
    for (let i = 0; i < n; i++) body += POOL[Math.floor(rand() * POOL.length)]
    return rand() < 0.85 ? `<r>${body}</r>` : body
  }

  it(`random valid/malformed documents give the same outcome for every feeding (${LARGE ? 'full' : 'reduced'})`, () => {
    const rand = createPrng(0x5eed)
    const docs = LARGE ? 20_000 : 250
    let comparisons = 0
    let wellFormed = 0
    const codes = new Map<string, number>()
    for (let d = 0; d < docs; d++) {
      const doc = makeDoc(rand)
      for (const limits of [undefined, { maxTextSegmentChars: 3 }]) {
        const base = key(outcomeOf(doc, [], limits))
        if (limits === undefined) {
          if (base.startsWith('error:')) codes.set(base, (codes.get(base) ?? 0) + 1)
          else wellFormed++
        }
        const variants: number[][] = [everyN(doc, 1)]
        for (let r = 0; r < 3; r++) variants.push(randomCuts(doc, rand, 7))
        for (let i = 1; i < doc.length; i++) variants.push([i])
        if (LARGE && doc.length <= 48) {
          for (let i = 1; i < doc.length; i++) for (let j = i + 1; j < doc.length; j++) variants.push([i, j])
        } else {
          for (let s = 0; s < (LARGE ? 300 : 40); s++) {
            const i = 1 + Math.floor(rand() * (doc.length - 1))
            const j = 1 + Math.floor(rand() * (doc.length - 1))
            if (i !== j) variants.push([Math.min(i, j), Math.max(i, j)])
          }
        }
        for (const cuts of variants) {
          const k = key(outcomeOf(doc, cuts, limits))
          if (k !== base) expect({ doc, cuts, limits, got: k }).toEqual({ doc, cuts, limits, got: base })
          comparisons++
        }
      }
    }
    // the pool must exercise both outcomes
    expect(wellFormed).toBeGreaterThan(docs / 50)
    expect(codes.get('error:parse-error') ?? 0).toBeGreaterThan(docs / 10)
    expect(codes.get('error:unsupported') ?? 0).toBeGreaterThan(0)
    expect(comparisons).toBeGreaterThan(docs * 20)
    if (LARGE) console.info(`tokenizer fuzz: ${docs} docs, ${comparisons} comparisons, ${wellFormed} well-formed`)
  }, LARGE ? 3_600_000 : undefined)

  it('exhaustive two-cut splits of a document with escapes, an astral character and a lone-surrogate escape', () => {
    const doc = '<sst><si><t>a_x000D_b\r\nc_x005F_x0041_😀_xD800_&amp;&#65;</t></si></sst>'
    const base = key(outcomeOf(doc, []))
    for (let i = 1; i < doc.length; i++) {
      for (let j = i + 1; j < doc.length; j++) expect(key(outcomeOf(doc, [i, j]))).toBe(base)
    }
  })
})

describe('malformed tokens (RED test 10)', () => {
  it.each<[string, string, string]>([
    ['bare &', '<r>a & b</r>', 'parse-error'],
    ['bare & in an attribute', '<r a="x&y"/>', 'parse-error'],
    ['&amp without ;', '<r>&amp</r>', 'parse-error'],
    ['&#;', '<r>&#;</r>', 'parse-error'],
    ['&#x;', '<r>&#x;</r>', 'parse-error'],
    ['&#xZ;', '<r>&#xZ;</r>', 'parse-error'],
    ['&#X41; (uppercase X)', '<r>&#X41;</r>', 'parse-error'],
    ['reference body of 33 characters', `<r>&${'a'.repeat(33)};</r>`, 'parse-error'],
    ['numeric reference over 32 characters', `<r>&#${'0'.repeat(32)}1;</r>`, 'parse-error'],
    ['&#xD800;', '<r>&#xD800;</r>', 'parse-error'],
    ['&#x110000;', '<r>&#x110000;</r>', 'parse-error'],
    ['&#1; (not an XML character)', '<r>&#1;</r>', 'parse-error'],
    ['numeric overflow', `<r>&#${'9'.repeat(30)};</r>`, 'parse-error'],
    ['duplicate attribute', '<r a="1" a="2"/>', 'parse-error'],
    ['unquoted attribute', '<r a=1/>', 'parse-error'],
    ['attribute without value', '<r a/>', 'parse-error'],
    ['attributes without separating space', '<r a="1"b="2"/>', 'parse-error'],
    ['unterminated comment at end of input', '<r><!-- x</r>', 'parse-error'],
    ['partial comment opener at end of input', '<r></r><!-', 'parse-error'],
    ['unterminated CDATA', '<r><![CDATA[x</r>', 'parse-error'],
    ['unterminated PI', '<r><?pi x</r>', 'parse-error'],
    ['unterminated start tag', '<r><a b="1"', 'parse-error'],
    ['unterminated end tag', '<r></r', 'parse-error'],
    ['mismatched end tag', '<r><a></b></r>', 'parse-error'],
    ['unclosed element', '<r><a></a>', 'parse-error'],
    ['text outside the root', '<r/>junk', 'parse-error'],
    ['second root', '<r/><s/>', 'parse-error'],
    ['empty document', '  ', 'parse-error'],
    ['CDATA outside the root', '<![CDATA[x]]><r/>', 'parse-error'],
    ['XML declaration not at the start', ' <?xml version="1.0"?><r/>', 'parse-error'],
    ['invalid element name', '<1r/>', 'parse-error'],
    ['unknown <! markup', '<r><!x></r>', 'parse-error'],
    ['DOCTYPE', '<!DOCTYPE r><r/>', 'unsupported'],
    ['DOCTYPE with an internal subset', '<!DOCTYPE r [<!ENTITY e "v">]><r>&e;</r>', 'unsupported'],
  ])('%s -> %s for every chunking', (_name, doc, code) => {
    const o = expectSameEverywhere(doc)
    expect(errorCode(o)).toBe(code)
  })

  it('duplicate attributes after prefix stripping are detected by the consumer helper', () => {
    const rec = recorder()
    const seen: string[] = []
    const ev: XmlEvents = {
      ...rec.handler,
      open(name, names, values, count, sc) {
        rec.handler.open(name, names, values, count, sc)
        const dup = findLocalNameCollision(names, count)
        if (dup !== undefined) seen.push(dup)
        if (name === 'bad') assertUniqueLocalAttributes(names, count)
      },
    }
    const tok = new XmlTokenizer(ev)
    tok.write('<r xmlns:x="u" xmlns:y="v" x:a="1" b="2"><ok xmlns="w" a="1"/>')
    expect(() => tok.write('<bad x:k="1" y:k="2"/></r>')).toThrow(PackageError)
    expect(seen).toEqual(['k'])
    expect(localName('x:sheetData')).toBe('sheetData')
    expect(localName('row')).toBe('row')
  })

  it('a failed tokenizer keeps failing with the same error', () => {
    const tok = new XmlTokenizer(recorder().handler)
    expect(() => tok.write('<r>&</r>')).toThrow(/reference/)
    expect(() => tok.write('<r/>')).toThrow(/reference/)
    expect(() => tok.end()).toThrow(/reference/)
  })
})

describe('value fidelity (RED test 11)', () => {
  function decodeAll(segments: string[]): string {
    const d = new XstringDecoder()
    let out = ''
    segments.forEach((s, i) => { out += d.push(s, i === segments.length - 1) })
    return out
  }

  it('ST_Xstring decoding: lone surrogates, astral, _x000D_ keeps CR, _x005F_x0041_ is literal _x0041_', () => {
    expect(decodeAll(['_xD800_-_xDE00_😀'])).toBe('\uD800-\uDE00😀')
    expect(decodeAll(['a_x000D_b'])).toBe('a\rb')
    expect(decodeAll(['_x005F_x0041_'])).toBe('_x0041_')
    expect(decodeAll(['_x00', '41_', '_x', '_x0042_'])).toBe('A_xB')
  })

  it('XstringDecoder equals decodeXlsxString for every two-cut chunking', () => {
    const samples = [
      '_x005F_x0041__x000D__xD800__x_x0041x_x0041_',
      '__x0041___x00_x0041_',
      ...Array.from({ length: 16 }, (_, i) => encodeXlsxString(f4StringFor(FIXTURE_SEEDS.F4, i)).replace(/&amp;/g, '&')),
      encodeXlsxString('a\rb\u0001c_x0041_\uDE00'),
    ]
    for (const s of samples) {
      const expected = decodeXlsxString(s)
      for (let i = 0; i <= s.length; i++) {
        for (let j = i; j <= s.length; j++) expect(decodeAll([s.slice(0, i), s.slice(i, j), s.slice(j)])).toBe(expected)
      }
    }
  })

  it('Utf8ChunkDecoder: split multi-byte sequences, BOM stripped, UTF-16 BOM refused', () => {
    const bytes = new TextEncoder().encode('﻿é😀 ok')
    for (let cut = 0; cut <= bytes.length; cut++) {
      const d = new Utf8ChunkDecoder()
      expect(d.decode(bytes.subarray(0, cut)) + d.decode(bytes.subarray(cut), true)).toBe('é😀 ok')
      expect(d.replacementCount).toBe(0)
    }
    for (const bom of [[0xff, 0xfe], [0xfe, 0xff]]) {
      const d = new Utf8ChunkDecoder()
      const err = (() => { try { d.decode(Uint8Array.of(bom[0])); d.decode(Uint8Array.of(bom[1], 0x3c, 0)); return null } catch (e) { return e } })()
      expect((err as PackageError).code).toBe('unsupported')
    }
  })

  it('invalid UTF-8 decodes with replacement characters plus a count; a genuine U+FFFD is not counted; fatal mode is a parse-error', () => {
    // 0xC3 0x28 invalid, EF BF BD genuine U+FFFD split across chunks, lone 0xFF
    const parts = [Uint8Array.of(0x61, 0xc3, 0x28, 0xef), Uint8Array.of(0xbf), Uint8Array.of(0xbd, 0xff, 0x62)]
    const d = new Utf8ChunkDecoder()
    const out = parts.map((p, i) => d.decode(p, i === parts.length - 1)).join('')
    expect(out).toBe('a�(��b')
    expect(d.replacementCount).toBe(2)
    const f = new Utf8ChunkDecoder({ fatal: true })
    expect(() => parts.forEach((p, i) => f.decode(p, i === parts.length - 1))).toThrow(PackageError)
  })

  it('lone surrogates in input strings survive the tokenizer unchanged (UTF-16 code units, no lossy step)', () => {
    const doc = '<r>\uD800x\uDE00</r>'
    expect(textOfOutcome(expectSameEverywhere(doc))).toBe('\uD800x\uDE00')
  })
})

describe('bounds (RED test 12)', () => {
  it('a token over maxTokenChars fails with resource-limit after at most maxTokenChars plus one chunk', () => {
    for (const piece of [1, 7, 100]) {
      const tok = new XmlTokenizer(recorder().handler, { maxTokenChars: 1000 })
      tok.write('<r><a b="')
      let fed = 9
      let error: unknown
      for (let i = 0; i < 5000 && error === undefined; i++) {
        try {
          tok.write('x'.repeat(piece))
          fed += piece
          expect(tok.bufferedChars).toBeLessThanOrEqual(1000)
        } catch (e) { error = e }
      }
      expect((error as PackageError).code).toBe('resource-limit')
      expect(fed).toBeLessThanOrEqual(1000 + piece + 9)
    }
  })

  it('an unterminated comment fails with resource-limit, not after the whole input', () => {
    const tok = new XmlTokenizer(recorder().handler, { maxTokenChars: 4096 })
    tok.write('<r><!--')
    expect(() => { for (let i = 0; i < 100; i++) tok.write('y'.repeat(1024)) }).toThrow(/token/)
  })

  it('a giant text node arrives in bounded segments while the tokenizer holds no more than a reference', () => {
    let segments = 0
    let maxSeg = 0
    let total = 0
    const ev: XmlEvents = { open() {}, close() {}, text(t) { segments++; total += t.length; maxSeg = Math.max(maxSeg, t.length) } }
    const tok = new XmlTokenizer(ev, { maxTextSegmentChars: 1000 })
    tok.write('<r>')
    const chunk = 'abcdefghij'.repeat(3277) + '&amp;'
    let maxBuffered = 0
    for (let i = 0; i < 200; i++) { tok.write(chunk); maxBuffered = Math.max(maxBuffered, tok.bufferedChars) }
    tok.write('</r>')
    tok.end()
    expect(total).toBe(200 * (32_770 + 1))
    expect(maxSeg).toBeLessThanOrEqual(1000)
    expect(segments).toBeGreaterThanOrEqual(6554)
    expect(maxBuffered).toBeLessThanOrEqual(33)
  })

  it('depth and attribute-count caps are resource-limit', () => {
    expect(errorCode(outcomeOf('<a>'.repeat(5) + '</a>'.repeat(5), [], { maxDepth: 4 }))).toBe('resource-limit')
    expect(errorCode(outcomeOf('<a>'.repeat(4) + '</a>'.repeat(4), [], { maxDepth: 4 }))).toBe('ok')
    const attrs = (n: number) => `<r ${Array.from({ length: n }, (_, i) => `a${i}="${i}"`).join(' ')}/>`
    expect(errorCode(outcomeOf(attrs(128), []))).toBe('ok')
    expect(errorCode(expectSameEverywhere(attrs(129)))).toBe('resource-limit')
  })
})

describe('parity with the current parser on the snippet corpus (RED test 13)', () => {
  /** Current path: parseXml + parse.ts sharedStrings() logic. */
  function legacyStrings(xml: string): string[] {
    const root = parseXml(xml)
    const out: string[] = []
    for (const si of getChildren(root, 'si')) {
      let s = ''
      const direct = si['t']
      if (direct !== undefined) s += decodeXlsxString(textOf(typeof direct === 'object' ? direct : { '#text': String(direct ?? '') } as XmlNode))
      for (const r of getChildren(si, 'r')) s += decodeXlsxString(textOf(getChildren(r, 't')[0]))
      out.push(s)
    }
    return out
  }
  function streamStrings(xml: string, cuts: number[]): string[] {
    const c = sharedStringsConsumer()
    const tok = new XmlTokenizer(c.handler)
    let at = 0
    for (const cut of cuts) { tok.write(xml.slice(at, cut)); at = cut }
    tok.write(xml.slice(at))
    tok.end()
    return c.strings
  }
  const SNIPPETS: Array<[string, string]> = [
    ['plain', '<t>hello</t>'],
    ['ST_Xstring escapes', '<t>a_x000D_b_x0009_c</t>'],
    ['escaped underscore', '<t>_x005F_x0041_</t>'],
    ['lone surrogate escapes', '<t>_xD800_-_xDE00_</t>'],
    ['predefined entities', '<t>&lt;&gt;&amp;&quot;&apos;</t>'],
    ['numeric references', '<t>&#65;&#x42;&#x1F600;&#9;</t>'],
    ['leading-zero reference', '<t>&#00000000000000065;</t>'],
    ['literal CRLF and CR', '<t>a\r\nb\rc</t>'],
    ['&#13; keeps CR', '<t>a&#13;b</t>'],
    ['rich runs', '<r><t>bold</t></r><r><rPr><b/></rPr><t xml:space="preserve"> text</t></r>'],
    ['phonetic run ignored', '<t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh>'],
    ['CDATA', '<t><![CDATA[<x>&]]></t>'],
    ['comment inside text', '<t>a<!-- c -->b</t>'],
    ['namespace prefixes', '<x:t xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">p</x:t>'],
    ['empty text', '<t/>'],
    ['whitespace preserved', '<t xml:space="preserve">  two  </t>'],
    ['attribute containing >', '<t foo="a>b">v</t>'],
    ['unknown entity stays literal', '<t>&foo;</t>'],
  ]
  it.each(SNIPPETS)('%s', (_name, si) => {
    const xml = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si>${si}</si><si><t>z</t></si></sst>`
    const expected = legacyStrings(xml)
    for (const cuts of [[], everyN(xml, 1), everyN(xml, 5), everyN(xml, 1_048_576)]) expect(streamStrings(xml, cuts)).toEqual(expected)
  })

  it('an invalid numeric reference fails on both paths', () => {
    const xml = '<sst><si><t>&#xD800;</t></si></sst>'
    expect(() => legacyStrings(xml)).toThrow()
    expect(() => streamStrings(xml, [])).toThrow(PackageError)
  })
})

describe('tokenizeBytes: cancellation and backpressure hooks', () => {
  async function* chunks(parts: string[]): AsyncGenerator<Uint8Array> {
    for (const p of parts) yield new TextEncoder().encode(p)
  }

  it('awaits afterChunk between chunks (the backpressure point) and decodes UTF-8 across chunk boundaries', async () => {
    const order: string[] = []
    const c = sharedStringsConsumer()
    const bytes = new TextEncoder().encode('<sst><si><t>é😀x</t></si></sst>')
    async function* split(): AsyncGenerator<Uint8Array> { for (let i = 0; i < bytes.length; i += 3) { order.push(`chunk${i}`); yield bytes.subarray(i, i + 3) } }
    let pending = 0
    const result = await tokenizeBytes(split(), c.handler, {
      async afterChunk() { pending++; order.push('drain'); await Promise.resolve(); pending-- },
    })
    expect(c.strings).toEqual(['é😀x'])
    expect(result.replacementCount).toBe(0)
    expect(pending).toBe(0)
    expect(order.filter((s) => s === 'drain').length).toBe(Math.ceil(bytes.length / 3))
    expect(order[1]).toBe('drain')
  })

  it('an aborted signal stops tokenizing with cancelled', async () => {
    const ac = new AbortController()
    const err = await tokenizeBytes(chunks(['<r>', 'a', 'b', '</r>']), recorder().handler, {
      signal: ac.signal,
      afterChunk() { ac.abort() },
    }).catch((e: unknown) => e)
    expect((err as PackageError).code).toBe('cancelled')
  })

  it('reports replacement characters as a counted diagnostic', async () => {
    const res = await tokenizeBytes((async function* () { yield Uint8Array.of(0x3c, 0x72, 0x3e, 0xff, 0x3c, 0x2f, 0x72, 0x3e) })(), recorder().handler)
    expect(res.replacementCount).toBe(1)
  })
})
