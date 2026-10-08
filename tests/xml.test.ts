import { describe, test, expect } from 'vitest'
import { parseXml, parseXmlOrdered, getChildren, attrs, textOf } from '../src/core/xml'

describe('xml normalize', () => {
  test('strips namespaces and lifts attributes', () => {
    const xml = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p w14:paraId="1">
      <w:pPr><w:jc w:val="center"/><w:spacing w:after="200" w:line="240"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Hello</w:t></w:r>
      <w:r><w:t xml:space="preserve"> world</w:t></w:r>
      <w:r><w:t>more</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
    const doc = parseXml(xml)
    const body = doc['body'] as never
    const paragraphs = getChildren(body as never, 'p')
    expect(paragraphs).toHaveLength(1)
    const p = paragraphs[0]
    expect(attrs(p)['w14:paraId'] ?? attrs(p)['paraId']).toBeDefined()
    const jc = getChildren(getChildren(p, 'pPr')[0], 'jc')[0]
    expect(attrs(jc).val).toBe('center')
    const runs = getChildren(p, 'r')
    expect(runs).toHaveLength(3)
    expect(textOf(runs[0])).toBe('Hello')
    expect(textOf(runs[1])).toBe(' world')
    expect(textOf(runs[2])).toBe('more')
    const rPr = getChildren(runs[0], 'rPr')[0]
    expect(attrs(getChildren(rPr, 'sz')[0]).val).toBe('28')
  })
})

describe('real-world XML edge cases', () => {
  test('prolog and trailing whitespace do not break root selection', () => {
    // PowerPoint emits a BOM, a prolog, and trailing newline after </Types>
    const xml = '\uFEFF<?xml version="1.0" encoding="utf-8"?>\n<Types xmlns="urn:x">\n  <Default a="1"/>\n</Types>\n'
    const doc = parseXml(xml)
    expect(attrs(getChildren(doc as never, 'Default')[0]).a).toBe('1')
  })
})

describe('XML numeric references and entity decoding', () => {
  test('decodes valid decimal and hex numeric references in attributes and text across parseXml and parseXmlOrdered', () => {
    const xml = `<root a="A&#10;B&#x9;C&#13;D" b="A&#xA;B" zwj="&#x1F469;&#8205;&#x1F4BB;">
      <elem>A&#10;B&#x9;C&#13;D</elem>
      <elemHex>A&#xA;B</elemHex>
      <elemZWJ>&#x1F469;&#8205;&#x1F4BB;</elemZWJ>
    </root>`

    const doc = parseXml(xml)
    expect(attrs(doc as any).a).toBe('A\nB\tC\rD')
    expect(attrs(doc as any).b).toBe('A\nB')
    expect(attrs(doc as any).zwj).toBe('👩‍💻')
    expect(textOf(getChildren(doc as any, 'elem')[0])).toBe('A\nB\tC\rD')
    expect(textOf(getChildren(doc as any, 'elemHex')[0])).toBe('A\nB')
    expect(textOf(getChildren(doc as any, 'elemZWJ')[0])).toBe('👩‍💻')

    const ordered = parseXmlOrdered(xml)
    expect(attrs(ordered).a).toBe('A\nB\tC\rD')
    expect(attrs(ordered).b).toBe('A\nB')
    expect(attrs(ordered).zwj).toBe('👩‍💻')
    expect(textOf(getChildren(ordered, 'elem')[0])).toBe('A\nB\tC\rD')
    expect(textOf(getChildren(ordered, 'elemHex')[0])).toBe('A\nB')
    expect(textOf(getChildren(ordered, 'elemZWJ')[0])).toBe('👩‍💻')
  })

  test('retains standard XML reserved entities and leaves HTML aliases unexpanded', () => {
    const xml = `<root reserved="&amp;&lt;&gt;&apos;&quot;" html="&nbsp;&copy;">
      <reserved>&amp;&lt;&gt;&apos;&quot;</reserved>
      <htmlNode>&nbsp;&copy;</htmlNode>
    </root>`

    const doc = parseXml(xml)
    expect(attrs(doc as any).reserved).toBe('&<>\'"')
    expect(attrs(doc as any).html).toBe('&nbsp;&copy;')
    expect(textOf(getChildren(doc as any, 'reserved')[0])).toBe('&<>\'"')
    expect(textOf(getChildren(doc as any, 'htmlNode')[0])).toBe('&nbsp;&copy;')

    const ordered = parseXmlOrdered(xml)
    expect(attrs(ordered).reserved).toBe('&<>\'"')
    expect(attrs(ordered).html).toBe('&nbsp;&copy;')
    expect(textOf(getChildren(ordered, 'reserved')[0])).toBe('&<>\'"')
    expect(textOf(getChildren(ordered, 'htmlNode')[0])).toBe('&nbsp;&copy;')
  })

  test('escaped amp;#10; decodes once only without manufactured newline', () => {
    const xml = `<root a="&amp;#10;"><elem>&amp;#10;</elem></root>`

    const doc = parseXml(xml)
    expect(attrs(doc as any).a).toBe('&#10;')
    expect(textOf(getChildren(doc as any, 'elem')[0])).toBe('&#10;')

    const ordered = parseXmlOrdered(xml)
    expect(attrs(ordered).a).toBe('&#10;')
    expect(textOf(getChildren(ordered, 'elem')[0])).toBe('&#10;')
  })

  test('illegal numeric XML scalars (0, surrogates, FFFE, FFFF, >10FFFF) must throw in both parsers', () => {
    const illegalCases = [
      '<root a="&#0;"/>',
      '<root>&#0;</root>',
      '<root a="&#x0;"/>',
      '<root>&#x0;</root>',
      '<root a="&#xD800;"/>',
      '<root>&#xD800;</root>',
      '<root a="&#55296;"/>',
      '<root>&#57343;</root>',
      '<root a="&#xDFFF;"/>',
      '<root a="&#xFFFE;"/>',
      '<root>&#65534;</root>',
      '<root a="&#xFFFF;"/>',
      '<root>&#65535;</root>',
      '<root a="&#1114112;"/>',
      '<root>&#x110000;</root>',
    ]

    for (const xml of illegalCases) {
      expect(() => parseXml(xml)).toThrow()
      expect(() => parseXmlOrdered(xml)).toThrow()
    }
  })

  test('legal supplementary noncharacters are accepted', () => {
    const xml = `<root a="&#x1FFFE;"><elem>&#x10FFFF;</elem></root>`
    const doc = parseXml(xml)
    expect(attrs(doc as any).a).toBe(String.fromCodePoint(0x1FFFE))
    expect(textOf(getChildren(doc as any, 'elem')[0])).toBe(String.fromCodePoint(0x10FFFF))

    const ordered = parseXmlOrdered(xml)
    expect(attrs(ordered).a).toBe(String.fromCodePoint(0x1FFFE))
    expect(textOf(getChildren(ordered, 'elem')[0])).toBe(String.fromCodePoint(0x10FFFF))
  })

  test('reference-looking CDATA, comments, PI and escaped amp;#0; stay literal/ignored', () => {
    const xml = `<root a="&amp;#0;">
      <cdata><![CDATA[&#0;&#xD800;]]></cdata>
      <!-- &#0;&#xFFFF; -->
      <?pi &#0;&#xD800;?>
    </root>`

    const doc = parseXml(xml)
    expect(attrs(doc as any).a).toBe('&#0;')
    expect(textOf(getChildren(doc as any, 'cdata')[0])).toBe('&#0;&#xD800;')

    const ordered = parseXmlOrdered(xml)
    expect(attrs(ordered).a).toBe('&#0;')
    expect(textOf(getChildren(ordered, 'cdata')[0])).toBe('&#0;&#xD800;')
  })
})
