import { describe, test, expect } from 'vitest'
import { parseXml, getChildren, attrs, textOf } from '../src/core/xml'

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
