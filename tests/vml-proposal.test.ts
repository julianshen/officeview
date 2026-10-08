import { describe, expect, test } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseAllVmlWordArt } from '../src/drawing/vml'

const GROUP_XML = (inner: string) =>
  parseXmlOrdered(`<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000">${inner}</v:group></xml>`)

describe('VML proposal RED: grouped transforms', () => {
  test('unitless-group child maps to (82,46,80,30)pt', async () => {
    const { parseVmlContainer } = await import('../src/drawing/vml') as any
    expect(typeof parseVmlContainer).toBe('function')
    const root = GROUP_XML(`<v:shape id="Grouped" style="position:absolute;left:150;top:400;width:400;height:600;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Grouped"/></v:shape>`)
    const tree = parseVmlContainer(root)
    const leaves: any[] = []
    const walk = (n: any) => { if (n.kind === 'shape') leaves.push(n); else (n.children ?? []).forEach(walk) }
    tree.nodes.forEach(walk)
    expect(leaves.map((l: any) => l.result.textBody.paragraphs[0].runs[0].text)).toEqual(['Grouped'])
    // Hierarchical group must be preserved, not flattened
    expect(tree.nodes[0].kind).toBe('group')
    expect(tree.nodes[0].children.length).toBe(1)
  })

  test('nested-group maps to (92,46,20,10)pt', async () => {
    const { parseVmlContainer } = await import('../src/drawing/vml') as any
    const root = parseXmlOrdered(`<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000"><v:group id="inner" style="position:absolute;left:100;top:200;width:500;height:1000;" coordorigin="-100,-200" coordsize="500,1000"><v:shape id="Nested" style="position:absolute;left:0;top:0;width:100;height:200;"><v:textpath on="t" style="font-family:Arial;font-size:24pt;" string="Nested"/></v:shape></v:group></v:group></xml>`)
    const tree = parseVmlContainer(root)
    expect(tree.nodes[0].kind).toBe('group')
    expect(tree.nodes[0].children[0].kind).toBe('group')
  })

  test('mixed-source-order retains Before/Grouped/Nested/After', async () => {
    const { parseVmlContainer } = await import('../src/drawing/vml') as any
    const root = parseXmlOrdered(`<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="Before" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;"><v:textpath on="t" string="Before"/></v:shape><v:group id="group" style="position:absolute;left:72pt;top:36pt;width:200pt;height:100pt;" coordorigin="100,200" coordsize="1000,2000"><v:shape id="Grouped" style="position:absolute;left:150;top:400;width:400;height:600;"><v:textpath on="t" string="Grouped"/></v:shape><v:group id="inner" style="position:absolute;left:100;top:200;width:500;height:1000;" coordorigin="-100,-200" coordsize="500,1000"><v:shape id="Nested" style="position:absolute;left:0;top:0;width:100;height:200;"><v:textpath on="t" string="Nested"/></v:shape></v:group></v:group><v:shape id="After" style="position:absolute;left:0pt;top:0pt;width:160pt;height:40pt;"><v:textpath on="t" string="After"/></v:shape></xml>`)
    const tree = parseVmlContainer(root)
    const texts: string[] = []
    const walk = (n: any) => { if (n.kind === 'shape') texts.push(n.result.textBody.paragraphs[0].runs[0].text); else (n.children ?? []).forEach(walk) }
    tree.nodes.forEach(walk)
    expect(texts).toEqual(['Before', 'Grouped', 'Nested', 'After'])
  })

  test('legacy direct parseAll retains direct shapes', () => {
    const root = parseXmlOrdered(`<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="A" style="width:160pt;height:40pt;"><v:textpath on="t" string="A"/></v:shape></xml>`)
    expect(parseAllVmlWordArt(root as any).length).toBeGreaterThan(0)
  })
})

describe('VML proposal RED: template and typography', () => {
  test('partial-template-override inherits Arial/center/activation, clears bold/italic/outline', async () => {
    const { parseVmlWordArtWithTemplates } = await import('../src/drawing/vml') as any
    expect(typeof parseVmlWordArtWithTemplates).toBe('function')
  })
  test('cm units convert completely', async () => {
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = parseXmlOrdered(`<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="C" style="position:absolute;left:2.54cm;top:1.27cm;width:5.08cm;height:2.54cm;"><v:textpath on="t" style="font-family:Arial;font-size:1cm" string="Centimeter"/></v:shape></xml>`)
    const r = parseVmlWordArt(root as any)
    expect(r?.leftPt).toBeCloseTo(72, 5)
    expect(r?.topPt).toBeCloseTo(36, 5)
    expect(r?.widthPt).toBeCloseTo(144, 5)
    expect(r?.heightPt).toBeCloseTo(72, 5)
    expect(r?.textBody.paragraphs[0].runs[0].fontSizePt).toBeCloseTo(28.346456692913385, 5)
  })
  test('font shorthand parses family/size/bold/italic', async () => {
    const { parseVmlWordArt } = await import('../src/drawing/vml')
    const root = parseXmlOrdered(`<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="S" style="width:160pt;height:40pt;"><v:textpath on="t" style="font:italic bold 24pt/1.2 &quot;Times New Roman&quot;, serif" string="Shorthand"/></v:shape></xml>`)
    const r = parseVmlWordArt(root as any)
    expect(r?.textBody.paragraphs[0].runs[0].fontFamily).toBe('Times New Roman')
    expect(r?.textBody.paragraphs[0].runs[0].fontSizePt).toBe(24)
    expect(r?.textBody.paragraphs[0].runs[0].bold).toBe(true)
    expect(r?.textBody.paragraphs[0].runs[0].italic).toBe(true)
  })
})
