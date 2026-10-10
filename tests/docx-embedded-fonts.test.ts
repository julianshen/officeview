// @vitest-environment node
import { readFileSync } from 'node:fs'
import { describe, expect, test, vi } from 'vitest'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { parseXml } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parseDocxFontDeclarations, deobfuscateOdttf, loadDocxEmbeddedFonts } from '../src/core/fonts/docx'
import { decodeEmbeddedFont, FONT_LIMITS } from '../src/core/fonts/decode'
import { acquireFonts, type FontRegistrationRequest, type RegisterFont } from '../src/core/fonts/register'
import { parseDocx } from '../src/docx/parse'
import type { DocxDocument } from '../src/docx/types'
import { getPaintables } from '../src/render/paint'
import { eot } from './fonts-decode.test'

describe('DOCX embedded fonts', () => {
  const FONT_TABLE_XML = `
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="CustomSerif">
    <w:embedRegular r:id="rId1" w:fontKey="{12345678-ABCD-EF01-2345-6789ABCDEF01}"/>
    <w:embedBold r:id="rId2" w:fontKey="{87654321-DCBA-10FE-5432-10FEDCBA9876}"/>
    <w:embedItalic r:id="rId3" w:fontKey="{A1B2C3D4-E5F6-7890-1234-56789ABCDEF0}"/>
    <w:embedBoldItalic r:id="rId4" w:fontKey="{F0E1D2C3-B4A5-0987-6543-210FEDCBA987}"/>
  </w:font>
  <w:font w:name="CustomSans">
    <w:embedRegular r:id="rId5" w:fontKey="{99999999-8888-7777-6666-555555555555}"/>
  </w:font>
</w:fonts>
`

  const FONT_TABLE_RELS_XML = `
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font2.odttf"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font3.odttf"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font4.odttf"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font5.odttf"/>
</Relationships>
`

  test('parses word/fontTable.xml and relationships to locate embedded font parts and fontKey GUIDs', () => {
    const fontTableNode = parseXml(FONT_TABLE_XML)
    const relsNode = parseXml(FONT_TABLE_RELS_XML)

    const declarations = parseDocxFontDeclarations(fontTableNode, relsNode, 'word/fontTable.xml')

    expect(declarations).toHaveLength(5)
    expect(declarations[0]).toEqual({
      family: 'CustomSerif',
      variant: 'regular',
      relationshipId: 'rId1',
      fontKey: '{12345678-ABCD-EF01-2345-6789ABCDEF01}',
      partPath: 'word/fonts/font1.odttf',
    })
    expect(declarations[1]).toEqual({
      family: 'CustomSerif',
      variant: 'bold',
      relationshipId: 'rId2',
      fontKey: '{87654321-DCBA-10FE-5432-10FEDCBA9876}',
      partPath: 'word/fonts/font2.odttf',
    })
    expect(declarations[2]).toEqual({
      family: 'CustomSerif',
      variant: 'italic',
      relationshipId: 'rId3',
      fontKey: '{A1B2C3D4-E5F6-7890-1234-56789ABCDEF0}',
      partPath: 'word/fonts/font3.odttf',
    })
    expect(declarations[3]).toEqual({
      family: 'CustomSerif',
      variant: 'boldItalic',
      relationshipId: 'rId4',
      fontKey: '{F0E1D2C3-B4A5-0987-6543-210FEDCBA987}',
      partPath: 'word/fonts/font4.odttf',
    })
    expect(declarations[4]).toEqual({
      family: 'CustomSans',
      variant: 'regular',
      relationshipId: 'rId5',
      fontKey: '{99999999-8888-7777-6666-555555555555}',
      partPath: 'word/fonts/font5.odttf',
    })
  })

  test('handles missing or external relationships gracefully', () => {
    const tableWithExternal = `
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="ExternalFont">
    <w:embedRegular r:id="rIdExt" w:fontKey="{11111111-2222-3333-4444-555555555555}"/>
    <w:embedBold r:id="rIdMissing"/>
  </w:font>
</w:fonts>
`
    const relsWithExternal = `
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdExt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="https://example.com/font.ttf" TargetMode="External"/>
</Relationships>
`
    const declarations = parseDocxFontDeclarations(parseXml(tableWithExternal), parseXml(relsWithExternal), 'word/fontTable.xml')
    expect(declarations).toHaveLength(2)
    expect(declarations[0].isExternal).toBe(true)
    expect(declarations[0].partPath).toBeUndefined()
    expect(declarations[1].partPath).toBeUndefined()
  })

  test('ODTTF deobfuscation correctly decrypts the 32-byte header using 16-byte reversed GUID XOR key', () => {
    // 64 bytes of mock font data
    const original = new Uint8Array(64)
    for (let i = 0; i < 64; i++) original[i] = (i * 7 + 13) & 0xff
    // Standard SFNT header start: 0x00010000
    original[0] = 0x00
    original[1] = 0x01
    original[2] = 0x00
    original[3] = 0x00

    const fontKey = '{12345678-ABCD-EF01-2345-6789ABCDEF01}'
    // 16 bytes hex in string order:
    // [0x12, 0x34, 0x56, 0x78, 0xab, 0xcd, 0xef, 0x01, 0x23, 0x45, 0x67, 0x89, 0xab, 0xcd, 0xef, 0x01]
    // Reversed key (ECMA-376):
    // [0x01, 0xef, 0xcd, 0xab, 0x89, 0x67, 0x45, 0x23, 0x01, 0xef, 0xcd, 0xab, 0x78, 0x56, 0x34, 0x12]
    const reversedKey = new Uint8Array([
      0x01, 0xef, 0xcd, 0xab, 0x89, 0x67, 0x45, 0x23,
      0x01, 0xef, 0xcd, 0xab, 0x78, 0x56, 0x34, 0x12,
    ])

    const obfuscated = original.slice()
    for (let i = 0; i < 32; i++) {
      obfuscated[i] ^= reversedKey[i % 16]
    }

    // Obfuscated data is altered in the first 32 bytes and identical thereafter
    expect(obfuscated.slice(0, 32)).not.toEqual(original.slice(0, 32))
    expect(obfuscated.slice(32)).toEqual(original.slice(32))

    // Deobfuscation restores the original bytes
    const deobfuscated = deobfuscateOdttf(obfuscated, fontKey)
    expect(deobfuscated).toEqual(original)

    // Symmetric property: deobfuscating the original with fontKey yields the obfuscated data
    expect(deobfuscateOdttf(original, fontKey)).toEqual(obfuscated)

    // Strips curly braces or handles plain GUID string identically
    const plainGuid = '12345678-ABCD-EF01-2345-6789ABCDEF01'
    expect(deobfuscateOdttf(obfuscated, plainGuid)).toEqual(original)
  })

  test('ODTTF deobfuscation rejects invalid or malformed GUID keys safely', () => {
    const data = new Uint8Array(64)
    expect(() => deobfuscateOdttf(data, 'invalid-guid')).toThrow(/guid/i)
    expect(() => deobfuscateOdttf(data, '')).toThrow(/guid/i)
  })

  test('real TrueType font obfuscated as ODTTF restores and decodes cleanly via decodeEmbeddedFont', () => {
    const literBytes = new Uint8Array(readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
    const guid = '{A5B6C7D8-1122-3344-5566-778899AABBCC}'

    // Obfuscating the font corrupts the SFNT header so decodeEmbeddedFont rejects it
    const obfuscated = deobfuscateOdttf(literBytes, guid)
    expect(() => decodeEmbeddedFont(obfuscated)).toThrow()

    // Deobfuscating restores the original font, which decodes and validates cleanly
    const restored = deobfuscateOdttf(obfuscated, guid)
    expect(restored).toEqual(literBytes)
    const decoded = decodeEmbeddedFont(restored)
    expect(decoded).toEqual(literBytes)
  })

  test('loadDocxEmbeddedFonts decodes deobfuscated ODTTF fonts with accurate metadata', async () => {
    const literBytes = new Uint8Array(readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
    const guid = '{12345678-ABCD-EF01-2345-6789ABCDEF01}'
    const obfuscated = deobfuscateOdttf(literBytes, guid)

    const zip = new JSZip()
    zip.file('word/fontTable.xml', `
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="Liter">
    <w:embedRegular r:id="rId1" w:fontKey="${guid}"/>
  </w:font>
</w:fonts>
`)
    zip.file('word/_rels/fontTable.xml.rels', `
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/liter.odttf"/>
</Relationships>
`)
    zip.file('word/fonts/liter.odttf', obfuscated)

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const result = await loadDocxEmbeddedFonts(pkg)

    expect(result.fontDiagnostics).toEqual([])
    expect(result.embeddedFonts).toHaveLength(1)
    expect(result.embeddedFonts[0]).toEqual({
      family: 'Liter',
      variant: 'regular',
      relationshipId: 'rId1',
      partPath: 'word/fonts/liter.odttf',
      bytes: literBytes,
    })
  })

  test('loadDocxEmbeddedFonts enforces permission and safety bounds (external, missing, malformed, restricted, limit)', async () => {
    const literBytes = new Uint8Array(readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
    const zip = new JSZip()
    zip.file('word/fontTable.xml', `
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="FontA">
    <w:embedRegular r:id="rIdOk"/>
    <w:embedBold r:id="rIdExternal"/>
    <w:embedItalic r:id="rIdMissing"/>
    <w:embedBoldItalic r:id="rIdBad"/>
  </w:font>
  <w:font w:name="FontB">
    <w:embedRegular r:id="rIdRestricted"/>
  </w:font>
</w:fonts>
`)
    zip.file('word/_rels/fontTable.xml.rels', `
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/ok.ttf"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="https://example.com/font.ttf" TargetMode="External"/>
  <Relationship Id="rIdMissing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/missing.ttf"/>
  <Relationship Id="rIdBad" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/bad.ttf"/>
  <Relationship Id="rIdRestricted" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/restricted.ttf"/>
</Relationships>
`)
    zip.file('word/fonts/ok.ttf', literBytes)
    zip.file('word/fonts/bad.ttf', new Uint8Array([1, 2, 3]))
    zip.file('word/fonts/restricted.ttf', eot(literBytes, 2))

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const result = await loadDocxEmbeddedFonts(pkg)

    expect(result.embeddedFonts).toHaveLength(1)
    expect(result.embeddedFonts[0].family).toBe('FontA')
    expect(result.embeddedFonts[0].variant).toBe('regular')

    expect(result.fontDiagnostics.map(d => d.kind)).toEqual([
      'external-font',
      'missing-font',
      'malformed-font',
      'restricted-font',
    ])
  })

  test('loadDocxEmbeddedFonts caps face enumeration at FONT_LIMITS.faces', async () => {
    const literBytes = new Uint8Array(readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
    const zip = new JSZip()
    const fontNodes: string[] = []
    const relNodes: string[] = []
    for (let i = 0; i < 65; i++) {
      fontNodes.push(`<w:font w:name="Font_${i}"><w:embedRegular r:id="rId_${i}"/></w:font>`)
      relNodes.push(`<Relationship Id="rId_${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font.ttf"/>`)
    }
    zip.file('word/fontTable.xml', `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${fontNodes.join('')}</w:fonts>`)
    zip.file('word/_rels/fontTable.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relNodes.join('')}</Relationships>`)
    zip.file('word/fonts/font.ttf', literBytes)

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const result = await loadDocxEmbeddedFonts(pkg)

    expect(result.embeddedFonts).toHaveLength(FONT_LIMITS.faces)
    expect(result.fontDiagnostics.at(-1)?.kind).toBe('font-limit')
  })

  test('parseDocx exposes embeddedFonts and diagnostics, registering faces in acquireFonts', async () => {
    const literBytes = new Uint8Array(readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
    const guid = '{A1B2C3D4-E5F6-7890-1234-56789ABCDEF0}'
    const obfuscated = deobfuscateOdttf(literBytes, guid)

    const zip = new JSZip()
    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Liter"/></w:rPr>
        <w:t>Hello embedded font</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
    )
    zip.file(
      'word/fontTable.xml',
      `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="Liter">
    <w:embedRegular r:id="rId1" w:fontKey="${guid}"/>
  </w:font>
</w:fonts>`
    )
    zip.file(
      'word/_rels/fontTable.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/liter.odttf"/>
</Relationships>`
    )
    zip.file('word/fonts/liter.odttf', obfuscated)

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const doc = await parseDocx(pkg)

    expect(doc.embeddedFonts).toBeDefined()
    expect(doc.embeddedFonts).toHaveLength(1)
    expect(doc.embeddedFonts![0].family).toBe('Liter')
    expect(doc.embeddedFonts![0].bytes).toEqual(literBytes)
    expect(doc.fontDiagnostics).toEqual([])

    const registered: Array<{ alias: string; family: string }> = []
    const lease = await acquireFonts(doc, async ({ alias, face }) => {
      registered.push({ alias, family: face.family })
    })

    expect(registered).toHaveLength(1)
    expect(registered[0].family).toBe('Liter')
    const resolved = lease.resolve('Liter')
    expect(resolved).toBe(registered[0].alias)
    expect(resolved).toMatch(/^OfficeviewFont_/)
    expect(lease.resolve('Arial')).toBe('Arial')
    lease.dispose()
  })

  test('DOCX canvas layout and text rendering measure and paint using registered embedded font aliases', async () => {
    const literBytes = new Uint8Array(readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
    const requests: FontRegistrationRequest[] = []
    const cleanup = vi.fn()
    const registerFont: RegisterFont = async request => {
      requests.push(request)
      return cleanup
    }

    const para = {
      align: 'left' as const,
      runs: [{
        text: 'Custom Embedded Font Text',
        fontFamily: 'Liter',
        fontSizePt: 16,
      }],
      images: [],
    }

    const doc: DocxDocument = {
      defaultFontFamily: 'Calibri',
      defaultFontSizePt: 11,
      styleDefaults: new Map(),
      embeddedFonts: [{
        family: 'Liter',
        variant: 'regular',
        relationshipId: 'rId1',
        partPath: 'word/fonts/liter.odttf',
        bytes: literBytes,
      }],
      fontDiagnostics: [],
      sections: [{
        margins: { topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, headerTwips: 720, footerTwips: 720, gutterTwips: 0 },
        pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' },
        paragraphs: [para],
        blocks: [{ kind: 'p', paragraph: para }],
      }],
    }

    const pages = await getPaintables(doc, { registerFont })
    expect(requests).toHaveLength(1)
    expect(requests[0].face.family).toBe('Liter')
    const alias = requests[0].alias
    expect(alias).toMatch(/^OfficeviewFont_/)

    const canvas = createCanvas(816, 1056)
    const ctx = canvas.getContext('2d')
    const fonts: string[] = []
    const fill = ctx.fillText.bind(ctx)
    ctx.fillText = function(text, x, y) {
      fonts.push(ctx.font)
      fill(text, x, y)
    }

    pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
    expect(fonts.length).toBeGreaterThan(0)
    expect(fonts.every(f => f.includes(alias))).toBe(true)

    pages.dispose()
    expect(cleanup).toHaveBeenCalledTimes(1)
  })

  test('F2: malformed font declarations do not consume the valid face budget', async () => {
    const literBytes = new Uint8Array(readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
    const zip = new JSZip()
    const fontNodes: string[] = []
    const relNodes: string[] = []
    // 5 malformed font nodes first
    for (let i = 0; i < 5; i++) {
      fontNodes.push(`<w:font><w:embedRegular r:id="badRel_${i}"/></w:font>`)
    }
    // Exactly 64 valid font nodes
    for (let i = 0; i < 64; i++) {
      fontNodes.push(`<w:font w:name="ValidFont_${i}"><w:embedRegular r:id="rId_${i}"/></w:font>`)
      relNodes.push(`<Relationship Id="rId_${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font.ttf"/>`)
    }
    zip.file('word/fontTable.xml', `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${fontNodes.join('')}</w:fonts>`)
    zip.file('word/_rels/fontTable.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relNodes.join('')}</Relationships>`)
    zip.file('word/fonts/font.ttf', literBytes)

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const result = await loadDocxEmbeddedFonts(pkg)

    // All 64 valid faces should be loaded, not curtailed by the 5 malformed ones
    expect(result.embeddedFonts).toHaveLength(64)
    expect(result.fontDiagnostics.filter(d => d.kind === 'malformed-font')).toHaveLength(5)
  })

  test('F4: nameless font entry emits a malformed-font diagnostic', async () => {
    const zip = new JSZip()
    zip.file('word/fontTable.xml', `
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font>
    <w:embedRegular r:id="rIdNameless"/>
  </w:font>
</w:fonts>
`)
    zip.file('word/_rels/fontTable.xml.rels', `
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdNameless" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/nameless.ttf"/>
</Relationships>
`)
    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const result = await loadDocxEmbeddedFonts(pkg)

    expect(result.fontDiagnostics).toHaveLength(1)
    expect(result.fontDiagnostics[0].kind).toBe('malformed-font')
  })

  test('F3: parseDocxFontDeclarations bounds parsed declarations against hostile inputs', () => {
    const fontNodes: string[] = []
    for (let i = 0; i < 200; i++) {
      fontNodes.push(`<w:font w:name="Font_${i}"><w:embedRegular r:id="rId_${i}"/></w:font>`)
    }
    const xml = `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${fontNodes.join('')}</w:fonts>`
    const decls = parseDocxFontDeclarations(parseXml(xml), undefined, 'word/fontTable.xml', 64)
    expect(decls).toHaveLength(64)
  })

  test('malformed font bytes do not consume the usable face budget', async () => {
    const zip = new JSZip()
    const malformed = Array.from({ length: FONT_LIMITS.faces }, (_, i) => `<w:font w:name="Bad_${i}"><w:embedRegular r:id="bad"/></w:font>`).join('')
    zip.file('word/fontTable.xml', `<w:fonts xmlns:w="w" xmlns:r="r">${malformed}<w:font w:name="Liter"><w:embedRegular r:id="good"/></w:font></w:fonts>`)
    zip.file('word/_rels/fontTable.xml.rels', '<Relationships><Relationship Id="bad" Type="x/font" Target="fonts/bad.ttf"/><Relationship Id="good" Type="x/font" Target="fonts/good.ttf"/></Relationships>')
    zip.file('word/fonts/bad.ttf', new Uint8Array([1, 2, 3]))
    zip.file('word/fonts/good.ttf', readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
    const result = await loadDocxEmbeddedFonts(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    expect(result.embeddedFonts.map(face => face.family)).toEqual(['Liter'])
    expect(result.fontDiagnostics).toHaveLength(FONT_LIMITS.faces)
    expect(result.fontDiagnostics.every(issue => issue.kind === 'malformed-font')).toBe(true)
  })
})
