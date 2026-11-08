/**
 * Stage-0 Task 4 — T4 font-probe tests.
 *
 * Round-10/round-9 alignment: `requiredFaces` classifies embedded faces
 * (bytes → verifiable, hash96 = 24 hex chars), bundled faces, and census-seen
 * system faces (no bytes → unverifiable ⇒ the units that use them would
 * raster per §7.4 rule 3). The embedded-face test builds a real fixture docx
 * (JSZip + ODTTF deobfuscation, the same shape as tests/docx-embedded-fonts.test.ts).
 */
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import path from 'node:path'
import { expect, it } from 'vitest'
import { deobfuscateOdttf } from '../../src/core/fonts/docx'
import { listCorpus, type CorpusFile } from '../../scripts/stream-s0/corpus'
import { hash96, fontFamilies, requiredFaces } from '../../scripts/stream-s0/t4-fonts'

const literBytes = (): Uint8Array => new Uint8Array(readFileSync(new URL('../fixtures/fonts/Liter-Regular.ttf', import.meta.url)))

/** A minimal docx with an embedded (ODTTF-obfuscated) Liter face. */
async function fixtureDocxWithEmbeddedFont(): Promise<Uint8Array> {
  const guid = '{12345678-ABCD-EF01-2345-6789ABCDEF01}'
  const obfuscated = deobfuscateOdttf(literBytes(), guid)
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Probe</w:t></w:r></w:p></w:body></w:document>`)
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>`)
  zip.file('word/fontTable.xml', `<?xml version="1.0"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="Liter"><w:embedRegular r:id="rId1" w:fontKey="${guid}"/></w:font></w:fonts>`)
  zip.file('word/_rels/fontTable.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/liter.odttf"/></Relationships>`)
  zip.file('word/fonts/liter.odttf', obfuscated)
  return (await zip.generateAsync({ type: 'uint8array' })) as Uint8Array
}

it('hash96 is 24 hex chars and stable', () => {
  const bytes = literBytes()
  const h1 = hash96(bytes)
  const h2 = hash96(new Uint8Array(bytes))
  expect(h1).toMatch(/^[0-9a-f]{24}$/)
  expect(h1).toBe(h2)
  const flipped = new Uint8Array(bytes)
  flipped[100] ^= 0xff
  expect(hash96(flipped)).not.toBe(h1)
})

it('system face is unverifiable (no bytes ⇒ §7.4 rule 3 raster)', async () => {
  // A census font string with no bytes classifies system; requiredFaces
  // synthesises it from the census font list — here directly via a fake file
  // whose census sees an unresolvable family.
  const dir = mkdtempSync(path.join(tmpdir(), 'stream-s0-t4-'))
  // A docx with NO embedded fonts: every census family is system/unverifiable.
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="NoSuchFace"/><w:rFonts w:hAnsi="NoSuchFace"/></w:rPr><w:t>Probe</w:t></w:r></w:p></w:body></w:document>`)
  const at = path.join(dir, 'system-probe.docx')
  writeFileSync(at, await zip.generateAsync({ type: 'uint8array' }))
  const result = await requiredFaces({ name: 'system-probe.docx', path: at, kind: 'docx' })
  const systemFaces = result.faces.filter((f) => f.source === 'system')
  expect(systemFaces.length).toBeGreaterThan(0)
  for (const face of systemFaces) {
    expect(face.bytes).toBeUndefined()
    expect(face.hash96).toBeUndefined()
    expect(face.source === 'system' && face.bytes === undefined).toBe(true)
  }
  // Round-9 §7.4 rule 3: any unit that uses a byte-less face would raster.
  const unitsUsingSystem = result.units.filter((u) => u.families.some((family) => !result.faces.some((f) => f.family.toLowerCase() === family && f.bytes)))
  expect(unitsUsingSystem.length).toBe(result.unitsThatWouldRaster)
})

it('embedded font in a fixture docx is verifiable', async () => {
  const bytes = await fixtureDocxWithEmbeddedFont()
  const dir = mkdtempSync(path.join(tmpdir(), 'stream-s0-t4-'))
  const at = path.join(dir, 'embedded-probe.docx')
  writeFileSync(at, bytes)
  const result = await requiredFaces({ name: 'embedded-probe.docx', path: at, kind: 'docx' })
  const liter = result.faces.find((f) => f.family === 'Liter' && f.source === 'embedded')
  expect(liter).toBeTruthy()
  expect(liter!.bytes?.length).toBeGreaterThan(0)
  expect(liter!.hash96).toMatch(/^[0-9a-f]{24}$/)
  expect(result.verifiable).toBeGreaterThanOrEqual(1)
  // The registered face must be verifiable: the unit(s) that used it do NOT
  // raster because of it.
  expect(result.faces.find((f) => f.family === 'Liter')?.source).toBe('embedded')
})

it('fontFamilies extracts quoted and unquoted families', () => {
  expect(fontFamilies('44pt "Aptos Display"')).toEqual(['Aptos Display'])
  expect(fontFamilies('bold 16px "Calibri"')).toEqual(['Calibri'])
  expect(fontFamilies('italic 12pt Liter')).toEqual(['Liter'])
  expect(fontFamilies('12px/1.4 system-ui, sans-serif')).toEqual([])
})
