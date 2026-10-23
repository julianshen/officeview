// Office-shaped synthetic fixtures; not Office evidence.
import { it, expect } from 'vitest'
import JSZip from 'jszip'
import { spawnSync } from 'node:child_process'
import { posix } from 'node:path'
import { buildPptxDeck, timingXmlOf, mainSeqXml } from '../src/testdata/ooxml-builders'
const p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
async function inspect(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes)
  const docs: Record<string, Document> = {}
  const lint = spawnSync('xmllint', ['--version']).status === 0
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir || !/\.(xml|rels)$/.test(path)) continue
    const xml = await file.async('text')
    const doc = new DOMParser().parseFromString(xml, 'application/xml')
    expect(doc.querySelector('parsererror'), path).toBeNull()
    for (const el of Array.from(doc.getElementsByTagName('*'))) {
      if (el.prefix) expect(el.lookupNamespaceURI(el.prefix), path).toBe(el.namespaceURI)
      for (const attr of Array.from(el.attributes)) if (attr.prefix && attr.prefix !== 'xmlns') expect(el.lookupNamespaceURI(attr.prefix), path).toBe(attr.namespaceURI)
    }
    if (lint) expect(spawnSync('xmllint', ['--noout', '-'], { input: xml }).status, path).toBe(0)
    docs[path] = doc
    if (path.endsWith('.rels')) {
      const ids = Array.from(doc.getElementsByTagName('Relationship')).map(el => el.getAttribute('Id'))
      expect(new Set(ids).size).toBe(ids.length)
      for (const el of Array.from(doc.getElementsByTagName('Relationship'))) {
        if (el.getAttribute('TargetMode') === 'External') continue
        const base = path === '_rels/.rels' ? '' : posix.dirname(posix.dirname(path))
        expect(zip.file(posix.normalize(posix.join(base, el.getAttribute('Target')!))), path).not.toBeNull()
      }
    }
    const ids = Array.from(doc.getElementsByTagNameNS(p, 'cNvPr')).map(el => el.getAttribute('id'))
    expect(new Set(ids).size).toBe(ids.length)
  }
  for (const [path, doc] of Object.entries(docs)) {
    if (path.endsWith('.rels')) continue
    const relPath = posix.join(posix.dirname(path), '_rels', posix.basename(path) + '.rels')
    const relIds = new Set(Array.from(docs[relPath]?.getElementsByTagName('Relationship') ?? []).map(el => el.getAttribute('Id')))
    for (const el of Array.from(doc.getElementsByTagName('*'))) for (const attr of Array.from(el.attributes)) {
      if (attr.namespaceURI === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') expect(relIds.has(attr.value), `${path}: ${attr.name}=${attr.value}`).toBe(true)
    }
  }
  return { zip, docs }
}
it('an AlternateContent transition keeps mc:Choice Requires and the fallback', async () => {
  const timing = timingXmlOf(mainSeqXml(['<p:par><p:cTn id="7" dur="100"/></p:par>']))
  const { docs } = await inspect(await buildPptxDeck({ slides: [{ shapes: [], timingXml: timing, transition: { xml: '<p:transition p14:dur="300"><p14:flip/></p:transition>', wrapInAlternateContent: { requires: 'p14 p15', fallbackXml: '<p:transition><p:fade/></p:transition>' } } }] }))
  const doc = docs['ppt/slides/slide1.xml']
  expect(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/markup-compatibility/2006', 'Choice')[0].getAttribute('Requires')).toBe('p14 p15')
  expect(doc.getElementsByTagNameNS(p, 'fade')).toHaveLength(1)
  const nodes = Array.from(doc.getElementsByTagNameNS(p, 'cTn'))
  expect(nodes.map(n => n.getAttribute('nodeType'))).toEqual(expect.arrayContaining(['tmRoot', 'mainSeq']))
  expect(new Set(nodes.map(n => n.getAttribute('id'))).size).toBe(nodes.length)
})
it('notes slides, hidden slides, hyperlink rels and media rels resolve', async () => {
  const { docs, zip } = await inspect(await buildPptxDeck({ slides: [0, 1].map(() => ({ hidden: true, notes: [['one & <two>', 'three'], []], shapes: [{ id: 2, paragraphs: [{ runs: [{ text: 'go', color: 'FF0000' }] }] }, { id: 3 }], links: [{ shapeId: 2, target: 'https://example.com/?a=1&b=2', external: true, onRun: true }, { shapeId: 3, action: 'ppaction://hlinkshowjump?jump=nextslide' }], media: [{ kind: 'audio', shapeId: 2, data: new Uint8Array([1]), ext: 'mp3', p14: true }, { kind: 'video', shapeId: 3, data: new Uint8Array([2]), ext: 'mp4' }] })) }))
  expect(docs['ppt/slides/slide1.xml'].documentElement.getAttribute('show')).toBe('0')
  expect(docs['ppt/notesSlides/notesSlide1.xml'].getElementsByTagNameNS(a, 't')[0].textContent).toBe('one & <two>')
  expect(docs['ppt/slides/slide1.xml'].getElementsByTagNameNS(a, 'hlinkClick')).toHaveLength(2)
  expect(Array.from(docs['ppt/slides/slide1.xml'].getElementsByTagNameNS(a, 'rPr')[0].children).map(el => el.localName)).toEqual(['solidFill', 'hlinkClick'])
  expect(docs['ppt/slides/slide1.xml'].getElementsByTagNameNS(a, 'audioFile')).toHaveLength(1)
  expect(docs['ppt/slides/slide1.xml'].getElementsByTagNameNS(a, 'videoFile')).toHaveLength(1)
  expect(await zip.file('[Content_Types].xml')!.async('text')).toContain('audio/mpeg')
  expect(await zip.file('[Content_Types].xml')!.async('text')).toContain('video/mp4')
})
it('presProps is written to ppt/presProps.xml with its relationship', async () => {
  const props = `<p:presentationPr xmlns:p="${p}"><p:showPr loop="1"><p:kiosk restart="300000"/></p:showPr></p:presentationPr>`
  const { zip } = await inspect(await buildPptxDeck({ slides: [{ shapes: [] }], presProps: props }))
  expect(await zip.file('ppt/presProps.xml')!.async('text')).toBe(props)
  expect(await zip.file('ppt/_rels/presentation.xml.rels')!.async('text')).toContain('Target="presProps.xml"')
})
it('rejects unresolved shapes and run links without a run', async () => {
  await expect(buildPptxDeck({ slides: [{ shapes: [], links: [{ shapeId: 99 }] }] })).rejects.toThrow('unknown show shape')
  await expect(buildPptxDeck({ slides: [{ shapes: [{ id: 2 }], links: [{ shapeId: 2, onRun: true }] }] })).rejects.toThrow('text run')
})
it('rejects duplicate attachments instead of dropping one or producing invalid children', async () => {
  await expect(buildPptxDeck({ slides: [{ shapes: [{ id: 2 }], links: [{ shapeId: 2 }, { shapeId: 2 }] }] })).rejects.toThrow('duplicate')
  const media = { kind: 'audio' as const, shapeId: 2, data: new Uint8Array([1]), ext: 'mp3' }
  await expect(buildPptxDeck({ slides: [{ shapes: [{ id: 2 }], media: [media, media] }] })).rejects.toThrow('duplicate')
})
