// @vitest-environment node
import { readFileSync } from 'node:fs'
import { describe, expect, test, vi } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { decodeEmbeddedFont, FONT_LIMITS } from '../src/core/fonts/decode'
import { parsePptx } from '../src/pptx/parse'
const liter = new Uint8Array(readFileSync(new URL('./fixtures/fonts/Liter-Regular.ttf', import.meta.url)))
export function eot(data: Uint8Array, permissions = 8, root = '', flags = 0): Uint8Array {
  const header = new Uint8Array(100 + root.length * 2)
  const view = new DataView(header.buffer)
  view.setUint32(0, header.length + data.length, true)
  view.setUint32(4, data.length, true)
  view.setUint32(8, 0x20001, true)
  view.setUint32(12, flags, true)
  view.setUint16(32, permissions, true)
  view.setUint16(34, 0x504c, true)
  view.setUint16(98, root.length * 2, true)
  for (let i = 0; i < root.length; i++) view.setUint16(100 + i * 2, root.charCodeAt(i), true)
  const out = new Uint8Array(header.length + data.length)
  out.set(header); out.set(data, header.length)
  return out
}
describe('bounded font decode', () => {
  test('decodes real open licensed SFNT and EOT with display permissions', () => {
    expect(decodeEmbeddedFont(liter)).toEqual(liter)
    for (const permissions of [0, 4, 8]) expect(decodeEmbeddedFont(eot(liter, permissions))).toEqual(liter)
  })
  test('rejects restricted permissions, RootString restrictions and truncation', () => {
    expect(() => decodeEmbeddedFont(eot(liter, 2))).toThrow(/restricted/i)
    expect(() => decodeEmbeddedFont(eot(liter, 0, 'https://example.com'))).toThrow(/root/i)
    for (const n of [0, 4, 11, 70, 101]) expect(() => decodeEmbeddedFont(eot(liter).slice(0, n))).toThrow()
    const hostile = eot(liter); new DataView(hostile.buffer).setUint32(4, 0xffffffff, true)
    expect(() => decodeEmbeddedFont(hostile)).toThrow()
    expect(() => decodeEmbeddedFont(new Uint8Array(FONT_LIMITS.inputBytes + 1))).toThrow(/limit|budget/i)
  })
})
async function fontPackage(faces: string, rels: string, parts: Record<string, Uint8Array>) {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="p" xmlns:r="r"><p:embeddedFontLst>${faces}</p:embeddedFontLst></p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships>${rels}</Relationships>`)
  for (const [path, data] of Object.entries(parts)) zip.file(path, data)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}
describe('font parts', () => {
  test('resolves internal part relative variants and preserves declaration identity', async () => {
    const variants = ['regular', 'bold', 'italic', 'boldItalic']
    const pkg = await fontPackage(`<p:embeddedFont><p:font typeface="Liter"/>${variants.map(v => `<p:${v} r:id="${v}"/>`).join('')}</p:embeddedFont>`, variants.map(v => `<Relationship Id="${v}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="../fonts/${v}.fntdata"/>`).join(''), Object.fromEntries(variants.map(v => [`fonts/${v}.fntdata`, eot(liter)])))
    const doc = await parsePptx(pkg)
    expect(doc.embeddedFonts?.map(f => [f.family, f.variant, f.partPath, f.bytes.length])).toEqual(variants.map(v => ['Liter', v, `fonts/${v}.fntdata`, liter.length]))
    expect(doc.fontDiagnostics).toEqual([])
  })
  test('each external, missing, malformed and restricted face degrades independently', async () => {
    const pkg = await fontPackage('<p:embeddedFont><p:font typeface="Liter"/><p:regular r:id="ok"/><p:bold r:id="external"/><p:italic r:id="missing"/><p:boldItalic r:id="bad"/></p:embeddedFont><p:embeddedFont><p:font typeface="Restricted"/><p:regular r:id="restricted"/></p:embeddedFont>', '<Relationship Id="ok" Target="fonts/ok"/><Relationship Id="external" Target="https://example.com/font" TargetMode="External"/><Relationship Id="missing" Target="fonts/missing"/><Relationship Id="bad" Target="fonts/bad"/><Relationship Id="restricted" Target="fonts/restricted"/>', { 'ppt/fonts/ok': liter, 'ppt/fonts/bad': new Uint8Array(2), 'ppt/fonts/restricted': eot(liter, 2) })
    const doc = await parsePptx(pkg)
    expect(doc.embeddedFonts).toHaveLength(1)
    expect(doc.fontDiagnostics?.map(d => d.kind)).toEqual(['external-font', 'missing-font', 'malformed-font', 'restricted-font'])
    expect(doc.fontDiagnostics?.every(d => d.family && d.variant && d.relationshipId)).toBe(true)
  })
  test('caps face enumeration at 64', async () => {
    const pkg = await fontPackage(Array.from({ length: 65 }, (_, i) => `<p:embeddedFont><p:font typeface="f${i}"/><p:regular r:id="f"/></p:embeddedFont>`).join(''), '<Relationship Id="f" Target="fonts/f"/>', { 'ppt/fonts/f': liter })
    const doc = await parsePptx(pkg)
    expect(doc.embeddedFonts).toHaveLength(64)
    expect(doc.fontDiagnostics?.at(-1)?.kind).toBe('font-limit')
  })
})

test('ZIP rejects actual streamed inflated bytes with misleading declared size before cache', async () => {
  const zip = new JSZip(); zip.file('font', new Uint8Array(1024 * 1024))
  const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  // central directory uncompressed length lies; local header too
  const dv = new DataView(data.buffer)
  for (let i = 0; i < data.length - 30; i++) {
    const sig = dv.getUint32(i, true)
    if (sig === 0x02014b50) dv.setUint32(i + 24, 1, true)
    if (sig === 0x04034b50) dv.setUint32(i + 22, 1, true)
  }
  const pkg = await OfficePackage.load(data)
  await expect(pkg.bytes('font', 32 * 1024)).rejects.toThrow(/limit|budget/)
  await expect(pkg.bytes('font', 16 * 1024)).rejects.toThrow(/limit|budget/)
  const normal = new JSZip(); normal.file('a', new Uint8Array(20))
  const ordinary = await OfficePackage.load(await normal.generateAsync({ type: 'uint8array' }))
  expect(await ordinary.bytes('a')).toHaveLength(20)
  await expect(ordinary.bytes('a', 10)).rejects.toThrow(/limit|budget/)
})

test('decodes an authored real MTX and MTX EOT; hostile table and truncations terminate', () => {
  const mtx = new Uint8Array(readFileSync(new URL('./fixtures/fonts/authored-empty.mtx', import.meta.url)))
  const bytes = decodeEmbeddedFont(mtx)
  expect(new DataView(bytes.buffer).getUint32(0)).toBe(0x10000)
  expect(bytes.length).toBeGreaterThan(150)
  expect(decodeEmbeddedFont(eot(mtx, 8, '', 4))).toEqual(bytes)
  const hostile = new Uint8Array(readFileSync(new URL('./fixtures/fonts/authored-hostile-table.mtx', import.meta.url)))
  expect(() => decodeEmbeddedFont(hostile)).toThrow(/budget/)
  for (let n = 0; n < mtx.length; n++) expect(() => decodeEmbeddedFont(mtx.slice(0, n))).toThrow()
})

test('reads direct latin/ea/cs family declarations without replacing requested names', async () => {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="s"/></p:sldIdLst></p:presentation>')
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="s" Target="slides/a.xml"/></Relationships>')
  zip.file('ppt/slides/a.xml', '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:rPr><a:latin typeface="Liter"/><a:ea typeface="MiSans"/><a:cs typeface="微软雅黑"/></a:rPr><a:t>abc</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>')
  const doc = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  expect(doc.slides[0].shapes[0].textBody?.paragraphs[0].runs[0]).toMatchObject({ fontFamily: 'Liter', fontFamilyEastAsia: 'MiSans', fontFamilyComplexScript: '微软雅黑', text: 'abc' })
})

test('conversion budget is enforced before SFNT copy and MTX stream/table allocation', () => {
  expect(() => decodeEmbeddedFont(liter, 1024)).toThrow(/budget|limit/)
  const mtx = new Uint8Array(readFileSync(new URL('./fixtures/fonts/authored-empty.mtx', import.meta.url)))
  expect(() => decodeEmbeddedFont(mtx, 128)).toThrow(/budget|limit/)
})

test('hostile glyph point counts fail before arrays are allocated', () => {
  const hostile = new Uint8Array(readFileSync(new URL('./fixtures/fonts/authored-hostile-points.mtx', import.meta.url)))
  expect(() => decodeEmbeddedFont(hostile)).toThrow(/glyph points budget/)
})
test('oversized compressed font part is rejected before font decode, while a neighboring face survives', async () => {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:embeddedFontLst><p:embeddedFont><p:font typeface="Liter"/><p:regular r:id="ok"/><p:bold r:id="large"/></p:embeddedFont></p:embeddedFontLst></p:presentation>')
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="ok" Target="fonts/ok"/><Relationship Id="large" Target="fonts/large"/></Relationships>')
  zip.file('ppt/fonts/ok', liter); zip.file('ppt/fonts/large', new Uint8Array(FONT_LIMITS.inputBytes + 1))
  const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  const dv = new DataView(data.buffer)
  // Alter only declared size of the large member, leaving its compressed stream intact.
  for (let i = 0; i < data.length - 30; i++) {
    const signature = dv.getUint32(i, true)
    if (signature === 0x02014b50 && dv.getUint32(i + 24, true) > FONT_LIMITS.inputBytes) dv.setUint32(i + 24, 1, true)
    if (signature === 0x04034b50 && dv.getUint32(i + 22, true) > FONT_LIMITS.inputBytes) dv.setUint32(i + 22, 1, true)
  }
  const doc = await parsePptx(await OfficePackage.load(data))
  expect(doc.embeddedFonts).toHaveLength(1)
  expect(doc.fontDiagnostics).toEqual([expect.objectContaining({ kind: 'font-limit', variant: 'bold' })])
})
test('total converted document bytes stop at 64 MiB before the next SFNT conversion', async () => {
  const padded = new Uint8Array(2 * 1024 * 1024); padded.set(liter)
  const faces = Array.from({ length: 33 }, (_, i) => `<p:embeddedFont><p:font typeface="f${i}"/><p:regular r:id="f"/></p:embeddedFont>`).join('')
  const doc = await parsePptx(await fontPackage(faces, '<Relationship Id="f" Target="fonts/f"/>', { 'ppt/fonts/f': padded }))
  expect(doc.embeddedFonts?.reduce((n, f) => n + f.bytes.length, 0)).toBe(FONT_LIMITS.documentBytes)
  expect(doc.embeddedFonts).toHaveLength(32)
  expect(doc.fontDiagnostics).toEqual([expect.objectContaining({ kind: 'font-limit', family: 'f32' })])
})

test('RLE expansion is bounded before growing output and unfinished RLE terminates', () => {
  const hostile = new Uint8Array(readFileSync(new URL('./fixtures/fonts/authored-hostile-rle.mtx', import.meta.url)))
  expect(() => decodeEmbeddedFont(hostile, 16384)).toThrow(/allocation budget/)
  const truncated = new Uint8Array(readFileSync(new URL('./fixtures/fonts/authored-truncated-rle.mtx', import.meta.url)))
  expect(() => decodeEmbeddedFont(truncated)).toThrow(/Truncated LZ run length/)
})


test('uncompressed EOT checks the remaining converted byte budget before copying font data', () => {
  const copy = Uint8Array.prototype.slice
  const copied: number[] = []
  const spy = vi.spyOn(Uint8Array.prototype, 'slice').mockImplementation(function(this: Uint8Array, ...args) { copied.push(this.length); return copy.apply(this, args) })
  try {
    expect(() => decodeEmbeddedFont(eot(liter), 1024)).toThrow(/budget|limit/)
    expect(copied).not.toContain(liter.length)
  } finally { spy.mockRestore() }
})

test('CVT aggregate allowance is checked before expanded buffers are copied', () => {
  const hostile = new Uint8Array(readFileSync(new URL('./fixtures/fonts/authored-hostile-cvt-aggregate.mtx', import.meta.url)))
  const copy = Uint8Array.prototype.slice
  let convertedBytes = 0, expandedReservations = 0
  const ByteArray = Uint8Array
  vi.stubGlobal('Uint8Array', new Proxy(ByteArray, { construct(target, args, newTarget) {
    if (args[0] === 65535 * 2) expandedReservations++
    return Reflect.construct(target, args, newTarget)
  } }))
  const spy = vi.spyOn(Uint8Array.prototype, 'slice').mockImplementation(function(this: Uint8Array, ...args) {
    const result = copy.apply(this, args)
    if (result.length === 65535 * 2) convertedBytes += result.length
    return result
  })
  try {
    expect(() => decodeEmbeddedFont(hostile)).toThrow(/allocation budget/)
    expect(convertedBytes).toBeLessThanOrEqual(FONT_LIMITS.decodedBytes)
    expect(expandedReservations).toBe(255) // The 256th reserve never allocates.
  } finally { spy.mockRestore(); vi.unstubAllGlobals() }
})

test('a unique expanding CVT honors its siblings before allocation and decodes when allowed', async () => {
  // Expose only the existing Stream class to construct parser inputs. This is
  // unchanged vendor code, not a production test-only API or mocked parser.
  const source = readFileSync(new URL('../src/core/fonts/vendor/mtx.mjs', import.meta.url), 'utf8')
  const vendor = await import(/* @vite-ignore */ `data:text/javascript;base64,${Buffer.from(source + '\nexport { Stream };').toString('base64')}`)
  const fixture = new Uint8Array(readFileSync(new URL('./fixtures/fonts/authored-cvt-sibling.mtx', import.meta.url)))
  const streams = vendor.unpackMtx(fixture, fixture.length).streams as Uint8Array[]
  const makeStreams = () => streams.map(bytes => new vendor.Stream(bytes, bytes.length))
  const copy = Uint8Array.prototype.slice
  const copied: number[] = []
  let expandedReservations = 0
  const ByteArray = Uint8Array
  vi.stubGlobal('Uint8Array', new Proxy(ByteArray, { construct(target, args, newTarget) {
    if (args[0] === 512) expandedReservations++
    return Reflect.construct(target, args, newTarget)
  } }))
  const spy = vi.spyOn(Uint8Array.prototype, 'slice').mockImplementation(function(this: Uint8Array, ...args) { const result = copy.apply(this, args); copied.push(result.length); return result })
  try {
    expect(() => vendor.withDecoderByteLimit(4096, () => vendor.parseCTF(makeStreams()))).toThrow(/allocation budget/)
    expect(copied).not.toContain(512)
    expect(expandedReservations).toBe(0)
    const font = vendor.withDecoderByteLimit(8192, () => vendor.parseCTF(makeStreams()))
    expect(font.tables.find((table: { tag: string }) => table.tag === 'cvt ').bufSize).toBe(512)
    expect(expandedReservations).toBe(1)
  } finally { spy.mockRestore(); vi.unstubAllGlobals() }
})
