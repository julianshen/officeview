/**
 * Shared ODF core tests: length units, container sniffing, style plumbing.
 */
import { describe, test, expect } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { lengthTwips, lengthEmu, fontSizePt, hexColor } from '../src/odf/units'
import { odfKind } from '../src/odf/container'
import { buildOdt } from '../src/testdata/odf-builders'
import { buildDocx } from '../src/testdata/ooxml-builders'

describe('odf length units', () => {
  test('absolute lengths convert to twips', () => {
    expect(lengthTwips('2.54cm')).toBeCloseTo(1440, 3)
    expect(lengthTwips('25.4mm')).toBeCloseTo(1440, 3)
    expect(lengthTwips('1in')).toBe(1440)
    expect(lengthTwips('12pt')).toBe(240)
    expect(lengthTwips('1pc')).toBe(240)
    expect(lengthTwips('96px')).toBeCloseTo(1440, 3)
    expect(lengthTwips('bogus')).toBeUndefined()
    expect(lengthTwips(undefined)).toBeUndefined()
  })

  test('lengths convert to EMU for image extents', () => {
    expect(lengthEmu('2.54cm')).toBeCloseTo(914400, 0)
    expect(lengthEmu('1in')).toBe(914400)
  })

  test('font sizes handle points and percent-of-base', () => {
    expect(fontSizePt('18pt')).toBe(18)
    expect(fontSizePt('150%', 12)).toBe(18)
    expect(fontSizePt('150%')).toBeUndefined()
    expect(fontSizePt('huge')).toBeUndefined()
  })

  test('hex colors normalize', () => {
    expect(hexColor('#ffcc00')).toBe('FFCC00')
    expect(hexColor('red')).toBeUndefined()
  })
})

describe('odf container sniffing', () => {
  test('a builder odt sniffs as text', async () => {
    const pkg = await OfficePackage.load(await buildOdt({ paras: [{ text: 'x' }] }))
    expect(await odfKind(pkg)).toBe('text')
  })

  test('manifest media-type backs up a missing mimetype entry', async () => {
    const buf = await buildOdt({ paras: [{ text: 'x' }] })
    const zip = await JSZip.loadAsync(buf)
    zip.remove('mimetype')
    const forged = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    expect(await odfKind(forged)).toBe('text')
  })

  test('a docx is not ODF', async () => {
    const pkg = await OfficePackage.load(await buildDocx([{ runs: [{ text: 'x' }] }]))
    expect(await odfKind(pkg)).toBeUndefined()
  })
})
