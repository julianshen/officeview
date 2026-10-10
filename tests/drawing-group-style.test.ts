import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { paintScene } from '../src/drawing/scene-paint'
import type { SceneNode } from '../src/drawing/scene'
import type { GeometryCommand } from '../src/drawing/geometry'
import {
  parseGroupShapeProperties,
  resolveDrawingStyle,
  type DrawingStyle,
  type ThemeContext,
} from '../src/drawing/style'

const dummyTheme: ThemeContext = {
  colors: {
    accent1: { kind: 'srgb', value: '4472C4', transforms: [] },
    accent2: { kind: 'srgb', value: 'ED7D31', transforms: [] },
  },
  palette: {},
  colorMap: {},
  fonts: { major: { supplemental: {} }, minor: { supplemental: {} } },
  fillStyles: [],
  bgFillStyles: [],
  lineStyles: [],
  effectStyles: [],
  issues: [],
}

describe('Group shape style cascading (<a:grpSpPr>)', () => {
  test('parses group shape properties (grpSpPr) including fills and leniently accepts nonconformant outline defaults (M2)', () => {
    const xml = parseXmlOrdered(
      `<p:grpSpPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:xfrm>` +
      `    <a:off x="100" y="200"/>` +
      `    <a:ext cx="300" cy="400"/>` +
      `    <a:chOff x="0" y="0"/>` +
      `    <a:chExt cx="300" cy="400"/>` +
      `  </a:xfrm>` +
      `  <a:solidFill>` +
      `    <a:srgbClr val="FF5500"/>` +
      `  </a:solidFill>` +
      `  <a:ln w="25400" cap="rnd">` +
      `    <a:solidFill><a:srgbClr val="00AA22"/></a:solidFill>` +
      `    <a:prstDash val="dash"/>` +
      `  </a:ln>` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="38100" dir="5400000" blurRad="19050">` +
      `      <a:srgbClr val="000000"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</p:grpSpPr>`
    )

    const style = parseGroupShapeProperties(xml, dummyTheme)
    expect(style).toBeDefined()
    // Fill
    expect(style.fill).toEqual({
      kind: 'solid',
      color: { r: 255, g: 85, b: 0, a: 1 },
    })
    // Line outline defaults
    expect(style.line).toBeDefined()
    expect(style.line?.width).toBeCloseTo(25400 / 9525, 2)
    expect(style.line?.cap).toBe('round')
    expect(style.line?.dash).toBe('dash')
    expect(style.line?.fill).toEqual({
      kind: 'solid',
      color: { r: 0, g: 170, b: 34, a: 1 },
    })
    // Shadow
    expect(style.shadow).toBeDefined()
    expect(style.shadow?.color).toEqual({ r: 0, g: 0, b: 0, a: 1 })
    expect(style.shadow?.offsetY).toBeCloseTo(4, 1)

    // Also verify pattern fill on grpSpPr
    const xmlPattern = parseXmlOrdered(
      `<a:grpSpPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:pattFill prst="pct50">` +
      `    <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>` +
      `    <a:bgClr><a:srgbClr val="0000FF"/></a:bgClr>` +
      `  </a:pattFill>` +
      `</a:grpSpPr>`
    )
    const patternStyle = parseGroupShapeProperties(xmlPattern, dummyTheme)
    expect(patternStyle.fill?.kind).toBe('pattern')
    if (patternStyle.fill?.kind === 'pattern') {
      expect(patternStyle.fill.preset).toBe('pct50')
      expect(patternStyle.fill.fgColor).toEqual({ r: 255, g: 0, b: 0, a: 1 })
      expect(patternStyle.fill.bgColor).toEqual({ r: 0, g: 0, b: 255, a: 1 })
    }
  })

  test('child shapes with <a:grpFill/> or omitted fill inherit parent group fill style', () => {
    const parentGroupStyle: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 128, b: 0, a: 1 } },
      line: { width: 3, cap: 'round', fill: { kind: 'solid', color: { r: 0, g: 128, b: 255, a: 1 } } },
      issues: [],
    }

    // 1. Child with omitted fill in spPr inherits group fill and line
    const spPrOmitted = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:prstGeom prst="rect"/>` +
      `</a:spPr>`
    )
    const styleOmitted = resolveDrawingStyle(spPrOmitted, undefined, dummyTheme, parentGroupStyle)
    expect(styleOmitted.fill).toEqual(parentGroupStyle.fill)
    expect(styleOmitted.line).toEqual(parentGroupStyle.line)

    // 2. Child with <a:grpFill/> inherits group fill even if a theme style fillRef is present
    const spPrUseBg = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:grpFill/>` +
      `</a:spPr>`
    )
    const styleNode = parseXmlOrdered(
      `<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
      `</p:style>`
    )
    const styleUseBg = resolveDrawingStyle(spPrUseBg, styleNode, dummyTheme, parentGroupStyle)
    expect(styleUseBg.fill).toEqual(parentGroupStyle.fill)

    // 3. Child with <a:grpFill/> inherits group fill
    const spPrGrpFill = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:grpFill/>` +
      `</a:spPr>`
    )
    const styleGrpFill = resolveDrawingStyle(spPrGrpFill, styleNode, dummyTheme, parentGroupStyle)
    expect(styleGrpFill.fill).toEqual(parentGroupStyle.fill)
    expect(styleGrpFill.issues.some(i => i.kind === 'unsupported-fill' && i.feature === 'grpFill')).toBe(false)

    // 4. Child with explicit <a:noFill/> does NOT inherit group fill
    const spPrNoFill = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:noFill/>` +
      `</a:spPr>`
    )
    const styleNoFill = resolveDrawingStyle(spPrNoFill, undefined, dummyTheme, parentGroupStyle)
    expect(styleNoFill.fill).toEqual({ kind: 'none' })

    // 5. Child with explicit solid fill overrides group fill, but still inherits line if line omitted
    const spPrExplicit = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>` +
      `</a:spPr>`
    )
    const styleExplicit = resolveDrawingStyle(spPrExplicit, undefined, dummyTheme, parentGroupStyle)
    expect(styleExplicit.fill).toEqual({ kind: 'solid', color: { r: 0, g: 255, b: 0, a: 1 } })
    expect(styleExplicit.line).toEqual(parentGroupStyle.line)
  })

  test('PPTX parse inherits parent group fill and outline style on child shapes with omitted fill or <a:grpFill/>', async () => {
    const zip = new JSZip()
    zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="2857500" cy="1905000"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>')
    zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>')
    zip.file('ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `  <p:cSld><p:spTree>` +
      `    <p:grpSp>` +
      `      <p:nvGrpSpPr><p:cNvPr id="2" name="Group 2"/><p:nvPr/></p:nvGrpSpPr>` +
      `      <p:grpSpPr>` +
      `        <a:xfrm><a:off x="100" y="200"/><a:ext cx="1000" cy="800"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="800"/></a:xfrm>` +
      `        <a:solidFill><a:srgbClr val="FF5500"/></a:solidFill>` +
      `        <a:ln w="25400"><a:solidFill><a:srgbClr val="00AA22"/></a:solidFill></a:ln>` +
      `      </p:grpSpPr>` +
      `      <p:sp>` +
      `        <p:nvSpPr><p:cNvPr id="3" name="Child Omitted"/><p:nvPr/></p:nvSpPr>` +
      `        <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="200" cy="200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
      `      </p:sp>` +
      `      <p:sp>` +
      `        <p:nvSpPr><p:cNvPr id="4" name="Child UseBg"/><p:nvPr/></p:nvSpPr>` +
      `        <p:spPr><a:xfrm><a:off x="300" y="200"/><a:ext cx="200" cy="200"/></a:xfrm><a:prstGeom prst="ellipse"/><a:grpFill/></p:spPr>` +
      `      </p:sp>` +
      `      <p:sp>` +
      `        <p:nvSpPr><p:cNvPr id="5" name="Child Explicit"/><p:nvPr/></p:nvSpPr>` +
      `        <p:spPr><a:xfrm><a:off x="500" y="200"/><a:ext cx="200" cy="200"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></p:spPr>` +
      `      </p:sp>` +
      `      <p:sp>` +
      `        <p:nvSpPr><p:cNvPr id="6" name="Child NoFill"/><p:nvPr/></p:nvSpPr>` +
      `        <p:spPr><a:xfrm><a:off x="700" y="200"/><a:ext cx="200" cy="200"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></p:spPr>` +
      `      </p:sp>` +
      `    </p:grpSp>` +
      `  </p:spTree></p:cSld>` +
      `</p:sld>`
    )
    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const doc = await parsePptx(pkg)
    const groupShape = doc.slides[0].shapes[0]
    expect(groupShape.children).toHaveLength(4)

    const [childOmitted, childUseBg, childExplicit, childNoFill] = groupShape.children!
    // Child 1: omitted fill inherits group solid fill and line
    expect(childOmitted.fill).toBe('#FF5500')
    expect(childOmitted.drawingStyle?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 85, b: 0, a: 1 } })
    expect(childOmitted.line?.color).toBe('#00AA22')

    // Child 2: <a:grpFill/> inherits group fill
    expect(childUseBg.fill).toBe('#FF5500')
    expect(childUseBg.drawingStyle?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 85, b: 0, a: 1 } })

    // Child 3: explicit fill overrides group fill
    expect(childExplicit.fill).toBe('#0000FF')
    expect(childExplicit.drawingStyle?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 0, b: 255, a: 1 } })
    expect(childExplicit.line?.color).toBe('#00AA22')

    // Child 4: explicit noFill does NOT inherit group fill
    expect(childNoFill.fill).toBeUndefined()
    expect(childNoFill.drawingStyle?.fill).toEqual({ kind: 'none' })
  })

  test('group style inheritance cascades cleanly across nested group hierarchies without mutating child geometry', async () => {
    const zip = new JSZip()
    zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="2857500" cy="1905000"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>')
    zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>')
    zip.file('ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `  <p:cSld><p:spTree>` +
      `    <p:grpSp>` +
      `      <p:nvGrpSpPr><p:cNvPr id="10" name="Root Group"/><p:nvPr/></p:nvGrpSpPr>` +
      `      <p:grpSpPr>` +
      `        <a:xfrm><a:off x="50" y="60"/><a:ext cx="2000" cy="1500"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1500"/></a:xfrm>` +
      `        <a:solidFill><a:srgbClr val="FF5500"/></a:solidFill>` +
      `        <a:effectLst>` +
      `          <a:outerShdw dist="38100" dir="5400000" blurRad="19050">` +
      `            <a:srgbClr val="000000"/>` +
      `          </a:outerShdw>` +
      `        </a:effectLst>` +
      `      </p:grpSpPr>` +
      `      <p:grpSp>` +
      `        <p:nvGrpSpPr><p:cNvPr id="20" name="Nested Child Group"/><p:nvPr/></p:nvGrpSpPr>` +
      `        <p:grpSpPr>` +
      `          <a:xfrm><a:off x="100" y="120"/><a:ext cx="1000" cy="800"/><a:chOff x="10" y="20"/><a:chExt cx="500" cy="400"/></a:xfrm>` +
      `          <a:ln w="25400" cap="rnd">` +
      `            <a:solidFill><a:srgbClr val="00AA22"/></a:solidFill>` +
      `            <a:prstDash val="dash"/>` +
      `          </a:ln>` +
      `        </p:grpSpPr>` +
      `        <p:sp>` +
      `          <p:nvSpPr><p:cNvPr id="30" name="Grandchild Shape"/><p:nvPr/></p:nvSpPr>` +
      `          <p:spPr>` +
      `            <a:xfrm rot="1800000" flipH="1">` +
      `              <a:off x="150" y="250"/><a:ext cx="350" cy="450"/>` +
      `            </a:xfrm>` +
      `            <a:prstGeom prst="roundRect">` +
      `              <a:avLst>` +
      `                <a:gd name="adj" fmla="val 30000"/>` +
      `              </a:avLst>` +
      `            </a:prstGeom>` +
      `          </p:spPr>` +
      `        </p:sp>` +
      `      </p:grpSp>` +
      `    </p:grpSp>` +
      `  </p:spTree></p:cSld>` +
      `</p:sld>`
    )
    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const doc = await parsePptx(pkg)

    const rootGroup = doc.slides[0].shapes[0]
    expect(rootGroup.group).toEqual({
      off: { x: 50, y: 60 },
      ext: { width: 2000, height: 1500 },
      chOff: { x: 0, y: 0 },
      chExt: { width: 2000, height: 1500 },
    })
    expect(rootGroup.drawingStyle?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 85, b: 0, a: 1 } })
    expect(rootGroup.drawingStyle?.shadow).toBeDefined()

    const childGroup = rootGroup.children![0]
    expect(childGroup.group).toEqual({
      off: { x: 100, y: 120 },
      ext: { width: 1000, height: 800 },
      chOff: { x: 10, y: 20 },
      chExt: { width: 500, height: 400 },
    })
    // Child group inherits fill and shadow from root group, and has its own line outline
    expect(childGroup.drawingStyle?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 85, b: 0, a: 1 } })
    expect(childGroup.drawingStyle?.line?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 170, b: 34, a: 1 } })
    expect(childGroup.drawingStyle?.shadow).toBeDefined()

    const grandchild = childGroup.children![0]
    // Grandchild inherits fill from root group, line from child group, and shadow from root group
    expect(grandchild.fill).toBe('#FF5500')
    expect(grandchild.drawingStyle?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 85, b: 0, a: 1 } })
    expect(grandchild.line?.color).toBe('#00AA22')
    expect(grandchild.drawingStyle?.line?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 170, b: 34, a: 1 } })
    expect(grandchild.drawingStyle?.shadow).toBeDefined()

    // Geometry, preset, and local coordinate transforms MUST remain unmutated
    expect(grandchild.geometry).toBe('roundRect')
    expect(grandchild.presetName).toBe('roundRect')
    expect(grandchild.drawingGeometry?.preset).toBe('roundRect')
    expect(grandchild.drawingGeometry?.adjustments).toEqual([['adj', 'val 30000']])
    expect(grandchild.xEmu).toBe(150)
    expect(grandchild.yEmu).toBe(250)
    expect(grandchild.widthEmu).toBe(350)
    expect(grandchild.heightEmu).toBe(450)
    expect(grandchild.rotationDeg).toBe(30)
    expect(grandchild.flipH).toBe(true)
    expect(grandchild.flipV).toBeFalsy()
  })

  test('diagnoses group-level gradient fill when inherited by child shape (M1)', () => {
    const groupStyle: DrawingStyle = {
      fill: {
        kind: 'gradient',
        gradient: 'linear',
        angle: 90,
        scaled: false,
        stops: [
          { position: 0, color: { r: 1, g: 0, b: 0, a: 1 } },
          { position: 1, color: { r: 0, g: 0, b: 1, a: 1 } },
        ],
      },
      issues: [],
    }
    const spPrOmitted = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:prstGeom prst="rect"/>` +
      `</a:spPr>`
    )
    const childStyle = resolveDrawingStyle(spPrOmitted, undefined, dummyTheme, groupStyle)
    expect(childStyle.fill?.kind).toBe('gradient')
    expect(childStyle.issues).toContainEqual({
      kind: 'unsupported-gradient',
      message: 'Group-level gradient fill across child shapes is evaluated in child local coordinates',
      feature: 'gradFill',
    })
  })

  test('group diagnostics are not duplicated onto child shapes (m3)', () => {
    const groupStyle: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 1, g: 0, b: 0, a: 1 } },
      issues: [{ kind: 'unsupported-effect', message: 'Deferred effect', feature: 'glow' }],
    }
    const spPrOmitted = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:prstGeom prst="rect"/>` +
      `</a:spPr>`
    )
    const childStyle = resolveDrawingStyle(spPrOmitted, undefined, dummyTheme, groupStyle)
    // Child shape should not duplicate the glow issue from parent group
    expect(childStyle.issues).toEqual([])
  })

  test('diagnoses non-auto bwMode on grpSpPr (s1)', () => {
    const xml = parseXmlOrdered(
      `<a:grpSpPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" bwMode="gray">` +
      `  <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm>` +
      `</a:grpSpPr>`
    )
    const style = parseGroupShapeProperties(xml, dummyTheme)
    expect(style.issues).toContainEqual({
      kind: 'unsupported-color-mode',
      message: 'Group black-and-white mode gray is deferred',
      feature: 'gray',
    })
  })

  test('parseGroupShapeProperties handles undefined grpSpPr cleanly', () => {
    const defaultStyle: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 1, g: 0, b: 0, a: 1 } },
      line: { width: 2 },
      shadow: { color: { r: 0, g: 0, b: 0, a: 1 }, blurPx: 5, offsetX: 2, offsetY: 2 },
      issues: [],
    }
    const result = parseGroupShapeProperties(undefined, dummyTheme, defaultStyle)
    expect(result.fill).toEqual(defaultStyle.fill)
    expect(result.line).toEqual(defaultStyle.line)
    expect(result.shadow).toEqual(defaultStyle.shadow)
    expect(result.shadow).not.toBe(defaultStyle.shadow) // Defensive copy
    expect(result.line).not.toBe(defaultStyle.line) // Defensive copy
  })

  test('child shape own shadow replaces inherited group shadow without stacking (m6)', () => {
    const parentStyle: DrawingStyle = {
      shadow: { color: { r: 1, g: 0, b: 0, a: 1 }, blurPx: 10, offsetX: 5, offsetY: 5 },
      issues: [],
    }
    const childSpPr = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="9525" dir="0" blurRad="0">` +
      `      <a:srgbClr val="0000FF"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const childStyle = resolveDrawingStyle(childSpPr, undefined, dummyTheme, parentStyle)
    expect(childStyle.shadow).toBeDefined()
    // Child's own shadow color (#0000FF = blue) replaces parent group's shadow (red)
    expect(childStyle.shadow?.color).toEqual({ r: 0, g: 0, b: 255, a: 1 })
    expect(childStyle.shadow?.offsetX).toBeCloseTo(1, 1)
  })

  test('child empty effectLst retains container group shadow under unit-effect model (m5)', () => {
    const parentStyle: DrawingStyle = {
      shadow: { color: { r: 1, g: 0, b: 0, a: 1 }, blurPx: 10, offsetX: 5, offsetY: 5 },
      issues: [],
    }
    const childSpPr = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:effectLst/>` +
      `</a:spPr>`
    )
    const childStyle = resolveDrawingStyle(childSpPr, undefined, dummyTheme, parentStyle)
    // Container group shadow applies
    expect(childStyle.shadow).toBeDefined()
    expect(childStyle.shadow?.color).toEqual({ r: 1, g: 0, b: 0, a: 1 })
  })

  test('paints child shape with inherited group fill onto Canvas 2D (m6)', () => {
    const canvas = createCanvas(200, 200)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D

    const groupNode: SceneNode = {
      xEmu: 10 * 9525,
      yEmu: 10 * 9525,
      widthEmu: 100 * 9525,
      heightEmu: 100 * 9525,
      group: {
        off: { x: 10 * 9525, y: 10 * 9525 },
        ext: { width: 100 * 9525, height: 100 * 9525 },
        chOff: { x: 0, y: 0 },
        chExt: { width: 100 * 9525, height: 100 * 9525 },
      },
      drawingStyle: {
        fill: { kind: 'solid' as const, color: { r: 255, g: 85, b: 0, a: 1 } },
        issues: [],
      },
      children: [
        {
          xEmu: 10 * 9525,
          yEmu: 10 * 9525,
          widthEmu: 80 * 9525,
          heightEmu: 80 * 9525,
          geometry: 'rect' as const,
          drawingGeometry: {
            preset: 'rect',
            adjustments: [],
            guides: [],
            paths: [{ commands: [['moveTo', '0', '0'], ['lnTo', 'w', '0'], ['lnTo', 'w', 'h'], ['lnTo', '0', 'h'], ['close']] as GeometryCommand[] }],
          },
          drawingStyle: {
            fill: { kind: 'solid' as const, color: { r: 255, g: 85, b: 0, a: 1 } },
            issues: [],
          },
        },
      ],
    }

    paintScene([groupNode], ctx)
    const rawData = (canvas.getContext('2d') as any).getImageData(50, 50, 1, 1).data
    // Pixel inside child rect must be painted with #FF5500 (r=255, g=85, b=0)
    expect(rawData[0]).toBe(255)
    expect(rawData[1]).toBe(85)
    expect(rawData[2]).toBe(0)
    expect(rawData[3]).toBe(255)
  })

  test('cascades style properties across 3-level deep group style hierarchy (s4)', () => {
    const rootGrpPr = parseXmlOrdered(
      `<a:grpSpPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
      `  <a:effectLst><a:outerShdw dist="9525" dir="0"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>` +
      `</a:grpSpPr>`
    )
    const rootStyle = parseGroupShapeProperties(rootGrpPr, dummyTheme)

    const midGrpPr = parseXmlOrdered(
      `<a:grpSpPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:ln w="19050"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>` +
      `</a:grpSpPr>`
    )
    const midStyle = parseGroupShapeProperties(midGrpPr, dummyTheme, rootStyle)
    expect(midStyle.fill).toEqual(rootStyle.fill)
    expect(midStyle.shadow).toEqual(rootStyle.shadow)
    expect(midStyle.line?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 255, b: 0, a: 1 } })

    const leafSpPr = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:prstGeom prst="rect"/>` +
      `</a:spPr>`
    )
    const leafStyle = resolveDrawingStyle(leafSpPr, undefined, dummyTheme, midStyle)
    expect(leafStyle.fill).toEqual({ kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } })
    expect(leafStyle.line?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 255, b: 0, a: 1 } })
    expect(leafStyle.shadow?.color).toEqual({ r: 0, g: 0, b: 0, a: 1 })
  })

  test('connector cxnSp inside group forces fill: none (m2)', () => {
    const parentGroupStyle: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 128, b: 0, a: 1 } },
      line: { width: 2, fill: { kind: 'solid', color: { r: 0, g: 0, b: 255, a: 1 } } },
      issues: [],
    }
    const cxnSpPr = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:prstGeom prst="line"/>` +
      `</a:spPr>`
    )
    // When element is connector, connector style override forces fill none
    const connectorDefaults = { ...parentGroupStyle, fill: { kind: 'none' as const } }
    const connectorStyle = resolveDrawingStyle(cxnSpPr, undefined, dummyTheme, connectorDefaults)
    expect(connectorStyle.fill).toEqual({ kind: 'none' })
    expect(connectorStyle.line).toEqual(parentGroupStyle.line)
  })

  test('blipFill shape preserves image fill without inheriting group fill (m6)', () => {
    const parentGroupStyle: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 128, b: 0, a: 1 } },
      issues: [],
    }
    const picSpPr = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:blipFill><a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></a:blipFill>` +
      `</a:spPr>`
    )
    const picStyle = resolveDrawingStyle(picSpPr, undefined, dummyTheme, parentGroupStyle)
    expect(picStyle.fill).toEqual({ kind: 'none' })
  })

  test('child shape with explicit theme fillRef uses theme fill over group fill (m6)', () => {
    const parentGroupStyle: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 128, b: 0, a: 1 } },
      issues: [],
    }
    const spPrOmitted = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:prstGeom prst="rect"/>` +
      `</a:spPr>`
    )
    const styleNode = parseXmlOrdered(
      `<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
      `</p:style>`
    )
    const themeWithFill: ThemeContext = {
      ...dummyTheme,
      fillStyles: [{ kind: 'solid', color: { kind: 'scheme', value: 'accent1', transforms: [] } }],
    }
    const resolved = resolveDrawingStyle(spPrOmitted, styleNode, themeWithFill, parentGroupStyle)
    // Theme fillRef takes precedence over parent group fill when child has no <a:grpFill/>
    expect(resolved.fill).toEqual({ kind: 'solid', color: { r: 68, g: 114, b: 196, a: 1 } })
  })
})
