/** Parse PPTX parts (presentation.xml, slides/slideN.xml) into PptxDocument. */
import type { OfficePackage } from '../core/zip'
import { attrs, elementChildren, getChildren, textOf, type XmlNode } from '../core/xml'
import { hexRgbToCss } from '../core/color'
import type { PptxDocument, PptxParagraph, PptxShape, PptxSlide, PptxTextBody, PptxTextRun } from './types'

const DEFAULT_INSET_LR = 91440
const DEFAULT_INSET_TB = 45720

function num(v: string | undefined, dflt = 0): number {
  const n = v === undefined ? NaN : parseFloat(v)
  return Number.isFinite(n) ? n : dflt
}

function colorOf(srgbClr: XmlNode | undefined): string | undefined {
  if (!srgbClr) return undefined
  return hexRgbToCss(attrs(srgbClr).val as string)
}

function parseAlign(v: string | undefined): PptxParagraph['align'] {
  switch (v) {
    case 'ctr': return 'center'
    case 'r': return 'right'
    case 'just': return 'justify'
    default: return 'left'
  }
}

function parseRun(r: XmlNode, inherited?: Partial<PptxTextRun>): PptxTextRun {
  const rPr = getChildren(r, 'rPr')[0]
  const run: PptxTextRun = { text: textOf(getChildren(r, 't')[0]), ...inherited }
  if (rPr) {
    const a = attrs(rPr)
    if (a.b !== undefined) run.bold = a.b === '1' || a.b === 'true'
    if (a.i !== undefined) run.italic = a.i === '1' || a.i === 'true'
    if (a.sz !== undefined) run.fontSizePt = num(a.sz as string) / 100
    if (a.typeface !== undefined) run.fontFamily = a.typeface as string
    run.color = colorOf(getChildren(getChildren(rPr, 'solidFill')[0], 'srgbClr')[0])
  }
  return run
}

function parseTextBody(txBody: XmlNode): PptxTextBody {
  const bodyPr = getChildren(txBody, 'bodyPr')[0]
  const ba = attrs(bodyPr)
  const body: PptxTextBody = {
    paragraphs: [],
    anchor: (ba.anchor as 't' | 'ctr' | 'b') ?? 't',
    insetLeftEmu: num(ba.lIns as string, DEFAULT_INSET_LR),
    insetRightEmu: num(ba.rIns as string, DEFAULT_INSET_LR),
    insetTopEmu: num(ba.tIns as string, DEFAULT_INSET_TB),
    insetBottomEmu: num(ba.bIns as string, DEFAULT_INSET_TB),
    wrap: ba.wrap !== 'none',
  }
  for (const p of getChildren(txBody, 'p')) {
    const pPr = getChildren(p, 'pPr')[0]
    const para: PptxParagraph = {
      runs: [],
      align: parseAlign(attrs(pPr).algn as string | undefined),
      bullet: pPr ? getChildren(pPr, 'buChar').length > 0 : false,
      level: num(attrs(pPr).lvl as string, 0),
    }
    let prevEnd: Partial<PptxTextRun> = {}
    for (const [name, node] of elementChildren(p)) {
      if (name === 'r') {
        para.runs.push(parseRun(node, prevEnd))
        const run = para.runs[para.runs.length - 1]
        prevEnd = { fontSizePt: run.fontSizePt, color: run.color, bold: run.bold, fontFamily: run.fontFamily }
      } else if (name === 'br') {
        para.runs.push({ text: '\n', ...prevEnd })
      }
    }
    // paragraph with no runs but a pPr still takes vertical space — keep it
    body.paragraphs.push(para)
  }
  return body
}

function parseShape(sp: XmlNode): PptxShape | undefined {
  const spPr = getChildren(sp, 'spPr')[0]
  if (!spPr) return undefined
  const xfrm = getChildren(spPr, 'xfrm')[0]
  const off = xfrm ? getChildren(xfrm, 'off')[0] : undefined
  const ext = xfrm ? getChildren(xfrm, 'ext')[0] : undefined
  const oa = attrs(off)
  const ea = attrs(ext)
  const prstGeom = getChildren(spPr, 'prstGeom')[0]
  const prst = attrs(prstGeom).prst as string | undefined
  const geometry: PptxShape['geometry'] =
    prst === 'ellipse' ? 'ellipse'
    : prst === 'roundRect' ? 'roundRect'
    : prst === 'rect' || prst === undefined ? 'rect'
    : 'other'
  const solidFill = getChildren(spPr, 'solidFill')[0]
  const fill = colorOf(solidFill ? getChildren(solidFill, 'srgbClr')[0] : undefined)
  const ln = getChildren(spPr, 'ln')[0]
  const lineColor = colorOf(ln ? getChildren(ln, 'solidFill').flatMap((sf) => getChildren(sf, 'srgbClr'))[0] : undefined)
  const shape: PptxShape = {
    xEmu: num(oa.x as string),
    yEmu: num(oa.y as string),
    widthEmu: num(ea.cx as string),
    heightEmu: num(ea.cy as string),
    geometry,
    fill,
    line: lineColor ? { color: lineColor, widthEmu: ln ? num(attrs(ln).w as string, 12700) : undefined } : undefined,
    rotationDeg: xfrm ? (num(attrs(xfrm).rot as string, 0) / 60000) : 0,
  }
  const txBody = getChildren(sp, 'txBody')[0]
  if (txBody) shape.textBody = parseTextBody(txBody)
  return shape
}

export async function parsePptx(pkg: OfficePackage): Promise<PptxDocument> {
  const presentation = await pkg.xml('ppt/presentation.xml')
  if (!presentation) throw new Error('ppt/presentation.xml missing — not a valid pptx?')
  const sldSz = getChildren(presentation, 'sldSz')[0]
  const sa = attrs(sldSz)
  const doc: PptxDocument = {
    slideWidthEmu: num(sa.cx as string, 9144000),
    slideHeightEmu: num(sa.cy as string, 6858000),
    slides: [],
  }
  // slide order from presentation rels
  const rels = await pkg.xml('ppt/_rels/presentation.xml.rels')
  const relMap = new Map<string, string>()
  if (rels) {
    for (const rel of getChildren(rels, 'Relationship')) {
      const a = attrs(rel)
      if (a.Id) relMap.set(a.Id, a.Target as string)
    }
  }
  const sldIdLst = getChildren(presentation, 'sldIdLst')[0]
  const slideIds = sldIdLst ? getChildren(sldIdLst, 'sldId') : []
  for (let i = 0; i < slideIds.length; i++) {
    const a = attrs(slideIds[i])
    const rid = (a['r:id'] ?? a.id) as string
    const target = relMap.get(rid) ?? ''
    const path = target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^\.\.\//, '')}`
    const slideRoot = await pkg.xml(path)
    const slide: PptxSlide = { index: i, widthEmu: doc.slideWidthEmu, heightEmu: doc.slideHeightEmu, shapes: [] }
    if (slideRoot) {
      const cSld = getChildren(slideRoot, 'cSld')[0]
      const spTree = cSld ? getChildren(cSld, 'spTree')[0] : undefined
      if (spTree) {
        for (const [name, node] of elementChildren(spTree)) {
          if (name === 'sp' || name === 'pic') {
            const shape = parseShape(node)
            if (shape) slide.shapes.push(shape)
          }
        }
      }
    }
    doc.slides.push(slide)
  }
  return doc
}
