/** Parse PPTX parts (presentation.xml, slides/slideN.xml) into PptxDocument. */
import type { OfficePackage } from '../core/zip'
import { attrs, elementChildren, getChildren, textOf, type XmlNode } from '../core/xml'
import { hexRgbToCss } from '../core/color'
import type { PptxDocument, PptxImageRef, PptxParagraph, PptxShape, PptxSlide, PptxTable, PptxTableCell, PptxTableRow, PptxTextBody, PptxTextRun } from './types'
import { sniffImageMime } from '../core/images'

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

/**
 * Placeholder geometry from a slide layout (or its master). Real decks rely on
 * this heavily: a title/body placeholder on the slide carries only `p:ph`, and
 * its position lives in the layout. Without it those shapes render at 0x0.
 */
async function layoutPlaceholderGeometry(
  pkg: OfficePackage,
  slidePath: string,
): Promise<Map<string, { x: number; y: number; w: number; h: number }>> {
  const out = new Map<string, { x: number; y: number; w: number; h: number }>()
  const dir = slidePath.slice(0, slidePath.lastIndexOf('/'))
  const base = slidePath.slice(slidePath.lastIndexOf('/') + 1)
  const rels = await pkg.xml(`${dir}/_rels/${base}.rels`)
  if (!rels) return out
  let layoutPath: string | undefined
  for (const rel of getChildren(rels, 'Relationship')) {
    const a = attrs(rel)
    if ((a.Type as string | undefined)?.endsWith('/slideLayout') && a.Target) {
      layoutPath = resolveTarget(slidePath, a.Target)
      break
    }
  }
  if (!layoutPath) return out

  const collect = (root: XmlNode | undefined): void => {
    if (!root) return
    const cSld = getChildren(root, 'cSld')[0]
    const spTree = cSld ? getChildren(cSld, 'spTree')[0] : undefined
    if (!spTree) return
    for (const [, node] of elementChildren(spTree)) {
      const spPr = getChildren(node, 'spPr')[0]
      const xfrm = spPr ? getChildren(spPr, 'xfrm')[0] : undefined
      if (!xfrm) continue
      const oa = attrs(getChildren(xfrm, 'off')[0])
      const ea = attrs(getChildren(xfrm, 'ext')[0])
      const nvSpPr = getChildren(node, 'nvSpPr')[0]
      const ph = nvSpPr ? getChildren(getChildren(nvSpPr, 'nvPr')[0], 'ph')[0] : undefined
      const pa = attrs(ph)
      const type = (pa.type as string | undefined) ?? 'body'
      const idx = pa.idx !== undefined ? parseInt(pa.idx as string, 10) || 0 : 0
      out.set(`${type}|${idx}`, {
        x: num(oa.x as string),
        y: num(oa.y as string),
        w: num(ea.cx as string),
        h: num(ea.cy as string),
      })
    }
  }

  collect(await pkg.xml(layoutPath))
  // a layout may itself defer geometry to the master
  const layoutDir = layoutPath.slice(0, layoutPath.lastIndexOf('/'))
  const layoutBase = layoutPath.slice(layoutPath.lastIndexOf('/') + 1)
  const layoutRels = await pkg.xml(`${layoutDir}/_rels/${layoutBase}.rels`)
  if (layoutRels) {
    for (const rel of getChildren(layoutRels, 'Relationship')) {
      const a = attrs(rel)
      if ((a.Type as string | undefined)?.endsWith('/slideMaster') && a.Target) {
        collect(await pkg.xml(resolveTarget(layoutPath, a.Target)))
        break
      }
    }
  }
  return out
}

function parseShape(sp: XmlNode, slideImages?: Map<string, PptxImageRef>): PptxShape | undefined {
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
  // p:ph marks a placeholder; its geometry usually comes from the layout
  const nvSpPr = getChildren(sp, 'nvSpPr')[0]
  const ph = nvSpPr ? getChildren(getChildren(nvSpPr, 'nvPr')[0], 'ph')[0] : undefined
  if (ph) {
    const pa = attrs(ph)
    shape.placeholder = {
      type: (pa.type as string | undefined) ?? 'body',
      idx: pa.idx !== undefined ? parseInt(pa.idx as string, 10) || 0 : 0,
    }
  }
  // p:pic: <p:blipFill><a:blip r:embed="rIdN"/><a:srcRect/></p:blipFill>
  if (slideImages) {
    const blipFill = getChildren(sp, 'blipFill')[0]
    if (blipFill) {
      const blip = getChildren(blipFill, 'blip')[0]
      const rid = attrs(blip).embed as string | undefined
      const image = rid ? slideImages.get(rid) : undefined
      if (image) {
        const srcRect = getChildren(blipFill, 'srcRect')[0]
        shape.image = image
        if (srcRect) {
          const sa = attrs(srcRect)
          // srcRect units are 1/1000 of a percent
          const frac = (v: string | undefined): number => (v !== undefined ? parseFloat(v) / 100000 : 0)
          shape.image = {
            ...image,
            srcRect: {
              l: frac(sa.l as string | undefined),
              t: frac(sa.t as string | undefined),
              r: frac(sa.r as string | undefined),
              b: frac(sa.b as string | undefined),
            },
          }
        }
      }
    }
  }
  return shape
}

/** Resolve a relationship Target against the part's directory. */
function resolveTarget(partPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const base = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/')) : ''
  const segs = base ? base.split('/') : []
  for (const seg of target.split('/')) {
    if (seg === '.' || seg === '') continue
    if (seg === '..') segs.pop()
    else segs.push(seg)
  }
  return segs.join('/')
}

/** Load image parts referenced by a part's .rels file, keyed by rId. */
async function loadSlideImages(pkg: OfficePackage, partPath: string): Promise<Map<string, PptxImageRef>> {
  const out = new Map<string, PptxImageRef>()
  const relsPath = `${partPath.slice(0, partPath.lastIndexOf('/'))}/_rels/${partPath.slice(partPath.lastIndexOf('/') + 1)}.rels`
  const rels = await pkg.xml(relsPath)
  if (!rels) return out
  // one PptxImageRef per media part, so multiple rIds for the same part dedupe
  const byPath = new Map<string, PptxImageRef>()
  for (const rel of getChildren(rels, 'Relationship')) {
    const a = attrs(rel)
    const type = a.Type as string | undefined
    if (!type || !type.includes('/image')) continue
    const path = resolveTarget(partPath, (a.Target as string) ?? '')
    let ref = byPath.get(path)
    if (!ref) {
      const data = await pkg.bytes(path)
      if (!data) continue
      ref = { data, mime: sniffImageMime(data) }
      byPath.set(path, ref)
    }
    if (a.Id) out.set(a.Id, ref)
  }
  return out
}

/** Depth-first search for the first a:srgbClr under a node. */
function firstSrgb(node: XmlNode | undefined): string | undefined {
  if (!node) return undefined
  for (const [name, child] of elementChildren(node)) {
    if (name === 'srgbClr') return colorOf(child)
    const found = firstSrgb(child)
    if (found) return found
  }
  return undefined
}

interface TableStyleEntry {
  fills: { firstRow?: string; band1?: string; band2?: string; wholeTable?: string; lastRow?: string; firstCol?: string }
  firstRowTextColor?: string
}

async function readTableStyles(pkg: OfficePackage): Promise<Map<string, TableStyleEntry>> {
  const out = new Map<string, TableStyleEntry>()
  const root = await pkg.xml('ppt/tableStyles.xml')
  if (!root) return out
  // the part root is a:tblStyleLst itself, but tolerate a wrapper
  const list = getChildren(root, 'tblStyleLst')[0] ?? root
  for (const style of getChildren(list, 'tblStyle')) {
    const id = attrs(style).styleId as string | undefined
    if (!id) continue
    const entry: TableStyleEntry = { fills: {} }
    // whole-table fill lives directly on the style
    const tableFill = getChildren(getChildren(style, 'tblPr')[0], 'solidFill')[0]
    const whole = tableFill ? colorOf(getChildren(tableFill, 'srgbClr')[0]) : undefined
    if (whole) entry.fills.wholeTable = whole
    for (const region of getChildren(style, 'tblStylePr')) {
      const kind = attrs(region).type as string | undefined
      if (!kind) continue
      const tcPr = getChildren(region, 'tcPr')[0]
      const solid = tcPr ? getChildren(tcPr, 'solidFill')[0] : undefined
      const fill = solid ? colorOf(getChildren(solid, 'srgbClr')[0]) : undefined
      switch (kind) {
        case 'firstRow':
          if (fill) entry.fills.firstRow = fill
          // a:txStyles is a sibling of a:tcPr under a:tblStylePr
          entry.firstRowTextColor = firstSrgb(getChildren(region, 'txStyles')[0])
          break
        case 'band1Horz': if (fill) entry.fills.band1 = fill; break
        case 'band2Horz': if (fill) entry.fills.band2 = fill; break
        case 'lastRow': if (fill) entry.fills.lastRow = fill; break
        case 'firstCol': if (fill) entry.fills.firstCol = fill; break
        default: break
      }
    }
    out.set(id, entry)
  }
  return out
}

/** p:graphicFrame -> a:graphic/a:graphicData/a:tbl */
function parseGraphicFrame(frame: XmlNode, tableStyles?: Map<string, TableStyleEntry>): PptxShape | undefined {
  const xfrm = getChildren(frame, 'xfrm')[0]
  const off = xfrm ? getChildren(xfrm, 'off')[0] : undefined
  const ext = xfrm ? getChildren(xfrm, 'ext')[0] : undefined
  const oa = attrs(off)
  const ea = attrs(ext)
  const graphic = getChildren(frame, 'graphic')[0]
  const graphicData = graphic ? getChildren(graphic, 'graphicData')[0] : undefined
  const tbl = graphicData ? getChildren(graphicData, 'tbl')[0] : undefined
  if (!tbl) return undefined

  const table: PptxTable = { colWidthsEmu: [], rows: [] }
  const tblPr = getChildren(tbl, 'tblPr')[0]
  if (tblPr) {
    const ta = attrs(tblPr)
    table.firstRow = ta.firstRow === '1' || ta.firstRow === 'true'
    table.bandRow = ta.bandRow === '1' || ta.bandRow === 'true'
    const styleId = getChildren(tblPr, 'tableStyleId')[0]
      ? textOf(getChildren(tblPr, 'tableStyleId')[0]).trim() || undefined
      : undefined
    if (styleId) {
      table.styleId = styleId
      const entry = tableStyles?.get(styleId)
      if (entry) {
        table.styleFills = entry.fills
        table.firstRowTextColor = entry.firstRowTextColor
      }
    }
  }
  const grid = getChildren(tbl, 'tblGrid')[0]
  if (grid) {
    for (const col of getChildren(grid, 'gridCol')) {
      table.colWidthsEmu.push(num(attrs(col).w as string))
    }
  }
  for (const tr of getChildren(tbl, 'tr')) {
    const ta = attrs(tr)
    const row: PptxTableRow = { cells: [], heightEmu: ta.h !== undefined ? num(ta.h as string) : undefined }
    for (const tc of getChildren(tr, 'tc')) {
      const ca = attrs(tc)
      const cell: PptxTableCell = {
        paragraphs: [],
        gridSpan: ca.gridSpan !== undefined ? num(ca.gridSpan as string) : 1,
        rowSpan: ca.rowSpan !== undefined ? num(ca.rowSpan as string) : 1,
        // hMerge/vMerge without a value marks a cell absorbed by a merge
        merged: ca.hMerge !== undefined || ca.vMerge !== undefined,
      }
      const tcPr = getChildren(tc, 'tcPr')[0]
      if (tcPr) {
        const solid = getChildren(tcPr, 'solidFill')[0]
        const color = colorOf(solid ? getChildren(solid, 'srgbClr')[0] : undefined)
        if (color) cell.fill = color
      }
      const txBody = getChildren(tc, 'txBody')[0]
      if (txBody) cell.paragraphs = parseTextBody(txBody).paragraphs
      row.cells.push(cell)
    }
    table.rows.push(row)
  }

  return {
    xEmu: num(oa.x as string),
    yEmu: num(oa.y as string),
    widthEmu: num(ea.cx as string),
    heightEmu: num(ea.cy as string),
    geometry: 'rect',
    rotationDeg: xfrm ? num(attrs(xfrm).rot as string, 0) / 60000 : 0,
    table,
  }
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
    images: [],
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
  let tableStyles: Map<string, TableStyleEntry> | undefined
  const sldIdLst = getChildren(presentation, 'sldIdLst')[0]
  const slideIds = sldIdLst ? getChildren(sldIdLst, 'sldId') : []
  for (let i = 0; i < slideIds.length; i++) {
    const a = attrs(slideIds[i])
    const rid = (a['r:id'] ?? a.id) as string
    const target = relMap.get(rid) ?? ''
    const path = target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^\.\.\//, '')}`
    const slideRoot = await pkg.xml(path)
    const slideImages = await loadSlideImages(pkg, path)
    if (!tableStyles) tableStyles = await readTableStyles(pkg)
    const slide: PptxSlide = { index: i, widthEmu: doc.slideWidthEmu, heightEmu: doc.slideHeightEmu, shapes: [] }
    if (slideRoot) {
      const cSld = getChildren(slideRoot, 'cSld')[0]
      const spTree = cSld ? getChildren(cSld, 'spTree')[0] : undefined
      if (spTree) {
        for (const [name, node] of elementChildren(spTree)) {
          if (name === 'sp' || name === 'pic') {
            const shape = parseShape(node, slideImages)
            if (shape) slide.shapes.push(shape)
          } else if (name === 'graphicFrame') {
            const frame = parseGraphicFrame(node, tableStyles)
            if (frame) slide.shapes.push(frame)
          }
        }
      }
    }
    // fill in placeholder geometry from the layout/master
    const needsGeometry = slide.shapes.some((s) => s.placeholder && (s.widthEmu === 0 || s.heightEmu === 0))
    if (needsGeometry) {
      const inherited = await layoutPlaceholderGeometry(pkg, path)
      for (const shape of slide.shapes) {
        if (!shape.placeholder || (shape.widthEmu !== 0 && shape.heightEmu !== 0)) continue
        const geom = inherited.get(`${shape.placeholder.type}|${shape.placeholder.idx}`)
        if (geom) {
          shape.xEmu = geom.x
          shape.yEmu = geom.y
          shape.widthEmu = geom.w
          shape.heightEmu = geom.h
        }
      }
    }
    doc.slides.push(slide)
  }
  // assign document-wide image indices in first-use order
  for (const slide of doc.slides) {
    for (const shape of slide.shapes) {
      if (!shape.image) continue
      let idx = doc.images.indexOf(shape.image)
      if (idx < 0) {
        idx = doc.images.length
        doc.images.push(shape.image)
      }
      shape.imageIndex = idx
    }
  }
  return doc
}
