/** Parse XLSX parts into the XlsxDocument model. */
import type { OfficePackage } from '../core/zip'
import { attrs, getChildren, textOf, type XmlNode } from '../core/xml'
import type { XlsxCell, XlsxDocument, XlsxMergeRange, XlsxRow, XlsxSheet } from './types'
import { parseWorksheetDrawings, collectXlsxImages } from './drawing'
import { computeMetrics } from './render'
import { resolvePartTarget } from '../drawing/parts'
import { parseThemeContext, type ThemeContext } from '../drawing/style'
import { evaluateWorkbookFormulas } from './formula/workbook'

/** Convert "A1" / "BC23" to 0-based [row, col]. */
export function parseRef(ref: string): [number, number] {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)
  if (!m) return [0, 0]
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  return [parseInt(m[2], 10) - 1, col - 1]
}

/** "A1:B3" -> 0-based inclusive range. */
export function parseMergeRange(ref: string): XlsxMergeRange {
  const [a, b] = ref.split(':')
  const [r1, c1] = parseRef(a)
  const [r2, c2] = b ? parseRef(b) : [r1, c1]
  return { minRow: Math.min(r1, r2), minCol: Math.min(c1, c2), maxRow: Math.max(r1, r2), maxCol: Math.max(c1, c2) }
}

async function sharedStrings(pkg: OfficePackage): Promise<string[]> {
  const root = await pkg.xml('xl/sharedStrings.xml')
  if (!root) return []
  const out: string[] = []
  for (const si of getChildren(root, 'si')) {
    // si may be plain t, or rich text runs (r/t), or phonetic — concatenate all t
    let s = ''
    const direct = si['t']
    if (direct !== undefined) {
      s += textOf(asNode(direct))
    }
    for (const r of getChildren(si, 'r')) {
      s += textOf(getChildren(r, 't')[0])
    }
    out.push(s)
  }
  return out
}

function asNode(v: unknown): XmlNode {
  if (v && typeof v === 'object') return v as XmlNode
  return { '@attrs': {}, '#text': v === null || v === undefined ? '' : String(v) }
}

interface Styles {
  numFmtIds: number[]
  fonts: Array<{ bold: boolean; italic: boolean; sizePt?: number; color?: string }>
  fills: Array<{ rgb?: string; pattern?: string }>
  borders: Array<{ left?: string; right?: string; top?: string; bottom?: string }>
  xfs: Array<{ numFmtId: number; fontId: number; fillId: number; borderId: number; textRotation?: number; horizontal?: string; vertical?: string; wrapText?: boolean }>
}

async function parseStyles(pkg: OfficePackage): Promise<Styles> {
  const root = await pkg.xml('xl/styles.xml')
  const styles: Styles = { numFmtIds: [], fonts: [], fills: [], borders: [], xfs: [] }
  if (!root) return styles
  for (const f of getChildren(root, 'fonts').flatMap((n) => getChildren(n, 'font'))) {
    const font = {
      bold: f['b'] !== undefined,
      italic: f['i'] !== undefined,
      sizePt: f['sz'] !== undefined ? parseFloat(attrs(asNode(f['sz'])).val ?? '11') : undefined,
      color: attrs(asNode(f['color'])).rgb as string | undefined,
    }
    styles.fonts.push(font)
  }
  for (const fillsN of getChildren(root, 'fills')) {
    for (const fill of getChildren(fillsN, 'fill')) {
      const pattern = getChildren(fill, 'patternFill')[0]
      const fg = pattern ? getChildren(pattern, 'fgColor')[0] : undefined
      styles.fills.push({ rgb: attrs(asNode(fg)).rgb as string | undefined, pattern: attrs(asNode(pattern)).patternType as string | undefined })
    }
  }
  for (const bordersN of getChildren(root, 'borders')) {
    for (const border of getChildren(bordersN, 'border')) {
      const side = (n: string) => {
        const node = getChildren(border, n)[0]
        return node ? (attrs(node).style as string | undefined) : undefined
      }
      styles.borders.push({ left: side('left'), right: side('right'), top: side('top'), bottom: side('bottom') })
    }
  }
  for (const cellXfsN of getChildren(root, 'cellXfs')) {
    for (const xf of getChildren(cellXfsN, 'xf')) {
      const a = attrs(xf)
      const rotation = Number(attrs(getChildren(xf, 'alignment')[0]).textRotation)
      const textRotation = Number.isInteger(rotation) && (rotation === 255 || (rotation >= 1 && rotation <= 180)) ? rotation : undefined
      const align = attrs(getChildren(xf, 'alignment')[0])
      const horizontal = typeof align.horizontal === 'string' && align.horizontal !== '' ? align.horizontal : undefined
      const vertical = typeof align.vertical === 'string' && align.vertical !== '' ? align.vertical : undefined
      const wrapText = align.wrapText === '1' || align.wrapText === 'true' ? true : align.wrapText === '0' || align.wrapText === 'false' ? false : undefined
      styles.xfs.push({
        numFmtId: parseInt(a.numFmtId ?? '0', 10),
        fontId: parseInt(a.fontId ?? '0', 10),
        fillId: parseInt(a.fillId ?? '0', 10),
        borderId: parseInt(a.borderId ?? '0', 10),
        textRotation,
        horizontal,
        vertical,
        wrapText,
      })
    }
  }
  return styles
}

export async function parseXlsx(pkg: OfficePackage): Promise<XlsxDocument> {
  const workbook = await pkg.xml('xl/workbook.xml')
  if (!workbook) throw new Error('xl/workbook.xml missing — not a valid xlsx?')
  const rels = await pkg.xml('xl/_rels/workbook.xml.rels')
  const relMap = new Map<string, string>()
  let themePath: string | undefined
  if (rels) {
    for (const rel of getChildren(rels, 'Relationship')) {
      const a = attrs(rel)
      if (a.Id && a.Target && a.TargetMode !== 'External') relMap.set(a.Id, a.Target as string)
      if (a.Type?.endsWith('/theme') && a.TargetMode !== 'External' && a.Target) themePath = resolvePartTarget('xl/workbook.xml', a.Target)
    }
  }
  const strings = await sharedStrings(pkg)
  const styles = await parseStyles(pkg)
  let theme: ThemeContext = parseThemeContext()
  if (themePath) {
    try { theme = parseThemeContext(await pkg.xmlOrdered(themePath)) }
    catch { /* optional theme cannot invalidate usable cells */ }
  }

  const sheets: XlsxSheet[] = []
  const sheetsNode = getChildren(workbook, 'sheets')[0]
  for (const sheetNode of getChildren(sheetsNode, 'sheet')) {
    const a = attrs(sheetNode)
    const name = (a.name as string) ?? 'Sheet'
    const target = relMap.get(a['r:id'] ?? a.id ?? '') ?? ''
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\.\//, '')}`
    const sheet = await parseSheet(pkg, path, name, strings, styles, theme)
    sheets.push(sheet)
  }
  const drawingCoverage = sheets.flatMap((sheet, unit) => {
    const entries = (sheet.drawingCoverage ?? []).map(entry => ({ ...entry, unit }))
    const metrics = sheet.drawings?.length ? computeMetrics(sheet) : undefined
    const requested = metrics?.requestedDrawingBounds, retained = metrics?.retainedDrawingBounds
    if (requested && retained && (requested.right > retained.right || requested.bottom > retained.bottom)) entries.push({
      partPath: sheet.sourcePartPath ?? 'xl/worksheets/unknown.xml', treePath: sheet.name, element: 'viewport', feature: 'viewport',
      status: 'unsupported' as const, selectedRepresentation: 'none' as const, representation: 'native' as const, reason: 'drawing viewport limit', scope: 'diagnostic' as const, unit,
      limit: requested.right > 16384 || requested.bottom > 16384 ? 16384 : 16777216,
      requestedExtent: { width: requested.right, height: requested.bottom }, retainedExtent: { width: retained.right, height: retained.bottom },
    })
    return entries
  })
  const calcPrNode = getChildren(workbook, 'calcPr')[0]
  const calcPrAttrs = calcPrNode ? attrs(calcPrNode) : {}
  const fullCalcOnLoad = calcPrAttrs.fullCalcOnLoad === '1' || calcPrAttrs.fullCalcOnLoad === 'true'

  const doc: XlsxDocument = { sheets, images: collectXlsxImages(sheets), drawingCoverage }
  evaluateWorkbookFormulas(doc, { fullCalcOnLoad })
  return doc
}

async function parseSheet(
  pkg: OfficePackage,
  path: string,
  name: string,
  strings: string[],
  styles: Styles,
  theme: ThemeContext,
): Promise<XlsxSheet> {
  const root = await pkg.xml(path)
  const sheet: XlsxSheet = { name, sourcePartPath: path, rows: [], cols: [], merges: [], mergeRanges: [] }
  if (!root) return sheet
  const data = getChildren(root, 'worksheet')[0] ?? root
  const colsNode = getChildren(data, 'cols')[0]
  if (colsNode) {
    for (const col of getChildren(colsNode, 'col')) {
      const a = attrs(col)
      sheet.cols.push({
        min: parseInt(a.min ?? '1', 10) - 1,
        max: parseInt(a.max ?? '1', 10) - 1,
        widthChars: a.width !== undefined ? parseFloat(a.width as string) : undefined,
        hidden: a.hidden === '1' || a.hidden === 'true',
      })
    }
  }
  const mergeCells = getChildren(data, 'mergeCells')[0]
  if (mergeCells) {
    for (const mc of getChildren(mergeCells, 'mergeCell')) {
      const ref = attrs(mc).ref as string
      if (ref) sheet.merges.push(ref)
    }
    sheet.mergeRanges = sheet.merges.map(parseMergeRange)
  }
  const setupPr = getChildren(getChildren(data, 'sheetPr')[0], 'pageSetUpPr')[0]
  const setupNode = getChildren(data, 'pageSetup')[0]
  if (setupNode || setupPr) {
    const a = attrs(setupNode)
    const num = (v: unknown): number | undefined => {
      const n = typeof v === 'string' ? parseFloat(v) : NaN
      return Number.isFinite(n) ? n : undefined
    }
    const fitToPageAttr = setupPr ? attrs(setupPr).fitToPage : undefined
    sheet.pageSetup = {
      ...(num(a.paperSize) !== undefined ? { paperSizeId: num(a.paperSize) } : {}),
      ...(a.orientation === 'landscape' || a.orientation === 'portrait' ? { orientation: a.orientation } : {}),
      ...(num(a.scale) !== undefined ? { scale: num(a.scale) } : {}),
      ...(num(a.fitToWidth) !== undefined ? { fitToWidth: num(a.fitToWidth) } : {}),
      ...(num(a.fitToHeight) !== undefined ? { fitToHeight: num(a.fitToHeight) } : {}),
      ...(fitToPageAttr === '1' || fitToPageAttr === 'true' ? { fitToPage: true } : fitToPageAttr !== undefined ? { fitToPage: false } : {}),
    }
  }
  // Excel defaults (inches) when the element is absent.
  const marginsNode = getChildren(data, 'pageMargins')[0]
  const inches = (v: unknown, fallback: number): number => {
    const n = typeof v === 'string' ? parseFloat(v) : NaN
    return Number.isFinite(n) && n >= 0 ? n : fallback
  }
  const ma = attrs(marginsNode)
  sheet.pageMargins = {
    left: inches(ma.left, 0.7), right: inches(ma.right, 0.7),
    top: inches(ma.top, 0.75), bottom: inches(ma.bottom, 0.75),
    header: inches(ma.header, 0.3), footer: inches(ma.footer, 0.3),
  }
  const sheetData = getChildren(data, 'sheetData')[0]
  for (const rowNode of getChildren(sheetData, 'row')) {
    const ra = attrs(rowNode)
    const row: XlsxRow = {
      index: parseInt(ra.r ?? '1', 10) - 1,
      heightPt: ra.ht !== undefined ? parseFloat(ra.ht as string) : undefined,
      customHeight: ra.customHeight === '1',
      hidden: ra.hidden === '1' || ra.hidden === 'true',
      cells: [],
    }
    for (const cNode of getChildren(rowNode, 'c')) {
      const ca = attrs(cNode)
      const ref = (ca.r as string) ?? ''
      const [rowIdx, colIdx] = parseRef(ref)
      const t = ca.t as string | undefined
      const sIdx = parseInt(ca.s ?? '0', 10)
      let value: string | number | boolean | null = null
      let formula: string | undefined
      let sharedFormula: { si: number; ref?: string } | undefined
      const vNode = getChildren(cNode, 'v')[0]
      const isNode = getChildren(cNode, 'is')[0]
      let calcAlways = ca.ca === '1' || ca.ca === 'true'
      const fNode = getChildren(cNode, 'f')[0]
      if (fNode) {
        const rawF = textOf(fNode)
        if (rawF !== '') formula = rawF
        const fa = attrs(fNode)
        // ECMA-376 Part 1 §18.3.1.40: ca attribute on <f>
        if (fa.ca === '1' || fa.ca === 'true') {
          calcAlways = true
        }
        if (fa.t === 'shared') {
          const si = parseInt(fa.si ?? '0', 10)
          sharedFormula = {
            si: Number.isFinite(si) ? si : 0,
            ref: (fa.ref as string) || undefined,
          }
        }
      }
      if (t === 's') {
        const idx = vNode ? parseInt(textOf(vNode), 10) : NaN
        value = Number.isFinite(idx) ? (strings[idx] ?? '') : ''
      } else if (t === 'inlineStr') {
        value = isNode ? textOf(isNode) : ''
      } else if (t === 'b') {
        value = vNode ? textOf(vNode) === '1' : false
      } else if (t === 'str') {
        value = vNode ? textOf(vNode) : ''
      } else if (t === 'e') {
        value = vNode ? textOf(vNode) : null
      } else {
        // numeric
        const raw = vNode ? textOf(vNode) : ''
        value = raw !== '' && Number.isFinite(parseFloat(raw)) ? parseFloat(raw) : raw === '' ? null : raw
      }
      const cell: XlsxCell = {
        ref,
        row: rowIdx,
        col: colIdx,
        value,
        styleIndex: Number.isFinite(sIdx) ? sIdx : 0,
        formula,
        sharedFormula,
        ca: calcAlways ? true : undefined,
      }
      const xf = styles.xfs[cell.styleIndex]
      if (xf) {
        const font = styles.fonts[xf.fontId]
        const fill = styles.fills[xf.fillId]
        const border = styles.borders[xf.borderId]
        cell.style = {
          numFmtId: xf.numFmtId,
          bold: font?.bold,
          italic: font?.italic,
          fontSizePt: font?.sizePt,
          color: font?.color,
          fillColor: fill?.pattern === 'solid' ? fill.rgb : undefined,
          borders: border ? { left: border.left, right: border.right, top: border.top, bottom: border.bottom } : undefined,
          textRotation: xf.textRotation,
          horizontal: xf.horizontal,
          vertical: xf.vertical,
          wrapText: xf.wrapText,
        }
      }
      row.cells.push(cell)
    }
    sheet.rows.push(row)
  }
  await parseWorksheetDrawings(pkg, sheet, path, data, theme)
  return sheet
}

export { parseStyles }
export type { Styles }
