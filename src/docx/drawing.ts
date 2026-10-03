/** Embedded DrawingML objects use their cached geometry/data, not a screenshot. */
import { attrs, elementChildren, getChildren, textOf, type XmlNode } from '../core/xml'
import type { OfficePackage } from '../core/zip'
import { emuToPx } from '../core/geometry'
import type { DocxDrawing, DocxDrawingShape, DocxImage, DocxParagraph, DocxTextRun } from './types'
import { fontFamilyCss, type DocxTheme } from './styles'
import { parseInk, paintInk } from './ink'

const num = (v: string | undefined, fallback = 0): number =>
  v !== undefined && Number.isFinite(Number(v)) ? Number(v) : fallback
const child = (n: XmlNode | undefined, name: string) => getChildren(n, name)[0]
function descendants(n: XmlNode | undefined, name: string): XmlNode[] {
  if (!n) return []
  return elementChildren(n).flatMap(([tag, node]) => [...(tag === name ? [node] : []), ...descendants(node, name)])
}
function target(part: string, path: string): string {
  const bits = path.startsWith('/') ? [] : part.split('/').slice(0, -1)
  for (const bit of path.split('/')) {
    if (bit === '..') bits.pop()
    else if (bit && bit !== '.') bits.push(bit)
  }
  return bits.join('/')
}

function color(fill: XmlNode | undefined, theme: DocxTheme): string | undefined {
  const c = child(fill, 'srgbClr') ?? child(fill, 'schemeClr') ?? child(fill, 'sysClr') ?? child(fill, 'prstClr')
  if (!c) return undefined
  const a = attrs(c)
  const preset: Record<string, string> = {
    black: '000000',
    white: 'FFFFFF',
    red: 'FF0000',
    blue: '0000FF',
    green: '008000'
  }
  const value = child(fill, 'schemeClr')
    ? theme.colors.get(a.val)
    : child(fill, 'prstClr')
      ? preset[a.val]
      : child(fill, 'sysClr')
        ? a.lastClr
        : a.val
  if (!value || !/^[\da-f]{6}$/i.test(value)) return undefined
  let channels = [0, 2, 4].map((i) => parseInt(value.slice(i, i + 2), 16) / 255)
  for (const [name, transform] of elementChildren(c)) {
    const amount = Math.min(1, Math.max(0, num(attrs(transform).val) / 100000))
    if (name === 'tint' || name === 'shade')
      channels = channels.map((channel) => {
        const linear = channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4
        const adjusted = name === 'tint' ? linear * amount + 1 - amount : linear * amount
        return adjusted <= 0.0031308 ? adjusted * 12.92 : 1.055 * adjusted ** (1 / 2.4) - 0.055
      })
    else if (name === 'lumMod') channels = channels.map((channel) => channel * amount)
    else if (name === 'lumOff') channels = channels.map((channel) => Math.min(1, channel + amount))
  }
  return channels
    .map((c) =>
      Math.round(c * 255)
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase()
}

function shape(node: XmlNode, theme: DocxTheme): DocxDrawingShape {
  const pr = child(node, 'spPr'),
    transform = child(pr, 'xfrm')
  const offset = attrs(child(transform, 'off')),
    extent = attrs(child(transform, 'ext'))
  const ln = child(pr, 'ln'),
    lineColor = color(child(ln, 'solidFill'), theme)
  const textColor = color(child(child(node, 'style'), 'fontRef'), theme)
  const fontFamily = theme.fonts.get('minorHAnsi') ?? 'Calibri'
  const paragraphs = getChildren(child(node, 'txBody'), 'p').map((p) => ({
    align:
      ({ ctr: 'center', r: 'right', just: 'justify' } as const)[attrs(child(p, 'pPr')).algn as 'ctr' | 'r' | 'just'] ??
      ('left' as const),
    runs: getChildren(p, 'r').map((r) => {
      const pr = child(r, 'rPr'),
        a = attrs(pr)
      const family = attrs(child(pr, 'latin')).typeface
      return {
        text: textOf(child(r, 't')),
        fontFamily: theme.fonts.get(family) ?? family ?? fontFamily,
        fontSizePt: num(a.sz, 2400) / 100,
        color: color(child(pr, 'solidFill'), theme) ?? textColor,
        bold: a.b === '1',
        italic: a.i === '1'
      }
    })
  }))
  return {
    xEmu: num(offset.x),
    yEmu: num(offset.y),
    widthEmu: num(extent.cx),
    heightEmu: num(extent.cy),
    geometry: attrs(child(pr, 'prstGeom')).prst ?? 'rect',
    rotationDeg: num(attrs(transform).rot) / 60000,
    fill: color(child(pr, 'solidFill'), theme),
    line: lineColor ? { color: lineColor, widthEmu: num(attrs(ln).w, 12700) } : undefined,
    paragraphs,
    fontFamily,
    textColor
  }
}

/** Cache points are indexed; missing points must retain their category slot. */
function cacheValues(node: XmlNode | undefined): Array<string | undefined> {
  const cache =
    child(child(node, 'strRef'), 'strCache') ??
    child(child(node, 'numRef'), 'numCache') ??
    child(node, 'strLit') ??
    child(node, 'numLit')
  const out: Array<string | undefined> = []
  for (const p of getChildren(cache, 'pt')) out[num(attrs(p).idx)] = textOf(child(p, 'v'))
  return Array.from({ length: out.length }, (_, index) => out[index])
}

function chart(root: XmlNode, theme: DocxTheme): DocxDrawing | undefined {
  const chart = child(root, 'chart'),
    plot = child(chart, 'plotArea'),
    bars = child(plot, 'barChart')
  if (
    !bars ||
    attrs(child(bars, 'barDir')).val === 'bar' ||
    !['clustered', undefined].includes(attrs(child(bars, 'grouping')).val)
  )
    return undefined
  const series = getChildren(bars, 'ser').map((s, index) => {
    const tx = child(s, 'tx')
    return {
      name: textOf(child(tx, 'v')) || cacheValues(tx)[0] || `Series ${index + 1}`,
      values: cacheValues(child(s, 'val')).map((v) =>
        v === undefined || !Number.isFinite(Number(v)) ? undefined : Number(v)
      ),
      color:
        color(child(child(s, 'spPr'), 'solidFill'), theme) ?? theme.colors.get(`accent${(index % 6) + 1}`) ?? '4F81BD'
    }
  })
  const categories = cacheValues(child(getChildren(bars, 'ser')[0], 'cat')).map((v) => v ?? '')
  const title = child(chart, 'title')
  const titleText = descendants(child(title, 'tx'), 't').map(textOf).join('') || cacheValues(child(title, 'tx'))[0]
  const locale = attrs(child(root, 'lang')).val ?? ''
  const axis = child(plot, 'valAx'),
    scaling = child(axis, 'scaling')
  const optional = (n: XmlNode | undefined) => (n ? num(attrs(n).val) : undefined)
  return {
    kind: 'chart',
    categories,
    series,
    title: title ? titleText || (locale.startsWith('zh') ? '圖表標題' : 'Chart Title') : undefined,
    min: optional(child(scaling, 'min')),
    max: optional(child(scaling, 'max')),
    majorUnit: optional(child(axis, 'majorUnit')),
    gapWidth: Math.max(0, num(attrs(child(bars, 'gapWidth')).val, 150)),
    overlap: Math.max(-100, Math.min(100, num(attrs(child(bars, 'overlap')).val))),
    legend: !!child(chart, 'legend'),
    fontFamily: theme.fonts.get('minorHAnsi') ?? 'Calibri',
    fontSizePt: num(attrs(descendants(child(axis, 'txPr'), 'defRPr')[0]).sz, 900) / 100
  }
}

export async function loadDrawingParts(
  pkg: OfficePackage,
  root: XmlNode,
  part: string,
  theme: DocxTheme,
  parseParagraph: (p: XmlNode) => DocxParagraph
): Promise<Map<XmlNode, DocxImage>> {
  const out = new Map<XmlNode, DocxImage>(),
    rels = new Map<string, string>()
  const dir = part.slice(0, part.lastIndexOf('/')),
    base = part.slice(part.lastIndexOf('/') + 1)
  for (const rel of getChildren(await pkg.xml(`${dir}/_rels/${base}.rels`), 'Relationship')) {
    const a = attrs(rel)
    if (a.TargetMode !== 'External' && a.Id && a.Target) rels.set(a.Id, target(part, a.Target))
  }
  for (const drawing of descendants(root, 'drawing')) {
    const wp = child(drawing, 'inline') ?? child(drawing, 'anchor'),
      graphic = child(child(wp, 'graphic'), 'graphicData')
    if (!wp || !graphic) continue
    let vector: DocxDrawing | undefined
    const ids = child(graphic, 'relIds')
    if (ids) {
      const path = rels.get(attrs(ids).dm)
      const data = path ? await pkg.xml(path) : undefined
      const cachedId = attrs(descendants(data, 'dataModelExt')[0]).relId
      let cachedPath = rels.get(cachedId)
      // Some producers put the cached drawing relationship on the diagram data part.
      if (!cachedPath && path && cachedId) {
        const slash = path.lastIndexOf('/')
        for (const rel of getChildren(
          await pkg.xml(`${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`),
          'Relationship'
        )) {
          const a = attrs(rel)
          if (a.Id === cachedId && a.TargetMode !== 'External') cachedPath = target(path, a.Target)
        }
      }
      const cached = cachedPath ? await pkg.xml(cachedPath) : undefined
      if (cached)
        vector = { kind: 'diagram', shapes: getChildren(child(cached, 'spTree'), 'sp').map((s) => shape(s, theme)) }
    }
    const chartRef = child(graphic, 'chart')
    if (chartRef) {
      const path = rels.get(attrs(chartRef).id),
        root = path ? await pkg.xml(path) : undefined
      if (root) vector = chart(root, theme)
    }
    const content = child(graphic, 'contentPart')
    if (content && attrs(graphic).uri?.endsWith('/wordprocessingInk')) {
      const path = rels.get(attrs(content).id)
      const root = path ? await pkg.xml(path) : undefined
      if (root) vector = parseInk(root)
    }
    const wsp = child(graphic, 'wsp')
    if (wsp) {
      const body = attrs(child(wsp, 'bodyPr')),
        pr = child(wsp, 'spPr'),
        ln = child(pr, 'ln'),
        lineColor = color(child(ln, 'solidFill'), theme)
      vector = {
        kind: 'textbox',
        paragraphs: getChildren(child(child(wsp, 'txbx'), 'txbxContent'), 'p').map(parseParagraph),
        vertical: body.vert === 'eaVert' || body.vert === 'vert',
        fontFamily: theme.fonts.get('minorHAnsi') ?? 'Calibri',
        fontSizePt: 12,
        insets: {
          left: num(body.lIns, 91440),
          top: num(body.tIns, 45720),
          right: num(body.rIns, 91440),
          bottom: num(body.bIns, 45720)
        },
        fill: color(child(pr, 'solidFill'), theme),
        line: lineColor ? { color: lineColor, widthEmu: num(attrs(ln).w, 6350) } : undefined
      }
    }
    if (vector) {
      const extent = attrs(child(wp, 'extent'))
      out.set(drawing, { data: new Uint8Array(), widthEmu: num(extent.cx), heightEmu: num(extent.cy), drawing: vector })
    }
  }
  return out
}

export function paintDrawing(drawing: DocxDrawing, ctx: CanvasRenderingContext2D, width: number, height: number): void {
  if (drawing.kind === 'ink') paintInk(drawing, ctx, width, height)
  else if (drawing.kind === 'diagram') {
    for (const s of drawing.shapes) {
      const x = emuToPx(s.xEmu),
        y = emuToPx(s.yEmu),
        w = emuToPx(s.widthEmu),
        h = emuToPx(s.heightEmu)
      ctx.save()
      ctx.translate(x + w / 2, y + h / 2)
      ctx.rotate(((s.rotationDeg ?? 0) * Math.PI) / 180)
      ctx.translate(-w / 2, -h / 2)
      ctx.beginPath()
      if (s.geometry === 'ellipse') ctx.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2)
      else if (s.geometry === 'rightArrow') {
        ctx.moveTo(0, h * 0.2)
        ctx.lineTo(w * 0.5, h * 0.2)
        ctx.lineTo(w * 0.5, 0)
        ctx.lineTo(w, h / 2)
        ctx.lineTo(w * 0.5, h)
        ctx.lineTo(w * 0.5, h * 0.8)
        ctx.lineTo(0, h * 0.8)
        ctx.closePath()
      } else ctx.rect(0, 0, w, h)
      if (s.fill) {
        ctx.fillStyle = `#${s.fill}`
        ctx.fill()
      }
      if (s.line) {
        ctx.strokeStyle = `#${s.line.color}`
        ctx.lineWidth = emuToPx(s.line.widthEmu)
        ctx.stroke()
      }
      const rows = s.paragraphs.filter((p) => p.runs.some((r) => r.text))
      let ty =
        h / 2 -
        rows.reduce((sum, p) => sum + Math.max(0, ...p.runs.map((r) => ((r.fontSizePt ?? 24) * 4) / 3)) * 1.2, 0) / 2
      for (const p of rows) {
        const fontSize = Math.max(0, ...p.runs.map((r) => ((r.fontSizePt ?? 24) * 4) / 3))
        const textWidth = p.runs.reduce((sum, r) => {
          ctx.font = font(r, s.fontFamily, 24)
          return sum + ctx.measureText(r.text).width
        }, 0)
        let tx = p.align === 'center' ? (w - textWidth) / 2 : p.align === 'right' ? w - textWidth : 0
        for (const r of p.runs) {
          ctx.font = font(r, s.fontFamily, 24)
          ctx.fillStyle = `#${r.color ?? s.textColor ?? '000000'}`
          ctx.textBaseline = 'alphabetic'
          ctx.fillText(r.text, tx, ty + fontSize * 0.9)
          tx += ctx.measureText(r.text).width
        }
        ty += fontSize * 1.2
      }
      ctx.restore()
    }
  } else if (drawing.kind === 'chart') paintChart(drawing, ctx, width, height)
  else {
    if (drawing.fill) {
      ctx.fillStyle = `#${drawing.fill}`
      ctx.fillRect(0, 0, width, height)
    }
    if (drawing.line) {
      ctx.strokeStyle = `#${drawing.line.color}`
      ctx.lineWidth = emuToPx(drawing.line.widthEmu)
      ctx.strokeRect(0, 0, width, height)
    }
    let x = width - emuToPx(drawing.insets.right),
      y = emuToPx(drawing.insets.top)
    for (const p of drawing.paragraphs) {
      const size = ((p.runs[0]?.fontSizePt ?? drawing.fontSizePt) * 4) / 3
      if (drawing.vertical) {
        x -= size
        y = emuToPx(drawing.insets.top)
      } else {
        x = emuToPx(drawing.insets.left)
      }
      for (const r of p.runs) {
        ctx.font = font(r, drawing.fontFamily, drawing.fontSizePt)
        ctx.fillStyle = `#${r.color ?? '000000'}`
        ctx.textBaseline = 'top'
        if (!drawing.vertical) {
          ctx.fillText(r.text, x, y)
          x += ctx.measureText(r.text).width
          continue
        }
        for (const token of r.text.match(/[\x20-\x7e]+|[^\x20-\x7e]/gu) ?? []) {
          if (/^[\x20-\x7e]+$/.test(token)) {
            ctx.save()
            ctx.translate(x + size, y)
            ctx.rotate(Math.PI / 2)
            ctx.fillText(token, 0, 0)
            ctx.restore()
            y += ctx.measureText(token).width
          } else {
            ctx.fillText(token, x, y)
            y += size
          }
        }
      }
      if (!drawing.vertical) y += size * 1.2
      else x -= size * 0.2
    }
  }
}

function font(run: DocxTextRun, family: string, size: number): string {
  return `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.fontSizePt ?? size}pt ${fontFamilyCss(run.fontFamily ?? family)}`
}
function paintChart(
  chart: Extract<DocxDrawing, { kind: 'chart' }>,
  ctx: CanvasRenderingContext2D,
  w: number,
  h: number
): void {
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, w, h)
  ctx.strokeStyle = '#D9D9D9'
  ctx.lineWidth = 1
  ctx.strokeRect(0, 0, w, h)
  const left = w * 0.046,
    right = w * 0.026,
    top = h * (chart.title?.length ? 0.152 : 0.09),
    bottom = h * (chart.legend ? 0.185 : 0.14)
  const pw = w - left - right,
    ph = h - top - bottom
  const values = chart.series.flatMap((s) => s.values).filter((v): v is number => v !== undefined)
  const min = chart.min ?? Math.min(0, ...values),
    maximum = Math.max(1, ...values)
  const raw = ((maximum - min) * 1.1) / 6,
    base = 10 ** Math.floor(Math.log10(raw)),
    step =
      chart.majorUnit && chart.majorUnit > 0
        ? chart.majorUnit
        : ([1, 2, 5, 10].map((v) => v * base).find((v) => v >= raw) ?? base * 10)
  const max = chart.max ?? Math.ceil(((maximum - min) * 1.1) / step) * step + min,
    range = Math.max(1, max - min)
  ctx.font = font({ text: '' }, chart.fontFamily, chart.fontSizePt)
  ctx.fillStyle = '#595959'
  ctx.textAlign = 'right'
  ctx.textBaseline = 'middle'
  for (let value = min, count = 0; value <= max + step / 1000 && count < 100; value += step, count++) {
    const y = top + ph - ((value - min) / range) * ph
    ctx.strokeStyle = '#D9D9D9'
    ctx.beginPath()
    ctx.moveTo(left, y)
    ctx.lineTo(w - right, y)
    ctx.stroke()
    ctx.fillText(String(Number(value.toPrecision(8))), left - 8, y)
  }
  const categories = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1),
    group = pw / categories
  const stride = 1 - chart.overlap / 100,
    span = 1 + (Math.max(1, chart.series.length) - 1) * stride
  const bar = group / (span + chart.gapWidth / 100),
    zero = top + ph - ((0 - min) / range) * ph
  chart.series.forEach((s, si) =>
    s.values.forEach((value, ci) => {
      if (value === undefined) return
      const x = left + group * ci + (group - bar * span) / 2 + bar * si * stride,
        y = top + ph - ((value - min) / range) * ph
      ctx.fillStyle = `#${s.color}`
      ctx.fillRect(x, Math.min(y, zero), bar, Math.abs(zero - y))
    })
  )
  ctx.fillStyle = '#595959'
  ctx.textAlign = 'center'
  ctx.textBaseline = 'top'
  chart.categories.forEach((label, i) => ctx.fillText(label, left + group * (i + 0.5), top + ph + 10))
  if (chart.title) {
    ctx.font = font({ text: '' }, chart.fontFamily, 14)
    ctx.fillText(chart.title, w / 2, 14)
  }
  if (chart.legend) {
    ctx.font = font({ text: '' }, chart.fontFamily, chart.fontSizePt)
    const sizes = chart.series.map((s) => ctx.measureText(s.name).width + 17),
      total = sizes.reduce((a, b) => a + b, 0)
    let x = (w - total) / 2
    const y = h - 24
    chart.series.forEach((s, i) => {
      ctx.fillStyle = `#${s.color}`
      ctx.fillRect(x, y + 3, 6, 6)
      ctx.fillStyle = '#595959'
      ctx.textAlign = 'left'
      ctx.fillText(s.name, x + 9, y)
      x += sizes[i]
    })
  }
}
