import { attrs, elementChildren, getChildren, textOf, type XmlNode } from '../core/xml'
import type { DocxDrawing } from './types'

function descendants(node: XmlNode | undefined, tag: string): XmlNode[] {
  return elementChildren(node).flatMap(([name, child]) => [
    ...(name === tag ? [child] : []),
    ...descendants(child, tag)
  ])
}

const LENGTH_EMU: Record<string, number> = { cm: 360000, mm: 36000, in: 914400, pt: 12700, px: 9525 }

/** Cartesian InkML traces. Unsupported encodings keep the compatibility picture. */
export function parseInk(root: XmlNode): Extract<DocxDrawing, { kind: 'ink' }> | undefined {
  const contexts = descendants(root, 'context')
  const brushes = descendants(root, 'brush')
  const strokes: Extract<DocxDrawing, { kind: 'ink' }>['strokes'] = []
  for (const trace of descendants(root, 'trace')) {
    if (attrs(trace).type === 'penUp') continue
    const context = contexts.find((node) => `#${attrs(node).id}` === attrs(trace).contextRef)
    const source = getChildren(context, 'inkSource')[0]
    const format = getChildren(source, 'traceFormat')[0] ?? getChildren(root, 'traceFormat')[0]
    const channels = getChildren(format, 'channel')
    if (channels.length < 2 || attrs(channels[0]).name !== 'X' || attrs(channels[1]).name !== 'Y') return undefined
    const scale = channels.slice(0, 2).map((channel) => {
      const a = attrs(channel)
      const resolution = descendants(source, 'channelProperty').find(
        (property) => attrs(property).channel === a.name && attrs(property).name === 'resolution'
      )
      const divisor = Number(attrs(resolution).value ?? 1)
      return (LENGTH_EMU[a.units ?? 'mm'] ?? 0) / divisor
    })
    if (scale.some((value) => !Number.isFinite(value) || value <= 0)) return undefined
    const values = channels.map(() => 0)
    const differences = channels.map(() => 0)
    const modes = channels.map(() => '!')
    const points: Array<[number, number]> = []
    // InkML prefixes persist per channel: ! absolute, ' first difference,
    // " second difference. See https://www.w3.org/TR/InkML/#trace.
    for (const sample of textOf(trace).split(',')) {
      const tokens = [...sample.matchAll(/([!'\"]?)([+-]?(?:\d*\.\d+|\d+)(?:[eE][+-]?\d+)?)/g)]
      if (tokens.length !== channels.length) return undefined
      for (let index = 0; index < tokens.length; index++) {
        const [, prefix, raw] = tokens[index]
        if (prefix) modes[index] = prefix
        const value = Number(raw)
        if (modes[index] === '!') {
          differences[index] = value - values[index]
          values[index] = value
        } else if (modes[index] === "'") {
          differences[index] = value
          values[index] += value
        } else {
          differences[index] += value
          values[index] += differences[index]
        }
      }
      if (values.some(value => !Number.isFinite(value))) return undefined
      points.push([values[0] * scale[0], values[1] * scale[1]])
    }
    const brush = brushes.find((node) => `#${attrs(node).id}` === attrs(trace).brushRef)
    const properties = getChildren(brush, 'brushProperty')
    const width = attrs(properties.find((node) => attrs(node).name === 'width'))
    const widthEmu = Number(width.value ?? 1) * (LENGTH_EMU[width.units ?? 'mm'] ?? 0)
    if (!Number.isFinite(widthEmu) || widthEmu <= 0) return undefined
    const color = attrs(properties.find((node) => attrs(node).name === 'color')).value?.replace(/^#/, '')
    strokes.push({ points, widthEmu, color: color && /^[\da-f]{6}$/i.test(color) ? color : '000000' })
  }
  if (!strokes.length) return undefined
  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity
  for (const stroke of strokes)
    for (const [x, y] of stroke.points) {
      minX = Math.min(minX, x)
      minY = Math.min(minY, y)
      maxX = Math.max(maxX, x)
      maxY = Math.max(maxY, y)
    }
  for (const stroke of strokes) stroke.points = stroke.points.map(([x, y]) => [x - minX, y - minY])
  return { kind: 'ink', strokes, widthEmu: maxX - minX, heightEmu: maxY - minY }
}

export function paintInk(
  ink: Extract<DocxDrawing, { kind: 'ink' }>,
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number
): void {
  const sx = ink.widthEmu > 0 ? width / ink.widthEmu : 1 / 9525
  const sy = ink.heightEmu > 0 ? height / ink.heightEmu : 1 / 9525
  ctx.lineCap = 'round'
  ctx.lineJoin = 'round'
  for (const stroke of ink.strokes) {
    ctx.strokeStyle = ctx.fillStyle = `#${stroke.color}`
    ctx.lineWidth = (stroke.widthEmu * (sx + sy)) / 2
    ctx.beginPath()
    stroke.points.forEach(([x, y], index) => {
      if (index === 0) ctx.moveTo(x * sx, y * sy)
      else ctx.lineTo(x * sx, y * sy)
    })
    if (stroke.points.length === 1) {
      const [x, y] = stroke.points[0]
      ctx.arc(x * sx, y * sy, ctx.lineWidth / 2, 0, Math.PI * 2)
      ctx.fill()
    } else ctx.stroke()
  }
}
