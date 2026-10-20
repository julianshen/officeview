/**
 * Chart title runs (spec 5.10; lazy chunk). The title's content is chosen by the parser (`title/tx/rich`, else the
 * `title/tx/strRef` cache, else the automatic title, unless `autoTitleDeleted`); this module resolves the runs'
 * formatting, attribute by attribute (size, bold, italic, colour, typeface, underline), in Office's order:
 *
 *   run `a:rPr` → the paragraph's `a:pPr/a:defRPr` → the title's `c:txPr` → the chart style part (`cs:title`)
 *   → the theme (left unset here: the painter's defaults, the theme minor font at 1.4 x the chart size in #595959).
 *
 * Runs are kept only when they carry formatting or there is more than one (paragraphs, several runs): a plain
 * single-run title keeps the engine's existing drawing and costs nothing. What is kept is charged before it is built.
 * Bounds: CHART_TITLE_RUN_LIMIT records (text runs and line breaks alike), CHART_TITLE_LINE_LIMIT lines and
 * CHART_TITLE_TEXT_LIMIT text units; whatever they cut is reported (`truncated`), never dropped silently.
 */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { parseDrawingColor, resolveDrawingColor, type ThemeContext } from '../drawing/style'
import type { ChartTextRun } from './model'
import type { ChartStyleText } from './style'
import { hexColor } from './fills'

/** Records (runs and line breaks), lines and text units kept per title. */
export const CHART_TITLE_RUN_LIMIT = 64
export const CHART_TITLE_LINE_LIMIT = 16
export const CHART_TITLE_TEXT_LIMIT = 1024
/** The diagnostic of a title cut by these bounds or by the placement (spec 6.6). */
export const CHART_TITLE_TRUNCATED = 'chart-title-truncated'
const RUN_BYTES = 128

type Props = Omit<ChartTextRun, 'text'>
const KEYS = ['sizePt', 'bold', 'italic', 'color', 'font', 'underline'] as const
const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const flag = (raw: string | undefined) => raw === undefined ? undefined : raw === '1' || raw === 'true'

/** A DrawingML typeface: theme references (`+mn-lt`, `+mj-lt`, and the ea/cs forms) resolve to the theme's latin font. */
export function typeface(raw: string | undefined, theme: ThemeContext | undefined): string | undefined {
  if (!raw) return undefined
  if (/^\+mn-/.test(raw)) return theme?.fonts.minor.latin || undefined
  if (/^\+mj-/.test(raw)) return theme?.fonts.major.latin || undefined
  return raw.slice(0, 64)
}
/** The formatting a DrawingML character-properties element states (`a:rPr`, `a:defRPr`). */
function propsOf(node: XmlNode | undefined, theme: ThemeContext | undefined): Props {
  if (!node) return {}
  const a = attrs(node), size = Number(a.sz), fill = child(node, 'solidFill')
  const color = fill ? hexColor(resolveDrawingColor(parseDrawingColor(fill), theme)) : undefined
  const font = typeface(attrs(child(node, 'latin')).typeface, theme)
  return {
    ...(a.sz !== undefined && size > 0 ? { sizePt: Math.max(6, Math.min(72, size / 100)) } : {}),
    ...(flag(a.b) !== undefined ? { bold: flag(a.b) } : {}),
    ...(flag(a.i) !== undefined ? { italic: flag(a.i) } : {}),
    ...(color ? { color: `#${color}` } : {}),
    ...(font ? { font } : {}),
    ...(a.u !== undefined ? { underline: a.u !== 'none' } : {}),
  }
}
function styleProps(style: ChartStyleText | undefined, theme: ThemeContext | undefined): Props {
  if (!style) return {}
  const color = style.color ? hexColor(resolveDrawingColor(style.color, theme)) : undefined
  const font = style.font === 'major' ? theme?.fonts.major.latin || undefined : style.font === 'minor' ? undefined : typeface(style.font, theme)
  return {
    ...(style.sizePt !== undefined ? { sizePt: style.sizePt } : {}),
    ...(style.bold !== undefined ? { bold: style.bold } : {}),
    ...(style.italic !== undefined ? { italic: style.italic } : {}),
    ...(color ? { color: `#${color}` } : {}),
    ...(font ? { font } : {}),
    ...(style.underline !== undefined ? { underline: style.underline } : {}),
  }
}
/** Each attribute from the first level that states it. */
function resolve(levels: Props[]): Props {
  const out: Props = {}
  for (const key of KEYS) {
    const level = levels.find(l => l[key] !== undefined)
    if (level) (out as Record<string, unknown>)[key] = level[key]
  }
  return out
}
const defRPr = (txBody: XmlNode | undefined) => child(child(child(txBody, 'p'), 'pPr'), 'defRPr')
const formatted = (run: ChartTextRun) => Object.keys(run).length > 1

export interface TitleRuns { runs: ChartTextRun[]; truncated: boolean }
/**
 * The title's runs, formatting resolved, or undefined for a plain single run (drawn as before). `text` is the
 * title the parser chose (used when the title is not rich text).
 */
export function titleRuns(title: XmlNode | undefined, text: string | undefined, theme: ThemeContext | undefined, style: ChartStyleText | undefined, charge: (bytes: number) => void): TitleRuns | undefined {
  if (!title || !text) return undefined
  const txPr = propsOf(defRPr(child(title, 'txPr')), theme), fromStyle = styleProps(style, theme)
  const blank = (p: XmlNode) => !orderedChildren(p).some(([kind, r]) => (kind === 'r' || kind === 'fld') && getChildren(r, 't').map(textOf).join('').trim())
  // Rich text with no text gave way to the cache or the automatic title (the parser's order): one run of `text`.
  const authored = child(child(title, 'tx'), 'rich')
  const rich = authored && getChildren(authored, 'p').some(p => !blank(p)) ? authored : undefined
  // The walk runs twice: first to size and decide (nothing retained is built), then, once charged, to build.
  // Every record (a text run, a paragraph separator, a line break) counts against the run bound.
  const walk = (emit: (run: ChartTextRun | '\n') => void): boolean => {
    if (!rich) { emit({ text: text.slice(0, CHART_TITLE_TEXT_LIMIT), ...resolve([txPr, fromStyle]) }); return text.length > CHART_TITLE_TEXT_LIMIT }
    const paragraphs = getChildren(rich, 'p')
    // Leading and trailing empty paragraphs are not lines (the joined title is trimmed the same way).
    let first = 0, last = paragraphs.length - 1
    while (first <= last && blank(paragraphs[first])) first++
    while (last >= first && blank(paragraphs[last])) last--
    let left = CHART_TITLE_TEXT_LIMIT, records = 0, lines = 1
    const room = () => records < CHART_TITLE_RUN_LIMIT
    const lineBreak = () => { if (!room() || lines >= CHART_TITLE_LINE_LIMIT) return false; emit('\n'); records++; lines++; return true }
    for (let i = first; i <= last; i++) {
      if (i > first && !lineBreak()) return true
      const p = paragraphs[i], paragraph = propsOf(child(child(p, 'pPr'), 'defRPr'), theme)
      for (const [kind, r] of orderedChildren(p)) {
        if (kind !== 'r' && kind !== 'fld' && kind !== 'br') continue
        if (kind === 'br') { if (!lineBreak()) return true; continue }
        const full = getChildren(r, 't').map(textOf).join('')
        if (!full) continue
        if (!room() || left <= 0) return true
        const value = full.slice(0, left)
        emit({ text: value, ...resolve([propsOf(child(r, 'rPr'), theme), paragraph, txPr, fromStyle]) })
        left -= value.length; records++
        if (value.length < full.length) return true
      }
    }
    return false
  }
  let runs = 0, bytes = 0, styled = false
  const truncated = walk(run => { runs++; bytes += RUN_BYTES + (run === '\n' ? 2 : run.text.length * 2 + (run.font?.length ?? 0) * 2); if (run !== '\n' && formatted(run)) styled = true })
  if (!runs || (runs === 1 && !styled && !truncated)) return undefined
  charge(bytes)
  const out: ChartTextRun[] = []
  walk(run => { out.push(run === '\n' ? { text: '\n' } : run) })
  return { runs: out, truncated }
}

/** The title's lines: runs split at `\n` (a paragraph or line break). */
export function titleLines(runs: readonly ChartTextRun[]): ChartTextRun[][] {
  const lines: ChartTextRun[][] = [[]]
  for (const run of runs) {
    if (run.text === '\n') lines.push([])
    else lines[lines.length - 1].push(run)
  }
  return lines
}
/** The text a title's runs show (the painted and accessible title when runs were cut). */
export const runsText = (runs: readonly ChartTextRun[]) => runs.map(r => r.text).join('')
