/**
 * Chart title runs (spec 5.10; lazy chunk). The title's content is chosen by the parser (`title/tx/rich`, else the
 * `title/tx/strRef` cache, else the automatic title, unless `autoTitleDeleted`); this module resolves the runs'
 * formatting, attribute by attribute, in Office's order:
 *
 *   run `a:rPr` → the paragraph's `a:pPr/a:defRPr` → the title's `c:txPr` → the chart style part (`cs:title`)
 *   → the theme (left unset here: the painter's defaults, the theme minor font at 1.4 x the chart size in #595959).
 *
 * Runs are kept only when they carry formatting or there is more than one (paragraphs, several runs): a plain
 * single-run title keeps the engine's existing drawing and costs nothing. What is kept is charged before it is built.
 */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { parseDrawingColor, resolveDrawingColor, type ThemeContext } from '../drawing/style'
import type { ChartTextRun } from './model'
import type { ChartStyleText } from './style'
import { hexColor } from './fills'

/** Runs and text units kept per title. */
export const CHART_TITLE_RUN_LIMIT = 64
export const CHART_TITLE_TEXT_LIMIT = 1024
const RUN_BYTES = 128

type Props = Omit<ChartTextRun, 'text'>
const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const flag = (raw: string | undefined) => raw === undefined ? undefined : raw === '1' || raw === 'true'

/** The formatting a DrawingML character-properties element states (`a:rPr`, `a:defRPr`). */
function propsOf(node: XmlNode | undefined, theme: ThemeContext | undefined): Props {
  if (!node) return {}
  const a = attrs(node), size = Number(a.sz), fill = child(node, 'solidFill')
  const color = fill ? hexColor(resolveDrawingColor(parseDrawingColor(fill), theme)) : undefined
  return {
    ...(a.sz !== undefined && size > 0 ? { sizePt: Math.max(6, Math.min(72, size / 100)) } : {}),
    ...(flag(a.b) !== undefined ? { bold: flag(a.b) } : {}),
    ...(flag(a.i) !== undefined ? { italic: flag(a.i) } : {}),
    ...(color ? { color: `#${color}` } : {}),
  }
}
function styleProps(style: ChartStyleText | undefined, theme: ThemeContext | undefined): Props {
  if (!style) return {}
  const color = style.color ? hexColor(resolveDrawingColor(style.color, theme)) : undefined
  return {
    ...(style.sizePt !== undefined ? { sizePt: style.sizePt } : {}),
    ...(style.bold !== undefined ? { bold: style.bold } : {}),
    ...(style.italic !== undefined ? { italic: style.italic } : {}),
    ...(color ? { color: `#${color}` } : {}),
  }
}
/** Each attribute from the first level that states it. */
function resolve(levels: Props[]): Props {
  const out: Props = {}
  for (const key of ['sizePt', 'bold', 'italic', 'color'] as const) {
    const level = levels.find(l => l[key] !== undefined)
    if (level) (out as Record<string, unknown>)[key] = level[key]
  }
  return out
}
const defRPr = (txBody: XmlNode | undefined) => child(child(child(txBody, 'p'), 'pPr'), 'defRPr')
const formatted = (run: ChartTextRun) => Object.keys(run).length > 1

/**
 * The title's runs, formatting resolved, or undefined for a plain single run (drawn as before). `text` is the
 * title the parser chose (used when the title is not rich text).
 */
export function titleRuns(title: XmlNode | undefined, text: string | undefined, theme: ThemeContext | undefined, style: ChartStyleText | undefined, charge: (bytes: number) => void): ChartTextRun[] | undefined {
  if (!title || !text) return undefined
  const txPr = propsOf(defRPr(child(title, 'txPr')), theme), fromStyle = styleProps(style, theme)
  const blank = (p: XmlNode) => !orderedChildren(p).some(([kind, r]) => (kind === 'r' || kind === 'fld') && getChildren(r, 't').map(textOf).join('').trim())
  // Rich text with no text gave way to the cache or the automatic title (the parser's order): one run of `text`.
  const authored = child(child(title, 'tx'), 'rich')
  const rich = authored && getChildren(authored, 'p').some(p => !blank(p)) ? authored : undefined
  // The walk runs twice: first to size and decide (nothing retained is built), then, once charged, to build.
  const walk = (emit: (run: ChartTextRun | '\n') => void) => {
    if (!rich) { emit({ text: text.slice(0, CHART_TITLE_TEXT_LIMIT), ...resolve([txPr, fromStyle]) }); return }
    const paragraphs = getChildren(rich, 'p')
    // Leading and trailing empty paragraphs are not lines (the joined title is trimmed the same way).
    let first = 0, last = paragraphs.length - 1
    while (first <= last && blank(paragraphs[first])) first++
    while (last >= first && blank(paragraphs[last])) last--
    let left = CHART_TITLE_TEXT_LIMIT, count = 0
    for (let i = first; i <= last; i++) {
      if (i > first) emit('\n')
      const p = paragraphs[i], paragraph = propsOf(child(child(p, 'pPr'), 'defRPr'), theme)
      for (const [kind, r] of orderedChildren(p)) {
        if (kind !== 'r' && kind !== 'fld' && kind !== 'br') continue
        if (count >= CHART_TITLE_RUN_LIMIT || left <= 0) return
        if (kind === 'br') { emit('\n'); continue }
        const value = getChildren(r, 't').map(textOf).join('').slice(0, left)
        if (!value) continue
        emit({ text: value, ...resolve([propsOf(child(r, 'rPr'), theme), paragraph, txPr, fromStyle]) })
        left -= value.length; count++
      }
    }
  }
  let runs = 0, bytes = 0, styled = false
  walk(run => { runs++; bytes += RUN_BYTES + (run === '\n' ? 2 : run.text.length * 2); if (run !== '\n' && formatted(run)) styled = true })
  if (!runs || (runs === 1 && !styled)) return undefined
  charge(bytes)
  const out: ChartTextRun[] = []
  walk(run => { out.push(run === '\n' ? { text: '\n' } : run) })
  return out
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
