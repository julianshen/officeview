/**
 * The captions rule (plan Stage 3 scope rulings, "Captions", closing the Stage 2 caption ruling): a layout cell with
 * `caption: 'implicit'` (Excel's default caption, which Excel localises) is filled from the pivot's own file, never
 * from an assumed English table. Lazy chunk `pivots`.
 *
 * 1. Measured tables: `MEASURED_CAPTION_TABLES` holds one pattern table per measured Excel locale. A pivot whose saved
 *    implicit captions ALL render exactly through one table (at least one caption compared) uses that table, locale
 *    'zh-TW'.
 * 2. Otherwise the pivot uses only the patterns its own saved cells show, kind by kind (locale 'file'): each saved
 *    implicit caption gives its kind's pattern by locating its arguments inside the saved text ("Sum of Sales 的加總"
 *    for "Total <data>" with data "Sum of Sales" gives "{data} 的加總"). Two samples of one kind that disagree, or a
 *    sample whose arguments do not appear in its text, leave that kind without a pattern.
 * 3. A layout that needs a kind with no pattern refuses with `pivot-caption-locale` (`resolveCaptions`).
 *
 * A pattern is keyed by kind, plus the function for the kinds that name one (`dataField:sum`, `subtotalFunction:avg`):
 * Excel localises the function name, so it is part of the pattern, not an argument.
 */
import type { PivotDiagnostic, PivotTableDefinition } from '../model'
import type { EngineResult } from '../engine/axes'
import type { CellReader } from '../engine/faithful'
import type { CaptionArg, CaptionArgs, CaptionKind, CaptionRef, PivotLayout, PivotLayoutCell } from '../engine/place'

export type { CaptionArg, CaptionArgs, CaptionKind, CaptionRef, CaptionSlot } from '../engine/place'
export interface CaptionRule { readonly locale: 'zh-TW' | 'file'; caption(kind: CaptionKind, args: CaptionArgs): string | undefined }
/** A measured pattern table: pattern key → template, where `{item}`, `{data}` and `{field}` stand for the arguments. */
export type CaptionTable = Readonly<Partial<Record<string, string>>>

/**
 * The measured zh-TW caption patterns (Excel 16.114 zh-TW). TASK 24 PLUGS ITS MEASURED TABLE IN HERE, one entry per
 * pattern key, each cited from docs/validation/xlsx-pivots.md ("Stage 3 interaction oracle and measurements") and
 * pinned by a test. Until then it is empty, so no pivot matches it and every pivot uses its own saved patterns.
 */
export const ZH_TW_CAPTIONS: CaptionTable = Object.freeze({})
export const MEASURED_CAPTION_TABLES: Readonly<Record<'zh-TW', CaptionTable>> = Object.freeze({ 'zh-TW': ZH_TW_CAPTIONS })

const FN_KINDS = new Set<CaptionKind>(['dataField', 'subtotalFunction', 'subtotalFunctionData'])
/** The pattern key of a caption: its kind, with the function for the kinds that name one. */
export const captionKey = (kind: CaptionKind, args: CaptionArgs) => FN_KINDS.has(kind) && args.fn !== undefined ? `${kind}:${args.fn}` : kind
/** The substituted arguments, in the order a pattern is derived. */
const ARG_NAMES = ['item', 'data', 'field'] as const
type ArgName = typeof ARG_NAMES[number]
type Segment = { text: string } | { arg: ArgName }
type Template = readonly Segment[]

/** `{item} 合計` → segments. Only the three argument names are placeholders. */
function parseTemplate(source: string): Template {
  const out: Segment[] = []
  const re = /\{(item|data|field)\}/g
  let at = 0
  for (let m = re.exec(source); m; m = re.exec(source)) {
    if (m.index > at) out.push({ text: source.slice(at, m.index) })
    out.push({ arg: m[1] as ArgName })
    at = re.lastIndex
  }
  if (at < source.length) out.push({ text: source.slice(at) })
  return out
}
const sameTemplate = (a: Template, b: Template) => a.length === b.length && a.every((s, k) => {
  const t = b[k]!
  return 'text' in s ? 'text' in t && t.text === s.text : 'arg' in t && t.arg === s.arg
})

/** A rule over pattern templates by key. Nested implicit arguments render through the same rule. */
function ruleOf(locale: CaptionRule['locale'], templates: ReadonlyMap<string, Template>): CaptionRule {
  const rule: CaptionRule = {
    locale,
    caption(kind, args) {
      const template = templates.get(captionKey(kind, args))
      if (!template) return undefined
      let out = ''
      for (const segment of template) {
        if ('text' in segment) { out += segment.text; continue }
        const value = argText(args[segment.arg], rule)
        if (value === undefined) return undefined
        out += value
      }
      return out
    },
  }
  return rule
}
function argText(arg: CaptionArg | undefined, rule: CaptionRule): string | undefined {
  if (arg === undefined) return undefined
  return typeof arg === 'string' ? arg : rule.caption(arg.kind, arg.args)
}

/** The template a saved text shows for `args` (resolved to text), or undefined when an argument is not in it. */
function derive(text: string, args: Partial<Record<ArgName, string>>): Template | undefined {
  let segments: Segment[] = [{ text }]
  for (const name of ARG_NAMES) {
    const value = args[name]
    if (value === undefined) continue
    if (value === '') return undefined
    const at = segments.findIndex(segment => 'text' in segment && segment.text.includes(value))
    if (at < 0) return undefined
    const literal = (segments[at] as { text: string }).text, index = literal.indexOf(value)
    const parts: Segment[] = [
      ...(index > 0 ? [{ text: literal.slice(0, index) }] : []),
      { arg: name },
      ...(index + value.length < literal.length ? [{ text: literal.slice(index + value.length) }] : []),
    ]
    segments = [...segments.slice(0, at), ...parts, ...segments.slice(at + 1)]
  }
  return segments
}

interface Sample { ref: CaptionRef; text: string }
/** Patterns from the pivot's own saved captions. Nested arguments resolve through the patterns found so far, so
 * samples are taken in passes until none makes progress; a sample left unresolved voids its kind. */
function filePatterns(samples: readonly Sample[]): Map<string, Template> {
  const found = new Map<string, Template>()
  const void_ = new Set<string>()
  let pending = [...samples]
  for (let progress = true; progress && pending.length;) {
    progress = false
    const rule = ruleOf('file', found)
    const next: Sample[] = []
    for (const sample of pending) {
      const args: Partial<Record<ArgName, string>> = {}
      let unresolved = false
      for (const name of ARG_NAMES) {
        const arg = sample.ref.args[name]
        if (arg === undefined) continue
        const value = argText(arg, rule)
        if (value === undefined) { unresolved = true; break }
        args[name] = value
      }
      if (unresolved) { next.push(sample); continue }
      progress = true
      const key = captionKey(sample.ref.kind, sample.ref.args)
      const template = derive(sample.text, args)
      const known = found.get(key)
      if (!template || (known && !sameTemplate(known, template))) void_.add(key)
      else found.set(key, template)
    }
    pending = next
  }
  for (const sample of pending) void_.add(captionKey(sample.ref.kind, sample.ref.args))
  for (const key of void_) found.delete(key)
  return found
}

/**
 * The caption rule of a pivot. `layout` is its SAVED state computed without saved captions (`computePivot` on
 * `initialViewState`), so its `captionSlots` name every implicit caption; `saved` reads the sheet's saved text there.
 */
export function captionRule(_def: PivotTableDefinition, saved: CellReader, layout: PivotLayout,
  tables: Readonly<Partial<Record<'zh-TW', CaptionTable>>> = MEASURED_CAPTION_TABLES): CaptionRule {
  const samples = (layout.captionSlots ?? []).flatMap((slot): Sample[] => {
    const cell = saved(slot.row, slot.col)
    return cell?.kind === 'text' ? [{ ref: slot, text: cell.value }] : []
  })
  const own = filePatterns(samples)
  for (const [locale, table] of Object.entries(tables) as Array<['zh-TW', CaptionTable]>) {
    const templates = new Map(Object.entries(table).flatMap(([key, source]) => source === undefined ? [] : [[key, parseTemplate(source)] as const]))
    const measured = ruleOf(locale, templates)
    if (samples.length && samples.every(sample => measured.caption(sample.ref.kind, sample.ref.args) === sample.text)) {
      // The measured table wins; a kind it does not list falls back to the pivot's own pattern.
      for (const [key, template] of own) if (!templates.has(key)) templates.set(key, template)
      return ruleOf(locale, templates)
    }
  }
  return ruleOf('file', own)
}

const refuse = (message: string): EngineResult<never> => ({ ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-caption-locale', message } satisfies PivotDiagnostic })
const at = (row: number, col: number) => `R${row + 1}C${col + 1}`
const resolved = (layout: PivotLayout, texts: ReadonlyMap<string, string>): PivotLayout => ({
  ...layout,
  cells: layout.cells.map((cell): PivotLayoutCell => {
    const value = cell.caption === 'implicit' ? texts.get(`${cell.row},${cell.col}`) : undefined
    return value === undefined ? cell : { row: cell.row, col: cell.col, value: { kind: 'text', value } }
  }),
  captionsUnverified: layout.cells.some(cell => cell.caption === 'implicit' && !texts.has(`${cell.row},${cell.col}`)),
  captionSlots: (layout.captionSlots ?? []).filter(slot => !texts.has(`${slot.row},${slot.col}`)),
})

/** Every implicit caption of `layout` filled through `rule`; a kind the rule has no pattern for refuses. */
export function resolveCaptions(layout: PivotLayout, rule: CaptionRule): EngineResult<PivotLayout> {
  const slots = layout.captionSlots ?? []
  const texts = new Map<string, string>()
  for (const slot of slots) {
    const text = rule.caption(slot.kind, slot.args)
    if (text === undefined) {
      const source = rule.locale === 'file' ? 'the pivot\'s saved cells show' : `the measured ${rule.locale} table holds`
      return refuse(`no caption pattern for ${captionKey(slot.kind, slot.args)} at ${at(slot.row, slot.col)}: none that ${source}`)
    }
    texts.set(`${slot.row},${slot.col}`, text)
  }
  const orphan = layout.cells.find(cell => cell.caption === 'implicit' && !texts.has(`${cell.row},${cell.col}`))
  if (orphan) return refuse(`the default caption at ${at(orphan.row, orphan.col)} has no caption kind`)
  return { ok: true, value: resolved(layout, texts) }
}

/** The saved state's implicit captions filled with the saved text itself (the saved state shows what was saved). */
export function savedCaptions(layout: PivotLayout, saved: CellReader): PivotLayout {
  const texts = new Map<string, string>()
  for (const slot of layout.captionSlots ?? []) {
    const cell = saved(slot.row, slot.col)
    if (cell?.kind === 'text') texts.set(`${slot.row},${slot.col}`, cell.value)
  }
  return resolved(layout, texts)
}

