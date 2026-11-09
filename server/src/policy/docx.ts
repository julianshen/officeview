import { all, attr, local, type Context, type Node } from './tree'

const visibility = new Set(['vanish', 'specVanish', 'webHidden'])
const child = (node: Node | undefined, name: string) => node?.children.find(n => local(n.name) === name)
type Flag = { node: Node; hidden: boolean; unknown: boolean }
/** Resolve the run-property cascade before removing anything from the styles part. */
export function projectDocx(ctx: Context): void {
  const stylesPart = ctx.parts.get('word/styles.xml')
  const styles = stylesPart && all(stylesPart, 'styles')[0]
  const byId = new Map((styles ? all(styles, 'style') : []).map(s => [attr(s, 'styleId'), s]))
  const defaults = styles && child(child(child(styles, 'docDefaults'), 'rPrDefault'), 'rPr')
  const defaultParagraph = styles && all(styles, 'style').find(s => attr(s, 'type') === 'paragraph' && ['1', 'true', 'on'].includes(attr(s, 'default') ?? ''))
  function properties(props: Node | undefined, flags: Map<string, Flag>) {
    for (const flag of props?.children ?? []) {
      if (!visibility.has(local(flag.name))) continue
      const value = attr(flag, 'val') ?? '1'
      const hidden = !['0', 'false', 'off'].includes(value)
      const unknown = !['0', 'false', 'off', '1', 'true', 'on'].includes(value)
      flags.set(local(flag.name), { node: flag, hidden, unknown })
      if (!unknown) ctx.handled.add(flag)
    }
  }
  function style(id: string | undefined, flags: Map<string, Flag>, seen = new Set<string>()) {
    if (!id || seen.has(id)) return
    seen.add(id)
    const entry = byId.get(id)
    if (!entry) return
    style(attr(child(entry, 'basedOn') ?? entry, 'val'), flags, seen)
    properties(child(entry, 'rPr'), flags)
  }
  // Known style declarations are consumed by the cascade, not independent drawables.
  if (styles) for (const props of all(styles, 'rPr')) properties(props, new Map())
  // Paragraph-mark properties (w:pPr/w:rPr) style only the pilcrow, which
  // carries no glyphs: a hidden mark reveals nothing, so its known flags are
  // recognised (Word would merge the paragraph with the next; we keep it).
  for (const [part, root] of ctx.parts) if (part.startsWith('word/')) for (const pPr of all(root, 'pPr')) properties(child(pPr, 'rPr'), new Map())
  for (const [part, root] of [...ctx.parts]) {
    if (!part.startsWith('word/')) continue
    if (/\/(?:comments[^/]*|people)\.xml$/.test(part)) { ctx.dropPart(part); continue }
    function visit(node: Node, ancestors: Node[]) {
      node.children = node.children.filter(current => {
        const name = local(current.name)
        if (['commentRangeStart', 'commentRangeEnd', 'commentReference'].includes(name)) {
          ctx.diagnostics.push({ code: 'hidden-stripped', part, element: current.name }); return false
        }
        const prefix = current.name.includes(':') ? current.name.split(':')[0] : ''
        const namespace = [...ancestors, node, current].reverse().map(n => n.attrs[prefix ? `xmlns:${prefix}` : 'xmlns']).find(Boolean)
        const wordRun = name === 'r' && (namespace ? /\/wordprocessingml\/(?:2006\/)?main$/.test(namespace) : prefix === 'w' || prefix === '')
        if (wordRun) {
          const flags = new Map<string, Flag>()
          properties(defaults, flags)
          const para = [...ancestors, node].reverse().find(a => local(a.name) === 'p')
          const paraId = attr(child(child(para, 'pPr'), 'pStyle') ?? {name:'',attrs:{},children:[]}, 'val')
          style(paraId ?? (defaultParagraph && attr(defaultParagraph, 'styleId')), flags)
          const rPr = child(current, 'rPr')
          style(attr(child(rPr, 'rStyle') ?? current, 'val'), flags)
          properties(rPr, flags)
          const hidden = [...flags.values()].find(f => f.hidden)
          if (hidden) {
            ctx.diagnostics.push({code: hidden.unknown ? 'hidden-unrecognised' : 'hidden-stripped', part, element: hidden.node.name, attr: 'val', value: attr(hidden.node, 'val') ?? '1'})
            return false
          }
          const drawing = all(current, 'docPr').find(p => ['1','true','on'].includes(attr(p,'hidden') ?? ''))
          if (drawing) { ctx.diagnostics.push({code:'hidden-stripped',part,element:drawing.name,attr:'hidden',value:attr(drawing,'hidden')}); return false }
        }
        if (['wsp','wgp','grpSp','sp','pic','drawing'].includes(name)) {
          // Only this shape's nonvisual properties: a child shape must not hide its siblings.
          const ownFlags = (n: Node): Node[] => n.children.flatMap(c => ['wsp','wgp','grpSp','sp','pic'].includes(local(c.name)) ? [] : ['cNvPr','docPr'].includes(local(c.name)) ? [c] : ownFlags(c))
          const hidden = ownFlags(current).find(p => ['1','true','on'].includes(attr(p,'hidden') ?? ''))
          if (hidden) { ctx.diagnostics.push({code:'hidden-stripped',part,element:hidden.name,attr:'hidden',value:attr(hidden,'hidden')}); return false }
        }
        visit(current, [...ancestors, node]); return true
      })
    }
    visit(root, [])
  }
}
