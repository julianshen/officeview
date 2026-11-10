import { all, attr, local, remove, type Context, type Node } from './tree'

const shapes = new Set(['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'])
const isTrue = (value: string | undefined) => ['1', 'true', 'on'].includes(value ?? '')
const isFalse = (value: string | undefined) => ['0', 'false', 'off'].includes(value ?? '')
function targetPath(target: string): string {
  const segments = (target.startsWith('/') ? target.slice(1) : `ppt/${target}`).split('/')
  const result: string[] = []
  for (const segment of segments) { if (segment === '..') result.pop(); else if (segment && segment !== '.') result.push(segment) }
  return result.join('/')
}
export function projectPptx(ctx: Context): void {
  const presentation = ctx.parts.get('ppt/presentation.xml')
  const rels = ctx.parts.get('ppt/_rels/presentation.xml.rels')
  const targets = new Map((rels ? all(rels, 'Relationship') : []).filter(r => attr(r,'TargetMode') !== 'External').map(r => [attr(r,'Id'), targetPath(attr(r,'Target') ?? '')]))
  const hidden = new Set<string>()
  for (const [part, root] of [...ctx.parts]) {
    if (!part.startsWith('ppt/')) continue
    if (/^ppt\/(notesSlides|notesMasters|comments|commentAuthors)\//.test(part) || /^ppt\/(commentAuthors|comments)[^/]*\.xml$/.test(part)) { ctx.dropPart(part); continue }
    if (/^ppt\/slides\/[^/]+\.xml$/.test(part)) {
      const slide = all(root, 'sld')[0]
      if (slide && isFalse(attr(slide,'show'))) { hidden.add(part); ctx.dropPart(part); ctx.diagnostics.push({code:'hidden-stripped',part,element:slide.name,attr:'show',value:attr(slide,'show')}); continue }
      if (slide && isTrue(attr(slide,'show'))) ctx.handled.add(slide)
    }
    function visit(node: Node) {
      node.children = node.children.filter(current => {
        if (shapes.has(local(current.name))) {
          const properties = (n: Node): Node[] => n.children.flatMap(c => shapes.has(local(c.name)) ? [] : local(c.name) === 'cNvPr' ? [c] : properties(c))
          const flag = properties(current).find(p => isTrue(attr(p, 'hidden')))
          if (flag) { ctx.diagnostics.push({code:'hidden-stripped',part,element:flag.name,attr:'hidden',value:attr(flag,'hidden')}); return false }
        }
        visit(current); return true
      })
    }
    visit(root)
  }
  if (presentation) for (const slideId of all(presentation, 'sldId')) {
    const path = targets.get(slideId.attrs['r:id'] ?? Object.entries(slideId.attrs).find(([key]) => key.includes(':') && local(key) === 'id')?.[1])
    if (isFalse(attr(slideId,'show')) || (path && hidden.has(path))) {
      remove(presentation, slideId)
      if (path && ctx.parts.has(path)) { ctx.dropPart(path); ctx.diagnostics.push({code:'hidden-stripped',part:'ppt/presentation.xml',element:slideId.name,attr:'show',value:attr(slideId,'show')}) }
    } else if (isTrue(attr(slideId,'show'))) ctx.handled.add(slideId)
  }
}
