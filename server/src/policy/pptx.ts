import type { PptxPackageHandle } from '../../../src/pptx/types'
import { all, attr, local, readTree, remove, type Context, type Node } from './tree'

const shapes = new Set(['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'])
const isTrue = (value: string | undefined) => ['1', 'true', 'on'].includes(value ?? '')
const isFalse = (value: string | undefined) => ['0', 'false', 'off'].includes(value ?? '')
function targetPath(target: string): string {
  const segments = (target.startsWith('/') ? target.slice(1) : `ppt/${target}`).split('/')
  const result: string[] = []
  for (const segment of segments) { if (segment === '..') result.pop(); else if (segment && segment !== '.') result.push(segment) }
  return result.join('/')
}
/** The hiding rule shared by the visible-only projection and the 'full'
 * flagger (`hiddenSlideUnits`), so the two cannot drift: a slide is hidden by
 * its root (`<p:sld show="0">`) or by its presentation entry (`<p:sldId show="0">`). */
const slideTargets = (rels: Node | undefined) => new Map((rels ? all(rels, 'Relationship') : []).filter(r => attr(r,'TargetMode') !== 'External').map(r => [attr(r,'Id'), targetPath(attr(r,'Target') ?? '')]))
const slideIdPath = (slideId: Node, targets: Map<string | undefined, string>) => targets.get(slideId.attrs['r:id'] ?? Object.entries(slideId.attrs).find(([key]) => key.includes(':') && local(key) === 'id')?.[1])
const hiddenByRoot = (root: Node) => { const slide = all(root, 'sld')[0]; return !!slide && isFalse(attr(slide,'show')) }
const hiddenByEntry = (slideId: Node) => isFalse(attr(slideId,'show'))

export function projectPptx(ctx: Context): void {
  const presentation = ctx.parts.get('ppt/presentation.xml')
  const targets = slideTargets(ctx.parts.get('ppt/_rels/presentation.xml.rels'))
  const hidden = new Set<string>()
  for (const [part, root] of [...ctx.parts]) {
    if (!part.startsWith('ppt/')) continue
    if (/^ppt\/(notesSlides|notesMasters|comments|commentAuthors)\//.test(part) || /^ppt\/(commentAuthors|comments)[^/]*\.xml$/.test(part)) { ctx.dropPart(part); continue }
    if (/^ppt\/slides\/[^/]+\.xml$/.test(part)) {
      const slide = all(root, 'sld')[0]
      if (slide && hiddenByRoot(root)) { hidden.add(part); ctx.dropPart(part); ctx.diagnostics.push({code:'hidden-stripped',part,element:slide.name,attr:'show',value:attr(slide,'show')}); continue }
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
    const path = slideIdPath(slideId, targets)
    if (hiddenByEntry(slideId) || (path && hidden.has(path))) {
      remove(presentation, slideId)
      if (path && ctx.parts.has(path)) { ctx.dropPart(path); ctx.diagnostics.push({code:'hidden-stripped',part:'ppt/presentation.xml',element:slideId.name,attr:'show',value:attr(slideId,'show')}) }
    } else if (isTrue(attr(slideId,'show'))) ctx.handled.add(slideId)
  }
}

/** Under `'full'` hidden slides stay units, so the show must be told to skip
 * them: the unit indices the visible-only projection above would drop (same
 * predicates), read from the original package once at open.
 * `handle.slideParts[i]` is unit i (the parser's lease keeps them aligned).
 * Its reads are uncached afterwards, whether or not it throws. */
export async function hiddenSlideUnits(handle: PptxPackageHandle): Promise<number[]> {
  try {
    const read = async (path: string) => { const text = await handle.pkg.text(path); return text === undefined ? undefined : readTree(text) }
    const presentation = await read('ppt/presentation.xml')
    const targets = slideTargets(await read('ppt/_rels/presentation.xml.rels'))
    const hiddenParts = new Set<string>()
    for (const slideId of presentation ? all(presentation, 'sldId') : []) {
      const path = slideIdPath(slideId, targets)
      if (path && hiddenByEntry(slideId)) hiddenParts.add(path)
    }
    const out: number[] = []
    for (const [i, path] of handle.slideParts.entries()) {
      if (hiddenParts.has(path)) { out.push(i); continue }
      const root = await read(path)
      if (root && hiddenByRoot(root)) { hiddenParts.add(path); out.push(i) }
    }
    return out
  } finally { handle.pkg.dropCache() }
}
