/** XML normalization shared by runtime parsing and the offline preset generator. */
import { attrs, getChild, getChildren, orderedChildren, type XmlNode } from '../core/xml'
import type { GeometryDefinition, GeometryCommand, Guide } from './geometry'

export const parseGuideList = (node: XmlNode | undefined): Guide[] => getChildren(node, 'gd').map(gd => {
  const a = attrs(gd)
  return [a.name ?? '', a.fmla ?? '']
})

export function parseCustomGeometry(node: XmlNode): GeometryDefinition {
  const definition: GeometryDefinition = {
    adjustments: parseGuideList(getChild(node, 'avLst')),
    guides: parseGuideList(getChild(node, 'gdLst')),
    paths: getChildren(getChild(node, 'pathLst'), 'path').map(path => {
      const a = attrs(path)
      const commands: GeometryCommand[] = []
      for (const [kind, child] of orderedChildren(path)) {
        if (kind === '#text') continue
        if (kind === 'arcTo') {
          const arc = attrs(child)
          commands.push([kind, arc.wR ?? '', arc.hR ?? '', arc.stAng ?? '', arc.swAng ?? ''])
        } else {
          const points = getChildren(child, 'pt').flatMap(pt => {
            const p = attrs(pt)
            return [p.x ?? '', p.y ?? '']
          })
          // Unknown commands are retained so their containing path is audited and skipped.
          commands.push([kind, ...points])
        }
      }
      return {
        ...(a.w !== undefined ? { width: a.w } : {}),
        ...(a.h !== undefined ? { height: a.h } : {}),
        ...(a.fill !== undefined ? { fill: a.fill } : {}),
        ...(a.stroke !== undefined ? { stroke: a.stroke !== 'false' && a.stroke !== '0' } : {}),
        commands,
      }
    }),
  }
  const rect = getChild(node, 'rect')
  if (rect) {
    const a = attrs(rect)
    definition.textRect = [a.l ?? '', a.t ?? '', a.r ?? '', a.b ?? '']
  }
  return definition
}

