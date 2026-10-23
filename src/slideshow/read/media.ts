import { attrs, elementChildren, getChild, getChildren, type XmlNode } from '../../core/xml'
import type { OfficePackage } from '../../core/zip'
import { alternateBranchSelector } from '../../pptx/parse'
import type { DiagnosticSink, MediaRecord } from '../types'
import type { Relationship } from './deck'
import { READER_LIMITS } from './limits'
/**
 * The shape each media relationship plays on (S2-7): `p:nvPr/a:videoFile|a:audioFile@r:link` and
 * `p:nvPr/p:extLst/p:ext/p14:media@r:embed`, walked through groups and the rendered branch of
 * `mc:AlternateContent` - the one the drawing parser selects (`alternateBranchSelector`, its compatibility
 * rules; PR #19 review) - bounded at 64 levels like the reader.
 */
function mediaShapes(slide: XmlNode | undefined): Map<string, string[]> {
  // Every shape a relationship plays on: two shapes may share one media part (PR #19 review 10).
  const out = new Map<string, string[]>()
  const select = alternateBranchSelector()
  const walk = (n: XmlNode | undefined, depth: number) => {
    if (!n || depth > 64) return
    for (const [kind, shape] of elementChildren(n)) {
      if (kind === 'grpSp') { walk(shape, depth + 1); continue }
      if (kind === 'AlternateContent') { walk(select(shape), depth + 1); continue }
      const nv = getChild(shape, 'nvPicPr') ?? getChild(shape, 'nvSpPr') ?? getChild(shape, 'nvGraphicFramePr')
      const spid = attrs(getChild(nv, 'cNvPr')).id
      const nvPr = getChild(nv, 'nvPr')
      if (!spid || !nvPr) continue
      const ids = [attrs(getChild(nvPr, 'videoFile')).link, attrs(getChild(nvPr, 'audioFile')).link]
      for (const ext of getChildren(getChild(nvPr, 'extLst'), 'ext')) ids.push(attrs(getChild(ext, 'media')).embed)
      for (const id of new Set(ids)) {
        if (!id) continue
        const list = out.get(id)
        if (!list) out.set(id, [spid])
        else if (!list.includes(spid)) list.push(spid)
      }
    }
  }
  walk(getChild(getChild(slide, 'cSld'), 'spTree'), 0)
  return out
}

export function readMedia(rels: readonly Relationship[], types: XmlNode | undefined, pkg: OfficePackage, index: number, sink: DiagnosticSink, slide?: XmlNode): MediaRecord[] {
  const out: MediaRecord[] = []
  const shapes = mediaShapes(slide)
  for (const r of rels) {
    if (!/\/(audio|video|media)$/.test(r.type))
      continue
    if (r.external) {
      sink.add({ code: 'media-refused', slideIndex: index, effect: r.id, message: 'External media is not read' })
      continue
    }
    const entry = pkg.entry(r.target), size = (entry as unknown as {
      _data?: {
        uncompressedSize?: number
      }
    })?._data?.uncompressedSize
    if (!entry || !Number.isSafeInteger(size) || size! < 0) {
      sink.add({ code: 'media-refused', slideIndex: index, effect: r.id, message: 'Missing media or invalid declared size' })
      continue
    }
    // OPC part names and extensions compare case-insensitively.
    const override = getChildren(types, 'Override').find(n => attrs(n).PartName?.toLowerCase() === '/' + r.target.toLowerCase())
    const def = getChildren(types, 'Default').find(n => attrs(n).Extension?.toLowerCase() === r.target.split('.').pop()?.toLowerCase())
    if (size! > READER_LIMITS.mediaBytes)
      sink.add({ code: 'media-limit', slideIndex: index, effect: r.id, message: 'Declared media exceeds 512 MiB; playback must refuse' })
    const record = { rId: r.id, partPath: r.target, contentType: attrs(override ?? def).ContentType ?? 'application/octet-stream', declaredBytes: size! }
    // One record per shape the relationship plays on (none: a sound with no shape, an effect or transition sound).
    const spids = shapes.get(r.id)
    if (!spids?.length) out.push(record)
    else for (const spid of spids) out.push({ ...record, spid })
  }
  return out
}
