/**
 * Layout and master artwork ids for the timing target resolver (S1-19 review M2), shared by the local show
 * (`player/session.ts`) and the stream server's show plan (piece 4 Task 6 follow-up: targets must resolve
 * identically on both sides). DOM-free.
 *
 * Parsed once per show and layout: the resolver only needs the ids, so the cache keeps an id-only skeleton (no
 * geometry, text or image bytes). A slide that hides the master shapes (`showMasterSp="0"`) has no underlay.
 * Inherited transitions are diagnosed on the way (they are not inherited).
 */
import { throwIfAborted } from '../../core/package-error'
import { attrs } from '../../core/xml'
import { HandlePackage } from '../../pptx/appearance/handle-package'
import { masterArtwork } from '../../pptx/appearance/master-art'
import { slideTheme } from '../../pptx/parse'
import type { PptxPackageHandle, PptxShape } from '../../pptx/types'
import type { DiagnosticSink, MetadataHandle, ShowIndex, SlideMetadataStore } from '../types'
import { readRelationships } from './deck'
import { diagnoseInheritedTransition } from './transition'

export interface UnderlayReader {
  /** Slide `i`'s layout and master artwork skeleton (`[]` when there is none, or it cannot be read: diagnosed). */
  underlayOf(i: number): Promise<readonly PptxShape[]>
  /** The resolver's underlay of slide `i` while its metadata is being read (`acquire`). */
  readonly underlays: Map<number, readonly PptxShape[]>
  /** Acquires slide `i`'s metadata from `store` with its underlay in place for the resolver. */
  acquire(store: SlideMetadataStore, i: number): Promise<MetadataHandle>
}

export function createUnderlayReader(handle: PptxPackageHandle | undefined, index: ShowIndex, sink: DiagnosticSink, signal?: AbortSignal): UnderlayReader {
  const underlayCache = new Map<string, Promise<readonly PptxShape[]>>()
  const underlays = new Map<number, readonly PptxShape[]>()
  const skeleton = (list: readonly PptxShape[], depth = 0): PptxShape[] =>
    depth > 64 ? [] : list.map((s) => ({ source: s.source && { ...s.source }, children: s.children && skeleton(s.children, depth + 1) }) as PptxShape)
  async function underlayOf(i: number): Promise<readonly PptxShape[]> {
    const part = index.slides[i]?.partPath
    if (!handle || !part) return []
    try {
      const rels = await readRelationships(handle, part, signal)
      const local = (list: typeof rels, kind: string) => list.find((r) => !r.external && r.type?.endsWith(`/${kind}`))?.target
      const layout = local(rels, 'slideLayout')
      const master = layout ? local(await readRelationships(handle, layout, signal), 'slideMaster') : local(rels, 'slideMaster')
      // A private, bounded view over the handle: its cache dies with this call.
      const pkg = HandlePackage.over(handle)
      for (const [path, kind] of [[layout, 'layout'], [master, 'master']] as const) {
        const root = path ? await pkg.xmlOrdered(path) : undefined
        if (root) diagnoseInheritedTransition(root, kind, i, sink)
      }
      if (['0', 'false'].includes(attrs(await pkg.xmlOrdered(part)).showMasterSp ?? '')) return []
      const key = `${layout ?? ''}|${master ?? ''}`
      let art = underlayCache.get(key)
      if (!art) {
        art = slideTheme(pkg, part).then((theme) => masterArtwork(pkg, part, theme.context)).then((a) => skeleton([...a.master, ...a.layout]))
        underlayCache.set(key, art)
        art.catch(() => underlayCache.delete(key))
      }
      return await art
    } catch (e) {
      throwIfAborted(signal)
      sink.add({ code: 'underlay', slideIndex: i, message: `Layout and master artwork could not be read (${e instanceof Error ? e.message : String(e)}); master-only targets read as missing` })
      return []
    }
  }
  async function acquire(store: SlideMetadataStore, i: number): Promise<MetadataHandle> {
    const mine = store.resident().includes(i) ? undefined : await underlayOf(i)
    if (mine) underlays.set(i, mine)
    try {
      return await store.acquire(i, signal)
    } finally {
      // A second visit to the same slide may have joined the same read with its own copy: leave that one.
      if (mine && underlays.get(i) === mine) underlays.delete(i)
    }
  }
  return { underlayOf, underlays, acquire }
}
