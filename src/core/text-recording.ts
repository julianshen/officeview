/** Optional recording protocol for renderers that know their logical source text. */
export const RECORD_TEXT = Symbol('officeview.recordText')
export interface LogicalTextSource {
  readonly text: string
  /** Paragraphs from one paint body share a scope and retain source order. */
  readonly scope?: object
  readonly order?: number
}
export interface LogicalTextRange {
  source: LogicalTextSource; start: number; end: number
  /** Optional physical viewport. Full source remains recorded; visible bands
   * and hit candidates are intersected with this caller-space rectangle. */
  clip?: { x: number; y: number; width: number; height: number;
    transform?: { a: number; b: number; c: number; d: number; e: number; f: number } }
  /** Optional measured ink in local alphabetic-baseline coordinates. */
  ink?: { ascent: number; descent: number }
  /** A source-ordered visual wrap/column can contain differently rotated runs. */
  line?: number
  /** Visual progression of a logical line, independent of individual glyph axes. */
  flow?: 'vertical'
  run?: number
  /** Absolute UTF-16 boundaries in the paragraph source. */
  graphemeBoundaries?: readonly number[]
  /** Appearance passes may paint offset ink while indexing this canonical band.
   * x/y are text coordinates under `transform` (or the current transform when
   * omitted); font, alignment and baseline are those of the recording call. */
  canonical?: {
    x: number; y: number; width: number
    transform?: { a: number; b: number; c: number; d: number; e: number; f: number }
  }
  /** Optional mapped visual geometry for curved ink (warp): the ink's mapped
   * polygon plus per-grapheme-cluster mapped centers, both in the recording
   * transform's coordinate space. Same logical record — never extra records;
   * search/highlights/hits may use it while copy stays canonical. */
  visual?: {
    polygon: ReadonlyArray<[number, number]>
    clusters?: ReadonlyArray<[number, number]>
    /** Small convex mesh cells covering the mapped band. Hit testing uses
     * this cell union (robust to global polygon winding on folded bands). */
    cells?: ReadonlyArray<readonly [[number, number], [number, number], [number, number], [number, number]]>
    /** Mesh triangles covering the mapped band (authoritative for containment;
     * triangles stay convex even when quads twist). */
    tris?: ReadonlyArray<readonly [[number, number], [number, number], [number, number]]>
    /** Per-grapheme-cluster mapped band bounds [x, y, width, height] in the
     * same order as `clusters`: the axis-aligned bound of the mapped mesh
     * cells for that cluster's own source interval. Highlights union these
     * (intersected with the viewport) instead of centering nominal font bands
     * on the cluster centers. Minimal type extension for the highlight
     * coverage fix: paint computes it where glyph advances are known. */
    bands?: ReadonlyArray<readonly [number, number, number, number]>
  }
}
export type RecordText = (text: string, x: number, y: number, width: number, logical: LogicalTextRange) => void
export type TextRecordingContext = CanvasRenderingContext2D & { [RECORD_TEXT]?: RecordText }

/** UTF-16 boundaries, shared by layout and public selection carets. */
export function graphemes(text: string): Array<{ text: string; start: number }> {
  // Intl.Segmenter is present in supported browsers/Node; the fallback retains
  // combining marks and ZWJ sequences even on an older JS host.
  const Segmenter = (Intl as unknown as { Segmenter?: new (locale?: string, options?: { granularity: string }) => { segment(text: string): Iterable<{ segment: string; index: number }> } }).Segmenter
  if (Segmenter) return [...new Segmenter(undefined, { granularity: 'grapheme' }).segment(text)].map(s => ({ text: s.segment, start: s.index }))
  const out: Array<{ text: string; start: number }> = []
  let offset = 0, joined = false, regionalCount = 0
  for (const char of text) {
    const regional = /\p{Regional_Indicator}/u.test(char)
    if (out.length && (joined || char === '\u200d' || /[\p{M}\p{Emoji_Modifier}\u{E0020}-\u{E007F}]/u.test(char) || regional && regionalCount % 2 === 1)) out[out.length - 1].text += char
    else out.push({ text: char, start: offset })
    joined = char === '\u200d'; regionalCount = regional ? regionalCount + 1 : 0; offset += char.length
  }
  return out
}
export function snapGrapheme(text: string, offset: number, edge: 'nearest' | 'floor' | 'ceil' = 'nearest'): number {
  const at = Math.max(0, Math.min(text.length, offset))
  for (const cluster of graphemes(text)) {
    const end = cluster.start + cluster.text.length
    if (at > cluster.start && at < end) return edge === 'floor' ? cluster.start : edge === 'ceil' ? end : at - cluster.start <= end - at ? cluster.start : end
  }
  return at
}
