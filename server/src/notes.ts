/**
 * Stage 3c speaker notes (§12a.5, §12a.8; plan calls C2–C4): the notes
 * slide's body text, read from the ORIGINAL package with the local reader's
 * rule (`readNotes`: body placeholders only, 65,536-char cap), laid out as
 * plain text on white, 720 CSS px wide, at most NOTES_MAX_PX device px tall.
 * Text that does not fit ends with NOTES_CONTINUATION, never a silent cut.
 * Notes travel only as this image (NOTES_RASTER, kind 7); never as text.
 */
import { posix } from 'node:path'
import { parseXmlOrdered } from '../../src/core/xml'
import type { CanvasBackend } from '../../src/core/canvas-backend'
import type { OfficePackage } from '../../src/core/zip'
import { READER_LIMITS } from '../../src/slideshow/read/limits'
import { readNotes } from '../../src/slideshow/read/notes'
import type { NoteParagraph } from '../../src/slideshow/types'
import { encodeRasterBody } from '../../src/stream/raster'
import { all, attr, readTree } from './policy/tree'

export const NOTES_WIDTH_CSS = 720
export const NOTES_MAX_PX = 8192
export const NOTES_RATE_PER_MIN = 30
/** §14 NOTES_RASTER (kind 7) encoded cap; its decoded cap is the UNIT_RASTER one (64 MiB). */
export const NOTES_ENCODED_CAP = 8 * 1024 * 1024
export const NOTES_CONTINUATION = 'Notes continue — too long to show in full.'
/** The pack family the notes are drawn in (a generic family would resolve per host). */
export const NOTES_FONT = '16px "Carlito"'
const PAD = 24, LINE = 24, INDENT = 24, BASELINE = 17

/** The Layer-1 basis label of a notes image: never the slide's `${unit}@${scale}`. */
export const notesBasisLabel = (unitIndex: number, scale: number) => `notes:${unitIndex}@${scale}`

/** NOTES_RASTER body: `u32 LE unitIndex` + the UNIT_RASTER body (`u32 w, u32 h, u32 pngLen, PNG`). */
export function encodeNotesBody(unitIndex: number, w: number, h: number, png: Uint8Array): Uint8Array {
  const raster = encodeRasterBody(w, h, png), out = new Uint8Array(4 + raster.length)
  new DataView(out.buffer).setUint32(0, unitIndex, true)
  out.set(raster, 4)
  return out
}

/** The notes of the slide at `slidePart` in `pkg` (the original package): its
 * non-external notesSlide relationship, read under the reader's byte caps.
 * No relationship or no notes part ⇒ no paragraphs. */
export async function readSlideNotes(pkg: OfficePackage, slidePart: string, index: number): Promise<NoteParagraph[]> {
  const dir = posix.dirname(slidePart)
  const rels = await pkg.boundedText(posix.join(dir, '_rels', `${posix.basename(slidePart)}.rels`), READER_LIMITS.relsBytes)
  if (rels === undefined) return []
  const rel = all(readTree(rels), 'Relationship').find((r) => attr(r, 'TargetMode') !== 'External' && (attr(r, 'Type') ?? '').endsWith('/notesSlide'))
  const target = rel && attr(rel, 'Target')
  if (!target) return []
  const path = target.startsWith('/') ? target.slice(1) : posix.normalize(posix.join(dir, target))
  const xml = await pkg.boundedText(path, READER_LIMITS.notesBytes)
  if (xml === undefined) return []
  return readNotes(parseXmlOrdered(xml), index, { add() {}, list: () => [] })
}

/** Local Presenter View's bullet rule: an authored character, or a plain bullet for an auto-number scheme. */
const prefixOf = (n: NoteParagraph) => (n.bullet ? `${n.bullet.length <= 2 ? n.bullet : '•'} ` : '')

/** `y`: the line's top in CSS px. */
interface Line { x: number; y: number; text: string; chars: number }
export interface NotesLayout { w: number; h: number; truncated: boolean; chars: number; lines: Line[] }

/** Greedy word wrap of one hard line into `width`; a word wider than the line breaks by code points. */
function wrap(text: string, width: number, measure: (s: string) => number, out: (text: string) => boolean): boolean {
  if (text === '') return out('')
  let line = ''
  for (const token of text.split(/(\s+)/).filter(Boolean)) {
    if (measure(line + token) <= width) { line += token; continue }
    if (line && !out(line.trimEnd())) return false
    line = ''
    if (/^\s+$/.test(token)) continue
    let rest = [...token]
    while (measure(rest.join('')) > width && rest.length > 1) {
      let lo = 1, hi = rest.length - 1
      while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (measure(rest.slice(0, mid).join('')) <= width) lo = mid; else hi = mid - 1 }
      if (!out(rest.slice(0, lo).join(''))) return false
      rest = rest.slice(lo)
    }
    line = rest.join('')
  }
  return out(line.trimEnd())
}

/** The notes laid out in CSS px at `scale`: every line that fits under
 * NOTES_MAX_PX device px; when some do not, the last line is the continuation
 * line and the image is exactly NOTES_MAX_PX tall. `chars` counts the notes
 * characters on drawn lines (prefixes and the continuation line excluded). */
export function layoutNotes(paragraphs: readonly NoteParagraph[], scale: 1 | 2, measure: (s: string) => number): NotesLayout {
  const maxH = Math.floor(NOTES_MAX_PX / scale), room = Math.floor((maxH - 2 * PAD) / LINE)
  const lines: Line[] = []
  let truncated = false
  for (const p of paragraphs) {
    const x = PAD + Math.min(p.level, 8) * INDENT, width = NOTES_WIDTH_CSS - PAD - x
    let prefix = prefixOf(p)
    for (const hard of p.text.split('\n')) {
      const ok = wrap(prefix + hard, width, measure, (text) => {
        if (lines.length >= room) { truncated = true; return false }
        // The prefix is not notes text.
        const chars = text.length - Math.min(prefix.length, text.length)
        prefix = prefix.slice(Math.min(prefix.length, text.length))
        lines.push({ x, y: PAD + lines.length * LINE, text, chars })
        return true
      })
      prefix = ''
      if (!ok) break
    }
    if (truncated) break
  }
  if (truncated) {
    lines.splice(room - 1)
    // Anchored to the bottom edge: the image's last rows always show that the notes go on.
    lines.push({ x: PAD, y: maxH - PAD - LINE, text: NOTES_CONTINUATION, chars: 0 })
  }
  const h = truncated ? maxH : 2 * PAD + Math.max(1, lines.length) * LINE
  return { w: NOTES_WIDTH_CSS * scale, h: h * scale, truncated, chars: lines.reduce((n, l) => n + l.chars, 0), lines }
}

const measurer = (backend: CanvasBackend) => {
  const ctx = backend.createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.font = NOTES_FONT
  return (s: string) => ctx.measureText(s).width
}

/** Plans the image without allocating it (the audit row is written between plan and draw). */
export function planNotes(paragraphs: readonly NoteParagraph[], scale: 1 | 2, backend: CanvasBackend): Omit<NotesLayout, 'lines'> {
  const { lines: _, ...plan } = layoutNotes(paragraphs, scale, measurer(backend))
  return plan
}

/** Draws the notes on white (device size `w × h`); the caller paints the visible mark on `ctx` (CSS units). */
export function drawNotes(paragraphs: readonly NoteParagraph[], scale: 1 | 2, backend: CanvasBackend): NotesLayout & { canvas: ReturnType<CanvasBackend['createCanvas']>; ctx: CanvasRenderingContext2D } {
  const layout = layoutNotes(paragraphs, scale, measurer(backend))
  const canvas = backend.createCanvas(layout.w, layout.h)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.scale(scale, scale)
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, NOTES_WIDTH_CSS, layout.h / scale)
  ctx.fillStyle = '#1a1a1a'
  ctx.font = NOTES_FONT
  ctx.textBaseline = 'alphabetic'
  for (const l of layout.lines) ctx.fillText(l.text, l.x, l.y + BASELINE)
  return { ...layout, canvas, ctx }
}
