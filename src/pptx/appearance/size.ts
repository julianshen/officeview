/**
 * Slide size validation for the lazy slideshow entry and the appearance enrichment (slideshow spec 7),
 * run before any allocation, projection or letterbox maths. ECMA-376 Part 1 pml.xsd
 * `ST_SlideSizeCoordinate`: xsd:int, minInclusive 914400, maxInclusive 51206400 (1 to 56 inches).
 */
export const SLIDE_EMU_MIN = 914_400
export const SLIDE_EMU_MAX = 51_206_400
/** The longer side may be at most this many times the shorter (1 : 20). */
export const SLIDE_ASPECT_MAX = 20

export type SlideSizeCheck = { ok: true } | { ok: false; diagnostic: { kind: 'slide-size'; message: string } }

export function validateSlideSize(widthEmu: number, heightEmu: number): SlideSizeCheck {
  const fail = (why: string): SlideSizeCheck => ({ ok: false, diagnostic: { kind: 'slide-size', message: `Slide size ${widthEmu} x ${heightEmu} EMU ${why}` } })
  if (!Number.isFinite(widthEmu) || !Number.isFinite(heightEmu) || widthEmu <= 0 || heightEmu <= 0) return fail('is not finite and positive')
  if (widthEmu < SLIDE_EMU_MIN || heightEmu < SLIDE_EMU_MIN || widthEmu > SLIDE_EMU_MAX || heightEmu > SLIDE_EMU_MAX) return fail(`is outside ${SLIDE_EMU_MIN}-${SLIDE_EMU_MAX} EMU`)
  if (Math.max(widthEmu, heightEmu) > SLIDE_ASPECT_MAX * Math.min(widthEmu, heightEmu)) return fail(`has an aspect ratio beyond 1 : ${SLIDE_ASPECT_MAX}`)
  return { ok: true }
}
