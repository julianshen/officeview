/** Show diagnostics (slideshow spec 6.4): deduplicated per code, slide and effect. */
import type { DiagnosticSink, SlideShowDiagnostic } from './types'

export function createDiagnosticSink(): DiagnosticSink {
  const seen = new Set<string>()
  const items: SlideShowDiagnostic[] = []
  return {
    add(d) {
      // A JSON tuple keeps an absent slide or effect distinct from any literal value.
      const key = JSON.stringify([d.code, d.slideIndex ?? null, d.effect ?? null])
      if (seen.has(key)) return
      seen.add(key)
      items.push(d)
    },
    list: () => items,
  }
}
