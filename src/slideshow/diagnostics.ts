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

/**
 * What a diagnostic says about the show (slideshow plan S1-22):
 * - `stand-in`: a catalogued stand-in plays (pending a later stage);
 * - `approximate`: the closest look-alike plays, or a part of an effect is dropped while the rest plays;
 * - `refused`: input is refused or dropped (a hostile or invalid value, a limit, an unreadable part);
 * - `resource`: memory, GPU or platform pressure changed how the show renders (ladder, cuts, fallbacks);
 * - `info`: the show works as the file asks, with a note (ignored presentation-only attributes, host calls).
 */
export type DiagnosticClass = 'stand-in' | 'approximate' | 'refused' | 'resource' | 'info'

/**
 * The diagnostic code catalogue: every code the slideshow emits, with its class.
 * tests/slideshow-corpus.test.ts checks that each emitted code and each code
 * literal under src/slideshow is listed here.
 */
export const DIAGNOSTIC_CODES: Readonly<Record<string, DiagnosticClass>> = Object.freeze({
  // Show setup and session
  'slide-size': 'refused', 'no-package': 'refused', 'package-reopen': 'refused', 'show-properties': 'refused',
  'browse': 'approximate', 'custom-show': 'approximate', 'slide-range': 'refused', 'slide-root': 'refused',
  'presenter-view-unavailable': 'stand-in', 'fullscreen-refused': 'resource', 'invalid-slide': 'info', 'kiosk-restart': 'refused',
  'underlay': 'refused', 'workspace': 'resource',
  // Slide metadata
  'notes-limit': 'refused', 'notes-truncated': 'refused', 'metadata-limit': 'resource', 'slide-refused': 'refused',
  'media-refused': 'refused', 'media-limit': 'refused', 'link-refused': 'refused', 'link-open-failed': 'resource',
  // Transitions
  'transition-approximate': 'approximate', 'transition-stand-in': 'stand-in', 'transition-unknown': 'refused',
  'transition-unknown-preset': 'approximate', 'transition-extension': 'refused', 'transition-invalid-value': 'refused',
  'transition-requires': 'info', 'transition-inherited': 'info', 'transition-sound': 'refused', 'transition-workspace': 'resource',
  // Timing tree (reader)
  'timing-limit': 'refused', 'timing-malformed': 'refused', 'timing-invalid-value': 'refused', 'timing-ignored': 'info',
  'timing-carried': 'info', 'timing-approximated': 'approximate', 'timing-requires': 'info', 'timing-duplicate-id': 'refused',
  'timing-missing-reference': 'refused', 'timing-syncbase-no-event': 'refused', 'timing-unknown-attribute': 'refused',
  'timing-unknown-attribute-name': 'refused', 'timing-unknown-container': 'refused', 'timing-unknown-element': 'refused',
  'motion-path-deferred': 'approximate', 'master-target': 'refused', 'unresolved-target': 'refused',
  // Scheduler
  'timing-cycle': 'refused', 'timing-too-deep': 'refused', 'timing-forward-reference': 'refused',
  'timing-too-many-activations': 'refused', 'timing-history-truncated': 'refused', 'timing-reference-outside-window': 'refused',
  'timing-query-budget': 'refused',
  // Behaviours
  'preset-approximate': 'approximate', 'behaviour-unsupported': 'approximate', 'behaviour-invalid-formula': 'refused',
  'anim-effect-unknown-filter': 'approximate', 'anim-clr-unreadable': 'approximate', 'media-stand-in': 'stand-in',
  // Rendering, memory and GPU
  'ladder-reduced': 'resource', 'ladder-flattened': 'resource', 'ladder-terminal': 'resource', 'layer-limit': 'resource',
  'reraster-limit': 'resource', 'layer-workspace': 'resource', 'layer-scale': 'resource', 'image-workspace': 'resource',
  'scratch-image': 'resource', 'paint-failed': 'resource', 'resize-kept-old': 'resource', 'mask-ops': 'resource',
  'mask-unknown': 'approximate', 'compositor-workspace': 'resource', 'compositor-handle': 'resource',
  'backend-fallback': 'resource', 'context-lost': 'resource',
  // Stage 2: media playback (S2-7)
  'media-autoplay': 'resource', 'media-codec': 'resource', 'media-error': 'resource', 'media-missing': 'refused',
  'media-command': 'refused', 'media-overlay': 'approximate', 'media-fullscreen': 'approximate',
  // Stage 2: text builds by word or letter (S2-8)
  'text-build-approximate': 'approximate', 'text-build-unsplit': 'approximate', 'text-build-coarsened': 'approximate',
  'text-build-cjk-chars': 'approximate', 'text-build-refused': 'refused', 'text-build-whole': 'resource',
  // Stage 2: Presenter View, its window and screen placement (S2-1 to S2-5)
  'popup-blocked': 'resource', 'presenter-window-failed': 'resource', 'presenter-styles-blocked': 'resource',
  'presenter-paint': 'resource', 'presenter-surface': 'resource', 'presenter-thumbnail-approximate': 'approximate',
  'preview-prime': 'resource', 'placement-permission-denied': 'resource', 'placement-api-missing': 'resource',
  'placement-single-screen': 'info', 'placement-failed': 'resource', 'placement-fullscreen-refused': 'resource',
})
