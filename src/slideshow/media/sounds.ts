/**
 * Transition sounds (slideshow plan S2-7; spec 2.1 goal 3 "transition sound
 * plays where the browser supports its format", 6.1 `p:sndAc`, 7 cross-slide
 * sound).
 *
 * `start(slide)` runs when a slide is entered: its `p:endSnd` stops the
 * sound still playing, and its `p:stSnd` replaces it. `stopAt(slide)` runs
 * when a slide is left: a sound that does not loop stops there ("everything
 * else stops on slide change"); a looping one plays on until the next
 * slide's own sound or `p:endSnd`. A looping sound's lease is retained
 * indefinitely and released when it stops; the element lets go of its URL
 * before the lease is released (S2-6). The format is checked with
 * `canPlayType` before any byte is read; an undecodable or refused sound,
 * or a rejected `play()`, is diagnosed and the show goes on (there is no
 * poster to click, so no "Click to play"); a rejected sound is stopped and
 * its lease released, so a later resume never starts it mid-slide.
 */
import type { DiagnosticSink, MediaRecord, ShowSlide } from '../types'
import type { MediaLease, MediaLeaseResult } from './lease'

export interface TransitionSound {
  start(slide: number): void
  stopAt(slide: number): void
}

export interface TransitionSoundDeps {
  root: HTMLElement
  sink: DiagnosticSink
  /** The slide's metadata, when it is in. */
  slide(index: number): ShowSlide | undefined
  /** The lease opener (the media player's, so the pressure rule applies). */
  open(rec: MediaRecord, signal: AbortSignal): Promise<MediaLeaseResult>
  /** @internal Element factory (tests). */
  create?(kind: 'audio'): HTMLMediaElement
}

export interface TransitionSoundController extends TransitionSound {
  setPaused(paused: boolean): void
  /**
   * Stops the sound now; true when one was playing. Workspace pressure and "Stop previous sound" stop any;
   * `onlyLooping` (an effect sound starting: "loop until next sound") stops only a looping one (S2-10).
   */
  release(onlyLooping?: boolean): boolean
  dispose(): void
}

interface Playing {
  slide: number
  loop: boolean
  el?: HTMLMediaElement
  lease?: MediaLease
  abort: AbortController
}

export function createTransitionSound(deps: TransitionSoundDeps): TransitionSoundController {
  const doc = deps.root.ownerDocument
  const create = deps.create ?? ((kind) => doc.createElement(kind))
  let playing: Playing | undefined
  let paused = false

  const diagnose = (slide: number, code: string, effect: string, message: string) => deps.sink.add({ code, slideIndex: slide, effect, message })

  function stop(): boolean {
    const p = playing
    if (!p) return false
    playing = undefined
    p.abort.abort()
    if (p.el) {
      try { p.el.pause() } catch { /* a detached realm */ }
      p.el.removeAttribute('src')
      try { p.el.load() } catch { /* nothing to load */ }
      p.el.remove()
    }
    p.lease?.release()
    return true
  }

  function play(p: Playing): void {
    const el = p.el
    if (!el || paused) return
    let r: Promise<void> | undefined
    // A refused sound retires: stopped and its lease released, so a later resume (a hidden tab coming back)
    // never starts it mid-slide (PR #19 review 3). There is no poster to click, so no retry is offered.
    const refuse = (e: unknown) => {
      diagnose(p.slide, 'media-autoplay', 'transition-sound', `The browser refused the transition sound (${String((e as { name?: unknown })?.name ?? e)}); the show continues`)
      if (playing === p) stop()
    }
    try {
      r = el.play() as Promise<void> | undefined
    } catch (e) {
      refuse(e)
      return
    }
    if (r && typeof r.then === 'function') {
      r.then(undefined, (e) => {
        if ((e as { name?: unknown })?.name === 'AbortError') return
        refuse(e)
      })
    }
  }

  return {
    start(index) {
      const slide = deps.slide(index)
      const sound = slide?.transition?.sound
      if (!slide || !sound) return
      // p:endSnd, or a new sound: the one still playing stops here.
      if (sound.end || sound.start) stop()
      const want = sound.start
      if (!want) return
      const rec = slide.media.find((m) => m.rId === want.rId)
      if (!rec) {
        diagnose(index, 'media-missing', 'transition-sound', 'The transition sound has no media part; the transition runs silent')
        return
      }
      const el = create('audio')
      if (el.canPlayType(rec.contentType) === '') {
        diagnose(index, 'media-codec', 'transition-sound', `The browser cannot decode the transition sound (${rec.contentType}); the transition runs silent`)
        return
      }
      const p: Playing = { slide: index, loop: want.loop, abort: new AbortController() }
      playing = p
      deps.open(rec, p.abort.signal).then((res) => {
        if ('refused' in res) {
          if (playing === p) playing = undefined
          diagnose(index, 'media-refused', 'transition-sound', `The transition sound was not read (${res.refused}); the transition runs silent`)
          return
        }
        if (playing !== p) { res.release(); return }
        if (p.loop) res.retain(Infinity)
        p.lease = res
        p.el = el
        el.setAttribute('data-officeview-show-sound', '')
        el.loop = p.loop
        el.src = res.url
        el.hidden = true
        deps.root.append(el)
        el.addEventListener('ended', () => { if (playing === p && !p.loop) stop() })
        play(p)
      }, () => { if (playing === p) playing = undefined })
    },
    stopAt(index) {
      void index
      // Everything but a looping sound stops on slide change.
      if (playing && !playing.loop) stop()
    },
    setPaused(on) {
      if (on === paused) return
      paused = on
      const p = playing
      if (!p?.el) return
      if (on) p.el.pause()
      else play(p)
    },
    release(onlyLooping = false) {
      if (onlyLooping && !playing?.loop) return false
      return stop()
    },
    dispose() { stop() },
  }
}
