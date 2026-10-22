// The D4 behavioural questions (S1-21): every behaviour the spec or a task
// report leaves "to verify against D4 captures" or "unverified". Each id has a
// section in docs/validation/slideshow-captures.md with fixed fields (Open, Do,
// Record, Engine assumption, Verifies, Answer); tests/slideshow-compare.test.ts
// checks that every id here has one.
//
// kind: 'corpus' = answered from a corpus deck; 'built' = needs a small deck
// the owner builds in PowerPoint (steps in the checklist); 'standard' = not
// visible in frames, answered by reading ECMA-376 (or a hand-edited deck).

export interface CaptureQuestion {
  id: string
  title: string
  kind: 'corpus' | 'built' | 'standard'
  /** Where the question was left open. */
  source: string
}

const q = (id: string, kind: CaptureQuestion['kind'], title: string, source: string): CaptureQuestion => Object.freeze({ id, kind, title, source })

export const BEHAVIOURAL_QUESTIONS: readonly CaptureQuestion[] = Object.freeze([
  // The seven named in the S1-21 brief and spec section 11 D4.
  q('Q-spd', 'corpus', 'How long are spd="fast", "med" and "slow"?', 'spec 6.1; S1-8 SPD_MS; validation doc spd table'),
  q('Q-advtm-pending', 'built', 'Does advTm wait while on-click steps are still unplayed?', 'spec 4.4; S1-10 report'),
  q('Q-trigger-link', 'built', 'A shape with both a trigger and a hyperlink: which runs?', 'spec 4.3; validation doc "trigger beats hyperlink | pending D4"'),
  q('Q-kiosk-restart', 'built', 'When does a kiosk show restart, and what does restart="0" do?', 'spec 4.4; S1-10 report (restart 0 refused)'),
  q('Q-kiosk-input', 'built', 'In a kiosk show, do hyperlinks and triggers respond?', 'spec 4.4; validation doc "kiosk link inertness | pending D4"'),
  q('Q-previous-into', 'built', 'Previous from a slide\'s start: is the previous slide shown fully built?', 'spec 4.3, 6.2 (revisiting a slide)'),
  q('Q-med-700', 'built', 'spd="med" beside p14:dur="700": which duration plays?', 'S1-8 report; validation doc spd table'),
  // Auto-advance details (spec 4.4 working assumption; S1-10 "auto-advance rule is unverified").
  q('Q-advtm-start', 'built', 'Does the advTm timer start when the incoming transition ends?', 'spec 4.4; S1-10 report'),
  q('Q-advtm-busy', 'built', 'Does a running finite effect hold the advance past advTm?', 'spec 4.4; S1-10 report'),
  q('Q-advtm-indefinite', 'built', 'Do indefinite repeats, trigger sequences and media hold the advance?', 'spec 4.4; S1-10 report'),
  q('Q-advtm-zero', 'built', 'What does advTm="0" do?', 'spec 4.4'),
  q('Q-advclick', 'corpus', 'With advClick="0", do mouse clicks and keys advance?', 'spec 4.4; tdf142915'),
  // Navigation (spec 6.2; S1-10).
  q('Q-next-during', 'corpus', 'Next during a running effect: does the step jump to its end?', 'spec 6.2; S1-9 item 7'),
  q('Q-previous-trigger', 'built', 'Previous after a trigger fired at step k: checkpoint k-1 or k?', 'S1-10 report'),
  q('Q-showanimation', 'corpus', 'What does "Show without animation" (showAnimation="0") display?', 'spec 4.4, 6.2 static state'),
  // Transitions (S1-3 classifications, S1-15 defaults and look-alikes, S1-8 inheritance).
  q('Q-transition-defaults', 'built', 'Default direction of Pan, Zoom (Box) and Fly Through', 'S1-15 report and review'),
  q('Q-transition-visuals', 'built', 'What do zoom, newsflash, fly through, reveal, flash and dissolve look like?', 'S1-15 report'),
  q('Q-transition-classes', 'built', 'Which XML does each disputed transition write, and how does it look?', 'S1-3 review M1; catalogue "D4 classification unverified" rows; Prism, Orbit'),
  q('Q-transition-inherit', 'standard', 'Is a transition on a layout or master inherited by its slides?', 'spec 4.4; S1-8 report; validation doc'),
  q('Q-transition-extlst', 'standard', 'Does PowerPoint read a transition placed in p:transition/p:extLst?', 'S1-8 report; validation doc'),
  // Timing attributes observable in a built deck (S1-9 "to verify (D4 capture)").
  q('Q-restart-default', 'built', 'An effect re-triggered while running: does it restart (absent restart)?', 'S1-9 item 1'),
  q('Q-autorev', 'built', 'Auto-reverse with repeat: how long, and how is the ease applied?', 'S1-9 item 4'),
  q('Q-accel-decel', 'built', 'Smooth start + smooth end longer than the effect: what happens?', 'S1-9 item 5'),
  q('Q-repeat-fraction', 'built', 'Is repeatCount in thousandths, and does 2.5 stop mid-iteration?', 'S1-8 report; validation doc repeatCount row'),
  q('Q-fill-absent', 'built', 'Does an effect keep its end state when its behaviours have no fill?', 'S1-9 items 2 and 3'),
  // Catalogue labels and appearance.
  q('Q-wipe-direction', 'corpus', 'animEffect wipe(up): which edge does the reveal start from?', 'S1-11 report (wipe motion-direction to reveal-edge inversion; subtype 8/2/1/4 table)'),
  q('Q-blinds-subtype', 'corpus', 'tdf93868: are the presetSubtype 10 blinds horizontal or vertical?', 'catalogue label "venetian-blinds: vertical" vs filter blinds(horizontal)'),
  q('Q-usebgfill', 'built', 'A shape with "slide background fill" over a picture background', 'S0-7 report (useBgFill precedence unverified)'),
  // Not visible in frames: answered from the standard text.
  q('Q-fmla-mod', 'standard', 'What does % mean in an animation formula?', 'S1-5 report'),
  q('Q-excl-pause', 'standard', 'Does excl pause or stop its other children?', 'S1-9 items 6 and 13'),
  q('Q-onclick-target', 'standard', 'onClick with no target or a slide target', 'S1-9 item 8'),
  q('Q-endsync-dur', 'standard', 'Does endSync apply when dur is set (including indefinite)?', 'S1-9 item 9'),
  q('Q-end-at-begin', 'standard', 'Does an end condition at the begin instant end the node?', 'S1-9 item 10'),
  q('Q-implicit-par-end', 'standard', 'Implicit par end with unresolved children', 'S1-9 items 11 and 12'),
  q('Q-additive-attrs', 'standard', 'cBhvr additive default and the attrName list', 'S1-8 report; validation doc'),
])
