# Slideshow validation

Validation record for the PowerPoint slideshow project (spec
`docs/superpowers/specs/2026-10-21-slideshow-design.md`, plan
`docs/superpowers/plans/2026-10-21-slideshow.md`). Created by task S0-1 as a stub;
later tasks fill in their sections. The controller merges the per-task ledger files
(`.superpowers/sdd/ledger/<task-id>.md`) into the Eager ledger and Catalogue sections
in merge order; implementer tasks do not edit those two sections except S1-3's explicitly assigned initial generated catalogue.

Synthetic builder decks are "Office-shaped synthetic" and are never reported as
Office evidence.

## Gates

Pending (filled at each stage close).

## Eager ledger

Measured with `bun run build && bun scripts/eager-budget.ts` (gzip of `dist/index.js`)
on fresh rebuilds only; rebuild variance is about 2 B. Ceiling 38,647 B, hard limit
38,720 B.

| Task | Item | Before | After | Delta | Running total |
|---|---|---|---|---|---|
| S0-1 | Baseline (main `b8a146c`, fresh rebuild on `965ad4e`, no `src` change since `b8a146c`) | | 38,584 | | 38,584 |
| S0-1 | `T0`: throwaway stubs of S0-4, S0-6 and S0-7's eager edits, measured then reverted | 38,584 | 38,634 | +50 | (not landed) |
| S0-1 | Recovery: delete 11 navigation-banner comments in `src/rtf/parse.ts` that the eager bundle kept verbatim | 38,584 | 38,453 | −131 | 38,453 |
| S0-1 | Check: the `T0` stubs re-applied on the recovered tree (then reverted) | 38,453 | 38,503 | +50 | (not landed) |

`.superpowers/sdd/ledger/S0-1.md` holds the full S0-1 rows, including the `T0` stub breakdown.

## Catalogue

<!-- slideshow-catalogue:start -->
Stage contract, not an implementation-completion claim. Every row is unverified; no Office frame compare has landed.
Animation names retain the LibreOffice reference spelling where the Microsoft UI correspondence is uncertain (see source audit below).

| ID | Name | Implementation | Status | Stage | Verified | Lookalike / stand-in | Mean abs | Row max |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| transition:p:blinds | blinds | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:checker | checker | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:circle | Shape (circle) | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:comb | comb | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:cover | cover | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:cut | cut | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:diamond | Shape (diamond) | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:dissolve | dissolve | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:fade | fade | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:newsflash | newsflash | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:plus | Shape (plus) | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:pull | Uncover | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:push | push | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:random | random | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:randomBar | randomBar | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:split | split | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:strips | strips | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:wedge | wedge | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:wheel | Clock | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:wipe | wipe | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:zoom | zoom | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:vortex | vortex | stand-in | pending-stage-3 | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:switch | switch | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:flip | flip | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:ripple | ripple | stand-in | pending-stage-3 | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:honeycomb | honeycomb | stand-in | pending-stage-3 | 3 | no | transition:p:checker | 0.08 | 0.3 |
| transition:p14:prism | Prism (Cube / Box / Rotate variants; D4 unverified) | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:doors | doors | stand-in | pending-stage-3 | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p14:window | window | stand-in | pending-stage-3 | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p14:ferris | ferris | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:gallery | gallery | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:conveyor | conveyor | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:pan | pan | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:glitter | glitter | stand-in | pending-stage-3 | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:warp | warp | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:flythrough | flythrough | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:flash | flash | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:shred | shred | stand-in | pending-stage-3 | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:reveal | reveal | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:wheelReverse | wheelReverse | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p15:prstTrans:fallOver | fallOver | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:drape | drape | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:curtains | curtains | stand-in | pending-stage-3 | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p15:prstTrans:wind | wind | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:prestige | prestige | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:fracture | fracture | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:crush | crush | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:peelOff | peelOff | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:pageCurlDouble | pageCurlDouble | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:pageCurlSingle | pageCurlSingle | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:airplane | airplane | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:origami | origami | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph | Morph (default) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byObject | Morph (byObject) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byWord | Morph (byWord) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byChar | Morph (byChar) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| behaviour:p:set | set | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:anim | anim | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animEffect | animEffect | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animMotion | animMotion | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animRot | animRot | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animScale | animScale | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animClr | animClr | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:cmd | cmd | stand-in | pending-stage-2 | 2 | no | silent timed no-op | 0.08 | 0.3 |
| behaviour:p:audio | audio | stand-in | pending-stage-2 | 2 | no | silent timed no-op | 0.08 | 0.3 |
| behaviour:p:video | video | stand-in | pending-stage-2 | 2 | no | poster frame | 0.08 | 0.3 |
| preset:entr:1 | appear | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2 | fly-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3 | venetian-blinds | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4 | box | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5 | checkerboard | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6 | circle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7 | fly-in-slow | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8 | diamond | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9 | dissolve-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10 | fade-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11 | flash-once | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12 | peek-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13 | plus | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14 | random-bars | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15 | spiral-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16 | split | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17 | stretchy | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18 | diagonal-squares | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19 | swivel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20 | wedge | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:21 | wheel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22 | wipe | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23 | zoom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24 | random | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25 | boomerang | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26 | bounce | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27 | colored-lettering | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28 | movie-credits | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29 | ease-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30 | float | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31 | turn-and-grow | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34 | breaks | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35 | pinwheel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37 | rise-up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38 | falling-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39 | thread | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40 | unfold | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41 | whip | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42 | ascend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43 | center-revolve | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45 | fade-in-and-swivel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47 | descend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48 | sling | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49 | spin-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50 | compress | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51 | magnify | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52 | curve-up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53 | fade-in-and-zoom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54 | glide | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55 | expand | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56 | flip | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58 | fold | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:1 | fill-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:2 | font | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:3 | font-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:4 | font-size | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:5 | font-style | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:6 | grow-and-shrink | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:7 | line-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:8 | spin | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:9 | transparency | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:10 | bold-flash | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:14 | blast | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:15 | bold-reveal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:16 | color-over-by-word | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:18 | reveal-underline | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:19 | color-blend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:20 | color-over-by-letter | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:21 | complementary-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:22 | complementary-color-2 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:23 | contrasting-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:24 | darken | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:25 | desaturate | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:26 | flash-bulb | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:27 | flicker | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:28 | grow-with-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:30 | lighten | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:31 | style-emphasis | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:32 | teeter | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:33 | vertical-highlight | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:34 | wave | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:35 | blink | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:36 | shimmer | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1 | disappear | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2 | fly-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3 | venetian-blinds | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4 | box | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5 | checkerboard | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6 | circle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7 | crawl-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8 | diamond | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9 | dissolve | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10 | fade-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11 | flash-once | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12 | peek-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13 | plus | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14 | random-bars | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15 | spiral-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16 | split | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17 | collapse | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18 | diagonal-squares | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19 | swivel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20 | wedge | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:21 | wheel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22 | wipe | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23 | zoom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24 | random | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25 | boomerang | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26 | bounce | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27 | colored-lettering | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28 | movie-credits | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29 | ease-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30 | float | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31 | turn-and-grow | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34 | breaks | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35 | pinwheel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37 | sink-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38 | swish | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39 | thread | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40 | unfold | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41 | whip | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42 | descend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43 | center-revolve | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45 | fade-out-and-swivel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47 | ascend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48 | sling | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53 | fade-out-and-zoom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55 | contract | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49 | spin-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50 | stretchy | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51 | magnify | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52 | curve-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54 | glide | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56 | flip | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58 | fold | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:16 | 4-point-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:5 | 5-point-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:11 | 6-point-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:17 | 8-point-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:1 | circle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:6 | crescent-moon | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:3 | diamond | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:13 | equal-triangle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:12 | oval | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:9 | heart | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:4 | hexagon | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:10 | octagon | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:14 | parallelogram | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:15 | pentagon | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:2 | right-triangle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:7 | square | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:18 | teardrop | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:8 | trapezoid | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:37 | arc-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:51 | arc-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:58 | arc-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:44 | arc-up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:41 | bounce-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:54 | bounce-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:48 | curvy-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:61 | curvy-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:60 | decaying-wave | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:49 | diagonal-down-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:56 | diagonal-up-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:42 | down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:52 | funnel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:53 | spring | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:62 | stairs-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:50 | turn-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:36 | turn-down-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:43 | turn-up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:57 | turn-up-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:64 | up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:47 | wave | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:38 | zigzag | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:31 | bean | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:25 | buzz-saw | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:20 | curved-square | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:21 | curved-x | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:23 | curvy-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:28 | figure-8-four | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:26 | horizontal-figure-8 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:34 | inverted-square | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:33 | inverted-triangle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:24 | loop-de-loop | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:29 | neutron | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:27 | peanut | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:32 | clover | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:19 | pointy-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:30 | swoosh | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:22 | vertical-figure-8 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:35 | left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:63 | right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:55 | spiral-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:46 | spiral-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:40 | sine-wave | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:59 | s-curve-1 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:39 | s-curve-2 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:45 | heartbeat | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:1 | fly-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:2 | fly-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:3 | fly-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:4 | fly-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:5 | fly-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:6 | fly-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:8 | fly-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:9 | fly-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:10 | fly-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:12 | fly-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:16 | fly-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:21 | fly-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:26 | fly-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:32 | fly-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:36 | fly-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:37 | fly-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:42 | fly-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:272 | fly-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:288 | fly-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:528 | fly-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:1 | fly-out: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:2 | fly-out: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:3 | fly-out: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:4 | fly-out: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:5 | fly-out: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:6 | fly-out: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:8 | fly-out: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:9 | fly-out: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:10 | fly-out: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:12 | fly-out: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:16 | fly-out: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:21 | fly-out: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:26 | fly-out: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:32 | fly-out: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:36 | fly-out: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:37 | fly-out: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:42 | fly-out: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:272 | fly-out: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:288 | fly-out: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:528 | fly-out: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| build:word | Word build (paragraph stand-in) | stand-in | pending-stage-2 | 2 | no | whole paragraph | 0.08 | 0.3 |
| build:letter | Letter build (paragraph stand-in) | stand-in | pending-stage-2 | 2 | no | whole paragraph | 0.08 | 0.3 |
| sound:transition | Transition / cross-slide sound | stand-in | pending-stage-2 | 2 | no | silent timed no-op | 0.08 | 0.3 |
| build:chart | Chart build (whole chart) | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| build:smartart | SmartArt build (fallback drawing) | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| build:fallback | Fallback picture / card animation | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| media:codec | Unavailable media codec (poster / silent no-op) | stand-in | approximate | 2 | no | — | 0.08 | 0.3 |
| never:macro | VBA / program execution | stand-in | never | 1 | no | — | 0.08 | 0.3 |
| never:ole | OLE activation | stand-in | never | 1 | no | — | 0.08 | 0.3 |
| never:recorded | Recorded ink / narration beyond stored timing | stand-in | never | 1 | no | — | 0.08 | 0.3 |

| Precedence question | State |
| --- | --- |
| trigger beats hyperlink | pending D4 |
| kiosk link inertness | pending D4 |
<!-- slideshow-catalogue:end -->

## Fixtures

Pending (S0-2: LibreOffice corpus, pinned commit SHA and provenance).

## Host requirements

Pending.

## Known limits

- **Slide size range (S0-4).** `validateSlideSize` accepts 914,400–51,206,400 EMU per side (1–56 in), both ends inclusive, finite and positive, with an aspect ratio of at most 1 : 20. The range is ECMA-376 Part 1 pml.xsd `ST_SlideSizeCoordinate` (`xsd:int`, `minInclusive 914400`, `maxInclusive 51206400`).
  - Microsoft's [SlideSize.Cx reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.slidesize.cx?view=openxml-3.0.1) confirms that `p:sldSz@cx` is typed `ST_SlideSizeCoordinate`, an integer. It does not state the facets.
  - The numbers were checked against python-pptx `oxml/simpletypes.py` `ST_SlideSizeCoordinate`, which raises outside `914400..51206400`, "(1-56 inches)".
  - **pml.xsd itself has not been checked.**
- **Package lease metric is a payload lower bound (S0-4).** `retainedArchiveBytes(handle)` follows the plan's formula: the sum of compressed entry sizes. JSZip keeps each entry's compressed bytes as a view into the whole input buffer, so the memory actually retained is at least the full archive, including local headers, the central directory and any comment or trailer. It can be much larger than the metric. `tests/pptx-appearance-read.test.ts` pins this with a 16 KiB comment. Corpus figures are pending, because the S0-2 fixtures are not on this branch. Parse-time read state is not retained: `parsePptx` drops the package read cache and the drawing parser's package context when it attaches the handle.

## Manual checks

Pending.

## S1-3 inventory provenance and name audit

This task's explicit generated-table requirement is the scoped exception to the
shared catalogue ownership rule above. The eager ledger is untouched. The table
is stage policy; it does not claim that the later rendering tasks have landed.

- **Animation source (controller ruling, 2026-10-21):** LibreOffice core
  [`oox/source/ppt/commontimenodecontext.cxx`](https://github.com/LibreOffice/core/blob/2578da962608a2123d637b80499123c258829582/oox/source/ppt/commontimenodecontext.cxx),
  commit `2578da962608a2123d637b80499123c258829582`, MPL-2.0.
  Saved source SHA256 `57efab6164d1a8d6852d58c40370c01cadca24b508f00e00f61e616547f391c4`.
  Only the factual keys and effect names from `preset_mapping::getList()` and
  `convert_subtype::getList()` were transcribed; no implementation code was copied.
  The inventory has 199 base families (52 entrance, 52 exit, 31 emphasis, 64 paths),
  plus 40 explicit fly-direction tuples (entr/exit id 2 × 20 source directions).
  The direction table is exported separately as `PRESET_DIRECTIONS`.
  **This is not a claim to a complete Microsoft legal subtype cross-product.**
  The source does not enumerate that cross-product. No directions were invented
  for other families. Unknown tuples return undefined for the reader to diagnose;
  a supplied subtype never silently falls back to an id-only row.
- **Transitions:** ECMA-376 Part 1, fifth edition, December 2016,
  `CT_SlideTransition` ([edition archive](https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip));
  the 21 schema children are the spec §6.4 list. Extension names checked against
  [MS-PPTX v20240820](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-PPTX/%5BMS-PPTX%5D.pdf),
  §5.1 (19 p14 transitions), §2.4.3.8 (12 `prstTrans` presets), §5.4
  (`morph`, options byObject/byWord/byChar). Morph's default and three explicit
  options have separate lookup keys; the optional `preset` lookup argument carries
  the Morph option. Counts total 56 transition identities.
  Cube/Box/Rotate are UI variants of prism, not invented XML child elements.
  Orbit has no independent element in these pinned lists; its precise UI-to-schema
  correspondence is unverified and is not invented as a new key.
- **Behaviours and policy-only rows:** project spec §6.1 (10 behaviour elements),
  §6.5 (best effort), §2.3 (never), stage table §12. The word/letter rows are
  pending Stage 2; shaders/Morph pending Stage 3. No tolerance overrides are used.
- **Names cross-check:** [Microsoft MsoAnimEffect](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect),
  page revision 2021-09-13, retrieved 2026-10-21. The check compares spelling,
  not VBA numeric values: e.g. LO path circle id 1 versus VBA PathCircle 86;
  LO emphasis spin id 8 versus VBA Spin 61. These are different identifier spaces.
  Common naming differences include LO venetian-blinds vs Blinds, fly-in/out vs Fly,
  fade-in/out vs Fade, movie-credits vs Credits. The table retains LO reference
  spelling. These likely correspondences are not promoted to verified Office UI
  aliases or rendering equivalence. Notable unresolved names include LO path oval
  vs VBA PathFootball, clover vs PathPlus, turn-down-right vs PathTurnRight, and
  entrance fly-in-slow vs VBA Crawl. No numeric keys were changed to match VBA.

The following audit covers all 199 source families. “No exact name match” means
only that spelling does not match; it does not assert that PowerPoint lacks the
feature. Exact spelling is likewise not proof of identical rendering. All frame
comparisons remain unverified. The base family names in `uiName` are reference
labels pending D4, rather than certified localized PowerPoint menu text.

| LO class | ID | LO reference name | Microsoft enum spelling check |
| --- | --- | --- | --- |
| ENTRANCE | 1 | appear | `msoAnimEffectAppear` |
| ENTRANCE | 2 | fly-in | No exact name match; unresolved |
| ENTRANCE | 3 | venetian-blinds | No exact name match; unresolved |
| ENTRANCE | 4 | box | `msoAnimEffectBox` |
| ENTRANCE | 5 | checkerboard | `msoAnimEffectCheckerboard` |
| ENTRANCE | 6 | circle | `msoAnimEffectCircle` |
| ENTRANCE | 7 | fly-in-slow | No exact name match; unresolved |
| ENTRANCE | 8 | diamond | `msoAnimEffectDiamond` |
| ENTRANCE | 9 | dissolve-in | No exact name match; unresolved |
| ENTRANCE | 10 | fade-in | No exact name match; unresolved |
| ENTRANCE | 11 | flash-once | `msoAnimEffectFlashOnce` |
| ENTRANCE | 12 | peek-in | No exact name match; unresolved |
| ENTRANCE | 13 | plus | `msoAnimEffectPlus` |
| ENTRANCE | 14 | random-bars | `msoAnimEffectRandomBars` |
| ENTRANCE | 15 | spiral-in | No exact name match; unresolved |
| ENTRANCE | 16 | split | `msoAnimEffectSplit` |
| ENTRANCE | 17 | stretchy | `msoAnimEffectStretchy` |
| ENTRANCE | 18 | diagonal-squares | No exact name match; unresolved |
| ENTRANCE | 19 | swivel | `msoAnimEffectSwivel` |
| ENTRANCE | 20 | wedge | `msoAnimEffectWedge` |
| ENTRANCE | 21 | wheel | `msoAnimEffectWheel` |
| ENTRANCE | 22 | wipe | `msoAnimEffectWipe` |
| ENTRANCE | 23 | zoom | `msoAnimEffectZoom` |
| ENTRANCE | 24 | random | No exact name match; unresolved |
| ENTRANCE | 25 | boomerang | `msoAnimEffectBoomerang` |
| ENTRANCE | 26 | bounce | `msoAnimEffectBounce` |
| ENTRANCE | 27 | colored-lettering | No exact name match; unresolved |
| ENTRANCE | 28 | movie-credits | No exact name match; unresolved |
| ENTRANCE | 29 | ease-in | `msoAnimEffectEaseIn` |
| ENTRANCE | 30 | float | `msoAnimEffectFloat` |
| ENTRANCE | 31 | turn-and-grow | No exact name match; unresolved |
| ENTRANCE | 34 | breaks | No exact name match; unresolved |
| ENTRANCE | 35 | pinwheel | `msoAnimEffectPinwheel` |
| ENTRANCE | 37 | rise-up | `msoAnimEffectRiseUp` |
| ENTRANCE | 38 | falling-in | No exact name match; unresolved |
| ENTRANCE | 39 | thread | No exact name match; unresolved |
| ENTRANCE | 40 | unfold | `msoAnimEffectUnfold` |
| ENTRANCE | 41 | whip | `msoAnimEffectWhip` |
| ENTRANCE | 42 | ascend | `msoAnimEffectAscend` |
| ENTRANCE | 43 | center-revolve | `msoAnimEffectCenterRevolve` |
| ENTRANCE | 45 | fade-in-and-swivel | No exact name match; unresolved |
| ENTRANCE | 47 | descend | `msoAnimEffectDescend` |
| ENTRANCE | 48 | sling | `msoAnimEffectSling` |
| ENTRANCE | 49 | spin-in | No exact name match; unresolved |
| ENTRANCE | 50 | compress | No exact name match; unresolved |
| ENTRANCE | 51 | magnify | No exact name match; unresolved |
| ENTRANCE | 52 | curve-up | No exact name match; unresolved |
| ENTRANCE | 53 | fade-in-and-zoom | No exact name match; unresolved |
| ENTRANCE | 54 | glide | `msoAnimEffectGlide` |
| ENTRANCE | 55 | expand | `msoAnimEffectExpand` |
| ENTRANCE | 56 | flip | `msoAnimEffectFlip` |
| ENTRANCE | 58 | fold | `msoAnimEffectFold` |
| EMPHASIS | 1 | fill-color | No exact name match; unresolved |
| EMPHASIS | 2 | font | No exact name match; unresolved |
| EMPHASIS | 3 | font-color | No exact name match; unresolved |
| EMPHASIS | 4 | font-size | No exact name match; unresolved |
| EMPHASIS | 5 | font-style | No exact name match; unresolved |
| EMPHASIS | 6 | grow-and-shrink | No exact name match; unresolved |
| EMPHASIS | 7 | line-color | No exact name match; unresolved |
| EMPHASIS | 8 | spin | `msoAnimEffectSpin` |
| EMPHASIS | 9 | transparency | `msoAnimEffectTransparency` |
| EMPHASIS | 10 | bold-flash | `msoAnimEffectBoldFlash` |
| EMPHASIS | 14 | blast | `msoAnimEffectBlast` |
| EMPHASIS | 15 | bold-reveal | `msoAnimEffectBoldReveal` |
| EMPHASIS | 16 | color-over-by-word | No exact name match; unresolved |
| EMPHASIS | 18 | reveal-underline | No exact name match; unresolved |
| EMPHASIS | 19 | color-blend | `msoAnimEffectColorBlend` |
| EMPHASIS | 20 | color-over-by-letter | No exact name match; unresolved |
| EMPHASIS | 21 | complementary-color | `msoAnimEffectComplementaryColor` |
| EMPHASIS | 22 | complementary-color-2 | `msoAnimEffectComplementaryColor2` |
| EMPHASIS | 23 | contrasting-color | `msoAnimEffectContrastingColor` |
| EMPHASIS | 24 | darken | `msoAnimEffectDarken` |
| EMPHASIS | 25 | desaturate | `msoAnimEffectDesaturate` |
| EMPHASIS | 26 | flash-bulb | `msoAnimEffectFlashBulb` |
| EMPHASIS | 27 | flicker | `msoAnimEffectFlicker` |
| EMPHASIS | 28 | grow-with-color | `msoAnimEffectGrowWithColor` |
| EMPHASIS | 30 | lighten | `msoAnimEffectLighten` |
| EMPHASIS | 31 | style-emphasis | `msoAnimEffectStyleEmphasis` |
| EMPHASIS | 32 | teeter | `msoAnimEffectTeeter` |
| EMPHASIS | 33 | vertical-highlight | No exact name match; unresolved |
| EMPHASIS | 34 | wave | `msoAnimEffectWave` |
| EMPHASIS | 35 | blink | No exact name match; unresolved |
| EMPHASIS | 36 | shimmer | `msoAnimEffectShimmer` |
| EXIT | 1 | disappear | No exact name match; unresolved |
| EXIT | 2 | fly-out | No exact name match; unresolved |
| EXIT | 3 | venetian-blinds | No exact name match; unresolved |
| EXIT | 4 | box | `msoAnimEffectBox` |
| EXIT | 5 | checkerboard | `msoAnimEffectCheckerboard` |
| EXIT | 6 | circle | `msoAnimEffectCircle` |
| EXIT | 7 | crawl-out | No exact name match; unresolved |
| EXIT | 8 | diamond | `msoAnimEffectDiamond` |
| EXIT | 9 | dissolve | `msoAnimEffectDissolve` |
| EXIT | 10 | fade-out | No exact name match; unresolved |
| EXIT | 11 | flash-once | `msoAnimEffectFlashOnce` |
| EXIT | 12 | peek-out | No exact name match; unresolved |
| EXIT | 13 | plus | `msoAnimEffectPlus` |
| EXIT | 14 | random-bars | `msoAnimEffectRandomBars` |
| EXIT | 15 | spiral-out | No exact name match; unresolved |
| EXIT | 16 | split | `msoAnimEffectSplit` |
| EXIT | 17 | collapse | No exact name match; unresolved |
| EXIT | 18 | diagonal-squares | No exact name match; unresolved |
| EXIT | 19 | swivel | `msoAnimEffectSwivel` |
| EXIT | 20 | wedge | `msoAnimEffectWedge` |
| EXIT | 21 | wheel | `msoAnimEffectWheel` |
| EXIT | 22 | wipe | `msoAnimEffectWipe` |
| EXIT | 23 | zoom | `msoAnimEffectZoom` |
| EXIT | 24 | random | No exact name match; unresolved |
| EXIT | 25 | boomerang | `msoAnimEffectBoomerang` |
| EXIT | 26 | bounce | `msoAnimEffectBounce` |
| EXIT | 27 | colored-lettering | No exact name match; unresolved |
| EXIT | 28 | movie-credits | No exact name match; unresolved |
| EXIT | 29 | ease-out | No exact name match; unresolved |
| EXIT | 30 | float | `msoAnimEffectFloat` |
| EXIT | 31 | turn-and-grow | No exact name match; unresolved |
| EXIT | 34 | breaks | No exact name match; unresolved |
| EXIT | 35 | pinwheel | `msoAnimEffectPinwheel` |
| EXIT | 37 | sink-down | No exact name match; unresolved |
| EXIT | 38 | swish | `msoAnimEffectSwish` |
| EXIT | 39 | thread | No exact name match; unresolved |
| EXIT | 40 | unfold | `msoAnimEffectUnfold` |
| EXIT | 41 | whip | `msoAnimEffectWhip` |
| EXIT | 42 | descend | `msoAnimEffectDescend` |
| EXIT | 43 | center-revolve | `msoAnimEffectCenterRevolve` |
| EXIT | 45 | fade-out-and-swivel | No exact name match; unresolved |
| EXIT | 47 | ascend | `msoAnimEffectAscend` |
| EXIT | 48 | sling | `msoAnimEffectSling` |
| EXIT | 53 | fade-out-and-zoom | No exact name match; unresolved |
| EXIT | 55 | contract | No exact name match; unresolved |
| EXIT | 49 | spin-out | No exact name match; unresolved |
| EXIT | 50 | stretchy | `msoAnimEffectStretchy` |
| EXIT | 51 | magnify | No exact name match; unresolved |
| EXIT | 52 | curve-down | No exact name match; unresolved |
| EXIT | 54 | glide | `msoAnimEffectGlide` |
| EXIT | 56 | flip | `msoAnimEffectFlip` |
| EXIT | 58 | fold | `msoAnimEffectFold` |
| MOTIONPATH | 16 | 4-point-star | `msoAnimEffectPath4PointStar` |
| MOTIONPATH | 5 | 5-point-star | `msoAnimEffectPath5PointStar` |
| MOTIONPATH | 11 | 6-point-star | `msoAnimEffectPath6PointStar` |
| MOTIONPATH | 17 | 8-point-star | `msoAnimEffectPath8PointStar` |
| MOTIONPATH | 1 | circle | `msoAnimEffectPathCircle` |
| MOTIONPATH | 6 | crescent-moon | `msoAnimEffectPathCrescentMoon` |
| MOTIONPATH | 3 | diamond | `msoAnimEffectPathDiamond` |
| MOTIONPATH | 13 | equal-triangle | `msoAnimEffectPathEqualTriangle` |
| MOTIONPATH | 12 | oval | No exact name match; unresolved |
| MOTIONPATH | 9 | heart | `msoAnimEffectPathHeart` |
| MOTIONPATH | 4 | hexagon | `msoAnimEffectPathHexagon` |
| MOTIONPATH | 10 | octagon | `msoAnimEffectPathOctagon` |
| MOTIONPATH | 14 | parallelogram | `msoAnimEffectPathParallelogram` |
| MOTIONPATH | 15 | pentagon | `msoAnimEffectPathPentagon` |
| MOTIONPATH | 2 | right-triangle | `msoAnimEffectPathRightTriangle` |
| MOTIONPATH | 7 | square | `msoAnimEffectPathSquare` |
| MOTIONPATH | 18 | teardrop | `msoAnimEffectPathTeardrop` |
| MOTIONPATH | 8 | trapezoid | `msoAnimEffectPathTrapezoid` |
| MOTIONPATH | 37 | arc-down | `msoAnimEffectPathArcDown` |
| MOTIONPATH | 51 | arc-left | `msoAnimEffectPathArcLeft` |
| MOTIONPATH | 58 | arc-right | `msoAnimEffectPathArcRight` |
| MOTIONPATH | 44 | arc-up | `msoAnimEffectPathArcUp` |
| MOTIONPATH | 41 | bounce-left | `msoAnimEffectPathBounceLeft` |
| MOTIONPATH | 54 | bounce-right | `msoAnimEffectPathBounceRight` |
| MOTIONPATH | 48 | curvy-left | `msoAnimEffectPathCurvyLeft` |
| MOTIONPATH | 61 | curvy-right | `msoAnimEffectPathCurvyRight` |
| MOTIONPATH | 60 | decaying-wave | `msoAnimEffectPathDecayingWave` |
| MOTIONPATH | 49 | diagonal-down-right | `msoAnimEffectPathDiagonalDownRight` |
| MOTIONPATH | 56 | diagonal-up-right | `msoAnimEffectPathDiagonalUpRight` |
| MOTIONPATH | 42 | down | `msoAnimEffectPathDown` |
| MOTIONPATH | 52 | funnel | `msoAnimEffectPathFunnel` |
| MOTIONPATH | 53 | spring | `msoAnimEffectPathSpring` |
| MOTIONPATH | 62 | stairs-down | `msoAnimEffectPathStairsDown` |
| MOTIONPATH | 50 | turn-down | `msoAnimEffectPathTurnDown` |
| MOTIONPATH | 36 | turn-down-right | No exact name match; unresolved |
| MOTIONPATH | 43 | turn-up | `msoAnimEffectPathTurnUp` |
| MOTIONPATH | 57 | turn-up-right | `msoAnimEffectPathTurnUpRight` |
| MOTIONPATH | 64 | up | `msoAnimEffectPathUp` |
| MOTIONPATH | 47 | wave | `msoAnimEffectPathWave` |
| MOTIONPATH | 38 | zigzag | `msoAnimEffectPathZigzag` |
| MOTIONPATH | 31 | bean | `msoAnimEffectPathBean` |
| MOTIONPATH | 25 | buzz-saw | `msoAnimEffectPathBuzzsaw` |
| MOTIONPATH | 20 | curved-square | `msoAnimEffectPathCurvedSquare` |
| MOTIONPATH | 21 | curved-x | `msoAnimEffectPathCurvedX` |
| MOTIONPATH | 23 | curvy-star | `msoAnimEffectPathCurvyStar` |
| MOTIONPATH | 28 | figure-8-four | `msoAnimEffectPathFigure8Four` |
| MOTIONPATH | 26 | horizontal-figure-8 | `msoAnimEffectPathHorizontalFigure8` |
| MOTIONPATH | 34 | inverted-square | `msoAnimEffectPathInvertedSquare` |
| MOTIONPATH | 33 | inverted-triangle | `msoAnimEffectPathInvertedTriangle` |
| MOTIONPATH | 24 | loop-de-loop | `msoAnimEffectPathLoopdeLoop` |
| MOTIONPATH | 29 | neutron | `msoAnimEffectPathNeutron` |
| MOTIONPATH | 27 | peanut | `msoAnimEffectPathPeanut` |
| MOTIONPATH | 32 | clover | No exact name match; unresolved |
| MOTIONPATH | 19 | pointy-star | `msoAnimEffectPathPointyStar` |
| MOTIONPATH | 30 | swoosh | `msoAnimEffectPathSwoosh` |
| MOTIONPATH | 22 | vertical-figure-8 | `msoAnimEffectPathVerticalFigure8` |
| MOTIONPATH | 35 | left | `msoAnimEffectPathLeft` |
| MOTIONPATH | 63 | right | `msoAnimEffectPathRight` |
| MOTIONPATH | 55 | spiral-left | `msoAnimEffectPathSpiralLeft` |
| MOTIONPATH | 46 | spiral-right | `msoAnimEffectPathSpiralRight` |
| MOTIONPATH | 40 | sine-wave | `msoAnimEffectPathSineWave` |
| MOTIONPATH | 59 | s-curve-1 | `msoAnimEffectPathSCurve1` |
| MOTIONPATH | 39 | s-curve-2 | `msoAnimEffectPathSCurve2` |
| MOTIONPATH | 45 | heartbeat | `msoAnimEffectPathHeartbeat` |

Audit totals: 140 exact normalized-name matches; 59 unresolved name matches. Hyphens/case and the LO motionpath prefix were normalized only.
