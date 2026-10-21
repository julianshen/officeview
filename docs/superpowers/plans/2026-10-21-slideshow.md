# PowerPoint Slideshow Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Present a PPTX deck full screen in officeview, with PowerPoint's transitions, animations, navigation, show settings, Presenter View and media. The work ships in four staged PRs, and Stage 0 improves the static viewer's slide backgrounds and master artwork along the way.

**Architecture:**
1. **Stage 0.** A lazy slide-appearance enrichment (`src/pptx/appearance/`) is loaded from `getPaintables` and paints backgrounds and master/layout artwork for every consumer. `parsePptx` keeps a package handle in a WeakMap side table.
2. **Stage 1.** A lazy slideshow chunk (`src/slideshow/`) holds:
   - a bounded show reader over that handle;
   - an event-driven timing scheduler (a pure function of tree, event history and time);
   - a z-order-preserving render plan rasterised by the unchanged `paintScene`;
   - a `Compositor` interface with WebGL2 and Canvas 2D backends.
   The only eager additions are host plumbing and a synchronous start sequence with a per-document show coordinator, paid for by recovered bytes.
3. **Stages 2 and 3.** Stage 2 adds Presenter View and media. Stage 3 adds 3-D transitions and Morph, without touching eager code.

**Tech Stack:** TypeScript, React 19, Vite 7, vitest + jsdom + node-canvas, Playwright (Chromium), bun 1.3.10, WebGL2 (no new dependency).

**Status:** revision 2, after plan review round 1 (codex, agy and opencode requested changes; pi approved). Dispositions are in the appendix at the end.

**Spec:** `docs/superpowers/specs/2026-10-21-slideshow-design.md`, at `b693077`: approved by the owner, review rounds 1 and 2 folded in, agent approvals recorded. It is the binding authority, and every task names the sections it implements. Where the spec says "to verify in the parser task", the owning task here is named and records the verified value in the validation doc.

## Global Constraints

From the spec (values verbatim):

- **Eager budget.** The project ceiling stays **38,647 B** (hard limit 38,720 B). The baseline is **38,584 B** (main `b8a146c`, fresh rebuild; 63 B headroom).
  - "No eager growth is allowed without first recovering at least as many bytes elsewhere." Recovery candidates "must not remove public exports". Only a fresh rebuild counts (there is 2 B of rebuild variance).
  - Measure with `bun run build && bun scripts/eager-budget.ts`.
  - Check the eager graph with `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts`.
  - The controller re-measures on main before each stage's first eager edit. Every task records before/after/running total in the ledger (`docs/validation/slideshow.md` § Eager ledger).
- **Slideshow parsing does not happen in the eager `parsePptx`.** What stays eager, and only this:
  - the package handle;
  - the `import()` seams;
  - `forwardRef`/handle plumbing, the pre-mounted host element, the per-document show coordinator, the synchronous start sequence and F5 key routing.

  The Present button lives in the lazy `PagedChrome`. New eager files must not live under `src/components/paged/`, because `EAGER_FORBIDDEN` forbids everything there except `metrics.ts`.
- **Lazy directories:** `src/slideshow/` and `src/pptx/appearance/`. Both get `EAGER_FORBIDDEN` rules and both go on the `vite.harness.config.ts` warm list (S0-1).
- **Esc always exits**, in every mode, including kiosk. Never pop a window without a user gesture. Never trap focus.
- **Effects never silently do nothing.** Every unsupported, approximated or stand-in effect is diagnosed and catalogued. Diagnostics are deduplicated per effect and slide, never emitted per frame.
- **No `eval`, `Function` or any JavaScript evaluation** for formulas or paths. Formula limits: 1 KiB per expression, 256 tokens, depth 32.
- **Reader limits:**
  - 4 MiB inflated XML per slide part, 1 MiB per notes part, 1 MiB for `presProps`;
  - 2,000 time nodes per slide; 64 nesting levels; 10,000 conditions per slide;
  - 1,000 `p:tav` per behaviour, 20,000 per slide;
  - motion path ≤ 64 KiB and 10,000 commands;
  - 64 KiB notes text per slide;
  - 64 MiB total show metadata;
  - 512 MiB per media part.

  Lazy reads are uncached and bounded (`entry()` plus a byte cap), never `OfficePackage.text()`/`bytes()`.
- **Workspace and rendering:**
  - Workspace 512 MiB desktop / 256 MiB mobile.
  - Permanent reserve: one output surface plus a scratch allowance of 64 MiB desktop / 32 MiB mobile.
  - Backing store capped at DPR 2 **and** 3840 × 2160 device pixels; no layer above `MAX_TEXTURE_SIZE`.
  - At most 128 composited layers and 8 re-rasterisations per frame.
  - p95 frame work ≤ 12 ms on the benchmark deck.
- **Timing and slide size:**
  - Context-restore deadline 3 s, on `performance.now()`.
  - Reduced motion: crossfade ≤ 200 ms.
  - Slide size valid only when finite, positive, within 914,400–51,206,400 EMU, and with aspect within 1 : 20.

Project rules:

- **CI is Linux on 2 CPUs.** Every clock is injected and driven: show clock, rAF, restore deadline, cursor-hide timer, kiosk timer. No test waits on wall time.
- **Goldens are Mac-only**, run with `OFFICEVIEW_STRICT_GOLDEN=1 bunx vitest run tests/golden.test.ts`, never in CI. Every changed golden is listed in the commit and inspected individually.
- **Browser evidence is Chromium only.** WebKit cannot launch on the build host. Slideshow specs run under the `slideshow` Playwright project (created in S1-2) whose launch args enable SwiftShader WebGL2 (`--use-gl=angle --use-angle=swiftshader --enable-unsafe-swiftshader`). They assert behaviour on both the WebGL2 and a forced Canvas 2D path, never GPU pixels, and never frame-time numbers. Each run uses a unique `IV_HARNESS_PORT`.
- **CI must exercise the WebGL2 compositor via SwiftShader** (controller ruling). If CI cannot create a WebGL2 context, the stage PR is blocked and escalated to the owner; covering only Canvas 2D in CI is not allowed. GPU goldens are taken locally in Mac Chromium.
- **codex cannot bind localhost or write git metadata.** For codex tasks the controller runs the Playwright specs and commits.
- **No new runtime dependencies.**
- **Realm-safe DOM:** every listener, `requestAnimationFrame`, `matchMedia`, focus call and element creation goes through the target node's `ownerDocument`/`defaultView`. This matters most for the popup in Stage 2.
- **Theme-root CSS scoping:** slideshow surfaces carry `themeRootProps` (`src/components/theme.ts`). The popup document injects officeview's stylesheet under a theme root. No unscoped selectors.
- **Fixtures:**
  - Real decks come from LibreOffice core `sd/qa/unit/data/pptx` (MPL-2.0). They are pinned to a commit SHA, unmodified, and screened like `tests/fixtures/chartex/PROVENANCE.md`.
  - Builder decks are labelled "Office-shaped synthetic" and are never reported as Office evidence.
- **Per task:**
  - `bunx tsc --noEmit`;
  - the focused vitest files;
  - full `bunx vitest run` once before reporting;
  - build + budget + check-eager;
  - the listed Playwright specs.
- **Review:** a different model family reviews each task (see table). The stage PR also needs CI and the Codex bot review.

## Review Focus

| # | Input or condition | Expected | Pinned by |
|---|---|---|---|
| 1 | A click on a trigger shape whose exit effect has already run (the shape is invisible) | The hit test uses current render state; the click advances the main sequence and the trigger does not fire | S1-20 `a click on an exited trigger shape advances instead of firing` |
| 2 | A host that `structuredClone`s or `JSON.stringify`s a parsed `PptxDocument`, then presents the clone | Clone and stringify succeed; the clone has no handle; presenting it plays static slides with cut transitions and a `'no-package'` diagnostic (or reopens a re-readable source) | S0-4 `a cloned document has no handle and stringifies`; S1-19 `a handle-less document presents static slides with no-package` |
| 3 | A `p:spTgt spid` that names no shape on the slide (deleted shape), or a shape only on the master | The behaviour is dropped with an `unresolved-target` diagnostic; its step still counts as a click step; siblings play | S1-8 `unresolved spid drops the behaviour, keeps the step`; S1-10 `a step whose only target is unresolved still advances` |
| 4 | No visible slide in the show: every slide hidden, or `p:sldRg` entirely outside the deck | `present({from:'beginning'})` resolves and shows the end-of-show slide; the next input exits; `sldRg` outside the deck is diagnosed and treated as the whole deck | S1-7 `sldRg outside the deck falls back to the whole deck`; S1-19 `a deck with every slide hidden opens on end-of-show` |
| 5 | A sandboxed iframe that refuses full screen **and** popups at once (`sandbox="allow-scripts allow-same-origin"`) with Presenter View requested | In-window, single-screen show; two separate diagnostics (`fullscreen-refused`, `popup-blocked`); `present()` resolves, never rejects | S2-2 `iframe refusing both fullscreen and popups runs in-window single-screen` |

---

## Streams, order and agents

Each task gets a branch off `feature/slideshow` and is merged back `--no-ff` after review, followed by the combined gates. A stage's PR opens from `feature/slideshow` when its closing task is done. The next stage branches from main after that merge.

**Implementer pool.**
- **Claude** subagents take integration-heavy work.
- **codex** (OpenAI): the controller runs its browser tests and commits for it.
- **agy** (Gemini Flash) takes small, fully pinned tasks only (S1-18) and reviews.
- **opencode** (GLM) takes focused modules.
- **pi** takes one small self-contained task, reviewed by codex.

The reviewer is always a different family from the implementer.

| Task | Scope | Implementer | Reviewer | Depends on |
|---|---|---|---|---|
| **S0-1** | Eager byte recovery (Stage 0), lazy-dir rules, ledger stub | Claude | codex | — |
| **S0-2** | LibreOffice corpus import, provenance, manifest | codex | Claude | — |
| **S0-3** | Deck builder: multi-slide, layouts, masters, backgrounds | opencode | Claude | — |
| **S0-4** | Package handle, bounded reader, `HandlePackage`, slide-size validation | Claude | codex | S0-1 |
| **S0-5** | Background resolver (solid, gradient, picture, pattern, `bgRef`; inheritance) | codex | Claude | S0-3 |
| **S0-6** | Master/layout non-placeholder artwork (`showMasterSp`) | Claude | opencode | S0-1, S0-3 (merges after S0-4) |
| **S0-7** | Enrichment entry, `getPaintables` seam, underlay paint, lazy export, goldens | Claude | codex | S0-2, S0-4, S0-5, S0-6 |
| **S0-8** | Stage 0 close: validation, README, CHANGELOG, gates, PR | Claude | codex + agy | all S0 |
| **S1-1** | Eager byte recovery (Stage 1) | Claude | codex | — |
| **S1-2** | Foundation contracts (`types.ts`, clock, rng, compositor, workspace and lease interfaces); `slideshow` Playwright project and CI smoke | Claude | codex | — |
| **S1-3** | Effect catalogue, independent inventory, completeness test | codex | Claude | S1-2 |
| **S1-4** | Builder extension: transitions, timing, notes, presProps, hidden, links, media | codex | Claude | — (S0-3 merged in Stage 0) |
| **S1-5** | Bounded formula interpreter | opencode | codex | S1-2 |
| **S1-6** | Bounded motion-path parser | pi | codex | S1-2 |
| **S1-7** | Show index and slide-metadata store: `showPr`, hidden, notes, links, media, limits, eviction | codex | Claude | S1-2, S1-4 |
| **S1-8** | Transition and timing-tree reader, requirements table, unknown-node policy | Claude | codex | S1-2, S1-3, S1-4 |
| **S1-9** | Scheduler core (intervals, repeat, fill, restart, conditions, composition order) | Claude | codex | S1-2 |
| **S1-10** | Navigation layer: click steps, triggers, seek/Previous, auto-advance, kiosk, pause | codex | Claude | S1-9 |
| **S1-11** | Behaviours → render state; preset rebuild; D1 stand-ins | opencode | Claude | S1-3, S1-5, S1-6, S1-9 |
| **S1-12** | Render plan, layer rasteriser, direct painter | Claude | codex | S1-2 |
| **S1-13** | Canvas 2D compositor and mask library | opencode | Claude | S1-2 |
| **S1-14** | WebGL2 compositor and context loss | Claude | codex | S1-2, S1-12, S1-13 |
| **S1-15** | All 2-D transitions | codex | Claude | S1-3, S1-13, S1-14 |
| **S1-16** | Slideshow workspace, fallback ladder, permanent reserve | codex | opencode | S1-2 |
| **S1-17** | Eager host API, show coordinator, start sequence, events, props | Claude | codex | S1-1, S1-2 |
| **S1-18** | Present button and toolbar preload (lazy); F5 routing lives in S1-17 (eager) | agy | Claude | S1-17 |
| **S1-19** | Session controller, navigation shell, audience surface (static slides, cuts) | Claude | codex | S0-4, S1-7, S1-8, S1-10, S1-17 |
| **S1-20** | Engine wiring, hit testing, triggers, links, a11y, reduced motion, resize | Claude | opencode | S1-8, S1-10, S1-11, S1-12, S1-13, S1-14, S1-15, S1-16, S1-19 |
| **S1-21** | D4 capture checklist, refs layout, `compareFrames` (pure) | opencode | Claude | S1-3 |
| **S1-22** | Corpus conformance, compare CLI, frame benchmark | Claude | codex | S1-20, S1-21 (S0-2 merged) |
| **S1-23** | Stage 1 close | Claude | codex + agy | all S1 |
| **S2-1** | Eager byte recovery (Stage 2: popup branch, Alt+F5, screen preload) | Claude | codex | — |
| **S2-2** | Presenter popup lifecycle, eager popup branch, stylesheet/CSP host check | Claude | codex | S2-1 |
| **S2-3** | Window Management placement and Swap displays | codex | Claude | S2-2 |
| **S2-4** | Next-checkpoint preview and presenter mirror (engine side), preview lease | opencode | Claude | — (Stage 1 merged: S1-10, S1-11, S1-12, S1-16) |
| **S2-5** | Presenter View UI and single-screen swap | Claude | agy | S2-2, S2-4 |
| **S2-6** | Media reader and lease lifecycle | opencode | codex | — (Stage 1 merged: S0-4, S1-16) |
| **S2-7** | Media playback, commands, media-end rules, sounds | Claude | codex | S2-6 |
| **S2-8** | Text builds by word and letter | codex | Claude | — |
| **S2-9** | Stage 2 close | Claude | codex + agy | all S2 |
| **S3-1** | 3-D transition framework (mesh, camera, registry, look-alike dispatch), per-row algorithm sheet | Claude | codex | — |
| **S3-2** | 3-D geometric group (`prism`, flip, gallery, doors, window, switch, conveyor, ferris, pan) | codex | Claude | S3-1 |
| **S3-3** | 3-D tile/particle group (vortex, shred, fracture, crush, glitter, honeycomb, ripple, flash, reveal) | Claude | opencode | S3-1 |
| **S3-4** | 3-D deformation group (page curl, peel off, airplane, origami, curtains, drape, wind, fall over, prestige, warp) | opencode | Claude | S3-1 |
| **S3-5** | Morph matching (pure) | codex | Claude | — |
| **S3-6** | Morph rendering | Claude | codex | S3-1, S3-5 |
| **S3-7** | Stage 3 close (catalogue complete) | Claude | codex + agy | all S3 |

**Planned parallelism** (the most tasks running at once; each wave starts when its dependencies merge):

- **Stage 0, 8 tasks.**
  - Wave 1: S0-1, S0-2, S0-3 (**3**).
  - Wave 2: S0-4, S0-5, S0-6 (**3**). S0-4 and S0-6 both edit `src/pptx/parse.ts`; S0-4 merges first and S0-6 rebases.
  - Then S0-7 and S0-8.
- **Stage 1, 23 tasks.**
  - Wave 1: S1-1, S1-2, S1-4 (3).
  - Wave 2: S1-3, S1-5, S1-6, S1-7, S1-9, S1-12, S1-13, S1-16, S1-17 (**9**).
  - Wave 3: S1-8, S1-10, S1-11, S1-14, S1-18, S1-21 (6).
  - Wave 4: S1-15, S1-19 (2).
  - Then S1-20 → S1-22 → S1-23.
- **Stage 2, 9 tasks.**
  - Wave 1: S2-1, S2-4, S2-6, S2-8 (**4**).
  - Wave 2: S2-2, S2-7 (2).
  - Then S2-3 and S2-5 (2), then S2-9.
- **Stage 3, 7 tasks.**
  - Wave 1: S3-1, S3-5 (2).
  - Wave 2: S3-2, S3-3, S3-4, S3-6 (**4**).
  - Then S3-7.

**Total: 47 tasks.**

**Shared-file policy** (plan review M1, opencode I-2):
- Implementer tasks never edit the validation doc's ledger or catalogue table. Each task writes its ledger row and any verified values to `.superpowers/sdd/ledger/<task-id>.md`. The controller appends them to `docs/validation/slideshow.md` at merge time, in merge order.
- Catalogue status flips (pending → exact/approximate) happen only in the stage-closing task (S1-23, S2-9, S3-7), which also regenerates the catalogue table. Implementer tasks add the code a row needs and a test that names the row; S1-3's completeness test stays green because rows exist at every stage.
- Within a wave, tasks touching the same source file merge in table order and the later one rebases; the implementer of the later task resolves conflicts.

Stage 3 makes no eager edits, so it has no recovery task. Its closing gate asserts that the eager size equals the Stage 2 merge's figure, give or take the 2 B of rebuild variance.

**Step pattern for every task** (only the specifics are listed below):
1. Write the failing tests named in the task.
2. Run them and see each fail for the stated reason.
3. Implement the listed signatures in the listed files.
4. Run them and see them pass, together with the regression files named.
5. Run `bunx tsc --noEmit`, the build, `bun scripts/eager-budget.ts` (running total ≤ 38,647, recorded in `.superpowers/sdd/ledger/<task-id>.md`) and `bun scripts/check-eager.ts`.
6. Run the listed Playwright specs with `bunx playwright test --config playwright.viewer.config.ts --project slideshow <spec>` and a unique `IV_HARNESS_PORT`. For codex tasks, the controller runs this step.
7. Commit with the given message.

---

## Stage 0: Slide appearance (D2)

### Task S0-1: Eager byte recovery (Stage 0), lazy-dir rules, ledger

**Spec:** §3 eager budget, §8 measuring and recovery, §11 D3.

**Files:**
- Modify:
  - `tests/helpers/eager-graph.ts`: append `/src\/slideshow\//` and `/src\/pptx\/appearance\//` to `EAGER_FORBIDDEN`;
  - `vite.harness.config.ts`: add `/src/pptx/appearance/index.ts` and `/src/slideshow/index.ts` to the `lazyPagedChunks` warm list (the slideshow entry is created as a stub here);
  - the recovered-byte sites, found by measurement.
- Create:
  - `src/slideshow/index.ts` (stub `export {}`) and `src/pptx/appearance/index.ts` (stub), so the rules have targets;
  - `docs/validation/slideshow.md`, a stub with sections Gates, Eager ledger, Catalogue, Fixtures, Host requirements, Known limits, Manual checks.
- Test: extend `tests/ui-eager-graph.test.ts`.

**Procedure:**
1. On main, rebuild and record the baseline in the ledger (expected 38,584 B; a fresh rebuild is authoritative).
2. **Measure the Stage 0 growth target.** Write throwaway stubs of S0-4's and S0-7's eager edits against their final signatures, build, measure, and revert:
   - the WeakMap handle and getter;
   - `OfficePackage.dropCache()`;
   - the enrichment `import()` in `getPaintables`;
   - the `slide.underlay` `paintScene` call in `renderSlide`;
   - `collectCachedPictures` over the underlay;
   - the lazy public `enrichSlideAppearance` export;
   - S0-6's `parseArtworkTree` wrapper and the `slideTheme` export.
   Record the delta as `T0` (expected 60–150 B).
3. **Hunt for at least `T0` bytes of recovery.**
   - Rank the modules in `dist/.module-map.json` by their contribution to `dist/index.js`.
   - Look for dead branches, duplicated helpers, and eager code that only the lazy `PagedChrome`/`zoomRouting`/chart paths use. Such code can move behind an existing `import()` seam.
   - Never remove a public export or change behaviour.
4. **If the hunt falls short of `T0`:** stop, record what was found, and escalate to the owner. D3 forbids growth without recovery.

- [ ] Tests:
  - `it('src/slideshow and src/pptx/appearance are not in the eager graph')`, from the module map;
  - `it('EAGER_FORBIDDEN rejects a synthetic module id under each new directory')`;
  - every existing test passes unchanged (the recovery is behaviour-neutral).
- [ ] Ledger rows: baseline, `T0` with its stub breakdown, each recovery item with its saving, and the new running total (≤ 38,584 − `T0`).
- [ ] Commit `perf(eager): recover bytes ahead of the slideshow package handle and appearance seam`.

### Task S0-2: LibreOffice corpus import

**Spec:** §10 fixtures (real), §11 D4.

**Files:**
- Create:
  - `tests/fixtures/slideshow/*.pptx` (unmodified copies);
  - `tests/fixtures/slideshow/PROVENANCE.md`, modelled on `tests/fixtures/chartex/PROVENANCE.md` (source, MPL-2.0 statement, screening method, rejected candidates);
  - `tests/fixtures/slideshow/manifest.json`: `Array<{ file: string; sha: string; blob: string; application: string; pins: string[] }>`.
- Test: `tests/slideshow-fixtures.test.ts`.

**Source:** the copies in `/private/tmp/claude-501/-Users-julianshen-prj-officeview/aa69b0c3-1aa3-4cfb-b283-9e7808c48e7e/scratchpad/lo-pptx/files/` came from `master` raw URLs. For each file:
1. Find the last commit touching it: `gh api 'repos/LibreOffice/core/commits?path=sd/qa/unit/data/pptx/<f>&per_page=1'`.
2. Fetch the blob SHA from `contents/<path>?ref=<sha>`.
3. Check byte identity with `git hash-object <file>`. The controller runs the `gh` calls.

**Screening rule:**
- `docProps/app.xml` `<Application>` must contain "Microsoft" (metadata-reported, as in the chartex precedent).
- Each file must contain what its `pins` claim.

**Accept** (initial scan; re-verify each):

| File | Pins |
|---|---|
| `tdf104788.pptx` | 6 slides with notes, `showNarration="1"`, entrance presets 10/34/37/42/47/58, `p:animEffect`/`p:animScale`/`p:bldP`. **Benchmark deck** (§7). |
| `tdf115394.pptx` | `fade` with `spd` only, `spd`+`p14:dur` (250, 4250): the `spd` table |
| `tdf142915.pptx` | `advClick="0" advTm="1000"`, `p14:dur="0"` and `"10"`: auto-advance, zero-length transition |
| `tdf44223.pptx` | `split`, `p:sndAc`, `p:bldP animBg="1"`, entrance preset 2 |
| `bnc591147.pptx` | `interactiveSeq` with `evt="onClick"`, `mediacall` `p:cmd`, `video.avi` (poster path), `audio.mp3` |
| `connector-shape-animations.pptx` | preset 22 on connectors (Mac PowerPoint) |
| `tdf104792-smart-art-animation.pptx` | SmartArt build (`p:bldDgm` path, §6.5) |
| `narration.pptx` | `p:audio`, `p14:media`, `media1.m4a`, transitions |
| `airplane-transition.pptx` | `p15:prstTrans prst="airplane"` |
| `tdf156808.pptx` | `p14:flythrough` |
| `tdf149588_transparentSolidFill.pptx` | `p159:morph` |

Add at least two decks with gradient or picture `p:bg`, or with master non-placeholder shapes. Find them by scanning the folder for `<p:bg>` with `gradFill`/`blipFill`/`pattFill` and for `p:sp` without `p:ph` in the masters. These pin Stage 0.

**Reject (documented):**
- `effectOrder.pptx`: no `p:timing` or `p:transition`; it is an effect-list order test.
- `animation-sound-embedded.pptx` and `animation-sound-external.pptx`: no `app.xml` producer, and a 52-byte WAV.
- `tdf148478_transitionMusic.pptx`: LibreOffice-authored.

- [ ] Tests:
  - `it('every manifest file exists, matches its blob sha and reports a Microsoft producer')`;
  - `it('each fixture contains the elements its pins claim')`: a per-file regex over the slide parts;
  - `it('PROVENANCE lists every manifest file and every rejected candidate')`.
- [ ] Commit `test(slideshow): pinned LibreOffice PowerPoint corpus with provenance`.

### Task S0-3: Deck builder (layouts, masters, backgrounds)

**Spec:** §10 synthetic fixtures, §11 D2.

**Files:**
- Modify: `src/testdata/ooxml-builders.ts`. `buildPptx` keeps its exact output.
- Test: `tests/ooxml-builders-deck.test.ts`.

**Interfaces (produced):**
- `interface PptxDeckSpec { slideSize?: [number, number]; masters?: PptxMasterSpec[]; slides: PptxSlideSpec[]; presProps?: string; theme?: PptxThemeSpec }`.
- `interface PptxMasterSpec { bg?: PptxBgSpec; shapes?: PptxShapeSpec[]; layouts: PptxLayoutSpec[] }`.
- `interface PptxLayoutSpec { bg?: PptxBgSpec; shapes?: PptxShapeSpec[]; showMasterSp?: boolean; placeholders?: Array<{ type: string; idx?: number; off: [string, string]; ext: [string, string] }> }`.
- `interface PptxSlideSpec { layout?: [master: number, layout: number]; bg?: PptxBgSpec; shapes: PptxShapeSpec[]; showMasterSp?: boolean; extraXml?: { beforeSpTree?: string; afterCSld?: string } }`.
- `type PptxBgSpec = { solid: string } | { gradient: { angle: number; stops: Array<[pos: number, rgb: string]> } } | { picture: { data: Uint8Array; tile?: boolean } } | { pattern: { prst: string; fg: string; bg: string } } | { bgRef: { idx: number; color: string } }`.
- `PptxThemeSpec { bgFillStyles?: string[] }`: the raw `a:bgFillStyleLst` children.
- `PptxShapeSpec` gains `id?: number`, `name?: string` and `placeholder?: { type: string; idx?: number }`. These default to today's values.
- `buildPptxDeck(spec: PptxDeckSpec): Promise<Uint8Array>`.

- [ ] Tests:
  - `it('buildPptx output is unchanged')`: unzip the outputs of the existing call sites and compare entry names and entry contents to a stored snapshot. JSZip stamps entries with the current time, so a raw zip hash is not stable;
  - `it('buildPptxDeck writes N slides, their layout and master rels, and a parseable theme')`: parse with `parsePptx` and assert the slide count and that every relationship resolves;
  - `it('each PptxBgSpec kind serialises to the matching p:bg child')`;
  - `it('slide and layout showMasterSp="0" are written only when false')`.
- [ ] Commit `test(pptx): multi-slide deck builder with layouts, masters and backgrounds`.

### Task S0-4: Package handle, bounded reader, slide-size validation

**Spec:** §3 package ownership, §6.1 bounded uncached reads, §7 slide size validation, §8 package handle (lease), §11 D3.

**Files:**
- Modify (eager, measured):
  - `src/pptx/parse.ts`: a module-level `WeakMap<PptxDocument, PptxPackageHandle>`, set at the end of `parsePptx` with the deck-order slide part paths already resolved in the loop, followed by `pkg.dropCache()`;
  - `src/pptx/types.ts` (type only);
  - `src/core/zip.ts`: `dropCache(): void`, one line.
- Create (lazy):
  - `src/pptx/appearance/read.ts`;
  - `src/pptx/appearance/handle-package.ts`;
  - `src/pptx/appearance/size.ts`.
- Test: `tests/pptx-package-handle.test.ts`, `tests/pptx-appearance-read.test.ts`.

**Interfaces (produced):**
- `interface PptxPackageHandle { readonly pkg: OfficePackage; readonly slideParts: readonly string[] }`. `slideParts[i]` is the part of `doc.slides[i]`.
- `export function pptxPackageHandle(doc: PptxDocument): PptxPackageHandle | undefined`, in `parse.ts`. It is the only accessor. Lazy chunks import it from the eager module.
- `readPartBounded(source: Pick<OfficePackage, 'entry'>, path: string, maxBytes: number, options?: { signal?: AbortSignal; admit?: (chunkBytes: number) => boolean; output?: 'bytes' | 'chunks' }): Promise<Uint8Array | Uint8Array[] | undefined>`.
  - It streams `entry(path)` chunks with the `internalStream` pattern of `embeddedWorkbook.ts` `Reader.stream` and `zip.ts` `boundedBytes`.
  - It throws `PackageError('resource-limit')` on the chunk that crosses `maxBytes`.
  - `admit` is called **before** each chunk is retained. A `false` aborts the inflate and throws `PackageError('resource-limit')`, so an aggregate owner (the slideshow workspace, S1-16) can refuse chunk by chunk.
  - `output: 'chunks'` returns the chunk list without a final concatenation, so a large part (media, S2-6) never holds two copies.
- `readXmlBounded(source, path, maxBytes, signal?): Promise<XmlNode | undefined>`, and an ordered variant `readXmlOrderedBounded`.
- `class HandlePackage extends OfficePackage`:
  - constructed with `HandlePackage.over(handle: PptxPackageHandle, caps?: { partBytes?: number })`;
  - the constructor calls `super(new JSZip())` and never touches the base class's private `zip`;
  - the synchronous metadata methods `has`, `entry` and `paths` delegate directly to `handle.pkg` (no bounded read, no recursion);
  - the materialising methods `text`, `bytes`, `boundedText`, `xml` and `xmlOrdered` go through `readPartBounded`, with the part cap and an instance-wide aggregate cap of 32 MiB (my call) across all parts one enrichment run reads;
  - default cap 4 MiB per part;
  - its private cache lives only as long as the instance. It lets existing parse helpers run over the handle without touching the document's package cache.
- `retainedArchiveBytes(handle: PptxPackageHandle): number` (lazy, in `read.ts`): the sum of the compressed entry sizes the handle keeps alive. It is the measured cost of the lease (spec §8), reported by the validation doc for the corpus and surfaced to the slideshow as an informational `package-retained` diagnostic. No new eager document ledger is added; see the plan review dispositions.
- `validateSlideSize(widthEmu: number, heightEmu: number): { ok: true } | { ok: false; diagnostic: { kind: 'slide-size'; message: string } }`, with `SLIDE_EMU_MIN = 914_400`, `SLIDE_EMU_MAX = 51_206_400` and `SLIDE_ASPECT_MAX = 20`. **Verify the schema range** against ECMA-376 `ST_SlideSizeCoordinate` and record it in the validation doc.

- [ ] Tests:
  - `it('parsePptx attaches a handle whose slideParts follow deck order')`;
  - `it('after parsePptx the package read cache is empty and handle reads go through entry() only')`: spy on `entry`, `text` and `bytes`;
  - `it('a cloned document has no handle and stringifies')`: `structuredClone(doc)` and `JSON.stringify(doc)` both succeed, and `pptxPackageHandle(clone)` is `undefined` (Review Focus 2);
  - `it('readPartBounded refuses on the chunk crossing the cap even when the declared size lies')`: a forged-size entry;
  - `it('HandlePackage reads never populate the source package cache')`;
  - `it('HandlePackage has, entry and paths delegate synchronously; text over the per-part or aggregate cap refuses')`;
  - `it('admit returning false on the third chunk aborts the read and retains no chunk')`;
  - `it('retainedArchiveBytes equals the sum of compressed entry sizes of a builder deck')`;
  - `it('the handle does not outlive the document')`: a `WeakRef` to the package is cleared after the document is dropped and `globalThis.gc()` runs (vitest with `--expose-gc` in this file's pool options);
  - `it('validateSlideSize accepts the range ends and rejects zero, negative, NaN, Infinity, one EMU outside each end and 1:21')`.
- [ ] Ledger: the measured eager delta, which must not exceed `T0`.
- [ ] Commit `feat(pptx): package handle lease and bounded lazy part reads`.

### Task S0-5: Background resolver

**Spec:** §3 backgrounds, §8 enrichment, §11 D2.

**Files:**
- Create: `src/pptx/appearance/background.ts` (lazy).
- Test: `tests/pptx-appearance-background.test.ts`, using S0-3 decks over a plain `OfficePackage`. `HandlePackage` is an `OfficePackage`, so S0-4 is not needed.

**Interfaces (produced):**
- `type SlideBackground = { kind: 'none' } | { kind: 'fill'; fill: DrawingFill } | { kind: 'picture'; image: PptxImageRef; tile?: { sx: number; sy: number; tx: number; ty: number; algn?: string; flip?: string }; stretch?: SceneRectFractions }`.
- `resolveSlideBackground(pkg: OfficePackage, slidePath: string, theme: ThemeContext): Promise<{ background: SlideBackground; diagnostics: PptxDiagnostic[] }>`:
  - inheritance slide → layout → master; the first `p:bg` found wins;
  - `p:bgPr` goes through the existing `parseFillDefinition`;
  - `p:bgRef idx` resolves 1–999 to `a:fillStyleLst`, and ≥ 1001 to `a:bgFillStyleLst[idx − 1001]` with the `bgRef` colour as `phClr`. **Verify** against ECMA-376 §20.1.4.1.
- `backgroundNode(bg: SlideBackground, widthEmu: number, heightEmu: number): PptxShape | undefined`: a full-slide rect `PptxShape` with `drawingGeometry` rect and `drawingStyle.fill`, or a picture shape. The existing painter draws it unchanged.

- [ ] Tests:
  - `it('linear and radial gradients become a gradient DrawingFill with stops and angle')`;
  - `it('a stretched picture background becomes a full-slide picture node; tile records its scale and offset')`;
  - `it('pattern background keeps preset, fg and bg')`;
  - `it('bgRef 1001+ uses bgFillStyleLst with phClr replaced; 1-999 uses fillStyleLst')`;
  - `it('slide bg beats layout beats master; absent everywhere is none')`;
  - `it('a missing picture part yields none plus a missing-image diagnostic')`.
- [ ] Commit `feat(pptx): lazy slide background resolution with inheritance`.

### Task S0-6: Master and layout artwork

**Spec:** §3 master art, §11 D2 (`showMasterSp`, placeholders excluded).

**Files:**
- Create: `src/pptx/appearance/master-art.ts` (lazy).
- Modify (eager, measured): `src/pptx/parse.ts`.
  - `parseShapeTree` takes the private `DrawingCompatibility` and `PptxImageContext`, so exporting it alone is not callable. Instead add a narrow exported wrapper, `parseArtworkTree(pkg: OfficePackage, partPath: string, theme: ThemeContext, diagnostics: PptxDiagnostic[]): Promise<PptxShape[]>`. It loads the part's images and table styles and builds the compatibility and image context internally, exactly as the slide loop does.
  - Also add `export` to `slideTheme`.
  - Record the cost in the ledger. S0-1's stub measurement includes this wrapper, and it must fit `T0`.
- Test: `tests/pptx-appearance-master-art.test.ts`.

**Interfaces (produced):**
- `masterArtwork(pkg: OfficePackage, slidePath: string, theme: ThemeContext): Promise<{ master: PptxShape[]; layout: PptxShape[]; diagnostics: PptxDiagnostic[] }>`.
  - Non-placeholder shapes only: any shape with `p:nvPr/p:ph` is excluded at every depth.
  - Slide `showMasterSp="0"` drops both master and layout art.
  - Layout `showMasterSp="0"` drops master art only.
  - Shape `source.partPath` is the layout or master part, so ids never collide with slide ids.
  - Images go through the same loaders as the slide.
  - Shapes are parsed by `parseArtworkTree` over the layout and master parts; non-placeholders are filtered afterwards.

- [ ] Tests:
  - `it('master logo and layout band are returned; title and body placeholders are not')`;
  - `it('slide showMasterSp=0 drops both; layout showMasterSp=0 drops master only')`;
  - `it('a placeholder nested in a group is excluded')`;
  - `it('artwork source.partPath names the layout or master part')`.
- [ ] Commit `feat(pptx): lazy master and layout artwork`.

### Task S0-7: Enrichment entry, seam, underlay paint, goldens

**Spec:** §8 slide appearance enrichment, §11 D2, §7 slide size validation.

**Files:**
- Modify:
  - `src/pptx/appearance/index.ts`: `enrichSlideAppearance`;
  - `src/pptx/types.ts`: `PptxSlide.underlay?: PptxShape[]`;
  - `src/pptx/render.ts`:
    - build the scene options once;
    - after the solid fill, call `paintScene(slide.underlay ?? [], ctx, options)`;
    - then the watermark, then `slide.shapes`;
  - `src/render/paint.ts`:
    - a module-level `const loadAppearance = () => import('../pptx/appearance')` in `src/render/paint.ts`, injectable for tests through `PaintOptions.__appearanceLoader`;
    - the PPTX branch's first statement is `await loadAppearance().then(m => m.enrichSlideAppearance(doc)).catch(() => {})`, with a diagnostic on failure;
    - `collectCachedPictures` also walks `underlay`;
  - `src/index.ts`: `export const enrichSlideAppearance = (doc: PptxDocument): Promise<void> => import('./pptx/appearance').then(m => m.enrichSlideAppearance(doc))`;
  - `tests/golden.test.ts`: add entries for two corpus decks from S0-2.
- Test: `tests/pptx-appearance.test.ts`, `tests/browser/pptx-appearance.spec.ts`.

**Interfaces:**
- Consumes: S0-4 `pptxPackageHandle`, `HandlePackage.over`, `validateSlideSize`; S0-5 `resolveSlideBackground`, `backgroundNode`; S0-6 `masterArtwork`.
- Produces: `enrichSlideAppearance(doc: PptxDocument): Promise<void>`. It:
  - is idempotent (a module `WeakSet` of enriched docs);
  - with no handle or an invalid size, adds a `slide-size`/`no-package` diagnostic and changes nothing;
  - sets `slide.underlay = [backgroundNode, ...master, ...layout]`;
  - appends new image refs to `doc.images` (deduplicated with the existing key) and sets `imageIndex`;
  - keeps `slide.background` as the solid approximation.

**Paint order (decided):** solid fill → underlay (background, master art, layout art) → watermark → slide shapes. The watermark stays above all inherited art and below slide content, as today.

- [ ] Tests:
  - `it('getPaintables paints a gradient background and a master logo for a builder deck')`: pixel probes at known points;
  - `it('enrichment is idempotent across two getPaintables calls')`: the underlay and image counts are unchanged;
  - `it('the watermark paints over an opaque picture background')`: pixel probe;
  - `it('renderSlide without enrichment paints the pre-Stage-0 appearance; after enrichSlideAppearance it paints the underlay')`;
  - `it('the worker path (handleWorkerRequest) returns the enriched slide')`;
  - `it('an invalid slide size skips enrichment with a slide-size diagnostic')`;
  - `it('an appearance chunk failure still returns pages')`: inject a failing loader;
  - `it('text search and coverage ignore underlay shapes')`.
- [ ] Browser: `test('a corpus deck shows its designed background and master art (browser)')`, `test('print output includes the background (browser)')`.
- [ ] Goldens: re-run on the Mac. List every changed golden with its reason and inspect each one.
- [ ] Ledger: S0-4 + S0-6 + S0-7 eager deltas ≤ `T0`; running total ≤ 38,647.
- [ ] Commit `feat(pptx): slide backgrounds and master artwork through the lazy appearance seam`.

### Task S0-8: Stage 0 close

**Files:** `docs/validation/slideshow.md`, `README.md` (a "Slide appearance" note: hosts calling `renderSlide` directly run `enrichSlideAppearance` first), `CHANGELOG.md` (Unreleased → Added).

- [ ] The validation doc covers:
  - gates;
  - the eager ledger: baseline, `T0`, recoveries, per-task deltas, final;
  - the package-handle memory note (the compressed archive is charged to the document);
  - the verified slide-size range and `bgRef` index rule;
  - fixtures and their SHAs;
  - the changed goldens.
- [ ] Full gates:
  - tsc and `typecheck:scripts`;
  - full vitest;
  - budget ≤ 38,647;
  - check-eager;
  - Mac goldens with the enumerated deltas;
  - full Playwright (Chromium).
- [ ] Commit `docs: slideshow stage 0 validation, README and CHANGELOG`.
- [ ] Open the PR `feature/slideshow` → `main`, titled "Slideshow stage 0: slide backgrounds and master artwork". Merge after CI and the Codex bot review.

---

## Stage 1: Core show

Internal checkpoints (spec §13) map to tasks as follows:
1. S1-17, S1-18, S1-19;
2. S1-7, S1-8;
3. S1-9, S1-10;
4. S1-12, S1-13;
5. S1-14;
6. S1-11;
7. S1-15;
8. S1-16.

Each checkpoint is gated by its cross-family review before S1-20 integrates it.

### Task S1-1: Eager byte recovery (Stage 1)

**Spec:** §8, §11 D3.

1. Re-measure on main after the Stage 0 merge.
2. **Measure the target `T1`.** Write throwaway stubs against S1-17's final interfaces, measure, and revert. The stubs cover:
   - `forwardRef` on `OfficeDoc` and `OfficeFile`;
   - the pre-mounted host element;
   - `src/components/present.ts` (the `SlideShowError` class, `PresentPlatform` defaults, the show coordinator, the synchronous start prologue and the `import('../slideshow')` loader);
   - the new props;
   - F5/Shift+F5 routing.
   Stage 1's eager code does **not** include the popup branch or Alt+F5; Stage 2 adds them (S2-1).
3. **Hunt** for ≥ `T1`, with the same method and rules as S0-1. Prefer moving `OfficeDoc.tsx` logic that only the lazy chrome needs behind the existing `PagedChrome`/`zoomRouting` seams.
4. **Fall-short rule:** stop and escalate to the owner. Do not start S1-17's eager edits.

- [ ] Tests: the existing suite passes unchanged; `tests/ui-eager-graph.test.ts` stays green.
- [ ] Ledger: `T1` with its breakdown, each recovery, running total.
- [ ] Commit `perf(eager): recover bytes ahead of the slideshow host plumbing`.

### Task S1-2: Foundation contracts

**Spec:** §4.2, §6.1 model, §6.2 states and scheduler inputs, §6.3 compositor, §6.4 rows, §7 workspace.

**Files:**
- Create `src/slideshow/types.ts` (types only), `src/slideshow/clock.ts`, `src/slideshow/rng.ts`, `src/slideshow/diagnostics.ts`.
- Modify `playwright.viewer.config.ts`: add the `slideshow` project (Chromium, `testMatch: /slideshow-.*\.spec\.ts/`, launch args `--use-gl=angle --use-angle=swiftshader --enable-unsafe-swiftshader`). The CI workflow under `.github/workflows/` runs `--project slideshow`.
- Create `tests/browser/slideshow-smoke.spec.ts` and a minimal harness page `src/debug/slideshow-smoke.html`.
- Test: `tests/slideshow-foundation.test.ts`.

**Interfaces (produced; later tasks use these names):**

*Model:*
- `ShowModel { index: ShowIndex; store: SlideMetadataStore }` (both defined below).
- `ShowSlide { index: number; hidden: boolean; transition?: TransitionSpec; timing: TimingTree | { refused: string } | undefined; notes: NoteParagraph[]; links: LinkSpec[]; media: MediaRecord[] }`.
- `ShowSettings { loop: boolean; showAnimation: boolean; useTimings: boolean; showNarration: boolean; type: 'present' | 'browse' | 'kiosk'; kioskRestartMs: number; range?: { st: number; end: number } }`.
- `TransitionSpec { element: string; ns: 'p' | 'p14' | 'p15' | 'p159'; preset?: string; attrs: Readonly<Record<string, string>>; durationMs: number; advClick: boolean; advTmMs?: number; sound?: { start?: { rId: string; loop: boolean }; end?: true } }`.

*Timing tree:*
- `TimingTree { root: TimeNode; build: BuildTemplate[] }`.
- `TimeNode = ContainerNode | BehaviourNode`, where `ContainerNode { kind: 'par' | 'seq' | 'excl'; ctn: CTn; children: TimeNode[]; seq?: { concurrent: boolean; nextAc: string; prevAc: string } }`.
- `CTn { id: number; presetClass?: string; presetId?: number; presetSubtype?: number; nodeType?: string; grpId?: number; durMs: number | 'indefinite' | undefined; repeatCount?: number | 'indefinite'; repeatDurMs?: number | 'indefinite'; accel: number; decel: number; autoRev: boolean; fill?: 'hold' | 'remove' | 'freeze' | 'transition'; restart?: 'always' | 'whenNotActive' | 'never'; syncBehavior?: 'canSlip' | 'locked'; masterRel?: 'sameClick' | 'lastClick' | 'nextClick'; afterEffect?: boolean; start: Condition[]; end: Condition[]; endSync?: Condition; iterate?: Iterate }`.
  - Every attribute of the spec §6.1 table has a field. A field the engine does not act on (for example `syncBehavior`) is still carried, and S1-8 records it as "carried, diagnosed" in the validation doc.
  - `accel` and `decel` are fractions, 0–1.
  - `repeatCount` is in iterations; the reader converts thousandths.
- `Condition { delayMs: number | 'indefinite'; evt?: string; tn?: number; target?: TargetRef; rtn?: string }`.
- `BehaviourNode`: discriminated by `kind`, one of `'set' | 'anim' | 'animEffect' | 'animMotion' | 'animRot' | 'animScale' | 'animClr' | 'cmd' | 'audio' | 'video'`. Each has `ctn`, `target: TargetRef`, `attrNames: string[]`, `additive`, `accumulate`, `override?`, `xfrmType?: 'pt' | 'img'`, `rctx?: string` and its kind's fields (`keys: Keyframe[]`, `path?: MotionPath`, `filter?`, `transition?`, `by/from/to`, `clrSpc`, `dir`, `cmdType`, `cmd`).
- `TargetRef = { spid: string; text?: { kind: 'p' | 'char'; st: number; end: number }; sub?: 'bg' | 'subSp' | 'chart' | 'graphic' | 'ink' } | { bg: true } | { sld: true } | { snd: string }`.
- `targetKey(ref: TargetRef): TargetKey` (a string).
- `BuildTemplate = { kind: 'bldP'; spid: string; grpId: number; build: 'whole' | 'p' | 'cust' | 'allAtOnce'; bldLvl: number; animBg: boolean; rev: boolean } | { kind: 'bldGraphic' | 'bldOleChart' | 'bldDgm'; spid: string; grpId: number; bld?: string }`.
- `Iterate { type: 'el' | 'wd' | 'lt'; backwards: boolean; tmAbsMs?: number; tmPct?: number }`.
- `Keyframe { tm: number | 'indefinite'; value: string | number; fmla?: string }`.
- `NoteParagraph { level: number; bullet?: string; text: string }`.
- `LinkSpec { spid: string; run?: { p: number; r: number }; action: { kind: 'slide'; index: number } | { kind: 'jump'; verb: 'nextslide' | 'previousslide' | 'firstslide' | 'lastslide' | 'endshow' } | { kind: 'url'; href: string } | { kind: 'refused'; reason: string } }`.
- `MediaRecord { rId: string; partPath: string; contentType: string; declaredBytes: number; spid?: string }`.
- `TargetResolver { shape(spid: string): 'slide' | 'master-only' | 'missing'; paragraphs(spid: string): number; runs(spid: string, p: number): number }`, built from the parsed `PptxSlide` (its shapes after compatibility selection) and its underlay.
- `BuildOwnership { spid: string; mode: 'whole' | 'paragraph'; ranges: ReadonlyArray<{ key: TargetKey; st: number; end: number }>; animBg: boolean; body: TargetKey }`: the normalised per-shape build record that the render planner consumes.
- `ShowIndex { slides: ReadonlyArray<{ index: number; hidden: boolean; partPath?: string }>; settings: ShowSettings; diagnostics: SlideShowDiagnostic[] }`: the cheap deck-level part of the model.
- `SlideMetadataStore { get(index: number): Promise<ShowSlide>; peek(index: number): ShowSlide | undefined; evict(index: number): void; resident(): readonly number[] }`. Every resident `ShowSlide` (notes, links, media records, timing, builds, compiled formulas and paths) is charged to one workspace lease per slide.

*Render state:*
- `RenderState = ReadonlyMap<TargetKey, TargetState>`.
- `TargetState { visible: boolean; opacity: number; transform: Affine; mask?: MaskState; color?: { fill?: ColorMatrix; line?: ColorMatrix; text?: ColorMatrix }; motion?: { dx: number; dy: number } }`.
- `Affine = readonly [a, b, c, d, e, f]`; `ColorMatrix = readonly number[]`, a 4×5 matrix.
- `MaskState { kind: MaskKind; progress: number; params: Readonly<Record<string, number | string>> }`.

*Events and clock:*
- `ShowEvent = { kind: 'next'; t: number } | { kind: 'previous'; t: number } | { kind: 'trigger'; spid: string; t: number } | { kind: 'mouseOver' | 'mouseOut'; spid: string; t: number } | { kind: 'mediaEnd'; nodeId: number; t: number; synthetic: boolean } | { kind: 'stopAudio'; t: number }`. These cover every `evt` value in the spec's condition list (`onBegin`/`onEnd`/`begin`/`end` are internal to the scheduler).
- `ShowClock { now(): number; paused: boolean }`, `ManualClock implements ShowClock { advance(ms: number): void; pause(): void; resume(): void }`, `RafClock(win: Window)`.

*Compositor:*
- `Compositor { readonly backend: 'webgl2' | 'canvas2d'; readonly maxTextureSize: number; upload(layer: LayerBitmap, lease: Lease): LayerHandle | { refused: 'workspace' | 'texture-size' }; release(handle: LayerHandle): void; frame(draws: readonly LayerDraw[], transition?: TransitionDraw): void; resize(widthPx: number, heightPx: number): void; dispose(): void }`.
- `LayerDraw { handle: LayerHandle; state: TargetState; part?: 'fill' | 'line' | 'text' }`.
- `TransitionDraw { from: LayerHandle; to: LayerHandle; effect: string; progress: number; params: Readonly<Record<string, unknown>>; rng: Rng }`.
- `CanvasFactory = (w: number, h: number) => HTMLCanvasElement | OffscreenCanvas`, injected so node-canvas drives tests.
- `LayerBitmap { canvas: HTMLCanvasElement | OffscreenCanvas; bounds: { x: number; y: number; w: number; h: number }; scale: number; bytes: number }`. Bounds are padded for stroke, shadow and glow.
- `LayerHandle` is an opaque `{ readonly id: number }`.
- `MaskKind`, the string union of every mask the catalogue names (filled in by S1-13).

*Workspace:*
- `SlideShowWorkspace { readonly limit: number; lease(bytes: number, tag: string): Lease | undefined; reserveTerminal(outputBytes: number): TerminalReserve | undefined; used(): number }`.
- `Lease { readonly tag: string; readonly bytes: number; grow(bytes: number): boolean; shrink(bytes: number): void; transfer(tag: string): Lease; release(): void }`. `release` is idempotent and returns the bytes exactly once; `transfer` moves ownership without a second charge.
- `TerminalReserve { readonly outputBytes: number; readonly scratch: SubOwner; resize(outputBytes: number): boolean; release(): void }`.
- `SubOwner { lease(bytes: number, tag: string): Lease | undefined }`: a bounded sub-owner carved from bytes already reserved. Its leases never touch the parent ledger, so terminal scratch is not charged twice.
- **Allocation rule (plan review I1):** every producer receives a `Lease` or `SlideShowWorkspace` in its constructor or call and reserves **before** allocating: the readers (S1-7, S2-6) through `readPartBounded`'s `admit`; the rasteriser (S1-12) before creating a canvas; the compositors (S1-13, S1-14) before each texture, framebuffer and mask upload, holding the CPU and GPU copies together during upload; the backend manager (S1-14) before mounting the 2-D canvas and before restoring.

*Catalogue row:*
- `CatalogueRow { id: string; kind: 'transition' | 'animation' | 'behaviour'; element?: string; ns?: string; preset?: { cls: string; id: number; subtype?: number }; uiName: string; implementation: 'file' | 'preset-table' | 'mask' | 'transform' | 'shader' | 're-raster' | 'stand-in'; status: 'exact' | 'approximate' | 'pending-stage-2' | 'pending-stage-3' | 'never'; verified: boolean; lookalike?: string; tolerance: Tolerance; stage: 1 | 2 | 3 }`.
- `Tolerance { meanAbs: number; rowMax: number }`, as fractions of 255.

*Diagnostics and random numbers:*
- `SlideShowDiagnostic { code: string; slideIndex?: number; effect?: string; message: string }`.
- `DiagnosticSink { add(d: SlideShowDiagnostic): void; list(): readonly SlideShowDiagnostic[] }`, deduplicated by `(code, slideIndex, effect)`.
- `type Rng = () => number`; `createRng(seed: number): Rng` (mulberry32).

- [ ] Tests:
  - `it('ManualClock advances only when told and stands still while paused')`;
  - `it('a lease releases exactly once; transfer does not double-charge; a SubOwner never reaches the parent')`, against a reference ledger in the test;
  - `it('createRng is deterministic per seed and differs across seeds')`;
  - `it('DiagnosticSink deduplicates per code, slide and effect')`;
  - `it('targetKey is distinct for shape, paragraph range, char range and bg')`.
- [ ] Browser (CI smoke, blocking): `test('the slideshow project creates a WebGL2 context under SwiftShader')` and `test('headless Chromium enters and exits element full screen on a click')`.
  - If WebGL2 cannot be created in CI, the stage PR is blocked and escalated to the owner (controller ruling). Covering only Canvas 2D in CI is not allowed.
  - If headless full screen is unsupported, the spec allows emulation (§10). The start tests then drive full screen through an injected `PresentPlatform` that records the calls, and the validation doc says which assertions are emulated. This is not a skip: the same assertions run.
- [ ] Commit `feat(slideshow): foundation contracts and slideshow browser project`.

### Task S1-3: Effect catalogue and independent inventory

**Spec:** §6.4, §11 D1, D4, §6.3 transition lists, §6.5.

**Files:**
- Create:
  - `src/slideshow/inventory.ts`: data only, enumerated from the standards and **not** imported from any reader;
  - `src/slideshow/catalogue.ts`.
- Modify: `docs/validation/slideshow.md` (generated table section).
- Test: `tests/slideshow-catalogue.test.ts`.

**Inventory sources** (record each source and version in the validation doc):
- **ECMA-376 `p:transition` children:** `blinds`, `checker`, `circle`, `comb`, `cover`, `cut`, `diamond`, `dissolve`, `fade`, `newsflash`, `plus`, `pull`, `push`, `random`, `randomBar`, `split`, `strips`, `wedge`, `wheel`, `wipe`, `zoom`.
- **[MS-PPTX] `p14` transitions** (verify the list): `vortex`, `switch`, `flip`, `ripple`, `honeycomb`, `prism`, `doors`, `window`, `ferris`, `gallery`, `conveyor`, `pan`, `glitter`, `warp`, `flythrough`, `flash`, `shred`, `reveal`, `wheelReverse`.
- **`p15:prstTrans@prst`** (verify the list): `fallOver`, `drape`, `curtains`, `wind`, `prestige`, `fracture`, `crush`, `peelOff`, `pageCurlDouble`, `pageCurlSingle`, `airplane`, `origami`.
- **`p159:morph`** with its `option` values.
- **The behaviour elements** of the §6.1 table.
- **PowerPoint's animation preset list** (`presetClass` × `presetID` × `presetSubtype`). Never derive it from the corpus or from our own parser.
  - **Step 0 (controller, before dispatch):** pin the exact source document, its version and a durable retrieval URL in `src/slideshow/inventory.ts`'s header. If no independent Microsoft-published list can be found, stop and ask the owner which source to accept; do not substitute corpus coverage.

**Rows:**
- Stage 1 statuses:
  - 2-D transitions and file-defined behaviours are `exact`;
  - preset-table rebuilds and §6.5 items are `approximate`;
  - 3-D, Morph, media commands, sounds and word/letter builds are `pending-stage-2` or `pending-stage-3` with their D1 stand-in;
  - `never` applies only to §2.3 items.
- Every 3-D row names its 2-D `lookalike` (for example `prism` → `push`, `vortex` → `dissolve`).
- Tolerance defaults (my call): exact `{ meanAbs: 0.02, rowMax: 0.10 }`, approximate `{ meanAbs: 0.08, rowMax: 0.30 }`. Shader rows may loosen with a reason.
- The table also records the precedence questions:
  - trigger beats hyperlink (pending D4);
  - kiosk link inertness (pending D4).

**Interfaces (produced):**
- `CATALOGUE: readonly CatalogueRow[]`;
- `rowForTransition(ns: string, element: string, preset?: string): CatalogueRow | undefined`;
- `rowForPreset(cls: string, id: number, subtype?: number): CatalogueRow | undefined`;
- `catalogueMarkdown(): string`, the table the validation doc embeds;
- `INVENTORY: { transitions: ReadonlyArray<{ ns: string; element: string; preset?: string }>; behaviours: readonly string[]; presets: ReadonlyArray<{ cls: string; id: number; subtype?: number }> }`.

- [ ] Tests:
  - `it('every inventory transition, behaviour and preset maps to exactly one catalogue row')`: the completeness test, run at every stage;
  - `it('every row has an implementation, a status, a tolerance and a stage; verified is false unless a compare record exists')`;
  - Stage 1 rows are written here with their Stage 1 status; later status flips happen only in stage-closing tasks (shared-file policy).
  - `it('every 3-D row names a lookalike that is itself a 2-D exact row')`;
  - `it('no row is pending a stage that has already landed')`. It reads `LANDED_STAGE`, a constant in `catalogue.ts` set to 1 here and raised by each stage close;
  - `it('the validation doc catalogue section equals catalogueMarkdown()')`;
  - `it('the inventory module imports nothing from src/slideshow/read')`.
- [ ] Commit `feat(slideshow): effect catalogue with an independent inventory`.

### Task S1-4: Builder extension for show content

**Spec:** §10 synthetic fixtures.

**Files:** Modify `src/testdata/ooxml-builders.ts`. Test: `tests/ooxml-builders-show.test.ts`.

**Interfaces (produced):** `PptxSlideSpec` gains:
- `transition?: { xml: string; wrapInAlternateContent?: { requires: string; fallbackXml: string } }`;
- `timingXml?: string`;
- `hidden?: boolean`;
- `notes?: string[][]`, paragraphs of runs, written as a `notesSlide` part with its rels;
- `links?: Array<{ shapeId: number; action?: string; target?: string; external?: boolean; onRun?: boolean }>`;
- `media?: Array<{ kind: 'audio' | 'video'; data: Uint8Array; ext: string; shapeId: number; p14?: boolean }>`.

`PptxDeckSpec.presProps` is written verbatim.

Helper: `timingXmlOf(nodes: string): string`, which wraps `p:tnLst/p:par/p:cTn[nodeType=tmRoot]/p:childTnLst`, plus `mainSeqXml(steps: string[]): string`.

- [ ] Tests:
  - `it('an AlternateContent transition keeps mc:Choice Requires and the fallback')`;
  - `it('notes slides, hidden slides, hyperlink rels and media rels resolve')`;
  - `it('presProps is written to ppt/presProps.xml with its relationship')`.
- [ ] Commit `test(pptx): builder support for transitions, timing, notes, links and media`.

### Task S1-5: Bounded formula interpreter

**Spec:** §6.1 formulas, §7 hostile input.

**Files:** Create `src/slideshow/formula.ts`. Test: `tests/slideshow-formula.test.ts`.

**Interfaces (produced):**
- `compileFormula(src: string): { ok: true; fn: CompiledFormula } | { ok: false; reason: 'too-long' | 'too-many-tokens' | 'too-deep' | 'syntax' | 'unknown-identifier' }`.
- `type CompiledFormula = (env: FormulaEnv) => number`.
- `FormulaEnv { $: number; ppt_x: number; ppt_y: number; ppt_w: number; ppt_h: number; current: (name: string) => number; rand: Rng }`.

Grammar: a recursive-descent parser to an AST, evaluated by a switch.
- Literals, `+ - * / ^` and unary minus, parentheses.
- Variables `$`, `ppt_x`, `ppt_y`, `ppt_w`, `ppt_h`, `#name`, `pi`, `e`.
- Functions `abs acos asin atan ceil cos deg exp floor ln max min rad rand sin sqrt tan`.

Verify the grammar against ECMA-376 §19.5.4 notes and record any differences.

- [ ] Tests:
  - `it('evaluates arithmetic, precedence, right-associative ^ and unary minus')`;
  - `it('#ppt_x and ppt_w read the environment')`;
  - `it('rand uses the injected Rng')`;
  - `it('accepts 1024 bytes, 256 tokens and depth 32; rejects 1025 bytes, 257 tokens, depth 33 and every identifier not in the list')`;
  - `it('rejects constructor, __proto__, this, globalThis, brackets, quotes, semicolons and backticks')`;
  - `it('the module contains no eval or Function')`: a source scan.
- [ ] Commit `feat(slideshow): bounded animation formula interpreter`.

### Task S1-6: Bounded motion-path parser

**Spec:** §6.1 `p:animMotion` path, reader limits.

**Files:** Create `src/slideshow/motion-path.ts`. Test: `tests/slideshow-motion-path.test.ts`.

**Interfaces (produced):**
- `parseMotionPath(src: string): { ok: true; path: MotionPath } | { ok: false; reason: 'too-long' | 'too-many-commands' | 'syntax' }`.
- `MotionPath { segments: ReadonlyArray<{ cmd: 'M' | 'L' | 'C' | 'Z'; pts: readonly number[] }>; length: number }`.
- `pointAt(path: MotionPath, fraction: number): { x: number; y: number; angle: number }`. Coordinates are fractions of the slide. `E` ends the path. Relative lower-case commands are supported. Cubics are arc-length sampled with 32 samples per segment.

- [ ] Tests:
  - `it('parses M L C Z E, absolute and relative')`;
  - `it('pointAt 0, 0.5 and 1 on a known line and a known cubic')`;
  - `it('accepts 65536 characters and 10000 commands; refuses 65537 characters and 10001 commands')`;
  - `it('refuses NaN, Infinity and stray tokens')`.
- [ ] Commit `feat(slideshow): bounded motion path parser`.

### Task S1-7: Show index and slide-metadata store

**Spec:** §6.1 (notes, hidden, links, media, reader limits), §4.4 defaults and show types, §4.3 hyperlinks, §2.2 custom shows, §7 metadata budget.

**Files:** Create `src/slideshow/read/deck.ts`, `read/store.ts`, `read/notes.ts`, `read/links.ts`, `read/media.ts`, `read/limits.ts`. Test: `tests/slideshow-read-deck.test.ts`, `tests/slideshow-read-store.test.ts`.

**Interfaces:**
- Consumes: S0-4 `pptxPackageHandle`, `readXmlBounded`, `validateSlideSize`; S1-2 types.
- Produces:
  - `READER_LIMITS`, a frozen object of every §6.1 value;
  - `readShowIndex(doc: PptxDocument, handle: PptxPackageHandle | undefined, sink: DiagnosticSink, signal?: AbortSignal): Promise<ShowIndex>`. It reads `presProps` and each slide's `show` attribute (a bounded read of the slide root only). With no handle, every slide is static with `no-package`.
  - `createSlideMetadataStore(args: { doc: PptxDocument; handle?: PptxPackageHandle; index: ShowIndex; workspace: SlideShowWorkspace; sink: DiagnosticSink; timing?: SlideTimingReader; resolver: (slide: PptxSlide) => TargetResolver; signal?: AbortSignal }): SlideMetadataStore`.
    - `get(i)` reads slide part `i` **once** (bounded), parses notes, links and media, and calls the injected `timing` reader on the same parsed XML. This is the single assembly point for `ShowSlide`; S1-19 passes S1-8's `readSlideTiming`.
    - Everything resident (notes, links, media records, timing, builds, compiled formulas and paths) is estimated and charged to one lease per slide, `meta:<i>`, which grows as later consumers attach compiled data. The aggregate across slides is capped at 64 MiB by a store-level counter on top of the workspace.
    - `evict(i)` releases the lease and drops every reference; a later `get(i)` re-reads.
  - `type SlideTimingReader = (slideXml: XmlNode, slideIndex: number, resolve: TargetResolver, sink: DiagnosticSink) => { transition?: TransitionSpec; timing: TimingTree | { refused: string } | undefined; builds: BuildOwnership[] }`.
  - `openShowPackage(source: Blob | ArrayBuffer | Uint8Array | string): Promise<PptxPackageHandle>`, used for re-readable sources only.
  - it fills S1-2's `LinkSpec`, `MediaRecord` and `NoteParagraph`.

Rules:
- Defaults: `loop`=false, `showNarration`=false, `showAnimation`=true, `useTimings`=true. `restart` defaults to 300,000. Verify every default against the schema.
- `sldRg` is 1-based and inclusive (verify). Out of the deck → whole deck, with a diagnostic.
- `custShow` → full deck, with a diagnostic. `browse` → present, with a diagnostic.
- URLs: only `http:`, `https:` and `mailto:`. Slide-jump verbs: verify the `ppaction://` set.
- The metadata estimate is charged as slides are read. A slide that would push the aggregate over 64 MiB is evicted-from-neighbours first; if it still does not fit, that slide is static, with a diagnostic.

- [ ] Tests:
  - `it('showPr defaults when attributes are absent; kiosk restart default 300000')`;
  - `it('sldRg st..end inclusive; sldRg outside the deck falls back to the whole deck')` (Review Focus 4);
  - `it('custShow and browse are diagnosed and play the full deck as present')`;
  - `it('hidden slides keep their index')`;
  - `it('notes keep level and bullet; 65536 chars of notes are kept and 65537 truncate with a diagnostic; a 1 MiB notes part is read and 1 MiB + 1 is refused')`;
  - `it('a 1 MiB presProps is read; 1 MiB + 1 is refused with defaults applied')`;
  - `it('metadata at 64 MiB is accepted; one byte more evicts a neighbour, then makes the slide static')`: through an injected estimator;
  - `it('notes, links, media and an injected timing tree that each fit but together exceed the cap are refused as one aggregate')`;
  - `it('evict releases the lease and drops references; a revisited slide is re-read')`: a `WeakRef` to the old `ShowSlide` clears after gc, and the read spy counts two reads;
  - `it('get reads each slide part once and hands the same XML to the timing reader')`;
  - `it('hlinkClick on shape and run resolves slide jumps, show jumps and http/https/mailto; javascript:, file: and ppaction://macro are refused')`;
  - `it('media records keep declared size and content type; no bytes are read')`;
  - `it('a 4 MiB slide part is read; 4 MiB + 1 drops that slide to static with a diagnostic')`;
  - `it('without a handle every slide is static and no-package is diagnosed once')`;
  - `it('tdf104788 reads 6 notes bodies and showNarration=1')`: corpus.
- [ ] Commit `feat(slideshow): show index and leased slide metadata store`.

### Task S1-8: Transition and timing-tree reader

**Spec:** §6.1 transition and semantics table, unknown-node policy, reader limits, §4.4 `advClick`/`advTm`.

**Files:** Create `src/slideshow/read/requirements.ts`, `read/transition.ts`, `read/timing.ts`. Test: `tests/slideshow-read-timing.test.ts`.

**Interfaces:**
- Consumes: S1-2 model; S1-3 `rowForTransition`; `supportedChoiceRequirements` (`src/drawing/coverage.ts`).
- Produces:
  - `SLIDESHOW_REQUIREMENTS: readonly string[]` (`p14`, `p15`, `p159`, plus any others the catalogue names);
  - `SPD_MS = { fast: 500, med: 750, slow: 1000 }`;
  - `readSlideTiming: SlideTimingReader` (signature in S1-7). Target resolution uses the injected `TargetResolver`, built from the parsed slide after compatibility selection, so an id present only in a discarded `mc:Choice` branch is unresolved;
  - `expandBuilds(build: readonly BuildTemplate[], resolve: TargetResolver): BuildOwnership[]`: `bldP build="p"` with `bldLvl` gives one range key per paragraph at or above the level (deeper paragraphs join their parent); `animBg` makes the body a target;
  - `shapeResolver(slide: PptxSlide): TargetResolver` (in `read/resolve.ts`).

**Verify in this task** (against ECMA-376 / [MS-PPTX] and the corpus) and record in the validation doc:
- the `spd` table. tdf115394 and tdf150770 write `spd="med"` beside `p14:dur="700"`; whether `med` maps to 750 or to the nearest-of value is a D4 question;
- `repeatCount` thousandths;
- the `cBhvr` attribute-name list;
- `p:extLst` transition placement;
- layout and master transition inheritance (diagnosed, not inherited);
- `numSld`.

Rules:
- An unknown element inside a known container is dropped with a diagnostic naming it.
- An unknown container drops its subtree.
- Exceeding a limit → `{ refused }`, and the slide is static.
- A `tn` reference cycle is broken with a diagnostic.
- An unresolved `spid` drops the behaviour with `unresolved-target`; its container stays. A master-only id gets `master-target` and is dropped the same way.
- Fields the engine carries but does not act on (`syncBehavior`, `masterRel`, `rctx`) are recorded as "carried, diagnosed" in the validation doc and diagnosed once per slide when present.

- [ ] Tests:
  - `it('mc:Choice p14 transition is chosen by namespace, not prefix; an aliased prefix works; unknown Requires falls back')`;
  - `it('p:extLst transition children are read; an unknown extension URI is diagnosed')`;
  - `it('spd without dur uses SPD_MS; p14:dur wins; advClick=0 and advTm parse; missing advTm is undefined')`;
  - `it('tdf115394 transition durations match the file')` and `it('tdf142915 reads advClick=0 advTm=1000 with dur 0 and 10')`: corpus;
  - `it('repeatCount 2500 is 2.5 iterations; accel 50000 is 0.5; animRot by 5400000 is 90 degrees; animScale 150000 is 1.5')`;
  - `it('every behaviour element of the semantics table parses its fields')`: one builder deck per behaviour;
  - `it('unknown element is dropped with siblings kept; unknown container drops its subtree')`;
  - `it('2000 time nodes, depth 64, 10000 conditions, 1000 tav and 20000 tav per slide are accepted')`;
  - `it('2001 time nodes, depth 65, 10001 conditions, 1001 tav and 20001 tav per slide each refuse to static')`;
  - `it('a tn cycle is broken with a diagnostic')`;
  - `it('unresolved spid drops the behaviour, keeps the step')` (Review Focus 3);
  - `it('an spid present only in a discarded mc:Choice branch is unresolved; a master-only id is master-target')`;
  - `it('identical shapes with bldLvl 1 vs 2 and animBg 0 vs 1 expand to different BuildOwnership')`;
  - `it('syncBehavior, masterRel, xfrmType and rctx are carried on the normalised nodes')`;
  - `it('a transition on a layout is diagnosed and not inherited')`;
  - `it('tdf104788 timing tree has six entrance steps with the file presets')`.
- [ ] Commit `feat(slideshow): transition and timing tree reader`.

### Task S1-9: Scheduler core

**Spec:** §6.2 scheduler, `fill`/`restart`/`autoRev`/`accel`/`decel`/repeat/`additive`/`accumulate`; §7 hostile input.

**Files:** Create `src/slideshow/engine/scheduler.ts`, `engine/ease.ts`. Test: `tests/slideshow-scheduler.test.ts`, which builds `TimingTree` values directly.

**Interfaces (produced):**
- `createScheduler(tree: TimingTree): Scheduler`.
- `Scheduler { evaluate(events: readonly ShowEvent[], t: number): SchedulerSnapshot; endTimeOf(nodeId: number, events: readonly ShowEvent[]): number | 'indefinite' | undefined }`. `evaluate` is pure: equal inputs give an equal output.
- `SchedulerSnapshot { active: ReadonlyArray<{ node: BehaviourNode; progress: number; iteration: number }>; frozen: ReadonlyArray<{ node: BehaviourNode; progress: number; iteration: number }>; ended: ReadonlySet<number> }`. `progress` is local simple time after accel/decel/autoRev, 0–1.
- `easeProgress(p: number, accel: number, decel: number, autoRev: boolean): number`.

- [ ] Tests:
  - `it('par children start together; seq children chain; excl pauses siblings')`;
  - `it('delay, dur and indefinite; endSync and end conditions cut a node short')`;
  - `it('repeat 2.5 iterations ends mid-iteration; repeatDur caps; indefinite never ends')`;
  - `it('autoRev doubles the active duration and mirrors progress')`;
  - `it('accel 0.3 decel 0.3 matches the closed-form ease at 0.1, 0.5 and 0.9')`;
  - `it('fill hold/freeze keeps the end value; remove drops it')`;
  - `it('restart always, whenNotActive and never while active')`;
  - `it('parallel effects ending at different times')`;
  - `it('a held effect keeps its final iteration count in frozen entries')`;
  - `it('onMouseOver, onPrev and onStopAudio start conditions fire from their events')`;
  - `it('evaluate is pure: same tree, events and t give deep-equal snapshots')`.
- [ ] Commit `feat(slideshow): event-driven timing scheduler`.

### Task S1-10: Navigation layer, auto-advance, kiosk, pause

**Spec:** §6.2 click steps, interactive sequences, navigation rules, slide states; §4.4 `advClick`, `advTm`, kiosk, pause, `showAnimation`; §4.5 reduced motion (jump to end states).

**Files:** Create `src/slideshow/engine/steps.ts`, `engine/advance.ts`. Test: `tests/slideshow-steps.test.ts`, `tests/slideshow-advance.test.ts`.

**Interfaces:**
- Consumes: S1-9 `Scheduler`.
- Produces:
  - `SlideTimeline { readonly steps: number; stateAt(cursor: SlideCursor, t: number): SchedulerSnapshot; next(cursor: SlideCursor, t: number): SlideCursor; checkpoint(n: number): SlideCursor; trigger(cursor: SlideCursor, spid: string, t: number): SlideCursor | undefined; triggers(): readonly string[] }`;
  - `SlideCursor { step: number; events: readonly ShowEvent[]; mode: 'static' | 'initial' | 'checkpoint' | 'live' }`;
  - `createSlideTimeline(tree: TimingTree | undefined, options: { showAnimation: boolean; reducedMotion: boolean }): SlideTimeline`;
  - `AutoAdvance { due(cursor: SlideCursor, transitionEndT: number, t: number): boolean }`, from `createAutoAdvance(slide: ShowSlide, settings: ShowSettings, timeline: SlideTimeline)`;
  - `KioskTimer { restartDue(t: number): boolean; restarted(t: number): void }`.

Rules:
- **Next during an active step** settles finite effects; indefinite ones continue.
- **`checkpoint(n)`** replays steps 0..n−1, settled, with no trigger events. Indefinite effects are taken at the end of their first iteration.
- **Previous** clears trigger history.
- **Auto-advance** follows the §4.4 working assumption:
  - the timer starts when the transition ends;
  - it advances when the timer has run out and no finite timed effect is running;
  - it defers while unplayed click steps remain;
  - it is never held by indefinite effects, interactive sequences or waiting media.

- [ ] Tests:
  - `it('clickEffect starts a step, withEffect joins, afterEffect chains')`;
  - `it('next during an active effect settles finite effects and keeps indefinite ones running')`;
  - `it('two trigger histories at the same step and time give different render states')`;
  - `it('checkpoint n ignores trigger history and takes indefinite effects at first-iteration end')`;
  - `it('previous across a trigger clears the trigger')`;
  - `it('a step whose only target is unresolved still advances')` (Review Focus 3);
  - `it('showAnimation=0 and refused timing are one static step; exit effects are not applied')`;
  - `it('entrance and exit on the same shape')`;
  - `it('reduced motion jumps each step to its settled state')`;
  - `it('advTm missing never advances; advTm=0 advances at transition end; pending click steps defer; indefinite repeat does not hold')`;
  - `it('advClick=0 refuses clicks but accepts keyboard and API next')`: an `AutoAdvance` input flag;
  - `it('kiosk restarts restart ms after the last start or restart; each loop restart resets the counter; pause stops it')`.
- [ ] Commit `feat(slideshow): click steps, triggers, seeking and auto-advance`.

### Task S1-11: Behaviours to render state

**Spec:** §6.2 output, presets, D1 stand-ins; §6.1 `cBhvr` composition; §6.3 part-addressed colour.

**Files:** Create `src/slideshow/engine/behaviours.ts`, `engine/presets.ts`, `engine/stand-ins.ts`. Test: `tests/slideshow-behaviours.test.ts`.

**Interfaces:**
- Consumes: S1-9 `SchedulerSnapshot`; S1-5 `compileFormula`; S1-6 `pointAt`; S1-3 `rowForPreset`, `rowForTransition`.
- Produces:
  - `renderStateAt(snapshot: SchedulerSnapshot, base: ReadonlyMap<TargetKey, BaseGeometry>, rng: Rng, sink: DiagnosticSink): RenderState`;
  - `BaseGeometry { x: number; y: number; w: number; h: number; rot: number }` (px);
  - `rebuildFromPreset(ctn: CTn, target: TargetRef): BehaviourNode[] | undefined` (marked approximate);
  - `standInFor(cls: string): 'visible' | 'none' | 'hidden' | 'path-end'`;
  - `transitionStandIn(): 'fade'`.

Rules:
- Composition order follows `additive`/`accumulate` (`base`, `sum`, `repl`, `mult`, `none`) in document order.
- `style.visibility` comes from `p:set`.
- `animEffect` filters map to `MaskState` kinds (`wipe(dir)`, `barn`, `box`, `checkerboard`, `circle`, `diamond`, `dissolve`, `fade`, `plus`, `randombar`, `slide`, `strips`, `wedge`, `wheel(n)`), per the catalogue.
- `animClr` addressed to `fillcolor` / `stroke.color` / `style.color` sets `color.fill` / `color.line` / `color.text`.

- [ ] Tests (numeric oracles at chosen times):
  - `it('anim ppt_x from #ppt_x-1 to #ppt_x at 0, 0.5 and 1')`;
  - `it('tav keyframes with fmla and calcmode lin and discrete')`;
  - `it('animRot by 90 at half time is 45; animScale 150% with zoomContents')`;
  - `it('animMotion follows the path and rAng; origin layout vs parent')`;
  - `it('animEffect wipe(down) progress maps to the mask')`;
  - `it('animClr hsl dir cw and rgb by; fill-only, line-only and text-only addressing')`;
  - `it('additive sum and mult compose in document order')`;
  - `it('accumulate=always with repeat 3 holds the accumulated value after the end')`;
  - `it('a builder deck with an onMouseOver trigger goes reader to scheduler to render state')`: S1-8's reader over an S1-4 deck, end to end;
  - `it('a preset with missing behaviours is rebuilt and diagnosed approximate')`;
  - `it('stand-ins: entrance visible, emphasis none, exit hidden, path end position; each diagnosed once')`;
  - `it('an invalid formula drops only its behaviour')`.
- [ ] Commit `feat(slideshow): behaviour evaluation to render state`.

### Task S1-12: Render plan, layer rasteriser, direct painter

**Spec:** §6.3 render plan, ownership, part layers, layer size, watermarks; §7 terminal fallback painter.

**Files:** Create `src/slideshow/render/plan.ts`, `render/raster.ts`, `render/direct.ts`. Test: `tests/slideshow-plan.test.ts` (node-canvas pixels).

**Interfaces:**
- Consumes: S1-2 types; `paintScene` and `renderSlide` helpers unchanged; S0-7 `slide.underlay`.
- Produces:
  - `buildRenderPlan(slide: PptxSlide, animated: ReadonlySet<TargetKey>, colourParts: ReadonlyMap<TargetKey, ReadonlySet<'fill' | 'line' | 'text'>>, builds: readonly BuildOwnership[]): RenderPlan`. Paragraph splitting comes only from `builds` (S1-2 type; produced by S1-8's `expandBuilds`; tests construct it directly);
  - `RenderPlan { layers: readonly PlanLayer[] }`, where `PlanLayer = { kind: 'static'; nodes: PptxShape[] } | { kind: 'target'; key: TargetKey; nodes: PptxShape[]; part?: 'fill' | 'line' | 'text' | 'body' | 'range'; sub?: PlanLayer[] } | { kind: 'watermark' }`. Nodes are pruned trees: the wanted nodes plus their ancestor groups;
  - `rasterLayer(layer: PlanLayer, slide: PptxSlide, ctx: RasterContext): LayerBitmap`;
  - `RasterContext { scale: number; images; resolveFont; canvas: CanvasFactory; workspace: SlideShowWorkspace; maxTextureSize: number }`;
  - it returns S1-2's `LayerBitmap`;
  - `paintDirect(ctx: CanvasRenderingContext2D, slide: PptxSlide, state: RenderState, assets: DirectAssets): void`, the terminal path, a single pass with no layers;
  - `DirectAssets { images: ReadonlyArray<CanvasImageSource | undefined>; resolveFont: FontResolver; scale: number; scratch: SlideShowWorkspace }`. An image that does not fit the scratch allowance paints the missing-image placeholder with a diagnostic.

Rules:
- The watermark is always the top static layer.
- Underlay nodes join the first static run.
- A group and an animated descendant become a group target with a sub-target; the descendant is not painted twice.
- A paragraph build splits the shape into a body layer and range layers.

- [ ] Tests:
  - `it('static A, animated B, static C becomes three layers in authored order')`;
  - `it('translucent overlap composites identically to renderSlide at the static state')`: the pixel diff is 0, on unwatermarked fixtures (the slideshow puts the watermark on top while `renderSlide` paints it under shapes);
  - `it('grouped child animation keeps the group transform and is painted once')`;
  - `it('shape background plus paragraph build paints the background once')`;
  - `it('bldP build=p bldLvl=2 expands into one TargetKey per paragraph, level-3 paragraphs joining their level-2 parent; animBg animates the shape body')`;
  - `it('same-colour fill and text, fill-only colour change with partial transparency, changes only the fill')`;
  - `it('a layer larger than maxTextureSize rasterises at reduced scale with a diagnostic')`;
  - `it('layers are cropped to padded bounds, not full slide')`;
  - `it('paintDirect at the static state matches renderSlide pixel-for-pixel')`, on unwatermarked fixtures;
  - `it('with a watermark, an opaque shape over the watermark area still shows the watermark on top')`: the slideshow watermark oracle;
  - `it('rasterLayer leases before creating its canvas and refuses without allocating')`: a canvas-factory spy;
  - `it('the watermark layer is last and no render state moves it')`.
- [ ] Commit `feat(slideshow): z-order render plan and layer rasteriser`.

### Task S1-13: Canvas 2D compositor and masks

**Spec:** §6.3 backends (Canvas 2D fallback, no `ctx.filter`), masks, colour pipeline.

**Files:** Create `src/slideshow/render/masks.ts`, `render/canvas2d.ts`. Test: `tests/slideshow-canvas2d.test.ts`.

**Interfaces (produced):**
- it fills in S1-2's `MaskKind`;
- `maskPath(kind: MaskKind, progress: number, params, w: number, h: number, rng: Rng): Path2D | MaskRaster`, with procedural geometry shared by both backends' tests;
- `MaskRaster { width: number; height: number; alpha: Uint8ClampedArray }`, used for noise masks (`dissolve`, `randomBar`) that a path cannot express;
- `createCanvas2DCompositor(canvas: HTMLCanvasElement | OffscreenCanvas, factory: CanvasFactory, workspace: SlideShowWorkspace): Compositor`. Mask rasters and scratch canvases are leased before creation;
- `applyColorMatrix(src: ImageData, m: ColorMatrix): ImageData`, operating on un-premultiplied colour.

- [ ] Tests:
  - `it('opacity, transform and z-order of three layers match a hand-composited reference')`;
  - `it('each mask kind at progress 0 hides and 1 shows the incoming layer fully')`;
  - `it('wipe left at 0.5 reveals exactly the left half')`;
  - `it('dissolve with a fixed seed is identical across runs')`;
  - `it('colour matrix on a premultiplied semi-transparent pixel returns the expected sRGB')`;
  - `it('the compositor never touches ctx.filter')`;
  - `it('upload with an exhausted lease returns refused and creates no canvas')`: a factory spy.
- [ ] Commit `feat(slideshow): canvas 2D compositor and mask library`.

### Task S1-14: WebGL2 compositor and context loss

**Spec:** §6.3 frames, colour pipeline, backends, context loss (round 2), texture budget, determinism.

**Files:** Create `src/slideshow/render/webgl2.ts`, `render/shaders.ts`, `render/backend.ts`. Test: `tests/slideshow-backend.test.ts` (a fake GL for lifecycle), `tests/browser/slideshow-gpu.spec.ts`.

**Interfaces:**
- Consumes: S1-2 `Lease`, `TerminalReserve`, the `slideshow` Playwright project; S1-12 `paintDirect`, `LayerBitmap`; S1-13 masks.
- Produces:
  - `createWebGL2Compositor(canvas: HTMLCanvasElement, workspace: SlideShowWorkspace): Compositor | undefined`. Each texture, framebuffer and mask is leased before allocation; during upload the CPU bitmap's lease and the texture's lease are both held, and the CPU lease is released after upload. It returns `undefined` when there is no context. Uploads use `UNPACK_PREMULTIPLY_ALPHA_WEBGL=true`, blending is `ONE, ONE_MINUS_SRC_ALPHA`, and the colour matrix is applied un-premultiplied in the fragment shader;
  - `BackendManager { readonly current: Compositor; onLoss(): void; dispose(): void }`, from `createBackendManager(host: HTMLElement, deps: { now: () => number; schedule: (cb: () => void, ms: number) => unknown; paintLive: (ctx: CanvasRenderingContext2D) => void; rebuildLayers: (c: Compositor) => Promise<void>; workspace: SlideShowWorkspace; terminal: TerminalReserve; sink: DiagnosticSink })`. The loss-time 2-D canvas is painted from `terminal` (the permanent reserve), so mounting it never needs a fresh admission. The deadline uses `deps.now`, which is `performance.now` in production.

Context loss (round 2):
1. The handler calls `preventDefault()`.
2. In the same task, it mounts a separate 2-D canvas over the WebGL canvas that paints live state via `paintLive` every frame. The clock and media keep running.
3. **Restored within 3 s:** recreate every resource and switch back at the first frame whose layers are ready.
4. **Not restored:** switch to the Canvas 2D compositor on the mounted canvas.
5. More than 2 losses → Canvas 2D for the rest of the show.
6. Late restores after a switch or exit are ignored.

- [ ] Unit tests:
  - `it('loss mounts a 2-D canvas in the same task and preventDefault is called')`;
  - `it('restore at 2999 ms recreates programs, buffers, VAOs, framebuffers and textures and switches back')`;
  - `it('no restore by 3000 ms switches to canvas2d; a later restore is ignored')`;
  - `it('pausing the show does not stretch the deadline')`;
  - `it('a third loss pins canvas2d')`;
  - `it('exit during restore wait ignores the restore')`;
  - `it('texture upload leases before texImage2D and refuses without calling it')`: a fake-GL spy;
  - `it('restoration re-leases every recreated resource; a refusal drops to canvas2d with a diagnostic')`.
- [ ] Browser (both paths):
  - `test('the compositor runs on WebGL2 in CI')`, building on S1-2's smoke test. A failure blocks the stage PR and is escalated to the owner (controller ruling); there is no Canvas-2D-only fallback for CI;
  - `test('forced context loss during a transition keeps a changing animation advancing on the 2-D canvas; the frame after recovery matches the show clock')`, using `WEBGL_lose_context`;
  - `test('never restored switches at the deadline')`;
  - `test('repeated loss')`;
  - `test('exit during restore')`;
  - `test('WebGL2 and Canvas 2D frames agree within the row tolerance on a static slide')`.
- [ ] Commit `feat(slideshow): WebGL2 compositor with context-loss recovery`.

### Task S1-15: 2-D transitions

**Spec:** §6.3 transitions (2-D list), §4.5 reduced motion, §10 seeded randomness.

**Files:** Create `src/slideshow/render/transitions2d.ts`. Test: `tests/slideshow-transitions2d.test.ts`; frame goldens in `tests/slideshow-goldens.test.ts` (Mac).

**Interfaces:**
- Consumes: S1-13 masks and Canvas 2D compositor; S1-14 WebGL2 compositor; S1-3 rows.
- Produces:
  - `transitionDraw(spec: TransitionSpec, progress: number, rng: Rng): TransitionDraw['params'] & { mask?: MaskState; from?: Affine; to?: Affine; fromOpacity?: number; toOpacity?: number }`, for every 2-D element: `blinds`, `checker`, `circle`, `comb`, `cover`, `cut` (including `thruBlk`), `diamond`, `dissolve`, `fade` (including `thruBlk`), `newsflash`, `plus`, `pull`, `push`, `random` (seeded pick), `randomBar`, `split`, `strips`, `wedge`, `wheel`, `wipe`, `zoom`, and the 2-D `p14` variants (`wheelReverse`, `flythrough`, and `pan` if catalogued 2-D);
  - `reducedTransition(spec): TransitionSpec`, a fade of ≤ 200 ms.

- [ ] Tests:
  - `it('every 2-D catalogue row has a transitionDraw and both compositors render it')`;
  - `it('progress 0 shows the outgoing and 1 the incoming slide for every row')`;
  - `it('direction attributes change the geometry: push l vs r, split horz in vs vert out, wipe d')`;
  - `it('random picks deterministically per seed')`;
  - `it('reduced motion replaces any transition with a fade of at most 200 ms')`;
  - `it('cut thruBlk paints black at 0.5')`.
- [ ] Frame goldens at 0/25/50/75/100 % for `fade`, `push`, `wipe`, `split`, `circle`, `dissolve`, `wheel` (Mac).
- [ ] Commit `feat(slideshow): two-dimensional transitions`.

### Task S1-16: Workspace, fallback ladder, permanent reserve

**Spec:** §7 slideshow workspace, charged items, residency, fallback ladder, permanent reserve (round 2), reserve refusal at start, metadata/media pressure, resizing, layer and re-raster limits.

**Files:** Create `src/slideshow/workspace.ts`, `src/slideshow/residency.ts`. Test: `tests/slideshow-workspace.test.ts`.

**Interfaces (produced):**
- `SLIDESHOW_WORKSPACE_LIMIT = { desktop: 512 * MiB, mobile: 256 * MiB }`;
- `TERMINAL_SCRATCH = { desktop: 64 * MiB, mobile: 32 * MiB }`;
- `LAYER_LIMIT = 128`, `RERASTER_LIMIT = 8`;
- `createSlideShowWorkspace(profile: 'desktop' | 'mobile'): SlideShowWorkspace`;
- `admitShow(ws, css: { w: number; h: number }, dpr: number): { ok: true; backing: { w: number; h: number; dpr: number }; terminal: TerminalReserve } | { ok: false }`. It tries the capped DPR, then steps down to DPR 1. `{ ok: false }` means `present()` rejects `'load-failed'` with a `workspace` diagnostic;
- `Residency { ensure(slide: number, level: LadderLevel): Promise<LadderLevel>; evict(slide: number): void; onPressure(): void; resize(css, dpr): 'resized' | 'kept-old' }`;
- `type LadderLevel = 'full' | 'reduced' | 'flattened' | 'terminal'`.

Rules:
- The reserve is never lent.
- Under pressure: stop cross-slide media, then call `SlideMetadataStore.evict` on neighbours (S1-7). Current-slide metadata is never dropped. `createResidency` takes the store as an argument (`createResidency(ws: SlideShowWorkspace, store: SlideMetadataStore, terminal: TerminalReserve)`), so eviction removes references rather than only ledger entries. Tests use a fake store implementing the S1-2 interface.
- A resize takes the new reserve before releasing the old one.
- A completion after its slide was left releases its reservations.

- [ ] Tests:
  - `it('refuses one byte below each peak: layer+texture during upload, mirror, transition surfaces')`;
  - `it('the terminal reserve is output plus scratch and is never lent to layers, textures, metadata or media')`;
  - `it('reserve at 3840x2160 is 33177600 plus scratch; failing that steps the DPR down; failing DPR 1 refuses')`;
  - `it('ladder: reduced, then flattened, then terminal, each diagnosed')`;
  - `it('residency keeps current, previous and next; evicts neighbours first')`;
  - `it('rapid navigation: an allocation completing after its slide was left is released and discarded')`;
  - `it('resize that cannot get a larger reserve keeps the old backing store with a diagnostic')`;
  - `it('pressure stops cross-slide media, then drops neighbour metadata, never current')`;
  - `it('129 layers or 9 re-rasters in a frame coarsen with a diagnostic')`.
- [ ] Commit `feat(slideshow): workspace, fallback ladder and permanent reserve`.

### Task S1-17: Eager host API, show coordinator, start sequence

**Spec:** §4.1 start sequence and failure handling, §4.2 (all, including the show coordinator), §5.3 detection, §8 eager list, §9 focus on start.

**Files:**
- Create (eager, measured): `src/components/present.ts`.
- Modify (eager, measured):
  - `src/components/OfficeDoc.tsx`: `forwardRef`, the pre-mounted host element (always rendered and empty, `visibility:hidden`, never `display:none` at the moment `requestFullscreen` runs), new props, F5/Shift+F5 routing when focus is inside the viewer;
  - `src/components/OfficeFile.tsx`: `forwardRef` passthrough, plus the re-readable source handed to the lazy reader.
- Create (lazy): `src/slideshow/index.ts`: `runShow`.
- Test: `tests/slideshow-present.test.tsx`, `tests/browser/slideshow-start.spec.ts`.

**Interfaces (produced):**
- Exported from `src/index.ts`: `OfficeViewerHandle`, `SlideShowSession`, `SlideShowError`, `SlideShowErrorCode`, `SlideShowDiagnostic`, all as in spec §4.2.
- `class SlideShowError extends Error { readonly code: 'gesture-required' | 'not-presentable' | 'already-presenting' | 'load-failed' | 'cancelled' | 'invalid-slide' }`.
- `interface PresentPlatform { hasActivation(win: Window): boolean | undefined; requestFullscreen(el: Element): Promise<void>; exitFullscreen(doc: Document): Promise<void>; fullscreenElement(doc: Document): Element | null; open(win: Window, url: string, features: string): Window | null }`. Defaults to the real APIs. The platform is injectable through the internal prop `__presentPlatform`, which is never documented.
- `acquireShowToken(host: HTMLElement): ShowToken | undefined`, keyed by `host.ownerDocument` in a per-document `WeakMap<Document, ShowToken>`. The token is bound to `host`; `owns(el)` is true when `el` is `host` or a descendant.
- `ShowToken { release(): void; readonly released: boolean; owns(el: Element | null): boolean }`.
- `PresentOptions { from?: 'beginning' | 'current' | number; presenterView?: boolean | 'auto' }`, as in spec §4.2.
- `ReexportableSource = Blob | ArrayBuffer | Uint8Array | string`: the bytes, Blob or URL `OfficeFile` loaded from. It is undefined for one-shot sources.
- `SlideShowCallbacks { onStart?(): void; onEnd?(): void; onSlideChange?(slideIndex: number, buildIndex: number): void; onDiagnostic?(d: SlideShowDiagnostic): void; currentSlide(): number; restoreFocus(): void; scrollToSlide(i: number): void }`.
- `startPresent(req: { doc: PptxDocument | OfficeDocSource; host: HTMLElement; opts?: PresentOptions; platform: PresentPlatform; generation: number; source?: ReexportableSource; callbacks: SlideShowCallbacks; allow: boolean }): Promise<SlideShowSession>`. Its synchronous prologue runs in this order:
  1. check `allow` and kind → `'not-presentable'`;
  2. check `from` → `'invalid-slide'`;
  3. check activation → `'gesture-required'`;
  4. take the token → `'already-presenting'`;
  5. show the host and call `requestFullscreen`;
  6. only then `import('../slideshow')`.
- New props: `allowPresent = true`, `presentShortcuts = true`, `slideShowMotion: 'auto' | 'full' | 'reduced' = 'auto'`, `onSlideShowStart`, `onSlideShowEnd`, `onSlideShowSlideChange`, `onSlideShowDiagnostic`. `presenterView: true` is diagnosed and treated as audience-only until Stage 2.
- `runShow(start: ShowStart): Promise<SessionController>`, where `ShowStart { doc; handle?: PptxPackageHandle; source?: ReexportableSource; host: HTMLElement; token: ShowToken; platform; fullscreen: Promise<'entered' | 'refused'>; retryFullscreen(): Promise<'entered' | 'refused'>; from; motion; callbacks; signal: AbortSignal; generation: number; isCurrent(): boolean }` and `SessionController` implements `SlideShowSession`. `fullscreen` is the promise from the synchronous request, so a later refusal reaches the session; `retryFullscreen` must be called from a new gesture ("Click for full screen"). Implemented here as a stub that S1-19 fills.

Rules:
- Cleanup runs exactly once and in a fixed order: close the popup (Stage 2), call `exitFullscreen` only if `token.owns(fullscreenElement)`, hide the host, release the token.
- A late completion from an older generation is discarded.
- Unmount, document change or `allowPresent → false` ends the session (or cancels the start) with exit semantics.

- [ ] Unit tests (jsdom with an injected platform):
  - `it('rejects gesture-required when userActivation reports none and opens nothing')`;
  - `it('not-presentable for DOCX and for allowPresent=false; invalid-slide for 1.5, -1 and slideCount')`;
  - `it('requestFullscreen is called synchronously before the chunk import')`: assert the call order with a deferred loader;
  - `it('a second present() in the same document rejects already-presenting, from the same or another viewer')`;
  - `it('A's load fails after B entered full screen: A does not exit B's full screen and releases its token')`;
  - `it('a fullscreenchange for another element is ignored')`;
  - `it('load failure rejects load-failed, shows the notice and leaves the viewer usable; cancel by Esc, unmount, document change and allowPresent=false rejects cancelled')`;
  - `it('a late chunk completion from an older generation is discarded')`;
  - `it('fullscreen resolving after runShow starts sets surface fullscreen; rejecting after it starts switches to in-window with a diagnostic; either outcome for a stale generation is ignored')`: through the real `startPresent` → `runShow` handoff with a deferred platform;
  - `it('a re-readable source reaches runShow; a one-shot source does not')`;
  - `it('onSlideShowEnd fires exactly once before ended resolves; exit() is idempotent')`;
  - `it('presenterView true before stage 2 is diagnosed and audience-only')`;
  - `it('the host shows a black letterbox with a busy indicator until the first paint, and hides it after')`;
  - `it('F5 and Shift+F5 inside the viewer start the show; presentShortcuts=false disables them; outside focus is ignored')`.
- [ ] Browser:
  - `test('present from the API inside a click enters full screen')`;
  - `test('present() from page.evaluate without a gesture rejects gesture-required')`;
  - `test('cold start: a delayed chunk still enters full screen first; a failed chunk closes everything')`, using a routed slow or failed chunk;
  - `test('iframe without allow=fullscreen runs in-window with a notice')`;
  - `test('two viewers on one page: the second present is rejected; browser Esc exits only the owner')`;
  - `test('unmount during load and document change during a show')`.
- [ ] Ledger: eager delta ≤ `T1`; running total ≤ 38,647.
- [ ] Commit `feat(slideshow): host handle, show coordinator and synchronous start`.

### Task S1-18: Present button and toolbar routing (lazy)

**Spec:** §4.1 toolbar, §4.1 busy state, §5.3.

**Files:** Modify `src/components/paged/PagedToolbar.tsx`, `src/components/paged/PagedChrome.tsx` (lazy), and `src/components/Icon.tsx` (add the Lucide `presentation` icon; it is already a dependency). Test: `tests/paged-present-button.test.tsx`, plus a case in `tests/browser/paged-toolbar.spec.ts`.

**Interfaces:** Consumes S1-17's `onPresent(from: 'beginning' | 'current')`, passed into `PagedChromeProps` as `present?: { busy: boolean; start(from: 'beginning' | 'current'): void; preload(): void }`.

- [ ] Tests:
  - `it('PPTX shows Present with From beginning and From current slide; Presenter View is hidden before stage 2')`;
  - `it('Present is hidden with showToolbar=false and allowPresent=false; never shown for DOCX or XLSX')`;
  - `it('pointer enter, focus and menu open call preload once')`;
  - `it('busy state while the chunk loads')`.
- [ ] Browser: `test('Present button starts the show from the current slide (browser)')`.
- [ ] Commit `feat(slideshow): Present button in the paged toolbar`.

### Task S1-19: Session controller, navigation shell, audience surface

**Spec:** §4.1 exit, end-of-show and after-exit; §4.2 session and starting slide; §4.3 key table and precedence (not hit testing); §4.4 hidden, loop, kiosk input, pause; §6.2 hidden tab; §7 slide size; §8 lease fallback; checkpoint 1 (static slides and cuts).

**Files:** Create `src/slideshow/player/session.ts`, `player/keys.ts`, `player/surface.tsx`, `player/order.ts`. Modify `src/slideshow/index.ts`. Test: `tests/slideshow-session.test.tsx`, `tests/browser/slideshow-nav.spec.ts`.

**Interfaces:**
- Consumes: S1-17 `ShowStart`, `ShowToken`; S1-7 `readShowIndex`, `createSlideMetadataStore`, `openShowPackage`; S1-8 `readSlideTiming`, `shapeResolver` (passed into the store, which makes this the model assembly point); S1-10 `createSlideTimeline` (for checkpoint entry on Previous); S0-4 `validateSlideSize`.
- Produces:
  - `SessionController implements SlideShowSession { dispatch(input: ShowInput): void; readonly state: SessionState }`;
  - `ShowInput = { kind: 'next' | 'previous' | 'first' | 'last' | 'exit' | 'black' | 'white' | 'pause' } | { kind: 'goto'; index: number } | { kind: 'digit'; d: number } | { kind: 'enter' } | { kind: 'click'; x: number; y: number }`;
  - `keyToInput(e: KeyboardEvent, buffer: string): ShowInput | undefined`, following the §4.3 table;
  - `visibleOrder(index: ShowIndex): number[]`, which honours hidden slides and `sldRg`;
  - `gestureToInput(e: WheelEvent | TouchEvent, state: GestureState): ShowInput | undefined`: wheel down/up → next/previous (one step per 120 delta units, debounced on the injected clock), horizontal swipe left/right ≥ 50 CSS px → next/previous;
  - `firstSlide(model, from): number | 'end'`.

Rules:
- Precedence: Esc first, then the number buffer, then black/white resume, then the table.
- Kiosk disables everything but Esc.
- Hidden document → clock paused.
- End-of-show slide unless `loop`.
- After exit, the viewer scrolls to the last shown slide and focus returns to the start element if it is connected, else to the viewer root.
- An invalid slide size → `'not-presentable'` with a `slide-size` diagnostic, before any allocation.

- [ ] Tests:
  - `it('key table: every next and previous key; Home/End; 12 Enter goes to slide 12; B and . black, W and , white; any key resumes; Esc exits; S and + pause')`;
  - `it('Esc beats the number buffer and black screen')`;
  - `it('wheel down and swipe left go next; wheel up and swipe right go previous; a 30 px swipe does nothing; kiosk ignores both')`;
  - `it('next/previous skip hidden slides; goTo shows a hidden slide; current and number starts on a hidden slide')`;
  - `it('a deck with every slide hidden opens on end-of-show')` (Review Focus 4);
  - `it('Previous from slide n+1 enters slide n at its last checkpoint; Next from n-1 enters it at initial')`: the working assumption, pending D4;
  - `it('loop continues past the last visible slide without end-of-show')`;
  - `it('kiosk ignores clicks and keys except Esc')`;
  - `it('a hidden document pauses the clock')`;
  - `it('zero, negative, non-finite and 1:21 slide sizes, including caller-constructed documents, are not-presentable')`;
  - `it('a handle-less document presents static slides with no-package')` (Review Focus 2);
  - `it('a re-readable source is reopened when the handle is missing; a one-shot source is not')`;
  - `it('focus returns to the start element or the viewer root; the viewer scrolls to the last slide')`;
  - `it('onSlideShowSlideChange reports slide and build index')`.
- [ ] Browser:
  - `test('keyboard navigation and key precedence')`;
  - `test('end-of-show click exits')`;
  - `test('hidden slides and loop')`;
  - `test('focus restore after an API-started show')`;
  - `test('leaving full screen through the browser exits the show')`.
- [ ] Commit `feat(slideshow): session controller and audience navigation`.

### Task S1-20: Engine wiring, interactions, accessibility

**Spec:** §4.3 click targets, hyperlinks, cursor; §4.4 `advClick`; §4.5; §6.3 resize/DPR; §9 (all); checkpoints 2–8 integrated.

**Files:** Create `src/slideshow/player/stage.ts` (per-slide engine: timeline, plan, residency, compositor, transitions), `player/hit-test.ts`, `player/links.ts`, `player/a11y.tsx`, `player/cursor.ts`. Modify `player/session.ts`. Test: `tests/slideshow-interaction.test.tsx`, `tests/browser/slideshow-play.spec.ts`.

**Interfaces:**
- Consumes: S1-19's `SlideMetadataStore` (assembled with S1-8 `readSlideTiming`) and S1-8 `expandBuilds` output (`ShowSlide` builds); S1-10 `createSlideTimeline`, `createAutoAdvance`, `KioskTimer`; S1-11 `renderStateAt`; S1-12 `buildRenderPlan`, `rasterLayer`, `paintDirect`; S1-14 `createBackendManager`; S1-15 `transitionDraw`; S1-16 `createSlideShowWorkspace`, `admitShow`, `Residency`.
- Produces:
  - `hitTest(plan: RenderPlan, state: RenderState, x: number, y: number, interactive: ReadonlySet<string>): string | undefined`. It walks top-down with current transform, clip and **visibility**;
  - `runLink(link: LinkSpec, session, win: Window): void`. URLs open with `noopener`;
  - `ShowA11y`: a live region, an offscreen visible-text list and trigger/link buttons.

Rules:
- A trigger beats a link on the same shape (pending D4).
- Kiosk links are inert (pending D4).
- `advClick="0"` refuses mouse clicks only.
- The cursor hides after 3 s without movement, on the injected clock.
- On a DPR or size change, re-rasterise current/previous/next while the old layers keep painting scaled.

- [ ] Tests:
  - `it('a click on a trigger runs its interactive sequence instead of advancing')`;
  - `it('a click on an exited trigger shape advances instead of firing')` (Review Focus 1);
  - `it('a shape with both trigger and link runs the trigger')`;
  - `it('hyperlink slide jump, endshow and url with noopener; refused schemes diagnose')`;
  - `it('advClick=0: click does not advance, Space does')`;
  - `it('live region announces Slide n of m: title once per change')`;
  - `it('offscreen text excludes not-yet-entered and exited text')`;
  - `it('Tab reaches trigger and link buttons in z-order; Enter activates')`;
  - `it('cursor hides after 3000 ms of no movement and returns on move')`;
  - `it('reduced motion: transitions crossfade within 200 ms, steps jump to end states')`;
  - `it('slideShowMotion full overrides prefers-reduced-motion')`;
  - `it('resize re-rasterises neighbours; old layers paint scaled until ready')`;
  - `it('watermark appears on the audience surface above animated layers')`;
  - `it('allowCopy=false keeps the offscreen text readable but not selectable; allowPrint=false keeps slideshow surfaces out of print')`: the viewer's protection policy on slideshow surfaces (§6.3).
- [ ] Browser:
  - `test('animations play in build order on the WebGL2 and forced Canvas 2D paths')`;
  - `test('keyboard activation of triggers and links; an animated link')`;
  - `test('no premature text disclosure')`;
  - `test('reduced motion')`;
  - `test('terminal fallback under a tight budget reaches direct-paint cuts without stalling')`, via an injected small workspace.
- [ ] Commit `feat(slideshow): engine wiring, interactions and accessibility`.

### Task S1-21: D4 capture checklist, refs layout, frame comparison

**Spec:** §11 D4, §10 frame goldens and tolerances.

**Files:**
- Create:
  - `docs/validation/slideshow-captures.md`, the checklist:
    - per corpus deck: slide, step, effect and timestamps at 0/25/50/75/100 %;
    - the behavioural questions: `spd` durations, `advTm` with pending clicks, trigger-versus-link precedence, kiosk restart, whether hyperlinks and triggers respond in kiosk mode, Previous into a slide, and `spd="med"` beside `p14:dur="700"`;
  - `tests/fixtures/slideshow-refs/README.md`, the layout `<deck>/<slide>-<step>-<pct>.png` plus `provenance.json` `{ powerpoint: string; os: string; resolution: [number, number]; capturedBy: string; date: string }`;
  - `scripts/slideshow-compare.ts`.
- Test: `tests/slideshow-compare.test.ts`.

**Interfaces (produced):**
- `compareFrames(a: ImageData, b: ImageData, tol: Tolerance): { pass: boolean; meanAbs: number; rowMax: number; worstRow: number }`;
- `readRefs(dir: string): Promise<RefSet>` and `writeCompareRecord(path: string, rows: CompareRecord[]): void`, where `CompareRecord { row: string; deck: string; slide: number; step: number; pct: number; pass: boolean; meanAbs: number; rowMax: number }`. A row becomes `verified` only from a passing record in `docs/validation/slideshow-compare.json`.
- The executable CLI that renders our frames through the engine moves to S1-22, which runs after the engine is integrated.

- [ ] Tests:
  - `it('identical frames pass; a one-row stripe fails rowMax but passes meanAbs')`;
  - `it('approximate rows use the looser tolerance')`;
  - `it('a missing ref leaves the row unverified, never failed')`;
  - `it('the checklist lists every corpus deck and every behavioural question')`.
- [ ] Commit `test(slideshow): PowerPoint capture checklist and frame compare tool`.

### Task S1-22: Corpus conformance and frame benchmark

**Spec:** §13 Stage 1 done-when, §7 frame work, §10 browser list remainder.

**Files:** Create `tests/slideshow-corpus.test.ts`, `tests/browser/slideshow-perf.spec.ts` (runs only with `OV_SLIDESHOW_BENCH=1`; not in CI), and `scripts/slideshow-compare.ts`: CLI `bun scripts/slideshow-compare.ts <deck> [--rows id,…]`, which renders our frames through the lazy engine at the checklist times and uses S1-21's `compareFrames` and `writeCompareRecord`.

- [ ] Tests:
  - `it('tdf104788 plays six steps in file order with file durations')`: numeric render-state oracles at each step's 50 %;
  - `it('tdf44223 split transition, bldP animBg and entrance preset 2')`; `split`'s sound is a stage-2 stand-in, diagnosed;
  - `it('tdf142915 auto-advances after 1000 ms and ignores clicks')`;
  - `it('bnc591147 trigger runs its interactive sequence; mediacall commands are stage-2 stand-ins; the avi shows its poster')`;
  - `it('airplane and morph decks play their catalogued pending-stage-3 stand-ins with diagnostics')`;
  - `it('tdf156808 plays the flythrough effect its catalogue row names (2-D, slow, 800 ms), not a stand-in')`: assert the effect identity and render state at 50 %;
  - `it('every corpus deck presents without throwing and every diagnostic code is catalogued')`.
- [ ] Benchmark: tdf104788 at 1920 × 1080 CSS px, DPR 2, Chromium, on the owner's reference laptop. Record p95 frame work (target ≤ 12 ms) and dropped frames in the validation doc. A miss is recorded and escalated, not hidden.
- [ ] Commit `test(slideshow): corpus conformance and frame benchmark`.

### Task S1-23: Stage 1 close

**Files:** `docs/validation/slideshow.md`, `README.md` (a "Slideshow" section: Present, keys, host API and gesture rule, props, events, embedding requirements per §5.3, the known limits macOS F5, `p:browse` and custom shows), `CHANGELOG.md`.

- [ ] Set `LANDED_STAGE = 1`. The catalogue table regenerates, so every row is exact, approximate, pending-stage-2/3 or never.
- [ ] The validation doc covers:
  - gates;
  - the eager ledger;
  - every verified-in-parser value;
  - the WebGL2-in-CI result;
  - host requirements;
  - known limits;
  - manual checks (real transient activation, real full screen on a Mac);
  - the benchmark.
- [ ] `it('README slideshow props and error codes match the exported types')`.
- [ ] Full gates as S0-8, plus `--project slideshow`.
- [ ] Commit `docs: slideshow stage 1 validation, README and CHANGELOG`.
- [ ] Open the PR "Slideshow stage 1: core show". Merge after CI and the Codex bot review.

---

## Stage 2: Presenter View and media

### Task S2-1: Eager byte recovery (Stage 2)

**Spec:** §8, §11 D3, §4.1 (popup step), §5.1 preload.

1. Re-measure on main.
2. Measure `T2` with stubs of:
   - the popup branch in `startPresent`: `platform.open` before `requestFullscreen`; `null` or closed means blocked;
   - `presenterView` resolution (`'auto'`);
   - Alt+F5 routing;
   - the `window-management` preload hint (`navigator.permissions.query` then `getScreenDetails` when granted, called from the toolbar's `preload`), cached in a per-document eager slot `preparedPlacement: { popupFeatures: string; audienceScreen?: ScreenDetailed } | undefined` so the synchronous prologue can open the popup with placement features and call `requestFullscreen({ screen })` without awaiting;
   - `PresentPlatform.requestFullscreen(el, options?: FullscreenOptions)`.
   Revert the stubs.
3. Hunt ≥ `T2` with the same rules. Fall-short → escalate.

- [ ] Ledger and unchanged suite.
- [ ] Commit `perf(eager): recover bytes ahead of the presenter popup start`.

### Task S2-2: Presenter popup lifecycle and CSP host check

**Spec:** §4.1 steps 1–2 and popup refusal, §5.1 second window, closing, single-screen fallback, §5.3 (all rows), §7 popup boundary, §9 Presenter View focus.

**Files:**
- Modify (eager, measured): `src/components/present.ts` (`PresentPlatform.open` used; `requestFullscreen(el, options?)`; `screenDetails?(): Promise<ScreenDetails | undefined>`; the `preparedPlacement` slot filled by the preload hint), `src/components/OfficeDoc.tsx` (Alt+F5).
- Create: `src/slideshow/presenter/window.ts`.
- Modify: `src/slideshow/player/session.ts`.
- Create: `src/debug/slideshow-csp.html`, a harness page with a strict CSP.
- Test: `tests/slideshow-presenter-window.test.tsx`, `tests/browser/slideshow-presenter.spec.ts`.

**Interfaces (produced):**
- `PresenterWindow { readonly doc: Document; readonly root: HTMLElement; close(): void; onClosed(cb: () => void): void }`, from `adoptPresenterWindow(popup: Window, opener: Window, sink: DiagnosticSink): PresenterWindow`. It:
  - writes a theme-root container;
  - injects officeview's stylesheet (`link` to the same `style.css` URL, or an adopted copy of the `CSSStyleSheet`). On a CSP refusal it diagnoses `presenter-styles-blocked`, and the controls stay usable;
  - attaches listeners on the popup's own document and window;
  - closes on opener `pagehide`.
- Session `presenter: 'window' | 'swap' | 'none'`.

Rules:
- Closing the popup falls back to single-screen and does not end the show.
- Exit closes the popup.

- [ ] Tests:
  - `it('popup is opened before requestFullscreen in the same synchronous prologue')`;
  - `it('null or already-closed popup is blocked: single-screen with popup-blocked')`;
  - `it('closing the presenter window keeps the show in single-screen')`;
  - `it('exit and opener pagehide close the popup')`;
  - `it('notes and labels are text nodes; deck text with <script> renders literally')`;
  - `it('a failed load closes the provisional popup')`.
- [ ] Browser:
  - `test('popup-then-full-screen in one activation; a refused full screen shows Click for full screen and retries on the next click')`;
  - `test('iframe refusing both fullscreen and popups runs in-window single-screen')` (Review Focus 5): assert two diagnostics and a resolved promise;
  - `test('CSP-hardened host: presenter controls render and work; the styles diagnostic appears when blocked')`;
  - `test('both views stay in sync on next and previous')`.
- [ ] Ledger: delta ≤ `T2`.
- [ ] Commit `feat(slideshow): presenter window lifecycle`.

### Task S2-3: Window Management placement

**Spec:** §5.1 placement in Chromium, other browsers, denial diagnostics.

**Files:** Create `src/slideshow/presenter/screens.ts`. Test: `tests/slideshow-screens.test.ts`, plus cases in `tests/browser/slideshow-presenter.spec.ts` with feature-detection mocks (the controller runs them).

**Interfaces (produced):**
- `planPlacement(details: ScreenDetailsLike | undefined): { popupFeatures: string; audienceScreen?: ScreenLike } | { manual: true; reason: 'api-missing' | 'permission-denied' | 'single-screen' }`;
- `useBothScreens(win: Window): Promise<ScreenDetailsLike | undefined>`, which asks for permission on the click; the second full-screen request with `{ screen }` happens on that or a later gesture via `ShowStart.retryFullscreen`;
- `preparePlacement(doc: Document, details: ScreenDetailsLike | undefined): void`: computes `planPlacement` from the preload hint and stores it in S2-1's eager `preparedPlacement` slot;
- `swapDisplays(session): void`.

- [ ] Tests:
  - `it('granted with two screens: popup on primary, audience requestFullscreen({screen}) on the external screen')`;
  - `it('not yet granted: start proceeds and Use both screens asks on click')`;
  - `it('denied, missing API and single screen are three distinct diagnostics, none is popup-blocked')`;
  - `it('swap displays exchanges the screens')`;
  - `it('other browsers show the one-time drag hint')`.
- [ ] Commit `feat(slideshow): screen placement for presenter and audience`.

### Task S2-4: Next-checkpoint preview and mirror

**Spec:** §5.2 current and next.

**Files:** Create `src/slideshow/presenter/preview.ts`. Test: `tests/slideshow-preview.test.ts`.

**Interfaces:**
- Consumes: S1-10 `SlideTimeline.checkpoint`, S1-11 `renderStateAt`, S1-12 `paintDirect`, S1-16 `SlideShowWorkspace` (all merged in Stage 1).
- Produces:
  - `nextCheckpoint(model: ShowModel, order: number[], slide: number, cursor: SlideCursor): { slide: number; cursor: SlideCursor } | 'end'`;
  - `paintPreview(ctx, target, assets): void`: effect-free with respect to live playback (it never touches the live scheduler, media or trigger history), but not resource-free;
  - `mirrorFrame(source: HTMLCanvasElement, ctx): void`: a scaled copy of the audience output, using the same render state;
  - `PresenterSurfaces { mirror: HTMLCanvasElement; preview: HTMLCanvasElement; grid(i: number): HTMLCanvasElement; resize(css, dpr): boolean; dispose(): void }`, from `createPresenterSurfaces(doc: Document, workspace: SlideShowWorkspace)`. Its leases (`presenter:*`) are separate from the terminal reserve and owned by the presenter lifecycle (S2-5 creates and disposes it). A refused lease drops to a lower DPR, then to a placeholder with a diagnostic.

- [ ] Tests:
  - `it('next is the next build of this slide, else the next visible slide checkpoint 0, else end')`;
  - `it('zero-duration effects applied; indefinite at first-iteration end')`;
  - `it('computing next leaves the live cursor, events and media untouched')`;
  - `it('the mirror shows the in-progress animation from the same state')`;
  - `it('presenter surfaces lease on open, re-lease on resize, release everything on dispose, and never touch the terminal reserve')`;
  - `it('a refused presenter lease falls back to a lower DPR, then a placeholder')`;
  - `it('preview and mirror paint the watermark on top')`.
- [ ] Commit `feat(slideshow): presenter next preview and mirror`.

### Task S2-5: Presenter View UI and single-screen swap

**Spec:** §5.2 (all), §5.1 single-screen mode, §9 Presenter View.

**Files:** Create `src/slideshow/presenter/PresenterView.tsx`, `presenter/notes.tsx`, `presenter/timer.ts`, `presenter/grid.tsx`, `presenter/diagnostics.tsx`. Modify `src/components/paged/PagedToolbar.tsx` (show the "Presenter View" item). Test: `tests/slideshow-presenter-view.test.tsx`.

**Interfaces:**
- Consumes: S2-2 `PresenterWindow`, S2-4.
- Produces:
  - `PresenterView(props: { session: SessionController; host: 'window' | 'swap'; doc: Document })`, rendered into the popup by a portal from the opener;
  - `ElapsedTimer { elapsed(t: number): number; pause(t): void; reset(t): void }`, on the injected clock.

- [ ] Tests:
  - `it('shows current, next, notes, elapsed, wall clock and Slide n of m')`;
  - `it('A+ and A- change the notes size; notes keep bullets and paragraph breaks')`;
  - `it('timer pause and reset on the injected clock')`;
  - `it('controls: previous, next, black, end show; grid jumps to a slide including hidden')`;
  - `it('diagnostics panel lists approximate, pending and refused effects once')`;
  - `it('keys typed in the slide-number field stay in the field except Esc')`;
  - `it('Alt+F5 and the toggle swap the full-screen surface between audience and presenter')`;
  - `it('labels and focus management inside the popup document')`;
  - `it('allowCopy=false makes notes unselectable; allowPrint=false hides the presenter popup in print; grid thumbnails carry the watermark')`.
- [ ] Browser: `test('single-screen Presenter View swap')`.
- [ ] Commit `feat(slideshow): presenter view and single-screen swap`.

### Task S2-6: Media reader and leases

**Spec:** §6.1 media, §7 media reads, Blob URLs, cross-slide retention, charging, pressure.

**Files:** Create `src/slideshow/media/read.ts`, `media/lease.ts`. Test: `tests/slideshow-media-read.test.ts`.

**Interfaces (produced):**
- `MEDIA_PART_LIMIT = 512 * MiB`;
- `openMediaLease(handle: PptxPackageHandle, rec: MediaRecord, ws: SlideShowWorkspace, win: Window, signal?): Promise<MediaLease | { refused: string }>`;
- `MediaLease { url: string; bytes: number; spans: number; retain(slides: number): void; release(): void }`.
  - Reads go through `readPartBounded` with `admit` growing one workspace lease chunk by chunk and `output: 'chunks'`, so the aggregate budget refuses mid-inflate.
  - The Blob is built from the chunk list; while `new Blob(chunks)` runs, both the chunks and the Blob are charged, then the chunk lease is released.
  - `release` revokes the URL exactly once.

- [ ] Tests:
  - `it('reads with the inflated counter; a forged declared size is caught at the cap')`;
  - `it('charges the inflated length; refuses one byte below the remaining budget')`;
  - `it('on the 256 MiB profile a media part larger than the remaining budget is refused at the crossing chunk, not after inflation')`;
  - `it('the Blob construction peak (chunks plus Blob) is charged, then the chunk lease is released')`;
  - `it('revokes the URL when the slide is left unless retained across slides')`;
  - `it('never uses OfficePackage.bytes')`: a spy.
- [ ] Commit `feat(slideshow): bounded media reads and leases`.

### Task S2-7: Media playback and timeline integration

**Spec:** §2.1 goal 7, §6.1 `p:cmd`/`p:audio`/`p:video`, §5.2 one audio output, §7 cross-slide sound, autoplay, media end and the failed-playback escape (round 2), retry, §6.5 codecs.

**Files:** Create `src/slideshow/media/player.ts`, `media/sounds.ts`. Modify `src/slideshow/engine/steps.ts` (media-end events), `player/stage.ts`, (media row status flips are left to S2-9). Test: `tests/slideshow-media.test.ts`, `tests/browser/slideshow-media.spec.ts`.

**Interfaces:**
- Produces:
  - `MediaController { play(nodeId: number): void; pause(nodeId): void; stop(nodeId): void; onEnd(cb: (nodeId: number, synthetic: boolean, t: number) => void): void; retry(nodeId): void }`;
  - `TransitionSound { start(slide: number): void; stopAt(slide: number): void }`.

Rules:
- **Successful playback** ends at the real clip end, or at the node's own `dur`/clip/repeat end.
- **Refusal or failure:** a synthetic end at the explicit finite `dur` from the node start, else immediately. Diagnosed.
- **Retry** plays the media only and never re-arms successors.
- Each node's end fires at most once per activation.
- A looping transition sound plays until the next sound or `p:endSnd`.
- Presenter mirror video is muted and copied from frames.

- [ ] Tests:
  - `it('a successful clip with no dur holds its successor until the clip ends')`;
  - `it('refused playback with no dur and refused indefinite media release successors at once, diagnosed')`;
  - `it('refused with dur 2000 releases at 2000 ms')`;
  - `it('a later retry does not fire successors twice')`;
  - `it('play rejection shows Click to play; the click calls play in its own activation')`;
  - `it('cmd play, pause and stop follow the timeline')`;
  - `it('a looping transition sound continues to the next slide sound or endSnd')`;
  - `it('numSld media spans slides; everything else stops on slide change')`;
  - `it('undecodable avi/wmv keeps the poster with a diagnostic')`;
  - `it('only the audience element has audio')`.
- [ ] Browser:
  - `test('media playback of a supported format')`;
  - `test('audio spanning slides')`;
  - `test('a rejected later play()')`;
  - `test('narration.pptx plays its m4a where the browser decodes it')`.
- [ ] Commit `feat(slideshow): media playback, commands and transition sounds`.

### Task S2-8: Text builds by word and letter

**Spec:** §6.5 text by word or letter, §7 layer limits.

**Files:** Create `src/slideshow/render/text-split.ts`. Modify `render/plan.ts` (range layers for `iterate` targets), and `engine/behaviours.ts` (`iterate` offsets via `tmAbs`/`tmPct`, `backwards`). Row flips are left to S2-9. Test: `tests/slideshow-text-builds.test.ts`.

**Interfaces:** Produces `splitTextBoxes(shape: PptxShape, layout: TextLayout, unit: 'wd' | 'lt'): Array<{ key: TargetKey; box: { x: number; y: number; w: number; h: number } }>`, which uses the laid-out runs and lines.

- [ ] Tests:
  - `it('by word: each word is its own layer, staggered by tmPct')`;
  - `it('by letter with tmAbs and backwards')`;
  - `it('word boxes come from the text layout runs and cover the line')`;
  - `it('more than 128 boxes coarsen to paragraph layers with a diagnostic')`.
- [ ] Commit `feat(slideshow): text builds by word and by letter`.

### Task S2-9: Stage 2 close

- [ ] Flip the Stage 2 rows (media, sounds, word/letter builds) to their final status, set `LANDED_STAGE = 2`, regenerate the table. No `pending-stage-2` rows remain.
- [ ] The validation doc covers: Presenter View and media rows, the CSP result, the manual two-screen check, host requirements with the popup sandbox row.
- [ ] README: Presenter View, `presenterView`, Alt+F5, iframe `allow-popups` and `allow="window-management"`, CSP.
- [ ] CHANGELOG.
- [ ] Full gates as S1-23. Commit `docs: slideshow stage 2 validation, README and CHANGELOG`.
- [ ] Open the PR "Slideshow stage 2: Presenter View and media". Merge after CI and the Codex bot review.

---

## Stage 3: 3-D transitions and Morph

No task edits eager code. Every task runs the budget gate and asserts that the running total is unchanged.

### Task S3-1: 3-D transition framework

**Spec:** §6.3 3-D transitions as compositor shaders, look-alikes when WebGL2 is unavailable.

**Files:** Create `src/slideshow/render/three/framework.ts`, `three/mesh.ts`, `three/registry.ts`. Test: `tests/slideshow-3d-framework.test.ts`, `tests/browser/slideshow-3d.spec.ts`.

**Interfaces (produced):**
- `Transition3D { element: string; preset?: string; draw(gl: WebGL2RenderingContext, from: WebGLTexture, to: WebGLTexture, progress: number, params: Readonly<Record<string, string>>, rng: Rng, mesh: MeshPool): void }`;
- `registerTransition3D(t: Transition3D): void`. Group tasks never edit the registry or an index: each exports `registerGeometric3D()`, `registerParticle3D()` or `registerDeform3D()`, and S3-7 wires them;
- `docs/validation/slideshow-3d-algorithms.md`: one entry per 3-D row with its algorithm (mesh, motion, camera, timing curve), its seeded parameters and a numeric oracle at 50 % (for example a vertex position or a mask coverage fraction). Reviewed by codex before S3-2 to S3-4 are dispatched;
- `MeshPool { grid(cols: number, rows: number): MeshHandle }`, where `MeshHandle { vao: WebGLVertexArrayObject; count: number }`,, with a perspective camera helper `perspective(fovDeg, aspect)`;
- `dispatchTransition(spec, compositor): 'shader' | 'lookalike'`. Canvas 2D uses the row's `lookalike` through `transitionDraw`, with a diagnostic.

- [ ] Tests:
  - `it('a registered 3-D row dispatches to its shader on webgl2 and to its lookalike on canvas2d with a diagnostic')`;
  - `it('mesh pool reuses grids and charges the workspace')`;
  - `it('progress 0 and 1 reproduce the outgoing and incoming frames for a sample shader')`.
- [ ] Commit `feat(slideshow): 3-D transition framework`.

### Tasks S3-2 to S3-4: 3-D transition groups

Each task implements one `Transition3D` per element/preset in its group, following S3-1's algorithm sheet, exports its group registration function, and adds its own browser spec `tests/browser/slideshow-3d-<group>.spec.ts` and frame goldens at 0/25/50/75/100 % (Mac, Chromium WebGL). It does not edit the catalogue, the registry or the validation doc (S3-7 does). Each runs these tests:
- `it('every row in the group has a registered shader and a lookalike')`;
- `it('direction and option attributes change the output')` (for example `prism isContent`, `isInverted`, `flip dir`, `vortex dir`, `shred pattern`);
- `it('seeded effects are deterministic')`;
- `it('each shader matches its algorithm-sheet oracle at 50 %')`.

- **S3-2 geometric** (`src/slideshow/render/three/geometric.ts`): `prism` (cube, box, rotate, orbit), `flip`, `gallery`, `doors`, `window`, `switch`, `conveyor`, `ferris`, `pan` (if not 2-D in S1-3). Commit `feat(slideshow): geometric 3-D transitions`.
- **S3-3 tile and particle** (`three/particles.ts`): `vortex`, `shred`, `fracture`, `crush`, `glitter`, `honeycomb`, `ripple`, `flash`, `reveal`. Commit `feat(slideshow): tile and particle 3-D transitions`.
- **S3-4 deformation** (`three/deform.ts`): `pageCurlSingle`, `pageCurlDouble`, `peelOff`, `airplane`, `origami`, `curtains`, `drape`, `wind`, `fallOver`, `prestige`, `warp`. Plus `it('airplane-transition.pptx plays the airplane shader')`. Commit `feat(slideshow): deformation 3-D transitions`.

### Task S3-5: Morph matching

**Spec:** §6.3 Morph matching order.

**Files:** Create `src/slideshow/morph/match.ts`. Test: `tests/slideshow-morph-match.test.ts`.

**Interfaces (produced):** `matchMorph(from: PptxSlide, to: PptxSlide, option: 'byObject' | 'byWord' | 'byChar'): { pairs: Array<{ from: TargetKey; to: TargetKey; rule: '!!' | 'name-kind' | 'text' | 'geometry' }>; fadeOut: TargetKey[]; fadeIn: TargetKey[] }`.

The rules apply in order:
1. `!!` names;
2. same name and kind;
3. same text (per `option`);
4. same preset geometry within a distance of 25 % of the slide diagonal (my call).

Each object pairs at most once.

- [ ] Tests:
  - `it('!! names pair first even across kinds')`;
  - `it('name and kind beat text; text beats geometry')`;
  - `it('byWord and byChar pair text units')`;
  - `it('unmatched objects fade out and in')`;
  - `it('one-to-one: no object pairs twice')`.
- [ ] Commit `feat(slideshow): morph object matching`.

### Task S3-6: Morph rendering

**Spec:** §6.3 Morph interpolation; catalogue "approximate".

**Files:** Create `src/slideshow/morph/render.ts`. The Morph row flip is left to S3-7. Test: `tests/slideshow-morph.test.ts`, plus a golden.

**Interfaces:** Consumes S3-5 `matchMorph`, S1-12 rasteriser, S3-1 framework. Produces `morphFrame(pairs, progress, compositor): void`. It interpolates position, size, rotation, fill and line colour and opacity, crossfades the content where the geometry differs, and works on both backends. Canvas 2D uses transforms plus crossfade.

- [ ] Tests:
  - `it('a matched rectangle interpolates position, size, rotation and colour linearly at 0.5')`;
  - `it('geometry change crossfades the rendered content')`;
  - `it('tdf149588 morph plays as approximate with a diagnostic')`;
  - `it('canvas2d morph renders without a GPU')`.
- [ ] Commit `feat(slideshow): morph transitions`.

### Task S3-7: Stage 3 close

- [ ] Flip every 3-D and Morph row from `pending-stage-3` to `approximate` with implementation `shader` (keeping each `lookalike`), wire the group registration functions into `src/slideshow/render/three/index.ts`, set `LANDED_STAGE = 3`, and regenerate the table. Run `it('no pending-stage rows remain')`.
- [ ] Gate: the eager size is unchanged from the Stage 2 merge.
- [ ] The validation doc covers: the final catalogue (every row exact or approximate with tolerance, verified only from compare records), 3-D look-alikes, Morph approximations.
- [ ] README and CHANGELOG.
- [ ] Full gates. Commit `docs: slideshow stage 3 validation, README and CHANGELOG`.
- [ ] Open the PR "Slideshow stage 3: 3-D transitions and Morph". Merge after CI and the Codex bot review.

---

## Self-review notes

**Spec coverage.**

| Spec section | Tasks |
|---|---|
| §4.1 | S1-17, S1-18, S1-19, S2-2 |
| §4.2, including the coordinator | S1-17, S1-19 |
| §4.3 | S1-19, S1-20 |
| §4.4 | S1-7, S1-10, S1-19, S1-20 |
| §4.5 | S1-10, S1-15, S1-20 |
| §5 | S2-2 to S2-5 |
| §6.1 | S0-4, S1-5 to S1-8 |
| §6.2 | S1-9 to S1-11 |
| §6.3 | S1-12 to S1-15, S1-20, S3 |
| §6.4 | S1-3, with each stage close raising `LANDED_STAGE` |
| §6.5 | S1-11 (charts and SmartArt as one target, through the plan), S2-7, S2-8 |
| §7 | S1-16, S1-19 (size), S2-6, S2-7, S1-22 |
| §8 | S0-1, S0-4, S0-7, S1-1, S1-17, S2-1 |
| §9 | S1-19, S1-20, S2-5 |
| §10 | the tests throughout |
| §11 D4 | S1-21 |

**Proportion.** The plan carries signatures, test names and decided values. No function bodies.

---

## Appendix: Plan review dispositions (round 1)

Reviews: `.superpowers/reviews/plan-codex.md`, `plan-agy.md`, `plan-opencode.md`, `plan-pi.md`. Each claim was checked against the spec (`b693077`) and the code before disposition.

| Finding | Disposition |
|---|---|
| codex I1 (allocation owners missing from readers/compositors) | Fixed: S1-2 adds `Lease`, `TerminalReserve`, `SubOwner` and the allocation rule; `readPartBounded` gains chunk-level `admit` and `output: 'chunks'` (S0-4); compositors and the backend manager take the workspace and lease before allocating (S1-13, S1-14); rasteriser leases before canvas creation (S1-12); allocation-time spy tests added |
| codex I2 (retained package cost has no owner) | Fixed in part: S0-4 adds `retainedArchiveBytes` and a gc test that the handle does not outlive the document; the validation doc reports corpus numbers. Rejected: a new eager per-document ledger. Spec §8 asks that the cost be charged to the document and recorded, not that a ledger be added, and resident PPTX has no document ledger today (only the chart workspace); adding one would be unbudgeted eager growth under D3 |
| codex I3 (deck-wide model vs evictable metadata) | Fixed: S1-2 splits `ShowIndex` from `SlideMetadataStore`; S1-7 builds the store with one lease per slide covering notes, links, media, timing, builds, formulas and paths; aggregate-cap and evict-and-reread tests added; S1-16 residency takes the store |
| codex I4 (timing foundation omits fields and events) | Fixed: S1-2 adds `syncBehavior`, `masterRel`, `afterEffect`, `xfrmType`, `rctx` and the full event vocabulary; frozen snapshot entries keep `iteration` (S1-9); held-accumulate and reader→scheduler→render-state tests (S1-9, S1-11); carried-but-unused fields are diagnosed (S1-8) |
| codex I5 (target resolution and bldP inputs) | Fixed: `TargetResolver` built from the parsed slide is injected into `readSlideTiming` (S1-8); `expandBuilds` produces `BuildOwnership`, which `buildRenderPlan` now takes (S1-12); discarded-Choice and bldLvl/animBg tests added |
| codex I6 (start/session and placement handoffs) | Fixed: host-bound `acquireShowToken(host)`; `ShowStart` carries `source`, the full-screen promise, `retryFullscreen` and `isCurrent` (S1-17) with real-handoff tests; S2-1 adds `requestFullscreen(el, options)` and an eager `preparedPlacement` slot filled by the preload hint, measured in `T2`; S2-3 `preparePlacement` |
| codex I7 (missing dependency edges) | Fixed: S0-6→S0-1, S0-7→S0-2, S1-14 and S1-15 edges, S1-19→S1-8/S1-10; the Playwright project moves to S1-2; the compare CLI moves from S1-21 to S1-22 |
| codex I8 (CI fallback contradicts WebGL2 gate) | Fixed per controller ruling: CI must run WebGL2 via SwiftShader; failure blocks the stage PR and escalates to the owner (Global Constraints, S1-2 smoke test, S1-14) |
| codex I9 (flythrough oracle contradiction) | Fixed: S1-22 asserts tdf156808 plays its catalogued 2-D flythrough; only airplane and Morph are stand-ins |
| codex I10 (preview not resource-free) | Fixed: S2-4 adds `PresenterSurfaces` with its own leases, separate from the terminal reserve, owned by S2-5; open/resize/dispose/refusal tests |
| codex M1 (shared files and merge policy) | Fixed: shared-file policy in Streams (per-task ledger files, controller-owned ledger and catalogue table, flips only in stage closes, in-wave merge order) |
| codex M2 (unstable oracles) | Fixed: S0-3 compares unzipped entries, not a zip hash (JSZip stamps current time); S1-12 static-equivalence tests use unwatermarked fixtures plus a separate watermark-on-top oracle |
| codex M3 (HandlePackage sync/async split) | Fixed: S0-4 specifies direct delegation for `has`/`entry`/`paths`, bounded overrides for materialising methods, and an aggregate cap |
| codex M4 (wheel/swipe; protection on presenter surfaces) | Fixed: S1-19 `gestureToInput` and tests; S2-4 and S2-5 watermark and protection tests |
| codex note (preset inventory source) | Fixed: S1-3 step 0 pins the exact source, version and URL before dispatch, else the owner decides |
| codex note (S1-4 too broad for agy) | Fixed: S1-4 reassigned to codex |
| codex note (S3-4 needs per-row algorithms) | Fixed: S3-1 produces a reviewed per-row algorithm sheet with 50 % oracles before S3-2 to S3-4 dispatch |
| agy 1 (readShowDeck / readSlideTiming integration) | Fixed: the S1-7 store reads each slide once and calls the injected timing reader on the same XML |
| agy 2 (`private zip` blocks HandlePackage) | Rejected: `HandlePackage` calls `super(new JSZip())` and overrides or delegates every accessor, so it never needs the base class's `zip`; changing visibility is unnecessary |
| agy 3 (Stage 3 registration collisions) | Fixed: each group exports a registration function; S3-7 wires them |
| agy 4 (S1-14 needs the Playwright project) | Fixed: project created in S1-2, which S1-14 depends on |
| agy 5 (protection on Presenter View) | Fixed: S2-5 test |
| opencode I-1 (ShowModel assembly point unnamed) | Fixed: S1-7 store is the assembly point; S1-19 injects S1-8's reader and depends on S1-8 |
| opencode I-2 (Stage 3 wave collisions) | Fixed per controller preference: catalogue flips and validation table edits move to S3-7; each group has its own browser spec file |
| opencode M-1 (S1-18 row stale) | Fixed: table row says F5 routing lives in S1-17 |
| opencode M-2 (table omits consumed tasks) | Fixed: table now mirrors Consumes (S1-19, S1-20, S2-4, S2-6) |
| opencode M-3 (S0-2 too judgement-heavy for agy) | Fixed: S0-2 reassigned to codex (the controller still runs the `gh` calls) |
| opencode M-4 (headless full screen/popup) | Fixed in part: S1-2 smoke test checks headless full screen; where unsupported, spec §10 allows emulation, so the same assertions run through an injected `PresentPlatform` and the validation doc lists them. A local-only skip was rejected because it would drop CI coverage |
| pi (S0-3 → S1-4 implicit edge) | Fixed: noted in the S1-4 table row |
| pi (verdict) | Approve; no change required |
