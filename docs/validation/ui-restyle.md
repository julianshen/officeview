# UI restyle validation (release gate)

Plan: `docs/superpowers/plans/2026-10-18-ui-restyle.md` (Task 8). Spec:
`docs/superpowers/specs/2026-10-18-ui-restyle-design.md`.

Measured on `feature/ui-h`: base `697cbf9` (all streams A–G merged), the Task 8
docs change (`0292978`) and Task 8 fix round 1 (`319c15c`: outline Lucide icons,
select chevrons, disabled primary styling, dialog button gaps, accent focus ring;
lazy files and `officeview.css` only). Host: macOS (Darwin 25.5.0), Bun, Node
22.22.3, 2026-10-19.

## Gates

| Gate | Command | Result |
|---|---|---|
| Type check | `bunx tsc --noEmit` | exit 0 |
| Script type check | `bun run typecheck:scripts` | exit 0 |
| Unit tests | `bunx vitest run` | 391 files passed; 7,599 passed, 2 skipped, 1 todo, 0 failed |
| Build + eager budget | `bun run build && bun scripts/eager-budget.ts` | exit 0; **38,696 B** (ceiling 38,720 B, 24 B headroom) |
| CSS isolation audit | `bun scripts/audit-css.ts dist/style.css` | `ok (150 rules, all scoped)` |
| Eager graph | `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts` | `eager modules: 151`, `eager graph ok` |
| Golden corpus | `bun scripts/golden-corpus.ts compare --threshold 0 --max-ratio 0` | **Baseline-identical** (controller ruling): 32 passed, 5 mismatches that `main` shows identically, 2 skipped (see below) |
| Browser (Chromium + WebKit) | `IV_HARNESS_PORT=4193 bunx playwright test --config playwright.viewer.config.ts` | Full suite at `0292978`: exit 0; 226 tests (113 Chromium, 113 WebKit), 212 passed, 14 skipped, 0 failed. After fix round 1, the affected specs (ui-icon-fill, ui-control-polish, ui-chrome-geometry, ui-host-isolation, ui-motion, interactive-viewer-visual, -mobile, -a11y, interactive-viewer, paged-toolbar, viewer-layout-shift): exit 0; 196 tests (99 Chromium, 97 WebKit), **188 passed, 8 skipped, 0 failed** |

Vitest count note: the ledger's last combined gate recorded 7,511 tests at
`697cbf9`. This run had the real-file corpus present in `corpus/` (copied from
the main checkout and verified against `corpus/corpus.lock.json` SHA-256), so the
corpus-backed tests ran instead of skipping, plus the new `readme-tokens`,
`ui-icon-fill` and `ui-control-polish` tests.

### Golden corpus

The five mismatches against the committed goldens:

- `poi-55733.docx.0`: 305 / 861,696 px (0.035%)
- `poi-picture-transparency.pptx.1`, `.2`, `.3`: 1,024 / 629,510 px each (0.163%)
- `pydocx-blk-inner-content.docx.0`: 3,725 / 861,696 px (0.432%)

They are not caused by the restyle. The same command on a `git archive` of
`main` (`849e1d2`, the merge base) with the same corpus fails on exactly the same
five units with the same pixel counts, and the `actual` renders from `main` and
from `feature/ui-h` are pixel-identical for all five (compared with PIL). The
restyle therefore leaves corpus rendering unchanged; the five goldens disagree
with this host's rendering on `main` too. Nothing was re-recorded.

**Ruling (controller, Task 8 fix round 1):** because the five failing units are
pixel-identical to `main`'s renders and fail identically on `main`, the golden
gate counts as baseline-identical for this branch. Re-run after fix round 1:
the same five units, with the same pixel counts.

## Eager budget: B0 to final

`dist/index.js`, Node-zlib gzip after a fresh `bun run build`
(`scripts/eager-budget.ts`). Ceiling 38,720 B.

| Point | Commit | Eager gzip (B) | Delta | Source |
|---|---|---|---|---|
| B0, before the restyle | `d81beca` | 38,697 | | Task 0 commit `f206518` message |
| Task 0: paged floating controls lazy | `f206518` | 37,583 | −1,114 | Task 0 commit message |
| Stream A merged (Task 1, foundation) | `37a3f0f` | 37,753 | +170 | ledger |
| Stream B merged (Task 2, shared controls) | `7900636` | 37,753 | 0 | ledger |
| Stream C merged (Task 3, paged toolbar) | `6b9bdf4` | 38,680 | +927 | ledger |
| Stream E merged (Task 5, sheet chrome) | `527d8cb` | 38,679 | −1 | ledger |
| Stream D merged (Task 4, paged print) | `7eed355` | 38,672 | −7 | ledger |
| Stream G merged (Task 7, canvas theming) | `1f0a9cd` | 38,720 | +48 | ledger |
| Stream F merged (Task 6, motion) | `697cbf9` | 38,696 | −24 | ledger |
| Task 8 docs | `0292978` | 38,696 | 0 | this run |
| Final (Task 8 fix round 1, lazy only) | `319c15c` | **38,696** | 0 | this run |

Net over the restyle: 38,697 → 38,696 (−1 B). Lucide icons, motion, tokens and
all new paged UI are outside the eager graph (`check-eager.ts`).

## Visual review

Before/after composites were captured for DOCX, PPTX, static XLSX and the
interactive XLSX viewer (filter popup, sort dialog, mobile filter sheet, zoom
50% and 200%), in light and dark, at 1280×800 desktop and 375×667 touch:
48 composite pairs over 24 shots. "Before" is `main` (`849e1d2`), which has no
dark mode, so each shot pairs before-light with after-light and with
after-dark. The "after" shots were re-captured after fix round 1 (tree `9cbdd5e`,
identical to `319c15c`, which only rewords the commit message). The gallery is
published separately for the owner; the cross-model
visual reviews (plan Task 8 step 3) are recorded with the PR.

## Known limitations (documented in the README)

- Page tracking and Previous/Next scroll the viewer's own container, so the host
  must give the viewer a constrained height (spec §6.11).
- Print suppression with `allowPrint={false}` in the page viewer relies on
  `style.css`; without it, an instance-scoped fallback rule is injected once the
  lazy toolbar chunk loads.
- Token overrides changed only through an ancestor stylesheet reach the canvas
  overlays at the next `theme` change or remount (spec §6.2).
- Host CSS at specificity (0,2,0) or higher, or with `!important`, can still
  override the chrome (spec §6.8).
