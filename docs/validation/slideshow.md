# Slideshow validation

Validation record for the PowerPoint slideshow project (spec
`docs/superpowers/specs/2026-10-21-slideshow-design.md`, plan
`docs/superpowers/plans/2026-10-21-slideshow.md`). Created by task S0-1 as a stub;
later tasks fill in their sections. The controller merges the per-task ledger files
(`.superpowers/sdd/ledger/<task-id>.md`) into the Eager ledger and Catalogue sections
in merge order; implementer tasks do not edit those two sections.

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

Pending (S1-3).

## Fixtures

Pending (S0-2: LibreOffice corpus, pinned commit SHA and provenance).

## Host requirements

Pending.

## Known limits

Pending.

## Manual checks

Pending.
