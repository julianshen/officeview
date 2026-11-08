# Server streaming — Stage 1 validation

Spec: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §17 row 1. Plans: `docs/superpowers/plans/2026-11-08-streaming-s1*.md`.

## T11 — §8.1 eager gates (Stage 1a)

`bun run gates:core` (OV_MODULE_MAP=1 build, then `scripts/core-gates.ts`), ea91322a:

```text
ok   core closure 299 B over 1 chunk(s) ≤ 38720
ok   engine modules in core closure: none
ok   core purity: no react imports, no document global
ok   index closure 417580 B (baseline 417436, Δ 144, recorded 144)
ok   index module set = baseline ± 1 recorded change(s) (156 modules)
```

- Baseline: `scripts/eager-baseline.json`, captured from the Stage-0 build (`feature/streaming-s0`), 417,436 B over 155 modules.
- Recorded change: the §6.1 `StreamSource` guard (+144 B, `src/stream/source.ts`).
- `scripts/eager-budget.ts` (index.js alone): 36,122 B ≤ 38,720.
- `bun run test`: 609 files, 12,304 passed, 62 skipped.
