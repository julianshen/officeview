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

## Corpus zero local-backend diff (Stage 1b)

`env S0_RESULTS_DIR=docs/validation/streaming-s1 bun run s0:napi scripts/stream-s1/zero-diff.ts` — every Stage-0 corpus unit (57 files) at scale 1 and 2 on `@napi-rs/canvas`: direct paint vs record → `encodeUnit` → OVSD frames (gzip 9, 64 KiB fragments) → `Reassembler` → `decodeUnit` → `replayUnit`, RGBA byte-identical.

| run | zero | diff | tile | skipped |
|---|---|---|---|---|
| first (1c12a1fd) | 336 | 114 | 68 | 1 |
| round 2 (64a2c180) | 481 | 37 | 0 | 1 |
| **round 3 (GREEN)** | **518** | **0** | **0** | 1 (malformed fuzzer input) |

Causes fixed on the way, each proven with a probe: Skia samples decoded images and canvases differently when scaling (image-sourced snapshots now replay as images); colour and font read-back strings lose precision (assigned strings are recorded); pre-composed transforms round differently from Skia's step-by-step f32 composition (transform calls are recorded as ops). Recorder options used for napi: `partialAlphaSnapshots: 'allow'`, `patternSources: 'snapshot'` (node-canvas defaults: refuse / live). Assets are passed in memory; their wire form (`ASSET_DEFINE`) is Stage 1c.
