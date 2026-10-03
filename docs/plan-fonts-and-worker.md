# Plan: font fidelity, substitution visibility, off-thread parsing

Design notes for the three improvements identified after reviewing
[embedpdf](https://github.com/embedpdf/embed-pdf-viewer). These replace the
"rewrite the core in Rust + WASM" idea, which would mostly buy what item 1
buys, at the cost of porting 10,445 lines across five formats.

## What already exists (measured, not assumed)

- **A measurement seam already exists.** `MeasureFn = (text, style) => number`,
  and `layoutDocx(document, measure)` takes it. `createMeasurer(ctx)` is only one
  implementation of that seam.
- **`OfficePackage.bytes(path)`** can already read arbitrary parts, so embedded
  font files are reachable.
- `fontFamily` is a plain `string` everywhere — no font object, so fonts can be
  resolved externally without a model change.
- **11 `measureText` call sites**; **zero font-table reads** (`fontTable.xml`,
  `.odttf`, `fontKey` are untouched; only RTF's font table is read, for names).
- **No Worker**: parse and layout both run on the main thread. Layout *cannot*
  move to a worker — `measurerFromDoc()` needs a DOM canvas.
- CI excludes the pixel goldens **"for font reasons"**.
- `createMeasurer`'s cache key is `text\0family|size|bold|italic` and does not
  include the fallback chain.

## Two corrections to the previous draft of this plan

**1. Deterministic metrics do NOT make the pixel goldens portable.** A golden is
canvas → PNG → bytes. PNG bytes depend on glyph rasterization, not on where the
line breaks fell. Metrics make *layout* deterministic; glyphs still come from
whatever font the machine has. Determinism must therefore be proven with
layout-level snapshots, never pixels. Pixel portability additionally requires a
fixed rendered face — item 1c, not item 1a.

**2. The sfnt parser is not the foundation — it is the optional extra.** If the
font is present, canvas already measures it correctly. If we hand a font to
`FontFace`, canvas measures *that* correctly for free. So:

| Situation | Does a hand-written font parser help? |
| --- | --- |
| Font is present | No — canvas reads the same advances |
| We load a metric-compatible or embedded face | No — `FontFace` makes canvas correct |
| Font absent, we render a substitute | Yes, but the text then **overflows** the measured box |

The parser only pays off when we must measure a face we cannot render, which is
exactly the case that produces overflow. It is therefore demoted to the last
item, gated on a specific need.

## The central risk, restated

Real metrics are only worth having if the **same font renders**. Lay out with
genuine Calibri advances but render a substitute and the glyphs overflow the line
box computed for them — arguably worse than today, where layout matches the
substituted glyphs exactly. Every item below must be gated on the font actually
being renderable.

---

## Item 3 — Font substitution visibility (do first)

Cheapest, and it produces the data item 1 needs: which fonts a document asks
for, and which of those the reader can actually render.

**Design**

- `inspectFonts(doc): FontReport` — `{ requested, available, missing, unknown }`,
  walking the parsed model for every `fontFamily` across all five formats.
- **Availability by width probe, not `document.fonts.check()`.** `check()`
  consults font matching and can report `true` for a family that actually
  resolves to a fallback — precisely the case we need to catch. Instead measure a
  sample string in the target family and in a known fallback: identical widths
  mean the font is not really present.
- Environments without a canvas (Node, jsdom) report `unknown` rather than
  pretending.
- `onFontsMissing?: (fonts: string[]) => void` on `OfficeDoc` / `OfficeFile`,
  fired once after parse.
- README: document the substitution behaviour.

**Acceptance**

- A doc naming an absent font reports it; one naming present fonts reports none.
- A font that resolves to a fallback is correctly reported missing.
- No canvas → `unknown`, no throw.
- Golden and corpus suites unchanged (default off ⇒ zero pixel effect).

**Risk** — low. **Effort** — small.

---

## Item 1c — Metric-compatible substitution (the load-bearing piece)

This is what actually buys both determinism and visual correctness.

- Calibri → **Carlito**, Cambria → **Caladea** (SIL OFL, metric-compatible).
  Calibri/Cambria themselves are not redistributable.
- Load on demand, **only** for a family a document actually requests. Never
  bundle: a weight is roughly 700KB.
- Opt-in via a config map, so a host can point Calibri at whatever face it
  already ships.

**The wrinkle that makes this non-trivial:** `FontFace.load()` is asynchronous,
while layout runs as soon as the document parses. If layout wins the race we
measure with the fallback *and cache those widths for the document's lifetime*.
So pending font loads must be awaited before layout, or the measurer cache must
be invalidated on load. This is architectural — decide it here, not at
implementation time.

**Acceptance**

- A document requesting Calibri lays out identically with and without Carlito
  loaded, once loads are awaited.
- After a late font load, subsequent measurement uses the real face.

**Risk** — licensing (solved by the OFL clones), and the load race above.

---

## Item 1b — Embedded fonts

- Read `word/fontTable.xml` + rels for `w:embedRegular` / `w:embedBold` / etc.
- **`.odttf` files are obfuscated**: the first 32 bytes are XORed with the GUID
  from `w:fontKey`. Not optional — skip it and every embedded font parses as
  garbage.
- De-obfuscate, then register via `FontFace`. **No font parser needed**; canvas
  measures the registered face.

**Acceptance**

- A corpus document with embedded fonts renders with that font, and layout
  matches the font's own metrics.
- Un-obfuscated bytes match the original sfnt header.

---

## Item 1a — sfnt parser (demoted; gate before building)

Only build if a concrete need appears — chiefly *headless determinism in an
environment with no browser font API*.

- `src/core/fonts/` — `head` (unitsPerEm), `hhea`+`hmtx`, `cmap`, optional
  `kern`. GPOS, shaping, ligatures and RTL stay out of scope.
- Route through the existing `MeasureFn` seam so all 11 call sites benefit.
- Extend the measurer cache key to include the fallback chain.

---

## Item 2 — Off-main-thread parsing

**Gate on measurement.** Parse may not dominate, and the clone may cost more than
it saves.

### 2a. Benchmark

- Largest corpus files; report **inflate vs XML parse vs layout vs paint**
  separately, plus **structured-clone cost of the parsed model**.
- If clone cost exceeds the parse saving, stop: the worker is a net loss.

### 2b. Worker

- Seam exists: `loadOfficeFile` does `OfficePackage.load(data).then(detectAndParse)`,
  and `detectAndParse` is pure CPU with no DOM.
- Worker owns **bytes → model**. Layout and paint stay on the main thread
  (`measureText` requires a canvas).
- `new Worker(new URL('./parse.worker.ts', import.meta.url), { type: 'module' })`.
- Opt-in (`useWorker?: boolean`), default off, with the main-thread path kept as a
  tested fallback — Node/jsdom have no `Worker`.

**Acceptance**

- Benchmark numbers recorded in the PR description.
- Identical model from worker and non-worker paths.
- With the worker on, a long parse does not block input (measured).

**Risks** — clone cost; Vite worker bundling; Node/SSR must keep working;
`Worker` test doubles in jsdom.

---

## Sequencing

```
0. Benchmark, incl. clone cost ── gate for item 2
1. Item 3   font visibility ──── width probe, not fonts.check
2. Item 1c  metric-compatible substitution ── load-bearing
3. Item 1b  embedded fonts ───── de-obfuscate + FontFace
4. Item 1a  sfnt parser ──────── only for headless-without-font-API
5. Item 2   worker ───────────── only if the benchmark says so
```

## Non-goals

- Shaping, ligatures, RTL, complex scripts.
- Pixel-exact Word reproduction. The ceiling is set by font availability.
- Replacing canvas as the default measurer. It stays the fallback everywhere.
- Making the pixel goldens portable. That needs a fixed rendered face (1c) and
  is not claimed by this plan.
