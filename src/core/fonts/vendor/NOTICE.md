# MTX decoder provenance

Vendored from `mtx-decompressor` **1.8.0**, npm distribution `dist/index.mjs`
and `dist/index.d.ts` (renamed `mtx.mjs` / `mtx.d.mts`).
Upstream: https://www.npmjs.com/package/mtx-decompressor/v/1.8.0
This TypeScript port derives from libeot (MPL 2.0). The exact upstream
LICENSE is retained alongside this notice. No fonts ship with this decoder.

Modified file: `mtx.mjs`. Officeview adds allocation budgets before stream,
table, glyph, metrics, LZ window/output and final SFNT allocations; aggregate
stream/table budgets (including CVT expansion charged against sibling tables
before reserve/copy); glyph point limits; LZ copy/distance bounds and RLE
termination checks. Upstream decoding arithmetic is otherwise retained.
`mtx.d.mts` retains upstream declarations and adds the scoped budget helper declaration.
