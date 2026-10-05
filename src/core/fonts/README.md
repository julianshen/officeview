# Embedded PPTX fonts

Parsing is DOM-free. `parsePptx` retains requested family/variant/source-part
identity and permission-checked SFNT bytes in optional `embeddedFonts`, and
reports per-face `fontDiagnostics`. Supported inputs are SFNT, EOT and MTX.
Restricted/bitmap-only fsType and nonempty EOT RootString restrictions are
rejected. Preview/print and editable embedding permissions both permit this
read-only display pipeline.

The public `getPaintables` pipeline awaits font registration before returning
paint closures. It resolves document-scoped aliases at the measurement and
painting boundary; requested style names and text remain unchanged.
Low-level `renderSlide` remains synchronous and uses requested names unless
passed a resolver. Full DrawingML style inheritance and rich paragraph layout
are separate work.

```ts
const pages = await getPaintables(doc)
try {
  pages[0].paint(context)
} finally {
  pages.dispose()
}
```

The returned array adds idempotent `dispose()` and `fontDiagnostics` properties.
Each consumer must release its extraction when it stops using its closures.
Active consumers of the same document and registration adapter share pending
and loaded faces. The final release deletes browser FontFaceSet entries and
evicts registration state. Different documents/adapters receive unique aliases.
`OfficeDoc` owns its viewer lease; asynchronous indexing owns a temporary lease,
and cancelled/late results are released when loading finishes.

A headless caller may pass `options.registerFont`, an async callback receiving
`{ bytes, alias, descriptors, face }`. It must resolve only when that alias is
ready for measurement and painting; it may return a face-specific cleanup
function. Reuse the same callback identity to share registrations. The library
never installs system fonts or fetches fonts over the network. Browser load
failures produce additive extraction diagnostics and degrade independently.

Limits: 64 faces, 16 MiB actual streamed inflated input per face, 32 MiB decoded
output per face, 64 MiB total converted output per document. ZIP byte limits are
checked on actual chunks before caching, including dishonest declared sizes.
The isolated pinned decoder guards intermediate streams, table/glyph/metric
buffers, LZ output/window and final output before allocation. A synchronous
scoped budget also constrains the remaining document conversion allowance.
The upstream license/provenance and Officeview modifications are in `vendor`;
exact LICENSE and NOTICE are emitted to `dist/fonts` by the build.
