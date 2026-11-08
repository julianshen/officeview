# @officeview/stream-server (demo)

The WS-only stream server of the server-streaming design
(`docs/superpowers/specs/2026-11-07-server-streaming-design.md`, Stage 1c).
Private workspace; it imports the library's TypeScript sources directly and
never React.

```bash
bun run stream:demo                 # --demo, serves ./corpus, ephemeral key
# ws://localhost:8787/stream?doc=pypptx-test.pptx            protected (final-raster)
# ws://localhost:8787/stream?doc=pypptx-test.pptx&relaxed=1  audited commands session (--demo only)
bun run test:server                 # bun test (napi preload via bunfig)
```

Outside `--demo` the server refuses to start without `OV_STREAM_KEY` (32 bytes,
base64) and an allowed origin (`--origin` / `OV_STREAM_ORIGINS`).

- **One endpoint.** `GET /stream` WebSocket upgrade from an allowlisted
  `Origin`; every other HTTP request is 404. There is no file route.
- **Your hooks.** `authorize(req)` decides the viewer and per-session policy;
  `resolveDocument(req, decision)` returns the bytes. `src/demo-hooks.ts` is
  the demo pair (bare file names inside one directory).
- **Protected by default.** Hidden content is removed before parsing
  (`src/policy.ts`, fail-closed); units ship as metadata-free PNGs with the
  visible mark composited. The forensic pixel mark is Stage 2: until then the
  manifest carries the `forensic-mark-stage-2` notice and the audit row
  records `forensic: false`.
- **Fonts.** `fontpack/` (20 OFL faces, `pack.json` pins their hashes);
  Calibri/Cambria/Arial/Helvetica/Times New Roman/Courier New are painted
  with their metric-compatible pack faces. Verify: `bun scripts/stream-s1/fontpack-lock.ts`.
