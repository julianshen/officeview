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
base64) or `OV_STREAM_KEYRING`, and an allowed origin (`--origin` /
`OV_STREAM_ORIGINS`). Setting both key variables is an error.

**Key rotation.** `OV_STREAM_KEYRING=<path>` names a JSON file
`{"current": 2, "keys": {"1": "<base64 32 B>", "2": "<base64 32 B>"}}`
(generations are integers 1..255). New sessions are issued under `current`;
each audit row records its `keyGeneration`. To rotate: add a new generation,
switch `current` to it, and keep every older generation until the audit rows
issued under it have expired, so leaked captures from that era stay traceable.
`OV_STREAM_KEY` alone is a keyring with generation 1.

**Audit store.** `OV_AUDIT_DB=<path>` (required outside `--demo`; `--demo`
keeps an in-memory store) names a SQLite file (`bun:sqlite`, WAL,
`synchronous=FULL`, created mode 0600; keep it on local disk and back it up
with its `-wal`/`-shm` files). It holds the write-once audit rows, one
delivery record per `(sid, unit, scale)` served to a forensic session (status
`l1` or `unmarked` with its reason, plus the SHA-256 of the canonical PNG the
mark was applied to), the per-generation token counters, and the trace tool's
detection log. Ordering: a session's row is durable before its manifest is
sent, and a delivery record before its frames; a failed write refuses (HELLO
closes 4009 `policy-refused`; a unit is refused `policy-refused`), never sends
unrecorded. Rows and their deliveries are deleted after 90 days (at startup
and hourly); token counters are never reset.

- **One endpoint.** `GET /stream` WebSocket upgrade from an allowlisted
  `Origin`; every other HTTP request is 404. There is no file route.
- **Your hooks.** `authorize(req)` decides the viewer and per-session policy;
  `resolveDocument(req, decision)` returns the bytes. `src/demo-hooks.ts` is
  the demo pair (bare file names inside one directory).
- **Protected by default.** Hidden content is removed before parsing
  (`src/policy.ts`, fail-closed); units ship as metadata-free PNGs with the
  visible mark composited. Each raster unit also carries the session's own
  invisible Layer-1 pixel mark (`l1-bch512x40@1`; `src/mark/`), added per
  session after the shared cache. The audit row records the token, salt and
  keyed message issued at HELLO, and separately the layers actually delivered
  (`forensicLayers`, counts of marked/unmarked raster units); `untraceable` is
  derived from what was delivered, never from issuance alone. Command units (`projection: 'relaxed'`) carry
  no forensic mark yet (`forensic-commands-unmarked` notice); `forensicMark:
  false` sends `forensic-off`. Audit rows live in memory until Stage 2b.
- **Fonts.** `fontpack/` (20 OFL faces, `pack.json` pins their hashes);
  Calibri/Cambria/Arial/Helvetica/Times New Roman/Courier New are painted
  with their metric-compatible pack faces. Verify: `bun scripts/stream-s1/fontpack-lock.ts`.

## Forensic mark (Stage 2a)

The render cache holds one canonical PNG per unit; a forensic-on session's
raster units then go through a per-session pool `encode` job (decode → embed
the session's 512-sign code over the page-keyed basis → lossless re-encode),
at most 2 in flight per session. Units too small to carry 512 cells ship with
the visible mark only and a `diag forensic-unmarked`. Integration proof on
served bytes (ΔE00, §13.5 survival at the Stage-0 τ, unmarked captures never
attributing): `docs/validation/streaming-s2.md`. Detection tooling, τ
recalibration and audit persistence are Stage 2b.

## The browser side (Stage 1d)

```tsx
<OfficeFile
  data={{ kind: 'stream', url: 'wss://viewer.example.com/stream?doc=…', token }}
  streamFonts={{ baseUrl: '/officeview-fontpack' }}   // optional: serve fontpack/ here
  onStreamNotice={(n) => console.log(n.code)}         // manifest notices, face-missing, resumed, unit-failed, session end codes
  streamClient={(s) => new WebSocket(s.url)}          // optional factory (the default)
/>
```

- **The mode is the branch.** A `StreamSource` (or a promise of one) never reaches `fetch` or a byte reader; the client code is a lazy chunk, so pages that only load files never download it, and local loads paint byte-for-byte as before (`docs/validation/streaming-s1.md`).
- **What a reader gets in Stage 1.** Rasters (protected) or replayed commands (audited relaxed sessions), painted at the scale bucket for the device pixel ratio. No text layer, search, copy, zoom or presentation; the view says "text unavailable in this viewing mode" once. The host `watermark` prop is ignored (the server bakes the mark); print follows the server's `can.print`.
- **Fonts.** The font pack is **not** in the npm package. Protected sessions need no fonts (every unit is a raster). For relaxed sessions, serve `fontpack/` (`pack.json` + TTFs) and pass `streamFonts.baseUrl`; each face is hash-checked against the manifest before it is registered, and any face that is missing, unverifiable or not hashed within 1.5 s is reported missing, so its units arrive as rasters.
- **Drops.** An abnormal drop (1001/1006) resumes the same session for up to 30 s (`detachGraceMs`, capped by the token's `exp`); the scroll position is untouched and pending units are re-requested. A terminal close (revoked, expired, oversized frame, …) shows "Viewing session ended (<code>)" and calls `onStreamNotice` with that code; recovery is a host re-open.
- **Test harness.** `bun run test:browser:stream` runs `tests/browser/stream-viewer.spec.ts` against `server/test/app.ts`, which adds test-only controls (drop / revoke / oversize) on a separate port.
