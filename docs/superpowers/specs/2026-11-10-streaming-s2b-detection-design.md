# Streaming Stage 2b — tracing a leaked capture (design)

**Parent spec:** `docs/superpowers/specs/2026-11-07-server-streaming-design.md`
(§9.3 audit lifecycle, §13.2 decision rule, §13.3 collusion, §13.4 key law,
§13.5 detection tool and gates, §16 T9/T10, §17 Stage 2). This document
narrows Stage 2's remaining scope; where it is silent the parent spec rules.

**Builds on:** Stage 2a (PR #65, ab499d2f): every final-raster unit a
forensic-on session receives carries that session's Layer-1 mark
(`l1-bch512x40@1`); the audit table issues token/salt/msg40 at HELLO and
records delivered layers — in memory only.

## Intent

A screenshot or phone photo of a protected page leaks. An **investigator**
gets the image and needs to know **which session it came from**. Stage 2a
makes every page traceable in principle; Stage 2b makes it traceable in
practice: the record survives, a tool reads it, and its answer is
trustworthy enough to act on.

## Who uses it, and how (owner decisions, 2026-11-10)

- The investigator **knows which document and page** the image shows. The
  tool is told; it never searches for the document.
- The tool is run **only by someone authorized**: it needs the server's
  key material and audit store, so it runs where those live (the server
  host or an equivalent operator machine). No web UI, no remote service.
- A **simple on-screen answer** is enough. No report document for HR/legal.
- The audit table **is written to disk** and survives restarts (owner
  approved this "ask first" item of parent §18).

## Goals

### G1 Durable audit store
- Audit rows persist to a local file and survive restart and crash, kept
  for the 90-day retention window (§9.3), then deleted.
- **Ordering:** a session's issuance row is durable *before* the server
  sends its first marked unit. A crash can lose an unmarked session; it
  can never leave a delivered mark without a row.
- The token counter and the set of live `msg40` values persist, so tokens
  never rewind and the collision re-draw (§13.2) still holds after restart.
- Rows additionally carry what the tool needs to reproduce the page the
  viewer saw: `docRef`, the content-policy id, and the visible-mark
  settings **including the rendered visible-mark text** (a `viewer` mark
  differs per session, and the detector's reference must match it). Every
  `(unitIndex, scale)` a session was delivered marked is recorded too, so
  the tool knows which references to build.
- Demo mode (`--demo`, generated key) keeps an in-memory store, as today.

### G2 Key generations (rotation)
- The server reads a **keyring** of master keys, each with a generation
  number; it marks with the current generation and keeps older ones so the
  tool can still read marks made under them for as long as their rows live.
- The single-key `OV_STREAM_KEY` setup keeps working as generation 1.
- The tool's answer names the key generation of the matched row (§13.5).

### G3 Detection tool
- `server/scripts/detect.ts`: given the leaked image(s), the original
  document file, the page number, the keyring and the audit store, it:
  1. re-renders the clean reference page through the same render path the
     server uses (never a screenshot), once per distinct reference the
     candidates need (policy × visible-mark text × scale);
  2. takes as candidates the audit rows for that document that were
     delivered that page marked, grouped by key generation;
  3. registers and decodes the capture with the Stage-0 detector, and
     applies §13.2's decision rule.
- **Four answers, plain text** (a `--json` flag for scripting is fine):
  - **Traced** — session id, viewer id if known, IP, start time, key
    generation, page, confidence (Z and margin).
  - **Possible collusion** — the candidate sessions for manual review;
    names no single viewer.
  - **No match** — signal present but below threshold. Never a guess.
  - **No usable signal** — flat/blank capture, page never marked, or
    registration failed (says which).
- **Every run is logged** in the audit store: when, who (OS user, plus an
  optional `--investigator` label), image hash, document, page, answer.
- The tool **never modifies** audit rows and **never prints key material**.

### G4 Calibration on production output (the Stage-2 gate)
- Re-run T9 against units marked by the **real server path** (the bytes a
  client receives), decoded by **the tool's own code path**, not a copy.
- Refit τ_abs/τ_margin over the complete decision (registration, crop and
  scale search, all candidates) at the declared table sizes
  M ∈ {2¹⁶, 2²⁰, 2²⁴}; the tool picks τ for the actual candidate count and
  **refuses** above 2²⁴.
- Gate: false attribution ≤ 10⁻⁶ per decision over the declared table,
  by the Stage-0 method (null decisions with zero false attributions, plus
  the fitted tail and union bound). State plainly that 10⁻⁶ rests on the
  fitted tail, not on direct observation.
- Re-measure the §13.5 Layer-1 survival rows end-to-end through the tool.
- Publish in `docs/validation/streaming-s2b.md` (+ `.json`).

### G5 Collusion battery on production output
- K-average K ∈ {2, 4, 6}; K-sample median K ∈ {3, 4, 5, 6}; two-stream
  differencing (§13.3).
- Gate: **innocent named = 0** in every attack. Colluder-named,
  collusion-suspected and none rates are published, not gated.

### G6 Small carry-overs
- Stage-2a deferred minor: record unmarked deliveries per unit with their
  reason (capacity / all-masked), so "this page was never marked" is an
  answer the tool can give.
- T10 rows for audit retention and key rotation.

## Capture geometry (owner decision, 2026-11-10)

- **In scope: screenshots at any browser zoom.** Registration searches a
  continuous scale range (not only {0.5, 1, 2}) plus crop ≤ 25 % per axis,
  combined with JPEG; the §13.5 rows stay the gates, and G4 adds a zoom
  battery (at least 75 %, 110 %, 125 %, 150 %, 175 %) whose survival is
  published. The scale search sits inside the calibrated decision.
- **Out of scope: phone photos and any rotated, skewed or perspective
  capture.** The tool detects a registration it cannot fit and answers
  **No usable signal — image geometry not supported**, never a guess.
  Photo tracing is a later stage of its own.

## Rulings from the builder check (my calls, 2026-11-10)

- **One decision per capture, over every candidate.** Scores from all
  reference groups and key generations are concatenated into one list;
  §13.2's rule (τ_abs, τ_margin, multi-hit ⇒ possible collusion) applies
  to that list, with τ chosen for the total candidate count. Two hits in
  different groups are a multi-hit like any other.
- **Register once, score many.** The capture is registered once per
  (content policy, scale) against that policy's reference; the pose is
  reused to score each visible-mark group. The Stage-0 detector API may
  be split into register + score to allow this. Runtime target: a run
  over 10,000 candidates finishes in ≤ 5 minutes on the server host.
- **Several images = several independent answers**, one per image, each
  logged separately. Pooling captures is out of scope.
- **Reference check.** Each marked delivery records the SHA-256 of the
  canonical reference PNG it was marked from. The tool hashes the
  reference it re-renders; a mismatch is **No usable signal — reference
  differs from what was served** (font or renderer drift), never a decode.
- **Delivery records are written before send.** `(sid, unitIndex,
  scale, reference hash)` is committed before the marked frame goes out.
  A crash after the write and before the send leaves a record of a page
  that wasn't received, which costs only a "No match". A failed write
  refuses that unit; a failed write of the issuance row at HELLO refuses
  the marked session (the existing 4009 `policy-refused` path).
- **Table size.** The count is the filtered candidates for this
  document and page; τ is taken from the next declared M at or above it.
  The 2²⁴ entry is the fitted-tail value (as Stage 0's τ ≈ 8.66); the tool
  refuses above 2²² candidates, the most the detector's matrix holds.
- **Token space is per key generation.** Each generation has its own
  2²⁴ token counter and `msg40` set; rotating to a new generation is how
  an operator recovers from token exhaustion.

## Non-goals
- Layer 2 (word-shift on command units). Command sessions stay
  `forensic-commands-unmarked` as in 2a.
- Working out the document or page from the image.
- Phone photos, rotation, skew and perspective (see Capture geometry).
- A report document, web UI, or exporting the audit table to anyone but
  the operator (parent §18 "ask first" — not asked for).
- Persisting *sessions* across restart: live sessions still end on
  restart; only the audit record survives.
- The M-rule prototype and the command-culling occlusion store (Stage 2c).

## Anti-goals (failure even if it "works")
- Naming an innocent viewer — in any attack, at any table size.
- Reporting a guess when the evidence is below threshold.
- A marked page delivered with no durable row behind it.
- Key material leaving the operator's machine, or appearing in output or logs.
- Silently decoding with the wrong key generation or wrong reference.

## Constraints
- `server/` workspace only; no change to the library's eager bundle
  (`gates:core` must stay green). Bun runtime; no new dependencies
  (`bun:sqlite` is built in).
- Codec `l1-bch512x40@1` and its parameters are unchanged; new τ values
  ship with the calibration evidence that produced them.
- Write-once rows; deletion only by retention expiry.

## Builder's calls (left open)
- Store engine and schema (SQLite via `bun:sqlite` in WAL mode is the
  expected choice: the server writes while the tool reads).
- Keyring file format and how the current generation is selected.
- Whether the store authenticates rows with `kAudit` (tamper evidence).
- CLI argument names beyond the parent §13.5 sketch; exit codes.
- Number of null decisions in calibration (Stage 0 used 2,000; more is
  better if the run time allows).
- Task split and test layout.
