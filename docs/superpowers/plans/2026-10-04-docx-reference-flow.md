# Latest DOCX reference flow implementation plan

> REQUIRED: Use subagent-driven-development, with one fresh implementer and spec review followed by quality review. No commits. Start only after the foundation final fixes pass their review gates.

**Goal:** Repair source-driven header/table/paragraph flow defects found by comparing the latest saved complex.docx with native Word.

**Spec:** `docs/superpowers/specs/2026-10-04-docx-reference-flow-design.md`

**Workspace:** `/tmp/officeview-wt-drawings`. Native and before references: `/tmp/officeview-complex-20261003`. Preserve reviewed drawing code and concurrent main work.

## Task 1: Source style and block flow corrections

**Files:** `src/docx/{types.ts,parse.ts,styles.ts,layout.ts}`; existing DOCX tests and a focused regression test file if useful. Temporary diagnostics/references stay in `/tmp`.

- [x] Inspect the frozen package and retained native references; reproduce dropped header tables, invisible first-row text and omitted default spacing independently.
- [x] Establish failing deterministic tests for ordered header/footer blocks and dual-model precedence (including empty authoritative blocks), document/paragraph/table style cascade, conditional cell properties, exact row heights, cell paragraph spacing, image indices/fields, field-induced relayout and header/footer occupied bounds/oversized-content termination.
- [x] Implement source-defined style cascade without mutating inherited definitions. Keep direct spacing zero and false toggles authoritative; resolve conditional styles by grid ownership and enabled table look.
- [x] Parse ordered header/footer blocks with part-local image/drawing relationships and inherited/default/first-page variants. Preserve paragraph-array consumers.
- [x] Preserve cached content inside transparent content-control wrappers at block/row/cell boundaries; independently test the source's wrapped first-row cells and their order.
- [x] Reuse block layout for header/footer tables, fields, images and text. Honor exact tiny row heights and cell-specific margins. Reserve body space from occupied header/footer bounds using source-defined positioning established by native reference evidence. Use at most eight field/pagination relayout passes with conservative maximum observed bounds, a fixed repeated-content overflow viewport and diagnostics at the limit. Oversized repeating blocks cannot create pages; retain a general minimum body line band and guarantee forward progress.
- [x] Render the frozen current complex snapshot and diagnose pagination from source properties. Compare all three native pages and feature inventory. Do not hardcode three pages or modify source Office files.
- [x] Run targeted DOCX tests and TypeScript, self-review, and report exact changes and any evidence-backed residual limitations. Spec review then quality review must approve before integration.

## Task 2: Native and full integration gate

- [x] Root independently renders the frozen snapshot and compares all three retained native pages, headers/body table and feature presence.
- [ ] Refresh the actual Canvas comparison in Dia when the desktop window becomes available; current app selection returns cgWindowNotFound. Native/source and code gates remain independently verifiable.
- [x] Inspect changed golden images; intentionally update only source-corrected baselines after native/source validation.
- [x] Run `bunx tsc --noEmit`, `OFFICEVIEW_STRICT_GOLDEN=1 bun run test`, `bun run build`, `bun scripts/corpus-report.ts`, `git diff --check`.
- [x] Record current input hash and native/reference metrics, distinguishing the previous externally changed input. Finish final review and guarded integration with no commit, then continue the focused NATS corrections before approved drawing stage 2.

Spec and quality gates approved after reproduced fixes for cell-anchored repeated drawings, empty story reservations, current-section continuous flow, and physical margin reference offsets/ranges. Root strict suite729/729,42files; tsc/build/corpus/diff passed. Only two source-corrected golden fixtures changed; hashes and independent native/source evidence retained. Final review approved; 11 changed/new files copied with source/destination fingerprint guards, one identical spec skipped, main HEAD unchanged. A refreshed live Dia check remains pending cgWindowNotFound; retain that limitation and continue unaffected authorized work after native/code review gates.

Final review also caught an explicitly finalized empty terminal section; parser/page finalization now retain its own header/footer and legitimate page while excluding unused builders. Spec/quality/final re-reviews approved109DOCXtests. Main integration is uncommitted.
