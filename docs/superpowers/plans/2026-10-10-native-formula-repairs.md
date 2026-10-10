# Native formula compatibility repairs

User-approved scope: fix the five native differences in `validation/excel-formulas/2026-10-10/REPORT.md`. Missing functions and other documented unsupported profiles remain separately scoped. No commit or push is authorized by this repair request.

Architecture: preserve the formula evaluator and spill engine. Decode OOXML escaped strings at the XLSX text boundary; resolve cell metadata through its package relationship and indexed XLDAPR records, and adopt verified dynamic output cells without destroying styles or unrelated input. LEFT/RIGHT count code points independently of LEN/MID compatibility settings. General formatting uses available cell width; built-in time formats use the shared civil/time formatter.

- [x] Root: after Pi's stalled exploration produced no changes, implement metadata and escaped-string repairs locally with failing regressions first. Parser-owned saved followers are leased by exact identity/content; generated descendants restore blank styled cells. Unverified metadata retains legacy behavior.
- [x] Root: native LEFT/RIGHT regressions pass while LEN/MID retain their version distinction.
- [x] Root: width-aware General and built-in time/date-format regressions pass.
- [x] Root: direct native whole-document baseline and changed-input replay pass; a new earlier-reader/grown-spill counterexample was independently found and repaired. Pi read-only review hit an upstream HTTP 500 and was stopped; no approval claim.
- [x] Root: consistent-baseline and latest committed-baseline TypeScript, full strict suites (latest 2,899 passed, two existing skips), build and corpus report pass; both changed goldens individually reviewed. Shared-checkout Word font/type and DOCX assertion failures after the concurrent branch switch were reported to OpenCode and left untouched.

Native oracle workbooks and extraction are retained at `/tmp/officeview-excel-gap-validation-20261010/`. Preserve every unrelated pending change and concurrent commit. Validation artifacts and new regression tests must explain native profile boundaries, not claim full Excel parity.
