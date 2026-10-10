# Formula compatibility implementation plan

User authorized implementing the reviewed fixes and filling named gaps. Keep main drawing work intact. Work in isolated worktree; no main commits/push until root review/release. Architecture: preserve public XLSX primitives and typed-error separation; bounded explicit AST frames, Excel-context-aware argument coercion; reference metadata and bounded matrices for subsequent functions/arrays. No eval/newFunction, no new dependencies without a concrete need. Existing formulas and caches remain the baseline.

Stage A:All confirmed defects.
- [x] Add native-backed RED cases from review/COMPARISON.json for supported formulas; correct old tests with disproved expectations.
- [x] Remove uncaught AST stack overflows on Node/V8; preserve lazy IF/IFERROR and dependency suspension/resume; bounded cell errors for policy limits.
- [x] Preserve sheet whitespace identity and quote shared names using lexer-safe grammar.
- [x] CorrectCOUNT error/empty-text behavior,COUNTA empty-text counting,explicit omitted args versus blank references,selected IF blanks/scalar blank results,blank/FALSE comparison,and Excel numeric-text coercion.
- [x] CorrectROUND decimal midpoints/extreme digits,INT premature rounding,0^0,MOD native floating behavior; check more matching controls before replacing numeric policies.
- [x] Distinguish empty numeric caches from valid empty strings/false; preserve unsupported formula caches on file-driven recalc with diagnostics; never replace unsupported shared formula source with diagnostic prose.
- [x] Support version-aware Unicode length/slicing with preserved legacy behavior and document-mode metadata where available.
- [x] Spec then quality review; root independent native/node probes; formula tests,tsc,full strict tests/build/corpus. Freeze and retain exact proofs.

Stage B:Reference foundation and conditional functions.
- [x] Named expressions with workbook/local scope,whole rows/columns,structured table refs,3D/union/intersection refs; translation preserves source. Sparse/bounded references must not change counts/positional semantics by clamping away blanks.
- [x] SUMIF/SUMIFS,COUNTIF/COUNTIFS,AVERAGEIF/AVERAGEIFS,IFNA,IFS,SWITCH and predicates; Excel wildcard/escape/coercion/error rules; lazy outcomes and native controls.

Stage C:Lookup/date/text/array features.
- [x] INDEX,MATCH,VLOOKUP,HLOOKUP,XLOOKUP,XMATCH; exact/approx/wildcard/search modes with native cases.
- [x] DATE,DATEVALUE,YEAR,MONTH,DAY,TIME,TIMEVALUE,HOUR,MINUTE,SECOND,DAYS,EDATE,EOMONTH,WEEKDAY,TODAY/NOW;1900/1904 systems and fictitious1900-02-29; TEXT formats using existing formatter where appropriate.
- [x] Array constants,elementwise evaluation and bounded spills;FILTER,SORT,SORTBY,UNIQUE,SEQUENCE,TRANSPOSE;modern errors,#/@ refs,collision/edge limits and legacy array behavior; scalar API remains compatible.
- [x] Respect workbook manual/iterative calculation settings or retain caches with explicit diagnosed policy rather than silently generating wrong values.
- [ ] Expand finite common financial/statistical families if requested; maintain a precise support matrix without claiming the full catalog before implementation/verification.

For every stage:real RED/GREEN,root-owned external expected values,independent SPEC then QUALITY review,shared-work protection,and final source/hash guard. No formulasource privacy/provider/memory/config changes. Do not modify main PLAN.md or drawing paths.

Status: the reviewed Stage A/B/C implementation scope is complete and independently accepted through C4. The reviewed formula changes were integrated into main while preserving concurrent drawing/font work. Final checks passed: TypeScript, production bundle and declarations, strict tests (2,984 passed; 2 existing native PowerPoint fixture skips), and corpus report (28 ok, 0 degraded, 1 empty, 1 intentionally rejected). Optional financial/statistical/catalog expansion remains excluded unless separately requested. No commit or push was performed. See [the compatibility matrix](../../formula-compatibility.md) for finite supported profiles and explicit limits.
