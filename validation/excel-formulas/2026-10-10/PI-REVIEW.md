# Pi formula-only approval

Decision: **APPROVE for the formula-only commit/push**, excluding the unrelated
`testTimeout: 20000` hunk in `vitest.config.ts`. Keep the formula worker-cap hunk;
leave all unrelated pending Word/drawing changes out of the commit.

Pi reviewed the frozen 63-path candidate against base commit
`aca4d465d772ed32211dd1cb5d9c603d4e792091`. It found no unresolved critical,
high or medium defect. This approval covers the finite documented formula scope,
not full Excel function-catalog compatibility.

Review packet: `/tmp/officeview-pi-formula-approval-csue9_zc/`; exact paths/hashes
are in `MANIFEST.json`. Fresh review session:
`01a1237e-857a-71ad-9a43-4cf144775f62`, Herdr agent `pi-formula-validation`.
Longcat encountered upstream HTTP 500 errors during review; the review completed
after restoring Pi's previously selected `opencode-go/step-5-preview-free` for
this session. No account, credentials, provider or model defaults were changed.

Pi independently ran TypeScript (exit 0), the 44 native regression tests, and
1,023 formula-family tests. It reviewed source and scratch counterexamples for
Unicode LEFT/RIGHT, XLDAPR metadata, saved-follower ownership and resizing,
General/time display, and single-pass escaped strings. It checked cache/error/
lazy/suspension/iteration boundaries and all three changed golden images.

Root separately ran the full frozen suite with the original five-second timeout:
2,899 passed, two existing skips, 104 files. The native baseline replay includes
valid cached unsupported results; matching those rows is **not** claimed as
successful recalculation of unsupported functions. The distinct ordinary,
Unicode, changed-input array, cache-decoding and display controls are documented
in the main report and native replay artifact.

Root rechecked the source fingerprints against the live formula files after
approval. No formula commit or push was performed by this review task.
