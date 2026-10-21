# Slideshow corpus: provenance

These unmodified PPTX files are untrusted test data for slideshow and slide-appearance work. Nothing in them is executed. Their producer is **metadata-reported** Microsoft PowerPoint; this is screening, not proof of authorship or Office-rendering fidelity. No file was independently re-saved or captured in Office for this task. PowerPoint reference captures remain the separate D4 gate.

## Source and licence

- Source: LibreOffice core, `sd/qa/unit/data/pptx`, at <https://github.com/LibreOffice/core>. Each file below is pinned to a full commit and Git blob SHA, never a moving branch.
- **Licence: MPL-2.0.** LibreOffice publishes its test data under MPL-2.0. These decks are fetched on demand, not redistributed or vendored in this repository. Several docProps creators are third parties, suggesting bug-report attachments; this is why the repository does not vendor the decks. Terms: <https://mozilla.org/MPL/2.0/>.
- Creator metadata is not a copyright clearance or proof of authorship. This task used controller-prepared bytes and pins without network access; origins of individual attachments were not independently investigated.
- Run `bun scripts/fetch-corpus.ts` to fetch `corpus/lo-<file>`. Run `bun scripts/fetch-corpus.ts --check` to verify SHA-256 against `corpus/corpus.lock.json`. The corpus directory is gitignored.

## Screening and reproducibility

1. The controller retrieved the last LibreOffice commit touching each path and its blob at that commit. Those pins were supplied in `S0-2-pins.tsv`.
2. Independently recomputed each local Git blob SHA (`SHA1("blob " + byteLength + NUL + bytes)`) and matched it to the supplied upstream blob. Copies are byte-identical to the prepared source.
3. Read `docProps/app.xml` and required Application to contain Microsoft, matching the supplied producer exactly.
4. Parsed XML by namespace, separately scanning slide, notes, master and other parts. The test has independent per-file assertions for every fetch-entry probe, including attributes and target relationships. It requires the complete 16-file set so an empty fetch list cannot pass.
5. The five appearance candidates were retained. `ooxtheme.pptx` has **ten** non-placeholder master artwork nodes: eight shapes and two connectors. `tdf123684.pptx` has six non-placeholder artwork nodes: two shapes and four pictures. `tdf146223.pptx` contributes a master picture, not a non-placeholder `p:sp`.

Raw URL pattern: `https://raw.githubusercontent.com/LibreOffice/core/<commit>/sd/qa/unit/data/pptx/<file>`.

## Files

| File | Commit (full SHA) | Git blob SHA | Application | Verified pins |
|---|---|---|---|---|
| [tdf104788.pptx](https://raw.githubusercontent.com/LibreOffice/core/319a6fb33c22156e83cebaf3447d7432beca8fb3/sd/qa/unit/data/pptx/tdf104788.pptx) | `319a6fb33c22156e83cebaf3447d7432beca8fb3` | `c46fb3da50a96142d99b125370679482e87758d7` | Microsoft Office PowerPoint | six slides and six notes bodies; showNarration=1; entrance presets 10/34/37/42/47/58; animEffect, animScale and bldP |
| [tdf115394.pptx](https://raw.githubusercontent.com/LibreOffice/core/a8a3e6a7fee5ac23bc2155b9391ead9402211147/sd/qa/unit/data/pptx/tdf115394.pptx) | `a8a3e6a7fee5ac23bc2155b9391ead9402211147` | `d01c90270f345e257c7b4d7bf0d5ad4685ed42cb` | Microsoft Office PowerPoint | fade; speed-only and explicit durations 250/4250 |
| [tdf142915.pptx](https://raw.githubusercontent.com/LibreOffice/core/f8ddaaf0f5e1fb61e0d4404ea28757bc652ae4be/sd/qa/unit/data/pptx/tdf142915.pptx) | `f8ddaaf0f5e1fb61e0d4404ea28757bc652ae4be` | `889a08fd50a1f069410de5edb67b7b769e2ff283` | Microsoft Office PowerPoint | advClick=0 advTm=1000; durations 0/10 |
| [tdf44223.pptx](https://raw.githubusercontent.com/LibreOffice/core/8a1321362a0229a25869e4e3d0422a5a51c5b5be/sd/qa/unit/data/pptx/tdf44223.pptx) | `8a1321362a0229a25869e4e3d0422a5a51c5b5be` | `6f0af688c2062607d261ecfcd84c39b347490b64` | Microsoft Office PowerPoint | split with sound; paragraph background build; entrance preset 2 |
| [bnc591147.pptx](https://raw.githubusercontent.com/LibreOffice/core/a3ef3b658dea9b59ded319f2f092e73e4deaad24/sd/qa/unit/data/pptx/bnc591147.pptx) | `a3ef3b658dea9b59ded319f2f092e73e4deaad24` | `afa382e65770d9ac80e7b9a9c5b14cb727ff4d8e` | Microsoft Office PowerPoint | interactive click sequence and media commands; AVI and MP3 media with picture poster |
| [connector-shape-animations.pptx](https://raw.githubusercontent.com/LibreOffice/core/8359e9974a5a43b094a7eecb73070b26fcb4df5a/sd/qa/unit/data/pptx/connector-shape-animations.pptx) | `8359e9974a5a43b094a7eecb73070b26fcb4df5a` | `2fc2cee6af20ff9cee04e372103e6a2cc0e9ac90` | Microsoft Macintosh PowerPoint | preset 22 targets connectors |
| [tdf104792-smart-art-animation.pptx](https://raw.githubusercontent.com/LibreOffice/core/7c354d1bf1ca29a5ae9ee9d6f18dd68b90530043/sd/qa/unit/data/pptx/tdf104792-smart-art-animation.pptx) | `7c354d1bf1ca29a5ae9ee9d6f18dd68b90530043` | `ac72639efd77ca2b64b9caa91052addb51f02b9e` | Microsoft Office PowerPoint | SmartArt bldDgm |
| [narration.pptx](https://raw.githubusercontent.com/LibreOffice/core/56f26399df894d05f24874a3cdc90694158e03bc/sd/qa/unit/data/pptx/narration.pptx) | `56f26399df894d05f24874a3cdc90694158e03bc` | `5226ce6072158d6e738a05e571d6674f02742f47` | Microsoft Office PowerPoint | audio, p14 media, M4A and transition |
| [airplane-transition.pptx](https://raw.githubusercontent.com/LibreOffice/core/2c6eb9d95930a33318bcec785edf735aeb8adced/sd/qa/unit/data/pptx/airplane-transition.pptx) | `2c6eb9d95930a33318bcec785edf735aeb8adced` | `18ee9aa1962107d954e8c87b51d6344961dd05c6` | Microsoft Office PowerPoint | p15 airplane transition |
| [tdf156808.pptx](https://raw.githubusercontent.com/LibreOffice/core/40641b868b20c73e694770c2c7ea71e232514fa1/sd/qa/unit/data/pptx/tdf156808.pptx) | `40641b868b20c73e694770c2c7ea71e232514fa1` | `34256151b60bec1085e50ef5b5d5af7a823d8068` | Microsoft Office PowerPoint | p14 flythrough |
| [tdf149588_transparentSolidFill.pptx](https://raw.githubusercontent.com/LibreOffice/core/0cc9aaae5dc6655490513e8e4a93967f6248e23c/sd/qa/unit/data/pptx/tdf149588_transparentSolidFill.pptx) | `0cc9aaae5dc6655490513e8e4a93967f6248e23c` | `8967590b92848a8d5f05f5fc599a17469e06d657` | Microsoft Office PowerPoint | p159 morph |
| [tdf105739.pptx](https://raw.githubusercontent.com/LibreOffice/core/319a6fb33c22156e83cebaf3447d7432beca8fb3/sd/qa/unit/data/pptx/tdf105739.pptx) | `319a6fb33c22156e83cebaf3447d7432beca8fb3` | `08e26f55a91dda83422135719107a0a53535e597` | Microsoft Office PowerPoint | slide gradient background |
| [tdf146223.pptx](https://raw.githubusercontent.com/LibreOffice/core/4574a1ea408c2ac30042dca32d02207ec7add4da/sd/qa/unit/data/pptx/tdf146223.pptx) | `4574a1ea408c2ac30042dca32d02207ec7add4da` | `73d5eea144cce4c7e7e1a1a3a7d80d4de28b5a7d` | Microsoft Office PowerPoint | slide and master picture backgrounds; master picture artwork |
| [tdf123684.pptx](https://raw.githubusercontent.com/LibreOffice/core/943a534ac7cb3df513583e226c986dafd8ba246b/sd/qa/unit/data/pptx/tdf123684.pptx) | `943a534ac7cb3df513583e226c986dafd8ba246b` | `87243422753bb3922828cbb382265d304cd34fb9` | Microsoft Office PowerPoint | master bgRef 1003 and six non-placeholder artwork nodes |
| [ooxtheme.pptx](https://raw.githubusercontent.com/LibreOffice/core/cd08835d1bd9274412c7326db32fb3752bc94fb3/sd/qa/unit/data/pptx/ooxtheme.pptx) | `cd08835d1bd9274412c7326db32fb3752bc94fb3` | `f0f3f2575894f715759a18439588cb149996140b` | Microsoft Office PowerPoint | ten non-placeholder master artwork nodes |
| [tdf93868.pptx](https://raw.githubusercontent.com/LibreOffice/core/f3d1ac75c4b7fa63022e54a9cbff46ba99535076/sd/qa/unit/data/pptx/tdf93868.pptx) | `f3d1ac75c4b7fa63022e54a9cbff46ba99535076` | `f54fb0496e117eb55e0c16a023c045f80ba2b6dd` | Microsoft Office PowerPoint | master gradient background and two non-placeholder shapes |

## Rejected candidates

The following exclusions are carried from the controller's initial scan in the task brief. Their bytes were not supplied among the 16 candidates; the reasons below are **not independently reverified** in this task, and none is imported.

- `effectOrder.pptx`: reported to contain no `p:timing` or `p:transition`; it tests effect-list order instead.
- `animation-sound-embedded.pptx` and `animation-sound-external.pptx`: reported to lack app.xml producer metadata and carry a 52-byte WAV.
- `tdf148478_transitionMusic.pptx`: reported LibreOffice-authored, so outside the metadata-reported Microsoft corpus.

## Evidence limits

The blob pins establish identity with the controller's upstream retrieval, not independent verification of historical authoring. Application metadata can survive later ZIP/XML edits. No Microsoft conformance claim, pixel baseline, precise Office build, or real-browser result follows from this import. `tdf104788.pptx` is the designated benchmark deck, but no frame benchmark was run here. Its six entrance presets are all on slide 6, not one per slide.

XML precision: the SmartArt fixture uses DrawingML `a:bldDgm` inside the PresentationML build path, not a `p:bldDgm` element. In the connector fixture, preset 22 targets connector 44 and also ordinary shape 52; it is not exclusively a connector preset.

## Per-file creator metadata

| File | docProps creator |
|---|---|
| tdf104788.pptx | blank |
| tdf115394.pptx | Szymon Kłos |
| tdf142915.pptx | fejleszto |
| tdf44223.pptx | luke |
| bnc591147.pptx | m |
| connector-shape-animations.pptx | Helmut Steuler/karthikgodha |
| tdf104792-smart-art-animation.pptx | Szymon Kłos |
| narration.pptx | Microsoft account |
| airplane-transition.pptx | blank |
| tdf156808.pptx | India Armstrong |
| tdf149588_transparentSolidFill.pptx | Keith Miracle |
| tdf105739.pptx | Windows User |
| tdf146223.pptx | blank |
| tdf123684.pptx | me |
| ooxtheme.pptx | Grzegorz Araminowicz |
| tdf93868.pptx | Your User Name |
