# Office text directions and inline icons

This is stage 3 of Julian's approved Office drawings program. Implementation starts after stage 2 passes review and verification. Native fixtures are retained in `/tmp/officeview-text-directions-20261003`; source packages, generators, XML and hashes accompany the local PDF exports.

## Shared text behavior

Preserve the source direction token, paragraph order, explicit breaks, run styles and script font assignments. A shared layout result contains placed text runs with local transforms; adapters supply available bounds and paragraph/run data. Do not flatten rotated text into an unsearchable bitmap. Existing affine text recording supplies search highlights and selection geometry.

Affine placement alone does not establish reading order. Retain logical line/run identity, source offsets and grapheme boundaries alongside placements. Search must join mixed upright/rotated runs into their logical line and find a word painted as stacked glyphs; selection/copy follows the declared column progression and source text. Carets do not split combining or emoji clusters. Existing handwritten indexes and ordinary horizontal reading order remain compatible.

DrawingML requires all seven tokens: `horz`, `vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert`, `wordArtVertRtl`. Distinguish whole-line rotation, mixed vertical orientation, individually stacked WordArt characters, and column progression. Shape rotation/flips compose outside the text layout. Retain wrapping, explicit line breaks, insets, horizontal alignment and block anchoring. Latin runs remain shaped together when the direction calls for run rotation; combining sequences and emoji must not be split into individual code units. Mongolian runs remain shaped together and use the source script font. Vertical punctuation orientation follows the controlled native export and Unicode vertical-orientation data.

Resolve text styles by field from master, layout, placeholder, list style, paragraph default, run properties and end-paragraph properties. Explicit false/zero and direct color/font choices override inheritance. Theme major/minor Latin, East Asian and complex-script font assignments remain distinct. The retained open a11 title must inherit its 44-point master size. WordArt text fills, outlines and shadows are text appearance properties: support the retained a11 patterned fill and outline/shadow without changing ordinary shape fills.

Cached diagrams and other DrawingML text payloads also retain body/list defaults, paragraph defRPr and fontRef before direct run overrides. This applies to all three adapters, including complex.docx's cached diagram labels; it is separate from WordprocessingML paragraph/table style resolution.

Multiple appearance passes produce one logical text record, including outline-only text. Search does not duplicate text because a shadow, fill and outline all paint the same glyphs.

## Format adapters

Word table cells preserve `lrTb`, `tbRl`, `btLr`, `lrTbV`, `tbRlV`, `tbLrV`. Layout measures in the appropriate local coordinate system, then projects into the cell; wrapping, merged-cell extent, margins, vertical alignment and following table/document flow retain their meaning. The six-column native Word reference pins the context-specific behavior. Word drawing text boxes use the shared direction engine while retaining Word paragraphs and placement.

Excel cell `textRotation` has values 0–90 for counterclockwise angles, 91–180 for clockwise angles 1–90, and 255 for stacked characters. Preserve explicit wrap, horizontal/vertical alignment, merged bounds and cell clipping. Retain ordinary drawing-free sheet dimensions and border behavior. Spreadsheet drawing text uses the same DrawingML directions as the other adapters.

## Icons and image representations

Choose one image representation. For `a:blip` with the standard SVG extension, prefer the referenced SVG when the browser can decode it and retain the PNG fallback when SVG is missing, invalid or unsupported. Relationship targets remain relative to their owner part and external relations are skipped. Embedded SVG is self-contained; do not fetch its external assets. Existing raster-only inputs retain their behavior.

Enforce that boundary before either browser or Node decoding. Reject external href/xlink:href, CSS url/import references, external entities, scripts/event handlers and foreignObject; retain internal fragment references needed by valid icons. Negative decode tests verify that forbidden content cannot trigger network or local-file access. A real-browser decode/paint case is required in addition to Node tests. Rejected or failed SVG selects the raster fallback once and records the reason.

Preserve mixed text/image source order and inline extents so four horizontal icons remain on a line when the available width permits and wrap predictably when it does not. The controlled icon fixture contains four SVG/PNG pairs between `Before` and `After`. Record selected SVG or raster representation and fallback reason in drawing coverage. Do not count both representations as painted objects.

## Acceptance

- Independent fixed-measure assertions cover each direction's glyph/run orientation, column progression, wrapping, alignment, mixed scripts, combining clusters and transform composition. Actual-paint search/copy tests cover logical lines across mixed orientation, stacked words and vertical columns; appearance passes produce one logical record. Painted pixel assertions verify representative directions and styled WordArt.
- Adapter tests cover all six Word directions and all Excel rotation modes, including merged cells and mixed ordinary/rotated content. Search and strict/non-strict selection tests cover actual transformed text without changing ordinary line selection.
- Both retained a11 versions and complex.docx preserve unit counts and source object counts. The open title and WordArt appearance improve against its retained native export; retained fallback content remains explicitly identified.
- Render each controlled native fixture at matching content coordinates, retain comparisons and quantify residual differences. Complete the pending native SVG-icon export when computer-use is available; do not claim native validation from a generated raster fallback alone.
- TypeScript, strict tests, build and corpus checks pass. Individually inspect intentional golden differences. No commits, original files unchanged, and concurrent workspace edits preserved.

Exact installed-font rasterization is not promised. Any missing direction, image representation or text appearance variant remains an explicit coverage entry and prevents a full-program completion claim.
