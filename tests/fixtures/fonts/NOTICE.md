# Test font provenance

`Liter-Regular.ttf` and `OFL.txt` were independently downloaded from Google
Fonts. `provenance.json` records exact source URLs, byte counts and SHA-256s.
The exact SIL Open Font License is retained. This font is only a test fixture.
It was not extracted from an Office document.

`authored-*.mtx` are authored synthetic CTF/MTX data: a minimal one-empty-glyph
font, plus hostile table, glyph-point, RLE and aggregate CVT expansion cases.
The CVT fixtures include repeated underdeclared aliases and one unique expanding
CVT beside sibling tables. They contain no user font
bytes or third-party outlines. These synthetic font bytes are dedicated to the
public domain under CC0 1.0 (https://creativecommons.org/publicdomain/zero/1.0/).
`generate-authored.mjs` documents reproducible construction and uses the pinned
upstream adaptive Huffman model solely to encode literal test streams.
These MTX fixtures test actual decoding; the licensed Liter font is the valid
browser registration fixture. No test fixture is shipped in the library build.
