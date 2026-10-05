# DrawingML preset data attribution

`presets.json` is a modified, normalized copy of the 187 preset shape definitions
from Apache POI, tag `REL_5_4_1`:

[poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml](https://github.com/apache/poi/blob/REL_5_4_1/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml)

Apache POI  
Copyright 2003-2025 The Apache Software Foundation

This product includes software developed at The Apache Software Foundation
(https://www.apache.org/).

The data is licensed under the Apache License, Version 2.0. The complete license
is reproduced in [LICENSE](./LICENSE). The original project attribution is in
[Apache POI's NOTICE](https://github.com/apache/poi/blob/REL_5_4_1/legal/NOTICE).

The pinned input SHA-256 is:

```
a7dad593d27bd70536b41da9b761fa16409536cc0c25ef2b6c7a61c5d9b3e738
```

OfficeView modifications: the source XML is converted into JSON containing
adjustment defaults, ordered guides, ordered path commands, path coordinate
spaces, fill/stroke flags, and text rectangles. Adjustment handles, connector
attachment metadata, namespace declarations, and extrusion flags are omitted.
Eight source formulas contain a redundant fourth zero operand after the
three-operand `+-` operator. Generation removes only that redundant operand in
`xB` and `yB` of `circularArrow`, `leftCircularArrow`, and
`leftRightCircularArrow`, and `xJ` and `yJ` of `leftRightCircularArrow`.
All other formulas retain their source strings and order.

Reproduce the JSON from a byte-identical local input:

```sh
bun scripts/generate-drawing-presets.ts /local/presetShapeDefinitions.xml
```

The generator rejects a different input hash. Generation and runtime resolution
make no network requests.
