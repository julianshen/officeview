# Slideshow 3-D transition algorithms (S3-1)

Status: **implemented** (written as the S3-1 proposal before S3-2 to S3-4 were dispatched; reviewed by agy, and kept in step with the code by S3-2 to S3-4 and S3-7). Each of the 26 catalogued 3-D rows has one entry below. An entry gives the row's mesh, motion, camera and timing curve, its seeded parameters, and a numeric oracle at 50 %. Every row is a **closest look-alike** (spec §1, §6.3). PowerPoint does not publish its 3-D rendering, so no row here is claimed to match Office. Each row stays `approximate` in the catalogue. Rows marked "D4 unverified" also have an unverified classification in the catalogue.

The executable reference is `scripts/slideshow-3d-oracles.ts`, which uses the framework's own maths (`src/slideshow/render/three/mesh.ts`). Every formula below is a function in that script, under the name given. `bun scripts/slideshow-3d-oracles.ts` regenerates the oracle table at the end of this document. If this document and the script disagree, the script wins, and the disagreement is a review finding. The script models only what the oracles measure. Terms that no oracle depends on are specified here only: drape's ripple and sway, curtains' folds, honeycomb's hexagon mask, glitter's sparkle mask, peelOff's colours and shadow, the fold back colours, shred `in`, warp `out` and the `invX`/`invY` mirrors.

**Stage 3 close (S3-7).** Every row below is implemented and wired: `src/slideshow/render/three/index.ts` (`registerTransitions3D`) registers the three groups, and the show's backend manager calls it before it creates a WebGL2 compositor. Each row is `approximate` in the catalogue at the default approximate tolerance (mean abs 0.08, row max 0.30) and is verified only from compare records; none exists yet, because the owner's PowerPoint captures (D4) have not arrived. drape, wind, fallOver, warp and airplane are look-alikes awaiting owner PowerPoint capture (D4) for their motion as well as their classification. The shader path emits no diagnostic of its own when it draws (by design, S3-1 M1): `transition-approximate` is said only when the look-alike draws, and `transition-shader-fallback` only when the shader could not draw a frame.

`flash`, `reveal` and `pan` are not on this sheet: S1-15 draws them as 2-D transitions (`exact` rows). The four Morph rows are S3-5/S3-6.

## Framework conventions (S3-1)

These are the framework's contracts. Every group task builds on them and must not change them.

### Dispatch and fallbacks

- **Dispatch.** `dispatchTransition(spec, compositor, sink?, slideIndex?)` (`three/framework.ts`) returns `'shader'` only when all of these hold:
  - the row is 3-D;
  - a `Transition3D` is registered for the row;
  - the compositor is WebGL2 and its context is not lost;
  - `prepare` succeeded.

  Otherwise it returns `'lookalike'`, and the stage draws `transitionDraw`, which draws the row's catalogued 2-D look-alike. A look-alike is diagnosed once per slide as `transition-approximate`. Reduced motion is applied first: `reducedTransition` turns every real transition into a fade of at most 200 ms, so reduced motion never reaches a shader and never adds motion.
- **Params.** A dispatched shader's frame params are `transition3DDraw(spec, progress, rng, aspect)`: the look-alike's own params, computed with the same rng, plus `three: { ns, element, preset?, attrs }`.
  - Canvas 2D ignores `three` and draws the look-alike.
  - WebGL2 draws the registered shader.
  - If a shader throws during a frame (for example a refused mesh lease or a failed compile), WebGL2 draws the look-alike for that frame instead, diagnosed as `transition-shader-fallback`.
- **Ends.** At progress exactly 0 and exactly 1, the compositor draws the flat outgoing or incoming frame and never calls `draw`. Every model below still has to be continuous into its ends: as p → 0⁺ it must tend to the outgoing frame at rest, and as p → 1⁻ to the incoming frame at rest (or to nothing covering it). Each group spec checks this at p = 0.001 and p = 0.999 against the flat frames, within the row's tolerance, as the S3-1 browser spec does for the sample shader.

### Drawing environment

- **Viewport.** `draw` runs with:
  - the GL viewport set to the slide's device rect;
  - scissoring on that rect;
  - the rect cleared to opaque black, which is also the backdrop every row shows where nothing covers it.

  The compositor resets every GL state afterwards, so a row may enable `CULL_FACE`, change blending, or use its own programs.
- **World space and camera.**
  - The slide is the rectangle x ∈ [−a/2, a/2], y ∈ [−½, ½] on the plane z = 0, where a = width / height (`mesh.aspect`). y points up, and the viewer is on +z.
  - Slide uv (u, v) has its origin at the top-left (v points down). `slidePoint(u, v) = ((u − ½)a, ½ − v, 0)`, and GLSL `slidePos(uv, aspect)` is the same map.
  - The camera is `C = slideCamera(45°, a) = perspective(45°, a) · translation(0, 0, −d)`, with d = `cameraDistance(45°)` = 0.5 / tan 22.5° ≈ 1.2071. Under this camera, an identity model fills the viewport exactly.
  - Every row uses `DEFAULT_FOV_DEG` (45°) unless its entry says otherwise. Near is 0.01 and far is 100. A surface that sweeps past the camera is clipped by the near plane, which is correct perspective.
- **Matrices.** Matrices are column-major. `mul4(A, B)` applies B first. T is `translation`, S is `scaling`, and Rx, Ry, Rz are right-handed rotations about +x, +y and +z. Ry(θ) turns z towards +x, so a positive θ moves the slide's front edge to the right.
- **No depth buffer.** The drawing buffer has none, and the 3-D path allocates no surface, so transitions still play at 4K inside the slide's own transition allowance (S1-22). Rows order their draws back to front (painter's order): flat `to` first when it is a backdrop, then pieces in order of increasing view-space depth of their centres. Rows may cull back faces. Front faces are counter-clockwise, and a grid at rest fronts the viewer. No row's pieces interpenetrate; the entries keep moving pieces on one side of any flat backdrop.

### Meshes and programs

- **`MeshPool`**:
  - `grid(cols, rows)` returns cols × rows unshared quads, cached. Each grid is leased (`webgl2:mesh`) before it is created. cols and rows run from 1 to `GRID_MAX` (128). Attributes: `layout(location = 0) in vec2 a_uv` (the vertex's slide uv) and `layout(location = 1) in vec2 a_cell` (its cell's centre uv). A vertex shader can therefore deform smoothly by `a_uv`, or move tiles as units by `a_cell`.
  - `program(key, vs, fs)` returns a cached program.
  - `draw(mesh, texture, mvp, { opacity, shade, uvRect })` draws with the standard program. `uvRect` maps the mesh's [0, 1]² to a slide-uv sub-rect, for both position and texture.
  - `drawMesh(mesh)` binds the mesh and issues its indexed draw.

  Handles are valid only for the current frame. Ask the pool again on every frame (a cache hit); context loss drops everything. Lease and compile in `prepare(mesh)`, so a failure there sends the whole effect to its look-alike before it starts.
- **Draw calls.** A row draws `to` flat with `mesh.draw(mesh.grid(1, 1), to, C)`. A tile row draws all of its tiles in **one** indexed draw with a custom vertex shader keyed on `a_cell`. The p95 ≤ 12 ms budget assumes O(1) draw calls per frame, never one call per tile.

### Timing, randomness and attributes

- **Timing.** e(t) = smoothstep(t) = t²(3 − 2t) on [0, 1] (`ease`, GLSL `ease`). Values: e(0) = 0, e(½) = ½, e(1) = 1. s(p) = sin(πp) is the "bulge" curve, 0 at both ends and 1 at the middle.
- **Seeds.** The shader's rng is re-seeded from the effect's seed on every frame. A row derives its effect seed from that rng's first value, `effectSeed(rng)` = ⌊rng() · 2³²⌋, and then draws every per-cell value from `cellRandom(seed, i, j, k)` (CPU) or GLSL `cellRandom` (`HASH_GLSL`). Both are the lowbias32 integer hash of `seed ^ hash(i + 1024 (j + 1024 k))` / 2³². The hash is exact on every GPU; the float division agrees to float precision. A row never calls `Math.random`, and never draws a per-frame random from the rng. The oracles use effect seed 2³¹, which is a shader rng whose first value is 0.5.
- **Attributes.** Boolean attributes are true for `"1"` or `"true"`. An omitted or unknown value takes the default listed in the entry.

### Tile template

The template is `tileTime` in the script. It is used by vortex, shred, fracture, crush, glitter, honeycomb and wind.

- **Cells.** A grid of cols × rows cells. Cell (i, j) has its centre at uv ((i + ½)/cols, (j + ½)/rows), and its world centre is c = `slidePoint` of that uv.
- **Delay.** Each cell has an order g ∈ [0, 1] and a jitter r₀ = `cellRandom(seed, i, j, 0)`. Its delay is δ = W((1 − J) g + J r₀) ∈ [0, W], and its local time is t = clamp((p − δ) / (1 − W), 0, 1). So every cell has t = 0 at p = 0 and t = 1 at p = 1.
- **Order g.** g is one of:
  - *leading(dir)*: u_c for `l`, 1 − u_c for `r`, v_c for `u`, 1 − v_c for `d`, so the edge the motion heads for goes first;
  - *fromCentre*: |(u_c − ½)a, v_c − ½| / |(a/2, ½)|, so the centre goes first;
  - 0: jitter only.
- **Fade.** Unless an entry says otherwise, a tile's opacity is fade(t) = 1 − e((t − 0.7)/0.3), so a tile is gone by t = 1.
- **Drawing.** `to` is drawn flat first. All tiles then draw in one call, in cell order.
- **Oracles.** "rest" is the fraction of cells with t = 0 at p = ½, and "done" is the fraction with t = 1.

### Fold template

The template is `fold` in the script. It is used by pageCurlSingle, pageCurlDouble and peelOff.

- **Fragment shader.** The fold is evaluated per fragment on a 1 × 1 grid, using a custom fragment shader with both frames bound. Pixel q = (x, y) is in slide units: x ∈ [0, a] to the right, y ∈ [0, 1] downwards. A page has an x range, a start corner P₀ and a target corner P₁.
- **Fold line.** The unit direction is n = (P₁ − P₀)/D, where D = |P₁ − P₀|. The fold line sits at L = D · e(p) from P₀. Each pixel has s = (q − P₀)·n.
- **Per-pixel rule:**
  - if s < L, the paper is gone: show `to`;
  - otherwise reflect q across the fold line, m = q − 2(s − L) n. If m is inside the page, show the **back** of the page at m (see each row's colour). Otherwise show `from` at q.
- **Crease shading.** Shade by 1 − 0.3 exp(−|s − L| / 0.05). This multiplies the colour and does not change the classification.
- **Ends.** At p = 0, L = 0, so the whole frame is `from`. At p = 1, L = D, so the whole page is `to`.

## S3-2: geometric rows (`three/geometric.ts`, `registerGeometric3D()`)

### Face shading and the doors back-face pass (S3-2 M1)

These apply to every geometric row; the per-row entries below give only the motion.

- **Shading.** Each piece is drawn with the standard program's `shade`, which multiplies rgb only (premultiplied, so alpha and coverage are unchanged). A piece whose model turns its normal by φ from the viewer axis is shaded by `0.7 + 0.3 |cos φ|` (`shadeOf`: `0.7 + 0.3 min(1, |m[10]|)`, `m[10]` being the z component of the turned normal). It is 1 at rest, so progress 0 and 1 stay exact, and it falls to 0.7 when the piece is edge-on. No light position, specular term or ambient colour is modelled; this is a look-alike of a face turning away from a light behind the viewer.
- **Culling.** Front faces are counter-clockwise, as the grid is built. Every piece is drawn with `CULL_FACE` on and `cullFace(BACK)`, so prism's hidden cube side and flip's other frame are never drawn from behind.
- **Two-sided pieces (doors and window).** A half has `sides: 'both'`. After its front pass the same quad is drawn again with `cullFace(FRONT)`, so only its back (the half seen past edge-on) is drawn, with the same texture (mirrored, as paper seen from behind) and `shade · BACK_SHADE` (`BACK_SHADE` = 0.55). Exactly one of the two passes covers a pixel, so a half never double-blends. gallery, switch, conveyor and ferris never turn far enough to face away and draw one pass.
- **Order.** Pieces are drawn back to front by the world z of each piece's centre (`byDepth`; the drawing buffer has no depth), stable on ties, with premultiplied `ONE, ONE_MINUS_SRC_ALPHA` blending.

### `transition:p14:prism`: Cube / Box / Rotate / Orbit (D4 unverified variants)

- **Attributes:**
  - `dir` is `l` (default), `r`, `u` or `d`, the direction the content moves in;
  - `isInverted` (default false);
  - `isContent` (default false).
- **Look-alike:** push.
- **Mesh:** two 1 × 1 grids (the from face and the to face), each slide-sized.
- **Motion** (`prism`). θ = (π/2) e(p). The cube depth is D = a for `l`/`r` (rotation about y) and D = 1 for `u`/`d` (rotation about x). σ is −1 for `l`/`u` and +1 for `r`/`d`. R(φ) is Ry(φ) for `l`/`r` and Rx(φ) for `u`/`d`.
  - **Outside (default).** The pivot is K = −D/2 and the turn is s = σ:
    - from = T(0,0,K) R(sθ) T(0,0,−K);
    - to = T(0,0,K) R(sθ − sπ/2) T(0,0,−K).
  - **`isInverted`: the inside of the box.** K = +D/2 and s = −σ. The faces sweep past the camera and are clipped by the near plane.
  - **`isContent`:** the whole cube is additionally pushed back by T(0, 0, −0.3 s(p)). This is an approximation: content and background are one frame here.
- **Camera:** C.
- **Timing:** e(p).
- **Draw:** cull back faces, and draw the farther face (by view-space z of its centre) first.
- **Seeded parameters:** none.
- **Oracle (dir `l`):** from uv(1, 0), the shared top edge, is at NDC (0, 1.4389).

### `transition:p14:flip`

- **Attributes:** `dir` is `l` (default) or `r`.
- **Look-alike:** push.
- **Mesh:** two 1 × 1 grids.
- **Motion** (`flip`). θ = π e(p). σ is −1 for `l` and +1 for `r`. z_r = −0.5 s(p).
  - from = T(0,0,z_r) Ry(σθ);
  - to = T(0,0,z_r) Ry(σθ − σπ), which is the back of the card and right way round at the end.
- **Draw:** cull back faces. Only one face is visible at a time; the card is edge-on at ½.
- **Seeded parameters:** none.
- **Oracle:** from uv(0, 0) is at world (0, 0.5, −1.3889) and NDC (0, 0.465).

### `transition:p14:gallery`

- **Attributes:** `dir` is `l` (default) or `r`.
- **Look-alike:** push.
- **Motion** (`gallery`). σ is −1 for `l` and +1 for `r`.
  - from = T(1.1aσ e, 0, −0.6 s) Ry(0.35σ s);
  - to = T(1.1aσ (e − 1), 0, −0.6 s) Ry(0.35σ s).

  This reads as two pictures on one wall, 0.1a apart, panning past.
- **Draw:** either order; the frames never overlap.
- **Seeded parameters:** none.
- **Oracle:** from uv(1, ½) is at NDC (−0.1291, 0), and to uv(0, ½) is at (0.0918, 0).

### `transition:p14:doors`

- **Attributes:** `dir` is `vert` (default; a vertical split, so the doors are left and right) or `horz`.
- **Look-alike:** split.
- **Mesh:** a 1 × 1 grid drawn three times. The halves use `uvRect` [0,0,½,1] and [½,0,½,1] (for `horz`, [0,0,1,½] and [0,½,1,½]).
- **Motion** (`doors`). θ = (π/2) e(p). The halves swing **towards the viewer** about their outer edges:
  - **`vert`:**
    - left = T(−a/2,0,0) Ry(−θ) T(a/2,0,0);
    - right = T(a/2,0,0) Ry(θ) T(−a/2,0,0).
  - **`horz`:**
    - top = T(0,½,0) Rx(−θ) T(0,−½,0);
    - bottom = T(0,−½,0) Rx(θ) T(0,½,0).
  - to = T(0, 0, −0.6(1 − e)). It approaches from behind and never crosses z = 0, while the doors stay at z ≥ 0.
- **Draw:** to, then the two halves.
- **Seeded parameters:** none.
- **Oracle:** the left half's uv(½, 0), its free top corner, is at NDC (−0.6111, 2.0864) (off-screen above, near the camera). to uv(0, 0) is at (−0.8009, 0.8009).

### `transition:p14:window`

- **Attributes:** `dir` is `vert` (default) or `horz`.
- **Look-alike:** split.
- **Motion** (`doors(…, windowed)`). The halves move as in doors, and also part outwards by T(∓0.25a e, 0, 0) (for `horz`, T(0, ±0.25 e, 0)). to = T(0, 0, −1.2(1 − e)).
- **Draw:** to, then the two halves.
- **Seeded parameters:** none.
- **Oracle:** the left half's uv(½, 0) is at NDC (−1.1327, 2.0864), and to uv(0, 0) is at (−0.668, 0.668).

### `transition:p14:switch`

- **Attributes:** `dir` is `l` (default) or `r`.
- **Look-alike:** push.
- **Motion** (`switchRow`). σ is −1 for `l` and +1 for `r`.
  - from = T(0.55aσ s, 0, −1.2 e) Ry(0.6σ s);
  - to = T(−0.55aσ s, 0, −1.2(1 − e)) Ry(−0.6σ s).

  The two frames swap depth, passing side by side.
- **Draw:** the frame with the smaller z translation first. On a tie (p = ½), draw from first.
- **Seeded parameters:** none.
- **Oracle:** from uv(½, ½) is at NDC (−0.7348, 0), to uv(½, ½) is at (0.7348, 0), and from uv(1, 0) is at (−0.254, 0.9248).

### `transition:p14:conveyor`

- **Attributes:** `dir` is `l` (default) or `r`.
- **Look-alike:** push.
- **Motion** (`conveyor`). σ is −1 for `l` and +1 for `r`, z = −0.8 s, and x_f = 1.1aσ e.
  - from = T(x_f, 0, z) Ry(0.25σ s);
  - to = T(x_f − 1.1aσ, 0, z) Ry(0.25σ s).
- **Draw:** either order; the frames never overlap.
- **Seeded parameters:** none.
- **Oracle:** from uv(1, ½) is at NDC (−0.0885, 0), and to uv(0, ½) is at (0.0711, 0).

### `transition:p14:ferris`

- **Attributes:** `dir` is `l` (default) or `r`.
- **Look-alike:** push.
- **Motion** (`ferris`). Both frames ride the top of a wheel and stay upright:
  - the wheel's radius is R = 0.9a and its centre is (0, −R);
  - φ = (π/3) e(p), and z = −0.3 s;
  - the position at angle ψ is (R sin ψ, −R + R cos ψ, z);
  - from is at ψ = σφ, and to is at ψ = σ(φ − π/3).
- **Draw:** either order.
- **Seeded parameters:** none.
- **Oracle:** from uv(½, ½) is at NDC (−0.7208, −0.3434), and to uv(½, ½) is at (0.7208, −0.3434).

## S3-3: tile and particle rows (`three/particles.ts`, `registerParticle3D()`)

### `transition:p14:vortex`

- **Attributes:** `dir` is `l` (default), `u`, `r` or `d`.
- **Look-alike:** dissolve.
- **Mesh:** a 32 × 18 grid.
- **Motion** (`vortexTile`). Tile template with W = 0.6, J = 0.35 and g = leading(dir). Let m be the direction's unit vector (for example `l` = (−1, 0)), with r₁, r₂, r₃ = `cellRandom(seed, i, j, 1..3)`.
  - swirl = 0.4 sin(πt)(2r₂ − 1);
  - the tile moves by 1.4a t² m + swirl (−m_y, m_x) in xy, and by 0.5t in z;
  - tile = T(c + offset) Rz(2π(0.75 + 0.5r₁)t) Ry(π t(r₃ − ½)) S(1 − 0.5t) T(−c);
  - opacity is fade(t).
- **Draw:** to flat, then all tiles in one draw.
- **Seeded parameters:** r₀ to r₃ per cell.
- **Oracle:**
  - rest = 0.059 and done = 0.059;
  - cell (0, 0) has t = 0.8158, and its centre is at NDC (−4.2778, 0.8061).

### `transition:p14:shred`

- **Attributes:**
  - `pattern` is `strip` (default) or `rectangle`;
  - `dir` is `out` (default) or `in`.
- **Look-alike:** dissolve.
- **Motion** (`shredTile`):
  - **strip:** a 24 × 1 grid, W = 0.5, J = 0.2, g = 0. Strip i moves by T(0, ±1.3t², 0): up for even i, down for odd i.
  - **rectangle:** a 24 × 14 grid, W = 0.5, J = 0.6, g = 0. With heading 2π r₁:
    - tile = T(c + 1.5t(cos, sin), 0.8t) Rz(2π r₂ t) S(1 − 0.6t) T(−c).
  - opacity is fade(t) for both.
  - **`dir in`:** the roles swap and time runs backwards. `from` is drawn flat, and the `to` tiles are drawn with each cell's local time evaluated at 1 − p.
- **Seeded parameters:** r₀ to r₂ per cell.
- **Oracle (strip, out):**
  - rest = 0 and done = 0 (rectangle: rest = 0, done = 0);
  - strip 0 has t = 0.8404 and dy = 0.9181.

### `transition:p15:prstTrans:fracture`

- **Attributes:** none.
- **Look-alike:** fade.
- **Motion** (`fractureTile`). A 12 × 7 grid, W = 0.6, J = 0.4, g = fromCentre.
  - tile = T(c + (0.4 ĉ t, −1.6t²), 0.3 t r₃) Rx(π r₁ t) Ry(π r₂ t) T(−c), where ĉ is c normalised;
  - opacity is fade(t).
- **Seeded parameters:** r₀ to r₃ per cell.
- **Oracle:** rest = 0.0714 and done = 0.0238. Cell (0, 0)'s centre is at NDC (−0.9167, 0.8571).

### `transition:p15:prstTrans:crush`

- **Attributes:** none.
- **Look-alike:** fade.
- **Motion** (`crushTile`). A 16 × 9 grid, W = 0.3, J = 1, g = 0. With k = 0.9 e(t):
  - tile = B · T(c(1 − k)) Rz(2π(r₁ − ½)t) Rx(π r₂ t) S(1 − 0.8 e(t)) T(−c);
  - the ball offset is B = T(0, 0, −2 e((p − 0.6)/0.4));
  - opacity is 1 − e((p − 0.75)/0.25), the same for every tile.
- **Seeded parameters:** r₀ to r₂ per cell.
- **Oracle:** rest = 0 and done = 0. Cell (0, 0)'s centre is at NDC (−0.6738, 0.6388), with opacity 1.

### `transition:p14:glitter`

- **Attributes:**
  - `pattern` is `diamond` (default) or `hexagon`;
  - `dir` is `l` (default), `u`, `r` or `d`.
- **Look-alike:** dissolve.
- **Motion** (`glitterTile`):
  - **grid:** 32 × 18 for diamond and 24 × 14 for hexagon, with W = 0.65, J = 0.4 and g = leading(dir);
  - **turn:** each cell turns about its centre by π e(t), using Ry(∓·) for `l`/`r` and Rx(∓·) for `u`/`d`;
  - **faces:** the front face shows `from` and the back face shows `to`, with culling;
  - **sparkle:** shade = 1 + 0.6 sin(πt) on both faces;
  - **pattern:** the pattern selects the grid and a sparkle mask (a diamond |x| + |y| ≤ ½, or a hexagon, in cell space), applied to the shade only. Cells at rest stay whole squares, so the ends are exact.
- **Draw:** two draws, front faces (from) and back faces (to), since no flat `to` backdrop is needed.
- **Seeded parameters:** r₀ per cell.
- **Oracle:**
  - rest = 0.1372 and done = 0.1319;
  - 0.4896 of the cells show `to` (e(t) > ½);
  - cell (0, 0) has t = 0.8183.

### `transition:p14:honeycomb`

- **Attributes:** none.
- **Look-alike:** checker.
- **Motion** (`honeycombTile`). A 20 × 11 grid, W = 0.6, J = 0.3, g = fromCentre.
  - tile = T(c, 0.6t) Rz((π/3)t) S(1 − e(t)) T(−c);
  - opacity is 1 − e(t);
  - while a tile is in flight (0 < t < 1), a hexagon inscribed in the cell masks it.
- **Seeded parameters:** r₀ per cell.
- **Oracle:** rest = 0.0455 and done = 0.0318. Cell (10, 5) has t = 1, and cell (0, 0) has t = 0.

### `transition:p14:ripple`

- **Attributes:** `dir` is `center` (default), `lu`, `ru`, `ld` or `rd`, the origin of the ripples.
- **Look-alike:** dissolve.
- **Mesh:** a 64 × 36 grid, deformed in the vertex shader by `a_uv`.
- **Motion** (`rippleZ`). Let O be the origin in uv (centre, or a corner) and ρ = |((u − O_u)a, v − O_v)|.
  - z = 0.06 · 4p(1 − p) · sin(2π(4ρ − 3p)) · exp(−1.5ρ);
  - from is drawn displaced, with opacity 1 − e(p), over to drawn flat.
- **Seeded parameters:** none.
- **Oracle:**
  - z at uv(0.75, ½) is 0.0303, and at the centre it is 0;
  - from's opacity is 0.5;
  - with `lu`, z at uv(0.25, 0.25) is −0.0069.

## S3-4: deformation rows (`three/deform.ts`, `registerDeform3D()`)

### `transition:p15:prstTrans:pageCurlSingle`

- **Attributes:** `invX` and `invY`, which mirror the start corner.
- **Look-alike:** fade.
- **Algorithm:** the fold template with one page over x ∈ [0, a]. P₀ is the bottom-right corner (a, 1): `invX` moves it to x = 0 and `invY` to y = 0. P₁ is the opposite corner. The back of the page is mix(from(m), white, ½).
- **Seeded parameters:** none.
- **Oracle:** 0.5 of the area shows `to` (exact, by symmetry), and 0.3291 shows the back.

### `transition:p15:prstTrans:pageCurlDouble`

- **Attributes:** `invY`.
- **Look-alike:** fade.
- **Algorithm:** the fold template with two pages:
  - the left page is x ∈ [0, a/2], from P₀ = (0, 1) to P₁ = (a/2, 0);
  - the right page is x ∈ [a/2, a], from P₀ = (a, 1) to P₁ = (a/2, 0);
  - `invY` uses the top corners;
  - the reflection stays within its own page.

  The back of the page is the same as in pageCurlSingle.
- **Seeded parameters:** none.
- **Oracle:** `to` = 0.5 and back = 0.4475.

### `transition:p15:prstTrans:peelOff`

- **Attributes:** `invX`.
- **Look-alike:** fade.
- **Algorithm:** the fold template. P₀ is the bottom-left corner (0, 1) (`invX`: bottom-right) and P₁ is the opposite corner. The back is the sticker's back, solid (0.92, 0.92, 0.92). A lift shadow darkens `to` by 0.6 in the band s ∈ [L − 0.04, L].
- **Seeded parameters:** none.
- **Oracle:** `to` = 0.5 and back = 0.3291 (the mirror image of pageCurlSingle).

### `transition:p15:prstTrans:airplane` (D4 unverified)

- **Status:** a diagonal paper-dart look-alike (the S3-4 fix round 1 model, replacing a flat fold that did not read as a plane), awaiting owner PowerPoint capture (D4). **Clean-room:** the panel cut, the fold angles (π/4 dihedral, 0.9π wing fold) and the flight constants of G (0.9a sideways, −2.4 down, −3 deep, the 0.3 / 0.5 / 0.6 bank, yaw and pitch, the 0.4 shrink) were chosen by eye for a paper dart leaving the frame. No PowerPoint frame, binary or other implementation was consulted.
- **Attributes:** `invX`, which flies the dart away to the left.
- **Look-alike:** fade.
- **Panels** (`AIRPLANE_TRIS`). The slide is cut into four rigid triangles along two diagonal creases, from the top corners to the nose N at the bottom centre (uv (½, 1)), and the vertical spine from the top centre to N:
  - body left (0,0)–(½,0)–(½,1) and body right (½,0)–(1,0)–(½,1);
  - wing left (0,0)–(0,1)–(½,1) and wing right (1,0)–(1,1)–(½,1).

  Each triangle is one quad of a 1 × 1 grid whose fourth corner collapses onto its third, so the row draws 4 triangles plus `to`, with no deformable grid.
- **Motion** (`airplane`). The motion has two phases:
  - **Fold** (f = e(2p), complete at p = ½). The body halves open a dihedral about the spine, body left = Ry(−β) and body right = Ry(β), β = (π/4) f, so the outer edges recede. Each wing folds over its body half about its crease by θ = 0.9π f, towards the viewer: wing left = body left · R(N→TL, θ) and wing right = body right · R(N→TR, −θ), where R(A→B, θ) is the right-handed rotation by θ about the axis through A and B (`axisRotation`). At p = ½ the slide reads as a pointed paper dart, nose down, with its wings 18° off the body.
  - **Fly** (τ = e(2p − 1)). All four panels are then placed by G = T(0.9aσ τ, −2.4τ, −3τ) Rz(0.3σ τ) Ry(0.5σ τ) Rx(0.6τ) S(1 − 0.4τ), with σ = +1, or −1 for `invX`. This banks the dart about its long axis (Ry), pitches its nose into depth (Rx), yaws the nose towards the flight (Rz) and flies it nose-first down, sideways and away. It is never edge-on: its outline keeps area until it leaves the frame.
  - **Opacity:** 1 − e((p − 0.85)/0.15).
- **Draw:** to flat, then the two body halves, then the two wings (each pair farthest first). Wings shade by 1 − 0.1 f (the back of the paper, darkening as they fold).
- **Seeded parameters:** none.
- **Oracle:** at p = ½ the nose is at NDC (0, −1), the folded left wing's uv(0, 1) corner is at world (−0.0887, 0.3612, 0.2017), and the outline (the convex hull of the projected panel corners, `airplaneSilhouette`) has area 0.93 in NDC units (the frame is 4). At p = 0.75 the nose is at NDC (0.3986, −1.3538) and the outline's area is 0.1961.

### `transition:p15:prstTrans:origami`

- **Attributes:** `invX`, which hinges the chain at the right edge.
- **Look-alike:** fade.
- **Motion** (`origami`). Four vertical panels (`uvRect` [k/4, 0, ¼, 1]) form an accordion hinged in a chain from the left edge:
  - θ = 0.9 (π/2) e(p), and panel k's absolute angle is α_k = (−1)^k θ;
  - the hinge points are P₀ = (−a/2, 0) and P_{k+1} = P_k + (a/4)(cos α_k, −sin α_k) in (x, z);
  - panel k = T(P_k) Ry(α_k) T(−rest hinge x);
  - opacity is 1 − e((p − 0.7)/0.3).
- **Draw:** to flat, then the panels, farthest first.
- **Seeded parameters:** none.
- **Oracle:** panel 2's uv(½, 0) is at world (−0.213, 0.5, 0), and panel 3's uv(1, 0) is at NDC (0.5208, 1).

### `transition:p15:prstTrans:curtains`

- **Attributes:** none.
- **Look-alike:** split.
- **Mesh:** two 32 × 18 grids, one per half (`uvRect` [0,0,½,1] and [½,0,½,1]).
- **Motion** (`curtainX`). Each half gathers towards its outer edge x_o = ∓a/2:
  - x' = x_o + (x − x_o)(1 − 0.85 e(p)(0.75 + 0.25v)) ± 0.2a e((p − 0.6)/0.4), where the last term slides the half outwards;
  - folds: z = 0.035 e(p) sin(16π u_l), where u_l is the vertex's uv across its half, measured from the outer edge.

  By p = 1 both halves are off the slide.
- **Draw:** to flat, then the curtains.
- **Seeded parameters:** none.
- **Oracle:** x at uv(½, ½) (the left half's inner edge) is −0.3306, and x at uv(¼, 1) is −0.6333. At p = 1, x at uv(½, 0) is −0.9222 (< −a/2, off the slide).

### `transition:p15:prstTrans:drape` (D4 unverified)

- **Attributes:** none.
- **Look-alike:** fade.
- **Mesh:** a 32 × 18 grid.
- **Motion** (`drapeY`). The slide falls like cloth:
  - y' = y − 1.4 e(p)(0.8 + 0.2(½ − y)), so the bottom edge falls faster;
  - a ripple z = 0.08 s(p) sin(3πu + 2πp)(½ − y) and a sway x += 0.03 s(p) sin(2πv) are added.
- **Draw:** to flat, then from.
- **Seeded parameters:** none.
- **Oracle:** y at uv(½, 0) is −0.06, and y at uv(½, 1) is −1.2.

### `transition:p15:prstTrans:wind` (D4 unverified)

- **Attributes:** `invX`, which blows the wind to the right.
- **Look-alike:** fade.
- **Motion** (`windTile`). An 8 × 18 grid, W = 0.55, J = 0.5, g = 0.6(1 − u_c) + 0.4 v_c (mirrored for `invX`).
  - tile = T(c + (−1.6a t², 0.25 sin(πt)(r₁ − ½), 0)) Rz(0.6(r₂ − ½)t) Ry(−(π/3)t) T(−c);
  - opacity is fade(t).
- **Seeded parameters:** r₀ to r₂ per cell.
- **Oracle:** rest = 0 and done = 0. Cell (7, 0) has t = 0.4831, and its centre is at NDC (0.1281, 1.1572).

### `transition:p15:prstTrans:fallOver` (D4 unverified)

- **Attributes:** `invX`, which mirrors the lean.
- **Look-alike:** fade.
- **Motion** (`fallOver`). The slide falls backwards about its bottom edge:
  - θ = (π/2) e(p);
  - from = T(0, −½, 0) Rz(0.08σ s(p)) Rx(−θ) T(0, ½, 0), with σ = +1, or −1 for `invX`;
  - opacity is 1 − e((p − 0.8)/0.2).
- **Draw:** to flat, then from.
- **Seeded parameters:** none.
- **Oracle:** from uv(½, 0) is at world (−0.0565, 0.2048, −0.7071) and NDC (−0.0401, 0.2584).

### `transition:p15:prstTrans:prestige`

- **Attributes:** none.
- **Look-alike:** fade.
- **Motion** (`prestige`):
  - from = T(0, 0, −1.2e) T(0, −½, 0) Rx(−0.6e) T(0, ½, 0), with opacity 1 − e;
  - to = T(0, −1.1(1 − e), 0), which rises from below the slide and is drawn over from.
- **Draw:** from, then to.
- **Seeded parameters:** none.
- **Oracle:** from uv(½, 0) is at NDC (0, 0.5228), to uv(½, 0) is at (0, −0.1), and from's opacity is 0.5.

### `transition:p14:warp` (D4 unverified)

- **Attributes:** `dir` is `in` (default) or `out`.
- **Look-alike:** fade.
- **Mesh:** a 32 × 18 grid, deformed by `a_uv`.
- **Motion** (`warpK`):
  - **`in`:** each `from` vertex q (world xy) is pulled to the centre, q' = q · k, where k = (1 − e(p))^(1 + 2|q|). Its opacity is 1 − e((p − ½)/½), over `to` drawn flat at scale S(0.7 + 0.3e).
  - **`out`:** time runs backwards with the frames swapped. `to` is warped with k evaluated at 1 − p, with opacity 1 − e(((1 − p) − ½)/½), over `from` drawn flat at scale S(0.7 + 0.3(1 − e)).
- **Seeded parameters:** none.
- **Oracle (in):** k at uv(1, ½) is 0.1458, so that vertex's world x is 0.1296. k at the centre is 0.5.

## Oracle table (generated)

Generated by `bun scripts/slideshow-3d-oracles.ts` at aspect 16 : 9, field of view 45°, effect seed 2³¹ and p = 0.5. Group tests assert positions to 1e-3 (float32 GPU maths) and fractions to 0.01. Fractions measured from a rendered frame are fractions of pixels, not of cells.

| Row | Task | Attributes | Oracle at 50 % |
| --- | --- | --- | --- |
| transition:p14:prism | S3-2 | dir l | from uv(1,0) NDC = (0, 1.4389); to uv(1,1) NDC = (0.8145, -0.5759); isInverted=1 from uv(1,0) NDC = (0, 0.7663); dir u from uv(0,0) NDC = (-0.7071, 1) |
| transition:p14:flip | S3-2 | dir l | from uv(0,0) NDC = (0, 0.465); from uv(0,0) world = (0, 0.5, -1.3889) |
| transition:p14:gallery | S3-2 | dir l | from uv(1,0.5) NDC = (-0.1291, 0); to uv(0,0.5) NDC = (0.0918, 0) |
| transition:p14:doors | S3-2 | dir vert | left half uv(0.5,0) NDC = (-0.6111, 2.0864); to uv(0,0) NDC = (-0.8009, 0.8009); horz top half uv(0,0.5) NDC = (-1.4142, 0.4142) |
| transition:p14:window | S3-2 | dir vert | left half uv(0.5,0) NDC = (-1.1327, 2.0864); to uv(0,0) NDC = (-0.668, 0.668) |
| transition:p14:switch | S3-2 | dir l | from uv(0.5,0.5) NDC = (-0.7348, 0); to uv(0.5,0.5) NDC = (0.7348, 0); from uv(1,0) NDC = (-0.254, 0.9248) |
| transition:p14:conveyor | S3-2 | dir l | from uv(1,0.5) NDC = (-0.0885, 0); to uv(0,0.5) NDC = (0.0711, 0) |
| transition:p14:ferris | S3-2 | dir l | from uv(0.5,0.5) NDC = (-0.7208, -0.3434); to uv(0.5,0.5) NDC = (0.7208, -0.3434) |
| transition:p14:vortex | S3-3 | dir l | rest = 0.059; done = 0.059; cell (0,0) centre NDC = (-4.2778, 0.8061); cell (0,0) t = 0.8158 |
| transition:p14:shred | S3-3 | pattern strip, dir out | rest = 0; done = 0; strip 0 t = 0.8404; strip 0 dy = 0.9181; rectangle rest = 0; rectangle done = 0 |
| transition:p15:prstTrans:fracture | S3-3 | — | rest = 0.0714; done = 0.0238; cell (0,0) centre NDC = (-0.9167, 0.8571) |
| transition:p15:prstTrans:crush | S3-3 | — | rest = 0; done = 0; cell (0,0) centre NDC = (-0.6738, 0.6388); opacity = 1 |
| transition:p14:glitter | S3-3 | pattern diamond, dir l | rest = 0.1372; done = 0.1319; cells showing to = 0.4896; cell (0,0) t = 0.8183 |
| transition:p14:honeycomb | S3-3 | — | rest = 0.0455; done = 0.0318; cell (10,5) t = 1; cell (0,0) t = 0 |
| transition:p14:ripple | S3-3 | dir center | z at uv(0.75,0.5) = 0.0303; z at uv(0.5,0.5) = 0; from opacity = 0.5; dir lu z at uv(0.25,0.25) = -0.0069 |
| transition:p15:prstTrans:pageCurlSingle | S3-4 | — | fraction to = 0.5; fraction back = 0.3291 |
| transition:p15:prstTrans:pageCurlDouble | S3-4 | — | fraction to = 0.5; fraction back = 0.4475 |
| transition:p15:prstTrans:peelOff | S3-4 | — | fraction to = 0.5; fraction back = 0.3291 |
| transition:p15:prstTrans:airplane | S3-4 | — | nose uv(0.5,1) NDC = (0, -1); left wing uv(0,1) world = (-0.0887, 0.3612, 0.2017); silhouette area (NDC) = 0.93; p 0.75 nose NDC = (0.3986, -1.3538); p 0.75 silhouette area (NDC) = 0.1961 |
| transition:p15:prstTrans:origami | S3-4 | — | panel 2 uv(0.5,0) world = (-0.213, 0.5, 0); panel 3 uv(1,0) NDC = (0.5208, 1) |
| transition:p15:prstTrans:curtains | S3-4 | — | x at uv(0.5,0.5) (left half) = -0.3306; x at uv(0.25,1) = -0.6333; p 1 x at uv(0.5,0) = -0.9222 |
| transition:p15:prstTrans:drape | S3-4 | — | y at uv(0.5,0) = -0.06; y at uv(0.5,1) = -1.2 |
| transition:p15:prstTrans:wind | S3-4 | — | rest = 0; done = 0; cell (7,0) t = 0.4831; cell (7,0) centre NDC = (0.1281, 1.1572) |
| transition:p15:prstTrans:fallOver | S3-4 | — | from uv(0.5,0) NDC = (-0.0401, 0.2584); from uv(0.5,0) world = (-0.0565, 0.2048, -0.7071) |
| transition:p15:prstTrans:prestige | S3-4 | — | from uv(0.5,0) NDC = (0, 0.5228); to uv(0.5,0) NDC = (0, -0.1); from opacity = 0.5 |
| transition:p14:warp | S3-4 | dir in | k at uv(1,0.5) = 0.1458; world x at uv(1,0.5) = 0.1296; k at uv(0.5,0.5) = 0.5 |
