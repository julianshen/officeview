/**
 * The WebGL2 compositor's one program (slideshow spec 6.3), plus the pure
 * maths that feeds its uniforms.
 *
 * - **Vertex.** A unit quad `[0, 1]²`; `u_matrix` maps it to output device
 *   px (layer bounds, target transform, motion and viewport all composed on
 *   the CPU), and the projection flips y so device row 0 is the top, as in
 *   Canvas 2D. Textures are uploaded unflipped, so quad uv `(0, 0)` is the
 *   bitmap's first row.
 * - **Fragment.** Samples the premultiplied layer. With a colour matrix it
 *   un-premultiplies (guarding `a = 0`), applies the 4×5 matrix (offsets
 *   already in [0, 1]), clamps and re-premultiplies, the same order as the
 *   Canvas 2D backend's `applyColorMatrix`. The mask is coverage in `.r`:
 *   mode 1 samples it at the quad's uv placed in `u_maskRect` (a layer mask,
 *   which moves with the layer: `[0, 0, 1, 1]` over the layer's own box, or
 *   the layer's place inside its shape's box for a split text item), mode 2 at the fragment's device position over `u_maskRect` (a
 *   transition mask, fixed to the incoming frame's device rect), and 0
 *   outside that rect. The result is premultiplied, × opacity × coverage, for
 *   `blendFunc(ONE, ONE_MINUS_SRC_ALPHA)`.
 */
import type { Affine, ColorMatrix } from '../types'

export const VERTEX_SHADER = `#version 300 es
in vec2 a_unit;
uniform mat3 u_matrix;
uniform vec2 u_output;
out vec2 v_uv;
void main() {
  vec2 p = (u_matrix * vec3(a_unit, 1.0)).xy;
  v_uv = a_unit;
  gl_Position = vec4(p.x / u_output.x * 2.0 - 1.0, 1.0 - p.y / u_output.y * 2.0, 0.0, 1.0);
}
`

export const FRAGMENT_SHADER = `#version 300 es
precision highp float;
in vec2 v_uv;
uniform sampler2D u_layer;
uniform sampler2D u_mask;
uniform int u_maskMode;
uniform vec4 u_maskRect;
uniform vec2 u_output;
uniform float u_opacity;
uniform bool u_tint;
uniform mat4 u_colorM;
uniform vec4 u_colorOff;
out vec4 outColor;
void main() {
  vec4 c = texture(u_layer, v_uv);
  if (u_tint) {
    vec4 s = c.a > 0.0 ? vec4(c.rgb / c.a, c.a) : vec4(0.0);
    vec4 r = clamp(u_colorM * s + u_colorOff, 0.0, 1.0);
    c = vec4(r.rgb * r.a, r.a);
  }
  float m = 1.0;
  if (u_maskMode == 1) {
    vec2 t = (v_uv - u_maskRect.xy) / u_maskRect.zw;
    m = (t.x < 0.0 || t.y < 0.0 || t.x > 1.0 || t.y > 1.0) ? 0.0 : texture(u_mask, t).r;
  } else if (u_maskMode == 2) {
    vec2 d = vec2(gl_FragCoord.x, u_output.y - gl_FragCoord.y);
    vec2 t = (d - u_maskRect.xy) / u_maskRect.zw;
    m = (t.x < 0.0 || t.y < 0.0 || t.x > 1.0 || t.y > 1.0) ? 0.0 : texture(u_mask, t).r;
  }
  outColor = c * (u_opacity * m);
}
`

/** The uniforms the program reads, looked up once per (re)link. */
export const UNIFORMS = [
  'u_matrix', 'u_output', 'u_layer', 'u_mask', 'u_maskMode', 'u_maskRect', 'u_opacity', 'u_tint', 'u_colorM', 'u_colorOff',
] as const
export type UniformName = typeof UNIFORMS[number]

/** `m1 ∘ m2`: apply `m2`, then `m1`. */
export function mulAffine(m1: Affine, m2: Affine): Affine {
  const [a, b, c, d, e, f] = m1, [A, B, C, D, E, F] = m2
  return [a * A + c * B, b * A + d * B, a * C + c * D, b * C + d * D, a * E + c * F + e, b * E + d * F + f]
}

/**
 * The column-major `mat3` taking the unit quad to device px: `m` applied to
 * the layer's bounds box (`unit → bounds → m`).
 */
export function quadMatrix(m: Affine, bounds: { x: number; y: number; w: number; h: number }): Float32Array {
  const [a, b, c, d, e, f] = mulAffine(m, [bounds.w, 0, 0, bounds.h, bounds.x, bounds.y])
  return new Float32Array([a, b, 0, c, d, 0, e, f, 1])
}

/**
 * A 4×5 row-major colour matrix as the column-major `mat4` plus offset `vec4`
 * the shader multiplies with: `out = M · rgba + off`, values and offsets in
 * [0, 1]. Non-finite entries count as 0, as in `applyColorMatrix`.
 */
export function colorUniforms(m: ColorMatrix): { matrix: Float32Array; offset: Float32Array } {
  const k = (i: number) => (Number.isFinite(m[i]) ? m[i] : 0)
  const matrix = new Float32Array(16)
  for (let row = 0; row < 4; row++) for (let col = 0; col < 4; col++) matrix[col * 4 + row] = k(row * 5 + col)
  return { matrix, offset: new Float32Array([k(4), k(9), k(14), k(19)]) }
}
