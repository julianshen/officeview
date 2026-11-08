/** Methods/properties recorded by §7.3/§7.5, verbatim from the spec text. */
export const RECORDED_SUBSET = {
  /** Verbatim §7.3/§7.5 methods. `getTransform` is the §7.3
   * "getTransform-consumers" entry (re-reads are recorded state reads). */
  methods: new Set([
    'setTransform', 'translate', 'scale', 'rotate', 'transform',
    'getTransform',
    'createLinearGradient', 'createRadialGradient', 'createPattern',
    'beginPath', 'moveTo', 'lineTo', 'quadraticCurveTo', 'bezierCurveTo',
    'arc', 'arcTo', 'ellipse', 'rect', 'roundRect', 'closePath',
    'fill', 'stroke', 'fillRect', 'strokeRect', 'clearRect', 'clip',
    'fillText', 'strokeText', 'measureText', 'drawImage', 'rasterBlock',
    'setLineDash', 'getLineDash', // dash recorded in STATE_SETUP (§7.5)
    // save/restore: §7.5 STATE_PUSH/STATE_POP
    'save', 'restore',
  ]),
  /** Verbatim §7.3/§7.5 state/property setters (STATE_SETUP fields). */
  setters: new Set([
    'globalAlpha', 'globalCompositeOperation',
    'imageSmoothingEnabled', 'imageSmoothingQuality',
    'lineWidth', 'lineCap', 'lineJoin', 'miterLimit', 'lineDashOffset',
    'letterSpacing', 'wordSpacing', 'direction', 'lang',
    'fillStyle', 'strokeStyle',
    'shadowColor', 'shadowBlur', 'shadowOffsetX', 'shadowOffsetY',
    'font', 'textAlign', 'textBaseline',
  ]),
} as const

